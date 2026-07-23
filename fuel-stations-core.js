(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.HeraFuelStations = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const AVAILABLE_VALUES = new Set(["yes", "true", "1", "available", "designated"]);
  const FUEL_LABELS = {
    cng: "metano",
    lpg: "GPL",
    diesel: "diesel",
    petrol: "benzina",
    electric: "ricarica elettrica"
  };

  function cleanText(value) {
    return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[_/,+-]+/g, " ").replace(/\s+/g, " ").trim();
  }

  function normalizeFuel(value) {
    const text = cleanText(value);
    // I vecchi ibridi vanno deliberatamente ricondotti al carburante principale.
    if (/\b(metano|cng|gnc)\b/.test(text)) return "cng";
    if (/\b(gpl|lpg)\b/.test(text)) return "lpg";
    if (/\b(diesel|gasolio)\b/.test(text)) return "diesel";
    if (/\b(benzina|petrol|gasoline)\b/.test(text)) return "petrol";
    if (/\b(elettric[oa]|electric|ev)\b/.test(text)) return "electric";
    return "";
  }

  function isAvailable(value) {
    return AVAILABLE_VALUES.has(cleanText(value));
  }

  function matchingFuel(tags, fuel) {
    const keys = Object.keys(tags || {}).filter((key) => isAvailable(tags[key]));
    const patterns = {
      cng: [/^fuel:(cng|compressed_natural_gas)$/],
      lpg: [/^fuel:(lpg|autogas)$/],
      diesel: [/^fuel:(diesel|hgv_diesel|biodiesel)$/],
      petrol: [/^fuel:(petrol|gasoline|e5|e10|octane_\d+)$/],
      electric: [/^socket:/, /^charging_station$/]
    };
    if (fuel === "electric" && cleanText(tags?.amenity) === "charging station") return FUEL_LABELS.electric;
    return keys.some((key) => (patterns[fuel] || []).some((pattern) => pattern.test(key.toLowerCase()))) ? FUEL_LABELS[fuel] : "";
  }

  function detectBrand(tags) {
    const text = cleanText([tags?.brand, tags?.name, tags?.operator, tags?.["brand:it"]].filter(Boolean).join(" "));
    if (/\bq8\b|kuwait petroleum/.test(text)) return "Q8";
    if (/\beni\b|\bagip\b/.test(text)) return "ENI";
    return "";
  }

  function formatAddress(tags) {
    const street = [tags?.["addr:street"], tags?.["addr:housenumber"]].filter(Boolean).join(" ");
    return [street, tags?.["addr:postcode"], tags?.["addr:city"] || tags?.["addr:place"]].filter(Boolean).join(", ") || "Indirizzo non disponibile";
  }

  function parseStations(elements, fuel, origin, distanceFn) {
    return (elements || []).map((item) => {
      const tags = item.tags || {};
      const lat = Number(item.lat ?? item.center?.lat);
      const lon = Number(item.lon ?? item.center?.lon);
      const brandLabel = detectBrand(tags);
      const availableFuel = matchingFuel(tags, fuel);
      if (!Number.isFinite(lat) || !Number.isFinite(lon) || !brandLabel || !availableFuel) return null;
      return {
        id: `${item.type || "node"}-${item.id}`,
        name: tags.name || tags.brand || tags.operator || "Distributore",
        brand: tags.brand || tags.operator || brandLabel,
        brandLabel,
        address: formatAddress(tags),
        availableFuel,
        lat,
        lon,
        distance: distanceFn(origin.lat, origin.lng, lat, lon)
      };
    }).filter(Boolean).sort((a, b) => a.distance - b.distance);
  }

  function buildQuery(lat, lng, radiusKm, fuel) {
    const radius = Math.round(radiusKm * 1000);
    const amenity = fuel === "electric" ? "charging_station" : "fuel";
    const brandPattern = "Q8|ENI|Agip|Kuwait Petroleum";
    const clauses = ["brand", "name", "operator"].flatMap((key) => ["node", "way"].map((type) => `${type}["amenity"="${amenity}"]["${key}"~"${brandPattern}",i](around:${radius},${lat},${lng});`)).join("");
    return `[out:json][timeout:25];(${clauses});out center tags;`;
  }

  return { FUEL_LABELS, normalizeFuel, matchingFuel, detectBrand, formatAddress, parseStations, buildQuery };
});
