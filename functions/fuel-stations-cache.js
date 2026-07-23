"use strict";

const CACHE_TTL_MS = 24 * 60 * 60 * 1000;
const ANAGRAFICA_URL = "https://www.mimit.gov.it/images/exportCSV/anagrafica_impianti_attivi.csv";
const PREZZI_URL = "https://www.mimit.gov.it/images/exportCSV/prezzo_alle_8.csv";

function cleanHeader(value) {
  return String(value || "")
    .replace(/^\uFEFF/, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function splitDelimitedLine(line, delimiter) {
  const fields = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (char === '"') {
      if (quoted && line[index + 1] === '"') {
        current += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
    } else if (char === delimiter && !quoted) {
      fields.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  fields.push(current.trim());
  return fields;
}

function detectDelimiter(line) {
  return (String(line).match(/\|/g) || []).length >= (String(line).match(/;/g) || []).length ? "|" : ";";
}

function parseDelimitedTable(text) {
  const lines = String(text || "").replace(/\r/g, "").split("\n").filter((line) => line.trim());
  const headerIndex = lines.findIndex((line) => /id\s*impianto/i.test(line));
  if (headerIndex < 0) throw new Error("Intestazione MIMIT non trovata");
  const delimiter = detectDelimiter(lines[headerIndex]);
  const headers = splitDelimitedLine(lines[headerIndex], delimiter).map(cleanHeader);
  const rows = [];
  for (let index = headerIndex + 1; index < lines.length; index += 1) {
    const values = splitDelimitedLine(lines[index], delimiter);
    if (!values.some(Boolean)) continue;
    const row = {};
    headers.forEach((header, column) => { row[header] = values[column] ?? ""; });
    rows.push(row);
  }
  const extractionMatch = lines.slice(0, headerIndex).join(" ").match(/\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}/);
  return { rows, extractionDate: extractionMatch?.[0] || "" };
}

function numberValue(value) {
  const parsed = Number(String(value ?? "").trim().replace(",", "."));
  return Number.isFinite(parsed) ? parsed : NaN;
}

function booleanValue(value) {
  return ["1", "true", "si", "yes"].includes(String(value || "").trim().toLowerCase());
}

function detectBrand(row) {
  const text = [row.bandiera, row.nomeimpianto, row.gestore].filter(Boolean).join(" ")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  if (/\bq8\b|kuwait petroleum/.test(text)) return "Q8";
  if (/\beni\b|\bagip\b/.test(text)) return "ENI";
  return "";
}

function isSupportedFuel(name) {
  const text = String(name || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  if (/\bgnl\b|\bhvo\b/.test(text)) return false;
  return /\b(metano|cng|gnc|gpl|lpg|gasolio|diesel|benzina|petrol|gasoline)\b/.test(text);
}

function buildNationalSnapshot(anagraficaText, prezziText, now = Date.now()) {
  const anagrafica = parseDelimitedTable(anagraficaText);
  const prezzi = parseDelimitedTable(prezziText);
  const fuelsByStation = new Map();

  prezzi.rows.forEach((row) => {
    const id = String(row.idimpianto || "").trim();
    const name = String(row.desccarburante || "").trim();
    if (!id || !name || !isSupportedFuel(name)) return;
    const price = numberValue(row.prezzo);
    const fuel = {
      name,
      price: Number.isFinite(price) ? price : null,
      isSelf: booleanValue(row.isself),
      dtComu: String(row.dtcomu || "").trim()
    };
    if (!fuelsByStation.has(id)) fuelsByStation.set(id, []);
    fuelsByStation.get(id).push(fuel);
  });

  const stations = anagrafica.rows.map((row) => {
    const id = String(row.idimpianto || "").trim();
    const lat = numberValue(row.latitudine);
    const lng = numberValue(row.longitudine);
    const brandLabel = detectBrand(row);
    const fuels = fuelsByStation.get(id) || [];
    if (!id || !brandLabel || !fuels.length || !Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    if (lat < 34 || lat > 48.5 || lng < 5 || lng > 20.5) return null;
    return {
      id,
      name: String(row.nomeimpianto || row.bandiera || brandLabel).trim(),
      brand: String(row.bandiera || brandLabel).trim(),
      address: [row.indirizzo, row.comune, row.provincia].filter(Boolean).join(", "),
      location: { lat, lng },
      fuels,
      insertDate: fuels.reduce((latest, fuel) => fuel.dtComu > latest ? fuel.dtComu : latest, "")
    };
  }).filter(Boolean);

  return {
    updatedAt: now,
    extractionDate: prezzi.extractionDate || anagrafica.extractionDate || "",
    stations
  };
}

async function fetchText(fetchImpl, url) {
  const response = await fetchImpl(url, { headers: { Accept: "text/csv,text/plain;q=0.9,*/*;q=0.8" } });
  if (!response.ok) throw Object.assign(new Error(`MIMIT CSV HTTP ${response.status}`), { status: response.status });
  return response.text();
}

async function downloadNationalSnapshot(fetchImpl = fetch) {
  const [anagraficaText, prezziText] = await Promise.all([
    fetchText(fetchImpl, ANAGRAFICA_URL),
    fetchText(fetchImpl, PREZZI_URL)
  ]);
  const snapshot = buildNationalSnapshot(anagraficaText, prezziText);
  if (!snapshot.stations.length) throw new Error("Archivio MIMIT nazionale vuoto");
  return snapshot;
}

module.exports = {
  CACHE_TTL_MS,
  ANAGRAFICA_URL,
  PREZZI_URL,
  cleanHeader,
  splitDelimitedLine,
  parseDelimitedTable,
  detectBrand,
  isSupportedFuel,
  buildNationalSnapshot,
  downloadNationalSnapshot
};
