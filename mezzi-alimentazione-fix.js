(function () {
  "use strict";

  function normalizeAlimentazione(value) {
    const original = String(value || "").trim();
    const normalized = original.toLowerCase();
    const hasBenzina = /\bbenzina\b/.test(normalized);
    const hasMetano = /\bmetano\b/.test(normalized);
    const hasGpl = /\bgpl\b/.test(normalized);

    if (!hasBenzina || (!hasMetano && !hasGpl)) return original;

    const alimentazioni = [];
    if (hasMetano) alimentazioni.push("METANO");
    if (hasGpl) alimentazioni.push("GPL");
    return alimentazioni.join(" + ");
  }

  const originalNormalizeMezzoDocument = window.normalizeMezzoDocument;
  if (typeof originalNormalizeMezzoDocument === "function") {
    window.normalizeMezzoDocument = function (doc) {
      const mezzo = originalNormalizeMezzoDocument(doc);
      return {
        ...mezzo,
        alimentazione: normalizeAlimentazione(mezzo.alimentazione)
      };
    };
  }

  document.getElementById("mezzi-form")?.addEventListener("submit", function () {
    const input = document.getElementById("mezzo-alimentazione");
    if (input) input.value = normalizeAlimentazione(input.value);
  }, true);

  // La bonifica storica è conclusa: nessuna scansione automatica della collezione mezzi.
  // La normalizzazione resta attiva durante lettura e salvataggio dei singoli mezzi.

  if (!document.querySelector('script[data-squadre-mezzi-pictograms]')) {
    const script = document.createElement("script");
    script.src = "./squadre-mezzi-pictograms.js?v=20260727a";
    script.defer = true;
    script.dataset.squadreMezziPictograms = "1";
    document.head.appendChild(script);
  }

  if (!document.querySelector('script[data-today-live-hours-vehicles]')) {
    const script = document.createElement("script");
    script.src = "./today-live-hours-vehicles.js?v=20260730b";
    script.defer = true;
    script.dataset.todayLiveHoursVehicles = "1";
    document.head.appendChild(script);
  }

  if (!document.querySelector('script[data-squad-operator-profile]')) {
    const script = document.createElement("script");
    script.src = "./squad-operator-profile.js?v=20260731a";
    script.defer = true;
    script.dataset.squadOperatorProfile = "1";
    document.head.appendChild(script);
  }
})();
