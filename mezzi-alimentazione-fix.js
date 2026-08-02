(function () {
  "use strict";

  const CLEANUP_VERSION = "20260802-cost2";
  const CLEANUP_KEY_PREFIX = "hera_mezzi_alimentazione_cleanup_";

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

  let cleanupInProgress = false;

  async function cleanupSnapshot(snapshot) {
    if (cleanupInProgress) return;

    const updates = snapshot.docs.flatMap((doc) => {
      const current = String(doc.data()?.alimentazione || "").trim();
      const corrected = normalizeAlimentazione(current);
      return corrected !== current ? [{ ref: doc.ref, alimentazione: corrected }] : [];
    });
    if (!updates.length) return;

    cleanupInProgress = true;
    try {
      const firestore = firebase.firestore();
      for (let start = 0; start < updates.length; start += 450) {
        const batch = firestore.batch();
        updates.slice(start, start + 450).forEach(({ ref, alimentazione }) => {
          batch.update(ref, { alimentazione });
        });
        await batch.commit();
      }
      console.log(`Alimentazione mezzi corretta: ${updates.length} record aggiornati.`);
    } catch (error) {
      console.error("Errore durante la correzione dell'alimentazione dei mezzi", error);
    } finally {
      cleanupInProgress = false;
    }
  }

  function cleanupKey(user) {
    return `${CLEANUP_KEY_PREFIX}${CLEANUP_VERSION}_${user?.uid || "anonymous"}`;
  }

  firebase.auth().onAuthStateChanged(async (user) => {
    if (!user) return;
    const key = cleanupKey(user);
    try {
      if (localStorage.getItem(key) === "done") return;
      const snapshot = await firebase.firestore().collection("mezzi").get();
      await cleanupSnapshot(snapshot);
      localStorage.setItem(key, "done");
    } catch (error) {
      console.error("Impossibile controllare l'alimentazione dei mezzi", error);
    }
  });

  // Carica il miglioramento grafico direttamente nelle schede delle squadre della schermata Oggi.
  if (!document.querySelector('script[data-squadre-mezzi-pictograms]')) {
    const script = document.createElement("script");
    script.src = "./squadre-mezzi-pictograms.js?v=20260727a";
    script.defer = true;
    script.dataset.squadreMezziPictograms = "1";
    document.head.appendChild(script);
  }

  // Mantiene separata e protetta la logica del riepilogo Oggi: ore live e mezzi della squadra.
  if (!document.querySelector('script[data-today-live-hours-vehicles]')) {
    const script = document.createElement("script");
    script.src = "./today-live-hours-vehicles.js?v=20260730b";
    script.defer = true;
    script.dataset.todayLiveHoursVehicles = "1";
    document.head.appendChild(script);
  }

  // Rende cliccabili i nomi delle squadre senza intervenire sulla logica FATTO/WhatsApp.
  if (!document.querySelector('script[data-squad-operator-profile]')) {
    const script = document.createElement("script");
    script.src = "./squad-operator-profile.js?v=20260731a";
    script.defer = true;
    script.dataset.squadOperatorProfile = "1";
    document.head.appendChild(script);
  }
})();