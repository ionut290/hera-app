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

  firebase.auth().onAuthStateChanged((user) => {
    if (!user) return;
    firebase.firestore().collection("mezzi").onSnapshot(cleanupSnapshot, (error) => {
      console.error("Impossibile controllare l'alimentazione dei mezzi", error);
    });
  });
})();
