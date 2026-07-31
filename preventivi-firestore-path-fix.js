(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');
  if (PV.firestorePathFixInstalled) return;
  PV.firestorePathFixInstalled = true;

  // Firestore riserva gli identificatori racchiusi tra doppi underscore.
  // Il precedente ID __preventivi__ veniva quindi rifiutato prima ancora
  // dell'applicazione delle regole di sicurezza.
  PV.collections.priceLists = 'commesse/preventivi_app/prezziari';
  PV.collections.quotes = 'commesse/preventivi_app/preventivi';

  const MIGRATION_KEY = 'varga_preventivi_firestore_valid_path_v1';
  const originalLoadLocal = PV.loadLocal.bind(PV);

  PV.loadLocal = () => {
    originalLoadLocal();

    try {
      if (localStorage.getItem(MIGRATION_KEY) === '1') return;

      // I dati locali creati mentre il percorso non era valido vengono
      // rimessi automaticamente in coda verso il nuovo percorso Firestore.
      PV.state.priceLists = (PV.state.priceLists || []).map((item) => ({
        ...item,
        syncPending: true
      }));
      PV.state.quotes = (PV.state.quotes || []).map((item) => ({
        ...item,
        syncPending: true
      }));
      PV.persistLocal();
      localStorage.setItem(MIGRATION_KEY, '1');
    } catch (error) {
      console.warn('Preventivi: migrazione al percorso Firestore valido non memorizzata.', error);
    }
  };
})();
