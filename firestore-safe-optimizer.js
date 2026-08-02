(() => {
  'use strict';

  // Ottimizzatore disattivato: Firestore usa integralmente il comportamento
  // originale dell'app. Questo evita interferenze con il salvataggio e il
  // ripristino della data amministrativa delle squadre per commessa.
  window.__vargaFirestoreSafeOptimizer = true;
  window.VargaFirestoreSafeOptimizer = Object.freeze({
    enabled: false,
    pendingCount: () => 0,
    clear: () => {}
  });

  console.info('Ottimizzatore Firestore disattivato: comportamento originale ripristinato.');
})();
