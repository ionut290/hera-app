(() => {
  'use strict';

  // PROTEZIONE PERMANENTE
  // Questo modulo non deve mai cancellare, svuotare o sovrascrivere in massa
  // la collezione Firestore "personale".
  window.__personaleDestructiveClearDisabled = true;
  window.__personaleAutomaticRestoreBlocked = true;
})();
