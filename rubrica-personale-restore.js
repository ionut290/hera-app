(() => {
  'use strict';

  // Protezione permanente: questo file non deve eliminare o modificare
  // i documenti della collezione Firestore "personale".
  // Il precedente script di svuotamento è stato disattivato perché poteva
  // rieseguirsi su browser o dispositivi diversi.
  window.__personaleDestructiveClearDisabled = true;
})();
