(() => {
  'use strict';
  if (window.__vargaRubricaPermissionsBridge) return;
  window.__vargaRubricaPermissionsBridge = true;

  const SOURCE_COLLECTION = 'rubricaContacts';
  const COMPATIBLE_COLLECTION = 'safetyContacts';

  function patchFirestore() {
    const Firestore = window.firebase?.firestore?.Firestore;
    const proto = Firestore?.prototype;
    if (!proto || typeof proto.collection !== 'function') return false;
    if (proto.collection.__vargaRubricaRedirect) return true;

    const original = proto.collection;
    const wrapped = function collectionWithRubricaCompatibility(path) {
      const resolved = String(path || '') === SOURCE_COLLECTION ? COMPATIBLE_COLLECTION : path;
      return original.call(this, resolved);
    };
    wrapped.__vargaRubricaRedirect = true;
    wrapped.__vargaOriginal = original;
    proto.collection = wrapped;
    console.info('Rubrica: raccolta compatibile Firestore attiva.');
    return true;
  }

  if (!patchFirestore()) {
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (patchFirestore() || attempts >= 80) clearInterval(timer);
    }, 100);
  }
})();