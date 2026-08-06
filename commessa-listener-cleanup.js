(() => {
  'use strict';

  if (window.__heraCommessaListenerCleanupInstalled) return;

  const originalSubscribeImpianti = window.subscribeImpianti;
  const originalStopImpiantiSubscription = window.stopImpiantiSubscription;
  const originalSubscribeCommessaNotes = window.subscribeCommessaNotes;
  const originalStopCommessaNotesSubscription = window.stopCommessaNotesSubscription;

  if (typeof originalSubscribeImpianti !== 'function' ||
      typeof originalSubscribeCommessaNotes !== 'function') {
    console.warn('Pulizia listener commessa non installata: funzioni principali non disponibili.');
    return;
  }

  let impiantiSubscriptionStarted = false;
  let notesSubscriptionStarted = false;

  window.subscribeImpianti = function subscribeImpiantiWithCleanup(...args) {
    if (impiantiSubscriptionStarted && typeof originalStopImpiantiSubscription === 'function') {
      originalStopImpiantiSubscription();
    }

    const result = originalSubscribeImpianti.apply(this, args);
    impiantiSubscriptionStarted = true;
    return result;
  };

  window.subscribeCommessaNotes = function subscribeCommessaNotesWithCleanup(...args) {
    if (notesSubscriptionStarted && typeof originalStopCommessaNotesSubscription === 'function') {
      originalStopCommessaNotesSubscription();
    }

    const result = originalSubscribeCommessaNotes.apply(this, args);
    notesSubscriptionStarted = true;
    return result;
  };

  window.__heraCommessaListenerCleanupInstalled = true;
  console.info('Pulizia listener commessa installata.');
})();
