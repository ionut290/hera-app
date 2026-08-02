(() => {
  'use strict';

  if (window.__heraFirestorePresenceCostGuardInstalled) return;
  window.__heraFirestorePresenceCostGuardInstalled = true;

  const HEARTBEAT_INTERVAL_MS = 15 * 60 * 1000;
  const originalStop = typeof window.stopPresenceHeartbeat === 'function'
    ? window.stopPresenceHeartbeat.bind(window)
    : null;
  const originalUpsert = typeof window.upsertCurrentPlatformUser === 'function'
    ? window.upsertCurrentPlatformUser.bind(window)
    : null;

  let timer = null;

  function clearGuardTimer() {
    if (timer !== null) window.clearInterval(timer);
    timer = null;
  }

  function stopPresenceHeartbeatOptimized() {
    clearGuardTimer();
    try { originalStop?.(); } catch (_) { /* Arresto già completato. */ }
  }

  function canWritePresence() {
    return Boolean(
      !document.hidden
      && navigator.onLine
      && window.firebase?.auth?.()?.currentUser
      && originalUpsert
    );
  }

  function startPresenceHeartbeatOptimized() {
    stopPresenceHeartbeatOptimized();
    if (!window.firebase?.auth?.()?.currentUser || document.hidden || !navigator.onLine) return;

    timer = window.setInterval(() => {
      if (!canWritePresence()) return;
      Promise.resolve(originalUpsert()).catch((error) => {
        console.warn('Aggiornamento presenza differito non riuscito:', error);
      });
    }, HEARTBEAT_INTERVAL_MS);
  }

  window.startPresenceHeartbeat = startPresenceHeartbeatOptimized;
  window.stopPresenceHeartbeat = stopPresenceHeartbeatOptimized;

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) stopPresenceHeartbeatOptimized();
    else startPresenceHeartbeatOptimized();
  });

  window.addEventListener('offline', stopPresenceHeartbeatOptimized);
  window.addEventListener('online', startPresenceHeartbeatOptimized);
  window.addEventListener('pagehide', stopPresenceHeartbeatOptimized);

  if (window.firebase?.auth) {
    window.firebase.auth().onAuthStateChanged((user) => {
      if (user) startPresenceHeartbeatOptimized();
      else stopPresenceHeartbeatOptimized();
    });
  }
})();
