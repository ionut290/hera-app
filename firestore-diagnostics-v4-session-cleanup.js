(() => {
  'use strict';

  const today = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Europe/Rome' });
  const key = `varga_fs_diag_v4_${today()}`;
  const now = Date.now();
  const closedAt = new Date(now).toISOString();

  try {
    const value = JSON.parse(localStorage.getItem(key) || 'null');
    if (!value || value.date !== today() || !value.listeners) return;

    let recovered = 0;
    Object.values(value.listeners).forEach((listener) => {
      if (!listener?.active) return;
      listener.active = false;
      listener.abandoned = true;
      listener.closeReason = 'page-session-ended-without-unsubscribe';
      listener.closedAt = closedAt;
      listener.durationMs = listener.openedAt
        ? Math.max(0, now - new Date(listener.openedAt).getTime())
        : null;
      recovered += 1;
    });

    if (!recovered) return;
    value.totals = value.totals || {};
    value.totals.abandonedListenersRecovered =
      Math.max(0, Number(value.totals.abandonedListenersRecovered) || 0) + recovered;
    value.lifecycle = Array.isArray(value.lifecycle) ? value.lifecycle : [];
    value.lifecycle.unshift({
      at: closedAt,
      type: 'previous-page-listeners-recovered',
      recoveredListeners: recovered,
      note: 'Listener rimasti attivi nel salvataggio locale dopo ricaricamento, crash o chiusura della pagina.'
    });
    value.updatedAt = closedAt;
    localStorage.setItem(key, JSON.stringify(value));
  } catch (_) {
    // La pulizia è solo diagnostica e non deve mai bloccare l'avvio dell'app.
  }
})();
