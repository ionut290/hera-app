(() => {
  'use strict';
  if (window.HeraSquadContext?.installed) return;

  const VERSION = '1.0.0';
  let lastSignature = '';
  let refreshTimer = 0;
  let pollCount = 0;

  const text = (value) => String(value ?? '').trim();
  const upper = (value) => text(value).toLocaleUpperCase('it-IT');

  function todayKey() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function selectedCommessaIdValue() {
    try {
      if (typeof selectedCommessaId !== 'undefined' && selectedCommessaId) return text(selectedCommessaId);
    } catch (_) {}
    return text(window.selectedCommessaId);
  }

  function splitValues(value) {
    if (Array.isArray(value)) return value.flatMap(splitValues).filter(Boolean);
    if (value && typeof value === 'object') {
      return [value.nome, value.name, value.codice, value.code, value.id, value.nId]
        .map(text)
        .filter(Boolean);
    }
    return text(value).split(/[,;|\n]+/).map(text).filter(Boolean);
  }

  function unique(values) {
    const seen = new Set();
    return values.filter((value) => {
      const key = upper(value);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function getHistoryComposition(commessaId, dateKey) {
    try {
      if (typeof squadreHistoryByDate === 'undefined' || !(squadreHistoryByDate instanceof Map)) return null;
      const byCommessa = squadreHistoryByDate.get(dateKey);
      if (!(byCommessa instanceof Map)) return null;
      return byCommessa.get(commessaId) || null;
    } catch (_) {
      return null;
    }
  }

  function rowMatchesCurrentUser(row) {
    try {
      if (typeof doesSquadraMemberMatchCurrentUser !== 'function') return false;
      if (row?.caposquadra && doesSquadraMemberMatchCurrentUser(row.caposquadra)) return true;
      return splitValues(row?.personale ?? row?.operatori).some((member) => doesSquadraMemberMatchCurrentUser(member));
    } catch (_) {
      return false;
    }
  }

  function rowScore(row) {
    if (!row || typeof row !== 'object') return -1;
    let score = 0;
    if (text(row.caposquadra)) score += 4;
    score += splitValues(row.personale ?? row.operatori).length * 3;
    score += splitValues(row.mezzi ?? row.veicoli ?? row.attrezzature).length * 2;
    if (Array.isArray(row.impiantiDettagli) && row.impiantiDettagli.length) score += 1;
    return score;
  }

  function chooseSquadRow(composition) {
    const rows = Array.isArray(composition?.squadre) ? composition.squadre.filter(Boolean) : [];
    if (!rows.length) return null;
    const mine = rows.find(rowMatchesCurrentUser);
    if (mine) return mine;
    return [...rows].sort((a, b) => rowScore(b) - rowScore(a))[0] || rows[0];
  }

  function classifyVehicle(raw) {
    const value = upper(raw).replace(/\s+/g, ' ');
    const compact = value.replace(/[\s._-]+/g, '');
    if (/^A\d{1,6}/.test(compact)) return { code: compact.match(/^A\d{1,6}/)?.[0] || compact, kind: 'daily' };
    if (/^T\d{1,6}/.test(compact)) return { code: compact.match(/^T\d{1,6}/)?.[0] || compact, kind: 'trincia-big' };
    if (/^R\d{1,6}/.test(compact)) return { code: compact.match(/^R\d{1,6}/)?.[0] || compact, kind: 'trincia-small' };
    return { code: text(raw), kind: 'other' };
  }

  function semanticEquipment(vehicleValues) {
    const parts = [];
    vehicleValues.forEach((raw) => {
      const meta = classifyVehicle(raw);
      if (meta.kind === 'daily') parts.push(`${meta.code} DAILY IVECO`);
      else if (meta.kind === 'trincia-big') parts.push(`${meta.code} TRINCIA TRINCIATRICE TRATTORE GRANDE`);
      else if (meta.kind === 'trincia-small') parts.push(`${meta.code} TRINCIA TRINCIATRICE TRATTORE PICCOLO`);
      else if (meta.code) parts.push(meta.code);
    });
    return parts.join(' ');
  }

  function buildContext() {
    const commessaId = selectedCommessaIdValue();
    const dateKey = todayKey();
    if (!commessaId) return null;

    const composition = getHistoryComposition(commessaId, dateKey);
    const row = chooseSquadRow(composition);
    if (!row) return null;

    const operators = unique([
      ...splitValues(row.caposquadra),
      ...splitValues(row.personale ?? row.operatori)
    ]);
    const vehicles = unique(splitValues(row.mezzi ?? row.veicoli ?? row.attrezzature));
    const classified = vehicles.map(classifyVehicle);
    const equipmentText = semanticEquipment(vehicles);
    const teamSize = Math.max(1, operators.length || Number(row.numeroOperatori || row.teamSize || 0) || 2);

    return {
      source: 'squadreHistoryByDate',
      commessaId,
      data: dateKey,
      date: dateKey,
      giorno: dateKey,
      teamSize,
      numeroOperatori: teamSize,
      operatorCount: teamSize,
      operatori: operators,
      personale: operators.join(', '),
      caposquadra: text(row.caposquadra),
      mezzi: equipmentText,
      mezziOriginali: vehicles.join(', '),
      mezziCodici: classified.map((item) => item.code).filter(Boolean),
      equipmentText,
      hasDaily: classified.some((item) => item.kind === 'daily'),
      hasBigTrincia: classified.some((item) => item.kind === 'trincia-big'),
      hasSmallTrincia: classified.some((item) => item.kind === 'trincia-small'),
      hasTrincia: classified.some((item) => item.kind === 'trincia-big' || item.kind === 'trincia-small'),
      row,
      composition
    };
  }

  function refreshConsumers() {
    clearTimeout(refreshTimer);
    refreshTimer = window.setTimeout(() => {
      try {
        const state = window.HeraRecommendedPlants?.getState?.();
        if (state?.open) window.HeraRecommendedPlants.refresh?.();
      } catch (_) {}
      window.setTimeout(() => {
        try { window.HeraEquipmentAdvisor?.refresh?.(); } catch (_) {}
        try { window.HeraAdaptiveWorkLearning?.applyToRecommendedPanel?.(); } catch (_) {}
      }, 220);
    }, 80);
  }

  function sync(options = {}) {
    const context = buildContext();
    if (!context) return null;
    const signature = JSON.stringify({
      commessaId: context.commessaId,
      data: context.data,
      operatori: context.operatori,
      mezzi: context.mezziOriginali
    });

    // I moduli Impianti consigliati leggono currentSquadre per retrocompatibilità.
    // Qui gli forniamo la stessa composizione realmente salvata nella schermata Squadre.
    window.currentSquadre = [context];

    if (signature !== lastSignature || options.force) {
      lastSignature = signature;
      window.dispatchEvent(new CustomEvent('hera:squad-context-updated', { detail: {
        commessaId: context.commessaId,
        teamSize: context.teamSize,
        mezzi: context.mezziCodici,
        source: context.source
      }}));
      refreshConsumers();
    }
    return context;
  }

  function startShortPolling() {
    const timer = window.setInterval(() => {
      pollCount += 1;
      sync();
      if (pollCount >= 20) window.clearInterval(timer);
    }, 1500);
  }

  document.addEventListener('click', (event) => {
    if (event.target?.closest?.('#recommended-plants-btn, #recommended-origin-avola, #recommended-origin-live, #recommended-start-route')) {
      sync({ force: true });
    }
  }, true);
  window.addEventListener('focus', () => sync());
  document.addEventListener('visibilitychange', () => { if (!document.hidden) sync(); });
  window.addEventListener('hera:squad-context-request', () => sync({ force: true }));

  window.HeraSquadContext = {
    installed: true,
    version: VERSION,
    getCurrent: buildContext,
    sync,
    classifyVehicle
  };

  sync();
  startShortPolling();
})();