/* Modulo Ferie / Disponibilità personale - estratto dal core senza modifiche funzionali. */
(function (global) {
  "use strict";

function getFerieEligibleOperators() {
  const commesseNames = Array.from(commesseById.values()).map((c) => String(c?.nome || "").trim()).filter(Boolean);
  return personaleRecords.filter((person) => {
    const allEnabled = Boolean(person?.abilitatoTutteCommesse || person?.allCommesseEnabled);
    if (allEnabled) return true;
    const enabled = Array.isArray(person?.commesseAbilitate)
      ? person.commesseAbilitate.map((v) => String(v || "").trim()).filter(Boolean)
      : [];
    if (!enabled.length) return false;
    if (!commesseNames.length) return true;
    return commesseNames.some((commessaName) => isPersonAbilitataForCommessa(person, commessaName));
  });
}

function refreshFerieProgrammazioneUi() {
  refreshFerieOperatorOptions();
}

function refreshFerieOperatorOptions() {
  if (!ui.ferieOperatore) return;
  const people = getFerieEligibleOperators();
  const prev = ui.ferieOperatore.value;
  ui.ferieOperatore.innerHTML = '<option value="">Operatore</option>' + people
    .map((p) => getPersonaleDisplayName(p)).filter(Boolean).sort((a,b)=>a.localeCompare(b,'it'))
    .map((name)=>`<option value="${escapeHTML(name)}">${escapeHTML(name)}</option>`).join("");
  if (prev) ui.ferieOperatore.value = prev;
}

async function saveFerieCollega(event) {
  event.preventDefault();
  if (!canManageData()) return;
  const operatore = String(ui.ferieOperatore?.value || "").trim();
  const dataInizio = String(ui.ferieInizio?.value || "").trim();
  const dataFine = String(ui.ferieFine?.value || "").trim();
  const note = String(ui.ferieNote?.value || "").trim();
  if (!operatore || !dataInizio || !dataFine) return alert('Compila tutti i campi obbligatori ferie.');
  if (dataFine < dataInizio) return alert('La data fine ferie deve essere successiva o uguale alla data inizio.');
  await db.collection('ferieColleghi').add({ operatore, dataInizio, dataFine, note, createdAt: firebase.firestore.FieldValue.serverTimestamp(), createdBy: currentUser?.email || '' });
  ui.ferieForm?.reset();
  renderFerieList();
}

function computeDayStats(dateKey, ferieItems) {
  const enabledPeople = getFerieEligibleOperators();
  const inFerie = new Set(ferieItems.filter((f) => f.dataInizio <= dateKey && f.dataFine >= dateKey).map((f) => normalizeSafetyKey(f.operatore)));
  const available = enabledPeople.filter((p) => !inFerie.has(normalizeSafetyKey(getPersonaleDisplayName(p))));
  const reqCounts = {
    primo: available.filter((p) => hasRequiredPersonaleCourse(p, 'primo soccorso')).length,
    antincendio: available.filter((p) => hasRequiredPersonaleCourse(p, 'antincendio')).length,
    preposto: available.filter((p) => hasRequiredPersonaleCourse(p, 'preposto')).length
  };
  const byPeople = Math.floor(available.length / 2);
  const validTeams = Math.max(0, Math.min(byPeople, reqCounts.primo, reqCounts.antincendio, reqCounts.preposto));
  return { enabledPeople, inFerie, available, validTeams };
}

async function renderFerieList() {
  if (!ui.ferieList) return;
  if (!canManageData()) { ui.ferieList.innerHTML = "<p class='muted'>Solo admin può gestire ferie.</p>"; return; }
  const snap = await db.collection('ferieColleghi').orderBy('dataInizio','asc').get().catch(()=>null);
  if (!snap) return;
  const rows = snap.docs.map((d)=>({id:d.id,...d.data()}));
  ui.ferieList.innerHTML = rows.map((r)=>`<article class='simple-list-item'><strong>${escapeHTML(r.operatore||'-')}</strong><p>${escapeHTML(r.dataInizio||'-')} → ${escapeHTML(r.dataFine||'-')}</p><p>${escapeHTML(r.note||'')}</p><div class='item-actions'><button type='button' class='btn' data-edit-ferie='${escapeHTML(r.id)}'>Modifica</button><button type='button' class='btn btn-danger' data-del-ferie='${escapeHTML(r.id)}'>Elimina</button></div></article>`).join('') || "<p class='muted'>Nessuna ferie inserita.</p>";
  ui.ferieList.querySelectorAll('[data-del-ferie]').forEach((btn)=>btn.addEventListener('click', async()=>{
    if (!canManageData()) return;
    if (!confirm('Eliminare ferie?')) return;
    await db.collection('ferieColleghi').doc(btn.getAttribute('data-del-ferie')||'').delete();
    renderFerieList();
  }));
  ui.ferieList.querySelectorAll('[data-edit-ferie]').forEach((btn)=>btn.addEventListener('click', async()=>{
    if (!canManageData()) return;
    const id = btn.getAttribute('data-edit-ferie') || '';
    const row = rows.find((x)=>x.id===id);
    if (!row) return;
    const note = prompt('Modifica note ferie', row.note || '');
    if (note === null) return;
    await db.collection('ferieColleghi').doc(id).set({ note, updatedAt: firebase.firestore.FieldValue.serverTimestamp() }, { merge: true });
    renderFerieList();
  }));
}

async function renderFerieDisponibilitaCalendar() {
  if (!ui.ferieCalendarResult) return;
  if (!canManageData()) return;
  const start = String(ui.ferieCheckStart?.value || '').trim();
  const end = String(ui.ferieCheckEnd?.value || '').trim();
  if (!start || !end) return alert('Seleziona periodo.');
  if (end < start) return alert('Intervallo date non valido.');

  const ferieSnap = await db.collection('ferieColleghi').get();
  const ferieItems = ferieSnap.docs.map((d) => ({ id: d.id, ...d.data() }));

  const startDate = new Date(`${start}T00:00:00`);
  const endDate = new Date(`${end}T00:00:00`);
  const dayStats = new Map();
  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const key = d.toISOString().slice(0, 10);
    const stats = computeDayStats(key, ferieItems);
    const combos = buildTeamCombinations(stats.available, stats.validTeams);
    dayStats.set(key, { stats, combos });
  }

  const monthStart = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const monthEnd = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  const monthBlocks = [];

  for (let m = new Date(monthStart); m <= monthEnd; m.setMonth(m.getMonth() + 1)) {
    const year = m.getFullYear();
    const month = m.getMonth();
    const firstDay = new Date(year, month, 1);
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const startOffset = (firstDay.getDay() + 6) % 7;
    const monthLabel = firstDay.toLocaleDateString('it-IT', { month: 'long', year: 'numeric' });
    const cells = [];

    for (let i = 0; i < startOffset; i += 1) cells.push('<div class="ferie-month-cell ferie-month-cell--empty"></div>');

    for (let day = 1; day <= daysInMonth; day += 1) {
      const dayDate = new Date(year, month, day);
      const key = dayDate.toISOString().slice(0, 10);
      const inRange = dayDate >= startDate && dayDate <= endDate;
      const payload = inRange ? dayStats.get(key) : null;
      const validTeams = payload?.stats?.validTeams || 0;
      const statusClass = !inRange ? 'ferie-month-cell--out' : (validTeams > 0 ? 'ferie-month-cell--ok' : 'ferie-month-cell--ko');
      const detailId = `ferie-day-detail-${key}`;
      const stats = payload?.stats;
      const combos = payload?.combos || [];
      const comboRows = combos.map((team, idx) => `<li><b>Squadra ${idx + 1}</b>: ${team.map((p) => `${escapeHTML(getPersonaleDisplayName(p) || '-')}${escapeHTML(formatPersonReqBadges(p))}`).join(' + ')}</li>`).join('');
      const inFerieNames = stats
        ? stats.enabledPeople.filter((p) => stats.inFerie.has(normalizeSafetyKey(getPersonaleDisplayName(p)))).map((p) => getPersonaleDisplayName(p)).filter(Boolean)
        : [];
      const detail = stats ? `<div id="${escapeHTML(detailId)}" class="ferie-day-detail hidden"><p><b>Data:</b> ${escapeHTML(key)}</p><p>Abilitati: ${stats.enabledPeople.length} • In ferie: ${stats.inFerie.size} • Disponibili: ${stats.available.length}</p><p>✅ Squadre complete creabili: ${validTeams}</p><p>${validTeams === 0 && stats.available.length > 0 ? `⚠️ Persone disponibili ma requisiti mancanti: ${stats.available.length}` : '⚠️ Persone disponibili ma requisiti mancanti: 0'}</p><p>${validTeams === 0 ? '❌ Giorno scoperto' : ''}</p><p><b>Colleghi in ferie:</b> ${escapeHTML(inFerieNames.join(', ') || '-')}</p>${comboRows ? `<ul>${comboRows}</ul>` : '<p class="muted">Nessuna combinazione valida.</p>'}</div>` : '';
      cells.push(`<button type="button" class="ferie-month-cell ${statusClass}" ${inRange ? `data-ferie-toggle="${escapeHTML(detailId)}"` : 'disabled'}><span class="ferie-month-daynum">${day}</span>${inRange ? `<small>Sq: ${validTeams}</small>` : ''}</button>${detail}`);
    }

    monthBlocks.push(`<section class="ferie-month"><h5>${escapeHTML(monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1))}</h5><div class="ferie-month-weekdays"><span>Lun</span><span>Mar</span><span>Mer</span><span>Gio</span><span>Ven</span><span>Sab</span><span>Dom</span></div><div class="ferie-month-grid">${cells.join('')}</div></section>`);
  }

  ui.ferieCalendarResult.innerHTML = `<div class="ferie-month-wrap">${monthBlocks.join('')}</div>`;
  ui.ferieCalendarResult.querySelectorAll('[data-ferie-toggle]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-ferie-toggle') || '';
      const detail = ui.ferieCalendarResult.querySelector(`#${cssEscapeValue(id)}`);
      if (!detail) return;
      detail.classList.toggle('hidden');
    });
  });
}

  Object.assign(global, {
    getFerieEligibleOperators,
    refreshFerieProgrammazioneUi,
    refreshFerieOperatorOptions,
    saveFerieCollega,
    computeDayStats,
    renderFerieList,
    renderFerieDisponibilitaCalendar
  });
  global.VargaAvailabilityModule = Object.freeze({
    functions: Object.freeze(["getFerieEligibleOperators","refreshFerieProgrammazioneUi","refreshFerieOperatorOptions","saveFerieCollega","computeDayStats","renderFerieList","renderFerieDisponibilitaCalendar"])
  });
})(globalThis);
