(function installAdministrativeCalendar(global) {
  "use strict";

  if (global.HeraAdministrativeCalendar?.installed) return;

  const money = new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" });
  const CACHE_PREFIX = "heraImpiantiPersistentCacheV1:";
  const CACHE_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
  const monthlyViews = new Map();
  const recoveryByMonth = new Map();
  let activeMonth = "";
  let unsubscribeMonth = null;

  function escape(value) {
    if (typeof escapeHTML === "function") return escapeHTML(String(value ?? ""));
    return String(value ?? "").replace(/[&<>'"]/g, (character) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
    }[character]));
  }

  function text(value) {
    return String(value ?? "").trim();
  }

  function normalize(value) {
    return text(value).toLocaleLowerCase("it-IT").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ");
  }

  function number(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const cleaned = text(value).replace(/\s/g, "").replace(/\.(?=\d{3}(?:\D|$))/g, "").replace(",", ".").replace(/[^0-9.-]/g, "");
    if (!cleaned) return null;
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function dateKey(value) {
    if (typeof normalizeHoursReportDateKey === "function") {
      try {
        const normalized = normalizeHoursReportDateKey(value);
        if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) return normalized;
      } catch (_) {}
    }
    if (typeof firestoreDateToMillis === "function") {
      try {
        const millis = firestoreDateToMillis(value);
        if (millis) return formatCalendarDateKey(new Date(millis));
      } catch (_) {}
    }
    const raw = text(value);
    const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
    const italian = raw.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})/);
    return italian ? `${italian[3]}-${italian[2].padStart(2, "0")}-${italian[1].padStart(2, "0")}` : "";
  }

  function dateTimeMillis(value) {
    if (typeof firestoreDateToMillis === "function") {
      try { return firestoreDateToMillis(value) || 0; } catch (_) {}
    }
    const parsed = Date.parse(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function formatHours(value) {
    const minutes = Math.round(Math.max(0, Number(value) || 0) * 60);
    return `${String(Math.floor(minutes / 60)).padStart(2, "0")}:${String(minutes % 60).padStart(2, "0")}`;
  }

  function commessaName(commessaId, fallback = "") {
    const commessa = typeof commesseById !== "undefined" && commesseById instanceof Map ? commesseById.get(commessaId) : null;
    return text(commessa?.nome || commessa?.name || fallback || "Commessa non identificata");
  }

  function commessaCode(commessaId) {
    const commessa = typeof commesseById !== "undefined" && commesseById instanceof Map ? commesseById.get(commessaId) : null;
    return text(commessa?.codice || commessa?.code || "");
  }

  function reportEntries(report) {
    if (Array.isArray(report?.entries)) return report.entries;
    if (Array.isArray(report?.commesse)) return report.commesse;
    return [report || {}];
  }

  function entryRows(entry) {
    if (Array.isArray(entry?.rows)) return entry.rows;
    if (Array.isArray(entry?.operatori)) return entry.operatori;
    if (Array.isArray(entry?.operators)) return entry.operators;
    return [];
  }

  function rowHours(row) {
    for (const candidate of [row?.ore, row?.hours, row?.totaleOre, row?.oreTotali]) {
      const parsed = number(candidate);
      if (parsed != null && parsed > 0) return parsed;
    }
    return 0;
  }

  function operatorName(row, entry = {}) {
    const candidates = [
      row?.operatore, row?.operatoreNome, row?.nomeOperatore, row?.displayName, row?.nomeCompleto, row?.name,
      `${row?.nome || ""} ${row?.cognome || ""}`, entry?.operatore, entry?.operatoreNome
    ];
    return candidates.map(text).find(Boolean) || "Operatore non indicato";
  }

  function monthViewForDate(selectedDate) {
    return monthlyViews.get(text(selectedDate).slice(0, 7)) || null;
  }

  function collectHoursForDate(selectedDate) {
    const rows = [];
    const monthly = monthViewForDate(selectedDate);
    const reports = Array.isArray(monthly?.reports)
      ? monthly.reports
      : (typeof allHoursReports !== "undefined" && Array.isArray(allHoursReports) ? allHoursReports : []);
    reports.forEach((report) => {
      if (dateKey(report?.date || report?.data || report?.giorno || report?.workDate) !== selectedDate) return;
      reportEntries(report).forEach((entry) => {
        const id = text(entry?.commessaId || entry?.projectId || report?.commessaId || report?.projectId);
        const name = commessaName(id, entry?.commessaName || entry?.commessa || report?.commessaName || report?.commessa);
        entryRows(entry).forEach((row) => {
          const hours = rowHours(row);
          if (hours <= 0) return;
          rows.push({
            commessaId: id,
            commessaName: name,
            operator: operatorName(row, entry),
            operatorId: text(row?.operatoreId || row?.personaleId || row?.uid || row?.userId || row?.email),
            hours,
            note: text(row?.note || row?.nota || entry?.note || entry?.nota || report?.note || report?.nota)
          });
        });
      });
    });
    return rows;
  }

  function persistentCacheScopePrefix() {
    const uid = text(typeof currentUser !== "undefined" ? currentUser?.uid : "");
    if (!uid) return "";
    let collectionName = "commesse";
    try { collectionName = text(getCommesseCollectionName?.()) || collectionName; } catch (_) {}
    return `${CACHE_PREFIX}${encodeURIComponent(uid)}:${encodeURIComponent(collectionName)}:`;
  }

  function collectCachedImpianti() {
    const byCommessa = new Map();
    if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId instanceof Map) {
      impiantiByCommessaId.forEach((items, commessaId) => {
        if (Array.isArray(items) && items.length) byCommessa.set(text(commessaId), items.map((item) => ({ ...item })));
      });
    }

    const prefix = persistentCacheScopePrefix();
    if (!prefix) return byCommessa;
    try {
      for (let index = 0; index < localStorage.length; index += 1) {
        const key = localStorage.key(index) || "";
        if (!key.startsWith(prefix)) continue;
        const commessaId = decodeURIComponent(key.slice(prefix.length));
        if (!commessaId || byCommessa.has(commessaId)) continue;
        const cached = JSON.parse(localStorage.getItem(key) || "null");
        const age = Date.now() - Number(cached?.savedAt || 0);
        if (cached?.schemaVersion !== 1 || !Array.isArray(cached?.items) || !cached.items.length || age < 0 || age > CACHE_MAX_AGE_MS) continue;
        byCommessa.set(commessaId, cached.items.map((item) => ({ ...item })));
      }
    } catch (error) {
      console.warn("Cache locale impianti non disponibile per il calendario amministrativo:", error);
    }
    return byCommessa;
  }

  function isCompleted(plant) {
    if (typeof isImpiantoDoneState === "function") {
      try { return Boolean(isImpiantoDoneState(plant)); } catch (_) {}
    }
    const status = normalize(plant?.stato || plant?.status);
    return Boolean(plant?.done || plant?.fatto || plant?.completed || ["fatto", "done", "completed", "completato"].includes(status));
  }

  function completedDateKey(plant) {
    return dateKey(plant?.doneAt || plant?.completedAt || plant?.dataEsecuzione || plant?.dataFatto || plant?.executionDate);
  }

  function completedAmount(plant) {
    for (const candidate of [plant?.totale, plant?.importo, plant?.totaleRiga, plant?.importoPrestazione, plant?.valoreProdotto, plant?.earnedAmount]) {
      const parsed = number(candidate);
      if (parsed != null && parsed >= 0) return parsed;
    }
    return null;
  }

  function plantIdentity(plant, index) {
    return text(plant?.physicalPlantId || plant?.id || plant?.idSap || plant?.denominazione || plant?.nome || index);
  }

  function sharedActivityToPlant(activity) {
    return {
      commessaId: text(activity?.commessaId),
      commessaName: commessaName(text(activity?.commessaId), activity?.commessaName),
      id: text(activity?.itemId || activity?.sourceKey),
      impiantoId: text(activity?.impiantoId || activity?.itemId),
      idSap: text(activity?.idSap),
      name: text(activity?.name || (activity?.kind === "lavorazione" ? "Lavorazione" : "Impianto")),
      comune: text(activity?.comune),
      address: text(activity?.address),
      work: text(activity?.work || activity?.workCode),
      workCode: text(activity?.workCode),
      operator: text(activity?.operator),
      operatorId: text(activity?.operatorId),
      time: text(activity?.time),
      amount: number(activity?.amount),
      note: text(activity?.note),
      quantity: number(activity?.quantity),
      unit: text(activity?.unit),
      kind: text(activity?.kind || "impianto"),
      sourceKey: text(activity?.sourceKey),
      shared: true
    };
  }

  function collectSharedActivitiesForDate(selectedDate) {
    const view = monthViewForDate(selectedDate);
    const activities = Array.isArray(view?.activities) ? view.activities : [];
    const daily = activities.filter((activity) => dateKey(activity?.date) === selectedDate);
    const plantsWithWorkItems = new Set(daily
      .filter((activity) => activity?.kind === "lavorazione")
      .map((activity) => `${text(activity?.commessaId)}|${text(activity?.impiantoId)}`)
      .filter((key) => !key.endsWith("|")));
    return daily
      .filter((activity) => activity?.kind !== "impianto" || !plantsWithWorkItems.has(`${text(activity?.commessaId)}|${text(activity?.impiantoId || activity?.itemId)}`))
      .map(sharedActivityToPlant);
  }

  function collectPlantsForDate(selectedDate, cachedByCommessa = collectCachedImpianti()) {
    const cachedResult = [];
    cachedByCommessa.forEach((items, commessaId) => {
      const seen = new Set();
      items.forEach((plant, index) => {
        const identity = plantIdentity(plant, index);
        if (seen.has(identity) || !isCompleted(plant) || completedDateKey(plant) !== selectedDate) return;
        seen.add(identity);
        const completedMillis = dateTimeMillis(plant?.doneAt || plant?.completedAt);
        const time = text(plant?.oraEsecuzione || plant?.oraFatto) || (completedMillis
          ? new Date(completedMillis).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false })
          : "");
        cachedResult.push({
          commessaId,
          commessaName: commessaName(commessaId),
          id: identity,
          impiantoId: text(plant?.physicalPlantId || plant?.impiantoId || plant?.id || identity),
          idSap: text(plant?.idSap || plant?.sapId),
          name: text(plant?.denominazione || plant?.nome || "Impianto"),
          comune: text(plant?.comune),
          address: text(plant?.indirizzo || plant?.descrizioneVia),
          work: text(plant?.tipologiaIntervento || plant?.lavorazioniRichieste || plant?.lavorazione || plant?.codicePrezzo),
          operator: text(plant?.doneBy || plant?.operatoreNome || plant?.operatore || plant?.completedBy),
          time,
          amount: completedAmount(plant),
          note: text(plant?.noteImpianto || plant?.note),
          kind: "impianto",
          shared: false
        });
      });
    });
    const shared = collectSharedActivitiesForDate(selectedDate);
    if (!shared.length) return cachedResult;

    const cachedByPlant = new Map(cachedResult.map((plant) => [
      `${plant.commessaId}|${plant.impiantoId || plant.id}`,
      plant
    ]));
    const sharedPlantKeys = new Set();
    const enrichedShared = shared.map((plant) => {
      const key = `${plant.commessaId}|${plant.impiantoId || plant.id}`;
      const cached = cachedByPlant.get(key);
      sharedPlantKeys.add(key);
      if (!cached) return plant;
      return {
        ...plant,
        idSap: plant.idSap || cached.idSap,
        name: plant.name === "Lavorazione" ? cached.name : (plant.name || cached.name),
        comune: plant.comune || cached.comune,
        address: plant.address || cached.address,
        note: plant.note || cached.note
      };
    });
    return [
      ...enrichedShared,
      ...cachedResult.filter((plant) => !sharedPlantKeys.has(`${plant.commessaId}|${plant.impiantoId || plant.id}`))
    ];
  }

  function groupDayData(selectedDate, cachedByCommessa) {
    const groups = new Map();
    const ensure = (id, name) => {
      const key = id ? `id:${id}` : `name:${normalize(name) || "unknown"}`;
      let group = groups.get(key);
      if (!group) {
        const normalizedName = normalize(commessaName(id, name));
        group = Array.from(groups.values()).find((candidate) => normalize(candidate.commessaName) === normalizedName) || null;
      }
      if (!group) {
        group = { key, commessaId: id, commessaName: commessaName(id, name), commessaCode: commessaCode(id), hoursRows: [], plants: [] };
        groups.set(key, group);
      } else if (!group.commessaId && id) {
        group.commessaId = id;
        group.commessaName = commessaName(id, group.commessaName);
        group.commessaCode = commessaCode(id);
      }
      return group;
    };
    collectHoursForDate(selectedDate).forEach((row) => ensure(row.commessaId, row.commessaName).hoursRows.push(row));
    collectPlantsForDate(selectedDate, cachedByCommessa).forEach((plant) => ensure(plant.commessaId, plant.commessaName).plants.push(plant));
    return Array.from(groups.values()).map((group) => {
      const operators = new Map();
      group.hoursRows.forEach((row) => {
        const key = row.operatorId || normalize(row.operator);
        const current = operators.get(key) || { name: row.operator, hours: 0, notes: new Set() };
        current.hours += row.hours;
        if (row.note) current.notes.add(row.note);
        operators.set(key, current);
      });
      const knownEarnings = group.plants.reduce((sum, plant) => sum + (plant.amount ?? 0), 0);
      const missingAmounts = group.plants.filter((plant) => plant.amount == null).length;
      const missingOperators = group.plants.filter((plant) => !plant.operator).length;
      const completedUnits = new Set(group.plants.map((plant) => text(plant.impiantoId || plant.id)).filter(Boolean)).size;
      return {
        ...group,
        operators: Array.from(operators.values()).map((item) => ({ ...item, notes: Array.from(item.notes) })),
        totalHours: group.hoursRows.reduce((sum, row) => sum + row.hours, 0),
        knownEarnings,
        missingAmounts,
        missingOperators,
        completedUnits
      };
    }).sort((a, b) => b.knownEarnings - a.knownEarnings || b.totalHours - a.totalHours || a.commessaName.localeCompare(b.commessaName, "it"));
  }

  function monthDaySummaries(year, month, cachedByCommessa) {
    const summaries = new Map();
    const days = new Date(year, month + 1, 0).getDate();
    for (let day = 1; day <= days; day += 1) {
      const key = formatCalendarDateKey(new Date(year, month, day, 12));
      const groups = groupDayData(key, cachedByCommessa);
      const summary = {
        groups,
        plants: groups.reduce((sum, group) => sum + group.completedUnits, 0),
        activities: groups.reduce((sum, group) => sum + group.plants.length, 0),
        hours: groups.reduce((sum, group) => sum + group.totalHours, 0),
        earnings: groups.reduce((sum, group) => sum + group.knownEarnings, 0),
        missingAmounts: groups.reduce((sum, group) => sum + group.missingAmounts, 0)
      };
      summaries.set(key, summary);
    }
    return summaries;
  }

  function visibleMonthKey() {
    if (!(calendarVisibleMonth instanceof Date) || Number.isNaN(calendarVisibleMonth.getTime())) return "";
    return `${calendarVisibleMonth.getFullYear()}-${String(calendarVisibleMonth.getMonth() + 1).padStart(2, "0")}`;
  }

  function deactivate() {
    unsubscribeMonth?.();
    unsubscribeMonth = null;
    activeMonth = "";
  }

  function recoverMonthActivities(month) {
    const current = monthlyViews.get(month);
    if (Array.isArray(current?.activities) && current.activities.length) return Promise.resolve(current);
    if (recoveryByMonth.has(month)) return recoveryByMonth.get(month);

    const request = (async () => {
      const callable = global.firebase?.app?.().functions?.("europe-west1")?.httpsCallable?.("getAdministrativeCalendarMonth");
      if (!callable) return null;
      monthlyViews.set(month, { ...current, reports: current?.reports || [], activities: [], recovering: true });
      try {
        const response = await callable({ month });
        const data = response?.data || {};
        const latest = monthlyViews.get(month) || {};
        const recovered = {
          ...latest,
          reports: Array.isArray(latest.reports) ? latest.reports : [],
          activities: Array.isArray(data.activities) ? data.activities : [],
          source: "recupero-mensile-controllato",
          recovering: false,
          recoveryError: ""
        };
        monthlyViews.set(month, recovered);
        return recovered;
      } catch (error) {
        console.warn("Recupero mensile controllato non disponibile:", error);
        const latest = monthlyViews.get(month) || {};
        monthlyViews.set(month, { ...latest, recovering: false, recoveryError: text(error?.message || error) });
        return null;
      } finally {
        if (typeof calendarMode !== "undefined" && calendarMode === "administrative" && visibleMonthKey() === month) {
          renderCalendarGrid();
        }
      }
    })();
    recoveryByMonth.set(month, request);
    return request;
  }

  function ensureMonthView(month) {
    if (!/^\d{4}-\d{2}$/.test(month) || activeMonth === month) return;
    deactivate();
    activeMonth = month;
    const api = global.HeraSharedStaticViews;
    const cached = api?.getCached?.("calendario", month);
    if (cached?.payload) {
      monthlyViews.set(month, {
        reports: Array.isArray(cached.payload.reports) ? cached.payload.reports : [],
        activities: Array.isArray(cached.payload.activities) ? cached.payload.activities : [],
        source: "cache"
      });
    }
    if (!api?.subscribe) return;
    unsubscribeMonth = api.subscribe("calendario", month, (view, metadata = {}) => {
      if (!view?.payload || activeMonth !== month) return;
      monthlyViews.set(month, {
        reports: Array.isArray(view.payload.reports) ? view.payload.reports : [],
        activities: Array.isArray(view.payload.activities) ? view.payload.activities : [],
        source: metadata.source || "firestore"
      });
      if (!view.payload.activities?.length) recoverMonthActivities(month);
      if (typeof calendarMode !== "undefined" && calendarMode === "administrative" && visibleMonthKey() === month) {
        renderCalendarGrid();
      }
    });
    Promise.resolve().then(() => recoverMonthActivities(month));
  }

  function renderCalendarGrid() {
    if (!ui?.calendarGrid || !ui?.calendarMonthTitle) return;
    const year = calendarVisibleMonth.getFullYear();
    const month = calendarVisibleMonth.getMonth();
    const monthKey = visibleMonthKey();
    ensureMonthView(monthKey);
    const cachedByCommessa = collectCachedImpianti();
    const summaries = monthDaySummaries(year, month, cachedByCommessa);
    ui.calendarMonthTitle.textContent = new Intl.DateTimeFormat("it-IT", { month: "long", year: "numeric" }).format(calendarVisibleMonth);
    const firstDay = new Date(year, month, 1, 12);
    const gridStart = new Date(year, month, 1 - ((firstDay.getDay() + 6) % 7), 12);
    const today = formatCalendarDateKey(new Date());
    const cells = [];

    for (let index = 0; index < 42; index += 1) {
      const date = new Date(gridStart);
      date.setDate(gridStart.getDate() + index);
      const key = formatCalendarDateKey(date);
      const summary = summaries.get(key) || { groups: [], plants: 0, activities: 0, hours: 0, earnings: 0, missingAmounts: 0 };
      const hasData = summary.groups.length > 0;
      const classes = ["calendar-day", "administrative-calendar-day", date.getMonth() === month ? "" : "is-outside", key === today ? "is-today" : "", key === calendarSelectedDate ? "is-selected" : "", hasData ? "has-administrative-data" : ""].filter(Boolean).join(" ");
      const aria = hasData ? `${summary.groups.length} commesse, ${summary.plants} impianti fatti, ${formatHours(summary.hours)} ore, ${money.format(summary.earnings)}` : "nessuna attività disponibile";
      cells.push(`<button class="${classes}" type="button" role="gridcell" data-calendar-date="${key}" aria-label="${escape(formatCalendarLongDate(key))}, ${escape(aria)}"><span class="calendar-day-number">${date.getDate()}</span>${hasData ? `<span class="administrative-day-counts"><b>${summary.plants} ✓</b><b>${formatHours(summary.hours)} h</b><b>${money.format(summary.earnings)}${summary.missingAmounts ? "+" : ""}</b></span>` : ""}</button>`);
    }
    ui.calendarGrid.innerHTML = cells.join("");
    ui.calendarGrid.querySelectorAll("[data-calendar-date]").forEach((button) => {
      button.addEventListener("click", () => selectCalendarDate(button.dataset.calendarDate || ""));
    });
    if (ui.calendarFeedback) {
      const cachedCount = cachedByCommessa.size;
      const commesseCount = typeof commesseById !== "undefined" && commesseById instanceof Map ? commesseById.size : 0;
      const shared = monthlyViews.get(monthKey);
      const activities = Array.isArray(shared?.activities) ? shared.activities.length : 0;
      ui.calendarFeedback.innerHTML = `<span class="administrative-data-source">✓ Riepilogo Firestore mensile · recupero mirato solo quando incompleto</span><span>Nessuna lettura per singolo giorno, commessa o impianto · ${activities} attività aggregate · ${cachedCount}/${commesseCount || cachedCount} commesse anche in cache</span>`;
      if (shared?.recovering) ui.calendarFeedback.innerHTML += `<span>Recupero mirato delle attività del mese in corso…</span>`;
      else if (shared?.recoveryError) ui.calendarFeedback.innerHTML += `<span>⚠️ Recupero mensile non disponibile; restano visibili i dati in cache.</span>`;
    }
    renderSelectedDay(cachedByCommessa);
  }

  function plantCard(plant, fallbackOperators) {
    const location = [plant.comune, plant.address].filter(Boolean).join(" · ");
    const operator = plant.operator || (fallbackOperators.length ? `Ore inserite da: ${fallbackOperators.join(", ")}` : "Operatore non registrato");
    const quantity = plant.quantity != null ? `${plant.quantity}${plant.unit ? ` ${plant.unit}` : ""}` : "";
    return `<li class="administrative-plant"><div class="administrative-plant-head"><div><strong>${escape(plant.name)}</strong>${plant.idSap ? `<small>ID SAP ${escape(plant.idSap)}</small>` : ""}</div><span class="administrative-amount ${plant.amount == null ? "is-missing" : ""}">${plant.amount == null ? "Da valorizzare" : money.format(plant.amount)}</span></div>${location ? `<p>📍 ${escape(location)}</p>` : ""}${plant.work ? `<p>🛠️ ${escape(plant.work)}${plant.workCode && !plant.work.includes(plant.workCode) ? ` · ${escape(plant.workCode)}` : ""}${quantity ? ` · ${escape(quantity)}` : ""}</p>` : ""}<p>👤 ${escape(operator)}${plant.time ? ` · ${escape(plant.time)}` : ""}</p>${plant.note ? `<p class="administrative-note">📝 ${escape(plant.note)}</p>` : ""}</li>`;
  }

  function commessaCard(group) {
    const fallbackOperators = group.operators.map((operator) => operator.name).filter(Boolean);
    const warnings = [
      group.plants.length && !group.hoursRows.length ? "Ore non ancora inserite" : "",
      group.missingAmounts ? `${group.missingAmounts} ${group.missingAmounts === 1 ? "importo da valorizzare" : "importi da valorizzare"}` : "",
      group.missingOperators ? `${group.missingOperators} ${group.missingOperators === 1 ? "impianto senza operatore" : "impianti senza operatore"}` : ""
    ].filter(Boolean);
    return `<details class="administrative-commessa-card" open><summary><div><span>COMMESSA</span><strong>${escape(group.commessaName)}</strong>${group.commessaCode ? `<small>Cod. ${escape(group.commessaCode)}</small>` : ""}</div><div class="administrative-commessa-total"><b>${money.format(group.knownEarnings)}${group.missingAmounts ? "+" : ""}</b><small>${group.completedUnits} impianti · ${group.plants.length} attività · ${formatHours(group.totalHours)} ore</small></div></summary>${warnings.length ? `<div class="administrative-warnings">${warnings.map((warning) => `<span>⚠️ ${escape(warning)}</span>`).join("")}</div>` : ""}<section><h4>✅ Impianti e lavorazioni completate</h4>${group.plants.length ? `<ul class="administrative-plant-list">${group.plants.map((plant) => plantCard(plant, fallbackOperators)).join("")}</ul>` : `<p class="administrative-empty-inline">Nessun impianto FATTO disponibile nel riepilogo mensile o nella cache.</p>`}</section><section><h4>🕒 Ore inserite dagli operatori</h4>${group.operators.length ? `<ul class="administrative-operator-list">${group.operators.map((operator) => `<li><div><strong>${escape(operator.name)}</strong>${operator.notes.length ? `<small>${escape(operator.notes.join(" · "))}</small>` : ""}</div><b>${formatHours(operator.hours)} ore</b></li>`).join("")}</ul>` : `<p class="administrative-empty-inline">Nessuna ora inserita per questa commessa.</p>`}</section><footer><span>Totale ore commessa <b>${formatHours(group.totalHours)}</b></span><span>Guadagno ${group.missingAmounts ? "parziale" : "del giorno"} <b>${money.format(group.knownEarnings)}${group.missingAmounts ? "+" : ""}</b></span></footer></details>`;
  }

  function renderSelectedDay(cachedByCommessa = collectCachedImpianti()) {
    if (!ui?.calendarDayEvents) return;
    const groups = groupDayData(calendarSelectedDate, cachedByCommessa);
    const totals = {
      plants: groups.reduce((sum, group) => sum + group.completedUnits, 0),
      activities: groups.reduce((sum, group) => sum + group.plants.length, 0),
      hours: groups.reduce((sum, group) => sum + group.totalHours, 0),
      earnings: groups.reduce((sum, group) => sum + group.knownEarnings, 0),
      missingAmounts: groups.reduce((sum, group) => sum + group.missingAmounts, 0),
      operators: new Set(groups.flatMap((group) => group.operators
        .map((operator) => normalize(operator.name))
        .filter((name) => name && name !== "operatore non indicato"))).size
    };
    if (ui.calendarSelectedDayTitle) ui.calendarSelectedDayTitle.textContent = formatCalendarLongDate(calendarSelectedDate);
    if (ui.calendarSelectedDaySummary) ui.calendarSelectedDaySummary.textContent = groups.length ? `${groups.length} ${groups.length === 1 ? "commessa con attività" : "commesse con attività"}` : "Nessuna attività disponibile per questo giorno";
    if (!groups.length) {
      ui.calendarDayEvents.innerHTML = `<div class="calendar-empty-day"><span>📊</span><p>Nessun impianto FATTO o ora inserita risulta nel riepilogo mensile.</p><small>Il calendario non esegue letture separate per questo giorno.</small></div>`;
      return;
    }
    ui.calendarDayEvents.innerHTML = `<div class="administrative-day-kpis"><article><span>✅</span><b>${totals.plants}</b><small>Impianti fatti · ${totals.activities} attività</small></article><article><span>🕒</span><b>${formatHours(totals.hours)}</b><small>Ore totali</small></article><article><span>👷</span><b>${totals.operators}</b><small>Operatori</small></article><article><span>💶</span><b>${money.format(totals.earnings)}${totals.missingAmounts ? "+" : ""}</b><small>${totals.missingAmounts ? "Guadagno parziale" : "Guadagno del giorno"}</small></article></div>${groups.map(commessaCard).join("")}`;
  }

  function render() {
    renderCalendarGrid();
  }

  function installInteractions() {
    [
      document.getElementById("calendar-choice-administrative-btn"),
      document.getElementById("calendar-administrative-tab")
    ].forEach((button) => {
      if (!button || button.dataset.administrativeCalendarBound === "true") return;
      button.dataset.administrativeCalendarBound = "true";
      button.addEventListener("click", () => setCalendarMode("administrative"));
    });
  }

  installInteractions();

  global.HeraAdministrativeCalendar = Object.freeze({
    installed: true,
    render,
    renderSelectedDay,
    collectHoursForDate,
    collectPlantsForDate,
    groupDayData,
    ensureMonthView,
    recoverMonthActivities,
    deactivate,
    installInteractions
  });
})(window);
