(() => {
  "use strict";

  if (window.__heraSharedStaticViewsUiInstalled) return;
  window.__heraSharedStaticViewsUiInstalled = true;

  let unsubscribe = null;
  let activeType = "squadre";
  let activeKey = "";

  const OPERATIONAL_BUTTONS = {
    squadre: "open-panel-squadre",
    calendario: "open-hours-btn"
  };

  const escapeHtml = (value) => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");

  function todayKey() {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Europe/Rome",
      year: "numeric",
      month: "2-digit",
      day: "2-digit"
    }).format(new Date());
  }

  function currentMonthKey() {
    return todayKey().slice(0, 7);
  }

  function isAdmin() {
    try { return typeof canManageData === "function" && canManageData(); } catch (_) { return false; }
  }

  function ensureStyles() {
    if (document.getElementById("shared-static-views-ui-style")) return;
    const style = document.createElement("style");
    style.id = "shared-static-views-ui-style";
    style.textContent = `
      .shared-view-dialog{width:min(94vw,760px);max-height:90vh;border:0;border-radius:18px;padding:0;background:#fff;color:#171717;box-shadow:0 18px 60px rgba(0,0,0,.35)}
      .shared-view-dialog::backdrop{background:rgba(0,0,0,.62)}
      .shared-view-head{position:sticky;top:0;z-index:2;background:#fff;padding:16px;border-bottom:1px solid #ddd;display:grid;gap:10px}
      .shared-view-head-row{display:flex;align-items:center;justify-content:space-between;gap:10px}
      .shared-view-controls{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
      .shared-view-controls input{min-height:42px;padding:8px;border:1px solid #aaa;border-radius:10px}
      .shared-view-body{padding:16px;overflow:auto}
      .shared-view-body--squadre{background:#f4f7fb}
      .shared-view-meta{font-size:.85rem;color:#64748b;margin-bottom:12px}
      .shared-view-body--squadre .shared-view-meta{margin:0 0 10px;padding:9px 11px;border:1px solid #dbe5ef;border-radius:12px;background:#fff}
      .shared-view-empty{padding:24px;border:1px dashed #aaa;border-radius:12px;text-align:center;color:#666;background:#fff}
      .shared-view-card{border:1px solid #ddd;border-radius:12px;padding:12px;margin-bottom:10px;background:#fafafa}
      .shared-view-card h4{margin:0 0 8px}
      .shared-view-list{display:grid;gap:8px}
      .shared-view-json{white-space:pre-wrap;overflow-wrap:anywhere;font:13px/1.45 ui-monospace,SFMono-Regular,Menlo,monospace;background:#f4f4f4;padding:12px;border-radius:10px}
      .shared-view-feedback{min-height:20px;font-size:.9rem}
      .shared-view-primary-menu{font-weight:800}
      .shared-view-operational-menu{opacity:.86}

      .shared-squadre-board{display:grid;gap:12px;--shared-blue:#2563eb;--shared-green:#16a34a;--shared-ink:#142033;--shared-muted:#64748b;--shared-line:#dbe5ef}
      .shared-squadre-summary{display:flex;flex-wrap:wrap;gap:7px;align-items:center}
      .shared-summary-chip{display:inline-flex;align-items:center;gap:5px;min-height:30px;padding:5px 10px;border:1px solid #cbd5e1;border-radius:999px;background:#fff;color:#334155;font-size:.78rem;font-weight:800}
      .shared-commessa-card{position:relative;overflow:hidden;border:1px solid #d8e4ef;border-radius:20px;padding:14px 15px 14px 19px;background:linear-gradient(160deg,#fff 0%,#f8fbff 100%);box-shadow:0 12px 28px rgba(15,23,42,.08);color:var(--shared-ink)}
      .shared-commessa-card::before{content:"";position:absolute;inset:0 auto 0 0;width:5px;background:linear-gradient(180deg,#2563eb,#60a5fa)}
      .shared-commessa-card:nth-child(even)::before{background:linear-gradient(180deg,#16a34a,#4ade80)}
      .shared-commessa-head{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;padding-bottom:9px;border-bottom:1px solid #e7eef6}
      .shared-commessa-title-wrap{min-width:0;display:flex;flex-wrap:wrap;align-items:center;gap:7px}
      .shared-commessa-title{min-width:0;font-size:1rem;font-weight:900;line-height:1.28;overflow-wrap:anywhere}
      .shared-code-badge,.shared-count-badge,.shared-status-badge{display:inline-flex;align-items:center;justify-content:center;border-radius:999px;font-size:.72rem;font-weight:850;line-height:1.2;white-space:nowrap}
      .shared-code-badge{padding:3px 8px;border:1px solid #c7d2fe;background:#eef2ff;color:#334155}
      .shared-count-badge{flex:0 0 auto;padding:5px 9px;border:1px solid #bae6fd;background:#f0f9ff;color:#075985}
      .shared-day-row{display:flex;align-items:center;gap:7px;margin:9px 0 0;color:#475569;font-size:.85rem}
      .shared-squad-list{display:grid;gap:9px;margin-top:10px}
      .shared-squad-row{display:grid;gap:8px;padding:11px;border:1px solid #dbe5ef;border-radius:14px;background:#fff;box-shadow:0 5px 14px rgba(15,23,42,.04)}
      .shared-squad-row:nth-child(even){background:#fbfefc}
      .shared-squad-row-head{display:flex;align-items:center;justify-content:space-between;gap:8px}
      .shared-squad-name{font-weight:900;color:#1e293b}
      .shared-status-badge{padding:4px 8px;border:1px solid #bbf7d0;background:#ecfdf3;color:#166534}
      .shared-status-badge--warning{border-color:#fde68a;background:#fffbeb;color:#92400e}
      .shared-status-badge--danger{border-color:#fecaca;background:#fef2f2;color:#991b1b}
      .shared-squad-line{display:grid;grid-template-columns:minmax(118px,auto) minmax(0,1fr);align-items:start;gap:8px;font-size:.88rem;line-height:1.45}
      .shared-squad-label{font-weight:850;color:#334155}
      .shared-values{display:flex;flex-wrap:wrap;gap:5px;min-width:0}
      .shared-value-chip{display:inline-flex;align-items:center;max-width:100%;padding:3px 8px;border:1px solid #d5dfeb;border-radius:999px;background:#f8fafc;color:#334155;font-size:.8rem;font-weight:750;overflow-wrap:anywhere}
      .shared-value-chip--person{border-color:#c7d2fe;background:#eef2ff;color:#3730a3}
      .shared-value-chip--vehicle{border-color:#bbf7d0;background:#ecfdf3;color:#166534}
      .shared-value-chip--plant{border-color:#bae6fd;background:#f0f9ff;color:#075985}
      .shared-value-empty{color:#94a3b8;font-style:italic}
      .shared-squad-note{padding:8px 9px;border-left:4px solid #f59e0b;border-radius:8px;background:#fffbeb;color:#78350f;font-size:.84rem;line-height:1.4}
      .shared-squad-alert{padding:8px 9px;border-left:4px solid #dc2626;border-radius:8px;background:#fef2f2;color:#991b1b;font-size:.84rem;font-weight:750;line-height:1.4}
      .shared-squad-flags{display:flex;flex-wrap:wrap;gap:6px}
      .shared-flag{display:inline-flex;align-items:center;gap:4px;padding:4px 8px;border:1px solid #fde68a;border-radius:999px;background:#fffbeb;color:#92400e;font-size:.74rem;font-weight:850}
      .shared-flag--danger{border-color:#fecaca;background:#fef2f2;color:#991b1b}

      @media (max-width:560px){
        .shared-view-dialog{width:96vw;max-height:92dvh;border-radius:20px}
        .shared-view-head{padding:14px}
        .shared-view-body{padding:12px}
        .shared-view-controls{display:grid;grid-template-columns:1fr 1fr}
        .shared-view-controls input{grid-column:1/-1;width:100%}
        .shared-view-controls .btn{width:100%;min-width:0;padding-inline:9px}
        .shared-view-controls #shared-view-publish{grid-column:1/-1}
        .shared-commessa-card{padding:13px 12px 13px 17px;border-radius:17px}
        .shared-commessa-head{align-items:flex-start}
        .shared-count-badge{font-size:.68rem}
        .shared-squad-line{grid-template-columns:1fr;gap:3px}
        .shared-values{gap:4px}
        .shared-value-chip{font-size:.77rem}
      }
    `;
    document.head.appendChild(style);
  }

  function ensureDialog() {
    let dialog = document.getElementById("shared-static-view-dialog");
    if (dialog) return dialog;
    ensureStyles();
    dialog = document.createElement("dialog");
    dialog.id = "shared-static-view-dialog";
    dialog.className = "shared-view-dialog";
    dialog.innerHTML = `
      <div class="shared-view-head">
        <div class="shared-view-head-row">
          <div><strong id="shared-view-title">Vista condivisa</strong><div id="shared-view-feedback" class="shared-view-feedback"></div></div>
          <button id="shared-view-close" class="btn" type="button">CHIUDI</button>
        </div>
        <div class="shared-view-controls">
          <input id="shared-view-key" aria-label="Data o mese">
          <button id="shared-view-refresh" class="btn" type="button">AGGIORNA</button>
          <button id="shared-view-open-operational" class="btn" type="button">APRI MODIFICA</button>
          <button id="shared-view-publish" class="btn btn-primary" type="button">PUBBLICA VERSIONE ATTUALE</button>
        </div>
      </div>
      <div id="shared-view-body" class="shared-view-body"></div>`;
    document.body.appendChild(dialog);

    dialog.querySelector("#shared-view-close").addEventListener("click", () => dialog.close());
    dialog.querySelector("#shared-view-refresh").addEventListener("click", () => {
      const key = dialog.querySelector("#shared-view-key").value.trim();
      openView(activeType, key);
    });
    dialog.querySelector("#shared-view-open-operational").addEventListener("click", openOperationalView);
    dialog.querySelector("#shared-view-publish").addEventListener("click", publishCurrent);
    dialog.addEventListener("close", () => {
      if (typeof unsubscribe === "function") unsubscribe();
      unsubscribe = null;
    });
    return dialog;
  }

  function stringifySafe(value) {
    try { return JSON.stringify(value, null, 2); } catch (_) { return String(value ?? ""); }
  }

  function hasMeaningfulValue(value) {
    if (value == null || value === false) return false;
    if (Array.isArray(value)) return value.length > 0;
    if (typeof value === "object") return Object.keys(value).length > 0;
    return String(value).trim() !== "";
  }

  function firstMeaningfulValue(...values) {
    return values.find(hasMeaningfulValue);
  }

  function textFromObject(value, keys = []) {
    if (value == null) return "";
    if (typeof value === "string" || typeof value === "number") return String(value).trim();
    if (typeof value !== "object") return "";
    const preferred = [...keys, "nomeCompleto", "displayName", "nominativo", "denominazione", "nome", "name", "label", "valore", "operatore", "codice", "targa", "id"];
    for (const key of preferred) {
      if (!hasMeaningfulValue(value[key])) continue;
      const nested = textFromObject(value[key]);
      if (nested) return nested;
    }
    return "";
  }

  function splitTextValues(value) {
    const text = String(value ?? "").trim();
    if (!text) return [];
    return text
      .split(/\s*(?:,|;|\n|\||•)\s*/)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  function toDisplayList(value, keys = []) {
    const values = [];
    const append = (entry) => {
      if (!hasMeaningfulValue(entry)) return;
      if (Array.isArray(entry)) {
        entry.forEach(append);
        return;
      }
      if (typeof entry === "string" || typeof entry === "number") {
        splitTextValues(entry).forEach((item) => values.push(item));
        return;
      }
      if (typeof entry !== "object") return;
      const label = textFromObject(entry, keys);
      if (label) {
        values.push(label);
        return;
      }
      Object.entries(entry).forEach(([key, nested]) => {
        if (nested === true) values.push(key);
      });
    };
    append(value);
    const seen = new Set();
    return values.filter((item) => {
      const clean = String(item || "").trim();
      if (!clean || clean === "-") return false;
      const key = clean.toLocaleLowerCase("it-IT");
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function normalizeDateValue(value) {
    if (!value) return "";
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
    if (typeof value === "object") {
      if (typeof value.toDate === "function") return normalizeDateValue(value.toDate());
      if (Number.isFinite(Number(value.seconds))) return normalizeDateValue(new Date(Number(value.seconds) * 1000));
    }
    const text = String(value).trim();
    const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return match ? `${match[1]}-${match[2]}-${match[3]}` : "";
  }

  function formatDateForDisplay(value) {
    const key = normalizeDateValue(value);
    if (!key) return "-";
    const date = new Date(`${key}T12:00:00`);
    return Number.isNaN(date.getTime()) ? key : date.toLocaleDateString("it-IT", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    });
  }

  function normalizeTime(value) {
    if (!hasMeaningfulValue(value)) return "";
    const text = String(value).trim();
    const match = text.match(/(?:^|T|\s)(\d{1,2}):(\d{2})/);
    if (!match) return text;
    return `${String(match[1]).padStart(2, "0")}:${match[2]}`;
  }

  function isTruthy(value) {
    if (value === true || value === 1) return true;
    return /^(true|1|si|sì|yes)$/i.test(String(value || "").trim());
  }

  function looksLikeSquadraRow(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) return false;
    return [
      "personale", "operatori", "operators", "mezzi", "mezzo", "caposquadra",
      "impianti", "impianto", "oraInizio", "oraFine", "note", "nome", "name",
      "stato", "senzaPausaPranzo", "conflittiConfermati"
    ].some((key) => hasMeaningfulValue(value[key]));
  }

  function flattenSquadrePayload(payload) {
    const rows = [];
    const visit = (value, inherited = {}) => {
      if (Array.isArray(value)) {
        value.forEach((item) => visit(item, inherited));
        return;
      }
      if (!value || typeof value !== "object") return;

      const next = {
        date: firstMeaningfulValue(value.date, value.data, value.dateKey, value.riferimentoData, inherited.date, activeKey),
        commessaId: firstMeaningfulValue(value.commessaId, value.commessaID, value.idCommessa, inherited.commessaId),
        commessa: firstMeaningfulValue(value.commessa, value.commessaNome, value.nomeCommessa, inherited.commessa),
        codiceCommessa: firstMeaningfulValue(value.codiceCommessa, value.commessaCodice, inherited.codiceCommessa)
      };

      const nestedRows = Array.isArray(value.squadre)
        ? value.squadre
        : (Array.isArray(value.items) ? value.items : (Array.isArray(value.rows) ? value.rows : null));
      if (nestedRows) {
        nestedRows.forEach((item) => visit(item, next));
        return;
      }

      if (looksLikeSquadraRow(value)) rows.push({ row: value, context: next });
    };

    visit(payload, { date: activeKey });
    return rows;
  }

  function findKnownCommessa(commessaId) {
    const id = String(commessaId || "").trim();
    if (!id) return null;
    try {
      if (typeof commesseById !== "undefined" && commesseById instanceof Map) {
        const record = commesseById.get(id);
        if (record) return record;
      }
    } catch (_) {}
    try {
      if (window.commesseById instanceof Map) return window.commesseById.get(id) || null;
    } catch (_) {}
    return null;
  }

  function resolveCommessa(entry) {
    const row = entry.row || {};
    const context = entry.context || {};
    const rawId = firstMeaningfulValue(row.commessaId, row.commessaID, row.idCommessa, context.commessaId);
    const id = textFromObject(rawId);
    const known = findKnownCommessa(id);
    const rawName = firstMeaningfulValue(row.commessa, row.commessaNome, row.nomeCommessa, context.commessa);
    let name = textFromObject(rawName, ["nome", "name", "denominazione"])
      || textFromObject(known, ["nome", "name", "denominazione"]);
    if (!name && id) name = `Commessa ${id}`;
    if (!name) name = "Commessa non indicata";
    const code = textFromObject(firstMeaningfulValue(
      row.codiceCommessa,
      row.commessaCodice,
      context.codiceCommessa,
      known?.codice,
      known?.code
    ));
    return { id, name, code, known };
  }

  function resolvePersonnel(row) {
    const source = firstMeaningfulValue(
      row.personale,
      row.operatori,
      row.operators,
      row.membri,
      row.componenti,
      row.conflittiConfermati?.operatori
    );
    return toDisplayList(source, ["nomeCompleto", "displayName", "nominativo", "nome", "name", "valore", "operatore"]);
  }

  function resolveVehicles(row) {
    const source = firstMeaningfulValue(row.mezzi, row.mezzo, row.veicoli, row.vehicles);
    return toDisplayList(source, ["codice", "targa", "nome", "name", "valore", "mezzo"]);
  }

  function resolvePlants(row) {
    const source = firstMeaningfulValue(row.impianti, row.impianto, row.impiantiNomi, row.impiantiDettagli);
    return toDisplayList(source, ["denominazione", "nome", "name", "label", "impianto", "idSap", "idSAP"]);
  }

  function resolveTeamName(row, index, commessaName) {
    const candidate = textFromObject(firstMeaningfulValue(row.nomeSquadra, row.squadraNome, row.teamName, row.nome, row.name));
    if (candidate && candidate.toLocaleLowerCase("it-IT") !== String(commessaName || "").toLocaleLowerCase("it-IT")) return candidate;
    return `Squadra ${index + 1}`;
  }

  function resolveSchedule(row) {
    const direct = textFromObject(firstMeaningfulValue(row.orario, row.fasciaOraria));
    if (direct) return direct;
    const start = normalizeTime(firstMeaningfulValue(row.oraInizio, row.orarioInizio, row.inizio, row.startTime));
    const end = normalizeTime(firstMeaningfulValue(row.oraFine, row.orarioFine, row.fine, row.endTime));
    if (start && end) return `${start} – ${end}`;
    return start || end || "";
  }

  function resolveEntryDate(entry) {
    const row = entry.row || {};
    const context = entry.context || {};
    return firstMeaningfulValue(row.riferimentoData, row.dateKey, row.date, row.data, row.giorno, context.date, activeKey);
  }

  function renderValues(values, variant = "") {
    if (!values.length) return `<span class="shared-value-empty">Non indicati</span>`;
    const suffix = variant ? ` shared-value-chip--${variant}` : "";
    return values.map((value) => `<span class="shared-value-chip${suffix}">${escapeHtml(value)}</span>`).join("");
  }

  function renderSquadraRow(entry, index, commessaName) {
    const row = entry.row || {};
    const personnel = resolvePersonnel(row);
    const vehicles = resolveVehicles(row);
    const plants = resolvePlants(row);
    const caposquadra = textFromObject(firstMeaningfulValue(row.caposquadra, row.capoSquadra, row.teamLeader), ["nomeCompleto", "displayName", "nome", "name"]);
    const schedule = resolveSchedule(row);
    const note = textFromObject(firstMeaningfulValue(row.note, row.notes, row.nota));
    const absenceWarning = textFromObject(firstMeaningfulValue(row.avvisoAutomaticoAssenze, row.avvisoAssenze));
    const status = textFromObject(row.stato);
    const withoutLunch = isTruthy(row.senzaPausaPranzo);
    const confirmedConflict = isTruthy(row.conflittoMezzoConfermato)
      || isTruthy(row.conflittoOperatoreConfermato)
      || isTruthy(row.conflittoConfermato);
    const teamName = resolveTeamName(row, index, commessaName);

    return `
      <section class="shared-squad-row">
        <div class="shared-squad-row-head">
          <span class="shared-squad-name">👥 ${escapeHtml(teamName)}</span>
          ${status ? `<span class="shared-status-badge">${escapeHtml(status)}</span>` : ""}
        </div>
        <div class="shared-squad-line">
          <span class="shared-squad-label">👷 Operatori:</span>
          <span class="shared-values">${renderValues(personnel, "person")}</span>
        </div>
        ${caposquadra ? `<div class="shared-squad-line"><span class="shared-squad-label">🧑‍✈️ Caposquadra:</span><span>${escapeHtml(caposquadra)}</span></div>` : ""}
        ${schedule ? `<div class="shared-squad-line"><span class="shared-squad-label">🕒 Orario:</span><span>${escapeHtml(schedule)}</span></div>` : ""}
        ${plants.length ? `<div class="shared-squad-line"><span class="shared-squad-label">📍 Impianti:</span><span class="shared-values">${renderValues(plants, "plant")}</span></div>` : ""}
        <div class="shared-squad-line">
          <span class="shared-squad-label">🚚 Mezzi ${index + 1}:</span>
          <span class="shared-values">${renderValues(vehicles, "vehicle")}</span>
        </div>
        ${(withoutLunch || confirmedConflict) ? `<div class="shared-squad-flags">${withoutLunch ? `<span class="shared-flag">🍽️ Senza pausa pranzo</span>` : ""}${confirmedConflict ? `<span class="shared-flag shared-flag--danger">⚠️ Conflitto confermato</span>` : ""}</div>` : ""}
        ${note ? `<div class="shared-squad-note"><b>📝 Note:</b> ${escapeHtml(note)}</div>` : ""}
        ${absenceWarning ? `<div class="shared-squad-alert"><b>⛔ Assenza calendario:</b> ${escapeHtml(absenceWarning)}</div>` : ""}
      </section>`;
  }

  function renderSquadreBoard(payload) {
    const entries = flattenSquadrePayload(payload);
    if (!entries.length) {
      return `<div class="shared-view-empty">Nessuna squadra presente nella versione pubblicata per questa data.</div>`;
    }

    const groups = new Map();
    entries.forEach((entry, originalIndex) => {
      const commessa = resolveCommessa(entry);
      const key = commessa.id || commessa.name.toLocaleLowerCase("it-IT") || `commessa-${originalIndex}`;
      if (!groups.has(key)) groups.set(key, { commessa, entries: [] });
      groups.get(key).entries.push(entry);
    });

    const groupList = Array.from(groups.values());
    const date = firstMeaningfulValue(payload?.date, payload?.dateKey, entries[0] && resolveEntryDate(entries[0]), activeKey);
    const summary = `
      <div class="shared-squadre-summary">
        <span class="shared-summary-chip">📅 ${escapeHtml(formatDateForDisplay(date))}</span>
        <span class="shared-summary-chip">📁 ${groupList.length} ${groupList.length === 1 ? "commessa" : "commesse"}</span>
        <span class="shared-summary-chip">👥 ${entries.length} ${entries.length === 1 ? "squadra" : "squadre"}</span>
      </div>`;

    const cards = groupList.map(({ commessa, entries: groupEntries }) => {
      const groupDate = firstMeaningfulValue(resolveEntryDate(groupEntries[0]), date, activeKey);
      return `
        <article class="shared-commessa-card">
          <div class="shared-commessa-head">
            <div class="shared-commessa-title-wrap">
              <span class="shared-commessa-title">📁 ${escapeHtml(commessa.name)}</span>
              ${commessa.code ? `<span class="shared-code-badge">${escapeHtml(commessa.code)}</span>` : ""}
            </div>
            <span class="shared-count-badge">${groupEntries.length} ${groupEntries.length === 1 ? "squadra" : "squadre"}</span>
          </div>
          <p class="shared-day-row"><b>📅 Giorno:</b> ${escapeHtml(formatDateForDisplay(groupDate))}</p>
          <div class="shared-squad-list">${groupEntries.map((entry, index) => renderSquadraRow(entry, index, commessa.name)).join("")}</div>
        </article>`;
    }).join("");

    return `<div class="shared-squadre-board">${summary}${cards}</div>`;
  }

  function renderPayload(documentValue) {
    const body = document.getElementById("shared-view-body");
    if (!body) return;
    body.classList.toggle("shared-view-body--squadre", activeType === "squadre");
    if (!documentValue?.payload) {
      body.innerHTML = `<div class="shared-view-empty">Nessuna vista condivisa pubblicata per questo periodo.<br>Un amministratore può aprire la modalità modifica e poi pubblicare la versione aggiornata.</div>`;
      return;
    }
    const updated = documentValue.updatedAt?.toDate?.() || documentValue.updatedAt || documentValue.updatedAtClient || "";
    const updatedDate = updated instanceof Date ? updated : (updated ? new Date(updated) : null);
    const updatedText = updatedDate && !Number.isNaN(updatedDate.getTime()) ? updatedDate.toLocaleString("it-IT") : String(updated || "");
    const payload = documentValue.payload;
    let content = "";

    if (activeType === "squadre") {
      content = renderSquadreBoard(payload);
    } else {
      const items = Array.isArray(payload) ? payload : (Array.isArray(payload?.items) ? payload.items : null);
      if (items) {
        content = `<div class="shared-view-list">${items.map((item, index) => `
          <article class="shared-view-card"><h4>${escapeHtml(item?.nome || item?.name || item?.commessa || item?.operatore || `Elemento ${index + 1}`)}</h4>
          <pre class="shared-view-json">${escapeHtml(stringifySafe(item))}</pre></article>`).join("")}</div>`;
      } else {
        content = `<pre class="shared-view-json">${escapeHtml(stringifySafe(payload))}</pre>`;
      }
    }

    body.innerHTML = `<div class="shared-view-meta">Versione ${escapeHtml(documentValue.version || 1)}${updatedText ? ` · aggiornata ${escapeHtml(updatedText)}` : ""}${documentValue.authorName ? ` · ${escapeHtml(documentValue.authorName)}` : ""}</div>${content}`;
  }

  function setFeedback(message, error = false) {
    const node = document.getElementById("shared-view-feedback");
    if (!node) return;
    node.textContent = message || "";
    node.style.color = error ? "#b00020" : "#146c2e";
  }

  function openView(type, key) {
    const api = window.HeraSharedStaticViews;
    const dialog = ensureDialog();
    activeType = type;
    activeKey = key || (type === "squadre" ? todayKey() : currentMonthKey());
    const input = dialog.querySelector("#shared-view-key");
    input.type = type === "squadre" ? "date" : "month";
    input.value = activeKey;
    dialog.querySelector("#shared-view-title").textContent = type === "squadre" ? "Composizione squadre" : "Calendario personale";
    dialog.querySelector("#shared-view-open-operational").textContent = type === "squadre" ? "MODIFICA SQUADRE" : "MODIFICA ORE";
    dialog.querySelector("#shared-view-publish").hidden = !isAdmin();
    setFeedback("");
    renderPayload(api?.getCached?.(type, activeKey));
    if (typeof unsubscribe === "function") unsubscribe();
    unsubscribe = api?.subscribe?.(type, activeKey, renderPayload) || null;
    if (!dialog.open) dialog.showModal();
  }

  function openOperationalView() {
    const dialog = ensureDialog();
    const buttonId = OPERATIONAL_BUTTONS[activeType];
    const operationalButton = buttonId ? document.getElementById(buttonId) : null;
    if (!operationalButton) {
      setFeedback("Schermata di modifica non disponibile.", true);
      return;
    }
    dialog.close();
    operationalButton.click();
  }

  async function publishCurrent() {
    if (!isAdmin()) return setFeedback("Solo un amministratore può pubblicare.", true);
    const api = window.HeraSharedStaticViews;
    const key = document.getElementById("shared-view-key")?.value?.trim() || activeKey;
    setFeedback("Pubblicazione in corso…");
    try {
      const result = activeType === "squadre"
        ? await api.publishSquadre(key)
        : await api.publishCalendar(key);
      setFeedback(result?.skipped ? "La versione condivisa era già aggiornata." : "Vista condivisa pubblicata su tutti i dispositivi.");
    } catch (error) {
      setFeedback(error?.message || "Pubblicazione non riuscita.", true);
    }
  }

  function renameOperationalButton(anchor, label) {
    if (!anchor || anchor.dataset.heraSharedOperationalRenamed === "true") return;
    anchor.dataset.heraSharedOperationalRenamed = "true";
    anchor.classList.add("shared-view-operational-menu");
    anchor.setAttribute("aria-label", label);
    anchor.setAttribute("title", label);
    anchor.innerHTML = `<span class="menu-item-icon" aria-hidden="true">✏️</span>${label}`;
  }

  function addPrimaryMenuButton(anchorId, id, label, type, operationalLabel) {
    const existing = document.getElementById(id);
    const anchor = document.getElementById(anchorId);
    if (!anchor?.parentNode) return false;
    renameOperationalButton(anchor, operationalLabel);
    if (existing) return true;
    const button = document.createElement("button");
    button.id = id;
    button.type = "button";
    button.className = "btn menu-title-btn shared-view-primary-menu";
    button.innerHTML = `<span class="menu-item-icon" aria-hidden="true">🖼️</span>${label}`;
    button.addEventListener("click", () => openView(type));
    anchor.insertAdjacentElement("beforebegin", button);
    return true;
  }

  function install() {
    const a = addPrimaryMenuButton("open-panel-squadre", "open-shared-squadre-view", "COMPOSIZIONE SQUADRE", "squadre", "MODIFICA SQUADRE");
    const b = addPrimaryMenuButton("open-hours-btn", "open-shared-calendar-view", "CALENDARIO PERSONALE", "calendario", "MODIFICA ORE / CALENDARIO");
    return a && b;
  }

  let attempts = 0;
  const timer = setInterval(() => {
    attempts += 1;
    if (install() || attempts >= 80) clearInterval(timer);
  }, 250);

  window.HeraSharedStaticViewsUi = {
    version: "2.0.0",
    openSquadre: (date) => openView("squadre", date),
    openCalendar: (month) => openView("calendario", month),
    openOperational: openOperationalView
  };
})();
