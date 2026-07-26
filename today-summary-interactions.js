"use strict";

(() => {
  const getAssignments = () => getCurrentUserAssignedCommesseForDate(getTodayDateKey());

  function getPlannedHours(assignments = getAssignments()) {
    const uniqueRows = new Set();
    let total = 0;
    assignments.forEach((assignment) => {
      (assignment.matchedRows || []).forEach(({ squadraIndex, row }) => {
        const rowKey = `${assignment.commessaId}:${squadraIndex}`;
        if (uniqueRows.has(rowKey)) return;
        uniqueRows.add(rowKey);
        total += getSquadraWorkedHours(row);
      });
    });
    return total;
  }

  function getAlertGroups(assignments = getAssignments()) {
    return assignments.map((assignment) => {
      const matchedRows = assignment.matchedRows || [];
      const assignedAlerts = matchedRows
        .map(({ squadraLabel, row }) => String(row?.avviso || "").trim()
          ? `⚠️ ${squadraLabel}: ${String(row.avviso).trim()}`
          : "")
        .filter(Boolean);
      const safetyAlerts = buildSquadraWarningDetails(
        assignment.commessa,
        matchedRows.map((item) => item.row)
      );
      return { assignment, issues: [...assignedAlerts, ...safetyAlerts] };
    }).filter((group) => group.issues.length);
  }

  function openChoice({ title, description, assignments, actionLabel, onSelect }) {
    if (!assignments.length) {
      alert("Oggi non risulti assegnato a nessuna commessa.");
      return;
    }
    if (assignments.length === 1) {
      onSelect(assignments[0]);
      return;
    }
    const overlay = document.createElement("div");
    overlay.className = "confirm-modal";
    overlay.innerHTML = `
      <div class="confirm-modal-card" role="dialog" aria-modal="true" aria-label="${escapeHTML(title)}">
        <h2>${escapeHTML(title)}</h2>
        <p>${escapeHTML(description)}</p>
        <div class="confirm-modal-actions">
          ${assignments.map((assignment, index) => `
            <button type="button" class="btn btn-primary" data-today-choice="${index}">
              ${escapeHTML(actionLabel)}: ${escapeHTML(assignment.commessaName || "Commessa")}
            </button>`).join("")}
          <button type="button" class="btn" data-today-choice-close>Annulla</button>
        </div>
      </div>`;
    const close = () => overlay.remove();
    overlay.querySelector("[data-today-choice-close]")?.addEventListener("click", close);
    overlay.addEventListener("click", (event) => { if (event.target === overlay) close(); });
    overlay.querySelectorAll("[data-today-choice]").forEach((button) => {
      button.addEventListener("click", () => {
        const assignment = assignments[Number(button.dataset.todayChoice)];
        close();
        if (assignment) onSelect(assignment);
      });
    });
    document.body.appendChild(overlay);
  }

  function openAssignedCommessa() {
    openChoice({
      title: "Commesse assegnate oggi",
      description: "Scegli la commessa da aprire.",
      assignments: getAssignments(),
      actionLabel: "Apri",
      onSelect: (assignment) => openCommessaFromSquadre(assignment.commessa)
    });
  }

  function openAssignedHours() {
    openChoice({
      title: "Inserisci ore di oggi",
      description: "Scegli la commessa: potrai inserire le ore soltanto per la squadra di cui fai parte.",
      assignments: getAssignments(),
      actionLabel: "Inserisci ore",
      onSelect: (assignment) => openHoursPageForCommessa(assignment.commessaId, getTodayDateKey())
    });
  }

  function openAssignedVehicles() {
    const assignments = getAssignments();
    if (!assignments.length) {
      alert("Oggi non risulti assegnato a nessuna commessa.");
      return;
    }
    const groups = assignments.map((assignment) => ({
      assignment,
      rows: (assignment.matchedRows || []).map(({ squadraLabel, row }) => ({
        squadraLabel,
        vehicles: parseMultiEntryValue(row?.mezzi || "")
      })).filter((item) => item.vehicles.length)
    })).filter((group) => group.rows.length);
    if (!groups.length) {
      alert("Oggi non risultano mezzi assegnati alla tua squadra.");
      return;
    }
    const overlay = document.createElement("div");
    overlay.className = "confirm-modal";
    overlay.innerHTML = `
      <div class="confirm-modal-card" role="dialog" aria-modal="true" aria-label="Mezzi assegnati oggi">
        <h2>🚚 Mezzi assegnati oggi</h2>
        ${groups.map(({ assignment, rows }) => `
          <section>
            <h3>${escapeHTML(assignment.commessaName || "Commessa")}</h3>
            ${rows.map(({ squadraLabel, vehicles }) => `
              <p><b>${escapeHTML(squadraLabel)}:</b> ${vehicles.map((vehicle) => escapeHTML(vehicle)).join(", ")}</p>
            `).join("")}
          </section>`).join("")}
        <div class="confirm-modal-actions">
          <button type="button" class="btn btn-primary" data-today-vehicles-close>Chiudi</button>
        </div>
      </div>`;
    const close = () => overlay.remove();
    overlay.querySelector("[data-today-vehicles-close]")?.addEventListener("click", close);
    overlay.addEventListener("click", (event) => { if (event.target === overlay) close(); });
    document.body.appendChild(overlay);
  }

  function openAlerts() {
    const groups = getAlertGroups();
    if (!groups.length) {
      alert("✅ Nessun avviso per le squadre a cui sei assegnato oggi.");
      return;
    }
    const overlay = document.createElement("div");
    overlay.className = "confirm-modal";
    overlay.innerHTML = `
      <div class="confirm-modal-card" role="dialog" aria-modal="true" aria-label="Avvisi di oggi">
        <h2>⚠️ Avvisi di oggi</h2>
        ${groups.map(({ assignment, issues }) => `
          <section>
            <h3>${escapeHTML(assignment.commessaName || "Commessa")}</h3>
            <ul>${issues.map((issue) => `<li>${escapeHTML(issue.replace(/^⚠️\s*/, ""))}</li>`).join("")}</ul>
          </section>`).join("")}
        <div class="confirm-modal-actions">
          <button type="button" class="btn btn-primary" data-today-alerts-close>Chiudi</button>
        </div>
      </div>`;
    const close = () => overlay.remove();
    overlay.querySelector("[data-today-alerts-close]")?.addEventListener("click", close);
    overlay.addEventListener("click", (event) => { if (event.target === overlay) close(); });
    document.body.appendChild(overlay);
  }

  function replaceSummaryButton(key, handler) {
    const current = ui[key];
    if (!current?.parentNode) return;
    const replacement = current.cloneNode(true);
    current.parentNode.replaceChild(replacement, current);
    ui[key] = replacement;
    replacement.addEventListener("click", handler);
  }

  replaceSummaryButton("todayCommesseBtn", openAssignedCommessa);
  replaceSummaryButton("todayHoursBtn", openAssignedHours);
  replaceSummaryButton("todayMezziBtn", openAssignedVehicles);
  replaceSummaryButton("todayAlertsBtn", openAlerts);

  renderTodaySummary = function renderInteractiveTodaySummary() {
    if (!ui.todayCommesseCount) return;
    const dateKey = getTodayDateKey();
    const assignments = getCurrentUserAssignedCommesseForDate(dateKey);
    const mezzi = new Set();
    const uniqueRows = new Set();

    assignments.forEach((assignment) => {
      (assignment.matchedRows || []).forEach(({ squadraIndex, row }) => {
        const rowKey = `${assignment.commessaId}:${squadraIndex}`;
        if (uniqueRows.has(rowKey)) return;
        uniqueRows.add(rowKey);
        parseMultiEntryValue(row?.mezzi || "").forEach((mezzo) => {
          const key = normalizeSquadraMemberIdentity(mezzo);
          if (key) mezzi.add(key);
        });
      });
    });

    const alerts = getAlertGroups(assignments).reduce((sum, group) => sum + group.issues.length, 0);
    ui.todayCommesseCount.textContent = String(assignments.length);
    ui.todayHoursCount.textContent = formatSquadraHours(getPlannedHours(assignments)) || "0";
    ui.todayMezziCount.textContent = String(mezzi.size);
    ui.todayAlertsCount.textContent = String(alerts);
    ui.todayAlertsBtn?.classList.toggle("has-alerts", alerts > 0);
  };

  // Selettore mezzi stabile su browser mobile e WebView Android.
  // Sostituisce il popup nativo del datalist, che può finire dietro la tastiera
  // o fuori dallo schermo, con una finestra scorrevole e ricercabile.
  function installVehiclePickerFix() {
    if (document.getElementById("vehicle-picker-fix-style")) return;
    const style = document.createElement("style");
    style.id = "vehicle-picker-fix-style";
    style.textContent = `
      .vehicle-picker-overlay{position:fixed;inset:0;z-index:5000;background:rgba(15,23,42,.55);display:flex;align-items:flex-end;justify-content:center;padding:12px}
      .vehicle-picker-card{width:min(560px,100%);max-height:min(78dvh,680px);background:#fff;border-radius:20px 20px 14px 14px;box-shadow:0 24px 60px rgba(0,0,0,.28);display:grid;grid-template-rows:auto auto minmax(0,1fr) auto;overflow:hidden}
      .vehicle-picker-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:16px;border-bottom:1px solid #d6dfec}
      .vehicle-picker-head h2{margin:0;font-size:1.15rem}
      .vehicle-picker-search{margin:12px 16px 8px;width:calc(100% - 32px);min-height:48px;font-size:16px;border:1px solid #b8c6da;border-radius:12px;padding:10px 12px}
      .vehicle-picker-list{overflow:auto;overscroll-behavior:contain;padding:6px 12px 12px;display:grid;gap:8px}
      .vehicle-picker-option{width:100%;min-height:50px;text-align:left;border:1px solid #d6dfec;border-radius:12px;background:#fff;padding:12px 14px;font-size:1rem;color:#172033}
      .vehicle-picker-option:active{background:#eef4ff}
      .vehicle-picker-empty{padding:18px;text-align:center;color:#60708a}
      .vehicle-picker-actions{padding:12px 16px calc(12px + env(safe-area-inset-bottom));border-top:1px solid #d6dfec;background:#fff}
      .vehicle-picker-actions button{width:100%;min-height:48px}
      @media (min-width:700px){.vehicle-picker-overlay{align-items:center}.vehicle-picker-card{border-radius:18px}}
    `;
    document.head.appendChild(style);

    const isVehicleDatalistInput = (input) => {
      if (!(input instanceof HTMLInputElement) || !input.list) return false;
      const text = `${input.placeholder || ""} ${input.getAttribute("aria-label") || ""} ${input.closest("label")?.textContent || ""} ${input.parentElement?.parentElement?.textContent || ""}`.toLowerCase();
      return /mezzo|mezzi/.test(text);
    };

    const openPicker = (input) => {
      const options = Array.from(input.list?.options || [])
        .map((option) => String(option.value || option.textContent || "").trim())
        .filter(Boolean);
      if (!options.length) return;

      input.blur();
      const overlay = document.createElement("div");
      overlay.className = "vehicle-picker-overlay";
      overlay.innerHTML = `
        <section class="vehicle-picker-card" role="dialog" aria-modal="true" aria-label="Seleziona mezzo squadra">
          <header class="vehicle-picker-head"><h2>🚚 Seleziona mezzo</h2><button type="button" class="btn" data-vehicle-close>Chiudi</button></header>
          <input class="vehicle-picker-search" type="search" inputmode="search" autocomplete="off" placeholder="Cerca targa o mezzo…" aria-label="Cerca mezzo">
          <div class="vehicle-picker-list"></div>
          <div class="vehicle-picker-actions"><button type="button" class="btn" data-vehicle-clear>Nessun mezzo</button></div>
        </section>`;
      const list = overlay.querySelector(".vehicle-picker-list");
      const search = overlay.querySelector(".vehicle-picker-search");
      const close = () => overlay.remove();
      const render = () => {
        const query = String(search.value || "").trim().toLowerCase();
        const filtered = options.filter((value) => value.toLowerCase().includes(query));
        list.innerHTML = filtered.length
          ? filtered.map((value, index) => `<button type="button" class="vehicle-picker-option" data-vehicle-index="${index}">${escapeHTML(value)}</button>`).join("")
          : '<p class="vehicle-picker-empty">Nessun mezzo trovato.</p>';
        list.querySelectorAll("[data-vehicle-index]").forEach((button) => {
          button.addEventListener("click", () => {
            const value = filtered[Number(button.dataset.vehicleIndex)];
            input.value = value || "";
            input.dispatchEvent(new Event("input", { bubbles: true }));
            input.dispatchEvent(new Event("change", { bubbles: true }));
            close();
          });
        });
      };
      overlay.querySelector("[data-vehicle-close]").addEventListener("click", close);
      overlay.querySelector("[data-vehicle-clear]").addEventListener("click", () => {
        input.value = "";
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));
        close();
      });
      overlay.addEventListener("click", (event) => { if (event.target === overlay) close(); });
      search.addEventListener("input", render);
      document.body.appendChild(overlay);
      render();
      setTimeout(() => search.focus({ preventScroll: true }), 80);
    };

    document.addEventListener("pointerdown", (event) => {
      const input = event.target.closest?.("input[list]");
      if (!isVehicleDatalistInput(input)) return;
      event.preventDefault();
      openPicker(input);
    }, true);
  }

  installVehiclePickerFix();
})();
