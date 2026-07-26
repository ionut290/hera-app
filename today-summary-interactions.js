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
})();
