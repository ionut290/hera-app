"use strict";
(function installVargaCalendarModule(global) {
  if (global.VargaCalendarModule) return;
  const api = {};
  function formatCalendarDateKey(date) {
    const value = date instanceof Date ? date : new Date(date);
    if (Number.isNaN(value.getTime())) return "";
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  api.formatCalendarDateKey = formatCalendarDateKey;
  function parseCalendarDateKey(dateKey) {
    const match = String(dateKey || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return null;
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  api.parseCalendarDateKey = parseCalendarDateKey;
  function formatCalendarLongDate(dateKey) {
    const date = parseCalendarDateKey(dateKey);
    return date
      ? new Intl.DateTimeFormat("it-IT", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(date)
      : "Data non disponibile";
  }
  api.formatCalendarLongDate = formatCalendarLongDate;
  function openCalendarPage() {
    if (!currentUser) {
      alert("Devi fare login per aprire il calendario.");
      return;
    }
    calendarMode = "choice";
    window.location.hash = "calendario";
    applyRoute();
    closeSideMenu();
  }
  api.openCalendarPage = openCalendarPage;
  function closeCalendarPage() {
    closeCalendarEventForm();
    calendarMode = "choice";
    window.location.hash = "";
    applyRoute();
  }
  api.closeCalendarPage = closeCalendarPage;
  function setCalendarMode(mode) {
    if (mode !== "hours" && mode !== "shared") return;
    calendarMode = mode;
    renderCalendarMode();
    renderCalendar();
  }
  api.setCalendarMode = setCalendarMode;
  function renderCalendarMode() {
    const isChoice = calendarMode === "choice";
    const isHours = calendarMode === "hours";
    ui.calendarChoiceCard?.classList.toggle("hidden", !isChoice);
    ui.calendarHeroCard?.classList.toggle("hidden", isChoice);
    ui.calendarMainCard?.classList.toggle("hidden", isChoice);
    ui.calendarDayCard?.classList.toggle("hidden", isChoice);
    ui.calendarNewEventBtn?.classList.toggle("hidden", isHours);
    ui.calendarAddSelectedDayBtn?.classList.toggle("hidden", isHours);
    ui.calendarHoursTab?.classList.toggle("is-active", isHours);
    ui.calendarSharedTab?.classList.toggle("is-active", calendarMode === "shared");
    ui.calendarHoursTab?.setAttribute("aria-selected", String(isHours));
    ui.calendarSharedTab?.setAttribute("aria-selected", String(calendarMode === "shared"));
    if (ui.calendarPageHeading) ui.calendarPageHeading.textContent = isHours ? "🕒 Le mie ore" : "🗓️ Calendario condiviso";
    if (ui.calendarPageDescription) {
      ui.calendarPageDescription.textContent = isHours
        ? "Ore lavorate personali recuperate dalla Gestione ore."
        : "Ferie, permessi, malattie, interventi e altre informazioni visibili a tutti gli utenti.";
    }
    if (ui.calendarGrid) ui.calendarGrid.setAttribute("aria-label", isHours ? "Calendario mensile delle mie ore" : "Calendario mensile condiviso");
  }
  api.renderCalendarMode = renderCalendarMode;
  function subscribeCalendarEvents() {
    if (!currentUser || !db || unsubscribeCalendarEvents) return;
    if (ui.calendarFeedback) ui.calendarFeedback.textContent = "Caricamento eventi...";
    const visibleEvents = new Map();
    const subscriptions = [];
    const applySnapshot = (snapshot) => {
      snapshot.docChanges().forEach((change) => {
        if (change.type === "removed") visibleEvents.delete(change.doc.id);
        else visibleEvents.set(change.doc.id, { id: change.doc.id, ...change.doc.data() });
      });
      calendarEvents = Array.from(visibleEvents.values());
      calendarAbsenceCache.clear();
      if (ui.calendarFeedback) {
        ui.calendarFeedback.textContent = calendarEvents.length
          ? `${calendarEvents.length} ${calendarEvents.length === 1 ? "evento condiviso" : "eventi condivisi"}`
          : "Nessun evento inserito.";
      }
      renderCalendar();
    };
    const handleError = (error) => {
      console.error("Errore caricamento calendario condiviso:", error);
      if (ui.calendarFeedback) ui.calendarFeedback.textContent = "Impossibile caricare gli eventi. Verifica la connessione e i permessi.";
    };
    // Keep private document expirations out of the broad shared-calendar query.
    // Firestore then validates each privacy-scoped query against the same rules used
    // for direct reads, so an administrator cannot enumerate personal expirations.
    const events = db.collection("calendarEvents");
    subscriptions.push(events.where("type", "!=", "SCADENZA_DOCUMENTO").onSnapshot(applySnapshot, handleError));
    subscriptions.push(events.where("ownerUserId", "==", currentUser.uid).onSnapshot(applySnapshot, handleError));
    subscriptions.push(events.where("authorizedUserIds", "array-contains", currentUser.uid).onSnapshot(applySnapshot, handleError));
    subscriptions.push(events.where("sharedToAll", "==", true).onSnapshot(applySnapshot, handleError));
    unsubscribeCalendarEvents = () => subscriptions.forEach((unsubscribe) => unsubscribe());
  }
  api.subscribeCalendarEvents = subscribeCalendarEvents;
  function stopCalendarEventsSubscription() {
    if (unsubscribeCalendarEvents) {
      unsubscribeCalendarEvents();
      unsubscribeCalendarEvents = null;
    }
  }
  api.stopCalendarEventsSubscription = stopCalendarEventsSubscription;
  function calendarEventIncludesDate(event, dateKey) {
    const start = String(event.startDate || "");
    const end = String(event.endDate || start);
    return Boolean(dateKey && start && start <= dateKey && dateKey <= end);
  }
  api.calendarEventIncludesDate = calendarEventIncludesDate;
  function getCalendarEventsForDate(dateKey) {
    return calendarEvents
      .filter((event) => calendarEventIncludesDate(event, dateKey))
      .sort((a, b) => {
        const aTime = a.allDay === false ? String(a.startTime || "23:59") : "00:00";
        const bTime = b.allDay === false ? String(b.startTime || "23:59") : "00:00";
        return aTime.localeCompare(bTime) || String(a.title || "").localeCompare(String(b.title || ""), "it");
      });
  }
  api.getCalendarEventsForDate = getCalendarEventsForDate;
  function renderCalendar() {
    renderCalendarMode();
    if (calendarMode === "choice") return;
    if (calendarMode === "hours") {
      renderPersonalHoursCalendar();
      return;
    }
    if (!ui.calendarGrid || !ui.calendarMonthTitle) return;
    const year = calendarVisibleMonth.getFullYear();
    const month = calendarVisibleMonth.getMonth();
    ui.calendarMonthTitle.textContent = new Intl.DateTimeFormat("it-IT", { month: "long", year: "numeric" }).format(calendarVisibleMonth);
    const firstDay = new Date(year, month, 1, 12);
    const mondayOffset = (firstDay.getDay() + 6) % 7;
    const gridStart = new Date(year, month, 1 - mondayOffset, 12);
    const todayKey = formatCalendarDateKey(new Date());
    const cells = [];
  
    for (let index = 0; index < 42; index += 1) {
      const date = new Date(gridStart);
      date.setDate(gridStart.getDate() + index);
      const dateKey = formatCalendarDateKey(date);
      const events = getCalendarEventsForDate(dateKey);
      const typeDots = [...new Set(events.map((event) => String(event.type || "altro")))]
        .slice(0, 3)
        .map((type) => `<span class="calendar-type-dot calendar-type-${escapeHTML(type)}"></span>`)
        .join("");
      const classes = [
        "calendar-day",
        date.getMonth() === month ? "" : "is-outside",
        dateKey === todayKey ? "is-today" : "",
        dateKey === calendarSelectedDate ? "is-selected" : "",
        events.length ? "has-events" : ""
      ].filter(Boolean).join(" ");
      const eventLabel = events.length ? `${events.length} ${events.length === 1 ? "evento" : "eventi"}` : "nessun evento";
      cells.push(`
        <button class="${classes}" type="button" role="gridcell" data-calendar-date="${dateKey}" aria-label="${escapeHTML(formatCalendarLongDate(dateKey))}, ${eventLabel}">
          <span class="calendar-day-number">${date.getDate()}</span>
          ${events.length ? `<span class="calendar-event-count">${events.length}</span>` : ""}
          <span class="calendar-type-dots">${typeDots}</span>
        </button>
      `);
    }
  
    ui.calendarGrid.innerHTML = cells.join("");
    ui.calendarGrid.querySelectorAll("[data-calendar-date]").forEach((button) => {
      button.addEventListener("click", () => selectCalendarDate(button.dataset.calendarDate || ""));
    });
    renderCalendarSelectedDay();
  }
  api.renderCalendar = renderCalendar;
  function selectCalendarDate(dateKey) {
    const date = parseCalendarDateKey(dateKey);
    if (!date) return;
    calendarSelectedDate = dateKey;
    if (date.getMonth() !== calendarVisibleMonth.getMonth() || date.getFullYear() !== calendarVisibleMonth.getFullYear()) {
      calendarVisibleMonth = new Date(date.getFullYear(), date.getMonth(), 1, 12);
    }
    renderCalendar();
  }
  api.selectCalendarDate = selectCalendarDate;
  function changeCalendarMonth(offset) {
    calendarVisibleMonth = new Date(calendarVisibleMonth.getFullYear(), calendarVisibleMonth.getMonth() + Number(offset || 0), 1, 12);
    calendarSelectedDate = formatCalendarDateKey(calendarVisibleMonth);
    renderCalendar();
  }
  api.changeCalendarMonth = changeCalendarMonth;
  function showCalendarToday() {
    const today = new Date();
    calendarVisibleMonth = new Date(today.getFullYear(), today.getMonth(), 1, 12);
    calendarSelectedDate = formatCalendarDateKey(today);
    renderCalendar();
  }
  api.showCalendarToday = showCalendarToday;
  function canModifyCalendarEvent(event) {
    if (event?.type === "SCADENZA_DOCUMENTO") return Boolean(currentUser && String(event?.ownerUserId || "") === String(currentUser.uid || ""));
    return Boolean(currentUser && (canManageData() || String(event?.createdByUid || "") === String(currentUser.uid || "")));
  }
  api.canModifyCalendarEvent = canModifyCalendarEvent;
  function formatCalendarEventPeriod(event) {
    const allDay = event.allDay !== false;
    const startDate = String(event.startDate || "");
    const endDate = String(event.endDate || startDate);
    if (allDay) return startDate === endDate ? "Tutto il giorno" : `Dal ${startDate} al ${endDate}`;
    const time = [event.startTime, event.endTime].filter(Boolean).join("–");
    return startDate === endDate ? (time || "Orario da definire") : `Dal ${startDate} al ${endDate}${time ? ` • ${time}` : ""}`;
  }
  api.formatCalendarEventPeriod = formatCalendarEventPeriod;
  function renderCalendarSelectedDay() {
    if (!ui.calendarDayEvents) return;
    const events = getCalendarEventsForDate(calendarSelectedDate);
    if (ui.calendarSelectedDayTitle) ui.calendarSelectedDayTitle.textContent = formatCalendarLongDate(calendarSelectedDate);
    if (ui.calendarSelectedDaySummary) {
      ui.calendarSelectedDaySummary.textContent = events.length
        ? `${events.length} ${events.length === 1 ? "evento programmato" : "eventi programmati"}`
        : "Nessun evento in questo giorno";
    }
    if (!events.length) {
      ui.calendarDayEvents.innerHTML = "<div class='calendar-empty-day'><span>🗓️</span><p>Nessun evento. Premi “Aggiungi” per inserirne uno.</p></div>";
      return;
    }
    ui.calendarDayEvents.innerHTML = events.map((event) => {
      const isDocumentExpiration = event.type === "SCADENZA_DOCUMENTO";
      const type = isDocumentExpiration ? { icon: "📄", label: "Scadenza documento" } : (CALENDAR_EVENT_TYPES[event.type] || CALENDAR_EVENT_TYPES.altro);
      const mayModify = canModifyCalendarEvent(event);
      const safeLink = /^https?:\/\//i.test(String(event.link || "")) ? String(event.link) : "";
      const detailRows = [
        event.worksite ? `<p><strong>Commessa / impianto:</strong> ${escapeHTML(event.worksite)}</p>` : "",
        event.location ? `<p><strong>Luogo:</strong> ${escapeHTML(event.location)}</p>` : "",
        event.participants ? `<p><strong>Persone:</strong> ${escapeHTML(event.participants)}</p>` : "",
        event.description ? `<p class="calendar-event-description">${escapeHTML(event.description)}</p>` : "",
        safeLink ? `<p><a class="btn calendar-link-btn" href="${escapeHTML(safeLink)}" target="_blank" rel="noopener noreferrer">🔗 Apri link</a></p>` : ""
      ].join("");
      return `
        <article class="calendar-event-card calendar-type-border-${escapeHTML(event.type || "altro")}">
          <div class="calendar-event-heading">
            <span class="calendar-event-icon" aria-hidden="true">${type.icon}</span>
            <div>
              <span class="calendar-event-type">${escapeHTML(type.label)}</span>
              <h3>${escapeHTML(isDocumentExpiration ? (event.compactTitle || event.title || "Documento") : (event.title || "Evento"))}</h3>
              <p class="calendar-event-period">${escapeHTML(formatCalendarEventPeriod(event))}</p>
            </div>
          </div>
          <div class="calendar-event-details">${detailRows}</div>
          <div class="calendar-event-footer">
            <span>Inserito da <strong>${escapeHTML(event.createdByName || event.createdByEmail || "Utente")}</strong></span>
            ${mayModify ? `
              <span class="calendar-event-actions">
                ${isDocumentExpiration ? `<button class="btn" type="button" data-calendar-document="${escapeHTML(event.documentId || "")}">Apri documento</button>` : `<button class="btn" type="button" data-calendar-edit="${escapeHTML(event.id)}">Modifica</button>`}
                ${isDocumentExpiration ? "" : `<button class="btn btn-danger" type="button" data-calendar-delete="${escapeHTML(event.id)}">Elimina</button>`}
              </span>
            ` : ""}
          </div>
        </article>
      `;
    }).join("");
    ui.calendarDayEvents.querySelectorAll("[data-calendar-edit]").forEach((button) => {
      button.addEventListener("click", () => {
        const event = calendarEvents.find((item) => item.id === button.dataset.calendarEdit);
        if (event) openCalendarEventForm(event.startDate, event);
      });
    });
    ui.calendarDayEvents.querySelectorAll("[data-calendar-document]").forEach((button) => {
      button.addEventListener("click", () => window.HeraDocuments?.open({ visibility: "personal" }));
    });
    ui.calendarDayEvents.querySelectorAll("[data-calendar-delete]").forEach((button) => {
      button.addEventListener("click", () => deleteCalendarEvent(button.dataset.calendarDelete || ""));
    });
  }
  api.renderCalendarSelectedDay = renderCalendarSelectedDay;
  function syncCalendarTimeFields() {
    const allDay = Boolean(ui.calendarEventAllDay?.checked);
    ui.calendarEventTimeFields?.classList.toggle("hidden", allDay);
    if (ui.calendarEventStartTime) ui.calendarEventStartTime.required = !allDay;
    if (ui.calendarEventEndTime) ui.calendarEventEndTime.required = false;
  }
  api.syncCalendarTimeFields = syncCalendarTimeFields;
  function getCalendarParticipantSnapshot(person = null, freeName = "") {
    const name = String(person ? getPersonaleDisplayName(person) : freeName).trim();
    return {
      id: String(person?.id || "").trim(),
      name,
      email: String(person?.email || "").trim(),
      freeText: !person
    };
  }
  api.getCalendarParticipantSnapshot = getCalendarParticipantSnapshot;
  function addCalendarParticipant(person = null, freeName = "") {
    const participant = getCalendarParticipantSnapshot(person, freeName);
    if (!participant.name) return;
    const key = normalizeSquadraMemberIdentity(participant.id || participant.email || participant.name);
    if (calendarSelectedParticipants.some((item) => normalizeSquadraMemberIdentity(item.id || item.email || item.name) === key)) return;
    calendarSelectedParticipants.push(participant);
    renderCalendarParticipantPicker();
  }
  api.addCalendarParticipant = addCalendarParticipant;
  function removeCalendarParticipant(index) {
    calendarSelectedParticipants.splice(index, 1);
    renderCalendarParticipantPicker();
  }
  api.removeCalendarParticipant = removeCalendarParticipant;
  function renderCalendarParticipantPicker() {
    if (!ui.calendarParticipantsChips || !ui.calendarEventParticipants) return;
    ui.calendarParticipantsChips.innerHTML = calendarSelectedParticipants.map((participant, index) => `
      <span class="calendar-participant-chip">
        <span>${escapeHTML(participant.name)}</span>
        <button type="button" data-calendar-participant-remove="${index}" aria-label="Rimuovi ${escapeHTML(participant.name)}">×</button>
      </span>
    `).join("");
    ui.calendarParticipantsChips.querySelectorAll("[data-calendar-participant-remove]").forEach((button) => {
      button.addEventListener("click", () => removeCalendarParticipant(Number(button.dataset.calendarParticipantRemove)));
    });
    ui.calendarEventParticipants.value = calendarSelectedParticipants.map((participant) => participant.name).join(", ");
  }
  api.renderCalendarParticipantPicker = renderCalendarParticipantPicker;
  function renderCalendarParticipantSuggestions() {
    if (!ui.calendarParticipantsSuggestions || !ui.calendarParticipantsSearch) return;
    const query = normalizeSquadraMemberIdentity(ui.calendarParticipantsSearch.value);
    const selectedKeys = new Set(calendarSelectedParticipants.map((item) => normalizeSquadraMemberIdentity(item.id || item.email || item.name)));
    const matches = personaleRecords
      .filter((person) => {
        const key = normalizeSquadraMemberIdentity(person.id || person.email || getPersonaleDisplayName(person));
        if (selectedKeys.has(key)) return false;
        return !query || normalizeSquadraMemberIdentity(`${getPersonaleDisplayName(person)} ${person.email || ""}`).includes(query);
      })
      .slice(0, 8);
    const freeValue = String(ui.calendarParticipantsSearch.value || "").trim();
    ui.calendarParticipantsSuggestions.innerHTML = [
      ...matches.map((person) => `<button type="button" role="option" data-calendar-person-id="${escapeHTML(person.id)}"><strong>${escapeHTML(getPersonaleDisplayName(person))}</strong>${person.email ? `<small>${escapeHTML(person.email)}</small>` : ""}</button>`),
      freeValue && !matches.some((person) => normalizeSquadraMemberIdentity(getPersonaleDisplayName(person)) === normalizeSquadraMemberIdentity(freeValue))
        ? `<button type="button" role="option" data-calendar-free-person="${escapeHTML(freeValue)}">＋ Usa nome libero: <strong>${escapeHTML(freeValue)}</strong></button>`
        : ""
    ].join("");
    ui.calendarParticipantsSuggestions.classList.toggle("hidden", !matches.length && !freeValue);
    ui.calendarParticipantsSuggestions.querySelectorAll("[data-calendar-person-id]").forEach((button) => {
      button.addEventListener("click", () => {
        const person = personaleRecords.find((item) => item.id === button.dataset.calendarPersonId);
        if (person) addCalendarParticipant(person);
        ui.calendarParticipantsSearch.value = "";
        renderCalendarParticipantSuggestions();
        ui.calendarParticipantsSearch.focus();
      });
    });
    ui.calendarParticipantsSuggestions.querySelectorAll("[data-calendar-free-person]").forEach((button) => {
      button.addEventListener("click", () => {
        addCalendarParticipant(null, button.dataset.calendarFreePerson || "");
        ui.calendarParticipantsSearch.value = "";
        renderCalendarParticipantSuggestions();
        ui.calendarParticipantsSearch.focus();
      });
    });
  }
  api.renderCalendarParticipantSuggestions = renderCalendarParticipantSuggestions;
  function handleCalendarParticipantSearchKeydown(event) {
    if (event.key === "Escape") {
      ui.calendarParticipantsSuggestions?.classList.add("hidden");
      return;
    }
    if (event.key !== "Enter" && event.key !== ",") return;
    event.preventDefault();
    const freeValue = String(ui.calendarParticipantsSearch?.value || "").trim();
    if (!freeValue) return;
    const exactPerson = personaleRecords.find((person) => normalizeSquadraMemberIdentity(getPersonaleDisplayName(person)) === normalizeSquadraMemberIdentity(freeValue));
    addCalendarParticipant(exactPerson || null, exactPerson ? "" : freeValue);
    ui.calendarParticipantsSearch.value = "";
    renderCalendarParticipantSuggestions();
  }
  api.handleCalendarParticipantSearchKeydown = handleCalendarParticipantSearchKeydown;
  function populateCalendarCommesse(selectedId = "", customName = "") {
    if (!ui.calendarEventCommessa) return;
    const commesse = sortCommesseByCreatedAtDesc(Array.from(commesseById.values()));
    ui.calendarEventCommessa.innerHTML = [
      '<option value="">Nessuna commessa</option>',
      ...commesse.map((commessa) => `<option value="${escapeHTML(commessa.id)}">${escapeHTML(getCommessaDisplayName(commessa))}</option>`),
      '<option value="__custom">＋ Scrivi una commessa non presente</option>'
    ].join("");
    if (selectedId && commesseById.has(selectedId)) ui.calendarEventCommessa.value = selectedId;
    else if (customName) ui.calendarEventCommessa.value = "__custom";
    else ui.calendarEventCommessa.value = "";
    ui.calendarEventCustomCommessa.value = customName || "";
    ui.calendarEventCustomCommessaField?.classList.toggle("hidden", ui.calendarEventCommessa.value !== "__custom");
  }
  api.populateCalendarCommesse = populateCalendarCommesse;
  function getCalendarImpiantoDisplayName(impianto = {}) {
    return String(impianto.denominazione || impianto.nome || impianto.idSap || "Impianto").trim();
  }
  api.getCalendarImpiantoDisplayName = getCalendarImpiantoDisplayName;
  function getCalendarImpiantoLocation(impianto = {}) {
    return [impianto.indirizzo || impianto.descrizioneVia, impianto.comune].map((value) => String(value || "").trim()).filter(Boolean).join(", ");
  }
  api.getCalendarImpiantoLocation = getCalendarImpiantoLocation;
  function populateCalendarImpianti(commessaId = "", selectedId = "", customName = "") {
    if (!ui.calendarEventImpianto) return;
    const impianti = commessaId ? getCommessaCachedImpianti(commessaId) : [];
    ui.calendarEventImpianto.innerHTML = [
      '<option value="">Nessun impianto</option>',
      ...impianti.map((impianto) => `<option value="${escapeHTML(getSquadraImpiantoId(impianto))}">${escapeHTML(getCalendarImpiantoDisplayName(impianto))}</option>`),
      '<option value="__custom">＋ Scrivi un impianto non presente</option>'
    ].join("");
    const hasCustomCommessa = ui.calendarEventCommessa?.value === "__custom";
    ui.calendarEventImpianto.disabled = !commessaId && !customName && !hasCustomCommessa;
    if (selectedId && impianti.some((impianto) => getSquadraImpiantoId(impianto) === selectedId)) ui.calendarEventImpianto.value = selectedId;
    else if (customName) ui.calendarEventImpianto.value = "__custom";
    else ui.calendarEventImpianto.value = "";
    ui.calendarEventCustomImpianto.value = customName || "";
    ui.calendarEventCustomImpiantoField?.classList.toggle("hidden", ui.calendarEventImpianto.value !== "__custom");
  }
  api.populateCalendarImpianti = populateCalendarImpianti;
  function handleCalendarCommessaChange() {
    const commessaId = String(ui.calendarEventCommessa?.value || "");
    const isCustom = commessaId === "__custom";
    ui.calendarEventCustomCommessaField?.classList.toggle("hidden", !isCustom);
    if (!isCustom) ui.calendarEventCustomCommessa.value = "";
    populateCalendarImpianti(isCustom ? "" : commessaId);
  }
  api.handleCalendarCommessaChange = handleCalendarCommessaChange;
  function handleCalendarImpiantoChange() {
    const commessaId = String(ui.calendarEventCommessa?.value || "");
    const impiantoId = String(ui.calendarEventImpianto?.value || "");
    const isCustom = impiantoId === "__custom";
    ui.calendarEventCustomImpiantoField?.classList.toggle("hidden", !isCustom);
    if (!isCustom) ui.calendarEventCustomImpianto.value = "";
    const impianto = getCommessaCachedImpianti(commessaId).find((item) => getSquadraImpiantoId(item) === impiantoId);
    if (impianto && ui.calendarEventLocation) {
      const location = getCalendarImpiantoLocation(impianto);
      if (location) ui.calendarEventLocation.value = location;
    }
  }
  api.handleCalendarImpiantoChange = handleCalendarImpiantoChange;
  function openCalendarEventForm(dateKey = calendarSelectedDate, event = null) {
    if (!currentUser) return;
    const fallbackDate = parseCalendarDateKey(dateKey) ? dateKey : formatCalendarDateKey(new Date());
    ui.calendarEventForm?.reset();
    ui.calendarEventId.value = event?.id || "";
    ui.calendarEventFormTitle.textContent = event ? "Modifica evento" : "Nuovo evento";
    ui.calendarEventType.value = event?.type || "ferie";
    ui.calendarEventTitle.value = event?.title || "";
    ui.calendarEventStartDate.value = event?.startDate || fallbackDate;
    ui.calendarEventEndDate.value = event?.endDate || event?.startDate || fallbackDate;
    ui.calendarEventAllDay.checked = event?.allDay !== false;
    ui.calendarEventStartTime.value = event?.startTime || "";
    ui.calendarEventEndTime.value = event?.endTime || "";
    const initialCommessaId = event?.commessaId || (!event ? selectedCommessaId : "");
    populateCalendarCommesse(initialCommessaId, event?.customCommessa || (!event?.commessaId ? event?.commessaName || event?.worksite || "" : ""));
    populateCalendarImpianti(initialCommessaId, event?.impiantoId || "", event?.customImpianto || (!event?.impiantoId ? event?.impiantoName || "" : ""));
    ui.calendarEventLocation.value = event?.location || "";
    calendarSelectedParticipants = Array.isArray(event?.participantSnapshots)
      ? event.participantSnapshots.map((participant) => ({
        id: String(participant?.id || ""),
        name: String(participant?.name || ""),
        email: String(participant?.email || ""),
        freeText: Boolean(participant?.freeText)
      })).filter((participant) => participant.name)
      : parseMultiEntryValue(event?.participants || "").map((name) => getCalendarParticipantSnapshot(null, name));
    if (!event && !calendarSelectedParticipants.length) {
      const currentPerson = getPersonaleByLoginEmail();
      addCalendarParticipant(currentPerson, currentPerson ? "" : getCurrentUserResolvedName("Utente"));
    } else {
      renderCalendarParticipantPicker();
    }
    if (ui.calendarParticipantsSearch) ui.calendarParticipantsSearch.value = "";
    ui.calendarParticipantsSuggestions?.classList.add("hidden");
    ui.calendarEventDescription.value = event?.description || "";
    ui.calendarEventLink.value = event?.link || "";
    ui.calendarEventFormFeedback.textContent = "";
    syncCalendarTimeFields();
    if (typeof ui.calendarEventDialog.showModal === "function") ui.calendarEventDialog.showModal();
    else ui.calendarEventDialog.setAttribute("open", "");
    setTimeout(() => ui.calendarEventTitle?.focus(), 50);
  }
  api.openCalendarEventForm = openCalendarEventForm;
  function closeCalendarEventForm() {
    if (!ui.calendarEventDialog) return;
    if (typeof ui.calendarEventDialog.close === "function" && ui.calendarEventDialog.open) ui.calendarEventDialog.close();
    else ui.calendarEventDialog.removeAttribute("open");
  }
  api.closeCalendarEventForm = closeCalendarEventForm;
  function getCalendarAuthorName() {
    const profile = platformUsers.find((user) => String(user.id || user.uid || "") === String(currentUser?.uid || ""));
    return String(profile?.displayName || profile?.nome || currentUser?.displayName || currentUser?.email || "Utente").trim();
  }
  api.getCalendarAuthorName = getCalendarAuthorName;
  async function saveCalendarEvent(event) {
    event.preventDefault();
    if (!currentUser || !db) return;
    const eventId = String(ui.calendarEventId.value || "").trim();
    const existing = calendarEvents.find((item) => item.id === eventId);
    if (existing && !canModifyCalendarEvent(existing)) {
      ui.calendarEventFormFeedback.textContent = "Non puoi modificare un evento inserito da un altro utente.";
      return;
    }
    const startDate = String(ui.calendarEventStartDate.value || "");
    const endDate = String(ui.calendarEventEndDate.value || startDate);
    const allDay = Boolean(ui.calendarEventAllDay.checked);
    const startTime = allDay ? "" : String(ui.calendarEventStartTime.value || "");
    const endTime = allDay ? "" : String(ui.calendarEventEndTime.value || "");
    if (!startDate || !endDate || endDate < startDate) {
      ui.calendarEventFormFeedback.textContent = "Controlla le date: la data finale non può precedere quella iniziale.";
      return;
    }
    if (!allDay && !startTime) {
      ui.calendarEventFormFeedback.textContent = "Inserisci almeno l'ora di inizio.";
      return;
    }
    if (!allDay && startDate === endDate && endTime && endTime < startTime) {
      ui.calendarEventFormFeedback.textContent = "L'ora finale non può precedere quella iniziale.";
      return;
    }
    const commessaSelection = String(ui.calendarEventCommessa.value || "");
    const customCommessa = commessaSelection === "__custom" ? String(ui.calendarEventCustomCommessa.value || "").trim() : "";
    const commessa = commessaSelection && commessaSelection !== "__custom" ? commesseById.get(commessaSelection) : null;
    const impiantoSelection = String(ui.calendarEventImpianto.value || "");
    const customImpianto = impiantoSelection === "__custom" ? String(ui.calendarEventCustomImpianto.value || "").trim() : "";
    const impianto = commessa
      ? getCommessaCachedImpianti(commessa.id).find((item) => getSquadraImpiantoId(item) === impiantoSelection)
      : null;
    const eventType = String(ui.calendarEventType.value || "altro");
    const generatedTitle = `${CALENDAR_EVENT_TYPES[eventType]?.label || "Evento"}${calendarSelectedParticipants[0]?.name ? ` • ${calendarSelectedParticipants[0].name}` : ""}`;
    const payload = {
      type: eventType,
      title: String(ui.calendarEventTitle.value || "").trim() || generatedTitle,
      titleWasGenerated: !String(ui.calendarEventTitle.value || "").trim(),
      startDate,
      endDate,
      allDay,
      startTime,
      endTime,
      commessaId: commessa?.id || "",
      commessaName: String(commessa?.nome || customCommessa || "").trim(),
      customCommessa,
      impiantoId: impianto ? getSquadraImpiantoId(impianto) : "",
      impiantoName: String(impianto ? getCalendarImpiantoDisplayName(impianto) : customImpianto).trim(),
      customImpianto,
      worksite: [commessa?.nome || customCommessa, impianto ? getCalendarImpiantoDisplayName(impianto) : customImpianto].filter(Boolean).join(" • "),
      location: String(ui.calendarEventLocation.value || "").trim(),
      participants: calendarSelectedParticipants.map((participant) => participant.name).join(", "),
      participantIds: calendarSelectedParticipants.map((participant) => participant.id).filter(Boolean),
      participantEmails: calendarSelectedParticipants.map((participant) => participant.email).filter(Boolean),
      participantSnapshots: calendarSelectedParticipants,
      description: String(ui.calendarEventDescription.value || "").trim(),
      link: String(ui.calendarEventLink.value || "").trim(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedByUid: currentUser.uid || "",
      updatedByEmail: currentUser.email || ""
    };
    ui.calendarEventSaveBtn.disabled = true;
    ui.calendarEventFormFeedback.textContent = "Salvataggio...";
    try {
      if (eventId) {
        await db.collection("calendarEvents").doc(eventId).set(payload, { merge: true });
      } else {
        await db.collection("calendarEvents").add({
          ...payload,
          createdByUid: currentUser.uid || "",
          createdByEmail: currentUser.email || "",
          createdByName: getCalendarAuthorName(),
          createdAt: firebase.firestore.FieldValue.serverTimestamp()
        });
      }
      calendarSelectedDate = startDate;
      const savedDate = parseCalendarDateKey(startDate);
      if (savedDate) calendarVisibleMonth = new Date(savedDate.getFullYear(), savedDate.getMonth(), 1, 12);
      closeCalendarEventForm();
    } catch (error) {
      console.error("Salvataggio evento calendario non riuscito:", error);
      ui.calendarEventFormFeedback.textContent = error?.message || "Errore durante il salvataggio dell'evento.";
    } finally {
      ui.calendarEventSaveBtn.disabled = false;
    }
  }
  api.saveCalendarEvent = saveCalendarEvent;
  async function deleteCalendarEvent(eventId) {
    const event = calendarEvents.find((item) => item.id === eventId);
    if (!event || !canModifyCalendarEvent(event)) {
      alert("Può eliminare questo evento solo chi lo ha inserito o un amministratore.");
      return;
    }
    if (!window.confirm(`Eliminare l'evento “${event.title || "Evento"}”?`)) return;
    try {
      await db.collection("calendarEvents").doc(eventId).delete();
    } catch (error) {
      console.error("Eliminazione evento calendario non riuscita:", error);
      alert(error?.message || "Impossibile eliminare l'evento.");
    }
  }
  api.deleteCalendarEvent = deleteCalendarEvent;
  async function getCalendarAbsencesForDate(dateKey, { force = false } = {}) {
    const key = String(dateKey || "").trim();
    if (!key || !db) return [];
    const cached = calendarAbsenceCache.get(key);
    if (!force && cached && Date.now() - cached.loadedAt < 60 * 1000) return cached.items;
    try {
      const snapshot = await db.collection("calendarEvents").where("startDate", "<=", key).get();
      const items = snapshot.docs
        .map((doc) => ({ id: doc.id, ...doc.data() }))
        .filter((event) => CALENDAR_ABSENCE_TYPES.has(String(event.type || "")) && String(event.endDate || event.startDate || "") >= key);
      calendarAbsenceCache.set(key, { loadedAt: Date.now(), items });
      return items;
    } catch (error) {
      console.warn("Controllo assenze calendario non disponibile:", error);
      return [];
    }
  }
  api.getCalendarAbsencesForDate = getCalendarAbsencesForDate;
  function getCalendarAbsenceParticipantNames(event = {}) {
    const snapshots = Array.isArray(event.participantSnapshots)
      ? event.participantSnapshots.map((participant) => participant?.name || participant?.email || participant?.id)
      : [];
    return [...new Set([
      ...snapshots,
      ...parseMultiEntryValue(event.participants || ""),
      ...(Array.isArray(event.participantEmails) ? event.participantEmails : []),
      ...(Array.isArray(event.participantIds) ? event.participantIds : [])
    ].map((value) => String(value || "").trim()).filter(Boolean))];
  }
  api.getCalendarAbsenceParticipantNames = getCalendarAbsenceParticipantNames;
  function calendarAbsenceMatchesOperator(event, operatorName) {
    const operatorKey = normalizeSquadraMemberIdentity(operatorName);
    if (!operatorKey) return false;
    return getCalendarAbsenceParticipantNames(event).some((participant) => {
      const participantKey = normalizeSquadraMemberIdentity(participant);
      if (!participantKey) return false;
      return participantKey === operatorKey || participantKey.includes(operatorKey) || operatorKey.includes(participantKey);
    });
  }
  api.calendarAbsenceMatchesOperator = calendarAbsenceMatchesOperator;
  function formatCalendarAbsenceWarning(event, operatorName, dateKey) {
    const type = CALENDAR_EVENT_TYPES[event.type] || CALENDAR_EVENT_TYPES.altro;
    return `⛔ ${operatorName} assente: ${type.label} il ${formatDateKeyForDisplay(dateKey)}`;
  }
  api.formatCalendarAbsenceWarning = formatCalendarAbsenceWarning;
  async function applyCalendarAbsenceWarningsToSquadraRows(rows, dateKey) {
    const absences = await getCalendarAbsencesForDate(dateKey, { force: true });
    for (const row of rows) {
      const members = getSquadraRowMemberNames(row);
      const warnings = [];
      const eventIds = [];
      for (const member of members) {
        const matching = absences.filter((event) => calendarAbsenceMatchesOperator(event, member));
        for (const event of matching) {
          const available = await validateSquadraOperatorAvailability(member, dateKey);
          if (!available) return false;
          warnings.push(formatCalendarAbsenceWarning(event, member, dateKey));
          eventIds.push(event.id);
        }
      }
      row.avvisoAutomaticoAssenze = [...new Set(warnings)].join("\n");
      row.calendarAbsenceEventIds = [...new Set(eventIds)];
    }
    return true;
  }
  api.applyCalendarAbsenceWarningsToSquadraRows = applyCalendarAbsenceWarningsToSquadraRows;
  function openNotificationCalendarView() {
    if (!ui.notificationCalendarView || !ui.notificationMainView) return;
    ui.notificationMainView.classList.add("hidden");
    ui.notificationCalendarView.classList.remove("hidden");
    renderNotificationCalendar();
  }
  api.openNotificationCalendarView = openNotificationCalendarView;
  function closeNotificationCalendarView() {
    if (!ui.notificationCalendarView || !ui.notificationMainView) return;
    ui.notificationCalendarView.classList.add("hidden");
    ui.notificationMainView.classList.remove("hidden");
  }
  api.closeNotificationCalendarView = closeNotificationCalendarView;
  function moveNotificationCalendarMonth(offset) {
    notificationCalendarCursor = new Date(notificationCalendarCursor.getFullYear(), notificationCalendarCursor.getMonth() + offset, 1);
    renderNotificationCalendar();
  }
  api.moveNotificationCalendarMonth = moveNotificationCalendarMonth;
  function renderNotificationCalendar() {
    if (!ui.notificationCalendarGrid || !ui.notificationCalendarMonthLabel) return;
    const monthStart = new Date(notificationCalendarCursor.getFullYear(), notificationCalendarCursor.getMonth(), 1);
    const monthLabel = monthStart.toLocaleDateString("it-IT", { month: "long", year: "numeric" });
    ui.notificationCalendarMonthLabel.textContent = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
    const weekdayLabels = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"];
    const dayMap = new Map();
    userAlerts.forEach((item) => {
      const key = getNotificationPrimaryDateKey(item);
      if (!key) return;
      if (!dayMap.has(key)) dayMap.set(key, []);
      dayMap.get(key).push(item);
    });
    const firstWeekday = (monthStart.getDay() + 6) % 7;
    const daysInMonth = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0).getDate();
    ui.notificationCalendarGrid.innerHTML = "";
    weekdayLabels.forEach((label) => {
      const header = document.createElement("div");
      header.className = "notification-calendar-cell notification-calendar-weekday";
      header.textContent = label;
      ui.notificationCalendarGrid.appendChild(header);
    });
    for (let i = 0; i < firstWeekday; i += 1) {
      const empty = document.createElement("div");
      empty.className = "notification-calendar-cell is-empty";
      ui.notificationCalendarGrid.appendChild(empty);
    }
    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = new Date(monthStart.getFullYear(), monthStart.getMonth(), day);
      const dateKey = getDateKeyFromLocalDate(date);
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "notification-calendar-cell notification-calendar-day";
      btn.textContent = String(day);
      if (dayMap.has(dateKey)) btn.classList.add("has-notification");
      if (selectedNotificationCalendarDateKey === dateKey) btn.classList.add("is-selected");
      btn.addEventListener("click", () => openNotificationDayDetail(dateKey));
      ui.notificationCalendarGrid.appendChild(btn);
    }
  }
  api.renderNotificationCalendar = renderNotificationCalendar;
  Object.assign(global, api);
  global.VargaCalendarModule = Object.freeze({ ...api });
})(window);
