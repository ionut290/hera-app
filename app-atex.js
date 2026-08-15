"use strict";
(function installVargaAtexModule(global) {
  if (global.VargaAtexModule) return;
  const api = {};
  function valueContainsAtex(value) {
    return /\bATEX\b/.test(normalizeAtexSearchValue(value));
  }
  api.valueContainsAtex = valueContainsAtex;
  function isTruthyAtexFlag(value) {
    if (value === true) return true;
    if (value === false || value === null || value === undefined) return false;
    const normalized = normalizeAtexSearchValue(value);
    return ["ATEX", "TRUE", "SI", "SÌ", "YES", "1", "EX"].includes(normalized) || valueContainsAtex(normalized);
  }
  api.isTruthyAtexFlag = isTruthyAtexFlag;
  function hasImpiantoAtexFlag(impianto = {}) {
    const explicitFlags = [
      impianto.atex,
      impianto.isAtex,
      impianto.flagAtex,
      impianto.atexFlag,
      impianto.areaAtex,
      impianto.zonaAtex,
      impianto.ex
    ];
    if (explicitFlags.some(isTruthyAtexFlag)) return true;
    return [
      impianto.category,
      impianto.categoria,
      impianto.flag,
      impianto.flags,
      impianto.tipo,
      impianto.tipologia,
      impianto.tipologiaImpianto,
      impianto.lavorazioniRichieste,
      impianto.note,
      impianto.codicePrezzo,
      impianto.voceRiferimento,
      impianto.denominazione
    ].some(valueContainsAtex);
  }
  api.hasImpiantoAtexFlag = hasImpiantoAtexFlag;
  function shouldShowAtexButtonForImpianto(impianto = {}) {
    if (isCurrentCommessaDepurazioneOrDiscariche()) return false;
    return Boolean(isCurrentCommessaInrete() || hasImpiantoAtexFlag(impianto));
  }
  api.shouldShowAtexButtonForImpianto = shouldShowAtexButtonForImpianto;
  function getAtexProcedureImpiantoByKey(key) {
    return getDettaglioMeteoImpiantoByKey(key);
  }
  api.getAtexProcedureImpiantoByKey = getAtexProcedureImpiantoByKey;
  function openAtexProcedurePage(impianto) {
    if (!selectedCommessaId || !impianto) return;
    const key = buildImpiantoKey(impianto);
    window.location.hash = `commessa=${encodeURIComponent(selectedCommessaId)}&atex=${encodeURIComponent(key)}`;
    applyRoute();
  }
  api.openAtexProcedurePage = openAtexProcedurePage;
  function closeAtexProcedurePage() {
    openImpiantiPage();
  }
  api.closeAtexProcedurePage = closeAtexProcedurePage;
  function handleAtexProcedureButtonClick(event) {
    const button = event.target?.closest?.("[data-atex-procedure]");
    if (!button) return;
    event.preventDefault();
    event.stopPropagation();
    const key = button.getAttribute("data-atex-procedure") || button.closest("[data-weather-card]")?.getAttribute("data-weather-card") || "";
    const impianto = findImpiantoByWeatherKey(key) || getAtexProcedureImpiantoByKey(key);
    if (!impianto) return;
    openAtexProcedurePage(impianto);
  }
  api.handleAtexProcedureButtonClick = handleAtexProcedureButtonClick;
  function getCurrentAtexProcedureContext() {
    const route = parseCommessaHash();
    const impiantoKey = route.atex || "";
    const impianto = getAtexProcedureImpiantoByKey(impiantoKey) || {};
    return { impiantoKey, impianto };
  }
  api.getCurrentAtexProcedureContext = getCurrentAtexProcedureContext;
  function formatAtexDateValue(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  api.formatAtexDateValue = formatAtexDateValue;
  function formatAtexTimeValue(date = new Date()) {
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    return `${hours}:${minutes}`;
  }
  api.formatAtexTimeValue = formatAtexTimeValue;
  function buildAtexChecklist(items = []) {
    return items.map((item, index) => `
      <label class="atex-check-row">
        <input type="checkbox" name="check_${index}" value="${escapeHTML(item)}">
        <span>${escapeHTML(item)}</span>
      </label>
    `).join("");
  }
  api.buildAtexChecklist = buildAtexChecklist;
  function buildAtexList(items = []) {
    return `<ul>${items.map((item) => `<li>${escapeHTML(item)}</li>`).join("")}</ul>`;
  }
  api.buildAtexList = buildAtexList;
  function normalizeAtexCommessaMatchValue(value) {
    return normalizeAtexSearchValue(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }
  api.normalizeAtexCommessaMatchValue = normalizeAtexCommessaMatchValue;
  function getSelectedAtexCommessaName() {
    const selected = commesseById.get(selectedCommessaId) || {};
    return selectedCommessaName || selected.nome || selected.name || selected.codice || selected.code || selectedCommessaId || "Commessa non indicata";
  }
  api.getSelectedAtexCommessaName = getSelectedAtexCommessaName;
  function getAtexClientContactsForCommessa() {
    const clients = [
      { area: "Bologna", name: "Carani Claudio", phone: "347 7614277" },
      { area: "Modena", name: "Montagnana Giorgio", phone: "320 4791013" },
      { area: "Ferrara", name: "Mateo Gardelini", phone: "348 0900290" }
    ];
    const selected = commesseById.get(selectedCommessaId) || {};
    const haystack = [
      selectedCommessaName,
      selectedCommessaId,
      selected.nome,
      selected.name,
      selected.codice,
      selected.code,
      selected.cliente,
      selected.customer,
      selected.category,
      selected.categoria,
      selected.tipologia
    ].map(normalizeAtexCommessaMatchValue).join(" ");
    const matched = clients.find((client) => haystack.includes(normalizeAtexCommessaMatchValue(client.area)));
    if (matched) return [{ ...matched, role: `Cliente INRETE ${matched.area}`, callFirst: true }];
    return clients.map((client) => ({
      ...client,
      role: `Cliente INRETE ${client.area}`,
      note: "cliente di riferimento da verificare",
      callFirst: true
    }));
  }
  api.getAtexClientContactsForCommessa = getAtexClientContactsForCommessa;
  function getAtexWhatsappText(impianto = {}) {
    const operator = currentUser?.displayName || currentUser?.email || "Operatore non indicato";
    return [
      "⚠️ PROBLEMA ATEX RILEVATO",
      `Commessa: ${getSelectedAtexCommessaName()}`,
      `Impianto: ${impianto.denominazione || impianto.nome || "Impianto non indicato"}`,
      `Comune: ${impianto.comune || "Comune non indicato"}`,
      `Operatore: ${operator}`,
      "Serve intervento/verifica urgente."
    ].join("\n");
  }
  api.getAtexWhatsappText = getAtexWhatsappText;
  function buildAtexContactCard(contact, whatsappText) {
    const phoneHref = sanitizePhoneHref(contact.phone);
    const whatsappPhone = sanitizeWhatsappPhone(contact.phone);
    const whatsappUrl = `https://wa.me/${encodeURIComponent(whatsappPhone)}?text=${encodeURIComponent(whatsappText)}`;
    const noteMarkup = contact.note ? `<span class="atex-contact-note">${escapeHTML(contact.note)}</span>` : "";
    const badgeMarkup = contact.callFirst ? `<span class="atex-call-first-badge">CHIAMARE PRIMA</span>` : "";
    return `
      <section class="atex-contact-card${contact.callFirst ? " is-primary" : ""}">
        <div class="atex-contact-head">
          <span class="atex-contact-role">${escapeHTML(contact.role)}</span>
          ${badgeMarkup}
        </div>
        ${noteMarkup}
        <strong class="atex-contact-name">${escapeHTML(contact.name)}</strong>
        <span class="atex-contact-phone">Tel: ${escapeHTML(contact.phone)}</span>
        <div class="atex-contact-actions">
          <a class="atex-contact-action call" href="tel:${escapeHTML(phoneHref)}">📞 CHIAMA</a>
          <a class="atex-contact-action whatsapp" href="${escapeHTML(whatsappUrl)}" target="_blank" rel="noopener noreferrer">💬 WHAZZUP</a>
        </div>
      </section>
    `;
  }
  api.buildAtexContactCard = buildAtexContactCard;
  function buildAtexEmergencyContactsSection(impianto = {}) {
    const contacts = [
      ...getAtexClientContactsForCommessa(),
      { role: "Capo squadra", name: "Varga Ionel", phone: "0039 389 2352575" },
      { role: "Responsabile commessa", name: "Alessandro Minarini", phone: "+39 335 6815371" },
      { role: "Numero unico emergenze", name: "Numero unico emergenze", phone: "112" },
      { role: "Vigili del Fuoco", name: "Vigili del Fuoco", phone: "115" },
      { role: "Emergenza sanitaria / ambulanza", name: "Emergenza sanitaria / ambulanza", phone: "118" },
      { role: "Carabinieri", name: "Carabinieri", phone: "112" },
      { role: "Polizia", name: "Polizia", phone: "113" }
    ];
    const whatsappText = getAtexWhatsappText(impianto);
    return `
      <article class="atex-procedure-section atex-emergency-section">
        <h3>6. CONTATTI EMERGENZA ATEX</h3>
        <div class="atex-contacts-grid">${contacts.map((contact) => buildAtexContactCard(contact, whatsappText)).join("")}</div>
      </article>
    `;
  }
  api.buildAtexEmergencyContactsSection = buildAtexEmergencyContactsSection;
  function getAtexIllustrationSvg(type = "safety") {
    const svgMap = {
      safety: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 240" role="img"><defs><linearGradient id="sky" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#fed7aa"/><stop offset="1" stop-color="#dcfce7"/></linearGradient></defs><rect width="420" height="240" rx="28" fill="url(#sky)"/><path d="M0 174c60-26 118-22 178 2s119 26 242-12v76H0z" fill="#86efac"/><rect x="48" y="71" width="118" height="86" rx="14" fill="#fff7ed" stroke="#9a3412" stroke-width="6"/><path d="M78 124h59M78 101h59" stroke="#9a3412" stroke-width="10" stroke-linecap="round"/><circle cx="292" cy="96" r="46" fill="#fb923c"/><path d="M292 57l36 70h-72z" fill="#fff7ed" stroke="#7c2d12" stroke-width="7"/><path d="M292 78v24" stroke="#7c2d12" stroke-width="8" stroke-linecap="round"/><circle cx="292" cy="114" r="5" fill="#7c2d12"/><path d="M208 163c26-30 65-30 91 0" fill="none" stroke="#166534" stroke-width="10" stroke-linecap="round"/></svg>`,
      checklist: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 210" role="img"><rect width="320" height="210" rx="24" fill="#ffedd5"/><rect x="76" y="30" width="168" height="150" rx="18" fill="#fffaf3" stroke="#fdba74" stroke-width="6"/><rect x="119" y="19" width="82" height="28" rx="12" fill="#9a3412"/><g fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M105 78l13 13 29-31" stroke="#16a34a" stroke-width="10"/><path d="M162 80h48" stroke="#7c2d12" stroke-width="8"/><path d="M105 124l13 13 29-31" stroke="#16a34a" stroke-width="10"/><path d="M162 126h48" stroke="#7c2d12" stroke-width="8"/></g></svg>`,
      altair: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 240" role="img"><rect width="320" height="240" rx="28" fill="#fff7ed"/><rect x="96" y="26" width="128" height="188" rx="30" fill="#2f2f2f" stroke="#111827" stroke-width="8"/><rect x="118" y="52" width="84" height="50" rx="10" fill="#93c5fd"/><text x="160" y="83" text-anchor="middle" font-family="Arial" font-size="20" font-weight="900" fill="#0f172a">4XR</text><circle cx="160" cy="139" r="28" fill="#f97316"/><circle cx="160" cy="139" r="13" fill="#ffedd5"/><circle cx="124" cy="179" r="10" fill="#fb923c"/><circle cx="160" cy="179" r="10" fill="#22c55e"/><circle cx="196" cy="179" r="10" fill="#ef4444"/><path d="M125 28c4-16 66-16 70 0" fill="none" stroke="#6b7280" stroke-width="10" stroke-linecap="round"/></svg>`,
      forbidden: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fef2f2"/><circle cx="110" cy="75" r="47" fill="#fff" stroke="#dc2626" stroke-width="14"/><path d="M78 107l64-64" stroke="#dc2626" stroke-width="14" stroke-linecap="round"/><path d="M107 43c17 17-17 22 0 39 16 17-15 24 6 39" fill="none" stroke="#7f1d1d" stroke-width="8" stroke-linecap="round"/></svg>`,
      dpi: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fff7ed"/><path d="M63 103c8-31 28-50 47-50s39 19 47 50z" fill="#fb923c" stroke="#9a3412" stroke-width="6"/><path d="M72 69c5-23 23-39 38-39s33 16 38 39" fill="#fdba74"/><rect x="54" y="101" width="112" height="18" rx="9" fill="#7c2d12"/><path d="M85 76h50" stroke="#fff7ed" stroke-width="8" stroke-linecap="round"/></svg>`,
      alarm: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fffbeb"/><path d="M110 24l70 102H40z" fill="#facc15" stroke="#92400e" stroke-width="8" stroke-linejoin="round"/><path d="M110 61v32" stroke="#7c2d12" stroke-width="10" stroke-linecap="round"/><circle cx="110" cy="111" r="6" fill="#7c2d12"/></svg>`,
      danger: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fff7ed"/><path d="M44 100c42-45 91-45 132 0" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round"/><rect x="73" y="45" width="74" height="56" rx="12" fill="#431407"/><circle cx="110" cy="73" r="22" fill="#fb923c"/><path d="M110 55v36M92 73h36" stroke="#fff7ed" stroke-width="7" stroke-linecap="round"/></svg>`
    };
    return svgMap[type] || svgMap.safety;
  }
  api.getAtexIllustrationSvg = getAtexIllustrationSvg;
  function buildAtexImageCard(type, alt, extraClass = "") {
    const src = `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(getAtexIllustrationSvg(type))}`;
    return `<figure class="atex-image-card ${escapeHTML(extraClass)}"><img src="${src}" alt="${escapeHTML(alt)}" loading="lazy"></figure>`;
  }
  api.buildAtexImageCard = buildAtexImageCard;
  function handleAtexProcedureContentClick(event) {
    const checklistToggle = event.target?.closest?.("[data-atex-checklist-toggle]");
    if (checklistToggle) {
      event.preventDefault();
      const targetId = checklistToggle.getAttribute("aria-controls");
      const panel = targetId ? document.getElementById(targetId) : null;
      const arrow = checklistToggle.querySelector("[data-atex-checklist-arrow]");
      const isOpen = checklistToggle.getAttribute("aria-expanded") === "true";
      checklistToggle.setAttribute("aria-expanded", String(!isOpen));
      if (panel) panel.hidden = isOpen;
      if (arrow) arrow.textContent = isOpen ? "▼" : "▲";
      return;
    }
  
    const openFormButton = event.target?.closest?.("[data-open-atex-form]");
    if (!openFormButton) return;
    event.preventDefault();
    const form = ui.atexProcedureContent?.querySelector?.("#atex-module-form");
    if (!form) return;
    form.classList.remove("hidden");
    openFormButton.classList.add("hidden");
    form.querySelector("input, select, textarea")?.focus?.();
  }
  api.handleAtexProcedureContentClick = handleAtexProcedureContentClick;
  async function saveAtexProcedureForm(event) {
    const form = event.target?.closest?.("#atex-module-form");
    if (!form) return;
    event.preventDefault();
    const feedback = form.querySelector("[data-atex-form-feedback]");
    const submitButton = form.querySelector("button[type='submit']");
    const { impiantoKey, impianto } = getCurrentAtexProcedureContext();
    const formData = new FormData(form);
    const payload = {
      commessaId: selectedCommessaId || "",
      commessaName: selectedCommessaName || "",
      impiantoKey,
      impiantoName: impianto.denominazione || formData.get("impianto") || "",
      impiantoComune: impianto.comune || "",
      operatore: String(formData.get("operatore") || "").trim(),
      squadra: String(formData.get("squadra") || "").trim(),
      impianto: String(formData.get("impianto") || "").trim(),
      data: String(formData.get("data") || "").trim(),
      ora: String(formData.get("ora") || "").trim(),
      presenzaGas: String(formData.get("presenzaGas") || "").trim(),
      altairVerificato: String(formData.get("altairVerificato") || "").trim(),
      dpiVerificati: String(formData.get("dpiVerificati") || "").trim(),
      noteOperative: String(formData.get("noteOperative") || "").trim(),
      firma: String(formData.get("firma") || "").trim(),
      checklist: Array.from(form.querySelectorAll(".atex-check-row input:checked")).map((input) => input.value),
      createdByUid: auth.currentUser?.uid || "",
      createdByEmail: auth.currentUser?.email || "",
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    };
    if (!payload.operatore || !payload.squadra || !payload.impianto || !payload.data || !payload.ora || !payload.firma) {
      if (feedback) feedback.textContent = "Compila operatore, squadra, impianto, data, ora e firma.";
      return;
    }
    try {
      if (submitButton) submitButton.disabled = true;
      if (feedback) feedback.textContent = "Salvataggio modulo ATEX in corso…";
      await db.collection("commesse").doc(selectedCommessaId).collection("atexModules").add(payload);
      form.reset();
      form.querySelector("[name='impianto']").value = payload.impianto;
      form.querySelector("[name='data']").value = formatAtexDateValue();
      form.querySelector("[name='ora']").value = formatAtexTimeValue();
      if (feedback) feedback.textContent = "Modulo ATEX salvato e collegato a commessa e impianto.";
    } catch (error) {
      console.error("Modulo ATEX non salvato:", error);
      if (feedback) feedback.textContent = "Errore durante il salvataggio del modulo ATEX. Riprova.";
    } finally {
      if (submitButton) submitButton.disabled = false;
    }
  }
  api.saveAtexProcedureForm = saveAtexProcedureForm;
  function renderAtexProcedurePage(impiantoKey) {
    if (!ui.atexProcedureContent) return;
    const impianto = getAtexProcedureImpiantoByKey(impiantoKey) || {};
    const impiantoName = impianto.denominazione || "Impianto";
    if (ui.atexProcedureSubtitle) {
      ui.atexProcedureSubtitle.textContent = `${impiantoName} • ${selectedCommessaName || "Commessa"}`;
    }
    const mandatoryChecks = [
      "Verifica presenza gas",
      "Controllo zona classificata",
      "Controllo DPI",
      "Verifica estintore",
      "Verifica vie di fuga",
      "Controllo vento e ventilazione area",
      "Controllo autorizzazione lavoro",
      "Compilazione modulo ATEX",
      "Verifica comunicazione squadra",
      "Controllo assenza inneschi/scintille"
    ];
    const altairChecks = [
      "Accendere il dispositivo prima di entrare o avvicinarsi all’area",
      "Attendere autotest completo",
      "Verificare batteria sufficiente",
      "Verificare sensori attivi",
      "Eseguire bump test se previsto dalla procedura aziendale",
      "Non entrare in zona se il dispositivo segnala errore",
      "Tenere il dispositivo vicino alla respirazione",
      "Controllare continuamente eventuali allarmi acustici, visivi o vibrazione"
    ];
    const statusIndicators = ["Vento", "Meteo", "Pioggia", "Segnalazioni", "Sensori", "Checklist"];
    ui.atexProcedureContent.innerHTML = `
      <article class="atex-procedure-alert">
        <strong>1. AVVISO IMPORTANTE</strong>
        <p>Leggere attentamente e completare tutti i controlli prima di iniziare ogni attività.</p>
      </article>
  
      <article class="atex-procedure-section atex-hero-section">
        ${buildAtexImageCard("safety", "Area verde esterna con segnalazione sicurezza ATEX", "atex-hero-image")}
        <div>
          <h3>SICUREZZA ATEX IN AREA VERDE ESTERNA</h3>
          <p>Procedura operativa sintetica per lavorare solo nelle aree verdi autorizzate, mantenendo controlli, DPI e comunicazioni sempre attivi.</p>
        </div>
      </article>
  
      <article class="atex-procedure-section atex-access-limits">
        <h3>2. LIMITAZIONE ACCESSI</h3>
        <p><strong>È vietato entrare all’interno delle strutture.</strong></p>
        <p>L’accesso è consentito solo nel cortile e nelle aree verdi esterne autorizzate.</p>
      </article>
  
      <article class="atex-procedure-section atex-team-warning">
        <strong>ATTENZIONE</strong>
        <p>Le squadre di manutenzione verde NON devono entrare nelle strutture, locali tecnici, vasche, pozzetti o ambienti confinati.</p>
        <p>È consentito operare solo nel cortile e nelle aree verdi esterne autorizzate.</p>
        <p>In caso di dubbio fermarsi e contattare il responsabile.</p>
      </article>
  
      <article class="atex-procedure-section atex-collapsible-section">
        <button class="atex-section-toggle" type="button" data-atex-checklist-toggle aria-expanded="false" aria-controls="atex-mandatory-checklist">
          <span>3. CONTROLLI OBBLIGATORI PRIMA ATTIVITÀ</span>
          <span class="atex-section-arrow" data-atex-checklist-arrow aria-hidden="true">▼</span>
        </button>
        <div id="atex-mandatory-checklist" class="atex-collapsible-panel" hidden>
          ${buildAtexImageCard("checklist", "Checklist sicurezza prima attività", "atex-checklist-image")}
          <div class="atex-checklist-grid">${buildAtexChecklist(mandatoryChecks)}</div>
        </div>
      </article>
  
      <article class="atex-procedure-section atex-altair-box">
        <h3>4. UTILIZZO DISPOSITIVO ALTAIR 4XR</h3>
        <div class="atex-altair-layout">
          ${buildAtexImageCard("altair", "Dispositivo rilevatore multigas Altair 4XR", "atex-altair-image")}
          <div class="atex-device-card">
            <p><strong>Nome dispositivo:</strong> Altair 4XR</p>
            <p><strong>Uso:</strong> rilevatore multigas personale</p>
            <p><strong>Obbligo:</strong> accendere prima di avvicinarsi alla zona ATEX</p>
            <p><strong>Posizione:</strong> tenere vicino alla zona di respirazione</p>
          </div>
        </div>
        ${buildAtexList(altairChecks)}
        <div class="atex-alarm-box">
          <strong>In caso di allarme:</strong>
          <ul>
            <li>interrompere immediatamente l’attività</li>
            <li>allontanarsi dalla zona in sicurezza</li>
            <li>avvisare squadra e responsabile</li>
            <li>non rientrare senza autorizzazione</li>
            <li>vietato riavviare attività senza verifica</li>
          </ul>
        </div>
      </article>
  
      <article class="atex-procedure-section">
        <h3>5. NORME OPERATIVE ATEX</h3>
        <div class="atex-rules-grid">
          <section class="atex-rule-box atex-rule-forbidden">${buildAtexImageCard("forbidden", "Divieti in zona ATEX", "atex-rule-image")}<h4>🔥 DIVIETI</h4>${buildAtexList(["Vietato fumare", "Vietato usare fiamme libere", "Vietato produrre scintille", "Vietato utilizzare utensili non certificati", "Vietato usare telefoni non autorizzati in area ATEX"])}</section>
          <section class="atex-rule-box">${buildAtexImageCard("dpi", "DPI obbligatori", "atex-rule-image")}<h4>🦺 DPI OBBLIGATORI</h4>${buildAtexList(["Scarpe antistatiche", "Guanti idonei", "Visiera/protezione occhi", "Alta visibilità", "DPI previsti dalla commessa"])}</section>
          <section class="atex-rule-box">${buildAtexImageCard("alarm", "Allarme sicurezza", "atex-rule-image")}<h4>⚠️ COMPORTAMENTO OPERATIVO</h4>${buildAtexList(["Lavorare sempre in squadra", "Mantenere distanza sicurezza", "Controllare costantemente atmosfera", "Fermare attività in caso di dubbio", "Segnalare immediatamente anomalie"])}</section>
          <section class="atex-rule-box">${buildAtexImageCard("danger", "Zona pericolosa esterna", "atex-rule-image")}<h4>🌬 GAS E VENTILAZIONE</h4>${buildAtexList(["Controllare direzione vento", "Evitare ristagni gas", "Non entrare in pozzetti senza verifica", "Aerare zona prima attività"])}</section>
        </div>
      </article>
  
      ${buildAtexEmergencyContactsSection(impianto)}
  
      <article class="atex-procedure-section atex-status-section">
        <h3>7. STATO SICUREZZA</h3>
        <div class="atex-status-grid">
          <div class="atex-status-card is-green">🟢 <strong>SICUREZZA OPERATIVA</strong></div>
          <div class="atex-status-card is-yellow">🟡 <strong>ATTENZIONE</strong></div>
          <div class="atex-status-card is-red">🔴 <strong>RISCHIO ALTO</strong></div>
        </div>
        <div class="atex-indicators">${statusIndicators.map((item) => `<span>${escapeHTML(item)}</span>`).join("")}</div>
      </article>
  
      <article class="atex-procedure-section atex-form-section">
        <h3>8. MODULO ATEX</h3>
        <button class="btn btn-primary atex-open-form-btn" type="button" data-open-atex-form>📄 COMPILA MODULO ATEX</button>
        <form id="atex-module-form" class="atex-module-form hidden">
          <div class="atex-form-grid">
            <label>Operatore<input name="operatore" type="text" autocomplete="name" required></label>
            <label>Squadra<input name="squadra" type="text" required></label>
            <label>Impianto<input name="impianto" type="text" value="${escapeHTML(impiantoName)}" required></label>
            <label>Data<input name="data" type="date" value="${formatAtexDateValue()}" required></label>
            <label>Ora<input name="ora" type="time" value="${formatAtexTimeValue()}" required></label>
            <label>Presenza gas<select name="presenzaGas" required><option value="NO">NO</option><option value="SI">SI</option></select></label>
            <label>Altair verificato<select name="altairVerificato" required><option value="SI">SI</option><option value="NO">NO</option></select></label>
            <label>DPI verificati<select name="dpiVerificati" required><option value="SI">SI</option><option value="NO">NO</option></select></label>
            <label class="atex-form-wide">Note operative<textarea name="noteOperative" rows="4" placeholder="Annotazioni, anomalie, valori o comunicazioni operative"></textarea></label>
            <label class="atex-form-wide">Firma semplice<input name="firma" type="text" placeholder="Nome e cognome" required></label>
          </div>
          <button class="btn btn-primary" type="submit">Salva modulo ATEX</button>
          <p class="muted" data-atex-form-feedback role="status" aria-live="polite"></p>
        </form>
      </article>
    `;
  }
  api.renderAtexProcedurePage = renderAtexProcedurePage;
  Object.assign(global, api);
  global.VargaAtexModule = Object.freeze({ ...api });
})(window);
