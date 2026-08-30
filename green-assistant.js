(() => {
  "use strict";

  if (window.HeraGreenAssistant?.installed) return;

  const ENDPOINT = "/api/green-assistant";
  const CACHE_KEY = "heraGreenAssistantArchiveV1";
  const MAX_ARCHIVE_ITEMS = 40;
  const state = {
    mounted: false,
    mode: "gardening",
    statusLoaded: false,
    configured: { plantnet: false, trefle: false, gemini: false, brave: false },
    plantImage: "",
    diseaseImage: "",
    equipmentImage: "",
    lastEquipment: null
  };

  const escapeHTML = (value) => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

  const safeUrl = (value) => {
    try {
      const url = new URL(String(value || ""), window.location.origin);
      return ["http:", "https:"].includes(url.protocol) ? url.href : "";
    } catch (_) {
      return "";
    }
  };

  function template() {
    return `
      <section id="green-assistant-overlay" class="green-assistant-overlay hidden" aria-hidden="true">
        <header class="green-assistant-header">
          <button id="green-assistant-close" class="green-assistant-back" type="button" aria-label="Chiudi assistente">←</button>
          <div>
            <p class="green-assistant-kicker">VARGA CANTIERI</p>
            <h1 id="green-assistant-title">Assistente Giardiniere</h1>
          </div>
          <span class="green-assistant-free-badge" title="Le chiavi API restano protette sul server">API PROTETTE</span>
        </header>

        <div class="green-assistant-scroll">
          <div id="green-assistant-provider-status" class="green-assistant-provider-status" role="status" aria-live="polite">
            Verifico i servizi disponibili…
          </div>

          <section id="gardening-assistant-view" class="green-assistant-view">
            <div class="green-assistant-tabs" role="tablist" aria-label="Funzioni giardinaggio">
              <button class="is-active" type="button" data-green-tab="identify">📷 Identifica</button>
              <button type="button" data-green-tab="search">🔎 Cerca</button>
              <button type="button" data-green-tab="disease">🦠 Malattia</button>
              <button type="button" data-green-tab="archive">🌳 Salvate</button>
            </div>

            <section class="green-assistant-panel" data-green-panel="identify">
              <div class="green-assistant-card green-assistant-hero-card">
                <span class="green-assistant-card-icon">🌿</span>
                <div><h2>Identifica una pianta</h2><p>Fotografa foglia, fiore, frutto o corteccia. L’operatore conferma sempre il risultato.</p></div>
              </div>
              <form id="green-identify-form" class="green-assistant-form">
                <label class="green-assistant-file">
                  <span>📷 Fotografia della pianta</span>
                  <input id="green-identify-image" type="file" accept="image/jpeg,image/png" capture="environment" required>
                </label>
                <img id="green-identify-preview" class="green-assistant-preview hidden" alt="Anteprima pianta">
                <label>Parte fotografata
                  <select id="green-identify-organ">
                    <option value="auto">Riconoscimento automatico</option>
                    <option value="leaf">Foglia</option>
                    <option value="flower">Fiore</option>
                    <option value="fruit">Frutto</option>
                    <option value="bark">Corteccia</option>
                  </select>
                </label>
                <button class="btn btn-primary green-assistant-submit" type="submit">IDENTIFICA PIANTA</button>
              </form>
              <div id="green-identify-result" class="green-assistant-result" aria-live="polite"></div>
            </section>

            <section class="green-assistant-panel hidden" data-green-panel="search">
              <div class="green-assistant-card green-assistant-hero-card">
                <span class="green-assistant-card-icon">📚</span>
                <div><h2>Cerca una scheda botanica</h2><p>Cerca per nome comune o scientifico nei dati gratuiti Trefle.</p></div>
              </div>
              <form id="green-search-form" class="green-assistant-form green-assistant-search-form">
                <label>Nome della pianta
                  <input id="green-search-query" type="search" minlength="2" maxlength="180" placeholder="Esempio: Platanus acerifolia" required>
                </label>
                <button class="btn btn-primary green-assistant-submit" type="submit">CERCA PIANTA</button>
              </form>
              <div id="green-search-result" class="green-assistant-result" aria-live="polite"></div>
              <div id="green-plant-detail" class="green-assistant-result" aria-live="polite"></div>
            </section>

            <section class="green-assistant-panel hidden" data-green-panel="disease">
              <div class="green-assistant-card green-assistant-hero-card">
                <span class="green-assistant-card-icon">🦠</span>
                <div><h2>Analizza malattia o parassita</h2><p>Inquadra bene la parte danneggiata. La copertura gratuita non comprende tutte le patologie.</p></div>
              </div>
              <form id="green-disease-form" class="green-assistant-form">
                <label class="green-assistant-file">
                  <span>📷 Fotografia della parte danneggiata</span>
                  <input id="green-disease-image" type="file" accept="image/jpeg,image/png" capture="environment" required>
                </label>
                <img id="green-disease-preview" class="green-assistant-preview hidden" alt="Anteprima parte danneggiata">
                <label>Parte fotografata
                  <select id="green-disease-organ">
                    <option value="auto">Riconoscimento automatico</option>
                    <option value="leaf">Foglia</option>
                    <option value="flower">Fiore</option>
                    <option value="fruit">Frutto</option>
                    <option value="bark">Corteccia</option>
                  </select>
                </label>
                <button class="btn btn-primary green-assistant-submit" type="submit">ANALIZZA PROBLEMA</button>
              </form>
              <div id="green-disease-result" class="green-assistant-result" aria-live="polite"></div>
            </section>

            <section class="green-assistant-panel hidden" data-green-panel="archive">
              <div class="green-assistant-card green-assistant-hero-card">
                <span class="green-assistant-card-icon">🌳</span>
                <div><h2>Piante salvate sul dispositivo</h2><p>Le schede già consultate si riaprono senza consumare altre richieste API.</p></div>
              </div>
              <div id="green-plant-archive" class="green-assistant-result"></div>
            </section>
          </section>

          <section id="equipment-assistant-view" class="green-assistant-view hidden">
            <div class="green-assistant-card green-assistant-hero-card green-assistant-equipment-hero">
              <span class="green-assistant-card-icon">🚜</span>
              <div><h2>Trova i dati del mezzo</h2><p>Inserisci marca e modello. I dati non confermati dal manuale saranno sempre indicati come da verificare.</p></div>
            </div>
            <form id="green-equipment-form" class="green-assistant-form green-equipment-form">
              <div class="green-assistant-registered-search">
                <label>Codice mezzo registrato
                  <input id="green-equipment-code" type="text" maxlength="120" placeholder="Esempio: R50" autocomplete="off" autocapitalize="characters" spellcheck="false">
                </label>
                <button id="green-equipment-load" class="btn" type="button">CARICA DATI</button>
              </div>
              <p id="green-equipment-record-status" class="green-assistant-note">Puoi cercare un mezzo già presente nell'app senza nuove letture Firestore.</p>
              <div class="green-assistant-form-grid">
                <label>Tipo mezzo o utensile
                  <input id="green-equipment-type" type="text" maxlength="120" placeholder="Motosega, trattore, decespugliatore…">
                </label>
                <label>Marca
                  <input id="green-equipment-brand" type="text" maxlength="120" placeholder="Esempio: STIHL" required>
                </label>
                <label>Modello
                  <input id="green-equipment-model" type="text" maxlength="160" placeholder="Esempio: FS 240" required>
                </label>
                <label>Anno, se conosciuto
                  <input id="green-equipment-year" type="text" maxlength="20" inputmode="numeric" placeholder="2024">
                </label>
                <label>Matricola, facoltativa
                  <input id="green-equipment-serial" type="text" maxlength="120" placeholder="Numero seriale">
                </label>
              </div>
              <label class="green-assistant-file">
                <span>📷 Foto targhetta, facoltativa</span>
                <input id="green-equipment-image" type="file" accept="image/jpeg,image/png" capture="environment">
              </label>
              <img id="green-equipment-preview" class="green-assistant-preview hidden" alt="Anteprima targhetta mezzo">
              <label>Testo del manuale ufficiale, facoltativo
                <textarea id="green-equipment-manual" rows="5" maxlength="24000" placeholder="Incolla qui la pagina del manuale con carburante, olio, manutenzione o ricambi."></textarea>
              </label>
              <p class="green-assistant-note">🔒 La chiave Gemini resta protetta su Netlify. Senza manuale, miscela, olio, capacità e ricambi restano “da verificare”.</p>
              <button class="btn btn-primary green-assistant-submit" type="submit">TROVA DATI DEL MEZZO</button>
            </form>
            <div id="green-equipment-result" class="green-assistant-result" aria-live="polite"></div>
            <div id="green-equipment-manual-results" class="green-assistant-result" aria-live="polite"></div>

            <div class="green-assistant-card green-assistant-archive-card">
              <div class="green-assistant-section-head"><h2>📋 Schede recenti</h2><button id="green-equipment-archive-clear" class="btn" type="button">Svuota</button></div>
              <div id="green-equipment-archive" class="green-assistant-result"></div>
            </div>
          </section>
        </div>
      </section>`;
  }

  function mount() {
    if (state.mounted) return;
    document.body.insertAdjacentHTML("beforeend", template());
    state.mounted = true;
    bindEvents();
    renderArchives();
  }

  function overlay() { return document.getElementById("green-assistant-overlay"); }

  function open(mode, equipment = null) {
    mount();
    state.mode = mode === "equipment" ? "equipment" : "gardening";
    document.getElementById("gardening-assistant-view")?.classList.toggle("hidden", state.mode !== "gardening");
    document.getElementById("equipment-assistant-view")?.classList.toggle("hidden", state.mode !== "equipment");
    document.getElementById("green-assistant-title").textContent = state.mode === "equipment"
      ? "Assistente Mezzi e Utensili"
      : "Assistente Giardiniere";
    overlay()?.classList.remove("hidden");
    overlay()?.setAttribute("aria-hidden", "false");
    document.body.classList.add("green-assistant-open");
    document.getElementById("menu-close-btn")?.click();
    if (!state.statusLoaded) void loadStatus();
    if (state.mode === "equipment") {
      renderEquipmentArchive();
      if (equipment) fillEquipmentForm(equipment);
    } else renderPlantArchive();
    setTimeout(() => document.getElementById("green-assistant-close")?.focus(), 0);
  }

  function close() {
    overlay()?.classList.add("hidden");
    overlay()?.setAttribute("aria-hidden", "true");
    document.body.classList.remove("green-assistant-open");
  }

  async function currentToken() {
    const user = window.firebase?.auth?.().currentUser;
    if (!user) throw new Error("Accedi all’app per utilizzare l’assistente.");
    return user.getIdToken();
  }

  async function api(action, payload = {}) {
    if (navigator.onLine === false) throw new Error("Questa ricerca richiede una connessione Internet.");
    const token = await currentToken();
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 50000);
    try {
      const response = await fetch(ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ action, ...payload }),
        signal: controller.signal
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.ok) throw new Error(data.error || "Assistente temporaneamente non disponibile.");
      return data.result ?? data;
    } catch (error) {
      if (error?.name === "AbortError") throw new Error("Il servizio sta impiegando troppo tempo. Riprova.");
      throw error;
    } finally {
      clearTimeout(timeout);
    }
  }

  async function loadStatus() {
    const element = document.getElementById("green-assistant-provider-status");
    try {
      const data = await api("status");
      state.configured = data.configured || state.configured;
      state.statusLoaded = true;
      const rows = [
        ["Pl@ntNet", state.configured.plantnet],
        ["Trefle", state.configured.trefle],
        ["Gemini", state.configured.gemini],
        ["Brave manuali", state.configured.brave]
      ];
      element.innerHTML = `${rows.map(([name, ready]) => `<span class="${ready ? "is-ready" : "is-missing"}">${ready ? "✓" : "!"} ${escapeHTML(name)}</span>`).join("")}${data.notice ? `<small>${escapeHTML(data.notice)}</small>` : ""}`;
    } catch (error) {
      element.textContent = error.message;
      element.classList.add("has-error");
    }
  }

  function setBusy(form, busy, busyText) {
    const button = form?.querySelector("button[type='submit']");
    if (!button) return;
    if (!button.dataset.defaultText) button.dataset.defaultText = button.textContent;
    button.disabled = busy;
    button.textContent = busy ? busyText : button.dataset.defaultText;
  }

  function showMessage(targetId, message, type = "info") {
    const target = document.getElementById(targetId);
    if (!target) return;
    target.innerHTML = `<div class="green-assistant-message is-${escapeHTML(type)}">${escapeHTML(message)}</div>`;
  }

  async function fileToCompressedDataUrl(file) {
    if (!file || !/^image\/(jpeg|png)$/i.test(file.type)) throw new Error("Seleziona una fotografia JPG o PNG.");
    if (file.size > 15 * 1024 * 1024) throw new Error("La fotografia originale è troppo grande.");
    const objectUrl = URL.createObjectURL(file);
    try {
      const image = await new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error("Impossibile leggere la fotografia."));
        img.src = objectUrl;
      });
      const maxSide = 1600;
      const scale = Math.min(1, maxSide / Math.max(image.naturalWidth || image.width, image.naturalHeight || image.height));
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round((image.naturalWidth || image.width) * scale));
      canvas.height = Math.max(1, Math.round((image.naturalHeight || image.height) * scale));
      canvas.getContext("2d", { alpha: false }).drawImage(image, 0, 0, canvas.width, canvas.height);
      return canvas.toDataURL("image/jpeg", 0.82);
    } finally {
      URL.revokeObjectURL(objectUrl);
    }
  }

  async function handleImageInput(input, stateKey, previewId) {
    const file = input.files?.[0];
    if (!file) return;
    try {
      const dataUrl = await fileToCompressedDataUrl(file);
      state[stateKey] = dataUrl;
      const preview = document.getElementById(previewId);
      preview.src = dataUrl;
      preview.classList.remove("hidden");
    } catch (error) {
      input.value = "";
      state[stateKey] = "";
      window.alert(error.message);
    }
  }

  function score(value) {
    return `${Math.round(Math.max(0, Math.min(1, Number(value || 0))) * 100)}%`;
  }

  function imageMarkup(url, alt) {
    const safe = safeUrl(url);
    return safe ? `<img class="green-assistant-result-thumb" src="${escapeHTML(safe)}" alt="${escapeHTML(alt)}" loading="lazy">` : "";
  }

  function renderPlantIdentification(data) {
    const target = document.getElementById("green-identify-result");
    const results = Array.isArray(data.results) ? data.results : [];
    if (!results.length) {
      showMessage("green-identify-result", "La fotografia non ha prodotto un risultato affidabile. Prova con più luce e avvicinati alla foglia o al fiore.", "warning");
      return;
    }
    target.innerHTML = `
      <div class="green-assistant-result-head"><strong>Possibili specie</strong><span>${data.remainingRequests == null ? "Quota gratuita" : `${data.remainingRequests} richieste rimaste oggi`}</span></div>
      ${results.map((item, index) => `
        <article class="green-assistant-result-card ${index === 0 ? "is-best" : ""}">
          ${imageMarkup(item.image, item.scientificName)}
          <div class="green-assistant-result-main">
            <span class="green-assistant-confidence">${score(item.score)}</span>
            <h3>${escapeHTML(item.commonNames?.[0] || item.scientificName || "Specie possibile")}</h3>
            <p><em>${escapeHTML(item.scientificName)}</em>${item.family ? ` · ${escapeHTML(item.family)}` : ""}</p>
            <div class="green-assistant-result-actions">
              <button class="btn btn-primary" type="button" data-confirm-plant="${escapeHTML(item.scientificName)}">CONFERMA</button>
              <button class="btn" type="button" data-search-plant-name="${escapeHTML(item.scientificName)}">VEDI SCHEDA</button>
            </div>
          </div>
        </article>`).join("")}
      <p class="green-assistant-attribution">Riconoscimento basato su Pl@ntNet. Il risultato deve essere confermato dall’operatore.</p>`;
  }

  function renderDisease(data) {
    const target = document.getElementById("green-disease-result");
    const results = Array.isArray(data.results) ? data.results : [];
    if (!results.length) {
      showMessage("green-disease-result", "Nessuna patologia riconosciuta. La versione gratuita copre soltanto alcune specie e patologie.", "warning");
      return;
    }
    target.innerHTML = `
      <div class="green-assistant-message is-warning">Queste sono possibilità, non una diagnosi certa. Verifica prima di effettuare trattamenti.</div>
      <div class="green-assistant-result-head"><strong>Possibili problemi</strong><span>${data.remainingRequests == null ? "Quota gratuita" : `${data.remainingRequests} richieste rimaste oggi`}</span></div>
      ${results.map((item) => `<article class="green-assistant-result-card">
        ${imageMarkup(item.image, item.name)}
        <div class="green-assistant-result-main"><span class="green-assistant-confidence">${score(item.score)}</span><h3>${escapeHTML(item.name)}</h3><p>Codice EPPO: ${escapeHTML(item.code || "n/d")}</p></div>
      </article>`).join("")}
      <p class="green-assistant-attribution">Analisi basata su Pl@ntNet Diseases; copertura limitata.</p>`;
  }

  function renderPlantSearch(data) {
    const target = document.getElementById("green-search-result");
    const results = Array.isArray(data.results) ? data.results : [];
    if (!results.length) {
      showMessage("green-search-result", "Nessuna pianta trovata con questo nome.", "warning");
      return;
    }
    target.innerHTML = results.map((item) => `<button class="green-assistant-search-row" type="button" data-plant-slug="${escapeHTML(item.slug)}">
      ${imageMarkup(item.image, item.scientificName)}
      <span><strong>${escapeHTML(item.commonName || item.scientificName)}</strong><small>${escapeHTML(item.scientificName)}${item.family ? ` · ${escapeHTML(item.family)}` : ""}</small></span><span aria-hidden="true">›</span>
    </button>`).join("") + `<p class="green-assistant-attribution">Dati botanici: Trefle e fonti indicate nelle singole schede.</p>`;
  }

  function displayValue(value) {
    if (value == null || value === "") return "Non disponibile";
    if (Array.isArray(value)) return value.filter(Boolean).join(", ") || "Non disponibile";
    if (typeof value === "object") {
      if (value.cm != null) return `${value.cm} cm`;
      if (value.celsius != null) return `${value.celsius} °C`;
      if (value.mm != null) return `${value.mm} mm`;
      return Object.entries(value).filter(([, item]) => item != null && typeof item !== "object").map(([key, item]) => `${key}: ${item}`).join(" · ") || "Non disponibile";
    }
    if (typeof value === "boolean") return value ? "Sì" : "No";
    return String(value);
  }

  function renderPlantDetails(data, { save = true } = {}) {
    const plant = data?.plant || data;
    const target = document.getElementById("green-plant-detail");
    if (!plant?.scientificName) {
      showMessage("green-plant-detail", "Scheda botanica non disponibile.", "warning");
      return;
    }
    const growth = plant.growth || {};
    const specs = plant.specifications || {};
    const fields = [
      ["Famiglia", plant.family],
      ["Durata", plant.duration],
      ["Forma", specs.ligneous_type || specs.growth_habit],
      ["Velocità di crescita", specs.growth_rate],
      ["Altezza media", specs.average_height],
      ["Altezza massima", specs.maximum_height],
      ["Tossicità", specs.toxicity],
      ["Luce richiesta (0–10)", growth.light],
      ["Umidità terreno (0–10)", growth.soil_humidity],
      ["pH minimo", growth.ph_minimum],
      ["pH massimo", growth.ph_maximum],
      ["Temperatura minima", growth.minimum_temperature],
      ["Temperatura massima", growth.maximum_temperature],
      ["Mesi di crescita", growth.growth_months],
      ["Mesi di fioritura", growth.bloom_months],
      ["Profondità minima radici", growth.minimum_root_depth]
    ].filter(([, value]) => value != null && value !== "" && (!Array.isArray(value) || value.length));
    target.innerHTML = `<article class="green-assistant-plant-sheet">
      <div class="green-assistant-plant-sheet-head">${imageMarkup(plant.image, plant.scientificName)}<div><span class="green-assistant-verified">DATI BOTANICI</span><h2>${escapeHTML(plant.commonName || plant.scientificName)}</h2><p><em>${escapeHTML(plant.scientificName)}</em></p></div></div>
      <dl>${fields.map(([label, value]) => `<div><dt>${escapeHTML(label)}</dt><dd>${escapeHTML(displayValue(value))}</dd></div>`).join("")}</dl>
      ${plant.observations ? `<p class="green-assistant-observations">${escapeHTML(plant.observations)}</p>` : ""}
      <div class="green-assistant-result-actions"><button class="btn btn-primary" type="button" data-save-plant-sheet>CONFERMA E SALVA</button></div>
      <p class="green-assistant-attribution">Fonte: Trefle. Verificare le attribuzioni delle fonti originali nella risposta del servizio.</p>
    </article>`;
    target.querySelector("[data-save-plant-sheet]")?.addEventListener("click", () => {
      saveArchive("plants", plant, plant.scientificName);
      renderPlantArchive();
      showMessage("green-plant-detail", `${plant.commonName || plant.scientificName} salvata sul dispositivo.`, "success");
      setTimeout(() => renderPlantDetails({ plant }, { save: false }), 1100);
    });
    if (save) saveArchive("plantCache", plant, plant.scientificName);
  }

  function statusLabel(status) {
    if (status === "confermato_manual") return ["🟢", "Confermato dal manuale", "is-confirmed"];
    if (status === "non_disponibile") return ["🔴", "Non disponibile", "is-unavailable"];
    return ["🟡", "Da verificare", "is-check"];
  }

  function equipmentRecords() {
    return typeof mezziRecords !== "undefined" && Array.isArray(mezziRecords) ? mezziRecords : [];
  }

  function equipmentCode(record) {
    return String(record?.codiceMezzo || record?.nId || record?.nome || record?.id || "").trim();
  }

  function findEquipmentRecord(value) {
    const key = String(value || "").trim().toLocaleUpperCase("it");
    if (!key) return null;
    return equipmentRecords().find((record) => [record.id, record.codiceMezzo, record.nId, record.nome, record.targa]
      .some((candidate) => String(candidate || "").trim().toLocaleUpperCase("it") === key)) || null;
  }

  function fillEquipmentForm(input) {
    const record = typeof input === "string" ? findEquipmentRecord(input) : input;
    const status = document.getElementById("green-equipment-record-status");
    if (!record || typeof record !== "object") {
      if (status) status.textContent = "Mezzo non trovato. Controlla il codice oppure selezionalo dall'elenco.";
      return false;
    }
    const values = {
      "green-equipment-code": equipmentCode(record),
      "green-equipment-type": record.categoria || record.tipo || "",
      "green-equipment-brand": record.marca || record.brand || "",
      "green-equipment-model": record.modello || record.model || "",
      "green-equipment-year": record.anno || record.annoImmatricolazione || record.annoAcquisto || "",
      "green-equipment-serial": record.matricola || record.seriale || record.numeroTelaio || ""
    };
    Object.entries(values).forEach(([id, value]) => {
      const field = document.getElementById(id);
      if (field) field.value = String(value || "");
    });
    if (status) {
      const missing = [values["green-equipment-brand"] ? "" : "marca", values["green-equipment-model"] ? "" : "modello"].filter(Boolean);
      status.textContent = missing.length
        ? `${equipmentCode(record)} caricato. Completa ${missing.join(" e ")} per cercare i manuali.`
        : `${equipmentCode(record)} caricato dai Mezzi già disponibili sul dispositivo.`;
    }
    return true;
  }

  function manualCacheKey(payload) {
    return [payload.brand, payload.model, payload.year, payload.type].map((value) => String(value || "").trim().toLowerCase()).join("|");
  }

  function cachedManualSearch(payload) {
    const key = manualCacheKey(payload);
    return (readArchive().manualSearchCache || []).find((item) => item.key === key)?.value || null;
  }

  function renderManualSearch(data, cached = false) {
    const target = document.getElementById("green-equipment-manual-results");
    if (!target) return;
    const results = Array.isArray(data?.results) ? data.results : [];
    target.innerHTML = `<article class="green-assistant-card green-assistant-manual-card">
      <div class="green-assistant-section-head"><h2>📚 Manuali e schede tecniche</h2><span class="green-assistant-verified">${cached ? "SALVATI" : "BRAVE"}</span></div>
      ${results.length ? results.map((item) => {
        const url = safeUrl(item.url);
        if (!url) return "";
        return `<a class="green-assistant-manual-row" href="${escapeHTML(url)}" target="_blank" rel="noopener noreferrer">
          <span><strong>${escapeHTML(item.title || "Documento tecnico")}</strong><small>${escapeHTML(item.domain || "Fonte web")}${item.description ? ` · ${escapeHTML(item.description)}` : ""}</small></span>
          <span class="green-assistant-data-status ${item.likelyOfficial ? "is-confirmed" : "is-check"}">${item.likelyOfficial ? "Produttore" : "Da verificare"}</span>
        </a>`;
      }).join("") : `<p class="green-assistant-empty-state">Nessun manuale trovato per marca e modello indicati.</p>`}
      <p class="green-assistant-attribution">I risultati del produttore vengono mostrati per primi. Verifica sempre che modello, variante e matricola coincidano prima di usare i dati tecnici.</p>
    </article>`;
  }

  async function searchEquipmentManuals(payload) {
    const cached = cachedManualSearch(payload);
    if (cached) return renderManualSearch(cached, true);
    showMessage("green-equipment-manual-results", "Cerco manuali e schede tecniche con Brave…");
    try {
      const data = await api("equipmentManuals", payload);
      saveArchive("manualSearchCache", data, manualCacheKey(payload));
      renderManualSearch(data);
    } catch (error) {
      showMessage("green-equipment-manual-results", error.message, "error");
    }
  }

  function renderEquipment(data, { save = true } = {}) {
    state.lastEquipment = data;
    const target = document.getElementById("green-equipment-result");
    const id = data.identification || {};
    const technical = Array.isArray(data.technicalData) ? data.technicalData : [];
    target.innerHTML = `<article class="green-assistant-equipment-sheet">
      <div class="green-assistant-equipment-sheet-head"><span class="green-assistant-big-icon">🚜</span><div><span class="green-assistant-verified">SCHEDA GENERATA</span><h2>${escapeHTML([id.brand, id.model].filter(Boolean).join(" ") || "Mezzo o utensile")}</h2><p>${escapeHTML(id.type || "Tipologia da verificare")}</p></div></div>
      ${id.summary ? `<p class="green-assistant-summary">${escapeHTML(id.summary)}</p>` : ""}
      <div class="green-assistant-technical-list">${technical.map((item) => {
        const [icon, label, className] = statusLabel(item.status);
        return `<div class="green-assistant-technical-row"><div><strong>${escapeHTML(item.label)}</strong><span>${escapeHTML(item.value || "Non disponibile")}</span>${item.sourceNote ? `<small>${escapeHTML(item.sourceNote)}</small>` : ""}</div><span class="green-assistant-data-status ${className}">${icon} ${label}</span></div>`;
      }).join("") || `<p class="muted">Nessun dato tecnico affidabile disponibile.</p>`}</div>
      ${listSection("🔧 Manutenzione generale", data.maintenance)}
      ${listSection("🦺 Sicurezza", data.safety)}
      ${listSection("✅ Controlli prima dell’uso", data.commonChecks)}
      ${listSection("❓ Dati mancanti", data.missingInformation)}
      <div class="green-assistant-message is-warning">${escapeHTML(data.warning || "Verificare sempre il manuale ufficiale.")}</div>
      <div class="green-assistant-result-actions"><button class="btn btn-primary" type="button" data-copy-equipment>COPIA NEL FORM MEZZI</button><button class="btn secondary" type="button" data-save-equipment>SALVA SCHEDA</button></div>
      <p class="green-assistant-attribution">Generato con Gemini (${escapeHTML(data.model || "modello gratuito")}). L’AI non sostituisce il manuale ufficiale.</p>
    </article>`;
    target.querySelector("[data-copy-equipment]")?.addEventListener("click", copyEquipmentToExistingForm);
    target.querySelector("[data-save-equipment]")?.addEventListener("click", () => {
      const key = `${id.brand || "mezzo"}-${id.model || Date.now()}`;
      saveArchive("equipment", data, key);
      renderEquipmentArchive();
      target.querySelector("[data-save-equipment]").textContent = "SALVATA ✓";
    });
    if (save) {
      const key = `${id.brand || "mezzo"}-${id.model || Date.now()}`;
      saveArchive("equipmentCache", data, key);
    }
  }

  function listSection(title, items) {
    const rows = Array.isArray(items) ? items.filter(Boolean) : [];
    return rows.length ? `<section class="green-assistant-sheet-section"><h3>${escapeHTML(title)}</h3><ul>${rows.map((item) => `<li>${escapeHTML(item)}</li>`).join("")}</ul></section>` : "";
  }

  function readArchive() {
    try {
      const value = JSON.parse(localStorage.getItem(CACHE_KEY) || "{}");
      return value && typeof value === "object" ? value : {};
    } catch (_) {
      return {};
    }
  }

  function saveArchive(group, value, key) {
    const archive = readArchive();
    const items = Array.isArray(archive[group]) ? archive[group] : [];
    const normalizedKey = String(key || "").trim().toLowerCase();
    archive[group] = [{ key: normalizedKey, savedAt: new Date().toISOString(), value }, ...items.filter((item) => item.key !== normalizedKey)].slice(0, MAX_ARCHIVE_ITEMS);
    try { localStorage.setItem(CACHE_KEY, JSON.stringify(archive)); } catch (_) {}
  }

  function renderPlantArchive() {
    const target = document.getElementById("green-plant-archive");
    if (!target) return;
    const items = (readArchive().plants || []).slice(0, MAX_ARCHIVE_ITEMS);
    target.innerHTML = items.length ? items.map((item, index) => `<button class="green-assistant-archive-row" type="button" data-open-saved-plant="${index}"><span>🌿</span><span><strong>${escapeHTML(item.value?.commonName || item.value?.scientificName || "Pianta")}</strong><small>${escapeHTML(item.value?.scientificName || "")}</small></span><span>›</span></button>`).join("") : `<p class="green-assistant-empty-state">Non hai ancora salvato schede di piante.</p>`;
  }

  function renderEquipmentArchive() {
    const target = document.getElementById("green-equipment-archive");
    if (!target) return;
    const items = (readArchive().equipment || []).slice(0, MAX_ARCHIVE_ITEMS);
    target.innerHTML = items.length ? items.map((item, index) => `<button class="green-assistant-archive-row" type="button" data-open-saved-equipment="${index}"><span>🚜</span><span><strong>${escapeHTML([item.value?.identification?.brand, item.value?.identification?.model].filter(Boolean).join(" ") || "Mezzo")}</strong><small>${escapeHTML(item.value?.identification?.type || "")}</small></span><span>›</span></button>`).join("") : `<p class="green-assistant-empty-state">Non hai ancora salvato schede di mezzi.</p>`;
  }

  function renderArchives() {
    renderPlantArchive();
    renderEquipmentArchive();
  }

  function switchTab(tab) {
    document.querySelectorAll("[data-green-tab]").forEach((button) => button.classList.toggle("is-active", button.dataset.greenTab === tab));
    document.querySelectorAll("[data-green-panel]").forEach((panel) => panel.classList.toggle("hidden", panel.dataset.greenPanel !== tab));
    if (tab === "archive") renderPlantArchive();
  }

  async function fetchPlantDetails(slug) {
    showMessage("green-plant-detail", "Caricamento scheda botanica…");
    try {
      renderPlantDetails(await api("plantDetails", { slug }));
      document.getElementById("green-plant-detail")?.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      showMessage("green-plant-detail", error.message, "error");
    }
  }

  async function searchByScientificName(name) {
    switchTab("search");
    const input = document.getElementById("green-search-query");
    input.value = name;
    showMessage("green-search-result", "Cerco la scheda gratuita…");
    try {
      const data = await api("searchPlant", { query: name });
      renderPlantSearch(data);
      const first = data.results?.find((item) => String(item.scientificName).toLowerCase() === String(name).toLowerCase()) || data.results?.[0];
      if (first?.slug) await fetchPlantDetails(first.slug);
    } catch (error) {
      showMessage("green-search-result", error.message, "error");
    }
  }

  function copyEquipmentToExistingForm() {
    const data = state.lastEquipment;
    if (!data) return;
    const id = data.identification || {};
    const fuel = (data.technicalData || []).find((item) => /carburante|alimentazione|batteria/i.test(item.label || ""))?.value || "";
    close();
    if (typeof window.openManagementPanel === "function") window.openManagementPanel("mezzi");
    const brand = document.getElementById("mezzo-marca");
    const model = document.getElementById("mezzo-modello");
    const power = document.getElementById("mezzo-alimentazione");
    if (brand && !brand.value) brand.value = id.brand || "";
    if (model && !model.value) model.value = id.model || "";
    if (power && !power.value) power.value = fuel || "";
    setTimeout(() => brand?.focus(), 50);
  }

  function bindEvents() {
    document.getElementById("green-assistant-close")?.addEventListener("click", close);
    overlay()?.addEventListener("click", (event) => { if (event.target === overlay()) close(); });
    document.querySelectorAll("[data-green-tab]").forEach((button) => button.addEventListener("click", () => switchTab(button.dataset.greenTab)));
    document.getElementById("green-identify-image")?.addEventListener("change", (event) => void handleImageInput(event.currentTarget, "plantImage", "green-identify-preview"));
    document.getElementById("green-disease-image")?.addEventListener("change", (event) => void handleImageInput(event.currentTarget, "diseaseImage", "green-disease-preview"));
    document.getElementById("green-equipment-image")?.addEventListener("change", (event) => void handleImageInput(event.currentTarget, "equipmentImage", "green-equipment-preview"));
    document.getElementById("green-equipment-load")?.addEventListener("click", () => {
      fillEquipmentForm(document.getElementById("green-equipment-code")?.value);
    });
    document.getElementById("green-equipment-code")?.addEventListener("change", (event) => {
      if (findEquipmentRecord(event.currentTarget.value)) fillEquipmentForm(event.currentTarget.value);
    });

    document.getElementById("green-identify-form")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!state.plantImage) return showMessage("green-identify-result", "Seleziona una fotografia della pianta.", "warning");
      setBusy(event.currentTarget, true, "ANALISI IN CORSO…");
      showMessage("green-identify-result", "Analizzo la fotografia con Pl@ntNet…");
      try {
        renderPlantIdentification(await api("identifyPlant", { image: state.plantImage, organ: document.getElementById("green-identify-organ").value }));
      } catch (error) { showMessage("green-identify-result", error.message, "error"); }
      finally { setBusy(event.currentTarget, false); }
    });

    document.getElementById("green-disease-form")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!state.diseaseImage) return showMessage("green-disease-result", "Seleziona una fotografia della parte danneggiata.", "warning");
      setBusy(event.currentTarget, true, "ANALISI IN CORSO…");
      showMessage("green-disease-result", "Confronto la fotografia con le patologie disponibili…");
      try {
        renderDisease(await api("identifyDisease", { image: state.diseaseImage, organ: document.getElementById("green-disease-organ").value }));
      } catch (error) { showMessage("green-disease-result", error.message, "error"); }
      finally { setBusy(event.currentTarget, false); }
    });

    document.getElementById("green-search-form")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      const query = document.getElementById("green-search-query").value.trim();
      setBusy(event.currentTarget, true, "RICERCA…");
      document.getElementById("green-plant-detail").innerHTML = "";
      showMessage("green-search-result", "Cerco nei dati botanici gratuiti…");
      try { renderPlantSearch(await api("searchPlant", { query })); }
      catch (error) { showMessage("green-search-result", error.message, "error"); }
      finally { setBusy(event.currentTarget, false); }
    });

    document.getElementById("green-equipment-form")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      const payload = {
        type: document.getElementById("green-equipment-type").value,
        brand: document.getElementById("green-equipment-brand").value,
        model: document.getElementById("green-equipment-model").value,
        year: document.getElementById("green-equipment-year").value,
        serial: document.getElementById("green-equipment-serial").value,
        manualText: document.getElementById("green-equipment-manual").value,
        image: state.equipmentImage || undefined
      };
      setBusy(event.currentTarget, true, "PREPARO LA SCHEDA…");
      showMessage("green-equipment-result", "Gemini sta preparando una scheda prudente…");
      const manualSearch = searchEquipmentManuals(payload);
      try { renderEquipment(await api("equipmentInfo", payload)); }
      catch (error) { showMessage("green-equipment-result", error.message, "error"); }
      finally { await manualSearch; setBusy(event.currentTarget, false); }
    });

    document.getElementById("green-equipment-archive-clear")?.addEventListener("click", () => {
      const archive = readArchive();
      archive.equipment = [];
      try { localStorage.setItem(CACHE_KEY, JSON.stringify(archive)); } catch (_) {}
      renderEquipmentArchive();
    });

    overlay()?.addEventListener("click", (event) => {
      const slugButton = event.target.closest("[data-plant-slug]");
      if (slugButton) return void fetchPlantDetails(slugButton.dataset.plantSlug);
      const searchButton = event.target.closest("[data-search-plant-name]");
      if (searchButton) return void searchByScientificName(searchButton.dataset.searchPlantName);
      const confirmButton = event.target.closest("[data-confirm-plant]");
      if (confirmButton) return void searchByScientificName(confirmButton.dataset.confirmPlant);
      const plantArchiveButton = event.target.closest("[data-open-saved-plant]");
      if (plantArchiveButton) {
        const item = (readArchive().plants || [])[Number(plantArchiveButton.dataset.openSavedPlant)];
        if (item?.value) { switchTab("search"); renderPlantDetails(item.value, { save: false }); }
        return;
      }
      const equipmentArchiveButton = event.target.closest("[data-open-saved-equipment]");
      if (equipmentArchiveButton) {
        const item = (readArchive().equipment || [])[Number(equipmentArchiveButton.dataset.openSavedEquipment)];
        if (item?.value) { renderEquipment(item.value, { save: false }); document.getElementById("green-equipment-result")?.scrollIntoView({ behavior: "smooth" }); }
      }
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    mount();
    document.getElementById("open-gardening-assistant-btn")?.addEventListener("click", () => open("gardening"));
    document.getElementById("open-equipment-assistant-btn")?.addEventListener("click", () => open("equipment"));
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !overlay()?.classList.contains("hidden")) close();
  });

  const openEquipment = (equipment) => open("equipment", equipment);
  window.HeraGreenAssistant = Object.freeze({ installed: true, version: "1.1.0", open, openEquipment, close });
})();
