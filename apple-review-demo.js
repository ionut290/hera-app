/* App Store review demonstration only: synthetic data in RAM, no networking,
 * Firebase, geolocation, persistent storage or production mutations. */
(() => {
  "use strict";
  const COMMESSE = Object.freeze([
    Object.freeze({ id: "DEMO-C01", name: "Verde Dimostrativo", category: "Manutenzione aree verdi", team: "Squadra Demo A" }),
    Object.freeze({ id: "DEMO-C02", name: "Impianti Dimostrativi", category: "Controlli e ispezioni", team: "Squadra Demo B" })
  ]);
  const PLANTS = Object.freeze([
    Object.freeze({ id: "DEMO-101", name: "Giardino di Esempio", commessaId: "DEMO-C01", activity: "Controllo verde e pulizia aiuole", status: "pending" }),
    Object.freeze({ id: "DEMO-102", name: "Parco Dimostrativo", commessaId: "DEMO-C01", activity: "Ispezione alberature", status: "pending" }),
    Object.freeze({ id: "DEMO-201", name: "Impianto di Prova", commessaId: "DEMO-C02", activity: "Verifica manutentiva periodica", status: "pending" }),
    Object.freeze({ id: "DEMO-202", name: "Area Tecnica Esempio", commessaId: "DEMO-C02", activity: "Controllo visivo", status: "done" })
  ]);
  // No actual user accounts, server endpoints, IDs or live company data.
  const state = {
    plants: PLANTS.map((plant) => ({ ...plant })),
    hours: [],
    selectedCommessa: "",
    expandedPlant: ""
  };
  const get = (id) => document.getElementById(id);
  function el(tag, className, textValue) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (textValue !== undefined) node.textContent = String(textValue);
    return node;
  }
  function append(parent, ...children) {
    for (const child of children) parent.appendChild(child);
    return parent;
  }
  function button(textValue, className, attrs = {}) {
    const result = el("button", className, textValue);
    result.type = "button";
    for (const [key, value] of Object.entries(attrs)) result.setAttribute(key, value);
    return result;
  }
  function showMessage(textValue) {
    get("demo-message").textContent = textValue;
  }
  function commessaById(id) {
    return COMMESSE.find((entry) => entry.id === id);
  }
  function setView(viewName) {
    const views = ["home", "commesse", "impianti", "squadre", "ore", "documenti"];
    if (!views.includes(viewName)) return;
    for (const view of views) {
      const section = get("view-" + view);
      const visible = view === viewName;
      section.hidden = !visible;
      section.classList.toggle("active", visible);
    }
    document.querySelectorAll(".tab[data-view]").forEach((tab) => {
      const active = tab.dataset.view === viewName;
      tab.classList.toggle("selected", active);
      tab.setAttribute("aria-pressed", String(active));
    });
    if (viewName === "impianti") renderPlants();
    if (viewName === "commesse") renderCommesse();
    if (viewName === "ore") renderHours();
    showMessage("");
    const heading = get("view-" + viewName).querySelector("h2");
    if (heading) {
      heading.setAttribute("tabindex", "-1");
      heading.focus({ preventScroll: true });
    }
    window.scrollTo?.({ top: 0, behavior: "instant" });
  }
  function renderMetrics() {
    get("metric-commesse").textContent = String(COMMESSE.length);
    get("metric-pending").textContent = String(state.plants.filter((p) => p.status === "pending").length);
    get("metric-squadre").textContent = "2";
  }
  function renderCommesse() {
    const root = get("commesse-list");
    root.replaceChildren();
    for (const job of COMMESSE) {
      const article = el("article", "panel");
      const head = el("div", "card-meta");
      append(head, el("strong", "code-pill", job.id), el("span", "pill-info", job.category));
      const total = state.plants.filter((plant) => plant.commessaId === job.id);
      const outstanding = total.filter((plant) => plant.status === "pending").length;
      append(article, head, el("h3", "", "📁 " + job.name),
        el("p", "", job.team + " · " + total.length + " impianti · " + outstanding + " da fare"));
      append(article, button("Visualizza impianti", "primary-button", {
        "data-action": "open-commessa", "data-id": job.id
      }));
      root.appendChild(article);
    }
  }
  function renderPlants() {
    const root = get("plants-list");
    root.replaceChildren();
    const query = get("plant-query").value.trim().toLocaleLowerCase("it");
    const status = get("plant-status").value;
    const matches = state.plants.filter((plant) =>
      (!state.selectedCommessa || plant.commessaId === state.selectedCommessa) &&
      (status === "all" || plant.status === status) &&
      (plant.name.toLocaleLowerCase("it").includes(query) || plant.id.toLocaleLowerCase("it").includes(query)));
    const filterText = state.selectedCommessa && commessaById(state.selectedCommessa)?.name;
    const heading = get("impianti-title");
    heading.textContent = filterText ? "Impianti · " + filterText : "Impianti di prova";
    for (const plant of matches) {
      const article = el("article", "panel");
      const top = el("div", "card-meta");
      append(top, el("strong", "code-pill", plant.id),
        el("span", "state-pill" + (plant.status === "done" ? " done" : ""),
          plant.status === "done" ? "Completato (demo)" : "Da fare (demo)"));
      append(article, top, el("h3", "", "📍 " + plant.name),
        el("p", "", commessaById(plant.commessaId)?.name || "Commessa demo"),
        el("p", "", "Attività: " + plant.activity));
      const actions = el("div", "plant-actions");
      append(actions,
        button("Dettagli", "secondary-button", { "data-action": "details", "data-id": plant.id }),
        button(plant.status === "pending" ? "Simula intervento" : "Ripristina stato demo",
          "primary-button", { "data-action": "simulate", "data-id": plant.id }));
      article.appendChild(actions);
      if (state.expandedPlant === plant.id) {
        const details = el("div", "plant-details");
        append(details,
          el("p", "", "Ubicazione: posizione inventata, nessuna coordinata GPS."),
          el("p", "", "Responsabile: Addetto dimostrativo."),
          el("p", "", "Documentazione: schede e immagini fittizie."));
        article.appendChild(details);
      }
      root.appendChild(article);
    }
    get("plants-empty").hidden = matches.length !== 0;
  }
  function renderHours() {
    const root = get("hours-list");
    root.replaceChildren();
    let totalMinutes = 0;
    for (const item of state.hours) {
      totalMinutes += item.minutes;
      root.appendChild(el("li", "", item.date + " · " + item.start + "–" + item.end + " · " +
        (item.minutes / 60).toLocaleString("it-IT", { maximumFractionDigits: 2 }) + " ore (demo)"));
    }
    if (state.hours.length === 0) root.appendChild(el("li", "", "Nessun turno aggiunto in questa sessione."));
    get("hours-total").textContent = (totalMinutes / 60).toLocaleString("it-IT", { maximumFractionDigits: 2 }) + " h";
  }
  function resetDemo() {
    state.plants = PLANTS.map((item) => ({ ...item }));
    state.hours = [];
    state.selectedCommessa = "";
    state.expandedPlant = "";
    get("plant-query").value = "";
    get("plant-status").value = "all";
    get("doc-preview").hidden = true;
    renderMetrics();
    renderCommesse();
    renderPlants();
    renderHours();
    showMessage("Demo ripristinata: nessun dato reale è stato modificato.");
  }
  document.querySelectorAll("button[data-view],button[data-go]").forEach((item) => {
    item.addEventListener("click", () => setView(item.dataset.view || item.dataset.go));
  });
  document.querySelectorAll(".reset-demo").forEach((item) => {
    item.addEventListener("click", resetDemo);
  });
  get("commesse-list").addEventListener("click", (event) => {
    const target = event.target.closest("button[data-action='open-commessa']");
    if (!target) return;
    if (!commessaById(target.dataset.id)) return;
    state.selectedCommessa = target.dataset.id;
    get("plant-query").value = "";
    get("plant-status").value = "all";
    setView("impianti");
  });
  get("plants-list").addEventListener("click", (event) => {
    const target = event.target.closest("button[data-action]");
    if (!target) return;
    const plant = state.plants.find((item) => item.id === target.dataset.id);
    if (!plant) return;
    if (target.dataset.action === "details") {
      state.expandedPlant = state.expandedPlant === plant.id ? "" : plant.id;
    } else if (target.dataset.action === "simulate") {
      plant.status = plant.status === "pending" ? "done" : "pending";
      showMessage("Intervento soltanto simulato su " + plant.id + ". Nessun invio al sistema reale.");
      renderMetrics();
    }
    renderPlants();
  });
  get("plant-query").addEventListener("input", () => { state.selectedCommessa = ""; renderPlants(); });
  get("plant-status").addEventListener("change", renderPlants);
  get("demo-hours-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const date = get("hours-date").value;
    const start = get("hours-start").value;
    const end = get("hours-end").value;
    const pause = Number(get("hours-break").value);
    const toMinutes = (time) => {
      const parts = time.split(":").map(Number);
      return parts.length === 2 ? parts[0] * 60 + parts[1] : NaN;
    };
    const minutes = toMinutes(end) - toMinutes(start) - pause;
    if (!date || !Number.isFinite(minutes) || minutes <= 0 || minutes > 1440) {
      showMessage("Controlla gli orari: la fine deve seguire l'inizio e superare la pausa.");
      return;
    }
    state.hours.push({ date, start, end, minutes });
    renderHours();
    showMessage("Turno fittizio aggiunto soltanto alla sessione dimostrativa.");
  });
  document.querySelectorAll("button[data-doc]").forEach((item) => {
    item.addEventListener("click", () => {
      get("doc-title").textContent = item.dataset.doc + " — esempio";
      get("doc-preview").hidden = false;
      showMessage("Documento di esempio. Nessun file reale è stato aperto.");
    });
  });
  renderMetrics();
  renderCommesse();
  renderPlants();
  renderHours();
})();