(function () {
  "use strict";

  const STORAGE_KEY = "hera-control-center-open-category";
  const PAGE_ID = "control-center-page";
  const CONTENT_ID = "control-center-content";

  const PINNED_SECTIONS = [
    { key: "firestore-usage", match: ["consumo firestore giornaliero"] },
    { key: "firestore-diagnostics-v3", match: ["diagnostica operazioni firestore v3"] }
  ];

  const CATEGORIES = [
    { key: "report", title: "📥 Report e diagnostica", description: "Scarica report, diagnostica, backup ed esportazioni.", keywords: ["report", "scarica", "download", "diagnost", "esporta", "export", "backup", "controllo app"] },
    { key: "firestore", title: "🔥 Firestore e database", description: "Altri controlli Firestore, listener, snapshot e quota.", keywords: ["firestore", "firebase", "database", "listener", "snapshot", "quota"] },
    { key: "sync", title: "🔄 Sincronizzazione e rete", description: "Stato online/offline, sincronizzazione, cache e code dati.", keywords: ["sincron", "sync", "online", "offline", "rete", "connession", "cache", "coda"] },
    { key: "access", title: "🔐 Accesso e sicurezza", description: "Login, utenti, autorizzazioni, sessioni e credenziali.", keywords: ["login", "accesso", "utente", "auth", "google", "password", "session", "permess", "autorizz"] },
    { key: "app", title: "📱 App, PWA e aggiornamenti", description: "Versione app, PWA, Service Worker, installazione e aggiornamenti.", keywords: ["pwa", "versione", "service worker", "aggiorna app", "aggiornamento", "install", "build", "release"] },
    { key: "device", title: "📲 Dispositivo e ambiente", description: "Browser, Android/iOS, memoria, storage e informazioni dispositivo.", keywords: ["dispositivo", "device", "browser", "android", "ios", "iphone", "capacitor", "storage", "memoria"] },
    { key: "other", title: "🧰 Altri controlli", description: "Tutte le altre informazioni e verifiche del Centro di Controllo.", keywords: [] }
  ];

  function cleanText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function normalizedText(value) {
    return cleanText(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  function getPinnedKey(node) {
    const text = normalizedText(node?.textContent);
    const pinned = PINNED_SECTIONS.find((section) => section.match.some((phrase) => text.includes(normalizedText(phrase))));
    return pinned?.key || "";
  }

  function classifyNode(node) {
    const text = normalizedText(node.textContent);
    for (const category of CATEGORIES) {
      if (category.key === "other") continue;
      if (category.keywords.some((keyword) => text.includes(normalizedText(keyword)))) return category;
    }
    return CATEGORIES[CATEGORIES.length - 1];
  }

  function isManagedNode(node) {
    return !(node instanceof HTMLElement)
      || node.classList.contains("control-center-category")
      || node.classList.contains("control-center-tools")
      || node.classList.contains("control-center-empty-search")
      || node.id === "control-center-results";
  }

  function setOpen(item, open, options = {}) {
    const trigger = item.querySelector(":scope > .control-center-category-trigger");
    const panel = item.querySelector(":scope > .control-center-category-panel");
    if (!trigger || !panel) return;
    item.classList.toggle("is-open", open);
    trigger.setAttribute("aria-expanded", String(open));
    panel.hidden = !open;
    if (open) {
      try { sessionStorage.setItem(STORAGE_KEY, item.dataset.controlCategory || ""); } catch (_) {}
      if (options.scroll) requestAnimationFrame(() => item.scrollIntoView({ behavior: "smooth", block: "start" }));
    }
  }

  function closeOthers(container, current) {
    container.querySelectorAll(":scope > .control-center-category.is-open").forEach((item) => {
      if (item !== current) setOpen(item, false);
    });
  }

  function createCategory(category, container) {
    const item = document.createElement("section");
    const trigger = document.createElement("button");
    const heading = document.createElement("span");
    const title = document.createElement("span");
    const count = document.createElement("span");
    const description = document.createElement("span");
    const panel = document.createElement("div");
    const content = document.createElement("div");
    const panelId = `control-center-category-${category.key}`;

    item.className = "control-center-category";
    item.dataset.controlCategory = category.key;
    item.dataset.controlSearchText = normalizedText(`${category.title} ${category.description}`);
    trigger.type = "button";
    trigger.className = "control-center-category-trigger";
    trigger.setAttribute("aria-expanded", "false");
    trigger.setAttribute("aria-controls", panelId);

    heading.className = "control-center-category-heading";
    title.className = "control-center-category-title";
    title.textContent = category.title;
    count.className = "control-center-category-count";
    count.textContent = "0";
    heading.append(title, count);

    description.className = "control-center-category-description";
    description.textContent = category.description;
    trigger.append(heading, description);

    panel.className = "control-center-category-panel";
    panel.id = panelId;
    panel.hidden = true;
    content.className = "control-center-category-content";
    panel.appendChild(content);

    trigger.addEventListener("click", () => {
      const shouldOpen = !item.classList.contains("is-open");
      closeOthers(container, item);
      setOpen(item, shouldOpen, { scroll: shouldOpen });
    });

    item.append(trigger, panel);
    return item;
  }

  function ensureCategories(container) {
    const map = new Map();
    CATEGORIES.forEach((category) => {
      let item = container.querySelector(`:scope > .control-center-category[data-control-category="${category.key}"]`);
      if (!item) {
        item = createCategory(category, container);
        container.appendChild(item);
      }
      map.set(category.key, item);
    });
    return map;
  }

  function updateCategoryMeta(item) {
    const content = item.querySelector(":scope > .control-center-category-panel > .control-center-category-content");
    const count = item.querySelector(":scope > .control-center-category-trigger .control-center-category-count");
    if (!content || !count) return;
    const nodes = Array.from(content.children);
    count.textContent = String(nodes.length);
    const base = normalizedText(item.querySelector(".control-center-category-trigger")?.textContent || "");
    item.dataset.controlSearchText = `${base} ${nodes.map((node) => normalizedText(node.textContent)).join(" ")}`.trim();
    item.hidden = nodes.length === 0;
  }

  function ensureTools(page, container) {
    let tools = page.querySelector(".control-center-tools");
    if (!tools) {
      tools = document.createElement("section");
      tools.className = "card control-center-tools";
      const label = document.createElement("strong");
      const input = document.createElement("input");
      const status = document.createElement("p");
      label.className = "control-center-tools-title";
      label.textContent = "Trova subito quello che ti serve";
      input.type = "search";
      input.className = "control-center-search";
      input.placeholder = "Cerca report, Firestore, PWA, login…";
      input.setAttribute("aria-label", "Cerca nel Centro di Controllo");
      status.className = "control-center-search-status";
      status.setAttribute("aria-live", "polite");

      input.addEventListener("input", () => {
        const query = normalizedText(input.value);
        let visible = 0;
        let firstMatch = null;
        container.querySelectorAll(":scope > .control-center-category").forEach((item) => {
          const hasContent = Number(item.querySelector(".control-center-category-count")?.textContent || "0") > 0;
          const matches = hasContent && (!query || (item.dataset.controlSearchText || "").includes(query));
          item.hidden = !matches;
          if (matches) {
            visible += 1;
            if (!firstMatch) firstMatch = item;
          }
        });
        if (query && visible === 1 && firstMatch) {
          closeOthers(container, firstMatch);
          setOpen(firstMatch, true);
        }
        status.textContent = query ? `${visible} categori${visible === 1 ? "a trovata" : "e trovate"}.` : "Tocca una categoria: vedrai solo le informazioni che ti servono.";
      });

      status.textContent = "Tocca una categoria: vedrai solo le informazioni che ti servono.";
      tools.append(label, input, status);
    }
    if (tools.parentElement !== container) container.appendChild(tools);
    return tools;
  }

  let enhancing = false;

  function enhance() {
    if (enhancing) return;
    enhancing = true;
    try {
      const page = document.getElementById(PAGE_ID);
      const container = document.getElementById(CONTENT_ID);
      if (!page || !container) return;

      const categories = ensureCategories(container);
      const tools = ensureTools(page, container);
      const candidates = [];

      Array.from(container.children).filter((node) => !isManagedNode(node)).forEach((node) => candidates.push(node));
      categories.forEach((item) => {
        const content = item.querySelector(":scope > .control-center-category-panel > .control-center-category-content");
        if (content) Array.from(content.children).forEach((node) => candidates.push(node));
      });

      const pinned = new Map();
      candidates.forEach((node) => {
        const pinnedKey = getPinnedKey(node);
        if (pinnedKey) {
          node.classList.add("control-center-pinned-section");
          node.dataset.controlPinned = pinnedKey;
          pinned.set(pinnedKey, node);
          return;
        }
        node.classList.remove("control-center-pinned-section");
        delete node.dataset.controlPinned;
        const category = classifyNode(node);
        const item = categories.get(category.key);
        const content = item?.querySelector(":scope > .control-center-category-panel > .control-center-category-content");
        if (content) content.appendChild(node);
      });

      categories.forEach((item) => updateCategoryMeta(item));

      PINNED_SECTIONS.forEach((section) => {
        const node = pinned.get(section.key);
        if (node) container.appendChild(node);
      });
      container.appendChild(tools);
      CATEGORIES.forEach((category) => {
        const item = categories.get(category.key);
        if (item) container.appendChild(item);
      });

      let remembered = "";
      try { remembered = sessionStorage.getItem(STORAGE_KEY) || ""; } catch (_) {}
      if (remembered) {
        const previous = categories.get(remembered);
        if (previous && !previous.hidden && !container.querySelector(":scope > .control-center-category.is-open")) setOpen(previous, true);
      }
    } finally {
      enhancing = false;
    }
  }

  const observer = new MutationObserver((mutations) => {
    const hasExternalNodes = mutations.some((mutation) => Array.from(mutation.addedNodes || []).some((node) => node instanceof HTMLElement && !node.classList.contains("control-center-category") && !node.classList.contains("control-center-tools")));
    if (hasExternalNodes) enhance();
  });

  function start() {
    enhance();
    const container = document.getElementById(CONTENT_ID);
    if (container) observer.observe(container, { childList: true });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
