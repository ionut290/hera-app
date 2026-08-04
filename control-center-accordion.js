(function () {
  "use strict";

  const STORAGE_KEY = "hera-control-center-open-section";
  const PAGE_ID = "control-center-page";
  const CONTENT_ID = "control-center-content";

  function cleanText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function slug(value, index) {
    const base = cleanText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");
    return base || `sezione-${index + 1}`;
  }

  function getTitle(node, index) {
    const titleNode = node.querySelector("h1, h2, h3, h4, strong, .title, [data-title]");
    return cleanText(titleNode?.textContent) || `Sezione ${index + 1}`;
  }

  function getDescription(node, title) {
    const candidates = Array.from(node.querySelectorAll("p, small, .muted"));
    const description = candidates
      .map((item) => cleanText(item.textContent))
      .find((text) => text && text !== title && text.length <= 180);
    return description || "Tocca per aprire questa sezione.";
  }

  function isIgnoredNode(node) {
    return !(node instanceof HTMLElement) ||
      node.classList.contains("control-center-accordion-item") ||
      node.classList.contains("control-center-empty-search") ||
      node.id === "control-center-results";
  }

  function setOpen(item, open, options = {}) {
    const trigger = item.querySelector(":scope > .control-center-accordion-trigger");
    const panel = item.querySelector(":scope > .control-center-accordion-panel");
    if (!trigger || !panel) return;

    item.classList.toggle("is-open", open);
    trigger.setAttribute("aria-expanded", String(open));
    panel.hidden = !open;

    if (open) {
      const key = item.dataset.controlSectionKey || "";
      try { sessionStorage.setItem(STORAGE_KEY, key); } catch (_) {}
      if (options.scroll) {
        requestAnimationFrame(() => item.scrollIntoView({ behavior: "smooth", block: "start" }));
      }
    }
  }

  function closeOthers(container, current) {
    container.querySelectorAll(":scope > .control-center-accordion-item.is-open").forEach((item) => {
      if (item !== current) setOpen(item, false);
    });
  }

  function buildItem(node, index, container) {
    const title = getTitle(node, index);
    const description = getDescription(node, title);
    const key = slug(title, index);
    const item = document.createElement("section");
    const trigger = document.createElement("button");
    const panel = document.createElement("div");
    const titleSpan = document.createElement("span");
    const descriptionSpan = document.createElement("span");
    const panelId = `control-center-panel-${key}-${index + 1}`;

    item.className = "control-center-accordion-item";
    item.dataset.controlSectionKey = key;
    item.dataset.controlSearchText = cleanText(node.textContent).toLowerCase();

    trigger.type = "button";
    trigger.className = "control-center-accordion-trigger";
    trigger.setAttribute("aria-expanded", "false");
    trigger.setAttribute("aria-controls", panelId);

    titleSpan.className = "control-center-accordion-title";
    titleSpan.textContent = title;
    descriptionSpan.className = "control-center-accordion-description";
    descriptionSpan.textContent = description;
    trigger.append(titleSpan, descriptionSpan);

    panel.className = "control-center-accordion-panel";
    panel.id = panelId;
    panel.hidden = true;
    panel.appendChild(node);

    trigger.addEventListener("click", () => {
      const shouldOpen = !item.classList.contains("is-open");
      closeOthers(container, item);
      setOpen(item, shouldOpen, { scroll: shouldOpen });
    });

    item.append(trigger, panel);
    return item;
  }

  function ensureTools(page, container) {
    let tools = page.querySelector(":scope > .control-center-tools");
    if (tools) return tools;

    tools = document.createElement("section");
    tools.className = "card control-center-tools";
    const input = document.createElement("input");
    const status = document.createElement("p");

    input.type = "search";
    input.className = "control-center-search";
    input.placeholder = "Cerca una funzione…";
    input.setAttribute("aria-label", "Cerca una funzione nel Centro di Controllo");
    status.className = "control-center-search-status";
    status.setAttribute("aria-live", "polite");

    input.addEventListener("input", () => {
      const query = cleanText(input.value).toLowerCase();
      let visible = 0;
      container.querySelectorAll(":scope > .control-center-accordion-item").forEach((item) => {
        const matches = !query || (item.dataset.controlSearchText || "").includes(query);
        item.hidden = !matches;
        if (matches) visible += 1;
      });
      status.textContent = query
        ? `${visible} sezione${visible === 1 ? "" : "i"} trovata${visible === 1 ? "" : "e"}.`
        : "Tocca una sezione per aprirla.";
    });

    status.textContent = "Tocca una sezione per aprirla.";
    tools.append(input, status);
    page.insertBefore(tools, container);
    return tools;
  }

  function enhance() {
    const page = document.getElementById(PAGE_ID);
    const container = document.getElementById(CONTENT_ID);
    if (!page || !container) return;

    ensureTools(page, container);

    const rawNodes = Array.from(container.children).filter((node) => !isIgnoredNode(node));
    rawNodes.forEach((node, index) => {
      if (node.closest(".control-center-accordion-item")) return;
      container.appendChild(buildItem(node, index, container));
    });

    let remembered = "";
    try { remembered = sessionStorage.getItem(STORAGE_KEY) || ""; } catch (_) {}
    if (remembered) {
      const previous = Array.from(container.querySelectorAll(":scope > .control-center-accordion-item"))
        .find((item) => item.dataset.controlSectionKey === remembered);
      if (previous && !previous.hidden) {
        closeOthers(container, previous);
        setOpen(previous, true);
      }
    }
  }

  const observer = new MutationObserver(() => enhance());

  function start() {
    enhance();
    const container = document.getElementById(CONTENT_ID);
    if (container) observer.observe(container, { childList: true });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }
})();
