(() => {
  "use strict";

  if (window.HeraOccasionalSquadSites?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const COMMESSA_NAME = "LAVORI OCCASIONALI";
  const CACHE_MS = 30000;
  const assignmentCache = new Map();
  let observer = null;
  let scanQueued = false;

  const normalize = (value) => String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleUpperCase("it-IT");

  function cleanSiteText(value) {
    const raw = String(value ?? "").trim().replace(/\s+/g, " ");
    if (!raw || raw.toLocaleUpperCase("it-IT") === "[OBJECT OBJECT]") return "";
    return raw;
  }

  function extractSiteName(value, seen = new WeakSet()) {
    if (value === null || value === undefined) return "";

    if (typeof value === "string") {
      const raw = cleanSiteText(value);
      if (!raw) return "";
      if ((raw.startsWith("{") && raw.endsWith("}")) || (raw.startsWith("[") && raw.endsWith("]"))) {
        try {
          const parsed = JSON.parse(raw);
          const nestedName = extractSiteName(parsed, seen);
          if (nestedName) return nestedName;
        } catch (_) {}
      }
      return raw;
    }

    if (typeof value === "number" || typeof value === "boolean") return cleanSiteText(value);
    if (typeof value !== "object" || seen.has(value)) return "";

    seen.add(value);
    const candidates = Array.isArray(value)
      ? value
      : [
          value.name,
          value.nome,
          value.denominazione,
          value.cantiere,
          value.lavoroOccasionaleNome,
          value.impiantoNome,
          value.plantName,
          value.placeName,
          value.displayName,
          value.label,
          value.title,
          value.luogo,
          value.value,
          value.text,
          value.metadata,
          value.data,
          value.impianto,
          value.plant,
          value.place
        ];

    for (const candidate of candidates) {
      const name = extractSiteName(candidate, seen);
      if (name) {
        seen.delete(value);
        return name;
      }
    }

    seen.delete(value);
    return "";
  }

  function extractSiteId(value) {
    if (!value || typeof value !== "object") return "";
    const candidates = [
      value.id,
      value.plantId,
      value.impiantoId,
      value.docId,
      value.cantiere?.id,
      value.cantiere?.plantId,
      value.impianto?.id,
      value.plant?.id
    ];
    for (const candidate of candidates) {
      const id = cleanSiteText(candidate);
      if (id) return id;
    }
    return "";
  }

  function uniqueSites(items) {
    const output = [];
    const seen = new Set();
    items.forEach((item) => {
      const name = extractSiteName(item);
      const key = normalize(name);
      if (!key || key === "[OBJECT OBJECT]" || seen.has(key)) return;
      seen.add(key);
      output.push({
        id: extractSiteId(item),
        name: name.toLocaleUpperCase("it-IT")
      });
    });
    return output;
  }

  function dateFromCard(card) {
    const match = String(card?.textContent || "").match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (match) return `${match[3]}-${match[2]}-${match[1]}`;
    const selected = document.getElementById("squadre-filter-date")?.value
      || document.getElementById("squadra-riferimento")?.value
      || "";
    return /^\d{4}-\d{2}-\d{2}$/.test(selected) ? selected : "";
  }

  function compositionSites(dateKey) {
    if (!dateKey) return [];
    try {
      if (!(squadreHistoryByDate instanceof Map)) return [];
      const composition = squadreHistoryByDate.get(dateKey)?.get(COMMESSA_ID);
      if (!composition) return [];
      const rows = Array.isArray(composition.squadre) ? composition.squadre : [];
      return uniqueSites([
        ...rows.map((row) => ({
          name: row?.lavoroOccasionaleNome,
          id: row?.lavoroOccasionalePlantId || row?.plantId || ""
        })),
        { name: composition.lavoroOccasionaleNome, id: composition.lavoroOccasionalePlantId || "" }
      ]);
    } catch (_) {
      return [];
    }
  }

  function activeCollectionName() {
    try {
      return typeof getCommesseCollectionName === "function"
        ? String(getCommesseCollectionName() || "commesse")
        : "commesse";
    } catch (_) {
      return "commesse";
    }
  }

  async function assignmentSites(dateKey) {
    if (!dateKey) return [];
    const cached = assignmentCache.get(dateKey);
    if (cached && Date.now() - cached.at < CACHE_MS) return cached.promise;

    const promise = (async () => {
      try {
        if (typeof db === "undefined" || !db?.collection) return [];
        const snapshot = await db.collection(activeCollectionName())
          .doc(COMMESSA_ID)
          .collection("assegnazioniOccasionali")
          .where("dateKey", "==", dateKey)
          .get();
        const sites = [];
        snapshot.forEach((doc) => {
          const data = doc.data() || {};
          const name = extractSiteName(data.cantiere)
            || extractSiteName(data.nome)
            || extractSiteName(data.denominazione)
            || extractSiteName(data.impianto)
            || extractSiteName(data);
          sites.push({
            id: cleanSiteText(data.plantId) || extractSiteId(data.cantiere) || cleanSiteText(doc.id),
            name
          });
        });
        return uniqueSites(sites);
      } catch (error) {
        console.warn("[Lavori occasionali] elenco cantieri squadra non disponibile", error);
        return [];
      }
    })();

    assignmentCache.set(dateKey, { at: Date.now(), promise });
    return promise;
  }

  function isOccasionalBadge(element) {
    const value = normalize(element?.textContent);
    return value === "OCCASIONALE" || value === "OCCASIONALI";
  }

  function isOccasionalCard(card) {
    if (!card) return false;
    if (String(card.getAttribute?.("data-commessa-id") || "") === COMMESSA_ID) return true;
    if (normalize(card.textContent).includes(COMMESSA_NAME)) return true;
    return Array.from(card.querySelectorAll?.("span, small") || []).some(isOccasionalBadge);
  }

  function findTitleTextNode(root, wantedValues) {
    if (!root) return null;
    const wanted = wantedValues.map(normalize).filter(Boolean);
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (node.parentElement?.closest?.(".occasional-squad-sites")) continue;
      const value = normalize(node.nodeValue);
      if (!value) continue;
      if (wanted.some((item) => value === item || value.includes(item))) return node;
    }
    return null;
  }

  function findHeaderRow(card, sites) {
    const badge = Array.from(card.querySelectorAll("span, small")).find(isOccasionalBadge);
    const headerZone = badge?.parentElement || card.firstElementChild || card;
    const wantedTitles = [COMMESSA_NAME, ...sites.map((site) => site.name)];
    let titleNode = findTitleTextNode(headerZone, wantedTitles);
    let titleHost = titleNode?.parentElement || null;

    if (!titleNode) {
      titleHost = headerZone.querySelector?.("strong, b, h1, h2, h3, h4") || null;
      if (titleHost) {
        const walker = document.createTreeWalker(titleHost, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
          if (String(walker.currentNode.nodeValue || "").trim()) {
            titleNode = walker.currentNode;
            break;
          }
        }
      }
    }

    if (titleNode && titleHost?.dataset.occasionalCommessaTitle !== "1") {
      const hadFolder = /📁|📂/.test(titleNode.nodeValue || titleHost.textContent || "");
      const fragment = document.createDocumentFragment();
      if (hadFolder) fragment.append(document.createTextNode("📁 "));
      const first = document.createElement("span");
      const second = document.createElement("span");
      first.textContent = "LAVORI ";
      second.textContent = "OCCASIONALI";
      fragment.append(first, second);
      titleNode.replaceWith(fragment);
      titleHost.dataset.occasionalCommessaTitle = "1";
    }

    let anchor = titleHost || headerZone;
    while (anchor?.parentElement && anchor.parentElement !== card) anchor = anchor.parentElement;
    return anchor?.parentElement === card ? anchor : null;
  }

  function installStyle() {
    if (document.getElementById("occasional-squad-sites-style")) return;
    const style = document.createElement("style");
    style.id = "occasional-squad-sites-style";
    style.textContent = `
      .occasional-squad-sites{display:grid;gap:5px;margin:7px 0 10px;padding:0}
      .occasional-squad-site{display:flex;align-items:flex-start;gap:7px;min-width:0;color:#0866e5;font-weight:800;line-height:1.25}
      .occasional-squad-site-number{flex:0 0 auto;display:inline-grid;place-items:center;min-width:21px;height:21px;padding:0 5px;border-radius:999px;background:#16a34a;color:#fff;font-size:.76rem;font-weight:900;box-shadow:0 1px 2px rgba(15,23,42,.14)}
      .occasional-squad-site-name{min-width:0;overflow-wrap:anywhere}
      @media(max-width:600px){.occasional-squad-site{font-size:.92rem}}
    `;
    document.head.appendChild(style);
  }

  function paint(card, dateKey, sites) {
    const clean = uniqueSites(sites);
    const key = `${dateKey}|${clean.map((site) => `${site.id}:${site.name}`).join("|")}`;
    let host = card.querySelector(":scope > .occasional-squad-sites");

    if (!clean.length) {
      host?.remove();
      delete card.dataset.occasionalSitesKey;
      return;
    }

    const anchor = findHeaderRow(card, clean);
    if (host && card.dataset.occasionalSitesKey === key) return;

    if (!host) {
      host = document.createElement("div");
      host.className = "occasional-squad-sites";
      host.setAttribute("aria-label", "Cantieri da fare");
      if (anchor) anchor.insertAdjacentElement("afterend", host);
      else card.insertBefore(host, card.firstChild);
    }

    host.replaceChildren(...clean.map((site, index) => {
      const row = document.createElement("div");
      row.className = "occasional-squad-site";
      const number = document.createElement("span");
      number.className = "occasional-squad-site-number";
      number.textContent = String(index + 1);
      const name = document.createElement("span");
      name.className = "occasional-squad-site-name";
      name.textContent = site.name;
      row.append(number, name);
      return row;
    }));
    card.dataset.occasionalSitesKey = key;
  }

  async function renderCard(card) {
    if (!isOccasionalCard(card)) return;
    installStyle();
    const dateKey = dateFromCard(card);
    const primary = compositionSites(dateKey);
    paint(card, dateKey, primary);
    const extras = await assignmentSites(dateKey);
    if (!card.isConnected || dateFromCard(card) !== dateKey) return;
    paint(card, dateKey, [...primary, ...extras]);
  }

  function scan() {
    const list = document.getElementById("squadre-lista");
    if (!list) return;
    Array.from(list.children).forEach((card) => {
      if (isOccasionalCard(card)) void renderCard(card);
      else {
        card.querySelector(":scope > .occasional-squad-sites")?.remove();
        delete card.dataset.occasionalSitesKey;
      }
    });
  }

  function scheduleScan() {
    if (scanQueued) return;
    scanQueued = true;
    requestAnimationFrame(() => {
      scanQueued = false;
      scan();
    });
  }

  function refresh({ clearCache = false } = {}) {
    if (clearCache) assignmentCache.clear();
    scheduleScan();
  }

  function installObserver() {
    const root = document.getElementById("today-squads-section") || document.body;
    if (!root || observer) return;
    observer = new MutationObserver(scheduleScan);
    observer.observe(root, { childList: true, subtree: true, characterData: true });
  }

  document.addEventListener("click", (event) => {
    if (!event.target?.closest?.(".occasional-save-extra")) return;
    window.setTimeout(() => refresh({ clearCache: true }), 1200);
  }, true);

  document.getElementById("squadra-form")?.addEventListener("submit", () => {
    window.setTimeout(() => refresh({ clearCache: true }), 1200);
  });
  document.getElementById("squadre-filter-date")?.addEventListener("change", scheduleScan);
  document.getElementById("squadra-riferimento")?.addEventListener("change", scheduleScan);

  const start = () => {
    installObserver();
    scheduleScan();
    window.setTimeout(scheduleScan, 1000);
    window.setTimeout(scheduleScan, 3000);
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();

  window.HeraOccasionalSquadSites = {
    installed: true,
    version: "1.0.2",
    refresh: () => refresh({ clearCache: true }),
    extractSiteName
  };
})();
