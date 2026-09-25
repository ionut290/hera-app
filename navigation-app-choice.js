(() => {
  "use strict";

  const STORAGE_KEY = "navigationApp";
  const VALID = new Set(["google", "waze"]);

  function getChoice() {
    try {
      const value = String(localStorage.getItem(STORAGE_KEY) || "google").toLowerCase();
      return VALID.has(value) ? value : "google";
    } catch (_) {
      return "google";
    }
  }

  function setChoice(value) {
    const normalized = VALID.has(String(value || "").toLowerCase())
      ? String(value).toLowerCase()
      : "google";
    try { localStorage.setItem(STORAGE_KEY, normalized); } catch (_) {}
    refreshUi();
    return normalized;
  }

  function buildGoogleUrl(lat, lon) {
    return `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${lat},${lon}`)}`;
  }

  function buildWazeUrl(lat, lon) {
    return `https://www.waze.com/ul?ll=${encodeURIComponent(`${lat},${lon}`)}&navigate=yes`;
  }

  function buildNavigationUrl(lat, lon) {
    const latitude = Number(lat);
    const longitude = Number(lon);
    if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) return "";
    return getChoice() === "waze"
      ? buildWazeUrl(latitude, longitude)
      : buildGoogleUrl(latitude, longitude);
  }

  function extractCoordinatesFromGoogleUrl(value) {
    try {
      const url = new URL(String(value || ""), window.location.href);
      if (!/google\./i.test(url.hostname)) return null;
      const raw = url.searchParams.get("destination") || "";
      const [lat, lon] = raw.split(",").map(Number);
      return Number.isFinite(lat) && Number.isFinite(lon) ? { lat, lon } : null;
    } catch (_) {
      return null;
    }
  }

  function refreshUi() {
    const choice = getChoice();
    document.querySelectorAll("[data-navigation-app]").forEach((input) => {
      input.checked = input.value === choice;
    });
    const status = document.getElementById("navigation-app-status");
    if (status) status.textContent = choice === "waze" ? "Waze selezionato" : "Google Maps selezionato";
    const globalButton = document.getElementById("global-impianto-navigate-btn");
    if (globalButton) globalButton.textContent = choice === "waze" ? "Apri Waze" : "Apri Google Maps";
  }

  function installSettingsUi() {
    const toolsSection = document.getElementById("menu-strumenti-title")?.closest(".menu-section");
    if (!toolsSection || document.getElementById("navigation-app-setting")) return;

    const wrap = document.createElement("div");
    wrap.id = "navigation-app-setting";
    wrap.className = "navigation-app-setting";
    wrap.innerHTML = `
      <div class="navigation-app-setting-title"><span aria-hidden="true">🧭</span><strong>App di navigazione</strong></div>
      <div class="navigation-app-options" role="radiogroup" aria-label="App di navigazione predefinita">
        <label><input type="radio" name="navigation-app-choice" value="google" data-navigation-app> Google Maps</label>
        <label><input type="radio" name="navigation-app-choice" value="waze" data-navigation-app> Waze</label>
      </div>
      <small id="navigation-app-status" class="muted"></small>
    `;
    toolsSection.insertBefore(wrap, toolsSection.querySelector("#open-gardening-assistant-btn") || toolsSection.firstChild?.nextSibling || null);
    wrap.addEventListener("change", (event) => {
      const target = event.target;
      if (target?.matches?.("[data-navigation-app]")) setChoice(target.value);
    });
    refreshUi();
  }

  function rewriteNavigationAnchors(root = document) {
    root.querySelectorAll?.('a[href*="google.com/maps/dir"], a[href*="www.google.com/maps/dir"]').forEach((anchor) => {
      const coords = extractCoordinatesFromGoogleUrl(anchor.getAttribute("href"));
      if (!coords) return;
      anchor.dataset.navigationLat = String(coords.lat);
      anchor.dataset.navigationLon = String(coords.lon);
      anchor.href = buildNavigationUrl(coords.lat, coords.lon);
    });
  }

  function interceptNavigationClick(event) {
    const anchor = event.target.closest?.("a[data-navigation-lat][data-navigation-lon], a[href*='google.com/maps/dir']");
    if (!anchor) return;
    const lat = Number(anchor.dataset.navigationLat);
    const lon = Number(anchor.dataset.navigationLon);
    const coords = Number.isFinite(lat) && Number.isFinite(lon)
      ? { lat, lon }
      : extractCoordinatesFromGoogleUrl(anchor.getAttribute("href"));
    if (!coords) return;
    anchor.href = buildNavigationUrl(coords.lat, coords.lon);
  }

  const observer = new MutationObserver((records) => {
    for (const record of records) {
      record.addedNodes.forEach((node) => {
        if (node.nodeType !== 1) return;
        rewriteNavigationAnchors(node);
      });
    }
    refreshUi();
  });

  window.VargaNavigation = {
    getChoice,
    setChoice,
    buildNavigationUrl,
    buildGoogleUrl,
    buildWazeUrl,
    refreshUi
  };

  document.addEventListener("DOMContentLoaded", () => {
    installSettingsUi();
    rewriteNavigationAnchors(document);
    observer.observe(document.body, { childList: true, subtree: true });
    document.addEventListener("click", interceptNavigationClick, true);
    refreshUi();
  });
})();
