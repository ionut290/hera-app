/* Regola A1/A11/A12: codici di lavoro ordinario.
 * Corregge la classificazione condivisa usata da schede e messaggio Whazzup,
 * senza intercettare pulsanti, listener, scritture Firestore o funzioni protette
 * dei flussi FATTO, RESET e Whazzup.
 */
(function () {
  "use strict";

  const ORDINARY_PRICE_CODES = new Set(["A1", "A11", "A12"]);
  const ROOT_ID = "impianti-lista";
  let observer = null;
  let scanScheduled = false;

  function splitDisplayPriceCodes(value) {
    return Array.from(new Set(
      String(value || "")
        .toUpperCase()
        .split(/[^A-Z0-9]+/)
        .map((code) => code.trim())
        .filter(Boolean)
    ));
  }

  function installA1OrdinaryClassificationGuard() {
    const originalHasOrdinario = typeof window.hasOrdinario === "function"
      ? window.hasOrdinario
      : null;
    const originalHasStraordinario = typeof window.hasStraordinario === "function"
      ? window.hasStraordinario
      : null;

    if (!originalHasOrdinario || !originalHasStraordinario) return;
    if (
      originalHasOrdinario.__heraA1OrdinaryGuard === true
      && originalHasStraordinario.__heraA1OrdinaryGuard === true
    ) {
      return;
    }

    function hasOrdinarioWithA1(codicePrezzo) {
      const codes = splitDisplayPriceCodes(codicePrezzo);
      return codes.includes("A1") || originalHasOrdinario(codicePrezzo);
    }

    function hasStraordinarioWithoutA1(codicePrezzo) {
      const codes = splitDisplayPriceCodes(codicePrezzo);
      if (codes.length > 0 && codes.every((code) => ORDINARY_PRICE_CODES.has(code))) {
        return false;
      }
      return originalHasStraordinario(codicePrezzo);
    }

    Object.defineProperty(hasOrdinarioWithA1, "__heraA1OrdinaryGuard", {
      value: true
    });
    Object.defineProperty(hasStraordinarioWithoutA1, "__heraA1OrdinaryGuard", {
      value: true
    });

    window.hasOrdinario = hasOrdinarioWithA1;
    window.hasStraordinario = hasStraordinarioWithoutA1;
  }

  function isA1OrdinaryOnly(impianto) {
    const codes = splitDisplayPriceCodes(impianto?.codicePrezzo || impianto?.voceRiferimento);
    return codes.includes("A1") && codes.every((code) => ORDINARY_PRICE_CODES.has(code));
  }

  function getVisibleImpiantoByKey(key) {
    try {
      if (!Array.isArray(currentImpianti) || typeof buildImpiantoKey !== "function") return null;
      return currentImpianti.find((impianto) => buildImpiantoKey(impianto) === key) || null;
    } catch (_) {
      return null;
    }
  }

  function correctCard(card) {
    const key = String(card?.dataset?.impiantoKey || "");
    const impianto = key ? getVisibleImpiantoByKey(key) : null;
    if (!impianto || !isA1OrdinaryOnly(impianto)) return;

    const typeBadge = Array.from(card.querySelectorAll(".impianto-summary-meta .badge"))
      .find((badge) => badge.classList.contains("badge-straordinaria")
        || /^Straordinaria$/i.test(String(badge.textContent || "").trim()));

    if (typeBadge) {
      typeBadge.textContent = "Ordinaria";
      typeBadge.classList.remove("badge-straordinaria");
      typeBadge.classList.add("badge-ordinaria");
      typeBadge.setAttribute("aria-label", "Lavoro ordinario");
    }

    card.querySelectorAll(".badge-extra-work").forEach((badge) => {
      badge.remove();
    });

    card.querySelectorAll(".impianto-details > p").forEach((row) => {
      const label = row.querySelector("b");
      if (!label || !/Lavori straordinari aperti/i.test(String(label.textContent || ""))) return;
      row.remove();
    });
  }

  function scan() {
    scanScheduled = false;
    const root = document.getElementById(ROOT_ID);
    if (!root) return;
    root.querySelectorAll(".impianto-item[data-impianto-key]").forEach(correctCard);
  }

  function scheduleScan() {
    if (scanScheduled) return;
    scanScheduled = true;
    window.requestAnimationFrame(scan);
  }

  function install() {
    installA1OrdinaryClassificationGuard();

    const root = document.getElementById(ROOT_ID);
    if (!root) return;

    observer?.disconnect();
    observer = new MutationObserver(scheduleScan);
    observer.observe(root, { childList: true, subtree: true });
    scheduleScan();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
