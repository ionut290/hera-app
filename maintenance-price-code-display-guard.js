/* Correzione esclusivamente visiva: A1/A11/A12 sono codici di lavoro ordinario.
 * Non modifica dati, Firestore, FATTO, RESET o il messaggio Whazzup protetto.
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

  function isA1OrdinaryOnly(impianto) {
    const codes = splitDisplayPriceCodes(impianto?.codicePrezzo || impianto?.voceRiferimento);
    return codes.includes("A1") && codes.every((code) => ORDINARY_PRICE_CODES.has(code));
  }

  function hasExplicitExtraWork(impianto) {
    return Boolean(String(impianto?.extraWorkText || "").trim());
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

    if (hasExplicitExtraWork(impianto)) return;

    card.querySelectorAll(".badge-extra-work").forEach((badge) => {
      badge.hidden = true;
      badge.setAttribute("aria-hidden", "true");
    });

    card.querySelectorAll(".impianto-details > p").forEach((row) => {
      const label = row.querySelector("b");
      if (!label || !/Lavori straordinari aperti/i.test(String(label.textContent || ""))) return;
      row.hidden = true;
      row.setAttribute("aria-hidden", "true");
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
