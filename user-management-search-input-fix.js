(() => {
  "use strict";

  const PATCH_VERSION = "1.0.0";
  let observer = null;

  function focusInput(input) {
    if (!input || document.activeElement === input) return;
    try {
      input.focus({ preventScroll: true });
    } catch (_) {
      input.focus();
    }
  }

  function patchSearchInput() {
    const input = document.getElementById("user-management-search-input");
    if (!input) return false;
    if (input.dataset.mobileTypingFix === PATCH_VERSION) return true;

    input.dataset.mobileTypingFix = PATCH_VERSION;

    // Su iOS il controllo search creato dinamicamente può non aprire in modo
    // affidabile la tastiera. Manteniamo la ricerca ma usiamo un normale campo
    // testo con tastiera/Invio impostati come ricerca.
    input.type = "text";
    input.setAttribute("inputmode", "search");
    input.setAttribute("enterkeyhint", "search");
    input.setAttribute("aria-label", "Cerca utente per nome o email");
    input.setAttribute("autocapitalize", "none");
    input.setAttribute("spellcheck", "false");
    input.removeAttribute("readonly");
    input.disabled = false;
    input.tabIndex = 0;

    const wrap = input.closest(".personale-search-input-wrap");
    const toolbar = input.closest(".user-management-search-toolbar");
    [toolbar, wrap, input].filter(Boolean).forEach((node) => {
      node.style.pointerEvents = "auto";
    });
    input.style.webkitUserSelect = "text";
    input.style.userSelect = "text";
    input.style.touchAction = "manipulation";

    // Toccando anche icona o bordo del campo, porta subito il cursore nel campo.
    wrap?.addEventListener("click", (event) => {
      if (event.target === input) return;
      focusInput(input);
    });

    // Evita che eventuali click delegati della schermata intercettino il tap
    // destinato alla casella, senza bloccare il comportamento nativo della tastiera.
    input.addEventListener("click", (event) => {
      event.stopPropagation();
      focusInput(input);
    });

    return true;
  }

  function initialize() {
    if (patchSearchInput()) return;
    if (!document.body) return;

    observer = new MutationObserver(() => {
      if (!patchSearchInput()) return;
      observer?.disconnect();
      observer = null;
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  window.HeraUserManagementSearchFix = {
    installed: true,
    version: PATCH_VERSION,
    refresh: patchSearchInput
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();
