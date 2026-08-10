(() => {
  "use strict";

  const PATCH_VERSION = "2.0.0";
  let observer = null;

  function isIosTouchDevice() {
    const ua = String(navigator.userAgent || "");
    return /iPad|iPhone|iPod/i.test(ua)
      || (navigator.platform === "MacIntel" && Number(navigator.maxTouchPoints || 0) > 1);
  }

  function focusInput(input) {
    if (!input || document.activeElement === input) return;
    try {
      input.focus({ preventScroll: true });
    } catch (_) {
      input.focus();
    }
  }

  function runSearch(input) {
    if (!input) return;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    try {
      window.HeraAdminPasswordManager?.search?.();
    } catch (_) {}
  }

  function openNativeSearch(input) {
    if (!input) return;
    const value = window.prompt(
      "Cerca utente per nome o email",
      String(input.value || "")
    );
    if (value === null) return;
    input.value = value;
    runSearch(input);
  }

  function makeWrapperSafe(input) {
    let wrap = input.closest(".personale-search-input-wrap");
    if (!wrap) return null;

    // Evita il doppio comportamento del <label> su Safari/iOS: il campo resta
    // visivamente identico ma il contenitore non prova più ad attivarlo due volte.
    if (wrap.tagName === "LABEL") {
      const replacement = document.createElement("div");
      replacement.className = wrap.className;
      const icon = wrap.querySelector("span[aria-hidden='true']");
      if (icon) replacement.appendChild(icon);
      replacement.appendChild(input);
      wrap.replaceWith(replacement);
      wrap = replacement;
    }

    return wrap;
  }

  function ensureNativeSearchButton(toolbar, input) {
    if (!toolbar || !input) return;
    let button = toolbar.querySelector("#user-management-native-search");
    if (!button) {
      button = document.createElement("button");
      button.id = "user-management-native-search";
      button.type = "button";
      button.className = "btn btn-small";
      button.textContent = "Cerca";
      const clearButton = toolbar.querySelector("#user-management-search-clear");
      if (clearButton) toolbar.insertBefore(button, clearButton);
      else toolbar.appendChild(button);
    }
    if (button.dataset.searchPromptBound === PATCH_VERSION) return;
    button.dataset.searchPromptBound = PATCH_VERSION;
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openNativeSearch(input);
    });
  }

  function patchSearchInput() {
    const input = document.getElementById("user-management-search-input");
    if (!input) return false;
    if (input.dataset.mobileTypingFix === PATCH_VERSION) return true;

    input.dataset.mobileTypingFix = PATCH_VERSION;
    input.type = "text";
    input.setAttribute("inputmode", "search");
    input.setAttribute("enterkeyhint", "search");
    input.setAttribute("aria-label", "Cerca utente per nome o email");
    input.setAttribute("autocapitalize", "none");
    input.setAttribute("autocomplete", "off");
    input.setAttribute("spellcheck", "false");
    input.removeAttribute("readonly");
    input.disabled = false;
    input.readOnly = false;
    input.tabIndex = 0;

    const wrap = makeWrapperSafe(input);
    const toolbar = input.closest(".user-management-search-toolbar");
    [toolbar, wrap, input].filter(Boolean).forEach((node) => {
      node.style.pointerEvents = "auto";
    });

    input.style.position = "relative";
    input.style.zIndex = "2";
    input.style.webkitUserSelect = "text";
    input.style.userSelect = "text";
    input.style.touchAction = "auto";
    input.style.fontSize = "16px";

    ensureNativeSearchButton(toolbar, input);

    if (wrap && wrap.dataset.userSearchWrapBound !== PATCH_VERSION) {
      wrap.dataset.userSearchWrapBound = PATCH_VERSION;
      wrap.addEventListener("click", (event) => {
        if (event.target === input) return;
        if (isIosTouchDevice()) openNativeSearch(input);
        else focusInput(input);
      });
    }

    if (input.dataset.userSearchInputBound !== PATCH_VERSION) {
      input.dataset.userSearchInputBound = PATCH_VERSION;
      input.addEventListener("click", (event) => {
        event.stopPropagation();
        if (isIosTouchDevice()) {
          event.preventDefault();
          openNativeSearch(input);
          return;
        }
        focusInput(input);
      });
      input.addEventListener("input", () => {
        try {
          window.HeraAdminPasswordManager?.search?.();
        } catch (_) {}
      });
    }

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
    refresh: patchSearchInput,
    openNativeSearch: () => openNativeSearch(document.getElementById("user-management-search-input"))
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();
