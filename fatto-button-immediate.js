(() => {
  "use strict";

  const FATTO_BUTTON_SELECTOR = [
    '.action-icon-btn[data-action-key="whatsapp"]:not(.is-completed-done)',
    'button[data-action-key="whatsapp"]:not(.is-completed-done)'
  ].join(",");

  function formatTodayDayMonth() {
    return new Intl.DateTimeFormat("it-IT", {
      day: "2-digit",
      month: "2-digit"
    }).format(new Date());
  }

  function showImmediateFattoState(button) {
    if (!(button instanceof HTMLButtonElement)) return;

    const dayMonth = formatTodayDayMonth();
    const label = `⚠️ ${dayMonth}`;

    button.textContent = "";
    button.classList.add("is-completed-done");
    button.dataset.doneLabel = label;
    button.setAttribute("aria-label", `Fatto il ${dayMonth}`);
    button.title = `Fatto il ${dayMonth}`;
  }

  function handleFattoInteraction(event) {
    const target = event.target;
    if (!(target instanceof Element)) return;

    const button = target.closest(FATTO_BUTTON_SELECTOR);
    if (!button) return;

    showImmediateFattoState(button);
  }

  // pointerdown provides the visual feedback before any existing FATTO handler
  // can disable or re-render the button. click remains as a keyboard fallback.
  document.addEventListener("pointerdown", handleFattoInteraction, { capture: true });
  document.addEventListener("click", handleFattoInteraction, { capture: true });
})();

(() => {
  "use strict";
  if (document.querySelector('script[data-password-access-manager="true"]')) return;
  const script = document.createElement("script");
  script.src = `password-access-manager.js?v=20260725a`;
  script.defer = true;
  script.dataset.passwordAccessManager = "true";
  document.head.appendChild(script);
})();
