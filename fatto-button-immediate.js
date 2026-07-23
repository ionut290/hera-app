(() => {
  "use strict";

  const FATTO_BUTTON_SELECTOR =
    '.impianto-primary-actions .action-icon-btn[data-action-key="whatsapp"]:not(.is-completed-done)';

  function formatTodayDayMonth() {
    return new Intl.DateTimeFormat("it-IT", {
      day: "2-digit",
      month: "2-digit"
    }).format(new Date());
  }

  function showImmediateFattoState(button) {
    if (!(button instanceof HTMLButtonElement) || button.disabled) return;

    const dayMonth = formatTodayDayMonth();
    const label = `⚠️ ${dayMonth}`;

    button.textContent = "";
    button.classList.add("is-completed-done");
    button.dataset.doneLabel = label;
    button.setAttribute("aria-label", `Fatto il ${dayMonth}`);
    button.title = `Fatto il ${dayMonth}`;
  }

  document.addEventListener(
    "click",
    (event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;

      const button = target.closest(FATTO_BUTTON_SELECTOR);
      if (!button) return;

      showImmediateFattoState(button);
    },
    { capture: true }
  );
})();
