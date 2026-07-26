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

  function loadNativeAndroidRuntime() {
    if (document.querySelector('script[data-hera-native-runtime="true"]')) return;
    const script = document.createElement("script");
    script.src = "native-android-runtime.js";
    script.defer = true;
    script.dataset.heraNativeRuntime = "true";
    document.head.appendChild(script);
  }

  // pointerup evita falsi FATTO quando il gesto viene annullato o diventa scroll.
  // click resta il fallback per tastiera e tecnologie assistive.
  document.addEventListener("pointerup", handleFattoInteraction, { capture: true });
  document.addEventListener("click", handleFattoInteraction, { capture: true });
  loadNativeAndroidRuntime();
})();
