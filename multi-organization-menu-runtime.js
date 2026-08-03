(function () {
  "use strict";

  const BUTTON_ID = "open-multi-organization-admin-btn";
  const ADMIN_SOURCE_BUTTON_ID = "open-panel-utenti";
  const TARGET_URL = "./multi-organization-admin.html";

  function isElementUnavailable(element) {
    if (!element) return true;
    return element.hidden
      || element.disabled
      || element.classList.contains("hidden")
      || element.getAttribute("aria-hidden") === "true";
  }

  function syncVisibility(sourceButton, organizationButton) {
    const unavailable = isElementUnavailable(sourceButton);
    organizationButton.classList.toggle("hidden", unavailable);
    organizationButton.hidden = unavailable;
    organizationButton.setAttribute("aria-hidden", unavailable ? "true" : "false");
    organizationButton.disabled = unavailable;
  }

  function createOrganizationButton(sourceButton) {
    if (document.getElementById(BUTTON_ID)) {
      return document.getElementById(BUTTON_ID);
    }

    const button = document.createElement("button");
    button.id = BUTTON_ID;
    button.type = "button";
    button.className = sourceButton.className;
    button.innerHTML = '<span class="menu-item-icon" aria-hidden="true">🏢</span>Organizzazioni';
    button.addEventListener("click", function () {
      window.location.assign(TARGET_URL);
    });

    sourceButton.insertAdjacentElement("afterend", button);
    syncVisibility(sourceButton, button);

    const observer = new MutationObserver(function () {
      syncVisibility(sourceButton, button);
    });
    observer.observe(sourceButton, {
      attributes: true,
      attributeFilter: ["class", "hidden", "disabled", "aria-hidden"]
    });

    return button;
  }

  function install() {
    const sourceButton = document.getElementById(ADMIN_SOURCE_BUTTON_ID);
    if (!sourceButton) return false;
    createOrganizationButton(sourceButton);
    return true;
  }

  function installWhenReady() {
    if (install()) return;

    const rootObserver = new MutationObserver(function (_mutations, observer) {
      if (install()) observer.disconnect();
    });
    rootObserver.observe(document.documentElement, { childList: true, subtree: true });

    window.setTimeout(function () {
      rootObserver.disconnect();
    }, 15000);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", installWhenReady, { once: true });
  } else {
    installWhenReady();
  }
})();
