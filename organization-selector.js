(function () {
  "use strict";

  function createOrganizationSelector(options) {
    const settings = options || {};
    const organizations = Array.isArray(settings.organizations) ? settings.organizations : [];
    const activeOrganizationId = String(settings.activeOrganizationId || "varga");
    const onSelect = typeof settings.onSelect === "function" ? settings.onSelect : function () {};

    const overlay = document.createElement("section");
    overlay.className = "organization-selector hidden";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.setAttribute("aria-labelledby", "organization-selector-title");

    const card = document.createElement("div");
    card.className = "organization-selector-card";

    const title = document.createElement("h2");
    title.id = "organization-selector-title";
    title.textContent = "Scegli organizzazione";

    const description = document.createElement("p");
    description.className = "muted";
    description.textContent = "Vedrai esclusivamente i dati dell’organizzazione selezionata.";

    const list = document.createElement("div");
    list.className = "organization-selector-list";

    organizations.forEach(function (organization) {
      if (!organization || !organization.id) return;

      const button = document.createElement("button");
      button.type = "button";
      button.className = "organization-selector-item";
      button.dataset.organizationId = String(organization.id);
      button.setAttribute("aria-pressed", String(organization.id) === activeOrganizationId ? "true" : "false");

      const name = document.createElement("strong");
      name.textContent = organization.name || organization.id;

      const role = document.createElement("span");
      role.textContent = organization.role || "operatore";

      button.append(name, role);
      button.addEventListener("click", function () {
        onSelect(String(organization.id));
      });
      list.appendChild(button);
    });

    card.append(title, description, list);
    overlay.appendChild(card);

    return {
      element: overlay,
      open: function () {
        overlay.classList.remove("hidden");
      },
      close: function () {
        overlay.classList.add("hidden");
      }
    };
  }

  window.HeraOrganizationSelector = Object.freeze({
    createOrganizationSelector
  });
})();
