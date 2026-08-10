(() => {
  const menu = document.getElementById("side-menu");
  const body = menu?.querySelector(".side-menu-body");
  if (!menu || !body || body.dataset.professionalMenuReady === "true") return;

  const sections = [
    {
      title: "Operatività",
      className: "menu-section-operational",
      items: [
        ["open-panel-commesse", "Commesse e cantieri", "📋"],
        ["open-panel-squadre", "Pianificazione squadre", "👥"],
        ["open-hours-btn", "Registro ore", "⏱️"],
        ["open-segnalazioni-btn", "Segnalazioni", "⚠️"],
        ["open-panel-programmazione", "Programmazione attività", "📅"]
      ]
    },
    {
      title: "Risorse aziendali",
      className: "menu-section-resources",
      items: [
        ["open-panel-personale", "Personale", "🪪"],
        ["open-panel-mezzi", "Mezzi e attrezzature", "🚜"],
        ["open-private-docs-upload-btn", "Carica documento", "➕"],
        ["open-pos-btn", "POS e sicurezza", "🦺"]
      ]
    },
    {
      title: "Comunicazioni",
      className: "menu-section-communications",
      items: [
        ["open-panel-notifiche", "Centro notifiche", "🔔"],
        ["open-panel-banner", "Comunicazioni in Home", "📣"],
        ["open-panel-banner-gestione", "Gestione comunicazioni", "🖥️"],
        ["open-panel-info-utili", "Informazioni utili", "ℹ️"]
      ]
    },
    {
      title: "Amministrazione",
      className: "menu-section-admin",
      items: [
        ["open-panel-utenti", "Gestione utenti", "🔐"],
        ["open-panel-global", "Archivio Global", "🌐"],
        ["open-control-center-btn", "Centro di controllo", "🛠️"]
      ]
    },
    {
      title: "Area personale",
      className: "menu-section-personal",
      items: [
        ["open-private-docs-btn", "I miei documenti", "📁"],
        ["open-personal-services-btn", "Servizi personali", "👤"]
      ]
    },
    {
      title: "Assistenza e applicazione",
      className: "menu-section-support",
      items: [
        ["open-howto-btn", "Guida all’utilizzo", "❓"],
        ["install-app-btn", "Installa applicazione", "📲"],
        ["open-book-pdf-btn", "Manuale PDF", "📘"]
      ]
    }
  ];

  const knownButtons = new Map(
    Array.from(body.querySelectorAll("button[id]")).map((button) => [button.id, button])
  );
  const fragment = document.createDocumentFragment();

  sections.forEach((sectionConfig) => {
    const validItems = sectionConfig.items.filter(([id]) => knownButtons.has(id));
    if (!validItems.length) return;

    const section = document.createElement("section");
    section.className = `menu-section ${sectionConfig.className}`;

    const heading = document.createElement("h3");
    heading.className = "menu-section-title";
    heading.textContent = sectionConfig.title;

    const items = document.createElement("div");
    items.className = "menu-section-items";

    validItems.forEach(([id, label, icon]) => {
      const button = knownButtons.get(id);
      button.setAttribute("data-menu-icon", icon);

      if (id === "open-panel-utenti") {
        const badge = button.querySelector("#pending-users-menu-badge");
        button.textContent = `${label} `;
        if (badge) button.appendChild(badge);
      } else {
        button.textContent = label;
      }

      items.appendChild(button);
      knownButtons.delete(id);
    });

    section.append(heading, items);
    fragment.appendChild(section);
  });

  if (knownButtons.size) {
    const section = document.createElement("section");
    section.className = "menu-section menu-section-other";
    const heading = document.createElement("h3");
    heading.className = "menu-section-title";
    heading.textContent = "Altri strumenti";
    const items = document.createElement("div");
    items.className = "menu-section-items";
    knownButtons.forEach((button) => {
      button.setAttribute("data-menu-icon", "•");
      items.appendChild(button);
    });
    section.append(heading, items);
    fragment.appendChild(section);
  }

  body.replaceChildren(fragment);
  body.dataset.professionalMenuReady = "true";

  if (!document.querySelector('script[data-admin-password-manager="true"]')) {
    const script = document.createElement("script");
    script.src = "admin-password-manager.js?v=20260810b";
    script.defer = true;
    script.dataset.adminPasswordManager = "true";
    document.head.appendChild(script);
  }

  if (!document.querySelector('script[data-user-management-search-fix="true"]')) {
    const script = document.createElement("script");
    script.src = "user-management-search-input-fix.js?v=20260810a";
    script.defer = true;
    script.dataset.userManagementSearchFix = "true";
    document.head.appendChild(script);
  }
})();
