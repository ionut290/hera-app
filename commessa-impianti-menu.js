/* Accesso rapido alla gestione impianti dalla testata della commessa. */
(() => {
  "use strict";

  const wrap = document.getElementById("commessa-plants-menu-wrap");
  const toggle = document.getElementById("commessa-plants-menu-btn");
  const menu = document.getElementById("commessa-plants-menu");
  const title = document.getElementById("commessa-focus-label");
  if (!wrap || !toggle || !menu || !title) return;

  function closeMenu() {
    menu.classList.add("hidden");
    menu.setAttribute("aria-hidden", "true");
    toggle.setAttribute("aria-expanded", "false");
  }

  function updateAvailability() {
    const available = typeof canManageData === "function"
      && canManageData()
      && typeof selectedCommessaId !== "undefined"
      && Boolean(String(selectedCommessaId || "").trim());
    wrap.classList.toggle("hidden", !available);
    if (!available) closeMenu();
  }

  function openImportSection() {
    const card = document.getElementById("impianti-import-card");
    card?.classList.remove("hidden");
    card?.scrollIntoView({ behavior: "smooth", block: "start" });
    document.getElementById("excel-file")?.focus();
  }

  async function openPlantsManagement(action) {
    closeMenu();
    if (typeof canManageData !== "function" || !canManageData()) {
      window.alert("La gestione impianti è riservata agli amministratori.");
      return;
    }
    const commessaId = typeof selectedCommessaId === "undefined" ? "" : String(selectedCommessaId || "").trim();
    const commessa = commessaId && typeof commesseById !== "undefined" ? commesseById.get(commessaId) : null;
    if (!commessa) {
      window.alert("Seleziona prima una commessa.");
      return;
    }

    openManagementPanel("commesse");
    await Promise.resolve(openImpiantiManagement(commessa));

    if (action === "add") {
      document.getElementById("add-management-impianto-btn")?.click();
      return;
    }
    if (action === "import") {
      openImportSection();
      return;
    }
    if (action === "export") {
      document.getElementById("export-all-impianti-btn")?.click();
      return;
    }
    if (action === "prices") {
      document.getElementById("open-prezziario-btn")?.click();
      return;
    }
    const search = document.getElementById("impianti-management-search");
    search?.scrollIntoView({ behavior: "smooth", block: "center" });
    search?.focus();
  }

  toggle.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    const opening = menu.classList.contains("hidden");
    menu.classList.toggle("hidden", !opening);
    menu.setAttribute("aria-hidden", String(!opening));
    toggle.setAttribute("aria-expanded", String(opening));
    if (opening) menu.querySelector("[role='menuitem']")?.focus();
  });

  menu.addEventListener("click", (event) => {
    const item = event.target.closest("[data-commessa-plants-action]");
    if (!item) return;
    event.preventDefault();
    void openPlantsManagement(item.dataset.commessaPlantsAction || "edit");
  });

  document.addEventListener("click", (event) => {
    if (!wrap.contains(event.target)) closeMenu();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeMenu();
      toggle.focus();
    }
  });

  new MutationObserver(updateAvailability).observe(title, { childList: true, characterData: true, subtree: true });
  window.addEventListener("pageshow", updateAvailability);
  updateAvailability();
})();
