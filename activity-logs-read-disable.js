(() => {
  "use strict";

  if (window.HeraActivityLogsReadGuard?.installed || typeof loadActiveUsersLogs !== "function") return;

  loadActiveUsersLogs = async function loadActiveUsersLogsDisabled() {
    activeUsersLogs = [];

    if (ui.activeUsersFilterOperator) {
      ui.activeUsersFilterOperator.innerHTML = '<option value="">Tutti operatori</option>';
    }
    if (ui.activeUsersFilterAction) {
      ui.activeUsersFilterAction.innerHTML = '<option value="">Tutte azioni</option>';
    }
    if (ui.activeUsersLogList) {
      ui.activeUsersLogList.classList.remove("hidden");
      ui.activeUsersLogList.innerHTML = '<p class="muted">Registro attività disattivato per evitare letture Firestore.</p>';
    }
    if (ui.activeUsersLogToggle) {
      ui.activeUsersLogToggle.classList.add("hidden");
      ui.activeUsersLogToggle.setAttribute("aria-hidden", "true");
    }

    renderActiveUsersDetail();
    return true;
  };

  window.HeraActivityLogsReadGuard = {
    installed: true,
    collection: "activityLogs",
    mode: "reads-disabled"
  };
})();
