window.firebaseConfig = {
  apiKey: "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k",
  authDomain: "hera-app-6cd2b.firebaseapp.com",
  projectId: "hera-app-6cd2b",
  storageBucket: "hera-app-6cd2b.firebasestorage.app",
  messagingSenderId: "645390631375",
  appId: "1:645390631375:web:df3659a23812560e4012ba"
};

function recoverAbandonedFirestoreDiagnosticListeners() {
  const today = () => new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Rome" });
  const key = `varga_fs_diag_v4_${today()}`;
  const now = Date.now();
  const closedAt = new Date(now).toISOString();
  try {
    const value = JSON.parse(localStorage.getItem(key) || "null");
    if (!value || value.date !== today() || !value.listeners) return;
    let recovered = 0;
    Object.values(value.listeners).forEach((listener) => {
      if (!listener?.active) return;
      listener.active = false;
      listener.abandoned = true;
      listener.closeReason = "page-session-ended-without-unsubscribe";
      listener.closedAt = closedAt;
      listener.durationMs = listener.openedAt ? Math.max(0, now - new Date(listener.openedAt).getTime()) : null;
      recovered += 1;
    });
    if (!recovered) return;
    value.totals = value.totals || {};
    value.totals.abandonedListenersRecovered = Math.max(0, Number(value.totals.abandonedListenersRecovered) || 0) + recovered;
    value.lifecycle = Array.isArray(value.lifecycle) ? value.lifecycle : [];
    value.lifecycle.unshift({ at: closedAt, type: "previous-page-listeners-recovered", recoveredListeners: recovered, note: "Listener rimasti attivi nel salvataggio locale dopo ricaricamento, crash o chiusura della pagina." });
    value.updatedAt = closedAt;
    localStorage.setItem(key, JSON.stringify(value));
  } catch (_) {}
}
recoverAbandonedFirestoreDiagnosticListeners();

const HERA_STORAGE_QUOTA_GUARD_SRC = "storage-quota-guard.js?v=20260812a";
const HERA_APP_NOTIFICATIONS_READ_GUARD_SRC = "app-notifications-read-guard.js?v=20260815a";
const HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC = "firestore-operation-diagnostics.js?v=20260806a";
const HERA_FIRESTORE_DIAGNOSTICS_V4_SRC = "firestore-diagnostics-dashboard-v4.js?v=20260805b";
const HERA_FIRESTORE_SAFE_OPTIMIZER_SRC = "firestore-safe-optimizer.js?v=20260805b";
const HERA_NATIVE_RUNTIME_SRC = "native-android-runtime.js?v=20260803-whatsapp-early2";
const HERA_FIRESTORE_INFLIGHT_COALESCER_SRC = "firestore-inflight-read-coalescer.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC = "firestore-diagnostics-optimizer-extension.js?v=20260804b";
const HERA_SHARED_STATIC_VIEWS_SRC = "shared-static-views.js?v=20260810-opera1";
const HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC = "active-commesse-first-boot-guard.js?v=20260816-fast1";
const HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC = "firestore-startup-cost-optimizer.js?v=20260816-fast1";
const HERA_SHARED_STATIC_VIEWS_UI_SRC = "shared-static-views-ui.js?v=20260815-squadre-board1";
const HERA_ADMIN_USER_ACCESS_TOOLS_SRC = "admin-user-access-tools.js?v=20260810a";
const HERA_ADMIN_USER_ACCESS_SHARE_FIX_SRC = "admin-user-access-share-fix.js?v=20260811a";
const HERA_OCCASIONAL_GOOGLE_PLACES_SRC = "lavori-occasionali-google-places.js?v=20260823-map2";
const HERA_OCCASIONAL_MULTI_SITE_HOURS_SRC = "lavori-occasionali-multi-cantiere-ore.js?v=20260823a";
const HERA_OCCASIONAL_PDF_STORAGE_SRC = "lavori-occasionali-pdf-storage.js?v=20260823a";

if (document.readyState === "loading") {
  document.write(`<script src="${HERA_STORAGE_QUOTA_GUARD_SRC}" data-storage-quota-guard="1"><\/script>`);
  document.write(`<script src="${HERA_APP_NOTIFICATIONS_READ_GUARD_SRC}" data-app-notifications-read-guard="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC}" data-firestore-operation-diagnostics="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_DIAGNOSTICS_V4_SRC}" data-firestore-diagnostics-v4="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_SAFE_OPTIMIZER_SRC}" data-firestore-safe-optimizer="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_INFLIGHT_COALESCER_SRC}"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC}"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_SRC}"><\/script>`);
  document.write(`<script src="${HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC}" data-active-commesse-first-boot-guard="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC}" data-firestore-startup-cost-optimizer="1"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_UI_SRC}"><\/script>`);
  document.write(`<script src="${HERA_NATIVE_RUNTIME_SRC}"><\/script>`);
  document.write(`<script src="${HERA_ADMIN_USER_ACCESS_TOOLS_SRC}"><\/script>`);
  document.write(`<script src="${HERA_ADMIN_USER_ACCESS_SHARE_FIX_SRC}"><\/script>`);
  document.write('<script src="notification-session-enhancements.js?v=20260727b"><\/script>');
  document.write('<script src="update-app-feature.js?v=20260824-oneclick1"><\/script>');
  document.write('<script src="google-sheet-two-way-sync.js?v=20260729b"><\/script>');
  document.write('<script src="personnel-training-manager.js?v=20260803a"><\/script>');
  document.write(`<script src="${HERA_OCCASIONAL_GOOGLE_PLACES_SRC}" data-occasional-google-places="1"><\/script>`);
  document.write(`<script src="${HERA_OCCASIONAL_MULTI_SITE_HOURS_SRC}" data-occasional-multi-site-hours="1"><\/script>`);
  document.write(`<script src="${HERA_OCCASIONAL_PDF_STORAGE_SRC}" data-occasional-pdf-storage="1"><\/script>`);
} else {
  function normalizeAssetPath(value) {
    try { return new URL(String(value || ""), document.baseURI).pathname; }
    catch (_) { return String(value || "").split("?")[0].split("#")[0]; }
  }
  function findExistingScript(src, dataName) {
    const byData = document.querySelector(`script[data-${dataName}="true"], script[data-${dataName}="1"]`);
    if (byData) return byData;
    const wantedPath = normalizeAssetPath(src);
    if (!wantedPath) return null;
    return Array.from(document.scripts || []).find((script) => normalizeAssetPath(script.src) === wantedPath) || null;
  }
  function loadOnce(src, dataName, ready, onLoad) {
    if (ready?.()) { onLoad?.(); return; }
    const existing = findExistingScript(src, dataName);
    if (existing) { if (onLoad) existing.addEventListener("load", onLoad, { once: true }); return; }
    const script = document.createElement("script");
    script.src = src;
    script.setAttribute(`data-${dataName}`, "true");
    if (onLoad) script.addEventListener("load", onLoad, { once: true });
    document.head.appendChild(script);
  }
  loadOnce(HERA_STORAGE_QUOTA_GUARD_SRC, "storage-quota-guard", () => window.HeraStorageQuotaGuard?.installed);
  loadOnce(HERA_APP_NOTIFICATIONS_READ_GUARD_SRC, "app-notifications-read-guard", () => window.VargaAppNotificationsReadGuard?.installed);
  const loadSafeOptimizer = () => loadOnce(HERA_FIRESTORE_SAFE_OPTIMIZER_SRC, "firestore-safe-optimizer", () => window.VargaFirestoreSafeOptimizer?.installed);
  const loadDiagnosticsV4 = () => loadOnce(HERA_FIRESTORE_DIAGNOSTICS_V4_SRC, "firestore-diagnostics-v4", () => window.VargaFirestoreDiagnosticsV4?.installed, loadSafeOptimizer);
  loadOnce(HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC, "firestore-operation-diagnostics", () => window.__vargaFsDiagV3, loadDiagnosticsV4);
  window.setTimeout(loadDiagnosticsV4, 150);
  window.setTimeout(loadSafeOptimizer, 300);
  loadOnce(HERA_FIRESTORE_INFLIGHT_COALESCER_SRC, "hera-firestore-inflight-coalescer", () => window.HeraFirestoreInflightReadCoalescer?.installed);
  loadOnce(HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC, "hera-firestore-diagnostics-optimizer", () => window.__vargaFsOptimizerDiagnosticsExtension);
  loadOnce(HERA_SHARED_STATIC_VIEWS_SRC, "hera-shared-static-views", () => window.HeraSharedStaticViews?.installed);
  loadOnce(HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC, "active-commesse-first-boot-guard", () => window.HeraActiveCommesseFirstBootGuard?.installed);
  loadOnce(HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC, "hera-firestore-startup-cost-optimizer", () => window.HeraFirestoreStartupCostOptimizer?.installed);
  loadOnce(HERA_SHARED_STATIC_VIEWS_UI_SRC, "hera-shared-static-views-ui", () => window.__heraSharedStaticViewsUiInstalled);
  loadOnce(HERA_NATIVE_RUNTIME_SRC, "hera-native-runtime", () => false);
  loadOnce(HERA_ADMIN_USER_ACCESS_TOOLS_SRC, "hera-admin-user-access-tools", () => window.HeraAdminUserAccessTools?.version === "1.0.0");
  loadOnce(HERA_ADMIN_USER_ACCESS_SHARE_FIX_SRC, "hera-admin-user-access-share-fix", () => window.HeraAdminUserAccessShareFix?.installed);
  loadOnce("notification-session-enhancements.js?v=20260727b", "hera-notification-session", () => false);
  loadOnce("update-app-feature.js?v=20260824-oneclick1", "hera-app-update", () => false);
  loadOnce("google-sheet-two-way-sync.js?v=20260729b", "hera-google-sheet-sync", () => false);
  loadOnce("personnel-training-manager.js?v=20260803a", "personnel-training-manager", () => false);
  loadOnce(HERA_OCCASIONAL_GOOGLE_PLACES_SRC, "occasional-google-places", () => Boolean(window.HeraLavoriOccasionaliGooglePlaces?.installed));
  loadOnce(HERA_OCCASIONAL_MULTI_SITE_HOURS_SRC, "occasional-multi-site-hours", () => Boolean(window.HeraOccasionalMultiSiteHours?.installed));
  loadOnce(HERA_OCCASIONAL_PDF_STORAGE_SRC, "occasional-pdf-storage", () => Boolean(window.HeraOccasionalPdfStorage?.installed));
}
