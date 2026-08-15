const CACHE_NAME = "varga-cantieri-shell-v133";
const CACHE_RESET_VERSION = "20260815-firestore-read-guard1";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=20260814-loading-humor1",
  "./management-v2.css?v=20260731",
  "./notification-center.css?v=20260729a",
  "./approval-access.css?v=20260731-legacy1",
  "./accounting-v2.css?v=20260728",
  "./calendar-feature.css?v=20260728b",
  "./squadre-restyle.css?v=20260731-mezzi1",
  "./app-pure-utils.js?v=20260815-mod1",
  "./app-worklimate.js?v=20260815-mod1",
  "./app-atex.js?v=20260815-mod1",
  "./app-documents.js?v=20260815-mod1",
  "./app.js?v=20260813-render-gate1",
  "./loading-humor.js?v=20260814a",
  "./management-core.js?v=20260731",
  "./management-v2.js?v=20260731",
  "./registry-google-sheet-sync.js?v=20260802-cost2",
  "./auth-login-fix.js?v=20260731-legacy1",
  "./login-retry-fix.js?v=20260726f",
  "./first-login-password.js?v=20260726b",
  "./approval-access.js?v=20260731-legacy1",
  "./coordinate-repair.js?v=20260728a",
  "./inrete-work-items-v2.js?v=20260728b",
  "./accounting-v2.js?v=20260812-modena2",
  "./accounting-view-guard.js?v=20260728a",
  "./operational-import-repair.js?v=20260728a",
  "./google-sheet-two-way-sync.js?v=20260729b",
  "./native-android-runtime.js?v=20260803-whatsapp-early2",
  "./notification-center.js?v=20260731-header1",
  "./today-summary-interactions.js?v=20260731-repair1",
  "./mezzi-alimentazione-fix.js?v=20260723",
  "./today-live-hours-vehicles.js?v=20260731-codes1",
  "./squad-operator-profile.js?v=20260731a",
  "./operator-profile-feature.js?v=20260802-cost1",
  "./personnel-training-manager.js?v=20260803a",
  "./fatto-button-immediate.js?v=20260727-fatto2",
  "./header-menu-runtime.js?v=20260804-diagnostics-reset1",
  "./firestore-presence-cost-guard.js?v=20260802a",
  "./preventivi-lazy-loader.js?v=20260801a",
  "./global-archive-sync.js?v=20260802-cost2",
  "./global-archive-new-commesse-fix.js?v=20260801-lazy1",
  "./varga-branding.js?v=20260731a",
  "./whazzup-preload-cache.js?v=20260803-installed-only2",
  "./fuel-stations-core.js",
  "./fuel-stations-national-cache.js",
  "./fuel-stations-search.js",
  "./fuel-stations-integration.js",
  "./registry-device-cache.js?v=20260804c",
  "./firestore-registry-read-optimizer.js?v=20260804b",
  "./firestore-safe-optimizer.js?v=20260805b",
  "./firestore-inflight-read-coalescer.js?v=20260804a",
  "./firestore-diagnostics-optimizer-extension.js?v=20260804a",
  "./app-notifications-read-guard.js?v=20260815a",
  "./firebase-config.js",
  "./manifest.webmanifest",
  "./icons/varga-cantieri-32.png",
  "./icons/varga-cantieri-180.png",
  "./icons/varga-cantieri-192.png",
  "./icons/varga-cantieri-512.png",
  "./icons/varga-cantieri-maskable-512.png"
];

const CACHEABLE_DESTINATIONS = new Set(["script", "style", "document", "image", "font"]);
const NETWORK_DOCUMENT_TIMEOUT_MS = 7000;
const NETWORK_FIRST_ASSET_PATHS = new Set([
  "/shared-static-views-client.js",
  "/firebase-config.js",
  "/auth-login-fix.js",
  "/login-retry-fix.js",
  "/first-login-password.js",
  "/approval-access.js",
  "/header-menu-runtime.js"
]);

const isDynamicEndpoint = (url) => [/^\/api(?:\/|$)/, /^\/graphql(?:\/|$)/, /^\/auth(?:\/|$)/, /^\/socket(?:\/|$)/]
  .some((pattern) => pattern.test(url.pathname));

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => Promise.allSettled(APP_SHELL.map((asset) => cache.add(asset))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin || isDynamicEndpoint(url)) return;

  if (event.request.mode === "navigate" || event.request.destination === "document") {
    event.respondWith((async () => {
      const timeout = new Promise((resolve) => setTimeout(() => resolve(null), NETWORK_DOCUMENT_TIMEOUT_MS));
      try {
        const response = await Promise.race([fetch(event.request), timeout]);
        if (response) {
          const cache = await caches.open(CACHE_NAME);
          cache.put(event.request, response.clone()).catch(() => {});
          return response;
        }
      } catch (_) {}
      return (await caches.match(event.request)) || (await caches.match("./index.html"));
    })());
    return;
  }

  const networkFirst = NETWORK_FIRST_ASSET_PATHS.has(url.pathname);
  if (networkFirst) {
    event.respondWith((async () => {
      try {
        const response = await fetch(event.request);
        const cache = await caches.open(CACHE_NAME);
        cache.put(event.request, response.clone()).catch(() => {});
        return response;
      } catch (_) {
        return (await caches.match(event.request)) || Response.error();
      }
    })());
    return;
  }

  if (!CACHEABLE_DESTINATIONS.has(event.request.destination)) return;
  event.respondWith((async () => {
    const cached = await caches.match(event.request);
    if (cached) return cached;
    try {
      const response = await fetch(event.request);
      const cache = await caches.open(CACHE_NAME);
      cache.put(event.request, response.clone()).catch(() => {});
      return response;
    } catch (_) {
      return Response.error();
    }
  })());
});
