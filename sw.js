const CACHE_NAME = "varga-cantieri-shell-v221";
const CACHE_RESET_VERSION = "20260904-menu-cache1";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=20260831-potature-followup1",
  "./management-v2.css?v=20260731",
  "./notification-center.css?v=20260729a",
  "./approval-access.css?v=20260828-email1",
  "./accounting-v2.css?v=20260901-current-gps1",
  "./calendar-feature.css?v=20260826-admin1",
  "./squadre-restyle.css?v=20260731-mezzi1",
  "./app-pure-utils.js?v=20260902-verde-levato2",
  "./verde-bologna.js?v=20260901-cobo-sfalcio1",
  "./verde-bologna-operativo.js?v=20260901-catasto-open2",
  "./verde-bologna-parchi-mobile.js?v=20260901-cobo-sfalcio1",
  "./verde-levato.js?v=20260902-verde-levato2",
  "./data-license-view.js?v=20260831-license1",
  "./app-worklimate.js?v=20260815-mod1",
  "./app-atex.js?v=20260815-mod1",
  "./app-documents.js?v=20260815-mod1",
  "./app-calendar.js?v=20260815-mod1",
  "./administrative-calendar.js?v=20260829-admin-hours1",
  "./app-snow.js?v=20260815-mod1",
  "./app-availability.js?v=20260815-mod1",
  "./app.js?v=20260902-special-finiti-stable1",
  "./green-assistant.css?v=20260830-brave-manuals3",
  "./green-assistant.js?v=20260830-brave-manuals3",
  "./android-play-store-install.js?v=20260828a",
  "./data-durability-runtime.js?v=20260818a",
  "./data-safety-layer.js?v=20260819a",
  "./critical-write-safety-bridge.js?v=20260824-oneclick1",
  "./update-app-feature.js?v=20260829-update-notice1",
  "./heavy-libs-lazy-loader.js?v=20260815a",
  "./identity-feature-lazy-loader.js?v=20260815a",
  "./loading-humor.js?v=20260829-firestore-quota-dedup1",
  "./client-error-reporter.js?v=20260816a",
  "./app-error-monitor.js?v=20260829-firestore-quota-dedup1",
  "./admin-error-center.js?v=20260829-chatgpt-category1",
  "./admin-error-center.css?v=20260824b",
  "./management-core.js?v=20260731",
  "./management-v2.js?v=20260830-brave-manuals3",
  "./registry-google-sheet-sync.js?v=20260802-cost2",
  "./auth-login-fix.js?v=20260731-legacy1",
  "./login-retry-fix.js?v=20260830-existing-account1",
  "./first-login-password.js?v=20260726b",
  "./approval-access.js?v=20260828-authfix",
  "./coordinate-repair.js?v=20260728a",
  "./inrete-work-items-v2.js?v=20260728b",
  "./accounting-v2.js?v=20260901-current-gps1",
  "./commessa-impianti-menu.js?v=20260826c",
  "./accounting-view-guard.js?v=20260728a",
  "./operational-import-repair.js?v=20260901-manual-plant-summary1",
  "./google-sheet-two-way-sync.js?v=20260729b",
  "./native-android-runtime.js?v=20260803-whatsapp-early2",
  "./notification-center.js?v=20260830-firestore-read-fix1",
  "./today-summary-interactions.js?v=20260731-repair1",
  "./mezzi-alimentazione-fix.js?v=20260723",
  "./today-live-hours-vehicles.js?v=20260731-codes1",
  "./squad-operator-profile.js?v=20260731a",
  "./operator-profile-feature.js?v=20260802-cost1",
  "./personnel-training-manager.js?v=20260803a",
  "./fatto-button-immediate.js?v=20260824-fatto-oneclick1",
  "./header-menu-runtime.js?v=20260830-map-gps-fix1",
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
  "./recommended-plants.css?v=20260823-stability3",
  "./recommended-plants.js?v=20260823-stability3",
  "./tree-search.css?v=20260831-potature1",
  "./cobo-mowing-work-orders.css?v=20260901-cobo-sfalcio1",
  "./cobo-mowing-work-orders.js?v=20260902-special-terminato1",
  "./tree-work-orders.js?v=20260831-potature1",
  "./potature-followup.js?v=20260902-special-finiti-stable1",
  "./tree-search.js?v=20260901-catasto-open2",
  "./green-areas.css?v=20260829a",
  "./green-area-sheet.css?v=20260829a",
  "./green-areas.js?v=20260829e",
  "./urban-furniture.css?v=20260830-completesheet1",
  "./urban-furniture.js?v=20260830-completesheet1",
  "./wastewater-plants.css?v=20260830a",
  "./wastewater-infrastructure.css?v=20260830a",
  "./wastewater-plants.js?v=20260830b",
  "./fatto-scroll-guard.js?v=20260824-oneclick2",
  "./squad-context-bridge.js?v=20260823-stability3",
  "./recommended-traffic-weather.js?v=20260823-stability3",
  "./street-view-cards.js?v=20260830-map-gps-fix1",
  "./firebase-config.js?v=20260830-firestore-read-fix2",
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
  "/app.js",
  "/app-pure-utils.js",
  "/verde-bologna.js",
  "/verde-bologna-operativo.js",
  "/verde-bologna-parchi-mobile.js",
  "/verde-levato.js",
  "/cobo-mowing-work-orders.js",
  "/cobo-mowing-work-orders.css",
  "/data-license-view.js",
  "/potature-followup.js",
  "/green-assistant.js",
  "/green-assistant.css",
  "/android-play-store-install.js",
  "/administrative-calendar.js",
  "/fatto-button-immediate.js",
  "/shared-static-views-client.js",
  "/data-durability-runtime.js",
  "/data-safety-layer.js",
  "/critical-write-safety-bridge.js",
  "/update-app-feature.js",
  "/firebase-config.js",
  "/notification-session-enhancements.js",
  "/auth-login-fix.js",
  "/login-retry-fix.js",
  "/first-login-password.js",
  "/approval-access.js",
  "/header-menu-runtime.js",
  "/operational-import-repair.js",
  "/commessa-impianti-menu.js",
  "/commessa-produced-widget.js",
  "/loading-humor.js",
  "/client-error-reporter.js",
  "/app-error-monitor.js",
  "/admin-error-center.js",
  "/admin-error-center.css",
  "/fatto-scroll-guard.js",
  "/recommended-plants.js",
  "/recommended-plants.css",
  "/adaptive-work-learning.js",
  "/equipment-recommendations.js",
  "/recommended-traffic-weather.js",
  "/street-view-cards.js",
  "/tree-search.js",
  "/urban-furniture.js",
  "/urban-furniture.css",
  "/wastewater-plants.js",
  "/wastewater-plants.css",
  "/wastewater-infrastructure.css",
  "/squad-context-bridge.js"
]);

const isDynamicEndpoint = (url) => [/^\/api(?:\/|$)/, /^\/graphql(?:\/|$)/, /^\/auth(?:\/|$)/, /^\/socket(?:\/|$)/]
  .some((pattern) => pattern.test(url.pathname));

const hasNoStoreDirective = (headers) => {
  const value = headers.get("cache-control");
  return typeof value === "string" && value.toLowerCase().includes("no-store");
};

const shouldHandleRequest = (request, url) => request.method === "GET"
  && url.origin === self.location.origin
  && CACHEABLE_DESTINATIONS.has(request.destination)
  && !hasNoStoreDirective(request.headers)
  && !isDynamicEndpoint(url);

const canCacheResponse = (request, response) => response.ok
  && response.type !== "opaque"
  && !hasNoStoreDirective(response.headers)
  && CACHEABLE_DESTINATIONS.has(request.destination);

const fetchWithTimeout = (request, timeoutMs) => {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(request, { signal: controller.signal }).finally(() => clearTimeout(timeoutId));
};

const networkFirstForDocument = async (request) => {
  try {
    const response = await fetchWithTimeout(request, NETWORK_DOCUMENT_TIMEOUT_MS);
    if (canCacheResponse(request, response)) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(request, response.clone());
    }
    return response;
  } catch (_) {
    return (await caches.match(request)) || caches.match("./index.html");
  }
};

const networkFirstForCriticalAsset = async (request) => {
  try {
    const response = await fetchWithTimeout(request, NETWORK_DOCUMENT_TIMEOUT_MS);
    if (canCacheResponse(request, response)) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(request, response.clone());
    }
    return response;
  } catch (_) {
    return caches.match(request);
  }
};

const staleWhileRevalidateForAsset = async (event) => {
  const cached = await caches.match(event.request);
  const networkUpdate = fetch(event.request).then(async (response) => {
    if (canCacheResponse(event.request, response)) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(event.request, response.clone());
    }
    return response;
  }).catch(() => null);

  if (cached) {
    event.waitUntil(networkUpdate);
    return cached;
  }
  return (await networkUpdate)
    || (event.request.destination === "image" ? caches.match("./icons/varga-cantieri-192.png") : caches.match("./index.html"));
};

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  // Non usare skipWaiting qui: un aggiornamento automatico non deve interrompere
  // una schermata con dati ancora in compilazione. Il refresh manuale può
  // autorizzare l'attivazione dopo avere creato un backup locale.
});

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)));
    await self.clients.claim();
    const windows = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
    windows.forEach((client) => {
      try {
        // Segnala che il nuovo controller è pronto, ma NON ricarica la pagina.
        // Il refresh resta una scelta esplicita e passa dal backup dati.
        client.postMessage({
          type: "HERA_SW_UPDATE_READY",
          version: CACHE_RESET_VERSION,
          activatedAt: Date.now()
        });
      } catch (_) {}
    });
  })());
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (!shouldHandleRequest(event.request, url)) return;
  if (event.request.destination === "document") {
    event.respondWith(networkFirstForDocument(event.request));
    return;
  }
  if (NETWORK_FIRST_ASSET_PATHS.has(url.pathname)) {
    event.respondWith(networkFirstForCriticalAsset(event.request));
    return;
  }
  event.respondWith(staleWhileRevalidateForAsset(event));
});

function normalizePushPayload(event) {
  let payload = {};
  try {
    payload = event.data ? event.data.json() : {};
  } catch (_) {
    payload = { body: event.data ? event.data.text() : "" };
  }
  const notification = payload.notification || {};
  const data = payload.data || {};
  return {
    title: payload.title || notification.title || data.title || "VARGA CANTIERI",
    body: payload.body || notification.body || data.body || "Nuovo aggiornamento disponibile.",
    destination: data.destination || data.page || data.route || "home",
    url: payload.url || data.url || "./index.html",
    id: data.notificationId || data.id || String(Date.now()),
    rawData: data
  };
}

self.addEventListener("push", (event) => {
  const message = normalizePushPayload(event);
  const options = {
    body: message.body,
    icon: "./icons/varga-cantieri-192.png",
    badge: "./icons/varga-cantieri-192.png",
    tag: message.id || "hera-push-default",
    renotify: true,
    actions: [
      { action: "read", title: "LEGGI" },
      { action: "open_app", title: "APRI NELL’APP" }
    ],
    data: message
  };
  event.waitUntil(self.registration.showNotification(message.title, options));
});

self.addEventListener("notificationclick", (event) => {
  const message = event.notification.data || {};
  event.notification.close();

  if (event.action === "read" || event.action === "") {
    event.waitUntil(clients.matchAll({ type: "window", includeUncontrolled: true }).then(async (windows) => {
      for (const client of windows) client.postMessage({ type: "HERA_NOTIFICATION_READ", notification: message });
    }));
    return;
  }

  if (event.action === "open_app") {
    const targetUrl = new URL(message.url || "./index.html", self.location.origin);
    targetUrl.searchParams.set("notification", "open");
    event.waitUntil(clients.matchAll({ type: "window", includeUncontrolled: true }).then(async (windows) => {
      const existing = windows[0];
      if (existing) {
        existing.postMessage({ type: "HERA_NOTIFICATION_READ", notification: message });
        await existing.focus();
        return;
      }
      return clients.openWindow(targetUrl.href);
    }));
  }
});

self.addEventListener("sync", (event) => {
  if (event.tag !== "hera-app-background-check") return;
  event.waitUntil(self.registration.showNotification("VARGA CANTIERI", {
    body: "Controllo in background completato.",
    icon: "./icons/varga-cantieri-192.png",
    badge: "./icons/varga-cantieri-192.png",
    tag: "hera-background-sync",
    actions: [{ action: "open_app", title: "APRI NELL’APP" }],
    data: { url: "./index.html", destination: "home" }
  }));
});