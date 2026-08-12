const CACHE_NAME = "varga-cantieri-shell-v112";
const CACHE_RESET_VERSION = "20260812-opera1";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=20260727-update2",
  "./management-v2.css?v=20260731",
  "./notification-center.css?v=20260729a",
  "./approval-access.css?v=20260731-legacy1",
  "./accounting-v2.css?v=20260728",
  "./calendar-feature.css?v=20260728b",
  "./squadre-restyle.css?v=20260731-mezzi1",
  "./app.js?v=20260801-cost1",
  "./management-core.js?v=20260731",
  "./management-v2.js?v=20260731",
  "./registry-google-sheet-sync.js?v=20260802-cost2",
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
  "./header-menu-runtime.js?v=20260801-cost1",
  "./firestore-presence-cost-guard.js?v=20260802a",
  "./preventivi-lazy-loader.js?v=20260801a",
  "./global-archive-sync.js?v=20260802-cost2",
  "./global-archive-new-commesse-fix.js?v=20260801-lazy1",
  "./auto-login-saved-credentials.js?v=20260801a",
  "./varga-branding.js?v=20260731a",
  "./whazzup-preload-cache.js?v=20260803-installed-only2",
  "./fuel-stations-core.js",
  "./fuel-stations-national-cache.js",
  "./fuel-stations-search.js",
  "./fuel-stations-integration.js",
  "./registry-device-cache.js?v=20260804c",
  "./firestore-registry-read-optimizer.js?v=20260804b",
  "./firestore-safe-optimizer.js?v=20260805b",
  "./firestore-nested-listener-optimizer.js?v=20260805a",
  "./firestore-inflight-read-coalescer.js?v=20260804a",
  "./firestore-diagnostics-optimizer-extension.js?v=20260804a",
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
const NETWORK_FIRST_ASSET_PATHS = new Set(["/shared-static-views-client.js"]);

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
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)));
    await self.clients.claim();
    const windows = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
    await Promise.all(windows.map(async (client) => {
      try {
        const url = new URL(client.url);
        if (url.origin !== self.location.origin || url.searchParams.get("cacheReset") === CACHE_RESET_VERSION) return;
        url.searchParams.set("cacheReset", CACHE_RESET_VERSION);
        url.searchParams.set("cacheResetTs", String(Date.now()));
        await client.navigate(url.toString());
      } catch (_) {}
    }));
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
