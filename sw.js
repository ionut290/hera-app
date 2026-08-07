const CACHE_NAME = "varga-cantieri-shell-v113";

const CORE_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=20260727-update2",
  "./app.js?v=20260804-squadre-restore1",
  "./firebase-config.js",
  "./auth-login-fix.js?v=20260731-legacy1",
  "./approval-access.js?v=20260731-legacy1",
  "./first-login-password.js?v=20260726b",
  "./login-retry-fix.js?v=20260726f",
  "./persistent-offline-auth.js?v=20260807a",
  "./offline-first-runtime.js?v=20260807a",
  "./firestore-safe-optimizer.js?v=20260805b",
  "./firestore-inflight-read-coalescer.js?v=20260805a",
  "./native-android-runtime.js?v=20260803-whatsapp-early2",
  "./fatto-button-immediate.js?v=20260727-fatto2",
  "./manifest.webmanifest",
  "./icons/varga-cantieri-32.png",
  "./icons/varga-cantieri-180.png",
  "./icons/varga-cantieri-192.png",
  "./icons/varga-cantieri-512.png",
  "./icons/varga-cantieri-maskable-512.png"
];

const FIREBASE_EXTERNAL_SHELL = [
  "https://www.gstatic.com/firebasejs/8.10.1/firebase-app.js",
  "https://www.gstatic.com/firebasejs/8.10.1/firebase-auth.js",
  "https://www.gstatic.com/firebasejs/8.10.1/firebase-firestore.js",
  "https://www.gstatic.com/firebasejs/8.10.1/firebase-storage.js",
  "https://www.gstatic.com/firebasejs/8.10.1/firebase-functions.js",
  "https://www.gstatic.com/firebasejs/8.10.1/firebase-messaging.js"
];

const EXTERNAL_CACHE_ORIGINS = new Set(["https://www.gstatic.com"]);
const OPTIONAL_SHELL = [
  "./management-v2.css?v=20260731",
  "./notification-center.css?v=20260729a",
  "./approval-access.css?v=20260731-legacy1",
  "./accounting-v2.css?v=20260728",
  "./calendar-feature.css?v=20260728b",
  "./squadre-restyle.css?v=20260731-mezzi1",
  "./management-core.js?v=20260731",
  "./management-v2.js?v=20260731",
  "./registry-google-sheet-sync.js?v=20260802-cost2",
  "./coordinate-repair.js?v=20260728a",
  "./inrete-work-items-v2.js?v=20260728b",
  "./accounting-v2.js?v=20260728e",
  "./accounting-view-guard.js?v=20260728a",
  "./operational-import-repair.js?v=20260728a",
  "./google-sheet-two-way-sync.js?v=20260729b",
  "./notification-center.js?v=20260731-header1",
  "./today-summary-interactions.js?v=20260731-repair1",
  "./mezzi-alimentazione-fix.js?v=20260723",
  "./today-live-hours-vehicles.js?v=20260731-codes1",
  "./squad-operator-profile.js?v=20260731a",
  "./operator-profile-feature.js?v=20260802-cost1",
  "./personnel-training-manager.js?v=20260803a",
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
  "./firestore-nested-listener-optimizer.js?v=20260805a",
  "./firestore-diagnostics-optimizer-extension.js?v=20260804a"
];

const CACHEABLE_DESTINATIONS = new Set(["script", "style", "document", "image", "font"]);
const NETWORK_TIMEOUT_MS = 3500;
const NETWORK_FIRST_ASSET_PATHS = new Set(["/shared-static-views-client.js"]);

const isDynamicEndpoint = (url) => [/^\/api(?:\/|$)/, /^\/graphql(?:\/|$)/, /^\/auth(?:\/|$)/, /^\/socket(?:\/|$)/].some((pattern) => pattern.test(url.pathname));
const hasNoStoreDirective = (headers) => {
  const value = headers.get("cache-control");
  return typeof value === "string" && value.toLowerCase().includes("no-store");
};
const isAllowedExternalAsset = (request, url) => request.method === "GET"
  && request.destination === "script"
  && EXTERNAL_CACHE_ORIGINS.has(url.origin)
  && FIREBASE_EXTERNAL_SHELL.includes(url.href);
const shouldHandleRequest = (request, url) => request.method === "GET" && (
  (url.origin === self.location.origin && CACHEABLE_DESTINATIONS.has(request.destination) && !hasNoStoreDirective(request.headers) && !isDynamicEndpoint(url))
  || isAllowedExternalAsset(request, url)
);
const canCacheResponse = (request, response) => Boolean(response)
  && (response.ok || response.type === "opaque")
  && !hasNoStoreDirective(response.headers)
  && CACHEABLE_DESTINATIONS.has(request.destination);

const fetchWithTimeout = (request, timeoutMs = NETWORK_TIMEOUT_MS) => {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(request, { signal: controller.signal }).finally(() => clearTimeout(timeoutId));
};

async function updateCache(request) {
  try {
    const response = await fetch(request);
    if (canCacheResponse(request, response)) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(request, response.clone());
    }
    return response;
  } catch (_) {
    return null;
  }
}

const cacheFirstForDocument = async (event) => {
  const cached = await caches.match(event.request) || await caches.match("./index.html") || await caches.match("./");
  const networkUpdate = updateCache(event.request);
  if (cached) {
    event.waitUntil(networkUpdate);
    return cached;
  }
  return (await networkUpdate) || new Response("App offline non ancora preparata. Apri una volta l'app con Internet e riprova.", { status: 503, headers: { "Content-Type": "text/plain; charset=utf-8" } });
};

const networkFirstForCriticalAsset = async (request) => {
  try {
    const response = await fetchWithTimeout(request);
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
  const networkUpdate = updateCache(event.request);
  if (cached) {
    event.waitUntil(networkUpdate);
    return cached;
  }
  return (await networkUpdate) || (event.request.destination === "image" ? caches.match("./icons/varga-cantieri-192.png") : Response.error());
};

const cacheFirstForExternalAsset = async (event) => {
  const cached = await caches.match(event.request);
  const networkUpdate = updateCache(event.request);
  if (cached) {
    event.waitUntil(networkUpdate);
    return cached;
  }
  return (await networkUpdate) || Response.error();
};

async function cacheOne(cache, url) {
  try {
    const request = new Request(url, { cache: "reload" });
    const response = await fetch(request);
    if (response && response.ok) {
      await cache.put(request, response.clone());
      return true;
    }
  } catch (_) {}
  return false;
}

async function cacheCoreAssets(cache) {
  const results = await Promise.allSettled(CORE_SHELL.map((url) => cacheOne(cache, url)));
  const indexCached = await caches.match("./index.html") || await caches.match("./");
  if (!indexCached) throw new Error("Impossibile preparare index.html per l'avvio offline");
  return results;
}

async function cacheOptionalAssets(cache) {
  await Promise.allSettled(OPTIONAL_SHELL.map((url) => cacheOne(cache, url)));
}

async function cacheFirebaseExternalAssets(cache) {
  await Promise.allSettled(FIREBASE_EXTERNAL_SHELL.map(async (url) => {
    try {
      const request = new Request(url, { mode: "no-cors", cache: "reload" });
      const response = await fetch(request);
      if (response && (response.ok || response.type === "opaque")) await cache.put(request, response);
    } catch (_) {}
  }));
}

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cacheCoreAssets(cache);
    await Promise.allSettled([cacheOptionalAssets(cache), cacheFirebaseExternalAssets(cache)]);
    self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)));
    if (self.registration.navigationPreload) {
      try { await self.registration.navigationPreload.enable(); } catch (_) {}
    }
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (!shouldHandleRequest(event.request, url)) return;
  if (isAllowedExternalAsset(event.request, url)) {
    event.respondWith(cacheFirstForExternalAsset(event));
    return;
  }
  if (event.request.destination === "document") {
    event.respondWith(cacheFirstForDocument(event));
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
  try { payload = event.data ? event.data.json() : {}; }
  catch (_) { payload = { body: event.data ? event.data.text() : "" }; }
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
  event.waitUntil(self.registration.showNotification(message.title, {
    body: message.body,
    icon: "./icons/varga-cantieri-192.png",
    badge: "./icons/varga-cantieri-192.png",
    tag: message.id || "hera-push-default",
    renotify: true,
    actions: [{ action: "read", title: "LEGGI" }, { action: "open_app", title: "APRI NELL’APP" }],
    data: message
  }));
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