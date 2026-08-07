const CACHE_NAME = "varga-cantieri-shell-v112";

// Solo i file indispensabili per mostrare rapidamente login, Home e commesse.
// Un errore in un modulo accessorio non deve più bloccare l'installazione del Service Worker.
const CORE_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=20260727-update2",
  "./app.js?v=20260804-squadre-restore1",
  "./firebase-config.js",
  "./disable-app-notifications.js?v=20260807a",
  "./firestore-safe-optimizer.js?v=20260805b",
  "./firestore-inflight-read-coalescer.js?v=20260804a",
  "./native-android-runtime.js?v=20260803-whatsapp-early2",
  "./fatto-button-immediate.js?v=20260727-fatto2",
  "./manifest.webmanifest",
  "./icons/varga-cantieri-32.png",
  "./icons/varga-cantieri-180.png",
  "./icons/varga-cantieri-192.png",
  "./icons/varga-cantieri-512.png",
  "./icons/varga-cantieri-maskable-512.png"
];

// Questi file restano disponibili offline, ma vengono memorizzati senza rallentare
// l'attivazione iniziale e senza far fallire tutto se un singolo asset non risponde.
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
  "./approval-access.js?v=20260731-legacy1",
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

const canCacheResponse = (request, response) => response?.ok
  && response.type !== "opaque"
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

// Dalla seconda apertura restituisce subito la pagina salvata e aggiorna in background.
// Alla prima apertura, quando la cache non esiste, usa normalmente la rete.
const cacheFirstForDocument = async (event) => {
  const cached = await caches.match(event.request) || await caches.match("./index.html");
  const networkUpdate = updateCache(event.request);

  if (cached) {
    event.waitUntil(networkUpdate);
    return cached;
  }

  return (await networkUpdate) || new Response(
    "App temporaneamente non disponibile. Controlla la connessione e riprova.",
    { status: 503, headers: { "Content-Type": "text/plain; charset=utf-8" } }
  );
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

  return (await networkUpdate)
    || (event.request.destination === "image"
      ? caches.match("./icons/varga-cantieri-192.png")
      : caches.match("./index.html"));
};

async function cacheOptionalAssets(cache) {
  await Promise.allSettled(OPTIONAL_SHELL.map(async (url) => {
    const request = new Request(url, { cache: "reload" });
    const response = await fetch(request);
    if (response.ok && response.type !== "opaque") await cache.put(url, response);
  }));
}

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(CORE_SHELL);
    self.skipWaiting();
    await cacheOptionalAssets(cache);
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