const CACHE_NAME = "hera-app-shell-v43";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=20260727-update2",
  "./calendar-feature.css?v=20260727a",
  "./squadre-restyle.css?v=20260728c",
  "./app.js?v=20260728-today3",
  "./native-android-runtime.js?v=20260726-fatto1",
  "./notification-session-enhancements.js?v=20260727b",
  "./today-summary-interactions.js?v=20260728f",
  "./fatto-button-immediate.js?v=20260727-fatto2",
  "./fuel-stations-core.js",
  "./fuel-stations-national-cache.js",
  "./fuel-stations-search.js",
  "./fuel-stations-integration.js",
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
  event.waitUntil(caches.keys().then((keys) => Promise.all(
    keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
  )));
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (!shouldHandleRequest(event.request, url)) return;
  if (event.request.destination === "document") {
    event.respondWith(networkFirstForDocument(event.request));
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
    title: payload.title || notification.title || data.title || "Varga Cantieri",
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
      for (const client of windows) {
        client.postMessage({ type: "HERA_NOTIFICATION_READ", notification: message });
      }
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
  event.waitUntil(self.registration.showNotification("Varga Cantieri", {
    body: "Controllo in background completato.",
    icon: "./icons/varga-cantieri-192.png",
    badge: "./icons/varga-cantieri-192.png",
    tag: "hera-background-sync",
    actions: [{ action: "open_app", title: "APRI NELL’APP" }],
    data: { url: "./index.html", destination: "home" }
  }));
});
