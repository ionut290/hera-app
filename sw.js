const CACHE_NAME = "hera-app-shell-v33";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=20260726-splash1",
  "./app.js?v=20260726-fatto1",
  "./native-android-runtime.js?v=20260726-fatto1",
  "./today-summary-interactions.js?v=20260726c",
  "./fatto-button-immediate.js?v=20260726-fatto1",
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
const OPAQUE_CACHE_WHITELIST = new Set([]);
const NETWORK_DOCUMENT_TIMEOUT_MS = 7000;

const isDynamicEndpoint = (url) => {
  const dynamicPathPatterns = [/^\/api(?:\/|$)/, /^\/graphql(?:\/|$)/, /^\/auth(?:\/|$)/, /^\/socket(?:\/|$)/];
  return dynamicPathPatterns.some((pattern) => pattern.test(url.pathname));
};

const hasNoStoreDirective = (headers) => {
  const cacheControl = headers.get("cache-control");
  return typeof cacheControl === "string" && cacheControl.toLowerCase().includes("no-store");
};

const shouldHandleRequest = (request, url) => {
  if (request.method !== "GET") return false;
  if (url.origin !== self.location.origin) return false;
  if (!CACHEABLE_DESTINATIONS.has(request.destination)) return false;
  if (hasNoStoreDirective(request.headers)) return false;
  if (isDynamicEndpoint(url)) return false;
  return true;
};

const canCacheResponse = (request, response, url) => {
  if (response.type === "opaque") {
    return OPAQUE_CACHE_WHITELIST.has(url.origin);
  }

  if (!response.ok) return false;
  if (hasNoStoreDirective(response.headers)) return false;

  return CACHEABLE_DESTINATIONS.has(request.destination);
};

const networkFirstForDocument = async (request) => {
  try {
    const response = await fetchWithTimeout(request, NETWORK_DOCUMENT_TIMEOUT_MS);
    const requestUrl = new URL(request.url);
    if (canCacheResponse(request, response, requestUrl)) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(request, response.clone());
    }
    return response;
  } catch (error) {
    const cached = await caches.match(request);
    if (cached) return cached;
    return caches.match("./index.html");
  }
};

const fetchWithTimeout = (request, timeoutMs) => {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(request, { signal: controller.signal }).finally(() => {
    clearTimeout(timeoutId);
  });
};

const staleWhileRevalidateForAsset = async (event) => {
  const { request } = event;
  const cached = await caches.match(request);

  const networkUpdate = fetch(request)
    .then(async (response) => {
      const requestUrl = new URL(request.url);
      if (canCacheResponse(request, response, requestUrl)) {
        const cache = await caches.open(CACHE_NAME);
        await cache.put(request, response.clone());
      }
      return response;
    })
    .catch(() => null);

  if (cached) {
    event.waitUntil(networkUpdate);
    return cached;
  }

  const response = await networkUpdate;
  if (response) return response;

  if (request.destination === "image") {
    return caches.match("./icons/varga-cantieri-192.png");
  }

  return caches.match("./index.html");
};

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  const url = new URL(request.url);

  if (!shouldHandleRequest(request, url)) {
    return;
  }

  if (request.destination === "document") {
    event.respondWith(networkFirstForDocument(request));
    return;
  }

  event.respondWith(staleWhileRevalidateForAsset(event));
});

self.addEventListener("push", (event) => {
  let payload = {};
  try {
    payload = event.data ? event.data.json() : {};
  } catch (error) {
    payload = { body: event.data ? event.data.text() : "" };
  }
  const title = payload.title || "Hera App";
  const options = {
    body: payload.body || "Nuovo aggiornamento disponibile.",
    icon: payload.icon || "./icons/varga-cantieri-192.png",
    badge: payload.badge || "./icons/varga-cantieri-192.png",
    tag: payload.tag || "hera-push-default",
    data: {
      url: payload.url || "./index.html"
    }
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = (event.notification.data && event.notification.data.url) || "./index.html";
  event.waitUntil(
    clients.matchAll({ type: "window", includeUncontrolled: true }).then((windows) => {
      const existing = windows.find((win) => win.url.includes("index.html"));
      if (existing) return existing.focus();
      return clients.openWindow(targetUrl);
    })
  );
});

self.addEventListener("sync", (event) => {
  if (event.tag !== "hera-app-background-check") return;
  event.waitUntil(
    self.registration.showNotification("Hera App", {
      body: "Controllo in background completato.",
      icon: "./icons/varga-cantieri-192.png",
      badge: "./icons/varga-cantieri-192.png",
      tag: "hera-background-sync",
      data: { url: "./index.html" }
    })
  );
});
