const CACHE_NAME = "hera-app-shell-v13";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./app.js?v=20260705",
  "./hours-export-range.js",
  "./hours-export-range.js?v=20260619",
  "./firebase-config.js",
  "./manifest.webmanifest",
  "./offline.html",
  "./icons/hera-icon.svg"
];

const CACHEABLE_DESTINATIONS = new Set(["script", "style", "document", "image", "font"]);
const OPAQUE_CACHE_WHITELIST = new Set(["https://www.gstatic.com", "https://cdn.jsdelivr.net", "https://unpkg.com"]);
const NETWORK_DOCUMENT_TIMEOUT_MS = 3500;
const MAX_DYNAMIC_CACHE_ENTRIES = 80;

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
  const isSameOrigin = url.origin === self.location.origin;
  const isStaticCdn = OPAQUE_CACHE_WHITELIST.has(url.origin);
  if (!isSameOrigin && !isStaticCdn) return false;
  if (!CACHEABLE_DESTINATIONS.has(request.destination)) return false;
  if (hasNoStoreDirective(request.headers)) return false;
  if (isDynamicEndpoint(url)) return false;
  return true;
};

const canCacheResponse = (request, response, url) => {
  if (/tile.openstreetmap.org|tilecache.rainviewer.com|googleapis.com|gstatic.com/.test(url.hostname)) return false;
  if (response.type === "opaque") {
    return OPAQUE_CACHE_WHITELIST.has(url.origin);
  }

  if (!response.ok) return false;
  if (hasNoStoreDirective(response.headers)) return false;

  return CACHEABLE_DESTINATIONS.has(request.destination);
};

const trimCache = async () => {
  const cache = await caches.open(CACHE_NAME);
  const keys = await cache.keys();
  if (keys.length <= MAX_DYNAMIC_CACHE_ENTRIES) return;
  await Promise.all(keys.slice(0, keys.length - MAX_DYNAMIC_CACHE_ENTRIES).map((key) => cache.delete(key)));
};

const networkFirstForDocument = async (request) => {
  try {
    const response = await fetchWithTimeout(request, NETWORK_DOCUMENT_TIMEOUT_MS);
    const requestUrl = new URL(request.url);
    if (canCacheResponse(request, response, requestUrl)) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(request, response.clone());
      await trimCache();
    }
    return response;
  } catch (error) {
    const cached = await caches.match(request);
    if (cached) return cached;
    return caches.match("./index.html").then((cachedIndex) => cachedIndex || caches.match("./offline.html"));
  }
};

const fetchWithTimeout = (request, timeoutMs) => {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(request, { signal: controller.signal }).finally(() => {
    clearTimeout(timeoutId);
  });
};

const cacheFirstForAsset = async (event) => {
  const { request } = event;
  const cached = await caches.match(request, { ignoreSearch: request.url.startsWith(self.location.origin) });
  if (cached) return cached;

  const networkUpdate = fetch(request)
    .then(async (response) => {
      const requestUrl = new URL(request.url);
      if (canCacheResponse(request, response, requestUrl)) {
        const cache = await caches.open(CACHE_NAME);
        await cache.put(request, response.clone());
        await trimCache();
      }
      return response;
    })
    .catch(() => null);

  const response = await networkUpdate;
  if (response) return response;

  if (request.destination === "image") {
    return caches.match("./icons/hera-icon.svg");
  }

  return caches.match("./index.html").then((cachedIndex) => cachedIndex || caches.match("./offline.html"));
};

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
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

  event.respondWith(cacheFirstForAsset(event));
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
    icon: payload.icon || "./icons/hera-icon.svg",
    badge: payload.badge || "./icons/hera-icon.svg",
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
      icon: "./icons/hera-icon.svg",
      badge: "./icons/hera-icon.svg",
      tag: "hera-background-sync",
      data: { url: "./index.html" }
    })
  );
});
