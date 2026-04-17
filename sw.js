const CACHE_NAME = "hera-app-shell-v4";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./firebase-config.js",
  "./manifest.webmanifest",
  "./icons/hera-icon.svg"
];

const CACHEABLE_DESTINATIONS = new Set(["script", "style", "document", "image", "font"]);
const OPAQUE_CACHE_WHITELIST = new Set([]);

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
    const response = await fetch(request);
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
    return caches.match("./icons/hera-icon.svg");
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
