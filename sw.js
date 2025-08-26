// sw.js — 강제 최신화용 서비스워커
const APP_VERSION = "2025-08-25-03";
const CACHE_NAME = "rent-label-cache-" + APP_VERSION;

// 즉시 컨트롤
self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME));
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.map((k) => {
      if (k !== CACHE_NAME) return caches.delete(k);
    }));
    await self.clients.claim();
  })());
});

// 네트워크 우선, 실패 시 캐시
self.addEventListener("fetch", (event) => {
  const req = event.request;

  // 항상 네트워크 우선으로 최신 가져오기
  event.respondWith((async () => {
    try {
      const fresh = await fetch(req, { cache: "no-store" });
      // 정적 파일만 캐싱(HTML 제외)
      if (req.method === "GET" && !req.headers.get("accept")?.includes("text/html")) {
        const cache = await caches.open(CACHE_NAME);
        cache.put(req, fresh.clone());
      }
      return fresh;
    } catch {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(req);
      if (cached) return cached;
      // 마지막 fallback
      return new Response("Offline", { status: 503, statusText: "Offline" });
    }
  })());
});

