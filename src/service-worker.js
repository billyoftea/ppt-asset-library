/**
 * Service Worker - 离线缓存策略
 * 
 * 策略：Cache First, Network Fallback
 * - 首次在线加载时缓存所有核心资源
 * - 之后优先从缓存读取，无需网络
 * - 有网时后台静默更新缓存
 */

const CACHE_NAME = "ppt-asset-library-v1";

// 核心资源列表 — 首次安装时预缓存
const PRECACHE_URLS = [
  "./taskpane.html",
  "./taskpane.bundle.js",
  "./data/assetIndex.json",
  "./data/slideIds.json",
  "./assets/bundle.pptx",
];

// 需要缓存的资源路径前缀（运行时动态缓存）
const CACHEABLE_PATHS = [
  "/assets/thumbnails/",
  "/assets/hd/",
  "/assets/slides/",
  "/assets/",
  "/data/",
];

/**
 * Install 事件 — 预缓存核心资源
 */
self.addEventListener("install", (event) => {
  console.log("[SW] Installing Service Worker...");
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then((cache) => {
        console.log("[SW] Pre-caching core resources...");
        return cache.addAll(PRECACHE_URLS);
      })
      .then(() => {
        console.log("[SW] Pre-cache complete!");
        // 跳过等待，立即激活
        return self.skipWaiting();
      })
      .catch((err) => {
        console.error("[SW] Pre-cache failed:", err);
        // 即使预缓存失败也激活，运行时再缓存
        return self.skipWaiting();
      })
  );
});

/**
 * Activate 事件 — 清理旧缓存
 */
self.addEventListener("activate", (event) => {
  console.log("[SW] Activating Service Worker...");
  event.waitUntil(
    caches
      .keys()
      .then((cacheNames) => {
        return Promise.all(
          cacheNames
            .filter((name) => name !== CACHE_NAME)
            .map((name) => {
              console.log("[SW] Deleting old cache:", name);
              return caches.delete(name);
            })
        );
      })
      .then(() => {
        // 接管所有打开的页面
        return self.clients.claim();
      })
  );
});

/**
 * Fetch 事件 — Cache First 策略
 * 
 * 1. 对于 Office.js CDN 请求 → 始终走网络（不缓存第三方 CDN）
 * 2. 对于本地资源 → 先查缓存，有则返回；无则走网络并缓存
 * 3. 网络请求成功后，后台更新缓存（Stale While Revalidate 混合模式）
 */
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // 不缓存 Office.js CDN 和其他第三方资源
  if (url.origin !== self.location.origin) {
    return;
  }

  // 不缓存非 GET 请求
  if (event.request.method !== "GET") {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        // 缓存命中 — 返回缓存，同时后台更新
        refreshCache(event.request);
        return cachedResponse;
      }

      // 缓存未命中 — 走网络并缓存
      return fetchAndCache(event.request);
    })
  );
});

/**
 * 网络请求并缓存结果
 */
async function fetchAndCache(request) {
  try {
    const response = await fetch(request);

    // 只缓存成功的响应
    if (response.ok) {
      const responseClone = response.clone();
      caches.open(CACHE_NAME).then((cache) => {
        cache.put(request, responseClone);
      });
    }

    return response;
  } catch (err) {
    // 网络失败且无缓存 — 返回离线提示页面（仅对 HTML 请求）
    if (request.headers.get("Accept")?.includes("text/html")) {
      return new Response(
        `<!DOCTYPE html>
        <html><head><meta charset="utf-8"><title>Offline</title></head>
        <body style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif;color:#666;">
          <div style="text-align:center;">
            <h2>📴 当前处于离线状态</h2>
            <p>请连接网络后重试。首次使用需联网加载资源。</p>
            <button onclick="location.reload()" style="padding:8px 24px;border-radius:6px;border:1px solid #ccc;cursor:pointer;">重试</button>
          </div>
        </body></html>`,
        { headers: { "Content-Type": "text/html; charset=utf-8" } }
      );
    }

    throw err;
  }
}

/**
 * 后台静默更新缓存（Stale While Revalidate）
 */
async function refreshCache(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(request, response);
    }
  } catch {
    // 后台更新失败没关系，继续用缓存
  }
}

/**
 * 监听来自主页面的消息
 */
self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }

  // 支持手动触发缓存所有缩略图
  if (event.data && event.data.type === "CACHE_THUMBNAILS") {
    const urls = event.data.urls || [];
    caches.open(CACHE_NAME).then((cache) => {
      console.log(`[SW] Caching ${urls.length} thumbnails...`);
      return Promise.allSettled(
        urls.map((url) =>
          cache.match(url).then((existing) => {
            if (!existing) {
              return cache.add(url).catch(() => {});
            }
          })
        )
      );
    });
  }
});
