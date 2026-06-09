// SOLTRI 피복관리 Service Worker — 오프라인 셸 캐싱
// v6: 오프라인 PDF 처리 + uploadPdf 멱등성
const CACHE = 'soltri-clothing-v7';
// GitHub Pages 하위경로(/clothing/) 대응 — 상대경로로 캐싱
const SHELL = ['./', './index.html', './manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    // 각 자원을 개별적으로 캐싱 → 하나 실패해도 나머지 계속
    await Promise.all(SHELL.map(url =>
      cache.add(url).catch(err => console.warn('[SW] cache fail', url, err))
    ));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // API 요청은 SW 우회 — 네트워크 실패는 앱이 큐에 저장
  if (url.pathname.startsWith('/api/')) return;
  // 다른 오리진(폰트/jsPDF CDN)은 브라우저 기본 캐시에 맡김
  if (url.origin !== location.origin) return;

  e.respondWith((async () => {
    // 1) 캐시 우선
    const cached = await caches.match(req);
    if (cached) {
      // 백그라운드에서 최신본 받아서 캐시 갱신 (stale-while-revalidate)
      fetch(req).then(resp => {
        if (resp && resp.ok) caches.open(CACHE).then(c => c.put(req, resp.clone()));
      }).catch(() => {});
      return cached;
    }
    // 2) 네트워크 시도 후 성공시 캐시에도 저장
    try {
      const resp = await fetch(req);
      if (resp && resp.ok) {
        const clone = resp.clone();
        caches.open(CACHE).then(c => c.put(req, clone));
      }
      return resp;
    } catch (err) {
      // 3) 네트워크 실패 + 캐시 없음 → navigation이면 index.html로 폴백
      if (req.mode === 'navigate' || (req.destination === 'document')) {
        const fallback = (await caches.match('./index.html')) || (await caches.match('./'));
        if (fallback) return fallback;
      }
      throw err;
    }
  })());
});
