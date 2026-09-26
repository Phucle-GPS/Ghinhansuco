/* CSKVTT Quy trình – service worker: chạy offline, nhận file chia sẻ, thông báo. Chỉ phụ trách quytrinh.html */
const VER = 'qt-v1';
const SHELL = ['quytrinh.html', 'quytrinh-manifest.json', 'qt-icon-192.png', 'qt-icon-512.png', 'qt-icon-maskable.png', 'qt-apple-icon.png'];
const CDN = ['cdnjs.cloudflare.com', 'cdn.jsdelivr.net', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VER).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('qt-') && k !== VER && k !== VER + '-cdn' && k !== 'qt-share').map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('message', e => { if (e.data === 'skipWaiting') self.skipWaiting(); });

self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  // Nhận ảnh, file chia sẻ từ Zalo, thư viện ảnh...
  if (req.method === 'POST' && url.pathname.endsWith('/quytrinh.html') && url.searchParams.has('share')) {
    e.respondWith((async () => {
      try {
        const fd = await req.formData(), c = await caches.open('qt-share'), fs = fd.getAll('files');
        await Promise.all(fs.map((f, i) => c.put(new Request('share/' + Date.now() + '_' + i), new Response(f, { headers: { 'content-type': f.type || 'application/octet-stream', 'x-name': encodeURIComponent(f.name || ('file' + i)) } }))));
      } catch (err) { }
      return Response.redirect(new URL('quytrinh.html?share=1', self.registration.scope).href, 303);
    })());
    return;
  }
  if (req.method !== 'GET') return;
  if (url.hostname.endsWith('script.google.com') || url.hostname.endsWith('googleusercontent.com') || url.hostname === 'api.anthropic.com') return;
  // Trang app: lấy bản mới khi có mạng, mất mạng thì dùng bản đã lưu
  if (req.mode === 'navigate' || (url.origin === location.origin && url.pathname.endsWith('/quytrinh.html'))) {
    e.respondWith(fetch(req).then(r => { if (r.ok) { const cp = r.clone(); caches.open(VER).then(c => c.put('quytrinh.html', cp)); } return r; })
      .catch(() => caches.match('quytrinh.html', { ignoreSearch: true })));
    return;
  }
  // Thư viện, phông chữ: dùng bản lưu, cập nhật ngầm
  if (CDN.includes(url.hostname)) {
    e.respondWith(caches.open(VER + '-cdn').then(async c => {
      const hit = await c.match(req);
      const net = fetch(req).then(r => { if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    }));
    return;
  }
  if (url.origin === location.origin) {
    e.respondWith(caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req)));
  }
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  const v = (e.notification.data && e.notification.data.v) || 's1022';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => {
    const c = cs.find(x => x.url.includes('quytrinh'));
    if (c) { c.postMessage({ go: v }); return c.focus(); }
    return self.clients.openWindow(new URL('quytrinh.html?v=' + v, self.registration.scope).href);
  }));
});
