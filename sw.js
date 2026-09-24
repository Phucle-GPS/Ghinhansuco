// v3 – luôn lấy bản mới nhất từ mạng (không dùng bộ nhớ đệm trình duyệt), chỉ dùng bản lưu khi mất mạng
const CACHE = "gnsc-v3";
const ASSETS = ["./","./index.html","./login.html","./cabinets.js","./config.js","./manifest.json","./icon-192.png","./icon-512.png"];
self.addEventListener("install", e=>{ e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS.map(u=>new Request(u,{cache:"reload"})))).then(()=>self.skipWaiting())); });
self.addEventListener("activate", e=>{ e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())); });
self.addEventListener("fetch", e=>{
  const u = new URL(e.request.url);
  if(e.request.method!=="GET" || u.origin!==location.origin) return;
  e.respondWith(
    fetch(e.request, {cache:"no-store"})
      .then(r=>{ const cp=r.clone(); caches.open(CACHE).then(c=>c.put(e.request,cp)); return r; })
      .catch(()=>caches.match(e.request, {ignoreSearch:true}))
  );
});
