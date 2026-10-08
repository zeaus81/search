const CACHE = "upche-shell-9d9448298e4f";
const SHELL = ["./", "manifest.webmanifest", "icon-192.png", "icon-512.png", "apple-touch-icon.png"];
self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k.startsWith("upche-shell-") && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  if (/version\.json$|data(\.gz)?\.bin$/.test(u.pathname)) return;  // 자료는 앱이 직접 받아 보관한다
  // 서버가 "10분 동안은 저장해 둔 것을 써도 된다"고 알려 주기 때문에, 그냥 받으면 새 화면을 올려도 한동안 예전 화면이 나온다.
  // 매번 서버에 바뀌었는지 물어보게 해서(no-cache) 새 화면이 바로 보이게 한다.
  e.respondWith(fetch(e.request, {cache: "no-cache"}).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request, {ignoreSearch: true}).then(r => r || caches.match("./"))));
});
