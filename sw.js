// Service Worker بسيط: يخلي الموقع يتثبت كتطبيق، ويفتح حتى لو النت ضعيف.
// بيانات جوجل شيت (script.google.com) عمرها ما بتتخزن عشان الداتا تكون دايماً حديثة.
const CACHE = "crm-shell-v1";
const SHELL = ["./", "index.html", "manifest.json", "icons/icon-192.png", "icons/icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  // لا نتدخل في طلبات قاعدة البيانات أبداً
  if (url.hostname.endsWith("google.com") || url.hostname.endsWith("googleusercontent.com")) return;

  // الصفحة نفسها: الشبكة أولاً (علشان التحديثات تظهر)، والكاش احتياطي
  if (req.mode === "navigate") {
    e.respondWith(fetch(req).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put("index.html", copy));
      return res;
    }).catch(() => caches.match("index.html")));
    return;
  }
  // باقي الملفات (مكتبات، أيقونات): الكاش أولاً
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  })));
});

// الضغط على إشعار المتابعة يفتح التطبيق
self.addEventListener("notificationclick", e => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: "window" }).then(list => {
    if (list.length) return list[0].focus();
    return self.clients.openWindow("./");
  }));
});
