/* Spațiul fraților — fișier de fundal doar pentru notificări.
   Nu interceptează rețeaua și nu păstrează nimic în cache: aplicația se încarcă exact ca înainte. */
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", event => event.waitUntil(self.clients.claim()));

self.addEventListener("push", event => {
  let p = {};
  try { p = event.data ? event.data.json() : {}; }
  catch (_) { p = { data: { body: event.data ? event.data.text() : "" } }; }
  const d = p.data || p.notification || {};
  event.waitUntil(self.registration.showNotification(d.title || "R∴L∴ Transilvania 216", {
    body: d.body || "",
    icon: "icon-192.png",
    tag: d.tag || "l216",
    data: { url: d.url || "./" }
  }));
});

self.addEventListener("notificationclick", event => {
  event.notification.close();
  const target = new URL((event.notification.data && event.notification.data.url) || "./", self.registration.scope).href;
  event.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(list => {
    for (const c of list) if (c.url.startsWith(self.registration.scope) && "focus" in c) {
      c.postMessage({ type: "l216-open", url: target });
      return c.focus();
    }
    return self.clients.openWindow(target);
  }));
});
