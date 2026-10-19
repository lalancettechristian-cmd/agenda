/* Service worker : l'app fonctionne hors ligne. Les fichiers de l'app sont servis depuis le cache, puis mis à jour en
   arrière-plan (la nouvelle version est prise au lancement suivant). Les appels à Microsoft (connexion, OneDrive) ne
   passent jamais par le cache. Changer VERSION à chaque publication. */
const VERSION = "1.0.3";
const CACHE = "agenda-" + VERSION;
const FICHIERS = ["./", "index.html", "style.css", "app.js", "config.js", "demo.json", "manifest.webmanifest",
  "icones/apple-touch-icon.png", "icones/icone-192.png", "icones/icone-512.png", "icones/icone-masquable-512.png",
  "icones/favicon-32.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FICHIERS.map((f) => new Request(f, {cache: "reload"}))))
    .then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((cles) => Promise.all(cles.filter((k) => k.startsWith("agenda-") && k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== "GET" || url.origin !== self.location.origin) return;
  // Toute page de l'app (aussi le retour de la connexion Microsoft, avec ?code=…) est index.html.
  const cle = req.mode === "navigate" ? new URL("index.html", self.registration.scope).href : url.origin + url.pathname;
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const enCache = await cache.match(cle, {ignoreSearch: true});
    const reseau = fetch(cle, {cache: "no-cache"})
      .then(async (r) => { if (r.ok) await cache.put(cle, r.clone()); return r; })
      .catch(() => null);
    if (enCache) {
      e.waitUntil(reseau);
      return enCache;
    }
    return (await reseau) || new Response("Hors ligne", {status: 503, headers: {"Content-Type": "text/plain; charset=utf-8"}});
  })());
});
