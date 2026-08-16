/*
 * Service worker : rend l'application utilisable hors-ligne, tout en se
 * mettant à jour automatiquement dès qu'une connexion est disponible.
 *
 * Stratégie :
 * - Fichiers de code de l'application (HTML, JS, manifeste) : on essaie
 *   TOUJOURS le réseau en premier pour récupérer la dernière version ;
 *   si hors-ligne, on retombe sur la version en cache.
 * - Icônes (changent rarement) : cache en priorité.
 * Les données de l'école (élèves, notes…) sont stockées séparément via
 * localStorage (voir src/storage.js) — ce fichier ne gère que le code.
 */

const CACHE_VERSION = "ecole-app-v5";
const PRECACHE_URLS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./bundle.js",
  "./firebase-config.js",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-192.png",
  "./icons/icon-maskable-512.png",
];

const APP_SHELL_PATHS = ["./", "/", "./index.html", "./bundle.js", "./manifest.json", "./firebase-config.js"];

function isAppShellRequest(request) {
  if (request.mode === "navigate") return true;
  const url = new URL(request.url);
  return APP_SHELL_PATHS.some((p) => url.pathname.endsWith(p.replace("./", "/")) || url.pathname === "/");
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(PRECACHE_URLS)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_VERSION).map((n) => caches.delete(n)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;

  // Code de l'application : réseau en priorité (auto-mise à jour), cache en secours (hors-ligne)
  if (isAppShellRequest(request) || request.url.includes("bundle.js")) {
    event.respondWith(
      fetch(request)
        .then((response) => {
          if (response && response.status === 200) {
            const copy = response.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(request, copy));
          }
          return response;
        })
        .catch(() => caches.match(request).then((cached) => cached || caches.match("./index.html")))
    );
    return;
  }

  // Autres ressources (icônes, polices…) : cache en priorité
  event.respondWith(
    caches.match(request).then((cached) => {
      if (cached) return cached;
      return fetch(request)
        .then((response) => {
          if (response && response.status === 200 && request.url.startsWith(self.location.origin)) {
            const copy = response.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(request, copy));
          }
          return response;
        })
        .catch(() => undefined);
    })
  );
});
