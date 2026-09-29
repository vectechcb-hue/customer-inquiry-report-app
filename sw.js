const CACHE_VERSION = "v60";
const CACHE_NAME = "vectech-customer-inquiry-" + CACHE_VERSION;
const APP_SHELL = [
  "./",
  "./index.html",
  "./login.html",
  "./manifest.webmanifest",
  "./app.js?v=20260929v60"
];

self.addEventListener("install", function(event){
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache){ return cache.addAll(APP_SHELL); })
      .then(function(){ return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function(event){
  event.waitUntil(
    caches.keys()
      .then(function(keys){
        return Promise.all(
          keys
            .filter(function(key){ return key.indexOf("vectech-customer-inquiry-") === 0 && key !== CACHE_NAME; })
            .map(function(key){ return caches.delete(key); })
        );
      })
      .then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function(event){
  if(event.request.method !== "GET") return;
  const url = new URL(event.request.url);
  if(url.origin !== self.location.origin) return;
  if(url.pathname.endsWith("/sw.js")) return;

  event.respondWith(
    fetch(event.request)
      .then(function(response){
        if(response && response.ok){
          const copy = response.clone();
          caches.open(CACHE_NAME).then(function(cache){ cache.put(event.request, copy); }).catch(function(){});
        }
        return response;
      })
      .catch(function(){
        return caches.match(event.request);
      })
  );
});