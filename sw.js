"use strict";

var CACHE = "twenty-study-os-v11-git-sync";
var APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css?v=11-git-sync",
  "./sync.js?v=11-git-sync",
  "./db.js?v=11-git-sync",
  "./app.js?v=11-git-sync",
  "./lucide.min.js?v=11-git-sync",
  "./manifest.webmanifest",
  "./icon.svg",
  "./icon-192.png",
  "./icon-512.png",
  "./data/academic-data.json",
  "./data/canteen-menu.json"
];

self.addEventListener("install", function (event) {
  event.waitUntil(caches.open(CACHE).then(function (cache) { return cache.addAll(APP_SHELL); }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener("activate", function (event) {
  event.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (key) { return key !== CACHE; }).map(function (key) { return caches.delete(key); }));
  }).then(function () { return self.clients.claim(); }));
});

function networkFirst(request) {
  return fetch(request).then(function (response) {
    if (response && response.ok) {
      var copy = response.clone();
      caches.open(CACHE).then(function (cache) { cache.put(request, copy); });
    }
    return response;
  }).catch(function () { return caches.match(request); });
}

function cacheFirst(request) {
  return caches.match(request).then(function (cached) {
    if (cached) return cached;
    return fetch(request).then(function (response) {
      if (response && response.ok && request.method === "GET") {
        var copy = response.clone();
        caches.open(CACHE).then(function (cache) { cache.put(request, copy); });
      }
      return response;
    });
  });
}

self.addEventListener("fetch", function (event) {
  var request = event.request;
  if (request.method !== "GET" || request.url.indexOf("blob:") === 0) return;
  var url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname.indexOf("/data/") >= 0) {
    event.respondWith(networkFirst(request));
  } else if (request.mode === "navigate") {
    event.respondWith(networkFirst(request).then(function (response) { return response || caches.match("./index.html"); }));
  } else {
    event.respondWith(cacheFirst(request));
  }
});
