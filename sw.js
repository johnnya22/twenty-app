"use strict";

var CACHE = "twenty-study-os-v18-lesson-ai";
var APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css?v=18-lesson-ai",
  "./sync.js?v=18-lesson-ai",
  "./db.js?v=18-lesson-ai",
  "./ai.js?v=18-lesson-ai",
  "./ai-worker.js?v=18-lesson-ai",
  "./app.js?v=18-lesson-ai",
  "./lucide.min.js?v=18-lesson-ai",
  "./manifest.webmanifest",
  "./icon.svg",
  "./icon-192.png",
  "./icon-512.png",
  "./data/academic-data.json",
  "./data/canteen-menu.json"
];

var SYNC_DB = "twenty-study-sync";
var SYNC_DB_VERSION = 1;
var APP_DB = "twenty-study-os";
var APP_DB_VERSION = 2;
var APP_STATE_KEY = "academic-state";

self.addEventListener("install", function (event) {
  event.waitUntil(caches.open(CACHE).then(function (cache) {
    return cache.addAll(APP_SHELL);
  }).then(function () {
    return self.skipWaiting();
  }));
});

self.addEventListener("activate", function (event) {
  event.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (key) { return key !== CACHE; }).map(function (key) { return caches.delete(key); }));
  }).then(function () {
    return self.clients.claim();
  }).then(function () {
    return backgroundUpdate().catch(function () {});
  }));
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
  } else if (/\.(?:js|css|html)$/.test(url.pathname)) {
    event.respondWith(networkFirst(request));
  } else {
    event.respondWith(cacheFirst(request));
  }
});

function openDb(name, version, upgrade) {
  return new Promise(function (resolve, reject) {
    var request = indexedDB.open(name, version);
    request.onupgradeneeded = function () {
      if (upgrade) upgrade(request.result);
    };
    request.onsuccess = function () { resolve(request.result); };
    request.onerror = function () { reject(request.error); };
  });
}

function withStore(db, storeName, mode, action) {
  return new Promise(function (resolve, reject) {
    var tx = db.transaction(storeName, mode);
    var store = tx.objectStore(storeName);
    var request;
    try { request = action(store); } catch (error) { reject(error); return; }
    if (request) {
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error); };
    } else {
      tx.oncomplete = function () { resolve(); };
    }
    tx.onerror = function () { reject(tx.error); };
  });
}

async function openSyncDb() {
  return openDb(SYNC_DB, SYNC_DB_VERSION, function (db) {
    if (!db.objectStoreNames.contains("meta")) db.createObjectStore("meta");
    if (!db.objectStoreNames.contains("queue")) db.createObjectStore("queue", { keyPath: "id" });
  });
}

async function openAppDb() {
  return openDb(APP_DB, APP_DB_VERSION, function (db) {
    if (!db.objectStoreNames.contains("kv")) db.createObjectStore("kv");
    if (!db.objectStoreNames.contains("files")) db.createObjectStore("files", { keyPath: "id" });
  });
}

async function syncMetaGet(key) {
  var db = await openSyncDb();
  return withStore(db, "meta", "readonly", function (store) { return store.get(key); });
}

async function syncMetaPut(key, value) {
  var db = await openSyncDb();
  return withStore(db, "meta", "readwrite", function (store) { return store.put(value, key); });
}

async function queueCount() {
  var db = await openSyncDb();
  return withStore(db, "queue", "readonly", function (store) { return store.count(); });
}

async function appStateGet() {
  var db = await openAppDb();
  return withStore(db, "kv", "readonly", function (store) { return store.get(APP_STATE_KEY); });
}

async function appStatePut(state) {
  var db = await openAppDb();
  return withStore(db, "kv", "readwrite", function (store) { return store.put(state, APP_STATE_KEY); });
}

function stateVersion(state) {
  var meta = state && state.meta && typeof state.meta === "object" ? state.meta : {};
  var value = Number(meta.gitRevision != null ? meta.gitRevision : meta.revision);
  return Number.isFinite(value) && value > 0 ? value : 0;
}

async function workerFetch(config, path) {
  var response = await fetch(String(config.endpoint || "").replace(/\/+$/, "") + path, {
    cache: "no-store",
    headers: {
      "Accept": "application/json",
      "Content-Type": "application/json",
      "X-Twenty-Key": String(config.key || "")
    }
  });
  if (!response.ok) throw new Error("Git background sync " + response.status);
  return response.json();
}

async function notifyClients(payload) {
  var windows = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
  windows.forEach(function (client) { client.postMessage(payload); });
  return windows.length;
}

async function backgroundUpdate() {
  var config = await syncMetaGet("serviceWorkerConfig");
  if (!config || config.enabled === false || !config.endpoint || !config.key) return false;
  if ((await queueCount()) > 0) return false;

  var versionPayload = await workerFetch(config, "/version");
  if (!versionPayload.exists) return false;
  var localState = await appStateGet();
  var localVersion = stateVersion(localState);
  var lastSha = await syncMetaGet("lastSha") || "";
  var remoteVersion = Number(versionPayload.version) || 0;
  var changed = !localState || remoteVersion > localVersion || (!!lastSha && !!versionPayload.sha && versionPayload.sha !== lastSha);
  if (!changed) {
    await syncMetaPut("lastBackgroundCheckAt", new Date().toISOString());
    return false;
  }

  var openClients = await notifyClients({
    type: "twenty:background-update-ready",
    version: remoteVersion,
    sha: versionPayload.sha || ""
  });
  if (openClients) return true;

  // Sem janelas abertas, o Service Worker atualiza diretamente o estado local.
  // Só o faz quando não existem alterações pendentes, para nunca apagar trabalho offline.
  var statePayload = await workerFetch(config, "/state");
  if (!statePayload.exists || !statePayload.state) return false;
  await appStatePut(statePayload.state);
  await Promise.all([
    syncMetaPut("lastRemote", statePayload.state),
    syncMetaPut("shadow", statePayload.state),
    syncMetaPut("lastSha", statePayload.sha || ""),
    syncMetaPut("lastVersion", Number(statePayload.version || stateVersion(statePayload.state)) || 0),
    syncMetaPut("lastBackgroundSyncAt", new Date().toISOString())
  ]);
  return true;
}

self.addEventListener("periodicsync", function (event) {
  if (event.tag === "twenty-git-background-v1") {
    event.waitUntil(backgroundUpdate().catch(function () {}));
  }
});

self.addEventListener("sync", function (event) {
  if (event.tag === "twenty-git-background-v1") {
    event.waitUntil(backgroundUpdate().catch(function () {}));
  }
});

self.addEventListener("message", function (event) {
  if (!event.data || event.data.type !== "twenty:background-check") return;
  event.waitUntil(backgroundUpdate().catch(function () {}));
});
