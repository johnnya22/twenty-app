(function () {
  "use strict";

  var DB_NAME = "twenty-study-sync";
  var DB_VERSION = 1;
  var CONFIG_KEY = "twenty-sync-config-v1";
  var DEVICE_KEY = "twenty-sync-device-v1";
  var dbPromise = null;
  var flushPromise = null;
  var flushTimer = null;
  var status = {
    state: "disabled",
    pending: 0,
    lastSyncAt: "",
    lastError: "",
    conflicts: 0
  };

  function clone(value) {
    return value == null ? value : JSON.parse(JSON.stringify(value));
  }

  function same(a, b) {
    try { return JSON.stringify(a) === JSON.stringify(b); } catch (_) { return false; }
  }

  function uid(prefix) {
    if (window.crypto && typeof window.crypto.randomUUID === "function") return (prefix || "id") + "_" + window.crypto.randomUUID();
    return (prefix || "id") + "_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 10);
  }

  function open() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      var request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = function () {
        var db = request.result;
        if (!db.objectStoreNames.contains("meta")) db.createObjectStore("meta");
        if (!db.objectStoreNames.contains("queue")) db.createObjectStore("queue", { keyPath: "id" });
      };
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error); };
    });
    return dbPromise;
  }

  function requestResult(request) {
    return new Promise(function (resolve, reject) {
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error); };
    });
  }

  function getMeta(key) {
    return open().then(function (db) {
      return requestResult(db.transaction("meta", "readonly").objectStore("meta").get(key));
    });
  }

  function setMeta(key, value) {
    return open().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction("meta", "readwrite");
        tx.objectStore("meta").put(clone(value), key);
        tx.oncomplete = function () { resolve(value); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function addQueue(item) {
    return open().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction("queue", "readwrite");
        tx.objectStore("queue").put(clone(item));
        tx.oncomplete = function () { resolve(item); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function deleteQueue(id) {
    return open().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction("queue", "readwrite");
        tx.objectStore("queue").delete(id);
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function listQueue() {
    return open().then(function (db) {
      return requestResult(db.transaction("queue", "readonly").objectStore("queue").getAll());
    }).then(function (items) {
      return (items || []).sort(function (a, b) {
        return String(a.createdAt || "").localeCompare(String(b.createdAt || "")) || String(a.id).localeCompare(String(b.id));
      });
    });
  }

  function clearQueue() {
    return open().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction("queue", "readwrite");
        tx.objectStore("queue").clear();
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function getConfig() {
    try {
      var parsed = JSON.parse(localStorage.getItem(CONFIG_KEY) || "null") || {};
      parsed.endpoint = String(parsed.endpoint || "").replace(/\/+$/, "");
      parsed.key = String(parsed.key || "");
      parsed.enabled = parsed.enabled !== false && !!parsed.endpoint && !!parsed.key;
      return parsed;
    } catch (_) {
      return { endpoint: "", key: "", enabled: false };
    }
  }

  function saveConfig(config) {
    var clean = {
      endpoint: String(config.endpoint || "").trim().replace(/\/+$/, ""),
      key: String(config.key || "").trim(),
      enabled: config.enabled !== false
    };
    localStorage.setItem(CONFIG_KEY, JSON.stringify(clean));
    updateStatus(clean.enabled && clean.endpoint && clean.key ? "idle" : "disabled", {});
    return clean;
  }

  function deviceId() {
    var id = localStorage.getItem(DEVICE_KEY);
    if (!id) {
      id = uid("device");
      localStorage.setItem(DEVICE_KEY, id);
    }
    return id;
  }

  function deviceName() {
    var ua = navigator.userAgent || "Dispositivo";
    if (/iPhone|iPad|iPod/i.test(ua)) return "iPhone/iPad";
    if (/Android/i.test(ua)) return "Android";
    if (/Macintosh|Mac OS X/i.test(ua)) return "Mac";
    if (/Windows/i.test(ua)) return "Windows";
    return "Browser";
  }

  function updateStatus(nextState, patch) {
    status = Object.assign({}, status, patch || {}, { state: nextState || status.state });
    window.dispatchEvent(new CustomEvent("twenty:sync-status", { detail: clone(status) }));
  }

  function getStatus() {
    var config = getConfig();
    return Object.assign({}, status, { configured: !!(config.endpoint && config.key), enabled: !!config.enabled });
  }

  async function api(path, options) {
    var config = getConfig();
    if (!config.enabled) throw new Error("A sincronização Git ainda não está configurada.");
    var response = await fetch(config.endpoint + path, Object.assign({
      cache: "no-store",
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Twenty-Key": config.key
      }
    }, options || {}));
    var payload = null;
    try { payload = await response.json(); } catch (_) { payload = null; }
    if (!response.ok) {
      var error = new Error(payload && payload.error ? payload.error : "Erro de sincronização (" + response.status + ")");
      error.status = response.status;
      error.payload = payload;
      throw error;
    }
    return payload || {};
  }

  async function refreshPending() {
    var items = await listQueue();
    status.pending = items.length;
    window.dispatchEvent(new CustomEvent("twenty:sync-status", { detail: clone(status) }));
    return items.length;
  }

  async function queueState(nextState) {
    var config = getConfig();
    if (!config.enabled) return nextState;
    var base = await getMeta("shadow");
    if (!base) base = clone(nextState);
    var mutation = {
      id: uid("mutation"),
      deviceId: deviceId(),
      deviceName: deviceName(),
      createdAt: new Date().toISOString(),
      base: clone(base),
      next: clone(nextState)
    };
    await addQueue(mutation);
    await setMeta("shadow", nextState);
    await refreshPending();
    scheduleFlush();
    return nextState;
  }

  function scheduleFlush() {
    clearTimeout(flushTimer);
    flushTimer = setTimeout(function () {
      flush().catch(function () {});
    }, 900);
  }

  async function pullRemote(dispatch) {
    var payload = await api("/state", { method: "GET" });
    if (!payload.exists || !payload.state) return null;
    var pendingItems = await listQueue();
    if (pendingItems.length) {
      updateStatus("syncing", { pending: pendingItems.length });
      scheduleFlush();
      return null;
    }
    await setMeta("lastRemote", payload.state);
    await setMeta("shadow", payload.state);
    await setMeta("lastSha", payload.sha || "");
    var conflictList = Array.isArray(payload.conflicts) ? payload.conflicts : [];
    updateStatus("synced", {
      lastSyncAt: new Date().toISOString(),
      lastError: "",
      conflicts: conflictList.length,
      pending: 0
    });
    if (dispatch !== false) {
      window.dispatchEvent(new CustomEvent("twenty:remote-state", {
        detail: { state: clone(payload.state), conflicts: conflictList, sha: payload.sha || "" }
      }));
    }
    return payload.state;
  }

  async function initializeRemote(localState) {
    var mutation = {
      id: uid("mutation"),
      deviceId: deviceId(),
      deviceName: deviceName(),
      createdAt: new Date().toISOString(),
      base: null,
      next: clone(localState)
    };
    await addQueue(mutation);
    await setMeta("shadow", localState);
  }

  async function flush() {
    if (flushPromise) return flushPromise;
    flushPromise = (async function () {
      var config = getConfig();
      if (!config.enabled) {
        updateStatus("disabled", { pending: 0 });
        return null;
      }
      if (!navigator.onLine) {
        await refreshPending();
        updateStatus("offline", { lastError: "Sem ligação à Internet." });
        return null;
      }
      updateStatus("syncing", { lastError: "" });
      try {
        var items = await listQueue();
        while (items.length) {
          for (var i = 0; i < items.length; i += 1) {
            var item = items[i];
            var payload = await api("/sync", { method: "POST", body: JSON.stringify(item) });
            await deleteQueue(item.id);
            if (payload.state) {
              await setMeta("lastRemote", payload.state);
              await setMeta("lastSha", payload.sha || "");
            }
            status.conflicts = Array.isArray(payload.conflicts) ? payload.conflicts.length : 0;
            status.pending = Math.max(0, status.pending - 1);
            window.dispatchEvent(new CustomEvent("twenty:sync-status", { detail: clone(status) }));
          }
          items = await listQueue();
        }
        return await pullRemote(true);
      } catch (error) {
        await refreshPending();
        updateStatus(navigator.onLine ? "error" : "offline", { lastError: error.message || "Falha de sincronização." });
        throw error;
      }
    })().finally(function () { flushPromise = null; });
    return flushPromise;
  }

  async function bootstrap(localState, initialBase) {
    var config = getConfig();
    await open();
    var shadow = await getMeta("shadow");
    if (!shadow) await setMeta("shadow", localState);
    await refreshPending();
    if (!config.enabled) {
      updateStatus("disabled", {});
      return null;
    }
    try {
      var queued = await listQueue();
      if (!queued.length) {
        var remote = await api("/state", { method: "GET" });
        if (!remote.exists) {
          await initializeRemote(localState);
        } else {
          var knownRemote = await getMeta("lastRemote");
          if (!knownRemote && initialBase && !same(localState, initialBase)) {
            await addQueue({
              id: uid("mutation"),
              deviceId: deviceId(),
              deviceName: deviceName(),
              createdAt: new Date().toISOString(),
              base: clone(initialBase),
              next: clone(localState)
            });
            await setMeta("shadow", localState);
          }
        }
      }
      return await flush();
    } catch (error) {
      updateStatus(navigator.onLine ? "error" : "offline", { lastError: error.message || "Falha de sincronização." });
      return null;
    }
  }

  async function syncNow(localState, initialBase) {
    var config = getConfig();
    if (!config.enabled) throw new Error("Configura primeiro o Worker e a chave de sincronização.");
    var queued = await listQueue();
    if (!queued.length) {
      var remote = await api("/state", { method: "GET" });
      if (!remote.exists) {
        await initializeRemote(localState);
      } else {
        var knownRemote = await getMeta("lastRemote");
        if (!knownRemote && initialBase && !same(localState, initialBase)) {
          await addQueue({
            id: uid("mutation"),
            deviceId: deviceId(),
            deviceName: deviceName(),
            createdAt: new Date().toISOString(),
            base: clone(initialBase),
            next: clone(localState)
          });
          await setMeta("shadow", localState);
        }
      }
    }
    return flush();
  }

  async function forcePull(options) {
    options = options || {};
    var config = getConfig();
    if (!config.enabled) throw new Error("Configura primeiro o Worker e a chave de sincronização.");
    updateStatus("syncing", { lastError: "" });
    try {
      var payload = await api("/state", { method: "GET" });
      if (!payload.exists || !payload.state) throw new Error("Ainda não existem dados sincronizados neste Git.");
      await clearQueue();
      await setMeta("lastRemote", payload.state);
      await setMeta("shadow", payload.state);
      await setMeta("lastSha", payload.sha || "");
      var conflictList = Array.isArray(payload.conflicts) ? payload.conflicts : [];
      updateStatus("synced", {
        lastSyncAt: new Date().toISOString(),
        lastError: "",
        conflicts: conflictList.length,
        pending: 0
      });
      if (options.dispatch !== false) {
        window.dispatchEvent(new CustomEvent("twenty:remote-state", {
          detail: { state: clone(payload.state), conflicts: conflictList, sha: payload.sha || "", forced: "pull" }
        }));
      }
      return clone(payload.state);
    } catch (error) {
      await refreshPending();
      updateStatus(navigator.onLine ? "error" : "offline", { lastError: error.message || "Falha no force pull." });
      throw error;
    }
  }

  async function forcePush(localState, options) {
    options = options || {};
    var config = getConfig();
    if (!config.enabled) throw new Error("Configura primeiro o Worker e a chave de sincronização.");
    if (!localState || typeof localState !== "object") throw new Error("Não existem dados locais válidos para enviar.");
    updateStatus("syncing", { lastError: "" });
    try {
      var payload = await api("/force-push", {
        method: "POST",
        body: JSON.stringify({
          operationId: uid("forcepush"),
          deviceId: deviceId(),
          deviceName: deviceName(),
          createdAt: new Date().toISOString(),
          state: clone(localState)
        })
      });
      if (!payload.state) throw new Error("O Worker não devolveu o estado confirmado.");
      await clearQueue();
      await setMeta("lastRemote", payload.state);
      await setMeta("shadow", payload.state);
      await setMeta("lastSha", payload.sha || "");
      var conflictList = Array.isArray(payload.conflicts) ? payload.conflicts : [];
      updateStatus("synced", {
        lastSyncAt: new Date().toISOString(),
        lastError: "",
        conflicts: conflictList.length,
        pending: 0
      });
      if (options.dispatch !== false) {
        window.dispatchEvent(new CustomEvent("twenty:remote-state", {
          detail: { state: clone(payload.state), conflicts: conflictList, sha: payload.sha || "", forced: "push" }
        }));
      }
      return clone(payload.state);
    } catch (error) {
      await refreshPending();
      updateStatus(navigator.onLine ? "error" : "offline", { lastError: error.message || "Falha no force push." });
      throw error;
    }
  }

  async function configure(endpoint, key) {
    var config = saveConfig({ endpoint: endpoint, key: key, enabled: true });
    if (!config.endpoint || !config.key) throw new Error("Falta o endereço do Worker ou a chave.");
    updateStatus("idle", { lastError: "" });
    return config;
  }

  function disable() {
    var config = getConfig();
    config.enabled = false;
    saveConfig(config);
    updateStatus("disabled", {});
  }

  async function adoptRemoteState(remoteState) {
    await setMeta("shadow", remoteState);
    await setMeta("lastRemote", remoteState);
  }

  async function resetLocal() {
    clearTimeout(flushTimer);
    await clearQueue();
    await setMeta("shadow", null);
    await setMeta("lastRemote", null);
    await setMeta("lastSha", "");
    updateStatus(getConfig().enabled ? "idle" : "disabled", { pending: 0, lastError: "", conflicts: 0 });
  }

  window.addEventListener("online", function () { scheduleFlush(); });

  window.TwentySync = {
    getConfig: getConfig,
    getStatus: getStatus,
    configure: configure,
    disable: disable,
    bootstrap: bootstrap,
    queueState: queueState,
    syncNow: syncNow,
    forcePull: forcePull,
    forcePush: forcePush,
    flush: flush,
    adoptRemoteState: adoptRemoteState,
    resetLocal: resetLocal
  };
})();
