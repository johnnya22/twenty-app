(function () {
  "use strict";

  var DB_NAME = "twenty-study-sync";
  var DB_VERSION = 1;
  var CONFIG_KEY = "twenty-sync-config-v1";
  var DEVICE_KEY = "twenty-sync-device-v1";
  var AUTO_CHECK_MS = 60000;
  var MIN_CHECK_GAP_MS = 5000;
  var PERIODIC_SYNC_MIN_MS = 15 * 60 * 1000;
  var dbPromise = null;
  var flushPromise = null;
  var checkPromise = null;
  var flushTimer = null;
  var autoTimer = null;
  var watchGeneration = 0;
  var lastCheckStartedAt = 0;
  var status = {
    state: "disabled",
    pending: 0,
    lastSyncAt: "",
    lastCheckedAt: "",
    lastError: "",
    conflicts: 0,
    localVersion: 0,
    remoteVersion: 0,
    outdated: false,
    backgroundEnabled: false
  };

  function clone(value) {
    return value == null ? value : JSON.parse(JSON.stringify(value));
  }

  function same(a, b) {
    try { return JSON.stringify(a) === JSON.stringify(b); } catch (_) { return false; }
  }

  function versionOf(state) {
    var meta = state && state.meta && typeof state.meta === "object" ? state.meta : {};
    var value = Number(meta.gitRevision != null ? meta.gitRevision : meta.revision);
    return Number.isFinite(value) && value > 0 ? value : 0;
  }

  function uid(prefix) {
    if (window.crypto && typeof window.crypto.randomUUID === "function") return (prefix || "id") + "_" + window.crypto.randomUUID();
    return (prefix || "id") + "_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 10);
  }

  function sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
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
    setMeta("serviceWorkerConfig", clean).catch(function () {});
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

  function fileRequest(method, path, options) {
    options = options || {};
    var config = getConfig();
    if (!config.enabled) return Promise.reject(new Error("A sincronização Git ainda não está configurada."));
    if (!navigator.onLine) return Promise.reject(new Error("Sem Internet para sincronizar o PowerPoint."));
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open(method, config.endpoint + path, true);
      xhr.setRequestHeader("X-Twenty-Key", config.key);
      xhr.setRequestHeader("X-Twenty-Device", deviceName());
      if (options.responseType) xhr.responseType = options.responseType;
      if (options.contentType) xhr.setRequestHeader("Content-Type", options.contentType);
      if (xhr.upload && typeof options.onUploadProgress === "function") {
        xhr.upload.onprogress = function (event) {
          options.onUploadProgress({
            loaded: Number(event.loaded) || 0,
            total: event.lengthComputable ? Number(event.total) || 0 : 0,
            progress: event.lengthComputable && event.total ? Math.round(event.loaded / event.total * 100) : null
          });
        };
        xhr.upload.onload = function () {
          if (typeof options.onUploadComplete === "function") options.onUploadComplete();
        };
      }
      if (typeof options.onDownloadProgress === "function") {
        xhr.onprogress = function (event) {
          options.onDownloadProgress({
            loaded: Number(event.loaded) || 0,
            total: event.lengthComputable ? Number(event.total) || 0 : 0,
            progress: event.lengthComputable && event.total ? Math.round(event.loaded / event.total * 100) : null
          });
        };
      }
      xhr.onerror = function () { reject(new Error("Não foi possível contactar o servidor de sincronização.")); };
      xhr.onabort = function () { reject(new Error("Transferência cancelada.")); };
      xhr.onload = function () {
        var payload = null;
        if (options.responseType === "blob") {
          if (xhr.status >= 200 && xhr.status < 300) {
            resolve({ blob: xhr.response, headers: xhr.getAllResponseHeaders(), status: xhr.status });
            return;
          }
        } else {
          try { payload = JSON.parse(xhr.responseText || "{}"); } catch (_) { payload = null; }
          if (xhr.status >= 200 && xhr.status < 300) {
            resolve(payload || {});
            return;
          }
        }
        var message = payload && payload.error ? payload.error : "Erro de transferência (" + xhr.status + ")";
        var error = new Error(message);
        error.status = xhr.status;
        error.payload = payload;
        reject(error);
      };
      xhr.send(options.body == null ? null : options.body);
      if (typeof options.onReady === "function") options.onReady(xhr);
    });
  }

  async function uploadFile(file, options) {
    options = options || {};
    if (!file) throw new Error("Escolhe um ficheiro primeiro.");
    var id = String(options.id || uid("file"));
    var name = String(options.name || file.name || "ficheiro");
    var path = "/files/upload?id=" + encodeURIComponent(id) + "&name=" + encodeURIComponent(name);
    var payload = await fileRequest("POST", path, {
      body: file,
      contentType: file.type || "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      onUploadProgress: options.onProgress,
      onUploadComplete: options.onUploadComplete,
      onReady: options.onReady
    });
    if (!payload.file) throw new Error("O servidor não confirmou o ficheiro enviado.");
    return payload.file;
  }

  async function downloadFile(file, options) {
    options = options || {};
    if (!file || !file.path) throw new Error("Este material não tem ficheiro sincronizado.");
    var path = "/files/download?path=" + encodeURIComponent(file.path) + "&name=" + encodeURIComponent(file.name || "ficheiro");
    var result = await fileRequest("GET", path, {
      responseType: "blob",
      onDownloadProgress: options.onProgress,
      onReady: options.onReady
    });
    return result.blob;
  }

  async function deleteFile(file) {
    if (!file || !file.path) return { ok: true, deleted: false };
    return api("/files/delete", {
      method: "POST",
      body: JSON.stringify({ path: file.path, deviceName: deviceName() })
    });
  }

  async function refreshPending() {
    var items = await listQueue();
    status.pending = items.length;
    window.dispatchEvent(new CustomEvent("twenty:sync-status", { detail: clone(status) }));
    return items.length;
  }

  async function rememberRemote(payload) {
    if (!payload || !payload.state) return;
    var remoteVersion = Number(payload.version || versionOf(payload.state)) || 0;
    await Promise.all([
      setMeta("lastRemote", payload.state),
      setMeta("shadow", payload.state),
      setMeta("lastSha", payload.sha || ""),
      setMeta("lastVersion", remoteVersion)
    ]);
    status.localVersion = remoteVersion;
    status.remoteVersion = remoteVersion;
    status.outdated = false;
  }

  async function queueState(nextState) {
    var config = getConfig();
    status.localVersion = versionOf(nextState);
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
      flush({ background: true }).catch(function () {});
    }, 900);
  }

  async function pullRemote(dispatch, options) {
    options = options || {};
    var payload = await api("/state", { method: "GET" });
    if (!payload.exists || !payload.state) return null;
    var pendingItems = await listQueue();
    if (pendingItems.length && !options.force) {
      updateStatus("syncing", { pending: pendingItems.length, remoteVersion: Number(payload.version || versionOf(payload.state)) || 0 });
      scheduleFlush();
      return null;
    }

    var lastSha = await getMeta("lastSha") || "";
    var knownRemote = await getMeta("lastRemote");
    var lastVersion = Number(await getMeta("lastVersion")) || status.localVersion || 0;
    var incomingVersion = Number(payload.version || versionOf(payload.state)) || 0;
    var changed = !!options.force || !knownRemote || incomingVersion > lastVersion || (!!payload.sha && payload.sha !== lastSha);
    var conflictList = Array.isArray(payload.conflicts) ? payload.conflicts : [];

    status.remoteVersion = incomingVersion;
    status.lastCheckedAt = new Date().toISOString();
    if (!changed) {
      updateStatus("synced", {
        lastError: "",
        conflicts: conflictList.length,
        pending: 0,
        localVersion: Math.max(status.localVersion || 0, incomingVersion),
        remoteVersion: incomingVersion,
        outdated: false,
        lastCheckedAt: status.lastCheckedAt
      });
      return null;
    }

    await rememberRemote(payload);
    updateStatus("synced", {
      lastSyncAt: new Date().toISOString(),
      lastCheckedAt: new Date().toISOString(),
      lastError: "",
      conflicts: conflictList.length,
      pending: 0,
      localVersion: incomingVersion,
      remoteVersion: incomingVersion,
      outdated: false
    });
    if (dispatch !== false) {
      window.dispatchEvent(new CustomEvent("twenty:remote-state", {
        detail: {
          state: clone(payload.state),
          conflicts: conflictList,
          sha: payload.sha || "",
          version: incomingVersion,
          background: !!options.background,
          forced: options.forced || ""
        }
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

  async function flush(options) {
    options = options || {};
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
        var latestPayload = null;
        while (items.length) {
          for (var i = 0; i < items.length; i += 1) {
            var item = items[i];
            var payload = await api("/sync", { method: "POST", body: JSON.stringify(item) });
            latestPayload = payload;
            await deleteQueue(item.id);
            if (payload.state) await rememberRemote(payload);
            status.conflicts = Array.isArray(payload.conflicts) ? payload.conflicts.length : 0;
            status.pending = Math.max(0, status.pending - 1);
            status.remoteVersion = Number(payload.version || versionOf(payload.state)) || status.remoteVersion || 0;
            status.localVersion = status.remoteVersion;
            window.dispatchEvent(new CustomEvent("twenty:sync-status", { detail: clone(status) }));
          }
          items = await listQueue();
        }
        if (latestPayload && latestPayload.state) {
          var finalVersion = Number(latestPayload.version || versionOf(latestPayload.state)) || 0;
          var finalConflicts = Array.isArray(latestPayload.conflicts) ? latestPayload.conflicts : [];
          updateStatus("synced", {
            lastSyncAt: new Date().toISOString(),
            lastCheckedAt: new Date().toISOString(),
            lastError: "",
            conflicts: finalConflicts.length,
            pending: 0,
            localVersion: finalVersion,
            remoteVersion: finalVersion,
            outdated: false
          });
          window.dispatchEvent(new CustomEvent("twenty:remote-state", {
            detail: {
              state: clone(latestPayload.state),
              conflicts: finalConflicts,
              sha: latestPayload.sha || "",
              version: finalVersion,
              background: !!options.background
            }
          }));
          return latestPayload.state;
        }
        return await pullRemote(true, { background: !!options.background });
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
    status.localVersion = versionOf(localState);
    status.remoteVersion = Number(await getMeta("lastVersion")) || status.localVersion || 0;
    var shadow = await getMeta("shadow");
    if (!shadow) await setMeta("shadow", localState);
    await refreshPending();
    if (!config.enabled) {
      updateStatus("disabled", {});
      return null;
    }
    setMeta("serviceWorkerConfig", config).catch(function () {});
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
      return await flush({ background: true });
    } catch (error) {
      updateStatus(navigator.onLine ? "error" : "offline", { lastError: error.message || "Falha de sincronização." });
      return null;
    }
  }

  async function syncNow(localState, initialBase) {
    var config = getConfig();
    if (!config.enabled) throw new Error("Configura primeiro o Worker e a chave de sincronização.");
    status.localVersion = versionOf(localState);
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
    return flush({ background: false });
  }

  async function checkForUpdates(options) {
    options = options || {};
    var config = getConfig();
    if (!config.enabled || !navigator.onLine) return null;
    if (flushPromise) return flushPromise;
    if (checkPromise) return checkPromise;
    var now = Date.now();
    if (!options.force && now - lastCheckStartedAt < MIN_CHECK_GAP_MS) return null;
    lastCheckStartedAt = now;

    checkPromise = (async function () {
      var queued = await listQueue();
      if (queued.length) {
        scheduleFlush();
        return null;
      }
      updateStatus("checking", { lastError: "" });
      try {
        var remote = await api("/version", { method: "GET" });
        var localVersion = status.localVersion || Number(await getMeta("lastVersion")) || 0;
        var lastSha = await getMeta("lastSha") || "";
        var remoteVersion = Number(remote.version) || 0;
        var knownRemote = await getMeta("lastRemote");
        var outdated = !!remote.exists && (!knownRemote || remoteVersion > localVersion || (!!remote.sha && !!lastSha && remote.sha !== lastSha));
        status.remoteVersion = remoteVersion;
        status.lastCheckedAt = new Date().toISOString();
        status.outdated = outdated;
        if (outdated) {
          updateStatus("checking", {
            localVersion: localVersion,
            remoteVersion: remoteVersion,
            outdated: true,
            lastCheckedAt: status.lastCheckedAt
          });
          return await pullRemote(options.dispatch !== false, { background: true });
        }
        if (!lastSha && remote.sha) await setMeta("lastSha", remote.sha);
        if (remoteVersion) await setMeta("lastVersion", remoteVersion);
        updateStatus("synced", {
          lastError: "",
          localVersion: Math.max(localVersion, remoteVersion),
          remoteVersion: remoteVersion,
          outdated: false,
          lastCheckedAt: status.lastCheckedAt
        });
        return null;
      } catch (error) {
        updateStatus(navigator.onLine ? "error" : "offline", { lastError: error.message || "Falha ao verificar atualizações." });
        if (options.throwOnError) throw error;
        return null;
      }
    })().finally(function () { checkPromise = null; });
    return checkPromise;
  }

  async function forcePull(options) {
    options = options || {};
    var config = getConfig();
    if (!config.enabled) throw new Error("Configura primeiro o Worker e a chave de sincronização.");
    updateStatus("syncing", { lastError: "" });
    try {
      await clearQueue();
      return await pullRemote(options.dispatch !== false, { force: true, forced: "pull", background: false });
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
      await rememberRemote(payload);
      var conflictList = Array.isArray(payload.conflicts) ? payload.conflicts : [];
      updateStatus("synced", {
        lastSyncAt: new Date().toISOString(),
        lastCheckedAt: new Date().toISOString(),
        lastError: "",
        conflicts: conflictList.length,
        pending: 0,
        localVersion: Number(payload.version || versionOf(payload.state)) || 0,
        remoteVersion: Number(payload.version || versionOf(payload.state)) || 0,
        outdated: false
      });
      if (options.dispatch !== false) {
        window.dispatchEvent(new CustomEvent("twenty:remote-state", {
          detail: { state: clone(payload.state), conflicts: conflictList, sha: payload.sha || "", version: payload.version || versionOf(payload.state), forced: "push" }
        }));
      }
      return clone(payload.state);
    } catch (error) {
      await refreshPending();
      updateStatus(navigator.onLine ? "error" : "offline", { lastError: error.message || "Falha no force push." });
      throw error;
    }
  }

  async function watchForCommits(generation) {
    while (generation === watchGeneration && getConfig().enabled) {
      if (document.hidden || !navigator.onLine) {
        await sleep(3000);
        continue;
      }
      try {
        var since = Math.max(status.localVersion || 0, Number(await getMeta("lastVersion")) || 0);
        var event = await api("/watch?since=" + encodeURIComponent(since), { method: "GET" });
        if (generation !== watchGeneration) return;
        if (event && event.changed && Number(event.version || 0) > since) {
          status.remoteVersion = Number(event.version) || status.remoteVersion || 0;
          status.outdated = true;
          updateStatus("checking", {
            remoteVersion: status.remoteVersion,
            outdated: true,
            lastCheckedAt: new Date().toISOString()
          });
          await checkForUpdates({ force: true });
        } else {
          status.lastCheckedAt = new Date().toISOString();
        }
      } catch (_) {
        await sleep(3000);
      }
    }
  }

  async function registerBackgroundSync() {
    if (!("serviceWorker" in navigator)) return false;
    try {
      var registration = await navigator.serviceWorker.ready;
      if (registration.periodicSync && typeof registration.periodicSync.register === "function") {
        await registration.periodicSync.register("twenty-git-background-v1", { minInterval: PERIODIC_SYNC_MIN_MS });
        updateStatus(status.state, { backgroundEnabled: true });
        return true;
      }
      updateStatus(status.state, { backgroundEnabled: false });
      return false;
    } catch (_) {
      updateStatus(status.state, { backgroundEnabled: false });
      return false;
    }
  }

  function startAutoSync() {
    clearInterval(autoTimer);
    watchGeneration += 1;
    if (!getConfig().enabled) return;
    var generation = watchGeneration;
    checkForUpdates({ force: true }).catch(function () {});
    watchForCommits(generation).catch(function () {});
    autoTimer = setInterval(function () {
      if (document.hidden) return;
      // Verificação de segurança para commits feitos fora da app ou após reinício do Worker.
      checkForUpdates().catch(function () {});
    }, AUTO_CHECK_MS);
    registerBackgroundSync().catch(function () {});
  }

  function stopAutoSync() {
    clearInterval(autoTimer);
    autoTimer = null;
    watchGeneration += 1;
  }

  async function configure(endpoint, key) {
    var config = saveConfig({ endpoint: endpoint, key: key, enabled: true });
    if (!config.endpoint || !config.key) throw new Error("Falta o endereço do Worker ou a chave.");
    updateStatus("idle", { lastError: "" });
    startAutoSync();
    return config;
  }

  function disable() {
    var config = getConfig();
    config.enabled = false;
    saveConfig(config);
    stopAutoSync();
    updateStatus("disabled", {});
  }

  async function adoptRemoteState(remoteState) {
    var remoteVersion = versionOf(remoteState);
    await Promise.all([
      setMeta("shadow", remoteState),
      setMeta("lastRemote", remoteState),
      setMeta("lastVersion", remoteVersion)
    ]);
    status.localVersion = remoteVersion;
    status.remoteVersion = Math.max(status.remoteVersion || 0, remoteVersion);
    status.outdated = false;
  }

  async function resetLocal() {
    clearTimeout(flushTimer);
    await clearQueue();
    await setMeta("shadow", null);
    await setMeta("lastRemote", null);
    await setMeta("lastSha", "");
    await setMeta("lastVersion", 0);
    updateStatus(getConfig().enabled ? "idle" : "disabled", { pending: 0, lastError: "", conflicts: 0, localVersion: 0, remoteVersion: 0, outdated: false });
  }

  window.addEventListener("online", function () {
    scheduleFlush();
    checkForUpdates({ force: true }).catch(function () {});
  });
  window.addEventListener("focus", function () {
    if (getConfig().enabled) checkForUpdates({ force: true }).catch(function () {});
  });
  window.addEventListener("pageshow", function () {
    if (getConfig().enabled) checkForUpdates({ force: true }).catch(function () {});
  });
  document.addEventListener("visibilitychange", function () {
    if (!document.hidden && getConfig().enabled) checkForUpdates({ force: true }).catch(function () {});
  });
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.addEventListener("message", function (event) {
      if (!event.data || event.data.type !== "twenty:background-update-ready") return;
      checkForUpdates({ force: true }).catch(function () {});
    });
  }

  window.TwentySync = {
    getConfig: getConfig,
    getStatus: getStatus,
    configure: configure,
    disable: disable,
    bootstrap: bootstrap,
    queueState: queueState,
    syncNow: syncNow,
    checkForUpdates: checkForUpdates,
    startAutoSync: startAutoSync,
    forcePull: forcePull,
    forcePush: forcePush,
    uploadFile: uploadFile,
    downloadFile: downloadFile,
    deleteFile: deleteFile,
    flush: flush,
    adoptRemoteState: adoptRemoteState,
    resetLocal: resetLocal
  };
})();
