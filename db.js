(function () {
  "use strict";

  var DB_NAME = "twenty-study-os";
  var DB_VERSION = 2;
  var STATE_KEY = "academic-state";
  var dbPromise = null;

  function open() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      if (!("indexedDB" in window)) {
        resolve(null);
        return;
      }
      var request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = function () {
        var db = request.result;
        if (!db.objectStoreNames.contains("kv")) db.createObjectStore("kv");
        if (!db.objectStoreNames.contains("files")) db.createObjectStore("files", { keyPath: "id" });
      };
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error); };
    });
    return dbPromise;
  }

  function transaction(store, mode, action) {
    return open().then(function (db) {
      if (!db) return action(null, null);
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(store, mode);
        var objectStore = tx.objectStore(store);
        var result;
        try { result = action(objectStore, tx); } catch (error) { reject(error); return; }
        tx.oncomplete = function () { resolve(result); };
        tx.onerror = function () { reject(tx.error); };
        tx.onabort = function () { reject(tx.error || new Error("Transação cancelada")); };
      });
    });
  }

  function getState() {
    return open().then(function (db) {
      if (!db) {
        try { return JSON.parse(localStorage.getItem(STATE_KEY) || "null"); } catch (_) { return null; }
      }
      return new Promise(function (resolve, reject) {
        var request = db.transaction("kv", "readonly").objectStore("kv").get(STATE_KEY);
        request.onsuccess = function () { resolve(request.result || null); };
        request.onerror = function () { reject(request.error); };
      });
    });
  }

  function saveState(state, options) {
    var clean = JSON.parse(JSON.stringify(state));
    options = options || {};
    return open().then(function (db) {
      if (!db) {
        localStorage.setItem(STATE_KEY, JSON.stringify(clean));
        return clean;
      }
      return transaction("kv", "readwrite", function (store) {
        store.put(clean, STATE_KEY);
        return clean;
      });
    }).then(function () {
      if (!options.skipSync && window.TwentySync) {
        return window.TwentySync.queueState(clean).catch(function (error) {
          console.error("Twenty sync queue:", error);
          return clean;
        });
      }
      return clean;
    });
  }

  function putFile(file, metadata) {
    var id = "file_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
    var record = {
      id: id,
      blob: file,
      name: file.name || "documento.pdf",
      type: file.type || "application/octet-stream",
      size: file.size || 0,
      createdAt: new Date().toISOString(),
      metadata: metadata || {}
    };
    return open().then(function (db) {
      if (!db) throw new Error("Este navegador não suporta armazenamento de ficheiros local.");
      return transaction("files", "readwrite", function (store) {
        store.put(record);
        return id;
      });
    });
  }

  function getFile(id) {
    return open().then(function (db) {
      if (!db) return null;
      return new Promise(function (resolve, reject) {
        var request = db.transaction("files", "readonly").objectStore("files").get(id);
        request.onsuccess = function () { resolve(request.result || null); };
        request.onerror = function () { reject(request.error); };
      });
    });
  }

  function deleteFile(id) {
    return transaction("files", "readwrite", function (store) {
      if (store) store.delete(id);
    });
  }

  function listFiles() {
    return open().then(function (db) {
      if (!db) return [];
      return new Promise(function (resolve, reject) {
        var request = db.transaction("files", "readonly").objectStore("files").getAll();
        request.onsuccess = function () { resolve(request.result || []); };
        request.onerror = function () { reject(request.error); };
      });
    });
  }

  function clearAll() {
    return open().then(function (db) {
      localStorage.removeItem(STATE_KEY);
      var localClear = !db ? Promise.resolve() : Promise.all([
        transaction("kv", "readwrite", function (store) { store.clear(); }),
        transaction("files", "readwrite", function (store) { store.clear(); })
      ]);
      return localClear.then(function () {
        if (window.TwentySync) return window.TwentySync.resetLocal();
      });
    });
  }

  window.TwentyDB = {
    open: open,
    getState: getState,
    saveState: saveState,
    putFile: putFile,
    getFile: getFile,
    deleteFile: deleteFile,
    listFiles: listFiles,
    clearAll: clearAll
  };
})();
