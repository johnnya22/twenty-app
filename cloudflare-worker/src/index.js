import { DurableObject } from "cloudflare:workers";

const VOLATILE_META_KEYS = new Set([
  "revision",
  "gitRevision",
  "updatedAt",
  "source",
  "externalFingerprint",
  "externalCheckedAt",
  "externalRevision",
  "sync",
  "syncConflicts"
]);

function clone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function deepEqual(a, b) {
  if (a === b) return true;
  if (typeof a !== typeof b || a === null || b === null) return false;
  if (Array.isArray(a)) {
    if (!Array.isArray(b) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) if (!deepEqual(a[i], b[i])) return false;
    return true;
  }
  if (isObject(a)) {
    if (!isObject(b)) return false;
    const aKeys = Object.keys(a).sort();
    const bKeys = Object.keys(b).sort();
    if (!deepEqual(aKeys, bKeys)) return false;
    for (const key of aKeys) if (!deepEqual(a[key], b[key])) return false;
    return true;
  }
  return false;
}

function cleanForMerge(input) {
  const state = clone(input || {});
  state.meta = isObject(state.meta) ? state.meta : {};
  for (const key of VOLATILE_META_KEYS) delete state.meta[key];
  return state;
}

function stateVersion(state) {
  const meta = state && isObject(state.meta) ? state.meta : {};
  const value = Number(meta.gitRevision != null ? meta.gitRevision : meta.revision);
  return Number.isFinite(value) && value > 0 ? value : 0;
}

function preview(value) {
  if (value === undefined) return "<apagado>";
  let raw;
  try { raw = JSON.stringify(value); } catch (_) { raw = String(value); }
  if (raw.length > 500) raw = raw.slice(0, 497) + "…";
  return raw;
}

function isEntityArray(value) {
  return Array.isArray(value) && value.every((item) => isObject(item) && typeof item.id === "string" && item.id);
}

function recordConflict(conflicts, path, kind, base, local, remote) {
  conflicts.push({
    path: path || "$",
    kind,
    base: preview(base),
    local: preview(local),
    remote: preview(remote)
  });
}

function mergeThreeWay(base, local, remote, path, conflicts) {
  const baseMissing = base === undefined;
  const localMissing = local === undefined;
  const remoteMissing = remote === undefined;

  if (deepEqual(local, base)) return clone(remote);
  if (deepEqual(remote, base)) return clone(local);
  if (deepEqual(local, remote)) return clone(local);

  if (localMissing) {
    if (remoteMissing) return undefined;
    recordConflict(conflicts, path, "apagado-localmente/alterado-remotamente", base, local, remote);
    return clone(remote); // nunca apaga uma versão que o outro dispositivo alterou
  }

  if (remoteMissing) {
    recordConflict(conflicts, path, "alterado-localmente/apagado-remotamente", base, local, remote);
    return clone(local); // preserva a alteração local em vez de a perder
  }

  if (Array.isArray(local) && Array.isArray(remote)) {
    const baseArray = Array.isArray(base) ? base : [];
    if (isEntityArray(local) && isEntityArray(remote) && (baseMissing || isEntityArray(baseArray))) {
      const baseMap = new Map(baseArray.map((item) => [item.id, item]));
      const localMap = new Map(local.map((item) => [item.id, item]));
      const remoteMap = new Map(remote.map((item) => [item.id, item]));
      const order = [];
      remote.forEach((item) => { if (!order.includes(item.id)) order.push(item.id); });
      local.forEach((item) => { if (!order.includes(item.id)) order.push(item.id); });
      baseArray.forEach((item) => { if (!order.includes(item.id)) order.push(item.id); });
      const result = [];
      for (const id of order) {
        const merged = mergeThreeWay(baseMap.get(id), localMap.get(id), remoteMap.get(id), path + "[id=" + id + "]", conflicts);
        if (merged !== undefined) result.push(merged);
      }
      return result;
    }
    recordConflict(conflicts, path, "lista-alterada-nos-dois-dispositivos", base, local, remote);
    return clone(local);
  }

  if (isObject(local) && isObject(remote)) {
    const baseObject = isObject(base) ? base : {};
    const result = {};
    const keys = new Set([...Object.keys(baseObject), ...Object.keys(local), ...Object.keys(remote)]);
    for (const key of keys) {
      const merged = mergeThreeWay(baseObject[key], local[key], remote[key], path ? path + "." + key : key, conflicts);
      if (merged !== undefined) result[key] = merged;
    }
    return result;
  }

  recordConflict(conflicts, path, "mesmo-campo-alterado-nos-dois-dispositivos", base, local, remote);
  return clone(local); // o último envio ganha apenas neste campo; o valor anterior fica registado no conflito e no histórico Git
}

function utf8ToBase64(text) {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function base64ToUtf8(value) {
  const binary = atob(String(value || "").replace(/\n/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

function bytesToBase64(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input || 0);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function safeFileId(value) {
  const clean = String(value || "").toLowerCase().replace(/[^a-z0-9_-]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 80);
  return clean || `file-${Date.now().toString(36)}`;
}

function safeFileName(value) {
  const raw = String(value || "ficheiro").split(/[\/]/).pop() || "ficheiro";
  const clean = raw.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9._ -]+/g, "-").replace(/\s+/g, "-").replace(/-+/g, "-").replace(/^[-.]+|[-.]+$/g, "").slice(0, 120);
  return clean || "ficheiro";
}

function fileRoot(env) {
  return String(env.FILE_ROOT || "data/files").replace(/^\/+|\/+$/g, "") || "data/files";
}

function buildFilePath(env, id, name) {
  return `${fileRoot(env)}/${safeFileId(id)}-${safeFileName(name)}`;
}

function validateFilePath(env, path) {
  const clean = String(path || "").replace(/^\/+/, "");
  const root = fileRoot(env) + "/";
  if (!clean.startsWith(root) || clean.includes("..")) throw Object.assign(new Error("Caminho de ficheiro inválido."), { status: 400 });
  return clean;
}

function githubHeaders(env) {
  return {
    "Accept": "application/vnd.github+json",
    "Authorization": "Bearer " + env.GITHUB_TOKEN,
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "Twenty-Study-OS-Sync"
  };
}

function githubContentUrl(env, path) {
  const owner = encodeURIComponent(env.GITHUB_OWNER);
  const repo = encodeURIComponent(env.GITHUB_REPO);
  const encodedPath = String(path || "").split("/").map(encodeURIComponent).join("/");
  return `https://api.github.com/repos/${owner}/${repo}/contents/${encodedPath}`;
}

function githubUrl(env) {
  return githubContentUrl(env, env.STATE_PATH || "data/twenty-state.json");
}

async function readGitHubState(env) {
  const url = githubUrl(env) + "?ref=" + encodeURIComponent(env.GITHUB_BRANCH || "main");
  const response = await fetch(url, { headers: githubHeaders(env) });
  if (response.status === 404) return { exists: false, state: null, sha: "", version: 0, updatedAt: "" };
  if (!response.ok) {
    const body = await response.text();
    throw new Error("GitHub GET " + response.status + ": " + body.slice(0, 300));
  }
  const payload = await response.json();
  const state = JSON.parse(base64ToUtf8(payload.content));
  return { exists: true, state, sha: payload.sha || "", version: stateVersion(state), updatedAt: state && state.meta && state.meta.updatedAt || "" };
}

async function writeGitHubState(env, state, sha, message) {
  const body = {
    message,
    content: utf8ToBase64(JSON.stringify(state, null, 2) + "\n"),
    branch: env.GITHUB_BRANCH || "main"
  };
  if (sha) body.sha = sha;
  const response = await fetch(githubUrl(env), {
    method: "PUT",
    headers: Object.assign({ "Content-Type": "application/json" }, githubHeaders(env)),
    body: JSON.stringify(body)
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error("GitHub PUT " + response.status + ": " + (payload.message || "erro desconhecido"));
    error.status = response.status;
    throw error;
  }
  return { sha: payload.content && payload.content.sha || "", commitSha: payload.commit && payload.commit.sha || "" };
}

async function readGitHubFileMeta(env, path) {
  const url = githubContentUrl(env, path) + "?ref=" + encodeURIComponent(env.GITHUB_BRANCH || "main");
  const response = await fetch(url, { headers: githubHeaders(env) });
  if (response.status === 404) return null;
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error("GitHub file GET " + response.status + ": " + (payload.message || "erro desconhecido"));
    error.status = response.status;
    throw error;
  }
  return payload;
}

async function writeGitHubBinary(env, path, bytes, metadata) {
  const existing = await readGitHubFileMeta(env, path);
  const body = {
    message: `Twenty: PowerPoint ${metadata.name} · ${metadata.deviceName || "dispositivo"}`,
    content: bytesToBase64(bytes),
    branch: env.GITHUB_BRANCH || "main"
  };
  if (existing && existing.sha) body.sha = existing.sha;
  const response = await fetch(githubContentUrl(env, path), {
    method: "PUT",
    headers: Object.assign({ "Content-Type": "application/json" }, githubHeaders(env)),
    body: JSON.stringify(body)
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error("GitHub file PUT " + response.status + ": " + (payload.message || "erro desconhecido"));
    error.status = response.status;
    throw error;
  }
  return { sha: payload.content && payload.content.sha || "", commitSha: payload.commit && payload.commit.sha || "" };
}

async function deleteGitHubBinary(env, path, deviceName) {
  const existing = await readGitHubFileMeta(env, path);
  if (!existing || !existing.sha) return { deleted: false, sha: "", commitSha: "" };
  const response = await fetch(githubContentUrl(env, path), {
    method: "DELETE",
    headers: Object.assign({ "Content-Type": "application/json" }, githubHeaders(env)),
    body: JSON.stringify({
      message: `Twenty: apagar PowerPoint · ${deviceName || "dispositivo"}`,
      sha: existing.sha,
      branch: env.GITHUB_BRANCH || "main"
    })
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error("GitHub file DELETE " + response.status + ": " + (payload.message || "erro desconhecido"));
    error.status = response.status;
    throw error;
  }
  return { deleted: true, sha: existing.sha, commitSha: payload.commit && payload.commit.sha || "" };
}

async function downloadGitHubBinary(env, path) {
  const url = githubContentUrl(env, path) + "?ref=" + encodeURIComponent(env.GITHUB_BRANCH || "main");
  const headers = githubHeaders(env);
  headers.Accept = "application/vnd.github.raw+json";
  const response = await fetch(url, { headers });
  if (!response.ok) {
    const body = await response.text();
    const error = new Error("GitHub file download " + response.status + ": " + body.slice(0, 240));
    error.status = response.status;
    throw error;
  }
  return response;
}

function mutationSummary(base, next) {
  const names = [];
  const keys = new Set([...Object.keys(base || {}), ...Object.keys(next || {})]);
  for (const key of keys) {
    if (key === "meta") continue;
    if (!deepEqual(base && base[key], next && next[key])) names.push(key);
  }
  if (!names.length) return "estado atualizado";
  if (names.length <= 3) return names.join(", ");
  return names.slice(0, 3).join(", ") + ` +${names.length - 3}`;
}

function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" }
  });
}

export class SyncCoordinator extends DurableObject {
  constructor(ctx, env) {
    super(ctx, env);
    this.ctx = ctx;
    this.env = env;
    this.tail = Promise.resolve();
    this.watchers = new Set();
  }

  async fetch(request) {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname.endsWith("/watch")) {
      const since = Number(url.searchParams.get("since")) || 0;
      return this.watchVersion(since);
    }
    if (request.method === "GET" && url.pathname.endsWith("/version")) {
      return this.serial(() => this.getVersion());
    }
    if (request.method === "GET" && url.pathname.endsWith("/state")) {
      return this.serial(() => this.getState());
    }
    if (request.method === "POST" && url.pathname.endsWith("/sync")) {
      const mutation = await request.json();
      return this.serial(() => this.applyMutation(mutation));
    }
    if (request.method === "POST" && url.pathname.endsWith("/force-push")) {
      const payload = await request.json();
      return this.serial(() => this.forcePush(payload));
    }
    if (request.method === "POST" && url.pathname.endsWith("/files/upload")) {
      const id = url.searchParams.get("id") || "";
      const name = url.searchParams.get("name") || "ficheiro";
      const mimeType = request.headers.get("Content-Type") || "application/octet-stream";
      const deviceName = request.headers.get("X-Twenty-Device") || "Dispositivo";
      const bytes = new Uint8Array(await request.arrayBuffer());
      return this.serial(() => this.uploadFile({ id, name, mimeType, deviceName }, bytes));
    }
    if (request.method === "GET" && url.pathname.endsWith("/files/download")) {
      const path = url.searchParams.get("path") || "";
      const name = url.searchParams.get("name") || "ficheiro";
      return this.downloadFile(path, name);
    }
    if (request.method === "POST" && url.pathname.endsWith("/files/delete")) {
      const payload = await request.json();
      return this.serial(() => this.deleteFile(payload));
    }
    return jsonResponse({ error: "Rota não encontrada." }, 404);
  }

  serial(task) {
    const run = this.tail.then(task, task);
    this.tail = run.catch(() => undefined);
    return run;
  }

  async currentVersionInfo() {
    const cached = await this.ctx.storage.get("current-version-info");
    if (cached && Number.isFinite(Number(cached.version))) return cached;
    const current = await readGitHubState(this.env);
    const info = {
      exists: current.exists,
      version: current.version || 0,
      sha: current.sha || "",
      updatedAt: current.updatedAt || ""
    };
    await this.ctx.storage.put("current-version-info", info);
    return info;
  }

  async publishVersion(info) {
    const clean = {
      exists: true,
      version: Number(info.version) || 0,
      sha: info.sha || "",
      updatedAt: info.updatedAt || new Date().toISOString()
    };
    await this.ctx.storage.put("current-version-info", clean);
    for (const watcher of Array.from(this.watchers)) {
      clearTimeout(watcher.timer);
      this.watchers.delete(watcher);
      watcher.resolve(jsonResponse(Object.assign({ changed: true }, clean)));
    }
  }

  async watchVersion(since) {
    const current = await this.currentVersionInfo();
    if ((current.version || 0) > since) {
      return jsonResponse(Object.assign({ changed: true }, current));
    }
    return new Promise((resolve) => {
      const watcher = { resolve, timer: null };
      watcher.timer = setTimeout(() => {
        this.watchers.delete(watcher);
        resolve(jsonResponse(Object.assign({ changed: false }, current)));
      }, 25000);
      this.watchers.add(watcher);
    });
  }

  async getVersion() {
    const current = await readGitHubState(this.env);
    const info = {
      exists: current.exists,
      version: current.version || 0,
      sha: current.sha || "",
      updatedAt: current.updatedAt || ""
    };
    await this.ctx.storage.put("current-version-info", info);
    return jsonResponse(info);
  }

  async getState() {
    const current = await readGitHubState(this.env);
    const conflictHistory = current.state && current.state.meta && Array.isArray(current.state.meta.syncConflicts)
      ? current.state.meta.syncConflicts.slice(-20)
      : [];
    // Um pull normal não volta a anunciar conflitos antigos como se fossem novos.
    return jsonResponse({ exists: current.exists, state: current.state, sha: current.sha, version: current.version || 0, updatedAt: current.updatedAt || "", conflicts: [], conflictHistory });
  }

  async uploadFile(metadata, bytes) {
    if (!metadata.id || !metadata.name || !bytes || !bytes.byteLength) return jsonResponse({ error: "Ficheiro inválido." }, 400);
    const allowed = /\.(pptx|pdf|png|jpe?g|webp|gif|txt|md)$/i;
    if (!allowed.test(metadata.name)) return jsonResponse({ error: "Formato não suportado. Usa PPTX, PDF, imagem, TXT ou Markdown." }, 400);
    const maxBytes = Math.max(1, Number(this.env.MAX_FILE_BYTES) || 25 * 1024 * 1024);
    if (bytes.byteLength > maxBytes) return jsonResponse({ error: `O ficheiro excede o limite de ${Math.round(maxBytes / 1024 / 1024)} MB.` }, 413);
    const path = buildFilePath(this.env, metadata.id, metadata.name);
    const written = await writeGitHubBinary(this.env, path, bytes, metadata);
    const uploadedAt = new Date().toISOString();
    return jsonResponse({
      ok: true,
      file: {
        id: safeFileId(metadata.id),
        path,
        name: metadata.name,
        size: bytes.byteLength,
        mimeType: metadata.mimeType,
        sha: written.sha,
        commitSha: written.commitSha,
        uploadedAt
      }
    });
  }

  async downloadFile(rawPath, rawName) {
    const path = validateFilePath(this.env, rawPath);
    const source = await downloadGitHubBinary(this.env, path);
    const headers = new Headers(source.headers);
    headers.set("Content-Type", source.headers.get("Content-Type") || "application/octet-stream");
    headers.set("Content-Disposition", `attachment; filename="${safeFileName(rawName)}"`);
    headers.set("X-Twenty-File-Name", safeFileName(rawName));
    headers.set("Cache-Control", "private, no-store");
    return new Response(source.body, { status: 200, headers });
  }

  async deleteFile(payload) {
    if (!payload || !payload.path) return jsonResponse({ error: "Falta o caminho do ficheiro." }, 400);
    const path = validateFilePath(this.env, payload.path);
    const result = await deleteGitHubBinary(this.env, path, payload.deviceName || "Dispositivo");
    return jsonResponse({ ok: true, deleted: result.deleted, commitSha: result.commitSha || "" });
  }

  async forcePush(payload) {
    if (!payload || !payload.operationId || !payload.deviceId || !payload.state || typeof payload.state !== "object") {
      return jsonResponse({ error: "Force push inválido." }, 400);
    }
    const cacheKey = "force-push:" + payload.operationId;
    const cached = await this.ctx.storage.get(cacheKey);
    if (cached) return jsonResponse(cached);

    let attempt = 0;
    while (attempt < 4) {
      attempt += 1;
      const current = await readGitHubState(this.env);
      const remoteRaw = current.exists ? current.state : {};
      const forced = cleanForMerge(payload.state);
      const now = new Date().toISOString();
      const previousConflicts = remoteRaw && remoteRaw.meta && Array.isArray(remoteRaw.meta.syncConflicts)
        ? remoteRaw.meta.syncConflicts
        : [];
      const nextVersion = Math.max(current.version || 0, stateVersion(remoteRaw)) + 1;
      forced.meta = Object.assign({}, forced.meta || {}, {
        revision: nextVersion,
        gitRevision: nextVersion,
        updatedAt: now,
        source: "git-force-push",
        sync: {
          lastMutationId: payload.operationId,
          lastDeviceId: payload.deviceId,
          lastDeviceName: payload.deviceName || "Dispositivo",
          lastCommitAt: now,
          forced: "push"
        },
        syncConflicts: previousConflicts.slice(-50)
      });
      try {
        const written = await writeGitHubState(
          this.env,
          forced,
          current.sha,
          `Twenty: force push · ${payload.deviceName || "dispositivo"}`
        );
        await this.publishVersion({ version: nextVersion, sha: written.sha, updatedAt: now });
        const finalPayload = {
          ok: true,
          forced: "push",
          state: forced,
          sha: written.sha,
          commitSha: written.commitSha,
          version: nextVersion,
          conflicts: []
        };
        await this.ctx.storage.put(cacheKey, finalPayload);
        return jsonResponse(finalPayload);
      } catch (error) {
        if (error.status !== 409 || attempt >= 4) throw error;
      }
    }
    return jsonResponse({ error: "Não foi possível concluir o force push após várias tentativas." }, 409);
  }

  async applyMutation(mutation) {
    if (!mutation || !mutation.id || !mutation.next || !mutation.deviceId) {
      return jsonResponse({ error: "Mutação inválida." }, 400);
    }
    const cacheKey = "mutation:" + mutation.id;
    const cached = await this.ctx.storage.get(cacheKey);
    if (cached) return jsonResponse(cached);

    let attempt = 0;
    let finalPayload;
    while (attempt < 4) {
      attempt += 1;
      const current = await readGitHubState(this.env);
      const remoteRaw = current.exists ? current.state : {};
      const remote = cleanForMerge(remoteRaw);
      const base = mutation.base ? cleanForMerge(mutation.base) : {};
      const local = cleanForMerge(mutation.next);
      const conflicts = [];
      const merged = mergeThreeWay(base, local, remote, "", conflicts) || {};
      const previousConflicts = remoteRaw && remoteRaw.meta && Array.isArray(remoteRaw.meta.syncConflicts)
        ? remoteRaw.meta.syncConflicts
        : [];
      const now = new Date().toISOString();
      const enrichedConflicts = conflicts.map((item) => Object.assign(item, {
        mutationId: mutation.id,
        deviceId: mutation.deviceId,
        deviceName: mutation.deviceName || "Dispositivo",
        detectedAt: now
      }));
      const nextVersion = Math.max(current.version || 0, stateVersion(remoteRaw)) + 1;
      merged.meta = Object.assign({}, merged.meta || {}, {
        revision: nextVersion,
        gitRevision: nextVersion,
        updatedAt: now,
        source: "git-sync",
        sync: {
          lastMutationId: mutation.id,
          lastDeviceId: mutation.deviceId,
          lastDeviceName: mutation.deviceName || "Dispositivo",
          lastCommitAt: now
        },
        syncConflicts: previousConflicts.concat(enrichedConflicts).slice(-50)
      });

      const summary = mutationSummary(base, local);
      const message = `Twenty: ${summary} · ${mutation.deviceName || "dispositivo"}`;
      try {
        const written = await writeGitHubState(this.env, merged, current.sha, message);
        await this.publishVersion({ version: nextVersion, sha: written.sha, updatedAt: now });
        finalPayload = {
          ok: true,
          state: merged,
          sha: written.sha,
          commitSha: written.commitSha,
          version: nextVersion,
          conflicts: enrichedConflicts
        };
        await this.ctx.storage.put(cacheKey, finalPayload);
        return jsonResponse(finalPayload);
      } catch (error) {
        if (error.status !== 409 || attempt >= 4) throw error;
      }
    }
    return jsonResponse({ error: "Não foi possível concluir o commit após várias tentativas." }, 409);
  }
}

function allowedOrigin(request, env) {
  const origin = request.headers.get("Origin");
  if (!origin) return "*";
  const allowed = String(env.ALLOWED_ORIGINS || "").split(",").map((item) => item.trim()).filter(Boolean);
  if (!allowed.length || allowed.includes(origin)) return origin;
  return "";
}

function withCors(response, origin) {
  const headers = new Headers(response.headers);
  headers.set("Access-Control-Allow-Origin", origin || "null");
  headers.set("Access-Control-Allow-Headers", "Content-Type, X-Twenty-Key, X-Twenty-Device");
  headers.set("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  headers.set("Access-Control-Expose-Headers", "Content-Disposition, Content-Length, X-Twenty-File-Name");
  headers.set("Vary", "Origin");
  return new Response(response.body, { status: response.status, statusText: response.statusText, headers });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = allowedOrigin(request, env);
    if (request.method === "OPTIONS") {
      if (!origin) return jsonResponse({ error: "Origem não autorizada." }, 403);
      return withCors(new Response(null, { status: 204 }), origin);
    }
    if (url.pathname === "/health") {
      return withCors(jsonResponse({ ok: true, service: "twenty-git-sync" }), origin || "*");
    }
    if (!origin) return jsonResponse({ error: "Origem não autorizada." }, 403);
    if (!env.SYNC_KEY || request.headers.get("X-Twenty-Key") !== env.SYNC_KEY) {
      return withCors(jsonResponse({ error: "Chave de sincronização inválida." }, 401), origin);
    }
    if (!env.GITHUB_TOKEN || !env.GITHUB_OWNER || !env.GITHUB_REPO) {
      return withCors(jsonResponse({ error: "Worker ainda não configurado." }, 503), origin);
    }
    try {
      const stub = env.SYNC_COORDINATOR.getByName("twenty-study-os");
      const response = await stub.fetch(request);
      return withCors(response, origin);
    } catch (error) {
      console.error("TWENTY_SYNC_ERROR", error);
      const message = error instanceof Error ? error.message : String(error || "Erro interno desconhecido.");
      const status = error && Number.isInteger(error.status) ? error.status : 500;
      return withCors(jsonResponse({ error: message, type: "worker-sync-error" }, status), origin);
    }
  }
};
