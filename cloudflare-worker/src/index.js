import { DurableObject } from "cloudflare:workers";

const VOLATILE_META_KEYS = new Set([
  "revision",
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

function githubHeaders(env) {
  return {
    "Accept": "application/vnd.github+json",
    "Authorization": "Bearer " + env.GITHUB_TOKEN,
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "Twenty-Study-OS-Sync"
  };
}

function githubUrl(env) {
  const owner = encodeURIComponent(env.GITHUB_OWNER);
  const repo = encodeURIComponent(env.GITHUB_REPO);
  const path = String(env.STATE_PATH || "data/twenty-state.json").split("/").map(encodeURIComponent).join("/");
  return `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
}

async function readGitHubState(env) {
  const url = githubUrl(env) + "?ref=" + encodeURIComponent(env.GITHUB_BRANCH || "main");
  const response = await fetch(url, { headers: githubHeaders(env) });
  if (response.status === 404) return { exists: false, state: null, sha: "" };
  if (!response.ok) {
    const body = await response.text();
    throw new Error("GitHub GET " + response.status + ": " + body.slice(0, 300));
  }
  const payload = await response.json();
  const state = JSON.parse(base64ToUtf8(payload.content));
  return { exists: true, state, sha: payload.sha || "" };
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
  }

  async fetch(request) {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname.endsWith("/state")) {
      return this.serial(() => this.getState());
    }
    if (request.method === "POST" && url.pathname.endsWith("/sync")) {
      const mutation = await request.json();
      return this.serial(() => this.applyMutation(mutation));
    }
    return jsonResponse({ error: "Rota não encontrada." }, 404);
  }

  serial(task) {
    const run = this.tail.then(task, task);
    this.tail = run.catch(() => undefined);
    return run;
  }

  async getState() {
    const current = await readGitHubState(this.env);
    const conflicts = current.state && current.state.meta && Array.isArray(current.state.meta.syncConflicts)
      ? current.state.meta.syncConflicts.slice(-20)
      : [];
    return jsonResponse({ exists: current.exists, state: current.state, sha: current.sha, conflicts });
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
      merged.meta = Object.assign({}, merged.meta || {}, {
        revision: Number(remoteRaw && remoteRaw.meta && remoteRaw.meta.revision || 0) + 1,
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
        finalPayload = {
          ok: true,
          state: merged,
          sha: written.sha,
          commitSha: written.commitSha,
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
  headers.set("Access-Control-Allow-Headers", "Content-Type, X-Twenty-Key");
  headers.set("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
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
    const stub = env.SYNC_COORDINATOR.getByName("twenty-study-os");
    const response = await stub.fetch(request);
    return withCors(response, origin);
  }
};
