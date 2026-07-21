import {
  CreateMLCEngine,
  prebuiltAppConfig
} from "https://esm.run/@mlc-ai/web-llm@0.2.84";

const APP_CONFIG = {
  ...prebuiltAppConfig,
  cacheBackend: "cache"
};

const FALLBACK_MODELS = {
  "Qwen2.5-0.5B-Instruct-q4f16_1-MLC": "Qwen2-0.5B-Instruct-q4f16_1-MLC",
  "Qwen2.5-1.5B-Instruct-q4f16_1-MLC": "Qwen2-1.5B-Instruct-q4f16_1-MLC"
};

let engine = null;
let loadedModel = "";
let requestedModel = "";

function reply(id, type, payload) {
  self.postMessage({ id, type, ...(payload || {}) });
}

function errorMessage(error) {
  return error && error.message ? error.message : String(error || "Erro desconhecido");
}

async function unloadEngine() {
  if (engine && typeof engine.unload === "function") {
    try { await engine.unload(); } catch (_) { /* libertação opcional */ }
  }
  engine = null;
  loadedModel = "";
}

async function createEngine(id, modelId) {
  requestedModel = modelId;
  reply(id, "model-progress", { progress: 0.01, text: "A preparar o motor de IA…", requestedModel: modelId });
  return CreateMLCEngine(modelId, {
    appConfig: APP_CONFIG,
    initProgressCallback(report) {
      const value = Number(report && report.progress);
      reply(id, "model-progress", {
        progress: Number.isFinite(value) ? value : null,
        text: report && report.text ? report.text : "A descarregar o modelo…",
        requestedModel: modelId
      });
    }
  });
}

async function ensureModel(id, modelId) {
  if (engine && loadedModel === modelId) return loadedModel;
  await unloadEngine();
  try {
    engine = await createEngine(id, modelId);
    loadedModel = modelId;
    reply(id, "model-ready", { modelId: loadedModel, requestedModel });
    return loadedModel;
  } catch (firstError) {
    const fallback = FALLBACK_MODELS[modelId];
    if (!fallback) throw firstError;
    reply(id, "model-warning", {
      text: "O modelo Qwen 2.5 não carregou neste navegador. A usar a variante Qwen 2 compatível.",
      requestedModel: modelId,
      fallbackModel: fallback
    });
    await unloadEngine();
    engine = await createEngine(id, fallback);
    loadedModel = fallback;
    reply(id, "model-ready", { modelId: loadedModel, requestedModel: modelId, fallback: true });
    return loadedModel;
  }
}

async function complete(id, payload) {
  const actualModel = await ensureModel(id, payload.modelId);
  const request = {
    messages: payload.messages,
    temperature: payload.temperature == null ? 0.25 : payload.temperature,
    top_p: payload.topP == null ? 0.9 : payload.topP,
    max_tokens: payload.maxTokens || 900,
    seed: payload.seed || 20
  };
  if (payload.json) request.response_format = { type: "json_object" };

  let response;
  try {
    response = await engine.chat.completions.create(request);
  } catch (error) {
    if (!payload.json) throw error;
    delete request.response_format;
    response = await engine.chat.completions.create(request);
  }
  const content = response && response.choices && response.choices[0] && response.choices[0].message
    ? response.choices[0].message.content
    : "";
  reply(id, "result", {
    content: String(content || ""),
    modelId: actualModel,
    requestedModel: payload.modelId,
    usage: response && response.usage ? response.usage : null
  });
}

self.onmessage = async function (event) {
  const message = event.data || {};
  const id = message.id;
  try {
    if (message.type === "load") {
      const modelId = await ensureModel(id, message.modelId);
      reply(id, "result", { content: "", modelId, requestedModel: message.modelId });
      return;
    }
    if (message.type === "complete") {
      await complete(id, message);
      return;
    }
    if (message.type === "unload") {
      await unloadEngine();
      reply(id, "result", { content: "", modelId: "" });
      return;
    }
    throw new Error("Pedido de IA desconhecido.");
  } catch (error) {
    reply(id, "error", { error: errorMessage(error), requestedModel: message.modelId || requestedModel });
  }
};
