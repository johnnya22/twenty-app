(function () {
  "use strict";

  var worker = null;
  var pending = new Map();
  var requestCounter = 0;
  var runtime = {
    modelId: "",
    requestedModel: "",
    warning: "",
    busy: false
  };

  var MODELS = {
    fast: {
      id: "Qwen2.5-0.5B-Instruct-q4f16_1-MLC",
      label: "Qwen 2.5 · 0.5B",
      shortLabel: "Rápido",
      size: "≈ 500–700 MB",
      note: "Menor consumo e melhor escolha para o MacBook Pro 2015."
    },
    quality: {
      id: "Qwen2.5-1.5B-Instruct-q4f16_1-MLC",
      label: "Qwen 2.5 · 1.5B",
      shortLabel: "Qualidade",
      size: "≈ 1–1,3 GB",
      note: "Melhores apontamentos e quizzes; recomendado no Galaxy S24 Ultra."
    }
  };

  function uid(prefix) {
    return (prefix || "ai") + "_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
  }

  function supportsWebGPU() {
    return !!(window.isSecureContext && navigator.gpu && window.Worker);
  }

  function recommendedMode() {
    var ua = navigator.userAgent || "";
    var memory = Number(navigator.deviceMemory) || 0;
    if (/Android/i.test(ua) && (memory >= 6 || /SM-S92/i.test(ua))) return "quality";
    if (/Macintosh|Mac OS X/i.test(ua) && !/Apple Silicon|arm64/i.test(ua)) return "fast";
    return "fast";
  }

  function selectedModel(mode) {
    var resolved = mode === "quality" || mode === "fast" ? mode : recommendedMode();
    return { mode: resolved, ...MODELS[resolved] };
  }

  function ensureWorker() {
    if (worker) return worker;
    worker = new Worker("ai-worker.js?v=18-lesson-ai", { type: "module" });
    worker.addEventListener("message", function (event) {
      var message = event.data || {};
      var item = pending.get(message.id);
      if (!item) return;
      if (message.type === "model-progress") {
        if (item.onProgress) item.onProgress({
          kind: "model",
          progress: message.progress,
          text: message.text || "A preparar o modelo…"
        });
        return;
      }
      if (message.type === "model-warning") {
        runtime.warning = message.text || "Foi usado um modelo de compatibilidade.";
        if (item.onProgress) item.onProgress({ kind: "warning", progress: null, text: runtime.warning });
        return;
      }
      if (message.type === "model-ready") {
        runtime.modelId = message.modelId || "";
        runtime.requestedModel = message.requestedModel || runtime.requestedModel;
        return;
      }
      if (message.type === "result") {
        pending.delete(message.id);
        runtime.busy = pending.size > 0;
        runtime.modelId = message.modelId || runtime.modelId;
        item.resolve(message);
        return;
      }
      if (message.type === "error") {
        pending.delete(message.id);
        runtime.busy = pending.size > 0;
        item.reject(new Error(message.error || "A IA local falhou."));
      }
    });
    function rejectWorkerRequests(message) {
      var error = new Error(message || "O motor de IA foi interrompido pelo navegador.");
      error.workerTerminated = true;
      pending.forEach(function (item) { item.reject(error); });
      pending.clear();
      runtime.busy = false;
      try { worker && worker.terminate(); } catch (_) {}
      worker = null;
    }
    worker.addEventListener("error", function (event) {
      rejectWorkerRequests(event.message || "O motor de IA terminou inesperadamente.");
    });
    worker.addEventListener("messageerror", function () {
      rejectWorkerRequests("O navegador interrompeu a comunicação com o motor de IA.");
    });
    return worker;
  }

  function askWorkerOnce(payload, onProgress) {
    if (!supportsWebGPU()) return Promise.reject(new Error("Este navegador não disponibiliza WebGPU. Abre a Twenty no Chrome atualizado."));
    var id = "ai_req_" + (++requestCounter) + "_" + Date.now();
    runtime.busy = true;
    return new Promise(function (resolve, reject) {
      pending.set(id, { resolve: resolve, reject: reject, onProgress: onProgress });
      ensureWorker().postMessage({ id: id, ...payload });
    });
  }

  function askWorker(payload, onProgress) {
    return askWorkerOnce(payload, onProgress).catch(function (error) {
      if (!error.workerTerminated || payload.__retried) throw error;
      if (onProgress) onProgress({ kind: "warning", progress: null, text: "O navegador libertou a memória da IA. A reiniciar no modelo rápido…" });
      var retry = Object.assign({}, payload, { modelId: MODELS.fast.id, __retried: true });
      return askWorkerOnce(retry, onProgress);
    });
  }

  function cleanText(value) {
    return String(value == null ? "" : value)
      .replace(/\u0000/g, "")
      .replace(/[\t\r ]+/g, " ")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  async function loadJSZip() {
    var module = await import("https://esm.run/jszip@3.10.1");
    return module.default || module;
  }

  function findZipEnd(view) {
    var minimum = Math.max(0, view.byteLength - 65557);
    for (var offset = view.byteLength - 22; offset >= minimum; offset -= 1) {
      if (view.getUint32(offset, true) === 0x06054b50) return offset;
    }
    return -1;
  }

  async function inflateZipEntry(bytes, method) {
    if (method === 0) return bytes;
    if (method !== 8) throw new Error("O PowerPoint usa um método de compressão não suportado.");
    if (typeof DecompressionStream !== "function") throw new Error("NATIVE_ZIP_UNAVAILABLE");
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function readPptxSlidesNative(file) {
    var buffer = await file.arrayBuffer();
    var bytes = new Uint8Array(buffer);
    var view = new DataView(buffer);
    var endOffset = findZipEnd(view);
    if (endOffset < 0) throw new Error("Este ficheiro não parece ser um PowerPoint válido.");
    var entries = view.getUint16(endOffset + 10, true);
    var directoryOffset = view.getUint32(endOffset + 16, true);
    var decoder = new TextDecoder("utf-8");
    var result = [];
    var cursor = directoryOffset;

    for (var index = 0; index < entries; index += 1) {
      if (cursor + 46 > view.byteLength || view.getUint32(cursor, true) !== 0x02014b50) {
        throw new Error("O índice interno deste PowerPoint está danificado.");
      }
      var flags = view.getUint16(cursor + 8, true);
      var method = view.getUint16(cursor + 10, true);
      var compressedSize = view.getUint32(cursor + 20, true);
      var fileNameLength = view.getUint16(cursor + 28, true);
      var extraLength = view.getUint16(cursor + 30, true);
      var commentLength = view.getUint16(cursor + 32, true);
      var localOffset = view.getUint32(cursor + 42, true);
      var nameStart = cursor + 46;
      var name = decoder.decode(bytes.slice(nameStart, nameStart + fileNameLength));

      if (/^ppt\/slides\/slide\d+\.xml$/i.test(name)) {
        if (flags & 1) throw new Error("Este PowerPoint está protegido por palavra-passe.");
        if (localOffset + 30 > view.byteLength || view.getUint32(localOffset, true) !== 0x04034b50) {
          throw new Error("Um dos slides está danificado.");
        }
        var localNameLength = view.getUint16(localOffset + 26, true);
        var localExtraLength = view.getUint16(localOffset + 28, true);
        var dataStart = localOffset + 30 + localNameLength + localExtraLength;
        var dataEnd = dataStart + compressedSize;
        if (dataEnd > bytes.length) throw new Error("Um dos slides está incompleto.");
        var inflated = await inflateZipEntry(bytes.slice(dataStart, dataEnd), method);
        result.push({ path: name, xml: decoder.decode(inflated) });
      }
      cursor += 46 + fileNameLength + extraLength + commentLength;
    }
    result.sort(function (a, b) { return slideNumber(a.path) - slideNumber(b.path); });
    return result;
  }

  async function readPptxSlides(file) {
    try {
      return await readPptxSlidesNative(file);
    } catch (error) {
      if (String(error && error.message) !== "NATIVE_ZIP_UNAVAILABLE") throw error;
      var JSZip = await loadJSZip();
      var zip = await JSZip.loadAsync(await file.arrayBuffer());
      var paths = Object.keys(zip.files)
        .filter(function (path) { return /^ppt\/slides\/slide\d+\.xml$/i.test(path); })
        .sort(function (a, b) { return slideNumber(a) - slideNumber(b); });
      return Promise.all(paths.map(async function (path) {
        return { path: path, xml: await zip.file(path).async("string") };
      }));
    }
  }

  function slideNumber(path) {
    var match = String(path).match(/slide(\d+)\.xml$/i);
    return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
  }

  function xmlParagraphs(xmlText) {
    var doc = new DOMParser().parseFromString(xmlText, "application/xml");
    if (doc.querySelector("parsererror")) throw new Error("Um dos slides contém XML inválido.");
    var paragraphs = Array.from(doc.getElementsByTagNameNS("*", "p")).map(function (paragraph) {
      return Array.from(paragraph.getElementsByTagNameNS("*", "t"))
        .map(function (node) { return node.textContent || ""; })
        .join("")
        .trim();
    }).filter(Boolean);
    if (!paragraphs.length) {
      paragraphs = Array.from(doc.getElementsByTagNameNS("*", "t"))
        .map(function (node) { return (node.textContent || "").trim(); })
        .filter(Boolean);
    }
    return paragraphs;
  }

  async function extractPptx(file, onProgress) {
    if (!file) throw new Error("Escolhe um PowerPoint primeiro.");
    if (!/\.pptx$/i.test(file.name || "")) throw new Error("A Twenty lê ficheiros .pptx. O formato antigo .ppt não é suportado.");
    if (file.size > 25 * 1024 * 1024) throw new Error("Este PowerPoint tem mais de 25 MB. Comprime imagens/vídeos ou divide a apresentação antes de a sincronizar.");
    if (onProgress) onProgress({ progress: 2, text: "A abrir o PowerPoint…" });
    var slideEntries = await readPptxSlides(file);
    if (!slideEntries.length) throw new Error("Não encontrei slides dentro deste ficheiro.");

    var slides = [];
    for (var index = 0; index < slideEntries.length; index += 1) {
      var xml = slideEntries[index].xml;
      var paragraphs = xmlParagraphs(xml);
      var number = slideNumber(slideEntries[index].path);
      var title = cleanText(paragraphs[0] || "Slide " + number).slice(0, 180);
      var text = cleanText(paragraphs.join("\n")).slice(0, 6000);
      slides.push({ number: number, title: title, text: text });
      if (onProgress) onProgress({
        progress: 5 + Math.round(((index + 1) / paths.length) * 25),
        text: "A extrair o slide " + (index + 1) + " de " + slideEntries.length + "…"
      });
    }
    var withText = slides.filter(function (slide) { return slide.text; });
    if (!withText.length) throw new Error("Os slides não têm texto selecionável. Se forem imagens digitalizadas, será necessário OCR.");
    var totalText = withText.reduce(function (sum, slide) { return sum + slide.text.length; }, 0);
    if (totalText > 650000) throw new Error("Esta apresentação tem texto a mais para sincronizar com segurança. Divide-a em duas apresentações e tenta novamente.");
    return {
      id: uid("slides"),
      fileName: file.name,
      fileSize: file.size,
      slideCount: slides.length,
      slides: slides,
      extractedAt: new Date().toISOString()
    };
  }

  function chunkSlides(slides) {
    var chunks = [];
    var current = [];
    var length = 0;
    (slides || []).forEach(function (slide) {
      var block = "[Slide " + slide.number + "] " + slide.title + "\n" + slide.text;
      if (current.length && (current.length >= 5 || length + block.length > 4800)) {
        chunks.push(current);
        current = [];
        length = 0;
      }
      current.push(slide);
      length += block.length;
    });
    if (current.length) chunks.push(current);
    return chunks;
  }

  function slidesPrompt(slides) {
    return slides.map(function (slide) {
      return "[SLIDE " + slide.number + "]\nTÍTULO: " + slide.title + "\n" + slide.text;
    }).join("\n\n");
  }

  function parseJSON(raw) {
    var text = String(raw || "").trim()
      .replace(/^```(?:json)?\s*/i, "")
      .replace(/\s*```$/i, "");
    var first = text.indexOf("{");
    var last = text.lastIndexOf("}");
    if (first >= 0 && last > first) text = text.slice(first, last + 1);
    return JSON.parse(text);
  }

  async function completeJSON(modelId, messages, options, onProgress) {
    options = options || {};
    var result = await askWorker({
      type: "complete",
      modelId: modelId,
      messages: messages,
      json: true,
      maxTokens: options.maxTokens || 900,
      temperature: options.temperature == null ? 0.2 : options.temperature,
      seed: options.seed || 20
    }, onProgress);
    try {
      return { data: parseJSON(result.content), modelId: result.modelId, raw: result.content };
    } catch (_) {
      var repair = await askWorker({
        type: "complete",
        modelId: modelId,
        messages: [
          { role: "system", content: "Corrige JSON. Responde apenas com um objeto JSON válido, sem markdown nem comentários." },
          { role: "user", content: "Converte esta resposta em JSON válido sem perder informação:\n\n" + String(result.content || "").slice(0, 12000) }
        ],
        json: true,
        maxTokens: options.maxTokens || 900,
        temperature: 0,
        seed: 20
      }, onProgress);
      return { data: parseJSON(repair.content), modelId: repair.modelId, raw: repair.content };
    }
  }

  function compactDigest(digest) {
    return {
      sourceSlides: Array.isArray(digest.sourceSlides) ? digest.sourceSlides.slice(0, 12) : [],
      summary: cleanText(digest.summary || "").slice(0, 1400),
      keyPoints: Array.isArray(digest.keyPoints) ? digest.keyPoints.slice(0, 8).map(function (item) { return cleanText(item).slice(0, 260); }) : [],
      terms: Array.isArray(digest.terms) ? digest.terms.slice(0, 8).map(function (item) {
        return { term: cleanText(item.term).slice(0, 120), definition: cleanText(item.definition).slice(0, 260) };
      }) : []
    };
  }

  async function digestChunks(modelId, chunks, onProgress) {
    var digests = [];
    for (var index = 0; index < chunks.length; index += 1) {
      if (onProgress) onProgress({
        kind: "generation",
        progress: 42 + Math.round((index / Math.max(1, chunks.length)) * 25),
        text: "A compreender os slides " + chunks[index][0].number + "–" + chunks[index][chunks[index].length - 1].number + "…"
      });
      var response = await completeJSON(modelId, [
        {
          role: "system",
          content: "És um assistente académico rigoroso. Usa apenas o conteúdo fornecido, escreve em português de Portugal e nunca inventes factos. Mantém os números dos slides de origem."
        },
        {
          role: "user",
          content: "Resume este bloco para servir de base a apontamentos e quizzes. Responde apenas neste JSON: {\"sourceSlides\":[1],\"summary\":\"...\",\"keyPoints\":[\"...\"],\"terms\":[{\"term\":\"...\",\"definition\":\"...\"}]}. Sê compacto.\n\n" + slidesPrompt(chunks[index])
        }
      ], { maxTokens: 700, seed: 20 + index }, onProgress);
      digests.push(compactDigest(response.data || {}));
    }
    return digests;
  }

  async function reduceDigests(modelId, digests, onProgress) {
    var current = digests.slice();
    var round = 0;
    while (JSON.stringify(current).length > 12500 && current.length > 2) {
      round += 1;
      var next = [];
      for (var i = 0; i < current.length; i += 4) {
        var group = current.slice(i, i + 4);
        if (onProgress) onProgress({ kind: "generation", progress: 68, text: "A juntar os blocos de matéria…" });
        var response = await completeJSON(modelId, [
          { role: "system", content: "Compacta matéria académica em português de Portugal. Preserva as referências aos slides e não inventes conteúdo." },
          { role: "user", content: "Funde os blocos seguintes. Responde apenas JSON no formato {\"sourceSlides\":[1],\"summary\":\"...\",\"keyPoints\":[\"...\"],\"terms\":[{\"term\":\"...\",\"definition\":\"...\"}]}.\n\n" + JSON.stringify(group) }
        ], { maxTokens: 750, seed: 60 + round + i }, onProgress);
        next.push(compactDigest(response.data || {}));
      }
      current = next;
    }
    return current;
  }

  function difficultyInstruction(value) {
    if (value === "easy") return "maioritariamente fáceis, focadas em definições e compreensão direta";
    if (value === "hard") return "exigentes, com aplicação e comparação de conceitos, mas sempre respondíveis pelos slides";
    if (value === "medium") return "de dificuldade média, misturando compreensão e aplicação";
    return "com dificuldade variada e adequada ao conteúdo";
  }

  function normalizeGeneratedGameType(value) {
    value = cleanText(value || "").toLowerCase();
    if (["formula", "short-answer", "write", "written", "type-answer"].indexOf(value) >= 0) return "type-answer";
    if (["spot-error", "find-error", "incorrect", "error"].indexOf(value) >= 0) return "spot-error";
    if (["order", "ordering", "sort"].indexOf(value) >= 0) return "order";
    if (["match", "matching", "pairs"].indexOf(value) >= 0) return "match";
    if (["memory", "sequence"].indexOf(value) >= 0) return "memory";
    return "choice";
  }

  function normalizeQuiz(data, count) {
    var questions = Array.isArray(data && data.questions) ? data.questions : [];
    return questions.slice(0, count).map(function (question, index) {
      question = question && typeof question === "object" ? question : {};
      var gameType = normalizeGeneratedGameType(question.gameType || question.mode);
      var normalized = {
        id: uid("aiq"),
        mode: gameType === "choice" ? "multiple-choice" : gameType,
        gameType: gameType,
        prompt: cleanText(question.question || question.prompt || "Pergunta " + (index + 1)),
        explanation: cleanText(question.explanation || ""),
        sourceSlides: Array.isArray(question.sourceSlides) ? question.sourceSlides.map(Number).filter(Number.isFinite) : [],
        difficulty: ["easy", "medium", "hard"].indexOf(question.difficulty) >= 0 ? question.difficulty : "medium",
        images: [],
        options: [],
        answerIndex: null,
        answerText: "",
        acceptedAnswers: [],
        orderItems: [],
        correctOrder: [],
        matchPairs: [],
        sequence: []
      };
      if (gameType === "type-answer") {
        normalized.answerText = cleanText(question.answerText || question.answer || "");
        normalized.acceptedAnswers = Array.isArray(question.acceptedAnswers) ? question.acceptedAnswers.map(cleanText).filter(Boolean).slice(0, 8) : [];
        if (normalized.answerText && normalized.acceptedAnswers.indexOf(normalized.answerText) < 0) normalized.acceptedAnswers.unshift(normalized.answerText);
        if (!normalized.acceptedAnswers.length) return null;
        return normalized;
      }
      if (gameType === "order") {
        normalized.correctOrder = Array.isArray(question.correctOrder) ? question.correctOrder.map(cleanText).filter(Boolean).slice(0, 8) : [];
        normalized.orderItems = Array.isArray(question.orderItems) ? question.orderItems.map(cleanText).filter(Boolean).slice(0, 8) : normalized.correctOrder.slice();
        if (normalized.correctOrder.length < 2) return null;
        return normalized;
      }
      if (gameType === "match") {
        normalized.matchPairs = Array.isArray(question.matchPairs) ? question.matchPairs.map(function (pair) {
          if (Array.isArray(pair)) return { left: cleanText(pair[0]), right: cleanText(pair[1]) };
          return { left: cleanText(pair && pair.left), right: cleanText(pair && pair.right) };
        }).filter(function (pair) { return pair.left && pair.right; }).slice(0, 6) : [];
        if (normalized.matchPairs.length < 2) return null;
        return normalized;
      }
      if (gameType === "memory") {
        normalized.sequence = Array.isArray(question.sequence) ? question.sequence.map(cleanText).filter(Boolean).slice(0, 8) : [];
        if (normalized.sequence.length < 3) return null;
        return normalized;
      }
      var options = Array.isArray(question.options) ? question.options.map(cleanText).filter(Boolean).slice(0, 4) : [];
      while (options.length < 4) options.push("Opção " + (options.length + 1));
      var correctIndex = Number(question.correctIndex != null ? question.correctIndex : question.answerIndex);
      if (!Number.isInteger(correctIndex) || correctIndex < 0 || correctIndex >= options.length) correctIndex = 0;
      normalized.options = options;
      normalized.answerIndex = correctIndex;
      return normalized;
    }).filter(Boolean);
  }

  function normalizeLessonPlan(data, source) {
    data = data && typeof data === "object" ? data : {};
    var homework = data.homework && typeof data.homework === "object" ? data.homework : {};
    var fallbackTitle = cleanText(source && source.fileName || "Nova aula").replace(/\.pptx$/i, "");
    return {
      title: cleanText(data.title || fallbackTitle || "Nova aula"),
      description: cleanText(data.description || data.summary || ""),
      objectives: Array.isArray(data.objectives) ? data.objectives.map(cleanText).filter(Boolean).slice(0, 6) : [],
      quizTitle: cleanText(data.quizTitle || ("Quiz da aula · " + (data.title || fallbackTitle || "Nova aula"))),
      homework: {
        title: cleanText(homework.title || ("TPC · " + (data.title || fallbackTitle || "Nova aula"))),
        instructions: cleanText(homework.instructions || homework.description || ""),
        estimatedMinutes: Math.max(10, Math.min(180, Number(homework.estimatedMinutes) || 30)),
        steps: Array.isArray(homework.steps) ? homework.steps.map(cleanText).filter(Boolean).slice(0, 6) : [],
        dueInDays: Math.max(0, Math.min(30, Number(homework.dueInDays) || 3))
      }
    };
  }

  function normalizeFlashcards(data) {
    var cards = Array.isArray(data && data.cards) ? data.cards : [];
    return cards.slice(0, 40).map(function (card) {
      return {
        id: uid("card"),
        front: cleanText(card.front || card.question || ""),
        back: cleanText(card.back || card.answer || ""),
        sourceSlides: Array.isArray(card.sourceSlides) ? card.sourceSlides.map(Number).filter(Number.isFinite) : []
      };
    }).filter(function (card) { return card.front && card.back; });
  }

  function normalizeNotes(data) {
    return {
      title: cleanText(data && data.title || "Apontamentos dos slides"),
      overview: cleanText(data && data.overview || data && data.summary || ""),
      sections: Array.isArray(data && data.sections) ? data.sections.slice(0, 16).map(function (section) {
        return {
          heading: cleanText(section.heading || "Tópico"),
          content: cleanText(section.content || ""),
          sourceSlides: Array.isArray(section.sourceSlides) ? section.sourceSlides.map(Number).filter(Number.isFinite) : []
        };
      }) : [],
      keyTakeaways: Array.isArray(data && data.keyTakeaways) ? data.keyTakeaways.slice(0, 14).map(cleanText).filter(Boolean) : []
    };
  }

  async function generateStudyPack(source, options, onProgress) {
    options = options || {};
    if (!source || !Array.isArray(source.slides) || !source.slides.length) throw new Error("Importa um PowerPoint antes de gerar conteúdo.");
    if (!supportsWebGPU()) throw new Error("WebGPU não está disponível. Usa o Chrome atualizado no S24 Ultra ou no Mac.");
    try { if (navigator.storage && navigator.storage.persist) await navigator.storage.persist(); } catch (_) { /* opcional */ }

    runtime.warning = "";
    var selected = selectedModel(options.modelMode || "auto");
    if (onProgress) onProgress({ kind: "model", progress: 31, text: "A preparar " + selected.label + "…" });
    var loaded = await askWorker({ type: "load", modelId: selected.id }, function (report) {
      if (!onProgress) return;
      var scaled = report.progress == null ? null : 31 + Math.round(Number(report.progress) * 10);
      onProgress({ kind: report.kind, progress: scaled, text: report.text });
    });
    var activeModelId = loaded && loaded.modelId ? loaded.modelId : selected.id;

    var chunks = chunkSlides(source.slides);
    var digests = await digestChunks(activeModelId, chunks, onProgress);
    digests = await reduceDigests(activeModelId, digests, onProgress);
    var material = JSON.stringify(digests);
    var output = options.output || "all";
    var wantsNotes = output === "all" || output === "notes" || output === "summary" || output === "lesson";
    var wantsQuiz = output === "all" || output === "quiz" || output === "lesson";
    var wantsCards = output === "all" || output === "flashcards";
    var actualModel = runtime.modelId || activeModelId;
    var notes = null;
    var summary = "";
    var questions = [];
    var flashcards = [];
    var lessonPlan = null;

    if (wantsNotes) {
      if (onProgress) onProgress({ kind: "generation", progress: 72, text: output === "summary" ? "A criar o resumo…" : "A organizar os apontamentos…" });
      var notesResponse = await completeJSON(activeModelId, [
        { role: "system", content: "És um explicador universitário. Escreve em português de Portugal, usa apenas a matéria fornecida e cita os slides de origem. Não inventes." },
        { role: "user", content: (output === "summary" ? "Cria um resumo rápido e muito claro." : "Cria apontamentos completos, organizados para estudar.") + " Responde apenas JSON: {\"title\":\"...\",\"overview\":\"...\",\"sections\":[{\"heading\":\"...\",\"content\":\"...\",\"sourceSlides\":[1]}],\"keyTakeaways\":[\"...\"]}.\n\nMATÉRIA:\n" + material }
      ], { maxTokens: output === "summary" ? 650 : 1050, seed: 101 }, onProgress);
      notes = normalizeNotes(notesResponse.data || {});
      summary = notes.overview;
      actualModel = notesResponse.modelId || actualModel;
    }

    if (wantsQuiz) {
      if (onProgress) onProgress({ kind: "generation", progress: 82, text: "A criar " + (options.questionCount || 10) + " perguntas…" });
      var count = Math.max(5, Math.min(30, Number(options.questionCount) || 10));
      var lessonQuiz = output === "lesson";
      var quizResponse = await completeJSON(activeModelId, [
        { role: "system", content: lessonQuiz ? "És um professor universitário a criar um Quiz-Aula curto e variado, fiel aos slides e em português de Portugal. Usa apenas a matéria fornecida. Mistura atividades adequadas ao conteúdo: choice, spot-error, type-answer, order, match ou memory. Não inventes." : "És um professor universitário a criar um quiz fiel aos slides. Português de Portugal. Não uses conhecimento externo. Cada pergunta tem exatamente quatro opções plausíveis e uma só correta." },
        { role: "user", content: lessonQuiz ? ("Cria " + count + " perguntas " + difficultyInstruction(options.difficulty) + ". Responde apenas JSON no formato {\"questions\":[{\"gameType\":\"choice|spot-error|type-answer|order|match|memory\",\"question\":\"...\",\"options\":[\"A\",\"B\",\"C\",\"D\"],\"correctIndex\":0,\"answerText\":\"...\",\"acceptedAnswers\":[\"...\"],\"orderItems\":[\"...\"],\"correctOrder\":[\"...\"],\"matchPairs\":[{\"left\":\"...\",\"right\":\"...\"}],\"sequence\":[\"...\"],\"explanation\":\"...\",\"sourceSlides\":[1],\"difficulty\":\"easy|medium|hard\"}]}. Preenche apenas os campos necessários para cada gameType. Usa pelo menos dois tipos diferentes, mas evita memória quando o conteúdo não tiver uma sequência natural.\n\nMATÉRIA:\n" + material) : ("Cria " + count + " perguntas " + difficultyInstruction(options.difficulty) + ". Responde apenas JSON: {\"questions\":[{\"question\":\"...\",\"options\":[\"A\",\"B\",\"C\",\"D\"],\"correctIndex\":0,\"explanation\":\"...\",\"sourceSlides\":[1],\"difficulty\":\"easy|medium|hard\"}]}. Evita perguntas repetidas.\n\nMATÉRIA:\n" + material) }
      ], { maxTokens: Math.min(1800, 420 + count * 88), seed: 202 }, onProgress);
      questions = normalizeQuiz(quizResponse.data || {}, count);
      actualModel = quizResponse.modelId || actualModel;
    }

    if (wantsCards) {
      if (onProgress) onProgress({ kind: "generation", progress: 91, text: "A transformar a matéria em flashcards…" });
      var cardCount = Math.max(10, Math.min(30, Number(options.questionCount) || 15));
      var cardResponse = await completeJSON(activeModelId, [
        { role: "system", content: "Cria flashcards curtos, claros e fiéis aos slides, em português de Portugal. Não inventes informação." },
        { role: "user", content: "Cria até " + cardCount + " flashcards. Responde apenas JSON: {\"cards\":[{\"front\":\"pergunta ou conceito\",\"back\":\"resposta clara\",\"sourceSlides\":[1]}]}.\n\nMATÉRIA:\n" + material }
      ], { maxTokens: 950, seed: 303 }, onProgress);
      flashcards = normalizeFlashcards(cardResponse.data || {});
      actualModel = cardResponse.modelId || actualModel;
    }


    if (output === "lesson") {
      if (onProgress) onProgress({ kind: "generation", progress: 94, text: "A preparar a aula e o TPC…" });
      var lessonResponse = await completeJSON(activeModelId, [
        { role: "system", content: "És um assistente académico universitário. Analisa apenas os slides fornecidos e cria uma proposta prática de aula em português de Portugal. Não inventes tópicos, datas, professores nem requisitos que não estejam na matéria." },
        { role: "user", content: "Cria os metadados da aula e um TPC curto que ajude a aplicar ou consolidar a matéria. O título deve ser específico e curto. A descrição deve explicar em 2 ou 3 frases o que é abordado. Responde apenas JSON: {\"title\":\"...\",\"description\":\"...\",\"objectives\":[\"...\"],\"quizTitle\":\"Quiz da aula · ...\",\"homework\":{\"title\":\"TPC · ...\",\"instructions\":\"...\",\"estimatedMinutes\":30,\"steps\":[\"...\"],\"dueInDays\":3}}. O TPC deve ser realizável apenas com o conteúdo dos slides, com 2 a 5 passos e sem fingir que será enviado ao professor.\n\nMATÉRIA:\n" + material }
      ], { maxTokens: 760, seed: 404 }, onProgress);
      lessonPlan = normalizeLessonPlan(lessonResponse.data || {}, source);
      actualModel = lessonResponse.modelId || actualModel;
    }

    if (onProgress) onProgress({ kind: "done", progress: 100, text: "Conteúdo criado e pronto a guardar." });
    return {
      id: uid("aiproject"),
      title: notes && notes.title ? notes.title : source.fileName.replace(/\.pptx$/i, ""),
      fileName: source.fileName,
      fileSize: source.fileSize,
      slideCount: source.slideCount,
      slides: source.slides,
      summary: summary,
      notes: notes,
      flashcards: flashcards,
      quizQuestions: questions,
      lessonPlan: lessonPlan,
      output: output,
      difficulty: options.difficulty || "auto",
      questionCount: Number(options.questionCount) || 10,
      requestedModel: selected.id,
      modelId: actualModel,
      modelMode: selected.mode,
      createdAt: new Date().toISOString(),
      warning: runtime.warning || ""
    };
  }

  function resetWorker() {
    pending.forEach(function (item) { item.reject(new Error("Geração cancelada.")); });
    pending.clear();
    if (worker) worker.terminate();
    worker = null;
    runtime.busy = false;
    runtime.modelId = "";
  }

  window.TwentyAI = {
    MODELS: MODELS,
    supportsWebGPU: supportsWebGPU,
    recommendedMode: recommendedMode,
    selectedModel: selectedModel,
    extractPptx: extractPptx,
    generateStudyPack: generateStudyPack,
    resetWorker: resetWorker,
    getRuntime: function () { return { ...runtime }; }
  };
})();
