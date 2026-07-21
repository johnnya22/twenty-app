(function () {
  "use strict";

  var DB = window.TwentyDB;
  var Sync = window.TwentySync;
  var AI = window.TwentyAI;
  var app = document.getElementById("app");
  var view = document.getElementById("view");
  var modalRoot = document.getElementById("modalRoot");
  var toastRegion = document.getElementById("toastRegion");
  var searchInput = document.getElementById("globalSearch");
  var searchResults = document.getElementById("searchResults");
  var importInput = document.getElementById("jsonImportInput");
  var pptxInput = document.getElementById("pptxImportInput");
  var syncActivity = document.getElementById("syncActivity");
  var syncActivityTitle = document.getElementById("syncActivityTitle");
  var syncActivityDetail = document.getElementById("syncActivityDetail");
  var syncProgressBar = document.getElementById("syncProgressBar");
  var manualSyncActivity = false;
  var syncActivityHideTimer = null;
  var state = null;
  var route = { name: "home", id: null, tab: "overview" };
  var onboarding = null;
  var activeObjectUrl = null;
  var searchTimer = null;
  var externalCheckTimer = null;
  var beOnlineTimer = null;
  var calendarCursor = todayISO();
  var guidedTour = null;
  var activeImageObjectUrls = [];
  var draggedStudyPayload = null;
  var aiDraft = null;
  var aiBusy = false;
  var aiTransferRequest = null;
  var aiProgress = { active: false, progress: null, title: "", detail: "" };
  var CANTEEN_API_URL = "https://sas.unl.pt/wp-json/wp/v2/pages/326?_fields=acf,link";
  var CANTEEN_INFO_API_URL = "https://sas.unl.pt/wp-json/wp/v2/pages/309?_fields=acf,link,modified";
  var CANTEEN_PAGE_URL = "https://sas.unl.pt/alimentacao/cantina-da-faculdade-de-ciencias-e-tecnologia-fct/";
  var CANTEEN_INFO_PAGE_URL = "https://sas.unl.pt/alimentacao/";
  var CANTEEN_CACHE_KEY = "twenty-canteen-menu-v2";
  var canteenMenu = loadCachedCanteen();
  var canteenStatus = canteenMenu ? "ready" : "idle";
  var canteenError = "";
  var canteenChecked = false;
  var canteenLoadPromise = null;
  var canteenSelectedDate = null;
  var canteenClockTimer = null;
  var COLORS = ["#a99df7", "#ff92ae", "#ffad72", "#79cdb8", "#80bee8", "#f3e873", "#cab6ea", "#87d7df"];
  var WEEKDAYS = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  var SHORT_WEEKDAYS = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  var ENTITY_ARRAYS = ["semesters", "courses", "schedule", "assessments", "events", "tasks", "lessons", "materials", "pastExams", "questions", "quizzes", "grades", "studyBlocks", "weeklyReviews", "aiProjects"];

  function uid(prefix) {
    return (prefix || "id") + "_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
  }

  function esc(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function attr(value) { return esc(value); }

  function nl2br(value) { return esc(value).replace(/\n/g, "<br>"); }

  function clone(value) { return JSON.parse(JSON.stringify(value)); }

  function asArray(value) { return Array.isArray(value) ? value : []; }

  function clamp(value, min, max) { return Math.min(max, Math.max(min, Number(value) || 0)); }

  function showSyncActivity(options) {
    options = options || {};
    if (!syncActivity) return;
    clearTimeout(syncActivityHideTimer);
    syncActivity.hidden = false;
    syncActivity.classList.toggle("is-blocking", !!options.blocking);
    syncActivityTitle.textContent = options.title || "A sincronizar dados…";
    syncActivityDetail.textContent = options.detail || "Aguarda enquanto a Twenty confirma a versão mais recente.";
    var track = syncProgressBar && syncProgressBar.parentElement;
    var hasProgress = options.progress !== null && options.progress !== undefined && Number.isFinite(Number(options.progress));
    if (track) track.classList.toggle("is-indeterminate", !hasProgress);
    if (syncProgressBar) syncProgressBar.style.width = hasProgress ? clamp(options.progress, 2, 100) + "%" : "38%";
    if (app) app.setAttribute("aria-busy", "true");
  }

  function setManualSyncActivity(title, detail, progress, blocking) {
    manualSyncActivity = true;
    showSyncActivity({ title: title, detail: detail, progress: progress, blocking: blocking !== false });
  }

  function finishManualSyncActivity(success) {
    if (!syncActivity) { manualSyncActivity = false; return; }
    if (success) {
      showSyncActivity({ title: "Sincronização concluída", detail: "Os dados já estão atualizados neste dispositivo.", progress: 100, blocking: syncActivity.classList.contains("is-blocking") });
    }
    clearTimeout(syncActivityHideTimer);
    syncActivityHideTimer = setTimeout(function () {
      manualSyncActivity = false;
      syncActivity.hidden = true;
      syncActivity.classList.remove("is-blocking");
      if (app) app.setAttribute("aria-busy", "false");
    }, success ? 520 : 120);
  }

  function syncDisplayInfo(info) {
    info = info || { state: "disabled", configured: false, pending: 0, conflicts: 0, lastError: "", localVersion: 0, remoteVersion: 0, outdated: false };
    var localVersion = Number(info.localVersion) || 0;
    var remoteVersion = Number(info.remoteVersion) || 0;
    var versionCopy = remoteVersion ? "Versão Git v" + remoteVersion + (localVersion && localVersion !== remoteVersion ? " · dispositivo v" + localVersion : "") : "A aguardar a primeira versão Git.";
    return {
      title: !info.configured ? "Por configurar" : info.outdated ? "A atualizar em segundo plano…" : info.state === "checking" ? "A verificar a versão…" : info.state === "syncing" ? "A sincronizar…" : info.state === "synced" ? "Sincronizado" : info.state === "offline" ? "Sem Internet" : info.state === "error" ? "Erro de sincronização" : "Pronto",
      detail: !info.configured ? "Liga a app ao Worker que cria os commits no teu repositório privado." : info.pending ? info.pending + " alteração(ões) à espera de push." : info.lastError || versionCopy,
      badgeClass: info.state === "error" || info.state === "offline" ? "badge-yellow" : info.outdated ? "badge-yellow" : info.state === "synced" ? "badge-mint" : "badge-violet"
    };
  }

  function updateGitSyncCard(info) {
    var card = document.getElementById("gitSyncCard");
    if (!card) return;
    info = info || (Sync ? Sync.getStatus() : null) || {};
    var display = syncDisplayInfo(info);
    var title = document.getElementById("gitSyncTitle");
    var detail = document.getElementById("gitSyncDetail");
    var summary = document.getElementById("gitSyncSummary");
    var badge = document.getElementById("gitSyncBadge");
    var inline = document.getElementById("gitSyncInlineProgress");
    if (title) title.textContent = display.title;
    if (detail) detail.textContent = display.detail;
    if (summary) {
      var localVersion = Number(info.localVersion) || 0;
      var remoteVersion = Number(info.remoteVersion) || 0;
      summary.textContent = info.pending ? info.pending + " por enviar" : remoteVersion ? (info.outdated || (localVersion && localVersion !== remoteVersion) ? "v" + localVersion + " → v" + remoteVersion : "Versão Git v" + remoteVersion) : "PC + telemóvel";
    }
    if (badge) {
      badge.className = "badge " + display.badgeClass;
      badge.textContent = info.conflicts ? info.conflicts + " conflito(s)" : info.outdated ? "Desatualizado" : info.remoteVersion ? "Atualizado" : "Protegido";
    }
    card.setAttribute("aria-busy", info.state === "syncing" ? "true" : "false");
    if (inline) {
      inline.classList.toggle("is-active", info.state === "syncing");
      inline.classList.toggle("is-indeterminate", info.state === "syncing");
    }
    card.querySelectorAll('[data-action="force-git-pull"], [data-action="force-git-push"]').forEach(function (button) {
      button.disabled = !info.configured || info.state === "syncing";
    });
  }

  function updateSyncActivityFromStatus(info) {
    updateGitSyncCard(info);
    if (manualSyncActivity) return;
    // A sincronização automática é silenciosa: só o cartão do Git muda de estado.
    clearTimeout(syncActivityHideTimer);
    if (syncActivity) {
      syncActivity.hidden = true;
      syncActivity.classList.remove("is-blocking");
    }
    if (app) app.setAttribute("aria-busy", "false");
  }

  function round(value, digits) {
    var p = Math.pow(10, digits == null ? 1 : digits);
    return Math.round((Number(value) || 0) * p) / p;
  }

  function todayISO(date) {
    var d = date || new Date();
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, "0");
    var day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  function nowMinutes(date) {
    var d = date || new Date();
    return d.getHours() * 60 + d.getMinutes();
  }

  function timeMinutes(value) {
    if (!value || value.indexOf(":") === -1) return 0;
    var parts = value.split(":");
    return Number(parts[0]) * 60 + Number(parts[1]);
  }

  function academicYearFor(date) {
    var d = date || new Date();
    var start = d.getMonth() >= 8 ? d.getFullYear() : d.getFullYear() - 1;
    return start + "/" + String(start + 1).slice(-2);
  }

  function formatDate(value, options) {
    if (!value) return "Sem data";
    var date = new Date(value + (String(value).length === 10 ? "T12:00:00" : ""));
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat("pt-PT", options || { day: "numeric", month: "short" }).format(date);
  }

  function formatLongDate(value) {
    if (!value) return "";
    var date = new Date(value + (String(value).length === 10 ? "T12:00:00" : ""));
    return new Intl.DateTimeFormat("pt-PT", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(date);
  }

  function relativeDate(value) {
    if (!value) return "Sem prazo";
    var today = new Date(todayISO() + "T12:00:00");
    var target = new Date(value + "T12:00:00");
    var diff = Math.round((target - today) / 86400000);
    if (diff === 0) return "Hoje";
    if (diff === 1) return "Amanhã";
    if (diff === -1) return "Ontem";
    if (diff > 1 && diff < 7) return "Daqui a " + diff + " dias";
    if (diff < -1) return "Há " + Math.abs(diff) + " dias";
    return formatDate(value);
  }

  function safeColor(value, fallback) {
    return /^#[0-9a-f]{6}$/i.test(value || "") ? value : (fallback || COLORS[0]);
  }

  function hashText(text) {
    var hash = 2166136261;
    for (var i = 0; i < text.length; i += 1) {
      hash ^= text.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16);
  }

  function defaultState() {
    return {
      schemaVersion: 5,
      meta: {
        revision: 0,
        updatedAt: "",
        externalFingerprint: "",
        externalCheckedAt: "",
        externalRevision: 0,
        source: "device"
      },
      profile: {
        name: "",
        institution: "",
        degree: "",
        targetGrade: 20,
        onboardingComplete: false,
        tutorialSeen: false
      },
      settings: {
        campusSimulation: true,
        jsonSync: true,
        reduceMotion: false,
        plannerView: "schedule",
        calendarView: "month",
        studyPlanDate: todayISO(),
        weeklyStudyHours: 16,
        studyDayStart: "09:00",
        studyDayEnd: "19:00",
        studySessionMinutes: 50,
        studyBreakMinutes: 10,
        studyLunchStart: "13:00",
        studyLunchMinutes: 60,
        aiModelMode: "auto",
        aiOutput: "all",
        aiQuestionCount: 10,
        aiDifficulty: "auto"
      },
      currentSemesterId: null,
      semesters: [],
      courses: [],
      schedule: [],
      assessments: [],
      events: [],
      tasks: [],
      lessons: [],
      materials: [],
      pastExams: [],
      questions: [],
      quizzes: [],
      grades: [],
      studyBlocks: [],
      weeklyReviews: [],
      aiProjects: []
    };
  }

  function normalizeState(input) {
    var base = defaultState();
    var source = input && typeof input === "object" ? input : {};
    base.schemaVersion = Math.max(5, Number(source.schemaVersion) || 0);
    base.meta = Object.assign(base.meta, source.meta || {});
    base.profile = Object.assign(base.profile, source.profile || {});
    base.settings = Object.assign(base.settings, source.settings || {});
    if (["day", "three", "week", "month"].indexOf(base.settings.calendarView) < 0) base.settings.calendarView = "month";
    if (["schedule", "calendar", "study-day"].indexOf(base.settings.plannerView) < 0) base.settings.plannerView = "schedule";
    if (!/^\d{4}-\d{2}-\d{2}$/.test(base.settings.studyPlanDate || "")) base.settings.studyPlanDate = todayISO();
    base.settings.weeklyStudyHours = clamp(base.settings.weeklyStudyHours || 16, 1, 80);
    base.settings.studySessionMinutes = clamp(base.settings.studySessionMinutes || 50, 20, 180);
    base.settings.studyBreakMinutes = clamp(base.settings.studyBreakMinutes || 10, 0, 60);
    base.settings.studyLunchMinutes = clamp(base.settings.studyLunchMinutes || 60, 0, 180);
    if (["auto", "fast", "quality"].indexOf(base.settings.aiModelMode) < 0) base.settings.aiModelMode = "auto";
    if (["all", "notes", "summary", "quiz", "flashcards"].indexOf(base.settings.aiOutput) < 0) base.settings.aiOutput = "all";
    base.settings.aiQuestionCount = clamp(base.settings.aiQuestionCount || 10, 5, 30);
    if (["auto", "easy", "medium", "hard"].indexOf(base.settings.aiDifficulty) < 0) base.settings.aiDifficulty = "auto";
    base.currentSemesterId = source.currentSemesterId || null;
    ENTITY_ARRAYS.forEach(function (key) {
      base[key] = asArray(source[key]).filter(function (item) { return item && typeof item === "object"; });
    });
    base.courses = base.courses.map(function (course, index) {
      var result = Object.assign({
        id: uid("course"), semesterId: base.currentSemesterId, name: "Cadeira", code: "", ects: 0,
        color: COLORS[index % COLORS.length], lessonTypes: ["T"], evaluation: { components: [], examReplacesTests: false, replacementPolicy: "if-higher" }
      }, course);
      result.color = safeColor(result.color, COLORS[index % COLORS.length]);
      result.lessonTypes = asArray(result.lessonTypes);
      result.evaluation = Object.assign({ components: [], examReplacesTests: false, replacementPolicy: "if-higher" }, result.evaluation || {});
      result.evaluation.components = asArray(result.evaluation.components).map(function (component) {
        return Object.assign({ id: uid("component"), label: "Componente", weight: 0, count: 1, kind: "other", replaceable: false, minimum: null, defenseEnabled: false, defenseType: "oral", defenseThreshold: null, maxWithoutDefense: null }, component);
      });
      return result;
    });
    base.semesters = base.semesters.map(function (semester) {
      return Object.assign({ id: uid("semester"), name: "Semestre", academicYear: academicYearFor(), startDate: "", endDate: "", archived: false }, semester);
    });
    base.lessons = base.lessons.map(function (lesson) {
      return Object.assign({ notes: "", aiNotes: [], mastered: false }, lesson, { aiNotes: asArray(lesson.aiNotes) });
    });
    base.materials = base.materials.map(function (material) {
      return Object.assign({ remoteFile: null, slides: [], slideCount: 0, uploadStatus: "ready" }, material, {
        slides: asArray(material.slides),
        remoteFile: material.remoteFile && typeof material.remoteFile === "object" ? material.remoteFile : null
      });
    });
    base.tasks = base.tasks.map(function (task) {
      if (task.type === "lesson-quiz" && /^Quiz beOnLine · /.test(task.title || "")) task.title = String(task.title).replace(/^Quiz beOnLine · /, "Quiz da aula · ");
      return task;
    });
    base.quizzes = base.quizzes.map(function (quiz) {
      if (/^Quiz beOnLine · /.test(quiz.title || "")) quiz.title = String(quiz.title).replace(/^Quiz beOnLine · /, "Quiz da aula · ");
      return quiz;
    });
    base.assessments = base.assessments.map(function (assessment) {
      return Object.assign({ requiresTestSheet: false, openBook: false, hasDefense: false, defenseType: "oral", defenseThreshold: null, maxWithoutDefense: null, replacementAssessmentIds: [], replacementPolicy: "if-higher" }, assessment, { replacementAssessmentIds: asArray(assessment.replacementAssessmentIds) });
    });
    base.pastExams = base.pastExams.map(function (exam) {
      return Object.assign({ id: uid("pastexam"), semesterId: base.currentSemesterId, courseId: null, title: "Teste anterior", academicYear: "", date: "", source: "", notes: "", createdAt: "" }, exam);
    });
    base.questions = base.questions.map(function (question) {
      return Object.assign({ pastExamId: null, number: "", images: [] }, question, { lessonIds: asArray(question.lessonIds), images: normalizeImageRefs(question.images) });
    });
    base.events = base.events.map(function (event) {
      return Object.assign({}, event, { images: normalizeImageRefs(event.images, "event") });
    });
    base.quizzes = base.quizzes.map(function (quiz) {
      quiz.questions = asArray(quiz.questions).map(function (question) { return Object.assign({}, question, { images: normalizeImageRefs(question.images) }); });
      return quiz;
    });
    base.studyBlocks = base.studyBlocks.map(function (block) {
      return Object.assign({ id: uid("studyblock"), semesterId: base.currentSemesterId, date: todayISO(), title: "Sessão de estudo", start: "09:00", end: "09:50", kind: "study", courseId: null, sourceType: "custom", sourceId: null, completed: false, notes: "" }, block);
    });
    base.grades = base.grades.map(function (grade) {
      return Object.assign({ defenseStatus: "not-applicable", defenseType: "", defenseFinalScore: null }, grade);
    });
    base.aiProjects = base.aiProjects.map(function (project) {
      return Object.assign({
        id: uid("aiproject"), semesterId: base.currentSemesterId, courseId: null, quizId: null,
        title: "Projeto de IA", fileName: "", fileSize: 0, slideCount: 0, slides: [], summary: "",
        notes: null, flashcards: [], quizQuestions: [], output: "all", difficulty: "auto",
        questionCount: 10, modelMode: "fast", modelId: "", createdAt: "", warning: ""
      }, project, {
        slides: asArray(project.slides), flashcards: asArray(project.flashcards), quizQuestions: asArray(project.quizQuestions)
      });
    });
    return base;
  }

  function mergeById(localItems, externalItems) {
    var map = new Map();
    asArray(localItems).forEach(function (item) { if (item && item.id) map.set(item.id, clone(item)); });
    asArray(externalItems).forEach(function (item) {
      if (!item || !item.id) return;
      var prior = map.get(item.id) || {};
      map.set(item.id, Object.assign(prior, clone(item)));
    });
    return Array.from(map.values());
  }

  function mergeExternal(local, external) {
    var merged = normalizeState(local);
    if (!external || typeof external !== "object") return merged;
    if (external.meta && external.meta.syncMode === "replace") return normalizeState(external);
    var isTemplate = !!(external.meta && external.meta.isTemplate) && ENTITY_ARRAYS.every(function (key) { return !asArray(external[key]).length; }) && !(external.profile && external.profile.onboardingComplete);
    if (!isTemplate) {
      if (external.profile) merged.profile = Object.assign(merged.profile, external.profile);
      if (external.settings) merged.settings = Object.assign(merged.settings, external.settings);
      if (external.currentSemesterId) merged.currentSemesterId = external.currentSemesterId;
      ENTITY_ARRAYS.forEach(function (key) {
        if (Array.isArray(external[key])) merged[key] = mergeById(merged[key], external[key]);
      });
    }
    return normalizeState(merged);
  }

  function touchState() {
    state.meta.revision = (Number(state.meta.revision) || 0) + 1;
    state.meta.updatedAt = new Date().toISOString();
    state.meta.source = "device";
  }

  function save(silent) {
    touchState();
    return DB.saveState(state).then(function () {
      if (!silent) toast("Alterações guardadas neste dispositivo.");
    }).catch(function (error) {
      console.error(error);
      toast("Não foi possível guardar os dados.", "error");
    });
  }

  function externalJSONUrl() {
    return "data/academic-data.json?check=" + Date.now();
  }

  function loadExternalJSON(options) {
    options = options || {};
    if (state && state.settings && state.settings.jsonSync === false && !options.force) return Promise.resolve(false);
    return fetch(externalJSONUrl(), { cache: "no-store" }).then(function (response) {
      if (!response.ok) throw new Error("JSON externo indisponível");
      return response.text();
    }).then(function (raw) {
      var fingerprint = hashText(raw);
      var external = JSON.parse(raw);
      if (!state) {
        state = normalizeState(external);
        state.meta.externalFingerprint = fingerprint;
        state.meta.externalCheckedAt = new Date().toISOString();
        return DB.saveState(state).then(function () { return true; });
      }
      if (state.meta.externalFingerprint === fingerprint) {
        state.meta.externalCheckedAt = new Date().toISOString();
        return DB.saveState(state).then(function () { return false; });
      }
      state = mergeExternal(state, external);
      state.meta.externalFingerprint = fingerprint;
      state.meta.externalCheckedAt = new Date().toISOString();
      state.meta.externalRevision = Number(external.meta && external.meta.revision) || 0;
      var templateOnly = !!(external.meta && external.meta.isTemplate) && ENTITY_ARRAYS.every(function (key) { return !asArray(external[key]).length; }) && !(external.profile && external.profile.onboardingComplete);
      return DB.saveState(state).then(function () {
        if (!options.silent && !templateOnly) toast("Alterações do academic-data.json aplicadas.");
        return !templateOnly;
      });
    }).catch(function (error) {
      if (!options.silent) toast("Não foi possível ler o academic-data.json. Mantiveram-se os dados locais.", "warning");
      return false;
    });
  }

  function cleanText(value) {
    return String(value == null ? "" : value).replace(/\s+/g, " ").trim();
  }

  function loadCachedCanteen() {
    try {
      var parsed = JSON.parse(localStorage.getItem(CANTEEN_CACHE_KEY) || "null");
      return parsed && Array.isArray(parsed.days) && parsed.days.length ? parsed : null;
    } catch (_) {
      return null;
    }
  }

  function cacheCanteen(data) {
    try { localStorage.setItem(CANTEEN_CACHE_KEY, JSON.stringify(data)); } catch (_) { /* cache opcional */ }
  }

  function menuDateISO(label) {
    var match = cleanText(label).match(/(\d{1,2})\s+de\s+([A-Za-zÀ-ÿ]+)\s+de\s+(\d{4})/i);
    if (!match) return "";
    var key = match[2].normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    var months = { janeiro: 1, fevereiro: 2, marco: 3, abril: 4, maio: 5, junho: 6, julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12 };
    var month = months[key];
    return month ? match[3] + "-" + String(month).padStart(2, "0") + "-" + String(Number(match[1])).padStart(2, "0") : "";
  }

  function extractDishAllergens(value) {
    var ids = [];
    var description = cleanText(value).replace(/([A-Za-zÀ-ÿ])((?:(?:1[0-4]|[1-9])(?:,(?:1[0-4]|[1-9]))*))(?=\s|$)/g, function (_, letter, list) {
      list.split(",").forEach(function (id) { if (ids.indexOf(id) === -1) ids.push(id); });
      return letter;
    });
    return { description: description, allergens: ids };
  }

  function normalizeCanteenData(payload) {
    if (payload && Array.isArray(payload.days)) {
      var normalized = clone(payload);
      normalized.pageUrl = normalized.pageUrl || CANTEEN_PAGE_URL;
      normalized.apiUrl = normalized.apiUrl || CANTEEN_API_URL;
      return normalized;
    }
    var sections = asArray(payload && payload.acf && payload.acf.seccao);
    var menuSection = sections.find(function (item) { return cleanText(item.titulo).toLowerCase() === "ementa"; });
    if (!menuSection || !menuSection.conteudo) throw new Error("A resposta oficial não contém uma ementa.");
    var doc = new DOMParser().parseFromString(menuSection.conteudo, "text/html");
    var allergenMap = {};
    doc.querySelectorAll(".lista-alergenios span").forEach(function (element) {
      var match = cleanText(element.textContent).match(/^\((\d+)\)\s*(.*?);?$/);
      if (match) allergenMap[match[1]] = match[2].replace(/;$/, "").trim();
    });
    var days = Array.from(doc.querySelectorAll(".day-slot")).map(function (slot) {
      var header = cleanText(slot.querySelector(".header") && slot.querySelector(".header").textContent);
      var meals = [];
      var currentMeal = null;
      Array.from(slot.children).forEach(function (child) {
        if (child.classList.contains("title")) {
          currentMeal = { name: cleanText(child.textContent), items: [] };
          meals.push(currentMeal);
        } else if (child.classList.contains("list") && currentMeal) {
          child.querySelectorAll(".row").forEach(function (row) {
            var columns = Array.from(row.querySelectorAll(".list-col"));
            if (columns.length < 2) return;
            var dish = extractDishAllergens(columns[1].textContent);
            var calorieColumn = columns[columns.length - 1];
            currentMeal.items.push({
              type: cleanText(columns[0].textContent),
              description: dish.description,
              allergens: dish.allergens,
              kcal: Number(cleanText(calorieColumn.textContent)) || null,
              calorieBand: calorieColumn.classList.contains("high") ? "high" : calorieColumn.classList.contains("medium") ? "medium" : calorieColumn.classList.contains("low") ? "low" : ""
            });
          });
        }
      });
      return {
        date: menuDateISO(header),
        label: header,
        meals: meals.filter(function (meal) {
          var name = String(meal.name || "").toLowerCase();
          return meal.items.length && (name.indexOf("almoço") >= 0 || name.indexOf("jantar") >= 0);
        })
      };
    }).filter(function (day) { return day.meals.length; });
    if (!days.length) throw new Error("A estrutura da ementa oficial mudou.");
    var hoursSection = sections.find(function (item) { return cleanText(item.titulo).toLowerCase() === "horário"; });
    var hours = [];
    if (hoursSection && hoursSection.conteudo) {
      var hoursDoc = new DOMParser().parseFromString(hoursSection.conteudo, "text/html");
      hours = Array.from(hoursDoc.body.children).map(function (element) { return cleanText(element.textContent); }).filter(function (line) {
        return line && line.toLowerCase().indexOf("snack") < 0;
      });
    }
    return {
      schemaVersion: 1,
      fetchedAt: new Date().toISOString(),
      pageUrl: payload.link || CANTEEN_PAGE_URL,
      apiUrl: CANTEEN_API_URL,
      days: days,
      hours: hours,
      allergens: allergenMap,
      allergenNotice: cleanText(doc.querySelector(".observacoes") && doc.querySelector(".observacoes").textContent)
    };
  }

  function sectionPlainText(sections, title) {
    var section = asArray(sections).find(function (item) {
      return cleanText(item && item.titulo).toLowerCase() === title.toLowerCase();
    });
    if (!section || !section.conteudo) return "";
    return cleanText(new DOMParser().parseFromString(section.conteudo, "text/html").body.textContent);
  }

  function normalizeCanteenInfo(payload) {
    if (!payload) return null;
    if (payload.socialMeal || payload.closures) return clone(payload);
    var sections = asArray(payload.acf && payload.acf.seccao);
    var pricingText = sectionPlainText(sections, "Preçário");
    var closureText = sectionPlainText(sections, "Períodos de Encerramento");
    var amountMatch = pricingText.match(/Refeição social[^:]*:\s*([0-9]+(?:[,.][0-9]{1,2})?)\s*€/i);
    var effectiveMatch = pricingText.match(/\(a partir do dia\s+([^)]+)\)/i);
    var includesMatch = pricingText.match(/A refeição completa é composta por\s+(.+?)(?:;|\.)/i);
    var summerMatch = closureText.match(/Férias de Verão:\s*(.*?Setembro\.)/i);
    var seasonalMatch = closureText.match(/As cantinas encerram ainda.*?universitárias\)\./i);
    var alternativesMatch = closureText.match(/Algumas cantinas.*?mesmas\./i);
    return {
      pageUrl: payload.link || CANTEEN_INFO_PAGE_URL,
      apiUrl: CANTEEN_INFO_API_URL,
      sourceModifiedAt: payload.modified || "",
      socialMeal: {
        amount: amountMatch ? amountMatch[1].replace(".", ",") + " €" : "",
        audience: "Alunos de licenciatura, mestrado e doutoramento",
        effectiveFrom: effectiveMatch ? cleanText(effectiveMatch[1]) : "",
        includes: includesMatch ? cleanText(includesMatch[1]) : ""
      },
      closures: {
        summer: summerMatch ? "Férias de verão: " + cleanText(summerMatch[1]) : "",
        seasonal: seasonalMatch ? cleanText(seasonalMatch[0]) : "",
        alternatives: alternativesMatch ? cleanText(alternativesMatch[0]) : ""
      }
    };
  }

  function setCanteenResult(data, status, error) {
    canteenMenu = data;
    canteenStatus = status;
    canteenError = error || "";
    canteenChecked = true;
    if (data) cacheCanteen(data);
    if (route.name === "canteen") render();
    return { data: data, status: status, error: error || "" };
  }

  function ensureCanteenMenu(force) {
    if (canteenLoadPromise) return canteenLoadPromise;
    if (canteenChecked && !force) return Promise.resolve({ data: canteenMenu, status: canteenStatus, error: canteenError });
    canteenStatus = "loading";
    canteenError = "";
    var fetchOfficialJson = function (url) {
      return fetch(url, { mode: "cors", cache: "no-store", headers: { Accept: "application/json" } }).then(function (response) {
        if (!response.ok) throw new Error("A SAS NOVA respondeu com o estado " + response.status + ".");
        return response.json();
      });
    };
    canteenLoadPromise = Promise.all([
      fetchOfficialJson(CANTEEN_API_URL),
      fetchOfficialJson(CANTEEN_INFO_API_URL).catch(function () { return null; })
    ]).then(function (payloads) {
      var data = normalizeCanteenData(payloads[0]);
      var info = normalizeCanteenInfo(payloads[1]);
      if (info) {
        data.info = info;
        data.infoSource = "official";
      } else if (canteenMenu && canteenMenu.info) {
        data.info = clone(canteenMenu.info);
        data.infoSource = "cache";
      } else {
        data.infoSource = "fallback";
      }
      data.fetchedAt = new Date().toISOString();
      data.source = "official";
      return setCanteenResult(data, "ready", "");
    }).catch(function (officialError) {
      if (canteenMenu) return setCanteenResult(canteenMenu, "stale", officialError.message);
      return fetch("data/canteen-menu.json", { cache: "no-store" }).then(function (response) {
        if (!response.ok) throw officialError;
        return response.json();
      }).then(function (payload) {
        var data = normalizeCanteenData(payload);
        data.source = "snapshot";
        return setCanteenResult(data, "stale", officialError.message);
      });
    }).catch(function (error) {
      return setCanteenResult(null, "error", error.message || "Não foi possível obter a ementa.");
    }).finally(function () {
      canteenLoadPromise = null;
    });
    return canteenLoadPromise;
  }

  function currentSemester() {
    return state.semesters.find(function (semester) { return semester.id === state.currentSemesterId; }) || null;
  }

  function semesterById(id) { return state.semesters.find(function (item) { return item.id === id; }) || null; }
  function courseById(id) { return state.courses.find(function (item) { return item.id === id; }) || null; }
  function lessonById(id) { return state.lessons.find(function (item) { return item.id === id; }) || null; }
  function scheduleById(id) { return state.schedule.find(function (item) { return item.id === id; }) || null; }
  function assessmentById(id) { return state.assessments.find(function (item) { return item.id === id; }) || null; }
  function pastExamById(id) { return state.pastExams.find(function (item) { return item.id === id; }) || null; }

  function localDate(value) {
    var date = new Date(String(value || "") + "T12:00:00");
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function scheduleMatchesDate(entry, dateValue) {
    var date = localDate(dateValue);
    return !!(entry && date && Number(entry.weekday) === date.getDay());
  }

  function lessonMatchesSchedule(lesson, entry) {
    if (!lesson || !entry) return false;
    if (lesson.scheduleId) return lesson.scheduleId === entry.id;
    return lesson.courseId === entry.courseId && scheduleMatchesDate(entry, lesson.date) && String(lesson.type || "") === String(entry.type || "") && (!lesson.start || lesson.start === entry.start);
  }

  function linkedLessonForSlot(entry, dateValue) {
    var date = typeof dateValue === "string" ? dateValue : todayISO(dateValue);
    return semesterItems("lessons").find(function (lesson) {
      return lesson.date === date && lessonMatchesSchedule(lesson, entry);
    }) || null;
  }

  function assessmentKind(type) {
    var value = String(type || "").toLowerCase();
    if (value.indexOf("exam") >= 0) return "exam";
    if (value.indexOf("teste") >= 0 || value.indexOf("test") >= 0) return "test";
    if (value.indexOf("projeto") >= 0 || value.indexOf("project") >= 0) return "project";
    if (value.indexOf("apresent") >= 0 || value.indexOf("oral") >= 0) return "presentation";
    if (value.indexOf("aula") >= 0 || value.indexOf("ficha") >= 0) return "class";
    return "other";
  }

  function suggestedComponentId(course, assessment) {
    if (!course || !assessment) return null;
    var components = asArray(course.evaluation && course.evaluation.components);
    if (assessment.componentId && components.some(function (component) { return component.id === assessment.componentId; })) return assessment.componentId;
    var kind = assessmentKind(assessment.type);
    var candidates = components.filter(function (component) { return component.kind === kind; });
    if (candidates.length === 1) return candidates[0].id;
    var search = (String(assessment.title || "") + " " + String(assessment.type || "")).toLowerCase();
    var labelMatches = candidates.filter(function (component) {
      var label = String(component.label || "").toLowerCase();
      return label && (search.indexOf(label) >= 0 || label.split(/\s+/).some(function (part) { return part.length > 3 && search.indexOf(part) >= 0; }));
    });
    return labelMatches.length === 1 ? labelMatches[0].id : null;
  }

  function componentOptionsForCourse(courseId, selectedId) {
    var course = courseById(courseId);
    var components = asArray(course && course.evaluation && course.evaluation.components);
    return components.map(function (component) {
      return '<option value="' + attr(component.id) + '" ' + (component.id === selectedId ? "selected" : "") + '>' + esc(component.label) + ' · ' + (Number(component.weight) || 0) + '%</option>';
    }).join("");
  }

  function activeCourses() {
    return state.courses.filter(function (course) { return course.semesterId === state.currentSemesterId; });
  }

  function semesterItems(key, semesterId) {
    var id = semesterId || state.currentSemesterId;
    return state[key].filter(function (item) { return item.semesterId === id; });
  }

  function refreshIcons(root) {
    if (window.lucide && window.lucide.createIcons) {
      try { window.lucide.createIcons({ root: root || document, attrs: { "stroke-width": 2 } }); } catch (_) { /* icon fallback */ }
    }
  }

  function toast(message, type) {
    var element = document.createElement("div");
    element.className = "toast" + (type ? " toast-" + type : "");
    element.innerHTML = '<i data-lucide="' + (type === "error" ? "circle-alert" : type === "warning" ? "triangle-alert" : "sparkles") + '"></i><span>' + esc(message) + "</span>";
    toastRegion.appendChild(element);
    refreshIcons(element);
    setTimeout(function () { element.remove(); }, 3800);
  }

  function initials(name) {
    var parts = String(name || "20").trim().split(/\s+/).filter(Boolean);
    return parts.slice(0, 2).map(function (part) { return part.charAt(0).toUpperCase(); }).join("") || "20";
  }

  function renderShell() {
    var semester = currentSemester();
    document.getElementById("avatarInitials").textContent = initials(state.profile.name);
    document.getElementById("semesterMini").innerHTML = semester
      ? "<small>Semestre ativo</small><strong>" + esc(semester.name) + "</strong><span>" + esc(semester.academicYear) + " · " + activeCourses().length + " cadeiras</span>"
      : "<small>Sem semestre</small><strong>Configura o próximo</strong>";
    document.documentElement.classList.toggle("reduce-motion", !!state.settings.reduceMotion);
    renderNav();
  }

  function navRouteName() {
    if (route.name === "course" || route.name === "lesson") return "courses";
    if (route.name === "settings") return "settings";
    return route.name;
  }

  function renderNav() {
    var active = navRouteName();
    document.querySelectorAll("[data-route]").forEach(function (button) {
      button.classList.toggle("is-active", button.getAttribute("data-route") === active);
    });
    document.querySelector(".side-settings").classList.toggle("is-active", active === "settings");
  }

  function setHeader(title, eyebrow) {
    document.getElementById("pageTitle").textContent = title;
    document.getElementById("eyebrow").textContent = eyebrow || "Twenty · Study OS";
    document.title = title + " · Twenty";
  }

  function setRoute(name, id, tab) {
    route = { name: name || "home", id: id || null, tab: tab || "overview" };
    render();
    history.replaceState(null, "", "#" + route.name + (route.id ? "/" + route.id : "") + (route.tab && route.tab !== "overview" ? "/" + route.tab : ""));
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function routeFromHash() {
    var parts = location.hash.replace(/^#/, "").split("/").filter(Boolean);
    if (!parts.length) return;
    route = { name: parts[0], id: parts[1] || null, tab: parts[2] || "overview" };
  }

  function render() {
    if (!state) return;
    revokeImageObjectUrls();
    if (canteenClockTimer) {
      clearTimeout(canteenClockTimer);
      canteenClockTimer = null;
    }
    renderShell();
    var html;
    if (route.name === "home") html = renderHome();
    else if (route.name === "courses") html = renderCourses();
    else if (route.name === "course") html = renderCourse(route.id, route.tab);
    else if (route.name === "lesson") html = renderLesson(route.id);
    else if (route.name === "planner") html = renderPlanner();
    else if (route.name === "study") html = renderStudy();
    else if (route.name === "grades") html = renderGrades();
    else if (route.name === "canteen") html = renderCanteen();
    else if (route.name === "settings") html = renderSettings();
    else { route.name = "home"; html = renderHome(); }
    view.innerHTML = '<div class="view-enter">' + html + "</div>";
    view.focus({ preventScroll: true });
    refreshIcons(document);
    hydrateLocalImages(view);
    if (route.name === "settings") {
      enhanceSettingsActions();
      updateStorageCount();
    }
    if (route.name === "canteen" && !canteenChecked && canteenStatus !== "loading") {
      setTimeout(function () { ensureCanteenMenu(false); }, 0);
    }
    if (route.name === "canteen") {
      canteenClockTimer = setTimeout(function () {
        if (route.name === "canteen") render();
      }, 60000 - (Date.now() % 60000) + 250);
    }
  }

  function getLiveLesson(date) {
    var now = date || new Date();
    var day = todayISO(now);
    var minutes = nowMinutes(now);
    var exact = semesterItems("lessons").find(function (lesson) {
      return lesson.date === day && timeMinutes(lesson.start) <= minutes && timeMinutes(lesson.end || lesson.start) >= minutes;
    });
    if (exact) return { type: "lesson", lesson: exact, course: courseById(exact.courseId), start: exact.start, end: exact.end, room: exact.room, title: exact.title };
    var schedule = semesterItems("schedule").find(function (entry) {
      return Number(entry.weekday) === now.getDay() && timeMinutes(entry.start) <= minutes && timeMinutes(entry.end) >= minutes;
    });
    if (!schedule) return null;
    var linked = linkedLessonForSlot(schedule, day);
    return { type: linked ? "lesson" : "schedule", lesson: linked || null, schedule: schedule, course: courseById(schedule.courseId), start: schedule.start, end: schedule.end, room: schedule.room, title: linked ? linked.title : "Aula em direto" };
  }

  function getNextClass(date, options) {
    options = options || {};
    var now = date || new Date();
    var currentMinutes = nowMinutes(now);
    var entries = semesterItems("schedule").filter(function (entry) {
      return !options.courseId || entry.courseId === options.courseId;
    });
    var candidates = [];
    for (var offset = 0; offset < 15; offset += 1) {
      var d = new Date(now);
      d.setDate(now.getDate() + offset);
      entries.forEach(function (entry) {
        if (Number(entry.weekday) !== d.getDay()) return;
        if (offset === 0 && timeMinutes(entry.start) <= currentMinutes) return;
        var dateValue = todayISO(d);
        var lesson = linkedLessonForSlot(entry, dateValue);
        if (options.unprepared && lesson) return;
        candidates.push({ entry: entry, lesson: lesson, date: d, dateISO: dateValue, distance: offset * 1440 + timeMinutes(entry.start) - (offset === 0 ? currentMinutes : 0) });
      });
    }
    candidates.sort(function (a, b) { return a.distance - b.distance; });
    if (!candidates.length) return null;
    var best = candidates[0];
    return { schedule: best.entry, lesson: best.lesson, course: courseById(best.entry.courseId), date: best.date, dateISO: best.dateISO };
  }

  function nextOccurrenceForSchedule(entry, date) {
    if (!entry) return null;
    var now = date || new Date();
    var occurrence = new Date(now);
    var add = (Number(entry.weekday) - occurrence.getDay() + 7) % 7;
    if (add === 0 && timeMinutes(entry.end || entry.start) < nowMinutes(now)) add = 7;
    occurrence.setDate(occurrence.getDate() + add);
    var dateISO = todayISO(occurrence);
    return { date: occurrence, dateISO: dateISO, lesson: linkedLessonForSlot(entry, dateISO) };
  }

  function simulatedPeople(courseId) {
    var seedText = (courseId || "twenty") + todayISO() + String(new Date().getHours());
    var seed = parseInt(hashText(seedText).slice(0, 6), 16) || 0;
    return 118 + (seed % 39);
  }

  function nullableNumber(value) {
    if (value === "" || value == null || Number.isNaN(Number(value))) return null;
    return Number(value);
  }

  function gradeDefenseRules(grade, component) {
    var assessment = assessmentById(grade.assessmentId);
    var assessmentDefense = !!(assessment && assessment.hasDefense);
    var componentDefense = !!(component && component.defenseEnabled);
    return {
      enabled: assessmentDefense || componentDefense,
      type: assessmentDefense ? assessment.defenseType : component && component.defenseType,
      threshold: nullableNumber(assessmentDefense ? assessment.defenseThreshold : component && component.defenseThreshold),
      maxWithoutDefense: nullableNumber(assessmentDefense ? assessment.maxWithoutDefense : component && component.maxWithoutDefense)
    };
  }

  function effectiveGrade(grade, component) {
    var original = clamp(grade.score, 0, 20);
    var rules = gradeDefenseRules(grade, component);
    var completed = grade.defenseStatus === "completed";
    var finalScore = completed ? nullableNumber(grade.defenseFinalScore) : null;
    var effective = finalScore == null ? original : clamp(finalScore, 0, 20);
    var capped = false;
    if (rules.enabled && !completed && rules.maxWithoutDefense != null && effective > rules.maxWithoutDefense) {
      effective = clamp(rules.maxWithoutDefense, 0, 20);
      capped = true;
    }
    var pending = rules.enabled && !completed && (rules.threshold == null || original >= rules.threshold);
    return { grade: grade, component: component, original: original, effective: effective, defensePending: pending, defenseCompleted: completed, capped: capped, replaced: false, replacedByAssessmentId: null };
  }

  function courseAverage(course, gradePool) {
    if (!course) return { value: null, knownWeight: 0, projected: false, components: [], requirementsMet: true, minimumFailures: [], defensePending: [] };
    var components = asArray(course.evaluation && course.evaluation.components);
    var grades = asArray(gradePool || state.grades).filter(function (grade) { return grade.courseId === course.id; });
    var gradeResults = grades.map(function (grade) {
      var component = components.find(function (item) { return item.id === grade.componentId; });
      return effectiveGrade(grade, component);
    });
    gradeResults.slice().forEach(function (sourceResult) {
      var sourceAssessment = assessmentById(sourceResult.grade.assessmentId);
      var targets = asArray(sourceAssessment && sourceAssessment.replacementAssessmentIds);
      targets.forEach(function (targetAssessmentId) {
        var targetAssessment = assessmentById(targetAssessmentId);
        if (!targetAssessment || targetAssessment.courseId !== course.id) return;
        var targetResult = gradeResults.find(function (item) { return item.grade.assessmentId === targetAssessmentId; });
        if (!targetResult) {
          var targetComponent = components.find(function (item) { return item.id === targetAssessment.componentId; });
          targetResult = effectiveGrade({ id: "replacement_" + sourceResult.grade.id + "_" + targetAssessmentId, assessmentId: targetAssessmentId, courseId: course.id, componentId: targetAssessment.componentId, score: sourceResult.effective, defenseStatus: "not-applicable", synthetic: true }, targetComponent);
          gradeResults.push(targetResult);
        }
        var policy = sourceAssessment.replacementPolicy || "if-higher";
        if (policy === "always" || sourceResult.effective > targetResult.effective) {
          targetResult.effective = sourceResult.effective;
          targetResult.replaced = true;
          targetResult.replacedByAssessmentId = sourceAssessment.id;
        }
      });
    });
    var results = components.map(function (component) {
      var componentGrades = gradeResults.filter(function (item) { return item.grade.componentId === component.id && !item.grade.synthetic; });
      var effectiveGrades = gradeResults.filter(function (item) { return item.grade.componentId === component.id; });
      var raw = componentGrades.length ? componentGrades.reduce(function (sum, item) { return sum + item.original; }, 0) / componentGrades.length : null;
      var effective = effectiveGrades.length ? effectiveGrades.reduce(function (sum, item) { return sum + item.effective; }, 0) / effectiveGrades.length : null;
      var expected = Math.max(1, Number(component.count) || 1);
      var minimum = nullableNumber(component.minimum);
      var minimumState = minimum == null || effective == null ? "not-applicable" : effectiveGrades.length < expected ? "pending" : effective >= minimum ? "met" : "failed";
      return { component: component, raw: raw, effective: effective, count: effectiveGrades.length, expectedCount: expected, replaced: effectiveGrades.some(function (item) { return item.replaced; }), minimum: minimum, minimumState: minimumState };
    });
    var hasExplicitReplacements = semesterItems("assessments", course.semesterId).some(function (assessment) { return assessment.courseId === course.id && asArray(assessment.replacementAssessmentIds).length; });
    var exams = results.filter(function (result) { return result.component.kind === "exam" && result.effective != null; });
    if (!hasExplicitReplacements && course.evaluation && course.evaluation.examReplacesTests && exams.length) {
      var examScore = Math.max.apply(null, exams.map(function (result) { return result.raw; }));
      results.forEach(function (result) {
        if (result.component.kind === "test" && result.raw != null) {
          var shouldReplace = course.evaluation.replacementPolicy === "always" || examScore > result.raw;
          if (shouldReplace) { result.effective = examScore; result.replaced = true; }
        }
      });
    }
    var weighted = 0;
    var knownWeight = 0;
    results.forEach(function (result) {
      if (result.effective == null) return;
      var weight = Number(result.component.weight) || 0;
      weighted += result.effective * weight;
      knownWeight += weight;
    });
    return {
      value: knownWeight ? weighted / knownWeight : null,
      knownWeight: knownWeight,
      projected: knownWeight < 100,
      components: results,
      requirementsMet: !results.some(function (result) { return result.minimumState === "failed"; }),
      minimumFailures: results.filter(function (result) { return result.minimumState === "failed"; }),
      defensePending: gradeResults.filter(function (result) { return result.defensePending; })
    };
  }

  function ectsAverage() {
    var total = 0;
    var ects = 0;
    activeCourses().forEach(function (course) {
      var avg = courseAverage(course).value;
      if (avg == null) return;
      var weight = Number(course.ects) || 0;
      total += avg * weight;
      ects += weight;
    });
    return { value: ects ? total / ects : null, ects: ects };
  }

  function overallProgress() {
    var courses = activeCourses();
    if (!courses.length) return 0;
    var total = courses.reduce(function (sum, course) {
      var lessons = state.lessons.filter(function (item) { return item.courseId === course.id; });
      var mastered = lessons.filter(function (item) { return item.mastered; }).length;
      return sum + (lessons.length ? mastered / lessons.length : 0);
    }, 0);
    return Math.round((total / courses.length) * 100);
  }

  function pastQuestionsForLesson(lessonId) {
    return state.questions.filter(function (question) {
      return asArray(question.lessonIds).indexOf(lessonId) >= 0;
    });
  }

  function quizQuestionFromPast(question) {
    var options = asArray(question.options).filter(Boolean);
    if (options.length >= 2 && question.answerIndex != null) {
      return {
        id: uid("quizq"),
        sourceQuestionId: question.id,
        sourceType: "past-test",
        mode: "multiple-choice",
        prompt: question.prompt,
        options: options,
        answerIndex: clamp(question.answerIndex, 0, options.length - 1),
        explanation: question.explanation || question.answer || "",
        images: clone(normalizeImageRefs(question.images))
      };
    }
    return {
      id: uid("quizq"),
      sourceQuestionId: question.id,
      sourceType: "past-test",
      mode: "self-check",
      prompt: question.prompt,
      answer: question.answer || "A resposta ainda não foi adicionada.",
      explanation: question.explanation || "",
      academicYear: question.academicYear || "",
      assessmentLabel: question.assessmentLabel || "Teste anterior",
      images: clone(normalizeImageRefs(question.images))
    };
  }

  function lessonHasEnded(lesson, date) {
    if (!lesson || !lesson.date) return false;
    var now = date || new Date();
    var today = todayISO(now);
    if (lesson.date < today) return true;
    if (lesson.date > today) return false;
    var end = lesson.end || lesson.start;
    return end ? timeMinutes(end) <= nowMinutes(now) : true;
  }

  function lessonIsBeOnline(lesson) {
    if (!lesson) return false;
    if (lesson.quizCompletedAt || lesson.beOnlineCompletedAt) return true;
    return state.quizzes.some(function (quiz) {
      return quiz.lessonId === lesson.id && !!quiz.lastCompletedAt;
    });
  }

  function beOnlineStatus() {
    var due = semesterItems("lessons").filter(function (lesson) { return lessonHasEnded(lesson); });
    var completed = due.filter(lessonIsBeOnline);
    var pending = due.filter(function (lesson) { return !lessonIsBeOnline(lesson); });
    return {
      due: due,
      completed: completed,
      pending: pending,
      progress: due.length ? Math.round(completed.length / due.length * 100) : 100,
      isOnline: pending.length === 0
    };
  }

  function ensureBeOnlineTasks() {
    if (!state.currentSemesterId) return false;
    var changed = false;
    semesterItems("lessons").filter(function (lesson) {
      return lessonHasEnded(lesson);
    }).forEach(function (lesson) {
      var task = state.tasks.find(function (item) { return item.type === "lesson-quiz" && item.lessonId === lesson.id; });
      if (lessonIsBeOnline(lesson)) {
        if (task && !task.done) { task.done = true; changed = true; }
        return;
      }
      if (!task) {
        var course = courseById(lesson.courseId);
        state.tasks.push({
          id: uid("task"), semesterId: lesson.semesterId, courseId: lesson.courseId, lessonId: lesson.id,
          title: "Quiz da aula · " + lesson.title, type: "lesson-quiz", dueDate: lesson.date,
          dueTime: lesson.end && lesson.end > "20:30" ? lesson.end : "20:30", priority: "high", done: false, autoGenerated: true,
          createdAt: new Date().toISOString(), courseName: course ? course.name : ""
        });
        changed = true;
      }
    });
    return changed;
  }

  function completeLessonBeOnline(lessonId) {
    var lesson = lessonById(lessonId);
    if (!lesson) return;
    lesson.quizCompletedAt = new Date().toISOString();
    state.tasks.forEach(function (task) {
      if (task.lessonId === lessonId && task.type === "lesson-quiz") task.done = true;
    });
  }

  function emptyState(icon, title, text, action, label) {
    return '<div class="empty-state"><span class="empty-icon"><i data-lucide="' + icon + '"></i></span><h3>' + esc(title) + "</h3><p>" + esc(text) + "</p>" + (action ? '<button class="button button-dark button-small" type="button" data-action="' + attr(action) + '"><i data-lucide="plus"></i>' + esc(label || "Adicionar") + "</button>" : "") + "</div>";
  }

  function renderHome() {
    setHeader("Hoje", "Resumo académico");
    var semester = currentSemester();
    if (!semester) {
      return '<div class="page-head"><div><h2>Novo semestre</h2><p>Configura as cadeiras, o horário e as avaliações do próximo período letivo.</p></div></div>' + emptyState("calendar-plus", "Sem semestre ativo", "Os semestres anteriores permanecem disponíveis no arquivo.", "new-semester", "Criar semestre");
    }
    var name = (state.profile.name || "estudante").split(/\s+/)[0];
    var live = getLiveLesson();
    var next = getNextClass();
    var progress = overallProgress();
    var ects = ectsAverage();
    var pendingTasks = semesterItems("tasks").filter(function (task) { return !task.done; }).sort(function (a, b) { return String(a.dueDate || "9999").localeCompare(String(b.dueDate || "9999")); });
    var upcoming = semesterItems("assessments").filter(function (item) { return !item.date || item.date >= todayISO(); }).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    var upcomingEvents = semesterItems("events").filter(function (item) { return !item.date || item.date >= todayISO(); }).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    var online = beOnlineStatus();
    var overdueTasks = pendingTasks.filter(function (task) { return task.dueDate && task.dueDate < todayISO(); });
    var todayEvents = upcomingEvents.filter(function (event) { return event.date === todayISO(); });
    var urgentAssessment = upcoming.find(function (item) {
      if (!item.date) return false;
      return Math.round((localDate(item.date) - localDate(todayISO())) / 86400000) <= 3;
    });
    var heroTitle = "Resumo de hoje";
    var heroCopy = "Consulta as próximas aulas, tarefas, avaliações e eventos.";
    if (live && live.course) {
      heroTitle = "Em curso: " + (live.lesson ? live.lesson.title : live.course.name);
      heroCopy = live.lesson ? "Slides, apontamentos, quiz e perguntas anteriores disponíveis." : "Este período ainda não tem uma aula preparada.";
    } else if (overdueTasks.length) {
      heroTitle = overdueTasks.length + " tarefa" + (overdueTasks.length === 1 ? " em atraso" : "s em atraso");
      heroCopy = "Revê os prazos e atualiza as prioridades.";
    } else if (online.pending.length) {
      heroTitle = online.pending.length + " aula" + (online.pending.length === 1 ? " por rever." : "s por rever.");
      heroCopy = "Conclui os quizzes pendentes e regista as dúvidas identificadas.";
    } else if (urgentAssessment) {
      heroTitle = urgentAssessment.title + " está a chegar.";
      heroCopy = relativeDate(urgentAssessment.date) + ". Abre a matéria definida, as aulas incluídas e as perguntas anteriores.";
    } else if (next && next.lesson && next.dateISO === todayISO()) {
      heroTitle = "Próxima aula preparada";
      heroCopy = next.lesson.title + " · " + next.schedule.start + "–" + next.schedule.end + ".";
    } else if (todayEvents.length) {
      heroTitle = todayEvents[0].title + " é hoje.";
      heroCopy = (todayEvents[0].time ? "Às " + todayEvents[0].time + ". " : "") + (todayEvents[0].location || "Consulta os detalhes no Calendário.");
    }

    var liveHtml;
    if (live && live.course) {
      var materials = live.lesson ? state.materials.filter(function (item) { return item.lessonId === live.lesson.id; }) : [];
      var questions = live.lesson ? state.questions.filter(function (item) { return asArray(item.lessonIds).indexOf(live.lesson.id) >= 0; }) : [];
      var quizzes = live.lesson ? state.quizzes.filter(function (item) { return item.lessonId === live.lesson.id; }) : [];
      liveHtml = '<article class="card live-card span-7"><div class="live-top"><span class="live-pill"><span class="live-dot"></span> Aula em direto</span>' + (state.settings.campusSimulation ? '<span class="simulated-label">' + simulatedPeople(live.course.id) + ' a acompanhar · simulação</span>' : "") + '</div><h3>' + esc(live.lesson ? live.lesson.title : live.course.name) + '</h3><p>' + esc(live.course.name) + ' · ' + esc(live.start || "") + '–' + esc(live.end || "") + (live.room ? " · " + esc(live.room) : "") + '</p><div class="live-meta"><span><i data-lucide="file-text"></i>' + materials.length + ' PDF</span><span><i data-lucide="circle-help"></i>' + questions.length + ' perguntas anteriores</span><span><i data-lucide="sparkles"></i>' + quizzes.length + ' quiz</span></div><div class="live-actions">' + (live.lesson ? '<button class="button button-yellow" type="button" data-route="lesson" data-id="' + attr(live.lesson.id) + '"><i data-lucide="play"></i>Abrir aula</button><button class="button" type="button" data-action="edit-lesson" data-id="' + attr(live.lesson.id) + '"><i data-lucide="pencil"></i>Editar aula</button>' : '<button class="button button-yellow" type="button" data-action="create-lesson-from-live" data-id="' + attr(live.schedule.id) + '"><i data-lucide="link"></i>Preparar esta aula</button>') + (live.lesson ? '<button class="button" type="button" data-action="quick-review" data-course="' + attr(live.course.id) + '" data-lesson="' + attr(live.lesson.id) + '"><i data-lucide="rotate-ccw"></i>Rever depois</button>' : '') + '</div></article>';
    } else if (next && next.course) {
      liveHtml = '<article class="card card-dark span-7"><div class="card-title-row"><div><p class="card-label">Próxima aula</p><h3>' + esc(next.lesson ? next.lesson.title : next.course.name) + '</h3><p class="card-subtitle">' + esc(next.course.name) + ' · ' + (next.dateISO === todayISO() ? "Hoje" : WEEKDAYS[next.date.getDay()]) + ' · ' + esc(next.schedule.start) + '–' + esc(next.schedule.end) + (next.schedule.room ? " · " + esc(next.schedule.room) : "") + '</p></div><span class="badge badge-yellow">' + esc(next.schedule.type || "Aula") + '</span></div><div class="live-meta"><span><i data-lucide="clock-3"></i>' + relativeDate(next.dateISO) + '</span><span><i data-lucide="map-pin"></i>' + esc(next.schedule.room || "Sala por definir") + '</span><span><i data-lucide="' + (next.lesson ? "check" : "file-plus-2") + '"></i>' + (next.lesson ? "Aula preparada" : "Por preparar") + '</span></div><div class="live-actions">' + (next.lesson ? '<button class="button button-yellow" type="button" data-route="lesson" data-id="' + attr(next.lesson.id) + '"><i data-lucide="arrow-right"></i>Abrir aula</button><button class="button" type="button" data-action="edit-lesson" data-id="' + attr(next.lesson.id) + '"><i data-lucide="pencil"></i>Editar aula</button>' : '<button class="button button-yellow" type="button" data-action="create-lesson" data-course="' + attr(next.course.id) + '" data-schedule="' + attr(next.schedule.id) + '" data-date="' + attr(next.dateISO) + '" data-start="' + attr(next.schedule.start) + '" data-end="' + attr(next.schedule.end) + '" data-room="' + attr(next.schedule.room || "") + '" data-type="' + attr(next.schedule.type || "T") + '"><i data-lucide="file-plus-2"></i>Preparar aula</button>') + '</div></article>';
    } else {
      liveHtml = '<article class="card card-dark span-7">' + emptyState("clock-3", "Nenhuma aula em curso", "Configura o horário para identificar automaticamente as aulas em curso.", "add-schedule", "Adicionar horário") + "</article>";
    }

    var taskHtml = pendingTasks.length ? pendingTasks.slice(0, 4).map(renderTaskRow).join("") : emptyState("check-check", "Sem tarefas pendentes", "Não existem tarefas por concluir.", "add-task", "Nova tarefa");
    var homeAgenda = upcoming.map(function (item) { return Object.assign({ agendaKind: "assessment" }, item); }).concat(upcomingEvents.map(function (item) { return Object.assign({ agendaKind: "event" }, item); })).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")) || String(a.time || "").localeCompare(String(b.time || "")); });
    var upcomingHtml = homeAgenda.length ? homeAgenda.slice(0, 4).map(function (item) {
      var course = courseById(item.courseId);
      var isEvent = item.agendaKind === "event";
      return '<div class="list-row"><span class="list-icon ' + (isEvent ? "pink" : "orange") + '"><i data-lucide="' + (isEvent ? "party-popper" : assessmentIcon(item.type)) + '"></i></span><span class="list-content"><strong>' + esc(item.title) + '</strong><small>' + esc(isEvent ? item.location || "Evento da faculdade" : course ? course.name : "Avaliação") + ' · ' + relativeDate(item.date) + (item.time ? " às " + esc(item.time) : "") + (!isEvent && item.requiresTestSheet ? ' · comprar folha de teste' : '') + '</small></span><span class="badge ' + (item.date === todayISO() ? "badge-danger" : isEvent ? "badge-pink" : "badge-yellow") + '">' + esc(isEvent ? "Evento" : item.type || "Avaliação") + "</span></div>";
    }).join("") : emptyState("calendar-check", "Sem eventos próximos", "Não existem avaliações ou eventos agendados.", "add-assessment", "Adicionar data");

    var onlineTitle = !online.due.length ? "Sem aulas concluídas" : online.isOnline ? "Revisão em dia" : online.pending.length + " aula" + (online.pending.length === 1 ? "" : "s") + " por rever";
    var onlineCopy = !online.due.length ? "O estado será atualizado após a primeira aula." : online.isOnline ? "Todos os quizzes das aulas concluídas estão atualizados." : "Conclui os quizzes pendentes e regista as dúvidas identificadas.";
    var onlineCard = '<article class="card card-yellow span-5 target-card beonline-card"><div class="target-copy"><p class="card-label befirst-wordmark">BEFIRST<sup>™</sup></p><h3>' + esc(onlineTitle) + '</h3><p>' + esc(onlineCopy) + '</p><div class="tiny-stats"><span><strong>' + online.completed.length + '/' + online.due.length + '</strong>aulas revistas</span><span><strong>' + progress + '%</strong>matéria dominada</span><span><strong>' + semesterItems("questions").length + '</strong>perguntas antigas</span></div>' + (online.pending.length ? '<button class="button button-dark button-small" style="margin-top:14px" type="button" data-action="beonline-next"><i data-lucide="play"></i>Rever próxima aula</button>' : '') + '</div><div class="progress-ring" style="--progress:' + online.progress + '%"><strong>' + online.progress + '%</strong></div></article>';
    return '<section class="card hero-card span-12"><div class="hero-copy"><p class="hello">Olá, ' + esc(name) + '.</p><h2>' + esc(heroTitle) + '</h2><p>' + esc(heroCopy) + '</p><div class="hero-actions"><button class="button button-dark" type="button" data-route="study"><i data-lucide="sparkles"></i>Estudar agora</button><button class="button" type="button" data-action="quick-add"><i data-lucide="plus"></i>Adicionar</button></div></div><div class="hero-art"><span class="hero-orbit one"></span><span class="hero-orbit two"></span><span class="hero-number">20</span><span class="floating-chip one"><i data-lucide="coffee"></i>Modo de estudo</span><span class="floating-chip two"><i data-lucide="book-open"></i>' + activeCourses().length + ' cadeiras</span></div></section><div class="bento-grid" style="margin-top:15px">' + liveHtml + onlineCard + '<article class="card card-pink span-3 metric-card"><div class="metric-top"><p class="card-label">Média ECTS</p><span class="metric-icon"><i data-lucide="chart-no-axes-column-increasing"></i></span></div><div><p class="metric-value">' + (ects.value == null ? "—" : round(ects.value, 1)) + '</p><p class="metric-caption">' + (ects.value == null ? "Adiciona notas para calcular" : "ponderada por " + ects.ects + " ECTS") + '</p></div></article><article class="card card-violet span-3 metric-card"><div class="metric-top"><p class="card-label">Tarefas</p><span class="metric-icon"><i data-lucide="list-checks"></i></span></div><div><p class="metric-value">' + pendingTasks.length + '</p><p class="metric-caption">' + overdueTasks.length + ' atrasadas · ' + pendingTasks.filter(function (task) { return task.dueDate === todayISO(); }).length + ' para hoje</p></div></article><article class="card card-orange span-3 metric-card"><div class="metric-top"><p class="card-label">Avaliações</p><span class="metric-icon"><i data-lucide="alarm-clock"></i></span></div><div><p class="metric-value">' + upcoming.length + '</p><p class="metric-caption">próximas neste semestre</p></div></article><article class="card card-mint span-3 metric-card"><div class="metric-top"><p class="card-label">Perguntas antigas</p><span class="metric-icon"><i data-lucide="message-circle-question"></i></span></div><div><p class="metric-value">' + semesterItems("questions").length + '</p><p class="metric-caption">ligadas à matéria</p></div></article><article class="card span-6"><div class="card-title-row"><div><h3>Tarefas prioritárias</h3></div><button class="button button-small" type="button" data-action="add-task"><i data-lucide="plus"></i>Tarefa</button></div><div class="list-stack">' + taskHtml + '</div></article><article class="card span-6"><div class="card-title-row"><div><h3>Próximas datas</h3></div><button class="button button-small" type="button" data-action="planner-mode" data-mode="calendar"><i data-lucide="arrow-right"></i>Ver calendário</button></div><div class="list-stack">' + upcomingHtml + "</div></article></div>";
  }

  function assessmentIcon(type) {
    var value = String(type || "").toLowerCase();
    if (value.indexOf("projeto") >= 0) return "folder-kanban";
    if (value.indexOf("exame") >= 0) return "graduation-cap";
    if (value.indexOf("apresent") >= 0) return "presentation";
    return "file-pen-line";
  }

  function taskIcon(type) {
    var value = String(type || "").toLowerCase();
    if (value === "project") return "folder-kanban";
    if (value === "review") return "rotate-ccw";
    if (value === "lesson-quiz") return "radio-tower";
    if (value === "reading") return "book-open-text";
    return "notebook-pen";
  }

  function renderTaskRow(task) {
    var course = courseById(task.courseId);
    var label = task.type === "lesson-quiz" ? "Quiz" : task.type === "review" ? "Revisão" : task.priority === "high" ? "Prioridade" : "Tarefa";
    var lessonAction = "";
    if (task.lessonId && task.type === "lesson-quiz" && !task.done) {
      lessonAction = '<button class="button button-dark button-small task-open-button" type="button" data-action="do-beonline-quiz" data-lesson="' + attr(task.lessonId) + '"><i data-lucide="play"></i>Fazer quiz</button>';
    } else if (task.lessonId) {
      lessonAction = '<button class="row-button task-open-button" type="button" data-route="lesson" data-id="' + attr(task.lessonId) + '" aria-label="Abrir aula"><i data-lucide="arrow-up-right"></i></button>';
    }
    return '<div class="list-row ' + (task.done ? "is-done" : "") + '"><button class="check-button ' + (task.done ? "is-done" : "") + '" type="button" data-action="toggle-task" data-id="' + attr(task.id) + '" aria-label="' + (task.done ? "Reabrir tarefa" : "Concluir tarefa") + '">' + (task.done ? '<i data-lucide="check"></i>' : "") + '</button><span class="list-icon ' + (task.type === "review" || task.type === "lesson-quiz" ? "yellow" : "") + '"><i data-lucide="' + taskIcon(task.type) + '"></i></span><span class="list-content"><strong>' + esc(task.title) + '</strong><small>' + esc(course ? course.name : "Pessoal") + ' · ' + relativeDate(task.dueDate) + (task.dueTime ? " às " + esc(task.dueTime) : "") + '</small></span><span class="badge ' + (task.type === "lesson-quiz" ? "badge-violet" : task.priority === "high" ? "badge-danger" : "") + '">' + esc(label) + '</span>' + lessonAction + "</div>";
  }

  function courseProgress(course) {
    var lessons = state.lessons.filter(function (lesson) { return lesson.courseId === course.id; });
    if (!lessons.length) return 0;
    return Math.round(lessons.filter(function (lesson) { return lesson.mastered; }).length / lessons.length * 100);
  }

  function renderCourses() {
    setHeader("Cadeiras", "Biblioteca académica");
    var courses = activeCourses();
    var archived = state.semesters.filter(function (semester) { return semester.archived; });
    var cards = courses.map(function (course) {
      var progress = courseProgress(course);
      var lessons = state.lessons.filter(function (item) { return item.courseId === course.id; }).length;
      var avg = courseAverage(course).value;
      return '<article class="card course-card" data-route="course" data-id="' + attr(course.id) + '" tabindex="0" role="button" aria-label="Abrir ' + attr(course.name) + '"><div class="course-cover" style="--course-color:' + safeColor(course.color) + '"><span class="course-code">' + esc(course.code || "Cadeira") + '</span><h3>' + esc(course.name) + '</h3></div><div class="course-body"><div class="course-meta"><span>' + (Number(course.ects) || 0) + ' ECTS</span><span>' + lessons + ' aulas</span><span>' + (avg == null ? "Sem nota" : round(avg, 1) + "/20") + '</span></div><div class="mini-progress"><span style="width:' + progress + '%"></span></div><div class="course-footer"><span>' + progress + '% da matéria dominada</span><span class="arrow"><i data-lucide="arrow-up-right"></i></span></div></div></article>';
    }).join("");
    var archiveHtml = archived.length ? '<section class="section-block"><div class="section-heading"><div><h3>Arquivo</h3><p>Semestres anteriores continuam consultáveis.</p></div></div><div class="list-stack">' + archived.map(function (semester) {
      var count = state.courses.filter(function (course) { return course.semesterId === semester.id; }).length;
      return '<div class="list-row"><span class="list-icon"><i data-lucide="archive"></i></span><span class="list-content"><strong>' + esc(semester.name) + '</strong><small>' + esc(semester.academicYear) + ' · ' + count + ' cadeiras · apenas consulta</small></span><button class="button button-small" type="button" data-action="view-archive" data-id="' + attr(semester.id) + '">Ver</button></div>';
    }).join("") + "</div></section>" : "";
    return '<div class="page-head"><div><h2>Cadeiras do semestre</h2><p>Cada cadeira reúne aulas, slides, perguntas de testes anteriores, quizzes, avaliações e notas.</p></div><div class="page-actions"><button class="button" type="button" data-action="import-courses"><i data-lucide="file-json-2"></i>Importar JSON</button><button class="button" type="button" data-action="add-course"><i data-lucide="plus"></i>Nova cadeira</button><button class="button button-dark" type="button" data-action="new-semester"><i data-lucide="calendar-plus"></i>Novo semestre</button></div></div>' + (courses.length ? '<div class="course-grid">' + cards + "</div>" : emptyState("library-big", "Ainda não há cadeiras", "Adiciona a primeira cadeira ou importa a configuração do semestre.", "import-courses", "Importar cadeiras")) + archiveHtml;
  }

  function courseTabs(course, active) {
    var tabs = [
      ["overview", "Resumo"], ["lessons", "Aulas"], ["materials", "Materiais"], ["assessments", "Avaliações"],
      ["questions", "Perguntas anteriores"], ["quizzes", "Quizzes"], ["grades", "Notas"]
    ];
    return '<div class="tabbar" role="tablist">' + tabs.map(function (tab) {
      return '<button type="button" class="' + (active === tab[0] ? "is-active" : "") + '" data-action="course-tab" data-id="' + attr(course.id) + '" data-tab="' + tab[0] + '">' + tab[1] + "</button>";
    }).join("") + "</div>";
  }

  function renderCourse(id, tab) {
    var course = courseById(id);
    if (!course) {
      setHeader("Cadeira", "Não encontrada");
      return emptyState("circle-alert", "Cadeira não encontrada", "Pode ter sido removida ou pertencer a outro ficheiro JSON.", "go-courses", "Voltar às cadeiras");
    }
    var semester = semesterById(course.semesterId);
    var archived = !!(semester && semester.archived);
    var average = courseAverage(course);
    setHeader(course.code || "Cadeira", semester ? semester.name + " · " + semester.academicYear : "Cadeira");
    var hero = '<section class="card course-hero" style="--course-color:' + safeColor(course.color) + '"><div class="course-hero-copy"><span class="badge badge-dark">' + esc(course.code || "Cadeira") + '</span><h2>' + esc(course.name) + '</h2><p>' + (Number(course.ects) || 0) + ' ECTS · ' + asArray(course.lessonTypes).map(lessonTypeLabel).join(" · ") + (archived ? " · Semestre arquivado" : "") + '</p></div><div class="course-score"><strong>' + (average.value == null ? "—" : round(average.value, 1)) + '</strong><span>' + (average.value == null ? "sem notas" : "média atual / 20") + "</span></div></section>";
    var controls = courseTabs(course, tab || "overview");
    var content = renderCourseTab(course, tab || "overview", archived);
    return '<div class="page-head"><div><button class="button button-ghost button-small" type="button" data-route="courses"><i data-lucide="arrow-left"></i>Cadeiras</button></div><div class="page-actions">' + (!archived ? '<button class="button" type="button" data-action="edit-course" data-id="' + attr(course.id) + '"><i data-lucide="settings-2"></i>Configurar</button><button class="button button-dark" type="button" data-action="create-lesson" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Nova aula</button>' : '<span class="badge badge-dark"><i data-lucide="archive"></i>Arquivo</span>') + "</div></div>" + hero + controls + content;
  }

  function lessonTypeLabel(type) {
    var map = { T: "Teóricas", TP: "Teórico-práticas", P: "Práticas", LAB: "Laboratórios", OT: "Orientação" };
    return map[type] || type || "Aulas";
  }

  function renderCourseTab(course, tab, archived) {
    if (tab === "lessons") return renderCourseLessons(course, archived);
    if (tab === "materials") return renderCourseMaterials(course, archived);
    if (tab === "assessments") return renderCourseAssessments(course, archived);
    if (tab === "questions") return renderCourseQuestions(course, archived);
    if (tab === "quizzes") return renderCourseQuizzes(course, archived);
    if (tab === "grades") return renderCourseGrades(course, archived);
    return renderCourseOverview(course, archived);
  }

  function evaluationFormula(course) {
    var components = asArray(course.evaluation && course.evaluation.components);
    if (!components.length) return "Método de avaliação ainda não configurado.";
    var formula = components.map(function (item) {
      var count = Math.max(1, Number(item.count) || 1);
      return '<span class="formula-part"><b>' + count + '×</b> ' + esc(item.label) + ' <strong>' + (Number(item.weight) || 0) + '%</strong></span>';
    }).join('<span class="formula-plus">+</span>');
    if (course.evaluation.examReplacesTests) formula += "<br>Exame pode substituir a nota dos testes" + (course.evaluation.replacementPolicy === "always" ? "." : " quando for superior.");
    return formula;
  }

  function defenseTypeLabel(value) {
    if (value === "practical") return "prática";
    if (value === "oral-practical") return "oral e prática";
    return "oral";
  }

  function componentRuleText(component) {
    var rules = [];
    if (component.minimum != null) rules.push("mínimo " + round(component.minimum, 1) + "/20");
    if (component.defenseEnabled) {
      var defense = "defesa " + defenseTypeLabel(component.defenseType);
      if (component.defenseThreshold != null) defense += " a partir de " + round(component.defenseThreshold, 1);
      rules.push(defense);
      if (component.maxWithoutDefense != null) rules.push("máximo sem defesa " + round(component.maxWithoutDefense, 1));
    }
    return rules;
  }

  function renderCourseOverview(course, archived) {
    var lessons = state.lessons.filter(function (item) { return item.courseId === course.id; }).sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
    var assessments = state.assessments.filter(function (item) { return item.courseId === course.id && (!item.date || item.date >= todayISO()); }).sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
    var avg = courseAverage(course);
    var evalItems = avg.components.length ? avg.components.map(function (result) {
      var rules = componentRuleText(result.component);
      return '<div class="evaluation-item ' + (result.minimumState === "failed" ? "has-failed-minimum" : "") + '"><div><strong>' + esc(result.component.label) + '</strong><b>' + (result.effective == null ? "—" : round(result.effective, 1)) + '</b></div><small>' + (Number(result.component.weight) || 0) + '% · ' + result.count + '/' + result.expectedCount + ' nota(s)' + (result.replaced ? " · substituição aplicada" : "") + (rules.length ? " · " + esc(rules.join(" · ")) : "") + "</small></div>";
    }).join("") : '<div class="form-note">Configura as percentagens para calcular automaticamente a nota.</div>';
    var latestLessons = lessons.length ? lessons.slice(0, 4).map(renderLessonRow).join("") : emptyState("presentation", "Sem aulas", "Cria a primeira aula e liga-lhe o PDF, quiz e perguntas antigas.", "create-lesson", "Criar aula");
    var nextAssessments = assessments.length ? assessments.slice(0, 3).map(function (item) {
      return '<div class="list-row"><span class="list-icon orange"><i data-lucide="' + assessmentIcon(item.type) + '"></i></span><span class="list-content"><strong>' + esc(item.title) + '</strong><small>' + relativeDate(item.date) + (item.time ? " · " + esc(item.time) : "") + ' · ' + asArray(item.lessonIds).length + ' aulas na matéria' + (item.requiresTestSheet ? ' · comprar folha de teste' : '') + '</small></span><span class="badge badge-yellow">' + (Number(item.weight) || 0) + "%</span></div>";
    }).join("") : emptyState("calendar-check", "Sem avaliações futuras", "Quando souberes uma data, adiciona-a aqui.", "add-assessment", "Adicionar avaliação");
    return '<div class="bento-grid"><article class="card span-5"><div class="card-title-row"><div><h3>Método de avaliação</h3></div>' + (!archived ? '<button class="row-button" type="button" data-action="edit-course" data-id="' + attr(course.id) + '" aria-label="Editar método"><i data-lucide="pencil"></i></button>' : "") + '</div><div class="formula" style="margin-top:15px">' + evaluationFormula(course) + '</div><div class="evaluation-grid">' + evalItems + '</div></article><article class="card card-yellow span-3 metric-card"><div class="metric-top"><p class="card-label">Média atual</p><span class="metric-icon"><i data-lucide="calculator"></i></span></div><div><p class="metric-value">' + (avg.value == null ? "—" : round(avg.value, 1)) + '</p><p class="metric-caption">' + (avg.knownWeight ? avg.knownWeight + "% já avaliado" : "Ainda sem notas") + '</p></div></article><article class="card card-violet span-4 target-card"><div class="target-copy"><p class="card-label">Domínio</p><h3>' + courseProgress(course) + '% da matéria</h3><p>Marca cada aula quando conseguires explicar os conceitos sem consultar os slides.</p><div class="tiny-stats"><span><strong>' + lessons.filter(function (item) { return item.mastered; }).length + '</strong>dominadas</span><span><strong>' + lessons.length + '</strong>aulas</span><span><strong>' + state.questions.filter(function (item) { return item.courseId === course.id; }).length + '</strong>perguntas</span></div></div></article><article class="card span-7"><div class="card-title-row"><div><h3>Últimas aulas</h3></div><button class="button button-small" type="button" data-action="course-tab" data-id="' + attr(course.id) + '" data-tab="lessons">Ver todas</button></div><div class="list-stack">' + latestLessons + '</div></article><article class="card span-5"><div class="card-title-row"><div><h3>Próximas avaliações</h3></div>' + (!archived ? '<button class="button button-small" type="button" data-action="add-assessment" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Adicionar</button>' : "") + '</div><div class="list-stack">' + nextAssessments + "</div></article></div>";
  }

  function renderLessonRow(lesson) {
    var materialCount = state.materials.filter(function (item) { return item.lessonId === lesson.id; }).length;
    var questionCount = state.questions.filter(function (item) { return asArray(item.lessonIds).indexOf(lesson.id) >= 0; }).length;
    return '<div class="list-row"><span class="list-icon ' + (lesson.mastered ? "mint" : "") + '"><i data-lucide="' + (lesson.mastered ? "badge-check" : "presentation") + '"></i></span><span class="list-content"><strong>' + esc(lesson.title) + '</strong><small>' + formatDate(lesson.date) + ' · ' + esc(lesson.type || "Aula") + ' · ' + materialCount + ' PDF · ' + questionCount + ' perguntas</small></span><button class="row-button" type="button" data-route="lesson" data-id="' + attr(lesson.id) + '" aria-label="Abrir aula"><i data-lucide="arrow-right"></i></button></div>';
  }

  function renderCourseLessons(course, archived) {
    var lessons = state.lessons.filter(function (item) { return item.courseId === course.id; }).sort(function (a, b) { return String(b.date).localeCompare(String(a.date)) || String(b.start).localeCompare(String(a.start)); });
    var groups = {};
    lessons.forEach(function (lesson) {
      var key = lesson.date ? formatDate(lesson.date, { month: "long", year: "numeric" }) : "Sem data";
      if (!groups[key]) groups[key] = [];
      groups[key].push(lesson);
    });
    var content = Object.keys(groups).map(function (key) {
      return '<section class="section-block"><div class="section-heading"><div><h3>' + esc(key) + '</h3><p>' + groups[key].length + ' aula(s)</p></div></div><div class="list-stack">' + groups[key].map(renderLessonRow).join("") + "</div></section>";
    }).join("");
    return '<div class="page-head"><div><h2>Todas as aulas</h2><p>Abre uma aula para ver os slides, quiz, matéria, perguntas anteriores e apontamentos no mesmo sítio.</p></div>' + (!archived ? '<div class="page-actions"><button class="button button-dark" type="button" data-action="create-lesson" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Nova aula</button></div>' : "") + '</div>' + (lessons.length ? content : emptyState("presentation", "Ainda não há aulas", "Cria uma aula associada a um período do horário.", "create-lesson", "Criar primeira aula"));
  }

  function materialYearBadge(material, course) {
    var semester = semesterById(course.semesterId);
    if (!material.academicYear || (semester && material.academicYear === semester.academicYear)) return "";
    return '<span class="badge badge-pink">' + esc(material.academicYear) + "</span>";
  }

  function isPptxMaterial(material) {
    return !!material && (/powerpoint|presentation/i.test(material.mimeType || "") || /\.pptx$/i.test(material.fileName || ""));
  }

  function lessonMaterialsWithSlides(lessonId) {
    return state.materials.filter(function (item) {
      return item.lessonId === lessonId && isPptxMaterial(item) && asArray(item.slides).length;
    });
  }

  function lessonHasAISource(lessonId) {
    return lessonMaterialsWithSlides(lessonId).length > 0;
  }

  function renderAINote(note) {
    if (!note) return "";
    var notes = note.notes || note;
    return '<article class="lesson-ai-note"><div class="lesson-ai-note-head"><div><span class="badge badge-violet"><i data-lucide="sparkles"></i>IA</span><h4>' + esc(notes.title || note.title || "Apontamentos gerados") + '</h4></div><small>' + formatDate((note.createdAt || "").slice(0, 10)) + '</small></div>' + (notes.overview ? '<p class="ai-overview">' + nl2br(notes.overview) + '</p>' : '') + asArray(notes.sections).map(function (section) { return '<details class="lesson-ai-section"><summary>' + esc(section.heading || "Tópico") + '</summary><p>' + nl2br(section.content || "") + '</p></details>'; }).join("") + (asArray(notes.keyTakeaways).length ? '<ul class="lesson-ai-takeaways">' + notes.keyTakeaways.map(function (item) { return '<li>' + esc(item) + '</li>'; }).join("") + '</ul>' : '') + '</article>';
  }

  function lessonAIAvailableLessons(courseId) {
    return state.lessons.filter(function (item) {
      return item.courseId === courseId && lessonHasAISource(item.id);
    }).sort(function (a, b) { return String(a.date || "").localeCompare(String(b.date || "")); });
  }

  function openLessonAIModal(lessonId, output, materialId) {
    var lesson = lessonById(lessonId);
    if (!lesson || !AI) return;
    var available = lessonAIAvailableLessons(lesson.courseId);
    if (!available.length) {
      toast("Carrega primeiro um PowerPoint numa aula desta cadeira.", "warning");
      return;
    }
    output = output === "notes" ? "notes" : "quiz";
    var selected = {};
    selected[lesson.id] = true;
    if (materialId) {
      var target = state.materials.find(function (item) { return item.id === materialId; });
      if (target && target.lessonId) selected[target.lessonId] = true;
    }
    var lessonRows = available.map(function (item) {
      var count = lessonMaterialsWithSlides(item.id).reduce(function (sum, material) { return sum + asArray(material.slides).length; }, 0);
      return '<label class="lesson-ai-choice"><input type="checkbox" name="lessonIds" value="' + attr(item.id) + '" ' + (selected[item.id] ? "checked" : "") + '><span><strong>' + esc(item.title) + '</strong><small>' + formatDate(item.date) + ' · ' + count + ' slides</small></span></label>';
    }).join("");
    var pastCount = state.questions.filter(function (question) { return asArray(question.lessonIds).some(function (id) { return available.some(function (lessonItem) { return lessonItem.id === id; }); }); }).length;
    var body = '<form id="lessonAIForm" data-lesson="' + attr(lesson.id) + '" data-output="' + output + '"><div class="lesson-ai-modal-intro"><span class="metric-icon"><i data-lucide="' + (output === "quiz" ? "brain" : "notebook-pen") + '"></i></span><div><h3>' + (output === "quiz" ? "Gerar quiz a partir das aulas" : "Gerar apontamentos a partir das aulas") + '</h3><p>Escolhe exatamente que aulas entram como fonte. A geração é guardada nesta aula e sincronizada no Git.</p></div></div><div class="field"><label>Aulas usadas pela IA</label><div class="lesson-ai-choices">' + lessonRows + '</div></div><label class="lesson-ai-toggle"><input type="checkbox" name="includePast" ' + (pastCount ? "checked" : "") + '><span><strong>Incluir perguntas de anos anteriores</strong><small>' + pastCount + ' pergunta(s) disponíveis nas aulas escolhidas. No quiz, também entram como perguntas reais.</small></span></label><div class="form-grid"><div class="field"><label>Modelo</label><select name="modelMode"><option value="auto">Automático · estável</option><option value="fast">Rápido · 0.5B</option><option value="quality">Qualidade · 1.5B</option></select></div>' + (output === "quiz" ? '<div class="field"><label>Número de perguntas IA</label><select name="questionCount"><option>5</option><option selected>10</option><option>15</option><option>20</option></select></div><div class="field"><label>Dificuldade</label><select name="difficulty"><option value="auto">Automática</option><option value="easy">Fácil</option><option value="medium">Média</option><option value="hard">Difícil</option></select></div>' : '') + '</div><div class="form-note"><strong>Primeira geração:</strong> o modelo pode demorar a descarregar. Não precisas de manter esta janela aberta depois de terminar.</div></form>';
    openModal(output === "quiz" ? "Criar quiz com IA" : "Criar apontamentos com IA", body, { className: "modal-wide lesson-ai-modal", footer: '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Cancelar</button><button class="button button-dark" type="button" data-action="run-lesson-ai"><i data-lucide="sparkles"></i>Gerar e sincronizar</button></footer>' });
  }

  function buildLessonAISource(lessonIds, includePast) {
    var slides = [];
    var sourceMap = {};
    var next = 1;
    lessonIds.forEach(function (lessonId) {
      var lesson = lessonById(lessonId);
      lessonMaterialsWithSlides(lessonId).forEach(function (material) {
        asArray(material.slides).forEach(function (slide) {
          sourceMap[next] = { lessonId: lessonId, materialId: material.id, originalSlide: slide.number, lessonTitle: lesson && lesson.title || "Aula" };
          slides.push({ number: next, title: (lesson ? lesson.title : "Aula") + " · slide " + slide.number + " · " + (slide.title || ""), text: slide.text || "" });
          next += 1;
        });
      });
    });
    var past = state.questions.filter(function (question) {
      return asArray(question.lessonIds).some(function (id) { return lessonIds.indexOf(id) >= 0; });
    });
    if (includePast) past.forEach(function (question) {
      sourceMap[next] = { questionId: question.id, pastQuestion: true };
      slides.push({ number: next, title: "Pergunta de " + (question.academicYear || "ano anterior"), text: "PERGUNTA: " + (question.prompt || "") + "\nRESPOSTA: " + (question.answer || "") + "\nEXPLICAÇÃO: " + (question.explanation || "") });
      next += 1;
    });
    return { fileName: "Aulas selecionadas", fileSize: 0, slideCount: slides.length, slides: slides, sourceMap: sourceMap, pastQuestions: past };
  }

  async function runLessonAI() {
    var form = document.getElementById("lessonAIForm");
    if (!form || aiBusy || !AI) return;
    var lesson = lessonById(form.dataset.lesson);
    if (!lesson) return;
    var selectedIds = Array.from(form.querySelectorAll('input[name="lessonIds"]:checked')).map(function (input) { return input.value; });
    if (!selectedIds.length) { toast("Escolhe pelo menos uma aula.", "warning"); return; }
    var output = form.dataset.output === "notes" ? "notes" : "quiz";
    var includePast = !!(form.elements.includePast && form.elements.includePast.checked);
    var source = buildLessonAISource(selectedIds, includePast);
    if (!source.slides.length) { toast("As aulas escolhidas não têm texto de slides disponível.", "warning"); return; }
    var options = {
      output: output,
      modelMode: form.elements.modelMode.value || "auto",
      questionCount: form.elements.questionCount ? Number(form.elements.questionCount.value) || 10 : 10,
      difficulty: form.elements.difficulty ? form.elements.difficulty.value || "auto" : "auto"
    };
    closeModal();
    aiBusy = true;
    setManualSyncActivity("A preparar a IA local…", "A organizar " + source.slideCount + " fontes das aulas escolhidas.", 4, true);
    try {
      var result = await AI.generateStudyPack(source, options, function (report) {
        setManualSyncActivity(report.kind === "model" ? "A preparar o modelo…" : output === "quiz" ? "A criar o quiz…" : "A criar os apontamentos…", report.text || "A processar a matéria.", report.progress, true);
      });
      if (output === "quiz") {
        var questions = asArray(result.quizQuestions);
        if (includePast) {
          var used = new Set(questions.map(function (item) { return String(item.prompt || "").trim().toLowerCase(); }));
          source.pastQuestions.forEach(function (question) {
            var key = String(question.prompt || "").trim().toLowerCase();
            if (!used.has(key)) { questions.push(quizQuestionFromPast(question)); used.add(key); }
          });
        }
        var quiz = { id: uid("quiz"), semesterId: lesson.semesterId, courseId: lesson.courseId, lessonId: lesson.id, lessonIds: selectedIds, title: "IA · " + lesson.title, questions: questions, generatedByAI: true, includesPastQuestions: includePast, createdAt: new Date().toISOString(), lastScore: null };
        state.quizzes.push(quiz);
      } else {
        lesson.aiNotes = asArray(lesson.aiNotes);
        lesson.aiNotes.unshift({ id: uid("ainote"), title: result.notes && result.notes.title || "Apontamentos · " + lesson.title, notes: result.notes, summary: result.summary || "", lessonIds: selectedIds, includesPastQuestions: includePast, modelId: result.modelId || "", createdAt: new Date().toISOString() });
      }
      setManualSyncActivity("A sincronizar no Git…", "A guardar o resultado para aparecer no PC e no telemóvel.", 96, true);
      await save(true);
      if (Sync && Sync.getStatus().configured) { try { await Sync.syncNow(state, defaultState()); } catch (_) {} }
      finishManualSyncActivity(true);
      aiBusy = false;
      render();
      toast(output === "quiz" ? "Quiz criado e sincronizado." : "Apontamentos criados e sincronizados.");
    } catch (error) {
      aiBusy = false;
      finishManualSyncActivity(false);
      toast(error.message || "A IA não conseguiu terminar.", "error");
    }
  }

  async function uploadMaterialFile(file, context) {
    context = context || {};
    if (!file || !file.size) return null;
    if (file.size > 25 * 1024 * 1024) throw new Error("O ficheiro tem mais de 25 MB.");
    if (!Sync || !Sync.getStatus().configured) throw new Error("Configura primeiro o Git em Admin & dados para sincronizar o ficheiro.");
    if (!navigator.onLine) throw new Error("Precisas de Internet para enviar o ficheiro. A aula ainda não foi alterada.");
    var materialId = context.id || uid("material");
    var blobId = await DB.putFile(file, { courseId: context.courseId, lessonId: context.lessonId });
    var extracted = null;
    if (/\.pptx$/i.test(file.name || "") && AI) {
      setManualSyncActivity("A extrair os slides…", "A preparar o PowerPoint para a IA.", 4, true);
      extracted = await AI.extractPptx(file, function (report) {
        setManualSyncActivity("A extrair os slides…", report.text || "A ler o PowerPoint.", Math.min(30, Number(report.progress) || 4), true);
      });
    }
    setManualSyncActivity("A enviar o material…", "A iniciar o upload para o repositório privado.", 32, true);
    var remoteFile = await Sync.uploadFile(file, {
      id: materialId,
      name: file.name,
      onProgress: function (report) {
        var progress = report.progress == null ? null : 32 + Math.round(report.progress * 0.58);
        var detail = report.total ? formatBytes(report.loaded) + " de " + formatBytes(report.total) + " enviados" : "A enviar o ficheiro…";
        setManualSyncActivity("A enviar o material…", detail, progress, true);
      },
      onUploadComplete: function () { setManualSyncActivity("A confirmar no GitHub…", "A aguardar a confirmação do commit do ficheiro.", 92, true); },
      onReady: function (request) { aiTransferRequest = request; }
    });
    aiTransferRequest = null;
    return { id: materialId, blobId: blobId, remoteFile: remoteFile, slides: extracted ? extracted.slides : [], slideCount: extracted ? extracted.slideCount : 0 };
  }

  async function syncExistingMaterial(id) {
    var material = state.materials.find(function (item) { return item.id === id; });
    if (!material) return;
    var record = material.blobId ? await DB.getFile(material.blobId) : null;
    if (!record || !record.blob) {
      toast("Abre a Twenty no dispositivo onde carregaste este ficheiro e sincroniza-o aí, ou volta a carregá-lo.", "warning");
      return;
    }
    try {
      var uploaded = await uploadMaterialFile(record.blob, { id: material.id, courseId: material.courseId, lessonId: material.lessonId });
      material.source = "remote";
      material.remoteFile = uploaded.remoteFile;
      material.slides = uploaded.slides;
      material.slideCount = uploaded.slideCount;
      material.fileName = material.fileName || record.name;
      material.mimeType = material.mimeType || record.type;
      await save(true);
      if (Sync && Sync.getStatus().configured) { try { await Sync.syncNow(state, defaultState()); } catch (_) {} }
      finishManualSyncActivity(true);
      render();
      toast("Material sincronizado com o Git.");
    } catch (error) {
      finishManualSyncActivity(false);
      toast(error.message || "Não foi possível sincronizar o material.", "error");
    }
  }

  function renderMaterialCard(material, course, archived) {
    var lesson = lessonById(material.lessonId);
    var kind = material.kind || "slides";
    var icon = kind === "slides" ? "presentation" : kind === "notes" ? "notebook-pen" : "file-text";
    var synced = !!(material.remoteFile && material.remoteFile.path);
    var aiReady = isPptxMaterial(material) && asArray(material.slides).length && lesson;
    return '<article class="material-card"><div class="material-preview"><i data-lucide="' + icon + '"></i>' + materialYearBadge(material, course) + '</div><h4>' + esc(material.title) + '</h4><p>' + esc(lesson ? lesson.title : "Biblioteca da cadeira") + (material.fileName ? " · " + esc(material.fileName) : "") + '</p><div class="material-actions"><span class="badge ' + (synced ? 'badge-mint' : 'badge-yellow') + '">' + (synced ? 'Sincronizado' : 'Local') + '</span><span class="list-actions">' + (aiReady && !archived ? '<button class="row-button row-button-ai" type="button" data-action="lesson-ai" data-output="quiz" data-lesson="' + attr(lesson.id) + '" data-material="' + attr(material.id) + '" aria-label="Criar quiz com IA"><i data-lucide="sparkles"></i></button>' : '') + (!synced && material.blobId && !archived ? '<button class="row-button" type="button" data-action="sync-material" data-id="' + attr(material.id) + '" aria-label="Sincronizar ficheiro com o Git"><i data-lucide="cloud-upload"></i></button>' : '') + '<button class="row-button" type="button" data-action="open-material" data-id="' + attr(material.id) + '" aria-label="Abrir material"><i data-lucide="eye"></i></button>' + (!archived ? '<button class="row-button" type="button" data-action="delete-entity" data-kind="materials" data-id="' + attr(material.id) + '" aria-label="Remover material"><i data-lucide="trash-2"></i></button>' : "") + '</span></div>' + (aiReady ? '<small class="material-ai-ready"><i data-lucide="brain"></i>' + asArray(material.slides).length + ' slides prontos para IA</small>' : '') + '</article>';
  }

  function renderCourseMaterials(course, archived) {
    var materials = state.materials.filter(function (item) { return item.courseId === course.id; }).sort(function (a, b) { return String(b.academicYear).localeCompare(String(a.academicYear)); });
    var semester = semesterById(course.semesterId);
    var current = materials.filter(function (item) { return !semester || !item.academicYear || item.academicYear === semester.academicYear; });
    var older = materials.filter(function (item) { return semester && item.academicYear && item.academicYear !== semester.academicYear; });
    var currentHtml = current.length ? '<div class="material-grid">' + current.map(function (item) { return renderMaterialCard(item, course, archived); }).join("") + "</div>" : emptyState("file-up", "Sem materiais deste ano", "Podes carregar o PDF depois da aula. Até lá, os slides antigos ficam disponíveis em baixo.", "add-material", "Carregar PDF");
    var olderHtml = older.length ? '<section class="section-block"><div class="section-heading"><div><h3>Anos letivos anteriores</h3><p>Cada ficheiro mantém o ano visível para não se confundir com a matéria atual.</p></div></div><div class="material-grid">' + older.map(function (item) { return renderMaterialCard(item, course, archived); }).join("") + "</div></section>" : "";
    return '<div class="page-head"><div><h2>Slides e PDFs</h2><p>Os documentos do ano atual aparecem sem etiqueta; materiais antigos mostram sempre o respetivo ano letivo.</p></div>' + (!archived ? '<div class="page-actions"><button class="button button-dark" type="button" data-action="add-material" data-course="' + attr(course.id) + '"><i data-lucide="file-up"></i>Carregar material</button></div>' : "") + '</div><section class="section-block"><div class="section-heading"><div><h3>' + esc(semester ? semester.academicYear : "Ano atual") + '</h3><p>Materiais principais desta cadeira</p></div></div>' + currentHtml + "</section>" + olderHtml;
  }

  function assessmentRuleBadges(item) {
    var badges = "";
    if (item.requiresTestSheet) badges += '<span class="badge badge-danger"><i data-lucide="shopping-basket"></i>Comprar folha de teste</span>';
    if (item.openBook) badges += '<span class="badge badge-mint"><i data-lucide="book-open-check"></i>Consulta</span>';
    if (item.hasDefense) badges += '<span class="badge badge-violet"><i data-lucide="messages-square"></i>Defesa ' + esc(defenseTypeLabel(item.defenseType)) + '</span>';
    if (asArray(item.replacementAssessmentIds).length) badges += '<span class="badge badge-mint"><i data-lucide="replace"></i>Substitui ' + asArray(item.replacementAssessmentIds).length + '</span>';
    return badges;
  }

  function assessmentRuleSummary(item) {
    var notes = [];
    if (item.defenseThreshold != null) notes.push("defesa a partir de " + round(item.defenseThreshold, 1) + "/20");
    if (item.maxWithoutDefense != null) notes.push("máximo sem defesa: " + round(item.maxWithoutDefense, 1) + "/20");
    if (asArray(item.replacementAssessmentIds).length) notes.push((item.replacementPolicy === "always" ? "substituição obrigatória" : "substitui apenas se melhorar"));
    return notes.length ? '<div class="assessment-rule-summary"><i data-lucide="info"></i><span>' + esc(notes.join(" · ")) + '</span></div>' : "";
  }

  function renderCourseAssessments(course, archived) {
    var items = state.assessments.filter(function (item) { return item.courseId === course.id; }).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    var html = items.map(function (item) {
      var lessonNames = asArray(item.lessonIds).map(function (id) { var lesson = lessonById(id); return lesson ? lesson.title : null; }).filter(Boolean);
      var component = asArray(course.evaluation && course.evaluation.components).find(function (entry) { return entry.id === item.componentId; });
      var actions = !archived ? '<div class="list-actions"><button class="row-button" type="button" data-action="add-grade" data-assessment="' + attr(item.id) + '" aria-label="Adicionar nota"><i data-lucide="chart-no-axes-combined"></i></button><button class="row-button" type="button" data-action="edit-assessment" data-id="' + attr(item.id) + '" aria-label="Editar avaliação"><i data-lucide="pencil"></i></button><button class="row-button" type="button" data-action="delete-entity" data-kind="assessments" data-id="' + attr(item.id) + '" aria-label="Remover avaliação"><i data-lucide="trash-2"></i></button></div>' : "";
      return '<article class="card span-6"><div class="card-title-row"><div><div class="question-meta"><span class="badge badge-yellow">' + esc(item.type || "Avaliação") + '</span>' + (component ? '<span class="badge badge-violet">' + esc(component.label) + '</span>' : '') + '<span class="badge">' + (Number(item.weight) || 0) + '%</span>' + assessmentRuleBadges(item) + '</div><h3 style="margin-top:12px">' + esc(item.title) + '</h3><p class="card-subtitle">' + formatLongDate(item.date) + (item.time ? " · " + esc(item.time) : "") + '</p></div>' + actions + '</div><div class="form-note" style="margin-top:15px"><strong>Matéria:</strong> ' + (lessonNames.length ? esc(lessonNames.join(" · ")) : "Ainda não foram selecionadas aulas.") + "</div>" + assessmentRuleSummary(item) + "</article>";
    }).join("");
    return '<div class="page-head"><div><h2>Avaliações e matéria</h2><p>Define exatamente que aulas teóricas, práticas ou teórico-práticas entram em cada teste.</p></div>' + (!archived ? '<div class="page-actions"><button class="button button-dark" type="button" data-action="add-assessment" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Nova avaliação</button></div>' : "") + '</div><div class="bento-grid">' + (items.length ? html : '<div class="span-12">' + emptyState("file-pen-line", "Sem avaliações", "Adiciona testes, projetos, apresentações ou exames e escolhe a matéria.", "add-assessment", "Adicionar avaliação") + "</div>") + "</div>";
  }

  function renderQuestionCard(question, archived) {
    var lessons = asArray(question.lessonIds).map(function (id) { var lesson = lessonById(id); return lesson ? lesson.title : null; }).filter(Boolean);
    var exam = pastExamById(question.pastExamId);
    return '<article class="question-card"><div class="question-meta"><span class="badge badge-pink">Pergunta de teste anterior</span>' + (question.number ? '<span class="badge badge-dark">' + esc(question.number) + '</span>' : '') + (question.academicYear ? '<span class="badge">' + esc(question.academicYear) + '</span>' : "") + (exam ? '<span class="badge badge-violet">' + esc(exam.title) + '</span>' : question.assessmentLabel ? '<span class="badge badge-violet">' + esc(question.assessmentLabel) + '</span>' : "") + '</div><h4>' + esc(question.prompt) + '</h4>' + renderImageGallery(question.images, "question", { compact: true, ownerId: question.id }) + '<p>' + (lessons.length ? "Associada a: " + esc(lessons.join(" · ")) : "Ainda sem aula associada") + '</p><div class="list-actions" style="margin-top:11px"><button class="button button-small" type="button" data-action="show-question-answer" data-id="' + attr(question.id) + '"><i data-lucide="eye"></i>Ver resposta</button>' + (!archived ? '<button class="row-button" type="button" data-action="edit-question" data-id="' + attr(question.id) + '" aria-label="Editar pergunta"><i data-lucide="pencil"></i></button><button class="row-button" type="button" data-action="delete-entity" data-kind="questions" data-id="' + attr(question.id) + '" aria-label="Remover pergunta"><i data-lucide="trash-2"></i></button>' : "") + "</div></article>";
  }

  function renderCourseQuestions(course, archived) {
    var questions = state.questions.filter(function (item) { return item.courseId === course.id; });
    var exams = state.pastExams.filter(function (item) { return item.courseId === course.id; }).sort(function (a, b) { return String(b.academicYear || b.date).localeCompare(String(a.academicYear || a.date)); });
    var examCards = exams.map(function (exam) {
      var count = questions.filter(function (question) { return question.pastExamId === exam.id; }).length;
      return '<article class="past-exam-card"><div><span class="badge badge-violet">' + esc(exam.academicYear || "Ano por indicar") + '</span><h3>' + esc(exam.title) + '</h3><p>' + count + ' pergunta(s)' + (exam.date ? ' · ' + esc(formatDate(exam.date)) : '') + (exam.source ? ' · ' + esc(exam.source) : '') + '</p></div><div class="list-actions"><button class="button button-small" type="button" data-action="add-question" data-course="' + attr(course.id) + '" data-past-exam="' + attr(exam.id) + '"><i data-lucide="plus"></i>Pergunta</button><button class="row-button" type="button" data-action="delete-entity" data-kind="pastExams" data-id="' + attr(exam.id) + '" aria-label="Remover teste anterior"><i data-lucide="trash-2"></i></button></div></article>';
    }).join("");
    var groupedQuestions = exams.map(function (exam) {
      var items = questions.filter(function (question) { return question.pastExamId === exam.id; });
      return items.length ? '<section class="section-block"><div class="section-heading"><div><h3>' + esc(exam.title) + '</h3><p>' + esc(exam.academicYear || "Ano letivo por indicar") + ' · ' + items.length + ' pergunta(s)</p></div></div><div class="list-stack">' + items.map(function (item) { return renderQuestionCard(item, archived); }).join("") + '</div></section>' : '';
    }).join("");
    var loose = questions.filter(function (question) { return !question.pastExamId || !pastExamById(question.pastExamId); });
    if (loose.length) groupedQuestions += '<section class="section-block"><div class="section-heading"><div><h3>Perguntas sem teste associado</h3><p>' + loose.length + ' pergunta(s)</p></div></div><div class="list-stack">' + loose.map(function (item) { return renderQuestionCard(item, archived); }).join("") + '</div></section>';
    return '<div class="page-head"><div><h2>Perguntas de testes anteriores</h2><p>Importa um teste completo ou adiciona perguntas individuais e liga-as às aulas relevantes.</p></div>' + (!archived ? '<div class="page-actions"><button class="button" type="button" data-action="add-question" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Pergunta</button><button class="button button-dark" type="button" data-action="add-past-exam" data-course="' + attr(course.id) + '"><i data-lucide="file-json-2"></i>Importar teste</button></div>' : "") + '</div>' + (exams.length ? '<div class="past-exam-grid">' + examCards + '</div>' : '') + (questions.length ? groupedQuestions : emptyState("message-circle-question", "Banco de perguntas vazio", "Adiciona uma pergunta ou importa um teste anterior em JSON.", "add-past-exam", "Importar teste anterior"));
  }

  function renderCourseQuizzes(course, archived) {
    var quizzes = state.quizzes.filter(function (item) { return item.courseId === course.id; });
    var html = quizzes.map(function (quiz) {
      var lesson = lessonById(quiz.lessonId);
      var hasPast = lesson && pastQuestionsForLesson(lesson.id).length;
      var origin = quiz.generatedFromPastQuestions ? "Perguntas anteriores" : "Manual";
      return '<article class="card span-4"><div class="card-title-row"><div><div class="question-meta"><span class="badge badge-violet">' + asArray(quiz.questions).length + ' pergunta(s)</span><span class="badge">' + origin + '</span></div><h3 style="margin-top:11px">' + esc(quiz.title) + '</h3><p class="card-subtitle">' + esc(lesson ? lesson.title : "Quiz geral da cadeira") + '</p></div><span class="metric-icon"><i data-lucide="sparkles"></i></span></div><div class="live-actions" style="margin-top:21px"><button class="button button-dark" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="play"></i>Começar</button>' + (!archived ? '<button class="button" type="button" data-action="add-quiz-question" data-id="' + attr(quiz.id) + '"><i data-lucide="plus"></i>Manual</button>' + (hasPast ? '<button class="button" type="button" data-action="add-past-to-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="history"></i>Anteriores</button>' : '') + '<button class="button button-danger" type="button" data-action="delete-entity" data-kind="quizzes" data-id="' + attr(quiz.id) + '"><i data-lucide="trash-2"></i></button>' : "") + "</div></article>";
    }).join("");
    return '<div class="page-head"><div><h2>Quizzes da cadeira</h2><p>Cria perguntas manualmente ou reutiliza as perguntas reais de testes anteriores ligadas a cada aula.</p></div>' + (!archived ? '<div class="page-actions"><button class="button button-dark" type="button" data-action="add-quiz" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Novo quiz</button></div>' : "") + '</div><div class="bento-grid">' + (quizzes.length ? html : '<div class="span-12">' + emptyState("sparkles", "Ainda não há quizzes", "Cria um quiz normal ou gera-o a partir do banco de perguntas anteriores de uma aula.", "add-quiz", "Criar quiz") + "</div>") + "</div>";
  }

  function renderCourseGrades(course, archived) {
    var avg = courseAverage(course);
    var grades = state.grades.filter(function (item) { return item.courseId === course.id; }).sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
    var components = asArray(course.evaluation && course.evaluation.components);
    var rows = grades.map(function (grade) {
      var component = components.find(function (item) { return item.id === grade.componentId; });
      var assessment = assessmentById(grade.assessmentId);
      var lesson = lessonById(grade.lessonId);
      var sourceTitle = assessment ? assessment.title : lesson ? lesson.title : component ? component.label : "Registo antigo";
      var sourceMeta = assessment ? (assessment.type || "Avaliação") : lesson ? "Nota de aula" : "Sem origem concreta";
      var gradeResult = effectiveGrade(grade, component);
      var states = [];
      if (gradeResult.defensePending) states.push("defesa pendente");
      if (gradeResult.defenseCompleted) states.push("defesa concluída");
      if (gradeResult.capped) states.push("limite sem defesa aplicado");
      var scoreCopy = '<span class="grade-number">' + round(grade.score, 1) + '</span>/20';
      if (Math.abs(gradeResult.effective - gradeResult.original) > .001) scoreCopy += '<small class="effective-grade">efetiva: ' + round(gradeResult.effective, 1) + '/20</small>';
      return '<tr><td><strong>' + esc(sourceTitle) + '</strong><small class="table-subtitle">' + esc(sourceMeta + (component ? " · " + component.label : "") + (states.length ? " · " + states.join(" · ") : "")) + '</small></td><td>' + formatDate(grade.date) + '</td><td>' + esc(grade.notes || "—") + '</td><td>' + scoreCopy + '</td><td>' + (!archived ? '<button class="row-button" type="button" data-action="delete-entity" data-kind="grades" data-id="' + attr(grade.id) + '"><i data-lucide="trash-2"></i></button>' : "") + "</td></tr>";
    }).join("");
    var componentsHtml = avg.components.map(function (result) {
      var minimumCopy = result.minimum == null ? "" : result.minimumState === "failed" ? " · mínimo não atingido" : result.minimumState === "met" ? " · mínimo atingido" : " · mínimo por confirmar";
      return '<div class="evaluation-item ' + (result.minimumState === "failed" ? "has-failed-minimum" : "") + '"><div><strong>' + esc(result.component.label) + '</strong><b>' + (result.effective == null ? "—" : round(result.effective, 1)) + '/20</b></div><small>' + (Number(result.component.weight) || 0) + '% · ' + result.count + '/' + result.expectedCount + ' nota(s)' + (result.replaced ? " · substituição aplicada" : "") + minimumCopy + "</small></div>";
    }).join("");
    var alerts = "";
    if (avg.minimumFailures.length) alerts += '<div class="grade-rule-alert is-danger"><i data-lucide="shield-alert"></i><span><strong>Mínimo não atingido.</strong> ' + esc(avg.minimumFailures.map(function (result) { return result.component.label + " exige " + result.minimum + "/20"; }).join(" · ")) + '</span></div>';
    if (avg.defensePending.length) alerts += '<div class="grade-rule-alert"><i data-lucide="messages-square"></i><span><strong>Defesa pendente.</strong> ' + avg.defensePending.length + ' nota(s) precisam de confirmação.</span></div>';
    return '<div class="page-head"><div><h2>Notas e cálculo</h2><p>Cada nota fica ligada à avaliação de origem e respeita mínimos, substituições e regras de defesa configuradas.</p></div>' + (!archived ? '<div class="page-actions"><button class="button button-dark" type="button" data-action="add-grade" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Adicionar nota</button></div>' : "") + '</div>' + alerts + '<div class="bento-grid"><article class="card card-yellow span-4 metric-card"><div class="metric-top"><p class="card-label">Média atual</p><span class="metric-icon"><i data-lucide="calculator"></i></span></div><div><p class="metric-value">' + (avg.value == null ? "—" : round(avg.value, 1)) + '</p><p class="metric-caption">' + avg.knownWeight + '% da avaliação com nota</p></div></article><article class="card span-8"><p class="card-label">Componentes</p><div class="evaluation-grid">' + (componentsHtml || '<div class="form-note">Configura primeiro o método de avaliação.</div>') + '</div></article><article class="card span-12"><div class="card-title-row"><div><h3>Todas as notas</h3></div></div>' + (grades.length ? '<div style="overflow:auto"><table class="grade-table"><thead><tr><th>Avaliação ou aula</th><th>Data</th><th>Notas</th><th>Valor</th><th></th></tr></thead><tbody>' + rows + "</tbody></table></div>" : emptyState("chart-no-axes-combined", "Sem notas registadas", "Adiciona notas para calcular a média da cadeira e a média ECTS.", "add-grade", "Adicionar nota")) + "</article></div>";
  }

  function renderLesson(id) {
    var lesson = lessonById(id);
    if (!lesson) {
      setHeader("Aula", "Não encontrada");
      return emptyState("circle-alert", "Aula não encontrada", "Pode ter sido removida ou alterada.", "go-courses", "Voltar");
    }
    var course = courseById(lesson.courseId);
    var semester = course ? semesterById(course.semesterId) : null;
    var archived = !!(semester && semester.archived);
    setHeader(lesson.title, course ? course.name : "Aula");
    var materials = state.materials.filter(function (item) { return item.lessonId === lesson.id; });
    var questions = pastQuestionsForLesson(lesson.id);
    var quizzes = state.quizzes.filter(function (item) { return item.lessonId === lesson.id; });
    var aiSources = lessonAIAvailableLessons(lesson.courseId);
    var canGenerateAI = aiSources.length > 0;
    var onlineComplete = lessonIsBeOnline(lesson);
    var lessonEnded = lessonHasEnded(lesson);
    var currentMaterials = materials.filter(function (item) { return !semester || !item.academicYear || item.academicYear === semester.academicYear; });
    var oldMaterials = materials.filter(function (item) { return semester && item.academicYear && item.academicYear !== semester.academicYear; });
    var materialsHtml = materials.length ? '<div class="material-grid">' + currentMaterials.concat(oldMaterials).map(function (item) { return renderMaterialCard(item, course || { semesterId: null }, archived); }).join("") + '</div>' : emptyState("file-up", "Ainda sem material", "Carrega os slides ou PDF desta aula. O ficheiro ficará sincronizado.", "add-material", "Carregar material");
    var questionsHtml = questions.length ? questions.map(function (item) { return renderQuestionCard(item, archived); }).join("") : emptyState("message-circle-question", "Sem perguntas anteriores", "Associa perguntas de testes antigos a esta aula.", "add-question", "Associar pergunta");
    var quizHtml = quizzes.length ? quizzes.map(function (quiz) {
      var origin = quiz.generatedByAI ? "IA" : quiz.generatedFromPastQuestions ? "perguntas anteriores" : "manual";
      return '<div class="list-row quiz-list-row"><span class="list-icon"><i data-lucide="sparkles"></i></span><span class="list-content"><strong>' + esc(quiz.title) + '</strong><small>' + asArray(quiz.questions).length + ' pergunta(s) · ' + origin + (quiz.lastScore != null ? ' · ' + quiz.lastScore + '%' : '') + '</small></span><div class="list-actions"><button class="button button-dark button-small" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="play"></i>Começar</button></div></div>';
    }).join("") : emptyState("sparkles", "Quiz por criar", "Gera um quiz com os slides de uma ou várias aulas.", canGenerateAI ? "lesson-ai" : "add-quiz", canGenerateAI ? "Gerar com IA" : "Criar quiz");
    var generatedNotes = asArray(lesson.aiNotes);
    var generatedNotesHtml = generatedNotes.length ? '<div class="lesson-ai-notes">' + generatedNotes.map(renderAINote).join("") + '</div>' : '';
    var statusLabel = onlineComplete ? "Revista" : lessonEnded ? "Por rever" : "Preparada";
    var statusCopy = onlineComplete ? "Quiz concluído. Esta aula continua disponível para praticares." : lessonEnded ? "Faz um quiz curto para não deixares a matéria acumular." : "Quando a aula terminar, faz o quiz para confirmares o que percebeste.";
    var aiQuizButton = !archived && canGenerateAI ? '<button class="button button-violet button-small" type="button" data-action="lesson-ai" data-output="quiz" data-lesson="' + attr(lesson.id) + '"><i data-lucide="sparkles"></i>Gerar com IA</button>' : '';
    var aiNotesButton = !archived && canGenerateAI ? '<button class="button button-violet button-small" type="button" data-action="lesson-ai" data-output="notes" data-lesson="' + attr(lesson.id) + '"><i data-lucide="sparkles"></i>Gerar com IA</button>' : '';
    return '<div class="page-head"><div><button class="button button-ghost button-small" type="button" data-route="course" data-id="' + attr(lesson.courseId) + '"><i data-lucide="arrow-left"></i>' + esc(course ? course.code || course.name : "Cadeira") + '</button><h2 style="margin-top:11px">' + esc(lesson.title) + '</h2><p>' + formatLongDate(lesson.date) + (lesson.start ? ' · ' + esc(lesson.start) + '–' + esc(lesson.end || '') : '') + ' · ' + esc(lessonTypeLabel(lesson.type)) + (lesson.room ? ' · ' + esc(lesson.room) : '') + '</p></div><div class="page-actions">' + (!archived ? '<button class="button" type="button" data-action="edit-lesson" data-id="' + attr(lesson.id) + '"><i data-lucide="pencil"></i>Editar aula</button><button class="button ' + (lesson.mastered ? 'button-yellow' : 'button-dark') + '" type="button" data-action="toggle-mastery" data-id="' + attr(lesson.id) + '"><i data-lucide="badge-check"></i>' + (lesson.mastered ? 'Dominada' : 'Marcar dominada') + '</button>' : '') + '</div></div><div class="bento-grid"><article class="card course-hero span-12" style="--course-color:' + safeColor(course && course.color) + ';min-height:220px"><div class="course-hero-copy"><span class="badge badge-dark">' + esc(lesson.type || 'Aula') + '</span><h2>' + esc(lesson.title) + '</h2><p>' + esc(lesson.topics || 'Adiciona os tópicos dados nesta aula.') + '</p></div><div class="course-score"><strong>' + (lesson.mastered ? '✓' : questions.length) + '</strong><span>' + (lesson.mastered ? 'matéria dominada' : 'perguntas antigas') + '</span></div></article><article class="card span-12 beonline-lesson-card ' + (onlineComplete ? 'is-online' : '') + '"><div class="beonline-lesson-copy"><span class="badge ' + (onlineComplete ? 'badge-mint' : lessonEnded ? 'badge-danger' : 'badge-violet') + '">' + statusLabel + '</span><h3>Revisão da aula</h3><p>' + esc(statusCopy) + '</p></div></article><article class="card span-12"><div class="card-title-row"><div><h3>Slides e PDFs</h3><p class="card-subtitle">Os ficheiros são enviados para o Git e ficam disponíveis nos teus dispositivos.</p></div>' + (!archived ? '<button class="button button-small" type="button" data-action="add-material" data-course="' + attr(lesson.courseId) + '" data-lesson="' + attr(lesson.id) + '"><i data-lucide="file-up"></i>Carregar</button>' : '') + '</div><div style="margin-top:15px">' + materialsHtml + '</div></article><article class="card span-7"><div class="card-title-row"><div><h3>Perguntas de testes anteriores</h3></div><div class="list-actions">' + (!archived ? '<button class="button button-small" type="button" data-action="add-question" data-course="' + attr(lesson.courseId) + '" data-lesson="' + attr(lesson.id) + '"><i data-lucide="plus"></i>Pergunta</button>' : '') + '</div></div><div style="margin-top:15px">' + questionsHtml + '</div></article><article class="card span-5"><div class="card-title-row"><div><h3>Quiz da aula</h3></div><div class="list-actions">' + aiQuizButton + (!archived ? '<button class="button button-small" type="button" data-action="add-quiz" data-course="' + attr(lesson.courseId) + '" data-lesson="' + attr(lesson.id) + '"><i data-lucide="plus"></i>Manual</button>' : '') + '</div></div><div class="list-stack">' + quizHtml + '</div></article><article class="card span-12"><div class="card-title-row"><div><h3>Apontamentos</h3><p class="card-subtitle">Podes combinar slides de várias aulas e incluir perguntas antigas.</p></div><div class="list-actions">' + aiNotesButton + (!archived ? '<button class="button button-small" type="button" data-action="edit-lesson-notes" data-id="' + attr(lesson.id) + '"><i data-lucide="pencil"></i>Editar manual</button>' : '') + '</div></div>' + generatedNotesHtml + '<div class="form-note" style="margin-top:15px">' + (lesson.notes ? nl2br(lesson.notes) : 'Ainda não escreveste apontamentos manuais nesta aula.') + '</div></article></div>';
  }

  function plannerModeControl(active) {
    return '<div class="planner-mode-control" role="group" aria-label="Vista da agenda"><button type="button" class="' + (active === "schedule" ? "is-active" : "") + '" data-action="planner-mode" data-mode="schedule"><i data-lucide="clock-3"></i>Horário</button><button type="button" class="' + (active === "calendar" ? "is-active" : "") + '" data-action="planner-mode" data-mode="calendar"><i data-lucide="calendar-days"></i>Calendário</button><button type="button" class="' + (active === "study-day" ? "is-active" : "") + '" data-action="planner-mode" data-mode="study-day"><i data-lucide="blocks"></i>Dia de estudo</button></div>';
  }

  function plannerSupportingCards() {
    var tasks = semesterItems("tasks").slice().sort(function (a, b) { return Number(a.done) - Number(b.done) || String(a.dueDate || "9999").localeCompare(String(b.dueDate || "9999")); });
    var assessments = semesterItems("assessments").slice().sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    var events = semesterItems("events").slice().sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    var taskHtml = tasks.length ? tasks.map(renderTaskRow).join("") : emptyState("list-checks", "Sem tarefas", "Adiciona trabalhos de casa, projetos, leituras ou revisões.", "add-task", "Nova tarefa");
    var agendaItems = assessments.map(function (item) {
      return { id: item.id, kind: "assessment", date: item.date, time: item.time, title: item.title, subtitle: (courseById(item.courseId) || {}).name || item.type, icon: assessmentIcon(item.type), color: "orange" };
    }).concat(events.map(function (item) {
      return { id: item.id, kind: "event", date: item.date, time: item.time, title: item.title, subtitle: item.location || "Evento da faculdade", icon: "party-popper", color: "pink" };
    })).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    var agendaHtml = agendaItems.length ? agendaItems.map(function (item) {
      var detailAction = item.kind === "event" ? "show-event" : "assessment-scope";
      return '<div class="list-row"><span class="list-icon ' + item.color + '"><i data-lucide="' + item.icon + '"></i></span><span class="list-content"><strong>' + esc(item.title) + '</strong><small>' + relativeDate(item.date) + (item.time ? " · " + esc(item.time) : "") + ' · ' + esc(item.subtitle) + '</small></span><span class="badge">' + (item.kind === "event" ? "Evento" : "Avaliação") + '</span><button class="row-button" type="button" data-action="' + detailAction + '" data-id="' + attr(item.id) + '" aria-label="Ver detalhes"><i data-lucide="arrow-right"></i></button></div>';
    }).join("") : emptyState("calendar-days", "Agenda livre", "Adiciona testes ou eventos da faculdade quando souberes as datas.", "add-assessment", "Adicionar data");
    return '<div class="bento-grid" style="margin-top:15px"><article class="card span-6"><div class="card-title-row"><div><h3>Tarefas e revisões</h3></div><button class="button button-small" type="button" data-action="add-task"><i data-lucide="plus"></i>Adicionar</button></div><div class="list-stack">' + taskHtml + '</div></article><article class="card span-6"><div class="card-title-row"><div><h3>Avaliações e eventos</h3></div><div class="list-actions"><button class="row-button" type="button" data-action="add-event" aria-label="Adicionar evento"><i data-lucide="party-popper"></i></button><button class="row-button" type="button" data-action="add-assessment" aria-label="Adicionar avaliação"><i data-lucide="file-plus-2"></i></button></div></div><div class="list-stack">' + agendaHtml + "</div></article></div>";
  }

  function renderScheduleView() {
    var schedule = semesterItems("schedule").slice().sort(function (a, b) { return Number(a.weekday) - Number(b.weekday) || String(a.start).localeCompare(String(b.start)); });
    var days = [1, 2, 3, 4, 5];
    var todayDay = new Date().getDay();
    var board = days.map(function (day) {
      var entries = schedule.filter(function (item) { return Number(item.weekday) === day; });
      return '<section class="day-column ' + (day === todayDay ? "is-today" : "") + '"><div class="day-title"><strong>' + WEEKDAYS[day] + '</strong><span>' + SHORT_WEEKDAYS[day] + '</span></div>' + (entries.length ? entries.map(function (entry) {
        var course = courseById(entry.courseId);
        var nextOccurrence = nextOccurrenceForSchedule(entry);
        var prepared = nextOccurrence && nextOccurrence.lesson;
        return '<button class="schedule-block" style="--block-color:' + safeColor(course && course.color) + '" type="button" data-action="schedule-detail" data-id="' + attr(entry.id) + '"><time>' + esc(entry.start) + '–' + esc(entry.end) + '</time><strong>' + esc(course ? course.code || course.name : "Cadeira") + '</strong><small>' + esc(lessonTypeLabel(entry.type)) + (entry.room ? " · " + esc(entry.room) : "") + '</small>' + (prepared ? '<span class="schedule-prepared"><i data-lucide="check"></i>' + esc(prepared.title) + '</span>' : '') + "</button>";
      }).join("") : '<p class="card-subtitle">Sem aulas</p>') + "</section>";
    }).join("");
    var weekend = schedule.filter(function (item) { return Number(item.weekday) === 0 || Number(item.weekday) === 6; });
    return '<section class="card"><div class="card-title-row"><div><h3>Horário semanal</h3></div></div><div class="week-board" style="margin-top:17px">' + board + '</div>' + (weekend.length ? '<div class="form-note" style="margin-top:12px"><strong>Fim de semana:</strong> ' + weekend.map(function (entry) { var c = courseById(entry.courseId); return WEEKDAYS[entry.weekday] + " " + entry.start + " · " + (c ? c.name : "Cadeira"); }).join(" · ") + "</div>" : "") + '</section>';
  }

  function calendarEntriesForDate(dateValue) {
    var entries = semesterItems("lessons").filter(function (lesson) { return lesson.date === dateValue; }).map(function (lesson) {
      var course = courseById(lesson.courseId);
      return { kind: "lesson", id: lesson.id, title: lesson.title, time: lesson.start, color: safeColor(course && course.color), subtitle: course ? course.code || course.name : "Aula" };
    });
    entries = entries.concat(semesterItems("schedule").filter(function (entry) {
      return scheduleMatchesDate(entry, dateValue) && !linkedLessonForSlot(entry, dateValue);
    }).map(function (entry) {
      var course = courseById(entry.courseId);
      return { kind: "schedule", id: entry.id, title: course ? course.code || course.name : "Aula", time: entry.start, color: safeColor(course && course.color), subtitle: lessonTypeLabel(entry.type) + (entry.room ? " · " + entry.room : "") };
    }));
    entries = entries.concat(semesterItems("assessments").filter(function (item) { return item.date === dateValue; }).map(function (item) {
      return { kind: "assessment", id: item.id, title: item.title, time: item.time, color: "#ffad72", subtitle: item.type || "Avaliação" };
    }));
    entries = entries.concat(semesterItems("events").filter(function (item) { return item.date === dateValue; }).map(function (item) {
      return { kind: "event", id: item.id, title: item.title, time: item.time, color: "#ff92ae", subtitle: item.location || "Evento" };
    }));
    entries = entries.concat(semesterItems("tasks").filter(function (item) { return !item.done && item.dueDate === dateValue; }).map(function (item) {
      return { kind: "task", id: item.id, title: item.title, time: item.dueTime, color: "#a99df7", subtitle: "Tarefa", lessonId: item.lessonId };
    }));
    entries = entries.concat(semesterItems("studyBlocks").filter(function (item) { return item.date === dateValue; }).map(function (item) {
      var course = courseById(item.courseId);
      return { kind: "study-block", id: item.id, title: item.title, time: item.start, color: item.kind === "break" || item.kind === "lunch" ? "#f3e873" : safeColor(course && course.color, "#79cdb8"), subtitle: item.kind === "break" ? "Pausa" : item.kind === "lunch" ? "Almoço" : "Estudo" };
    }));
    return entries.sort(function (a, b) { return String(a.time || "99:99").localeCompare(String(b.time || "99:99")); });
  }

  function calendarEntryAction(item) {
    if (item.kind === "lesson") return 'data-route="lesson" data-id="' + attr(item.id) + '"';
    if (item.kind === "schedule") return 'data-action="schedule-detail" data-id="' + attr(item.id) + '"';
    if (item.kind === "assessment") return 'data-action="assessment-scope" data-id="' + attr(item.id) + '"';
    if (item.kind === "event") return 'data-action="show-event" data-id="' + attr(item.id) + '"';
    if (item.kind === "study-block") return 'data-action="edit-study-block" data-id="' + attr(item.id) + '"';
    if (item.lessonId) return 'data-route="lesson" data-id="' + attr(item.lessonId) + '"';
    return 'data-action="show-task" data-id="' + attr(item.id) + '"';
  }

  function renderCalendarEntry(item) {
    return '<button class="calendar-entry is-' + item.kind + '" type="button" ' + calendarEntryAction(item) + ' style="--entry-color:' + safeColor(item.color, "#a99df7") + '" title="' + attr(item.title) + '"><span></span><strong>' + esc(item.time ? item.time + " · " + item.title : item.title) + '</strong></button>';
  }

  function renderCalendarAgendaEntry(item) {
    return '<button class="calendar-agenda-entry is-' + item.kind + '" type="button" ' + calendarEntryAction(item) + ' style="--entry-color:' + safeColor(item.color, "#a99df7") + '"><time>' + esc(item.time || "Dia inteiro") + '</time><span><strong>' + esc(item.title) + '</strong><small>' + esc(item.subtitle || "") + '</small></span><i data-lucide="arrow-up-right"></i></button>';
  }

  function addCalendarDays(value, amount) {
    var date = localDate(value) || new Date();
    date.setDate(date.getDate() + Number(amount || 0));
    return todayISO(date);
  }

  function calendarWeekStart(value) {
    var date = localDate(value) || new Date();
    var offset = (date.getDay() + 6) % 7;
    date.setDate(date.getDate() - offset);
    return todayISO(date);
  }

  function calendarViewControl(active) {
    var views = [
      { id: "day", label: "Dia" },
      { id: "three", label: "3 dias" },
      { id: "week", label: "Semana" },
      { id: "month", label: "Mês" }
    ];
    return '<div class="calendar-view-control" role="group" aria-label="Intervalo do calendário">' + views.map(function (item) { return '<button type="button" class="' + (active === item.id ? "is-active" : "") + '" data-action="calendar-view" data-view="' + item.id + '">' + item.label + '</button>'; }).join("") + '</div>';
  }

  function calendarLegend() {
    return '<div class="calendar-legend"><span><i class="lesson"></i>Aula</span><span><i class="schedule"></i>Horário</span><span><i class="assessment"></i>Avaliação</span><span><i class="event"></i>Evento</span><span><i class="task"></i>Tarefa</span><span><i class="study-block"></i>Estudo</span></div>';
  }

  function calendarToolbar(title, activeView) {
    var unit = activeView === "month" ? "período" : activeView === "week" ? "semana" : activeView === "three" ? "3 dias" : "dia";
    return '<div class="calendar-toolbar"><div><p class="card-label">Calendário académico</p><h3>' + esc(title) + '</h3></div><div class="calendar-toolbar-actions">' + calendarViewControl(activeView) + '<div class="calendar-nav"><button class="row-button" type="button" data-action="calendar-shift" data-delta="-1" aria-label="' + esc(unit + ' anterior') + '"><i data-lucide="chevron-left"></i></button><button class="button button-small" type="button" data-action="calendar-today">Hoje</button><button class="row-button" type="button" data-action="calendar-shift" data-delta="1" aria-label="' + esc(unit + ' seguinte') + '"><i data-lucide="chevron-right"></i></button></div></div></div>';
  }

  function renderMonthCalendar() {
    var cursorDate = localDate(calendarCursor) || new Date();
    var year = cursorDate.getFullYear();
    var month = cursorDate.getMonth();
    var first = new Date(year, month, 1, 12);
    var daysInMonth = new Date(year, month + 1, 0, 12).getDate();
    var leading = (first.getDay() + 6) % 7;
    var total = Math.ceil((leading + daysInMonth) / 7) * 7;
    var cells = [];
    for (var index = 0; index < total; index += 1) {
      var day = index - leading + 1;
      if (day < 1 || day > daysInMonth) {
        cells.push('<div class="calendar-day is-outside" aria-hidden="true"></div>');
        continue;
      }
      var dateValue = year + "-" + String(month + 1).padStart(2, "0") + "-" + String(day).padStart(2, "0");
      var items = calendarEntriesForDate(dateValue);
      cells.push('<div class="calendar-day ' + (dateValue === todayISO() ? "is-today" : "") + '"><div class="calendar-day-number"><time datetime="' + dateValue + '">' + day + '</time>' + (dateValue === todayISO() ? '<span>Hoje</span>' : '') + '</div><div class="calendar-day-items">' + items.slice(0, 3).map(renderCalendarEntry).join("") + (items.length > 3 ? '<small class="calendar-more">+' + (items.length - 3) + ' itens</small>' : '') + '</div></div>');
    }
    var monthTitle = new Intl.DateTimeFormat("pt-PT", { month: "long", year: "numeric" }).format(first);
    monthTitle = monthTitle.charAt(0).toUpperCase() + monthTitle.slice(1);
    return '<section class="card calendar-card">' + calendarToolbar(monthTitle, "month") + '<div class="calendar-weekdays">' + ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"].map(function (dayName) { return '<span>' + dayName + '</span>'; }).join("") + '</div><div class="month-grid">' + cells.join("") + '</div>' + calendarLegend() + '</section>';
  }

  function renderCalendarRange(activeView) {
    var count = activeView === "week" ? 7 : activeView === "three" ? 3 : 1;
    var startISO = activeView === "week" ? calendarWeekStart(calendarCursor) : calendarCursor;
    var dates = Array.from({ length: count }, function (_, index) { return addCalendarDays(startISO, index); });
    var first = localDate(dates[0]);
    var last = localDate(dates[dates.length - 1]);
    var title;
    if (count === 1) title = new Intl.DateTimeFormat("pt-PT", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(first);
    else if (first.getMonth() === last.getMonth()) title = first.getDate() + "–" + last.getDate() + " de " + new Intl.DateTimeFormat("pt-PT", { month: "long", year: "numeric" }).format(last);
    else title = new Intl.DateTimeFormat("pt-PT", { day: "numeric", month: "short" }).format(first) + " – " + new Intl.DateTimeFormat("pt-PT", { day: "numeric", month: "short", year: "numeric" }).format(last);
    title = title.charAt(0).toUpperCase() + title.slice(1);
    var columns = dates.map(function (dateValue) {
      var date = localDate(dateValue);
      var items = calendarEntriesForDate(dateValue);
      var dayName = new Intl.DateTimeFormat("pt-PT", { weekday: "short" }).format(date).replace(".", "");
      var dateLabel = new Intl.DateTimeFormat("pt-PT", { day: "numeric", month: "short" }).format(date).replace(".", "");
      return '<section class="calendar-agenda-day ' + (dateValue === todayISO() ? "is-today" : "") + '"><header><span>' + esc(dayName) + '</span><strong>' + esc(dateLabel) + '</strong>' + (dateValue === todayISO() ? '<small>Hoje</small>' : '') + '</header><div class="calendar-agenda-list">' + (items.length ? items.map(renderCalendarAgendaEntry).join("") : '<p class="calendar-agenda-empty">Sem aulas, prazos ou eventos.</p>') + '</div></section>';
    }).join("");
    return '<section class="card calendar-card">' + calendarToolbar(title, activeView) + '<div class="calendar-range-grid view-' + activeView + '" style="--calendar-days:' + count + '">' + columns + '</div>' + calendarLegend() + '</section>';
  }

  function renderCalendarView() {
    var activeView = state.settings.calendarView || "month";
    return activeView === "month" ? renderMonthCalendar() : renderCalendarRange(activeView);
  }

  function minutesToTime(value) {
    var minutes = Math.max(0, Math.min(1439, Math.round(Number(value) || 0)));
    return String(Math.floor(minutes / 60)).padStart(2, "0") + ":" + String(minutes % 60).padStart(2, "0");
  }

  function daysUntil(value) {
    if (!value) return 999;
    return Math.round((new Date(value + "T12:00:00") - new Date(todayISO() + "T12:00:00")) / 86400000);
  }

  function studySource(type, id) {
    var item;
    var course;
    if (type === "task") {
      item = state.tasks.find(function (entry) { return entry.id === id; });
      course = item && courseById(item.courseId);
      return item ? { type: type, id: item.id, title: item.title, courseId: item.courseId || null, icon: taskIcon(item.type), duration: item.type === "project" ? 90 : Number(state.settings.studySessionMinutes), meta: (course ? course.name + " · " : "") + relativeDate(item.dueDate), score: (item.priority === "high" ? 90 : 50) - Math.min(20, daysUntil(item.dueDate)) } : null;
    }
    if (type === "lesson") {
      item = lessonById(id);
      course = item && courseById(item.courseId);
      var pastCount = item ? pastQuestionsForLesson(item.id).length : 0;
      return item ? { type: type, id: item.id, title: "Rever " + item.title, courseId: item.courseId, icon: "presentation", duration: Number(state.settings.studySessionMinutes), meta: (course ? course.name + " · " : "") + pastCount + " pergunta(s) anterior(es)", score: 55 + pastCount * 4 } : null;
    }
    if (type === "quiz") {
      item = state.quizzes.find(function (entry) { return entry.id === id; });
      course = item && courseById(item.courseId);
      return item ? { type: type, id: item.id, title: item.title, courseId: item.courseId, icon: "sparkles", duration: 30, meta: (course ? course.name + " · " : "") + asArray(item.questions).length + " pergunta(s)", score: item.lastScore == null ? 60 : 45 + Math.max(0, 100 - item.lastScore) / 4 } : null;
    }
    if (type === "assessment") {
      item = assessmentById(id);
      course = item && courseById(item.courseId);
      return item ? { type: type, id: item.id, title: "Preparar " + item.title, courseId: item.courseId, icon: assessmentIcon(item.type), duration: 90, meta: (course ? course.name + " · " : "") + relativeDate(item.date), score: 110 - Math.min(60, Math.max(0, daysUntil(item.date)) * 3) + (Number(item.weight) || 0) / 2 } : null;
    }
    return null;
  }

  function studyBacklog() {
    var items = [];
    semesterItems("tasks").filter(function (task) { return !task.done && task.type !== "lesson-quiz"; }).forEach(function (task) { var source = studySource("task", task.id); if (source) items.push(source); });
    semesterItems("lessons").filter(function (lesson) { return !lesson.mastered && lessonHasEnded(lesson); }).forEach(function (lesson) { var source = studySource("lesson", lesson.id); if (source) items.push(source); });
    semesterItems("quizzes").filter(function (quiz) { return quiz.lastScore == null || Number(quiz.lastScore) < 85; }).forEach(function (quiz) { var source = studySource("quiz", quiz.id); if (source) items.push(source); });
    semesterItems("assessments").filter(function (assessment) { return !assessment.date || daysUntil(assessment.date) >= 0; }).forEach(function (assessment) { var source = studySource("assessment", assessment.id); if (source) items.push(source); });
    return items.sort(function (a, b) { return b.score - a.score; });
  }

  function studyBlocksForDate(dateValue) {
    return semesterItems("studyBlocks").filter(function (block) { return block.date === dateValue; }).sort(function (a, b) { return String(a.start).localeCompare(String(b.start)); });
  }

  function firstFreeStudyTime(dateValue, duration) {
    var start = timeMinutes(state.settings.studyDayStart || "09:00");
    var end = timeMinutes(state.settings.studyDayEnd || "19:00");
    var needed = Number(duration) || Number(state.settings.studySessionMinutes) || 50;
    var blocks = studyBlocksForDate(dateValue);
    for (var cursor = start; cursor + needed <= end; cursor += 10) {
      var collision = blocks.some(function (block) { return cursor < timeMinutes(block.end) && cursor + needed > timeMinutes(block.start); });
      if (!collision) return minutesToTime(cursor);
    }
    return state.settings.studyDayStart || "09:00";
  }

  function newStudyBlockFromSource(source, dateValue, startValue) {
    var start = timeMinutes(startValue || firstFreeStudyTime(dateValue, source.duration));
    var endLimit = timeMinutes(state.settings.studyDayEnd || "19:00");
    var end = Math.min(endLimit, start + (Number(source.duration) || 50));
    return { id: uid("studyblock"), semesterId: state.currentSemesterId, date: dateValue, title: source.title, start: minutesToTime(start), end: minutesToTime(end), kind: "study", courseId: source.courseId || null, sourceType: source.type, sourceId: source.id, completed: false, notes: "" };
  }

  async function scheduleStudySource(type, id, startValue) {
    var source = studySource(type, id);
    if (!source) { toast("Este item já não está disponível.", "warning"); return; }
    var dateValue = state.settings.studyPlanDate || todayISO();
    var duplicate = studyBlocksForDate(dateValue).some(function (block) { return block.sourceType === type && block.sourceId === id; });
    if (duplicate) { toast("Este item já está planeado neste dia.", "warning"); return; }
    state.studyBlocks.push(newStudyBlockFromSource(source, dateValue, startValue));
    await save(true);
    render();
    toast("Bloco adicionado ao dia de estudo.");
  }

  function renderStudyBacklogCard(source) {
    var course = courseById(source.courseId);
    return '<article class="study-backlog-item" draggable="true" data-study-source-type="' + attr(source.type) + '" data-study-source-id="' + attr(source.id) + '" style="--study-color:' + safeColor(course && course.color, "#a99df7") + '"><span class="list-icon"><i data-lucide="' + source.icon + '"></i></span><div><strong>' + esc(source.title) + '</strong><small>' + esc(source.meta) + '</small></div><button class="row-button" type="button" data-action="schedule-study-source" data-source-type="' + attr(source.type) + '" data-source-id="' + attr(source.id) + '" aria-label="Agendar"><i data-lucide="plus"></i></button></article>';
  }

  function studyBlockClass(block) {
    if (block.kind === "break") return "is-break";
    if (block.kind === "lunch") return "is-lunch";
    return block.completed ? "is-completed" : "";
  }

  function renderStudyDay() {
    var dateValue = state.settings.studyPlanDate || todayISO();
    var start = timeMinutes(state.settings.studyDayStart || "09:00");
    var end = timeMinutes(state.settings.studyDayEnd || "19:00");
    if (end <= start) end = start + 600;
    var slotSize = 10;
    var slotCount = Math.ceil((end - start) / slotSize);
    var blocks = studyBlocksForDate(dateValue);
    var timeline = "";
    for (var slot = 0; slot < slotCount; slot += 1) {
      var slotTime = minutesToTime(start + slot * slotSize);
      var majorSlot = ((start + slot * slotSize) % 30) === 0;
      timeline += (majorSlot ? '<time class="study-time-label" style="grid-row:' + (slot + 1) + '">' + esc(slotTime) + '</time>' : '') + '<div class="study-drop-slot ' + (majorSlot ? "is-major" : "") + '" data-study-drop="true" data-time="' + attr(slotTime) + '" style="grid-row:' + (slot + 1) + '" aria-label="Agendar às ' + attr(slotTime) + '"></div>';
    }
    timeline += blocks.map(function (block) {
      var rowStart = Math.max(1, Math.floor((timeMinutes(block.start) - start) / slotSize) + 1);
      var span = Math.max(1, Math.ceil((timeMinutes(block.end) - timeMinutes(block.start)) / slotSize));
      var course = courseById(block.courseId);
      return '<article class="study-time-block ' + studyBlockClass(block) + '" draggable="true" data-study-source-type="block" data-study-source-id="' + attr(block.id) + '" style="--study-color:' + safeColor(course && course.color, block.kind === "study" ? "#79cdb8" : "#f3e873") + ';grid-row:' + rowStart + ' / span ' + span + '"><button type="button" data-action="edit-study-block" data-id="' + attr(block.id) + '"><span><time>' + esc(block.start) + '–' + esc(block.end) + '</time><strong>' + esc(block.title) + '</strong><small>' + esc(course ? course.code || course.name : block.kind === "break" ? "Pausa" : block.kind === "lunch" ? "Almoço" : "Estudo") + '</small></span></button>' + (block.kind === "study" ? '<button class="study-block-check" type="button" data-action="toggle-study-block" data-id="' + attr(block.id) + '" aria-label="' + (block.completed ? "Reabrir" : "Concluir") + '"><i data-lucide="check"></i></button>' : '') + '</article>';
    }).join("");
    var scheduledKeys = blocks.map(function (block) { return block.sourceType + ":" + block.sourceId; });
    var backlog = studyBacklog().filter(function (source) { return scheduledKeys.indexOf(source.type + ":" + source.id) < 0; });
    var totalMinutes = blocks.filter(function (block) { return block.kind === "study"; }).reduce(function (sum, block) { return sum + Math.max(0, timeMinutes(block.end) - timeMinutes(block.start)); }, 0);
    var longDate = formatLongDate(dateValue);
    return '<section class="study-day-shell"><aside class="study-backlog"><div class="study-panel-head"><div><p class="card-label">Por planear</p><h3>' + backlog.length + ' itens</h3></div><span class="badge badge-violet">' + round(totalMinutes / 60, 1) + ' h</span></div><div class="study-backlog-list">' + (backlog.length ? backlog.slice(0, 14).map(renderStudyBacklogCard).join("") : '<div class="past-question-empty"><i data-lucide="check-check"></i><span>Sem itens pendentes para este dia.</span></div>') + '</div></aside><section class="card study-timeline-card"><div class="study-day-toolbar"><div><p class="card-label">' + esc(longDate.charAt(0).toUpperCase() + longDate.slice(1)) + '</p><h3>Plano do dia</h3></div><div class="study-day-nav"><button class="row-button" type="button" data-action="study-date-shift" data-delta="-1"><i data-lucide="chevron-left"></i></button><input type="date" data-role="study-plan-date" value="' + attr(dateValue) + '"><button class="row-button" type="button" data-action="study-date-shift" data-delta="1"><i data-lucide="chevron-right"></i></button></div></div><div class="study-timeline-grid" style="--study-rows:' + slotCount + '">' + timeline + '</div></section></section>';
  }

  function intervalIsFree(start, end, blocks) {
    return !blocks.some(function (block) { return start < timeMinutes(block.end) && end > timeMinutes(block.start); });
  }

  async function autoFillStudyDay() {
    var dateValue = state.settings.studyPlanDate || todayISO();
    var dayStart = timeMinutes(state.settings.studyDayStart || "09:00");
    var dayEnd = timeMinutes(state.settings.studyDayEnd || "19:00");
    var lunchStart = timeMinutes(state.settings.studyLunchStart || "13:00");
    var lunchEnd = lunchStart + Number(state.settings.studyLunchMinutes || 0);
    var occupancy = studyBlocksForDate(dateValue).slice();
    var used = occupancy.map(function (block) { return block.sourceType + ":" + block.sourceId; });
    var candidates = studyBacklog().filter(function (source) { return used.indexOf(source.type + ":" + source.id) < 0; });
    if (!candidates.length) { toast("Não existem itens pendentes para preencher.", "warning"); return; }
    if (state.settings.studyLunchMinutes > 0 && intervalIsFree(lunchStart, lunchEnd, occupancy) && lunchStart >= dayStart && lunchEnd <= dayEnd) {
      var lunch = { id: uid("studyblock"), semesterId: state.currentSemesterId, date: dateValue, title: "Almoço", start: minutesToTime(lunchStart), end: minutesToTime(lunchEnd), kind: "lunch", courseId: null, sourceType: "routine", sourceId: null, completed: false, notes: "" };
      occupancy.push(lunch); state.studyBlocks.push(lunch);
    }
    var cursor = dayStart;
    var added = 0;
    candidates.slice(0, 12).forEach(function (source) {
      var duration = Math.min(Number(source.duration) || Number(state.settings.studySessionMinutes), dayEnd - dayStart);
      var found = null;
      for (var minute = cursor; minute + duration <= dayEnd; minute += 10) {
        if (intervalIsFree(minute, minute + duration, occupancy)) { found = minute; break; }
      }
      if (found == null) {
        for (var restart = dayStart; restart + duration <= dayEnd; restart += 10) {
          if (intervalIsFree(restart, restart + duration, occupancy)) { found = restart; break; }
        }
      }
      if (found == null) return;
      var block = newStudyBlockFromSource(source, dateValue, minutesToTime(found));
      occupancy.push(block); state.studyBlocks.push(block); added += 1;
      cursor = timeMinutes(block.end);
      var breakMinutes = Number(state.settings.studyBreakMinutes || 0);
      if (breakMinutes && cursor + breakMinutes <= dayEnd && intervalIsFree(cursor, cursor + breakMinutes, occupancy)) {
        var breakBlock = { id: uid("studyblock"), semesterId: state.currentSemesterId, date: dateValue, title: "Pausa", start: minutesToTime(cursor), end: minutesToTime(cursor + breakMinutes), kind: "break", courseId: null, sourceType: "routine", sourceId: null, completed: false, notes: "" };
        occupancy.push(breakBlock); state.studyBlocks.push(breakBlock); cursor += breakMinutes;
      }
    });
    if (!added) { toast("Não há espaço livre suficiente neste dia.", "warning"); return; }
    await save(true); render(); toast(added + " bloco(s) de estudo adicionados.");
  }

  function renderPlanner() {
    var mode = ["calendar", "study-day"].indexOf(state.settings.plannerView) >= 0 ? state.settings.plannerView : "schedule";
    setHeader(mode === "calendar" ? "Calendário" : mode === "study-day" ? "Dia de estudo" : "Horário", "Agenda académica");
    var primary = mode === "calendar" ? renderCalendarView() : mode === "study-day" ? renderStudyDay() : renderScheduleView();
    var title = mode === "calendar" ? "Calendário do semestre" : mode === "study-day" ? "Planeamento diário" : "Horário semanal";
    var copy = mode === "calendar" ? "Aulas, avaliações, eventos, tarefas e blocos de estudo." : mode === "study-day" ? "Arrasta itens para uma hora ou usa Agendar no telemóvel." : "Os períodos do horário determinam a aula em curso.";
    var actions = mode === "study-day" ? '<button class="button" type="button" data-action="study-planner-settings"><i data-lucide="sliders-horizontal"></i>Configurar</button><button class="button" type="button" data-action="copy-study-day"><i data-lucide="copy"></i>Copiar rotina</button><button class="button" type="button" data-action="add-study-block"><i data-lucide="plus"></i>Bloco</button><button class="button button-dark" type="button" data-action="auto-fill-study-day"><i data-lucide="sparkles"></i>Preencher dia</button>' : '<button class="button" type="button" data-action="add-schedule"><i data-lucide="calendar-plus"></i>Bloco do horário</button><button class="button button-dark" type="button" data-action="create-lesson"><i data-lucide="plus"></i>Preparar aula</button>';
    return '<div class="page-head"><div><h2>' + title + '</h2><p>' + copy + '</p></div><div class="page-actions">' + plannerModeControl(mode) + actions + '</div></div>' + primary + (mode === "study-day" ? "" : plannerSupportingCards());
  }

  function weeklyStudyEstimates() {
    var courses = activeCourses();
    var budget = Number(state.settings.weeklyStudyHours || 16);
    if (!courses.length) return [];
    var scored = courses.map(function (course) {
      var pendingTasks = semesterItems("tasks").filter(function (task) { return task.courseId === course.id && !task.done; }).length;
      var pendingLessons = semesterItems("lessons").filter(function (lesson) { return lesson.courseId === course.id && !lesson.mastered && lessonHasEnded(lesson); }).length;
      var nextAssessment = semesterItems("assessments").filter(function (assessment) { return assessment.courseId === course.id && (!assessment.date || daysUntil(assessment.date) >= 0); }).sort(function (a, b) { return daysUntil(a.date) - daysUntil(b.date); })[0];
      var urgency = nextAssessment ? Math.max(0, 30 - Math.min(30, daysUntil(nextAssessment.date))) / 6 : 0;
      var assessmentWeight = nextAssessment ? Number(nextAssessment.weight) || 0 : 0;
      var score = Math.max(1, Number(course.ects) || 1) + pendingTasks * 1.25 + pendingLessons * .75 + urgency + assessmentWeight / 12;
      return { course: course, score: score, pendingTasks: pendingTasks, pendingLessons: pendingLessons, nextAssessment: nextAssessment };
    });
    var totalScore = scored.reduce(function (sum, item) { return sum + item.score; }, 0) || 1;
    scored.forEach(function (item) { item.hours = Math.max(.5, Math.round((budget * item.score / totalScore) * 2) / 2); });
    var difference = Math.round((budget - scored.reduce(function (sum, item) { return sum + item.hours; }, 0)) * 2) / 2;
    if (scored.length && difference) scored[0].hours = Math.max(.5, scored[0].hours + difference);
    return scored.sort(function (a, b) { return b.hours - a.hours; });
  }

  function renderStudyHourEstimate() {
    var estimates = weeklyStudyEstimates();
    return '<article class="card span-12 study-hours-card"><div class="card-title-row"><div><p class="card-label">Estimativa semanal</p><h3>' + round(state.settings.weeklyStudyHours || 16, 1) + ' horas distribuídas</h3><p class="card-subtitle">ECTS, proximidade das avaliações, respetivo peso e trabalho pendente.</p></div><button class="button button-small" type="button" data-action="study-planner-settings"><i data-lucide="sliders-horizontal"></i>Editar horas</button></div><div class="study-hours-grid">' + (estimates.length ? estimates.map(function (item) {
      var next = item.nextAssessment ? item.nextAssessment.title + " · " + relativeDate(item.nextAssessment.date) : "Sem avaliação próxima";
      return '<div class="study-hours-item" style="--course-color:' + safeColor(item.course.color) + '"><span></span><div><strong>' + esc(item.course.code || item.course.name) + '</strong><small>' + esc(next) + '</small></div><b>' + round(item.hours, 1) + ' h</b></div>';
    }).join("") : '<p class="card-subtitle">Adiciona cadeiras para calcular a distribuição.</p>') + '</div></article>';
  }

  function currentWeekStart() { return calendarWeekStart(todayISO()); }

  function weeklyReviewRecord() {
    var start = currentWeekStart();
    return semesterItems("weeklyReviews").find(function (review) { return review.weekStart === start; }) || null;
  }

  function renderWeeklyReview() {
    setHeader("Revisão semanal", "Estudar");
    var weekStart = currentWeekStart();
    var weekEnd = addCalendarDays(weekStart, 6);
    var review = weeklyReviewRecord();
    var overdue = semesterItems("tasks").filter(function (task) { return !task.done && task.dueDate && task.dueDate < todayISO(); });
    var unreviewed = beOnlineStatus().pending;
    var quizCompleted = semesterItems("quizzes").filter(function (quiz) { return quiz.lastCompletedAt && quiz.lastCompletedAt.slice(0, 10) >= weekStart && quiz.lastCompletedAt.slice(0, 10) <= weekEnd; });
    var upcoming = semesterItems("assessments").filter(function (assessment) { var distance = daysUntil(assessment.date); return distance >= 0 && distance <= 14; }).sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
    var plannedMinutes = semesterItems("studyBlocks").filter(function (block) { return block.date >= weekStart && block.date <= weekEnd && block.kind === "study"; }).reduce(function (sum, block) { return sum + Math.max(0, timeMinutes(block.end) - timeMinutes(block.start)); }, 0);
    var priorityHtml = review && asArray(review.priorities).length ? '<ol class="review-priority-list">' + review.priorities.map(function (item) { return '<li>' + esc(item) + '</li>'; }).join("") + '</ol>' : '<p class="card-subtitle">Ainda não definiste prioridades para esta semana.</p>';
    var doubtHtml = review && asArray(review.doubts).length ? review.doubts.map(function (item) { return '<span class="review-doubt"><i data-lucide="circle-help"></i>' + esc(item) + '</span>'; }).join("") : '<p class="card-subtitle">Sem dúvidas registadas nesta revisão.</p>';
    return '<div class="page-head"><div><button class="button button-ghost button-small" type="button" data-route="study"><i data-lucide="arrow-left"></i>Estudar</button><h2 style="margin-top:11px">Revisão da semana</h2><p>' + esc(formatDate(weekStart)) + '–' + esc(formatDate(weekEnd)) + '</p></div><div class="page-actions"><span class="badge ' + (review && review.completedAt ? "badge-mint" : "badge-violet") + '">' + (review && review.completedAt ? "Concluída" : "Por concluir") + '</span><button class="button button-dark" type="button" data-action="weekly-review"><i data-lucide="clipboard-check"></i>' + (review ? "Atualizar revisão" : "Fazer revisão") + '</button></div></div><div class="bento-grid"><article class="card card-pink span-3 metric-card"><div class="metric-top"><p class="card-label">Tarefas atrasadas</p><span class="metric-icon"><i data-lucide="triangle-alert"></i></span></div><div><p class="metric-value">' + overdue.length + '</p><p class="metric-caption">por concluir</p></div></article><article class="card card-yellow span-3 metric-card"><div class="metric-top"><p class="card-label">Aulas por rever</p><span class="metric-icon"><i data-lucide="rotate-ccw"></i></span></div><div><p class="metric-value">' + unreviewed.length + '</p><p class="metric-caption">quiz de aula pendente</p></div></article><article class="card card-mint span-3 metric-card"><div class="metric-top"><p class="card-label">Quizzes concluídos</p><span class="metric-icon"><i data-lucide="badge-check"></i></span></div><div><p class="metric-value">' + quizCompleted.length + '</p><p class="metric-caption">esta semana</p></div></article><article class="card card-violet span-3 metric-card"><div class="metric-top"><p class="card-label">Estudo planeado</p><span class="metric-icon"><i data-lucide="timer"></i></span></div><div><p class="metric-value">' + round(plannedMinutes / 60, 1) + '</p><p class="metric-caption">horas em blocos</p></div></article><article class="card span-6"><div class="card-title-row"><div><h3>Prioridades</h3></div></div>' + priorityHtml + '</article><article class="card span-6"><div class="card-title-row"><div><h3>Dúvidas a esclarecer</h3></div></div><div class="review-doubts">' + doubtHtml + '</div></article><article class="card span-12"><div class="card-title-row"><div><h3>Próximas duas semanas</h3></div><span class="badge badge-orange">' + upcoming.length + ' avaliações</span></div><div class="list-stack">' + (upcoming.length ? upcoming.map(function (assessment) { var course = courseById(assessment.courseId); return '<div class="list-row"><span class="list-icon orange"><i data-lucide="' + assessmentIcon(assessment.type) + '"></i></span><span class="list-content"><strong>' + esc(assessment.title) + '</strong><small>' + esc(course ? course.name : "Cadeira") + ' · ' + relativeDate(assessment.date) + '</small></span><button class="row-button" type="button" data-action="assessment-scope" data-id="' + attr(assessment.id) + '"><i data-lucide="arrow-right"></i></button></div>'; }).join("") : '<p class="card-subtitle">Não existem avaliações marcadas nos próximos 14 dias.</p>') + '</div></article>' + renderStudyHourEstimate() + '</div>';
  }


  function aiProjectById(id) {
    return state.aiProjects.find(function (project) { return project.id === id; }) || null;
  }

  function aiOutputLabel(value) {
    if (value === "notes") return "Apontamentos";
    if (value === "summary") return "Resumo rápido";
    if (value === "quiz") return "Quiz";
    if (value === "flashcards") return "Flashcards";
    return "Tudo";
  }

  function aiDifficultyLabel(value) {
    if (value === "easy") return "Fácil";
    if (value === "medium") return "Média";
    if (value === "hard") return "Difícil";
    return "Automática";
  }

  function setAIProgress(title, detail, progress) {
    aiProgress = { active: true, title: title || "A preparar a IA…", detail: detail || "Aguarda um momento.", progress: progress == null ? null : clamp(progress, 0, 100) };
    updateAIProgressDOM();
  }

  function clearAIProgress() {
    aiProgress = { active: false, progress: null, title: "", detail: "" };
    updateAIProgressDOM();
  }

  function updateAIProgressDOM() {
    var card = document.getElementById("aiProgressCard");
    if (!card) return;
    card.hidden = !aiProgress.active;
    var title = document.getElementById("aiProgressTitle");
    var detail = document.getElementById("aiProgressDetail");
    var bar = document.getElementById("aiProgressBar");
    var track = bar && bar.parentElement;
    if (title) title.textContent = aiProgress.title || "A preparar a IA…";
    if (detail) detail.textContent = aiProgress.detail || "Aguarda um momento.";
    if (track) track.classList.toggle("is-indeterminate", aiProgress.progress == null);
    if (bar) bar.style.width = aiProgress.progress == null ? "38%" : clamp(aiProgress.progress, 2, 100) + "%";
  }

  function aiSourceButtons(project, slides) {
    var unique = Array.from(new Set(asArray(slides).map(Number).filter(Number.isFinite))).slice(0, 12);
    if (!unique.length) return "";
    return '<div class="ai-source-list">' + unique.map(function (number) {
      return '<button class="ai-source-chip" type="button" data-action="ai-open-slide" data-project="' + attr(project.id) + '" data-slide="' + number + '"><i data-lucide="presentation"></i>Slide ' + number + '</button>';
    }).join("") + '</div>';
  }

  function renderAIProgress() {
    return '<article id="aiProgressCard" class="card card-dark span-12 ai-progress-card" ' + (aiProgress.active ? '' : 'hidden') + '><div class="ai-progress-head"><span class="ai-model-spinner" aria-hidden="true"></span><div><p class="card-label">Twenty AI · sincronização</p><h3 id="aiProgressTitle">' + esc(aiProgress.title || "A preparar a IA…") + '</h3><p id="aiProgressDetail" class="card-subtitle">' + esc(aiProgress.detail || "Aguarda um momento.") + '</p></div><button class="icon-button ai-cancel-button" type="button" data-action="ai-cancel" aria-label="Cancelar geração"><i data-lucide="x"></i></button></div><div class="ai-progress-track ' + (aiProgress.progress == null ? 'is-indeterminate' : '') + '"><span id="aiProgressBar" style="width:' + (aiProgress.progress == null ? '38' : clamp(aiProgress.progress, 2, 100)) + '%"></span></div><small>Não feches este separador durante o upload, download ou geração.</small></article>';
  }

  function renderAIProjectCard(project) {
    var course = courseById(project.courseId);
    var outputCount = asArray(project.quizQuestions).length + asArray(project.flashcards).length + asArray(project.notes && project.notes.sections).length;
    var hasOutput = outputCount > 0 || !!project.summary;
    var hasFile = !!(project.remoteFile && project.remoteFile.path);
    var badge = hasOutput ? (project.modelMode === "quality" ? "IA qualidade" : "IA pronta") : hasFile ? "PPT sincronizado" : "Só neste dispositivo";
    var badgeClass = hasOutput ? "badge-violet" : hasFile ? "badge-mint" : "badge-yellow";
    var icon = hasOutput ? "brain" : "presentation";
    var actions = '<button class="button button-dark button-small" type="button" data-action="ai-open-project" data-id="' + attr(project.id) + '"><i data-lucide="arrow-up-right"></i>Abrir</button>';
    if (hasFile) actions += '<button class="button button-small" type="button" data-action="ai-download-pptx" data-id="' + attr(project.id) + '"><i data-lucide="download"></i>PPT</button>';
    if (!hasOutput) actions += '<button class="button button-small" type="button" data-action="ai-use-project" data-id="' + attr(project.id) + '"><i data-lucide="sparkles"></i>Gerar</button>';
    else if (project.quizId) actions += '<button class="button button-small" type="button" data-action="ai-start-quiz" data-id="' + attr(project.id) + '"><i data-lucide="play"></i>Quiz</button>';
    return '<article class="card ai-project-card span-4"><div class="card-title-row"><div><span class="badge ' + badgeClass + '">' + esc(badge) + '</span><h3 style="margin-top:12px">' + esc(project.title || project.fileName || "Slides") + '</h3><p class="card-subtitle">' + esc(course ? course.code || course.name : "Sem cadeira") + ' · ' + Number(project.slideCount || 0) + ' slides</p></div><span class="metric-icon"><i data-lucide="' + icon + '"></i></span></div><div class="ai-project-stats"><span><i data-lucide="cloud-check"></i>' + (hasFile ? "Guardado no Git" : "Local") + '</span><span><i data-lucide="hard-drive"></i>' + esc(formatBytes(project.fileSize || project.remoteFile && project.remoteFile.size || 0)) + '</span></div><div class="list-actions">' + actions + '</div></article>';
  }

  function renderStudyAI() {
    setHeader("IA de estudo", "Estudar");
    var supported = !!(AI && AI.supportsWebGPU && AI.supportsWebGPU());
    var recommendation = AI && AI.selectedModel ? AI.selectedModel("auto") : { mode: "fast", label: "Modelo rápido", size: "" };
    var projects = semesterItems("aiProjects").slice().sort(function (a, b) { return String(b.createdAt || "").localeCompare(String(a.createdAt || "")); });
    var courseValue = aiDraft && aiDraft.courseId || "";
    var modelValue = state.settings.aiModelMode || "auto";
    var outputValue = state.settings.aiOutput || "all";
    var difficultyValue = state.settings.aiDifficulty || "auto";
    var questionValue = Number(state.settings.aiQuestionCount) || 10;
    var draftHtml = aiDraft ? '<div class="ai-file-ready"><span class="ai-file-icon"><i data-lucide="presentation"></i></span><div><strong>' + esc(aiDraft.fileName) + '</strong><small>' + Number(aiDraft.slideCount || 0) + ' slides · ' + formatBytes(aiDraft.fileSize || 0) + ' · ' + (aiDraft.remoteFile && aiDraft.remoteFile.path ? 'sincronizado no Git' : 'só neste dispositivo') + '</small></div><button class="icon-button" type="button" data-action="ai-clear-draft" aria-label="Fechar PowerPoint selecionado"><i data-lucide="x"></i></button></div>' : '<button class="ai-upload-zone" type="button" data-action="ai-pick-pptx"><span class="ai-upload-icon"><i data-lucide="upload-cloud"></i></span><strong>Escolher e sincronizar PowerPoint</strong><small>.pptx · até 25 MB · upload com progresso para o repositório privado</small></button>';
    var modelWarning = supported ? '<span class="badge badge-mint"><i data-lucide="cpu"></i>WebGPU disponível</span>' : '<span class="badge badge-yellow"><i data-lucide="triangle-alert"></i>WebGPU indisponível</span>';
    var supportCopy = supported ? 'O modelo corre no browser e fica guardado neste dispositivo depois do primeiro download.' : 'Abre esta página no Chrome atualizado. Sem WebGPU, a geração local não consegue arrancar neste dispositivo.';
    var projectHtml = projects.length ? projects.map(renderAIProjectCard).join("") : '<div class="span-12">' + emptyState("brain", "Ainda não tens projetos de IA", "Envia um PowerPoint para o Git e cria apontamentos, quizzes e flashcards sem API.", "ai-pick-pptx", "Escolher PowerPoint") + '</div>';

    return '<div class="page-head"><div><button class="button button-ghost button-small" type="button" data-route="study"><i data-lucide="arrow-left"></i>Estudar</button><h2 style="margin-top:11px">Transforma slides em estudo.</h2><p>Apontamentos, resumo, quiz e flashcards gerados localmente no browser.</p></div><div class="page-actions">' + modelWarning + '<button class="button" type="button" data-action="ai-pick-pptx"><i data-lucide="file-plus-2"></i>Importar .pptx</button></div></div><div class="bento-grid ai-page"><article class="card card-violet span-12 ai-hero"><div><p class="card-label">Twenty AI · offline depois do download</p><h2>Dos slides para o modo estudo.</h2><p>' + esc(supportCopy) + '</p></div><div class="ai-hero-model"><span>Recomendado neste dispositivo</span><strong>' + esc(recommendation.label) + '</strong><small>' + esc(recommendation.size || "") + '</small></div></article>' + renderAIProgress() + '<article class="card span-7 ai-generator-card"><div class="card-title-row"><div><p class="card-label">1 · Importar</p><h3>PowerPoint</h3><p class="card-subtitle">A Twenty extrai o texto e envia o .pptx para a pasta data/files do repositório privado.</p></div><span class="metric-icon"><i data-lucide="presentation"></i></span></div>' + draftHtml + '<form id="aiGeneratorForm" class="form-grid ai-generator-form"><div class="field"><label>Cadeira</label><select name="courseId"><option value="">Sem cadeira específica</option>' + courseOptions(courseValue) + '</select></div><div class="field"><label>Modelo</label><select name="modelMode"><option value="auto" ' + (modelValue === "auto" ? 'selected' : '') + '>Automático · recomendado</option><option value="fast" ' + (modelValue === "fast" ? 'selected' : '') + '>Rápido · Qwen 0.5B</option><option value="quality" ' + (modelValue === "quality" ? 'selected' : '') + '>Qualidade · Qwen 1.5B</option></select></div><div class="field"><label>O que criar</label><select name="output"><option value="all" ' + (outputValue === "all" ? 'selected' : '') + '>Tudo</option><option value="notes" ' + (outputValue === "notes" ? 'selected' : '') + '>Apontamentos completos</option><option value="summary" ' + (outputValue === "summary" ? 'selected' : '') + '>Resumo rápido</option><option value="quiz" ' + (outputValue === "quiz" ? 'selected' : '') + '>Quiz</option><option value="flashcards" ' + (outputValue === "flashcards" ? 'selected' : '') + '>Flashcards</option></select></div><div class="field"><label>Dificuldade</label><select name="difficulty"><option value="auto" ' + (difficultyValue === "auto" ? 'selected' : '') + '>Automática</option><option value="easy" ' + (difficultyValue === "easy" ? 'selected' : '') + '>Fácil</option><option value="medium" ' + (difficultyValue === "medium" ? 'selected' : '') + '>Média</option><option value="hard" ' + (difficultyValue === "hard" ? 'selected' : '') + '>Difícil</option></select></div><div class="field field-full"><label>Número de perguntas / flashcards</label><div class="ai-range-row"><input name="questionCount" type="range" min="5" max="30" step="5" value="' + questionValue + '" data-role="ai-question-range"><output id="aiQuestionCountOutput">' + questionValue + '</output></div></div></form><button class="button button-dark ai-generate-button" type="button" data-action="ai-generate" ' + (!aiDraft || aiBusy || !supported ? 'disabled' : '') + '><i data-lucide="sparkles"></i>' + (aiBusy ? 'A gerar…' : 'Gerar material de estudo') + '</button><p class="form-note"><strong>Primeira utilização:</strong> o modelo pode ocupar centenas de MB. Depois fica em cache neste dispositivo. Slides compostos apenas por imagens ainda não têm OCR.</p></article><article class="card span-5 ai-device-card"><div class="card-title-row"><div><p class="card-label">Como funciona</p><h3>Privado por defeito</h3></div><span class="metric-icon"><i data-lucide="shield-check"></i></span></div><div class="ai-steps"><div><span>1</span><p><strong>Extrai texto</strong><small>JSZip abre o .pptx no browser.</small></p></div><div><span>2</span><p><strong>Gera localmente</strong><small>WebLLM usa a GPU com WebGPU.</small></p></div><div><span>3</span><p><strong>Guarda na Twenty</strong><small>PowerPoint, apontamentos e quiz sincronizam com o Git; o modelo continua local.</small></p></div></div><div class="form-note"><strong>Modelo recomendado:</strong> ' + esc(recommendation.label) + ' (' + esc(recommendation.size || "tamanho variável") + '). Podes escolher manualmente antes de gerar.</div></article><section class="span-12 section-block"><div class="section-heading"><div><h3>Apresentações sincronizadas</h3><p>Os PowerPoints aparecem em todos os teus dispositivos; os resultados da IA ficam associados ao mesmo ficheiro.</p></div><span class="badge badge-violet">' + projects.length + '</span></div><div class="bento-grid">' + projectHtml + '</div></section></div>';
  }

  async function handleAIPptxFile(file) {
    if (!file || !AI) return;
    if (!Sync || !Sync.getStatus().configured) {
      toast("Configura primeiro o Git em Admin & dados para sincronizar o PowerPoint.", "warning");
      if (pptxInput) pptxInput.value = "";
      return;
    }
    if (!navigator.onLine) {
      toast("Precisas de Internet para enviar o PowerPoint para o repositório privado.", "warning");
      if (pptxInput) pptxInput.value = "";
      return;
    }
    aiBusy = true;
    aiDraft = null;
    var projectId = uid("aiproject");
    setAIProgress("A abrir o PowerPoint…", "A preparar o leitor de slides.", 2);
    if (route.name !== "study" || route.tab !== "ai") setRoute("study", null, "ai");
    else render();
    try {
      var extracted = await AI.extractPptx(file, function (report) {
        var progress = report.progress == null ? null : Math.min(30, Number(report.progress));
        setAIProgress("A extrair os slides…", report.text || "A ler o PowerPoint.", progress);
      });
      setAIProgress("A enviar o PowerPoint…", "A iniciar o upload seguro para o repositório privado.", 34);
      var remoteFile = await Sync.uploadFile(file, {
        id: projectId,
        name: file.name,
        onProgress: function (report) {
          var progress = report.progress == null ? null : 34 + Math.round(report.progress * 0.52);
          var detail = report.total ? formatBytes(report.loaded) + " de " + formatBytes(report.total) + " enviados" : "A enviar o ficheiro…";
          setAIProgress("A enviar o PowerPoint…", detail, progress);
        },
        onUploadComplete: function () {
          setAIProgress("A confirmar no GitHub…", "O upload terminou. A aguardar a criação do commit do ficheiro.", 88);
        },
        onReady: function (request) { aiTransferRequest = request; }
      });
      aiTransferRequest = null;
      var project = {
        id: projectId,
        semesterId: state.currentSemesterId,
        courseId: null,
        quizId: null,
        title: String(extracted.fileName || file.name).replace(/\.pptx$/i, ""),
        fileName: extracted.fileName || file.name,
        fileSize: extracted.fileSize || file.size,
        slideCount: extracted.slideCount,
        slides: extracted.slides,
        remoteFile: remoteFile,
        summary: "",
        notes: null,
        flashcards: [],
        quizQuestions: [],
        output: "pending",
        difficulty: "auto",
        questionCount: 10,
        modelMode: "",
        modelId: "",
        status: "uploaded",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      state.aiProjects.unshift(project);
      aiDraft = project;
      setAIProgress("A guardar na Twenty…", "A sincronizar os metadados e o texto dos slides para os outros dispositivos.", 94);
      await save(true);
      try { await Sync.syncNow(state, defaultState()); } catch (syncError) { console.warn("PPT metadata sync queued:", syncError); }
      aiDraft = aiProjectById(projectId) || project;
      aiBusy = false;
      setAIProgress("PowerPoint sincronizado", "O ficheiro já pode aparecer nos teus outros dispositivos.", 100);
      setTimeout(function () { clearAIProgress(); render(); }, 520);
      render();
      toast(extracted.slideCount + " slides enviados e sincronizados.");
    } catch (error) {
      aiBusy = false;
      clearAIProgress();
      render();
      toast(error.message || "Não foi possível enviar o PowerPoint.", "error");
    } finally {
      aiTransferRequest = null;
      if (pptxInput) pptxInput.value = "";
    }
  }

  async function generateAIProject() {
    if (!AI || !aiDraft || aiBusy) return;
    var form = document.getElementById("aiGeneratorForm");
    if (!form) return;
    var sourceProject = aiProjectById(aiDraft.id) || aiDraft;
    var courseId = form.elements.courseId.value || null;
    var options = {
      modelMode: form.elements.modelMode.value || "auto",
      output: form.elements.output.value || "all",
      difficulty: form.elements.difficulty.value || "auto",
      questionCount: Number(form.elements.questionCount.value) || 10
    };
    state.settings.aiModelMode = options.modelMode;
    state.settings.aiOutput = options.output;
    state.settings.aiDifficulty = options.difficulty;
    state.settings.aiQuestionCount = options.questionCount;
    aiBusy = true;
    setAIProgress("A preparar a IA local…", "Na primeira utilização, o modelo é descarregado para este dispositivo.", 31);
    render();
    try {
      var project = await AI.generateStudyPack(sourceProject, options, function (report) {
        var title = report.kind === "model" ? "A preparar o modelo…" : report.kind === "warning" ? "Modo de compatibilidade" : report.kind === "done" ? "A terminar…" : "A criar o teu material…";
        setAIProgress(title, report.text || "A processar os slides.", report.progress);
      });
      project.id = sourceProject.id;
      project.semesterId = state.currentSemesterId;
      project.courseId = courseId;
      project.remoteFile = sourceProject.remoteFile || null;
      project.createdAt = sourceProject.createdAt || project.createdAt;
      project.updatedAt = new Date().toISOString();
      project.status = "ready";
      if (sourceProject.quizId) state.quizzes = state.quizzes.filter(function (quiz) { return quiz.id !== sourceProject.quizId; });
      if (asArray(project.quizQuestions).length) {
        var quiz = { id: uid("quiz"), semesterId: state.currentSemesterId, courseId: courseId, lessonId: null, title: "IA · " + (project.title || project.fileName || "Slides"), questions: project.quizQuestions, generatedByAI: true, aiProjectId: project.id, createdAt: new Date().toISOString(), lastScore: null };
        state.quizzes.push(quiz);
        project.quizId = quiz.id;
      }
      var index = state.aiProjects.findIndex(function (item) { return item.id === project.id; });
      if (index >= 0) state.aiProjects[index] = project;
      else state.aiProjects.unshift(project);
      setAIProgress("A sincronizar o resultado…", "A enviar apontamentos, flashcards e quiz para o Git.", 96);
      await save(true);
      try {
        await Sync.syncNow(state, defaultState());
        setAIProgress("Material sincronizado", "O PowerPoint e o quiz já estão disponíveis nos outros dispositivos.", 100);
      } catch (syncError) {
        console.warn("AI project sync queued:", syncError);
        setAIProgress("Material guardado", "Ficou na fila e será enviado automaticamente quando a ligação estabilizar.", 100);
      }
      aiDraft = null;
      aiBusy = false;
      setTimeout(function () { clearAIProgress(); render(); }, 520);
      render();
      toast("Material criado e sincronizado na Twenty.");
      openAIProject(project.id);
    } catch (error) {
      aiBusy = false;
      clearAIProgress();
      render();
      if (String(error.message || "").toLowerCase().indexOf("cancel") < 0) toast(error.message || "A IA local não conseguiu terminar.", "error");
    }
  }

  function openAIProject(id) {
    var project = aiProjectById(id);
    if (!project) return;
    var course = courseById(project.courseId);
    var notes = project.notes || null;
    var hasFile = !!(project.remoteFile && project.remoteFile.path);
    var notesHtml = notes ? '<section class="ai-result-section"><div class="section-heading"><div><p class="card-label">Apontamentos</p><h3>' + esc(notes.title || project.title) + '</h3></div></div>' + (notes.overview ? '<p class="ai-overview">' + nl2br(notes.overview) + '</p>' : '') + asArray(notes.sections).map(function (section) { return '<article class="ai-note-section"><h4>' + esc(section.heading) + '</h4><p>' + nl2br(section.content) + '</p>' + aiSourceButtons(project, section.sourceSlides) + '</article>'; }).join("") + (asArray(notes.keyTakeaways).length ? '<div class="ai-takeaways"><p class="card-label">O essencial</p><ul>' + notes.keyTakeaways.map(function (item) { return '<li>' + esc(item) + '</li>'; }).join("") + '</ul></div>' : '') + '</section>' : (project.summary ? '<section class="ai-result-section"><p class="card-label">Resumo</p><p class="ai-overview">' + nl2br(project.summary) + '</p></section>' : '<section class="ai-result-section ai-awaiting-generation"><span class="metric-icon"><i data-lucide="sparkles"></i></span><div><p class="card-label">Pronto para gerar</p><h3>O PowerPoint já está sincronizado</h3><p class="card-subtitle">Podes gerar os apontamentos e o quiz neste dispositivo ou noutro com WebGPU.</p></div><button class="button button-dark" type="button" data-action="ai-use-project" data-id="' + attr(project.id) + '"><i data-lucide="sparkles"></i>Gerar material</button></section>');
    var fileHtml = '<section class="ai-result-section ai-file-sync-section"><span class="ai-file-icon"><i data-lucide="presentation"></i></span><div><p class="card-label">PowerPoint original</p><h3>' + esc(project.fileName || "Apresentação") + '</h3><p class="card-subtitle">' + formatBytes(project.fileSize || project.remoteFile && project.remoteFile.size || 0) + ' · ' + Number(project.slideCount || 0) + ' slides · ' + (hasFile ? 'guardado em data/files' : 'apenas local') + '</p></div>' + (hasFile ? '<button class="button" type="button" data-action="ai-download-pptx" data-id="' + attr(project.id) + '"><i data-lucide="download"></i>Descarregar</button>' : '') + '</section>';
    var cardsHtml = asArray(project.flashcards).length ? '<section class="ai-result-section"><div class="section-heading"><div><p class="card-label">Flashcards</p><h3>' + project.flashcards.length + ' cartões</h3></div></div><div class="ai-flashcard-grid">' + project.flashcards.map(function (card, index) { return '<details class="ai-flashcard"><summary><span>' + (index + 1) + '</span>' + esc(card.front) + '</summary><div><p>' + nl2br(card.back) + '</p>' + aiSourceButtons(project, card.sourceSlides) + '</div></details>'; }).join("") + '</div></section>' : '';
    var quizHtml = asArray(project.quizQuestions).length ? '<section class="ai-result-section ai-quiz-summary"><div><p class="card-label">Quiz</p><h3>' + project.quizQuestions.length + ' perguntas prontas</h3><p class="card-subtitle">Dificuldade ' + esc(aiDifficultyLabel(project.difficulty).toLowerCase()) + ' · resposta e explicação incluídas.</p></div><button class="button button-dark" type="button" data-action="ai-start-quiz" data-id="' + attr(project.id) + '"><i data-lucide="play"></i>Começar quiz</button></section>' : '';
    var warning = project.warning ? '<div class="form-note"><strong>Compatibilidade:</strong> ' + esc(project.warning) + '</div>' : '';
    var body = '<div class="ai-project-modal-head"><div><span class="badge ' + (hasFile ? 'badge-mint' : 'badge-yellow') + '">' + (hasFile ? 'PPT sincronizado' : 'PPT local') + '</span><h2>' + esc(project.title) + '</h2><p>' + esc(course ? course.name : "Sem cadeira") + ' · ' + Number(project.slideCount || 0) + ' slides' + (project.modelMode ? ' · ' + esc(project.modelMode === "quality" ? "modelo de qualidade" : "modelo rápido") : '') + '</p></div><span class="ai-file-icon"><i data-lucide="' + (notes || project.summary ? 'brain' : 'presentation') + '"></i></span></div>' + fileHtml + warning + notesHtml + cardsHtml + quizHtml;
    var footer = '<footer class="modal-foot"><button class="button button-danger" type="button" data-action="ai-delete-project" data-id="' + attr(project.id) + '"><i data-lucide="trash-2"></i>Apagar</button><button class="button" type="button" data-action="close-modal">Fechar</button></footer>';
    openModal("Projeto de IA", body, { className: "modal-wide ai-project-modal", footer: footer });
  }

  function useAIProject(id) {
    var project = aiProjectById(id);
    if (!project) return;
    aiDraft = project;
    closeModal();
    setRoute("study", null, "ai");
    render();
    toast("PowerPoint selecionado. Escolhe o que queres gerar.");
  }

  async function downloadAIProjectFile(id) {
    var project = aiProjectById(id);
    if (!project || !project.remoteFile || !project.remoteFile.path || !Sync) { toast("Este PowerPoint não está disponível no servidor.", "warning"); return; }
    closeModal();
    setAIProgress("A descarregar o PowerPoint…", "A preparar o ficheiro guardado no repositório privado.", 4);
    render();
    try {
      var blob = await Sync.downloadFile(project.remoteFile, { onProgress: function (report) { var progress = report.progress == null ? null : 5 + Math.round(report.progress * 0.9); var detail = report.total ? formatBytes(report.loaded) + " de " + formatBytes(report.total) + " descarregados" : "A receber o ficheiro…"; setAIProgress("A descarregar o PowerPoint…", detail, progress); }, onReady: function (request) { aiTransferRequest = request; } });
      aiTransferRequest = null;
      setAIProgress("PowerPoint pronto", "A abrir o download neste dispositivo.", 100);
      var url = URL.createObjectURL(blob);
      var link = document.createElement("a");
      link.href = url;
      link.download = project.fileName || project.remoteFile.name || "apresentacao.pptx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(function () { URL.revokeObjectURL(url); clearAIProgress(); render(); }, 800);
    } catch (error) {
      aiTransferRequest = null;
      clearAIProgress();
      render();
      toast(error.message || "Não foi possível descarregar o PowerPoint.", "error");
    }
  }

  function openAISlide(projectId, number) {
    var project = aiProjectById(projectId);
    if (!project) return;
    var slide = asArray(project.slides).find(function (item) { return Number(item.number) === Number(number); });
    if (!slide) { toast("O texto deste slide não está disponível.", "warning"); return; }
    var body = '<div class="ai-slide-preview"><span class="badge badge-violet">Slide ' + Number(slide.number) + '</span><h3>' + esc(slide.title || "Slide " + slide.number) + '</h3><div class="form-note">' + (slide.text ? nl2br(slide.text) : "Este slide não tinha texto selecionável.") + '</div></div>';
    var footer = '<footer class="modal-foot"><button class="button" type="button" data-action="ai-open-project" data-id="' + attr(project.id) + '"><i data-lucide="arrow-left"></i>Voltar ao projeto</button></footer>';
    openModal("Fonte dos apontamentos", body, { footer: footer });
  }

  function confirmDeleteAIProject(id) {
    var project = aiProjectById(id);
    if (!project) return;
    var fileCopy = project.remoteFile && project.remoteFile.path ? " O PowerPoint sincronizado também será apagado de data/files." : "";
    openModal("Apagar projeto de IA?", '<p class="onboarding-copy" style="margin-top:0">Vais apagar os apontamentos, flashcards e o quiz de <strong>' + esc(project.title) + '</strong>.' + fileCopy + '</p>', { footer: '<footer class="modal-foot"><button class="button" type="button" data-action="ai-open-project" data-id="' + attr(project.id) + '">Cancelar</button><button class="button button-danger" type="button" data-action="confirm-ai-delete-project" data-id="' + attr(project.id) + '"><i data-lucide="trash-2"></i>Apagar</button></footer>' });
  }

  async function deleteAIProject(id) {
    var project = aiProjectById(id);
    if (!project) return;
    closeModal();
    setAIProgress("A apagar o projeto…", project.remoteFile && project.remoteFile.path ? "A remover o PowerPoint do repositório privado." : "A remover os dados da Twenty.", 18);
    render();
    try {
      if (project.remoteFile && project.remoteFile.path && Sync) { await Sync.deleteFile(project.remoteFile); setAIProgress("A atualizar a data…", "A remover apontamentos e quiz do estado sincronizado.", 72); }
      state.aiProjects = state.aiProjects.filter(function (item) { return item.id !== id; });
      if (project.quizId) state.quizzes = state.quizzes.filter(function (quiz) { return quiz.id !== project.quizId; });
      if (aiDraft && aiDraft.id === id) aiDraft = null;
      await save(true);
      if (Sync && Sync.getStatus().configured) { try { await Sync.syncNow(state, defaultState()); } catch (_) {} }
      setAIProgress("Projeto apagado", "A alteração já foi guardada.", 100);
      setTimeout(function () { clearAIProgress(); render(); }, 450);
      render();
      toast("Projeto de IA e PowerPoint apagados.");
    } catch (error) {
      clearAIProgress();
      render();
      toast(error.message || "Não foi possível apagar o projeto.", "error");
    }
  }

  function renderStudy() {
    if (route.tab === "weekly") return renderWeeklyReview();
    setHeader("Estudar", "Quizzes e revisões");
    var quizzes = semesterItems("quizzes");
    var questions = semesterItems("questions");
    var upcoming = semesterItems("assessments").filter(function (item) { return !item.date || item.date >= todayISO(); }).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    var weakLessons = semesterItems("lessons").filter(function (item) { return !item.mastered; }).sort(function (a, b) {
      var aq = state.questions.filter(function (q) { return asArray(q.lessonIds).indexOf(a.id) >= 0; }).length;
      var bq = state.questions.filter(function (q) { return asArray(q.lessonIds).indexOf(b.id) >= 0; }).length;
      return bq - aq;
    });
    var featured = upcoming[0] || null;
    var featuredCourse = featured ? courseById(featured.courseId) : null;
    var scopedLessons = featured ? asArray(featured.lessonIds).map(lessonById).filter(Boolean) : [];
    var relatedQuestions = featured ? questions.filter(function (question) {
      return asArray(question.lessonIds).some(function (id) { return asArray(featured.lessonIds).indexOf(id) >= 0; });
    }) : questions;
    var quizCards = quizzes.length ? quizzes.map(function (quiz) {
      var course = courseById(quiz.courseId);
      var lesson = lessonById(quiz.lessonId);
      return '<article class="card span-4"><div class="card-title-row"><div><span class="badge badge-violet">' + esc(course ? course.code || course.name : "Quiz") + '</span><h3 style="margin-top:12px">' + esc(quiz.title) + '</h3><p class="card-subtitle">' + asArray(quiz.questions).length + ' perguntas · ' + esc(lesson ? lesson.title : "Revisão geral") + '</p></div><span class="metric-icon"><i data-lucide="brain"></i></span></div><button class="button button-dark" style="margin-top:20px" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="play"></i>Começar quiz</button></article>';
    }).join("") : '<div class="span-12">' + emptyState("sparkles", "Ainda não criaste quizzes", "Usa o lado admin para importar manualmente as perguntas e associá-las às aulas.", "add-quiz", "Criar quiz") + "</div>";
    var weakHtml = weakLessons.length ? weakLessons.slice(0, 5).map(function (lesson) {
      var course = courseById(lesson.courseId);
      var count = state.questions.filter(function (q) { return asArray(q.lessonIds).indexOf(lesson.id) >= 0; }).length;
      return '<div class="list-row"><span class="list-icon yellow"><i data-lucide="rotate-ccw"></i></span><span class="list-content"><strong>' + esc(lesson.title) + '</strong><small>' + esc(course ? course.name : "Cadeira") + ' · ' + count + ' perguntas anteriores</small></span><button class="row-button" type="button" data-route="lesson" data-id="' + attr(lesson.id) + '"><i data-lucide="arrow-right"></i></button></div>';
    }).join("") : emptyState("badge-check", "Sem revisões pendentes", "Todas as aulas estão marcadas como dominadas.", null);
    var questionHtml = relatedQuestions.length ? relatedQuestions.slice(0, 6).map(function (item) { return renderQuestionCard(item, false); }).join("") : emptyState("message-circle-question", "Sem perguntas para esta matéria", "Adiciona perguntas antigas e associa-as às aulas que saem na avaliação.", "add-question", "Adicionar pergunta");

    var hero = featured ? '<article class="card card-violet span-12"><div class="card-title-row"><div><span class="badge badge-dark">Próxima avaliação · ' + relativeDate(featured.date) + '</span><h2 style="margin:18px 0 8px;font-size:clamp(1.8rem,4vw,3.5rem);letter-spacing:-.07em">' + esc(featured.title) + '</h2><p class="card-subtitle" style="color:rgba(24,25,31,.7)">' + esc(featuredCourse ? featuredCourse.name : "Avaliação") + ' · ' + scopedLessons.length + ' aulas na matéria · ' + relatedQuestions.length + ' perguntas anteriores</p></div><span class="hero-number" style="position:relative;right:auto;bottom:auto;font-size:7rem;color:rgba(255,255,255,.55)">20</span></div><div class="hero-actions"><button class="button button-dark" type="button" data-action="study-assessment" data-id="' + attr(featured.id) + '"><i data-lucide="play"></i>Estudar esta matéria</button><button class="button" type="button" data-action="assessment-scope" data-id="' + attr(featured.id) + '"><i data-lucide="list-tree"></i>Ver aulas incluídas</button></div></article>' : '<article class="card card-violet span-12"><div class="page-head"><div><h2>Sessão livre</h2><p>Sem avaliação marcada. Escolhe uma cadeira ou uma aula por rever.</p></div><button class="button button-dark" type="button" data-action="add-assessment">Marcar avaliação</button></div></article>';
    return '<div class="page-head"><div><h2>Estudar com contexto.</h2><p>Matéria, slides, perguntas anteriores e quizzes no mesmo fluxo.</p></div><div class="page-actions"><button class="button" type="button" data-route="study" data-tab="weekly"><i data-lucide="clipboard-check"></i>Revisão semanal</button><button class="button" type="button" data-route="planner" data-planner-view="study-day"><i data-lucide="blocks"></i>Planear dia</button><button class="button" type="button" data-action="add-question"><i data-lucide="plus"></i>Pergunta antiga</button><button class="button button-dark" type="button" data-action="add-quiz"><i data-lucide="sparkles"></i>Novo quiz</button></div></div><div class="bento-grid">' + hero + '<article class="card span-5"><div class="card-title-row"><div><h3>Aulas por rever</h3></div><span class="badge badge-yellow">' + weakLessons.length + '</span></div><div class="list-stack">' + weakHtml + '</div></article><article class="card span-7"><div class="card-title-row"><div><p class="card-label">Perguntas de testes anteriores</p><h3>' + (featured ? "Ligadas à próxima avaliação" : "Banco geral") + '</h3></div><span class="badge badge-pink">' + relatedQuestions.length + '</span></div><div style="margin-top:14px">' + questionHtml + '</div></article>' + renderStudyHourEstimate() + '<section class="span-12 section-block"><div class="section-heading"><div><h3>Quizzes disponíveis</h3><p>Quizzes manuais e perguntas anteriores ligadas às aulas.</p></div></div><div class="bento-grid">' + quizCards + "</div></section></div>";
  }

  function gradeSimulatorAssessmentFields(courseId) {
    var course = courseById(courseId);
    if (!course) return "";
    var gradedIds = state.grades.map(function (grade) { return grade.assessmentId; }).filter(Boolean);
    var pending = semesterItems("assessments").filter(function (assessment) { return assessment.courseId === course.id && gradedIds.indexOf(assessment.id) < 0; }).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); });
    if (!pending.length) return '<div class="past-question-empty"><i data-lucide="check-check"></i><span>Não existem avaliações sem nota nesta cadeira.</span></div>';
    return '<div class="simulator-assessments">' + pending.map(function (assessment) {
      var component = asArray(course.evaluation && course.evaluation.components).find(function (item) { return item.id === assessment.componentId; });
      return '<label class="simulator-score-row"><span><strong>' + esc(assessment.title) + '</strong><small>' + esc((component ? component.label : assessment.type || "Avaliação") + (assessment.date ? " · " + formatDate(assessment.date) : "")) + '</small></span><span><input name="simScore" data-assessment-id="' + attr(assessment.id) + '" type="number" min="0" max="20" step="0.1" placeholder="—"><b>/20</b></span></label>';
    }).join("") + '</div>';
  }

  function updateGradeSimulator(form) {
    if (!form) return;
    var course = courseById(form.elements.courseId && form.elements.courseId.value);
    var target = form.querySelector("#gradeSimulatorResult");
    if (!course || !target) return;
    var simulated = Array.from(form.querySelectorAll('[name="simScore"]')).filter(function (input) { return input.value !== ""; }).map(function (input) {
      var assessment = assessmentById(input.dataset.assessmentId);
      return { id: "sim_" + input.dataset.assessmentId, semesterId: course.semesterId, courseId: course.id, assessmentId: assessment.id, componentId: assessment.componentId, score: clamp(input.value, 0, 20), defenseStatus: "not-applicable", defenseType: "", defenseFinalScore: null };
    });
    var result = courseAverage(course, state.grades.concat(simulated));
    var current = courseAverage(course).value;
    var components = result.components.map(function (entry) {
      return '<span><strong>' + esc(entry.component.label) + '</strong><b>' + (entry.effective == null ? "—" : round(entry.effective, 1)) + '</b><small>' + (Number(entry.component.weight) || 0) + '%</small></span>';
    }).join("");
    var notes = [];
    if (result.minimumFailures.length) notes.push(result.minimumFailures.length + " mínimo(s) não atingido(s)");
    if (result.defensePending.length) notes.push(result.defensePending.length + " defesa(s) pendente(s)");
    target.innerHTML = '<div><p class="card-label">Média projetada</p><strong class="simulator-result-number">' + (result.value == null ? "—" : round(result.value, 2)) + '</strong><span>/20</span></div><p>' + (simulated.length ? "Com " + simulated.length + " nota(s) simulada(s)" : "Introduz uma nota para ver a projeção") + (current == null ? "" : " · atual " + round(current, 2) + "/20") + '</p><div class="simulator-component-grid">' + components + '</div>' + (notes.length ? '<div class="form-note">' + esc(notes.join(" · ")) + '</div>' : '');
    refreshIcons(target);
  }

  function openGradeSimulator(courseId) {
    var selected = courseId || (activeCourses()[0] && activeCourses()[0].id) || "";
    var body = '<form id="gradeSimulatorForm"><div class="field"><label>Cadeira</label><select name="courseId" data-role="simulator-course"><option value="">Escolher…</option>' + courseOptions(selected) + '</select></div><div id="gradeSimulatorFields" style="margin-top:15px">' + gradeSimulatorAssessmentFields(selected) + '</div><article id="gradeSimulatorResult" class="grade-simulator-result"></article><p class="form-note">Esta simulação não guarda notas. Usa o método de avaliação, mínimos, defesas e substituições já configurados.</p></form>';
    openModal("Simular próximas notas", body, { footer: '<footer class="modal-foot"><button class="button button-dark" type="button" data-action="close-modal">Fechar</button></footer>' });
    updateGradeSimulator(modalRoot.querySelector("#gradeSimulatorForm"));
  }

  function renderGrades() {
    setHeader("Notas", "Média e progresso");
    var ects = ectsAverage();
    var courses = activeCourses();
    var totalEcts = courses.reduce(function (sum, course) { return sum + (Number(course.ects) || 0); }, 0);
    var courseRows = courses.map(function (course) {
      var avg = courseAverage(course);
      return '<tr><td><span class="badge" style="background:' + safeColor(course.color) + '">' + esc(course.code || "Cadeira") + '</span> <strong style="margin-left:6px">' + esc(course.name) + '</strong></td><td>' + (Number(course.ects) || 0) + '</td><td>' + avg.knownWeight + '%</td><td><span class="grade-number">' + (avg.value == null ? "—" : round(avg.value, 1)) + '</span>' + (avg.value == null ? "" : "/20") + '</td><td><button class="row-button" type="button" data-route="course" data-id="' + attr(course.id) + '" data-tab="grades"><i data-lucide="arrow-right"></i></button></td></tr>';
    }).join("");
    var best = courses.map(function (course) { return { course: course, avg: courseAverage(course).value }; }).filter(function (item) { return item.avg != null; }).sort(function (a, b) { return b.avg - a.avg; })[0];
    var known = courses.filter(function (course) { return courseAverage(course).value != null; }).length;
    return '<div class="page-head"><div><h2>Média e desempenho</h2><p>A média global é ponderada pelos ECTS. Cada cadeira segue o método de avaliação e as regras configuradas.</p></div><div class="page-actions"><button class="button" type="button" data-action="grade-simulator"><i data-lucide="calculator"></i>Simular notas</button><button class="button" type="button" onclick="window.print()"><i data-lucide="printer"></i>Imprimir</button><button class="button button-dark" type="button" data-action="add-grade"><i data-lucide="plus"></i>Adicionar nota</button></div></div><div class="bento-grid"><article class="card card-pink span-5 target-card"><div class="target-copy"><p class="card-label">Média ECTS atual</p><h3 style="font-size:3.4rem">' + (ects.value == null ? "—" : round(ects.value, 2)) + '</h3><p>' + (ects.value == null ? "Adiciona as primeiras notas para iniciar o cálculo." : "Calculada com " + ects.ects + " ECTS que já têm avaliação.") + '</p></div><div class="progress-ring" style="--progress:' + (ects.value == null ? 0 : clamp(ects.value / 20 * 100, 0, 100)) + '%"><strong>' + (ects.value == null ? "0" : Math.round(ects.value / 20 * 100)) + '%</strong></div></article><article class="card card-yellow span-3 metric-card"><div class="metric-top"><p class="card-label">ECTS inscritos</p><span class="metric-icon"><i data-lucide="graduation-cap"></i></span></div><div><p class="metric-value">' + totalEcts + '</p><p class="metric-caption">' + ects.ects + ' já entram na média</p></div></article><article class="card card-mint span-4 metric-card"><div class="metric-top"><p class="card-label">Melhor cadeira</p><span class="metric-icon"><i data-lucide="trophy"></i></span></div><div><p class="metric-value" style="font-size:2.35rem">' + (best ? round(best.avg, 1) : "—") + '</p><p class="metric-caption">' + (best ? esc(best.course.name) : "Ainda sem notas") + '</p></div></article><article class="card span-12"><div class="card-title-row"><div><p class="card-label">Resumo do semestre</p><h3>' + known + ' de ' + courses.length + ' cadeiras com notas</h3></div><span class="badge badge-violet">Meta ' + (Number(state.profile.targetGrade) || 20) + '/20</span></div><div style="overflow:auto;margin-top:13px">' + (courses.length ? '<table class="grade-table"><thead><tr><th>Cadeira</th><th>ECTS</th><th>Avaliado</th><th>Média</th><th></th></tr></thead><tbody>' + courseRows + "</tbody></table>" : emptyState("graduation-cap", "Sem cadeiras", "Configura o semestre para começar o cálculo.", "new-semester", "Criar semestre")) + '</div></article></div>';
  }

  function canteenPortugalParts(date) {
    var values = {};
    new Intl.DateTimeFormat("en-GB", {
      timeZone: "Europe/Lisbon",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hourCycle: "h23"
    }).formatToParts(date || new Date()).forEach(function (part) {
      if (part.type !== "literal") values[part.type] = part.value;
    });
    var iso = values.year + "-" + values.month + "-" + values.day;
    return {
      iso: iso,
      year: Number(values.year),
      month: Number(values.month),
      day: Number(values.day),
      weekday: new Date(iso + "T12:00:00Z").getUTCDay(),
      minutes: Number(values.hour) * 60 + Number(values.minute)
    };
  }

  function addISODays(value, amount) {
    var date = new Date(String(value) + "T12:00:00Z");
    date.setUTCDate(date.getUTCDate() + Number(amount || 0));
    return date.toISOString().slice(0, 10);
  }

  function isBusinessDayISO(value) {
    var weekday = new Date(String(value) + "T12:00:00Z").getUTCDay();
    return weekday >= 1 && weekday <= 5;
  }

  function lastBusinessDayISO(year, month) {
    var date = new Date(Date.UTC(year, month, 0, 12));
    while (date.getUTCDay() === 0 || date.getUTCDay() === 6) date.setUTCDate(date.getUTCDate() - 1);
    return date.toISOString().slice(0, 10);
  }

  function nthBusinessDayISO(year, month, target) {
    var date = new Date(Date.UTC(year, month - 1, 1, 12));
    var count = 0;
    while (count < target) {
      if (date.getUTCDay() >= 1 && date.getUTCDay() <= 5) count += 1;
      if (count < target) date.setUTCDate(date.getUTCDate() + 1);
    }
    return date.toISOString().slice(0, 10);
  }

  function canteenSummerClosure(value) {
    var year = Number(String(value).slice(0, 4));
    var start = lastBusinessDayISO(year, 7);
    var reopen = nthBusinessDayISO(year, 9, 2);
    return value >= start && value < reopen ? { start: start, reopen: reopen } : null;
  }

  function nextCanteenOpeningDate(value) {
    for (var offset = 1; offset <= 80; offset += 1) {
      var candidate = addISODays(value, offset);
      if (isBusinessDayISO(candidate) && !canteenSummerClosure(candidate)) return candidate;
    }
    return "";
  }

  function canteenDateLabel(value, reference) {
    if (value === addISODays(reference, 1)) return "amanhã";
    var date = new Date(String(value) + "T12:00:00Z");
    return new Intl.DateTimeFormat("pt-PT", { timeZone: "UTC", weekday: "long", day: "numeric", month: "long" }).format(date);
  }

  function canteenServiceHours(menu) {
    var text = asArray(menu && menu.hours).join(" ");
    function period(name, fallbackStart, fallbackEnd) {
      var match = text.match(new RegExp(name + "\\s+das\\s+(\\d{1,2}:\\d{2})\\s*h?\\s*às\\s*(\\d{1,2}:\\d{2})", "i"));
      return {
        start: match ? match[1] : fallbackStart,
        end: match ? match[2] : fallbackEnd
      };
    }
    return {
      lunch: period("Almoço", "11:30", "14:30"),
      dinner: period("Jantar", "18:30", "20:30"),
      weekendClosed: !text || /encerrad[ao]\s+ao\s+fim\s+de\s+semana/i.test(text)
    };
  }

  function canteenOpeningStatus(date, hours) {
    var now = canteenPortugalParts(date || new Date());
    hours = hours || canteenServiceHours(null);
    var lunchStart = timeMinutes(hours.lunch.start);
    var lunchEnd = timeMinutes(hours.lunch.end);
    var dinnerStart = timeMinutes(hours.dinner.start);
    var dinnerEnd = timeMinutes(hours.dinner.end);
    var summer = canteenSummerClosure(now.iso);
    if (summer) {
      return {
        open: false,
        icon: "calendar-off",
        title: "Fechada agora",
        detail: "Férias de verão · reabre " + canteenDateLabel(summer.reopen, now.iso) + " às " + hours.lunch.start
      };
    }
    if (now.weekday >= 1 && now.weekday <= 5) {
      if (now.minutes >= lunchStart && now.minutes < lunchEnd) return { open: true, icon: "door-open", title: "Aberta agora", detail: "Almoço até às " + hours.lunch.end };
      if (now.minutes >= dinnerStart && now.minutes < dinnerEnd) return { open: true, icon: "door-open", title: "Aberta agora", detail: "Jantar até às " + hours.dinner.end };
      if (now.minutes < lunchStart) return { open: false, icon: "door-closed", title: "Fechada agora", detail: "Abre às " + hours.lunch.start + " para almoço" };
      if (now.minutes < dinnerStart) return { open: false, icon: "door-closed", title: "Fechada agora", detail: "Reabre às " + hours.dinner.start + " para jantar" };
    }
    var next = nextCanteenOpeningDate(now.iso);
    return {
      open: false,
      icon: "door-closed",
      title: "Fechada agora",
      detail: next ? "Abre " + canteenDateLabel(next, now.iso) + " às " + hours.lunch.start : "Consulta a próxima abertura na SAS NOVA"
    };
  }

  function canteenMealIcon(name) {
    var value = String(name || "").toLowerCase();
    if (value.indexOf("jantar") >= 0) return "moon-star";
    return "sun";
  }

  function canteenMealTone(name) {
    var value = String(name || "").toLowerCase();
    if (value.indexOf("jantar") >= 0) return "dinner";
    return "lunch";
  }

  function canteenDishIcon(type) {
    var value = String(type || "").toLowerCase();
    if (value.indexOf("sopa") >= 0) return "soup";
    if (value.indexOf("veg") >= 0) return "leaf";
    return "utensils";
  }

  function canteenDayChip(day) {
    var date = localDate(day.date);
    var today = canteenPortugalParts(new Date()).iso;
    var label = day.date === today ? "Hoje" : date ? new Intl.DateTimeFormat("pt-PT", { weekday: "short" }).format(date).replace(".", "") : "Dia";
    var dateLabel = date ? new Intl.DateTimeFormat("pt-PT", { day: "2-digit", month: "short" }).format(date).replace(".", "") : day.label;
    return '<button class="canteen-day-chip ' + (day.date === canteenSelectedDate ? "is-active" : "") + '" type="button" data-action="canteen-day" data-date="' + attr(day.date) + '"><span>' + esc(label) + '</span><strong>' + esc(dateLabel) + '</strong></button>';
  }

  function renderCanteenMeal(meal, menu, options) {
    options = options || {};
    var tone = canteenMealTone(meal.name);
    var items = asArray(meal.items);
    var soups = items.filter(function (item) { return String(item.type || "").toLowerCase().indexOf("sopa") >= 0; });
    var mainDishes = items.filter(function (item) { return String(item.type || "").toLowerCase().indexOf("sopa") < 0; });
    function renderDish(item, showType) {
      var codes = asArray(item.allergens);
      var names = codes.map(function (id) { return menu.allergens && menu.allergens[id] ? menu.allergens[id] : "Alergénio " + id; });
      var allergenCodes = codes.length ? '<span class="canteen-allergen-codes" title="' + attr(names.join(", ")) + '" aria-label="Alergénios ' + attr(codes.join(", ")) + '"><i data-lucide="shield-alert"></i>' + esc(codes.join(", ")) + '</span>' : "";
      return '<div class="canteen-dish"><span class="canteen-dish-icon"><i data-lucide="' + canteenDishIcon(item.type) + '"></i></span><div class="canteen-dish-copy"><div>' + (showType ? '<span class="canteen-kind">' + esc(item.type || "Opção") + '</span>' : "") + allergenCodes + '</div><strong>' + esc(item.description || "Descrição indisponível") + '</strong></div>' + (item.kcal ? '<span class="canteen-kcal ' + attr(item.calorieBand || "") + '"><strong>' + Number(item.kcal) + '</strong><small>kcal</small></span>' : "") + '</div>';
    }
    var soupRows = soups.length ? soups.map(function (item) { return renderDish(item, false); }).join("") : '<p class="canteen-unavailable">Sem sopa indicada para este serviço.</p>';
    var dishRows = mainDishes.length ? mainDishes.map(function (item) { return renderDish(item, true); }).join("") : '<p class="canteen-unavailable">Sem pratos indicados para este serviço.</p>';
    var head = '<span class="canteen-meal-icon"><i data-lucide="' + canteenMealIcon(meal.name) + '"></i></span><h3>' + esc(meal.name) + '</h3><span class="canteen-option-count">' + mainDishes.length + (mainDishes.length === 1 ? ' prato' : ' pratos') + '</span>';
    var classes = "card canteen-meal canteen-" + tone;
    if (options.primary) classes += " canteen-primary-meal";
    if (options.nested) classes += " canteen-nested-meal";
    return '<article class="' + classes + '">' + (options.hideHead ? "" : '<div class="canteen-meal-head">' + head + '</div>') + '<div class="canteen-menu-groups"><section class="canteen-menu-group canteen-soup-group"><h4>Sopa</h4><div class="canteen-dishes">' + soupRows + '</div></section><section class="canteen-menu-group"><div class="canteen-group-head"><h4>Escolhe o prato</h4><span>' + mainDishes.length + (mainDishes.length === 1 ? ' opção' : ' opções') + '</span></div><div class="canteen-dishes">' + dishRows + '</div></section></div></article>';
  }

  function renderCanteenDinner(meal, menu, hours) {
    var mainCount = meal ? asArray(meal.items).filter(function (item) { return String(item.type || "").toLowerCase().indexOf("sopa") < 0; }).length : 0;
    var dinnerHours = hours && hours.dinner ? hours.dinner.start + "–" + hours.dinner.end : "18:30–20:30";
    var copy = meal ? mainCount + (mainCount === 1 ? " prato" : " pratos") + " · " + dinnerHours : "Sem ementa publicada";
    var content = meal ? renderCanteenMeal(meal, menu, { nested: true, hideHead: true }) : '<div class="canteen-dinner-empty"><i data-lucide="utensils-crossed"></i><p>O jantar não está indicado para este dia.</p></div>';
    return '<details class="canteen-dinner-disclosure"><summary><span class="canteen-dinner-icon"><i data-lucide="moon-star"></i></span><span><strong>Jantar</strong><small>' + esc(copy) + '</small></span><i class="canteen-dinner-chevron" data-lucide="chevron-down"></i></summary><div class="canteen-dinner-content">' + content + '</div></details>';
  }

  function renderCanteen() {
    setHeader("Cantina", "Campus · SAS NOVA");
    var refreshButton = '<button class="button" type="button" data-action="refresh-canteen" ' + (canteenStatus === "loading" ? "disabled" : "") + '><i data-lucide="refresh-cw"></i>' + (canteenStatus === "loading" ? "A atualizar…" : "Atualizar") + '</button>';
    var head = '<div class="page-head canteen-page-head"><div><h2>Cantina da FCT</h2><p>Ementa do refeitório.</p></div><div class="page-actions">' + refreshButton + '</div></div>';
    if (!canteenMenu && (canteenStatus === "idle" || canteenStatus === "loading")) {
      return head + '<section class="card canteen-loading"><span class="loading-orb"></span><div><h3>A carregar ementa</h3><p>A consultar a ementa oficial da SAS NOVA.</p></div></section>';
    }
    if (!canteenMenu) {
      return head + '<section class="card canteen-error"><span class="canteen-empty-icon"><i data-lucide="wifi-off"></i></span><div><h3>Não foi possível abrir a ementa.</h3><p>' + esc(canteenError || "A fonte oficial está temporariamente indisponível.") + '</p><button class="button button-dark" type="button" data-action="refresh-canteen"><i data-lucide="refresh-cw"></i>Tentar novamente</button></div></section>';
    }
    var now = canteenPortugalParts(new Date());
    var hours = canteenServiceHours(canteenMenu);
    var service = canteenOpeningStatus(new Date(), hours);
    var info = canteenMenu.info || {};
    var socialMeal = info.socialMeal || {};
    var closures = info.closures || {};
    var price = socialMeal.amount || "3,10 €";
    var effectiveFrom = socialMeal.effectiveFrom || "2 de março de 2026";
    var includes = socialMeal.includes || "um prato, sopa, pão, uma bebida e uma sobremesa";
    var days = asArray(canteenMenu.days).slice().sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
    if (!canteenSelectedDate || !days.some(function (day) { return day.date === canteenSelectedDate; })) {
      var bestDay = days.find(function (day) { return day.date === now.iso; }) || days.find(function (day) { return day.date >= now.iso; }) || days[days.length - 1];
      canteenSelectedDate = bestDay ? bestDay.date : "";
    }
    var selected = days.find(function (day) { return day.date === canteenSelectedDate; }) || days[0];
    var selectedDate = selected && localDate(selected.date);
    var longDate = selectedDate ? new Intl.DateTimeFormat("pt-PT", { weekday: "long", day: "numeric", month: "long" }).format(selectedDate) : selected.label;
    var consultedAt = canteenMenu.fetchedAt ? new Intl.DateTimeFormat("pt-PT", { timeZone: "Europe/Lisbon", day: "2-digit", month: "long", year: "numeric", hour: "2-digit", minute: "2-digit" }).format(new Date(canteenMenu.fetchedAt)) : "data indisponível";
    var allOfficial = canteenStatus === "ready" && canteenMenu.infoSource === "official";
    var statusClass = allOfficial ? "is-fresh" : "is-stale";
    var verificationTitle = allOfficial ? "Informação oficial SAS NOVA" : (canteenStatus === "ready" ? "Ementa oficial · condições guardadas" : "Última informação guardada");
    var verificationCopy = allOfficial ? "Ementa, preço e funcionamento consultados nas páginas oficiais." : (canteenStatus === "ready" ? "A ementa foi atualizada; o preço e as condições usam a última cópia disponível." : "A fonte oficial não respondeu. A informação apresentada pode estar desatualizada.");
    var verificationTime = (canteenStatus === "ready" ? "Consultada em " : "Última consulta em ") + consultedAt;
    var meals = selected ? asArray(selected.meals) : [];
    var lunch = meals.find(function (meal) { return String(meal.name || "").toLowerCase().indexOf("almoço") >= 0; });
    var dinner = meals.find(function (meal) { return String(meal.name || "").toLowerCase().indexOf("jantar") >= 0; });
    var lunchCard = lunch ? renderCanteenMeal(lunch, canteenMenu, { primary: true }) : '<article class="card canteen-no-menu"><i data-lucide="utensils-crossed"></i><div><h3>Sem almoço publicado</h3><p>Confirma a informação na SAS NOVA.</p></div></article>';
    var allergenEntries = Object.keys(canteenMenu.allergens || {}).sort(function (a, b) { return Number(a) - Number(b); }).map(function (id) { return '<span><strong>' + esc(id) + '</strong>' + esc(canteenMenu.allergens[id]) + '</span>'; }).join("");
    var summerCopy = closures.summer || "Férias de verão: encerramento no último dia útil de julho e reabertura no 2.º dia útil de setembro.";
    var seasonalCopy = closures.seasonal || "Existem ainda encerramentos curtos no Natal, Carnaval e Páscoa, comunicados por aviso.";
    var alternativesCopy = closures.alternatives || "Nesses períodos, algumas cantinas de outras universidades públicas podem permanecer abertas a alunos da NOVA.";
    var serviceCard = '<section class="canteen-service-status ' + (service.open ? "is-open" : "is-closed") + '"><span class="canteen-service-icon"><i data-lucide="' + service.icon + '"></i></span><div class="canteen-service-copy"><h3>' + esc(service.title) + '</h3><p>' + esc(service.detail) + '</p><small>Segundo o horário regular publicado.</small></div><div class="canteen-price"><span>Refeição social</span><strong>' + esc(price) + '</strong><small>desde ' + esc(effectiveFrom) + '</small></div></section>';
    var footer = '<div class="canteen-footer-grid"><article class="card canteen-closures"><div class="card-title-row"><div><p class="card-label">Funcionamento</p><h3>Encerramentos</h3></div><span class="metric-icon"><i data-lucide="calendar-off"></i></span></div><div class="canteen-closure-list"><p>' + esc(summerCopy) + '</p><p>' + esc(seasonalCopy) + '</p><small>' + esc(alternativesCopy) + '</small></div></article><article class="card canteen-allergens"><div class="card-title-row"><div><p class="card-label">Informação alimentar</p><h3>Alergénios</h3></div><span class="metric-icon"><i data-lucide="shield-alert"></i></span></div><details><summary>Ver legenda completa</summary><div class="canteen-allergen-grid">' + allergenEntries + '</div></details><p>' + esc(canteenMenu.allergenNotice || "Confirma sempre os alergénios com o responsável da unidade.") + '</p></article></div><aside class="canteen-meal-note"><i data-lucide="info"></i><p>A refeição social inclui ' + esc(includes) + '. A ementa pode sofrer alterações no próprio dia.</p></aside><section class="canteen-verification ' + statusClass + '"><span class="canteen-verified-icon"><i data-lucide="' + (allOfficial ? "badge-check" : "cloud-off") + '"></i></span><div><strong>' + esc(verificationTitle) + '</strong><p>' + esc(verificationCopy) + '</p></div><div class="canteen-source-meta"><time>' + esc(verificationTime) + '</time><span><a href="' + attr(canteenMenu.pageUrl || CANTEEN_PAGE_URL) + '" target="_blank" rel="noopener noreferrer">Ementa oficial <i data-lucide="arrow-up-right"></i></a><a href="' + attr(info.pageUrl || CANTEEN_INFO_PAGE_URL) + '" target="_blank" rel="noopener noreferrer">Preço e condições <i data-lucide="arrow-up-right"></i></a></span></div></section>';
    return head + serviceCard + '<div class="canteen-days" role="group" aria-label="Escolher dia">' + days.map(canteenDayChip).join("") + '</div><div class="canteen-date-heading"><h3>' + esc(longDate.charAt(0).toUpperCase() + longDate.slice(1)) + '</h3><span>Almoço</span></div><div class="canteen-menu-stack">' + lunchCard + renderCanteenDinner(dinner, canteenMenu, hours) + '</div>' + footer;
  }

  function renderSettings() {
    setHeader("Admin & dados", "Configuração local-first");
    var semester = currentSemester();
    var archived = state.semesters.filter(function (item) { return item.archived; }).length;
    var lastCheck = state.meta.externalCheckedAt ? new Intl.DateTimeFormat("pt-PT", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }).format(new Date(state.meta.externalCheckedAt)) : "Nunca";
    var syncInfo = Sync ? Sync.getStatus() : { state: "disabled", configured: false, pending: 0, conflicts: 0, lastError: "", localVersion: 0, remoteVersion: 0, outdated: false };
    var syncDisplay = syncDisplayInfo(syncInfo);
    var syncLocalVersion = Number(syncInfo.localVersion) || 0;
    var syncRemoteVersion = Number(syncInfo.remoteVersion) || 0;
    var syncVersionSummary = syncInfo.pending ? syncInfo.pending + " por enviar" : syncRemoteVersion ? (syncInfo.outdated || (syncLocalVersion && syncLocalVersion !== syncRemoteVersion) ? "v" + syncLocalVersion + " → v" + syncRemoteVersion : "Versão Git v" + syncRemoteVersion) : "PC + telemóvel";
    var syncVersionBadge = syncInfo.conflicts ? syncInfo.conflicts + " conflito(s)" : syncInfo.outdated ? "Desatualizado" : syncRemoteVersion ? "Atualizado" : "Protegido";
    var forceDisabled = syncInfo.configured && syncInfo.state !== "syncing" ? "" : " disabled";
    var forceControls = '<span class="sync-force-actions" role="group" aria-label="Substituição manual de dados"><button class="button button-small sync-icon-button" type="button" data-action="force-git-pull" aria-label="Forçar pull: substituir este dispositivo pelos dados do Git" title="Forçar pull"' + forceDisabled + '><i data-lucide="arrow-down-to-line"></i></button><button class="button button-dark button-small sync-icon-button" type="button" data-action="force-git-push" aria-label="Forçar push: substituir o Git pelos dados deste dispositivo" title="Forçar push"' + forceDisabled + '><i data-lucide="arrow-up-to-line"></i></button></span>';
    var syncProgress = '<div id="gitSyncInlineProgress" class="sync-inline-progress ' + (syncInfo.state === "syncing" ? "is-active is-indeterminate" : "") + '"><span></span></div>';
    var syncCard = '<article id="gitSyncCard" class="card settings-card card-violet" aria-busy="' + (syncInfo.state === "syncing" ? "true" : "false") + '"><div class="card-title-row"><div><p class="card-label">Git como base de dados</p><h3 id="gitSyncTitle">' + esc(syncDisplay.title) + '</h3><p id="gitSyncDetail" class="card-subtitle">' + esc(syncDisplay.detail) + '</p></div><span class="metric-icon"><i data-lucide="git-commit-horizontal"></i></span></div><div class="settings-row"><div><strong id="gitSyncSummary">' + esc(syncVersionSummary) + '</strong><small>Fusão por campos · fila offline · histórico em commits.</small></div><span id="gitSyncBadge" class="badge ' + syncDisplay.badgeClass + '">' + esc(syncVersionBadge) + '</span></div>' + syncProgress + '<div class="list-actions"><button class="button button-small" type="button" data-action="configure-git-sync"><i data-lucide="settings-2"></i>Configurar</button><button class="button button-dark button-small" type="button" data-action="sync-now"><i data-lucide="refresh-cw"></i>Sincronizar agora</button>' + forceControls + (syncInfo.configured ? '<button class="button button-small" type="button" data-action="disable-git-sync"><i data-lucide="pause"></i>Pausar</button>' : '') + '</div></article>';
    return '<div class="page-head"><div><h2>Definições</h2><p>Perfil, semestres, conteúdo e dados locais.</p></div><div class="page-actions"><button class="button" type="button" data-action="show-tutorial"><i data-lucide="map"></i>Visita guiada</button><button class="button button-dark" type="button" data-action="quick-add"><i data-lucide="plus"></i>Adicionar conteúdo</button></div></div><div class="settings-grid">' + syncCard + '<article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Perfil académico</p><h3>' + esc(state.profile.name || "Estudante") + '</h3><p class="card-subtitle">' + esc(state.profile.degree || "Curso por configurar") + (state.profile.institution ? " · " + esc(state.profile.institution) : "") + '</p></div><span class="metric-icon"><i data-lucide="user-round"></i></span></div><div class="settings-row"><div><strong>Meta</strong><small>Objetivo utilizado nos indicadores de desempenho.</small></div><span class="badge badge-yellow">' + (Number(state.profile.targetGrade) || 20) + '/20</span></div><button class="button button-small" type="button" data-action="edit-profile"><i data-lucide="pencil"></i>Editar perfil</button></article><article class="card settings-card card-violet"><div class="card-title-row"><div><p class="card-label">Semestre ativo</p><h3>' + esc(semester ? semester.name : "Nenhum") + '</h3><p class="card-subtitle">' + esc(semester ? semester.academicYear : "Cria o próximo semestre") + '</p></div><span class="metric-icon"><i data-lucide="calendar-range"></i></span></div><div class="settings-row"><div><strong>' + activeCourses().length + ' cadeiras</strong><small>' + archived + ' semestre(s) no arquivo.</small></div><span class="badge badge-dark">' + activeCourses().reduce(function (sum, course) { return sum + (Number(course.ects) || 0); }, 0) + ' ECTS</span></div><div class="list-actions"><button class="button button-small" type="button" data-action="new-semester"><i data-lucide="calendar-plus"></i>Novo</button>' + (semester ? '<button class="button button-danger button-small" type="button" data-action="archive-semester"><i data-lucide="archive"></i>Arquivar semestre</button>' : "") + '</div></article><article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Ficheiro JSON</p><h3>academic-data.json</h3><p class="card-subtitle">Editável fora da app; relido ao abrir ou regressar à janela.</p></div><span class="metric-icon"><i data-lucide="braces"></i></span></div><div class="settings-row"><div><strong>Última verificação</strong><small>' + esc(lastCheck) + ' · revisão local ' + (Number(state.meta.revision) || 0) + '</small></div><button class="switch ' + (state.settings.jsonSync ? "is-on" : "") + '" type="button" data-action="toggle-json-sync" aria-label="Ativar sincronização JSON"><span></span></button></div><div class="list-actions"><button class="button button-small" type="button" data-action="reload-json"><i data-lucide="refresh-cw"></i>Reler</button><button class="button button-small" type="button" data-action="export-json"><i data-lucide="download"></i>Exportar</button><button class="button button-small" type="button" data-action="import-json"><i data-lucide="upload"></i>Importar</button></div></article><article class="card settings-card card-yellow"><div class="card-title-row"><div><h3>Atividade simulada no campus</h3><p class="card-subtitle">Apresenta indicadores simulados de atividade no campus.</p></div><span class="metric-icon"><i data-lucide="users-round"></i></span></div><div class="settings-row"><div><strong>Contador simulado</strong><small>Mostra “pessoas a acompanhar” com etiqueta de simulação.</small></div><button class="switch ' + (state.settings.campusSimulation ? "is-on" : "") + '" type="button" data-action="toggle-campus"><span></span></button></div><span class="badge badge-dark">Local · privado · transparente</span></article><article class="card settings-card"><div class="card-title-row"><div><h3>Dados no dispositivo</h3><p class="card-subtitle">Os metadados ficam em IndexedDB; os PDFs enviados ficam separados do JSON.</p></div><span class="metric-icon"><i data-lucide="hard-drive"></i></span></div><div class="settings-row"><div><strong id="storageFileCount">A contar ficheiros…</strong><small>PDFs e documentos enviados na app.</small></div><span class="badge badge-mint">Local-first</span></div><button class="button button-small" type="button" data-action="export-json"><i data-lucide="shield-check"></i>Criar backup JSON</button></article><article class="card settings-card card-dark"><div class="card-title-row"><div><h3>Adicionar conteúdo</h3><p class="card-subtitle">Aulas, materiais, perguntas antigas, quizzes, notas e avaliações.</p></div><span class="metric-icon"><i data-lucide="wrench"></i></span></div><div class="quick-grid" style="grid-template-columns:repeat(3,1fr);margin-top:17px"><button type="button" data-action="create-lesson"><i data-lucide="presentation"></i>Aula</button><button type="button" data-action="add-material"><i data-lucide="file-up"></i>PDF</button><button type="button" data-action="add-question"><i data-lucide="message-circle-question"></i>Pergunta</button><button type="button" data-action="add-quiz"><i data-lucide="sparkles"></i>Quiz</button><button type="button" data-action="add-grade"><i data-lucide="chart-no-axes-combined"></i>Nota</button><button type="button" data-action="add-assessment"><i data-lucide="file-pen-line"></i>Avaliação</button></div></article><article class="card settings-card span-12"><div class="card-title-row"><div><p class="card-label">Segurança</p><h3>Recomeçar neste dispositivo</h3><p class="card-subtitle">Remove o estado local e os PDFs guardados. O ficheiro academic-data.json não é apagado.</p></div><button class="button button-danger" type="button" data-action="reset-app"><i data-lucide="trash-2"></i>Apagar dados locais</button></div></article></div>';
  }

  function updateStorageCount() {
    DB.listFiles().then(function (files) {
      var target = document.getElementById("storageFileCount");
      if (!target) return;
      var bytes = files.reduce(function (sum, file) { return sum + (Number(file.size) || 0); }, 0);
      target.textContent = files.length + " ficheiro(s) · " + formatBytes(bytes);
    }).catch(function () {});
  }

  function enhanceSettingsActions() {
    var grid = view.querySelector(".settings-card.card-dark .quick-grid");
    if (!grid || grid.querySelector('[data-action="add-past-exam"]')) return;
    grid.insertAdjacentHTML("beforeend", '<button type="button" data-action="add-past-exam"><i data-lucide="file-json-2"></i>Teste anterior</button><button type="button" data-action="import-courses"><i data-lucide="braces"></i>Cadeiras JSON</button><button type="button" data-action="study-planner-settings"><i data-lucide="sliders-horizontal"></i>Planeamento</button><button type="button" data-route="study" data-tab="weekly"><i data-lucide="clipboard-check"></i>Revisão semanal</button>');
    refreshIcons(grid);
  }

  function formatBytes(bytes) {
    if (!bytes) return "0 KB";
    var units = ["B", "KB", "MB", "GB"];
    var index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    return round(bytes / Math.pow(1024, index), index ? 1 : 0) + " " + units[index];
  }

  function openModal(title, body, options) {
    options = options || {};
    closeModal();
    modalRoot.innerHTML = '<div class="modal-layer" role="presentation"><section class="modal ' + (options.className || "") + '" role="dialog" aria-modal="true" aria-labelledby="modalTitle"><header class="modal-head"><h2 id="modalTitle">' + esc(title) + '</h2><button class="modal-close" type="button" data-action="close-modal" aria-label="Fechar"><i data-lucide="x"></i></button></header><div class="modal-body">' + body + "</div>" + (options.footer || "") + "</section></div>";
    document.body.style.overflow = "hidden";
    refreshIcons(modalRoot);
    hydrateLocalImages(modalRoot);
    var first = modalRoot.querySelector("input:not([type=hidden]), select, textarea, button");
    if (first) setTimeout(function () { first.focus(); }, 30);
  }

  function closeModal() {
    if (activeObjectUrl) {
      URL.revokeObjectURL(activeObjectUrl);
      activeObjectUrl = null;
    }
    modalRoot.innerHTML = "";
    revokeImageObjectUrls();
    hydrateLocalImages(view);
    if (!onboarding) document.body.style.overflow = "";
  }

  function formFooter(label) {
    return '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Cancelar</button><button class="button button-dark" type="submit" form="entityForm"><i data-lucide="check"></i>' + esc(label || "Guardar") + "</button></footer>";
  }

  function courseOptions(selected, includeArchived) {
    var courses = includeArchived ? state.courses : activeCourses();
    return courses.map(function (course) {
      return '<option value="' + attr(course.id) + '" ' + (selected === course.id ? "selected" : "") + '>' + esc(course.code ? course.code + " · " + course.name : course.name) + "</option>";
    }).join("");
  }

  function pastExamOptions(courseId, selectedId) {
    return state.pastExams.filter(function (exam) { return !courseId || exam.courseId === courseId; }).sort(function (a, b) { return String(b.academicYear || b.date).localeCompare(String(a.academicYear || a.date)); }).map(function (exam) {
      return '<option value="' + attr(exam.id) + '" ' + (selectedId === exam.id ? "selected" : "") + '>' + esc(exam.title + (exam.academicYear ? " · " + exam.academicYear : "")) + '</option>';
    }).join("");
  }

  function lessonOptions(courseId, selectedIds, includeAnyCourse) {
    selectedIds = asArray(selectedIds);
    var lessons = state.lessons.filter(function (lesson) {
      return includeAnyCourse ? lesson.semesterId === state.currentSemesterId : (!courseId || lesson.courseId === courseId);
    }).sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
    return lessons.map(function (lesson) {
      var course = courseById(lesson.courseId);
      return '<option value="' + attr(lesson.id) + '" ' + (selectedIds.indexOf(lesson.id) >= 0 ? "selected" : "") + '>' + esc((includeAnyCourse && course ? (course.code || course.name) + " · " : "") + lesson.title + (lesson.date ? " · " + formatDate(lesson.date) : "")) + "</option>";
    }).join("");
  }

  function scheduleOptionsForLesson(courseId, lessonType, dateValue, selectedId) {
    var entries = semesterItems("schedule").filter(function (entry) {
      return entry.courseId === courseId && String(entry.type || "") === String(lessonType || "") && scheduleMatchesDate(entry, dateValue);
    }).sort(function (a, b) { return String(a.start).localeCompare(String(b.start)); });
    return entries.map(function (entry) {
      var label = entry.start + "–" + entry.end + " · " + lessonTypeLabel(entry.type) + (entry.room ? " · " + entry.room : "");
      return '<option value="' + attr(entry.id) + '" ' + (entry.id === selectedId ? "selected" : "") + '>' + esc(label) + '</option>';
    }).join("");
  }

  function inferredScheduleId(lesson) {
    if (!lesson) return "";
    if (lesson.scheduleId && scheduleById(lesson.scheduleId)) return lesson.scheduleId;
    var matches = semesterItems("schedule").filter(function (entry) { return lessonMatchesSchedule(lesson, entry); });
    return matches.length === 1 ? matches[0].id : "";
  }

  function updateLessonScheduleOptions(form, preferredId) {
    if (!form) return;
    var courseId = form.elements.courseId && form.elements.courseId.value;
    var type = form.elements.lessonType && form.elements.lessonType.value;
    var dateValue = form.elements.date && form.elements.date.value;
    var select = form.elements.scheduleId;
    if (!select) return;
    var current = preferredId || select.value;
    var options = scheduleOptionsForLesson(courseId, type, dateValue, current);
    select.innerHTML = '<option value="">' + (options ? "Escolher bloco…" : "Sem bloco compatível no horário") + '</option>' + options;
    if (current && Array.from(select.options).some(function (option) { return option.value === current; })) select.value = current;
    if (!select.value && select.options.length === 2) select.selectedIndex = 1;
    applyLessonScheduleSelection(form);
  }

  function applyLessonScheduleSelection(form) {
    if (!form || !form.elements.scheduleId) return;
    var entry = scheduleById(form.elements.scheduleId.value);
    if (!entry) return;
    if (form.elements.courseId) form.elements.courseId.value = entry.courseId;
    if (form.elements.lessonType) form.elements.lessonType.value = entry.type;
    if (form.elements.start) form.elements.start.value = entry.start;
    if (form.elements.end) form.elements.end.value = entry.end;
    if (form.elements.room && !form.elements.room.value.trim()) form.elements.room.value = entry.room || "";
  }

  function updateAssessmentComponentOptions(form) {
    if (!form || !form.elements.componentId) return;
    var course = courseById(form.elements.courseId.value);
    var select = form.elements.componentId;
    var prior = select.value;
    var suggestion = suggestedComponentId(course, { type: form.elements.assessmentType.value, title: form.elements.title && form.elements.title.value, componentId: prior });
    var options = componentOptionsForCourse(course && course.id, prior || suggestion);
    select.innerHTML = '<option value="">' + (options ? "Escolher componente…" : "Configura o método da cadeira") + '</option>' + options;
    if (suggestion && Array.from(select.options).some(function (option) { return option.value === suggestion; })) select.value = suggestion;
  }

  function assessmentReplacementOptions(courseId, currentId, selectedIds) {
    selectedIds = asArray(selectedIds);
    return semesterItems("assessments").filter(function (assessment) { return assessment.courseId === courseId && assessment.id !== currentId; }).sort(function (a, b) { return String(a.date || "9999").localeCompare(String(b.date || "9999")); }).map(function (assessment) {
      return '<option value="' + attr(assessment.id) + '" ' + (selectedIds.indexOf(assessment.id) >= 0 ? "selected" : "") + '>' + esc(assessment.title + " · " + (assessment.type || "Avaliação") + (assessment.date ? " · " + formatDate(assessment.date) : "")) + '</option>';
    }).join("");
  }

  function updateAssessmentLinkedOptions(form) {
    if (!form) return;
    updateAssessmentComponentOptions(form);
    var courseId = form.elements.courseId && form.elements.courseId.value;
    var lessonSelect = form.elements.lessonIds;
    if (lessonSelect) {
      var selectedLessons = selectedValues(lessonSelect);
      lessonSelect.innerHTML = lessonOptions(courseId, selectedLessons, false);
    }
    var replacementSelect = form.elements.replacementAssessmentIds;
    if (replacementSelect) {
      var selectedReplacements = selectedValues(replacementSelect);
      replacementSelect.innerHTML = assessmentReplacementOptions(courseId, form.getAttribute("data-id") || "", selectedReplacements);
    }
  }

  function defenseTypeOptions(selected) {
    return '<option value="oral" ' + (selected === "oral" ? "selected" : "") + '>Oral</option><option value="practical" ' + (selected === "practical" ? "selected" : "") + '>Prática</option><option value="oral-practical" ' + (selected === "oral-practical" ? "selected" : "") + '>Oral e prática</option>';
  }

  function renderAssessmentForm(assessment, existingAssessment, assessmentCourse, assessmentType, assessmentComponent) {
    var replacementOptions = assessmentReplacementOptions(assessmentCourse, existingAssessment && existingAssessment.id, asArray(assessment.replacementAssessmentIds));
    return '<form id="entityForm" data-type="assessment" data-id="' + attr(existingAssessment && existingAssessment.id) + '"><div class="form-grid"><div class="field"><label>Cadeira</label><select name="courseId" data-role="assessment-course" required><option value="">Escolher…</option>' + courseOptions(assessmentCourse) + '</select></div><div class="field"><label>Tipo</label><select name="assessmentType" data-role="assessment-type"><option ' + (assessmentType === "Teste" ? "selected" : "") + '>Teste</option><option ' + (assessmentType === "Exame" ? "selected" : "") + '>Exame</option><option ' + (assessmentType === "Projeto" ? "selected" : "") + '>Projeto</option><option ' + (assessmentType === "Apresentação" ? "selected" : "") + '>Apresentação</option><option ' + (assessmentType === "Trabalho" ? "selected" : "") + '>Trabalho</option><option ' + (assessmentType === "Mini-teste" ? "selected" : "") + '>Mini-teste</option><option ' + (assessmentType === "Oral" ? "selected" : "") + '>Oral</option><option ' + (assessmentType === "Personalizada" ? "selected" : "") + '>Personalizada</option></select></div><div class="field field-full"><label>Nome</label><input name="title" data-role="assessment-title" required placeholder="Ex.: Teste 1" value="' + attr(assessment.title || "") + '"></div><div class="field field-full"><label>Componente do método de avaliação</label><select name="componentId" data-role="assessment-component" required><option value="">Escolher componente…</option>' + componentOptionsForCourse(assessmentCourse, assessmentComponent) + '</select><small>Liga esta avaliação ao grupo certo: Testes, Projetos, Quiz, Mini-projetos ou outro.</small></div><div class="field"><label>Data</label><input name="date" type="date" value="' + attr(assessment.date || "") + '"></div><div class="field"><label>Hora</label><input name="time" type="time" value="' + attr(assessment.time || "10:00") + '"></div><div class="field"><label>Peso informativo (%)</label><input name="weight" type="number" min="0" max="100" step="0.5" value="' + attr(assessment.weight || 0) + '"><small>O peso total continua definido na componente.</small></div><div class="field"><label>Local</label><input name="location" placeholder="Ex.: Auditório 1" value="' + attr(assessment.location || "") + '"></div><div class="field field-full"><label>Aulas que saem nesta avaliação</label><select name="lessonIds" multiple size="7">' + lessonOptions(assessmentCourse, asArray(assessment.lessonIds), false) + '</select><small>Usa Ctrl/Cmd para escolher várias. Podes misturar teóricas e práticas.</small></div><section class="assessment-rules field-full"><div><h3>Condições da avaliação</h3><p>Regista o que tens de levar, se há consulta, defesa ou limites de nota.</p></div><div class="checkbox-line"><label class="checkbox-chip"><input type="checkbox" name="requiresTestSheet" ' + (assessment.requiresTestSheet ? "checked" : "") + '> Precisa de folha de teste</label><label class="checkbox-chip"><input type="checkbox" name="openBook" ' + (assessment.openBook ? "checked" : "") + '> É de consulta</label><label class="checkbox-chip"><input type="checkbox" name="hasDefense" ' + (assessment.hasDefense ? "checked" : "") + '> Tem defesa</label></div><div class="assessment-defense-grid"><div class="field"><label>Tipo de defesa</label><select name="defenseType">' + defenseTypeOptions(assessment.defenseType || "oral") + '</select></div><div class="field"><label>Defesa necessária a partir de</label><input name="defenseThreshold" type="number" min="0" max="20" step="0.1" placeholder="Ex.: 14" value="' + attr(assessment.defenseThreshold == null ? "" : assessment.defenseThreshold) + '"></div><div class="field"><label>Nota máxima sem defesa</label><input name="maxWithoutDefense" type="number" min="0" max="20" step="0.1" placeholder="Ex.: 12" value="' + attr(assessment.maxWithoutDefense == null ? "" : assessment.maxWithoutDefense) + '"></div></div></section><section class="assessment-rules field-full"><div><h3>Substituição de notas</h3><p>Normalmente usada num exame: escolhe exatamente que avaliações podem ser substituídas.</p></div><div class="form-grid"><div class="field field-full"><label>Avaliações substituídas</label><select name="replacementAssessmentIds" multiple size="5">' + replacementOptions + '</select><small>Sem seleção, esta avaliação não substitui nenhuma nota.</small></div><div class="field"><label>Regra</label><select name="replacementPolicy"><option value="if-higher" ' + (assessment.replacementPolicy !== "always" ? "selected" : "") + '>Só se a nota for superior</option><option value="always" ' + (assessment.replacementPolicy === "always" ? "selected" : "") + '>Substitui sempre</option></select></div></div></section></div></form>';
  }

  function gradeDefenseConfiguration(assessmentId) {
    var assessment = assessmentById(assessmentId);
    if (!assessment) return { enabled: false, type: "oral", threshold: null, maxWithoutDefense: null };
    var course = courseById(assessment.courseId);
    var component = asArray(course && course.evaluation && course.evaluation.components).find(function (item) { return item.id === assessment.componentId; });
    var assessmentDefense = !!assessment.hasDefense;
    return {
      enabled: assessmentDefense || !!(component && component.defenseEnabled),
      type: assessmentDefense ? assessment.defenseType : component && component.defenseType || "oral",
      threshold: nullableNumber(assessmentDefense ? assessment.defenseThreshold : component && component.defenseThreshold),
      maxWithoutDefense: nullableNumber(assessmentDefense ? assessment.maxWithoutDefense : component && component.maxWithoutDefense)
    };
  }

  function renderGradeDefenseFields(assessmentId) {
    var config = gradeDefenseConfiguration(assessmentId);
    var notes = [];
    if (config.threshold != null) notes.push("defesa a partir de " + round(config.threshold, 1) + "/20");
    if (config.maxWithoutDefense != null) notes.push("máximo sem defesa: " + round(config.maxWithoutDefense, 1) + "/20");
    return '<section id="gradeDefenseFields" class="assessment-rules field-full" ' + (config.enabled ? "" : "hidden") + '><div><h3>Defesa da nota</h3><p>' + esc(notes.length ? notes.join(" · ") : "Regista o estado e, quando existir, a nota final após a defesa.") + '</p></div><div class="assessment-defense-grid"><div class="field"><label>Estado</label><select name="defenseStatus"><option value="pending" ' + (config.enabled ? "selected" : "") + '>Por realizar</option><option value="completed">Concluída</option><option value="not-applicable" ' + (!config.enabled ? "selected" : "") + '>Não aplicável</option></select></div><div class="field"><label>Tipo</label><select name="gradeDefenseType">' + defenseTypeOptions(config.type) + '</select></div><div class="field"><label>Nota final após defesa</label><input name="defenseFinalScore" type="number" min="0" max="20" step="0.1" placeholder="Opcional"></div></div></section>';
  }

  function updateGradeDefenseFields(form) {
    if (!form || !form.elements.target) return;
    var parts = String(form.elements.target.value || "").split("|");
    var current = form.querySelector("#gradeDefenseFields");
    if (!current) return;
    var wrapper = document.createElement("div");
    wrapper.innerHTML = renderGradeDefenseFields(parts[0] === "assessment" ? parts[1] : "");
    current.replaceWith(wrapper.firstElementChild);
  }

  function evaluationToText(course) {
    return asArray(course && course.evaluation && course.evaluation.components).map(function (item) {
      return item.label + " | " + (Number(item.weight) || 0) + " | " + (item.kind || "other");
    }).join("\n");
  }

  function evaluationKindOptions(selected) {
    var options = [
      ["test", "Testes"], ["project", "Projetos"], ["exam", "Exame"], ["presentation", "Apresentações"], ["class", "Aulas / participação"], ["other", "Personalizado"]
    ];
    return options.map(function (item) { return '<option value="' + item[0] + '" ' + (selected === item[0] ? "selected" : "") + '>' + item[1] + '</option>'; }).join("");
  }

  function renderEvaluationComponentRow(component) {
    component = Object.assign({ id: uid("component"), label: "Componente", kind: "other", count: 1, weight: 0, minimum: null, defenseEnabled: false, defenseType: "oral", defenseThreshold: null, maxWithoutDefense: null }, component || {});
    return '<article class="evaluation-builder-row"><input type="hidden" name="componentId" value="' + attr(component.id) + '"><div class="evaluation-builder-main"><div class="field"><label>Nome</label><input name="componentLabel" required placeholder="Ex.: Mini-projetos" value="' + attr(component.label) + '"></div><div class="field"><label>Tipo</label><select name="componentKind">' + evaluationKindOptions(component.kind) + '</select></div><div class="field"><label>Quantidade</label><input name="componentCount" type="number" min="1" max="30" step="1" value="' + attr(Math.max(1, Number(component.count) || 1)) + '"></div><div class="field"><label>Peso total (%)</label><input name="componentWeight" data-role="component-weight" type="number" min="0" max="100" step="0.5" value="' + attr(Number(component.weight) || 0) + '"></div><button class="remove-evaluation-row" type="button" data-action="remove-evaluation-component" aria-label="Remover componente"><i data-lucide="trash-2"></i></button></div><details class="evaluation-advanced"><summary>Regras avançadas</summary><div class="evaluation-rule-grid"><div class="field"><label>Nota mínima para aprovação</label><input name="componentMinimum" type="number" min="0" max="20" step="0.1" placeholder="Sem mínimo" value="' + attr(component.minimum == null ? "" : component.minimum) + '"><small>A média desta componente tem de atingir este valor.</small></div><label class="checkbox-chip evaluation-defense-toggle"><input name="componentDefenseEnabled" type="checkbox" ' + (component.defenseEnabled ? "checked" : "") + '> Pode exigir defesa</label><div class="field"><label>Tipo de defesa</label><select name="componentDefenseType"><option value="oral" ' + (component.defenseType === "oral" ? "selected" : "") + '>Oral</option><option value="practical" ' + (component.defenseType === "practical" ? "selected" : "") + '>Prática</option><option value="oral-practical" ' + (component.defenseType === "oral-practical" ? "selected" : "") + '>Oral e prática</option></select></div><div class="field"><label>Defesa necessária a partir de</label><input name="componentDefenseThreshold" type="number" min="0" max="20" step="0.1" placeholder="Ex.: 14" value="' + attr(component.defenseThreshold == null ? "" : component.defenseThreshold) + '"></div><div class="field"><label>Nota máxima sem defesa</label><input name="componentMaxWithoutDefense" type="number" min="0" max="20" step="0.1" placeholder="Ex.: 12" value="' + attr(component.maxWithoutDefense == null ? "" : component.maxWithoutDefense) + '"></div></div></details></article>';
  }

  function renderEvaluationBuilder(course) {
    var components = asArray(course && course.evaluation && course.evaluation.components);
    if (!components.length) components = [
      { id: uid("component"), label: "Testes", kind: "test", count: 2, weight: 60 },
      { id: uid("component"), label: "Projetos", kind: "project", count: 1, weight: 40 },
      { id: uid("component"), label: "Exame", kind: "exam", count: 1, weight: 0 }
    ];
    return '<section class="evaluation-builder field-full"><div class="evaluation-builder-head"><div><h3>Método de avaliação</h3><p>Indica quantos elementos existem e o peso total de cada grupo.</p></div><button class="button button-small" type="button" data-action="add-evaluation-component"><i data-lucide="plus"></i>Componente</button></div><div class="evaluation-builder-list">' + components.map(renderEvaluationComponentRow).join("") + '</div><div class="evaluation-weight-summary" data-role="evaluation-weight-summary"></div><p class="form-note">Podes criar Testes, Projetos, Exame ou qualquer componente personalizada, como Quiz ou Mini-projeto. As defesas e notas mínimas ficam nas regras avançadas.</p></section>';
  }

  function readEvaluationBuilder(form) {
    return Array.from(form.querySelectorAll(".evaluation-builder-row")).map(function (row) {
      var kind = row.querySelector('[name="componentKind"]').value || "other";
      return {
        id: row.querySelector('[name="componentId"]').value || uid("component"),
        label: row.querySelector('[name="componentLabel"]').value.trim(),
        kind: kind,
        count: Math.max(1, Number(row.querySelector('[name="componentCount"]').value) || 1),
        weight: clamp(row.querySelector('[name="componentWeight"]').value, 0, 100),
        minimum: nullableNumber(row.querySelector('[name="componentMinimum"]').value),
        defenseEnabled: row.querySelector('[name="componentDefenseEnabled"]').checked,
        defenseType: row.querySelector('[name="componentDefenseType"]').value || "oral",
        defenseThreshold: nullableNumber(row.querySelector('[name="componentDefenseThreshold"]').value),
        maxWithoutDefense: nullableNumber(row.querySelector('[name="componentMaxWithoutDefense"]').value),
        replaceable: kind === "test"
      };
    }).filter(function (component) { return component.label; });
  }

  function updateEvaluationBuilderSummary(form) {
    if (!form) return;
    var target = form.querySelector('[data-role="evaluation-weight-summary"]');
    if (!target) return;
    var total = Array.from(form.querySelectorAll('[data-role="component-weight"]')).reduce(function (sum, input) { return sum + (Number(input.value) || 0); }, 0);
    target.className = "evaluation-weight-summary " + (Math.abs(total - 100) < .01 ? "is-complete" : total > 100 ? "is-over" : "is-pending");
    target.innerHTML = '<span><i data-lucide="' + (Math.abs(total - 100) < .01 ? "circle-check" : "circle-alert") + '"></i>Peso configurado</span><strong>' + round(total, 1) + '%</strong>';
    refreshIcons(target);
  }

  function pastExamJSONExample() {
    return {
      title: "Teste 1 — época normal",
      academicYear: "2024/2025",
      date: "2025-01-15",
      source: "PDF disponibilizado pelo professor",
      questions: [{
        number: "1.1",
        prompt: "Transcrição exata do enunciado.",
        answer: "",
        explanation: "",
        points: 2,
        lessonTitles: ["Aula 03 · Tema"],
        tags: ["tema"],
        options: [],
        answerIndex: null,
        images: {
          question: ["assets/perguntas/teste-1-q1.png"],
          solution: ["assets/solucoes/teste-1-q1.png"],
          explanation: []
        }
      }]
    };
  }

  function courseJSONExample() {
    return {
      courses: [{
        name: "Programação Orientada a Objetos",
        code: "POO",
        ects: 6,
        color: "#a99df7",
        lessonTypes: ["T", "TP"],
        evaluation: {
          components: [
            { label: "Testes", kind: "test", count: 2, weight: 60, minimum: 9.5, defenseEnabled: false, defenseType: "oral", defenseThreshold: null, maxWithoutDefense: null },
            { label: "Projeto", kind: "project", count: 1, weight: 40, minimum: 10, defenseEnabled: true, defenseType: "oral-practical", defenseThreshold: null, maxWithoutDefense: 12 },
            { label: "Exame", kind: "exam", count: 1, weight: 0, minimum: null, defenseEnabled: false, defenseType: "oral", defenseThreshold: null, maxWithoutDefense: null }
          ],
          examReplacesTests: true,
          replacementPolicy: "if-higher"
        }
      }]
    };
  }

  function importPrompt(kind) {
    if (kind === "course") {
      return "Analisa apenas a informação visível nos documentos fornecidos sobre as cadeiras e converte-a para o JSON abaixo. Não inventes nomes, ECTS, tipos de aula, quantidades, percentagens, notas mínimas, defesas ou regras de exame. Quando um valor não estiver explícito, usa string vazia, null ou [] conforme o campo. Preserva exatamente os nomes usados pela instituição. Os tipos permitidos em lessonTypes são T, TP, P, LAB e OT. Os tipos permitidos em evaluation.components[].kind são test, project, exam, presentation, class e other. A soma dos pesos não pode ultrapassar 100. Responde apenas com JSON válido, sem markdown nem explicações. Formato:\n\n" + JSON.stringify(courseJSONExample(), null, 2);
    }
    return "Analisa o teste ou exame anterior fornecido e transcreve-o para o JSON abaixo. Não inventes texto, valores, unidades, opções, soluções, explicações, cotações, nomes de ficheiro ou ligações a aulas. Preserva a numeração, símbolos, fórmulas e unidades exatamente como aparecem. Se uma parte estiver ilegível, escreve [ILEGÍVEL] apenas nesse ponto; se um campo não estiver visível, deixa-o vazio, null ou []. Só preenche answer e explanation quando a solução estiver efetivamente presente. Em lessonTitles usa apenas títulos de aulas que eu tenha fornecido; caso contrário usa []. Se o enunciado depender de uma figura e não te tiver sido dado um caminho real, assinala [FIGURA — UPLOAD NECESSÁRIO] no prompt e deixa images vazio. Só usa um caminho de imagem quando esse nome ou caminho tiver sido fornecido. Responde apenas com JSON válido, sem markdown nem comentários. Formato:\n\n" + JSON.stringify(pastExamJSONExample(), null, 2);
  }

  function importTools(kind, textareaId) {
    return '<div class="import-tool-row"><button class="button button-small" type="button" data-action="fill-import-example" data-kind="' + attr(kind) + '" data-target="' + attr(textareaId) + '"><i data-lucide="braces"></i>Usar exemplo</button><button class="button button-small" type="button" data-action="copy-import-prompt" data-kind="' + attr(kind) + '"><i data-lucide="copy"></i>Copiar prompt para IA</button></div>';
  }

  function openQuickAdd() {
    openModal("Adicionar à Twenty", '<p class="onboarding-copy" style="margin-top:0">Escolhe o tipo de conteúdo. Tudo fica ligado ao semestre atual.</p><div class="quick-grid"><button type="button" data-action="create-lesson"><i data-lucide="presentation"></i>Nova aula</button><button type="button" data-action="add-material"><i data-lucide="file-up"></i>Slides / PDF</button><button type="button" data-action="add-task"><i data-lucide="notebook-pen"></i>TPC / tarefa</button><button type="button" data-action="add-assessment"><i data-lucide="file-pen-line"></i>Avaliação</button><button type="button" data-action="add-past-exam"><i data-lucide="file-json-2"></i>Teste anterior</button><button type="button" data-action="add-question"><i data-lucide="message-circle-question"></i>Pergunta antiga</button><button type="button" data-action="add-quiz"><i data-lucide="sparkles"></i>Quiz manual</button><button type="button" data-action="add-grade"><i data-lucide="chart-no-axes-combined"></i>Nota</button><button type="button" data-action="add-event"><i data-lucide="party-popper"></i>Evento</button><button type="button" data-action="add-course"><i data-lucide="library-big"></i>Cadeira</button><button type="button" data-action="import-courses"><i data-lucide="braces"></i>Cadeiras JSON</button></div>');
  }

  function openEntityForm(type, preset) {
    preset = preset || {};
    var title = "Adicionar";
    var body = "";
    var submitLabel = "Guardar";
    var semester = currentSemester();
    var year = semester ? semester.academicYear : academicYearFor();

    if (type === "course") {
      var course = preset.id ? courseById(preset.id) : null;
      title = course ? "Configurar cadeira" : "Nova cadeira";
      var selectedTypes = asArray(course && course.lessonTypes).length ? course.lessonTypes : ["T", "TP"];
      body = '<form id="entityForm" data-type="course" data-id="' + attr(course && course.id) + '"><div class="form-grid"><div class="field field-full"><label>Nome da cadeira</label><input name="name" required placeholder="Ex.: Programação Orientada a Objetos" value="' + attr(course && course.name) + '"></div><div class="field"><label>Código curto</label><input name="code" placeholder="Ex.: POO" value="' + attr(course && course.code) + '"></div><div class="field"><label>ECTS</label><input name="ects" type="number" min="0" max="60" step="0.5" value="' + attr(course ? course.ects : 6) + '"></div><div class="field"><label>Cor</label><input name="color" type="color" value="' + safeColor(course && course.color, COLORS[activeCourses().length % COLORS.length]) + '"></div><div class="field"><label>Tipos de aula</label><div class="checkbox-line">' + ["T", "TP", "P", "LAB", "OT"].map(function (value) { return '<label class="checkbox-chip"><input type="checkbox" name="lessonTypes" value="' + value + '" ' + (selectedTypes.indexOf(value) >= 0 ? "checked" : "") + '>' + value + "</label>"; }).join("") + '</div></div>' + renderEvaluationBuilder(course) + '</div></form>';
    } else if (type === "lesson") {
      title = preset.id ? "Editar aula" : "Nova aula";
      var existingLesson = preset.id ? lessonById(preset.id) : null;
      var lesson = Object.assign({}, existingLesson || {}, preset);
      if (!existingLesson && !lesson.scheduleId) {
        var suggestedClass = getNextClass(null, { courseId: lesson.courseId || "", unprepared: true });
        if (suggestedClass) {
          lesson = Object.assign({
            courseId: suggestedClass.schedule.courseId,
            scheduleId: suggestedClass.schedule.id,
            date: suggestedClass.dateISO,
            start: suggestedClass.schedule.start,
            end: suggestedClass.schedule.end,
            type: suggestedClass.schedule.type,
            room: suggestedClass.schedule.room || ""
          }, lesson);
        }
      }
      var selectedCourse = lesson.courseId || (activeCourses()[0] && activeCourses()[0].id) || "";
      var selectedLessonType = lesson.type || ((courseById(selectedCourse) || {}).lessonTypes || ["T"])[0] || "T";
      var selectedLessonDate = lesson.date || todayISO();
      var selectedScheduleId = inferredScheduleId(lesson);
      var selectedSlot = scheduleById(selectedScheduleId);
      if (selectedSlot) {
        selectedCourse = selectedSlot.courseId;
        selectedLessonType = selectedSlot.type;
        lesson.start = selectedSlot.start;
        lesson.end = selectedSlot.end;
        lesson.room = lesson.room || selectedSlot.room || "";
      }
      body = '<form id="entityForm" data-type="lesson" data-id="' + attr(existingLesson && existingLesson.id) + '"><div class="form-grid"><div class="field"><label>Cadeira</label><select name="courseId" data-role="lesson-course" required><option value="">Escolher…</option>' + courseOptions(selectedCourse) + '</select></div><div class="field"><label>Data</label><input name="date" data-role="lesson-date" type="date" required value="' + attr(selectedLessonDate) + '"></div><div class="field"><label>Tipo de aula</label><select name="lessonType" data-role="lesson-type"><option value="T" ' + (selectedLessonType === "T" ? "selected" : "") + '>Teórica</option><option value="TP" ' + (selectedLessonType === "TP" ? "selected" : "") + '>Teórico-prática</option><option value="P" ' + (selectedLessonType === "P" ? "selected" : "") + '>Prática</option><option value="LAB" ' + (selectedLessonType === "LAB" ? "selected" : "") + '>Laboratório</option><option value="OT" ' + (selectedLessonType === "OT" ? "selected" : "") + '>Orientação</option></select></div><div class="field"><label>Bloco compatível do horário</label><select name="scheduleId" data-role="lesson-schedule" required><option value="">Escolher bloco…</option>' + scheduleOptionsForLesson(selectedCourse, selectedLessonType, selectedLessonDate, selectedScheduleId) + '</select><small>Só aparecem blocos da mesma cadeira, dia e tipo.</small></div><div class="field field-full"><label>Nome da aula</label><input name="title" required placeholder="Ex.: TP08 · Herança e polimorfismo" value="' + attr(lesson.title) + '"></div><div class="field"><label>Início</label><input name="start" type="time" readonly value="' + attr(lesson.start || "") + '"></div><div class="field"><label>Fim</label><input name="end" type="time" readonly value="' + attr(lesson.end || "") + '"></div><div class="field field-full"><label>Sala</label><input name="room" placeholder="Ex.: B2.14" value="' + attr(lesson.room) + '"></div><div class="field field-full"><label>Matéria / tópicos</label><textarea name="topics" placeholder="Conceitos dados, capítulos, exercícios…">' + esc(lesson.topics) + '</textarea></div>' + (!existingLesson ? '<div class="field"><label>PDF opcional</label><input name="file" type="file" accept="application/pdf,image/*,.pptx,.txt,.md"></div><div class="field"><label>Ano letivo do PDF</label><input name="materialYear" placeholder="2025/26" value="' + attr(year) + '"></div>' : "") + '</div><div class="form-note" style="margin-top:13px">A aula fica ligada ao período real do horário. O nome que escreveres aparecerá na aula em direto e no Calendário.</div></form>';
    } else if (type === "material") {
      title = "Carregar slides ou PDF";
      var materialCourse = preset.courseId || (activeCourses()[0] && activeCourses()[0].id) || "";
      body = '<form id="entityForm" data-type="material"><div class="form-grid"><div class="field"><label>Cadeira</label><select name="courseId" required><option value="">Escolher…</option>' + courseOptions(materialCourse) + '</select></div><div class="field"><label>Aula associada</label><select name="lessonId"><option value="">Biblioteca geral</option>' + lessonOptions(null, preset.lessonId ? [preset.lessonId] : [], true) + '</select></div><div class="field field-full"><label>Título</label><input name="title" required placeholder="Ex.: Slides · Polimorfismo" value="' + attr(preset.title) + '"></div><div class="field"><label>Tipo</label><select name="kind"><option value="slides">Slides</option><option value="pdf">PDF / texto</option><option value="notes">Apontamentos</option><option value="worksheet">Ficha prática</option></select></div><div class="field"><label>Ano letivo</label><input name="academicYear" required placeholder="2024/25" value="' + attr(preset.academicYear || year) + '"><small>Se for o ano atual, a etiqueta fica oculta.</small></div><div class="field field-full"><label>Ficheiro no dispositivo</label><input name="file" type="file" accept="application/pdf,image/*,.pptx,.txt,.md"><small>O ficheiro é enviado para o repositório privado e fica também em cache neste dispositivo.</small></div><div class="field field-full"><label>Ou caminho / URL</label><input name="url" placeholder="assets/slides/aula-08.pdf ou https://…"><small>Ideal para PDFs colocados manualmente na pasta do projeto e referenciados no JSON.</small></div></div></form>';
    } else if (type === "task") {
      title = "Nova tarefa";
      body = '<form id="entityForm" data-type="task"><div class="form-grid"><div class="field field-full"><label>Título</label><input name="title" required placeholder="Ex.: Rever aula 08"></div><div class="field"><label>Tipo</label><select name="taskType"><option value="homework">Trabalho de casa</option><option value="project">Projeto</option><option value="review" ' + (preset.type === "review" ? "selected" : "") + '>Rever aula</option><option value="reading">Leitura</option><option value="other">Outro</option></select></div><div class="field"><label>Cadeira</label><select name="courseId"><option value="">Pessoal / geral</option>' + courseOptions(preset.courseId || "") + '</select></div><div class="field"><label>Prazo</label><input name="dueDate" type="date" value="' + attr(preset.dueDate || todayISO()) + '"></div><div class="field"><label>Hora</label><input name="dueTime" type="time" value="' + attr(preset.dueTime || "18:00") + '"></div><div class="field"><label>Prioridade</label><select name="priority"><option value="normal">Normal</option><option value="high">Alta</option><option value="low">Baixa</option></select></div><div class="field"><label>Ligada à aula</label><select name="lessonId"><option value="">Nenhuma</option>' + lessonOptions(null, preset.lessonId ? [preset.lessonId] : [], true) + '</select></div></div></form>';
    } else if (type === "assessment") {
      var existingAssessment = preset.id ? assessmentById(preset.id) : null;
      var assessment = Object.assign({}, existingAssessment || {}, preset);
      title = existingAssessment ? "Editar avaliação" : "Nova avaliação";
      var assessmentCourse = assessment.courseId || (activeCourses()[0] && activeCourses()[0].id) || "";
      var assessmentType = assessment.assessmentType || assessment.type || "Teste";
      var assessmentComponent = assessment.componentId || suggestedComponentId(courseById(assessmentCourse), { type: assessmentType, title: assessment.title || "" });
      body = renderAssessmentForm(assessment, existingAssessment, assessmentCourse, assessmentType, assessmentComponent);
    } else if (type === "event") {
      var existingEvent = preset.id ? state.events.find(function (item) { return item.id === preset.id; }) : null;
      var eventData = Object.assign({}, existingEvent || {}, preset);
      title = existingEvent ? "Editar evento" : "Novo evento da faculdade";
      body = '<form id="entityForm" data-type="event" data-id="' + attr(existingEvent && existingEvent.id) + '"><div class="form-grid"><div class="field field-full"><label>Evento</label><input name="title" required placeholder="Ex.: Feira de emprego" value="' + attr(eventData.title || "") + '"></div><div class="field"><label>Data</label><input name="date" type="date" required value="' + attr(eventData.date || todayISO()) + '"></div><div class="field"><label>Hora</label><input name="time" type="time" value="' + attr(eventData.time || "14:00") + '"></div><div class="field"><label>Local</label><input name="location" placeholder="Campus / sala" value="' + attr(eventData.location || "") + '"></div><div class="field"><label>Ligação</label><input name="url" type="url" placeholder="https://…" value="' + attr(eventData.url || "") + '"></div><div class="field field-full"><label>Notas</label><textarea name="notes" placeholder="O que levar, inscrição, detalhes…">' + esc(eventData.notes || "") + '</textarea></div><section class="field field-full media-input-section"><div><h3>Imagens do evento</h3><p>Carrega fotografias ou usa caminhos relativos ao projeto.</p></div>' + renderExistingImageManager(eventData.images) + '<div class="form-grid"><div class="field"><label>Carregar imagens</label><input name="eventImageFiles" type="file" accept="image/*" multiple></div><div class="field"><label>Caminhos / URLs</label><textarea name="eventImagePaths" placeholder="assets/eventos/cartaz.png&#10;https://…"></textarea></div></div></section></div></form>';
    } else if (type === "question") {
      var existingQuestion = preset.id ? state.questions.find(function (item) { return item.id === preset.id; }) : null;
      var questionData = Object.assign({}, existingQuestion || {}, preset);
      title = existingQuestion ? "Editar pergunta anterior" : "Pergunta de teste anterior";
      var questionCourse = questionData.courseId || (activeCourses()[0] && activeCourses()[0].id) || "";
      var selectedQuestionLessons = asArray(questionData.lessonIds).concat(questionData.lessonId ? [questionData.lessonId] : []);
      body = '<form id="entityForm" data-type="question" data-id="' + attr(existingQuestion && existingQuestion.id) + '"><div class="form-grid"><div class="field"><label>Cadeira</label><select name="courseId" required><option value="">Escolher…</option>' + courseOptions(questionCourse) + '</select></div><div class="field"><label>Teste anterior</label><select name="pastExamId"><option value="">Sem teste associado</option>' + pastExamOptions(questionCourse, questionData.pastExamId) + '</select></div><div class="field"><label>Ano letivo do teste</label><input name="academicYear" required placeholder="2024/25" value="' + attr(questionData.academicYear || year) + '"></div><div class="field"><label>Número</label><input name="number" placeholder="Ex.: 1.2" value="' + attr(questionData.number || "") + '"></div><div class="field"><label>Origem</label><input name="assessmentLabel" placeholder="Ex.: Teste 1 · Grupo II" value="' + attr(questionData.assessmentLabel || "") + '"></div><div class="field"><label>Cotação</label><input name="points" type="number" min="0" step="0.1" placeholder="2" value="' + attr(questionData.points == null ? "" : questionData.points) + '"></div><div class="field field-full"><label>Pergunta</label><textarea name="prompt" required placeholder="Escreve a pergunta tal como apareceu…">' + esc(questionData.prompt || "") + '</textarea></div><div class="field field-full"><label>Aulas associadas</label><select name="lessonIds" multiple size="7">' + lessonOptions(questionCourse, selectedQuestionLessons, false) + '</select><small>A pergunta aparecerá dentro de todas as aulas selecionadas.</small></div><div class="field field-full"><label>Resposta / solução</label><textarea name="answer" placeholder="Resposta esperada…">' + esc(questionData.answer || "") + '</textarea></div><div class="field field-full"><label>Explicação</label><textarea name="explanation" placeholder="Raciocínio, armadilhas, critérios…">' + esc(questionData.explanation || "") + '</textarea></div><section class="field field-full media-input-section"><div><h3>Imagens</h3><p>Separa imagens do enunciado, da solução e da explicação.</p></div>' + renderExistingImageManager(questionData.images) + '<div class="media-input-grid">' + [["question", "Enunciado"], ["solution", "Solução"], ["explanation", "Explicação"]].map(function (entry) { var role = entry[0]; return '<article><h4>' + entry[1] + '</h4><label>Carregar</label><input name="' + role + 'ImageFiles" type="file" accept="image/*" multiple><label>Caminhos / URLs</label><textarea name="' + role + 'ImagePaths" placeholder="assets/perguntas/' + role + '.png"></textarea></article>'; }).join("") + '</div></section></div></form>';
    } else if (type === "quiz") {
      title = "Novo quiz da aula";
      var quizCourse = preset.courseId || (activeCourses()[0] && activeCourses()[0].id) || "";
      var selectedQuizLesson = lessonById(preset.lessonId);
      var quizTitle = selectedQuizLesson ? "Quiz · " + selectedQuizLesson.title : "";
      body = '<form id="entityForm" data-type="quiz"><div class="form-grid"><div class="field"><label>Cadeira</label><select name="courseId" required><option value="">Escolher…</option>' + courseOptions(quizCourse) + '</select></div><div class="field"><label>Aula associada</label><select name="lessonId" data-role="quiz-lesson-select"><option value="">Quiz geral</option>' + lessonOptions(null, preset.lessonId ? [preset.lessonId] : [], true) + '</select></div><div class="field field-full"><label>Título do quiz</label><input name="title" required placeholder="Ex.: Quiz · Herança e polimorfismo" value="' + attr(quizTitle) + '"></div><div class="field field-full"><label>Primeira pergunta manual <span class="optional-label">opcional</span></label><textarea name="prompt" placeholder="Escreve uma pergunta tua ou escolhe perguntas antigas abaixo…"></textarea></div>' + [0, 1, 2, 3].map(function (index) { return '<div class="field"><label>Opção ' + (index + 1) + '</label><input name="option' + index + '" placeholder="Resposta ' + (index + 1) + '"></div>'; }).join("") + '<div class="field"><label>Resposta certa</label><select name="answerIndex"><option value="0">Opção 1</option><option value="1">Opção 2</option><option value="2">Opção 3</option><option value="3">Opção 4</option></select></div><div class="field"><label>Explicação</label><input name="explanation" placeholder="Porquê esta resposta?"></div><div class="field field-full"><label>Perguntas de testes anteriores desta aula</label><div id="quizPastQuestionPicker" class="past-question-picker"></div><small>Seleciona primeiro uma aula. Podes misturar perguntas anteriores com a pergunta manual.</small></div></div></form>';
    } else if (type === "quiz-question") {
      title = "Adicionar pergunta ao quiz";
      body = '<form id="entityForm" data-type="quiz-question" data-id="' + attr(preset.quizId) + '"><div class="form-grid"><div class="field field-full"><label>Pergunta</label><textarea name="prompt" required></textarea></div>' + [0, 1, 2, 3].map(function (index) { return '<div class="field"><label>Opção ' + (index + 1) + '</label><input name="option' + index + '" ' + (index < 2 ? "required" : "") + '></div>'; }).join("") + '<div class="field"><label>Resposta certa</label><select name="answerIndex"><option value="0">Opção 1</option><option value="1">Opção 2</option><option value="2">Opção 3</option><option value="3">Opção 4</option></select></div><div class="field"><label>Explicação</label><input name="explanation"></div></div></form>';
    } else if (type === "grade") {
      title = "Adicionar nota";
      var courseAssessments = semesterItems("assessments").filter(function (item) { return !preset.courseId || item.courseId === preset.courseId; });
      var defaultAssessmentId = preset.assessmentId || (courseAssessments[0] && courseAssessments[0].id) || "";
      var gradeAssessmentOptions = activeCourses().map(function (courseItem) {
        var items = semesterItems("assessments").filter(function (assessment) { return assessment.courseId === courseItem.id; });
        if (!items.length) return "";
        return '<optgroup label="' + attr(courseItem.name) + ' · avaliações">' + items.map(function (assessment) {
          return '<option value="assessment|' + attr(assessment.id) + '" ' + (assessment.id === defaultAssessmentId ? "selected" : "") + '>' + esc(assessment.title) + ' · ' + esc(assessment.type || "Avaliação") + (assessment.date ? ' · ' + formatDate(assessment.date) : '') + '</option>';
        }).join("") + '</optgroup>';
      }).join("");
      var gradeLessonOptions = activeCourses().map(function (courseItem) {
        var items = semesterItems("lessons").filter(function (lessonItem) { return lessonItem.courseId === courseItem.id; }).sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
        if (!items.length) return "";
        return '<optgroup label="' + attr(courseItem.name) + ' · aulas">' + items.map(function (lessonItem) {
          return '<option value="lesson|' + attr(lessonItem.id) + '" ' + (preset.lessonId === lessonItem.id ? "selected" : "") + '>' + esc(lessonItem.title) + ' · ' + formatDate(lessonItem.date) + '</option>';
        }).join("") + '</optgroup>';
      }).join("");
      body = '<form id="entityForm" data-type="grade"><div class="form-grid"><div class="field field-full"><label>Avaliação ou aula a que pertence</label><select name="target" data-role="grade-target" required><option value="">Escolher origem concreta…</option>' + gradeAssessmentOptions + gradeLessonOptions + '</select><small>Uma nota nunca fica solta: escolhe Teste 1, Projeto, Exame ou a aula onde foi obtida.</small></div><div class="field"><label>Nota (0–20)</label><input name="score" type="number" min="0" max="20" step="0.1" required></div><div class="field"><label>Data</label><input name="date" type="date" value="' + todayISO() + '"></div>' + renderGradeDefenseFields(defaultAssessmentId) + '<div class="field field-full"><label>Notas</label><input name="notes" placeholder="Ex.: cotação, feedback do professor, tentativa…"></div></div><div class="form-note" style="margin-top:13px">As notas de avaliações entram no componente associado ao teste. Uma nota de aula só entra na média se a cadeira tiver uma componente do tipo “class”.</div></form>';
    } else if (type === "schedule") {
      title = "Adicionar aula ao horário";
      body = '<form id="entityForm" data-type="schedule"><div class="form-grid"><div class="field field-full"><label>Cadeira</label><select name="courseId" required><option value="">Escolher…</option>' + courseOptions(preset.courseId || "") + '</select></div><div class="field"><label>Dia</label><select name="weekday">' + WEEKDAYS.map(function (label, index) { return '<option value="' + index + '" ' + (Number(preset.weekday) === index ? "selected" : "") + '>' + label + "</option>"; }).join("") + '</select></div><div class="field"><label>Tipo</label><select name="lessonType"><option value="T">Teórica</option><option value="TP">Teórico-prática</option><option value="P">Prática</option><option value="LAB">Laboratório</option><option value="OT">Orientação</option></select></div><div class="field"><label>Início</label><input name="start" type="time" required value="' + attr(preset.start || "09:00") + '"></div><div class="field"><label>Fim</label><input name="end" type="time" required value="' + attr(preset.end || "10:30") + '"></div><div class="field field-full"><label>Sala</label><input name="room" placeholder="Ex.: B2.14" value="' + attr(preset.room) + '"></div></div></form>';
    } else if (type === "study-block") {
      var existingStudyBlock = preset.id ? state.studyBlocks.find(function (item) { return item.id === preset.id; }) : null;
      var sourceForStudy = !existingStudyBlock && preset.sourceType ? studySource(preset.sourceType, preset.sourceId) : null;
      var studyDate = (existingStudyBlock && existingStudyBlock.date) || preset.date || state.settings.studyPlanDate || todayISO();
      var studyDuration = sourceForStudy ? sourceForStudy.duration : Number(state.settings.studySessionMinutes || 50);
      var studyStart = (existingStudyBlock && existingStudyBlock.start) || preset.start || firstFreeStudyTime(studyDate, studyDuration);
      var studyEnd = (existingStudyBlock && existingStudyBlock.end) || minutesToTime(timeMinutes(studyStart) + studyDuration);
      var studyCourseId = (existingStudyBlock && existingStudyBlock.courseId) || (sourceForStudy && sourceForStudy.courseId) || preset.courseId || "";
      var studyKind = (existingStudyBlock && existingStudyBlock.kind) || preset.kind || "study";
      title = existingStudyBlock ? "Editar bloco de estudo" : "Novo bloco no dia";
      body = '<form id="entityForm" data-type="study-block" data-id="' + attr(existingStudyBlock && existingStudyBlock.id) + '" data-source-type="' + attr((existingStudyBlock && existingStudyBlock.sourceType) || preset.sourceType || "custom") + '" data-source-id="' + attr((existingStudyBlock && existingStudyBlock.sourceId) || preset.sourceId || "") + '"><div class="form-grid"><div class="field field-full"><label>Nome</label><input name="title" required value="' + attr((existingStudyBlock && existingStudyBlock.title) || (sourceForStudy && sourceForStudy.title) || "Sessão de estudo") + '"></div><div class="field"><label>Tipo</label><select name="kind"><option value="study" ' + (studyKind === "study" ? "selected" : "") + '>Estudo</option><option value="break" ' + (studyKind === "break" ? "selected" : "") + '>Pausa</option><option value="lunch" ' + (studyKind === "lunch" ? "selected" : "") + '>Almoço</option></select></div><div class="field"><label>Cadeira</label><select name="courseId"><option value="">Sem cadeira</option>' + courseOptions(studyCourseId) + '</select></div><div class="field"><label>Data</label><input name="date" type="date" required value="' + attr(studyDate) + '"></div><div class="field"><label>Início</label><input name="start" type="time" required value="' + attr(studyStart) + '"></div><div class="field"><label>Fim</label><input name="end" type="time" required value="' + attr(studyEnd) + '"></div><div class="field field-full"><label>Notas</label><textarea name="notes" placeholder="Objetivo, exercícios, capítulos…">' + esc(existingStudyBlock && existingStudyBlock.notes || "") + '</textarea></div></div></form>';
    } else if (type === "study-planner-settings") {
      title = "Configurar planeamento";
      body = '<form id="entityForm" data-type="study-planner-settings"><div class="form-grid"><div class="field"><label>Início do dia</label><input name="studyDayStart" type="time" required value="' + attr(state.settings.studyDayStart) + '"></div><div class="field"><label>Fim do dia</label><input name="studyDayEnd" type="time" required value="' + attr(state.settings.studyDayEnd) + '"></div><div class="field"><label>Duração de uma sessão</label><input name="studySessionMinutes" type="number" min="20" max="180" step="5" value="' + attr(state.settings.studySessionMinutes) + '"><small>Minutos</small></div><div class="field"><label>Pausa entre sessões</label><input name="studyBreakMinutes" type="number" min="0" max="60" step="5" value="' + attr(state.settings.studyBreakMinutes) + '"><small>Minutos</small></div><div class="field"><label>Hora de almoço</label><input name="studyLunchStart" type="time" value="' + attr(state.settings.studyLunchStart) + '"></div><div class="field"><label>Duração do almoço</label><input name="studyLunchMinutes" type="number" min="0" max="180" step="5" value="' + attr(state.settings.studyLunchMinutes) + '"><small>Minutos</small></div><div class="field field-full"><label>Horas de estudo por semana</label><input name="weeklyStudyHours" type="number" min="1" max="80" step="0.5" value="' + attr(state.settings.weeklyStudyHours) + '"><small>Usadas na estimativa por cadeira.</small></div></div></form>';
    } else if (type === "weekly-review") {
      var existingReview = weeklyReviewRecord();
      title = existingReview ? "Atualizar revisão semanal" : "Revisão semanal";
      body = '<form id="entityForm" data-type="weekly-review" data-id="' + attr(existingReview && existingReview.id) + '"><div class="form-grid"><div class="field field-full"><label>Prioridades da próxima semana</label><textarea name="priorities" placeholder="Uma prioridade por linha">' + esc(existingReview ? asArray(existingReview.priorities).join("\n") : "") + '</textarea></div><div class="field field-full"><label>Dúvidas a esclarecer</label><textarea name="doubts" placeholder="Uma dúvida por linha">' + esc(existingReview ? asArray(existingReview.doubts).join("\n") : "") + '</textarea><small>Podes levar esta lista às aulas ou ao horário de dúvidas.</small></div><div class="field field-full"><label>Notas da semana</label><textarea name="notes" placeholder="O que funcionou, o que ajustar…">' + esc(existingReview && existingReview.notes || "") + '</textarea></div></div></form>';
      submitLabel = "Concluir revisão";
    } else if (type === "past-exam-import") {
      title = "Importar teste anterior";
      var examCourse = preset.courseId || (activeCourses()[0] && activeCourses()[0].id) || "";
      body = '<form id="entityForm" data-type="past-exam-import"><div class="form-grid"><div class="field"><label>Cadeira</label><select name="courseId" required><option value="">Escolher…</option>' + courseOptions(examCourse) + '</select></div><div class="field"><label>Ano letivo</label><input name="academicYear" placeholder="2024/2025" value="' + attr(preset.academicYear || year) + '"></div><div class="field field-full"><label>Nome do teste</label><input name="title" required placeholder="Ex.: Teste 1 — época normal"></div><div class="field"><label>Data <span class="optional-label">opcional</span></label><input name="date" type="date"></div><div class="field"><label>Origem</label><input name="source" placeholder="PDF do professor, arquivo pessoal…"></div><div class="field field-full"><label>Ficheiro JSON <span class="optional-label">opcional</span></label><input name="jsonFile" data-role="local-json-file" data-target="pastExamJson" type="file" accept="application/json,.json"></div><div class="field field-full"><label>JSON das perguntas</label><textarea id="pastExamJson" name="json" class="json-editor" spellcheck="false" placeholder="Cola aqui o JSON gerado ou cria primeiro o teste vazio."></textarea>' + importTools("past-exam", "pastExamJson") + '<small>A importação é validada por inteiro antes de criar o teste. Os caminhos de imagem ficam guardados; também podes fazer upload ao editar cada pergunta.</small></div></div></form>';
      submitLabel = "Importar teste";
    } else if (type === "course-import") {
      title = "Importar cadeiras em JSON";
      body = '<form id="entityForm" data-type="course-import"><div class="form-grid"><div class="field field-full"><label>Ficheiro JSON</label><input name="jsonFile" data-role="local-json-file" data-target="courseImportJson" type="file" accept="application/json,.json"></div><div class="field field-full"><label>JSON das cadeiras</label><textarea id="courseImportJson" name="json" class="json-editor" spellcheck="false" placeholder="Cola uma cadeira ou um objeto com courses: […]."></textarea>' + importTools("course", "courseImportJson") + '<small>Cada cadeira é ligada ao semestre atual. O método de avaliação, mínimos e defesas são importados sem preencher campos em falta.</small></div></div></form>';
      submitLabel = "Importar cadeiras";
    } else if (type === "profile") {
      title = "Editar perfil académico";
      body = '<form id="entityForm" data-type="profile"><div class="form-grid"><div class="field field-full"><label>Nome</label><input name="name" required value="' + attr(state.profile.name) + '"></div><div class="field"><label>Instituição</label><input name="institution" value="' + attr(state.profile.institution) + '"></div><div class="field"><label>Curso</label><input name="degree" value="' + attr(state.profile.degree) + '"></div><div class="field"><label>Meta (0–20)</label><input name="targetGrade" type="number" min="0" max="20" step="0.1" value="' + attr(state.profile.targetGrade || 20) + '"></div></div></form>';
    } else if (type === "lesson-notes") {
      var notesLesson = lessonById(preset.id);
      title = "Apontamentos da aula";
      body = '<form id="entityForm" data-type="lesson-notes" data-id="' + attr(preset.id) + '"><div class="field"><label>Notas</label><textarea name="notes" style="min-height:260px" placeholder="Resumo, dúvidas, referências…">' + esc(notesLesson && notesLesson.notes) + "</textarea></div></form>";
    } else {
      toast("Este formulário ainda não está disponível.", "warning");
      return;
    }
    openModal(title, body, { footer: formFooter(submitLabel) });
    if (type === "study-block" && existingStudyBlock) {
      var blockForm = modalRoot.querySelector('#entityForm[data-type="study-block"]');
      if (blockForm) blockForm.insertAdjacentHTML("beforeend", '<button class="button button-danger" style="margin-top:14px" type="button" data-action="delete-entity" data-kind="studyBlocks" data-id="' + attr(existingStudyBlock.id) + '"><i data-lucide="trash-2"></i>Remover bloco</button>');
      refreshIcons(blockForm);
    }
    if (type === "course") updateEvaluationBuilderSummary(modalRoot.querySelector('#entityForm[data-type="course"]'));
    if (type === "quiz") updateQuizPastQuestionPicker(preset.lessonId || "");
    if (type === "lesson") updateLessonScheduleOptions(modalRoot.querySelector('#entityForm[data-type="lesson"]'), selectedScheduleId || "");
    if (type === "assessment") updateAssessmentLinkedOptions(modalRoot.querySelector('#entityForm[data-type="assessment"]'));
  }

  function renderPastQuestionChoices(questions, name) {
    if (!questions.length) return '<div class="past-question-empty"><i data-lucide="message-circle-question"></i><span>Esta aula ainda não tem perguntas de testes anteriores.</span></div>';
    return questions.map(function (question) {
      return '<label class="past-question-option"><input type="checkbox" name="' + attr(name || "pastQuestionIds") + '" value="' + attr(question.id) + '"><span><strong>' + esc(question.prompt) + '</strong><small>' + esc([question.assessmentLabel, question.academicYear].filter(Boolean).join(" · ") || "Teste anterior") + '</small></span></label>';
    }).join("");
  }

  function updateQuizPastQuestionPicker(lessonId) {
    var picker = modalRoot.querySelector("#quizPastQuestionPicker");
    if (!picker) return;
    if (!lessonId) {
      picker.innerHTML = '<div class="past-question-empty"><i data-lucide="mouse-pointer-2"></i><span>Escolhe uma aula para veres as perguntas disponíveis.</span></div>';
    } else {
      picker.innerHTML = renderPastQuestionChoices(pastQuestionsForLesson(lessonId), "pastQuestionIds");
    }
    refreshIcons();
  }

  function openPastQuestionPicker(quizId) {
    var quiz = state.quizzes.find(function (item) { return item.id === quizId; });
    if (!quiz) return;
    if (!quiz.lessonId) {
      toast("Associa primeiro o quiz a uma aula para usar perguntas anteriores.", "warning");
      return;
    }
    var used = asArray(quiz.questions).map(function (question) { return question.sourceQuestionId; }).filter(Boolean);
    var available = pastQuestionsForLesson(quiz.lessonId).filter(function (question) { return used.indexOf(question.id) < 0; });
    var lesson = lessonById(quiz.lessonId);
    var body = '<form id="pastQuestionForm" data-quiz-id="' + attr(quiz.id) + '"><p class="onboarding-copy" style="margin-top:0">' + esc(lesson ? lesson.title : "Aula") + ' · escolhe as perguntas reais que queres juntar ao quiz.</p><div class="past-question-picker">' + (available.length ? renderPastQuestionChoices(available, "pastQuestionIds") : '<div class="past-question-empty"><i data-lucide="check-check"></i><span>Todas as perguntas anteriores desta aula já estão no quiz.</span></div>') + '</div></form>';
    openModal("Adicionar perguntas anteriores", body, { footer: available.length ? formFooter("Adicionar selecionadas") : '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Fechar</button></footer>' });
  }

  async function handlePastQuestionSubmit(event) {
    event.preventDefault();
    var form = event.target;
    var quiz = state.quizzes.find(function (item) { return item.id === form.dataset.quizId; });
    if (!quiz) return;
    var ids = new FormData(form).getAll("pastQuestionIds");
    if (!ids.length) {
      setFormError(form, "Escolhe pelo menos uma pergunta.");
      return;
    }
    var used = asArray(quiz.questions).map(function (question) { return question.sourceQuestionId; });
    var additions = state.questions.filter(function (question) {
      return ids.indexOf(question.id) >= 0 && used.indexOf(question.id) < 0 && asArray(question.lessonIds).indexOf(quiz.lessonId) >= 0;
    }).map(quizQuestionFromPast);
    quiz.questions = asArray(quiz.questions).concat(additions);
    await save(true);
    closeModal();
    render();
    toast(additions.length + " pergunta(s) anterior(es) adicionada(s) ao quiz.");
  }

  function parseEvaluation(text, existing) {
    var prior = asArray(existing);
    var used = {};
    return String(text || "").split(/\n+/).map(function (line) {
      var parts = line.split("|").map(function (part) { return part.trim(); });
      if (!parts[0]) return null;
      var label = parts[0];
      var weight = clamp(String(parts[1] || "0").replace(",", "."), 0, 100);
      var lower = (parts[2] || label).toLowerCase();
      var kind = lower.indexOf("exam") >= 0 ? "exam" : lower.indexOf("test") >= 0 ? "test" : lower.indexOf("proj") >= 0 ? "project" : lower.indexOf("apresent") >= 0 ? "presentation" : lower.indexOf("aula") >= 0 || lower.indexOf("class") >= 0 ? "class" : "other";
      var match = prior.find(function (item) { return !used[item.id] && String(item.label).toLowerCase() === label.toLowerCase(); });
      var id = match ? match.id : uid("component");
      used[id] = true;
      return { id: id, label: label, weight: weight, kind: kind, replaceable: kind === "test" };
    }).filter(Boolean);
  }

  function selectedValues(select) {
    if (!select) return [];
    return Array.from(select.selectedOptions || []).map(function (option) { return option.value; }).filter(Boolean);
  }

  function setFormError(form, message) {
    var existing = form.querySelector(".form-error");
    if (existing) existing.remove();
    var error = document.createElement("p");
    error.className = "form-error";
    error.textContent = message;
    form.appendChild(error);
  }

  function safeResourceUrl(value) {
    var url = String(value || "").trim();
    if (!url) return "";
    if (/^(javascript|data|vbscript):/i.test(url)) return "";
    return url;
  }

  function normalizeImageRefs(value, fallbackRole) {
    var refs = [];
    var fallback = fallbackRole || "question";
    if (value && !Array.isArray(value) && typeof value === "object" && !value.url && !value.path && !value.blobId) {
      ["question", "solution", "explanation", "event"].forEach(function (role) {
        asArray(value[role]).forEach(function (entry) {
          if (typeof entry === "string") refs.push({ id: uid("image"), role: role, source: "path", url: safeResourceUrl(entry), name: entry.split("/").pop(), caption: "" });
          else if (entry && typeof entry === "object") refs.push(Object.assign({ role: role }, entry));
        });
      });
    } else {
      asArray(value).forEach(function (entry) {
        if (typeof entry === "string") refs.push({ id: uid("image"), role: fallback, source: "path", url: safeResourceUrl(entry), name: entry.split("/").pop(), caption: "" });
        else if (entry && typeof entry === "object") refs.push(entry);
      });
    }
    return refs.map(function (entry) {
      var url = safeResourceUrl(entry.url || entry.path || "");
      return {
        id: entry.id || uid("image"),
        role: ["question", "solution", "explanation", "event"].indexOf(entry.role) >= 0 ? entry.role : fallback,
        source: entry.blobId ? "indexeddb" : (entry.source === "indexeddb" ? "indexeddb" : "path"),
        blobId: entry.blobId || "",
        url: entry.blobId ? "" : url,
        name: entry.name || (url ? url.split("/").pop() : "imagem"),
        caption: String(entry.caption || "").trim(),
        mimeType: entry.mimeType || ""
      };
    }).filter(function (entry) { return entry.blobId || entry.url; });
  }

  function imageRoleLabel(role) {
    if (role === "solution") return "Solução";
    if (role === "explanation") return "Explicação";
    if (role === "event") return "Evento";
    return "Enunciado";
  }

  function renderImageGallery(images, role, options) {
    options = options || {};
    var refs = normalizeImageRefs(images).filter(function (image) { return !role || image.role === role; });
    if (!refs.length) return "";
    return '<div class="media-gallery ' + (options.compact ? "is-compact" : "") + '">' + refs.map(function (image) {
      var source = image.source === "indexeddb"
        ? 'data-local-image-id="' + attr(image.blobId) + '" src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw=="'
        : 'src="' + attr(safeResourceUrl(image.url)) + '"';
      var caption = image.caption || (options.showRole ? imageRoleLabel(image.role) : "");
      return '<figure><button type="button" class="media-open" data-action="open-image" data-image-id="' + attr(image.id) + '" data-image-owner="' + attr(options.ownerId || "") + '" aria-label="Abrir imagem"><img ' + source + ' alt="' + attr(image.name || imageRoleLabel(image.role)) + '" loading="lazy"></button>' + (caption ? '<figcaption>' + esc(caption) + '</figcaption>' : '') + '</figure>';
    }).join("") + '</div>';
  }

  function renderExistingImageManager(images) {
    var refs = normalizeImageRefs(images);
    if (!refs.length) return "";
    return '<div class="existing-image-manager">' + refs.map(function (image) {
      return '<label class="existing-image-item"><span>' + esc(imageRoleLabel(image.role)) + ' · ' + esc(image.name || "imagem") + '</span><input type="checkbox" name="removeImageIds" value="' + attr(image.id) + '"> Remover</label>';
    }).join("") + '</div>';
  }

  function revokeImageObjectUrls() {
    activeImageObjectUrls.forEach(function (url) { try { URL.revokeObjectURL(url); } catch (_) { /* sem efeito */ } });
    activeImageObjectUrls = [];
  }

  function hydrateLocalImages(root) {
    if (!root) return;
    Array.from(root.querySelectorAll("img[data-local-image-id]")).forEach(function (image) {
      var id = image.getAttribute("data-local-image-id");
      if (!id || image.dataset.hydrated === "true") return;
      image.dataset.hydrated = "loading";
      DB.getFile(id).then(function (record) {
        if (!record || !record.blob || !image.isConnected) return;
        var url = URL.createObjectURL(record.blob);
        activeImageObjectUrls.push(url);
        image.src = url;
        image.dataset.hydrated = "true";
      }).catch(function () { image.dataset.hydrated = "error"; });
    });
  }

  function imagePathRefs(raw, role) {
    return String(raw || "").split(/\n+/).map(function (line) { return safeResourceUrl(line.trim()); }).filter(Boolean).map(function (url) {
      return { id: uid("image"), role: role, source: "path", url: url, name: url.split("/").pop(), caption: "" };
    });
  }

  async function storeImageFiles(data, fieldName, role, metadata) {
    var files = data.getAll(fieldName).filter(function (file) { return file && file.size; });
    var refs = [];
    for (var index = 0; index < files.length; index += 1) {
      var file = files[index];
      if (!/^image\//.test(file.type || "")) throw new Error("O ficheiro “" + file.name + "” não é uma imagem.");
      var blobId = await DB.putFile(file, Object.assign({ imageRole: role }, metadata || {}));
      refs.push({ id: uid("image"), role: role, source: "indexeddb", blobId: blobId, url: "", name: file.name, caption: "", mimeType: file.type });
    }
    return refs;
  }

  async function collectImagesFromForm(form, existing, groups, metadata) {
    var data = new FormData(form);
    var removed = data.getAll("removeImageIds");
    var original = normalizeImageRefs(existing);
    for (var removedIndex = 0; removedIndex < original.length; removedIndex += 1) {
      if (removed.indexOf(original[removedIndex].id) >= 0 && original[removedIndex].blobId) await DB.deleteFile(original[removedIndex].blobId);
    }
    var refs = original.filter(function (image) { return removed.indexOf(image.id) < 0; });
    for (var index = 0; index < groups.length; index += 1) {
      var group = groups[index];
      refs = refs.concat(await storeImageFiles(data, group.files, group.role, metadata));
      refs = refs.concat(imagePathRefs(data.get(group.paths), group.role));
    }
    return refs;
  }

  function parseJSONText(raw, label) {
    var text = String(raw || "").trim();
    if (!text) return null;
    try { return JSON.parse(text); }
    catch (_) { throw new Error((label || "O conteúdo") + " não contém JSON válido."); }
  }

  function importLessonIds(courseId, titles) {
    var wanted = asArray(titles).map(function (title) { return cleanText(title).toLocaleLowerCase("pt-PT"); }).filter(Boolean);
    if (!wanted.length) return [];
    return state.lessons.filter(function (lesson) {
      return lesson.courseId === courseId && wanted.indexOf(cleanText(lesson.title).toLocaleLowerCase("pt-PT")) >= 0;
    }).map(function (lesson) { return lesson.id; });
  }

  function importedCourseRecord(source, index) {
    if (!source || typeof source !== "object") throw new Error("A cadeira " + (index + 1) + " não é um objeto válido.");
    var name = String(source.name || "").trim();
    if (!name) throw new Error("A cadeira " + (index + 1) + " não tem nome.");
    var allowedLessonTypes = ["T", "TP", "P", "LAB", "OT"];
    var lessonTypes = asArray(source.lessonTypes).filter(function (type) { return allowedLessonTypes.indexOf(type) >= 0; });
    var evaluation = source.evaluation && typeof source.evaluation === "object" ? source.evaluation : {};
    var allowedKinds = ["test", "project", "exam", "presentation", "class", "other"];
    var components = asArray(evaluation.components).map(function (component, componentIndex) {
      if (!component || typeof component !== "object" || !String(component.label || "").trim()) throw new Error("A componente " + (componentIndex + 1) + " de “" + name + "” não tem nome.");
      var kind = allowedKinds.indexOf(component.kind) >= 0 ? component.kind : "other";
      return {
        id: uid("component"),
        label: String(component.label).trim(),
        kind: kind,
        count: Math.max(1, Number(component.count) || 1),
        weight: clamp(component.weight, 0, 100),
        minimum: nullableNumber(component.minimum),
        defenseEnabled: component.defenseEnabled === true,
        defenseType: ["oral", "practical", "oral-practical"].indexOf(component.defenseType) >= 0 ? component.defenseType : "oral",
        defenseThreshold: nullableNumber(component.defenseThreshold),
        maxWithoutDefense: nullableNumber(component.maxWithoutDefense),
        replaceable: kind === "test"
      };
    });
    var total = components.reduce(function (sum, component) { return sum + component.weight; }, 0);
    if (total > 100.01) throw new Error("As percentagens de “" + name + "” somam mais de 100%.");
    return {
      id: uid("course"), semesterId: state.currentSemesterId, name: name,
      code: String(source.code || "").trim(), ects: clamp(source.ects, 0, 60),
      color: safeColor(source.color, COLORS[(activeCourses().length + index) % COLORS.length]),
      lessonTypes: lessonTypes,
      evaluation: { components: components, examReplacesTests: evaluation.examReplacesTests === true, replacementPolicy: evaluation.replacementPolicy === "always" ? "always" : "if-higher" }
    };
  }

  async function handleEntitySubmit(event) {
    event.preventDefault();
    var form = event.target;
    if (form.id !== "entityForm") return;
    var type = form.getAttribute("data-type");
    var id = form.getAttribute("data-id");
    var data = new FormData(form);
    var submit = modalRoot.querySelector('[type="submit"]');
    var postSaveMessage = "Conteúdo guardado.";
    if (submit) submit.disabled = true;

    try {
      if (type === "course") {
        if (!state.currentSemesterId) throw new Error("Cria primeiro um semestre ativo.");
        var existingCourse = id ? courseById(id) : null;
        var components = readEvaluationBuilder(form);
        if (!components.length) throw new Error("Adiciona pelo menos uma componente ao método de avaliação.");
        var totalEvaluationWeight = components.reduce(function (sum, component) { return sum + (Number(component.weight) || 0); }, 0);
        if (totalEvaluationWeight > 100.01) throw new Error("As percentagens do método de avaliação não podem ultrapassar 100%.");
        var removedInUse = existingCourse && state.assessments.some(function (assessment) { return assessment.courseId === existingCourse.id && !components.some(function (component) { return component.id === assessment.componentId; }); });
        if (removedInUse) throw new Error("Não podes remover uma componente que já tem avaliações. Altera primeiro essas avaliações.");
        var courseData = Object.assign(existingCourse || {}, {
          id: existingCourse ? existingCourse.id : uid("course"),
          semesterId: existingCourse ? existingCourse.semesterId : state.currentSemesterId,
          name: String(data.get("name") || "").trim(),
          code: String(data.get("code") || "").trim(),
          ects: Number(data.get("ects")) || 0,
          color: safeColor(data.get("color"), COLORS[activeCourses().length % COLORS.length]),
          lessonTypes: data.getAll("lessonTypes"),
          evaluation: {
            components: components,
            examReplacesTests: false,
            replacementPolicy: "if-higher"
          }
        });
        if (!courseData.name) throw new Error("Escreve o nome da cadeira.");
        if (!existingCourse) state.courses.push(courseData);
        postSaveMessage = existingCourse ? "Cadeira atualizada." : "Cadeira adicionada.";
      } else if (type === "course-import") {
        if (!state.currentSemesterId) throw new Error("Cria primeiro um semestre ativo.");
        var coursePayload = parseJSONText(data.get("json"), "A importação de cadeiras");
        if (!coursePayload) throw new Error("Cola o JSON ou escolhe um ficheiro.");
        var importedCourses = Array.isArray(coursePayload) ? coursePayload : asArray(coursePayload.courses).length ? coursePayload.courses : coursePayload.name ? [coursePayload] : [];
        if (!importedCourses.length) throw new Error("O JSON não contém nenhuma cadeira.");
        var courseRecords = importedCourses.map(importedCourseRecord);
        state.courses = state.courses.concat(courseRecords);
        postSaveMessage = courseRecords.length + " cadeira(s) importada(s).";
      } else if (type === "lesson") {
        var course = courseById(data.get("courseId"));
        if (!course) throw new Error("Escolhe uma cadeira.");
        var lessonSchedule = scheduleById(data.get("scheduleId"));
        if (!lessonSchedule) throw new Error("Escolhe um bloco compatível do horário.");
        if (lessonSchedule.courseId !== course.id) throw new Error("O bloco escolhido pertence a outra cadeira.");
        if (String(lessonSchedule.type || "") !== String(data.get("lessonType") || "")) throw new Error("O tipo da aula tem de ser igual ao tipo do bloco do horário.");
        if (!scheduleMatchesDate(lessonSchedule, data.get("date"))) throw new Error("A data escolhida não corresponde ao dia desse bloco do horário.");
        var existingLesson = id ? lessonById(id) : null;
        var duplicateLesson = semesterItems("lessons").find(function (item) {
          return item.id !== (existingLesson && existingLesson.id) && item.date === data.get("date") && lessonMatchesSchedule(item, lessonSchedule);
        });
        if (duplicateLesson) throw new Error("Esse período do horário já está ligado à aula “" + duplicateLesson.title + "”.");
        var lessonData = Object.assign(existingLesson || {}, {
          id: existingLesson ? existingLesson.id : uid("lesson"),
          semesterId: course.semesterId,
          courseId: course.id,
          scheduleId: lessonSchedule.id,
          title: String(data.get("title") || "").trim(),
          date: data.get("date"),
          start: lessonSchedule.start,
          end: lessonSchedule.end,
          type: lessonSchedule.type,
          room: String(data.get("room") || lessonSchedule.room || "").trim(),
          topics: String(data.get("topics") || "").trim(),
          notes: existingLesson ? existingLesson.notes || "" : "",
          mastered: existingLesson ? !!existingLesson.mastered : false
        });
        if (!lessonData.title || !lessonData.date) throw new Error("Preenche o título e a data da aula.");
        if (!existingLesson) state.lessons.push(lessonData);
        var lessonFile = data.get("file");
        if (lessonFile && lessonFile.size) {
          var uploadedLessonFile = await uploadMaterialFile(lessonFile, { courseId: course.id, lessonId: lessonData.id });
          state.materials.push({ id: uploadedLessonFile.id, semesterId: course.semesterId, courseId: course.id, lessonId: lessonData.id, title: "Slides · " + lessonData.title, academicYear: String(data.get("materialYear") || (semesterById(course.semesterId) || {}).academicYear || ""), kind: "slides", source: "remote", blobId: uploadedLessonFile.blobId, remoteFile: uploadedLessonFile.remoteFile, slides: uploadedLessonFile.slides, slideCount: uploadedLessonFile.slideCount, fileName: lessonFile.name, mimeType: lessonFile.type, uploadedAt: new Date().toISOString() });
        }
      } else if (type === "material") {
        var linkedLesson = lessonById(data.get("lessonId"));
        var materialCourse = courseById(linkedLesson ? linkedLesson.courseId : data.get("courseId"));
        if (!materialCourse) throw new Error("Escolhe uma cadeira.");
        var file = data.get("file");
        var url = safeResourceUrl(data.get("url"));
        if ((!file || !file.size) && !url) throw new Error("Escolhe um ficheiro ou indica um caminho/URL.");
        var uploadedMaterial = file && file.size ? await uploadMaterialFile(file, { courseId: materialCourse.id, lessonId: linkedLesson && linkedLesson.id }) : null;
        state.materials.push({
          id: uploadedMaterial ? uploadedMaterial.id : uid("material"), semesterId: materialCourse.semesterId, courseId: materialCourse.id,
          lessonId: linkedLesson ? linkedLesson.id : null, title: String(data.get("title") || "").trim(),
          academicYear: String(data.get("academicYear") || "").trim(), kind: data.get("kind") || "slides",
          source: uploadedMaterial ? "remote" : "url", blobId: uploadedMaterial ? uploadedMaterial.blobId : null,
          remoteFile: uploadedMaterial ? uploadedMaterial.remoteFile : null, slides: uploadedMaterial ? uploadedMaterial.slides : [], slideCount: uploadedMaterial ? uploadedMaterial.slideCount : 0,
          url: uploadedMaterial ? "" : url, fileName: uploadedMaterial ? file.name : url.split("/").pop(), mimeType: uploadedMaterial ? file.type : "", uploadedAt: new Date().toISOString()
        });
      } else if (type === "task") {
        state.tasks.push({ id: uid("task"), semesterId: state.currentSemesterId, courseId: data.get("courseId") || null, lessonId: data.get("lessonId") || null, title: String(data.get("title") || "").trim(), type: data.get("taskType") || "homework", dueDate: data.get("dueDate") || "", dueTime: data.get("dueTime") || "", priority: data.get("priority") || "normal", done: false, createdAt: new Date().toISOString() });
      } else if (type === "assessment") {
        var assessmentCourse = courseById(data.get("courseId"));
        if (!assessmentCourse) throw new Error("Escolhe uma cadeira.");
        var assessmentComponentId = String(data.get("componentId") || "");
        if (!asArray(assessmentCourse.evaluation && assessmentCourse.evaluation.components).some(function (component) { return component.id === assessmentComponentId; })) throw new Error("Escolhe uma componente válida do método de avaliação.");
        var existingAssessment = id ? assessmentById(id) : null;
        var replacementAssessmentIds = selectedValues(form.elements.replacementAssessmentIds).filter(function (assessmentId) { var target = assessmentById(assessmentId); return target && target.courseId === assessmentCourse.id && target.id !== (existingAssessment && existingAssessment.id); });
        var assessmentData = Object.assign(existingAssessment || {}, {
          id: existingAssessment ? existingAssessment.id : uid("assessment"),
          semesterId: assessmentCourse.semesterId,
          courseId: assessmentCourse.id,
          componentId: assessmentComponentId,
          type: data.get("assessmentType"),
          title: String(data.get("title") || "").trim(),
          date: data.get("date") || "",
          time: data.get("time") || "",
          location: String(data.get("location") || "").trim(),
          weight: Number(data.get("weight")) || 0,
          lessonIds: selectedValues(form.elements.lessonIds).filter(function (lessonId) { var lesson = lessonById(lessonId); return lesson && lesson.courseId === assessmentCourse.id; }),
          requiresTestSheet: data.get("requiresTestSheet") === "on",
          openBook: data.get("openBook") === "on",
          hasDefense: data.get("hasDefense") === "on",
          defenseType: data.get("defenseType") || "oral",
          defenseThreshold: nullableNumber(data.get("defenseThreshold")),
          maxWithoutDefense: nullableNumber(data.get("maxWithoutDefense")),
          replacementAssessmentIds: replacementAssessmentIds,
          replacementPolicy: data.get("replacementPolicy") === "always" ? "always" : "if-higher",
          replacesTests: replacementAssessmentIds.length > 0
        });
        if (!assessmentData.title) throw new Error("Escreve o nome da avaliação.");
        if (!existingAssessment) state.assessments.push(assessmentData);
        state.grades.forEach(function (grade) {
          if (grade.assessmentId === assessmentData.id) {
            grade.semesterId = assessmentData.semesterId;
            grade.courseId = assessmentData.courseId;
            grade.componentId = assessmentData.componentId;
          }
        });
      } else if (type === "event") {
        var existingEvent = id ? state.events.find(function (item) { return item.id === id; }) : null;
        var eventImages = await collectImagesFromForm(form, existingEvent && existingEvent.images, [{ role: "event", files: "eventImageFiles", paths: "eventImagePaths" }], { entityType: "event", entityId: id || "new" });
        var savedEvent = Object.assign(existingEvent || {}, { id: existingEvent ? existingEvent.id : uid("event"), semesterId: existingEvent ? existingEvent.semesterId : state.currentSemesterId, title: String(data.get("title") || "").trim(), date: data.get("date"), time: data.get("time") || "", location: String(data.get("location") || "").trim(), url: safeResourceUrl(data.get("url")), notes: String(data.get("notes") || "").trim(), images: eventImages });
        if (!savedEvent.title || !savedEvent.date) throw new Error("Preenche o nome e a data do evento.");
        if (!existingEvent) state.events.push(savedEvent);
      } else if (type === "question") {
        var questionCourse = courseById(data.get("courseId"));
        if (!questionCourse) throw new Error("Escolhe uma cadeira.");
        var existingQuestion = id ? state.questions.find(function (item) { return item.id === id; }) : null;
        var selectedPastExam = pastExamById(data.get("pastExamId"));
        if (selectedPastExam && selectedPastExam.courseId !== questionCourse.id) throw new Error("O teste anterior pertence a outra cadeira.");
        var questionImages = await collectImagesFromForm(form, existingQuestion && existingQuestion.images, [
          { role: "question", files: "questionImageFiles", paths: "questionImagePaths" },
          { role: "solution", files: "solutionImageFiles", paths: "solutionImagePaths" },
          { role: "explanation", files: "explanationImageFiles", paths: "explanationImagePaths" }
        ], { entityType: "question", entityId: id || "new", courseId: questionCourse.id });
        var savedQuestion = Object.assign(existingQuestion || {}, { id: existingQuestion ? existingQuestion.id : uid("question"), semesterId: questionCourse.semesterId, courseId: questionCourse.id, pastExamId: selectedPastExam ? selectedPastExam.id : null, lessonIds: selectedValues(form.elements.lessonIds).filter(function (lessonId) { var lesson = lessonById(lessonId); return lesson && lesson.courseId === questionCourse.id; }), academicYear: String(data.get("academicYear") || "").trim(), number: String(data.get("number") || "").trim(), sourceType: "past-test", assessmentLabel: String(data.get("assessmentLabel") || "").trim(), prompt: String(data.get("prompt") || "").trim(), answer: String(data.get("answer") || "").trim(), explanation: String(data.get("explanation") || "").trim(), points: nullableNumber(data.get("points")), tags: asArray(existingQuestion && existingQuestion.tags), images: questionImages });
        if (!savedQuestion.prompt) throw new Error("Escreve a pergunta tal como apareceu.");
        if (!existingQuestion) state.questions.push(savedQuestion);
      } else if (type === "quiz" || type === "quiz-question") {
        var options = [0, 1, 2, 3].map(function (index) { return String(data.get("option" + index) || "").trim(); }).filter(Boolean);
        var manualPrompt = String(data.get("prompt") || "").trim();
        var questionData = null;
        if (manualPrompt) {
          if (options.length < 2) throw new Error("A pergunta manual precisa de pelo menos duas opções.");
          questionData = { id: uid("quizq"), mode: "multiple-choice", prompt: manualPrompt, options: options, answerIndex: clamp(data.get("answerIndex"), 0, options.length - 1), explanation: String(data.get("explanation") || "").trim() };
        } else if (type === "quiz-question") {
          throw new Error("Escreve a pergunta manual.");
        } else if (options.length) {
          throw new Error("Escreve a pergunta manual ou deixa as opções vazias.");
        }
        if (type === "quiz") {
          var quizCourse = courseById(data.get("courseId"));
          if (!quizCourse) throw new Error("Escolhe uma cadeira.");
          var quizLesson = lessonById(data.get("lessonId"));
          if (quizLesson && quizLesson.courseId !== quizCourse.id) quizLesson = null;
          var selectedPastIds = data.getAll("pastQuestionIds");
          var selectedPast = quizLesson ? pastQuestionsForLesson(quizLesson.id).filter(function (question) { return selectedPastIds.indexOf(question.id) >= 0; }) : [];
          var quizQuestions = (questionData ? [questionData] : []).concat(selectedPast.map(quizQuestionFromPast));
          if (!quizQuestions.length) throw new Error("Adiciona uma pergunta manual ou escolhe uma pergunta de teste anterior.");
          state.quizzes.push({ id: uid("quiz"), semesterId: quizCourse.semesterId, courseId: quizCourse.id, lessonId: quizLesson ? quizLesson.id : null, title: String(data.get("title") || "").trim(), questions: quizQuestions, generatedFromPastQuestions: selectedPast.length > 0 && !questionData, createdAt: new Date().toISOString(), lastScore: null });
        } else {
          var quiz = state.quizzes.find(function (item) { return item.id === id; });
          if (!quiz) throw new Error("Quiz não encontrado.");
          quiz.questions = asArray(quiz.questions);
          quiz.questions.push(questionData);
        }
      } else if (type === "grade") {
        var target = String(data.get("target") || "").split("|");
        var gradeKind = target[0];
        var gradeTargetId = target[1];
        var gradeCourse;
        var gradeData;
        if (gradeKind === "assessment") {
          var gradeAssessment = assessmentById(gradeTargetId);
          if (!gradeAssessment) throw new Error("Escolhe uma avaliação concreta.");
          gradeCourse = courseById(gradeAssessment.courseId);
          var gradeComponentId = suggestedComponentId(gradeCourse, gradeAssessment);
          if (!gradeComponentId) throw new Error("Esta avaliação ainda não está ligada a uma componente válida. Edita ou recria a avaliação depois de configurares o método da cadeira.");
          gradeAssessment.componentId = gradeComponentId;
          gradeData = { assessmentId: gradeAssessment.id, componentId: gradeComponentId, lessonId: null };
        } else if (gradeKind === "lesson") {
          var gradeLesson = lessonById(gradeTargetId);
          if (!gradeLesson) throw new Error("Escolhe a aula onde recebeste a nota.");
          gradeCourse = courseById(gradeLesson.courseId);
          var classComponents = asArray(gradeCourse && gradeCourse.evaluation && gradeCourse.evaluation.components).filter(function (component) { return component.kind === "class"; });
          gradeData = { assessmentId: null, componentId: classComponents.length === 1 ? classComponents[0].id : "class-note", lessonId: gradeLesson.id };
        } else {
          throw new Error("Escolhe o teste, projeto, exame ou aula desta nota.");
        }
        if (!gradeCourse) throw new Error("A origem da nota não pertence a uma cadeira válida.");
        var submittedDefenseStatus = data.get("defenseStatus") || "not-applicable";
        var submittedDefenseScore = nullableNumber(data.get("defenseFinalScore"));
        if (submittedDefenseStatus === "completed" && submittedDefenseScore == null) throw new Error("Indica a nota final obtida após a defesa.");
        state.grades.push(Object.assign({
          id: uid("grade"),
          semesterId: gradeCourse.semesterId,
          courseId: gradeCourse.id,
          score: clamp(data.get("score"), 0, 20),
          date: data.get("date") || todayISO(),
          notes: String(data.get("notes") || "").trim(),
          defenseStatus: submittedDefenseStatus,
          defenseType: data.get("gradeDefenseType") || "",
          defenseFinalScore: submittedDefenseScore
        }, gradeData));
      } else if (type === "schedule") {
        var scheduleCourse = courseById(data.get("courseId"));
        if (!scheduleCourse) throw new Error("Escolhe uma cadeira.");
        if (asArray(scheduleCourse.lessonTypes).length && asArray(scheduleCourse.lessonTypes).indexOf(data.get("lessonType")) < 0) throw new Error("Esse tipo de aula não está ativado nesta cadeira.");
        if (timeMinutes(data.get("end")) <= timeMinutes(data.get("start"))) throw new Error("A hora de fim deve ser posterior ao início.");
        state.schedule.push({ id: uid("schedule"), semesterId: scheduleCourse.semesterId, courseId: scheduleCourse.id, weekday: Number(data.get("weekday")), start: data.get("start"), end: data.get("end"), type: data.get("lessonType"), room: String(data.get("room") || "").trim() });
      } else if (type === "study-block") {
        if (timeMinutes(data.get("end")) <= timeMinutes(data.get("start"))) throw new Error("A hora de fim deve ser posterior ao início.");
        var existingBlock = id ? state.studyBlocks.find(function (item) { return item.id === id; }) : null;
        var overlappingBlock = semesterItems("studyBlocks").find(function (item) { return item.id !== (existingBlock && existingBlock.id) && item.date === data.get("date") && timeMinutes(data.get("start")) < timeMinutes(item.end) && timeMinutes(data.get("end")) > timeMinutes(item.start); });
        if (overlappingBlock) throw new Error("Este horário sobrepõe-se a “" + overlappingBlock.title + "”.");
        var studyCourse = data.get("courseId") ? courseById(data.get("courseId")) : null;
        var savedBlock = Object.assign(existingBlock || {}, {
          id: existingBlock ? existingBlock.id : uid("studyblock"), semesterId: existingBlock ? existingBlock.semesterId : state.currentSemesterId,
          date: data.get("date"), title: String(data.get("title") || "").trim(), start: data.get("start"), end: data.get("end"),
          kind: ["study", "break", "lunch"].indexOf(data.get("kind")) >= 0 ? data.get("kind") : "study", courseId: studyCourse ? studyCourse.id : null,
          sourceType: form.dataset.sourceType || (existingBlock && existingBlock.sourceType) || "custom", sourceId: form.dataset.sourceId || (existingBlock && existingBlock.sourceId) || null,
          completed: existingBlock ? !!existingBlock.completed : false, notes: String(data.get("notes") || "").trim()
        });
        if (!savedBlock.title) throw new Error("Escreve o nome do bloco.");
        if (!existingBlock) state.studyBlocks.push(savedBlock);
        state.settings.studyPlanDate = savedBlock.date;
        postSaveMessage = existingBlock ? "Bloco atualizado." : "Bloco adicionado ao plano.";
      } else if (type === "study-planner-settings") {
        if (timeMinutes(data.get("studyDayEnd")) <= timeMinutes(data.get("studyDayStart"))) throw new Error("O fim do dia deve ser posterior ao início.");
        state.settings.studyDayStart = data.get("studyDayStart");
        state.settings.studyDayEnd = data.get("studyDayEnd");
        state.settings.studySessionMinutes = clamp(data.get("studySessionMinutes"), 20, 180);
        state.settings.studyBreakMinutes = clamp(data.get("studyBreakMinutes"), 0, 60);
        state.settings.studyLunchStart = data.get("studyLunchStart") || "13:00";
        state.settings.studyLunchMinutes = clamp(data.get("studyLunchMinutes"), 0, 180);
        state.settings.weeklyStudyHours = clamp(data.get("weeklyStudyHours"), 1, 80);
        postSaveMessage = "Planeamento atualizado.";
      } else if (type === "weekly-review") {
        var existingWeeklyReview = id ? state.weeklyReviews.find(function (item) { return item.id === id; }) : weeklyReviewRecord();
        var lineValues = function (value) { return String(value || "").split(/\n+/).map(function (line) { return line.trim(); }).filter(Boolean); };
        var reviewData = Object.assign(existingWeeklyReview || {}, {
          id: existingWeeklyReview ? existingWeeklyReview.id : uid("review"), semesterId: state.currentSemesterId,
          weekStart: currentWeekStart(), priorities: lineValues(data.get("priorities")), doubts: lineValues(data.get("doubts")),
          notes: String(data.get("notes") || "").trim(), completedAt: new Date().toISOString()
        });
        if (!existingWeeklyReview) state.weeklyReviews.push(reviewData);
        postSaveMessage = "Revisão semanal guardada.";
      } else if (type === "past-exam-import") {
        var pastExamCourse = courseById(data.get("courseId"));
        if (!pastExamCourse) throw new Error("Escolhe uma cadeira.");
        var examPayload = parseJSONText(data.get("json"), "A importação do teste");
        if (Array.isArray(examPayload)) examPayload = { questions: examPayload };
        if (examPayload && typeof examPayload !== "object") throw new Error("O JSON do teste tem um formato inválido.");
        examPayload = examPayload || {};
        var examTitle = String(data.get("title") || examPayload.title || "").trim();
        if (!examTitle) throw new Error("Escreve o nome do teste anterior.");
        var examYear = String(data.get("academicYear") || examPayload.academicYear || "").trim();
        var incomingQuestions = asArray(examPayload.questions);
        var validatedQuestions = incomingQuestions.map(function (question, index) {
          if (!question || typeof question !== "object") throw new Error("A pergunta " + (index + 1) + " não é um objeto válido.");
          var prompt = String(question.prompt || "").trim();
          if (!prompt) throw new Error("A pergunta " + (question.number || index + 1) + " não tem enunciado. Usa [ILEGÍVEL] quando o original não puder ser lido.");
          var options = asArray(question.options).map(function (option) { return String(option == null ? "" : option).trim(); }).filter(Boolean);
          var answerIndex = question.answerIndex == null ? null : Number(question.answerIndex);
          if (answerIndex != null && (!options.length || answerIndex < 0 || answerIndex >= options.length)) throw new Error("A resposta certa da pergunta " + (question.number || index + 1) + " não corresponde às opções.");
          return {
            id: uid("question"), semesterId: pastExamCourse.semesterId, courseId: pastExamCourse.id,
            lessonIds: importLessonIds(pastExamCourse.id, question.lessonTitles), academicYear: String(question.academicYear || examYear).trim(),
            number: String(question.number || "").trim(), sourceType: "past-test", assessmentLabel: examTitle,
            prompt: prompt, answer: String(question.answer || "").trim(), explanation: String(question.explanation || "").trim(),
            points: nullableNumber(question.points), tags: asArray(question.tags).map(function (tag) { return String(tag).trim(); }).filter(Boolean),
            options: options, answerIndex: answerIndex, images: normalizeImageRefs(question.images)
          };
        });
        var newPastExam = { id: uid("pastexam"), semesterId: pastExamCourse.semesterId, courseId: pastExamCourse.id, title: examTitle, academicYear: examYear, date: data.get("date") || examPayload.date || "", source: String(data.get("source") || examPayload.source || "").trim(), notes: String(examPayload.notes || "").trim(), createdAt: new Date().toISOString() };
        validatedQuestions.forEach(function (question) { question.pastExamId = newPastExam.id; });
        state.pastExams.push(newPastExam);
        state.questions = state.questions.concat(validatedQuestions);
        postSaveMessage = incomingQuestions.length ? "Teste anterior e " + incomingQuestions.length + " pergunta(s) importados." : "Teste anterior criado. Podes adicionar as perguntas depois.";
      } else if (type === "profile") {
        state.profile.name = String(data.get("name") || "").trim();
        state.profile.institution = String(data.get("institution") || "").trim();
        state.profile.degree = String(data.get("degree") || "").trim();
        state.profile.targetGrade = clamp(data.get("targetGrade"), 0, 20) || 20;
      } else if (type === "lesson-notes") {
        var lessonNotes = lessonById(id);
        if (!lessonNotes) throw new Error("Aula não encontrada.");
        lessonNotes.notes = String(data.get("notes") || "").trim();
      } else {
        throw new Error("Tipo de formulário desconhecido.");
      }

      ensureBeOnlineTasks();
      await save(true);
      if ((type === "material" || type === "lesson") && manualSyncActivity) {
        setManualSyncActivity("A sincronizar a aula…", "A guardar a referência do ficheiro para aparecer nos outros dispositivos.", 97, true);
        if (Sync && Sync.getStatus().configured) { try { await Sync.syncNow(state, defaultState()); } catch (_) {} }
        finishManualSyncActivity(true);
      }
      closeModal();
      render();
      toast(type === "material" ? "Material enviado, guardado e ligado à aula." : type === "question" ? (id ? "Pergunta anterior atualizada." : "Pergunta anterior adicionada.") : type === "event" ? (id ? "Evento atualizado." : "Evento adicionado.") : type === "quiz" || type === "quiz-question" ? "Quiz atualizado." : postSaveMessage);
    } catch (error) {
      if (manualSyncActivity) finishManualSyncActivity(false);
      setFormError(form, error.message || "Não foi possível guardar.");
      if (submit) submit.disabled = false;
    }
  }

  function defaultSemesterDates() {
    var now = new Date();
    var yearStart = now.getMonth() >= 8 ? now.getFullYear() : now.getFullYear() - 1;
    if (now.getMonth() >= 1 && now.getMonth() < 8) {
      return { name: "2.º semestre", startDate: (yearStart + 1) + "-02-01", endDate: (yearStart + 1) + "-07-31" };
    }
    return { name: "1.º semestre", startDate: yearStart + "-09-01", endDate: (yearStart + 1) + "-01-31" };
  }

  function startOnboarding(mode) {
    var dates = defaultSemesterDates();
    onboarding = {
      mode: mode || "first",
      step: mode === "new-semester" ? 1 : 0,
      tutorialPage: 0,
      tutorialSkipped: false,
      draft: {
        profile: {
          name: state.profile.name || "",
          institution: state.profile.institution || "",
          degree: state.profile.degree || "",
          targetGrade: Number(state.profile.targetGrade) || 20
        },
        semester: { name: dates.name, academicYear: academicYearFor(), startDate: dates.startDate, endDate: dates.endDate },
        courses: [{ tempId: uid("draftcourse"), name: "", code: "", ects: 6, color: COLORS[0], lessonTypes: ["T", "TP"], evaluation: "Testes | 60 | test\nProjeto | 40 | project\nExame | 0 | exam", examReplacesTests: true }],
        schedule: [],
        assessments: []
      }
    };
    renderOnboarding();
  }

  function onboardingProgress() {
    var total = 5;
    var activeStep = onboarding.step === 0 ? 0 : onboarding.step;
    return '<div class="onboarding-progress" aria-label="Progresso">' + Array.from({ length: total }, function (_, index) { return '<span class="' + (index < activeStep ? "is-done" : "") + '"></span>'; }).join("") + "</div>";
  }

  function renderOnboarding() {
    if (!onboarding) return;
    document.body.style.overflow = "hidden";
    var main;
    if (onboarding.step === 0) main = renderOnboardingIntro();
    else if (onboarding.step === 1) main = renderOnboardingProfile();
    else if (onboarding.step === 2) main = renderOnboardingCourses();
    else if (onboarding.step === 3) main = renderOnboardingSchedule();
    else if (onboarding.step === 4) main = renderOnboardingAssessments();
    else main = renderOnboardingFinish();
    modalRoot.innerHTML = '<div class="onboarding-layer"><section class="onboarding-shell" role="dialog" aria-modal="true" aria-label="Configurar a Twenty"><aside class="onboarding-aside"><div class="onboarding-logo"><span class="brand-mark">20</span> twenty · study os</div><h2>O teu semestre, num só lugar.</h2><p>Organiza cadeiras, horário, aulas, avaliações, materiais e notas desde o primeiro dia.</p><div class="onboarding-quote">Semestre · horário · aulas · avaliações · estudo</div></aside><main class="onboarding-main">' + main + "</main></section></div>";
    refreshIcons(modalRoot);
  }

  function renderOnboardingIntro() {
    var pages = [
      {
        title: "Vamos montar o teu semestre.",
        copy: "A Twenty começa vazia para não inventar informação. Em poucos passos configuras as tuas cadeiras, horários e avaliações.",
        visual: '<div class="tutorial-visual"><div class="tutorial-phone"><div></div><div></div><div></div></div></div>',
        button: "Ver tutorial"
      },
      {
        title: "Cada aula fica acompanhada.",
        copy: "Na aula encontras slides, apontamentos e perguntas anteriores. Quando termina, fazes um quiz curto: detetas dúvidas cedo e não deixas matéria acumular.",
        visual: '<div class="tutorial-visual"><div style="width:min(500px,86%);display:grid;grid-template-columns:1.2fr .8fr;gap:10px"><div class="card card-dark" style="min-height:190px"><span class="badge badge-yellow">Aula em direto</span><h3 style="margin-top:18px">TP08 · Polimorfismo</h3><div class="live-meta"><span>PDF</span><span>Quiz</span><span>Perguntas</span></div></div><div style="display:grid;gap:10px"><div class="card card-yellow">Slides</div><div class="card card-pink">Teste 2024/25</div></div></div></div>',
        button: "Continuar"
      },
      {
        title: "Os dados ficam contigo.",
        copy: "A app guarda tudo no dispositivo. O ficheiro academic-data.json pode ser editado à mão e é relido automaticamente; os PDFs enviados ficam no armazenamento local.",
        visual: '<div class="tutorial-visual"><div class="card" style="width:min(520px,86%);text-align:left"><span class="metric-icon"><i data-lucide="braces"></i></span><h3 style="margin-top:17px">data/academic-data.json</h3><p class="card-subtitle">Horário · cadeiras · testes · aulas · perguntas</p><div class="formula" style="margin-top:14px">{ <strong>"schemaVersion"</strong>: 4, <strong>"courses"</strong>: [] }</div></div></div>',
        button: "Configurar agora"
      }
    ];
    var page = pages[onboarding.tutorialPage] || pages[0];
    return onboardingProgress() + '<h1>' + page.title + '</h1><p>' + page.copy + '</p>' + page.visual + '<div class="onboarding-actions"><div style="display:flex;gap:10px;flex-wrap:wrap"><button class="button button-ghost" type="button" data-action="import-json"><i data-lucide="upload"></i>Já tenho um JSON</button><button class="button button-yellow" type="button" data-action="onboarding-connect-git"><i data-lucide="arrow-down-to-line"></i>Usar dados sincronizados</button></div><div><button class="button" type="button" data-action="tutorial-skip">Pular tutorial</button><button class="button button-dark" type="button" data-action="tutorial-next">' + page.button + '<i data-lucide="arrow-right"></i></button></div></div>';
  }

  function renderOnboardingProfile() {
    var profile = onboarding.draft.profile;
    var semester = onboarding.draft.semester;
    return onboardingProgress() + '<h1>' + (onboarding.mode === "new-semester" ? "Novo semestre" : "Configuração inicial") + '</h1><p>Define o perfil académico e o semestre ativo.</p><form id="onboardingForm" data-step="1"><div class="form-grid"><div class="field field-full"><label>O teu nome</label><input name="name" required placeholder="Ex.: Matilde" value="' + attr(profile.name) + '"></div><div class="field"><label>Instituição</label><input name="institution" placeholder="Faculdade / universidade" value="' + attr(profile.institution) + '"></div><div class="field"><label>Curso</label><input name="degree" placeholder="Ex.: Engenharia Informática" value="' + attr(profile.degree) + '"></div><div class="field"><label>Meta</label><input name="targetGrade" type="number" min="0" max="20" step="0.1" value="' + attr(profile.targetGrade) + '"></div><div class="field"><label>Nome do semestre</label><input name="semesterName" required value="' + attr(semester.name) + '"></div><div class="field"><label>Ano letivo</label><input name="academicYear" required placeholder="2025/26" value="' + attr(semester.academicYear) + '"></div><div class="field"><label>Início</label><input name="startDate" type="date" value="' + attr(semester.startDate) + '"></div><div class="field"><label>Fim</label><input name="endDate" type="date" value="' + attr(semester.endDate) + '"></div></div></form><div class="onboarding-actions"><div style="display:flex;gap:8px;flex-wrap:wrap"><button class="button" type="button" data-action="onboarding-back"><i data-lucide="arrow-left"></i>Voltar</button><button class="button button-yellow" type="button" data-action="onboarding-connect-git"><i data-lucide="arrow-down-to-line"></i>Usar dados sincronizados</button></div><div><button class="button button-dark" type="button" data-action="onboarding-next">Cadeiras<i data-lucide="arrow-right"></i></button></div></div>';
  }

  function renderOnboardingCourses() {
    var courses = onboarding.draft.courses;
    var rows = courses.map(function (course, index) {
      return '<article class="setup-row" data-temp-id="' + attr(course.tempId) + '" style="grid-template-columns:1.3fr .55fr .4fr .42fr auto"><input name="name" aria-label="Nome da cadeira" required placeholder="Nome da cadeira" value="' + attr(course.name) + '"><input name="code" aria-label="Código" placeholder="Código" value="' + attr(course.code) + '"><input name="ects" aria-label="ECTS" type="number" min="0" max="60" step=".5" value="' + attr(course.ects) + '"><input name="color" aria-label="Cor" type="color" value="' + safeColor(course.color, COLORS[index % COLORS.length]) + '"><button class="remove-row" type="button" data-action="remove-onboarding-course" data-index="' + index + '" aria-label="Remover"><i data-lucide="trash-2"></i></button><div class="field" style="grid-column:1/-1"><label>Tipos de aula</label><div class="checkbox-line">' + ["T", "TP", "P", "LAB"].map(function (type) { return '<label class="checkbox-chip"><input type="checkbox" name="type-' + type + '" ' + (asArray(course.lessonTypes).indexOf(type) >= 0 ? "checked" : "") + '>' + type + "</label>"; }).join("") + '</div></div><div class="field" style="grid-column:1/-1"><label>Avaliação: nome | percentagem | tipo</label><textarea name="evaluation" placeholder="Testes | 60 | test\nProjeto | 40 | project">' + esc(course.evaluation) + '</textarea><label class="checkbox-chip"><input name="examReplacesTests" type="checkbox" ' + (course.examReplacesTests ? "checked" : "") + '> Exame pode substituir testes se melhorar a nota</label></div></article>';
    }).join("");
    return onboardingProgress() + '<h1>Que cadeiras vais ter?</h1><p>Suporta cadeiras teóricas, práticas ou mistas. Podes afinar tudo mais tarde na configuração da cadeira.</p><form id="onboardingForm" data-step="2"><div class="setup-list">' + rows + '</div></form><button class="button button-yellow" style="margin-top:11px" type="button" data-action="add-onboarding-course"><i data-lucide="plus"></i>Adicionar cadeira</button><div class="onboarding-actions"><button class="button" type="button" data-action="onboarding-back"><i data-lucide="arrow-left"></i>Voltar</button><div><button class="button button-dark" type="button" data-action="onboarding-next">Horário<i data-lucide="arrow-right"></i></button></div></div>';
  }

  function draftCourseOptions(selected) {
    return onboarding.draft.courses.filter(function (course) { return course.name; }).map(function (course) {
      return '<option value="' + attr(course.tempId) + '" ' + (selected === course.tempId ? "selected" : "") + '>' + esc(course.code ? course.code + " · " + course.name : course.name) + "</option>";
    }).join("");
  }

  function renderOnboardingSchedule() {
    var rows = onboarding.draft.schedule.map(function (entry, index) {
      return '<div class="setup-row schedule-row" data-index="' + index + '"><select name="courseRef"><option value="">Cadeira…</option>' + draftCourseOptions(entry.courseRef) + '</select><select name="weekday">' + WEEKDAYS.map(function (day, dayIndex) { return '<option value="' + dayIndex + '" ' + (Number(entry.weekday) === dayIndex ? "selected" : "") + '>' + day + "</option>"; }).join("") + '</select><input name="start" type="time" value="' + attr(entry.start || "09:00") + '"><input name="end" type="time" value="' + attr(entry.end || "10:30") + '"><select name="lessonType"><option value="T">T</option><option value="TP" ' + (entry.type === "TP" ? "selected" : "") + '>TP</option><option value="P" ' + (entry.type === "P" ? "selected" : "") + '>P</option><option value="LAB" ' + (entry.type === "LAB" ? "selected" : "") + '>LAB</option></select><button class="remove-row" type="button" data-action="remove-onboarding-schedule" data-index="' + index + '"><i data-lucide="trash-2"></i></button><input name="room" style="grid-column:1/-2" placeholder="Sala (opcional)" value="' + attr(entry.room) + '"></div>';
    }).join("");
    return onboardingProgress() + '<h1>Monta o horário.</h1><p>É este horário que permite mostrar a aula em direto. Se ainda não o souberes, podes saltar e editar depois — também diretamente no JSON.</p><form id="onboardingForm" data-step="3"><div class="setup-list">' + (rows || '<div class="empty-state" style="min-height:140px"><h3>Horário ainda vazio</h3><p>Podes continuar sem adicionar nada.</p></div>') + '</div></form><button class="button button-yellow" style="margin-top:11px" type="button" data-action="add-onboarding-schedule"><i data-lucide="plus"></i>Adicionar aula recorrente</button><div class="onboarding-actions"><button class="button" type="button" data-action="onboarding-back"><i data-lucide="arrow-left"></i>Voltar</button><div><button class="button button-dark" type="button" data-action="onboarding-next">Avaliações<i data-lucide="arrow-right"></i></button></div></div>';
  }

  function renderOnboardingAssessments() {
    var rows = onboarding.draft.assessments.map(function (item, index) {
      return '<div class="setup-row assessment-row" data-index="' + index + '"><select name="courseRef"><option value="">Cadeira…</option>' + draftCourseOptions(item.courseRef) + '</select><input name="title" placeholder="Teste 1 / Projeto" value="' + attr(item.title) + '"><select name="type"><option>Teste</option><option ' + (item.type === "Exame" ? "selected" : "") + '>Exame</option><option ' + (item.type === "Projeto" ? "selected" : "") + '>Projeto</option><option ' + (item.type === "Apresentação" ? "selected" : "") + '>Apresentação</option></select><input name="date" type="date" value="' + attr(item.date) + '"><button class="remove-row" type="button" data-action="remove-onboarding-assessment" data-index="' + index + '"><i data-lucide="trash-2"></i></button><input name="time" type="time" value="' + attr(item.time || "10:00") + '"><input name="weight" type="number" min="0" max="100" placeholder="Peso %" value="' + attr(item.weight || 0) + '"></div>';
    }).join("");
    return onboardingProgress() + '<h1>Já sabes alguma data?</h1><p>Adiciona testes, projetos ou exames que já estejam marcados. A matéria específica pode ser ligada às aulas mais tarde.</p><form id="onboardingForm" data-step="4"><div class="setup-list">' + (rows || '<div class="empty-state" style="min-height:140px"><h3>Nenhuma avaliação conhecida</h3><p>Sem problema — adiciona quando o professor anunciar.</p></div>') + '</div></form><button class="button button-yellow" style="margin-top:11px" type="button" data-action="add-onboarding-assessment"><i data-lucide="plus"></i>Adicionar avaliação</button><div class="onboarding-actions"><button class="button" type="button" data-action="onboarding-back"><i data-lucide="arrow-left"></i>Voltar</button><div><button class="button button-dark" type="button" data-action="onboarding-next">Rever tudo<i data-lucide="arrow-right"></i></button></div></div>';
  }

  function renderOnboardingFinish() {
    var draft = onboarding.draft;
    return onboardingProgress() + '<h1>O teu sistema está pronto.</h1><p>Confirma o essencial. Depois podes acrescentar aulas, carregar PDFs e construir os quizzes no lado admin.</p><div class="finish-card"><span class="badge badge-dark">' + esc(draft.semester.academicYear) + '</span><h2 style="margin:13px 0 4px">' + esc(draft.semester.name) + '</h2><p class="card-subtitle">' + esc(draft.profile.degree || "Curso") + (draft.profile.institution ? " · " + esc(draft.profile.institution) : "") + '</p><div class="finish-list"><div><strong>' + draft.courses.filter(function (c) { return c.name; }).length + '</strong><span>cadeiras</span></div><div><strong>' + draft.schedule.filter(function (s) { return s.courseRef; }).length + '</strong><span>aulas no horário</span></div><div><strong>' + draft.assessments.filter(function (a) { return a.courseRef && a.title; }).length + '</strong><span>avaliações marcadas</span></div></div></div><div class="form-note" style="margin-top:15px">A seguir: abre uma cadeira, cria uma aula com data e carrega os slides — mesmo que sejam de um ano letivo anterior.</div><div class="onboarding-actions"><button class="button" type="button" data-action="onboarding-back"><i data-lucide="arrow-left"></i>Voltar</button><div><button class="button button-dark" type="button" data-action="finish-onboarding"><i data-lucide="sparkles"></i>Entrar na Twenty</button></div></div>';
  }

  function captureOnboardingStep(validate) {
    if (validate == null) validate = true;
    var form = document.getElementById("onboardingForm");
    if (!form) return true;
    var step = Number(form.getAttribute("data-step"));
    if (step === 1) {
      var data = new FormData(form);
      if (validate && (!String(data.get("name") || "").trim() || !String(data.get("semesterName") || "").trim() || !String(data.get("academicYear") || "").trim())) {
        toast("Preenche o nome, semestre e ano letivo.", "warning");
        return false;
      }
      onboarding.draft.profile = { name: String(data.get("name")).trim(), institution: String(data.get("institution") || "").trim(), degree: String(data.get("degree") || "").trim(), targetGrade: clamp(data.get("targetGrade"), 0, 20) || 20 };
      onboarding.draft.semester = { name: String(data.get("semesterName")).trim(), academicYear: String(data.get("academicYear")).trim(), startDate: data.get("startDate") || "", endDate: data.get("endDate") || "" };
    } else if (step === 2) {
      var courseRows = Array.from(form.querySelectorAll(".setup-row"));
      var courses = courseRows.map(function (row, index) {
        var courseName = row.querySelector('[name="name"]').value.trim();
        return { tempId: row.getAttribute("data-temp-id") || uid("draftcourse"), name: courseName, code: row.querySelector('[name="code"]').value.trim(), ects: Number(row.querySelector('[name="ects"]').value) || 0, color: safeColor(row.querySelector('[name="color"]').value, COLORS[index % COLORS.length]), lessonTypes: ["T", "TP", "P", "LAB"].filter(function (type) { return row.querySelector('[name="type-' + type + '"]').checked; }), evaluation: row.querySelector('[name="evaluation"]').value, examReplacesTests: row.querySelector('[name="examReplacesTests"]').checked };
      });
      if (validate && !courses.some(function (course) { return course.name; })) {
        toast("Adiciona pelo menos uma cadeira.", "warning");
        return false;
      }
      onboarding.draft.courses = courses.filter(function (course) { return course.name; });
    } else if (step === 3) {
      onboarding.draft.schedule = Array.from(form.querySelectorAll(".setup-row")).map(function (row) {
        return { courseRef: row.querySelector('[name="courseRef"]').value, weekday: Number(row.querySelector('[name="weekday"]').value), start: row.querySelector('[name="start"]').value, end: row.querySelector('[name="end"]').value, type: row.querySelector('[name="lessonType"]').value, room: row.querySelector('[name="room"]').value.trim() };
      }).filter(function (entry) { return entry.courseRef; });
    } else if (step === 4) {
      onboarding.draft.assessments = Array.from(form.querySelectorAll(".setup-row")).map(function (row) {
        return { courseRef: row.querySelector('[name="courseRef"]').value, title: row.querySelector('[name="title"]').value.trim(), type: row.querySelector('[name="type"]').value, date: row.querySelector('[name="date"]').value, time: row.querySelector('[name="time"]').value, weight: Number(row.querySelector('[name="weight"]').value) || 0 };
      }).filter(function (item) { return item.courseRef && item.title; });
    }
    return true;
  }

  async function finishOnboarding() {
    var draft = onboarding.draft;
    var semesterId = uid("semester");
    var semester = { id: semesterId, name: draft.semester.name, academicYear: draft.semester.academicYear, startDate: draft.semester.startDate, endDate: draft.semester.endDate, archived: false, createdAt: new Date().toISOString() };
    var map = {};
    var courses = draft.courses.filter(function (course) { return course.name; }).map(function (course, index) {
      var id = uid("course");
      map[course.tempId] = id;
      var components = parseEvaluation(course.evaluation, []);
      if (course.examReplacesTests && !components.some(function (component) { return component.kind === "exam"; })) components.push({ id: uid("component"), label: "Exame", weight: 0, kind: "exam", replaceable: false });
      return { id: id, semesterId: semesterId, name: course.name, code: course.code, ects: Number(course.ects) || 0, color: safeColor(course.color, COLORS[index % COLORS.length]), lessonTypes: asArray(course.lessonTypes), evaluation: { components: components, examReplacesTests: !!course.examReplacesTests, replacementPolicy: "if-higher" } };
    });
    if (!courses.length) {
      toast("Adiciona pelo menos uma cadeira.", "warning");
      return;
    }
    if (onboarding.mode === "first") {
      state = normalizeState(state);
    }
    state.profile = Object.assign(state.profile, draft.profile, { onboardingComplete: true, tutorialSeen: !onboarding.tutorialSkipped });
    state.semesters.push(semester);
    state.courses = state.courses.concat(courses);
    state.currentSemesterId = semesterId;
    state.schedule = state.schedule.concat(draft.schedule.filter(function (entry) { return map[entry.courseRef]; }).map(function (entry) {
      return { id: uid("schedule"), semesterId: semesterId, courseId: map[entry.courseRef], weekday: Number(entry.weekday), start: entry.start, end: entry.end, type: entry.type, room: entry.room };
    }));
    state.assessments = state.assessments.concat(draft.assessments.filter(function (item) { return map[item.courseRef]; }).map(function (item) {
      var courseId = map[item.courseRef];
      var linkedCourse = courses.find(function (course) { return course.id === courseId; });
      return { id: uid("assessment"), semesterId: semesterId, courseId: courseId, componentId: suggestedComponentId(linkedCourse, item), type: item.type, title: item.title, date: item.date, time: item.time, weight: Number(item.weight) || 0, lessonIds: [], replacesTests: item.type === "Exame" };
    }));
    await save(true);
    onboarding = null;
    closeModal();
    setRoute("home");
    toast("Semestre configurado. Bem-vinda à Twenty ✨");
  }

  var quizRuntime = null;
  var pendingImport = null;

  function allImageRefs() {
    var refs = [];
    state.questions.forEach(function (question) { refs = refs.concat(normalizeImageRefs(question.images)); });
    state.events.forEach(function (event) { refs = refs.concat(normalizeImageRefs(event.images, "event")); });
    state.quizzes.forEach(function (quiz) { asArray(quiz.questions).forEach(function (question) { refs = refs.concat(normalizeImageRefs(question.images)); }); });
    return refs;
  }

  async function openImage(id) {
    var image = allImageRefs().find(function (item) { return item.id === id; });
    if (!image) { toast("Imagem não encontrada.", "error"); return; }
    var src = safeResourceUrl(image.url);
    var objectUrl = null;
    if (image.blobId) {
      var record = await DB.getFile(image.blobId);
      if (!record || !record.blob) { toast("A imagem local já não está disponível neste dispositivo.", "error"); return; }
      objectUrl = URL.createObjectURL(record.blob);
      src = objectUrl;
    }
    if (!src) { toast("A imagem não tem um caminho válido.", "error"); return; }
    openModal(imageRoleLabel(image.role), '<div class="image-lightbox"><img src="' + attr(src) + '" alt="' + attr(image.name || imageRoleLabel(image.role)) + '">' + (image.caption ? '<p>' + esc(image.caption) + '</p>' : '') + '<a class="button button-small" href="' + attr(src) + '" target="_blank" rel="noopener"><i data-lucide="external-link"></i>Abrir original</a></div>', { className: "modal-image" });
    if (objectUrl) activeObjectUrl = objectUrl;
  }

  async function copyText(value) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(value);
      return;
    }
    var textarea = document.createElement("textarea");
    textarea.value = value;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
  }

  async function openMaterial(id) {
    var material = state.materials.find(function (item) { return item.id === id; });
    if (!material) { toast("Material não encontrado.", "error"); return; }
    var src = "";
    var newObjectUrl = null;
    var mime = material.mimeType || "";
    var record = material.blobId ? await DB.getFile(material.blobId) : null;
    if (record && record.blob) {
      newObjectUrl = URL.createObjectURL(record.blob);
      src = newObjectUrl;
      mime = record.type || mime;
    } else if (material.remoteFile && material.remoteFile.path && Sync) {
      setManualSyncActivity("A descarregar o material…", "A obter o ficheiro sincronizado do Git.", 4, true);
      try {
        var blob = await Sync.downloadFile(material.remoteFile, { onProgress: function (report) {
          var progress = report.progress == null ? null : 5 + Math.round(report.progress * 0.9);
          var detail = report.total ? formatBytes(report.loaded) + " de " + formatBytes(report.total) + " descarregados" : "A receber o ficheiro…";
          setManualSyncActivity("A descarregar o material…", detail, progress, true);
        } });
        newObjectUrl = URL.createObjectURL(blob);
        src = newObjectUrl;
        mime = blob.type || mime;
        finishManualSyncActivity(true);
      } catch (error) {
        finishManualSyncActivity(false);
        toast(error.message || "Não foi possível descarregar o material.", "error");
        return;
      }
    } else {
      src = safeResourceUrl(material.url);
    }
    if (!src) { toast("Este material ainda não tem uma cópia sincronizada disponível.", "error"); return; }
    var isOffice = /powerpoint|presentation/i.test(mime) || /\.pptx?(?:$|\?)/i.test(material.fileName || src);
    var body = isOffice
      ? '<div class="empty-state"><span class="empty-icon"><i data-lucide="presentation"></i></span><h3>PowerPoint pronto</h3><p>O navegador não o pré-visualiza diretamente. Abre ou guarda o ficheiro original.</p><a class="button button-dark" href="' + attr(src) + '" target="_blank" rel="noopener" download="' + attr(material.fileName || "apresentacao.pptx") + '"><i data-lucide="download"></i>Abrir ficheiro</a></div>'
      : '<iframe class="pdf-frame" src="' + attr(src) + '#view=FitH" title="' + attr(material.title) + '"></iframe><div class="list-actions" style="margin-top:12px"><a class="button button-small" href="' + attr(src) + '" target="_blank" rel="noopener"><i data-lucide="external-link"></i>Abrir noutra janela</a></div>';
    openModal(material.title, body, { className: "modal-pdf" });
    if (newObjectUrl) activeObjectUrl = newObjectUrl;
  }

  function showQuestionAnswer(id) {
    var question = state.questions.find(function (item) { return item.id === id; });
    if (!question) return;
    var exam = pastExamById(question.pastExamId);
    openModal("Pergunta de teste anterior", '<div class="question-card question-detail"><div class="question-meta"><span class="badge badge-pink">' + esc(question.academicYear || "Ano não indicado") + '</span>' + (question.number ? '<span class="badge badge-dark">' + esc(question.number) + '</span>' : '') + '<span class="badge badge-violet">' + esc(exam ? exam.title : question.assessmentLabel || "Teste anterior") + '</span></div><h3 style="margin:17px 0 8px;line-height:1.4">' + esc(question.prompt) + '</h3>' + renderImageGallery(question.images, "question", { ownerId: question.id }) + '<div class="answer-box"><strong>Resposta:</strong><br>' + (question.answer ? nl2br(question.answer) : "Ainda não foi adicionada uma resposta.") + renderImageGallery(question.images, "solution", { ownerId: question.id }) + '</div>' + (question.explanation || normalizeImageRefs(question.images).some(function (image) { return image.role === "explanation"; }) ? '<div class="form-note question-explanation"><strong>Explicação:</strong><br>' + (question.explanation ? nl2br(question.explanation) : "") + renderImageGallery(question.images, "explanation", { ownerId: question.id }) + "</div>" : "") + '<div class="list-actions" style="margin-top:14px"><button class="button" type="button" data-action="edit-question" data-id="' + attr(question.id) + '"><i data-lucide="pencil"></i>Editar</button></div></div>');
  }

  async function generateQuizFromPast(lessonId, startAfter) {
    var lesson = lessonById(lessonId);
    if (!lesson) return null;
    var past = pastQuestionsForLesson(lesson.id);
    if (!past.length) {
      toast("Esta aula ainda não tem perguntas de testes anteriores.", "warning");
      return null;
    }
    var quiz = state.quizzes.find(function (item) {
      return item.lessonId === lesson.id && item.generatedFromPastQuestions;
    });
    if (!quiz) {
      quiz = {
        id: uid("quiz"), semesterId: lesson.semesterId, courseId: lesson.courseId, lessonId: lesson.id,
        title: "Quiz da aula · " + lesson.title, questions: [], generatedFromPastQuestions: true,
        createdAt: new Date().toISOString(), lastScore: null
      };
      state.quizzes.push(quiz);
    }
    var used = asArray(quiz.questions).map(function (question) { return question.sourceQuestionId; }).filter(Boolean);
    var additions = past.filter(function (question) { return used.indexOf(question.id) < 0; }).map(quizQuestionFromPast);
    quiz.questions = asArray(quiz.questions).concat(additions);
    await save(true);
    render();
    if (additions.length) toast(additions.length + " pergunta(s) anterior(es) preparadas para o quiz.");
    if (startAfter !== false) startQuiz(quiz.id);
    return quiz;
  }

  async function doLessonQuiz(lessonId) {
    var lesson = lessonById(lessonId);
    if (!lesson) return;
    var linked = state.quizzes.filter(function (quiz) { return quiz.lessonId === lesson.id && asArray(quiz.questions).length; });
    if (linked.length) {
      startQuiz(linked[0].id);
    } else if (pastQuestionsForLesson(lesson.id).length) {
      await generateQuizFromPast(lesson.id, true);
    } else {
      setRoute("lesson", lesson.id);
      toast("Esta aula ainda não tem quiz. Cria um quiz normal para concluíres a revisão.", "warning");
    }
  }

  function startQuiz(id) {
    var quiz = state.quizzes.find(function (item) { return item.id === id; });
    if (!quiz || !asArray(quiz.questions).length) { toast("Este quiz ainda não tem perguntas.", "warning"); return; }
    quizRuntime = { quizId: id, index: 0, answers: [], selected: null, revealed: false };
    renderQuizQuestion();
  }

  function renderQuizQuestion() {
    var quiz = state.quizzes.find(function (item) { return quizRuntime && item.id === quizRuntime.quizId; });
    if (!quiz) return;
    var questions = asArray(quiz.questions);
    if (quizRuntime.index >= questions.length) { finishQuiz(quiz); return; }
    var question = questions[quizRuntime.index];
    var selected = quizRuntime.selected;
    var progress = '<div class="quiz-progress"><span style="width:' + ((quizRuntime.index + 1) / questions.length * 100) + '%"></span></div><p class="card-label" style="margin-top:14px">Pergunta ' + (quizRuntime.index + 1) + ' de ' + questions.length + '</p>';
    var body;
    var footer;
    if (question.mode === "self-check" || !asArray(question.options).length) {
      var source = [question.assessmentLabel, question.academicYear].filter(Boolean).join(" · ") || "Pergunta de teste anterior";
      body = progress + '<div class="self-check-source"><span class="badge badge-pink"><i data-lucide="history"></i>Pergunta anterior</span><small>' + esc(source) + '</small></div><h3 class="quiz-question">' + esc(question.prompt) + '</h3>' + renderImageGallery(question.images, "question", { ownerId: question.sourceQuestionId || "" });
      if (quizRuntime.revealed) {
        body += '<div class="self-check-answer"><p class="card-label">Resposta guardada</p><div>' + nl2br(question.answer || "A resposta ainda não foi adicionada.") + '</div>' + renderImageGallery(question.images, "solution", { ownerId: question.sourceQuestionId || "" }) + (question.explanation || normalizeImageRefs(question.images).some(function (image) { return image.role === "explanation"; }) ? '<div class="self-check-explanation"><strong>Explicação:</strong> ' + nl2br(question.explanation || "") + renderImageGallery(question.images, "explanation", { ownerId: question.sourceQuestionId || "" }) + '</div>' : '') + '</div><p class="self-check-prompt">Compara a tua resposta com a solução guardada e regista se precisas de rever.</p>';
        footer = '<footer class="modal-foot self-check-actions"><button class="button" type="button" data-action="close-modal">Sair</button><button class="button" type="button" data-action="quiz-self-rate" data-value="0"><i data-lucide="rotate-ccw"></i>Preciso rever</button><button class="button button-dark" type="button" data-action="quiz-self-rate" data-value="1"><i data-lucide="check"></i>Sabia</button></footer>';
      } else {
        body += '<div class="form-note self-check-note"><strong>Responde primeiro sem consultar os apontamentos.</strong><br>Quando estiveres pronta, revela a solução que guardaste na pergunta original.</div>';
        footer = '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Sair</button><button class="button button-dark" type="button" data-action="quiz-reveal"><i data-lucide="eye"></i>Revelar resposta</button></footer>';
      }
    } else {
      body = progress + '<h3 class="quiz-question">' + esc(question.prompt) + '</h3>' + renderImageGallery(question.images, "question", { ownerId: question.sourceQuestionId || "" }) + '<div class="quiz-options">' + asArray(question.options).map(function (option, index) {
        return '<button class="quiz-option ' + (selected === index ? "is-selected" : "") + '" type="button" data-action="quiz-answer" data-index="' + index + '"><span>' + String.fromCharCode(65 + index) + '</span>' + esc(option) + "</button>";
      }).join("") + '</div><div id="quizFeedback"></div>';
      footer = '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Sair</button><button class="button button-dark" type="button" data-action="quiz-next" ' + (selected == null ? "disabled" : "") + '>' + (quizRuntime.index === questions.length - 1 ? "Terminar" : "Seguinte") + '<i data-lucide="arrow-right"></i></button></footer>';
    }
    openModal(quiz.title, body, { footer: footer });
  }

  async function finishQuiz(quiz) {
    var questions = asArray(quiz.questions);
    var correct = quizRuntime.answers.reduce(function (sum, answer, index) {
      var question = questions[index] || {};
      if (question.mode === "self-check" || !asArray(question.options).length) return sum + (answer === 1 ? 1 : 0);
      return sum + (answer === Number(question.answerIndex) ? 1 : 0);
    }, 0);
    var score = Math.round(correct / questions.length * 100);
    quiz.lastScore = score;
    quiz.lastCompletedAt = new Date().toISOString();
    if (quiz.lessonId) completeLessonBeOnline(quiz.lessonId);
    ensureBeOnlineTasks();
    await save(true);
    quizRuntime = null;
    var closesLesson = !!quiz.lessonId;
    var resultCopy = closesLesson
      ? (score === 100 ? "Aula acompanhada. Mantiveste-te em linha e sem matéria acumulada." : score >= 70 ? "Aula acompanhada. Anota o que falhou e esclarece as dúvidas cedo." : "Aula acompanhada, mas merece revisão: volta aos slides e leva as dúvidas ao professor.")
      : (score === 100 ? "Excelente domínio deste quiz." : score >= 70 ? "Bom caminho. Revê os itens que falharam." : "Volta aos materiais e tenta novamente.");
    openModal(closesLesson ? "Aula revista" : "Quiz concluído", '<div class="finish-card" style="text-align:center"><span class="badge badge-dark"><i data-lucide="' + (closesLesson ? "book-check" : "sparkles") + '"></i>' + (closesLesson ? "Aula acompanhada" : "Resultado") + '</span><h2 style="margin:17px 0 3px;font-size:4rem;letter-spacing:-.08em">' + score + '%</h2><p class="card-subtitle">' + correct + ' de ' + questions.length + ' itens dominados</p><div class="progress-ring" style="--progress:' + score + '%;margin:23px auto"><strong>' + score + '%</strong></div><p style="font-size:.7rem;font-weight:700;color:var(--muted)">' + resultCopy + '</p><button class="button button-dark" type="button" data-action="close-modal"><i data-lucide="check"></i>Fechar</button></div>');
  }

  function showAssessmentScope(id) {
    var assessment = state.assessments.find(function (item) { return item.id === id; });
    if (!assessment) return;
    var course = courseById(assessment.courseId);
    var lessons = asArray(assessment.lessonIds).map(lessonById).filter(Boolean);
    var component = asArray(course && course.evaluation && course.evaluation.components).find(function (item) { return item.id === assessment.componentId; });
    var replacementNames = asArray(assessment.replacementAssessmentIds).map(function (assessmentId) { var target = assessmentById(assessmentId); return target && target.title; }).filter(Boolean);
    var body = '<div class="question-meta"><span class="badge badge-yellow">' + esc(assessment.type) + '</span><span class="badge">' + relativeDate(assessment.date) + '</span><span class="badge badge-violet">' + esc(course ? course.name : "Cadeira") + '</span>' + (component ? '<span class="badge badge-mint">' + esc(component.label) + '</span>' : '') + assessmentRuleBadges(assessment) + '</div><h3 style="margin:16px 0 4px">' + esc(assessment.title) + '</h3><p class="card-subtitle">' + esc([assessment.date ? formatLongDate(assessment.date) : "Data por definir", assessment.time || "", assessment.location || ""].filter(Boolean).join(" · ")) + '</p>' + assessmentRuleSummary(assessment) + (replacementNames.length ? '<div class="form-note" style="margin-top:12px"><strong>Substitui:</strong> ' + esc(replacementNames.join(" · ")) + '</div>' : '') + '<div class="section-heading assessment-scope-heading"><div><h3>Matéria</h3><p>Aulas incluídas nesta avaliação</p></div></div><div class="list-stack">' + (lessons.length ? lessons.map(renderLessonRow).join("") : emptyState("list-tree", "Matéria ainda não definida", "Edita a avaliação e seleciona as aulas que saem.", null)) + '</div><div class="list-actions" style="margin-top:15px"><button class="button button-dark" type="button" data-action="add-grade" data-assessment="' + attr(assessment.id) + '"><i data-lucide="chart-no-axes-combined"></i>Adicionar nota</button><button class="button" type="button" data-action="edit-assessment" data-id="' + attr(assessment.id) + '"><i data-lucide="pencil"></i>Editar avaliação</button></div>';
    openModal("Matéria da avaliação", body);
  }

  function showEventDetail(id) {
    var event = state.events.find(function (item) { return item.id === id; });
    if (!event) return;
    var body = '<div class="question-meta"><span class="badge badge-pink">Evento</span><span class="badge">' + esc(formatLongDate(event.date)) + '</span></div><h3 style="margin:16px 0 5px">' + esc(event.title) + '</h3><p class="card-subtitle">' + (event.time ? esc(event.time) + ' · ' : '') + esc(event.location || "Local por definir") + '</p>' + renderImageGallery(event.images, "event", { ownerId: event.id }) + (event.notes ? '<div class="form-note" style="margin-top:14px">' + nl2br(event.notes) + '</div>' : '') + '<div class="list-actions" style="margin-top:14px">' + (event.url ? '<a class="button button-dark" href="' + attr(event.url) + '" target="_blank" rel="noopener"><i data-lucide="external-link"></i>Abrir ligação</a>' : '') + '<button class="button" type="button" data-action="edit-event" data-id="' + attr(event.id) + '"><i data-lucide="pencil"></i>Editar</button></div>';
    openModal("Evento da faculdade", body);
  }

  function showTaskDetail(id) {
    var task = state.tasks.find(function (item) { return item.id === id; });
    if (!task) return;
    var course = courseById(task.courseId);
    var body = '<div class="question-meta"><span class="badge badge-violet">Tarefa</span><span class="badge">' + esc(relativeDate(task.dueDate)) + '</span></div><h3 style="margin:16px 0 5px">' + esc(task.title) + '</h3><p class="card-subtitle">' + esc(course ? course.name : "Pessoal") + (task.dueTime ? ' · ' + esc(task.dueTime) : '') + '</p><div class="list-actions" style="margin-top:15px">' + (task.lessonId ? '<button class="button button-dark" type="button" data-route="lesson" data-id="' + attr(task.lessonId) + '"><i data-lucide="arrow-right"></i>Abrir aula</button>' : '') + '<button class="button" type="button" data-action="toggle-task" data-id="' + attr(task.id) + '"><i data-lucide="check"></i>' + (task.done ? "Reabrir" : "Concluir") + '</button></div>';
    openModal("Detalhes da tarefa", body);
  }

  function showScheduleDetail(id) {
    var entry = state.schedule.find(function (item) { return item.id === id; });
    if (!entry) return;
    var course = courseById(entry.courseId);
    var occurrence = nextOccurrenceForSchedule(entry);
    var dateValue = occurrence.dateISO;
    var prepared = occurrence.lesson;
    var primary = prepared
      ? '<button class="button button-dark" type="button" data-route="lesson" data-id="' + attr(prepared.id) + '"><i data-lucide="arrow-right"></i>Abrir “' + esc(prepared.title) + '”</button><button class="button" type="button" data-action="edit-lesson" data-id="' + attr(prepared.id) + '"><i data-lucide="pencil"></i>Editar aula</button>'
      : '<button class="button button-dark" type="button" data-action="create-lesson" data-course="' + attr(entry.courseId) + '" data-schedule="' + attr(entry.id) + '" data-date="' + attr(dateValue) + '" data-start="' + attr(entry.start) + '" data-end="' + attr(entry.end) + '" data-room="' + attr(entry.room || "") + '" data-type="' + attr(entry.type || "T") + '"><i data-lucide="presentation"></i>Preparar ' + esc(formatDate(dateValue)) + '</button>';
    openModal("Bloco do horário", '<div class="card" style="background:' + safeColor(course && course.color) + '"><span class="badge badge-dark">' + esc(lessonTypeLabel(entry.type)) + '</span><h2 style="margin:14px 0 5px">' + esc(course ? course.name : "Cadeira") + '</h2><p class="card-subtitle">' + WEEKDAYS[entry.weekday] + ' · ' + esc(entry.start) + '–' + esc(entry.end) + (entry.room ? " · " + esc(entry.room) : "") + '</p></div><div class="form-note" style="margin-top:14px">Próxima ocorrência: <strong>' + esc(formatLongDate(dateValue)) + '</strong>' + (prepared ? ' · preparada como “' + esc(prepared.title) + '”' : ' · ainda sem nome de aula') + '</div><div class="list-actions" style="margin-top:15px">' + primary + '<button class="button button-danger" type="button" data-action="delete-entity" data-kind="schedule" data-id="' + attr(entry.id) + '"><i data-lucide="trash-2"></i>Remover do horário</button></div>');
  }

  function viewArchive(id) {
    var semester = semesterById(id);
    if (!semester) return;
    var courses = state.courses.filter(function (course) { return course.semesterId === semester.id; });
    var body = '<div class="finish-card"><span class="badge badge-dark">Arquivo</span><h2 style="margin:12px 0 4px">' + esc(semester.name) + '</h2><p class="card-subtitle">' + esc(semester.academicYear) + ' · consulta preservada</p></div><div class="list-stack">' + (courses.length ? courses.map(function (course) {
      return '<div class="list-row"><span class="list-icon" style="background:' + safeColor(course.color) + '"><i data-lucide="book-open"></i></span><span class="list-content"><strong>' + esc(course.name) + '</strong><small>' + (Number(course.ects) || 0) + ' ECTS · ' + state.lessons.filter(function (lesson) { return lesson.courseId === course.id; }).length + ' aulas</small></span><button class="row-button" type="button" data-route="course" data-id="' + attr(course.id) + '"><i data-lucide="arrow-right"></i></button></div>';
    }).join("") : emptyState("archive", "Sem cadeiras", "Este semestre não tem cadeiras guardadas.", null)) + "</div>";
    openModal("Semestre arquivado", body);
  }

  function guidedTourSteps() {
    return [
      { route: "settings", selector: ".page-head", page: "Admin", title: "Admin e dados", copy: "É aqui que configuras o semestre, geres os dados locais e inicias ações administrativas." },
      { route: "settings", selector: ".settings-grid", page: "Admin", title: "Configuração do sistema", copy: "Cada cartão trata de uma área: perfil, semestre, JSON, atividade simulada, armazenamento e segurança." },
      { route: "home", selector: ".hero-card", page: "Hoje", title: "Resumo do dia", copy: "A página inicial adapta-se ao momento: aula em curso, próxima aula, tarefas, avaliações e progresso." },
      { route: "home", selector: ".live-card, .beonline-card", page: "Hoje", title: "Aula em direto e BEFIRST™", copy: "O horário ativa a aula atual. O BEFIRST™ acompanha as aulas que já reviste através do respetivo quiz." },
      { route: "home", selector: ".metric-card", page: "Hoje", title: "Indicadores rápidos", copy: "Média ECTS, tarefas, avaliações e perguntas anteriores são atualizados a partir dos teus dados." },
      { route: "courses", selector: ".course-grid, .empty-state", page: "Cadeiras", title: "Cadeiras do semestre", copy: "Abre uma cadeira para consultar aulas, materiais, avaliações, perguntas, quizzes e notas." },
      { route: "planner", plannerMode: "calendar", selector: ".planner-mode-control", page: "Calendário", title: "Horário ou calendário", copy: "O Horário guarda os blocos recorrentes. O Calendário combina aulas, testes, eventos e tarefas com data." },
      { route: "planner", plannerMode: "calendar", selector: ".calendar-view-control", page: "Calendário", title: "Quatro vistas", copy: "Alterna entre Dia, 3 dias, Semana e Mês. As setas avançam exatamente o intervalo selecionado." },
      { route: "planner", plannerMode: "calendar", selector: ".calendar-card", page: "Calendário", title: "Agenda académica", copy: "Uma aula preparada aparece pelo nome. Os blocos ainda não preparados continuam visíveis através do horário." },
      { route: "planner", plannerMode: "study-day", selector: ".study-day-shell", page: "Dia de estudo", title: "Blocos de tempo", copy: "Arrasta tarefas, aulas, quizzes e avaliações para uma hora. Em ecrãs táteis, usa o botão de agendar." },
      { route: "study", selector: ".page-head", page: "Estudar", title: "Centro de estudo", copy: "Aqui encontras aulas por rever, perguntas de testes anteriores e todos os quizzes disponíveis." },
      { route: "study", selector: ".study-hours-card", page: "Estudar", title: "Horas por cadeira", copy: "A estimativa distribui as horas semanais por ECTS, avaliações próximas e trabalho pendente." },
      { route: "grades", selector: ".target-card", page: "Notas", title: "Média ECTS", copy: "A média global usa a nota atual de cada cadeira e os respetivos ECTS. Cada nota mantém a avaliação de origem." },
      { route: "canteen", selector: ".canteen-days, .canteen-loading", page: "Cantina", title: "Ementa oficial", copy: "Escolhe o dia e confirma quando a informação da SAS NOVA foi consultada." },
      { route: "canteen", selector: ".canteen-meal-grid, .canteen-loading", page: "Cantina", title: "Sopa, pratos e alergénios", copy: "Almoço e jantar mostram a sopa primeiro, depois todas as opções, informação alimentar e alergénios." },
      { route: "settings", selector: ".quick-grid", page: "Admin", title: "Adicionar conteúdo", copy: "Usa estes atalhos para criar aulas, carregar PDFs, importar testes anteriores, guardar perguntas, quizzes, notas e avaliações." }
    ];
  }

  function startGuidedTour() {
    closeModal();
    guidedTour = {
      index: 0,
      returnRoute: clone(route),
      plannerView: state.settings.plannerView,
      calendarView: state.settings.calendarView
    };
    renderGuidedTourStep();
  }

  function stopGuidedTour(restore) {
    var tour = guidedTour;
    var root = document.getElementById("guidedTourRoot");
    if (root) root.remove();
    guidedTour = null;
    if (!tour || restore === false) return;
    state.settings.plannerView = tour.plannerView;
    state.settings.calendarView = tour.calendarView;
    route = tour.returnRoute || { name: "settings", id: null, tab: "overview" };
    render();
    history.replaceState(null, "", "#" + route.name + (route.id ? "/" + route.id : "") + (route.tab && route.tab !== "overview" ? "/" + route.tab : ""));
  }

  function positionGuidedTour(step) {
    if (!guidedTour) return;
    var oldRoot = document.getElementById("guidedTourRoot");
    if (oldRoot) oldRoot.remove();
    var target = document.querySelector(step.selector) || view;
    if (target.scrollIntoView) {
      try { target.scrollIntoView({ block: "center", behavior: "smooth" }); } catch (error) {}
    }
    setTimeout(function () {
      if (!guidedTour) return;
      var rect = target.getBoundingClientRect();
      var root = document.createElement("div");
      root.id = "guidedTourRoot";
      root.className = "guided-tour-root";
      var steps = guidedTourSteps();
      root.innerHTML = '<div class="guided-tour-blocker"></div><div class="guided-tour-highlight" aria-hidden="true"></div><aside class="guided-tour-popover" role="dialog" aria-modal="true" aria-labelledby="guidedTourTitle"><div class="guided-tour-top"><span>' + esc(step.page) + '</span><strong>' + (guidedTour.index + 1) + ' / ' + steps.length + '</strong></div><h2 id="guidedTourTitle">' + esc(step.title) + '</h2><p>' + esc(step.copy) + '</p><div class="guided-tour-progress"><span style="width:' + ((guidedTour.index + 1) / steps.length * 100) + '%"></span></div><div class="guided-tour-actions"><button class="button button-ghost button-small" type="button" data-action="tour-close">Sair</button><div><button class="button button-small" type="button" data-action="tour-back" ' + (guidedTour.index === 0 ? "disabled" : "") + '><i data-lucide="arrow-left"></i>Anterior</button><button class="button button-dark button-small" type="button" data-action="tour-next">' + (guidedTour.index === steps.length - 1 ? "Concluir" : "Seguinte") + '<i data-lucide="arrow-right"></i></button></div></div></aside>';
      document.body.appendChild(root);
      var highlight = root.querySelector(".guided-tour-highlight");
      var visible = rect.width > 0 && rect.height > 0;
      highlight.style.left = Math.max(8, rect.left - 6) + "px";
      highlight.style.top = Math.max(8, rect.top - 6) + "px";
      highlight.style.width = Math.max(40, visible ? rect.width + 12 : window.innerWidth - 32) + "px";
      highlight.style.height = Math.max(40, visible ? Math.min(rect.height + 12, window.innerHeight - 32) : 90) + "px";
      refreshIcons(root);
    }, 180);
  }

  function renderGuidedTourStep() {
    if (!guidedTour) return;
    var steps = guidedTourSteps();
    var step = steps[guidedTour.index];
    if (!step) { stopGuidedTour(true); return; }
    route = { name: step.route, id: null, tab: "overview" };
    if (step.plannerMode) state.settings.plannerView = step.plannerMode;
    if (step.route === "planner") state.settings.calendarView = "month";
    render();
    history.replaceState(null, "", "#" + route.name);
    positionGuidedTour(step);
  }

  function openTutorial() {
    startGuidedTour();
  }

  function downloadJSON() {
    var exported = clone(state);
    exported.meta = Object.assign({}, exported.meta, { exportedAt: new Date().toISOString(), exportedFromApp: true, syncMode: "merge" });
    var blob = new Blob([JSON.stringify(exported, null, 2)], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = "academic-data.json";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    toast("Backup JSON criado. Os bytes dos PDFs locais não são incluídos.");
  }

  async function applyImportedJSON(mode) {
    if (!pendingImport) return;
    state = mode === "replace" ? normalizeState(pendingImport) : mergeExternal(state, pendingImport);
    state.meta.source = "json-import";
    await save(true);
    pendingImport = null;
    onboarding = null;
    closeModal();
    render();
    if (!state.profile.onboardingComplete || !state.currentSemesterId) startOnboarding("first");
    else toast("JSON importado e aplicado.");
  }

  async function archiveCurrentSemester() {
    var semester = currentSemester();
    if (!semester) return;
    if (!window.confirm("Arquivar “" + semester.name + "”? As cadeiras e materiais ficam disponíveis apenas para consulta.")) return;
    semester.archived = true;
    semester.archivedAt = new Date().toISOString();
    state.currentSemesterId = null;
    await save(true);
    render();
    startOnboarding("new-semester");
  }

  async function deleteEntity(kind, id) {
    if (ENTITY_ARRAYS.indexOf(kind) < 0) return;
    var item = state[kind].find(function (entry) { return entry.id === id; });
    if (!item) return;
    if (kind === "schedule" && state.lessons.some(function (lesson) { return lesson.scheduleId === item.id || (!lesson.scheduleId && lessonMatchesSchedule(lesson, item)); })) {
      toast("Este bloco já tem aulas preparadas. Edita essas aulas antes de o removeres.", "warning");
      return;
    }
    if (kind === "assessments" && state.grades.some(function (grade) { return grade.assessmentId === item.id; })) {
      toast("Esta avaliação já tem notas. Remove primeiro as notas associadas.", "warning");
      return;
    }
    var linkedPastQuestions = kind === "pastExams" ? state.questions.filter(function (question) { return question.pastExamId === id; }) : [];
    var confirmation = linkedPastQuestions.length ? "Remover este teste anterior e as " + linkedPastQuestions.length + " pergunta(s) associadas?" : "Remover este item? Esta ação não pode ser desfeita.";
    if (!window.confirm(confirmation)) return;
    if (kind === "materials" && item.remoteFile && item.remoteFile.path && Sync && Sync.getStatus().configured) {
      setManualSyncActivity("A apagar o material…", "A remover também o ficheiro do repositório privado.", 25, true);
      try { await Sync.deleteFile(item.remoteFile); } catch (error) { finishManualSyncActivity(false); toast(error.message || "Não foi possível apagar o ficheiro remoto.", "error"); return; }
    }
    if (kind === "materials" && item.blobId) await DB.deleteFile(item.blobId);
    var imageOwners = kind === "questions" || kind === "events" ? [item] : linkedPastQuestions;
    for (var ownerIndex = 0; ownerIndex < imageOwners.length; ownerIndex += 1) {
      var ownerImages = normalizeImageRefs(imageOwners[ownerIndex].images);
      for (var imageIndex = 0; imageIndex < ownerImages.length; imageIndex += 1) {
        if (ownerImages[imageIndex].blobId) await DB.deleteFile(ownerImages[imageIndex].blobId);
      }
    }
    if (kind === "pastExams") state.questions = state.questions.filter(function (question) { return question.pastExamId !== id; });
    state[kind] = state[kind].filter(function (entry) { return entry.id !== id; });
    await save(true);
    if (kind === "materials" && manualSyncActivity) finishManualSyncActivity(true);
    closeModal();
    render();
    toast("Item removido.");
  }

  function addQuickReview(courseId, lessonId) {
    var course = courseById(courseId);
    var lesson = lessonById(lessonId);
    var tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    var title = "Rever " + (lesson ? lesson.title : course ? course.name : "aula");
    var exists = state.tasks.some(function (task) { return !task.done && task.lessonId === lessonId && task.type === "review"; });
    if (exists) { toast("Esta revisão já está na agenda.", "warning"); return; }
    state.tasks.push({ id: uid("task"), semesterId: state.currentSemesterId, courseId: courseId || null, lessonId: lessonId || null, title: title, type: "review", dueDate: todayISO(tomorrow), dueTime: "18:00", priority: "normal", done: false, createdAt: new Date().toISOString() });
    save(true).then(function () { render(); toast("“" + title + "” adicionada para amanhã."); });
  }

  function buildSearchIndex() {
    var semesterId = state.currentSemesterId;
    var items = [];
    activeCourses().forEach(function (course) { items.push({ type: "Cadeira", title: course.name, meta: course.code + " · " + course.ects + " ECTS", route: "course", id: course.id, icon: "library-big" }); });
    state.lessons.filter(function (item) { return item.semesterId === semesterId; }).forEach(function (lesson) { var course = courseById(lesson.courseId); items.push({ type: "Aula", title: lesson.title, meta: (course ? course.name : "") + " · " + formatDate(lesson.date), route: "lesson", id: lesson.id, icon: "presentation" }); });
    state.materials.filter(function (item) { return item.semesterId === semesterId; }).forEach(function (material) { var course = courseById(material.courseId); items.push({ type: "Material", title: material.title, meta: (course ? course.name : "") + " · " + (material.academicYear || ""), action: "open-material", id: material.id, icon: "file-text" }); });
    state.questions.filter(function (item) { return item.semesterId === semesterId; }).forEach(function (question) { items.push({ type: "Pergunta", title: question.prompt, meta: (question.academicYear || "") + " · " + (question.assessmentLabel || "Teste anterior"), action: "show-question-answer", id: question.id, icon: "message-circle-question" }); });
    state.pastExams.filter(function (item) { return item.semesterId === semesterId; }).forEach(function (exam) { var course = courseById(exam.courseId); items.push({ type: "Teste anterior", title: exam.title, meta: (course ? course.name + " · " : "") + (exam.academicYear || ""), route: "course", id: exam.courseId, icon: "file-json-2" }); });
    state.assessments.filter(function (item) { return item.semesterId === semesterId; }).forEach(function (assessment) { items.push({ type: "Avaliação", title: assessment.title, meta: formatDate(assessment.date) + " · " + assessment.type, route: "planner", icon: assessmentIcon(assessment.type) }); });
    return items;
  }

  function updateSearch() {
    var query = String(searchInput.value || "").trim().toLocaleLowerCase("pt-PT");
    if (!query) { searchResults.hidden = true; searchResults.innerHTML = ""; return; }
    var results = buildSearchIndex().filter(function (item) { return (item.title + " " + item.meta + " " + item.type).toLocaleLowerCase("pt-PT").indexOf(query) >= 0; }).slice(0, 9);
    searchResults.innerHTML = results.length ? results.map(function (item) {
      return '<button type="button" data-action="search-open" data-route-target="' + attr(item.route || "") + '" data-action-target="' + attr(item.action || "") + '" data-id="' + attr(item.id || "") + '"><span class="result-icon"><i data-lucide="' + item.icon + '"></i></span><span><strong>' + esc(item.title) + '</strong><small>' + esc(item.type + " · " + item.meta) + '</small></span><i data-lucide="arrow-up-right"></i></button>';
    }).join("") : '<div class="empty-state" style="min-height:130px"><h3>Sem resultados</h3><p>Tenta outro termo.</p></div>';
    searchResults.hidden = false;
    refreshIcons(searchResults);
  }

  async function startNewSemester() {
    var semester = currentSemester();
    if (semester) {
      if (!window.confirm("Para iniciar um novo semestre, “" + semester.name + "” será arquivado. Continuar?")) return;
      semester.archived = true;
      semester.archivedAt = new Date().toISOString();
      state.currentSemesterId = null;
      await save(true);
    }
    closeModal();
    render();
    startOnboarding("new-semester");
  }

  async function resetApp() {
    if (!window.confirm("Apagar todos os dados locais e PDFs deste dispositivo? O academic-data.json externo não será apagado.")) return;
    await DB.clearAll();
    state = defaultState();
    await loadExternalJSON({ force: true, silent: true });
    if (!state) state = defaultState();
    render();
    startOnboarding("first");
  }

  function openCopyStudyDay() {
    var target = state.settings.studyPlanDate || todayISO();
    var previous = addCalendarDays(target, -7);
    openModal("Copiar rotina de estudo", '<div class="form-grid"><div class="field"><label>Copiar blocos de</label><input id="copyStudySourceDate" type="date" value="' + attr(previous) + '"></div><div class="field"><label>Para</label><input type="date" value="' + attr(target) + '" disabled></div></div><p class="form-note" style="margin-top:14px">Os blocos já existentes no dia de destino são preservados; duplicados com a mesma hora e nome não são criados.</p><button class="button button-dark" style="margin-top:14px" type="button" data-action="apply-copy-study-day"><i data-lucide="copy"></i>Copiar blocos</button>');
  }

  async function copyStudyDay(sourceDate) {
    var targetDate = state.settings.studyPlanDate || todayISO();
    if (!sourceDate || sourceDate === targetDate) { toast("Escolhe outro dia como origem.", "warning"); return; }
    var sourceBlocks = studyBlocksForDate(sourceDate);
    if (!sourceBlocks.length) { toast("O dia escolhido não tem blocos para copiar.", "warning"); return; }
    var targetBlocks = studyBlocksForDate(targetDate);
    var additions = sourceBlocks.filter(function (block) {
      return !targetBlocks.some(function (target) { return target.start === block.start && target.end === block.end && target.title === block.title; });
    }).map(function (block) { return Object.assign({}, clone(block), { id: uid("studyblock"), date: targetDate, completed: false }); });
    if (!additions.length) { toast("Esta rotina já existe no dia de destino.", "warning"); return; }
    state.studyBlocks = state.studyBlocks.concat(additions);
    await save(true); closeModal(); render(); toast(additions.length + " bloco(s) copiados.");
  }


  async function connectGitFromOnboarding() {
    if (!Sync) { toast("O módulo de sincronização não foi carregado.", "error"); return; }
    var current = Sync.getConfig();
    var endpoint = window.prompt("Endereço do Cloudflare Worker", current.endpoint || "https://twenty-git-sync.TEU-SUBDOMINIO.workers.dev");
    if (endpoint == null) return;
    endpoint = endpoint.trim();
    if (!/^https:\/\//i.test(endpoint)) { toast("O endereço do Worker tem de começar por https://", "warning"); return; }
    var key = window.prompt("Chave privada de sincronização", current.key || "");
    if (key == null) return;
    key = key.trim();
    if (!key) { toast("Falta a chave privada de sincronização.", "warning"); return; }
    setManualSyncActivity("A ligar ao Git…", "A validar a chave e a procurar os teus dados sincronizados.", 12, true);
    try {
      await Sync.configure(endpoint, key);
      setManualSyncActivity("A descarregar os dados…", "O Git tem prioridade neste primeiro arranque.", 42, true);
      var remoteState = await Sync.forcePull({ dispatch: false });
      setManualSyncActivity("A aplicar neste dispositivo…", "A preparar as cadeiras, tarefas e definições.", 78, true);
      state = normalizeState(remoteState);
      await DB.saveState(state, { skipSync: true });
      await Sync.adoptRemoteState(state);
      onboarding = null;
      closeModal();
      setRoute("home");
      finishManualSyncActivity(true);
      toast("Force pull concluído. Este dispositivo já usa os dados do Git.");
    } catch (error) {
      finishManualSyncActivity(false);
      renderOnboarding();
      toast("Não foi possível carregar os dados do Git: " + error.message, "error");
    }
  }

  async function configureGitSync() {
    if (!Sync) { toast("O módulo de sincronização não foi carregado.", "error"); return; }
    var current = Sync.getConfig();
    var endpoint = window.prompt("Endereço do Cloudflare Worker", current.endpoint || "https://twenty-git-sync.TEU-SUBDOMINIO.workers.dev");
    if (endpoint == null) return;
    endpoint = endpoint.trim();
    if (!/^https:\/\//i.test(endpoint)) { toast("O endereço do Worker tem de começar por https://", "warning"); return; }
    var key = window.prompt("Chave privada de sincronização", current.key || "");
    if (key == null) return;
    key = key.trim();
    if (!key) { toast("Falta a chave privada de sincronização.", "warning"); return; }
    setManualSyncActivity("A configurar a sincronização…", "A testar o Worker e a confirmar o repositório privado.", 18, true);
    try {
      await Sync.configure(endpoint, key);
      setManualSyncActivity("A criar ou atualizar o Git…", "A sincronizar os dados deste dispositivo com segurança.", 52, true);
      await Sync.syncNow(state, defaultState());
      updateGitSyncCard(Sync.getStatus());
      finishManualSyncActivity(true);
      toast("Git sincronizado com sucesso.");
    } catch (error) {
      updateGitSyncCard(Sync.getStatus());
      finishManualSyncActivity(false);
      toast("Não foi possível ligar ao Git: " + error.message, "error");
    }
  }

  async function syncGitNow() {
    if (!Sync || !Sync.getStatus().configured) { await configureGitSync(); return; }
    setManualSyncActivity("A sincronizar dados…", "A enviar alterações e a confirmar a versão final no Git.", null, false);
    try {
      await Sync.syncNow(state, defaultState());
      updateGitSyncCard(Sync.getStatus());
      finishManualSyncActivity(true);
      toast("Alterações enviadas e dados atualizados.");
    } catch (error) {
      updateGitSyncCard(Sync.getStatus());
      finishManualSyncActivity(false);
      toast("A sincronização ficou na fila: " + error.message, "warning");
    }
  }

  function openForceGitConfirmation(direction) {
    if (!Sync || !Sync.getStatus().configured) { configureGitSync(); return; }
    var isPull = direction === "pull";
    var title = isPull ? "Forçar pull?" : "Forçar push?";
    var icon = isPull ? "arrow-down-to-line" : "arrow-up-to-line";
    var cardClass = isPull ? "card-violet" : "card-pink";
    var copy = isPull
      ? "Os dados deste dispositivo serão substituídos pela versão atual do Git. As alterações locais ainda não enviadas serão descartadas."
      : "A versão atual do Git será substituída pelos dados deste dispositivo. Alterações mais recentes feitas noutro dispositivo podem ser perdidas.";
    var action = isPull ? "confirm-force-git-pull" : "confirm-force-git-push";
    openModal(title, '<article class="card ' + cardClass + ' force-sync-warning"><span class="metric-icon"><i data-lucide="' + icon + '"></i></span><h3>' + (isPull ? "O Git fica com prioridade" : "Este dispositivo fica com prioridade") + '</h3><p class="card-subtitle">' + copy + '</p></article>', {
      footer: '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Cancelar</button><button class="button ' + (isPull ? "button-dark" : "button-danger") + '" type="button" data-action="' + action + '"><i data-lucide="' + icon + '"></i>' + (isPull ? "Forçar pull" : "Forçar push") + '</button></footer>'
    });
  }

  async function forceGitPull() {
    if (!Sync || !Sync.getStatus().configured) { await configureGitSync(); return; }
    closeModal();
    setManualSyncActivity("A fazer force pull…", "A descarregar a versão do Git. Não feches a aplicação.", 18, true);
    try {
      var remoteState = await Sync.forcePull({ dispatch: false });
      setManualSyncActivity("A aplicar os dados…", "A versão do Git está a substituir os dados deste dispositivo.", 72, true);
      state = normalizeState(remoteState);
      await DB.saveState(state, { skipSync: true });
      await Sync.adoptRemoteState(state);
      onboarding = null;
      render();
      finishManualSyncActivity(true);
      toast("Force pull concluído. Os dados locais foram atualizados.");
    } catch (error) {
      finishManualSyncActivity(false);
      updateGitSyncCard(Sync.getStatus());
      toast("Não foi possível forçar o pull: " + error.message, "error");
    }
  }

  async function forceGitPush() {
    if (!Sync || !Sync.getStatus().configured) { await configureGitSync(); return; }
    closeModal();
    setManualSyncActivity("A fazer force push…", "A preparar os dados deste dispositivo para substituir a versão do Git.", 18, true);
    try {
      var confirmedState = await Sync.forcePush(state, { dispatch: false });
      setManualSyncActivity("A confirmar o commit…", "O GitHub já recebeu os dados; falta confirmar a versão final.", 78, true);
      state = normalizeState(confirmedState);
      await DB.saveState(state, { skipSync: true });
      await Sync.adoptRemoteState(state);
      render();
      finishManualSyncActivity(true);
      toast("Force push concluído. Foi criado um novo commit no Git.");
    } catch (error) {
      finishManualSyncActivity(false);
      updateGitSyncCard(Sync.getStatus());
      toast("Não foi possível forçar o push: " + error.message, "error");
    }
  }

  async function handleAction(button) {
    var action = button.getAttribute("data-action");
    if (!action) return;
    if (action === "close-modal") {
      if (quizRuntime) quizRuntime = null;
      closeModal();
      if (onboarding) renderOnboarding();
    } else if (action === "quick-add") {
      openQuickAdd();
    } else if (action === "add-course") {
      if (!state.currentSemesterId) startOnboarding("new-semester");
      else openEntityForm("course", {});
    } else if (action === "import-courses") {
      if (!state.currentSemesterId) startOnboarding("new-semester");
      else openEntityForm("course-import", {});
    } else if (action === "edit-course") {
      openEntityForm("course", { id: button.dataset.id });
    } else if (action === "add-evaluation-component") {
      var evaluationForm = button.closest("form");
      var evaluationList = evaluationForm && evaluationForm.querySelector(".evaluation-builder-list");
      if (evaluationList) {
        evaluationList.insertAdjacentHTML("beforeend", renderEvaluationComponentRow({ label: "Nova componente", kind: "other", count: 1, weight: 0 }));
        updateEvaluationBuilderSummary(evaluationForm);
        refreshIcons(evaluationList);
      }
    } else if (action === "remove-evaluation-component") {
      var evaluationRow = button.closest(".evaluation-builder-row");
      var courseForm = button.closest("form");
      if (evaluationRow) evaluationRow.remove();
      updateEvaluationBuilderSummary(courseForm);
    } else if (action === "edit-lesson") {
      openEntityForm("lesson", { id: button.dataset.id });
    } else if (action === "create-lesson") {
      var lessonPreset = {};
      if (button.dataset.course) lessonPreset.courseId = button.dataset.course;
      if (button.dataset.schedule) lessonPreset.scheduleId = button.dataset.schedule;
      if (button.dataset.date) lessonPreset.date = button.dataset.date;
      if (button.dataset.start) lessonPreset.start = button.dataset.start;
      if (button.dataset.end) lessonPreset.end = button.dataset.end;
      if (button.dataset.room) lessonPreset.room = button.dataset.room;
      if (button.dataset.type) lessonPreset.type = button.dataset.type;
      if (button.dataset.title) lessonPreset.title = button.dataset.title;
      openEntityForm("lesson", lessonPreset);
    } else if (action === "create-lesson-from-live") {
      var liveEntry = state.schedule.find(function (item) { return item.id === button.dataset.id; });
      if (liveEntry) openEntityForm("lesson", { courseId: liveEntry.courseId, scheduleId: liveEntry.id, date: todayISO(), start: liveEntry.start, end: liveEntry.end, room: liveEntry.room, type: liveEntry.type, title: (liveEntry.type || "Aula") + " · " + formatDate(todayISO()) });
    } else if (action === "ai-pick-pptx") {
      if (pptxInput) pptxInput.click();
    } else if (action === "ai-clear-draft") {
      aiDraft = null;
      render();
    } else if (action === "ai-generate") {
      await generateAIProject();
    } else if (action === "ai-cancel") {
      if (aiTransferRequest) { try { aiTransferRequest.abort(); } catch (_) {} aiTransferRequest = null; }
      if (AI && AI.resetWorker) AI.resetWorker();
      aiBusy = false;
      clearAIProgress();
      render();
      toast("Geração cancelada.", "warning");
    } else if (action === "ai-open-project") {
      openAIProject(button.dataset.id);
    } else if (action === "ai-use-project") {
      useAIProject(button.dataset.id);
    } else if (action === "ai-download-pptx") {
      await downloadAIProjectFile(button.dataset.id);
    } else if (action === "ai-open-slide") {
      openAISlide(button.dataset.project, button.dataset.slide);
    } else if (action === "ai-start-quiz") {
      var aiProject = aiProjectById(button.dataset.id);
      if (aiProject && aiProject.quizId) { closeModal(); startQuiz(aiProject.quizId); }
    } else if (action === "ai-delete-project") {
      confirmDeleteAIProject(button.dataset.id);
    } else if (action === "confirm-ai-delete-project") {
      await deleteAIProject(button.dataset.id);
    } else if (action === "sync-material") {
      await syncExistingMaterial(button.dataset.id);
    } else if (action === "lesson-ai") {
      openLessonAIModal(button.dataset.lesson || route.id, button.dataset.output || "quiz", button.dataset.material || "");
    } else if (action === "run-lesson-ai") {
      await runLessonAI();
    } else if (action === "add-material") {
      openEntityForm("material", { courseId: button.dataset.course || "", lessonId: button.dataset.lesson || "" });
    } else if (action === "add-task") {
      openEntityForm("task", { courseId: button.dataset.course || "", lessonId: button.dataset.lesson || "" });
    } else if (action === "add-assessment") {
      openEntityForm("assessment", { courseId: button.dataset.course || "" });
    } else if (action === "edit-assessment") {
      closeModal();
      openEntityForm("assessment", { id: button.dataset.id });
    } else if (action === "add-event") {
      openEntityForm("event", {});
    } else if (action === "edit-event") {
      openEntityForm("event", { id: button.dataset.id });
    } else if (action === "add-past-exam") {
      openEntityForm("past-exam-import", { courseId: button.dataset.course || "" });
    } else if (action === "add-question") {
      openEntityForm("question", { courseId: button.dataset.course || "", lessonId: button.dataset.lesson || "", pastExamId: button.dataset.pastExam || "" });
    } else if (action === "edit-question") {
      openEntityForm("question", { id: button.dataset.id });
    } else if (action === "add-quiz") {
      openEntityForm("quiz", { courseId: button.dataset.course || "", lessonId: button.dataset.lesson || "" });
    } else if (action === "generate-past-quiz") {
      await generateQuizFromPast(button.dataset.lesson || "", true);
    } else if (action === "add-past-to-quiz") {
      openPastQuestionPicker(button.dataset.id);
    } else if (action === "add-quiz-question") {
      openEntityForm("quiz-question", { quizId: button.dataset.id });
    } else if (action === "add-grade") {
      openEntityForm("grade", { courseId: button.dataset.course || "", assessmentId: button.dataset.assessment || "", lessonId: button.dataset.lesson || "" });
    } else if (action === "grade-simulator") {
      openGradeSimulator(button.dataset.course || "");
    } else if (action === "add-schedule") {
      openEntityForm("schedule", { courseId: button.dataset.course || "" });
    } else if (action === "add-study-block") {
      openEntityForm("study-block", { date: state.settings.studyPlanDate || todayISO() });
    } else if (action === "edit-study-block") {
      openEntityForm("study-block", { id: button.dataset.id });
    } else if (action === "schedule-study-source") {
      openEntityForm("study-block", { sourceType: button.dataset.sourceType, sourceId: button.dataset.sourceId, date: state.settings.studyPlanDate || todayISO() });
    } else if (action === "toggle-study-block") {
      var studyBlock = state.studyBlocks.find(function (item) { return item.id === button.dataset.id; });
      if (studyBlock) { studyBlock.completed = !studyBlock.completed; await save(true); render(); }
    } else if (action === "study-planner-settings") {
      openEntityForm("study-planner-settings", {});
    } else if (action === "weekly-review") {
      openEntityForm("weekly-review", {});
    } else if (action === "auto-fill-study-day") {
      await autoFillStudyDay();
    } else if (action === "copy-study-day") {
      openCopyStudyDay();
    } else if (action === "apply-copy-study-day") {
      var copySource = modalRoot.querySelector("#copyStudySourceDate");
      await copyStudyDay(copySource && copySource.value);
    } else if (action === "study-date-shift") {
      state.settings.studyPlanDate = addCalendarDays(state.settings.studyPlanDate || todayISO(), Number(button.dataset.delta || 0));
      await save(true); render();
    } else if (action === "edit-profile") {
      openEntityForm("profile", {});
    } else if (action === "edit-lesson-notes") {
      openEntityForm("lesson-notes", { id: button.dataset.id });
    } else if (action === "toggle-task") {
      var task = state.tasks.find(function (item) { return item.id === button.dataset.id; });
      if (task && task.type === "lesson-quiz") {
        await doLessonQuiz(task.lessonId || "");
      } else if (task) {
        task.done = !task.done;
        await save(true);
        render();
        toast(task.done ? "Tarefa concluída." : "Tarefa reaberta.");
      }
    } else if (action === "toggle-mastery") {
      var lesson = lessonById(button.dataset.id);
      if (lesson) { lesson.mastered = !lesson.mastered; await save(true); render(); toast(lesson.mastered ? "Aula marcada como dominada." : "Aula voltou à lista de revisão."); }
    } else if (action === "quick-review") {
      addQuickReview(button.dataset.course || "", button.dataset.lesson || "");
    } else if (action === "course-tab") {
      closeModal();
      setRoute("course", button.dataset.id, button.dataset.tab);
    } else if (action === "open-material") {
      await openMaterial(button.dataset.id);
    } else if (action === "show-question-answer") {
      showQuestionAnswer(button.dataset.id);
    } else if (action === "open-image") {
      await openImage(button.dataset.imageId);
    } else if (action === "start-quiz") {
      startQuiz(button.dataset.id);
    } else if (action === "do-beonline-quiz") {
      await doLessonQuiz(button.dataset.lesson || "");
    } else if (action === "beonline-next") {
      var pendingOnline = beOnlineStatus().pending.slice().sort(function (a, b) {
        return String(a.date || "").localeCompare(String(b.date || "")) || String(a.start || "").localeCompare(String(b.start || ""));
      });
      if (pendingOnline.length) await doLessonQuiz(pendingOnline[0].id);
    } else if (action === "quiz-answer") {
      if (quizRuntime) { quizRuntime.selected = Number(button.dataset.index); renderQuizQuestion(); }
    } else if (action === "quiz-reveal") {
      if (quizRuntime) { quizRuntime.revealed = true; renderQuizQuestion(); }
    } else if (action === "quiz-self-rate") {
      if (quizRuntime) {
        quizRuntime.answers.push(Number(button.dataset.value) === 1 ? 1 : 0);
        quizRuntime.selected = null;
        quizRuntime.revealed = false;
        quizRuntime.index += 1;
        renderQuizQuestion();
      }
    } else if (action === "quiz-next") {
      if (quizRuntime && quizRuntime.selected != null) { quizRuntime.answers.push(quizRuntime.selected); quizRuntime.selected = null; quizRuntime.revealed = false; quizRuntime.index += 1; renderQuizQuestion(); }
    } else if (action === "assessment-scope" || action === "study-assessment") {
      showAssessmentScope(button.dataset.id);
    } else if (action === "show-event") {
      showEventDetail(button.dataset.id);
    } else if (action === "show-task") {
      showTaskDetail(button.dataset.id);
    } else if (action === "schedule-detail") {
      showScheduleDetail(button.dataset.id);
    } else if (action === "canteen-day") {
      canteenSelectedDate = button.dataset.date || canteenSelectedDate;
      render();
    } else if (action === "refresh-canteen") {
      var canteenResult = await ensureCanteenMenu(true);
      if (canteenResult.status === "ready") toast("Ementa atualizada a partir da SAS NOVA.");
      else if (canteenResult.data) toast("A SAS NOVA não respondeu; mantive a última ementa guardada.", "warning");
      else toast("Não foi possível atualizar a ementa.", "error");
    } else if (action === "planner-mode") {
      state.settings.plannerView = ["schedule", "calendar", "study-day"].indexOf(button.dataset.mode) >= 0 ? button.dataset.mode : "schedule";
      await save(true);
      if (route.name !== "planner") setRoute("planner");
      else render();
    } else if (action === "calendar-view") {
      state.settings.calendarView = ["day", "three", "week", "month"].indexOf(button.dataset.view) >= 0 ? button.dataset.view : "month";
      await save(true);
      render();
    } else if (action === "calendar-shift") {
      var calendarDelta = Number(button.dataset.delta || 0);
      var activeCalendarView = state.settings.calendarView || "month";
      if (activeCalendarView === "month") {
        var shifted = localDate(calendarCursor) || new Date();
        shifted.setMonth(shifted.getMonth() + calendarDelta);
        calendarCursor = todayISO(shifted);
      } else {
        var dayJump = activeCalendarView === "week" ? 7 : activeCalendarView === "three" ? 3 : 1;
        calendarCursor = addCalendarDays(calendarCursor, calendarDelta * dayJump);
      }
      render();
    } else if (action === "calendar-today") {
      calendarCursor = todayISO();
      render();
    } else if (action === "view-archive") {
      viewArchive(button.dataset.id);
    } else if (action === "delete-entity") {
      await deleteEntity(button.dataset.kind, button.dataset.id);
    } else if (action === "go-courses") {
      setRoute("courses");
    } else if (action === "new-semester") {
      await startNewSemester();
    } else if (action === "archive-semester") {
      await archiveCurrentSemester();
    } else if (action === "configure-git-sync") {
      await configureGitSync();
    } else if (action === "sync-now") {
      await syncGitNow();
    } else if (action === "force-git-pull") {
      openForceGitConfirmation("pull");
    } else if (action === "force-git-push") {
      openForceGitConfirmation("push");
    } else if (action === "confirm-force-git-pull") {
      await forceGitPull();
    } else if (action === "confirm-force-git-push") {
      await forceGitPush();
    } else if (action === "disable-git-sync") {
      if (Sync) Sync.disable();
      render();
      toast("Sincronização Git pausada neste dispositivo.");
    } else if (action === "toggle-campus") {
      state.settings.campusSimulation = !state.settings.campusSimulation;
      await save(true); render();
    } else if (action === "toggle-json-sync") {
      state.settings.jsonSync = !state.settings.jsonSync;
      await save(true); render(); toast(state.settings.jsonSync ? "Sincronização JSON ativada." : "Sincronização JSON pausada.");
    } else if (action === "reload-json") {
      button.disabled = true;
      var changed = await loadExternalJSON({ force: true, silent: false });
      render();
      if (!changed) toast("JSON verificado; não há alterações novas.");
    } else if (action === "export-json") {
      downloadJSON();
    } else if (action === "import-json") {
      importInput.value = "";
      importInput.click();
    } else if (action === "fill-import-example") {
      var importTarget = modalRoot.querySelector("#" + button.dataset.target);
      var examplePayload = button.dataset.kind === "course" ? courseJSONExample() : pastExamJSONExample();
      if (importTarget) importTarget.value = JSON.stringify(examplePayload, null, 2);
      var importForm = button.closest("form");
      if (importForm && button.dataset.kind !== "course") {
        if (importForm.elements.title && !importForm.elements.title.value) importForm.elements.title.value = examplePayload.title;
        if (importForm.elements.academicYear && !importForm.elements.academicYear.value) importForm.elements.academicYear.value = examplePayload.academicYear;
        if (importForm.elements.date && !importForm.elements.date.value) importForm.elements.date.value = examplePayload.date;
        if (importForm.elements.source && !importForm.elements.source.value) importForm.elements.source.value = examplePayload.source;
      }
    } else if (action === "copy-import-prompt") {
      await copyText(importPrompt(button.dataset.kind === "course" ? "course" : "past-exam"));
      toast("Prompt copiado.");
    } else if (action === "apply-import-merge") {
      await applyImportedJSON("merge");
    } else if (action === "apply-import-replace") {
      await applyImportedJSON("replace");
    } else if (action === "show-tutorial") {
      openTutorial();
    } else if (action === "tour-close") {
      stopGuidedTour(true);
    } else if (action === "tour-back") {
      if (guidedTour && guidedTour.index > 0) { guidedTour.index -= 1; renderGuidedTourStep(); }
    } else if (action === "tour-next") {
      if (guidedTour && guidedTour.index < guidedTourSteps().length - 1) { guidedTour.index += 1; renderGuidedTourStep(); }
      else stopGuidedTour(true);
    } else if (action === "reset-app") {
      await resetApp();
    } else if (action === "search-open") {
      searchResults.hidden = true;
      searchInput.value = "";
      document.querySelector(".search-box").classList.remove("is-open");
      if (button.dataset.routeTarget) setRoute(button.dataset.routeTarget, button.dataset.id || null);
      else if (button.dataset.actionTarget === "open-material") await openMaterial(button.dataset.id);
      else if (button.dataset.actionTarget === "show-question-answer") showQuestionAnswer(button.dataset.id);
    } else if (action === "tutorial-next") {
      if (onboarding.tutorialPage < 2) onboarding.tutorialPage += 1;
      else onboarding.step = 1;
      renderOnboarding();
    } else if (action === "onboarding-connect-git") {
      await connectGitFromOnboarding();
    } else if (action === "tutorial-skip") {
      onboarding.tutorialSkipped = true;
      onboarding.step = 1;
      renderOnboarding();
    } else if (action === "onboarding-next") {
      if (captureOnboardingStep(true)) { onboarding.step += 1; renderOnboarding(); }
    } else if (action === "onboarding-back") {
      captureOnboardingStep(false);
      if (onboarding.step === 1 && onboarding.mode === "new-semester") {
        if (window.confirm("Sair da configuração do novo semestre?")) { onboarding = null; closeModal(); render(); }
      } else {
        onboarding.step = Math.max(0, onboarding.step - 1);
        renderOnboarding();
      }
    } else if (action === "add-onboarding-course") {
      captureOnboardingStep(false);
      onboarding.draft.courses.push({ tempId: uid("draftcourse"), name: "", code: "", ects: 6, color: COLORS[onboarding.draft.courses.length % COLORS.length], lessonTypes: ["T", "TP"], evaluation: "Testes | 60 | test\nProjeto | 40 | project\nExame | 0 | exam", examReplacesTests: true });
      renderOnboarding();
    } else if (action === "remove-onboarding-course") {
      captureOnboardingStep(false);
      onboarding.draft.courses.splice(Number(button.dataset.index), 1);
      if (!onboarding.draft.courses.length) onboarding.draft.courses.push({ tempId: uid("draftcourse"), name: "", code: "", ects: 6, color: COLORS[0], lessonTypes: ["T"], evaluation: "", examReplacesTests: false });
      renderOnboarding();
    } else if (action === "add-onboarding-schedule") {
      captureOnboardingStep(false);
      onboarding.draft.schedule.push({ courseRef: onboarding.draft.courses[0] ? onboarding.draft.courses[0].tempId : "", weekday: 1, start: "09:00", end: "10:30", type: "T", room: "" });
      renderOnboarding();
    } else if (action === "remove-onboarding-schedule") {
      captureOnboardingStep(false);
      onboarding.draft.schedule.splice(Number(button.dataset.index), 1);
      renderOnboarding();
    } else if (action === "add-onboarding-assessment") {
      captureOnboardingStep(false);
      onboarding.draft.assessments.push({ courseRef: onboarding.draft.courses[0] ? onboarding.draft.courses[0].tempId : "", title: "", type: "Teste", date: "", time: "10:00", weight: 0 });
      renderOnboarding();
    } else if (action === "remove-onboarding-assessment") {
      captureOnboardingStep(false);
      onboarding.draft.assessments.splice(Number(button.dataset.index), 1);
      renderOnboarding();
    } else if (action === "finish-onboarding") {
      await finishOnboarding();
    }
  }

  function handleDocumentClick(event) {
    var actionButton = event.target.closest("[data-action]");
    if (actionButton) {
      event.preventDefault();
      handleAction(actionButton).catch(function (error) { console.error(error); toast("Ocorreu um erro: " + error.message, "error"); });
      return;
    }
    var routeButton = event.target.closest("[data-route]");
    if (routeButton) {
      event.preventDefault();
      closeModal();
      if (routeButton.getAttribute("data-route") === "planner" && routeButton.dataset.plannerView) {
        state.settings.plannerView = ["schedule", "calendar", "study-day"].indexOf(routeButton.dataset.plannerView) >= 0 ? routeButton.dataset.plannerView : "calendar";
        save(true);
      }
      setRoute(routeButton.getAttribute("data-route"), routeButton.dataset.id || null, routeButton.dataset.tab || "overview");
      return;
    }
    if (!event.target.closest(".search-box") && !event.target.closest(".search-results")) {
      searchResults.hidden = true;
      if (!searchInput.value) document.querySelector(".search-box").classList.remove("is-open");
    }
    if (event.target.classList.contains("modal-layer")) {
      closeModal();
      if (onboarding) renderOnboarding();
    }
  }

  async function handleImportFile(file) {
    if (!file) return;
    try {
      var text = await file.text();
      var parsed = JSON.parse(text);
      if (!parsed || typeof parsed !== "object") throw new Error("Formato inválido");
      pendingImport = parsed;
      openModal("Importar dados JSON", '<p class="onboarding-copy" style="margin-top:0">O ficheiro tem ' + asArray(parsed.courses).length + ' cadeira(s), ' + asArray(parsed.lessons).length + ' aula(s) e ' + asArray(parsed.materials).length + ' material(is).</p><div class="bento-grid"><article class="card card-violet span-6"><h3>Juntar aos dados atuais</h3><p class="card-subtitle">Itens com o mesmo ID são atualizados; os restantes são preservados.</p><button class="button button-dark" style="margin-top:17px" type="button" data-action="apply-import-merge">Juntar</button></article><article class="card card-pink span-6"><h3>Substituir metadados</h3><p class="card-subtitle">Troca o estado académico pelo ficheiro. PDFs locais continuam no dispositivo, mas podem ficar sem ligação.</p><button class="button button-danger" style="margin-top:17px" type="button" data-action="apply-import-replace">Substituir</button></article></div>');
    } catch (error) {
      toast("O ficheiro não é um JSON válido.", "error");
      if (onboarding) renderOnboarding();
    }
  }

  async function init() {
    try {
      state = normalizeState(await DB.getState());
      var hadLocal = !!(state.meta.updatedAt || state.profile.onboardingComplete || state.semesters.length);
      if (!hadLocal) {
        state = null;
        await loadExternalJSON({ force: true, silent: true });
      } else {
        await loadExternalJSON({ silent: true });
      }
      if (!state) state = defaultState();
      state = normalizeState(state);
      if (Sync) {
        var remoteState = await Sync.bootstrap(state, defaultState());
        if (remoteState) {
          state = normalizeState(remoteState);
          await DB.saveState(state, { skipSync: true });
          await Sync.adoptRemoteState(state);
        }
      }
      if (ensureBeOnlineTasks()) await save(true);
      routeFromHash();
      app.setAttribute("aria-busy", "false");
      render();
      clearInterval(beOnlineTimer);
      beOnlineTimer = setInterval(function () {
        if (!state || onboarding) return;
        if (ensureBeOnlineTasks()) {
          save(true).then(function () { render(); toast("A aula terminou: o quiz de revisão está pronto."); });
        }
      }, 60000);
      if (!state.profile.onboardingComplete || !state.currentSemesterId || !activeCourses().length) startOnboarding(state.semesters.length ? "new-semester" : "first");
      if ("serviceWorker" in navigator && location.protocol !== "file:") {
        navigator.serviceWorker.register("sw.js?v=16-ai-slides", { updateViaCache: "none" }).then(function () {
          if (Sync && Sync.getStatus().configured) Sync.startAutoSync();
        }).catch(function () {
          if (Sync && Sync.getStatus().configured) Sync.startAutoSync();
        });
      } else if (Sync && Sync.getStatus().configured) {
        Sync.startAutoSync();
      }
    } catch (error) {
      console.error(error);
      state = defaultState();
      app.setAttribute("aria-busy", "false");
      render();
      startOnboarding("first");
      toast("Os dados locais não foram carregados. Foi iniciado um estado seguro.", "warning");
    }
  }

  document.addEventListener("click", handleDocumentClick);
  document.addEventListener("change", function (event) {
    if (event.target === pptxInput) {
      handleAIPptxFile(event.target.files && event.target.files[0]);
      return;
    }
    if (event.target.matches('[data-role="ai-question-range"]')) {
      var output = document.getElementById("aiQuestionCountOutput");
      if (output) output.textContent = event.target.value;
      return;
    }
    if (!event.target.matches('[data-role="study-plan-date"]')) return;
    state.settings.studyPlanDate = event.target.value || todayISO();
    save(true).then(render);
  });
  document.addEventListener("input", function (event) {
    if (!event.target.matches('[data-role="ai-question-range"]')) return;
    var output = document.getElementById("aiQuestionCountOutput");
    if (output) output.textContent = event.target.value;
  });
  document.addEventListener("dragstart", function (event) {
    var source = event.target.closest("[data-study-source-type][data-study-source-id]");
    if (!source) return;
    draggedStudyPayload = { type: source.dataset.studySourceType, id: source.dataset.studySourceId };
    source.classList.add("is-dragging");
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", JSON.stringify(draggedStudyPayload));
    }
  });
  document.addEventListener("dragend", function (event) {
    var source = event.target.closest("[data-study-source-type][data-study-source-id]");
    if (source) source.classList.remove("is-dragging");
    document.querySelectorAll(".study-drop-slot.is-over").forEach(function (slot) { slot.classList.remove("is-over"); });
    draggedStudyPayload = null;
  });
  document.addEventListener("dragover", function (event) {
    var slot = event.target.closest("[data-study-drop]");
    if (!slot) return;
    event.preventDefault();
    slot.classList.add("is-over");
    if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
  });
  document.addEventListener("dragleave", function (event) {
    var slot = event.target.closest("[data-study-drop]");
    if (slot && !slot.contains(event.relatedTarget)) slot.classList.remove("is-over");
  });
  document.addEventListener("drop", function (event) {
    var slot = event.target.closest("[data-study-drop]");
    if (!slot) return;
    event.preventDefault();
    slot.classList.remove("is-over");
    var payload = draggedStudyPayload;
    if (!payload && event.dataTransfer) {
      try { payload = JSON.parse(event.dataTransfer.getData("text/plain")); } catch (_) { payload = null; }
    }
    if (!payload) return;
    var startValue = slot.dataset.time;
    if (payload.type !== "block") {
      scheduleStudySource(payload.type, payload.id, startValue).catch(function (error) { toast(error.message, "error"); });
      return;
    }
    var block = state.studyBlocks.find(function (item) { return item.id === payload.id; });
    if (!block) return;
    var duration = Math.max(10, timeMinutes(block.end) - timeMinutes(block.start));
    var newStart = timeMinutes(startValue);
    var newEnd = newStart + duration;
    var endLimit = timeMinutes(state.settings.studyDayEnd || "19:00");
    if (newEnd > endLimit) { toast("O bloco não cabe antes do fim do dia.", "warning"); return; }
    var collision = studyBlocksForDate(state.settings.studyPlanDate || todayISO()).some(function (other) { return other.id !== block.id && newStart < timeMinutes(other.end) && newEnd > timeMinutes(other.start); });
    if (collision) { toast("Essa hora já está ocupada.", "warning"); return; }
    block.date = state.settings.studyPlanDate || todayISO();
    block.start = minutesToTime(newStart);
    block.end = minutesToTime(newEnd);
    save(true).then(render);
  });
  modalRoot.addEventListener("submit", function (event) {
    if (event.target.id === "pastQuestionForm") {
      handlePastQuestionSubmit(event).catch(function (error) { console.error(error); setFormError(event.target, "Não foi possível guardar as perguntas."); });
    } else {
      handleEntitySubmit(event);
    }
  });
  modalRoot.addEventListener("input", function (event) {
    var simulatorForm = event.target.closest("#gradeSimulatorForm");
    if (simulatorForm && event.target.matches('[name="simScore"]')) {
      updateGradeSimulator(simulatorForm);
      return;
    }
    var evaluationForm = event.target.closest('#entityForm[data-type="course"]');
    if (evaluationForm && event.target.matches('[data-role="component-weight"]')) updateEvaluationBuilderSummary(evaluationForm);
  });
  modalRoot.addEventListener("change", function (event) {
    var simulatorForm = event.target.closest("#gradeSimulatorForm");
    if (simulatorForm && event.target.matches('[data-role="simulator-course"]')) {
      var simulatorFields = simulatorForm.querySelector("#gradeSimulatorFields");
      if (simulatorFields) simulatorFields.innerHTML = gradeSimulatorAssessmentFields(event.target.value);
      updateGradeSimulator(simulatorForm);
      refreshIcons(simulatorForm);
      return;
    }
    if (event.target.matches('[data-role="local-json-file"]')) {
      var jsonFile = event.target.files && event.target.files[0];
      var targetId = event.target.dataset.target;
      if (jsonFile && targetId) {
        jsonFile.text().then(function (text) {
          var target = modalRoot.querySelector("#" + targetId);
          if (target) target.value = text;
        }).catch(function () { toast("Não foi possível ler o ficheiro JSON.", "error"); });
      }
      return;
    }
    var questionForm = event.target.closest('#entityForm[data-type="question"]');
    if (questionForm && event.target.matches('[name="courseId"]')) {
      var questionCourseId = event.target.value;
      var examSelect = questionForm.elements.pastExamId;
      var lessonSelectForQuestion = questionForm.elements.lessonIds;
      if (examSelect) examSelect.innerHTML = '<option value="">Sem teste associado</option>' + pastExamOptions(questionCourseId, "");
      if (lessonSelectForQuestion) lessonSelectForQuestion.innerHTML = lessonOptions(questionCourseId, [], false);
      return;
    }
    var lessonSelect = event.target.closest('#entityForm[data-type="quiz"] [data-role="quiz-lesson-select"]');
    if (lessonSelect) {
      var lesson = lessonById(lessonSelect.value);
      updateQuizPastQuestionPicker(lessonSelect.value);
      if (lesson) {
        var quizForm = lessonSelect.closest("form");
        var courseSelect = quizForm.querySelector('[name="courseId"]');
        var titleInput = quizForm.querySelector('[name="title"]');
        if (courseSelect) courseSelect.value = lesson.courseId;
        if (titleInput && (!titleInput.value.trim() || /^Quiz · /.test(titleInput.value))) titleInput.value = "Quiz · " + lesson.title;
      }
      return;
    }
    var lessonForm = event.target.closest('#entityForm[data-type="lesson"]');
    if (lessonForm) {
      if (event.target.matches('[data-role="lesson-schedule"]')) applyLessonScheduleSelection(lessonForm);
      else if (event.target.matches('[data-role="lesson-course"], [data-role="lesson-date"], [data-role="lesson-type"]')) updateLessonScheduleOptions(lessonForm, "");
      return;
    }
    var gradeForm = event.target.closest('#entityForm[data-type="grade"]');
    if (gradeForm && event.target.matches('[data-role="grade-target"]')) {
      updateGradeDefenseFields(gradeForm);
      return;
    }
    var assessmentForm = event.target.closest('#entityForm[data-type="assessment"]');
    if (assessmentForm && event.target.matches('[data-role="assessment-course"]')) {
      updateAssessmentLinkedOptions(assessmentForm);
      return;
    }
    if (assessmentForm && event.target.matches('[data-role="assessment-type"], [data-role="assessment-title"]')) {
      updateAssessmentComponentOptions(assessmentForm);
      return;
    }
  });
  document.getElementById("quickAddButton").addEventListener("click", openQuickAdd);
  document.getElementById("profileButton").addEventListener("click", function () { setRoute("settings"); });
  importInput.addEventListener("change", function () { handleImportFile(importInput.files && importInput.files[0]); });
  searchInput.addEventListener("input", function () { clearTimeout(searchTimer); searchTimer = setTimeout(updateSearch, 80); });
  document.querySelector(".search-box").addEventListener("click", function () {
    if (window.innerWidth <= 820) {
      this.classList.add("is-open");
      searchInput.focus();
    }
  });
  searchInput.addEventListener("blur", function () {
    setTimeout(function () {
      if (!searchInput.value && searchResults.hidden) document.querySelector(".search-box").classList.remove("is-open");
    }, 180);
  });
  document.addEventListener("keydown", function (event) {
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") { event.preventDefault(); searchInput.focus(); searchInput.select(); }
    if (event.key === "Escape") {
      if (guidedTour) stopGuidedTour(true);
      else if (!onboarding) closeModal();
      searchResults.hidden = true;
    }
    if ((event.key === "Enter" || event.key === " ") && event.target.matches('.course-card[role="button"]')) { event.preventDefault(); setRoute("course", event.target.dataset.id); }
  });
  window.addEventListener("hashchange", function () { routeFromHash(); render(); });
  window.addEventListener("focus", function () {
    clearTimeout(externalCheckTimer);
    externalCheckTimer = setTimeout(function () {
      if (Sync && Sync.getStatus().configured) return;
      if (!state || !state.settings.jsonSync || onboarding) return;
      loadExternalJSON({ silent: true }).then(async function (changed) {
        var tasksChanged = ensureBeOnlineTasks();
        if (tasksChanged) await save(true);
        if (changed || tasksChanged) {
          render();
          toast(changed ? "Alterações do JSON sincronizadas." : "O quiz da aula de hoje está pronto para revisão.");
        }
      });
    }, 350);
  });

  window.addEventListener("twenty:remote-state", function (event) {
    if (!event.detail || !event.detail.state || !state) return;
    state = normalizeState(event.detail.state);
    DB.saveState(state, { skipSync: true }).then(function () {
      if (Sync) Sync.adoptRemoteState(state);
      if (!onboarding) render();
      if (event.detail.conflicts && event.detail.conflicts.length) {
        toast("Foram detetadas alterações simultâneas. Nenhum registo foi apagado sem aviso.", "warning");
      } else if (event.detail.forced === "pull") {
        toast("Dados atualizados a partir do Git.");
      }
      // Atualizações automáticas não mostram popup nem toast.
    });
  });

  window.addEventListener("twenty:sync-status", function () {
    updateSyncActivityFromStatus(Sync ? Sync.getStatus() : null);
  });
  document.addEventListener("visibilitychange", function () {
    if (!document.hidden && Sync && state && Sync.getStatus().configured) Sync.checkForUpdates({ force: true }).catch(function () {});
  });

  init();
})();
