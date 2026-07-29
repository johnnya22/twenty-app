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
  var draggedNotebookSticker = null;
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
  var CANTEEN_AI_CACHE_KEY = "twenty-canteen-ai-v12-resilient-stream";
  var CANTEEN_WEATHER_CACHE_KEY = "twenty-canteen-weather-v1";
  var CANTEEN_WEATHER_URL = "https://api.open-meteo.com/v1/forecast?latitude=38.661150&longitude=-9.205777&current=temperature_2m,apparent_temperature,precipitation,rain,weather_code,cloud_cover&timezone=Europe%2FLisbon";
  var CANTEEN_DEFAULT_ALLERGENS = {
    "1": "Cereais com glúten",
    "2": "Crustáceos",
    "3": "Ovo",
    "4": "Peixe",
    "5": "Amendoim",
    "6": "Soja",
    "7": "Leite",
    "8": "Frutos de casca rija",
    "9": "Aipo",
    "10": "Mostarda",
    "11": "Sésamo",
    "12": "Sulfitos",
    "13": "Tremoço",
    "14": "Moluscos"
  };
  var canteenAIState = { key: "", status: "idle", availability: "unknown", source: "rules", data: null, error: "", progress: null, streamText: "" };
  var canteenAIPromise = null;
  var canteenAIAbortController = null;
  var canteenAIRequestKey = "";
  var canteenAISession = null;
  var canteenAISessionProviderKind = "";
  var canteenAISessionUses = 0;
  var canteenWeatherState = loadCachedCanteenWeather();
  var canteenWeatherPromise = null;
  var canteenMenu = loadCachedCanteen();
  var canteenStatus = canteenMenu ? "ready" : "idle";
  var canteenError = "";
  var canteenChecked = false;
  var canteenLoadPromise = null;
  var canteenSelectedDate = null;
  var canteenMealTab = null;
  var canteenSelections = {};
  var canteenExpandedCompleted = {};
  var canteenClockTimer = null;
  var homeClockTimer = null;
  var homeworkClockTimer = null;
  var homeworkSessionRuntime = null;
  var homeDebug = null;
  var COLORS = ["#a99df7", "#ff92ae", "#ffad72", "#79cdb8", "#80bee8", "#f3e873", "#cab6ea", "#87d7df"];
  var WEEKDAYS = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  var SHORT_WEEKDAYS = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  var ENTITY_ARRAYS = ["semesters", "courses", "schedule", "assessments", "events", "tasks", "lessons", "materials", "pastExams", "questions", "quizzes", "grades", "studyBlocks", "weeklyReviews", "aiProjects", "canteenVisits"];

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


  function normalizeContentBlocks(value) {
    if (value == null) return [];
    if (typeof value === "string" || typeof value === "number") {
      return [{ type: "text", text: String(value) }];
    }
    if (!Array.isArray(value)) {
      if (value && typeof value === "object" && value.type) return [value];
      return [];
    }
    return value.map(function (block) {
      if (typeof block === "string" || typeof block === "number") return { type: "text", text: String(block) };
      return block && typeof block === "object" ? block : null;
    }).filter(Boolean);
  }

  function contentBlocksPlainText(value) {
    return normalizeContentBlocks(value).map(function (block) {
      if (block.type === "text") return block.text || block.content || "";
      if (block.type === "latex") return block.latex || "";
      if (block.type === "code") return block.code || "";
      if (block.type === "slide-image") return block.description || block.alt || ("Imagem do slide " + (block.slideNumber || ""));
      if (block.type === "image") return block.alt || block.caption || "Imagem";
      if (block.type === "svg") return block.alt || "Diagrama SVG";
      return block.text || block.content || "";
    }).filter(Boolean).join("\n");
  }

  function sanitizeSvgMarkup(markup) {
    var source = String(markup || "").trim();
    if (!source || source.length > 60000 || !/^<svg[\s>]/i.test(source)) return "";
    try {
      var doc = new DOMParser().parseFromString(source, "image/svg+xml");
      if (doc.querySelector("parsererror") || !doc.documentElement || doc.documentElement.nodeName.toLowerCase() !== "svg") return "";
      var allowed = new Set(["svg", "g", "path", "rect", "circle", "ellipse", "line", "polyline", "polygon", "text", "tspan", "defs", "marker", "lineargradient", "radialgradient", "stop", "clippath"]);
      Array.from(doc.querySelectorAll("*")).forEach(function (node) {
        if (!allowed.has(node.nodeName.toLowerCase())) { node.remove(); return; }
        Array.from(node.attributes).forEach(function (attribute) {
          var name = attribute.name.toLowerCase();
          var value = attribute.value || "";
          var allowedAttr = ["viewbox", "width", "height", "x", "y", "x1", "x2", "y1", "y2", "cx", "cy", "r", "rx", "ry", "d", "points", "fill", "stroke", "stroke-width", "stroke-linecap", "stroke-linejoin", "opacity", "transform", "font-size", "font-weight", "text-anchor", "dominant-baseline", "offset", "stop-color", "stop-opacity", "marker-end", "marker-start", "id", "class", "xmlns"].indexOf(name) >= 0;
          if (!allowedAttr || /^on/i.test(name) || /javascript:|data:text\/html/i.test(value)) node.removeAttribute(attribute.name);
        });
      });
      var root = doc.documentElement;
      if (!root.getAttribute("viewBox") && !root.getAttribute("viewbox")) root.setAttribute("viewBox", "0 0 640 360");
      root.setAttribute("role", "img");
      root.setAttribute("focusable", "false");
      return new XMLSerializer().serializeToString(root);
    } catch (_) {
      return "";
    }
  }

  function renderContentBlock(block, options) {
    options = options || {};
    var type = String(block && block.type || "text").toLowerCase();
    if (type === "text") {
      var textValue = block.text != null ? block.text : block.content;
      return '<div class="content-block content-text">' + nl2br(textValue || "") + '</div>';
    }
    if (type === "latex") {
      var latex = String(block.latex || block.value || "").trim();
      if (!latex) return "";
      return '<div class="content-block math-content ' + (block.display === false ? "is-inline" : "is-display") + '">' + (block.display === false ? '\\(' + esc(latex) + '\\)' : '\\[' + esc(latex) + '\\]') + '</div>';
    }
    if (type === "code") {
      var language = String(block.language || "text").replace(/[^a-z0-9_+\-#.]/gi, "").slice(0, 24) || "text";
      return '<div class="content-block ide-code"><div class="ide-code-head"><span></span><span></span><span></span><b>' + esc(language) + '</b></div><pre><code>' + esc(block.code || "") + '</code></pre></div>';
    }
    if (type === "svg") {
      var svg = sanitizeSvgMarkup(block.svg || block.code || "");
      if (!svg) return '<div class="content-block content-warning"><i data-lucide="triangle-alert"></i>O SVG desta resposta não passou a validação.</div>';
      return '<figure class="content-block svg-content">' + svg + (block.alt || block.caption ? '<figcaption>' + esc(block.alt || block.caption) + '</figcaption>' : '') + '</figure>';
    }
    if (type === "slide-image") {
      var slideNumber = Number(block.slideNumber || block.slide || 0) || "?";
      var description = block.description || block.alt || "Usar a imagem relevante deste slide.";
      var materialButton = block.materialId ? '<button class="button button-small" type="button" data-action="open-material" data-id="' + attr(block.materialId) + '"><i data-lucide="presentation"></i>Abrir slides</button>' : '';
      return '<figure class="content-block slide-image-request"><div><span class="metric-icon"><i data-lucide="image"></i></span><p><strong>Imagem do slide ' + esc(slideNumber) + '</strong><small>' + esc(description) + '</small></p></div>' + materialButton + (block.caption ? '<figcaption>' + esc(block.caption) + '</figcaption>' : '') + '</figure>';
    }
    if (type === "image") {
      var src = safeResourceUrl(block.url || block.src || "");
      if (!src) return '<div class="content-block slide-image-request"><div><span class="metric-icon"><i data-lucide="image-off"></i></span><p><strong>Imagem pedida</strong><small>' + esc(block.description || block.alt || "A imagem ainda precisa de ser associada.") + '</small></p></div></div>';
      return '<figure class="content-block image-content"><img src="' + attr(src) + '" alt="' + attr(block.alt || block.caption || "Imagem") + '">' + (block.caption ? '<figcaption>' + esc(block.caption) + '</figcaption>' : '') + '</figure>';
    }
    if (type === "heading") {
      var level = Number(block.level) === 3 ? 3 : 2;
      return '<h' + level + ' class="content-block content-heading">' + esc(block.text || block.content || "") + '</h' + level + '>';
    }
    if (type === "list") {
      var ordered = !!block.ordered;
      return '<' + (ordered ? 'ol' : 'ul') + ' class="content-block content-list">' + asArray(block.items).map(function (item) { return '<li>' + renderContentBlocks(item, options) + '</li>'; }).join('') + '</' + (ordered ? 'ol' : 'ul') + '>';
    }
    return '<div class="content-block content-text">' + nl2br(block.text || block.content || "") + '</div>';
  }

  function renderContentBlocks(value, options) {
    return '<div class="rich-content">' + normalizeContentBlocks(value).map(function (block) { return renderContentBlock(block, options); }).join("") + '</div>';
  }

  function renderQuizOptionContent(option) {
    if (Array.isArray(option) || (option && typeof option === "object")) return renderContentBlocks(option);
    return '<div class="rich-content"><div class="content-block content-text">' + esc(option || "") + '</div></div>';
  }

  function blocksToNotebookHTML(value) {
    return normalizeContentBlocks(value).map(function (block) {
      var type = String(block.type || "text").toLowerCase();
      if (type === "text") return '<p>' + nl2br(block.text != null ? block.text : block.content || "") + '</p>';
      if (type === "heading") return '<h' + (Number(block.level) === 3 ? 3 : 2) + '>' + esc(block.text || block.content || "") + '</h' + (Number(block.level) === 3 ? 3 : 2) + '>';
      return '<div class="notebook-embed" contenteditable="false">' + renderContentBlock(block) + '</div>';
    }).join("");
  }

  function sanitizeNotebookHTML(value) {
    try {
      var doc = new DOMParser().parseFromString('<div id="noteRoot">' + String(value || "") + '</div>', "text/html");
      var root = doc.getElementById("noteRoot");
      var allowed = new Set(["P", "BR", "STRONG", "B", "EM", "I", "U", "S", "H2", "H3", "UL", "OL", "LI", "BLOCKQUOTE", "PRE", "CODE", "DIV", "SPAN", "A", "FIGURE", "FIGCAPTION", "SMALL", "IMG"]);
      Array.from(root.querySelectorAll("script,style,iframe,object,embed,form,input,button")).forEach(function (node) { node.remove(); });
      Array.from(root.querySelectorAll("*")).forEach(function (node) {
        if (node.nodeName.toLowerCase() === "svg") {
          var safeSvg = sanitizeSvgMarkup(new XMLSerializer().serializeToString(node));
          var holder = doc.createElement("div");
          holder.innerHTML = safeSvg || '<span>SVG removido</span>';
          node.replaceWith(holder.firstElementChild || holder.firstChild);
          return;
        }
        if (!allowed.has(node.nodeName)) {
          var fragment = doc.createDocumentFragment();
          while (node.firstChild) fragment.appendChild(node.firstChild);
          node.replaceWith(fragment);
          return;
        }
        Array.from(node.attributes).forEach(function (attribute) {
          var name = attribute.name.toLowerCase();
          var keep = name === "class" || name === "contenteditable" || name === "draggable" || name.indexOf("data-") === 0 || (node.nodeName === "A" && name === "href") || (node.nodeName === "IMG" && ["src", "alt", "width", "height"].indexOf(name) >= 0);
          if (!keep || /^on/.test(name)) node.removeAttribute(attribute.name);
        });
        if (node.nodeName === "A") {
          var href = safeResourceUrl(node.getAttribute("href") || "");
          if (href) { node.setAttribute("href", href); node.setAttribute("target", "_blank"); node.setAttribute("rel", "noopener"); }
          else node.removeAttribute("href");
        }
        if (node.nodeName === "IMG") {
          var remotePath = node.getAttribute("data-remote-path") || "";
          var localImageId = node.getAttribute("data-local-image-id") || "";
          var imageSrc = safeNotebookImageUrl(node.getAttribute("src") || "");
          if (imageSrc) node.setAttribute("src", imageSrc);
          else node.removeAttribute("src");
          if (!imageSrc && !remotePath && !localImageId) { node.remove(); return; }
          var imageWidth = clamp(node.getAttribute("data-width") || node.getAttribute("width") || 280, 80, 720);
          node.setAttribute("width", String(imageWidth));
          node.setAttribute("data-width", String(imageWidth));
          node.removeAttribute("height");
          node.setAttribute("alt", node.getAttribute("alt") || "Imagem dos apontamentos");
        }
      });
      return root.innerHTML;
    } catch (_) {
      return '<p>' + nl2br(String(value || "")) + '</p>';
    }
  }


  function notebookStickerMarkup(options) {
    options = options || {};
    var width = clamp(options.width || 280, 80, 720);
    var align = ["left", "center", "right"].indexOf(options.align) >= 0 ? options.align : "center";
    var imageAttributes = 'width="' + width + '" data-width="' + width + '" alt="' + attr(options.alt || options.name || "Imagem dos apontamentos") + '"';
    if (options.src) imageAttributes += ' src="' + attr(options.src) + '"';
    if (options.remote && options.remote.path) {
      imageAttributes += ' data-remote-path="' + attr(options.remote.path) + '" data-remote-name="' + attr(options.remote.name || options.name || "imagem") + '"';
    }
    if (options.localImageId) imageAttributes += ' data-local-image-id="' + attr(options.localImageId) + '"';
    return '<figure class="notebook-sticker align-' + align + '" contenteditable="false" draggable="true" data-width="' + width + '"><img ' + imageAttributes + '>' + (options.caption ? '<figcaption>' + esc(options.caption) + '</figcaption>' : '') + '</figure><p><br></p>';
  }

  function preparePastedNotebookHTML(value) {
    var safe = sanitizeNotebookHTML(value || "");
    var doc = new DOMParser().parseFromString('<div id="pasteRoot">' + safe + '</div>', "text/html");
    var root = doc.getElementById("pasteRoot");
    Array.from(root.querySelectorAll("img")).forEach(function (image) {
      var existing = image.closest(".notebook-sticker");
      var width = clamp(image.getAttribute("width") || image.getAttribute("data-width") || 280, 80, 720);
      image.setAttribute("width", String(width));
      image.setAttribute("data-width", String(width));
      image.removeAttribute("height");
      if (existing) {
        existing.setAttribute("contenteditable", "false");
        existing.setAttribute("draggable", "true");
        existing.setAttribute("data-width", String(width));
        return;
      }
      var figure = doc.createElement("figure");
      figure.className = "notebook-sticker align-center";
      figure.setAttribute("contenteditable", "false");
      figure.setAttribute("draggable", "true");
      figure.setAttribute("data-width", String(width));
      image.parentNode.insertBefore(figure, image);
      figure.appendChild(image);
    });
    return root.innerHTML;
  }

  function notebookEditorRange(editor) {
    var selection = window.getSelection();
    if (!selection || !selection.rangeCount) return null;
    var range = selection.getRangeAt(0);
    if (!editor.contains(range.commonAncestorContainer)) return null;
    return range.cloneRange();
  }

  function insertNotebookHTML(editor, html, savedRange) {
    if (!editor || !html) return;
    editor.focus();
    var selection = window.getSelection();
    var range = savedRange;
    if (!range || !editor.contains(range.commonAncestorContainer)) {
      range = document.createRange();
      range.selectNodeContents(editor);
      range.collapse(false);
    }
    selection.removeAllRanges();
    selection.addRange(range);
    if (!document.execCommand("insertHTML", false, html)) {
      var fragment = range.createContextualFragment(html);
      range.deleteContents();
      range.insertNode(fragment);
      range.collapse(false);
      selection.removeAllRanges();
      selection.addRange(range);
    }
    enhanceNotebookStickers(editor);
    hydrateLocalImages(editor);
    hydrateNotebookImages(editor);
  }

  function enhanceNotebookStickers(root) {
    if (!root) return;
    Array.from(root.querySelectorAll(".notebook-sticker")).forEach(function (sticker) {
      sticker.setAttribute("contenteditable", "false");
      sticker.setAttribute("draggable", "true");
      var image = sticker.querySelector("img");
      var width = clamp(sticker.getAttribute("data-width") || (image && (image.getAttribute("data-width") || image.getAttribute("width"))) || 280, 80, 720);
      sticker.setAttribute("data-width", String(width));
      if (image) { image.setAttribute("width", String(width)); image.setAttribute("data-width", String(width)); }
      if (!sticker.querySelector(".notebook-sticker-controls")) {
        sticker.insertAdjacentHTML("beforeend", '<span class="notebook-sticker-controls" contenteditable="false"><button type="button" data-action="notebook-sticker-smaller" title="Diminuir"><i data-lucide="minus"></i></button><button type="button" data-action="notebook-sticker-larger" title="Aumentar"><i data-lucide="plus"></i></button><button type="button" data-action="notebook-sticker-align" data-align="left" title="Alinhar à esquerda"><i data-lucide="align-left"></i></button><button type="button" data-action="notebook-sticker-align" data-align="center" title="Centrar"><i data-lucide="align-center"></i></button><button type="button" data-action="notebook-sticker-align" data-align="right" title="Alinhar à direita"><i data-lucide="align-right"></i></button><button type="button" class="is-danger" data-action="notebook-sticker-delete" title="Apagar imagem"><i data-lucide="trash-2"></i></button></span>');
      }
    });
    refreshIcons(root);
  }

  function hydrateNotebookImages(root) {
    if (!root || !Sync) return;
    Array.from(root.querySelectorAll('img[data-remote-path]')).forEach(function (image) {
      if (image.dataset.remoteHydrated === "loading" || image.dataset.remoteHydrated === "true") return;
      var path = image.getAttribute("data-remote-path");
      if (!path) return;
      image.dataset.remoteHydrated = "loading";
      Sync.downloadFile({ path: path, name: image.getAttribute("data-remote-name") || "imagem" }, {}).then(function (blob) {
        if (!blob || !image.isConnected) return;
        var url = URL.createObjectURL(blob);
        activeImageObjectUrls.push(url);
        image.src = url;
        image.dataset.remoteHydrated = "true";
      }).catch(function () {
        image.dataset.remoteHydrated = "error";
        image.alt = "Não foi possível descarregar esta imagem.";
      });
    });
  }

  async function storeNotebookImageFile(file) {
    if (!file || !/^image\//i.test(file.type || "")) throw new Error("Só podes colar ficheiros de imagem no caderno.");
    if (file.size > 12 * 1024 * 1024) throw new Error("Esta imagem tem mais de 12 MB.");
    var syncConfigured = !!(Sync && Sync.getStatus && Sync.getStatus().configured);
    var localImageId = "";
    var remoteFile = null;
    var objectUrl = URL.createObjectURL(file);
    activeImageObjectUrls.push(objectUrl);
    if (syncConfigured && navigator.onLine) {
      setManualSyncActivity("A enviar a imagem…", "A preparar a imagem para aparecer em todos os dispositivos.", 4, true);
      remoteFile = await Sync.uploadFile(file, {
        id: uid("noteimage"),
        name: file.name || "imagem-colada.png",
        onProgress: function (report) {
          var progress = report.progress == null ? null : 6 + Math.round(report.progress * 0.86);
          var detail = report.total ? formatBytes(report.loaded) + " de " + formatBytes(report.total) + " enviados" : "A enviar a imagem…";
          setManualSyncActivity("A enviar a imagem…", detail, progress, true);
        },
        onUploadComplete: function () { setManualSyncActivity("A confirmar no GitHub…", "A guardar a imagem no caderno sincronizado.", 94, true); }
      });
      finishManualSyncActivity(true);
    } else {
      localImageId = await DB.putFile(file, { kind: "notebook-sticker" });
      toast(syncConfigured ? "Sem Internet: a imagem ficou local até voltares a ter ligação." : "A imagem ficou apenas neste dispositivo. Configura o Git para a sincronizar.", "warning");
    }
    return { objectUrl: objectUrl, remoteFile: remoteFile, localImageId: localImageId, file: file };
  }

  async function uploadNotebookImage(file, editor, savedRange) {
    var stored = await storeNotebookImageFile(file);
    insertNotebookHTML(editor, notebookStickerMarkup({ src: stored.objectUrl, remote: stored.remoteFile, localImageId: stored.localImageId, name: file.name, alt: file.name || "Imagem colada", width: 280 }), savedRange);
  }

  async function dataUrlToNotebookFile(value, fallbackName) {
    var response = await fetch(value);
    var blob = await response.blob();
    var extension = (blob.type || "image/png").split("/")[1].replace("jpeg", "jpg") || "png";
    return new File([blob], (fallbackName || "imagem-colada") + "." + extension, { type: blob.type || "image/png" });
  }

  async function materializeNotebookImages(editor) {
    if (!editor) return;
    var images = Array.from(editor.querySelectorAll('img:not([data-remote-path]):not([data-local-image-id])'));
    var keptExternal = false;
    for (var index = 0; index < images.length; index += 1) {
      var image = images[index];
      var src = image.getAttribute("src") || "";
      if (!src) continue;
      var file = null;
      try {
        if (/^data:image\//i.test(src)) {
          file = await dataUrlToNotebookFile(src, "imagem-colada");
        } else if (/^blob:/i.test(src)) {
          var blobResponse = await fetch(src);
          var blob = await blobResponse.blob();
          file = new File([blob], "imagem-colada." + ((blob.type || "image/png").split("/")[1] || "png"), { type: blob.type || "image/png" });
        } else if (/^https?:/i.test(src)) {
          try {
            var remoteResponse = await fetch(src, { mode: "cors", credentials: "omit" });
            if (remoteResponse.ok && /^image\//i.test(remoteResponse.headers.get("content-type") || "")) {
              var remoteBlob = await remoteResponse.blob();
              file = new File([remoteBlob], "imagem-web." + ((remoteBlob.type || "image/png").split("/")[1] || "png"), { type: remoteBlob.type || "image/png" });
            }
          } catch (_) {
            keptExternal = true;
          }
        }
        if (!file) continue;
        var stored = await storeNotebookImageFile(file);
        if (stored.remoteFile && stored.remoteFile.path) {
          image.setAttribute("data-remote-path", stored.remoteFile.path);
          image.setAttribute("data-remote-name", stored.remoteFile.name || file.name);
          image.removeAttribute("data-local-image-id");
        } else if (stored.localImageId) {
          image.setAttribute("data-local-image-id", stored.localImageId);
          image.removeAttribute("data-remote-path");
          image.removeAttribute("data-remote-name");
        }
        image.src = stored.objectUrl;
      } catch (error) {
        console.warn("Não foi possível guardar uma imagem colada", error);
        keptExternal = true;
      }
    }
    if (keptExternal) toast("Algumas imagens ficaram ligadas ao site original porque esse site bloqueou a cópia direta.", "warning");
  }

  async function handleNotebookPaste(event, editor) {
    var clipboard = event.clipboardData || window.clipboardData;
    if (!clipboard) return;
    event.preventDefault();
    var savedRange = notebookEditorRange(editor);
    var imageFiles = Array.from(clipboard.items || []).filter(function (item) { return item.kind === "file" && /^image\//i.test(item.type || ""); }).map(function (item) { return item.getAsFile(); }).filter(Boolean);
    if (!imageFiles.length) imageFiles = Array.from(clipboard.files || []).filter(function (file) { return /^image\//i.test(file.type || ""); });
    if (imageFiles.length) {
      try {
        for (var index = 0; index < imageFiles.length; index += 1) await uploadNotebookImage(imageFiles[index], editor, savedRange);
      } catch (error) {
        finishManualSyncActivity(false);
        toast(error.message || "Não foi possível colar a imagem.", "error");
      }
      return;
    }
    var html = clipboard.getData("text/html");
    if (html) {
      insertNotebookHTML(editor, preparePastedNotebookHTML(html), savedRange);
      return;
    }
    var text = clipboard.getData("text/plain");
    if (text) insertNotebookHTML(editor, '<p>' + nl2br(text) + '</p>', savedRange);
  }

  function typesetMath(root) {
    if (!root || !root.querySelector(".math-content")) return;
    if (window.MathJax && typeof window.MathJax.typesetPromise === "function") {
      setTimeout(function () { window.MathJax.typesetPromise([root]).catch(function () {}); }, 0);
    }
  }

  function parseJSONReply(value) {
    var text = String(value || "").trim();
    if (!text) throw new Error("Cola primeiro a resposta JSON da IA.");
    var fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fenced) text = fenced[1].trim();
    var first = text.indexOf("{");
    var last = text.lastIndexOf("}");
    if (first >= 0 && last > first) text = text.slice(first, last + 1);
    try { return JSON.parse(text); } catch (_) { throw new Error("A resposta não é JSON válido. Pede à IA para devolver apenas o objeto JSON."); }
  }

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
      schemaVersion: 7,
      meta: {
        revision: 0,
        updatedAt: "",
        externalFingerprint: "",
        externalCheckedAt: "",
        externalRevision: 0,
        source: "device",
        completedLessonQuizIds: [],
        completedHomeworkIds: []
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
        aiDifficulty: "auto",
        canteenAIEnabled: true,
        canteenAIDescriptions: true,
        canteenAIChefNote: true,
        canteenTheme: "diner",
        canteenAllergenFilters: {
          selected: [],
          hideDishes: true
        },
        canteenFoodPreferences: {
          dislikes: ["ervilhas"],
          lowPreference: ["peixe da cantina"],
          goal: "ganhar peso"
        }
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
      aiProjects: [],
      canteenVisits: []
    };
  }

  function normalizeState(input) {
    var base = defaultState();
    var source = input && typeof input === "object" ? input : {};
    base.schemaVersion = Math.max(7, Number(source.schemaVersion) || 0);
    base.meta = Object.assign(base.meta, source.meta || {});
    base.meta.completedLessonQuizIds = Array.from(new Set(asArray(base.meta.completedLessonQuizIds).map(String).filter(Boolean)));
    base.meta.completedHomeworkIds = Array.from(new Set(asArray(base.meta.completedHomeworkIds).map(String).filter(Boolean)));
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
    base.settings.canteenAIEnabled = base.settings.canteenAIEnabled !== false;
    base.settings.canteenAIDescriptions = base.settings.canteenAIDescriptions !== false;
    base.settings.canteenAIChefNote = base.settings.canteenAIChefNote !== false;
    if (["diner", "leaf"].indexOf(base.settings.canteenTheme) < 0) base.settings.canteenTheme = "diner";
    base.settings.canteenAllergenFilters = Object.assign({ selected: [], hideDishes: true }, base.settings.canteenAllergenFilters || {});
    base.settings.canteenAllergenFilters.selected = Array.from(new Set(asArray(base.settings.canteenAllergenFilters.selected).map(String).filter(function (id) { return Object.prototype.hasOwnProperty.call(CANTEEN_DEFAULT_ALLERGENS, id); })));
    base.settings.canteenAllergenFilters.hideDishes = base.settings.canteenAllergenFilters.hideDishes !== false;
    base.settings.canteenFoodPreferences = Object.assign({ dislikes: ["ervilhas"], lowPreference: ["peixe da cantina"], goal: "ganhar peso" }, base.settings.canteenFoodPreferences || {});
    base.settings.canteenFoodPreferences.dislikes = asArray(base.settings.canteenFoodPreferences.dislikes).map(cleanText).filter(Boolean);
    base.settings.canteenFoodPreferences.lowPreference = asArray(base.settings.canteenFoodPreferences.lowPreference).map(cleanText).filter(Boolean);
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
      var normalized = Object.assign({ notes: "", notesHtml: "", notesPaper: "lined", notesFont: "app", aiNotes: [], mastered: false }, lesson, { aiNotes: asArray(lesson.aiNotes) });
      if (!normalized.notesHtml && normalized.notes) normalized.notesHtml = '<p>' + nl2br(normalized.notes) + '</p>';
      if (["lined", "grid", "blank"].indexOf(normalized.notesPaper) < 0) normalized.notesPaper = "lined";
      normalized.notesHtml = sanitizeNotebookHTML(normalized.notesHtml || "");
      return normalized;
    });
    base.materials = base.materials.map(function (material) {
      return Object.assign({ remoteFile: null, slides: [], slideCount: 0, uploadStatus: "ready" }, material, {
        slides: asArray(material.slides),
        remoteFile: material.remoteFile && typeof material.remoteFile === "object" ? material.remoteFile : null
      });
    });
    base.canteenVisits = base.canteenVisits.map(function (visit) {
      var normalizedVisit = Object.assign({
        id: uid("canteen"), date: "", mealType: "lunch", mealLabel: "Almoço",
        ticketIssuedAt: "", completedAt: "", orderNumber: "",
        price: "3,10 €", totalKcal: 0, dish: null, soup: null, dessert: null
      }, visit, {
        dish: visit.dish && typeof visit.dish === "object" ? visit.dish : null,
        soup: visit.soup && typeof visit.soup === "object" ? visit.soup : null,
        dessert: visit.dessert && typeof visit.dessert === "object" ? visit.dessert : null
      });
      if (!normalizedVisit.ticketIssuedAt && normalizedVisit.completedAt) normalizedVisit.ticketIssuedAt = normalizedVisit.completedAt;
      if (!normalizedVisit.orderNumber && normalizedVisit.ticketIssuedAt) {
        normalizedVisit.orderNumber = String((parseInt(hashText(normalizedVisit.id + normalizedVisit.ticketIssuedAt).slice(-6), 16) % 900) + 100);
      }
      return normalizedVisit;
    });
    base.tasks = base.tasks.map(function (task) {
      if (task.type === "lesson-quiz" && /^Quiz beOnLine · /.test(task.title || "")) task.title = String(task.title).replace(/^Quiz beOnLine · /, "Quiz da aula · ");
      task.contentBlocks = normalizeContentBlocks(task.contentBlocks || task.instructionsBlocks || []);
      task.solutionBlocks = normalizeContentBlocks(task.solutionBlocks || []);
      task.checklist = asArray(task.checklist).map(String).filter(Boolean);
      task.extraTasks = asArray(task.extraTasks).map(function (item, index) {
        if (typeof item === "string") return { title: item, blocks: [], optional: false };
        item = item && typeof item === "object" ? item : {};
        return {
          title: String(item.title || item.label || ("Tarefa extra " + (index + 1))),
          blocks: normalizeContentBlocks(item.blocks || item.contentBlocks || item.instructions || ""),
          optional: !!item.optional
        };
      });
      task.completedAt = task.completedAt || "";
      task.completedOnce = !!(task.completedOnce || task.done || task.completedAt || base.meta.completedHomeworkIds.indexOf(String(task.id)) >= 0);
      if (task.completedOnce && (task.type === "homework" || task.type === "tpc")) task.done = true;
      if ((task.type === "homework" || task.type === "tpc") && task.completedOnce && base.meta.completedHomeworkIds.indexOf(String(task.id)) < 0) base.meta.completedHomeworkIds.push(String(task.id));
      task.actualSeconds = Math.max(0, Number(task.actualSeconds) || 0);
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
      quiz.questions = asArray(quiz.questions).map(function (question) {
        var normalized = Object.assign({}, question, { images: normalizeImageRefs(question.images) });
        normalized.promptBlocks = normalizeContentBlocks(normalized.promptBlocks || normalized.prompt || "");
        normalized.explanationBlocks = normalizeContentBlocks(normalized.explanationBlocks || normalized.explanation || normalized.answer || "");
        normalized.optionBlocks = asArray(normalized.optionBlocks).length ? normalized.optionBlocks.map(normalizeContentBlocks) : asArray(normalized.options).map(function (option) { return normalizeContentBlocks(option); });
        if (normalized.answerIndex == null && normalized.correctIndex != null) normalized.answerIndex = normalized.correctIndex;
        normalized.options = normalized.optionBlocks.map(contentBlocksPlainText);
        normalized.prompt = contentBlocksPlainText(normalized.promptBlocks);
        normalized.explanation = contentBlocksPlainText(normalized.explanationBlocks);
        return normalized;
      });
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
    base.lessons.forEach(function (lesson) {
      var completedQuiz = base.quizzes.filter(function (quiz) { return quiz.lessonId === lesson.id; }).sort(function (a, b) {
        return String(b.lastCompletedAt || "").localeCompare(String(a.lastCompletedAt || ""));
      })[0];
      var completedTask = base.tasks.find(function (task) { return task.lessonId === lesson.id && task.type === "lesson-quiz" && task.done; });
      var quizHasResult = completedQuiz && (completedQuiz.completedOnce || completedQuiz.lastScore != null || completedQuiz.lastCompletedAt);
      var completedAt = lesson.quizCompletedAt || lesson.beOnlineCompletedAt || (lesson.quizCompleted ? lesson.updatedAt || lesson.date || new Date().toISOString() : "") || (quizHasResult ? completedQuiz.lastCompletedAt || completedQuiz.updatedAt || completedQuiz.createdAt || new Date().toISOString() : "") || (completedTask && (completedTask.completedAt || completedTask.updatedAt || completedTask.createdAt || new Date().toISOString()));
      if (!completedAt) return;
      lesson.quizCompleted = true;
      if (base.meta.completedLessonQuizIds.indexOf(String(lesson.id)) < 0) base.meta.completedLessonQuizIds.push(String(lesson.id));
      lesson.quizCompletedAt = completedAt;
      lesson.beOnlineCompletedAt = lesson.beOnlineCompletedAt || completedAt;
      base.quizzes.forEach(function (quiz) {
        if (quiz.lessonId === lesson.id && asArray(quiz.questions).length) {
          quiz.completedOnce = true;
          quiz.lastCompletedAt = quiz.lastCompletedAt || completedAt;
        }
      });
      base.tasks.forEach(function (task) {
        if (task.lessonId === lesson.id && task.type === "lesson-quiz") {
          task.done = true;
          task.completedOnce = true;
          task.completedAt = task.completedAt || completedAt;
        }
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
    if (homeDebug && homeDebug.active) {
      if (!silent) toast("Alteração aplicada apenas à simulação.");
      return Promise.resolve();
    }
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
    if (homeDebug && homeDebug.active) return Promise.resolve(false);
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

  function canteenAllergenSettings() {
    var raw = state && state.settings && state.settings.canteenAllergenFilters || {};
    return {
      selected: Array.from(new Set(asArray(raw.selected).map(String).filter(function (id) { return Object.prototype.hasOwnProperty.call(CANTEEN_DEFAULT_ALLERGENS, id); }))),
      hideDishes: raw.hideDishes !== false
    };
  }

  function canteenAllergenMap() {
    return Object.assign({}, CANTEEN_DEFAULT_ALLERGENS, canteenMenu && canteenMenu.allergens || {});
  }

  function canteenItemMatchesSelectedAllergen(item) {
    var selected = canteenAllergenSettings().selected;
    if (!selected.length) return false;
    return asArray(item && item.allergens).map(String).some(function (id) { return selected.indexOf(id) >= 0; });
  }

  function canteenItemVisible(item) {
    var settings = canteenAllergenSettings();
    return !(settings.hideDishes && settings.selected.length && canteenItemMatchesSelectedAllergen(item));
  }

  function canteenAllergenFilterSignature() {
    var settings = canteenAllergenSettings();
    return (settings.hideDishes ? "hide" : "show") + ":" + settings.selected.slice().sort(function (a, b) { return Number(a) - Number(b); }).join(",");
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
    if (route.name === "canteen" || route.name === "home") render();
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
    document.body.dataset.canteenTheme = state.settings.canteenTheme === "leaf" ? "leaf" : "diner";
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

  function refreshCanteenClockView() {
    if (route.name !== "canteen" || !view || !canteenMenu) return;
    var stamp = view.querySelector(".campus-dining-stamp");
    if (!stamp) return;
    var now = canteenPortugalParts(new Date());
    var selectedIsToday = canteenSelectedDate === now.iso;
    var service = canteenOpeningStatus(new Date(), canteenServiceHours(canteenMenu));
    var nextLabel = selectedIsToday ? (service.open ? "ABERTO" : "FECHADO") : "EMENTA";
    var labelTarget = stamp.querySelector("[data-canteen-status-label]");
    if (labelTarget) labelTarget.textContent = nextLabel;
    else stamp.textContent = nextLabel;
    stamp.classList.toggle("is-open", !!(selectedIsToday && service.open));
  }

  function scheduleCanteenClockRefresh() {
    if (canteenClockTimer) clearTimeout(canteenClockTimer);
    canteenClockTimer = setTimeout(function () {
      canteenClockTimer = null;
      if (route.name !== "canteen") return;
      refreshCanteenClockView();
      scheduleCanteenClockRefresh();
    }, 60000 - (Date.now() % 60000) + 250);
  }

  function render() {
    if (!state) return;
    revokeImageObjectUrls();
    if (canteenClockTimer) {
      clearTimeout(canteenClockTimer);
      canteenClockTimer = null;
    }
    if (homeClockTimer) {
      clearTimeout(homeClockTimer);
      homeClockTimer = null;
    }
    if (homeworkClockTimer) {
      clearTimeout(homeworkClockTimer);
      homeworkClockTimer = null;
    }
    renderShell();
    document.body.dataset.route = route.name || "home";
    var html;
    if (route.name === "home") html = renderHome();
    else if (route.name === "courses") html = renderCourses();
    else if (route.name === "course") html = renderCourse(route.id, route.tab);
    else if (route.name === "lesson") html = renderLesson(route.id);
    else if (route.name === "planner") html = renderPlanner();
    else if (route.name === "homework") html = renderHomeworkSession();
    else if (route.name === "study") html = renderStudy();
    else if (route.name === "grades") html = renderGrades();
    else if (route.name === "canteen") html = renderCanteen();
    else if (route.name === "settings") html = renderSettings();
    else { route.name = "home"; html = renderHome(); }
    view.innerHTML = '<div class="view-enter">' + html + "</div>";
    view.focus({ preventScroll: true });
    refreshIcons(document);
    hydrateLocalImages(view);
    hydrateNotebookImages(view);
    typesetMath(view);
    if (route.name === "settings") {
      enhanceSettingsActions();
      updateStorageCount();
    }
    if ((route.name === "canteen" || route.name === "home") && !canteenChecked && canteenStatus !== "loading") {
      setTimeout(function () { ensureCanteenMenu(false); }, 0);
    }
    if (route.name === "canteen") {
      setTimeout(function () { ensureCanteenAIForCurrentMeal(false); }, 0);
      scheduleCanteenClockRefresh();
    }
    if (route.name === "home") {
      homeClockTimer = setTimeout(function () {
        if (route.name === "home" && !modalRoot.querySelector(".modal-card")) render();
      }, 60000 - (Date.now() % 60000) + 250);
    }
    if (route.name === "homework") scheduleHomeworkClock();
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
    var now = date || activeHomeNow();
    var today = todayISO(now);
    if (lesson.date < today) return true;
    if (lesson.date > today) return false;
    var end = lesson.end || lesson.start;
    return end ? timeMinutes(end) <= nowMinutes(now) : true;
  }


  function configuredQuizForLesson(lessonId) {
    return state.quizzes.find(function (quiz) {
      return quiz.lessonId === lessonId && asArray(quiz.questions).length;
    }) || null;
  }

  function homeworkForLesson(lessonId) {
    return state.tasks.find(function (task) {
      return task.lessonId === lessonId && (task.type === "homework" || task.type === "tpc") && (task.lockedContent || asArray(task.contentBlocks).length || task.configuredFromPrompt);
    }) || null;
  }

  function lessonIsBeOnline(lesson) {
    if (!lesson) return false;
    var quizzes = state.quizzes.filter(function (quiz) { return quiz.lessonId === lesson.id && asArray(quiz.questions).length; });
    if (!quizzes.length) return false;
    if (asArray(state.meta && state.meta.completedLessonQuizIds).indexOf(String(lesson.id)) >= 0) return true;
    if (lesson.quizCompleted || lesson.quizCompletedAt || lesson.beOnlineCompletedAt) return true;
    if (quizzes.some(function (quiz) { return !!quiz.completedOnce || quiz.lastScore != null || !!quiz.lastCompletedAt; })) return true;
    return state.tasks.some(function (task) { return task.lessonId === lesson.id && task.type === "lesson-quiz" && (task.done || task.completedOnce); });
  }

  function beOnlineStatus() {
    var due = semesterItems("lessons").filter(function (lesson) {
      return lessonHasEnded(lesson) && !!configuredQuizForLesson(lesson.id);
    });
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
    var before = state.tasks.length;
    var seenLessonQuizTasks = new Set();
    state.tasks = state.tasks.filter(function (task) {
      if (task.type !== "lesson-quiz" || !task.autoGenerated) return true;
      if (!configuredQuizForLesson(task.lessonId)) return false;
      if (seenLessonQuizTasks.has(task.lessonId)) return false;
      seenLessonQuizTasks.add(task.lessonId);
      return true;
    });
    if (state.tasks.length !== before) changed = true;
    semesterItems("lessons").filter(function (lesson) {
      return lessonHasEnded(lesson) && !!configuredQuizForLesson(lesson.id);
    }).forEach(function (lesson) {
      var quiz = configuredQuizForLesson(lesson.id);
      var task = state.tasks.find(function (item) { return item.type === "lesson-quiz" && item.lessonId === lesson.id; });
      if (lessonIsBeOnline(lesson)) {
        if (task && !task.done) { task.done = true; changed = true; }
        return;
      }
      if (!task) {
        var course = courseById(lesson.courseId);
        state.tasks.push({
          id: uid("task"), semesterId: lesson.semesterId, courseId: lesson.courseId, lessonId: lesson.id,
          quizId: quiz && quiz.id, title: "Quiz da aula · " + lesson.title, type: "lesson-quiz", dueDate: lesson.date,
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
    var completedAt = new Date().toISOString();
    lesson.quizCompleted = true;
    state.meta.completedLessonQuizIds = Array.from(new Set(asArray(state.meta.completedLessonQuizIds).concat(String(lessonId))));
    lesson.quizCompletedAt = completedAt;
    lesson.beOnlineCompletedAt = completedAt;
    state.quizzes.forEach(function (quiz) {
      if (quiz.lessonId === lessonId && asArray(quiz.questions).length) {
        quiz.completedOnce = true;
        quiz.lastCompletedAt = completedAt;
      }
    });
    state.tasks.forEach(function (task) {
      if (task.lessonId === lessonId && task.type === "lesson-quiz") {
        task.done = true;
        task.completedOnce = true;
        task.completedAt = completedAt;
      }
    });
  }

  function emptyState(icon, title, text, action, label) {
    return '<div class="empty-state"><span class="empty-icon"><i data-lucide="' + icon + '"></i></span><h3>' + esc(title) + "</h3><p>" + esc(text) + "</p>" + (action ? '<button class="button button-dark button-small" type="button" data-action="' + attr(action) + '"><i data-lucide="plus"></i>' + esc(label || "Adicionar") + "</button>" : "") + "</div>";
  }

  function activeHomeNow() {
    if (homeDebug && homeDebug.active && homeDebug.now) {
      var simulated = new Date(homeDebug.now);
      if (!Number.isNaN(simulated.getTime())) return simulated;
    }
    return new Date();
  }

  function homeDebugScenarios(baseDate) {
    var base = new Date(baseDate || new Date());
    function at(hour, minute) {
      var value = new Date(base);
      value.setHours(hour, minute, 0, 0);
      return value.toISOString();
    }
    return [
      { id: "morning", time: at(8, 10), label: "Manhã", title: "Antes das aulas", copy: "A Home apresenta a primeira aula, a sala e o que convém preparar." },
      { id: "soon", time: at(8, 52), label: "08:52", title: "Quase a começar", copy: "Faltam poucos minutos. A prioridade é abrir a aula e deixar os materiais prontos." },
      { id: "live", time: at(9, 35), label: "09:35", title: "Aula a decorrer", copy: "A Home centra-se na aula atual sem encher o ecrã de informação secundária." },
      { id: "closing", time: at(10, 22), label: "10:22", title: "Aula a terminar", copy: "Mesmo com outra aula colada, a Home fecha primeiro a aula atual e avisa que o Quiz da aula pode continuar nos primeiros 10 minutos seguintes." },
      { id: "settling", time: at(10, 35), label: "10:35", title: "Primeiros 10 minutos", copy: "Enquanto os alunos chegam, a Home usa esta janela para o quiz da aula anterior." },
      { id: "between", time: at(12, 12), label: "12:12", title: "Entre aulas", copy: "A Home fecha o que ficou pendente e mostra claramente quando começa a próxima aula." },
      { id: "after", time: at(15, 40), label: "15:40", title: "Depois das aulas", copy: "Primeiro aparecem os Quizzes da aula. Só depois surgem os TPCs." },
      { id: "homework", time: at(18, 10), label: "18:10", title: "TPC em casa", copy: "Com os Quizzes da aula concluídos, a Home transforma-se numa fila simples de TPCs." },
      { id: "complete", time: at(21, 5), label: "21:05", title: "Dia fechado", copy: "Com tudo concluído, aparece o Report Card e a sensação de dia escolar terminado." }
    ];
  }

  function buildHomeTutorialState(baseDate) {
    var now = new Date(baseDate || new Date());
    var day = todayISO(now);
    var tomorrow = new Date(now);
    tomorrow.setDate(tomorrow.getDate() + 1);
    var semesterId = "demo_semester";
    var courses = [
      { id: "demo_programacao", semesterId: semesterId, name: "Programação", code: "PROG", ects: 6, color: "#a99df7", lessonTypes: ["T"], evaluation: { components: [], examReplacesTests: false, replacementPolicy: "if-higher" } },
      { id: "demo_design", semesterId: semesterId, name: "Design de Interfaces", code: "DI", ects: 6, color: "#ffad72", lessonTypes: ["TP"], evaluation: { components: [], examReplacesTests: false, replacementPolicy: "if-higher" } },
      { id: "demo_ia", semesterId: semesterId, name: "Inteligência Artificial", code: "IA", ects: 6, color: "#79cdb8", lessonTypes: ["T"], evaluation: { components: [], examReplacesTests: false, replacementPolicy: "if-higher" } }
    ];
    var schedule = [
      { id: "demo_schedule_1", semesterId: semesterId, courseId: "demo_programacao", weekday: now.getDay(), start: "09:00", end: "10:30", room: "A1.12", type: "T" },
      { id: "demo_schedule_2", semesterId: semesterId, courseId: "demo_design", weekday: now.getDay(), start: "10:30", end: "12:00", room: "Lab 2", type: "TP" },
      { id: "demo_schedule_3", semesterId: semesterId, courseId: "demo_ia", weekday: now.getDay(), start: "14:00", end: "15:30", room: "B2.08", type: "T" }
    ];
    var lessons = [
      { id: "demo_lesson_1", semesterId: semesterId, courseId: "demo_programacao", scheduleId: "demo_schedule_1", title: "Arrays e ciclos", date: day, start: "09:00", end: "10:30", room: "A1.12", type: "T", notes: "", aiNotes: [], mastered: false },
      { id: "demo_lesson_2", semesterId: semesterId, courseId: "demo_design", scheduleId: "demo_schedule_2", title: "Hierarquia visual", date: day, start: "10:30", end: "12:00", room: "Lab 2", type: "TP", notes: "", aiNotes: [], mastered: false },
      { id: "demo_lesson_3", semesterId: semesterId, courseId: "demo_ia", scheduleId: "demo_schedule_3", title: "Pesquisa heurística", date: day, start: "14:00", end: "15:30", room: "B2.08", type: "T", notes: "", aiNotes: [], mastered: false }
    ];
    function quiz(id, lessonId, courseId, title) {
      return { id: id, semesterId: semesterId, courseId: courseId, lessonId: lessonId, title: title, questions: [
        { id: id + "_q1", mode: "multiple-choice", prompt: "Qual foi a ideia principal desta aula?", options: ["Aplicar o conceito central", "Ignorar os exemplos", "Memorizar sem perceber", "Saltar a revisão"], answerIndex: 0, explanation: "O Quiz da aula confirma a compreensão do conceito principal." },
        { id: id + "_q2", mode: "multiple-choice", prompt: "Qual é o melhor próximo passo quando algo não ficou claro?", options: ["Anotar a dúvida", "Esperar pelo exame", "Apagar os apontamentos", "Fingir que ficou claro"], answerIndex: 0, explanation: "Detetar a dúvida cedo evita matéria acumulada." }
      ] };
    }
    var demo = defaultState();
    demo.profile = { name: "Johnny", institution: "Twenty Campus", degree: "Licenciatura", targetGrade: 18, onboardingComplete: true, tutorialSeen: true };
    demo.settings = Object.assign(demo.settings, { jsonSync: false });
    demo.currentSemesterId = semesterId;
    demo.semesters = [{ id: semesterId, name: "Semestre de demonstração", academicYear: academicYearFor(now), startDate: day, endDate: "", archived: false }];
    demo.courses = courses;
    demo.schedule = schedule;
    demo.lessons = lessons;
    demo.quizzes = [
      quiz("demo_quiz_1", "demo_lesson_1", "demo_programacao", "Quiz da aula · Arrays e ciclos"),
      quiz("demo_quiz_2", "demo_lesson_2", "demo_design", "Quiz da aula · Hierarquia visual"),
      quiz("demo_quiz_3", "demo_lesson_3", "demo_ia", "Quiz da aula · Pesquisa heurística")
    ];
    demo.tasks = [
      { id: "demo_check_1", semesterId: semesterId, courseId: "demo_programacao", lessonId: "demo_lesson_1", title: "Quiz da aula · Arrays e ciclos", type: "lesson-quiz", dueDate: day, dueTime: "20:30", priority: "high", done: false, autoGenerated: true },
      { id: "demo_check_2", semesterId: semesterId, courseId: "demo_design", lessonId: "demo_lesson_2", title: "Quiz da aula · Hierarquia visual", type: "lesson-quiz", dueDate: day, dueTime: "20:30", priority: "high", done: false, autoGenerated: true },
      { id: "demo_check_3", semesterId: semesterId, courseId: "demo_ia", lessonId: "demo_lesson_3", title: "Quiz da aula · Pesquisa heurística", type: "lesson-quiz", dueDate: day, dueTime: "20:30", priority: "high", done: false, autoGenerated: true },
      { id: "demo_tpc_1", semesterId: semesterId, courseId: "demo_programacao", lessonId: "demo_lesson_1", title: "Resolver exercícios 4–8", type: "homework", dueDate: day, dueTime: "19:00", priority: "high", done: false },
      { id: "demo_tpc_2", semesterId: semesterId, courseId: "demo_design", lessonId: "demo_lesson_2", title: "Refazer o ecrã com melhor hierarquia", type: "homework", dueDate: day, dueTime: "20:00", priority: "normal", done: false }
    ];
    demo.assessments = [{ id: "demo_assessment", semesterId: semesterId, courseId: "demo_ia", title: "Mini-teste de pesquisa", type: "Teste", date: todayISO(tomorrow), time: "10:00", location: "B2.08", lessonIds: ["demo_lesson_3"] }];
    demo.meta.updatedAt = "";
    demo.meta.revision = 0;
    return normalizeState(demo);
  }

  function prepareHomeDebugScenario(index) {
    if (!homeDebug || !homeDebug.active) return;
    var scenarios = homeDebugScenarios(homeDebug.baseDate);
    var nextIndex = clamp(index, 0, scenarios.length - 1);
    var scenario = scenarios[nextIndex];
    state = normalizeState(clone(homeDebug.templateState));
    if (scenario.id === "homework" || scenario.id === "complete") {
      state.lessons.forEach(function (lesson, lessonIndex) {
        lesson.quizCompletedAt = scenario.id === "complete" || lessonIndex < 3 ? scenario.time : "";
      });
      state.tasks.forEach(function (task) {
        if (task.type === "lesson-quiz") task.done = true;
        if (scenario.id === "complete" && task.type !== "lesson-quiz") task.done = true;
      });
      state.quizzes.forEach(function (quiz, quizIndex) {
        quiz.lastCompletedAt = scenario.time;
        quiz.lastScore = [90, 80, 85][quizIndex] || 85;
      });
    }
    homeDebug.index = nextIndex;
    homeDebug.now = scenario.time;
    homeDebug.scenario = scenario.id;
    route = { name: "home", id: null, tab: "overview" };
    render();
    history.replaceState(null, "", "#home");
  }

  function startHomeDebugTutorial() {
    closeModal();
    var baseDate = new Date();
    homeDebug = {
      active: true,
      tutorial: true,
      originalState: clone(state),
      templateState: buildHomeTutorialState(baseDate),
      baseDate: baseDate.toISOString(),
      index: 0,
      now: "",
      scenario: "morning"
    };
    prepareHomeDebugScenario(0);
  }

  function startHomeDebugAt(value) {
    var simulated = new Date(value);
    if (Number.isNaN(simulated.getTime())) {
      toast("Escolhe uma data e hora válidas.", "warning");
      return;
    }
    if (!homeDebug || !homeDebug.active) {
      homeDebug = { active: true, tutorial: false, originalState: clone(state), templateState: clone(state), baseDate: simulated.toISOString(), index: 0, now: simulated.toISOString(), scenario: "custom" };
    } else {
      homeDebug.tutorial = false;
      homeDebug.templateState = clone(state);
      homeDebug.now = simulated.toISOString();
      homeDebug.scenario = "custom";
    }
    closeModal();
    setRoute("home");
  }

  function stopHomeDebug() {
    if (!homeDebug || !homeDebug.active) return;
    var original = homeDebug.originalState;
    homeDebug = null;
    state = normalizeState(original || state);
    closeModal();
    route = { name: "settings", id: null, tab: "overview" };
    render();
    history.replaceState(null, "", "#settings");
    toast("Simulação terminada. Os teus dados reais não foram alterados.");
  }

  function homeDebugSnapshot() {
    var context = homeContext(activeHomeNow());
    return {
      time: activeHomeNow().toLocaleString("pt-PT", { weekday: "short", hour: "2-digit", minute: "2-digit" }),
      phase: context.phase,
      label: context.phaseLabel,
      current: context.current && (context.current.course ? context.current.course.name : context.current.type),
      next: context.next && (context.next.course ? context.next.course.name : context.next.type),
      pendingChecks: context.pendingChecks.length,
      pendingHomework: context.tasks.homeworkPending.length + context.tasks.overdue.length
    };
  }

  function openHomeDebugLab() {
    var snapshot = homeDebugSnapshot();
    var now = activeHomeNow();
    var localValue = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0") + "-" + String(now.getDate()).padStart(2, "0") + "T" + String(now.getHours()).padStart(2, "0") + ":" + String(now.getMinutes()).padStart(2, "0");
    var scenarios = homeDebugScenarios(now).map(function (scenario, index) {
      return '<button class="debug-scenario-button" type="button" data-action="debug-start-scenario" data-index="' + index + '"><strong>' + esc(scenario.title) + '</strong><small>' + esc(scenario.label) + ' · ' + esc(scenario.copy) + '</small></button>';
    }).join("");
    var body = '<div class="debug-summary-grid"><span><small>Hora usada</small><strong>' + esc(snapshot.time) + '</strong></span><span><small>Estado</small><strong>' + esc(snapshot.label) + '</strong></span><span><small>Quizzes da aula</small><strong>' + snapshot.pendingChecks + ' pendente(s)</strong></span><span><small>TPC</small><strong>' + snapshot.pendingHomework + ' pendente(s)</strong></span></div><div class="form-note">Este laboratório é apenas visual. Não faz commits, não mexe no Git e restaura os teus dados quando sais.</div><div class="section-heading"><div><h3>Tutorial com dados de exemplo</h3><p>Cria temporariamente horário, aulas, teste e TPCs para percorreres um dia completo.</p></div></div><button class="button button-dark" type="button" data-action="debug-start-tutorial"><i data-lucide="play"></i>Iniciar tutorial do dia escolar</button><div class="section-heading"><div><h3>Ir diretamente para um cenário</h3><p>Usa o mesmo template e abre uma situação específica.</p></div></div><div class="debug-scenario-list">' + scenarios + '</div><div class="section-heading"><div><h3>Simular uma hora nos teus dados atuais</h3><p>Útil para testar o teu horário real sem esperar pela hora verdadeira.</p></div></div><div class="debug-time-row"><input id="debugCustomTime" type="datetime-local" value="' + attr(localValue) + '"><button class="button" type="button" data-action="debug-apply-time"><i data-lucide="clock-3"></i>Aplicar</button></div>' + (homeDebug && homeDebug.active ? '<button class="button button-danger" style="margin-top:15px" type="button" data-action="debug-exit"><i data-lucide="x"></i>Sair da simulação</button>' : '');
    openModal("Laboratório da Home", body, { className: "debug-lab-modal" });
  }

  function renderHomeDebugBar() {
    if (!homeDebug || !homeDebug.active) return "";
    var scenarios = homeDebugScenarios(homeDebug.baseDate);
    var scenario = homeDebug.tutorial ? scenarios[homeDebug.index] : null;
    var title = scenario ? scenario.title : "Hora simulada";
    var copy = scenario ? scenario.copy : "A Home está a usar uma hora de teste. Os teus dados reais continuam protegidos.";
    return '<aside class="home-debug-bar"><div><span class="badge badge-dark"><i data-lucide="flask-conical"></i>Admin · simulação</span><strong>' + esc(title) + '</strong><p>' + esc(copy) + '</p></div><div class="home-debug-time"><strong>' + esc(activeHomeNow().toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" })) + '</strong><small>' + (scenario ? (homeDebug.index + 1) + ' / ' + scenarios.length : 'hora manual') + '</small></div><div class="home-debug-actions">' + (scenario ? '<button class="icon-button" type="button" data-action="debug-prev" aria-label="Cenário anterior" ' + (homeDebug.index === 0 ? "disabled" : "") + '><i data-lucide="arrow-left"></i></button><button class="button button-dark button-small" type="button" data-action="debug-next">' + (homeDebug.index === scenarios.length - 1 ? "Terminar" : "Seguinte") + '<i data-lucide="arrow-right"></i></button>' : '') + '<button class="icon-button" type="button" data-action="debug-exit" aria-label="Sair da simulação"><i data-lucide="x"></i></button></div></aside>';
  }

  function homeGreeting(date) {
    var hour = (date || new Date()).getHours();
    if (hour < 12) return "Bom dia";
    if (hour < 19) return "Boa tarde";
    return "Boa noite";
  }

  function homeAtmosphere(date) {
    var now = date || new Date();
    var month = now.getMonth();
    var hour = now.getHours();
    if (hour >= 20 || hour < 6) return { label: "Sessão de biblioteca", icon: "moon-star", className: "is-night" };
    if (month >= 8 && month <= 10) return { label: "Autumn term", icon: "leaf", className: "is-autumn" };
    if (month === 11 || month <= 1) return { label: "Winter term", icon: "snowflake", className: "is-winter" };
    if (month >= 2 && month <= 4) return { label: "Spring term", icon: "flower-2", className: "is-spring" };
    return { label: "Campus hours", icon: "sun", className: "is-day" };
  }

  function homeMinutesCopy(value) {
    var minutes = Math.max(0, Math.round(Number(value) || 0));
    if (minutes < 1) return "agora";
    if (minutes === 1) return "1 min";
    if (minutes < 60) return minutes + " min";
    var hours = Math.floor(minutes / 60);
    var rest = minutes % 60;
    return hours + "h" + (rest ? String(rest).padStart(2, "0") : "");
  }

  function todayClassBlocks(date) {
    var now = date || new Date();
    var day = todayISO(now);
    var weekday = now.getDay();
    var blocks = semesterItems("schedule").filter(function (entry) {
      return Number(entry.weekday) === weekday;
    }).map(function (entry) {
      var lesson = linkedLessonForSlot(entry, day);
      return {
        key: "schedule_" + entry.id,
        schedule: entry,
        lesson: lesson || null,
        course: courseById(entry.courseId),
        dateISO: day,
        start: entry.start || "",
        end: entry.end || entry.start || "",
        startMin: timeMinutes(entry.start),
        endMin: timeMinutes(entry.end || entry.start),
        room: entry.room || "",
        type: entry.type || "Aula"
      };
    });

    semesterItems("lessons").filter(function (lesson) {
      return lesson.date === day;
    }).forEach(function (lesson) {
      var represented = blocks.some(function (block) {
        return block.lesson && block.lesson.id === lesson.id;
      });
      if (represented) return;
      blocks.push({
        key: "lesson_" + lesson.id,
        schedule: null,
        lesson: lesson,
        course: courseById(lesson.courseId),
        dateISO: day,
        start: lesson.start || "",
        end: lesson.end || lesson.start || "",
        startMin: timeMinutes(lesson.start),
        endMin: timeMinutes(lesson.end || lesson.start),
        room: lesson.room || "",
        type: lesson.type || "Aula"
      });
    });

    return blocks.sort(function (a, b) {
      return a.startMin - b.startMin || a.endMin - b.endMin;
    });
  }

  function latestQuizForLesson(lessonId) {
    return state.quizzes.filter(function (quiz) {
      return quiz.lessonId === lessonId && quiz.lastScore != null;
    }).sort(function (a, b) {
      return String(b.lastCompletedAt || "").localeCompare(String(a.lastCompletedAt || ""));
    })[0] || null;
  }

  function availableQuizForLesson(lessonId) {
    return state.quizzes.find(function (quiz) {
      return quiz.lessonId === lessonId && asArray(quiz.questions).length;
    }) || null;
  }

  function lessonCheckButton(lesson, label, className) {
    if (!lesson) return "";
    var quiz = configuredQuizForLesson(lesson.id);
    if (!quiz) return "";
    return '<button class="button ' + (className || "button-dark") + '" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="check-check"></i>' + esc(label || "Fazer quiz") + "</button>";
  }

  function blockOpenButton(block, label, className) {
    if (!block) return "";
    if (block.lesson) {
      return '<button class="button ' + (className || "button-dark") + '" type="button" data-route="lesson" data-id="' + attr(block.lesson.id) + '"><i data-lucide="book-open"></i>' + esc(label || "Abrir aula") + "</button>";
    }
    if (!block.schedule || !block.course) {
      return '<button class="button ' + (className || "button-dark") + '" type="button" data-route="courses"><i data-lucide="library-big"></i>' + esc(label || "Ver cadeiras") + "</button>";
    }
    return '<button class="button ' + (className || "button-dark") + '" type="button" data-action="create-lesson" data-course="' + attr(block.course.id) + '" data-schedule="' + attr(block.schedule.id) + '" data-date="' + attr(block.dateISO) + '" data-start="' + attr(block.start) + '" data-end="' + attr(block.end) + '" data-room="' + attr(block.room) + '" data-type="' + attr(block.type) + '"><i data-lucide="file-plus-2"></i>' + esc(label || "Preparar aula") + "</button>";
  }

  function homeTaskBuckets(day) {
    var today = day || todayISO();
    var all = semesterItems("tasks");
    var todayTasks = all.filter(function (task) {
      return task.dueDate === today;
    });
    var overdue = all.filter(function (task) {
      return !task.done && task.dueDate && task.dueDate < today;
    });
    var homeworkDue = todayTasks.filter(function (task) {
      return task.type !== "lesson-quiz";
    });
    var homeworkPending = homeworkDue.filter(function (task) { return !task.done; });
    return {
      all: all,
      today: todayTasks,
      overdue: overdue,
      homeworkDue: homeworkDue,
      homeworkPending: homeworkPending,
      homeworkDone: homeworkDue.filter(function (task) { return task.done; })
    };
  }

  function homeDailyReport(blocks, date) {
    var day = todayISO(date || new Date());
    var endedLessons = blocks.filter(function (block) {
      return block.lesson && block.endMin <= nowMinutes(date || new Date());
    });
    var courseScores = {};
    blocks.forEach(function (block) {
      if (!block.lesson || !block.course) return;
      var quiz = latestQuizForLesson(block.lesson.id);
      if (!quiz) return;
      if (!courseScores[block.course.id]) {
        courseScores[block.course.id] = { course: block.course, scores: [] };
      }
      courseScores[block.course.id].scores.push(Number(quiz.lastScore) || 0);
    });
    var subjects = Object.keys(courseScores).map(function (courseId) {
      var item = courseScores[courseId];
      var percentage = item.scores.reduce(function (sum, score) { return sum + score; }, 0) / item.scores.length;
      return { course: item.course, percentage: percentage, grade: percentage / 5 };
    });
    var academic = subjects.length ? subjects.reduce(function (sum, item) { return sum + item.grade; }, 0) / subjects.length : null;
    var tasks = homeTaskBuckets(day);
    var quizLessons = endedLessons.filter(function (block) { return !!configuredQuizForLesson(block.lesson.id); });
    var checksDue = quizLessons.length;
    var checksDone = quizLessons.filter(function (block) { return lessonIsBeOnline(block.lesson); }).length;
    var routineTotal = checksDue + tasks.homeworkDue.length;
    var routineDone = checksDone + tasks.homeworkDone.length;
    return {
      academic: academic,
      subjects: subjects,
      checksDue: checksDue,
      checksDone: checksDone,
      homeworkDue: tasks.homeworkDue.length,
      homeworkDone: tasks.homeworkDone.length,
      routine: routineTotal ? Math.round(routineDone / routineTotal * 100) : 100,
      complete: routineTotal ? routineDone >= routineTotal : blocks.length > 0,
      day: day
    };
  }

  function homeContext(date) {
    var now = date || activeHomeNow();
    var minute = nowMinutes(now);
    var blocks = todayClassBlocks(now);
    var current = blocks.find(function (block) {
      return block.startMin <= minute && block.endMin > minute;
    }) || null;
    var next = blocks.find(function (block) {
      return block.startMin > minute;
    }) || null;
    var ended = blocks.filter(function (block) {
      return block.endMin <= minute;
    });
    var previous = ended.length ? ended[ended.length - 1] : null;
    var pendingChecks = ended.filter(function (block) {
      return block.lesson && !!configuredQuizForLesson(block.lesson.id) && !lessonIsBeOnline(block.lesson);
    });
    var tasks = homeTaskBuckets(todayISO(now));
    var report = homeDailyReport(blocks, now);
    var atmosphere = homeAtmosphere(now);
    var context = {
      now: now,
      blocks: blocks,
      current: current,
      previous: previous,
      next: next,
      ended: ended,
      pendingChecks: pendingChecks,
      tasks: tasks,
      report: report,
      atmosphere: atmosphere,
      phase: "quiet",
      phaseLabel: "Hoje",
      title: "",
      copy: "",
      icon: "book-open",
      stat: "",
      statLabel: "",
      primary: "",
      secondary: ""
    };
    var name = (state.profile.name || "estudante").split(/\s+/)[0];
    var greeting = homeGreeting(now) + ", " + name + ".";

    if (current) {
      var elapsed = Math.max(0, minute - current.startMin);
      var remaining = Math.max(0, current.endMin - minute);
      var previousPending = pendingChecks.length ? pendingChecks[pendingChecks.length - 1] : null;
      var following = blocks.find(function (block) { return block.startMin >= current.endMin; }) || null;

      if (elapsed < 10) {
        if (previousPending) {
          context.phase = "settling";
          context.phaseLabel = "Primeiros minutos";
          context.title = "Enquanto a sala enche, fecha " + (previousPending.course ? previousPending.course.name : "a aula anterior") + ".";
          context.copy = (current.course ? current.course.name : "A aula") + " já começou, mas estes primeiros 10 minutos são uma janela tranquila para fazeres o Quiz da aula anterior sem ficares para trás.";
          context.icon = "book-check";
          context.stat = homeMinutesCopy(10 - elapsed);
          context.statLabel = "da janela inicial";
          context.primary = lessonCheckButton(previousPending.lesson, "Fazer Quiz da aula", "button-yellow");
          context.secondary = blockOpenButton(current, "Abrir aula atual", "");
        } else {
          context.phase = "starting";
          context.phaseLabel = "A começar";
          context.title = (current.course ? current.course.name : "A aula") + " irá começar em breve.";
          context.copy = "Abre os materiais e prepara o que precisas para a aula.";
          context.icon = "door-open";
          context.stat = homeMinutesCopy(10 - elapsed);
          context.statLabel = "de chegada";
          context.primary = blockOpenButton(current, "Abrir aula", "button-yellow");
        }
        return context;
      }

      if (remaining <= 15 && current.lesson && configuredQuizForLesson(current.lesson.id) && !lessonIsBeOnline(current.lesson)) {
        var isBackToBack = following && following.startMin - current.endMin <= 10;
        context.phase = "closing";
        context.phaseLabel = "A aula está a terminar";
        context.title = "Fecha " + (current.course ? current.course.name : "a aula") + " com um Quiz da aula.";
        context.copy = "A aula termina às " + current.end + ". " + (isBackToBack ? "A próxima começa às " + following.start + ", por isso podes começar agora e continuar durante os primeiros 10 minutos seguintes." : "Faz uma verificação curta agora para descobrires logo o que ficou menos claro.");
        context.icon = "sparkles";
        context.stat = homeMinutesCopy(remaining);
        context.statLabel = "até terminar";
        context.primary = lessonCheckButton(current.lesson, "Fazer Quiz da aula", "button-yellow");
        context.secondary = isBackToBack ? blockOpenButton(following, "Ver próxima aula", "") : blockOpenButton(current, "Abrir aula", "");
      } else {
        context.phase = "live";
        context.phaseLabel = "Em aula";
        context.title = (current.course ? current.course.name : "A aula") + " está a acontecer agora.";
        context.copy = (current.lesson ? current.lesson.title + ". " : "") + (current.room ? current.room + " · " : "") + "Tens tudo da aula reunido num só sítio.";
        context.icon = "radio-tower";
        context.stat = homeMinutesCopy(remaining);
        context.statLabel = "restantes";
        context.primary = blockOpenButton(current, "Entrar na aula", "button-yellow");
        if (current.lesson && configuredQuizForLesson(current.lesson.id) && !lessonIsBeOnline(current.lesson)) context.secondary = lessonCheckButton(current.lesson, "Quiz da aula", "");
      }
      return context;
    }

    if (next) {
      var untilNext = Math.max(0, next.startMin - minute);
      var lastPending = pendingChecks.length ? pendingChecks[pendingChecks.length - 1] : null;
      if (untilNext <= 10) {
        context.phase = "soon";
        context.phaseLabel = "Quase a começar";
        context.title = (next.course ? next.course.name : "A próxima aula") + " começa daqui a " + homeMinutesCopy(untilNext) + ".";
        context.copy = (next.lesson ? next.lesson.title + " · " : "") + next.start + "–" + next.end + (next.room ? " · " + next.room : "") + ". " + (lastPending ? "O Quiz da aula anterior continua disponível nos primeiros 10 minutos da aula." : "Abre a aula e deixa tudo preparado.");
        context.icon = "alarm-clock";
        context.stat = next.start;
        context.statLabel = "começa";
        context.primary = blockOpenButton(next, next.lesson ? "Preparar-me" : "Preparar aula", "button-yellow");
        if (lastPending) context.secondary = lessonCheckButton(lastPending.lesson, "Fechar aula anterior", "");
      } else if (lastPending) {
        context.phase = "between";
        context.phaseLabel = "Entre aulas";
        context.title = "A aula terminou. Fecha o ciclo antes da próxima.";
        context.copy = "Faz o Quiz da aula de " + (lastPending.course ? lastPending.course.name : lastPending.lesson.title) + ". Depois, a próxima aula começa às " + next.start + ".";
        context.icon = "book-check";
        context.stat = homeMinutesCopy(untilNext);
        context.statLabel = "até à próxima";
        context.primary = lessonCheckButton(lastPending.lesson, "Fazer Quiz da aula", "button-yellow");
        context.secondary = blockOpenButton(next, "Ver próxima aula", "");
      } else if (untilNext <= 45) {
        context.phase = "soon";
        context.phaseLabel = "Quase a começar";
        context.title = (next.course ? next.course.name : "A próxima aula") + " começa daqui a " + homeMinutesCopy(untilNext) + ".";
        context.copy = (next.lesson ? next.lesson.title + " · " : "") + next.start + "–" + next.end + (next.room ? " · " + next.room : "") + ". Abre a aula e deixa tudo preparado.";
        context.icon = "alarm-clock";
        context.stat = next.start;
        context.statLabel = "começa";
        context.primary = blockOpenButton(next, next.lesson ? "Preparar-me" : "Preparar aula", "button-yellow");
        context.secondary = '<button class="button" type="button" data-route="planner"><i data-lucide="calendar-days"></i>Ver o dia</button>';
      } else if (!ended.length) {
        context.phase = "before";
        context.phaseLabel = "Antes das aulas";
        context.title = greeting + " O teu dia começa às " + next.start + ".";
        context.copy = blocks.length + " aula" + (blocks.length === 1 ? "" : "s") + " hoje. A primeira é " + (next.course ? next.course.name : "aula") + (next.room ? " em " + next.room : "") + ".";
        context.icon = atmosphere.icon;
        context.stat = homeMinutesCopy(untilNext);
        context.statLabel = "até começar";
        context.primary = blockOpenButton(next, "Preparar primeira aula", "button-yellow");
        context.secondary = '<button class="button" type="button" data-route="planner"><i data-lucide="calendar-days"></i>Ver horário</button>';
      } else {
        context.phase = "between";
        context.phaseLabel = "Intervalo";
        context.title = "Respira. A próxima aula é " + (next.course ? next.course.name : "daqui a pouco") + ".";
        context.copy = "Começa às " + next.start + (next.room ? " em " + next.room : "") + ". Podes rever o essencial ou simplesmente fazer uma pausa.";
        context.icon = "coffee";
        context.stat = homeMinutesCopy(untilNext);
        context.statLabel = "de intervalo";
        context.primary = blockOpenButton(next, "Ver próxima aula", "button-yellow");
        context.secondary = '<button class="button" type="button" data-route="study"><i data-lucide="book-open-check"></i>Revisão curta</button>';
      }
      return context;
    }

    if (pendingChecks.length) {
      var firstPending = pendingChecks[0];
      context.phase = "after";
      context.phaseLabel = "Depois das aulas";
      context.title = "As aulas acabaram. Faltam " + pendingChecks.length + " Quiz da aula" + (pendingChecks.length === 1 ? "" : "s") + ".";
      context.copy = "Faz primeiro a verificação rápida do que aprendeste. O TPC vem depois — são duas coisas diferentes.";
      context.icon = "book-check";
      context.stat = String(pendingChecks.length);
      context.statLabel = pendingChecks.length === 1 ? "aula por fechar" : "aulas por fechar";
      context.primary = lessonCheckButton(firstPending.lesson, "Fazer o próximo", "button-yellow");
      context.secondary = '<button class="button" type="button" data-route="study"><i data-lucide="sparkles"></i>Ver estudo</button>';
      return context;
    }

    if (tasks.homeworkPending.length || tasks.overdue.length) {
      var totalPending = tasks.homeworkPending.length + tasks.overdue.length;
      context.phase = "homework";
      context.phaseLabel = blocks.length ? "Depois das aulas" : "Sessão de estudo";
      context.title = blocks.length ? "Foi um bom dia de escola. Agora, fecha os TPCs." : "Hoje a biblioteca chama por ti.";
      context.copy = totalPending + " tarefa" + (totalPending === 1 ? "" : "s") + " à tua espera. Faz uma de cada vez e o relatório do dia fecha sozinho.";
      context.icon = "notebook-pen";
      context.stat = String(totalPending);
      context.statLabel = totalPending === 1 ? "TPC pendente" : "TPCs pendentes";
      context.primary = '<button class="button button-yellow" type="button" data-action="start-homework-session"><i data-lucide="play"></i>Começar TPCs</button>';
      context.secondary = '<button class="button" type="button" data-action="add-task"><i data-lucide="plus"></i>Adicionar TPC</button>';
      return context;
    }

    if (blocks.length) {
      context.phase = "complete";
      context.phaseLabel = "Dia concluído";
      context.title = "Mochila fechada. O dia escolar está completo.";
      context.copy = report.academic == null ? "Cumpriste a rotina de hoje. Amanhã, a Home volta a adaptar-se ao teu horário." : "Nota académica provisória: " + round(report.academic, 1) + "/20. Agora podes descansar sem aquela sensação de que te esqueceste de alguma coisa.";
      context.icon = "badge-check";
      context.stat = report.academic == null ? "✓" : round(report.academic, 1);
      context.statLabel = report.academic == null ? "rotina completa" : "nota do dia";
      context.primary = '<button class="button button-yellow" type="button" data-action="view-report-card"><i data-lucide="award"></i>Ver Report Card</button>';
      context.secondary = '<button class="button" type="button" data-route="planner"><i data-lucide="calendar-days"></i>Ver amanhã</button>';
      return context;
    }

    context.phase = "quiet";
    context.phaseLabel = atmosphere.label;
    context.title = greeting + " Hoje tens espaço para estudar ao teu ritmo.";
    context.copy = "Sem aulas marcadas. Podes adiantar um TPC, rever matéria ou simplesmente aproveitar o dia livre.";
    context.icon = atmosphere.icon;
    context.stat = "Livre";
    context.statLabel = "agenda escolar";
    context.primary = '<button class="button button-yellow" type="button" data-route="study"><i data-lucide="sparkles"></i>Estudar um pouco</button>';
    context.secondary = '<button class="button" type="button" data-route="planner"><i data-lucide="calendar-days"></i>Ver calendário</button>';
    return context;
  }

  function renderHomeStep(icon, label, meta, status) {
    return '<div class="school-step ' + (status || "") + '"><span><i data-lucide="' + icon + '"></i></span><div><strong>' + esc(label) + '</strong><small>' + esc(meta) + '</small></div>' + (status === "is-done" ? '<i class="school-step-check" data-lucide="check"></i>' : "") + "</div>";
  }

  function renderHomeCanteenStep(date) {
    var day = todayISO(date || new Date());
    var visit = canteenVisitFor(day, "lunch");
    var completedVisit = visit && visit.completedAt ? visit : null;
    var menu = canteenMealForDate(day, "lunch");
    if (!visit && !menu) return "";
    var status = completedVisit ? "is-done" : "";
    var meta = completedVisit
      ? (completedVisit.dish && completedVisit.dish.description || "Almoço concluído")
      : visit && visit.ticketIssuedAt
        ? "Ticket emitido · levantar refeição"
        : "Escolher e emitir ticket";
    return '<button class="school-step school-step-button ' + status + '" type="button" data-route="canteen"><span><i data-lucide="utensils"></i></span><div><strong>Cantina</strong><small>' + esc(meta) + '</small></div>' + (completedVisit ? '<i class="school-step-check" data-lucide="check"></i>' : '<i class="school-step-arrow" data-lucide="chevron-right"></i>') + '</button>';
  }

  function renderLessonCheckRow(block) {
    if (!block || !block.lesson) return "";
    var course = block.course;
    var quiz = configuredQuizForLesson(block.lesson.id);
    if (!quiz || lessonIsBeOnline(block.lesson)) return "";
    var action = '<button class="button button-dark button-small" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="play"></i>Fazer</button>';
    return '<div class="list-row school-action-row"><span class="list-icon yellow"><i data-lucide="check-check"></i></span><span class="list-content"><strong>Quiz da aula · ' + esc(block.lesson.title) + '</strong><small>' + esc(course ? course.name : "Aula") + ' · 3–7 minutos</small></span><span class="badge badge-violet">Agora</span>' + action + "</div>";
  }

  function renderDailyReportCard(report, blocks) {
    var subjectRows = report.subjects.length ? report.subjects.slice(0, 4).map(function (item) {
      return '<div class="report-subject"><span><i data-lucide="book-open"></i>' + esc(item.course.name) + '</span><strong>' + round(item.grade, 1) + '</strong></div>';
    }).join("") : '<p class="report-empty">Faz os Quizzes da aula para construíres a nota académica de hoje.</p>';
    var grade = report.academic == null ? "—" : round(report.academic, 1);
    var status = report.complete ? "Fechado" : "Em construção";
    return '<article class="card span-5 daily-report-card ' + (report.complete ? "is-complete" : "") + '"><div class="card-title-row"><div><p class="card-label">Report Card · Hoje</p><h3>O teu dia escolar</h3></div><span class="badge ' + (report.complete ? "badge-mint" : "badge-yellow") + '">' + status + '</span></div><div class="report-grade"><strong>' + grade + '</strong><span>/20<br>nota académica</span></div><div class="report-subjects">' + subjectRows + '</div><div class="report-routine"><span><strong>' + report.routine + '%</strong> rotina escolar</span><div class="mini-progress"><span style="width:' + report.routine + '%"></span></div><small>' + report.checksDone + '/' + report.checksDue + ' Quizzes da aula · ' + report.homeworkDone + '/' + report.homeworkDue + ' TPCs</small></div><button class="button button-small" type="button" data-action="view-report-card"><i data-lucide="arrow-down"></i>Ver Report Card</button></article>';
  }

  function renderHome() {
    setHeader("Hoje", "O teu dia escolar");
    var semester = currentSemester();
    if (!semester) {
      return '<div class="page-head"><div><h2>Novo semestre</h2><p>Configura as cadeiras e o horário para a Home começar a acompanhar o teu dia.</p></div></div>' + emptyState("calendar-plus", "Sem semestre ativo", "Quando houver horário, a Home passa a mostrar automaticamente o que está a acontecer e o que vem a seguir.", "new-semester", "Criar semestre");
    }

    var context = homeContext(activeHomeNow());
    var blocks = context.blocks;
    var report = context.report;
    var currentMinute = nowMinutes(context.now);
    var completedClasses = blocks.filter(function (block) { return block.endMin < currentMinute; }).length;
    var endedLessons = blocks.filter(function (block) { return block.lesson && block.endMin < currentMinute; });
    var configuredEndedLessons = endedLessons.filter(function (block) { return !!configuredQuizForLesson(block.lesson.id); });
    var completedChecks = configuredEndedLessons.filter(function (block) { return lessonIsBeOnline(block.lesson); }).length;
    var tasks = context.tasks;
    var afterSchoolItems = context.pendingChecks.map(renderLessonCheckRow).join("");
    var homeworkList = tasks.overdue.concat(tasks.homeworkPending).filter(function (task, index, array) {
      return array.findIndex(function (candidate) { return candidate.id === task.id; }) === index;
    }).slice(0, 4);
    afterSchoolItems += homeworkList.map(renderTaskRow).join("");
    if (!afterSchoolItems) {
      afterSchoolItems = emptyState("check-check", "Nada pendente depois das aulas", "Quando houver Quizzes da aula ou TPCs, aparecem aqui pela ordem certa.", "add-task", "Adicionar TPC");
    }

    var nextAssessment = semesterItems("assessments").filter(function (item) {
      return item.date && item.date >= todayISO();
    }).sort(function (a, b) {
      return String(a.date).localeCompare(String(b.date));
    })[0] || null;
    var nextAssessmentCourse = nextAssessment ? courseById(nextAssessment.courseId) : null;
    var nextAssessmentHtml = nextAssessment
      ? '<div class="next-date-feature"><span class="list-icon orange"><i data-lucide="' + assessmentIcon(nextAssessment.type) + '"></i></span><div><p class="card-label">Próxima avaliação</p><h3>' + esc(nextAssessment.title) + '</h3><p>' + esc(nextAssessmentCourse ? nextAssessmentCourse.name : "Avaliação") + ' · ' + relativeDate(nextAssessment.date) + (nextAssessment.time ? " às " + esc(nextAssessment.time) : "") + '</p></div><span class="badge badge-yellow">' + esc(nextAssessment.type || "Avaliação") + "</span></div>"
      : emptyState("calendar-check", "Sem avaliações próximas", "Quando adicionares uma avaliação, a Home liga-a ao teu plano diário.", "add-assessment", "Adicionar avaliação");

    var classesStepStatus = !blocks.length ? "" : completedClasses >= blocks.length ? "is-done" : context.current ? "is-active" : completedClasses ? "is-active" : "";
    var checksStepStatus = !configuredEndedLessons.length ? "" : completedChecks >= configuredEndedLessons.length ? "is-done" : context.pendingChecks.length ? "is-active" : "";
    var tpcDone = tasks.homeworkDone.length;
    var tpcTotal = tasks.homeworkDue.length;
    var tpcStepStatus = !tpcTotal ? (context.phase === "complete" ? "is-done" : "") : tpcDone >= tpcTotal ? "is-done" : (context.phase === "homework" ? "is-active" : "");
    var reportStatus = report.complete && blocks.length ? "is-done" : context.phase === "complete" ? "is-active" : "";

    var timeline = renderHomeStep("graduation-cap", "Aulas", blocks.length ? completedClasses + "/" + blocks.length + " concluídas" : "Sem aulas hoje", classesStepStatus)
      + renderHomeCanteenStep(context.now)
      + renderHomeStep("check-check", "Quizzes da aula", configuredEndedLessons.length ? completedChecks + "/" + configuredEndedLessons.length + " feitos" : "Configura-os dentro de cada aula", checksStepStatus)
      + renderHomeStep("notebook-pen", "TPC", tpcTotal ? tpcDone + "/" + tpcTotal + " concluídos" : "Sem TPC para hoje", tpcStepStatus)
      + renderHomeStep("award", "Report Card", report.complete && blocks.length ? "Dia fechado" : "Atualiza ao longo do dia", reportStatus);

    var phaseBadge = context.phaseLabel && context.phaseLabel !== context.atmosphere.label ? '<b>' + esc(context.phaseLabel) + '</b>' : '';
    var hero = '<section class="school-now-card span-12 ' + attr(context.atmosphere.className) + ' phase-' + attr(context.phase) + '"><div class="school-now-copy"><div class="school-now-kicker"><span><i data-lucide="' + attr(context.atmosphere.icon) + '"></i>' + esc(context.atmosphere.label) + '</span>' + phaseBadge + '</div><h2>' + esc(context.title) + '</h2><p>' + esc(context.copy) + '</p><div class="school-now-actions">' + context.primary + context.secondary + '</div></div><div class="school-now-status"><span class="school-now-icon"><i data-lucide="' + attr(context.icon) + '"></i></span><strong>' + esc(context.stat) + '</strong><small>' + esc(context.statLabel) + '</small><div class="school-now-glow"></div></div></section>';

    return renderHomeDebugBar() + hero
      + '<div class="bento-grid home-school-grid" style="margin-top:15px">'
      + '<article class="card span-5 school-path-card"><div class="card-title-row"><div><p class="card-label">O teu dia</p><h3>Da aula ao fim do dia</h3><p class="card-subtitle">A Home muda o próximo passo à medida que vais avançando.</p></div><span class="metric-icon"><i data-lucide="route"></i></span></div><div class="school-steps">' + timeline + "</div></article>"
      + '<article class="card span-7 after-school-card"><div class="card-title-row"><div><p class="card-label">A seguir</p><h3>Quizzes da aula e TPC</h3><p class="card-subtitle">Primeiro verificas o que aprendeste. Depois fazes os TPCs numa sessão guiada.</p></div><div class="list-actions"><button class="button button-small" type="button" data-action="start-homework-session"><i data-lucide="play"></i>Fazer TPCs</button><button class="row-button" type="button" data-action="add-task" aria-label="Adicionar TPC"><i data-lucide="plus"></i></button></div></div><div class="list-stack after-school-list">' + afterSchoolItems + "</div></article>"
      + renderDailyReportCard(report, blocks)
      + '<article class="card span-7 next-date-card"><div class="card-title-row"><div><p class="card-label">No horizonte</p><h3>O que merece atenção</h3></div><button class="button button-small" type="button" data-route="planner"><i data-lucide="calendar-days"></i>Calendário</button></div><div style="margin-top:15px">' + nextAssessmentHtml + '</div><div class="home-mini-stats"><span><strong>' + tasks.overdue.length + '</strong> em atraso</span><span><strong>' + semesterItems("questions").length + '</strong> perguntas antigas</span><span><strong>' + overallProgress() + '%</strong> matéria dominada</span></div></article>'
      + "</div>";
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
    var lockedHomework = (task.type === "homework" || task.type === "tpc") && (task.lockedContent || task.configuredFromPrompt || asArray(task.contentBlocks).length);
    var oneTimeComplete = task.done && (task.type === "lesson-quiz" || lockedHomework);
    var lessonAction = "";
    if (task.lessonId && task.type === "lesson-quiz" && !task.done) {
      lessonAction = '<button class="button button-dark button-small task-open-button" type="button" data-action="do-beonline-quiz" data-lesson="' + attr(task.lessonId) + '"><i data-lucide="play"></i>Fazer quiz</button>';
    } else if (lockedHomework && !task.done) {
      lessonAction = '<button class="button button-dark button-small task-open-button" type="button" data-action="start-homework-session" data-task="' + attr(task.id) + '"><i data-lucide="play"></i>Fazer</button>';
    } else if (lockedHomework && task.done) {
      lessonAction = '<button class="button button-small task-open-button" type="button" data-action="view-lesson-homework" data-id="' + attr(task.id) + '"><i data-lucide="eye"></i>Ver</button>';
    } else if (task.lessonId) {
      lessonAction = '<button class="row-button task-open-button" type="button" data-route="lesson" data-id="' + attr(task.lessonId) + '" aria-label="Abrir aula"><i data-lucide="arrow-up-right"></i></button>';
    }
    var checkControl = oneTimeComplete
      ? '<span class="check-button is-done" aria-label="Concluído"><i data-lucide="check"></i></span>'
      : '<button class="check-button ' + (task.done ? "is-done" : "") + '" type="button" data-action="toggle-task" data-id="' + attr(task.id) + '" aria-label="' + (task.done ? "Reabrir tarefa" : "Concluir tarefa") + '">' + (task.done ? '<i data-lucide="check"></i>' : "") + '</button>';
    return '<div class="list-row ' + (task.done ? "is-done" : "") + '">' + checkControl + '<span class="list-icon ' + (task.type === "review" || task.type === "lesson-quiz" ? "yellow" : "") + '"><i data-lucide="' + taskIcon(task.type) + '"></i></span><span class="list-content"><strong>' + esc(task.title) + '</strong><small>' + esc(course ? course.name : "Pessoal") + ' · ' + relativeDate(task.dueDate) + (task.dueTime ? " às " + esc(task.dueTime) : "") + '</small></span><span class="badge ' + (task.type === "lesson-quiz" ? "badge-violet" : task.priority === "high" ? "badge-danger" : "") + '">' + esc(label) + '</span>' + lessonAction + "</div>";
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
      ["overview", "Resumo"], ["lessons", "Aulas"], ["notebook", "Caderno"], ["materials", "Materiais"], ["assessments", "Avaliações"],
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
    if (tab === "notebook") return renderCourseNotebook(course, archived);
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
    if (!lesson) return;
    output = output === "notes" ? "notes" : "quiz";
    if (!AI) {
      openModal("IA indisponível", '<div class="lesson-ai-empty"><span class="metric-icon"><i data-lucide="triangle-alert"></i></span><div><h3>O módulo de IA não carregou</h3><p>Atualiza a página. Se continuar, abre a Twenty no Chrome com WebGPU disponível.</p></div></div>', { footer: '<footer class="modal-foot"><button class="button button-dark" type="button" data-action="close-modal">Fechar</button></footer>' });
      return;
    }
    var available = lessonAIAvailableLessons(lesson.courseId);
    if (!available.length) {
      openModal(output === "quiz" ? "Criar quiz com IA" : "Criar apontamentos com IA", '<div class="lesson-ai-empty"><span class="metric-icon"><i data-lucide="presentation"></i></span><div><span class="badge badge-violet"><i data-lucide="sparkles"></i>Twenty AI</span><h3>Primeiro adiciona um PowerPoint</h3><p>Carrega um ficheiro .pptx nesta aula. A Twenty extrai os slides, sincroniza o ficheiro e depois deixa-te escolher as aulas usadas pela IA.</p></div></div>', { footer: '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Cancelar</button><button class="button button-dark" type="button" data-action="add-material" data-course="' + attr(lesson.courseId) + '" data-lesson="' + attr(lesson.id) + '"><i data-lucide="file-up"></i>Carregar PowerPoint</button></footer>' });
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
    return '<article class="material-card"><div class="material-preview"><i data-lucide="' + icon + '"></i>' + materialYearBadge(material, course) + '</div><h4>' + esc(material.title) + '</h4><p>' + esc(lesson ? lesson.title : "Biblioteca da cadeira") + (material.fileName ? " · " + esc(material.fileName) : "") + '</p><div class="material-actions"><span class="badge ' + (synced ? 'badge-mint' : 'badge-yellow') + '">' + (synced ? 'Sincronizado' : 'Local') + '</span><span class="list-actions">' + (aiReady && !archived ? '<button class="row-button row-button-ai" type="button" data-action="configure-lesson-content" data-kind="quiz" data-lesson="' + attr(lesson.id) + '" aria-label="Configurar quiz da aula"><i data-lucide="sparkles"></i></button>' : '') + (!synced && material.blobId && !archived ? '<button class="row-button" type="button" data-action="sync-material" data-id="' + attr(material.id) + '" aria-label="Sincronizar ficheiro com o Git"><i data-lucide="cloud-upload"></i></button>' : '') + '<button class="row-button" type="button" data-action="open-material" data-id="' + attr(material.id) + '" aria-label="Abrir material"><i data-lucide="eye"></i></button>' + (!archived ? '<button class="row-button" type="button" data-action="delete-entity" data-kind="materials" data-id="' + attr(material.id) + '" aria-label="Remover material"><i data-lucide="trash-2"></i></button>' : "") + '</span></div>' + (aiReady ? '<small class="material-ai-ready"><i data-lucide="brain"></i>' + asArray(material.slides).length + ' slides prontos para IA</small>' : '') + '</article>';
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
      var origin = quiz.configuredFromPrompt ? "Prompt IA" : quiz.generatedFromPastQuestions ? "Perguntas anteriores" : "Manual";
      var actions = '<button class="button button-dark" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="play"></i>Começar</button>';
      if (lesson) {
        actions = '<button class="button" type="button" data-action="view-lesson-quiz" data-lesson="' + attr(lesson.id) + '"><i data-lucide="eye"></i>Ver</button>' + actions;
      } else if (!archived) {
        actions += '<button class="button" type="button" data-action="add-quiz-question" data-id="' + attr(quiz.id) + '"><i data-lucide="plus"></i>Manual</button>' + (hasPast ? '<button class="button" type="button" data-action="add-past-to-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="history"></i>Anteriores</button>' : '') + '<button class="button button-danger" type="button" data-action="delete-entity" data-kind="quizzes" data-id="' + attr(quiz.id) + '"><i data-lucide="trash-2"></i></button>';
      }
      return '<article class="card span-4"><div class="card-title-row"><div><div class="question-meta"><span class="badge badge-violet">' + asArray(quiz.questions).length + ' pergunta(s)</span><span class="badge">' + origin + '</span>' + (lesson ? '<span class="badge badge-mint"><i data-lucide="lock-keyhole"></i> Aula</span>' : '') + '</div><h3 style="margin-top:11px">' + esc(quiz.title) + '</h3><p class="card-subtitle">' + esc(lesson ? lesson.title : "Quiz geral da cadeira") + '</p></div><span class="metric-icon"><i data-lucide="sparkles"></i></span></div><div class="live-actions" style="margin-top:21px">' + actions + '</div></article>';
    }).join("");
    return '<div class="page-head"><div><h2>Quizzes da cadeira</h2><p>Os quizzes associados a aulas ficam bloqueados depois de configurados. Os quizzes gerais continuam editáveis.</p></div>' + (!archived ? '<div class="page-actions"><button class="button button-dark" type="button" data-action="add-quiz" data-course="' + attr(course.id) + '"><i data-lucide="plus"></i>Novo quiz geral</button></div>' : "") + '</div><div class="bento-grid">' + (quizzes.length ? html : '<div class="span-12">' + emptyState("sparkles", "Ainda não há quizzes", "Configura um quiz em cada aula ou cria um quiz geral para a cadeira.", "add-quiz", "Criar quiz geral") + "</div>") + "</div>";
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


  function lessonSlides(lessonId) {
    return state.materials.filter(function (material) {
      return material.lessonId === lessonId && (material.kind === "slides" || /\.pptx$/i.test(material.fileName || material.title || ""));
    });
  }

  function lessonBuilderLabel(kind) {
    if (kind === "quiz") return "quiz da aula";
    if (kind === "homework") return "TPC";
    return "apontamentos";
  }

  function lessonBuilderSourceText(lesson, materialIds, includePast) {
    var materials = lessonSlides(lesson.id).filter(function (material) { return materialIds.indexOf(material.id) >= 0; });
    var chunks = materials.map(function (material) {
      var slides = asArray(material.slides).slice(0, 120).map(function (slide) {
        return '[Slide ' + (slide.number || "?") + '] ' + [slide.title, slide.text].filter(Boolean).join(" — ");
      }).filter(Boolean).join("\n");
      return 'FICHEIRO: ' + (material.fileName || material.title || "Slides") + '\n' + (slides || "O texto dos slides não está extraído; analisa o ficheiro PPTX anexado.");
    });
    if (includePast) {
      var past = pastQuestionsForLesson(lesson.id).map(function (question) {
        return '- ' + question.prompt + (question.answer ? '\n  Resposta guardada: ' + question.answer : "");
      });
      if (past.length) chunks.push('PERGUNTAS DE ANOS ANTERIORES:\n' + past.join("\n"));
    }
    return chunks.join("\n\n").slice(0, 30000);
  }


  function attachSlideMaterialReferences(value, materials) {
    materials = asArray(materials);
    return normalizeContentBlocks(value).map(function (sourceBlock) {
      var block = Object.assign({}, sourceBlock);
      if (String(block.type || "").toLowerCase() === "slide-image" && !block.materialId && materials.length) {
        var requested = String(block.sourceFile || block.fileName || "").trim().toLowerCase();
        var match = requested ? materials.find(function (material) {
          return [material.fileName, material.title].some(function (name) {
            var normalized = String(name || "").trim().toLowerCase();
            return normalized && (normalized === requested || normalized.indexOf(requested) >= 0 || requested.indexOf(normalized) >= 0);
          });
        }) : null;
        block.materialId = (match || materials[0]).id;
      }
      if (String(block.type || "").toLowerCase() === "list") {
        block.items = asArray(block.items).map(function (item) { return attachSlideMaterialReferences(item, materials); });
      }
      return block;
    });
  }

  function lessonBuilderSchema(kind, lesson) {
    var commonBlocks = '[{"type":"text","text":"..."},{"type":"latex","latex":"\\\\frac{a}{b}","display":true},{"type":"code","language":"python","code":"..."},{"type":"svg","svg":"<svg viewBox=\\"0 0 640 360\\">...</svg>","alt":"..."},{"type":"slide-image","sourceFile":"nome-do-ficheiro.pptx","slideNumber":12,"description":"imagem ou diagrama exato a usar","alt":"..."}]';
    if (kind === "quiz") {
      return '{\n  "title": "Quiz da aula · ' + lesson.title.replace(/"/g, "\\\"") + '",\n  "questions": [\n    {\n      "promptBlocks": ' + commonBlocks + ',\n      "options": [' + commonBlocks + ', ' + commonBlocks + ', ' + commonBlocks + ', ' + commonBlocks + '],\n      "correctIndex": 0,\n      "explanationBlocks": ' + commonBlocks + ',\n      "sourceSlides": [1, 2],\n      "difficulty": "medium"\n    }\n  ]\n}';
    }
    if (kind === "homework") {
      return '{\n  "title": "TPC · ' + lesson.title.replace(/"/g, "\\\"") + '",\n  "estimatedMinutes": 30,\n  "instructionsBlocks": ' + commonBlocks + ',\n  "solutionBlocks": ' + commonBlocks + ',\n  "checklist": ["Passo 1", "Passo 2"],\n  "extraTasks": [\n    {"title":"Tarefa adicional pedida pelo professor","blocks":' + commonBlocks + ',"optional":false}\n  ]\n}';
    }
    return '{\n  "title": "Apontamentos · ' + lesson.title.replace(/"/g, "\\\"") + '",\n  "paper": "lined",\n  "blocks": ' + commonBlocks + '\n}';
  }

  function buildLessonAIPrompt(lesson, kind, materialIds, includePast) {
    var course = courseById(lesson.courseId);
    var sourceText = lessonBuilderSourceText(lesson, materialIds, includePast);
    var purpose = kind === "quiz"
      ? "Cria um quiz curto para confirmar imediatamente o que foi aprendido nesta aula. As perguntas devem ser claras, úteis e baseadas exclusivamente nas fontes."
      : kind === "homework"
        ? "Cria um TPC para fazer em casa depois das aulas. Não repitas simplesmente o quiz: usa aplicação, prática, exercícios e consolidação."
        : "Cria apontamentos completos, bem estruturados e úteis para revisão, sem inventar matéria que não esteja nas fontes.";
    return [
      "CONTEXTO",
      "Estou a configurar conteúdo para a app Twenty, um sistema académico pessoal.",
      "Cadeira: " + (course ? course.name : "Cadeira"),
      "Aula: " + lesson.title,
      "Data: " + (lesson.date || "não indicada"),
      "Tópicos: " + (lesson.topics || "não indicados"),
      "",
      "OBJETIVO",
      purpose,
      "",
      "REGRAS DE CONTEÚDO",
      "- Devolve APENAS JSON válido, sem texto antes ou depois e sem blocos markdown.",
      "- Português de Portugal.",
      "- Usa texto normal sempre que for suficiente.",
      "- Para fórmulas, usa blocos type=latex e escreve LaTeX sem delimitadores $, $$, \\( ou \\[. Exemplo: \\\\frac{v^2}{2a}.",
      "- Para código, usa type=code com language e code. O código será mostrado numa interface estilo IDE.",
      "- Para diagramas simples, prefere type=svg. Usa SVG puro com viewBox e sem JavaScript, scripts, foreignObject ou ligações externas.",
      "- Se uma pergunta ou resposta depender de uma imagem específica dos slides, usa type=slide-image com sourceFile, slideNumber, description e alt. Só pede imagens quando forem realmente necessárias, por exemplo gráficos, circuitos, geometria, física, matemática visual ou interpretação de diagramas.",
      "- Uma resposta, opção ou explicação pode ser composta por texto, imagem pedida, SVG, LaTeX ou código, conforme o conteúdo exigir.",
      "- Não inventes imagens que não existem nos slides; quando necessário, descreve exatamente que zona/diagrama deve aparecer.",
      kind === "quiz" ? "- Cria entre 5 e 10 perguntas. Usa quatro opções por pergunta e um único correctIndex." : "",
      kind === "homework" ? "- O TPC deve ser diferente do quiz da aula e adequado para fazer em casa. Inclui solução ou critérios de correção." : "",
      kind === "homework" ? "- instructionsBlocks representa a tarefa principal. Usa extraTasks apenas para pedidos adicionais concretos do professor; cada item tem title, blocks e optional." : "",
      kind === "homework" ? "- Dá uma estimativa realista em estimatedMinutes, porque a Twenty vai usar esse valor no timer da sessão de TPC." : "",
      kind === "notes" ? "- Este conteúdo será ACRESCENTADO depois dos apontamentos já escritos pelo utilizador. Não reescrevas, apagues nem substituas os apontamentos existentes; evita repetir ideias já presentes." : "",
      kind === "notes" ? "- Organiza a nova continuação por secções e usa fórmulas, código ou diagramas apenas quando melhorarem realmente a compreensão." : "",
      kind === "notes" && lesson.notes ? "- Apontamentos existentes para não repetires: " + lesson.notes : "",
      "",
      "FORMATO EXATO ESPERADO PELA TWENTY",
      lessonBuilderSchema(kind, lesson),
      "",
      "FONTES",
      sourceText || "Não existe texto extraído. Analisa o PowerPoint anexado à conversa.",
    ].filter(function (line) { return line !== "" || true; }).join("\n");
  }

  function lessonBuilderMaterialChoices(lesson) {
    var materials = lessonSlides(lesson.id);
    if (!materials.length) return '<div class="builder-empty"><i data-lucide="presentation"></i><p><strong>Sem PowerPoint nesta aula</strong><small>Carrega primeiro os slides. Depois volta aqui para criar o prompt.</small></p><button class="button button-small" type="button" data-action="add-material" data-course="' + attr(lesson.courseId) + '" data-lesson="' + attr(lesson.id) + '"><i data-lucide="file-up"></i>Carregar slides</button></div>';
    return '<div class="builder-source-list">' + materials.map(function (material) {
      return '<label class="builder-source"><input type="checkbox" name="materialIds" value="' + attr(material.id) + '" checked><span class="list-icon"><i data-lucide="presentation"></i></span><span><strong>' + esc(material.title || material.fileName || "PowerPoint") + '</strong><small>' + (material.slideCount || asArray(material.slides).length || 0) + ' slides · ' + (material.remoteFile ? "sincronizado" : "local") + '</small></span></label>';
    }).join("") + '</div>';
  }

  function openLessonBuilder(lessonId, kind) {
    var lesson = lessonById(lessonId);
    if (!lesson) return;
    if (kind === "quiz" && configuredQuizForLesson(lesson.id)) { viewLessonQuiz(lesson.id); return; }
    if (kind === "homework" && homeworkForLesson(lesson.id)) { viewLessonHomework(homeworkForLesson(lesson.id).id); return; }
    var materials = lessonSlides(lesson.id);
    var ids = materials.map(function (material) { return material.id; });
    var includePast = pastQuestionsForLesson(lesson.id).length > 0;
    var prompt = buildLessonAIPrompt(lesson, kind, ids, includePast);
    var body = '<form id="lessonBuilderForm" data-kind="' + attr(kind) + '" data-lesson="' + attr(lesson.id) + '"><div class="builder-intro"><span class="metric-icon"><i data-lucide="' + (kind === "quiz" ? "check-check" : kind === "homework" ? "notebook-pen" : "book-open-text") + '"></i></span><div><p class="card-label">Configurar ' + esc(lessonBuilderLabel(kind)) + '</p><h3>' + esc(lesson.title) + '</h3><p>Escolhe as fontes, copia o prompt para a IA juntamente com o PowerPoint e cola aqui o JSON devolvido.</p></div></div><div class="section-heading"><div><h3>Slides da aula</h3><p>O prompt inclui o texto extraído e identifica os ficheiros que deves anexar à IA.</p></div></div>' + lessonBuilderMaterialChoices(lesson) + '<label class="builder-past-toggle"><input type="checkbox" name="includePast" ' + (includePast ? "checked" : "") + '><span><strong>Incluir perguntas de anos anteriores</strong><small>' + pastQuestionsForLesson(lesson.id).length + ' pergunta(s) associada(s) a esta aula</small></span></label><div class="builder-prompt-head"><div><h3>Prompt para a IA</h3><p>O formato inclui texto, imagens dos slides, SVG, LaTeX e código.</p></div><button class="button button-small" type="button" data-action="copy-lesson-builder-prompt"><i data-lucide="copy"></i>Copiar prompt</button></div><textarea id="lessonBuilderPrompt" class="builder-prompt" readonly>' + esc(prompt) + '</textarea><div class="builder-prompt-head"><div><h3>Resposta da IA</h3><p>Cola apenas o JSON. A Twenty valida antes de guardar.</p></div><span class="badge badge-violet">JSON</span></div><textarea id="lessonBuilderResponse" class="builder-response" placeholder="{&#10;  &quot;title&quot;: &quot;...&quot;&#10;}"></textarea><div class="form-error" hidden></div></form>';
    openModal("Configurar " + lessonBuilderLabel(kind), body, { className: "modal-builder", footer: '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Cancelar</button><button class="button button-dark" type="button" data-action="import-lesson-builder"><i data-lucide="check"></i>Criar na Twenty</button></footer>' });
  }

  function refreshLessonBuilderPrompt() {
    var form = modalRoot.querySelector("#lessonBuilderForm");
    if (!form) return;
    var lesson = lessonById(form.dataset.lesson);
    if (!lesson) return;
    var ids = Array.from(form.querySelectorAll('[name="materialIds"]:checked')).map(function (input) { return input.value; });
    var includePast = !!form.querySelector('[name="includePast"]:checked');
    var target = form.querySelector("#lessonBuilderPrompt");
    if (target) target.value = buildLessonAIPrompt(lesson, form.dataset.kind, ids, includePast);
  }

  function normalizeAIQuizQuestion(question, materials) {
    var promptBlocks = attachSlideMaterialReferences(question.promptBlocks || question.prompt || "", materials);
    var optionBlocks = asArray(question.options).map(function (option) { return attachSlideMaterialReferences(option, materials); });
    if (!promptBlocks.length) throw new Error("Uma pergunta do quiz não tem enunciado.");
    if (optionBlocks.length < 2) throw new Error("Cada pergunta do quiz precisa de pelo menos duas opções.");
    var correctIndex = Number(question.correctIndex != null ? question.correctIndex : question.answerIndex);
    if (!Number.isInteger(correctIndex) || correctIndex < 0 || correctIndex >= optionBlocks.length) throw new Error("Uma pergunta tem correctIndex inválido.");
    return {
      id: uid("quizq"), mode: "multiple-choice", promptBlocks: promptBlocks, prompt: contentBlocksPlainText(promptBlocks),
      optionBlocks: optionBlocks, options: optionBlocks.map(contentBlocksPlainText), answerIndex: correctIndex,
      explanationBlocks: attachSlideMaterialReferences(question.explanationBlocks || question.explanation || "", materials),
      explanation: contentBlocksPlainText(question.explanationBlocks || question.explanation || ""),
      sourceSlides: asArray(question.sourceSlides).map(Number).filter(Boolean), difficulty: question.difficulty || "auto", images: []
    };
  }

  async function importLessonBuilder() {
    var form = modalRoot.querySelector("#lessonBuilderForm");
    if (!form) return;
    var errorBox = form.querySelector(".form-error");
    try {
      var lesson = lessonById(form.dataset.lesson);
      if (!lesson) throw new Error("Aula não encontrada.");
      var payload = parseJSONReply(form.querySelector("#lessonBuilderResponse").value);
      var kind = form.dataset.kind;
      var selectedMaterialIds = Array.from(form.querySelectorAll('[name="materialIds"]:checked')).map(function (input) { return input.value; });
      var selectedMaterials = lessonSlides(lesson.id).filter(function (material) { return selectedMaterialIds.indexOf(material.id) >= 0; });
      if (kind === "quiz") {
        if (configuredQuizForLesson(lesson.id)) throw new Error("Esta aula já tem um quiz configurado.");
        var questions = asArray(payload.questions).map(function (question) { return normalizeAIQuizQuestion(question, selectedMaterials); });
        if (!questions.length) throw new Error("O JSON não contém perguntas.");
        state.quizzes.push({ id: uid("quiz"), semesterId: lesson.semesterId, courseId: lesson.courseId, lessonId: lesson.id, title: String(payload.title || "Quiz da aula · " + lesson.title), questions: questions, configuredFromPrompt: true, lockedContent: true, createdAt: new Date().toISOString(), lastScore: null });
      } else if (kind === "homework") {
        if (homeworkForLesson(lesson.id)) throw new Error("Esta aula já tem um TPC configurado.");
        var instructions = attachSlideMaterialReferences(payload.instructionsBlocks || payload.contentBlocks || payload.instructions || "", selectedMaterials);
        if (!instructions.length) throw new Error("O JSON não contém instruções para o TPC.");
        state.tasks.push({ id: uid("task"), semesterId: lesson.semesterId, courseId: lesson.courseId, lessonId: lesson.id, title: String(payload.title || "TPC · " + lesson.title), type: "homework", dueDate: lesson.date || todayISO(), dueTime: "20:30", priority: "normal", done: false, estimatedMinutes: clamp(payload.estimatedMinutes || 30, 5, 240), contentBlocks: instructions, solutionBlocks: attachSlideMaterialReferences(payload.solutionBlocks || payload.solution || "", selectedMaterials), checklist: asArray(payload.checklist).map(String).filter(Boolean), extraTasks: asArray(payload.extraTasks).map(function (item, index) { if (typeof item === "string") return { title: item, blocks: [], optional: false }; item = item && typeof item === "object" ? item : {}; return { title: String(item.title || ("Tarefa extra " + (index + 1))), blocks: attachSlideMaterialReferences(item.blocks || item.contentBlocks || item.instructions || "", selectedMaterials), optional: !!item.optional }; }), configuredFromPrompt: true, lockedContent: true, createdAt: new Date().toISOString() });
      } else {
        var blocks = attachSlideMaterialReferences(payload.blocks || payload.notesBlocks || payload.contentBlocks || payload.notes || "", selectedMaterials);
        if (!blocks.length) throw new Error("O JSON não contém apontamentos.");
        var generatedHtml = sanitizeNotebookHTML(blocksToNotebookHTML(blocks));
        var existingHtml = sanitizeNotebookHTML(lesson.notesHtml || (lesson.notes ? '<p>' + nl2br(lesson.notes) + '</p>' : ''));
        var aiContinuation = '<div class="notebook-ai-continuation" data-ai-added-at="' + attr(new Date().toISOString()) + '"><p class="notebook-ai-label"><span>✦</span> Continuação adicionada com IA</p>' + generatedHtml + '</div>';
        lesson.notesHtml = sanitizeNotebookHTML((existingHtml ? existingHtml + '<p><br></p>' : '') + aiContinuation + '<p><br></p>');
        lesson.notes = contentBlocksPlainText([{ type: "text", text: lesson.notes || "" }]) + (lesson.notes ? "\n\n" : "") + contentBlocksPlainText(blocks);
        lesson.notesPaper = ["lined", "grid", "blank"].indexOf(payload.paper) >= 0 ? payload.paper : (lesson.notesPaper || "lined");
        lesson.notesUpdatedAt = new Date().toISOString();
      }
      ensureBeOnlineTasks();
      await save(true);
      closeModal();
      render();
      toast((kind === "quiz" ? "Quiz da aula" : kind === "homework" ? "TPC" : "Apontamentos") + " criado(s) e sincronizado(s).");
    } catch (error) {
      if (errorBox) { errorBox.hidden = false; errorBox.textContent = error.message; }
      else toast(error.message, "error");
    }
  }

  function viewLessonQuiz(lessonId) {
    var quiz = configuredQuizForLesson(lessonId);
    var lesson = lessonById(lessonId);
    if (!quiz || !lesson) return;
    var body = '<div class="view-only-banner"><i data-lucide="lock-keyhole"></i><span><strong>Conteúdo fechado</strong><small>Há apenas um quiz por aula. Depois de concluído, fica disponível apenas para consulta.</small></span></div><div class="quiz-preview-list">' + asArray(quiz.questions).map(function (question, index) {
      var options = asArray(question.optionBlocks).length ? question.optionBlocks : asArray(question.options);
      return '<article class="quiz-preview-card"><p class="card-label">Pergunta ' + (index + 1) + '</p>' + renderContentBlocks(question.promptBlocks || question.prompt) + '<div class="quiz-preview-options">' + options.map(function (option, optionIndex) { return '<div class="quiz-preview-option ' + (optionIndex === Number(question.answerIndex) ? "is-correct" : "") + '"><span>' + String.fromCharCode(65 + optionIndex) + '</span>' + renderQuizOptionContent(option) + '</div>'; }).join("") + '</div><details><summary>Ver explicação</summary>' + renderContentBlocks(question.explanationBlocks || question.explanation) + '</details></article>';
    }).join("") + '</div>';
    var completed = lessonIsBeOnline(lesson);
    var footer = '<footer class="modal-foot"><button class="button' + (completed ? ' button-dark' : '') + '" type="button" data-action="close-modal">Fechar</button>' + (completed ? '' : '<button class="button button-dark" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="play"></i>Fazer quiz</button>') + '</footer>';
    openModal("Quiz da aula", body, { className: "modal-wide", footer: footer });
  }

  function viewLessonHomework(taskId) {
    var task = state.tasks.find(function (item) { return item.id === taskId; });
    if (!task) return;
    var lesson = lessonById(task.lessonId);
    var checklist = asArray(task.checklist).length ? '<div class="homework-checklist"><p class="card-label">Checklist</p>' + task.checklist.map(function (item) { return '<span><i data-lucide="square"></i>' + esc(item) + '</span>'; }).join("") + '</div>' : '';
    var extras = asArray(task.extraTasks).length ? '<div class="homework-extra-list"><p class="card-label">Tarefas adicionais</p>' + task.extraTasks.map(function (item) { return '<article><div><i data-lucide="' + (item.optional ? 'circle-dashed' : 'circle-check-big') + '"></i><strong>' + esc(item.title) + '</strong>' + (item.optional ? '<span class="badge">Opcional</span>' : '') + '</div>' + renderContentBlocks(item.blocks) + '</article>'; }).join("") + '</div>' : '';
    var solution = asArray(task.solutionBlocks).length ? '<details class="homework-solution"><summary>Ver solução / critérios de correção</summary>' + renderContentBlocks(task.solutionBlocks) + '</details>' : '';
    var body = '<div class="view-only-banner"><i data-lucide="lock-keyhole"></i><span><strong>TPC da aula</strong><small>O conteúdo fica em modo de visualização depois de criado.</small></span></div><div class="question-meta"><span class="badge badge-violet">' + esc(task.estimatedMinutes || 30) + ' min</span><span class="badge">' + esc(relativeDate(task.dueDate)) + '</span></div><h3 style="margin:16px 0 8px">' + esc(task.title) + '</h3>' + renderContentBlocks(task.contentBlocks) + checklist + extras + solution;
    var footer = '<footer class="modal-foot"><button class="button' + (task.done ? ' button-dark' : '') + '" type="button" data-action="close-modal">Fechar</button>' + (task.done ? '' : '<button class="button button-dark" type="button" data-action="start-homework-session" data-task="' + attr(task.id) + '"><i data-lucide="play"></i>Fazer TPC</button>') + '</footer>';
    openModal("TPC", body, { className: "modal-wide", footer: footer });
  }

  function notebookPaperLabel(value) {
    return value === "grid" ? "Quadriculado" : value === "blank" ? "Branco" : "Pautado";
  }

  function renderNotebookPage(lesson, compact) {
    var html = lesson.notesHtml || (lesson.notes ? '<p>' + nl2br(lesson.notes) + '</p>' : '');
    if (!html && asArray(lesson.aiNotes).length) html = '<div class="lesson-ai-notes">' + lesson.aiNotes.map(renderAINote).join("") + '</div>';
    if (!html) html = '<p class="notebook-placeholder">Ainda não há apontamentos nesta aula.</p>';
    return '<div class="notebook-page paper-' + attr(lesson.notesPaper || "lined") + ' ' + (compact ? "is-compact" : "") + '"><div class="notebook-content">' + html + '</div></div>';
  }

  function openNotebookEditor(lessonId) {
    var lesson = lessonById(lessonId);
    if (!lesson) return;
    var html = lesson.notesHtml || (lesson.notes ? '<p>' + nl2br(lesson.notes) + '</p>' : '<p><br></p>');
    var toolbar = '<div class="notebook-toolbar" role="toolbar" aria-label="Formatar apontamentos"><button type="button" data-action="notebook-command" data-command="bold" title="Negrito"><i data-lucide="bold"></i></button><button type="button" data-action="notebook-command" data-command="italic" title="Itálico"><i data-lucide="italic"></i></button><button type="button" data-action="notebook-command" data-command="underline" title="Sublinhar"><i data-lucide="underline"></i></button><button type="button" data-action="notebook-command" data-command="strikeThrough" title="Rasurar"><i data-lucide="strikethrough"></i></button><span></span><button type="button" data-action="notebook-block" data-block="h2">H2</button><button type="button" data-action="notebook-block" data-block="h3">H3</button><button type="button" data-action="notebook-command" data-command="insertUnorderedList" title="Lista"><i data-lucide="list"></i></button><button type="button" data-action="notebook-command" data-command="insertOrderedList" title="Lista numerada"><i data-lucide="list-ordered"></i></button><button type="button" data-action="notebook-command" data-command="formatBlock" data-value="blockquote" title="Citação"><i data-lucide="quote"></i></button><span></span><button type="button" data-action="notebook-add-image" title="Adicionar imagem ou sticker"><i data-lucide="image-plus"></i></button><button type="button" data-action="configure-lesson-content" data-kind="notes" data-lesson="' + attr(lesson.id) + '" title="Acrescentar com IA"><i data-lucide="sparkles"></i></button></div>';
    var body = '<form id="notebookForm" data-lesson="' + attr(lesson.id) + '"><div class="notebook-settings"><label>Folha<select id="notebookPaper"><option value="lined" ' + (lesson.notesPaper === "lined" ? "selected" : "") + '>Pautado</option><option value="grid" ' + (lesson.notesPaper === "grid" ? "selected" : "") + '>Quadriculado</option><option value="blank" ' + (lesson.notesPaper === "blank" ? "selected" : "") + '>Branco</option></select></label><label>Fonte<select disabled><option>Fonte da Twenty</option></select></label></div>' + toolbar + '<input id="notebookImageInput" type="file" accept="image/*" multiple hidden><div id="notebookEditor" class="notebook-page notebook-editor paper-' + attr(lesson.notesPaper || "lined") + '" contenteditable="true" spellcheck="true">' + html + '</div><div class="form-note"><strong>Colar de sites:</strong> usa “Copiar imagem” ou copia conteúdo de uma página e cola aqui. Texto e imagens são aceites diretamente no caderno. As imagens funcionam como stickers: arrasta para mudar de lugar e usa os controlos para redimensionar, alinhar ou apagar. O conteúdo criado pela IA é sempre acrescentado no fim e nunca apaga o que escreveste.</div></form>';
    openModal("Apontamentos da aula", body, { className: "modal-notebook", footer: '<footer class="modal-foot"><button class="button" type="button" data-action="close-modal">Sair e guardar</button><button class="button button-dark" type="button" data-action="save-notebook"><i data-lucide="check"></i>Guardar apontamentos</button></footer>' });
  }

  async function saveNotebook(options) {
    options = options || {};
    var form = modalRoot.querySelector("#notebookForm");
    if (!form) return false;
    var lesson = lessonById(form.dataset.lesson);
    if (!lesson) return false;
    var editor = form.querySelector("#notebookEditor");
    var paper = form.querySelector("#notebookPaper");
    await materializeNotebookImages(editor);
    var editorClone = editor ? editor.cloneNode(true) : document.createElement("div");
    Array.from(editorClone.querySelectorAll(".notebook-sticker-controls")).forEach(function (controls) { controls.remove(); });
    Array.from(editorClone.querySelectorAll('img[data-remote-path], img[data-local-image-id]')).forEach(function (image) { image.removeAttribute("src"); image.removeAttribute("data-hydrated"); image.removeAttribute("data-remote-hydrated"); });
    lesson.notesHtml = sanitizeNotebookHTML(editorClone.innerHTML || "");
    lesson.notes = contentBlocksPlainText([{ type: "text", text: editor ? editor.innerText : "" }]);
    lesson.notesPaper = paper && ["lined", "grid", "blank"].indexOf(paper.value) >= 0 ? paper.value : "lined";
    lesson.notesUpdatedAt = new Date().toISOString();
    await save(true);
    if (options.close !== false) closeModal();
    if (options.render !== false) render();
    if (!options.silent) toast("Apontamentos guardados.");
    return true;
  }

  async function closeModalSavingNotebook() {
    if (modalRoot.querySelector("#notebookForm")) {
      try { await saveNotebook({ close: false, render: false, silent: true }); }
      catch (error) { toast(error.message || "Não foi possível guardar os apontamentos.", "error"); return false; }
    }
    closeModal();
    render();
    return true;
  }

  function renderCourseNotebook(course, archived) {
    var lessons = state.lessons.filter(function (lesson) { return lesson.courseId === course.id; }).sort(function (a, b) { return String(a.date || "").localeCompare(String(b.date || "")) || String(a.start || "").localeCompare(String(b.start || "")); });
    var pages = lessons.map(function (lesson, index) {
      return '<article class="course-notebook-entry"><header><div><p class="card-label">Aula ' + (index + 1) + ' · ' + esc(formatDate(lesson.date)) + '</p><h3>' + esc(lesson.title) + '</h3><small>' + esc([lesson.start, lesson.room].filter(Boolean).join(" · ")) + '</small></div><div class="list-actions"><span class="badge">' + esc(notebookPaperLabel(lesson.notesPaper)) + '</span><button class="button button-small" type="button" data-route="lesson" data-id="' + attr(lesson.id) + '"><i data-lucide="arrow-right"></i>Abrir aula</button></div></header>' + renderNotebookPage(lesson, true) + '</article>';
    }).join("");
    return '<div class="page-head"><div><h2>Caderno da cadeira</h2><p>Todos os apontamentos, aula a aula e pela ordem em que aconteceram.</p></div></div><div class="course-notebook">' + (pages || emptyState("notebook", "Caderno vazio", "Escreve apontamentos dentro de uma aula para os veres aqui.", null)) + '</div>';
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
    var quiz = configuredQuizForLesson(lesson.id);
    var homework = homeworkForLesson(lesson.id);
    var onlineComplete = lessonIsBeOnline(lesson);
    var lessonEnded = lessonHasEnded(lesson);
    var currentMaterials = materials.filter(function (item) { return !semester || !item.academicYear || item.academicYear === semester.academicYear; });
    var oldMaterials = materials.filter(function (item) { return semester && item.academicYear && item.academicYear !== semester.academicYear; });
    var materialsHtml = materials.length ? '<div class="material-grid">' + currentMaterials.concat(oldMaterials).map(function (item) { return renderMaterialCard(item, course || { semesterId: null }, archived); }).join("") + '</div>' : emptyState("file-up", "Ainda sem material", "Carrega os slides ou PDF desta aula. O ficheiro ficará sincronizado.", "add-material", "Carregar material");
    var questionsHtml = questions.length ? questions.map(function (item) { return renderQuestionCard(item, archived); }).join("") : emptyState("message-circle-question", "Sem perguntas anteriores", "Associa perguntas de testes antigos a esta aula.", "add-question", "Associar pergunta");
    var quizHtml = quiz
      ? '<div class="lesson-configured-card"><span class="list-icon yellow"><i data-lucide="check-check"></i></span><div><strong>' + esc(quiz.title) + '</strong><small>' + asArray(quiz.questions).length + ' perguntas' + (quiz.lastScore != null ? ' · resultado ' + quiz.lastScore + '%' : '') + ' · conteúdo fechado</small></div><div class="list-actions"><button class="button ' + (onlineComplete ? 'button-dark' : 'button') + ' button-small" type="button" data-action="view-lesson-quiz" data-lesson="' + attr(lesson.id) + '"><i data-lucide="eye"></i>Ver</button>' + (onlineComplete ? '' : '<button class="button button-dark button-small" type="button" data-action="start-quiz" data-id="' + attr(quiz.id) + '"><i data-lucide="play"></i>Fazer</button>') + '</div></div>'
      : '<div class="lesson-unconfigured"><span class="metric-icon"><i data-lucide="check-check"></i></span><div><strong>Quiz ainda não configurado</strong><small>Escolhe os slides, copia o prompt para a IA e cola o JSON devolvido.</small></div><button class="button button-dark button-small" type="button" data-action="configure-lesson-content" data-kind="quiz" data-lesson="' + attr(lesson.id) + '"><i data-lucide="wand-sparkles"></i>✅ Configurar quiz</button></div>';
    var homeworkHtml = homework
      ? '<div class="lesson-configured-card"><span class="list-icon orange"><i data-lucide="notebook-pen"></i></span><div><strong>' + esc(homework.title) + '</strong><small>' + (homework.estimatedMinutes || 30) + ' min · ' + (homework.done ? 'concluído' : 'por fazer') + ' · conteúdo fechado</small></div><div class="list-actions"><button class="button ' + (homework.done ? 'button-dark' : 'button') + ' button-small" type="button" data-action="view-lesson-homework" data-id="' + attr(homework.id) + '"><i data-lucide="eye"></i>Ver</button>' + (homework.done ? '' : '<button class="button button-dark button-small" type="button" data-action="start-homework-session" data-task="' + attr(homework.id) + '"><i data-lucide="play"></i>Fazer</button>') + '</div></div>'
      : '<div class="lesson-unconfigured"><span class="metric-icon"><i data-lucide="notebook-pen"></i></span><div><strong>TPC ainda não configurado</strong><small>O TPC é separado do quiz e pensado para fazer em casa.</small></div><button class="button button-dark button-small" type="button" data-action="configure-lesson-content" data-kind="homework" data-lesson="' + attr(lesson.id) + '"><i data-lucide="wand-sparkles"></i>✅ Configurar TPC</button></div>';
    var quizStatusCard = onlineComplete ? '' : '<article class="card span-12 beonline-lesson-card"><div class="beonline-lesson-copy"><span class="badge ' + (quiz && lessonEnded ? 'badge-danger' : 'badge-violet') + '">' + (quiz && lessonEnded ? 'Quiz pendente' : quiz ? 'Quiz preparado' : 'Sem quiz') + '</span><h3>Quiz da aula</h3><p>' + esc(quiz && lessonEnded ? 'O quiz está pronto para fazer.' : quiz ? 'Quando a aula estiver a terminar, a Home pode sugerir este quiz.' : 'O quiz só aparece na Home depois de o configurares.') + '</p></div></article>';
    return '<div class="page-head"><div><button class="button button-ghost button-small" type="button" data-route="course" data-id="' + attr(lesson.courseId) + '"><i data-lucide="arrow-left"></i>' + esc(course ? course.code || course.name : "Cadeira") + '</button><h2 style="margin-top:11px">' + esc(lesson.title) + '</h2><p>' + formatLongDate(lesson.date) + (lesson.start ? ' · ' + esc(lesson.start) + '–' + esc(lesson.end || '') : '') + ' · ' + esc(lessonTypeLabel(lesson.type)) + (lesson.room ? ' · ' + esc(lesson.room) : '') + '</p></div><div class="page-actions">' + (!archived ? '<button class="button" type="button" data-action="edit-lesson" data-id="' + attr(lesson.id) + '"><i data-lucide="pencil"></i>Editar aula</button><button class="button ' + (lesson.mastered ? 'button-yellow' : 'button-dark') + '" type="button" data-action="toggle-mastery" data-id="' + attr(lesson.id) + '"><i data-lucide="badge-check"></i>' + (lesson.mastered ? 'Dominada' : 'Marcar dominada') + '</button>' : '') + '</div></div><div class="bento-grid"><article class="card course-hero span-12" style="--course-color:' + safeColor(course && course.color) + ';min-height:220px"><div class="course-hero-copy"><span class="badge badge-dark">' + esc(lesson.type || 'Aula') + '</span><h2>' + esc(lesson.title) + '</h2><p>' + esc(lesson.topics || 'Adiciona os tópicos dados nesta aula.') + '</p></div><div class="course-score"><strong>' + (lesson.mastered ? '✓' : questions.length) + '</strong><span>' + (lesson.mastered ? 'matéria dominada' : 'perguntas antigas') + '</span></div></article>' + quizStatusCard + '<article class="card span-12"><div class="card-title-row"><div><h3>Slides e PDFs</h3><p class="card-subtitle">Os ficheiros são enviados para o Git e ficam disponíveis nos teus dispositivos.</p></div>' + (!archived ? '<button class="button button-small" type="button" data-action="add-material" data-course="' + attr(lesson.courseId) + '" data-lesson="' + attr(lesson.id) + '"><i data-lucide="file-up"></i>Carregar</button>' : '') + '</div><div style="margin-top:15px">' + materialsHtml + '</div></article><article class="card span-7"><div class="card-title-row"><div><h3>Perguntas de anos anteriores</h3></div><div class="list-actions">' + (!archived ? '<button class="button button-small" type="button" data-action="add-question" data-course="' + attr(lesson.courseId) + '" data-lesson="' + attr(lesson.id) + '"><i data-lucide="plus"></i>Pergunta</button>' : '') + '</div></div><div style="margin-top:15px">' + questionsHtml + '</div></article><article class="card span-5 lesson-config-card"><div class="card-title-row"><div><h3>✅ Quiz da aula</h3><p class="card-subtitle">Um por aula. Depois de criado, fica em visualização.</p></div></div>' + quizHtml + '</article><article class="card span-12 lesson-config-card"><div class="card-title-row"><div><h3>✅ TPC da aula</h3><p class="card-subtitle">Aplicação e prática para fazer em casa, separada do quiz.</p></div></div>' + homeworkHtml + '</article><article class="card span-12 notebook-card"><div class="card-title-row"><div><h3>Apontamentos</h3><p class="card-subtitle">Escreve como num caderno ou gera conteúdo com um prompt estruturado.</p></div><div class="list-actions">' + (!archived ? '<button class="button button-small" type="button" data-action="configure-lesson-content" data-kind="notes" data-lesson="' + attr(lesson.id) + '"><i data-lucide="wand-sparkles"></i>Gerar por prompt</button><button class="button button-dark button-small" type="button" data-action="open-notebook-editor" data-lesson="' + attr(lesson.id) + '"><i data-lucide="pencil"></i>Escrever</button>' : '') + '</div></div><div style="margin-top:15px">' + renderNotebookPage(lesson, false) + '</div><div class="notebook-meta"><span><i data-lucide="notebook"></i>' + esc(notebookPaperLabel(lesson.notesPaper)) + '</span><button class="button button-small" type="button" data-action="course-tab" data-id="' + attr(lesson.courseId) + '" data-tab="notebook"><i data-lucide="library-big"></i>Ver caderno da cadeira</button></div></article></div>';
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
    return '<div class="bento-grid" style="margin-top:15px"><article class="card span-6"><div class="card-title-row"><div><h3>Tarefas e revisões</h3></div><div class="list-actions"><button class="button button-small" type="button" data-action="start-homework-session"><i data-lucide="play"></i>Fazer TPCs</button><button class="row-button" type="button" data-action="add-task" aria-label="Adicionar tarefa"><i data-lucide="plus"></i></button></div></div><div class="list-stack">' + taskHtml + '</div></article><article class="card span-6"><div class="card-title-row"><div><h3>Avaliações e eventos</h3></div><div class="list-actions"><button class="row-button" type="button" data-action="add-event" aria-label="Adicionar evento"><i data-lucide="party-popper"></i></button><button class="row-button" type="button" data-action="add-assessment" aria-label="Adicionar avaliação"><i data-lucide="file-plus-2"></i></button></div></div><div class="list-stack">' + agendaHtml + "</div></article></div>";
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


  function isHomeworkTask(task) {
    return !!task && (task.type === "homework" || task.type === "tpc");
  }

  function homeworkSessionTasks() {
    var today = todayISO(activeHomeNow());
    var all = semesterItems("tasks").filter(isHomeworkTask).slice().sort(function (a, b) {
      return Number(a.done) - Number(b.done) || String(a.dueDate || "9999").localeCompare(String(b.dueDate || "9999")) || String((courseById(a.courseId) || {}).name || "").localeCompare(String((courseById(b.courseId) || {}).name || ""));
    });
    var relevant = all.filter(function (task) {
      return !task.done || String(task.completedAt || "").slice(0, 10) === today;
    });
    return relevant;
  }

  function homeworkElapsedSeconds() {
    if (!homeworkSessionRuntime || !homeworkSessionRuntime.taskId) return 0;
    var base = Number(homeworkSessionRuntime.elapsedByTask[homeworkSessionRuntime.taskId] || 0);
    if (homeworkSessionRuntime.running && homeworkSessionRuntime.startedAt) {
      base += Math.max(0, Math.floor((Date.now() - homeworkSessionRuntime.startedAt) / 1000));
    }
    return base;
  }

  function storeHomeworkElapsed() {
    if (!homeworkSessionRuntime || !homeworkSessionRuntime.taskId) return;
    homeworkSessionRuntime.elapsedByTask[homeworkSessionRuntime.taskId] = homeworkElapsedSeconds();
    homeworkSessionRuntime.startedAt = homeworkSessionRuntime.running ? Date.now() : null;
  }

  function setHomeworkCurrent(taskId, autoStart) {
    var tasks = homeworkSessionTasks();
    var task = tasks.find(function (item) { return item.id === taskId && !item.done; }) || tasks.find(function (item) { return !item.done; }) || null;
    if (!homeworkSessionRuntime) homeworkSessionRuntime = { taskId: null, elapsedByTask: {}, startedAt: null, running: false, complete: false };
    storeHomeworkElapsed();
    homeworkSessionRuntime.taskId = task ? task.id : null;
    homeworkSessionRuntime.complete = !task;
    homeworkSessionRuntime.running = !!task && autoStart !== false;
    homeworkSessionRuntime.startedAt = homeworkSessionRuntime.running ? Date.now() : null;
    return task;
  }

  function startHomeworkSession(taskId) {
    if (!homeworkSessionTasks().some(function (task) { return !task.done; })) {
      homeworkSessionRuntime = { taskId: null, elapsedByTask: {}, startedAt: null, running: false, complete: true };
    } else {
      homeworkSessionRuntime = { taskId: null, elapsedByTask: {}, startedAt: null, running: false, complete: false };
      setHomeworkCurrent(taskId || "", true);
    }
    setRoute("homework");
  }

  function formatHomeworkClock(seconds) {
    seconds = Math.max(0, Math.floor(Number(seconds) || 0));
    var hours = Math.floor(seconds / 3600);
    var minutes = Math.floor(seconds % 3600 / 60);
    var rest = seconds % 60;
    return (hours ? String(hours).padStart(2, "0") + ":" : "") + String(minutes).padStart(2, "0") + ":" + String(rest).padStart(2, "0");
  }

  function updateHomeworkClock() {
    if (route.name !== "homework") return;
    var timer = document.getElementById("homeworkSessionTimer");
    var ring = document.querySelector(".homework-timer-ring");
    var elapsed = homeworkElapsedSeconds();
    if (timer) timer.textContent = formatHomeworkClock(elapsed);
    if (ring) {
      var duration = Math.max(1, Number(ring.dataset.duration || 30) * 60);
      var progress = Math.min(100, Math.round(elapsed / duration * 100));
      ring.style.setProperty("--homework-progress", progress + "%");
      var caption = ring.querySelector("small");
      if (caption) caption.textContent = elapsed >= duration ? "tempo estimado atingido" : Math.max(0, Math.ceil((duration - elapsed) / 60)) + " min estimados restantes";
    }
    homeworkClockTimer = setTimeout(updateHomeworkClock, 1000);
  }

  function scheduleHomeworkClock() {
    if (homeworkClockTimer) clearTimeout(homeworkClockTimer);
    homeworkClockTimer = setTimeout(updateHomeworkClock, 120);
  }

  function renderHomeworkChecklist(tasks, activeId) {
    return tasks.map(function (task, index) {
      var course = courseById(task.courseId);
      var status = task.done ? '<i data-lucide="check"></i>' : index + 1;
      return '<button class="homework-queue-item ' + (task.done ? 'is-done' : '') + ' ' + (task.id === activeId ? 'is-active' : '') + '" type="button" data-action="homework-select" data-id="' + attr(task.id) + '" ' + (task.done ? 'disabled' : '') + '><span class="homework-queue-check">' + status + '</span><span><strong>' + esc(course ? course.name : 'TPC') + '</strong><small>' + esc(task.title) + ' · ' + esc(task.estimatedMinutes || 30) + ' min</small></span></button>';
    }).join("");
  }

  function renderHomeworkTaskBody(task) {
    var checklist = asArray(task.checklist).length ? '<section class="homework-focus-section"><p class="card-label">Passos</p><div class="homework-focus-checklist">' + task.checklist.map(function (item) { return '<span><i data-lucide="square"></i>' + esc(item) + '</span>'; }).join("") + '</div></section>' : '';
    var extras = asArray(task.extraTasks).length ? '<section class="homework-focus-section"><p class="card-label">Tarefas adicionais</p><div class="homework-focus-extras">' + task.extraTasks.map(function (item) { return '<article><div><i data-lucide="' + (item.optional ? 'circle-dashed' : 'circle-check-big') + '"></i><strong>' + esc(item.title) + '</strong>' + (item.optional ? '<span class="badge">Opcional</span>' : '') + '</div>' + renderContentBlocks(item.blocks) + '</article>'; }).join("") + '</div></section>' : '';
    return '<section class="homework-focus-section homework-main-task"><p class="card-label">Tarefa</p>' + (asArray(task.contentBlocks).length ? renderContentBlocks(task.contentBlocks) : '<p>' + esc(task.notes || task.title) + '</p>') + '</section>' + checklist + extras;
  }

  function renderHomeworkComplete(tasks) {
    return '<section class="homework-complete-screen"><span class="homework-complete-icon"><i data-lucide="check"></i></span><p class="card-label">After-school complete</p><h2>TPCs feitos!</h2><p>Fechaste ' + tasks.length + ' tarefa' + (tasks.length === 1 ? '' : 's') + '. Já podes descansar sem aquela sensação de “estou-me a esquecer de alguma coisa”.</p><div class="homework-complete-actions"><button class="button button-dark" type="button" data-route="home"><i data-lucide="house"></i>Voltar à Home</button><button class="button" type="button" data-action="view-report-card"><i data-lucide="award"></i>Ver o dia</button></div></section>';
  }

  function renderHomeworkSession() {
    setHeader("TPCs", "Sessão depois das aulas");
    var tasks = homeworkSessionTasks();
    if (!tasks.length) {
      return '<div class="page-head"><div><h2>Não há TPCs para fazer.</h2><p>Quando adicionares um TPC numa aula, ele aparece aqui numa sessão guiada.</p></div><div class="page-actions"><button class="button" type="button" data-route="home"><i data-lucide="arrow-left"></i>Voltar</button></div></div>' + emptyState("party-popper", "Tudo livre", "O teu fim de tarde está livre. Aproveita.", "add-task", "Adicionar TPC");
    }
    if (!homeworkSessionRuntime) {
      homeworkSessionRuntime = { taskId: null, elapsedByTask: {}, startedAt: null, running: false, complete: false };
      setHomeworkCurrent("", true);
    }
    var pending = tasks.filter(function (task) { return !task.done; });
    if (!pending.length || homeworkSessionRuntime.complete) return renderHomeworkComplete(tasks);
    var current = pending.find(function (task) { return task.id === homeworkSessionRuntime.taskId; }) || setHomeworkCurrent(pending[0].id, true);
    var course = courseById(current.courseId);
    var position = pending.findIndex(function (task) { return task.id === current.id; }) + 1;
    var queue = renderHomeworkChecklist(tasks, current.id);
    var elapsed = homeworkElapsedSeconds();
    var body = renderHomeworkTaskBody(current);
    return '<div class="page-head homework-session-head"><div><button class="button button-ghost button-small" type="button" data-route="home"><i data-lucide="arrow-left"></i>Home</button><h2 style="margin-top:11px">TPCs depois das aulas</h2><p>Uma disciplina de cada vez. A Twenty trata da fila; tu só tens de fazer a tarefa em frente.</p></div><div class="page-actions"><span class="badge badge-violet">' + (tasks.length - pending.length) + '/' + tasks.length + ' feitos</span></div></div><div class="homework-session-layout"><aside class="homework-queue card"><div class="card-title-row"><div><p class="card-label">Checklist</p><h3>Disciplinas</h3></div><span class="badge">' + pending.length + ' restantes</span></div><div class="homework-queue-list">' + queue + '</div></aside><main class="homework-focus card"><div class="homework-focus-top"><div><p class="card-label">TPC em curso · ' + position + ' de ' + pending.length + '</p><span class="badge badge-dark">' + esc(course ? course.name : 'TPC') + '</span><h2>' + esc(current.title) + '</h2></div><div class="homework-timer-ring" data-duration="' + attr(current.estimatedMinutes || 30) + '" style="--homework-progress:0%"><strong id="homeworkSessionTimer">' + formatHomeworkClock(elapsed) + '</strong><small>' + esc(current.estimatedMinutes || 30) + ' min estimados</small></div></div><div class="homework-focus-content">' + body + '</div><div class="homework-focus-actions">' + (homeworkSessionRuntime.running ? '<button class="button" type="button" data-action="homework-pause"><i data-lucide="pause"></i>Pausar</button>' : '<button class="button" type="button" data-action="homework-resume"><i data-lucide="play"></i>Continuar</button>') + (current.lessonId ? '<button class="button" type="button" data-route="lesson" data-id="' + attr(current.lessonId) + '"><i data-lucide="book-open"></i>Abrir aula</button>' : '') + '<button class="button button-dark homework-finish-button" type="button" data-action="homework-finish"><i data-lucide="check"></i>Terminar e avançar</button></div></main></div>';
  }

  async function finishHomeworkTask() {
    if (!homeworkSessionRuntime || !homeworkSessionRuntime.taskId) return;
    storeHomeworkElapsed();
    var task = state.tasks.find(function (item) { return item.id === homeworkSessionRuntime.taskId; });
    if (!task) return;
    task.done = true;
    task.completedOnce = true;
    state.meta.completedHomeworkIds = Array.from(new Set(asArray(state.meta.completedHomeworkIds).concat(String(task.id))));
    task.completedAt = new Date().toISOString();
    task.actualSeconds = Number(homeworkSessionRuntime.elapsedByTask[task.id] || 0);
    var next = homeworkSessionTasks().find(function (item) { return !item.done; });
    if (next) {
      setHomeworkCurrent(next.id, true);
      await save(true);
      render();
      toast("TPC concluído. A seguir: " + next.title);
      return;
    }
    homeworkSessionRuntime.taskId = null;
    homeworkSessionRuntime.running = false;
    homeworkSessionRuntime.complete = true;
    await save(true);
    render();
    setTimeout(launchHomeworkConfetti, 80);
  }

  function launchHomeworkConfetti() {
    var layer = document.createElement("div");
    layer.className = "homework-confetti";
    for (var index = 0; index < 72; index += 1) {
      var piece = document.createElement("i");
      piece.style.setProperty("--x", Math.round(Math.random() * 100) + "vw");
      piece.style.setProperty("--delay", (Math.random() * .7).toFixed(2) + "s");
      piece.style.setProperty("--spin", Math.round(Math.random() * 720 - 360) + "deg");
      piece.style.setProperty("--drift", Math.round(Math.random() * 180 - 90) + "px");
      piece.dataset.tone = String(index % 6);
      layer.appendChild(piece);
    }
    document.body.appendChild(layer);
    setTimeout(function () { layer.remove(); }, 3800);
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

  function canteenMealType(value) {
    var name = typeof value === "string" ? value : String(value && value.name || "");
    return name.toLowerCase().indexOf("jantar") >= 0 ? "dinner" : "lunch";
  }

  function canteenMealLabel(type) {
    return type === "dinner" ? "Jantar" : "Almoço";
  }

  function canteenMealVerb(type) {
    return type === "dinner" ? "jantado" : "almoçado";
  }

  function canteenVisitId(date, mealType) {
    return "canteen_" + String(date || "").replace(/[^0-9]/g, "") + "_" + (mealType || "lunch");
  }

  function canteenVisitFor(date, mealType) {
    var id = canteenVisitId(date, mealType);
    return asArray(state && state.canteenVisits).find(function (visit) {
      return visit.id === id || (visit.date === date && visit.mealType === mealType);
    }) || null;
  }

  function canteenDayForDate(date) {
    return canteenMenu && asArray(canteenMenu.days).find(function (day) { return day.date === date; }) || null;
  }

  function canteenMealForDate(date, mealType) {
    var day = canteenDayForDate(date);
    if (!day) return null;
    return asArray(day.meals).find(function (meal) { return canteenMealType(meal) === mealType; }) || null;
  }

  function canteenSelectionKey(date, mealType) {
    return String(date || "") + "|" + String(mealType || "lunch");
  }

  function canteenSelectionFor(date, mealType) {
    var key = canteenSelectionKey(date, mealType);
    if (!canteenSelections[key]) canteenSelections[key] = { dishIndex: null, dessertId: "" };
    return canteenSelections[key];
  }

  function loadCanteenAICache() {
    try {
      var parsed = JSON.parse(localStorage.getItem(CANTEEN_AI_CACHE_KEY) || "{}");
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch (error) { return {}; }
  }

  function saveCanteenAICache(cache) {
    try { localStorage.setItem(CANTEEN_AI_CACHE_KEY, JSON.stringify(cache || {})); } catch (error) {}
  }

  function destroyCanteenAISession() {
    try {
      if (canteenAIAbortController) canteenAIAbortController.abort();
    } catch (error) {}
    canteenAIAbortController = null;
    try {
      if (canteenAISession && typeof canteenAISession.destroy === "function") canteenAISession.destroy();
    } catch (error) {}
    canteenAISession = null;
    canteenAISessionProviderKind = "";
    canteenAISessionUses = 0;
  }

  function clearCanteenAICache() {
    try { localStorage.removeItem(CANTEEN_AI_CACHE_KEY); } catch (error) {}
    destroyCanteenAISession();
    canteenAIState = { key: "", status: "idle", availability: "unknown", source: "rules", data: null, error: "", progress: null, streamText: "" };
  }

  function loadCachedCanteenWeather() {
    try {
      var parsed = JSON.parse(localStorage.getItem(CANTEEN_WEATHER_CACHE_KEY) || "null");
      if (!parsed || !parsed.fetchedAt || !parsed.current) return null;
      if (Date.now() - Number(parsed.fetchedAt) > 45 * 60 * 1000) return null;
      return parsed;
    } catch (error) { return null; }
  }

  function saveCachedCanteenWeather(weather) {
    try { localStorage.setItem(CANTEEN_WEATHER_CACHE_KEY, JSON.stringify(weather)); } catch (error) {}
  }

  function canteenWeatherContext(weather) {
    var current = weather && weather.current;
    if (!current) return { kind: "unknown", label: "", prompt: "Sem contexto meteorológico disponível." };
    var temperature = Math.round(Number(current.temperature_2m));
    var apparent = Math.round(Number(current.apparent_temperature));
    var precipitation = Number(current.precipitation) || Number(current.rain) || 0;
    var code = Number(current.weather_code);
    var rainy = precipitation > 0 || [51, 53, 55, 56, 57, 61, 63, 65, 66, 67, 80, 81, 82, 95, 96, 99].indexOf(code) >= 0;
    var snowy = [71, 73, 75, 77, 85, 86].indexOf(code) >= 0;
    var foggy = [45, 48].indexOf(code) >= 0;
    var cloudy = [2, 3].indexOf(code) >= 0;
    var kind = rainy ? "rainy" : snowy ? "snowy" : foggy ? "foggy" : temperature >= 27 ? "hot" : temperature <= 13 ? "cold" : cloudy ? "cloudy" : "mild";
    var mood = {
      rainy: "Está a chover no Campus; o ambiente pede uma escolha mais cozy e reconfortante.",
      snowy: "O tempo está muito frio; o ambiente pede uma refeição quente e reconfortante.",
      foggy: "A manhã está enevoada e calma; combina com uma pausa acolhedora.",
      hot: "Está um dia quente no Campus; favorece uma escolha mais leve e fresca.",
      cold: "Está fresco no Campus; combina com uma escolha quente e mais composta.",
      cloudy: "O céu está nublado; o menu pode ganhar um mood mais cozy.",
      mild: "O tempo está ameno no Campus; escolhe pela vibe da ementa."
    }[kind];
    var labelMap = { rainy: "Chuva", snowy: "Muito frio", foggy: "Nevoeiro", hot: "Calor", cold: "Fresco", cloudy: "Nublado", mild: "Ameno" };
    return {
      kind: kind,
      temperature: isFinite(temperature) ? temperature : null,
      apparentTemperature: isFinite(apparent) ? apparent : null,
      label: labelMap[kind] + (isFinite(temperature) ? " · " + temperature + " °C" : ""),
      prompt: mood + (isFinite(apparent) && apparent !== temperature ? " Sensação térmica de " + apparent + " °C." : "")
    };
  }

  async function ensureCanteenWeather(date) {
    var today = canteenPortugalParts(new Date()).iso;
    if (!date || date !== today) return null;
    if (canteenWeatherState && Date.now() - Number(canteenWeatherState.fetchedAt || 0) < 30 * 60 * 1000) return canteenWeatherState;
    if (canteenWeatherPromise) return canteenWeatherPromise;
    canteenWeatherPromise = (async function () {
      var controller = typeof AbortController !== "undefined" ? new AbortController() : null;
      var timer = controller ? setTimeout(function () { controller.abort(); }, 1800) : null;
      try {
        var response = await fetch(CANTEEN_WEATHER_URL, { cache: "no-store", signal: controller && controller.signal });
        if (!response.ok) throw new Error("Weather HTTP " + response.status);
        var raw = await response.json();
        if (!raw || !raw.current) throw new Error("Sem condições atuais.");
        canteenWeatherState = { fetchedAt: Date.now(), current: raw.current };
        saveCachedCanteenWeather(canteenWeatherState);
        return canteenWeatherState;
      } catch (error) {
        return canteenWeatherState || null;
      } finally {
        if (timer) clearTimeout(timer);
        canteenWeatherPromise = null;
      }
    })();
    return canteenWeatherPromise;
  }

  function canteenBuiltInAIProvider() {
    if (typeof self.LanguageModel !== "undefined" && self.LanguageModel && typeof self.LanguageModel.create === "function") {
      return { kind: "language-model", api: self.LanguageModel };
    }
    if (window.ai && window.ai.languageModel && typeof window.ai.languageModel.create === "function") {
      return { kind: "legacy-window-ai", api: window.ai.languageModel };
    }
    return null;
  }

  function canteenAIKey(date, mealType, meal, weather) {
    var compact = asArray(meal && meal.items).map(function (item) {
      return [item.type || "", item.description || "", Number(item.kcal) || 0, asArray(item.allergens).join(",")];
    });
    var weatherContext = canteenWeatherContext(weather);
    var dayContext = canteenPersonalContext(date);
    var weatherSignature = weatherContext.kind + "-" + (weatherContext.temperature == null ? "x" : Math.round(weatherContext.temperature / 3));
    return String(date || "") + "-" + String(mealType || "") + "-" + weatherSignature + "-" + hashText(JSON.stringify(compact) + "|" + dayContext.signature + "|" + canteenAllergenFilterSignature());
  }
  function canteenAIFallbackDish(item) {
    var tone = canteenDishTone(item.type, item.description);
    var descriptions = {
      soup: "Uma entrada quente e simples para começar a refeição.",
      vegetarian: "Uma opção vegetal com aquele conforto de almoço no campus.",
      fish: "Um clássico de peixe apresentado de forma simples e familiar.",
      meat: "Um prato reconfortante para uma pausa de almoço mais composta.",
      poultry: "Uma opção familiar e versátil para o menu do dia.",
      classic: "Uma escolha clássica da ementa, sem complicações."
    };
    var tags = {
      soup: ["Quente", "Entrada"],
      vegetarian: ["Vegetal", "Campus comfort"],
      fish: ["Peixe", "Clássico"],
      meat: ["Reconfortante", "Clássico"],
      poultry: ["Familiar", "Clássico"],
      classic: ["Menu do dia", "Clássico"]
    };
    return {
      officialName: cleanText(item.description || "Prato do dia"),
      displayName: cleanText(item.description || "Prato do dia"),
      description: descriptions[tone] || descriptions.classic,
      tags: tags[tone] || tags.classic
    };
  }

  function canteenMainItems(meal) {
    return asArray(meal && meal.items).filter(function (item) {
      return !/sopa|creme|caldo/i.test(String(item.type || "") + " " + String(item.description || ""));
    });
  }

  function canteenHasProteinSource(item) {
    var value = (String(item && item.type || "") + " " + String(item && item.description || "")).toLowerCase();
    return /peixe|pesc|bacalhau|salmão|atum|pescada|dourad|carapau|carne|porco|vaca|vitela|febr|bife|hambúrg|almôndeg|frango|peru|ovo|omelete|tofu|seitan|grão|lentilh|feijão/.test(value);
  }

  function canteenFallbackRecommendedItem(items, weatherContext) {
    if (!items.length) return null;
    var preferences = state && state.settings && state.settings.canteenFoodPreferences || {};
    return items.slice().sort(function (a, b) {
      function score(item) {
        var name = cleanText(item && item.description || "").toLowerCase();
        var tone = canteenDishTone(item && item.type, item && item.description);
        var value = Number(item && item.kcal) || 0;
        if (/ervilha/.test(name)) value -= 1200;
        if (tone === "fish") value -= weatherContext.kind === "hot" ? 35 : 150;
        if (tone === "meat" || tone === "poultry") value += 65;
        if (tone === "vegetarian") value += 15;
        if (["rainy", "snowy", "cold", "cloudy", "foggy"].indexOf(weatherContext.kind) >= 0) value += Number(item && item.kcal) * .08;
        if (weatherContext.kind === "hot" && tone !== "fish" && tone !== "vegetarian") value -= 45;
        return value;
      }
      return score(b) - score(a);
    })[0] || items[0];
  }
  function canteenSafeCompleteCopy(value, fallback, maxChars) {
    var text = cleanText(value || fallback || "");
    if (!text) return cleanText(fallback || "");
    if (!maxChars || text.length <= maxChars) return text;
    var clipped = text.slice(0, maxChars + 1);
    var punctuation = Math.max(clipped.lastIndexOf("."), clipped.lastIndexOf("!"), clipped.lastIndexOf("?"));
    if (punctuation >= Math.min(70, Math.floor(maxChars * .55))) return clipped.slice(0, punctuation + 1).trim();
    var words = clipped.slice(0, maxChars).replace(/\s+\S*$/, "").trim();
    return (words || cleanText(fallback || "")).replace(/[,:;\-–—]+$/, "") + ".";
  }

  function canteenPersonalContext(date) {
    var selectedDate = localDate(date) || new Date();
    var weekday = selectedDate.getDay();
    var now = canteenPortugalParts(new Date());
    var isToday = date === now.iso;
    var isPastDay = String(date || "") < now.iso;
    var isFutureDay = String(date || "") > now.iso;
    var referenceMinutes = isToday ? now.minutes : (isPastDay ? 24 * 60 : 0);
    var semester = currentSemester();
    var semesterStart = cleanText(semester && semester.startDate || "");
    var semesterEnd = cleanText(semester && semester.endDate || "");
    var dateInsideSemester = !semester || ((!semesterStart || String(date || "") >= semesterStart) && (!semesterEnd || String(date || "") <= semesterEnd));

    var scheduleEntries = dateInsideSemester ? semesterItems("schedule").filter(function (entry) {
      return Number(entry.weekday) === weekday;
    }).sort(function (a, b) {
      return timeMinutes(a.start) - timeMinutes(b.start);
    }) : [];

    var classes = scheduleEntries.map(function (entry) {
      var course = courseById(entry.courseId);
      var lesson = linkedLessonForSlot(entry, date);
      return {
        start: cleanText(entry.start || lesson && lesson.start || ""),
        end: cleanText(entry.end || lesson && lesson.end || ""),
        course: cleanText(course && (course.name || course.code) || "Aula"),
        title: cleanText(lesson && lesson.title || ""),
        type: cleanText(entry.type || lesson && lesson.type || ""),
        room: cleanText(entry.room || lesson && lesson.room || "")
      };
    });

    semesterItems("lessons").filter(function (lesson) {
      if (lesson.date !== date) return false;
      return !scheduleEntries.some(function (entry) { return lessonMatchesSchedule(lesson, entry); });
    }).forEach(function (lesson) {
      var course = courseById(lesson.courseId);
      classes.push({
        start: cleanText(lesson.start || ""),
        end: cleanText(lesson.end || ""),
        course: cleanText(course && (course.name || course.code) || "Aula"),
        title: cleanText(lesson.title || ""),
        type: cleanText(lesson.type || ""),
        room: cleanText(lesson.room || "")
      });
    });
    classes.sort(function (a, b) { return timeMinutes(a.start) - timeMinutes(b.start); });

    var allAssessments = semesterItems("assessments").slice().sort(function (a, b) {
      return String(a.date || "").localeCompare(String(b.date || "")) || timeMinutes(a.time || "23:59") - timeMinutes(b.time || "23:59");
    });
    var dayAssessments = allAssessments.filter(function (item) { return item.date === date; }).map(function (item) {
      var course = courseById(item.courseId);
      var minutes = item.time ? timeMinutes(item.time) : null;
      var status = "today-unspecified";
      if (isPastDay) status = "completed";
      else if (isFutureDay) status = "upcoming";
      else if (minutes != null && minutes < referenceMinutes) status = "completed";
      else if (minutes != null && minutes >= referenceMinutes) status = "upcoming";
      return {
        id: item.id,
        title: cleanText(item.title || item.type || "Avaliação"),
        type: cleanText(item.type || "Avaliação"),
        course: cleanText(course && (course.name || course.code) || ""),
        time: cleanText(item.time || ""),
        minutes: minutes,
        status: status
      };
    });

    var completedToday = dayAssessments.filter(function (item) { return item.status === "completed"; });
    var upcomingToday = dayAssessments.filter(function (item) { return item.status === "upcoming"; });
    var latestCompletedToday = completedToday.length ? completedToday[completedToday.length - 1] : null;
    var nextAssessmentToday = upcomingToday.length ? upcomingToday[0] : null;
    var nextAssessmentAfterLunch = upcomingToday.find(function (item) {
      return item.minutes != null && item.minutes >= 12 * 60 + 30;
    }) || null;

    var futureAssessments = allAssessments.filter(function (item) {
      if (String(item.date || "") > String(date || "")) return true;
      if (item.date !== date || !isToday || !item.time) return false;
      return timeMinutes(item.time) >= referenceMinutes;
    });
    var noFutureAssessments = !!(latestCompletedToday && futureAssessments.length === 0);
    var semesterLooksFinished = !!(semester && semester.endDate && String(date || "") >= addISODays(semester.endDate, -21));
    var lastAssessmentOfAcademicPeriod = !!(noFutureAssessments && semesterLooksFinished);

    var afterLunchClasses = classes.filter(function (entry) { return timeMinutes(entry.start) >= 12 * 60 + 30; });
    var completedClasses = classes.filter(function (entry) { return isPastDay || (isToday && timeMinutes(entry.end || entry.start) <= referenceMinutes); });
    var remainingClasses = classes.filter(function (entry) { return isFutureDay || (isToday && timeMinutes(entry.start) >= referenceMinutes); });
    var lastClass = classes.length ? classes[classes.length - 1] : null;

    var privateBits = [
      "Evita recomendar pratos com ervilhas.",
      "Os pratos de peixe da cantina têm preferência baixa, mas podem ganhar se forem claramente a melhor escolha.",
      "O objetivo pessoal é ganhar peso, por isso privilegia opções mais energéticas quando fizer sentido."
    ];

    var academicContext = {
      date: date,
      today: isToday,
      currentTime: isToday ? minutesToTime(referenceMinutes) : "",
      semesterActiveOnDate: dateInsideSemester,
      classes: classes,
      classesAfterLunch: afterLunchClasses,
      completedClasses: completedClasses,
      remainingClasses: remainingClasses,
      lastClassEnd: lastClass && (lastClass.end || lastClass.start) || "",
      assessmentsToday: dayAssessments,
      nextAssessmentToday: nextAssessmentToday,
      nextAssessmentAfterLunch: nextAssessmentAfterLunch,
      latestCompletedAssessmentToday: latestCompletedToday,
      noMoreAssessmentsRegisteredThisSemester: noFutureAssessments,
      lastAssessmentOfAcademicPeriod: lastAssessmentOfAcademicPeriod,
      semesterName: cleanText(semester && semester.name || ""),
      academicYear: cleanText(semester && semester.academicYear || "")
    };

    var promptContext = {
      date: date,
      today: isToday,
      classes: classes.map(function (entry) {
        return [entry.start, entry.end, entry.course, entry.title].filter(Boolean);
      }),
      assessments: dayAssessments.map(function (item) {
        return [item.time, item.course, item.title, item.status].filter(Boolean);
      }),
      nextAssessmentAfterLunch: nextAssessmentAfterLunch ? [nextAssessmentAfterLunch.time, nextAssessmentAfterLunch.course, nextAssessmentAfterLunch.title] : null,
      latestCompletedAssessmentToday: latestCompletedToday ? [latestCompletedToday.time, latestCompletedToday.course, latestCompletedToday.title] : null,
      lastAssessmentOfAcademicPeriod: lastAssessmentOfAcademicPeriod,
      remainingClasses: remainingClasses.map(function (entry) {
        return [entry.start, entry.end, entry.course, entry.title].filter(Boolean);
      })
    };
    var signatureContext = {
      date: date,
      semesterActiveOnDate: dateInsideSemester,
      classes: classes.map(function (entry) { return [entry.start, entry.end, entry.course, entry.title]; }),
      assessments: dayAssessments.map(function (item) { return [item.id, item.time, item.status]; }),
      nextAssessmentAfterLunch: nextAssessmentAfterLunch && nextAssessmentAfterLunch.id || "",
      latestCompletedAssessmentToday: latestCompletedToday && latestCompletedToday.id || "",
      lastAssessmentOfAcademicPeriod: lastAssessmentOfAcademicPeriod,
      completedClassCount: completedClasses.length,
      remainingClassCount: remainingClasses.length
    };

    return {
      assessment: nextAssessmentToday,
      nextAssessmentAfterLunch: nextAssessmentAfterLunch,
      latestCompletedAssessmentToday: latestCompletedToday,
      lastAssessmentOfAcademicPeriod: lastAssessmentOfAcademicPeriod,
      assessmentsToday: dayAssessments,
      assessmentCourse: nextAssessmentToday ? { name: nextAssessmentToday.course } : null,
      nextClass: remainingClasses[0] || afterLunchClasses[0] || null,
      nextClassCourse: remainingClasses[0] ? { name: remainingClasses[0].course } : (afterLunchClasses[0] ? { name: afterLunchClasses[0].course } : null),
      academicPrompt: JSON.stringify(promptContext),
      privatePrompt: privateBits.join(" "),
      signature: hashText(JSON.stringify({
        academic: signatureContext,
        privateRules: privateBits
      }))
    };
  }

  function canteenChefNoteSanitize(value, fallback) {
    var text = canteenSafeCompleteCopy(value, fallback, 260)
      .replace(/[“”"]/g, "")
      .replace(/[–—]/g, ". ")
      .replace(/^\s*-\s*/, "")
      .replace(/\s-\s/g, ". ")
      .replace(/[\u2600-\u27BF]/g, "")
      .replace(/[\uD83C-\uDBFF][\uDC00-\uDFFF]/g, "")
      .replace(/\s{2,}/g, " ")
      .trim();
    return text || cleanText(fallback || "");
  }

  function canteenChefContextInstruction(dayContext) {
    var nextTest = dayContext && dayContext.nextAssessmentAfterLunch;
    var completedTest = dayContext && dayContext.latestCompletedAssessmentToday;
    if (nextTest) {
      return "Existe uma avaliação confirmada depois do almoço: " + JSON.stringify(nextTest) + ". Podes desejar boa sorte, mas apenas porque esta avaliação está explicitamente registada.";
    }
    if (completedTest && dayContext.lastAssessmentOfAcademicPeriod) {
      return "Já aconteceu hoje a última avaliação registada do período: " + JSON.stringify(completedTest) + ". Podes sugerir que agora aproveite a refeição para relaxar.";
    }
    if (completedTest) {
      return "Já aconteceu hoje uma avaliação confirmada: " + JSON.stringify(completedTest) + ". Podes reconhecer que ficou para trás, sem assumir resultado ou nota.";
    }
    return "Não existe nenhuma avaliação confirmada relevante para esta refeição. É proibido mencionar teste, exame, avaliação, prova, apresentação, boa sorte ou último teste.";
  }

  function canteenCompactAcademicContext(dayContext) {
    var parts = [];
    if (dayContext && dayContext.nextAssessmentAfterLunch) {
      var nextTest = dayContext.nextAssessmentAfterLunch;
      parts.push("teste depois do almoço " + [nextTest.time, nextTest.course, nextTest.title].filter(Boolean).join(" "));
    } else if (dayContext && dayContext.latestCompletedAssessmentToday) {
      var completed = dayContext.latestCompletedAssessmentToday;
      parts.push((dayContext.lastAssessmentOfAcademicPeriod ? "último teste do período já terminou " : "teste de hoje já terminou ") + [completed.time, completed.course, completed.title].filter(Boolean).join(" "));
    } else {
      parts.push("sem teste confirmado relevante");
    }
    try {
      var academic = JSON.parse(dayContext && dayContext.academicPrompt || "{}");
      var remaining = asArray(academic.remainingClasses).slice(0, 5).map(function (entry) {
        return asArray(entry).filter(Boolean).join(" ");
      });
      if (remaining.length) parts.push("aulas restantes " + remaining.join("; "));
      else parts.push("sem aulas restantes confirmadas");
    } catch (error) {
      if (dayContext && dayContext.nextClass) parts.push("próxima aula " + [dayContext.nextClass.start, dayContext.nextClass.course, dayContext.nextClass.title].filter(Boolean).join(" "));
    }
    return parts.join(" | ");
  }

  function canteenChefAcademicBrief(dayContext) {
    var parts = [];
    var nextTest = dayContext && dayContext.nextAssessmentAfterLunch;
    var completedTest = dayContext && dayContext.latestCompletedAssessmentToday;
    if (nextTest) {
      parts.push("TESTE CONFIRMADO DEPOIS DO ALMOÇO " + [nextTest.time, nextTest.course, nextTest.title].filter(Boolean).join(" "));
    } else if (completedTest) {
      parts.push((dayContext.lastAssessmentOfAcademicPeriod ? "ÚLTIMO TESTE DO PERÍODO JÁ TERMINOU " : "TESTE DE HOJE JÁ TERMINOU ") + [completedTest.time, completedTest.course, completedTest.title].filter(Boolean).join(" "));
    } else {
      parts.push("NENHUM TESTE CONFIRMADO. Não menciones teste, avaliação nem boa sorte");
    }
    try {
      var academic = JSON.parse(dayContext && dayContext.academicPrompt || "{}");
      var remaining = asArray(academic.remainingClasses).slice(0, 4).map(function (entry) {
        return asArray(entry).filter(Boolean).join(" ");
      });
      parts.push(remaining.length ? "AULAS RESTANTES " + remaining.join("; ") : "SEM AULAS RESTANTES CONFIRMADAS");
    } catch (error) {
      if (dayContext && dayContext.nextClass) {
        parts.push("PRÓXIMA AULA " + [dayContext.nextClass.start, dayContext.nextClass.course, dayContext.nextClass.title].filter(Boolean).join(" "));
      }
    }
    return parts.join(" | ");
  }

  function canteenChefNoteMentionsDish(value, recommendedDish) {
    var note = cleanText(value || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    var dish = cleanText(recommendedDish || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\d+/g, " ");
    if (!note || !dish) return false;
    var ignored = {
      com: 1, sem: 1, de: 1, da: 1, do: 1, das: 1, dos: 1, no: 1, na: 1, nos: 1, nas: 1,
      e: 1, a: 1, o: 1, as: 1, os: 1, um: 1, uma: 1, arroz: 1, massa: 1, batata: 1,
      legumes: 1, molho: 1, salada: 1, forno: 1, fritas: 1, frito: 1, cozida: 1, cozido: 1
    };
    var words = dish.split(/[^a-zà-ÿ]+/i).filter(function (word) {
      return word.length >= 4 && !ignored[word];
    });
    if (!words.length) words = dish.split(/\s+/).filter(function (word) { return word.length >= 4; });
    return words.some(function (word) { return note.indexOf(word) >= 0; });
  }

  function canteenChefNoteValidate(value, dayContext, recommendedDish) {
    var text = canteenChefNoteSanitize(value, "");
    var lower = text.toLowerCase();
    if (!text) throw new Error("A IA local não devolveu uma Chef’s Note válida.");

    // Só falham a publicação regras factuais ou de privacidade. As regras de estilo
    // continuam no prompt, mas nunca apagam uma nota que o utilizador já viu a ser escrita.
    var hasAssessmentWords = /\b(teste|testes|exame|exames|avaliação|avaliações|prova|provas|apresentação|apresentações)\b/.test(lower);
    var hasGoodLuck = /\bboa sorte\b/.test(lower);
    var hasUpcoming = !!(dayContext && dayContext.nextAssessmentAfterLunch);
    var hasCompleted = !!(dayContext && dayContext.latestCompletedAssessmentToday);
    if (!hasUpcoming && hasGoodLuck) throw new Error("A IA desejou boa sorte sem existir uma avaliação confirmada depois do almoço.");
    if (!hasUpcoming && !hasCompleted && hasAssessmentWords) throw new Error("A IA inventou uma avaliação que não existe nos dados.");
    if (/último teste|última avaliação|último exame/.test(lower) && !(dayContext && dayContext.lastAssessmentOfAcademicPeriod)) {
      throw new Error("A IA afirmou que era a última avaliação sem confirmação nos dados.");
    }
    if (/não gostas?|não curtes?|evitas?|queres engordar|ganhar peso|objetivo de peso|peixe da cantina|não gostas? de ervilhas/.test(lower)) {
      throw new Error("A IA expôs uma preferência pessoal que devia permanecer privada.");
    }

    var styleIssues = [];
    if (/bom dia a todos|boa tarde a todos|olá a todos|espero que gostes|espero que gostem|pessoal|malta/.test(lower)) styleIssues.push("saudação genérica");
    if (/\bparece(?:-me)?\b|\bexcelente\b|\bótim[oa]s?\b|\baproveita\b/.test(lower)) styleIssues.push("tom publicitário");
    if (!/\b(eu|preparei|servia|guardava|escolhia|ia)\b/.test(lower)) styleIssues.push("sem primeira pessoa");
    if (recommendedDish && !canteenChefNoteMentionsDish(text, recommendedDish)) styleIssues.push("prato não mencionado literalmente");
    if (styleIssues.length) {
      console.info("[Twenty Cantina AI] nota aceite com avisos de estilo", { issues: styleIssues, recommendedDish: recommendedDish || "" });
    }
    return text;
  }

  function canteenDishWithArticle(value) {
    var name = cleanText(value || "prato do dia");
    var lower = name.charAt(0).toLowerCase() + name.slice(1);
    if (/^(febras|lentilhas|ervilhas|almôndegas|massas|salsichas|iscas)\b/i.test(lower)) return "as " + lower;
    if (/^(douradinhos|hambúrgueres|carapauzinhos|filetes|bifes)\b/i.test(lower)) return "os " + lower;
    if (/^(pescada|solha|salada|lasanha|massinha|perninha|perna|massa|omelete)\b/i.test(lower)) return "a " + lower;
    return "o " + lower;
  }

  function canteenAIFallbackData(meal, weather) {
    var items = canteenMainItems(meal).filter(canteenItemVisible);
    var weatherContext = canteenWeatherContext(weather);
    var recommended = canteenFallbackRecommendedItem(items, weatherContext);
    var energetic = items.slice().sort(function (a, b) { return (Number(b.kcal) || 0) - (Number(a.kcal) || 0); })[0] || null;
    var proteinDishes = items.filter(canteenHasProteinSource).map(function (item) { return cleanText(item.description); });
    var recommendedName = cleanText(recommended && recommended.description || "");
    return {
      chefNote: "",
      recommendedDish: recommendedName,
      recommendationReason: "A escolha do chef para o teu dia.",
      weatherLabel: weatherContext.label,
      mostEnergeticDish: cleanText(energetic && energetic.description || ""),
      proteinDishes: proteinDishes,
      dishes: items.map(canteenAIFallbackDish),
      generatedBy: "rules"
    };
  }
  function parseCanteenAIResponse(value) {
    var text = String(value || "").trim().replace(/^```(?:json|text)?\s*/i, "").replace(/\s*```$/, "");
    if (!text) throw new Error("A IA local não devolveu uma resposta.");
    if (text.charAt(0) === "{") {
      var parsedJSON = JSON.parse(text);
      if (!parsedJSON || typeof parsedJSON !== "object") throw new Error("Resposta inválida da IA local.");
      return parsedJSON;
    }
    var noteMatch = /(?:^|\n)\s*NOTA\s*[:=]\s*([\s\S]*?)(?=\n\s*ID\s*[:=]|$)/i.exec(text);
    var idMatch = /(?:^|\n)\s*ID\s*[:=]\s*(\d+)/i.exec(text);
    var note = cleanText(noteMatch && noteMatch[1] || text.replace(/(?:^|\n)\s*ID\s*[:=].*$/im, "").replace(/^\s*NOTA\s*[:=]\s*/i, ""));
    if (!note) throw new Error("A IA local não devolveu uma Chef’s Note válida.");
    return { chefNote: note, dishId: idMatch ? Number(idMatch[1]) : null };
  }

  function canteenMergeAIStreamChunk(current, chunk) {
    var previous = String(current || "");
    var next = String(chunk || "");
    if (!next) return previous;
    if (!previous) return next;
    if (next.indexOf(previous) === 0) return next;
    if (previous.indexOf(next) === 0) return previous;
    var maxOverlap = Math.min(previous.length, next.length, 120);
    for (var size = maxOverlap; size > 0; size -= 1) {
      if (previous.slice(-size) === next.slice(0, size)) return previous + next.slice(size);
    }
    return previous + next;
  }

  function canteenStreamingRecommendedDish(value, dishes) {
    var match = /(?:^|\n)\s*ID\s*[:=]\s*(\d+)/i.exec(String(value || ""));
    if (!match) return null;
    var id = Number(match[1]);
    return asArray(dishes).find(function (item) { return item.id === id; }) || null;
  }

  function canteenStreamingChefNote(value) {
    var text = String(value || "");
    var protocolMatch = /(?:^|\n)\s*NOTA\s*[:=]\s*/i.exec(text);
    if (protocolMatch) {
      var protocolOutput = text.slice(protocolMatch.index + protocolMatch[0].length).split(/\n\s*ID\s*[:=]/i)[0];
      return protocolOutput
        .replace(/[“”"]/g, "")
        .replace(/[–—]/g, ". ")
        .replace(/^\s*-\s*/, "")
        .replace(/\s-\s/g, ". ")
        .replace(/[\u2600-\u27BF]/g, "")
        .replace(/[\uD83C-\uDBFF][\uDC00-\uDFFF]/g, "")
        .replace(/\s{2,}/g, " ")
        .slice(0, 260)
        .trim();
    }
    var match = /"chefNote"\s*:\s*"/.exec(text);
    var output = "";
    if (match) {
      var index = match.index + match[0].length;
      var escaped = false;
      for (; index < text.length; index += 1) {
        var character = text.charAt(index);
        if (escaped) {
          if (character === "n" || character === "r" || character === "t") output += " ";
          else if (character === "u" && /^[0-9a-fA-F]{4}$/.test(text.slice(index + 1, index + 5))) {
            output += String.fromCharCode(parseInt(text.slice(index + 1, index + 5), 16));
            index += 4;
          } else output += character;
          escaped = false;
        } else if (character === "\\") {
          escaped = true;
        } else if (character === '"') {
          break;
        } else {
          output += character;
        }
      }
    } else {
      output = text.replace(/^```(?:text)?\s*/i, "").replace(/\s*```$/, "");
    }
    return output
      .replace(/[“”"]/g, "")
      .replace(/[–—]/g, ". ")
      .replace(/^\s*-\s*/, "")
      .replace(/\s-\s/g, ". ")
      .replace(/[\u2600-\u27BF]/g, "")
      .replace(/[\uD83C-\uDBFF][\uDC00-\uDFFF]/g, "")
      .replace(/\s{2,}/g, " ")
      .slice(0, 260)
      .trim();
  }

  function canteenStreamingChefNoteAllowed(value, dayContext) {
    var lower = cleanText(value || "").toLowerCase();
    if (!lower) return false;
    if (/bom dia a todos|boa tarde a todos|olá a todos|espero que gostes|espero que gostem|pessoal|malta/.test(lower)) return false;
    if (/\bparece(?:-me)?\b|\bexcelente\b|\bótim[oa]s?\b|\baproveita\b/.test(lower)) return false;
    var hasAssessmentWords = /\b(teste|testes|exame|exames|avaliação|avaliações|prova|provas|apresentação|apresentações)\b/.test(lower);
    var hasGoodLuck = /\bboa sorte\b/.test(lower);
    var hasUpcoming = !!(dayContext && dayContext.nextAssessmentAfterLunch);
    var hasCompleted = !!(dayContext && dayContext.latestCompletedAssessmentToday);
    if (!hasUpcoming && hasGoodLuck) return false;
    if (!hasUpcoming && !hasCompleted && hasAssessmentWords) return false;
    if (/último teste|última avaliação|último exame/.test(lower) && !(dayContext && dayContext.lastAssessmentOfAcademicPeriod)) return false;
    return true;
  }

  function updateCanteenAIStreamingNote(value, dayContext) {
    var preview = canteenStreamingChefNote(value);
    if (!canteenStreamingChefNoteAllowed(preview, dayContext)) return;
    canteenAIState.streamText = preview;
    if (route.name !== "canteen" || !view) return;
    var writing = view.querySelector(".campus-chef-writing");
    if (!writing) return;
    writing.textContent = preview;
    writing.classList.add("has-stream");
  }

  function normalizeCanteenAIData(raw, meal, weather, dayContext) {
    var fallback = canteenAIFallbackData(meal, weather);
    var officialNames = fallback.dishes.map(function (item) { return item.officialName; });
    var rawRecommended = cleanText(raw && (raw.recommendedDish || raw.recommendation || ""));
    var recommended = officialNames.find(function (name) { return name.toLowerCase() === rawRecommended.toLowerCase(); }) || fallback.recommendedDish;
    var rawDishes = asArray(raw && raw.dishes);
    var dishes = fallback.dishes.map(function (base) {
      var match = rawDishes.find(function (item) {
        return cleanText(item && (item.officialName || item.name || "")).toLowerCase() === base.officialName.toLowerCase();
      });
      if (!match) return base;
      return {
        officialName: base.officialName,
        displayName: cleanText(match.displayName || base.displayName).slice(0, 90) || base.displayName,
        description: canteenSafeCompleteCopy(match.description, base.description, 180),
        tags: asArray(match.tags).map(cleanText).filter(Boolean).slice(0, 3)
      };
    });
    var chefNote = canteenChefNoteValidate(raw && raw.chefNote, dayContext, recommended);
    return {
      chefNote: chefNote,
      recommendedDish: recommended,
      recommendationReason: canteenSafeCompleteCopy(raw && raw.recommendationReason, fallback.recommendationReason, 150),
      weatherLabel: fallback.weatherLabel,
      mostEnergeticDish: fallback.mostEnergeticDish,
      proteinDishes: fallback.proteinDishes,
      dishes: dishes,
      generatedBy: "built-in-ai"
    };
  }

  function canteenAIInsightFor(description) {
    var data = canteenAIState && canteenAIState.data;
    if (!data || !state.settings.canteenAIDescriptions) return null;
    var key = cleanText(description || "").toLowerCase();
    return asArray(data.dishes).find(function (item) { return cleanText(item.officialName || "").toLowerCase() === key; }) || null;
  }

  function canteenAIRecommendationFor(description) {
    var data = canteenAIState && canteenAIState.data;
    if (!data || data.generatedBy !== "built-in-ai") return null;
    var key = cleanText(description || "").toLowerCase();
    if (!key || cleanText(data.recommendedDish || "").toLowerCase() !== key) return null;
    return { reason: data.recommendationReason || "A escolha do chef para hoje." };
  }

  function canteenDishSmartBadges(item) {
    var data = canteenAIState && canteenAIState.data;
    if (!data || !state.settings.canteenAIDescriptions) return "";
    var key = cleanText(item && item.description || "").toLowerCase();
    var badges = [];
    if (data.generatedBy === "built-in-ai" && cleanText(data.recommendedDish || "").toLowerCase() === key) {
      badges.push('<span class="is-chef" title="Escolhido pela IA local a partir da ementa, do horário e das avaliações do dia"><i data-lucide="chef-hat"></i>Chef recommendation</span>');
    }
    if (cleanText(data.mostEnergeticDish || "").toLowerCase() === key && Number(item && item.kcal) > 0) {
      badges.push('<span class="is-energy" title="Comparação feita apenas com as kcal oficiais publicadas"><i data-lucide="flame"></i>Mais energético</span>');
    }
    if (asArray(data.proteinDishes).some(function (name) { return cleanText(name).toLowerCase() === key; })) {
      badges.push('<span class="is-protein" title="Identificado pelo nome/categoria do prato; não substitui informação nutricional"><i data-lucide="dumbbell"></i>Proteína em destaque</span>');
    }
    return badges.length ? '<div class="diner-smart-badges">' + badges.join("") + '</div>' : "";
  }

  function canteenAIAvailabilityLabel() {
    var provider = canteenBuiltInAIProvider();
    if (!state.settings.canteenAIEnabled) return { label: "Desativada", className: "badge" };
    if (!provider) return { label: "IA indisponível", className: "badge badge-yellow" };
    if (canteenAIState.status === "ready" && canteenAIState.source === "built-in-ai") return { label: "IA no dispositivo", className: "badge badge-mint" };
    if (canteenAIState.status === "loading") return { label: canteenAIState.phase === "writing" ? "A escrever" : "A escolher", className: "badge badge-violet" };
    if (canteenAIState.status === "downloadable") return { label: "Modelo disponível", className: "badge badge-violet" };
    if (canteenAIState.status === "error") return { label: "Erro na IA", className: "badge badge-yellow" };
    if (canteenAIState.status === "unavailable") return { label: "IA indisponível", className: "badge badge-yellow" };
    return { label: "Chrome AI compatível", className: "badge badge-mint" };
  }

  async function createCanteenAISession(provider, progressCallback) {
    if (provider.kind === "language-model") {
      return provider.api.create({
        monitor: function (monitor) {
          monitor.addEventListener("downloadprogress", function (event) {
            progressCallback(Math.round((Number(event.loaded) || 0) * 100));
          });
        }
      });
    }
    return provider.api.create({
      monitor: function (monitor) {
        if (monitor && monitor.addEventListener) monitor.addEventListener("downloadprogress", function (event) {
          progressCallback(Math.round((Number(event.loaded) || 0) * 100));
        });
      }
    });
  }

  async function getCanteenAISession(provider, progressCallback) {
    if (canteenAISession && canteenAISessionProviderKind === provider.kind && canteenAISessionUses < 6) return canteenAISession;
    destroyCanteenAISession();
    canteenAISession = await createCanteenAISession(provider, progressCallback);
    canteenAISessionProviderKind = provider.kind;
    canteenAISessionUses = 0;
    return canteenAISession;
  }

  function canteenParseChoiceId(value, dishes) {
    var text = String(value || "").trim();
    var match = /(?:^|\n)\s*ID\s*[:=]\s*(\d+)/i.exec(text) || /^\s*(\d+)\s*$/.exec(text);
    if (!match) return null;
    var id = Number(match[1]);
    return asArray(dishes).find(function (item) { return item.id === id; }) || null;
  }

  function canteenPublishAIChoice(selected, meal, weather) {
    if (!selected) return;
    var partial = canteenAIFallbackData(meal, weather);
    partial.recommendedDish = selected.officialName;
    partial.recommendationReason = "A escolha do chef para o teu dia.";
    partial.generatedBy = "built-in-ai";
    partial.chefNote = "";
    canteenAIState.data = partial;
    canteenAIState.source = "built-in-ai";
    canteenAIState.phase = "writing";
    canteenAIState.streamText = "";
    refreshCanteenAIView();
  }

  function canteenWaitForPaint() {
    return new Promise(function (resolve) {
      if (typeof requestAnimationFrame !== "function") {
        setTimeout(resolve, 0);
        return;
      }
      requestAnimationFrame(function () { setTimeout(resolve, 0); });
    });
  }

  async function generateCanteenAIData(provider, meal, weather) {
    var weatherContext = canteenWeatherContext(weather);
    var dayContext = canteenPersonalContext(canteenSelectedDate || todayISO());
    var visibleItems = canteenMainItems(meal).filter(canteenItemVisible);
    var dishes = visibleItems.map(function (item, index) {
      return {
        id: index,
        officialName: cleanText(item.description),
        category: cleanText(item.type || ""),
        kcal: Number(item.kcal) || null,
        sourceItem: item
      };
    });
    if (!dishes.length) throw new Error("Sem pratos disponíveis para escrever a Chef’s Note.");

    var compactDishes = dishes.map(function (item) {
      return [item.id, item.officialName, item.category, item.kcal];
    });
    var academic = canteenChefAcademicBrief(dayContext);
    var choicePrompt = [
      "Escolhe o prato que o chef da Cantina FCT recomendaria hoje ao Johnny.",
      "Usa o contexto apenas para decidir. Nunca o expliques.",
      "Preferências silenciosas: evitar ervilhas; peixe tem preferência baixa; favorecer pratos mais energéticos quando fizer sentido.",
      "CONTEXTO " + academic,
      "TEMPO " + weatherContext.prompt,
      "PRATOS [id,nome,tipo,kcal] " + JSON.stringify(compactDishes),
      "Responde apenas com ID=<número>. Nada mais."
    ].join("\n");

    destroyCanteenAISession();
    var session = await createCanteenAISession(provider, function (progress) {
      canteenAIState.progress = progress;
    });
    canteenAISession = session;
    canteenAISessionProviderKind = provider.kind;
    canteenAISessionUses = 0;
    var startedAt = performance.now();
    try {
      var choiceRaw = "";
      var choiceFirstChunkAt = 0;
      var choiceAt = 0;
      var selected = null;
      if (typeof session.promptStreaming === "function") {
        var choiceStream = session.promptStreaming(choicePrompt);
        for await (var choiceChunk of choiceStream) {
          if (!choiceFirstChunkAt) choiceFirstChunkAt = performance.now();
          choiceRaw = canteenMergeAIStreamChunk(choiceRaw, choiceChunk);
          if (!selected) {
            selected = canteenParseChoiceId(choiceRaw, dishes);
            if (selected) {
              choiceAt = performance.now();
              canteenPublishAIChoice(selected, meal, weather);
            }
          }
        }
      } else {
        choiceRaw = await session.prompt(choicePrompt);
      }

      if (!selected) selected = canteenParseChoiceId(choiceRaw, dishes);
      if (!selected) {
        var weatherChoice = canteenFallbackRecommendedItem(visibleItems, weatherContext);
        selected = dishes.find(function (item) { return item.sourceItem === weatherChoice; }) || dishes[0];
      }
      if (!choiceAt) choiceAt = performance.now();
      canteenPublishAIChoice(selected, meal, weather);
      await canteenWaitForPaint();

      var notePrompt = [
        "Agora escreve apenas a Chef’s Note para o prato " + selected.officialName + ".",
        "Fala em privado com o Johnny como o chef que cozinhou e lhe está a servir o almoço.",
        "Escreve 18 a 30 palavras em português europeu. Menciona naturalmente o prato. Usa primeira pessoa.",
        "Soa humano e específico. Não soes a publicidade, crítica de restaurante ou chatbot.",
        "Não escrevas parece excelente, ótimo, aproveita, espero que gostes, bom dia a todos ou frases vazias.",
        "Só menciona aulas ou testes quando este contexto os confirma: " + academic,
        "Sem emojis, sem travessões e sem dois pontos. Podes usar :) uma vez.",
        "Não inventes ingredientes, alergénios, kcal, benefícios ou factos académicos.",
        "Responde apenas com a nota, sem título, ID, aspas ou explicações."
      ].join("\n");

      var rawNote = "";
      var noteFirstChunkAt = 0;
      var noteStreamRecovered = false;
      try {
        if (typeof session.promptStreaming === "function") {
          var noteStream = session.promptStreaming(notePrompt);
          for await (var noteChunk of noteStream) {
            if (!noteFirstChunkAt) noteFirstChunkAt = performance.now();
            rawNote = canteenMergeAIStreamChunk(rawNote, noteChunk);
            updateCanteenAIStreamingNote(rawNote, dayContext);
          }
        } else {
          rawNote = await session.prompt(notePrompt);
        }
      } catch (noteError) {
        var recoveredNote = canteenChefNoteSanitize(rawNote, "");
        var recoveredWords = recoveredNote ? recoveredNote.split(/\s+/).filter(Boolean).length : 0;
        if (recoveredWords < 8) throw noteError;
        rawNote = recoveredNote;
        noteStreamRecovered = true;
        console.warn("[Twenty Cantina AI] stream interrompido; texto já recebido foi preservado", {
          words: recoveredWords,
          error: noteError && noteError.message || String(noteError)
        });
      }
      canteenAISessionUses += 2;
      var chefNote = canteenChefNoteValidate(rawNote, dayContext, selected.officialName);
      var result = normalizeCanteenAIData({
        chefNote: chefNote,
        recommendedDish: selected.officialName,
        recommendationReason: "A escolha do chef para o teu dia."
      }, meal, weather, dayContext);
      console.info("[Twenty Cantina AI] geração concluída", {
        choiceFirstTextMs: choiceFirstChunkAt ? Math.round(choiceFirstChunkAt - startedAt) : null,
        choiceMs: Math.round(choiceAt - startedAt),
        noteFirstTextMs: noteFirstChunkAt ? Math.round(noteFirstChunkAt - choiceAt) : null,
        durationMs: Math.round(performance.now() - startedAt),
        noArtificialTimeout: true,
        streamRecovered: noteStreamRecovered,
        choicePromptChars: choicePrompt.length,
        notePromptChars: notePrompt.length,
        dishes: dishes.length,
        recommendedDish: result.recommendedDish
      });
      return result;
    } catch (error) {
      console.warn("[Twenty Cantina AI] geração falhou; nenhuma nota genérica foi publicada", {
        durationMs: Math.round(performance.now() - startedAt),
        error: error && error.message || String(error)
      });
      throw error;
    } finally {
      canteenAIAbortController = null;
      destroyCanteenAISession();
    }
  }

  function refreshCanteenAIView() {
    if (route.name !== "canteen" || !view) return;
    var day = canteenDayForDate(canteenSelectedDate);
    var meal = day && asArray(day.meals).find(function (candidate) { return canteenMealType(candidate) === "lunch"; });
    if (!meal) return;
    var currentNote = view.querySelector(".campus-chef-note");
    if (currentNote) {
      var holder = document.createElement("div");
      holder.innerHTML = renderCanteenPosterAINote(meal);
      if (holder.firstElementChild) currentNote.replaceWith(holder.firstElementChild);
    }
    Array.from(view.querySelectorAll(".campus-menu-item[data-canteen-dish]")).forEach(function (node) {
      var officialName = node.getAttribute("data-canteen-dish") || "";
      var recommendation = canteenAIRecommendationFor(officialName);
      var insight = canteenAIInsightFor(officialName);
      node.classList.toggle("is-chef-recommended", !!recommendation);
      var oldBadge = node.querySelector(".campus-recommended");
      if (oldBadge) oldBadge.remove();
      if (recommendation) {
        var badge = document.createElement("span");
        badge.className = "campus-recommended";
        badge.innerHTML = '<i data-lucide="map-pin"></i>RECOMENDADO';
        var outline = node.querySelector(".campus-pencil-outline");
        if (outline) outline.insertAdjacentElement("afterend", badge);
        else node.prepend(badge);
      }
      var description = node.querySelector(".campus-menu-description");
      if (description && insight && insight.description) description.textContent = insight.description;
    });
    refreshIcons(view);
  }

  async function ensureCanteenAIForCurrentMeal(forceDownload) {
    if (!state || !state.settings.canteenAIEnabled || !canteenMenu || route.name !== "canteen") return;
    if (canteenAIPromise) return canteenAIPromise;
    var requestedDate = canteenSelectedDate;
    var day = canteenDayForDate(requestedDate);
    var mealType = "lunch";
    var meal = day && asArray(day.meals).find(function (candidate) { return canteenMealType(candidate) === mealType; });
    if (!meal) return;
    if (!canteenMainItems(meal).some(canteenItemVisible)) {
      canteenAIState = { key: requestedDate + "|blocked|" + canteenAllergenFilterSignature(), status: "blocked", availability: "available", source: "rules", data: null, error: "Todos os pratos foram escondidos pelos filtros de alergénios.", progress: null, streamText: "" };
      refreshCanteenAIView();
      return;
    }
    var weather = await ensureCanteenWeather(requestedDate);
    var key = canteenAIKey(requestedDate, mealType, meal, weather);
    if (canteenAIState.key === key && ["ready", "error", "unavailable", "downloadable", "loading"].indexOf(canteenAIState.status) >= 0 && !forceDownload) return;

    var fallback = canteenAIFallbackData(meal, weather);
    var cache = loadCanteenAICache();
    if (!forceDownload && cache[key]) {
      canteenAIState = { key: key, status: "ready", availability: "available", source: cache[key].generatedBy || "built-in-ai", data: cache[key], error: "", progress: 100 };
      refreshCanteenAIView();
      return;
    }

    var provider = canteenBuiltInAIProvider();
    if (!provider) {
      canteenAIState = { key: key, status: "unavailable", availability: "unavailable", source: "rules", data: fallback, error: "A IA local não está disponível neste navegador.", progress: null };
      refreshCanteenAIView();
      return;
    }

    canteenAIRequestKey = key;
    canteenAIPromise = (async function () {
      try {
        var availability = "available";
        if (provider.api && typeof provider.api.availability === "function") {
          availability = await provider.api.availability();
        } else if (provider.api && typeof provider.api.capabilities === "function") {
          var capabilities = await provider.api.capabilities();
          availability = capabilities && (capabilities.available || capabilities.availability) || "available";
        }
        canteenAIState = { key: key, status: "idle", availability: availability, source: "rules", data: null, error: "", progress: null };
        if ((availability === "downloadable" || availability === "downloading" || availability === "after-download") && !forceDownload) {
          canteenAIState.status = "downloadable";
          canteenAIState.data = fallback;
          refreshCanteenAIView();
          return;
        }
        if (availability === "unavailable" || availability === "no") {
          canteenAIState.status = "unavailable";
          canteenAIState.data = fallback;
          canteenAIState.error = "A IA local não está disponível neste navegador.";
          refreshCanteenAIView();
          return;
        }
        canteenAIState.status = "loading";
        canteenAIState.phase = "choosing";
        canteenAIState.progress = 0;
        canteenAIState.streamText = "";
        refreshCanteenAIView();
        var result = await generateCanteenAIData(provider, meal, weather);
        cache[key] = result;
        if (canteenSelectedDate === requestedDate && canteenAIRequestKey === key) {
          canteenAIState = { key: key, status: "ready", availability: "available", source: "built-in-ai", data: result, error: "", progress: 100 };
        }
        var keys = Object.keys(cache);
        if (keys.length > 24) keys.slice(0, keys.length - 24).forEach(function (oldKey) { delete cache[oldKey]; });
        saveCanteenAICache(cache);
      } catch (error) {
        var partialData = canteenAIState.data && canteenAIState.data.generatedBy === "built-in-ai" ? canteenAIState.data : fallback;
        canteenAIState = { key: key, status: "error", availability: "error", source: partialData.generatedBy === "built-in-ai" ? "built-in-ai" : "rules", data: partialData, error: error && error.message || "A IA local não conseguiu escrever a nota.", progress: null, streamText: "" };
      } finally {
        canteenAIPromise = null;
        canteenAIRequestKey = "";
        refreshCanteenAIView();
        if (route.name === "canteen" && canteenSelectedDate !== requestedDate) {
          setTimeout(function () { ensureCanteenAIForCurrentMeal(false); }, 0);
        }
      }
    })();
    return canteenAIPromise;
  }
  function renderCanteenAICard(meal) {
    if (!state.settings.canteenAIEnabled || !meal) return "";
    var data = canteenAIState.data;
    var status = canteenAIAvailabilityLabel();
    var progress = canteenAIState.status === "loading"
      ? '<div class="canteen-ai-progress"><span style="width:' + (canteenAIState.progress == null ? 34 : clamp(canteenAIState.progress, 4, 100)) + '%"></span></div>'
      : "";
    var action = canteenAIState.status === "downloadable"
      ? '<button class="button button-small" type="button" data-action="canteen-ai-prepare"><i data-lucide="sparkles"></i>Preparar IA local</button>'
      : (canteenAIState.status === "error" || canteenAIState.status === "unavailable") && canteenBuiltInAIProvider()
        ? '<button class="button button-small" type="button" data-action="canteen-ai-prepare"><i data-lucide="rotate-cw"></i>Tentar novamente</button>'
        : "";
    var copy = canteenAIState.status === "loading"
      ? "A IA está a ler a ementa, o horário e as avaliações do dia."
      : canteenAIState.source === "built-in-ai"
        ? "Criada no teu dispositivo. Os alergénios continuam a vir apenas da fonte oficial."
        : "A Twenty não publica uma Chef’s Note genérica quando a IA não responde.";
    var note = state.settings.canteenAIChefNote && canteenAIState.status === "ready" && data && data.chefNote ? '<blockquote>“' + esc(data.chefNote) + '”</blockquote>' : canteenAIState.status === "loading" ? '<blockquote class="is-loading">' + esc(canteenAIState.phase === "choosing" ? "a escolher o prato..." : "a escrever...") + '</blockquote>' : "";
    return '<section class="canteen-ai-card"><div class="canteen-ai-mark"><i data-lucide="sparkles"></i></div><div class="canteen-ai-copy"><div class="canteen-ai-title"><span>Chef\'s Note</span><span class="' + status.className + '">' + esc(status.label) + '</span></div>' + note + '<p>' + esc(copy) + '</p>' + progress + '</div>' + action + '</section>';
  }

  function canteenDesserts() {
    return [
      { id: "gelatin", label: "Gelatina", kcal: 60, seasonal: false },
      { id: "yogurt", label: "Iogurte Mimosa", kcal: 95, seasonal: false },
      { id: "fruit", label: "Fruta à disposição", kcal: 80, seasonal: false },
      { id: "mousse", label: "Mousse de chocolate", kcal: null, seasonal: true },
      { id: "aletria", label: "Aletria", kcal: null, seasonal: true },
      { id: "rice-pudding", label: "Arroz doce", kcal: null, seasonal: true },
      { id: "cake", label: "Bolo", kcal: null, seasonal: true },
      { id: "semifreddo", label: "Semifrio", kcal: null, seasonal: true }
    ];
  }
  function canteenDessertById(id) {
    return canteenDesserts().find(function (item) { return item.id === id; }) || null;
  }

  function canteenDishTone(type, description) {
    var value = (String(type || "") + " " + String(description || "")).toLowerCase();
    if (/sopa|creme|caldo/.test(value)) return "soup";
    if (/veget|vegan|tofu|seitan|grão|lentilh|legum/.test(value)) return "vegetarian";
    if (/peixe|pesc|bacalhau|salmão|atum|pescada|dourad|douradinho|carapau/.test(value)) return "fish";
    if (/carne|porco|vaca|vitela|febr|bife|hambúrg|almôndeg/.test(value)) return "meat";
    if (/frango|peru/.test(value)) return "poultry";
    return "classic";
  }

  function canteenDishEmoji(type, description) {
    var value = (String(type || "") + " " + String(description || "")).toLowerCase();
    if (/sopa|creme|caldo/.test(value)) return "🥣";
    if (/veget|vegan|tofu|seitan|grão|lentilh|legum/.test(value)) return "🌿";
    if (/peixe|pesc|bacalhau|salmão|atum|pescada|dourad|douradinho|carapau/.test(value)) return "🐟";
    if (/frango|peru/.test(value)) return "🍗";
    if (/carne|porco|vaca|vitela|febr|bife|hambúrg|almôndeg/.test(value)) return "🥩";
    if (/massa|esparguete|lasanha|tagliatelle/.test(value)) return "🍝";
    if (/arroz|risoto/.test(value)) return "🍚";
    if (/ovo|omelete/.test(value)) return "🍳";
    return "🍽️";
  }

  function canteenDishLabel(type) {
    var value = cleanText(type || "Opção");
    if (!value) return "Opção";
    return value.replace(/^prato\s+/i, "");
  }

  function canteenAllergenEmoji(name) {
    var value = String(name || "").toLowerCase();
    if (/glúten|trigo|cereal/.test(value)) return "🌾";
    if (/crustáce/.test(value)) return "🦐";
    if (/ovo/.test(value)) return "🥚";
    if (/peixe/.test(value)) return "🐟";
    if (/amendoim/.test(value)) return "🥜";
    if (/soja/.test(value)) return "🌱";
    if (/leite|lácteo/.test(value)) return "🥛";
    if (/frutos de casca|noz|avelã|amêndoa/.test(value)) return "🌰";
    if (/aipo/.test(value)) return "🥬";
    if (/mostarda/.test(value)) return "🟡";
    if (/sésamo/.test(value)) return "⚪";
    if (/sulfit/.test(value)) return "🍇";
    if (/tremoço/.test(value)) return "🌿";
    if (/molusc/.test(value)) return "🦑";
    return "⚠️";
  }

  function canteenAllergenPills(item, menu) {
    var codes = asArray(item && item.allergens);
    if (!codes.length) return '<span class="canteen-allergen-pill is-clear">✓ Sem alergénios indicados</span>';
    return codes.map(function (id) {
      var name = menu.allergens && menu.allergens[id] ? menu.allergens[id] : "Alergénio " + id;
      return '<span class="canteen-allergen-pill" title="Código ' + attr(id) + '"><span>' + canteenAllergenEmoji(name) + '</span>' + esc(name) + '</span>';
    }).join("");
  }

  function canteenDayChip(day) {
    var date = localDate(day.date);
    var today = canteenPortugalParts(new Date()).iso;
    var isToday = day.date === today;
    var label = isToday ? "Hoje" : date ? new Intl.DateTimeFormat("pt-PT", { weekday: "short" }).format(date).replace(".", "") : "Dia";
    var dateLabel = date ? new Intl.DateTimeFormat("pt-PT", { day: "2-digit", month: "short" }).format(date).replace(".", "") : day.label;
    return '<button class="canteen-day-chip ' + (day.date === canteenSelectedDate ? "is-active" : "") + '" type="button" data-action="canteen-day" data-date="' + attr(day.date) + '"><span>' + esc(label) + '</span><strong>' + esc(dateLabel) + '</strong>' + (isToday ? '<i aria-hidden="true"></i>' : '') + '</button>';
  }

  function canteenDishIconName(item) {
    var tone = canteenDishTone(item && item.type, item && item.description);
    if (tone === "fish") return "fish-symbol";
    if (tone === "vegetarian") return "sprout";
    if (tone === "poultry") return "drumstick";
    if (tone === "meat") return "beef";
    if (tone === "soup") return "soup";
    return "utensils";
  }

  function canteenDishAllergenIcons(item, menu) {
    var icons = asArray(item && item.allergens).slice(0, 3).map(function (id) {
      var name = menu && menu.allergens && menu.allergens[id] ? menu.allergens[id] : "Alergénio " + id;
      var value = String(name).toLowerCase();
      var icon = /glúten|cereal/.test(value) ? "wheat" : /peixe/.test(value) ? "fish-symbol" : /ovo/.test(value) ? "egg" : /leite/.test(value) ? "milk" : /soja|sésamo|aipo/.test(value) ? "sprout" : "circle-alert";
      return '<i data-lucide="' + icon + '" title="' + attr(name) + '"></i>';
    });
    return icons.length ? '<span class="campus-menu-allergen-icons">' + icons.join("") + '</span>' : "";
  }

  function canteenDishCard(item, menu, options) {
    options = options || {};
    var tone = canteenDishTone(item.type, item.description);
    var insight = canteenAIInsightFor(item.description) || canteenAIFallbackDish(item);
    var recommendation = canteenAIRecommendationFor(item.description);
    var displayName = item.description || "Descrição indisponível";
    var description = insight && insight.description || "Uma escolha simples da ementa de hoje.";
    var kcalText = item.kcal ? Number(item.kcal) : "—";
    var tagName = options.interactive ? "button" : "article";
    var attrs = (options.interactive ? ' type="button" data-action="canteen-select-dish" data-index="' + Number(options.index) + '" aria-pressed="' + (!!options.selected) + '"' : "") + ' data-canteen-dish="' + attr(displayName) + '"';
    var selected = options.selected ? " is-selected" : "";
    var recommended = recommendation ? " is-chef-recommended" : "";
    var badge = recommendation ? '<span class="campus-recommended"><i data-lucide="map-pin"></i>RECOMENDADO</span>' : "";
    return '<' + tagName + ' class="campus-menu-item is-' + tone + selected + recommended + '"' + attrs + '>' +
      '<span class="campus-pencil-outline" aria-hidden="true"></span>' + badge +
      '<span class="campus-menu-item-main"><strong>' + esc(displayName) + '</strong><span class="campus-menu-item-icons"><i data-lucide="' + canteenDishIconName(item) + '"></i>' + canteenDishAllergenIcons(item, menu) + '</span><span class="campus-menu-kcal">' + esc(kcalText) + '</span></span>' +
      '<small class="campus-menu-description">' + esc(description) + '</small>' +
      '</' + tagName + '>';
  }
  function canteenDessertChips(selection, readonly, chosenId) {
    var desserts = canteenDesserts();
    var standard = desserts.filter(function (dessert) { return !dessert.seasonal; });
    var seasonal = desserts.filter(function (dessert) { return dessert.seasonal; });
    function rows(list) {
      return list.map(function (dessert) {
        var selected = (readonly ? chosenId : selection.dessertId) === dessert.id;
        var tagName = readonly ? "span" : "button";
        var attrs = readonly ? "" : ' type="button" data-action="canteen-select-dessert" data-dessert="' + attr(dessert.id) + '" aria-pressed="' + selected + '"';
        return '<' + tagName + ' class="campus-menu-item campus-dessert-item' + (selected ? ' is-selected' : '') + '"' + attrs + '>' +
          '<span class="campus-pencil-outline" aria-hidden="true"></span>' +
          '<span class="campus-menu-item-main"><strong>' + esc(dessert.label) + '</strong><span class="campus-menu-kcal">' + (dessert.kcal == null ? '—' : '≈ ' + dessert.kcal) + '</span></span>' +
          '</' + tagName + '>';
      }).join("");
    }
    return rows(standard) + '<div class="campus-menu-subdivider"><span>SAZONAIS</span></div>' + rows(seasonal);
  }
  function canteenTicketOrderNumber(date, mealType) {
    var seed = String(date || "") + '|' + String(mealType || "lunch") + '|' + String(Date.now()) + '|' + String(Math.random());
    return String((parseInt(hashText(seed).slice(-6), 16) % 900) + 100);
  }

  function canteenTicketCodeHTML(visit) {
    var seed = hashText(String(visit && visit.id || "ticket") + String(visit && visit.orderNumber || "000"));
    var bars = "";
    for (var index = 0; index < 34; index += 1) {
      var digit = parseInt(seed.charAt(index % seed.length), 16) || 0;
      bars += '<i style="--bar:' + (1 + digit % 4) + '"></i>';
    }
    var cells = "";
    for (var cell = 0; cell < 81; cell += 1) {
      var value = parseInt(seed.charAt(cell % seed.length), 16) || 0;
      var finder = (cell < 18 && (cell % 9 < 3 || cell % 9 > 5)) || (cell > 62 && cell % 9 < 3);
      cells += '<i class="' + ((finder || ((value + cell) % 3 === 0)) ? 'is-dark' : '') + '"></i>';
    }
    return '<div class="diner-ticket-codes"><div class="diner-barcode" aria-hidden="true">' + bars + '</div><div class="diner-qr" aria-hidden="true">' + cells + '</div></div>';
  }

  function canteenTicketDraft(date, mealType) {
    var day = canteenDayForDate(date);
    var meal = day && asArray(day.meals).find(function (candidate) { return canteenMealType(candidate) === mealType; });
    var selection = canteenSelectionFor(date, mealType);
    var items = asArray(meal && meal.items);
    var soups = items.filter(function (item) { return /sopa|creme|caldo/i.test(String(item.type || "") + " " + String(item.description || "")); });
    var dishes = items.filter(function (item) { return soups.indexOf(item) < 0; });
    var dish = dishes[Number(selection.dishIndex)];
    if (dish && !canteenItemVisible(dish)) dish = null;
    var dessert = canteenDessertById(selection.dessertId);
    if (!meal || !dish || !dessert) return null;
    var info = canteenMenu.info || {};
    var price = info.socialMeal && info.socialMeal.amount || "3,10 €";
    var soup = soups[0] || null;
    return {
      id: canteenVisitId(date, mealType),
      date: date,
      mealType: mealType,
      mealLabel: canteenMealLabel(mealType),
      ticketIssuedAt: "",
      completedAt: "",
      orderNumber: "",
      price: price,
      totalKcal: (Number(soup && soup.kcal) || 0) + (Number(dish.kcal) || 0) + (Number(dessert.kcal) || 0),
      soup: soup ? { description: soup.description || "Sopa do dia", kcal: Number(soup.kcal) || 0 } : null,
      dish: { description: dish.description || "Prato", type: dish.type || "", kcal: Number(dish.kcal) || 0, allergens: asArray(dish.allergens) },
      dessert: clone(dessert)
    };
  }

  function canteenReceiptHTML(visit, options) {
    if (!visit) return "";
    options = options || {};
    var issued = visit.ticketIssuedAt ? new Date(visit.ticketIssuedAt) : new Date();
    var dateLabel = new Intl.DateTimeFormat("pt-PT", { weekday: "long", day: "2-digit", month: "long", year: "numeric" }).format(issued);
    var timeLabel = new Intl.DateTimeFormat("pt-PT", { hour: "2-digit", minute: "2-digit" }).format(issued);
    var soupRow = visit.soup ? '<li><span>Sopa</span><strong>' + esc(visit.soup.description || "Sopa do dia") + '</strong></li>' : "";
    var orderNumber = String(visit.orderNumber || "000").padStart(3, "0");
    var stateClass = visit.completedAt || options.stamped ? " is-stamped" : "";
    var printClass = options.printing ? " is-printing" : "";
    var stamp = visit.completedAt || options.stamped ? '<div class="diner-ticket-stamp">TABULEIRO<br>SACIADO!</div>' : "";
    return '<div class="diner-ticket' + stateClass + printClass + '">' +
      '<div class="diner-ticket-teeth top"></div><div class="diner-ticket-teeth bottom"></div>' +
      '<header><div><span>CAMPUS DINER FCT</span><small>Passe pessoal Twenty</small></div><strong>#' + esc(orderNumber) + '</strong></header>' +
      '<h3>' + esc(visit.mealLabel || canteenMealLabel(visit.mealType)) + '</h3><p class="diner-ticket-date">' + esc(dateLabel) + ' · ' + esc(timeLabel) + '</p>' +
      '<div class="diner-ticket-rule"></div><ul>' + soupRow + '<li><span>Prato</span><strong>' + esc(visit.dish && visit.dish.description || "—") + '</strong></li><li><span>Sobremesa</span><strong>' + esc(visit.dessert && visit.dessert.label || "—") + '</strong></li></ul>' +
      '<div class="diner-ticket-rule is-dashed"></div><div class="diner-ticket-total"><span><small>Total energético</small><strong>≈ ' + Number(visit.totalKcal || 0) + ' kcal</strong></span><span><small>Preço social</small><strong>' + esc(visit.price || "3,10 €") + '</strong></span></div>' +
      canteenTicketCodeHTML(visit) + '<footer>Apresenta este ticket na tua mente e bom apetite!</footer>' + stamp + '</div>';
  }

  function showCanteenReceipt(visit, options) {
    options = options || {};
    var body = options.printing
      ? '<div class="diner-printer"><div class="diner-printer-top"><span></span><strong>A imprimir o teu passe…</strong></div><div class="diner-printer-slot"></div>' + canteenReceiptHTML(visit, { printing: true }) + '<div class="diner-paper-cut"></div></div>'
      : canteenReceiptHTML(visit, { stamped: options.stamped });
    openModal(options.printing ? "Ticket emitido" : "Passe de refeição", body, { className: "canteen-ticket-modal", footer: '<footer class="modal-foot"><button class="button button-dark" type="button" data-action="close-modal">Fechar</button></footer>' });
    refreshIcons(modalRoot);
  }

  function canteenFeedback() {
    if (navigator.vibrate) navigator.vibrate(28);
    try {
      var AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      var context = new AudioCtx();
      var oscillator = context.createOscillator();
      var gain = context.createGain();
      oscillator.type = "sine";
      oscillator.frequency.value = 720;
      gain.gain.setValueAtTime(0.0001, context.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.055, context.currentTime + 0.01);
      gain.gain.exponentialRampToValueAtTime(0.0001, context.currentTime + 0.09);
      oscillator.connect(gain);
      gain.connect(context.destination);
      oscillator.start();
      oscillator.stop(context.currentTime + 0.1);
      setTimeout(function () { context.close().catch(function () {}); }, 180);
    } catch (error) {}
  }

  function canteenPrinterFeedback() {
    if (navigator.vibrate) navigator.vibrate([18, 22, 18, 22, 32]);
    try {
      var AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      var context = new AudioCtx();
      var gain = context.createGain();
      gain.gain.setValueAtTime(0.0001, context.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.035, context.currentTime + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.0001, context.currentTime + 0.72);
      gain.connect(context.destination);
      [0, .12, .24, .36, .48].forEach(function (offset, index) {
        var oscillator = context.createOscillator();
        oscillator.type = index % 2 ? "square" : "sawtooth";
        oscillator.frequency.value = 105 + index * 17;
        oscillator.connect(gain);
        oscillator.start(context.currentTime + offset);
        oscillator.stop(context.currentTime + offset + .085);
      });
      setTimeout(function () { context.close().catch(function () {}); }, 900);
    } catch (error) {}
  }

  function renderCanteenActiveTicket(visit) {
    return '<section class="campus-dining-state campus-order-active"><div><small>PEDIDO FEITO</small><h3>O teu ticket está pronto.</h3><p>Quando terminares a refeição, fecha o almoço no teu dia.</p></div>' + canteenReceiptHTML(visit) + '<button class="campus-order-button" type="button" data-action="canteen-finish-meal" data-id="' + attr(visit.id) + '">Já almocei</button></section>';
  }
  function renderCanteenCompletedCard(visit, expanded) {
    var mealType = visit.mealType || "lunch";
    if (!expanded) {
      return '<section class="campus-dining-state campus-order-complete"><div><small>ALMOÇO FEITO</small><h3>Refeição concluída.</h3><p>' + esc(visit.dish && visit.dish.description || "Prato escolhido") + ' · ≈ ' + Number(visit.totalKcal || 0) + ' kcal</p></div><div class="campus-state-actions"><button type="button" data-action="canteen-open-receipt" data-id="' + attr(visit.id) + '">Ver ticket</button><button type="button" data-action="canteen-expand-completed" data-key="' + attr(canteenSelectionKey(visit.date, mealType)) + '">Ver ementa</button></div></section>';
    }
    return "";
  }
  function renderCanteenPosterAINote(meal) {
    if (!state.settings.canteenAIEnabled || !meal) return "";
    var status = canteenAIState.status || "idle";
    var data = canteenAIState.data;
    var readyNote = status === "ready" && data && data.generatedBy === "built-in-ai"
      ? canteenChefNoteSanitize(data.chefNote, "")
      : "";
    if (readyNote && state.settings.canteenAIChefNote) {
      return '<aside class="campus-chef-note"><h3>Chef’s Note</h3><p>' + esc(readyNote) + '</p></aside>';
    }
    if (status === "blocked") {
      return '<aside class="campus-chef-note is-error"><h3>Chef’s Note</h3><div class="campus-chef-status"><span>Não há nenhum prato visível para eu recomendar.</span></div></aside>';
    }
    if (status === "error" || status === "unavailable") {
      return '<aside class="campus-chef-note is-error"><h3>Chef’s Note</h3><div class="campus-chef-status"><span>Não consegui escrever a nota agora.</span><button type="button" data-action="canteen-ai-prepare">Tentar novamente</button></div></aside>';
    }
    if (status === "downloadable") {
      return '<aside class="campus-chef-note is-loading"><h3>Chef’s Note</h3><div class="campus-chef-status"><span>A preparar a IA local...</span><button type="button" data-action="canteen-ai-prepare">Preparar agora</button></div></aside>';
    }
    var streamingNote = status === "loading" ? cleanText(canteenAIState.streamText || "") : "";
    var waitingCopy = canteenAIState.phase === "choosing" ? "a escolher o prato..." : "a escrever...";
    return '<aside class="campus-chef-note is-loading" aria-live="polite" aria-busy="true"><h3>Chef’s Note</h3><p class="campus-chef-writing' + (streamingNote ? ' has-stream' : '') + '">' + esc(streamingNote || waitingCopy) + '</p></aside>';
  }
  function canteenMealComparable(meal) {
    return asArray(meal && meal.items).map(function (item) {
      return [cleanText(item.type || "").toLowerCase(), cleanText(item.description || "").toLowerCase(), Number(item.kcal) || 0];
    });
  }

  function canteenMealsEqual(first, second) {
    return JSON.stringify(canteenMealComparable(first)) === JSON.stringify(canteenMealComparable(second));
  }

  function canteenInfoMealRows(meal) {
    if (!meal) return '<p>Sem ementa publicada.</p>';
    return '<div class="campus-info-meal-list">' + asArray(meal.items).map(function (item) {
      return '<div><span>' + esc(canteenDishLabel(item.type)) + '</span><strong>' + esc(item.description || "—") + '</strong><small>' + (item.kcal ? Number(item.kcal) + ' kcal' : '—') + '</small></div>';
    }).join("") + '</div>';
  }

  function showCanteenInfo() {
    var day = canteenDayForDate(canteenSelectedDate);
    var meals = day ? asArray(day.meals) : [];
    var lunch = meals.find(function (meal) { return canteenMealType(meal) === "lunch"; });
    var dinner = meals.find(function (meal) { return canteenMealType(meal) === "dinner"; });
    var hours = canteenServiceHours(canteenMenu);
    var info = canteenMenu.info || {};
    var social = info.socialMeal || {};
    var closures = info.closures || {};
    var days = asArray(canteenMenu.days).slice().sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
    var dinnerDifference = dinner && !canteenMealsEqual(lunch, dinner)
      ? '<section class="campus-info-section"><h3>Jantar diferente</h3><p>A ementa principal mostra o almoço. Para este dia, o jantar muda para:</p>' + canteenInfoMealRows(dinner) + '</section>'
      : "";
    var allergenRows = Object.keys(canteenMenu.allergens || {}).map(function (id) {
      return '<span><b>' + esc(id) + '</b>' + esc(canteenMenu.allergens[id]) + '</span>';
    }).join("");
    var aiAction = canteenAIState.status === "downloadable" ? '<button class="campus-info-action" type="button" data-action="canteen-ai-prepare">Preparar IA local</button>' : "";
    var source = canteenMenu.pageUrl || CANTEEN_PAGE_URL;
    var body = '<div class="campus-info-days"><strong>Escolher dia</strong><div>' + days.map(canteenDayChip).join("") + '</div></div>' +
      '<section class="campus-info-section campus-info-facts"><div><span>Almoço</span><strong>' + esc(hours.lunch.start) + ' às ' + esc(hours.lunch.end) + '</strong></div><div><span>Jantar</span><strong>' + esc(hours.dinner.start) + ' às ' + esc(hours.dinner.end) + '</strong></div><div><span>Preço social</span><strong>' + esc(social.amount || '3,10 €') + '</strong></div></section>' +
      (social.includes ? '<section class="campus-info-section"><h3>Incluído na refeição</h3><p>' + esc(social.includes) + '</p></section>' : '') +
      dinnerDifference +
      '<section class="campus-info-section"><h3>Alergénios</h3><div class="campus-info-allergens">' + allergenRows + '</div><p class="campus-info-notice">' + esc(canteenMenu.allergenNotice || 'Confirma sempre os alergénios junto da unidade.') + '</p></section>' +
      '<section class="campus-info-section"><h3>Funcionamento</h3><p>' + esc(closures.summer || 'O funcionamento pode mudar durante férias e pausas letivas.') + '</p>' + (closures.seasonal ? '<p>' + esc(closures.seasonal) + '</p>' : '') + (closures.alternatives ? '<p>' + esc(closures.alternatives) + '</p>' : '') + '</section>' +
      '<section class="campus-info-section campus-info-source"><a href="' + attr(source) + '" target="_blank" rel="noopener noreferrer">Abrir fonte oficial</a><button type="button" data-action="refresh-canteen">Atualizar ementa</button>' + aiAction + '</section>';
    openModal('Informação da cantina', body, { className: 'canteen-info-modal', footer: '<footer class="modal-foot"><button class="campus-info-close" type="button" data-action="close-modal">Fechar</button></footer>' });
  }

  function renderCanteenMeal(meal, menu, options) {
    options = options || {};
    var mealType = "lunch";
    var date = options.date || canteenSelectedDate;
    var visit = canteenVisitFor(date, mealType);
    var expandKey = canteenSelectionKey(date, mealType);
    var expanded = !!canteenExpandedCompleted[expandKey];
    if (visit && visit.completedAt && !expanded) return renderCanteenCompletedCard(visit, false);
    if (visit && visit.ticketIssuedAt && !visit.completedAt) return renderCanteenActiveTicket(visit);

    var items = asArray(meal && meal.items);
    var soups = items.filter(function (item) { return /sopa|creme|caldo/i.test(String(item.type || "") + " " + String(item.description || "")); });
    var mainDishes = items.filter(function (item) { return soups.indexOf(item) < 0; });
    var visibleSoups = visit ? soups : soups.filter(canteenItemVisible);
    var visibleMainEntries = mainDishes.map(function (item, index) { return { item: item, index: index }; }).filter(function (entry) { return visit || canteenItemVisible(entry.item); });
    var hiddenCount = visit ? 0 : (soups.length - visibleSoups.length) + (mainDishes.length - visibleMainEntries.length);
    var selection = canteenSelectionFor(date, mealType);
    var selectedDescription = visit && visit.dish && visit.dish.description;
    if (visit && expanded) {
      selection.dishIndex = mainDishes.findIndex(function (item) { return item.description === selectedDescription; });
      selection.dessertId = visit.dessert && visit.dessert.id || "";
    } else if (selection.dishIndex !== null && !canteenItemVisible(mainDishes[Number(selection.dishIndex)])) {
      selection.dishIndex = null;
    }
    var soupLines = visibleSoups.length ? visibleSoups.map(function (item) { return canteenDishCard(item, menu, { readonly: true }); }).join("") : (soups.length ? '<p class="campus-menu-empty">Sopa escondida pelos teus alergénios.</p>' : '<p class="campus-menu-empty">Sem sopa indicada.</p>');
    var veganDividerShown = false;
    var dishLines = visibleMainEntries.length ? visibleMainEntries.map(function (entry) {
      var item = entry.item;
      var index = entry.index;
      var divider = "";
      if (!veganDividerShown && canteenDishTone(item.type, item.description) === "vegetarian") {
        veganDividerShown = true;
        divider = '<div class="campus-menu-subdivider"><span>ALTERNATIVA VEGAN</span></div>';
      }
      return divider + canteenDishCard(item, menu, { interactive: !visit, readonly: !!visit, index: index, selected: visit ? item.description === selectedDescription : (selection.dishIndex !== null && Number(selection.dishIndex) === index) });
    }).join("") : (mainDishes.length ? '<p class="campus-menu-empty">As opções deste almoço foram escondidas pelos teus alergénios.</p>' : '<p class="campus-menu-empty">Sem pratos publicados.</p>');
    var allergenNotice = hiddenCount ? '<div class="campus-allergen-hidden-note"><i data-lucide="eye-off"></i><span>' + hiddenCount + ' ' + (hiddenCount === 1 ? 'opção escondida' : 'opções escondidas') + ' pelos teus alergénios.</span><button type="button" data-route="settings" data-tab="experience">Alterar</button></div>' : "";
    var ready = selection.dishIndex !== null && selection.dishIndex !== "" && Number.isInteger(Number(selection.dishIndex)) && Number(selection.dishIndex) >= 0 && !!selection.dessertId;
    var readonlyActions = visit && expanded
      ? '<div class="campus-readonly-actions"><button type="button" data-action="canteen-open-receipt" data-id="' + attr(visit.id) + '">Ver ticket</button><button type="button" data-action="canteen-collapse-completed" data-key="' + attr(expandKey) + '">Esconder ementa</button></div>'
      : '<button class="campus-order-button" type="button" data-action="canteen-issue-ticket" ' + (ready ? "" : "disabled") + '>Fazer pedido</button>';

    return '<div class="campus-dining-columns">' +
      '<div class="campus-dining-column campus-left-column">' + allergenNotice +
        '<section class="campus-menu-section campus-soups"><header><h3>SOPAS</h3><span>calorias</span></header><div>' + soupLines + '</div></section>' +
        '<section class="campus-menu-section campus-mains"><header><h3>PRATOS</h3><span>calorias</span></header><div>' + dishLines + '</div></section>' +
        '<section class="campus-menu-section campus-drinks"><header><h3>BEBIDAS</h3></header><div class="campus-drink-row">Água</div></section>' +
      '</div>' +
      '<div class="campus-dining-column campus-center-column">' +
        '<section class="campus-menu-section campus-desserts"><header><h3>SOBREMESAS</h3><span>calorias</span></header><div>' + canteenDessertChips(selection, !!visit, visit && visit.dessert && visit.dessert.id) + '</div></section>' +
        readonlyActions +
      '</div>' +
      renderCanteenPosterAINote(meal) +
      '</div>';
  }
  function renderCanteen() {
    setHeader("Cantina", "Campus · SAS NOVA");
    var theme = state.settings.canteenTheme === "leaf" ? "leaf" : "diner";
    var themeClass = theme === "leaf" ? " is-leaf-theme" : "";
    if (!canteenMenu && (canteenStatus === "idle" || canteenStatus === "loading")) {
      return '<section class="campus-dining-page' + themeClass + '"><div class="campus-dining-loading"><span></span><h2>A preparar a ementa</h2><p>A consultar a informação oficial da SAS NOVA.</p></div></section>';
    }
    if (!canteenMenu) {
      return '<section class="campus-dining-page' + themeClass + '"><div class="campus-dining-loading is-error"><h2>A ementa não chegou.</h2><p>' + esc(canteenError || "A fonte oficial está temporariamente indisponível.") + '</p><button type="button" data-action="refresh-canteen">Tentar novamente</button></div></section>';
    }
    var now = canteenPortugalParts(new Date());
    var hours = canteenServiceHours(canteenMenu);
    var service = canteenOpeningStatus(new Date(), hours);
    var days = asArray(canteenMenu.days).slice().sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
    if (!canteenSelectedDate || !days.some(function (day) { return day.date === canteenSelectedDate; })) {
      var bestDay = days.find(function (day) { return day.date === now.iso; }) || days.find(function (day) { return day.date >= now.iso; }) || days[days.length - 1];
      canteenSelectedDate = bestDay ? bestDay.date : "";
    }
    canteenMealTab = "lunch";
    var selected = days.find(function (day) { return day.date === canteenSelectedDate; }) || days[0];
    var selectedIsToday = selected && selected.date === now.iso;
    var meals = selected ? asArray(selected.meals) : [];
    var lunch = meals.find(function (meal) { return canteenMealType(meal) === "lunch"; });
    var statusLabel = selectedIsToday ? (service.open ? "ABERTO" : "FECHADO") : "EMENTA";
    var body = lunch
      ? renderCanteenMeal(lunch, canteenMenu, { date: selected.date, hours: hours, service: service, selectedIsToday: selectedIsToday })
      : '<div class="campus-dining-state"><div><small>SEM ALMOÇO</small><h3>Não há ementa publicada para este dia.</h3><p>Consulta outro dia no botão de informação.</p></div></div>';

    if (theme === "leaf") {
      var leafHero = '<header class="campus-dining-hero leaf-dining-hero">' +
        '<div class="campus-dining-brand leaf-dining-brand"><span class="leaf-menu-kicker"><i data-lucide="leaf"></i>TWENTY CAMPUS CAFÉ<i data-lucide="leaf"></i></span><h2>MENU DA<br>CANTINA</h2><p>PREPARADO HOJE NA FCT</p></div>' +
        '<div class="campus-dining-stamp ' + (service.open && selectedIsToday ? 'is-open' : '') + '"><i data-lucide="leaf"></i><span data-canteen-status-label>' + esc(statusLabel) + '</span><i data-lucide="leaf"></i></div>' +
        '<button class="campus-dining-info" type="button" data-action="canteen-open-info" aria-label="Informação da cantina">i</button>' +
        '</header>';
      return '<section class="campus-dining-page is-leaf-theme">' +
        '<div class="leaf-browser-shell">' +
          '<div class="leaf-browser-bar"><span class="leaf-window-dots"><i></i><i></i><i></i></span><span class="leaf-address-bar"></span><i data-lucide="leaf"></i></div>' +
          '<div class="leaf-browser-nav"><span>EMENTA</span><span>HOJE</span><span>FCT</span><span>REFEIÇÃO SOCIAL</span></div>' +
          '<div class="leaf-menu-surface">' + leafHero + body + '</div>' +
          '<footer class="leaf-menu-footer"><span><i data-lucide="leaf"></i><i data-lucide="leaf"></i><i data-lucide="leaf"></i></span><small>TWENTY · CAMPUS DINING</small></footer>' +
        '</div>' +
      '</section>';
    }

    return '<section class="campus-dining-page">' +
      '<header class="campus-dining-hero"><div class="campus-dining-brand"><h2>CAMPUS<br>DINING</h2><p>DINE IN CAFETERIA</p></div><div class="campus-dining-stamp ' + (service.open && selectedIsToday ? 'is-open' : '') + '">' + esc(statusLabel) + '</div><button class="campus-dining-info" type="button" data-action="canteen-open-info" aria-label="Informação da cantina">i</button></header>' +
      body +
      '</section>';
  }
  function settingsNavButton(id, icon, label, active) {
    return '<button type="button" class="settings-nav-button' + (active === id ? ' is-active' : '') + '" data-action="settings-section" data-section="' + attr(id) + '"><i data-lucide="' + attr(icon) + '"></i><span>' + esc(label) + '</span><i data-lucide="chevron-right"></i></button>';
  }

  function renderSettings() {
    setHeader("Definições", "Twenty · controlo do sistema");
    var section = ["overview", "academic", "data", "experience", "developer"].indexOf(route.tab) >= 0 ? route.tab : "overview";
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
    var forceControls = '<span class="sync-force-actions" role="group" aria-label="Substituição manual de dados"><button class="button button-small sync-icon-button" type="button" data-action="force-git-pull" aria-label="Forçar pull" title="Forçar pull"' + forceDisabled + '><i data-lucide="arrow-down-to-line"></i></button><button class="button button-dark button-small sync-icon-button" type="button" data-action="force-git-push" aria-label="Forçar push" title="Forçar push"' + forceDisabled + '><i data-lucide="arrow-up-to-line"></i></button></span>';
    var syncProgress = '<div id="gitSyncInlineProgress" class="sync-inline-progress ' + (syncInfo.state === "syncing" ? "is-active is-indeterminate" : "") + '"><span></span></div>';
    var syncCard = '<article id="gitSyncCard" class="card settings-card card-violet" aria-busy="' + (syncInfo.state === "syncing" ? "true" : "false") + '"><div class="card-title-row"><div><p class="card-label">Git como base de dados</p><h3 id="gitSyncTitle">' + esc(syncDisplay.title) + '</h3><p id="gitSyncDetail" class="card-subtitle">' + esc(syncDisplay.detail) + '</p></div><span class="metric-icon"><i data-lucide="git-commit-horizontal"></i></span></div><div class="settings-row"><div><strong id="gitSyncSummary">' + esc(syncVersionSummary) + '</strong><small>Fusão por campos · fila offline · histórico em commits.</small></div><span id="gitSyncBadge" class="badge ' + syncDisplay.badgeClass + '">' + esc(syncVersionBadge) + '</span></div>' + syncProgress + '<div class="list-actions"><button class="button button-small" type="button" data-action="configure-git-sync"><i data-lucide="settings-2"></i>Configurar</button><button class="button button-dark button-small" type="button" data-action="sync-now"><i data-lucide="refresh-cw"></i>Sincronizar agora</button>' + forceControls + (syncInfo.configured ? '<button class="button button-small" type="button" data-action="disable-git-sync"><i data-lucide="pause"></i>Pausar</button>' : '') + '</div></article>';
    var debugSnapshot = homeDebugSnapshot();
    var debugCard = '<article class="card settings-card debug-admin-card"><div class="card-title-row"><div><p class="card-label">Ferramentas de desenvolvimento</p><h3>Laboratório da Home</h3><p class="card-subtitle">Simula manhã, aulas coladas, TPC e Report Card sem esperar pela hora real.</p></div><span class="metric-icon"><i data-lucide="flask-conical"></i></span></div><div class="settings-row"><div><strong>' + esc(debugSnapshot.label) + '</strong><small>' + esc(debugSnapshot.time) + ' · motor: ' + esc(debugSnapshot.phase) + '</small></div><span class="badge ' + (homeDebug && homeDebug.active ? "badge-yellow" : "badge-mint") + '">' + (homeDebug && homeDebug.active ? "Simulação ativa" : "Dados reais") + '</span></div><div class="list-actions"><button class="button button-dark button-small" type="button" data-action="debug-start-tutorial"><i data-lucide="play"></i>Tutorial do dia</button><button class="button button-small" type="button" data-action="debug-open-lab"><i data-lucide="bug"></i>Abrir debug</button>' + (homeDebug && homeDebug.active ? '<button class="button button-danger button-small" type="button" data-action="debug-exit"><i data-lucide="x"></i>Sair</button>' : '') + '</div></article>';
    var profileCard = '<article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Perfil académico</p><h3>' + esc(state.profile.name || "Estudante") + '</h3><p class="card-subtitle">' + esc(state.profile.degree || "Curso por configurar") + (state.profile.institution ? " · " + esc(state.profile.institution) : "") + '</p></div><span class="metric-icon"><i data-lucide="user-round"></i></span></div><div class="settings-row"><div><strong>Meta académica</strong><small>Usada nos indicadores de desempenho.</small></div><span class="badge badge-yellow">' + (Number(state.profile.targetGrade) || 20) + '/20</span></div><button class="button button-small" type="button" data-action="edit-profile"><i data-lucide="pencil"></i>Editar perfil</button></article>';
    var semesterCard = '<article class="card settings-card card-violet"><div class="card-title-row"><div><p class="card-label">Semestre ativo</p><h3>' + esc(semester ? semester.name : "Nenhum") + '</h3><p class="card-subtitle">' + esc(semester ? semester.academicYear : "Cria o próximo semestre") + '</p></div><span class="metric-icon"><i data-lucide="calendar-range"></i></span></div><div class="settings-row"><div><strong>' + activeCourses().length + ' cadeiras</strong><small>' + archived + ' semestre(s) no arquivo.</small></div><span class="badge badge-dark">' + activeCourses().reduce(function (sum, course) { return sum + (Number(course.ects) || 0); }, 0) + ' ECTS</span></div><div class="list-actions"><button class="button button-small" type="button" data-action="new-semester"><i data-lucide="calendar-plus"></i>Novo</button>' + (semester ? '<button class="button button-danger button-small" type="button" data-action="archive-semester"><i data-lucide="archive"></i>Arquivar</button>' : "") + '</div></article>';
    var quickCard = '<article class="card settings-card card-dark span-12"><div class="card-title-row"><div><p class="card-label">Criação rápida</p><h3>Adicionar conteúdo</h3><p class="card-subtitle">Atalhos para o que normalmente configuras no início da semana.</p></div><span class="metric-icon"><i data-lucide="wand-sparkles"></i></span></div><div class="quick-grid"><button type="button" data-action="create-lesson"><i data-lucide="presentation"></i>Aula</button><button type="button" data-action="add-material"><i data-lucide="file-up"></i>Material</button><button type="button" data-action="add-question"><i data-lucide="message-circle-question"></i>Pergunta</button><button type="button" data-action="add-quiz"><i data-lucide="circle-check-big"></i>Quiz</button><button type="button" data-action="add-grade"><i data-lucide="chart-no-axes-combined"></i>Nota</button><button type="button" data-action="add-assessment"><i data-lucide="file-pen-line"></i>Avaliação</button></div></article>';
    var jsonCard = '<article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Ficheiro local</p><h3>academic-data.json</h3><p class="card-subtitle">Importação e exportação manual, separada do Git.</p></div><span class="metric-icon"><i data-lucide="braces"></i></span></div><div class="settings-row"><div><strong>Última verificação</strong><small>' + esc(lastCheck) + ' · revisão local ' + (Number(state.meta.revision) || 0) + '</small></div><button class="switch ' + (state.settings.jsonSync ? "is-on" : "") + '" type="button" data-action="toggle-json-sync" aria-label="Ativar sincronização JSON"><span></span></button></div><div class="list-actions"><button class="button button-small" type="button" data-action="reload-json"><i data-lucide="refresh-cw"></i>Reler</button><button class="button button-small" type="button" data-action="export-json"><i data-lucide="download"></i>Exportar</button><button class="button button-small" type="button" data-action="import-json"><i data-lucide="upload"></i>Importar</button></div></article>';
    var storageCard = '<article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Neste dispositivo</p><h3>Armazenamento local</h3><p class="card-subtitle">Documentos, imagens e cache usados por esta instalação.</p></div><span class="metric-icon"><i data-lucide="hard-drive"></i></span></div><div class="settings-row"><div><strong id="storageFileCount">A contar ficheiros…</strong><small>Ficheiros guardados no browser.</small></div><span class="badge badge-mint">Local-first</span></div><button class="button button-small" type="button" data-action="export-json"><i data-lucide="shield-check"></i>Criar backup JSON</button></article>';
    var canteenAIStatus = canteenAIAvailabilityLabel();
    var canteenAICard = '<article class="card settings-card settings-feature-card"><div class="card-title-row"><div><p class="card-label">Cantina</p><h3>Brilho com IA local</h3><p class="card-subtitle">Chef’s Note contextual, escolha do chef e badges úteis. A ementa, as kcal e os alergénios oficiais nunca são alterados.</p></div><span class="metric-icon"><i data-lucide="sparkles"></i></span></div><div class="settings-row"><div><strong>IA integrada do Chrome</strong><small>Quando não existe, a Twenty não inventa uma Chef’s Note.</small></div><span class="' + canteenAIStatus.className + '">' + esc(canteenAIStatus.label) + '</span></div><div class="settings-toggle-list"><div class="settings-row"><div><strong>Ativar enriquecimento</strong><small>Permite escrever uma Chef’s Note com a IA local. Os badges factuais continuam determinísticos.</small></div><button class="switch ' + (state.settings.canteenAIEnabled ? "is-on" : "") + '" type="button" data-action="toggle-canteen-ai"><span></span></button></div><div class="settings-row"><div><strong>Chef\'s Note</strong><small>Escolhe primeiro o prato e depois escreve uma nota curta ligada à ementa, ao horário e às avaliações do dia.</small></div><button class="switch ' + (state.settings.canteenAIChefNote ? "is-on" : "") + '" type="button" data-action="toggle-canteen-ai-note"><span></span></button></div><div class="settings-row"><div><strong>Recomendação e badges</strong><small>Mostra a escolha do chef, a opção mais energética e fontes de proteína identificáveis pelo nome.</small></div><button class="switch ' + (state.settings.canteenAIDescriptions ? "is-on" : "") + '" type="button" data-action="toggle-canteen-ai-descriptions"><span></span></button></div></div><div class="list-actions"><button class="button button-dark button-small" type="button" data-route="canteen"><i data-lucide="utensils"></i>Abrir Cantina</button><button class="button button-small" type="button" data-action="canteen-ai-clear-cache"><i data-lucide="eraser"></i>Limpar cache IA</button></div></article>';
    var canteenTheme = state.settings.canteenTheme === "leaf" ? "leaf" : "diner";
    var canteenThemeCard = '<article class="card settings-card canteen-theme-settings-card"><div class="card-title-row"><div><p class="card-label">Cantina</p><h3>Estilo da ementa</h3><p class="card-subtitle">Alterna o visual sem mudar a ementa, a IA, os filtros, o ticket ou o estado da refeição.</p></div><span class="metric-icon"><i data-lucide="palette"></i></span></div><div class="canteen-theme-picker" role="radiogroup" aria-label="Tema da Cantina"><button class="canteen-theme-option' + (canteenTheme === "diner" ? ' is-selected' : '') + '" type="button" data-action="canteen-set-theme" data-theme="diner" role="radio" aria-checked="' + (canteenTheme === "diner") + '"><span class="canteen-theme-preview is-diner"><i></i><b>CAMPUS<br>DINING</b><em>ABERTO</em></span><strong>Diner</strong><small>O visual atual, a preto e branco.</small></button><button class="canteen-theme-option' + (canteenTheme === "leaf" ? ' is-selected' : '') + '" type="button" data-action="canteen-set-theme" data-theme="leaf" role="radio" aria-checked="' + (canteenTheme === "leaf") + '"><span class="canteen-theme-preview is-leaf"><i data-lucide="leaf"></i><b>MENU DA<br>CANTINA</b><em>FCT CAFÉ</em></span><strong>Leaf</strong><small>Papel quente, verdes suaves e detalhes naturais.</small></button></div></article>';
    var allergenSettings = canteenAllergenSettings();
    var allergenMap = canteenAllergenMap();
    var allergenOptions = Object.keys(CANTEEN_DEFAULT_ALLERGENS).sort(function (a, b) { return Number(a) - Number(b); }).map(function (id) {
      var selected = allergenSettings.selected.indexOf(id) >= 0;
      return '<button class="settings-allergen-option' + (selected ? ' is-selected' : '') + '" type="button" data-action="canteen-toggle-allergen" data-allergen="' + attr(id) + '" aria-pressed="' + selected + '"><span>' + esc(id) + '</span><strong>' + esc(allergenMap[id] || CANTEEN_DEFAULT_ALLERGENS[id]) + '</strong><i data-lucide="' + (selected ? 'check' : 'plus') + '"></i></button>';
    }).join("");
    var allergenCountLabel = allergenSettings.selected.length ? allergenSettings.selected.length + ' selecionado' + (allergenSettings.selected.length === 1 ? '' : 's') : 'Nenhum selecionado';
    var canteenAllergenCard = '<article class="card settings-card canteen-allergen-settings-card"><div class="card-title-row"><div><p class="card-label">Cantina</p><h3>Alergénios</h3><p class="card-subtitle">Escolhe os códigos que queres evitar. A Twenty esconde apenas opções que tenham esses alergénios indicados na ementa oficial.</p></div><span class="metric-icon"><i data-lucide="shield-alert"></i></span></div><div class="settings-row"><div><strong>Esconder opções incompatíveis</strong><small>Também impede a Chef’s Note de recomendar esses pratos.</small></div><button class="switch ' + (allergenSettings.hideDishes ? 'is-on' : '') + '" type="button" data-action="canteen-toggle-hide-allergens" aria-label="Esconder pratos com alergénios selecionados"><span></span></button></div><div class="settings-allergen-summary"><span class="badge ' + (allergenSettings.selected.length ? 'badge-yellow' : 'badge-mint') + '">' + esc(allergenCountLabel) + '</span><button type="button" data-action="canteen-clear-allergens" ' + (allergenSettings.selected.length ? '' : 'disabled') + '>Limpar seleção</button></div><div class="settings-allergen-grid">' + allergenOptions + '</div><p class="settings-allergen-note">O filtro depende dos códigos publicados pela cantina. As sobremesas não são filtradas porque a fonte oficial não publica os respetivos códigos. Confirma sempre a informação com a unidade, sobretudo em casos de alergia grave.</p></article>';
    var motionCard = '<article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Interface</p><h3>Movimento e conforto</h3><p class="card-subtitle">Controla animações sem perder a informação importante.</p></div><span class="metric-icon"><i data-lucide="accessibility"></i></span></div><div class="settings-row"><div><strong>Reduzir movimento</strong><small>Desativa transições e celebrações mais intensas.</small></div><button class="switch ' + (state.settings.reduceMotion ? "is-on" : "") + '" type="button" data-action="toggle-reduce-motion"><span></span></button></div></article>';
    var campusCard = '<article class="card settings-card card-yellow"><div class="card-title-row"><div><p class="card-label">Home</p><h3>Atividade simulada</h3><p class="card-subtitle">Mostra indicadores de presença no campus claramente assinalados como simulação.</p></div><span class="metric-icon"><i data-lucide="users-round"></i></span></div><div class="settings-row"><div><strong>Contador simulado</strong><small>É apenas ambiente visual; não representa utilizadores reais.</small></div><button class="switch ' + (state.settings.campusSimulation ? "is-on" : "") + '" type="button" data-action="toggle-campus"><span></span></button></div></article>';
    var tutorialCard = '<article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Ajuda</p><h3>Aprender a Twenty</h3><p class="card-subtitle">Revê a visita guiada ou testa um dia escolar completo.</p></div><span class="metric-icon"><i data-lucide="map"></i></span></div><div class="list-actions"><button class="button button-dark button-small" type="button" data-action="show-tutorial"><i data-lucide="map"></i>Visita guiada</button><button class="button button-small" type="button" data-action="debug-start-tutorial"><i data-lucide="play"></i>Simular dia</button></div></article>';
    var planningCard = '<article class="card settings-card"><div class="card-title-row"><div><p class="card-label">Rotina</p><h3>Planeamento de estudo</h3><p class="card-subtitle">' + Number(state.settings.weeklyStudyHours || 16) + ' h/semana · sessões de ' + Number(state.settings.studySessionMinutes || 50) + ' min.</p></div><span class="metric-icon"><i data-lucide="timer-reset"></i></span></div><div class="settings-row"><div><strong>' + esc(state.settings.studyDayStart || "09:00") + '–' + esc(state.settings.studyDayEnd || "19:00") + '</strong><small>Pausa de ' + Number(state.settings.studyBreakMinutes || 10) + ' minutos entre sessões.</small></div><span class="badge badge-violet">Plano ativo</span></div><button class="button button-small" type="button" data-action="study-planner-settings"><i data-lucide="sliders-horizontal"></i>Configurar rotina</button></article>';
    var safetyCard = '<article class="card settings-card span-12 settings-danger-zone"><div class="card-title-row"><div><p class="card-label">Zona de segurança</p><h3>Recomeçar neste dispositivo</h3><p class="card-subtitle">Remove o estado local e os ficheiros deste browser. O repositório Git não é apagado.</p></div><button class="button button-danger" type="button" data-action="reset-app"><i data-lucide="trash-2"></i>Apagar dados locais</button></div></article>';
    var systemCard = '<article class="card settings-card settings-system-card"><div class="card-title-row"><div><p class="card-label">Resumo</p><h3>Estado da Twenty</h3><p class="card-subtitle">Um olhar rápido antes de entrares nos detalhes.</p></div><span class="metric-icon"><i data-lucide="gauge"></i></span></div><div class="settings-system-grid"><div><span>Git</span><strong>' + esc(syncVersionBadge) + '</strong></div><div><span>Semestre</span><strong>' + activeCourses().length + ' cadeiras</strong></div><div><span>IA Cantina</span><strong>' + esc(canteenAIStatus.label) + '</strong></div><div><span>Dados</span><strong>Revisão ' + (Number(state.meta.revision) || 0) + '</strong></div></div></article>';

    var titles = {
      overview: ["Visão geral", "O essencial da tua conta académica e do sistema."],
      academic: ["Académico", "Perfil, semestre, rotina e criação de conteúdo."],
      data: ["Dados e sincronização", "Git, JSON, armazenamento e backups num só lugar."],
      experience: ["Experiência", "Cantina inteligente, movimento, ambiente e ajuda."],
      developer: ["Laboratório", "Debug, simulação e ferramentas avançadas."]
    };
    var content = {
      overview: systemCard + profileCard + semesterCard + quickCard,
      academic: profileCard + semesterCard + planningCard + quickCard,
      data: syncCard + jsonCard + storageCard + safetyCard,
      experience: canteenThemeCard + canteenAICard + canteenAllergenCard + motionCard + campusCard + tutorialCard,
      developer: debugCard + jsonCard + storageCard + safetyCard
    }[section];
    var nav = settingsNavButton("overview", "layout-dashboard", "Visão geral", section) + settingsNavButton("academic", "graduation-cap", "Académico", section) + settingsNavButton("data", "database", "Dados e sync", section) + settingsNavButton("experience", "sparkles", "Experiência", section) + settingsNavButton("developer", "flask-conical", "Laboratório", section);
    return '<div class="page-head settings-page-head"><div><p class="settings-eyebrow">Twenty control room</p><h2>Definições</h2><p>Organizadas por contexto, não por ordem aleatória de cartões.</p></div><div class="page-actions"><button class="button button-dark" type="button" data-action="quick-add"><i data-lucide="plus"></i>Adicionar conteúdo</button></div></div><div class="settings-shell"><aside class="settings-nav"><div class="settings-nav-title"><span><i data-lucide="sliders-horizontal"></i></span><div><strong>Definições</strong><small>Escolhe uma área</small></div></div>' + nav + '</aside><section class="settings-panel"><header class="settings-section-head"><div><span>' + esc(titles[section][0]) + '</span><h3>' + esc(titles[section][0]) + '</h3><p>' + esc(titles[section][1]) + '</p></div><i data-lucide="' + (section === "overview" ? "layout-dashboard" : section === "academic" ? "graduation-cap" : section === "data" ? "database" : section === "experience" ? "sparkles" : "flask-conical") + '"></i></header><div class="settings-grid">' + content + '</div></section></div>';
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
    hydrateNotebookImages(modalRoot);
    enhanceNotebookStickers(modalRoot);
    typesetMath(modalRoot);
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
    hydrateNotebookImages(view);
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

  function safeNotebookImageUrl(value) {
    var url = String(value || "").trim();
    if (!url) return "";
    if (/^data:image\/(png|jpe?g|gif|webp);base64,/i.test(url) && url.length <= 18 * 1024 * 1024) return url;
    if (/^blob:/i.test(url)) return url;
    return safeResourceUrl(url);
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
        var taskLessonId = data.get("lessonId") || null;
        var taskTypeValue = data.get("taskType") || "homework";
        if (taskLessonId && (taskTypeValue === "homework" || taskTypeValue === "tpc") && homeworkForLesson(taskLessonId)) throw new Error("Esta aula já tem um TPC configurado.");
        state.tasks.push({ id: uid("task"), semesterId: state.currentSemesterId, courseId: data.get("courseId") || null, lessonId: taskLessonId, title: String(data.get("title") || "").trim(), type: taskTypeValue, dueDate: data.get("dueDate") || "", dueTime: data.get("dueTime") || "", priority: data.get("priority") || "normal", done: false, createdAt: new Date().toISOString() });
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
          if (quizLesson && configuredQuizForLesson(quizLesson.id)) throw new Error("Esta aula já tem um quiz configurado.");
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
        lessonNotes.notesHtml = sanitizeNotebookHTML("<p>" + nl2br(lessonNotes.notes) + "</p>");
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
    if (quiz.lessonId && lessonIsBeOnline(lessonById(quiz.lessonId))) {
      viewLessonQuiz(quiz.lessonId);
      return;
    }
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
      var quizOptions = asArray(question.optionBlocks).length ? question.optionBlocks : asArray(question.options);
      body = progress + '<div class="quiz-question">' + renderContentBlocks(question.promptBlocks || question.prompt) + '</div>' + renderImageGallery(question.images, "question", { ownerId: question.sourceQuestionId || "" }) + '<div class="quiz-options">' + quizOptions.map(function (option, index) {
        return '<button class="quiz-option ' + (selected === index ? "is-selected" : "") + '" type="button" data-action="quiz-answer" data-index="' + index + '"><span>' + String.fromCharCode(65 + index) + '</span>' + renderQuizOptionContent(option) + "</button>";
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
    quiz.completedOnce = true;
    quiz.lastCompletedAt = new Date().toISOString();
    if (quiz.lessonId) completeLessonBeOnline(quiz.lessonId);
    ensureBeOnlineTasks();
    await save(true);
    quizRuntime = null;
    render();
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
    if ((task.type === "homework" || task.type === "tpc") && asArray(task.contentBlocks).length) { viewLessonHomework(task.id); return; }
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
      { route: "home", selector: ".school-now-card", page: "Hoje", title: "O que está a acontecer agora", copy: "A Home escolhe um próximo passo: preparar, acompanhar a aula, fazer o Quiz da aula, tratar do TPC ou fechar o dia." },
      { route: "home", selector: ".after-school-card", page: "Hoje", title: "Quizzes da aula antes do TPC", copy: "A verificação rápida fecha cada aula. O trabalho de casa fica separado para a sessão depois das aulas." },
      { route: "home", selector: ".daily-report-card", page: "Hoje", title: "Report Card diário", copy: "A nota académica vem dos Quizzes da aula e a rotina mostra o trabalho escolar concluído." },
      { route: "courses", selector: ".course-grid, .empty-state", page: "Cadeiras", title: "Cadeiras do semestre", copy: "Abre uma cadeira para consultar aulas, materiais, avaliações, perguntas, quizzes e notas." },
      { route: "planner", plannerMode: "calendar", selector: ".planner-mode-control", page: "Calendário", title: "Horário ou calendário", copy: "O Horário guarda os blocos recorrentes. O Calendário combina aulas, testes, eventos e tarefas com data." },
      { route: "planner", plannerMode: "calendar", selector: ".calendar-view-control", page: "Calendário", title: "Quatro vistas", copy: "Alterna entre Dia, 3 dias, Semana e Mês. As setas avançam exatamente o intervalo selecionado." },
      { route: "planner", plannerMode: "calendar", selector: ".calendar-card", page: "Calendário", title: "Agenda académica", copy: "Uma aula preparada aparece pelo nome. Os blocos ainda não preparados continuam visíveis através do horário." },
      { route: "planner", plannerMode: "study-day", selector: ".study-day-shell", page: "Dia de estudo", title: "Blocos de tempo", copy: "Arrasta tarefas, aulas, quizzes e avaliações para uma hora. Em ecrãs táteis, usa o botão de agendar." },
      { route: "study", selector: ".page-head", page: "Estudar", title: "Centro de estudo", copy: "Aqui encontras aulas por rever, perguntas de testes anteriores e todos os quizzes disponíveis." },
      { route: "study", selector: ".study-hours-card", page: "Estudar", title: "Horas por cadeira", copy: "A estimativa distribui as horas semanais por ECTS, avaliações próximas e trabalho pendente." },
      { route: "grades", selector: ".target-card", page: "Notas", title: "Média ECTS", copy: "A média global usa a nota atual de cada cadeira e os respetivos ECTS. Cada nota mantém a avaliação de origem." },
      { route: "canteen", selector: ".canteen-days, .canteen-loading", page: "Cantina", title: "Ementa oficial", copy: "Escolhe o dia e confirma quando a informação da SAS NOVA foi consultada." },
      { route: "canteen", selector: ".canteen-menu-board, .canteen-loading", page: "Cantina", title: "Escolher e registar a refeição", copy: "Escolhe almoço ou jantar, seleciona prato e sobremesa e guarda um talão no teu dia." },
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
      await closeModalSavingNotebook();
      if (onboarding) renderOnboarding();
    } else if (action === "view-report-card") {
      closeModal();
      setRoute("home");
      setTimeout(function () {
        var card = document.querySelector(".daily-report-card");
        if (!card) return;
        card.classList.add("is-highlighted");
        card.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(function () { card.classList.remove("is-highlighted"); }, 1800);
      }, 80);
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
    } else if (action === "configure-lesson-content") {
      if (modalRoot.querySelector("#notebookForm")) await saveNotebook({ close: false, render: false, silent: true });
      openLessonBuilder(button.dataset.lesson || route.id, button.dataset.kind || "quiz");
    } else if (action === "copy-lesson-builder-prompt") {
      var promptArea = modalRoot.querySelector("#lessonBuilderPrompt");
      if (promptArea) { await copyText(promptArea.value); toast("Prompt copiado."); }
    } else if (action === "import-lesson-builder") {
      await importLessonBuilder();
    } else if (action === "view-lesson-quiz") {
      viewLessonQuiz(button.dataset.lesson || route.id);
    } else if (action === "view-lesson-homework") {
      viewLessonHomework(button.dataset.id);
    } else if (action === "open-notebook-editor") {
      openNotebookEditor(button.dataset.lesson || route.id);
    } else if (action === "save-notebook") {
      await saveNotebook({ close: true, render: true, silent: false });
    } else if (action === "notebook-command") {
      var noteEditor = modalRoot.querySelector("#notebookEditor");
      if (noteEditor) { noteEditor.focus(); document.execCommand(button.dataset.command, false, button.dataset.value || null); }
    } else if (action === "notebook-block") {
      var blockEditor = modalRoot.querySelector("#notebookEditor");
      if (blockEditor) { blockEditor.focus(); document.execCommand("formatBlock", false, button.dataset.block || "p"); }
    } else if (action === "notebook-add-image") {
      var notebookImageInput = modalRoot.querySelector("#notebookImageInput");
      if (notebookImageInput) notebookImageInput.click();
    } else if (action === "notebook-sticker-delete") {
      var deleteSticker = button.closest(".notebook-sticker");
      if (deleteSticker) deleteSticker.remove();
    } else if (action === "notebook-sticker-smaller" || action === "notebook-sticker-larger") {
      var resizeSticker = button.closest(".notebook-sticker");
      var resizeImage = resizeSticker && resizeSticker.querySelector("img");
      if (resizeSticker && resizeImage) {
        var currentWidth = clamp(resizeSticker.getAttribute("data-width") || resizeImage.getAttribute("width") || 280, 80, 720);
        var nextWidth = clamp(currentWidth + (action === "notebook-sticker-larger" ? 40 : -40), 80, 720);
        resizeSticker.setAttribute("data-width", String(nextWidth));
        resizeImage.setAttribute("width", String(nextWidth));
        resizeImage.setAttribute("data-width", String(nextWidth));
      }
    } else if (action === "notebook-sticker-align") {
      var alignSticker = button.closest(".notebook-sticker");
      if (alignSticker) {
        alignSticker.classList.remove("align-left", "align-center", "align-right");
        alignSticker.classList.add("align-" + (["left", "center", "right"].indexOf(button.dataset.align) >= 0 ? button.dataset.align : "center"));
      }
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
      var pastTargetQuiz = state.quizzes.find(function (item) { return item.id === button.dataset.id; });
      if (pastTargetQuiz && (pastTargetQuiz.lessonId || pastTargetQuiz.lockedContent)) {
        if (pastTargetQuiz.lessonId) viewLessonQuiz(pastTargetQuiz.lessonId);
        toast("O quiz desta aula está bloqueado depois de configurado.");
      } else {
        openPastQuestionPicker(button.dataset.id);
      }
    } else if (action === "add-quiz-question") {
      var manualTargetQuiz = state.quizzes.find(function (item) { return item.id === button.dataset.id; });
      if (manualTargetQuiz && (manualTargetQuiz.lessonId || manualTargetQuiz.lockedContent)) {
        if (manualTargetQuiz.lessonId) viewLessonQuiz(manualTargetQuiz.lessonId);
        toast("O quiz desta aula está bloqueado depois de configurado.");
      } else {
        openEntityForm("quiz-question", { quizId: button.dataset.id });
      }
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
      } else if (task && (task.type === "homework" || task.type === "tpc") && (task.lockedContent || task.configuredFromPrompt)) {
        if (task.done || task.completedOnce) {
          viewLessonHomework(task.id);
        } else {
          closeModal();
          startHomeworkSession(task.id);
        }
      } else if (task) {
        task.done = !task.done;
        await save(true);
        closeModal();
        render();
        toast(task.done ? "Tarefa concluída." : "Tarefa reaberta.");
      }
    } else if (action === "start-homework-session") {
      closeModal();
      startHomeworkSession(button.dataset.task || "");
    } else if (action === "homework-select") {
      setHomeworkCurrent(button.dataset.id || "", true);
      render();
    } else if (action === "homework-pause") {
      storeHomeworkElapsed();
      if (homeworkSessionRuntime) { homeworkSessionRuntime.running = false; homeworkSessionRuntime.startedAt = null; }
      render();
    } else if (action === "homework-resume") {
      if (homeworkSessionRuntime) { homeworkSessionRuntime.running = true; homeworkSessionRuntime.startedAt = Date.now(); }
      render();
    } else if (action === "homework-finish") {
      await finishHomeworkTask();
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
    } else if (action === "canteen-open-info") {
      showCanteenInfo();
    } else if (action === "canteen-ai-prepare") {
      await ensureCanteenAIForCurrentMeal(true);
    } else if (action === "canteen-ai-clear-cache") {
      clearCanteenAICache();
      render();
      toast("Cache da IA da Cantina limpa.");
    } else if (action === "canteen-toggle-allergen") {
      var allergenId = String(button.dataset.allergen || "");
      if (Object.prototype.hasOwnProperty.call(CANTEEN_DEFAULT_ALLERGENS, allergenId)) {
        var allergenConfig = state.settings.canteenAllergenFilters || { selected: [], hideDishes: true };
        var selectedAllergens = asArray(allergenConfig.selected).map(String);
        if (selectedAllergens.indexOf(allergenId) >= 0) selectedAllergens = selectedAllergens.filter(function (id) { return id !== allergenId; });
        else selectedAllergens.push(allergenId);
        state.settings.canteenAllergenFilters = { selected: selectedAllergens, hideDishes: allergenConfig.hideDishes !== false };
        clearCanteenAICache();
        await save(true);
        render();
      }
    } else if (action === "canteen-toggle-hide-allergens") {
      var hideConfig = state.settings.canteenAllergenFilters || { selected: [], hideDishes: true };
      state.settings.canteenAllergenFilters = { selected: asArray(hideConfig.selected).map(String), hideDishes: hideConfig.hideDishes === false };
      clearCanteenAICache();
      await save(true);
      render();
    } else if (action === "canteen-clear-allergens") {
      var clearConfig = state.settings.canteenAllergenFilters || { selected: [], hideDishes: true };
      state.settings.canteenAllergenFilters = { selected: [], hideDishes: clearConfig.hideDishes !== false };
      clearCanteenAICache();
      await save(true);
      render();
      toast("Seleção de alergénios limpa.");
    } else if (action === "canteen-day") {
      canteenSelectedDate = button.dataset.date || canteenSelectedDate;
      canteenMealTab = "lunch";
      if (modalRoot && modalRoot.innerHTML) closeModal();
      render();
    } else if (action === "canteen-meal-tab") {
      canteenMealTab = button.dataset.meal === "dinner" ? "dinner" : "lunch";
      render();
    } else if (action === "canteen-select-dish") {
      var canteenDishSelection = canteenSelectionFor(canteenSelectedDate, canteenMealTab || "lunch");
      var nextDishIndex = Number(button.dataset.index);
      canteenDishSelection.dishIndex = Number(canteenDishSelection.dishIndex) === nextDishIndex ? null : nextDishIndex;
      render();
    } else if (action === "canteen-select-dessert") {
      var canteenDessertSelection = canteenSelectionFor(canteenSelectedDate, canteenMealTab || "lunch");
      var nextDessertId = button.dataset.dessert || "";
      canteenDessertSelection.dessertId = canteenDessertSelection.dessertId === nextDessertId ? "" : nextDessertId;
      render();
    } else if (action === "canteen-issue-ticket") {
      var ticketDraft = canteenTicketDraft(canteenSelectedDate, canteenMealTab || "lunch");
      if (!ticketDraft) {
        toast("Escolhe um prato e uma sobremesa antes de emitir o ticket.", "warning");
      } else {
        var ticketSummary = '<div class="diner-campus-card-confirm"><span class="diner-campus-card-icon"><i data-lucide="credit-card"></i></span><div><small>CARTÃO DO CAMPUS</small><h3>Validar ' + esc(ticketDraft.mealLabel.toLowerCase()) + '</h3><p>' + esc(ticketDraft.dish.description) + ' · ' + esc(ticketDraft.dessert.label) + '</p><strong>' + esc(ticketDraft.price) + '</strong></div></div>';
        openModal("Confirmar pedido", ticketSummary, { className: "canteen-ticket-confirm-modal", footer: '<footer class="modal-foot"><button class="campus-modal-secondary" type="button" data-action="close-modal">Voltar</button><button class="campus-order-button" type="button" data-action="canteen-confirm-ticket">Confirmar pedido</button></footer>' });
      }
    } else if (action === "canteen-confirm-ticket") {
      var confirmedTicket = canteenTicketDraft(canteenSelectedDate, canteenMealTab || "lunch");
      if (!confirmedTicket) {
        closeModal();
        toast("A seleção já não está disponível. Volta a escolher o menu.", "warning");
      } else {
        confirmedTicket.ticketIssuedAt = new Date().toISOString();
        confirmedTicket.orderNumber = canteenTicketOrderNumber(confirmedTicket.date, confirmedTicket.mealType);
        var confirmedIndex = state.canteenVisits.findIndex(function (visit) { return visit.id === confirmedTicket.id; });
        if (confirmedIndex >= 0) state.canteenVisits[confirmedIndex] = confirmedTicket;
        else state.canteenVisits.push(confirmedTicket);
        closeModal();
        await save(true);
        canteenPrinterFeedback();
        render();
        showCanteenReceipt(confirmedTicket, { printing: true });
        toast("Ticket de " + confirmedTicket.mealLabel.toLowerCase() + " emitido.");
      }
    } else if (action === "canteen-finish-meal") {
      var finishedVisit = state.canteenVisits.find(function (visit) { return visit.id === button.dataset.id; });
      if (finishedVisit && !finishedVisit.completedAt) {
        finishedVisit.completedAt = new Date().toISOString();
        canteenExpandedCompleted[canteenSelectionKey(finishedVisit.date, finishedVisit.mealType)] = false;
        await save(true);
        canteenFeedback();
        render();
        showCanteenReceipt(finishedVisit, { stamped: true });
        toast(finishedVisit.mealLabel + " concluído. Bom apetite!");
      }
    } else if (action === "canteen-open-receipt") {
      var receiptVisit = state.canteenVisits.find(function (visit) { return visit.id === button.dataset.id; });
      if (receiptVisit) showCanteenReceipt(receiptVisit);
    } else if (action === "canteen-expand-completed") {
      canteenExpandedCompleted[button.dataset.key || ""] = true;
      render();
    } else if (action === "canteen-collapse-completed") {
      canteenExpandedCompleted[button.dataset.key || ""] = false;
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
    } else if (action === "debug-open-lab") {
      openHomeDebugLab();
    } else if (action === "debug-start-tutorial") {
      startHomeDebugTutorial();
    } else if (action === "debug-start-scenario") {
      var chosenScenario = Number(button.dataset.index || 0);
      if (!homeDebug || !homeDebug.active || !homeDebug.tutorial) startHomeDebugTutorial();
      prepareHomeDebugScenario(chosenScenario);
    } else if (action === "debug-apply-time") {
      var debugInput = modalRoot.querySelector("#debugCustomTime");
      startHomeDebugAt(debugInput && debugInput.value);
    } else if (action === "debug-prev") {
      prepareHomeDebugScenario((homeDebug && homeDebug.index || 0) - 1);
    } else if (action === "debug-next") {
      var debugSteps = homeDebugScenarios(homeDebug && homeDebug.baseDate);
      if (!homeDebug || homeDebug.index >= debugSteps.length - 1) stopHomeDebug();
      else prepareHomeDebugScenario(homeDebug.index + 1);
    } else if (action === "debug-exit") {
      stopHomeDebug();
    } else if (action === "settings-section") {
      route.tab = ["overview", "academic", "data", "experience", "developer"].indexOf(button.dataset.section) >= 0 ? button.dataset.section : "overview";
      render();
    } else if (action === "canteen-set-theme") {
      var nextCanteenTheme = button.dataset.theme === "leaf" ? "leaf" : "diner";
      if (state.settings.canteenTheme !== nextCanteenTheme) {
        state.settings.canteenTheme = nextCanteenTheme;
        await save(true);
        render();
        toast(nextCanteenTheme === "leaf" ? "Tema Leaf ativado na Cantina." : "Tema Diner ativado na Cantina.");
      }
    } else if (action === "toggle-canteen-ai") {
      state.settings.canteenAIEnabled = !state.settings.canteenAIEnabled;
      if (!state.settings.canteenAIEnabled) canteenAIState = { key: "", status: "idle", availability: "unknown", source: "rules", data: null, error: "", progress: null, streamText: "" };
      await save(true); render();
    } else if (action === "toggle-canteen-ai-note") {
      state.settings.canteenAIChefNote = !state.settings.canteenAIChefNote;
      await save(true); render();
    } else if (action === "toggle-canteen-ai-descriptions") {
      state.settings.canteenAIDescriptions = !state.settings.canteenAIDescriptions;
      await save(true); render();
    } else if (action === "toggle-reduce-motion") {
      state.settings.reduceMotion = !state.settings.reduceMotion;
      await save(true); render();
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

  async function handleDocumentClick(event) {
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
      await closeModalSavingNotebook();
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
        if (!state || onboarding || (homeDebug && homeDebug.active)) return;
        if (ensureBeOnlineTasks()) {
          save(true).then(function () { render(); toast("A aula terminou: o quiz de revisão está pronto."); });
        }
      }, 60000);
      if (!state.profile.onboardingComplete || !state.currentSemesterId || !activeCourses().length) startOnboarding(state.semesters.length ? "new-semester" : "first");
      if ("serviceWorker" in navigator && location.protocol !== "file:") {
        navigator.serviceWorker.register("sw.js?v=27.10-chef-resilient-stream", { updateViaCache: "none" }).then(function () {
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

  modalRoot.addEventListener("paste", function (event) {
    var editor = event.target.closest("#notebookEditor");
    if (!editor) return;
    handleNotebookPaste(event, editor);
  });
  modalRoot.addEventListener("keydown", function (event) {
    var editor = event.target.closest("#notebookEditor");
    if (!editor || event.key !== "Enter" || event.shiftKey) return;
    var selection = window.getSelection();
    var node = selection && selection.anchorNode;
    var element = node && (node.nodeType === 1 ? node : node.parentElement);
    var quote = element && element.closest ? element.closest("blockquote") : null;
    if (!quote || !editor.contains(quote)) return;
    event.preventDefault();
    var paragraph = document.createElement("p");
    paragraph.innerHTML = "<br>";
    quote.insertAdjacentElement("afterend", paragraph);
    var range = document.createRange();
    range.selectNodeContents(paragraph);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  });
  modalRoot.addEventListener("dragstart", function (event) {
    var sticker = event.target.closest(".notebook-sticker");
    if (!sticker) return;
    draggedNotebookSticker = sticker;
    sticker.classList.add("is-dragging");
    if (event.dataTransfer) { event.dataTransfer.effectAllowed = "move"; event.dataTransfer.setData("text/plain", "twenty-notebook-sticker"); }
  });
  modalRoot.addEventListener("dragend", function (event) {
    var sticker = event.target.closest(".notebook-sticker");
    if (sticker) sticker.classList.remove("is-dragging");
    draggedNotebookSticker = null;
  });
  modalRoot.addEventListener("dragover", function (event) {
    var editor = event.target.closest("#notebookEditor");
    if (!editor || !draggedNotebookSticker) return;
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
  });
  modalRoot.addEventListener("drop", function (event) {
    var editor = event.target.closest("#notebookEditor");
    if (!editor || !draggedNotebookSticker) return;
    event.preventDefault();
    var range = null;
    if (document.caretRangeFromPoint) range = document.caretRangeFromPoint(event.clientX, event.clientY);
    else if (document.caretPositionFromPoint) {
      var position = document.caretPositionFromPoint(event.clientX, event.clientY);
      if (position) { range = document.createRange(); range.setStart(position.offsetNode, position.offset); range.collapse(true); }
    }
    if (range && editor.contains(range.commonAncestorContainer)) range.insertNode(draggedNotebookSticker);
    else editor.appendChild(draggedNotebookSticker);
    draggedNotebookSticker.classList.remove("is-dragging");
    draggedNotebookSticker = null;
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
    if (event.target.matches("#notebookImageInput")) {
      var notebookEditorForFile = modalRoot.querySelector("#notebookEditor");
      var selectedFiles = Array.from(event.target.files || []);
      var savedNotebookRange = notebookEditorForFile ? notebookEditorRange(notebookEditorForFile) : null;
      (async function () {
        try {
          for (var imageIndex = 0; imageIndex < selectedFiles.length; imageIndex += 1) await uploadNotebookImage(selectedFiles[imageIndex], notebookEditorForFile, savedNotebookRange);
        } catch (error) {
          finishManualSyncActivity(false);
          toast(error.message || "Não foi possível adicionar a imagem.", "error");
        } finally {
          event.target.value = "";
        }
      })();
      return;
    }
    var builderForm = event.target.closest("#lessonBuilderForm");
    if (builderForm && event.target.matches('[name="materialIds"], [name="includePast"]')) {
      refreshLessonBuilderPrompt();
      return;
    }
    if (event.target.matches("#notebookPaper")) {
      var notebookEditor = modalRoot.querySelector("#notebookEditor");
      if (notebookEditor) notebookEditor.className = "notebook-page notebook-editor paper-" + event.target.value;
      return;
    }
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
  document.addEventListener("keydown", async function (event) {
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") { event.preventDefault(); searchInput.focus(); searchInput.select(); }
    if (event.key === "Escape") {
      if (guidedTour) stopGuidedTour(true);
      else if (!onboarding) await closeModalSavingNotebook();
      searchResults.hidden = true;
    }
    if ((event.key === "Enter" || event.key === " ") && event.target.matches('.course-card[role="button"]')) { event.preventDefault(); setRoute("course", event.target.dataset.id); }
  });
  window.addEventListener("hashchange", function () { routeFromHash(); render(); });
  window.addEventListener("focus", function () {
    clearTimeout(externalCheckTimer);
    externalCheckTimer = setTimeout(function () {
      if (Sync && Sync.getStatus().configured) return;
      if (!state || !state.settings.jsonSync || onboarding || (homeDebug && homeDebug.active)) return;
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
    if (homeDebug && homeDebug.active) return;
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
  document.addEventListener("twenty:math-ready", function () {
    typesetMath(view);
    typesetMath(modalRoot);
  });
  document.addEventListener("visibilitychange", function () {
    if (!document.hidden && !(homeDebug && homeDebug.active) && Sync && state && Sync.getStatus().configured) Sync.checkForUpdates({ force: true }).catch(function () {});
  });

  init();
})();
