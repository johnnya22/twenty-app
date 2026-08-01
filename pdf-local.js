(function () {
  "use strict";

  function uid(prefix) {
    return (prefix || "pdf") + "_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
  }

  function bytesToBinary(bytes) {
    var out = "";
    var size = 32768;
    for (var index = 0; index < bytes.length; index += size) {
      out += String.fromCharCode.apply(null, bytes.subarray(index, Math.min(index + size, bytes.length)));
    }
    return out;
  }

  function binaryToBytes(text) {
    var bytes = new Uint8Array(text.length);
    for (var index = 0; index < text.length; index += 1) bytes[index] = text.charCodeAt(index) & 255;
    return bytes;
  }

  function cleanText(value) {
    return String(value || "")
      .replace(/\u0000/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n[ \t]+/g, "\n")
      .replace(/[ \t]{2,}/g, " ")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function streamBounds(raw, objectStart, objectEnd) {
    var marker = raw.indexOf("stream", objectStart);
    if (marker < 0 || marker > objectEnd) return null;
    var after = marker + 6;
    if (raw.slice(after, after + 2) === "\r\n") after += 2;
    else if (raw.charAt(after) === "\n" || raw.charAt(after) === "\r") after += 1;
    else return null;
    var end = raw.indexOf("endstream", after);
    if (end < 0) return null;
    var dataEnd = end;
    while (dataEnd > after && (raw.charAt(dataEnd - 1) === "\n" || raw.charAt(dataEnd - 1) === "\r")) dataEnd -= 1;
    return { marker: marker, start: after, end: dataEnd, endstream: end };
  }

  function parseObjects(bytes) {
    var raw = bytesToBinary(bytes);
    var objects = new Map();
    var pattern = /(?:^|[\r\n\t \f])(\d+)\s+(\d+)\s+obj\b/g;
    var match;
    while ((match = pattern.exec(raw))) {
      var number = Number(match[1]);
      var generation = Number(match[2]);
      var start = pattern.lastIndex;
      var firstEnd = raw.indexOf("endobj", start);
      if (firstEnd < 0) break;
      var bounds = streamBounds(raw, start, firstEnd);
      var end = bounds ? raw.indexOf("endobj", bounds.endstream + 9) : firstEnd;
      if (end < 0) end = firstEnd;
      var bodyEnd = bounds ? bounds.marker : end;
      objects.set(number, {
        number: number,
        generation: generation,
        body: raw.slice(start, bodyEnd).trim(),
        streamBytes: bounds ? bytes.slice(bounds.start, bounds.end) : null
      });
      pattern.lastIndex = end + 6;
    }
    return objects;
  }

  async function inflate(bytes) {
    if (typeof DecompressionStream !== "function") throw new Error("Este browser não suporta a descompressão local necessária para este PDF.");
    var attempts = ["deflate", "deflate-raw"];
    var lastError = null;
    for (var index = 0; index < attempts.length; index += 1) {
      try {
        var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(attempts[index]));
        return new Uint8Array(await new Response(stream).arrayBuffer());
      } catch (error) { lastError = error; }
    }
    throw lastError || new Error("Não foi possível descomprimir o PDF.");
  }

  function decodeAscii85(bytes) {
    var text = bytesToBinary(bytes).replace(/\s+/g, "").replace(/^<~/, "").replace(/~>.*$/, "");
    var output = [];
    var group = [];
    function flush(finalGroup) {
      if (!group.length) return;
      var originalLength = group.length;
      while (group.length < 5) group.push(84);
      var value = 0;
      for (var index = 0; index < 5; index += 1) value = value * 85 + group[index];
      var decoded = [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];
      var count = finalGroup ? Math.max(0, originalLength - 1) : 4;
      for (var byteIndex = 0; byteIndex < count; byteIndex += 1) output.push(decoded[byteIndex]);
      group = [];
    }
    for (var cursor = 0; cursor < text.length; cursor += 1) {
      var char = text.charAt(cursor);
      if (char === "z") {
        if (group.length) throw new Error("ASCII85 inválido no PDF.");
        output.push(0, 0, 0, 0);
        continue;
      }
      var code = char.charCodeAt(0) - 33;
      if (code < 0 || code > 84) continue;
      group.push(code);
      if (group.length === 5) flush(false);
    }
    flush(true);
    return new Uint8Array(output);
  }

  function filtersFor(body) {
    var match = String(body || "").match(/\/Filter\s*(\[[^\]]+\]|\/[A-Za-z0-9]+)/);
    if (!match) return [];
    return (match[1].match(/\/[A-Za-z0-9]+/g) || []).map(function (item) { return item.slice(1); });
  }

  async function decodedStream(object) {
    if (!object || !object.streamBytes) return null;
    if (object.decodedBytes) return object.decodedBytes;
    var bytes = object.streamBytes;
    var filters = filtersFor(object.body);
    for (var index = 0; index < filters.length; index += 1) {
      if (filters[index] === "FlateDecode" || filters[index] === "Fl") bytes = await inflate(bytes);
      else if (filters[index] === "ASCII85Decode" || filters[index] === "A85") bytes = decodeAscii85(bytes);
      else if (filters[index] === "ASCIIHexDecode") {
        var hex = bytesToBinary(bytes).replace(/\s+/g, "").replace(/>.*$/, "");
        if (hex.length % 2) hex += "0";
        var output = new Uint8Array(hex.length / 2);
        for (var cursor = 0; cursor < hex.length; cursor += 2) output[cursor / 2] = parseInt(hex.slice(cursor, cursor + 2), 16) || 0;
        bytes = output;
      } else {
        throw new Error("Este PDF usa uma compressão ainda não suportada localmente: " + filters[index] + ".");
      }
    }
    object.decodedBytes = bytes;
    return bytes;
  }

  async function expandObjectStreams(objects) {
    var packed = Array.from(objects.values()).filter(function (object) { return /\/Type\s*\/ObjStm\b/.test(object.body); });
    for (var objectIndex = 0; objectIndex < packed.length; objectIndex += 1) {
      var object = packed[objectIndex];
      var countMatch = object.body.match(/\/N\s+(\d+)/);
      var firstMatch = object.body.match(/\/First\s+(\d+)/);
      if (!countMatch || !firstMatch) continue;
      var count = Number(countMatch[1]);
      var first = Number(firstMatch[1]);
      var decoded = await decodedStream(object);
      var text = bytesToBinary(decoded || new Uint8Array());
      var header = text.slice(0, first).trim().split(/\s+/).map(Number);
      for (var index = 0; index < count; index += 1) {
        var number = header[index * 2];
        var offset = header[index * 2 + 1];
        var nextOffset = index + 1 < count ? header[(index + 1) * 2 + 1] : text.length - first;
        if (!Number.isFinite(number) || !Number.isFinite(offset)) continue;
        objects.set(number, { number: number, generation: 0, body: text.slice(first + offset, first + nextOffset).trim(), streamBytes: null, packed: true });
      }
    }
  }

  function references(value) {
    var refs = [];
    String(value || "").replace(/(\d+)\s+\d+\s+R/g, function (_, number) { refs.push(Number(number)); return _; });
    return refs;
  }

  function refFor(body, name) {
    var match = String(body || "").match(new RegExp("/" + name + "\\s+(\\d+)\\s+\\d+\\s+R"));
    return match ? Number(match[1]) : null;
  }

  function balancedDictionary(text, start) {
    start = text.indexOf("<<", start || 0);
    if (start < 0) return "";
    var depth = 0;
    for (var index = start; index < text.length - 1; index += 1) {
      var pair = text.slice(index, index + 2);
      if (pair === "<<") { depth += 1; index += 1; }
      else if (pair === ">>") {
        depth -= 1; index += 1;
        if (depth === 0) return text.slice(start, index + 1);
      }
    }
    return "";
  }

  function namedDictionary(body, name, objects) {
    var source = String(body || "");
    var ref = refFor(source, name);
    if (ref && objects.has(ref)) return objects.get(ref).body;
    var position = source.search(new RegExp("/" + name + "\\s*<<"));
    return position >= 0 ? balancedDictionary(source, source.indexOf("<<", position)) : "";
  }

  function inheritedDictionary(page, name, objects) {
    var cursor = page;
    var visited = new Set();
    while (cursor && !visited.has(cursor.number)) {
      visited.add(cursor.number);
      var dictionary = namedDictionary(cursor.body, name, objects);
      if (dictionary) return dictionary;
      var parent = refFor(cursor.body, "Parent");
      cursor = parent ? objects.get(parent) : null;
    }
    return "";
  }

  function decodeUtf16Hex(hex) {
    hex = String(hex || "").replace(/[^0-9A-Fa-f]/g, "");
    if (!hex) return "";
    if (hex.length % 4 && hex.length % 2 === 0) {
      var oneByte = "";
      for (var byteIndex = 0; byteIndex < hex.length; byteIndex += 2) oneByte += String.fromCharCode(parseInt(hex.slice(byteIndex, byteIndex + 2), 16));
      return oneByte;
    }
    var output = "";
    for (var index = 0; index + 3 < hex.length; index += 4) output += String.fromCharCode(parseInt(hex.slice(index, index + 4), 16));
    try { return decodeURIComponent(escape(output)); } catch (_) { return output; }
  }

  function incrementHex(hex, amount) {
    var width = hex.length;
    var value = parseInt(hex, 16) + amount;
    return value.toString(16).toUpperCase().padStart(width, "0");
  }

  function parseCMap(text) {
    var map = new Map();
    var lengths = new Set();
    text = String(text || "");
    text.replace(/begincodespacerange([\s\S]*?)endcodespacerange/g, function (_, block) {
      block.replace(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g, function (_m, start) { lengths.add(start.length / 2); return _m; });
      return _;
    });
    text.replace(/beginbfchar([\s\S]*?)endbfchar/g, function (_, block) {
      block.replace(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g, function (_m, source, target) {
        map.set(source.toUpperCase(), decodeUtf16Hex(target)); lengths.add(source.length / 2); return _m;
      });
      return _;
    });
    text.replace(/beginbfrange([\s\S]*?)endbfrange/g, function (_, block) {
      var lines = block.split(/[\r\n]+/);
      lines.forEach(function (line) {
        var arrayMatch = line.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[([^\]]+)\]/);
        if (arrayMatch) {
          var start = parseInt(arrayMatch[1], 16), end = parseInt(arrayMatch[2], 16);
          var targets = arrayMatch[3].match(/<([0-9A-Fa-f]+)>/g) || [];
          for (var value = start, index = 0; value <= end && index < targets.length; value += 1, index += 1) {
            var source = value.toString(16).toUpperCase().padStart(arrayMatch[1].length, "0");
            map.set(source, decodeUtf16Hex(targets[index].slice(1, -1)));
          }
          lengths.add(arrayMatch[1].length / 2);
          return;
        }
        var simple = line.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/);
        if (!simple) return;
        var from = parseInt(simple[1], 16), to = parseInt(simple[2], 16);
        for (var code = from; code <= to; code += 1) {
          var sourceHex = code.toString(16).toUpperCase().padStart(simple[1].length, "0");
          map.set(sourceHex, decodeUtf16Hex(incrementHex(simple[3], code - from)));
        }
        lengths.add(simple[1].length / 2);
      });
      return _;
    });
    return { map: map, lengths: Array.from(lengths).filter(Boolean).sort(function (a, b) { return b - a; }) };
  }

  async function fontDecoder(fontObject, objects) {
    if (!fontObject) return null;
    if (fontObject.decoder !== undefined) return fontObject.decoder;
    var toUnicode = refFor(fontObject.body, "ToUnicode");
    if (toUnicode && objects.has(toUnicode)) {
      try {
        var decoded = await decodedStream(objects.get(toUnicode));
        fontObject.decoder = parseCMap(bytesToBinary(decoded || new Uint8Array()));
        return fontObject.decoder;
      } catch (_) {}
    }
    fontObject.decoder = null;
    return null;
  }

  function win1252(bytes) {
    try { return new TextDecoder("windows-1252").decode(bytes); }
    catch (_) { return bytesToBinary(bytes); }
  }

  function decodeTextBytes(bytes, decoder) {
    if (!decoder || !decoder.map || !decoder.map.size) return win1252(bytes);
    var hex = Array.from(bytes).map(function (value) { return value.toString(16).toUpperCase().padStart(2, "0"); }).join("");
    var output = "";
    var cursor = 0;
    while (cursor < hex.length) {
      var found = false;
      for (var index = 0; index < decoder.lengths.length; index += 1) {
        var chars = decoder.lengths[index] * 2;
        var key = hex.slice(cursor, cursor + chars);
        if (key.length === chars && decoder.map.has(key)) {
          output += decoder.map.get(key); cursor += chars; found = true; break;
        }
      }
      if (!found) {
        output += String.fromCharCode(parseInt(hex.slice(cursor, cursor + 2), 16) || 32);
        cursor += 2;
      }
    }
    return output;
  }

  function parseLiteral(text, start) {
    var bytes = [];
    var depth = 1;
    var index = start + 1;
    while (index < text.length && depth > 0) {
      var code = text.charAt(index);
      if (code === "\\") {
        index += 1;
        if (index >= text.length) break;
        var escaped = text.charAt(index);
        var simple = { n: 10, r: 13, t: 9, b: 8, f: 12, "(": 40, ")": 41, "\\": 92 };
        if (Object.prototype.hasOwnProperty.call(simple, escaped)) bytes.push(simple[escaped]);
        else if (/[0-7]/.test(escaped)) {
          var octal = escaped;
          while (octal.length < 3 && /[0-7]/.test(text.charAt(index + 1))) { index += 1; octal += text.charAt(index); }
          bytes.push(parseInt(octal, 8) & 255);
        } else if (escaped === "\r" || escaped === "\n") {
          if (escaped === "\r" && text.charAt(index + 1) === "\n") index += 1;
        } else bytes.push(escaped.charCodeAt(0) & 255);
      } else if (code === "(") { depth += 1; bytes.push(40); }
      else if (code === ")") { depth -= 1; if (depth > 0) bytes.push(41); }
      else bytes.push(code.charCodeAt(0) & 255);
      index += 1;
    }
    return { token: { type: "bytes", value: new Uint8Array(bytes) }, end: index };
  }

  function tokenizeContent(text) {
    var tokens = [];
    var index = 0;
    while (index < text.length) {
      var char = text.charAt(index);
      if (/\s/.test(char)) { index += 1; continue; }
      if (char === "%") { while (index < text.length && text.charAt(index) !== "\n" && text.charAt(index) !== "\r") index += 1; continue; }
      if (char === "(") { var literal = parseLiteral(text, index); tokens.push(literal.token); index = literal.end; continue; }
      if (char === "<" && text.charAt(index + 1) !== "<") {
        var close = text.indexOf(">", index + 1); if (close < 0) break;
        var hex = text.slice(index + 1, close).replace(/\s+/g, ""); if (hex.length % 2) hex += "0";
        var hexBytes = new Uint8Array(hex.length / 2);
        for (var h = 0; h < hex.length; h += 2) hexBytes[h / 2] = parseInt(hex.slice(h, h + 2), 16) || 0;
        tokens.push({ type: "bytes", value: hexBytes }); index = close + 1; continue;
      }
      if (char === "[") {
        var array = []; index += 1;
        while (index < text.length) {
          while (/\s/.test(text.charAt(index))) index += 1;
          if (text.charAt(index) === "]") { index += 1; break; }
          if (text.charAt(index) === "(") { var itemLiteral = parseLiteral(text, index); array.push(itemLiteral.token); index = itemLiteral.end; continue; }
          if (text.charAt(index) === "<" && text.charAt(index + 1) !== "<") {
            var itemClose = text.indexOf(">", index + 1); if (itemClose < 0) break;
            var itemHex = text.slice(index + 1, itemClose).replace(/\s+/g, ""); if (itemHex.length % 2) itemHex += "0";
            var itemBytes = new Uint8Array(itemHex.length / 2);
            for (var ih = 0; ih < itemHex.length; ih += 2) itemBytes[ih / 2] = parseInt(itemHex.slice(ih, ih + 2), 16) || 0;
            array.push({ type: "bytes", value: itemBytes }); index = itemClose + 1; continue;
          }
          var itemEnd = index;
          while (itemEnd < text.length && !/[\s\]]/.test(text.charAt(itemEnd))) itemEnd += 1;
          var item = text.slice(index, itemEnd);
          array.push(/^[-+]?\d*\.?\d+$/.test(item) ? { type: "number", value: Number(item) } : { type: "word", value: item });
          index = itemEnd;
        }
        tokens.push({ type: "array", value: array }); continue;
      }
      if (char === "/") {
        var nameEnd = index + 1;
        while (nameEnd < text.length && !/[\s\[\]()<>/%]/.test(text.charAt(nameEnd))) nameEnd += 1;
        tokens.push({ type: "name", value: text.slice(index + 1, nameEnd) }); index = nameEnd; continue;
      }
      var end = index;
      while (end < text.length && !/[\s\[\]()<>/%]/.test(text.charAt(end))) end += 1;
      var word = text.slice(index, end);
      tokens.push(/^[-+]?\d*\.?\d+$/.test(word) ? { type: "number", value: Number(word) } : { type: "word", value: word });
      index = end || index + 1;
    }
    return tokens;
  }

  async function pageFontMap(page, objects) {
    var resources = inheritedDictionary(page, "Resources", objects);
    var fontDictionary = namedDictionary(resources, "Font", objects);
    var fonts = {};
    var pattern = /\/([^\s<>{}\[\]()/%]+)\s+(\d+)\s+\d+\s+R/g;
    var match;
    while ((match = pattern.exec(fontDictionary))) {
      var font = objects.get(Number(match[2]));
      fonts[match[1]] = { object: font, decoder: await fontDecoder(font, objects) };
    }
    return fonts;
  }

  async function extractContentText(contentObject, fonts) {
    var decoded = await decodedStream(contentObject);
    if (!decoded) return "";
    var tokens = tokenizeContent(bytesToBinary(decoded));
    var stack = [];
    var currentFont = null;
    var output = [];
    function add(value) {
      value = String(value || "").replace(/\u0000/g, "");
      if (!value) return;
      output.push(value);
    }
    function decodeToken(token) {
      if (!token || token.type !== "bytes") return "";
      var font = currentFont && fonts[currentFont];
      return decodeTextBytes(token.value, font && font.decoder);
    }
    for (var index = 0; index < tokens.length; index += 1) {
      var token = tokens[index];
      if (token.type !== "word") { stack.push(token); continue; }
      var operator = token.value;
      if (operator === "Tf") {
        var fontName = stack.length >= 2 ? stack[stack.length - 2] : null;
        if (fontName && fontName.type === "name") currentFont = fontName.value;
      } else if (operator === "Tj") {
        add(decodeToken(stack[stack.length - 1]));
      } else if (operator === "TJ") {
        var array = stack[stack.length - 1];
        if (array && array.type === "array") {
          var line = "";
          array.value.forEach(function (part) {
            if (part.type === "bytes") line += decodeToken(part);
            else if (part.type === "number" && part.value < -160 && line && !/\s$/.test(line)) line += " ";
          });
          add(line);
        }
      } else if (operator === "'" || operator === '"') {
        output.push("\n"); add(decodeToken(stack[stack.length - 1]));
      } else if (operator === "T*" || operator === "Td" || operator === "TD" || operator === "ET") {
        output.push("\n");
      }
      stack = [];
    }
    return cleanText(output.join(" ").replace(/\s*\n\s*/g, "\n"));
  }

  function contentReferences(page) {
    var body = page.body;
    var arrayMatch = body.match(/\/Contents\s*\[([\s\S]*?)\]/);
    if (arrayMatch) return references(arrayMatch[1]);
    var single = refFor(body, "Contents");
    return single ? [single] : [];
  }

  function textQuality(text) {
    if (!text) return 0;
    var visible = (text.match(/[\p{L}\p{N}]/gu) || []).length;
    var bad = (text.match(/[\u0000-\u0008\u000E-\u001F\uFFFD]/g) || []).length;
    return visible / Math.max(1, text.length) - bad / Math.max(1, text.length);
  }

  async function extract(file, onProgress) {
    if (!file) throw new Error("Escolhe um PDF primeiro.");
    if (!/\.pdf$/i.test(file.name || "") && !/application\/pdf/i.test(file.type || "")) throw new Error("Este ficheiro não é um PDF.");
    if (file.size > 25 * 1024 * 1024) throw new Error("Este PDF tem mais de 25 MB. Comprime-o ou divide-o antes de o carregar.");
    if (onProgress) onProgress({ progress: 2, text: "A abrir o PDF localmente…" });
    var bytes = new Uint8Array(await file.arrayBuffer());
    if (bytesToBinary(bytes.slice(0, 8)).indexOf("%PDF-") !== 0) throw new Error("O ficheiro não parece ser um PDF válido.");
    var objects = parseObjects(bytes);
    if (!objects.size) throw new Error("Não foi possível ler a estrutura deste PDF.");
    await expandObjectStreams(objects);
    var pages = Array.from(objects.values()).filter(function (object) { return /\/Type\s*\/Page\b/.test(object.body) && !/\/Type\s*\/Pages\b/.test(object.body); });
    if (!pages.length) throw new Error("Não foram encontradas páginas neste PDF.");
    var slides = [];
    for (var pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
      var page = pages[pageIndex];
      var fonts = await pageFontMap(page, objects);
      var refs = contentReferences(page);
      var chunks = [];
      for (var refIndex = 0; refIndex < refs.length; refIndex += 1) {
        var content = objects.get(refs[refIndex]);
        if (content) chunks.push(await extractContentText(content, fonts));
      }
      var text = cleanText(chunks.filter(Boolean).join("\n"));
      slides.push({ number: pageIndex + 1, title: cleanText((text.split("\n")[0] || "Página " + (pageIndex + 1))).slice(0, 180), text: text.slice(0, 8000), sourceType: "pdf-page" });
      if (onProgress) onProgress({ progress: 5 + Math.round(((pageIndex + 1) / pages.length) * 25), text: "A extrair a página " + (pageIndex + 1) + " de " + pages.length + "…" });
    }
    var withText = slides.filter(function (page) { return textQuality(page.text) > 0.18 && page.text.length > 2; });
    if (!withText.length) throw new Error("O PDF não tem texto selecionável compatível com o leitor local.");
    return { id: uid("pdf"), fileName: file.name, fileSize: file.size, slideCount: slides.length, slides: slides, extractedAt: new Date().toISOString(), sourceType: "pdf", extractor: "twenty-local" };
  }

  window.TwentyPDF = { extract: extract };
})();
