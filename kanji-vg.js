(function (global) {
  "use strict";

  function resolveKanjiVgTxtUrl() {
    try {
      var u = global.KANJI_VG_TXT_URL;
      if (u && String(u).trim()) return String(u).trim();
    } catch (e) {}
    try {
      return new URL("KanjiVG.txt", global.location.href).href;
    } catch (e2) {
      return "KanjiVG.txt";
    }
  }

  function firstIdeographFromTsvCell(cell) {
    var s = String(cell || "")
      .normalize("NFC")
      .trim();
    if (!s) return "";
    for (var i = 0; i < s.length; ) {
      var cp = s.codePointAt(i);
      i += cp > 0xffff ? 2 : 1;
      if (
        (cp >= 0x4e00 && cp <= 0x9fff) ||
        (cp >= 0x3400 && cp <= 0x4dbf) ||
        (cp >= 0xf900 && cp <= 0xfaff)
      ) {
        return String.fromCodePoint(cp);
      }
    }
    return "";
  }

  function unicodeColToCharHexOnly(raw) {
    var t = String(raw || "").trim();
    if (!t) return "";
    var hexMatch = t.match(/^(?:U\+|0x)?([0-9A-Fa-f]{4,6})$/i);
    if (hexMatch) {
      var cp = parseInt(hexMatch[1], 16);
      if (!isNaN(cp) && cp > 0 && cp <= 0x10ffff) return String.fromCodePoint(cp);
    }
    return "";
  }

  function parseKanjiVgTsv(text) {
    var map = {};
    var lines = String(text || "").split(/\r?\n/);
    for (var li = 0; li < lines.length; li++) {
      var row = lines[li].trim();
      if (!row || row.charAt(0) === "#") continue;
      var cols = row.split("\t");
      if (cols.length < 3) continue;
      var kanji = firstIdeographFromTsvCell(cols[0]);
      var unicodeCol = String(cols[1] || "").trim();
      var strokesCol = String(cols[2] || "").trim();
      if (!kanji) kanji = unicodeColToCharHexOnly(unicodeCol);
      if (!kanji || !strokesCol) continue;
      var paths = strokesCol
        .split("|")
        .map(function (p) {
          return String(p || "").trim();
        })
        .filter(function (p) {
          return p && (p.charAt(0) === "M" || p.charAt(0) === "m");
        });
      if (paths.length > 0) map[kanji] = paths;
    }
    return map;
  }

  function fetchMap(customFetch) {
    var fn = customFetch || global.fetch;
    if (typeof fn !== "function") return Promise.reject(new Error("fetch unavailable"));
    var url = resolveKanjiVgTxtUrl();
    return fn(url).then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.text();
    }).then(function (text) {
      var data = parseKanjiVgTsv(text);
      if (!Object.keys(data).length) throw new Error("TSVの解析結果が空です");
      return data;
    });
  }

  global.KanjiVg = {
    resolveTxtUrl: resolveKanjiVgTxtUrl,
    parseTsv: parseKanjiVgTsv,
    fetchMap: fetchMap
  };
})(typeof window !== "undefined" ? window : globalThis);
