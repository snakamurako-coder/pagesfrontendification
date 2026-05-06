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

  /** TSV の1行から { kanji, paths } を返す。不正行は null */
  function parseKanjiVgLine(line) {
    var row = String(line || "").trim();
    if (!row || row.charAt(0) === "#") return null;
    var cols = row.split("\t");
    if (cols.length < 3) return null;
    var kanji = firstIdeographFromTsvCell(cols[0]);
    var unicodeCol = String(cols[1] || "").trim();
    var strokesCol = String(cols[2] || "").trim();
    if (!kanji) kanji = unicodeColToCharHexOnly(unicodeCol);
    if (!kanji || !strokesCol) return null;
    var paths = strokesCol
      .split("|")
      .map(function (p) {
        return String(p || "").trim();
      })
      .filter(function (p) {
        return p && (p.charAt(0) === "M" || p.charAt(0) === "m");
      });
    if (paths.length < 1) return null;
    return { kanji: kanji.normalize("NFC"), paths: paths };
  }

  /**
   * KanjiVG.txt の全文から、指定した1文字のストロークパスだけを検索（①問題の漢字 → ②TSV直接参照）
   */
  function pathsForChar(text, char) {
    var want = String(char || "").normalize("NFC");
    if (!want) return null;
    var lines = String(text || "").split(/\r?\n/);
    for (var li = 0; li < lines.length; li++) {
      var parsed = parseKanjiVgLine(lines[li]);
      if (parsed && parsed.kanji === want) return parsed.paths;
    }
    return null;
  }

  function parseKanjiVgTsv(text) {
    var map = {};
    var lines = String(text || "").split(/\r?\n/);
    for (var li = 0; li < lines.length; li++) {
      var parsed = parseKanjiVgLine(lines[li]);
      if (parsed) map[parsed.kanji] = parsed.paths;
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
    parseLine: parseKanjiVgLine,
    pathsForChar: pathsForChar,
    fetchMap: fetchMap
  };
})(typeof window !== "undefined" ? window : globalThis);
