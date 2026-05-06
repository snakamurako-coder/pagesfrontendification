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

const KP_IFRAME_HTML = "<!DOCTYPE html>\n<html>\n<head>\n  <meta charset=\"UTF-8\">\n  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n  <link href=\"https://fonts.googleapis.com/css2?family=Klee+One:wght@400;600&display=swap\" rel=\"stylesheet\">\n  \n  <style>\n    body { font-family: 'Klee One', sans-serif; background: #f0f2f5; padding: 20px; display: flex; flex-direction: column; align-items: center; margin: 0; overflow-x: hidden; }\n    h2 { color: #333; margin-bottom: 5px; }\n    canvas { touch-action: none; background-image: linear-gradient(#eee 1px, transparent 1px), linear-gradient(90deg, #eee 1px, transparent 1px); background-size: 50% 50%; }\n    #canvas { background-color: #fff; border: 2px solid #333; border-radius: 4px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); cursor: crosshair; }\n    #result-box { margin-top: 15px; padding: 15px; background: #fff; border-radius: 8px; width: 100%; max-width: 350px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }\n    .score { font-size: 38px; font-weight: bold; color: #e53935; }\n    .msg { font-size: 13px; color: #555; margin-top: 10px; font-weight: bold; line-height: 1.5; }\n    button { padding: 10px 20px; border-radius: 20px; border: none; background: #1a73e8; color: white; font-weight: bold; cursor: pointer; font-size: 14px; transition: all 0.2s; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }\n    button:active { transform: translateY(1px); box-shadow: none; }\n    .hidden { display: none !important; }\n    .controls { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; margin-top: 10px; }\n    \n    #kanji-selector-area { margin-bottom: 12px; display:flex; flex-direction: column; align-items:center; gap: 10px; background: #fff; padding: 15px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); font-weight: bold; width: 100%; max-width: 350px; box-sizing: border-box; }\n    .selector-row { display: flex; gap: 8px; width: 100%; justify-content: center; align-items: center; }\n    select, input[type=\"text\"] { padding: 8px; border-radius: 6px; border: 2px solid #1a73e8; font-family: 'Klee One'; font-size: 14px; outline: none; flex: 1; background: #fff;}\n    \n    #loading-overlay { position: fixed; top:0; left:0; right:0; bottom:0; background: rgba(255,255,255,0.9); display:flex; flex-direction:column; justify-content:center; align-items:center; z-index: 1000; font-family: 'Klee One'; font-size: 18px; color: #1a73e8; }\n    .loader { border: 4px solid #f3f3f3; border-top: 4px solid #1a73e8; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin-bottom: 15px; }\n    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }\n\n    /* モーダル汎用 */\n    .modal-overlay { position: fixed; top:0; left:0; right:0; bottom:0; background: rgba(0,0,0,0.7); display: flex; justify-content: center; align-items: center; z-index: 2000; }\n    .modal-content { background: #fff; padding: 20px; border-radius: 12px; width: 95%; max-width: 400px; max-height: 95vh; overflow-y: auto; box-shadow: 0 10px 25px rgba(0,0,0,0.3); }\n    .modal-content h3 { margin: 0 0 15px 0; color: #1a73e8; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; font-size: 16px; }\n\n    /* 設定用 */\n    .slider-group { margin-bottom: 12px; }\n    .slider-group label { display: flex; justify-content: space-between; font-size: 12px; font-weight: bold; color: #555; margin-bottom: 3px; }\n    .slider-group input[type=range] { width: 100%; cursor: pointer; margin: 0; }\n    #preview-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 15px; background: #f0f2f5; padding: 10px; border-radius: 8px; }\n    .preview-box { background: #fff; border: 1px solid #ddd; border-radius: 4px; position: relative; aspect-ratio: 1; display: flex; flex-direction: column; align-items: center; }\n    .preview-canvas { width: 100%; height: 100%; background: transparent; }\n    .preview-label { position: absolute; top: 2px; left: 4px; font-size: 10px; color: #999; font-weight: bold; }\n\n    /* ウィザード用 */\n    #wizard-canvas { background-color: #fff; border: 2px dashed #aaa; border-radius: 4px; cursor: crosshair; margin: 10px auto; display: block; }\n    .wiz-btn-group { display: flex; gap: 5px; justify-content: center; margin-bottom: 8px; flex-wrap: wrap; }\n    .wiz-btn { flex: 1 1 30%; padding: 8px 5px; font-size: 12px; background: #fff; color: #333; border: 2px solid #ccc; font-weight: bold; }\n    .wiz-btn-ignore { flex: 1 1 100%; padding: 8px; font-size: 12px; background: #e0e0e0; color: #555; border: 2px solid #bbb; font-weight: bold; }\n    .wiz-stats { background: #f9f9f9; padding: 10px; border-radius: 6px; font-size: 12px; color: #555; }\n    .profile-slot { display: flex; gap: 5px; margin-bottom: 8px; align-items: center; }\n    .profile-slot input { flex: 1; font-size: 12px; padding: 6px; }\n    .profile-slot button { font-size: 11px; padding: 6px 10px; border-radius: 4px; }\n  </style>\n</head>\n<body>\n\n  <div id=\"loading-overlay\"><div class=\"loader\"></div><span>準備中...</span></div>\n\n  <h2>漢字書き順ドリル</h2>\n\n  <div id=\"kanji-selector-area\">\n    <div class=\"selector-row\">\n      <select id=\"sheet-selector\" onchange=\"loadKanjiData()\"></select>\n      <button onclick=\"syncData()\" style=\"margin:0; background:#34a853; padding:6px 12px; font-size:12px;\">更新</button>\n    </div>\n    <div class=\"selector-row\">\n      <select id=\"target-kanji\" onchange=\"initTargetKanji()\"></select>\n      <button onclick=\"switchMode('demo')\" style=\"margin:0; background:#f4b400; padding:6px 12px; font-size:12px;\">お手本</button>\n    </div>\n    <div class=\"selector-row\">\n      <select id=\"active-profile-select\" onchange=\"switchProfile(this.value)\" style=\"flex: 1.5; font-size:12px; padding: 6px;\"></select>\n      <button onclick=\"toggleWizard(true)\" style=\"margin:0; background:#e91e63; padding:6px 10px; font-size:12px; flex:1;\">🎯 調整</button>\n      <button onclick=\"toggleSettings(true)\" style=\"margin:0; background:#7b1fa2; padding:6px 10px; font-size:12px; flex:1;\">🖋 設定</button>\n    </div>\n  </div>\n\n  <div id=\"mode-msg\" style=\"font-size:12px; color:#d32f2f; margin-bottom:10px; font-weight: bold;\">正確に書いてください</div>\n\n  <canvas id=\"canvas\"></canvas>\n\n  <div class=\"controls\" style=\"margin-top:10px;\">\n    <label style=\"font-size:12px; font-weight:bold; color:#1a73e8; display:flex; align-items:center; cursor:pointer; background:#e8f0fe; padding:5px 10px; border-radius:15px;\">\n      <input type=\"checkbox\" id=\"cb-strict-mode\" style=\"margin-right:5px; transform: scale(1.2);\"> ☑ 厳密採点 (とめ・はね・はらい判定)\n    </label>\n    <button style=\"background:#ff9800; padding:5px 10px; font-size:12px; border-radius:15px; box-shadow:none;\" onclick=\"removeLastEffect()\">直前画の飾りを消す</button>\n  </div>\n\n  <div id=\"score-controls\" class=\"controls\">\n    <button onclick=\"evaluateKanji(false)\">採点する</button>\n    <button style=\"background:#8e8e93;\" onclick=\"clearCanvas()\">クリア</button>\n  </div>\n  \n  <div id=\"trace-controls\" class=\"controls hidden\">\n    <button style=\"background:#e53935;\" onclick=\"evaluateKanji(true)\">なぞりを採点</button>\n    <button style=\"background:#8e8e93;\" onclick=\"clearCanvas()\">書き直す</button>\n    <button style=\"background:#333;\" onclick=\"switchMode('score')\">戻る</button>\n  </div>\n\n  <div id=\"demo-controls\" class=\"controls hidden\">\n    <button style=\"background:#1a73e8;\" onclick=\"playAnimation()\">再生</button>\n    <button style=\"background:#e53935;\" onclick=\"switchMode('trace')\">なぞり練習</button>\n    <button style=\"background:#333;\" onclick=\"switchMode('score')\">戻る</button>\n  </div>\n\n  <div id=\"result-box\">\n    <div id=\"score\" class=\"score\">- 点</div>\n    <div id=\"msg\" class=\"msg\">書いて「採点する」を押してください</div>\n  </div>\n\n  <div id=\"settings-modal\" class=\"modal-overlay hidden\">\n    <div class=\"modal-content\">\n      <h3>🖋 ペン・エフェクト表現設定</h3>\n      <div class=\"slider-group\">\n        <label>基本の太さ <span id=\"val-baseWidth\"></span></label>\n        <input type=\"range\" id=\"sl-baseWidth\" min=\"4\" max=\"18\" step=\"0.5\" oninput=\"updateParam('baseWidth', this.value)\">\n      </div>\n      <div style=\"background:#fff3e0; padding:10px; border-radius:8px; margin-bottom:10px; border-left: 4px solid #ff9800;\">\n        <label style=\"font-size:11px; color:#ef6c00; display:block; margin-bottom:5px;\">【とめ】インク溜まり</label>\n        <div class=\"slider-group\">\n          <label>開始位置 <span id=\"val-tomeStart\"></span></label>\n          <input type=\"range\" id=\"sl-tomeStart\" min=\"0.5\" max=\"0.95\" step=\"0.05\" oninput=\"updateParam('tomeStart', this.value)\">\n        </div>\n        <div class=\"slider-group\">\n          <label>膨張率 <span id=\"val-tomeScale\"></span></label>\n          <input type=\"range\" id=\"sl-tomeScale\" min=\"1.0\" max=\"2.0\" step=\"0.1\" oninput=\"updateParam('tomeScale', this.value)\">\n        </div>\n      </div>\n      <div style=\"background:#e8f5e9; padding:10px; border-radius:8px; margin-bottom:10px; border-left: 4px solid #4caf50;\">\n        <label style=\"font-size:11px; color:#2e7d32; display:block; margin-bottom:5px;\">【はね】跳ね上げ</label>\n        <div class=\"slider-group\">\n          <label>開始位置 <span id=\"val-haneStart\"></span></label>\n          <input type=\"range\" id=\"sl-haneStart\" min=\"0.5\" max=\"0.95\" step=\"0.05\" oninput=\"updateParam('haneStart', this.value)\">\n        </div>\n        <div class=\"slider-group\">\n          <label>先端の鋭さ <span id=\"val-haneSharp\"></span></label>\n          <input type=\"range\" id=\"sl-haneSharp\" min=\"0.5\" max=\"1.0\" step=\"0.05\" oninput=\"updateParam('haneSharp', this.value)\">\n        </div>\n      </div>\n      <div style=\"background:#e3f2fd; padding:10px; border-radius:8px; border-left: 4px solid #2196f3;\">\n        <label style=\"font-size:11px; color:#1565c0; display:block; margin-bottom:5px;\">【はらい】</label>\n        <div class=\"slider-group\">\n          <label>開始位置 <span id=\"val-haraiStart\"></span></label>\n          <input type=\"range\" id=\"sl-haraiStart\" min=\"0.3\" max=\"0.9\" step=\"0.05\" oninput=\"updateParam('haraiStart', this.value)\">\n        </div>\n        <div class=\"slider-group\">\n          <label>先端の鋭さ <span id=\"val-haraiSharp\"></span></label>\n          <input type=\"range\" id=\"sl-haraiSharp\" min=\"0.5\" max=\"1.0\" step=\"0.05\" oninput=\"updateParam('haraiSharp', this.value)\">\n        </div>\n      </div>\n      <div id=\"preview-grid\">\n        <div class=\"preview-box\"><span class=\"preview-label\">永</span><canvas id=\"pre-永\" class=\"preview-canvas\"></canvas></div>\n        <div class=\"preview-box\"><span class=\"preview-label\">校</span><canvas id=\"pre-校\" class=\"preview-canvas\"></canvas></div>\n        <div class=\"preview-box\"><span class=\"preview-label\">木</span><canvas id=\"pre-木\" class=\"preview-canvas\"></canvas></div>\n        <div class=\"preview-box\"><span class=\"preview-label\">文</span><canvas id=\"pre-文\" class=\"preview-canvas\"></canvas></div>\n      </div>\n      <div class=\"controls\" style=\"margin-top:20px;\">\n        <button onclick=\"resetParams()\" style=\"background:#8e8e93; font-size:13px;\">リセット</button>\n        <button onclick=\"toggleSettings(false)\">完了</button>\n      </div>\n    </div>\n  </div>\n\n  <div id=\"wizard-modal\" class=\"modal-overlay hidden\">\n    <div class=\"modal-content\" style=\"max-width: 350px;\">\n      <h3>🎯 自動キャリブレーション</h3>\n      <p style=\"font-size:11px; color:#555; margin-top:0;\">文字をなぞって「今の画はどれですか？」とシステムに教えてください。あなたのペンの判定基準を学習します。</p>\n      \n      <div style=\"display:flex; align-items:center; gap:10px; margin-bottom:10px;\">\n        <span style=\"font-size:12px; font-weight:bold;\">下敷き:</span>\n        <select id=\"wiz-kanji-select\" onchange=\"drawWizGuide()\"></select>\n      </div>\n\n      <canvas id=\"wizard-canvas\"></canvas>\n\n      <div id=\"wiz-prompt-area\" class=\"hidden\">\n        <div style=\"text-align:center; font-weight:bold; font-size:13px; margin-bottom:6px; color:#e91e63;\">いま書いた画の意図は？</div>\n        <div class=\"wiz-btn-group\">\n          <button class=\"wiz-btn\" onclick=\"submitWizData('tome')\">とめ</button>\n          <button class=\"wiz-btn\" onclick=\"submitWizData('hane')\">はね</button>\n          <button class=\"wiz-btn\" onclick=\"submitWizData('harai')\">はらい</button>\n          <button class=\"wiz-btn-ignore\" onclick=\"submitWizData('ignore')\">いずれでもない（無視・消去）</button>\n        </div>\n      </div>\n\n      <div class=\"wiz-stats\" style=\"margin-top:10px;\">\n        <div style=\"font-weight:bold; margin-bottom:3px;\">📊 学習データ記録数</div>\n        <div id=\"wiz-counts\">とめ: 0 / はね: 0 / はらい: 0</div>\n        <hr style=\"border:0; border-top:1px solid #ddd; margin:8px 0;\">\n        <div style=\"font-weight:bold; margin-bottom:3px;\">⚙️ 計算されたしきい値</div>\n        <div id=\"wiz-thresholds\" style=\"font-family:monospace; color:#1a73e8; font-size:11px;\">未計算</div>\n      </div>\n\n      <hr style=\"border:0; border-top:2px solid #eee; margin:15px 0;\">\n      <div style=\"font-size:12px; font-weight:bold; margin-bottom:8px;\">💾 プロフィール保存先</div>\n      <div id=\"profile-slots-area\"></div>\n\n      <div class=\"controls\" style=\"margin-top:15px;\">\n        <button onclick=\"resetWizard()\" style=\"background:#f44336; font-size:12px;\">記録リセット</button>\n        <button onclick=\"toggleWizard(false)\" style=\"background:#8e8e93; font-size:12px;\">キャンセル</button>\n      </div>\n    </div>\n  </div>\n\n<script>\n  // ==========================================\n  // 1. 高解像度Canvasセットアップ機構\n  // ==========================================\n  function setupHighDPICanvas(canvasId, logicalWidth, logicalHeight) {\n    const canvas = document.getElementById(canvasId);\n    if (!canvas) return null;\n    const dpr = window.devicePixelRatio || 1;\n    canvas.width = logicalWidth * dpr;\n    canvas.height = logicalHeight * dpr;\n    canvas.style.width = logicalWidth + 'px';\n    canvas.style.height = logicalHeight + 'px';\n    \n    const context = canvas.getContext('2d');\n    context.scale(dpr, dpr);\n    context.lineCap = 'round';\n    context.lineJoin = 'round';\n    return context;\n  }\n\n  const CANVAS_SIZE = 300;\n  const WIZ_CANVAS_SIZE = 280;\n  const PREVIEW_SIZE = 120;\n\n  const canvas = document.getElementById('canvas');\n  const ctx = setupHighDPICanvas('canvas', CANVAS_SIZE, CANVAS_SIZE);\n  const wizCanvas = document.getElementById('wizard-canvas');\n  const wctx = setupHighDPICanvas('wizard-canvas', WIZ_CANVAS_SIZE, WIZ_CANVAS_SIZE);\n\n  // ==========================================\n  // 2. グローバル変数宣言\n  // ==========================================\n  let KANJI_DATA = {};\n  const NUM_SAMPLES = 45; \n  let referenceStrokes = []; \n  let userStrokes = [];       \n  let mode = 'score';\n  let animationId = null; \n\n  const TEST_KANJI = {\n    '永': ['M45.5,13.25c5.12,2.4,8.62,5.75,10.75,9','M31.75,37.25c1.88,0.75,4.12,1,6.07,0.55c3.62-0.83,9.41-2.51,13.04-3.35c3-0.7,5.16,0.8,5.16,4.05c0,14.5-0.26,45.45-0.26,50c0,12-7.02,2.84-8.52,1.7','M14.25,58.68c1.75,0.45,3.46,0.3,5-0.02c2.5-0.53,12.84-3.54,15.34-4.49c2.5-0.95,4.65,0.77,3.75,2.85C33.5,68.25,25.75,79.75,16,86.37','M81.22,36.52c-0.1,1.11-0.78,2.03-1.52,2.7c-4.83,4.38-13.96,9.73-21.46,11.92','M58,52.74c9.88,9.52,20.02,18.85,29.07,24.84c2.01,1.33,4.05,2.66,6.43,3.56'],\n    '校': ['M11.53,40.68c1.1,0.32,2.6,0.45,4.53,0.32c5.4-0.35,16.57-3,23.14-4.04c1.25-0.2,2.3-0.18,3.07,0','M28.99,17.25c1.07,1.07,1.76,3.25,1.76,5.25c0,0.77-0.03,48.09-0.18,65.25c-0.03,3.03-0.05,5.16-0.07,6','M30.25,40.75c0,1.25-0.49,2.66-0.96,3.77C25.28,53.91,20.88,62.25,15,70','M33.75,51.25c2.75,1.5,6,5.25,7.25,7.75','M66.39,15.5c0.99,0.99,1.38,1.88,1.38,3.62c0,4.25-0.02,7.62-0.08,10.41','M48.12,31.71c2.3,0.29,3.9,0.44,6.09,0.2c10.28-1.16,20.32-2.66,32.45-3.53c2.35-0.17,4.03-0.01,5.33,0.32','M59.24,38.93c0.2,0.53,0.06,2.27-0.4,3.14C57,45.5,53.75,49.5,50,52.5','M79.27,38.5c4.34,3.07,8.73,8.68,10.9,12.41','M79.15,49.18c0.35,1.32,0.17,2.62-0.54,4.18C72.25,67.25,58.75,83.25,44,91.25','M55.95,55.88c6.3,3.37,21.64,22.12,31.45,30.33c2.64,2.21,5.07,4.15,8.6,4.44'],\n    '木': ['M19.5,39.86c2.45,0.57,5.23,0.8,8.04,0.57C40.75,39.38,63,36.5,79.78,36.15c2.8-0.06,4.54,0.1,7.34,0.5','M51.75,10.5c1.19,1.19,2,3,2,5c0,8.65,0,55.15-0.14,74.75c-0.03,4.19-0.07,7.15-0.11,8.25','M50.75,39.5c0,1.12-0.61,2.44-1.42,3.95C41.75,57.5,26.7,73.93,15.75,80.25','M54.5,39c4.62,6,23,25.75,31.76,34.61c2.27,2.29,4.61,4.39,7.49,5.64'],\n    '文': ['M51.62,12.75c1.06,1.06,1.73,2.5,1.73,4.01c0,4.32-0.11,7.61-0.11,12.15','M16.88,32.72c2.7,0.66,5.71,0.86,8.6,0.54c15.77-1.76,40.15-5.01,56.42-6.1c3.67-0.25,6.15,0.05,8.88,0.69','M69.89,32.5c0.36,2.12,0.06,3.82-0.93,6.27C61.62,57,43.75,80.25,18.75,93.75','M31,43.75c6,0,27.27,26.79,49.26,42.1c4.13,2.87,7.49,4.77,11.49,6.04']\n  };\n\n  // ==========================================\n  // 3. ストレージとプロファイル管理\n  // ==========================================\n  const SafeStorage = {\n    get: function(k) { try { return localStorage.getItem(k); } catch(e) { return null; } },\n    set: function(k, v) { try { localStorage.setItem(k, v); } catch(e) {} },\n    remove: function(k) { try { localStorage.removeItem(k); } catch(e) {} },\n    clearData: function() { try { for (let i = localStorage.length - 1; i >= 0; i--) { const k = localStorage.key(i); if (k && k.startsWith('kanjiData_')) localStorage.removeItem(k); } } catch(e) {} }\n  };\n\n  // 表現パラメータ\n  const DEFAULT_PARAMS = { baseWidth: 11.0, tomeStart: 0.90, tomeScale: 1.3, haneStart: 0.90, haneSharp: 0.55, haraiStart: 0.65, haraiSharp: 0.80 };\n  let STROKE_PARAMS = { ...DEFAULT_PARAMS };\n  const savedEff = SafeStorage.get('kanjiStrokeParamsV2');\n  if (savedEff) { try { STROKE_PARAMS = { ...DEFAULT_PARAMS, ...JSON.parse(savedEff) }; } catch(e) {} }\n\n  // 判定しきい値プロファイル\n  const DEFAULT_THRESHOLDS = { minVelocity: 1.2, hookAngleDiff: 0.6 };\n  let USER_PROFILES = [\n    { id: 0, name: \"汎用ペン (デフォルト)\", thresholds: { ...DEFAULT_THRESHOLDS }, isEmpty: false },\n    { id: 1, name: \"未登録 1\", thresholds: null, isEmpty: true },\n    { id: 2, name: \"未登録 2\", thresholds: null, isEmpty: true }\n  ];\n  let ACTIVE_PROFILE_ID = 0;\n\n  const savedProfiles = SafeStorage.get('kanjiPenProfiles');\n  if (savedProfiles) {\n    try { \n      const parsed = JSON.parse(savedProfiles);\n      USER_PROFILES = parsed.profiles || USER_PROFILES;\n      ACTIVE_PROFILE_ID = parsed.activeId !== undefined ? parsed.activeId : 0;\n    } catch(e) {}\n  }\n\n  function getActiveThresholds() {\n    const prof = USER_PROFILES[ACTIVE_PROFILE_ID];\n    return prof.thresholds ? prof.thresholds : DEFAULT_THRESHOLDS;\n  }\n\n  function updateProfileSelectUI() {\n    const sel = document.getElementById('active-profile-select');\n    sel.innerHTML = USER_PROFILES.map(p => `<option value=\"${p.id}\" ${p.id === ACTIVE_PROFILE_ID ? 'selected' : ''}>${p.isEmpty ? '(未登録)' : '🖋 ' + p.name}</option>`).join('');\n  }\n\n  function switchProfile(id) {\n    ACTIVE_PROFILE_ID = parseInt(id);\n    const prof = USER_PROFILES[ACTIVE_PROFILE_ID];\n    if (prof.isEmpty) alert(\"このスロットは未登録です。「ペン調整」から設定を記録してください。\");\n    SafeStorage.set('kanjiPenProfiles', JSON.stringify({ profiles: USER_PROFILES, activeId: ACTIVE_PROFILE_ID }));\n    if (referenceStrokes.length > 0) redrawAllUserStrokes(); \n  }\n\n  // ==========================================\n  // 4. 表現設定UIロジック\n  // ==========================================\n  const previewContexts = {}; \n\n  function toggleSettings(show) {\n    const modal = document.getElementById('settings-modal');\n    if (show) {\n      Object.keys(STROKE_PARAMS).forEach(k => {\n        const el = document.getElementById(`sl-${k}`);\n        if(el) { el.value = STROKE_PARAMS[k]; document.getElementById(`val-${k}`).innerText = STROKE_PARAMS[k]; }\n      });\n      modal.classList.remove('hidden'); updateAllPreviews();\n    } else {\n      modal.classList.add('hidden');\n      if (referenceStrokes.length > 0) redrawAllUserStrokes(); \n    }\n  }\n\n  function updateParam(k, v) { STROKE_PARAMS[k] = parseFloat(v); document.getElementById(`val-${k}`).innerText = v; SafeStorage.set('kanjiStrokeParamsV2', JSON.stringify(STROKE_PARAMS)); updateAllPreviews(); }\n  function resetParams() { if(confirm(\"エフェクト設定を初期化しますか？\")) { STROKE_PARAMS = { ...DEFAULT_PARAMS }; toggleSettings(true); } }\n\n  // ==========================================\n  // 5. キャリブレーション・ウィザード機能\n  // ==========================================\n  let wizDrawing = false, wizCurrentPoints = [], wizLastTime = 0, wizStats = null;\n  let wizHistory = { tome: [], hane: [], harai: [] };\n\n  function toggleWizard(show) {\n    const modal = document.getElementById('wizard-modal');\n    if (show) {\n      const sel = document.getElementById('wiz-kanji-select');\n      const loadedKeys = Object.keys(KANJI_DATA);\n      const keysToUse = loadedKeys.length > 0 ? loadedKeys : Object.keys(TEST_KANJI);\n      sel.innerHTML = `<option value=\"\">(なし・自由書き)</option>` + keysToUse.map(k => `<option value=\"${k}\">${k}</option>`).join('');\n      \n      wizHistory = { tome: [], hane: [], harai: [] };\n      renderProfileSaveSlots(); updateWizUI();\n      modal.classList.remove('hidden'); drawWizGuide();\n    } else {\n      modal.classList.add('hidden'); updateProfileSelectUI();\n    }\n  }\n\n  function drawWizGuide() {\n    wctx.clearRect(0, 0, WIZ_CANVAS_SIZE, WIZ_CANVAS_SIZE);\n    const char = document.getElementById('wiz-kanji-select').value;\n    if (!char) return;\n    let paths = TEST_KANJI[char] || KANJI_DATA[char];\n    if (!paths) return;\n\n    let strokes = paths.map(p => sampleSvgPath(p, 30));\n    let minX=Infinity, minY=Infinity, maxX=-Infinity, maxY=-Infinity;\n    strokes.forEach(s => s.forEach(p => { minX=Math.min(minX,p.x); minY=Math.min(minY,p.y); maxX=Math.max(maxX,p.x); maxY=Math.max(maxY,p.y); }));\n    \n    const pad = 30, scale = Math.min((WIZ_CANVAS_SIZE - pad * 2) / Math.max(1, maxX - minX), (WIZ_CANVAS_SIZE - pad * 2) / Math.max(1, maxY - minY));\n    const offX = (WIZ_CANVAS_SIZE - (maxX - minX) * scale) / 2 - minX * scale, offY = (WIZ_CANVAS_SIZE - (maxY - minY) * scale) / 2 - minY * scale;\n\n    wctx.fillStyle = '#e0e0e0';\n    strokes.forEach(s => { s.forEach(pt => { wctx.beginPath(); wctx.arc(pt.x * scale + offX, pt.y * scale + offY, 4, 0, Math.PI * 2); wctx.fill(); }); });\n  }\n\n  function clearWizCanvas() { wizCurrentPoints = []; wizStats = null; document.getElementById('wiz-prompt-area').classList.add('hidden'); drawWizGuide(); }\n\n  wizCanvas.addEventListener('pointerdown', e => {\n    wizCanvas.setPointerCapture(e.pointerId); wizDrawing = true; wizCurrentPoints = [];\n    const r = wizCanvas.getBoundingClientRect();\n    const pt = { x: (e.clientX - r.left)*(WIZ_CANVAS_SIZE/r.width), y: (e.clientY - r.top)*(WIZ_CANVAS_SIZE/r.height) };\n    wizCurrentPoints.push(pt); wizLastTime = Date.now();\n    wctx.fillStyle = '#e53935'; wctx.beginPath(); wctx.arc(pt.x, pt.y, STROKE_PARAMS.baseWidth/2, 0, Math.PI * 2); wctx.fill();\n    document.getElementById('wiz-prompt-area').classList.add('hidden');\n  });\n\n  wizCanvas.addEventListener('pointermove', e => {\n    if (!wizDrawing) return;\n    const r = wizCanvas.getBoundingClientRect();\n    const pt = { x: (e.clientX - r.left)*(WIZ_CANVAS_SIZE/r.width), y: (e.clientY - r.top)*(WIZ_CANVAS_SIZE/r.height) };\n    const last = wizCurrentPoints[wizCurrentPoints.length-1];\n    fillBetween(wctx, last, pt, STROKE_PARAMS.baseWidth); wizCurrentPoints.push(pt);\n  });\n\n  wizCanvas.addEventListener('pointerup', e => {\n    if (!wizDrawing) return;\n    wizDrawing = false;\n    if (wizCurrentPoints.length < 3) { clearWizCanvas(); return; }\n\n    const n = wizCurrentPoints.length, pStart = wizCurrentPoints[Math.max(0, n - 6)], pEnd = wizCurrentPoints[n - 1];\n    const velocity = Math.hypot(pEnd.x - pStart.x, pEnd.y - pStart.y) / ((Date.now() - wizLastTime) || 1) * 10;\n    const pMid = wizCurrentPoints[Math.floor(n * 0.5)], pPre = wizCurrentPoints[Math.floor(n * 0.85)];\n    const angleMain = Math.atan2(pPre.y - pMid.y, pPre.x - pMid.x), angleTip = Math.atan2(pEnd.y - pPre.y, pEnd.x - pPre.x);\n    let angleDiff = Math.abs(angleTip - angleMain); if (angleDiff > Math.PI) angleDiff = 2 * Math.PI - angleDiff;\n\n    wizStats = { v: velocity, a: angleDiff };\n    document.getElementById('wiz-prompt-area').classList.remove('hidden');\n  });\n\n  function submitWizData(type) {\n    if (type === 'ignore') { clearWizCanvas(); return; }\n    if (!wizStats) return;\n    wizHistory[type].push(wizStats); recalcThresholds(); updateWizUI(); clearWizCanvas();\n  }\n\n  let TEMPORARY_THRESHOLDS = { ...DEFAULT_THRESHOLDS };\n  function recalcThresholds() {\n    let maxTomeV = 0, minHaneHaraiV = 999;\n    if (wizHistory.tome.length > 0) maxTomeV = Math.max(...wizHistory.tome.map(d => d.v));\n    const hhArray = wizHistory.hane.concat(wizHistory.harai);\n    if (hhArray.length > 0) minHaneHaraiV = Math.min(...hhArray.map(d => d.v));\n\n    let newMinV = DEFAULT_THRESHOLDS.minVelocity;\n    if (wizHistory.tome.length > 0 && hhArray.length > 0) newMinV = (maxTomeV + minHaneHaraiV) / 2;\n    else if (wizHistory.tome.length > 0) newMinV = maxTomeV * 1.2; else if (hhArray.length > 0) newMinV = minHaneHaraiV * 0.8;\n\n    let newAngleDiff = DEFAULT_THRESHOLDS.hookAngleDiff;\n    if (wizHistory.hane.length > 0) newAngleDiff = Math.max(0.2, Math.min(...wizHistory.hane.map(d => d.a)) * 0.75); \n    \n    TEMPORARY_THRESHOLDS.minVelocity = parseFloat(newMinV.toFixed(2)); TEMPORARY_THRESHOLDS.hookAngleDiff = parseFloat(newAngleDiff.toFixed(2));\n  }\n\n  function updateWizUI() {\n    document.getElementById('wiz-counts').innerText = `とめ: ${wizHistory.tome.length} / はね: ${wizHistory.hane.length} / はらい: ${wizHistory.harai.length}`;\n    const t = wizHistory.tome.length + wizHistory.hane.length + wizHistory.harai.length > 0 ? TEMPORARY_THRESHOLDS : DEFAULT_THRESHOLDS;\n    document.getElementById('wiz-thresholds').innerHTML = `v >= ${t.minVelocity} で「払・跳」<br>a >= ${t.hookAngleDiff} で「跳ね」`;\n  }\n\n  function resetWizard() {\n    if (confirm(\"学習データをクリアしますか？\")) {\n      wizHistory = { tome: [], hane: [], harai: [] }; TEMPORARY_THRESHOLDS = { ...DEFAULT_THRESHOLDS };\n      updateWizUI(); clearWizCanvas();\n    }\n  }\n\n  function renderProfileSaveSlots() {\n    const area = document.getElementById('profile-slots-area');\n    area.innerHTML = USER_PROFILES.map(p => `\n      <div class=\"profile-slot\">\n        <input type=\"text\" id=\"prof-name-${p.id}\" value=\"${p.isEmpty ? '' : p.name}\" placeholder=\"ペン・端末名 (未登録)\">\n        <button style=\"background:${p.isEmpty ? '#4caf50' : '#f4b400'}; color:#fff;\" onclick=\"saveProfileToSlot(${p.id})\">${p.isEmpty ? 'ここに保存' : '上書き保存'}</button>\n      </div>`).join('');\n  }\n\n  function saveProfileToSlot(id) {\n    const t = wizHistory.tome.length + wizHistory.hane.length + wizHistory.harai.length > 0 ? TEMPORARY_THRESHOLDS : DEFAULT_THRESHOLDS;\n    const nameInput = document.getElementById(`prof-name-${id}`).value.trim() || `ペン設定 ${id+1}`;\n    if (!USER_PROFILES[id].isEmpty) if (!confirm(`「${USER_PROFILES[id].name}」に上書き保存しますか？\\n(以前の判定設定は消去されます)`)) return;\n    \n    USER_PROFILES[id] = { id: id, name: nameInput, thresholds: { ...t }, isEmpty: false };\n    ACTIVE_PROFILE_ID = id;\n    SafeStorage.set('kanjiPenProfiles', JSON.stringify({ profiles: USER_PROFILES, activeId: ACTIVE_PROFILE_ID }));\n    alert(`「${nameInput}」として保存し、現在のペンに設定しました。`);\n    renderProfileSaveSlots(); updateProfileSelectUI();\n  }\n\n  // ==========================================\n  // 6. アプリケーション初期化 (安全対策済)\n  // ==========================================\n  function showLoading(show, msg = \"ロード中...\") {\n    const overlay = document.getElementById('loading-overlay');\n    if (overlay) { overlay.style.display = show ? 'flex' : 'none'; if(show) overlay.querySelector('span').textContent = msg; }\n  }\n\n  window.onload = () => {\n    updateProfileSelectUI();\n    var sheetSelect = document.getElementById(\"sheet-selector\");\n    if (sheetSelect && !String(sheetSelect.innerHTML || \"\").trim()) {\n      sheetSelect.innerHTML = '<option value=\"KanjiVG_txt\">KanjiVG.txt</option>';\n    }\n    showLoading(true, \"初期化中...\");\n    loadKanjiData();\n  };\n\n  const KANJI_VG_CACHE_KEY = \"kanjiData_KanjiVG_txt\";\n\n  function parseUnicodeToChar(rawUnicode) {\n    const raw = String(rawUnicode || \"\").trim();\n    if (!raw) return \"\";\n    const hexMatch = raw.match(/^(?:U\\+|0x)?([0-9A-Fa-f]{4,6})$/);\n    if (hexMatch) {\n      const cp = parseInt(hexMatch[1], 16);\n      if (!isNaN(cp) && cp > 0) return String.fromCodePoint(cp);\n    }\n    const sci = Number(raw);\n    if (!isNaN(sci)) {\n      const cp = Math.round(sci);\n      if (cp > 0 && cp <= 0x10FFFF) return String.fromCodePoint(cp);\n    }\n    return \"\";\n  }\n\n  function parseKanjiVgTsv(text) {\n    const map = {};\n    String(text || \"\").split(/\\r?\\n/).forEach((line) => {\n      const row = line.trim();\n      if (!row) return;\n      const cols = row.split(\"\\t\");\n      if (cols.length < 3) return;\n      let kanji = String(cols[0] || \"\").trim();\n      const unicodeCol = String(cols[1] || \"\").trim();\n      const strokesCol = String(cols[2] || \"\").trim();\n      if (!kanji || kanji.length !== 1) {\n        const parsed = parseUnicodeToChar(unicodeCol);\n        if (parsed) kanji = parsed;\n      }\n      if (!kanji || !strokesCol) return;\n      const paths = strokesCol\n        .split(\"|\")\n        .map((p) => String(p || \"\").trim())\n        .filter((p) => p && (p.startsWith(\"M\") || p.startsWith(\"m\")));\n      if (paths.length > 0) map[kanji] = paths;\n    });\n    return map;\n  }\n\n  function syncData() {\n    try { SafeStorage.remove(KANJI_VG_CACHE_KEY); } catch (_) {}\n    loadKanjiData();\n  }\n\n  function loadKanjiData() {\n    if (window.__kpKanjiInflight) return window.__kpKanjiInflight;\n    const cached = SafeStorage.get(KANJI_VG_CACHE_KEY);\n    if (cached) {\n      try {\n        KANJI_DATA = JSON.parse(cached);\n        populateKanjiSelector();\n        showLoading(false);\n        return Promise.resolve();\n      } catch (_) {\n        SafeStorage.remove(KANJI_VG_CACHE_KEY);\n      }\n    }\n    showLoading(true, \"KanjiVG.txt を読込中...\");\n    window.__kpKanjiInflight = fetch(\"./KanjiVG.txt\")\n      .then((r) => {\n        if (!r.ok) throw new Error(\"HTTP \" + r.status);\n        return r.text();\n      })\n      .then((text) => {\n        const data = parseKanjiVgTsv(text);\n        if (!Object.keys(data).length) throw new Error(\"KanjiVG.txt が空または形式が不正です\");\n        KANJI_DATA = data;\n        SafeStorage.set(KANJI_VG_CACHE_KEY, JSON.stringify(data));\n        populateKanjiSelector();\n        showLoading(false);\n      })\n      .catch((e) => {\n        alert(\"KanjiVG.txt の読込に失敗: \" + (e && e.message ? e.message : e));\n        showLoading(false);\n      })\n      .finally(() => {\n        window.__kpKanjiInflight = null;\n      });\n    return window.__kpKanjiInflight;\n  }\n\n  function applyQuizKanjiData(entries) {\n    try {\n      if (!entries || typeof entries !== \"object\") return;\n      Object.keys(entries).forEach(function (k) {\n        var paths = entries[k];\n        if (Array.isArray(paths) && paths.length) KANJI_DATA[k] = paths;\n      });\n      var select = document.getElementById(\"target-kanji\");\n      if (!select) return;\n      var keys = Object.keys(entries).filter(function (k) { return KANJI_DATA[k]; });\n      select.innerHTML = keys.length\n        ? keys.map(function (k) { return '<option value=\"' + k + '\">' + k + '</option>'; }).join(\"\")\n        : \"<option>データなし</option>\";\n      notifyParentKanjiDataReady();\n    } catch (e) {}\n  }\n\n  function notifyParentKanjiDataReady() {\n    try {\n      var n = Object.keys(KANJI_DATA || {}).length;\n      if (n < 1 || !window.parent || window.parent === window) return;\n      window.parent.postMessage({ type: \"kpKanjiDataReady\", count: n }, \"*\");\n    } catch (_) {}\n  }\n\n  function populateKanjiSelector() {\n    const select = document.getElementById(\"target-kanji\");\n    const prevSelected = select && typeof select.value === \"string\" ? select.value : \"\";\n    const pending =\n      typeof window.__kpPendingKanjiSelect === \"string\" && window.__kpPendingKanjiSelect\n        ? window.__kpPendingKanjiSelect\n        : \"\";\n    const keys = Object.keys(KANJI_DATA);\n    select.innerHTML = keys.length\n      ? keys.map((k) => `<option value=\"${k}\">${k}</option>`).join(\"\")\n      : \"<option>データなし</option>\";\n    var restore = \"\";\n    if (pending && keys.indexOf(pending) >= 0) restore = pending;\n    else if (prevSelected && keys.indexOf(prevSelected) >= 0) restore = prevSelected;\n    if (restore) select.value = restore;\n    if (pending && keys.indexOf(pending) >= 0) {\n      try {\n        delete window.__kpPendingKanjiSelect;\n      } catch (_) {\n        window.__kpPendingKanjiSelect = \"\";\n      }\n    }\n    initTargetKanji();\n    notifyParentKanjiDataReady();\n  }\n\n  function initTargetKanji() {\n    const char = document.getElementById('target-kanji').value;\n    if(!char || !KANJI_DATA[char]) return;\n    try {\n      referenceStrokes = KANJI_DATA[char].map(p => {\n        const points = sampleSvgPath(p, NUM_SAMPLES);\n        return { type: detectStrokeType(points, false), points: points };\n      });\n      switchMode('score');\n    } catch(e) {\n      console.warn(\"SVG解析スキップ:\", char);\n    }\n  }\n\n  // ==========================================\n  // 7. 解析と描画エンジン (ゼロ除算保護済)\n  // ==========================================\n  function sampleSvgPath(d, n) {\n    if (!d || typeof d !== 'string' || d.trim() === '') return Array(n).fill({x:0, y:0});\n    try {\n      const p = document.createElementNS('http://www.w3.org/2000/svg', 'path'); p.setAttribute('d', d); const len = p.getTotalLength();\n      if (len === 0) return Array(n).fill({x:0, y:0});\n      return Array.from({length: n}, (_, i) => { const pt = p.getPointAtLength(i * len / (Math.max(1, n - 1))); return {x: pt.x, y: pt.y}; });\n    } catch(e) { return Array(n).fill({x:0, y:0}); }\n  }\n\n  function detectStrokeType(points, isUser = false, terminalInfo = null) {\n    if (points.length < 5) return 'tome';\n    const n = points.length, thr = getActiveThresholds(); \n    \n    if (!isUser) {\n        const pMid = points[Math.floor(n * 0.5)], pPre = points[Math.floor(n * 0.85)], pEnd = points[n - 1];\n        const angleMain = Math.atan2(pPre.y - pMid.y, pPre.x - pMid.x), angleTip = Math.atan2(pEnd.y - pPre.y, pEnd.x - pPre.x);\n        let diff = Math.abs(angleTip - angleMain); if (diff > Math.PI) diff = 2 * Math.PI - diff;\n        if (diff > thr.hookAngleDiff) return 'hane';\n        const deg = angleMain * 180 / Math.PI;\n        if ((deg > 10 && deg < 85) || (deg > 95 && deg < 170)) return 'harai';\n        return 'tome';\n    } \n    \n    if (terminalInfo) {\n        const { velocity, dy, angleDiff } = terminalInfo;\n        if (velocity < thr.minVelocity) return 'tome';\n        if (dy < 0 || angleDiff > thr.hookAngleDiff) return 'hane';\n        return 'harai';\n    }\n    return 'tome';\n  }\n\n  function renderStrokeWithEffect(context, points, type, color, scale = 1.0) {\n    if (points.length === 0) return;\n    context.fillStyle = color;\n    const bW = STROKE_PARAMS.baseWidth * scale;\n    if (points.length === 1) { context.beginPath(); context.arc(points[0].x, points[0].y, bW / 2, 0, Math.PI * 2); context.fill(); return; }\n\n    const n = points.length;\n    for (let i = 0; i < n - 1; i++) {\n      const ratio = i / (n - 1);\n      let w = bW;\n      \n      // ★ ゼロ除算・NaN対策\n      if (type === 'harai' && ratio > STROKE_PARAMS.haraiStart) {\n        const denom = (1 - STROKE_PARAMS.haraiStart) || 0.001;\n        w *= Math.max(0.1, 1 - STROKE_PARAMS.haraiSharp * ((ratio - STROKE_PARAMS.haraiStart) / denom));\n      } else if (type === 'hane' && ratio > STROKE_PARAMS.haneStart) {\n        const denom = (1 - STROKE_PARAMS.haneStart) || 0.001;\n        w *= Math.max(0.1, 1 - STROKE_PARAMS.haneSharp * ((ratio - STROKE_PARAMS.haneStart) / denom));\n      } else if (type === 'tome' && ratio > STROKE_PARAMS.tomeStart) {\n        const denom = (1 - STROKE_PARAMS.tomeStart) || 0.001;\n        w *= (1 + (STROKE_PARAMS.tomeScale - 1) * ((ratio - STROKE_PARAMS.tomeStart) / denom));\n      }\n      fillBetween(context, points[i], points[i+1], Math.max(0.5, w));\n    }\n  }\n\n  function fillBetween(context, start, end, width) {\n    const dist = Math.hypot(end.x - start.x, end.y - start.y); \n    const steps = Math.max(1, Math.ceil(dist * 2.5)); \n    for (let i = 0; i <= steps; i++) {\n      const t = i / steps; context.beginPath(); context.arc(start.x + (end.x - start.x) * t, start.y + (end.y - start.y) * t, width / 2, 0, Math.PI * 2); context.fill();\n    }\n  }\n\n  // ==========================================\n  // 8. 手書き処理（メインキャンバス）\n  // ==========================================\n  let isDrawingMain = false, currentPointsMain = [], mainLastTime = 0, mainLastPos = null;\n\n  function getMainXY(e) {\n    const r = canvas.getBoundingClientRect();\n    return { x: (e.clientX - r.left) * (CANVAS_SIZE / r.width), y: (e.clientY - r.top) * (CANVAS_SIZE / r.height), t: Date.now(), p: e.pressure || 0.5 };\n  }\n\n  canvas.addEventListener('pointerdown', e => {\n    if (mode === 'demo') return;\n    canvas.setPointerCapture(e.pointerId);\n    isDrawingMain = true; currentPointsMain = []; mainLastPos = getMainXY(e); mainLastTime = Date.now();\n    currentPointsMain.push({x: mainLastPos.x, y: mainLastPos.y});\n    ctx.fillStyle = (mode === 'trace' ? '#e53935' : '#333');\n    ctx.beginPath(); ctx.arc(mainLastPos.x, mainLastPos.y, STROKE_PARAMS.baseWidth / 2, 0, Math.PI * 2); ctx.fill();\n  });\n\n  canvas.addEventListener('pointermove', e => {\n    if (!isDrawingMain) return;\n    const pos = getMainXY(e);\n    ctx.fillStyle = (mode === 'trace' ? '#e53935' : '#333');\n    fillBetween(ctx, mainLastPos, pos, STROKE_PARAMS.baseWidth); \n    currentPointsMain.push({x: pos.x, y: pos.y}); mainLastPos = pos;\n  });\n\n  canvas.addEventListener('pointerup', e => {\n    if (!isDrawingMain) return;\n    isDrawingMain = false;\n    if (currentPointsMain.length < 3) { \n      userStrokes.push({ type: 'tome', points: [...currentPointsMain] }); redrawAllUserStrokes(); return; \n    }\n\n    const n = currentPointsMain.length, pStart = currentPointsMain[Math.max(0, n - 6)], pEnd = currentPointsMain[n - 1];\n    const velocity = Math.hypot(pEnd.x - pStart.x, pEnd.y - pStart.y) / ((Date.now() - mainLastTime) || 1) * 10;\n    const dy = pEnd.y - pStart.y;\n    const pMid = currentPointsMain[Math.floor(n * 0.5)], pPre = currentPointsMain[Math.floor(n * 0.85)];\n    const angleMain = Math.atan2(pPre.y - pMid.y, pPre.x - pMid.x), angleTip = Math.atan2(pEnd.y - pPre.y, pEnd.x - pPre.x);\n    let angleDiff = Math.abs(angleTip - angleMain); if (angleDiff > Math.PI) angleDiff = 2 * Math.PI - angleDiff;\n\n    const type = detectStrokeType(currentPointsMain, true, { velocity, dy, angleDiff });\n    userStrokes.push({ type, points: [...currentPointsMain] });\n    redrawAllUserStrokes();\n  });\n\n  function redrawAllUserStrokes() {\n    ctx.clearRect(0, 0, CANVAS_SIZE, CANVAS_SIZE);\n    if (mode === 'trace') drawTraceGuide();\n    userStrokes.forEach(s => { renderStrokeWithEffect(ctx, s.points, s.type, (mode === 'trace' ? '#e53935' : '#333')); });\n  }\n\n  function removeLastEffect() {\n    if (userStrokes.length > 0) {\n      userStrokes[userStrokes.length - 1].type = 'none';\n      redrawAllUserStrokes();\n    }\n  }\n\n  function updateAllPreviews() {\n    ['永', '校', '木', '文'].forEach(char => {\n      if (!previewContexts[char]) { previewContexts[char] = setupHighDPICanvas(`pre-${char}`, PREVIEW_SIZE, PREVIEW_SIZE); }\n      const pctx = previewContexts[char]; if (!pctx) return;\n      pctx.clearRect(0, 0, PREVIEW_SIZE, PREVIEW_SIZE);\n      const paths = TEST_KANJI[char]; if (!paths) return;\n      const strokes = paths.map(p => { return { type: detectStrokeType(sampleSvgPath(p, NUM_SAMPLES), false), points: sampleSvgPath(p, NUM_SAMPLES) }; });\n      const off = 10; strokes.forEach(s => { const scaledPoints = s.points.map(pt => ({x: pt.x + off, y: pt.y + off})); renderStrokeWithEffect(pctx, scaledPoints, s.type, \"#333\", 0.4); });\n    });\n  }\n\n  // ==========================================\n  // 9. モード切替と採点エンジン（厳密判定対応）\n  // ==========================================\n  function switchMode(m) {\n    mode = m; const msg = document.getElementById('mode-msg');\n    ['score-controls', 'trace-controls', 'demo-controls'].forEach(id => document.getElementById(id).classList.add('hidden'));\n    document.getElementById(`${mode}-controls`).classList.remove('hidden');\n    if (mode === 'demo') document.getElementById('result-box').classList.add('hidden'); else document.getElementById('result-box').classList.remove('hidden');\n    \n    clearCanvas();\n    if(mode==='trace') msg.innerHTML = \"お手本を<span style='color:#e53935'>赤ペン</span>でなぞってください\";\n    else if(mode==='demo') { msg.innerText = \"書き順を確認しましょう\"; playAnimation(); }\n    else msg.innerText = \"正確に書いてください\";\n  }\n\n  function clearCanvas() {\n    if(animationId) cancelAnimationFrame(animationId);\n    ctx.clearRect(0, 0, CANVAS_SIZE, CANVAS_SIZE); userStrokes = [];\n    if(mode === 'trace') drawTraceGuide();\n    document.getElementById('score').innerText = \"- 点\"; document.getElementById('msg').innerText = \"書いて「採点する」を押してください\";\n  }\n\n  function getScaledRefs() {\n    let minX=Infinity, minY=Infinity, maxX=-Infinity, maxY=-Infinity;\n    referenceStrokes.forEach(s => { s.points.forEach(p => { minX = Math.min(minX, p.x); minY = Math.min(minY, p.y); maxX = Math.max(maxX, p.x); maxY = Math.max(maxY, p.y); }); });\n    const pad = 40, scale = Math.min((CANVAS_SIZE - pad * 2) / Math.max(1, maxX - minX), (CANVAS_SIZE - pad * 2) / Math.max(1, maxY - minY));\n    const offX = (CANVAS_SIZE - (maxX - minX) * scale) / 2 - minX * scale, offY = (CANVAS_SIZE - (maxY - minY) * scale) / 2 - minY * scale;\n    return referenceStrokes.map(s => ({ type: s.type, points: s.points.map(p => ({x: p.x * scale + offX, y: p.y * scale + offY})) }));\n  }\n\n  function drawTraceGuide() { getScaledRefs().forEach(s => renderStrokeWithEffect(ctx, s.points, s.type, \"#f0f0f0\")); }\n\n  function playAnimation() {\n    if(animationId) cancelAnimationFrame(animationId);\n    const refs = getScaledRefs(); let sIdx = 0, pIdx = 0;\n    function frame() {\n      ctx.clearRect(0, 0, CANVAS_SIZE, CANVAS_SIZE);\n      refs.forEach(s => renderStrokeWithEffect(ctx, s.points, s.type, \"#f5f5f5\"));\n      for (let i = 0; i < sIdx; i++) renderStrokeWithEffect(ctx, refs[i].points, refs[i].type, \"#9e9e9e\");\n      if (sIdx < refs.length) {\n        renderStrokeWithEffect(ctx, refs[sIdx].points.slice(0, pIdx + 1), refs[sIdx].type, \"#9e9e9e\");\n        const cur = refs[sIdx].points[pIdx]; ctx.beginPath(); ctx.arc(cur.x, cur.y, STROKE_PARAMS.baseWidth/2, 0, Math.PI * 2); ctx.fillStyle = '#e53935'; ctx.fill();\n        pIdx++;\n        if (pIdx >= refs[sIdx].points.length) { sIdx++; pIdx = 0; setTimeout(() => { if (mode === 'demo') animationId = requestAnimationFrame(frame); }, 400); return; }\n      } else return;\n      if (mode === 'demo') animationId = requestAnimationFrame(frame);\n    }\n    frame();\n  }\n\n  // ==========================================\n  // 9b. 採点用ヘルパ（リサンプル・正規化・DTW・交差・サイズ）\n  // ==========================================\n  // 採点をやや易しくする都合: 重み（軌道の比重↓・サイズ↓をやや緩和寄り）と下記スケールを併用\n  const SCORE_WEIGHTS = { trajectory: 0.40, startEnd: 0.22, structure: 0.13, size: 0.25 };\n\n  function resamplePolyline(sArray, n) {\n    const s = [...sArray];\n    if (s.length < 2) return Array(n).fill(s[0] || {x:0, y:0});\n    let len = 0; for (let i = 1; i < s.length; i++) len += Math.hypot(s[i].x - s[i-1].x, s[i].y - s[i-1].y);\n    const step = len / (n - 1), res = [s[0]]; let d = 0, i = 1;\n    while (i < s.length && res.length < n) {\n      const seg = Math.hypot(s[i].x - s[i-1].x, s[i].y - s[i-1].y);\n      if (d + seg >= step) { const t = (step - d) / Math.max(0.1, seg); res.push({ x: s[i-1].x + (s[i].x - s[i-1].x) * t, y: s[i-1].y + (s[i].y - s[i-1].y) * t }); s.splice(i, 0, res[res.length - 1]); d = 0; }\n      else { d += seg; i++; }\n    }\n    while (res.length < n) res.push(s[s.length - 1]); return res;\n  }\n\n  function normalizeStrokesToUnitSquare(strokes) {\n    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;\n    strokes.forEach(s => { s.forEach(p => { minX = Math.min(minX, p.x); minY = Math.min(minY, p.y); maxX = Math.max(maxX, p.x); maxY = Math.max(maxY, p.y); }); });\n    const scale = Math.max(maxX - minX, maxY - minY) || 1;\n    return strokes.map(s => s.map(p => ({x: (p.x - minX) / scale, y: (p.y - minY) / scale})));\n  }\n\n  function dtwMeanDistance(a, b) {\n    const n = a.length, m = b.length;\n    if (n === 0 || m === 0) return 1;\n    const inf = 1e12;\n    const dtw = Array(n + 1);\n    for (let i = 0; i <= n; i++) {\n      dtw[i] = new Float64Array(m + 1);\n      dtw[i].fill(inf);\n    }\n    dtw[0][0] = 0;\n    for (let i = 1; i <= n; i++) {\n      for (let j = 1; j <= m; j++) {\n        const cost = Math.hypot(a[i - 1].x - b[j - 1].x, a[i - 1].y - b[j - 1].y);\n        dtw[i][j] = cost + Math.min(dtw[i - 1][j], dtw[i][j - 1], dtw[i - 1][j - 1]);\n      }\n    }\n    const pathLen = n + m - 1;\n    return dtw[n][m] / Math.max(1, pathLen);\n  }\n\n  function distTo100(meanDist, scale) {\n    return Math.max(0, Math.min(100, Math.round(100 * (1 - meanDist / scale))));\n  }\n\n  function calcTrajectoryScoreDtw(uN, rN) {\n    let sum = 0;\n    for (let i = 0; i < rN.length; i++) {\n      const m = dtwMeanDistance(uN[i], rN[i]);\n      sum += distTo100(m, 0.18);\n    }\n    return Math.round(sum / rN.length);\n  }\n\n  function calcStartEndScore(uN, rN) {\n    let total = 0, count = 0;\n    for (let i = 0; i < rN.length; i++) {\n      const us = uN[i][0], ue = uN[i][uN[i].length - 1], rs = rN[i][0], re = rN[i][rN[i].length - 1];\n      const dS = Math.hypot(us.x - rs.x, us.y - rs.y), dE = Math.hypot(ue.x - re.x, ue.y - re.y);\n      // 正規化座標上の先端/末端差。スケールを大きいほうが同じ距離で得点が上がる（易しい）\n      const pairMean = (dS + dE) / 2;\n      total += distTo100(pairMean, 0.30);\n      count++;\n    }\n    return count ? Math.round(total / count) : 0;\n  }\n\n  function cross2(ax, ay, bx, by) { return ax * by - ay * bx; }\n  function sub(a, b) { return {x: a.x - b.x, y: a.y - b.y}; }\n  function segIntersects(p1, p2, p3, p4) {\n    const r = sub(p2, p1), s = sub(p4, p3);\n    const cr = cross2(r.x, r.y, s.x, s.y);\n    if (Math.abs(cr) < 1e-9) return false;\n    const t = cross2(sub(p3, p1).x, sub(p3, p1).y, s.x, s.y) / cr;\n    const u = cross2(sub(p3, p1).x, sub(p3, p1).y, r.x, r.y) / cr;\n    return t >= 0 && t <= 1 && u >= 0 && u <= 1;\n  }\n\n  function strokePolylinesIntersect(A, B) {\n    for (let i = 0; i < A.length - 1; i++) {\n      for (let j = 0; j < B.length - 1; j++) {\n        if (segIntersects(A[i], A[i + 1], B[j], B[j + 1])) return true;\n      }\n    }\n    return false;\n  }\n\n  function collectIntersectingStrokePairs(polys) {\n    const pairs = new Set();\n    for (let i = 0; i < polys.length; i++) {\n      for (let j = i + 1; j < polys.length; j++) {\n        if (strokePolylinesIntersect(polys[i], polys[j])) {\n          const key = i * 256 + j;\n          pairs.add(key);\n        }\n      }\n    }\n    return pairs;\n  }\n\n  function calcStructureScore(uN, rN) {\n    const refPolys = rN, userPolys = uN;\n    const refSet = collectIntersectingStrokePairs(refPolys);\n    if (refSet.size === 0) return 100;\n    let matches = 0;\n    for (const key of refSet) {\n      const j = key % 256, i = (key - j) / 256;\n      if (strokePolylinesIntersect(userPolys[i], userPolys[j])) matches++;\n    }\n    // 交差の一致度に最低点を与え、わずかな不一致で大きく落ちにくくする\n    return Math.min(100, Math.round(22 + 78 * (matches / refSet.size)));\n  }\n\n  function boundingBoxFromStrokesInCanvas(strokes) {\n    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;\n    let any = false;\n    strokes.forEach(s => { s.forEach(p => { any = true; minX = Math.min(minX, p.x); minY = Math.min(minY, p.y); maxX = Math.max(maxX, p.x); maxY = Math.max(maxY, p.y); }); });\n    if (!any) return { w: 1, h: 1, area: 1 };\n    const w = Math.max(0.1, maxX - minX), h = Math.max(0.1, maxY - minY);\n    return { w, h, area: w * h };\n  }\n\n  function calcSizeScoreCanvas(userStrokesResampled, scaledRefStrokes) {\n    const uB = boundingBoxFromStrokesInCanvas(userStrokesResampled);\n    const rB = boundingBoxFromStrokesInCanvas(scaledRefStrokes.map(s => s.points));\n    const areaRatio = uB.area / rB.area;\n    const lo0 = 0.50, lo1 = 0.63, hi1 = 1.40, hi2 = 1.68, lo2 = 0.32, hi3 = 2.1;\n    if (areaRatio >= lo1 && areaRatio <= hi1) return 100;\n    if (areaRatio >= lo0 && areaRatio < lo1) return Math.round(90 + 10 * (areaRatio - lo0) / (lo1 - lo0));\n    if (areaRatio > hi1 && areaRatio <= hi2) return Math.round(90 + 10 * (hi2 - areaRatio) / (hi2 - hi1));\n    if (areaRatio < lo0 && areaRatio >= lo2) return Math.round(60 + 30 * (areaRatio - lo2) / (lo0 - lo2));\n    if (areaRatio > hi2 && areaRatio < hi3) return Math.round(60 + 30 * (hi3 - areaRatio) / (hi3 - hi2));\n    if (areaRatio < lo2) return Math.max(0, 40 - Math.min(40, Math.round(80 * (lo2 - areaRatio) / lo2)));\n    if (areaRatio > hi3) return Math.max(0, 40 - Math.min(40, Math.round(20 * (areaRatio - hi3))));\n    return 40;\n  }\n\n  function evaluateKanji(isTrace) {\n    const sEl = document.getElementById('score'), mEl = document.getElementById('msg');\n    if (userStrokes.length !== referenceStrokes.length) { sEl.innerText = \"0 点\"; mEl.innerText = `画数が違います（正解：${referenceStrokes.length}画）`; try { var _kan=document.getElementById('target-kanji'); var _kch=_kan&&_kan.value?String(_kan.value):\"\"; if(window.parent)window.parent.postMessage({type:\"kanjiQuizHandAnalytics\",handScore:0,passedThreshold:false,referenceStrokeCount:referenceStrokes.length,hasStrokeOrderIssue:false,strokeCountMismatch:true,brushEndingAllOk:false,kanjiChar:_kch},\"*\");} catch(_e) {} return; }\n\n    const userResampled = userStrokes.map(s => resamplePolyline(s.points, NUM_SAMPLES));\n    const uN = normalizeStrokesToUnitSquare(userResampled);\n    const rN = normalizeStrokesToUnitSquare(referenceStrokes.map(s => s.points));\n\n    const trajS = calcTrajectoryScoreDtw(uN, rN);\n    const seS = calcStartEndScore(uN, rN);\n    const strS = calcStructureScore(uN, rN);\n    const scaledRefs = getScaledRefs();\n    const sizeS = calcSizeScoreCanvas(userResampled, scaledRefs);\n\n    let score = Math.round(\n      SCORE_WEIGHTS.trajectory * trajS + SCORE_WEIGHTS.startEnd * seS + SCORE_WEIGHTS.structure * strS + SCORE_WEIGHTS.size * sizeS\n    );\n    score = Math.max(0, Math.min(100, score));\n\n    const isStrict = document.getElementById('cb-strict-mode').checked;\n    let typeErrors = 0;\n    if (isStrict) {\n      for (let i = 0; i < rN.length; i++) {\n        const rType = referenceStrokes[i].type;\n        const uType = userStrokes[i].type;\n        if (uType === 'none') continue;\n        // はね・はらいがあるべき場所にない場合は減点\n        if (rType === 'hane' && uType !== 'hane') {\n          typeErrors++; score = Math.max(0, score - 6);\n        } else if (rType === 'harai' && uType !== 'harai') {\n          typeErrors++; score = Math.max(0, score - 6);\n        }\n        // とめ(rType === 'tome')の箇所に不要なはね・はらいが存在しても減点しない\n      }\n    }\n\n    const uB2 = boundingBoxFromStrokesInCanvas(userResampled);\n    const rB2 = boundingBoxFromStrokesInCanvas(scaledRefs.map(s => s.points));\n    const areaRatio = (uB2.area / rB2.area) || 1;\n    const sizeLabel = (function () {\n      const lo0 = 0.50, lo1 = 0.63, hi1 = 1.40, hi2 = 1.68, lo2 = 0.32, hi3 = 2.1;\n      const ar = areaRatio;\n      if (ar >= lo1 && ar <= hi1) return '大きさは適正';\n      if (ar >= lo0 && ar < lo1) return '少し小さい';\n      if (ar > hi1 && ar <= hi2) return 'やや大きい';\n      if (ar < lo0 && ar >= lo2) return 'やや小さい';\n      if (ar > hi2 && ar < hi3) return 'やや大きい';\n      if (ar < lo2) return '小さすぎ';\n      if (ar > hi3) return '大きすぎ';\n      return '大きさは要調整';\n    })();\n\n    sEl.innerText = score + \" 点\";\n    mEl.style.color = '#555';\n    const subLine = `内訳: 軌道${trajS} 始終${seS} 構造${strS} 大きさ${sizeS}（${sizeLabel}）`;\n    if (isStrict && typeErrors > 0) {\n      mEl.style.color = '#e91e63';\n      mEl.innerText = `形は書けていますが、\\nとめ・はね・はらいのミスが ${typeErrors} 箇所あります。\\n` + subLine;\n    } else {\n      if (score > 85) mEl.innerText = \"素晴らしい！完璧に近いです。\\n\" + subLine;\n      else if (score > 65) mEl.innerText = `良い感じです。${sizeLabel}。お手本の形に近づけましょう。\\n` + subLine;\n      else mEl.innerText = `もう一歩。${sizeLabel}。ゆっくりなぞると形が安定しやすいです。\\n` + subLine;\n    }\n    try { var _bk=(!isStrict)||(typeErrors===0); var _hso=strS<42; var _kx=document.getElementById('target-kanji'); var _kcc=_kx&&_kx.value?String(_kx.value):\"\"; if(window.parent)window.parent.postMessage({type:\"kanjiQuizHandAnalytics\",handScore:score,passedThreshold:score>=60,referenceStrokeCount:referenceStrokes.length,hasStrokeOrderIssue:_hso,strokeCountMismatch:false,brushEndingAllOk:_bk,kanjiChar:_kcc},\"*\");} catch(_z) {}\n  }\n  (function(){try{if(window.parent&&window.parent!==window)window.parent.postMessage({type:\"kpIframeBootReady\"},\"*\");}catch(_e){}})();\n<\\/script>\n</body>\n</html>";

// ▼ GASのURLを貼り付けてください（※前回と同じURLならそのままで大丈夫） ▼
    const GAS_API_URL = "https://script.google.com/macros/s/AKfycbwlge5DwehGNIcovI9yPIvlxdnL4MG6mULsVSdHyiwbrUcopX-ea4R3JLxR6bLh_cSUZw/exec";
    
    let selectedUserId = null; let currentPin = ""; let isPinResetMode = false;
    let currentModeId = ""; let currentModeName = ""; let currentUnitName = "";
    let currentQuestions = []; let filteredQuestions = []; let currentQuestionIndex = 0; 
    let quizResults = []; let questionStartTime = 0; let currentIsReviewMode = false;
    let userInventoryData = []; let showAllInventory = false;
    function getUserIdForPref() {
      try {
        const user = JSON.parse(localStorage.getItem('app_kid_user'));
        return user && user.id ? user.id : 'guest';
      } catch(e) { return 'guest'; }
    }
    function getUserPref(key, defaultVal) {
      const uid = getUserIdForPref();
      const val = localStorage.getItem(`${uid}_${key}`);
      if (val !== null) return val;
      const oldVal = localStorage.getItem(key);
      return oldVal !== null ? oldVal : defaultVal;
    }
    function setUserPref(key, val) {
      const uid = getUserIdForPref();
      localStorage.setItem(`${uid}_${key}`, val);
    }
    function clearAppCacheAndReload() {
      if (!confirm("最新の学習データと設定を読み込み直します。\\nよろしいですか？（画面が再読み込みされます）")) return;
      Object.keys(localStorage).forEach(k => {
        if (k.startsWith('app_cached_')) {
          localStorage.removeItem(k);
        }
      });
      location.reload();
    }
    let appSettings = {}; let maxDeduction = 0; let autoNextTimer = null; let materialsData = []; let externalMenus = [];
    let kanjiQuizSession = null; let lastKanjiQuizContext = null;
    let __kanjiQuizSubmitInFlight = false;
    let currentMaterialsCategory = "english";
    let stopwatches = {
      home: { timerId: null, startAt: 0, elapsed: 0 },
      external: { timerId: null, startAt: 0, elapsed: 0 }
    };
    
    // ★ 穴埋め用ステート
    let fillBlanksData = []; 
    let activeFillBlankIndex = 0;
    let voiceFailCount = 0;
    const LS_INPUT_MODE = 'input_method_mode';
    const LS_QUIZ_RECOVERY_DRAFT = 'quiz_recovery_draft_v1';
    const LS_KANJI_QUIZ_RECOVERY_DRAFT = 'kanji_quiz_recovery_draft_v1';
    const LS_PEN_GUIDE_SPREAD = 'pen_guide_line_spread';
    const LS_PEN_GUIDE_SHOW = 'pen_guide_lines_show';
    const LS_PEN_CANVAS_MAX_WIDTH = 'pen_canvas_max_width_px';
    let inputMethodMode = getUserPref(LS_INPUT_MODE, 'keyboard');
    let isPenRecognitionInFlight = false;
    let resumePromptShownToken = "";
    let kanjiResumePromptShownToken = "";
    let handwritingState = {
      isDrawing: false,
      currentStroke: [],
      allStrokes: [],
      strokesBackupBeforeClear: null,
      pointerId: null,
      pointerType: '',
      activePenRect: null,
      lastTapAt: 0,
      penWidth: parseInt(getUserPref('pen_width', '4'), 10),
      penMode: getUserPref('pen_mode', 'pen'),
      guideLineSpread: (() => {
        const v = parseInt(getUserPref(LS_PEN_GUIDE_SPREAD, '50'), 10);
        return isNaN(v) ? 50 : Math.max(0, Math.min(100, v));
      })(),
      showGuideLines: getUserPref(LS_PEN_GUIDE_SHOW, '1') !== '0',
      canvasMaxWidthPx: (() => {
        const w = parseInt(getUserPref(LS_PEN_CANVAS_MAX_WIDTH, '1000'), 10);
        if (!isNaN(w)) return Math.max(400, Math.min(1200, w));
        const legacyH = parseInt(getUserPref('pen_canvas_height_px', ''), 10);
        if (!isNaN(legacyH)) return Math.max(400, Math.min(1200, Math.round(legacyH * 1.2)));
        return 1000;
      })(),
      pendingCandidates: [],
      pendingTempText: ""
    };

    const digitWords1to19 = ["zero","one","two","three","four","five","six","seven","eight","nine","ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"];
    const digitWordsTens = ["","","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"];
    function numberToWordsEn(num) {
      num = Number(num);
      if (!isFinite(num) || num < 0 || num > 9999) return String(num);
      if (num < 20) return digitWords1to19[num];
      if (num < 100) {
        const t = Math.floor(num / 10), r = num % 10;
        return digitWordsTens[t] + (r ? " " + digitWords1to19[r] : "");
      }
      if (num < 1000) {
        const h = Math.floor(num / 100), r = num % 100;
        return digitWords1to19[h] + " hundred" + (r ? " " + numberToWordsEn(r) : "");
      }
      const th = Math.floor(num / 1000), r = num % 1000;
      return digitWords1to19[th] + " thousand" + (r ? " " + numberToWordsEn(r) : "");
    }
    function numberToWordsSplitYear(num) {
      // 1000-9999 を「twenty twenty five」のように 2桁+2桁で読む
      num = Number(num);
      if (!isFinite(num) || num < 1000 || num > 9999) return "";
      const firstTwo = Math.floor(num / 100);
      const lastTwo = num % 100;
      return numberToWordsEn(firstTwo) + (lastTwo ? " " + numberToWordsEn(lastTwo) : "");
    }
    function numberToWordsVariants(numStr) {
      const n = Number(numStr);
      const out = new Set();
      out.add(numStr); // 数字そのもの
      if (!isFinite(n) || n < 0 || n > 9999) return [...out];
      out.add(numberToWordsEn(n)); // 標準
      if (n >= 1000 && n <= 3000) {
        const yearStyle = numberToWordsEn(n);
        out.add(yearStyle);
        const split = numberToWordsSplitYear(n);
        if (split) out.add(split);
      }
      return [...out];
    }
    function expandTextVariants(text) {
      const variants = [""];
      const str = String(text == null ? "" : text);
      let last = 0;
      str.replace(/\d+/g, (m, idx) => {
        const before = str.slice(last, idx);
        const numberForms = numberToWordsVariants(m);
        const next = [];
        variants.forEach(v => {
          numberForms.forEach(form => next.push(v + before + form));
        });
        variants.length = 0;
        variants.push(...next);
        last = idx + m.length;
      });
      if (last < str.length) {
        variants.forEach((v, i) => variants[i] = v + str.slice(last));
      }
      return variants;
    }
    function convertDigitsToWordsInText(text) {
      return String(text == null ? "" : text).replace(/\d+/g, n => numberToWordsEn(n));
    }
    const normalizeText = (text) => String(text == null ? "" : text).toLowerCase().replace(/[.,\-!?。、！？]/g, " ").replace(/\s+/g, ' ').trim();

    /** showQuestion と同じルールで正解文字列を得る（欠損データでも例外にしない） */
    function getCorrectAnswerForQuestion(q, format) {
      if (format === "ja_to_en") return String((q["英単語"] || q["英文"] || "")).trim();
      if (format === "en_to_ja") return String(q["日本語"] || "").trim();
      if (format === "qtext_to_en") return String(q["英文"] || "").trim();
      if (format === "en_audio_to_ja") return String(q["日本語"] || "").trim();
      if (format === "qaudio_to_en") return String(q["英文"] || "").trim();
      if (format === "en_audio_to_en") return String(q["英単語"] || q["英文"] || "").trim();
      if (format === "en_to_en") return String(q["英単語"] || q["英文"] || "").trim();
      if (format === "ja_to_en_sort") return getSortPrimaryCorrectDisplay(q);
      return "";
    }

    function parseSortPhraseTokens(q) {
      const list = [];
      for (let i = 1; i <= 40; i++) {
        const key = "並び替え語句" + i;
        if (q[key] != null && String(q[key]).trim() !== "") list.push(String(q[key]).trim());
      }
      return list;
    }

    function getDummyTokenForSort(q) {
      const raw = String(q["並び替え語句ダミー"] || "").trim();
      if (!raw) return "";
      const parts = raw.split(/[，,]/).map(s => s.trim()).filter(Boolean);
      return parts[0] || "";
    }

    function getSortPrimaryCorrectDisplay(q) {
      return String(q["並び替え箇所"] || q["英文"] || "").trim();
    }

    function getSortAcceptableRawStrings(q) {
      const list = [];
      const add = (s) => { const t = String(s || "").trim(); if (t) list.push(t); };
      add(q["並び替え箇所"]);
      add(q["英文"]);
      for (let i = 1; i <= 5; i++) add(q["別解" + i]);
      return [...new Set(list)];
    }

    function isSortAnswerCorrect(userJoined, q) {
      const u = normalizeText(userJoined);
      if (!u) return false;
      for (const s of getSortAcceptableRawStrings(q)) {
        const variants = expandTextVariants(s).map(t => normalizeText(t));
        if (variants.some(v => v === u)) return true;
      }
      return false;
    }

    function shuffleArrayInPlace(arr) {
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const t = arr[i]; arr[i] = arr[j]; arr[j] = t;
      }
      return arr;
    }

    let sortQuizState = null;

    function destroySortQuizIfAny() {
      if (sortQuizState && Array.isArray(sortQuizState.sortables)) {
        sortQuizState.sortables.forEach(s => { try { s.destroy(); } catch (e) {} });
      }
      sortQuizState = null;
      window.__sortMissingModeActive = false;
    }

    function checkSortQuizSubmit() {
      const btn = document.getElementById("sort-submit-btn");
      if (!btn || !sortQuizState) return;
      const slot = document.getElementById("sort-answer-slot");
      if (!slot) return;
      const ok = slot.children.length === sortQuizState.expectedCount;
      btn.style.display = ok ? "block" : "none";
    }

    function addSortMissingWordCard() {
      const inp = document.getElementById("sort-missing-input");
      const pool = document.getElementById("sort-word-pool");
      if (!inp || !pool) return;
      const val = inp.value.trim();
      if (!val) return;
      const card = document.createElement("div");
      card.className = "word-card";
      card.textContent = val;
      card.dataset.userMade = "1";
      pool.appendChild(card);
      inp.value = "";
      checkSortQuizSubmit();
    }

    function submitSortQuizAnswer() {
      const q = filteredQuestions[currentQuestionIndex];
      const slot = document.getElementById("sort-answer-slot");
      if (!slot || !sortQuizState) return;
      const userJoined = Array.from(slot.children).map(c => c.textContent).join(" ");
      checkAnswer(userJoined, sortQuizState.primaryCorrect, q);
    }

    function placeQuizFeedbackAtTopOfAnswerArea(answerAreaEl) {
      const fb = document.getElementById("quiz-feedback");
      if (!fb || !answerAreaEl) return;
      if (answerAreaEl.firstChild !== fb) answerAreaEl.insertBefore(fb, answerAreaEl.firstChild);
    }

    function setupSortQuiz(q, answerType) {
      const variant = answerType === "sort_all" ? "all" : answerType === "sort_dummy" ? "dummy" : "missing";
      const tokens = parseSortPhraseTokens(q);
      const expectedCount = tokens.length;
      let poolTokens = tokens.slice();
      let missingIdx = -1;
      if (variant === "missing") {
        missingIdx = Math.floor(Math.random() * poolTokens.length);
        poolTokens = poolTokens.filter((_, i) => i !== missingIdx);
      }
      const extras = [];
      if (variant === "dummy") {
        const d = getDummyTokenForSort(q);
        if (d) extras.push(d);
      }
      const allCards = poolTokens.concat(extras);
      shuffleArrayInPlace(allCards);
      const primaryDisplay = getSortPrimaryCorrectDisplay(q);
      sortQuizState = {
        expectedCount,
        variant,
        missingIdx,
        primaryCorrect: primaryDisplay,
        sortables: []
      };
      const trashLabel = variant === "dummy" ? "使わないカード（除外置き場）" : "使わないカード（整理用）";
      const showMissing = variant === "missing";
      const answerArea = document.getElementById("quiz-answer-area");
      answerArea.innerHTML =
        `<div id="sort-quiz-root" class="sort-quiz-root">` +
        (showMissing
          ? `<div class="sort-missing-toolbar"><p class="sort-missing-desc">たりないカードはキーボードで入力して「カードにする」で追加できます。</p><input type="text" id="sort-missing-input" class="large-input" readonly placeholder="キーボードで入力"><div id="keyboard-container"></div></div>`
          : "") +
        `<div class="sort-area-header"><span class="sort-area-label">回答欄（ここにカードを並べてください）</span></div>` +
        `<div id="sort-answer-slot" class="sortable-area sort-answer-slot"></div>` +
        `<div class="sort-area-header"><span class="sort-area-label">仮置き場（自由に並べ替え可能）</span>` +
        `<button type="button" id="sort-move-all-btn" class="sort-btn-small">↑ この順で回答欄にいれる</button></div>` +
        `<div id="sort-word-pool" class="sortable-area sort-word-pool"></div>` +
        `<div class="sort-area-header"><span class="sort-area-label">${escapeHtml(trashLabel)}</span></div>` +
        `<div id="sort-trash-slot" class="sortable-area sort-trash-slot"></div>` +
        `<button type="button" id="sort-submit-btn" class="submit-btn btn-green" style="display:none;width:100%;max-width:420px;margin:16px auto 0;">これで回答する</button></div>`;
      const pool = document.getElementById("sort-word-pool");
      allCards.forEach(w => {
        const card = document.createElement("div");
        card.className = "word-card";
        card.textContent = w;
        pool.appendChild(card);
      });
      if (showMissing) {
        shiftMode = 0; isShiftHoldMode = false;
        renderCustomKeyboard(false, true);
      }
      const opts = { group: "sort-shared", animation: 150, ghostClass: "sortable-ghost", onSort: () => checkSortQuizSubmit() };
      ["sort-answer-slot", "sort-word-pool", "sort-trash-slot"].forEach(id => {
        const el = document.getElementById(id);
        if (el && typeof Sortable !== "undefined") sortQuizState.sortables.push(Sortable.create(el, opts));
      });
      const moveBtn = document.getElementById("sort-move-all-btn");
      if (moveBtn) {
        moveBtn.onclick = () => {
          const wp = document.getElementById("sort-word-pool");
          const ans = document.getElementById("sort-answer-slot");
          if (!wp || !ans) return;
          Array.from(wp.children).forEach(c => ans.appendChild(c));
          checkSortQuizSubmit();
        };
      }
      const sub = document.getElementById("sort-submit-btn");
      if (sub) sub.onclick = () => submitSortQuizAnswer();
      placeQuizFeedbackAtTopOfAnswerArea(answerArea);
    }

    function computeQuizBasePoint(format, ansType, isWord) {
      const settingKey = `基本Pt_${format}_${ansType}`;
      const customPoint = Number(appSettings[settingKey]);
      const hasCustom = !isNaN(customPoint);

      if (ansType.startsWith("fill_")) {
        if (hasCustom) return customPoint;
        const numBlanks = parseInt(document.getElementById('setting-blank-count').value) || 1;
        return 5 + numBlanks;
      }
      if (format === "ja_to_en_sort") {
        const sk = `基本Pt_${format}_${ansType}`;
        const cp = Number(appSettings[sk]);
        if (!isNaN(cp)) return cp;
        if (ansType === "sort_all") return 25;
        if (ansType === "sort_dummy") return 28;
        if (ansType === "sort_missing") return 30;
        return 25;
      }
      if (hasCustom) return customPoint;
      if (isWord) {
        if (ansType === "4choice") return format.includes("to_en") ? 3 : 2;
        if (ansType === "typing" || ansType === "voice") return 20;
      } else {
        if (ansType === "4choice") return format.includes("to_en") ? 3 : 2;
        if (ansType === "typing" || ansType === "voice") return (format.includes("qtext") || format.includes("qaudio")) ? 30 : 25;
      }
      return 0;
    }

    /** GAS parseUnitSheetPointPercent_ と同じ（シート名末尾 _数字 ＝得点％） */
    function parseUnitSheetPointPercentClient(sheetName) {
      const s = String(sheetName || "");
      const m = s.match(/_(\d+)$/);
      if (!m) return 100;
      let p = parseInt(m[1], 10);
      if (isNaN(p)) return 100;
      if (p < 0) p = 0;
      if (p > 100) p = 100;
      return p;
    }

    /** handleSaveLearningSession と同じ係数（時間・ニガテ・ランダム） */
    function computePointsMultiplierClient() {
      const user = JSON.parse(localStorage.getItem('app_kid_user') || '{}');
      const unitId = getDetailedUnitId();
      let mult = 1.0;
      const lastStudyTimeStr = user.lastStudyJson && user.lastStudyJson[unitId];
      if (lastStudyTimeStr) {
        const lastTime = new Date(lastStudyTimeStr);
        const diffHours = (Date.now() - lastTime) / (1000 * 60 * 60);
        let basePercent = 10 + Math.floor(diffHours / 2) * 10;
        if (basePercent > 100) basePercent = 100;
        mult = basePercent / 100;
      }
      if (currentIsReviewMode) mult += 0.4;
      if (document.getElementById('setting-order') && document.getElementById('setting-order').value === 'random') mult += 0.1;
      return mult;
    }

    /** GAS と同じ二段階の丸め */
    function applySessionEarnedFromRaw(sessionRawPoints, multiplier, sheetPointPercent) {
      let earned = Math.round(sessionRawPoints * multiplier * 100) / 100;
      if (sheetPointPercent !== 100) {
        earned = Math.round(earned * (sheetPointPercent / 100) * 100) / 100;
      }
      return earned;
    }

    function rawPointsFromQuizResults(results) {
      let sum = 0;
      (results || []).forEach(res => {
        if (res.isCorrect) {
          const qpt = Math.max(1, (Number(res.basePoint) || 2) - (Number(res.maxDeduction) || 0));
          sum += qpt;
        }
      });
      return sum;
    }

    function sumMaxPossibleRawForSession() {
      if (!filteredQuestions || filteredQuestions.length === 0) return 0;
      const format = document.getElementById('setting-format').value;
      const ansType = document.getElementById('setting-answer-type').value;
      const isWord = currentModeName.includes("単語");
      const bp = computeQuizBasePoint(format, ansType, isWord);
      return filteredQuestions.length * bp;
    }

    function formatPointDisplayNum(v) {
      const n = Number(v);
      if (isNaN(n)) return '0';
      if (Math.abs(n - Math.round(n)) < 1e-9) return String(Math.round(n));
      return String(Math.round(n * 100) / 100);
    }

    function updateSessionScoreDisplay() {
      const el = document.getElementById('quiz-score-bar');
      if (!el || !filteredQuestions || filteredQuestions.length === 0) {
        if (el) el.innerHTML = "";
        return;
      }
      const mult = computePointsMultiplierClient();
      const sheetPct = parseUnitSheetPointPercentClient(currentUnitName);
      const maxRaw = sumMaxPossibleRawForSession();
      const maxEarned = applySessionEarnedFromRaw(maxRaw, mult, sheetPct);
      const curRaw = rawPointsFromQuizResults(quizResults);
      const curEarned = applySessionEarnedFromRaw(curRaw, mult, sheetPct);
      const parts = [];
      parts.push(`このセッションの得点 <strong style="color:#FFD54F;font-size:20px;">${formatPointDisplayNum(curEarned)}</strong> / <span style="color:#aaa;">${formatPointDisplayNum(maxEarned)}</span> 点`);
      parts.push(`<span style="font-size:12px;color:#888;">（全問せいかい・ヒントなしのときの最大。時間・ニガテ・ランダム・単元シート名の補正％を含む）</span>`);
      el.innerHTML = parts.join('<br>');
    }

    let recognition = null;
    let recognitionJa = null;
    if ('SpeechRecognition' in window || 'webkitSpeechRecognition' in window) {
      const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
      recognition = new SpeechRecognition();
      recognition.lang = 'en-US'; recognition.interimResults = true; recognition.continuous = false;
    }
    let preferredVoice = null;
    function loadVoices() {
      const voices = window.speechSynthesis.getVoices();
      if (voices.length === 0) return;
      const englishVoices = voices.filter(v => v.lang.startsWith('en'));
      if (englishVoices.length > 0) {
        const premiumKeywords = ['google us english', 'enhanced', 'premium', 'samantha', 'alex'];
        for (const keyword of premiumKeywords) { const match = englishVoices.find(v => v.name.toLowerCase().includes(keyword)); if (match) { preferredVoice = match; break; } }
        if (!preferredVoice) preferredVoice = englishVoices[0];
      }
    }
    if ('speechSynthesis' in window) { window.speechSynthesis.onvoiceschanged = loadVoices; loadVoices(); }
    function speakText(text, rate = 0.9) {
      if ('speechSynthesis' in window) { window.speechSynthesis.cancel(); const uttr = new SpeechSynthesisUtterance(text); uttr.lang = 'en-US'; uttr.rate = rate; if (preferredVoice) uttr.voice = preferredVoice; window.speechSynthesis.speak(uttr); }
    }

    function formatStopwatch(ms) {
      const totalSeconds = Math.floor(ms / 1000);
      const h = String(Math.floor(totalSeconds / 3600)).padStart(2, '0');
      const m = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');
      const s = String(totalSeconds % 60).padStart(2, '0');
      return `${h}:${m}:${s}`;
    }
    function renderStopwatch(target) {
      const el = document.getElementById(`${target}-sw-display`);
      if (!el) return;
      const st = stopwatches[target];
      const now = st.timerId ? Date.now() : st.startAt;
      const elapsed = st.timerId ? (now - st.startAt) : st.elapsed;
      el.innerText = formatStopwatch(elapsed);
    }
    function startStopwatch(target) {
      const st = stopwatches[target];
      if (!st || st.timerId) return;
      st.startAt = Date.now() - st.elapsed;
      st.timerId = setInterval(() => renderStopwatch(target), 200);
      renderStopwatch(target);
    }
    function stopStopwatch(target) {
      const st = stopwatches[target];
      if (!st) return;
      if (st.timerId) {
        clearInterval(st.timerId);
        st.timerId = null;
        st.elapsed = Date.now() - st.startAt;
      }
      renderStopwatch(target);
    }
    function resetStopwatch(target) {
      const st = stopwatches[target];
      if (!st) return;
      if (st.timerId) clearInterval(st.timerId);
      st.timerId = null; st.startAt = 0; st.elapsed = 0;
      renderStopwatch(target);
    }

    function getDetailedUnitId() { 
      // isTrainingMode が true なら特訓ルート専用のIDにする
      if (isTrainingMode) return `TrainingRoute_${currentTrainingStepIndex}_${currentUnitName}`;
      const format = document.getElementById('setting-format').value; 
      const ansType = document.getElementById('setting-answer-type').value; 
      let unitId = `${currentModeName}_${currentUnitName}_${format}_${ansType}`; 
      if (ansType.startsWith("fill_")) {
        const blankCount = document.getElementById('setting-blank-count').value;
        unitId += `_${blankCount}blanks`;
      }
      return unitId; 
    }

    function checkIsNigate(historyObj) { if(!historyObj || !historyObj.results || historyObj.results.length === 0) return false; const rMissLimit = Number(appSettings["苦手_直近ミス判定"]) || 1; const timeLimit = Number(appSettings["苦手_時間判定_秒"]) || 15; const rateLimit = Number(appSettings["苦手_全体正答率_未満"]) || 80; const recentResults = historyObj.results.slice(-3); if (recentResults.filter(r => r === 0).length >= rMissLimit) return true; if (((historyObj.results.filter(r => r === 1).length / historyObj.results.length) * 100) < rateLimit) return true; if (historyObj.times && historyObj.times.length > 0 && (historyObj.times.reduce((a,b)=>a+b, 0) / historyObj.times.length) >= timeLimit) return true; return false; }
    
    // ★ 特訓ルート用の変数
    let dailyRouteData = []; 
    let currentProgressData = {}; 
    let currentTrainingStepIndex = null;
    let currentTrainingMenuId = 1;
    let isTrainingMode = false;

    window.onload = () => { 
        const saved = localStorage.getItem('app_kid_user'); 
        if (inputMethodMode !== 'pen' && inputMethodMode !== 'keyboard') inputMethodMode = 'keyboard';
        initKeyboardAndSoundSettings();
        applyKanjiHwDominantHandToBody();
        syncKanjiHwHandSwitchUI();
        if (saved) { showHome(JSON.parse(saved)); fetchAppSettings(); }
        else fetchUsers(); 
    };

    function fetchAppSettings() { 
      const cacheKey = 'app_cached_settings';
      const cached = localStorage.getItem(cacheKey);
      if (cached) {
        try {
          const d = JSON.parse(cached);
          appSettings = d.settings;
          return Promise.resolve(d);
        } catch(e) {}
      }
      return fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_app_settings" }) }).then(r=>r.json()).then(d=>{ 
        if(d.status==="success") {
          appSettings = d.settings; 
          localStorage.setItem(cacheKey, JSON.stringify(d));
        }
        return d; 
      }); 
    }

    const LS_KBD_SCALE = 'vk_scale_pct';
    const LS_KBD_FONT = 'vk_key_font_px';
    const LS_ANSWER_SOUND = 'answer_sound_enabled';
    const LS_IPAD_STYLUS_OPT = 'ipad_stylus_opt_enabled';
    const LS_PEN_ADVANCED_VISIBLE = 'pen_advanced_visible';

    function initKeyboardAndSoundSettings() {
      if (getUserPref(LS_KBD_SCALE, null) === null) setUserPref(LS_KBD_SCALE, '100');
      if (getUserPref(LS_KBD_FONT, null) === null) setUserPref(LS_KBD_FONT, '18');
      if (getUserPref(LS_ANSWER_SOUND, null) === null) setUserPref(LS_ANSWER_SOUND, '1');
      if (getUserPref(LS_IPAD_STYLUS_OPT, null) === null) setUserPref(LS_IPAD_STYLUS_OPT, '0');
      if (getUserPref(LS_PEN_ADVANCED_VISIBLE, null) === null) setUserPref(LS_PEN_ADVANCED_VISIBLE, '0');
      const sc = document.getElementById('kbd-scale-range');
      const ft = document.getElementById('kbd-font-range');
      if (sc) sc.value = getUserPref(LS_KBD_SCALE, '100');
      if (ft) ft.value = getUserPref(LS_KBD_FONT, '18');
      const qsc = document.getElementById('quiz-kbd-scale-range');
      const qft = document.getElementById('quiz-kbd-font-range');
      if (qsc) qsc.value = getUserPref(LS_KBD_SCALE, '100');
      if (qft) qft.value = getUserPref(LS_KBD_FONT, '18');
      const on = getUserPref(LS_ANSWER_SOUND, '1') !== '0';
      const snd = document.getElementById('answer-sound-enabled');
      const qsnd = document.getElementById('quiz-answer-sound-enabled');
      const ipadStylus = document.getElementById('ipad-stylus-opt-enabled');
      if (snd) snd.checked = on;
      if (qsnd) qsnd.checked = on;
      if (ipadStylus) ipadStylus.checked = getUserPref(LS_IPAD_STYLUS_OPT, '0') === '1';
      applyKeyboardSettingsFromControls('settings');
    }

    function syncAnswerSoundFromCheckbox(el) {
      if (!el) return;
      setUserPref(LS_ANSWER_SOUND, el.checked ? '1' : '0');
      const a = document.getElementById('answer-sound-enabled');
      const q = document.getElementById('quiz-answer-sound-enabled');
      if (a && el !== a) a.checked = el.checked;
      if (q && el !== q) q.checked = el.checked;
    }

    function syncIpadStylusSettingsFromCheckbox(el) {
      if (!el) return;
      const checked = !!el.checked;
      setUserPref(LS_IPAD_STYLUS_OPT, checked ? '1' : '0');
      const panelCheckbox = document.getElementById('pen-ipad-stylus-opt');
      if (panelCheckbox && panelCheckbox !== el) panelCheckbox.checked = checked;
      refreshPenStatusHint();
    }

    function isIpadStylusOptimizationEnabled() {
      return getUserPref(LS_IPAD_STYLUS_OPT, '0') === '1';
    }

    function isPenAdvancedVisible() {
      return getUserPref(LS_PEN_ADVANCED_VISIBLE, '0') === '1';
    }

    function syncPenAdvancedVisibility() {
      const shown = isPenAdvancedVisible();
      const panel = document.getElementById('pen-advanced-controls');
      const btn = document.getElementById('pen-advanced-toggle-btn');
      if (panel) panel.style.display = shown ? 'flex' : 'none';
      if (btn) btn.innerText = shown ? '詳細設定ボタンを非表示' : '詳細設定ボタンを表示';
    }

    function togglePenAdvancedSettings() {
      const next = isPenAdvancedVisible() ? '0' : '1';
      setUserPref(LS_PEN_ADVANCED_VISIBLE, next);
      syncPenAdvancedVisibility();
    }

    function applyKeyboardSettingsFromControls(from) {
      let scalePct = parseInt(getUserPref(LS_KBD_SCALE, '100'), 10);
      let fontPx = parseInt(getUserPref(LS_KBD_FONT, '18'), 10);
      if (from === 'settings' || !from) {
        const sc = document.getElementById('kbd-scale-range');
        const ft = document.getElementById('kbd-font-range');
        if (sc) scalePct = parseInt(sc.value, 10);
        if (ft) fontPx = parseInt(ft.value, 10);
      }
      if (from === 'quiz') {
        const sc = document.getElementById('quiz-kbd-scale-range');
        const ft = document.getElementById('quiz-kbd-font-range');
        if (sc) scalePct = parseInt(sc.value, 10);
        if (ft) fontPx = parseInt(ft.value, 10);
      }
      if (isNaN(scalePct)) scalePct = 100;
      if (isNaN(fontPx)) fontPx = 18;
      setUserPref(LS_KBD_SCALE, String(scalePct));
      setUserPref(LS_KBD_FONT, String(fontPx));
      const scale = scalePct / 100;
      const padPx = Math.max(8, Math.round(fontPx * 0.72));
      document.documentElement.style.setProperty('--vk-font-px', String(fontPx));
      document.documentElement.style.setProperty('--vk-pad-px', String(padPx));
      const syncLabels = () => {
        const l1 = document.getElementById('kbd-scale-label');
        const l2 = document.getElementById('kbd-font-label');
        const q1 = document.getElementById('quiz-kbd-scale-label');
        const q2 = document.getElementById('quiz-kbd-font-label');
        if (l1) l1.innerText = String(scalePct);
        if (l2) l2.innerText = String(fontPx);
        if (q1) q1.innerText = String(scalePct);
        if (q2) q2.innerText = String(fontPx);
      };
      const syncRanges = () => {
        const ids = ['kbd-scale-range', 'kbd-font-range', 'quiz-kbd-scale-range', 'quiz-kbd-font-range'];
        ids.forEach(id => { const el = document.getElementById(id); if (!el) return; if (id.indexOf('scale') >= 0) el.value = scalePct; else el.value = fontPx; });
      };
      syncRanges();
      syncLabels();
      document.querySelectorAll('.keyboard-scale-wrap').forEach(wrap => {
        wrap.style.setProperty('--kb-scale', String(scale));
        wrap.style.setProperty('--vk-font-px', String(fontPx));
        wrap.style.setProperty('--vk-pad-px', String(padPx));
      });
      const quizSec = document.getElementById('section-quiz');
      if (quizSec && quizSec.classList.contains('active')) {
        const at = document.getElementById('setting-answer-type') && document.getElementById('setting-answer-type').value;
        const fmt = document.getElementById('setting-format') && document.getElementById('setting-format').value;
        if (at === 'typing') renderCustomKeyboard(false);
        else if (at === 'fill_typing') renderCustomKeyboard(true);
        else if (fmt === 'ja_to_en_sort' && at === 'sort_missing') renderCustomKeyboard(false, true);
      }
    }

    function toggleQuizKeyboardPanel() {
      const body = document.getElementById('quiz-keyboard-settings-body');
      if (!body) return;
      body.style.display = body.style.display === 'none' ? 'block' : 'none';
    }

    let answerAudioCtx = null;
    function playAnswerSound(isCorrect) {
      if (getUserPref(LS_ANSWER_SOUND, '1') === '0') return;
      try {
        const AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) return;
        if (!answerAudioCtx) answerAudioCtx = new AC();
        const ctx = answerAudioCtx;
        if (ctx.state === 'suspended') ctx.resume();
        const g = ctx.createGain();
        g.connect(ctx.destination);
        g.gain.value = 0.2;
        if (isCorrect) {
          [880, 1174].forEach((freq, i) => {
            const o = ctx.createOscillator();
            o.type = 'sine';
            o.frequency.value = freq;
            o.connect(g);
            const t0 = ctx.currentTime + i * 0.09;
            o.start(t0);
            o.stop(t0 + 0.1);
          });
        } else {
          const o = ctx.createOscillator();
          o.type = 'square';
          o.frequency.value = 180;
          o.connect(g);
          const t0 = ctx.currentTime;
          o.start(t0);
          o.stop(t0 + 0.22);
        }
      } catch (e) { /* ignore */ }
    }

    function escapeHtml(s) {
      if (s == null) return '';
      return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    /** materials シート名「単語A_40」→ 一覧・タイトル用「単語A（補正40％）」（API・保存は raw 名のまま） */
    function formatUnitSheetDisplayLabel(sheetName) {
      const s = String(sheetName || "");
      const m = s.match(/^(.*)_(\d+)$/);
      if (!m) return s;
      const pct = parseInt(m[2], 10);
      if (isNaN(pct) || pct < 0 || pct > 100) return s;
      return m[1] + "（補正" + pct + "％）";
    }
    function normalizeUnitNameForCompare(name) {
      const s = String(name || "").trim().replace(/\s+/g, "");
      return s.replace(/_(\d+)$/, "");
    }

    function buildFeedbackContentHtml(userA, resolvedCorrect, isCorrect, isSkip, maxDeduction, fillMode, answerType, pointsPlusHtml) {
      const correctDisp = String(resolvedCorrect != null ? resolvedCorrect : "");
      const isChoiceAnswer = (answerType === "4choice" || answerType === "fill_4choice");

      let userSlotHtml = "";
      if (isSkip) {
        userSlotHtml = `<div class="quiz-feedback-mirror-input quiz-feedback-mirror-skip">スキップしました</div>`;
      } else if (fillMode && fillBlanksData && fillBlanksData.length) {
        const word = String(resolvedCorrect != null ? resolvedCorrect : "");
        const ua = String(userA != null ? userA : "");
        let userLineHtml = "";
        for (let i = 0; i < word.length; i++) {
          const bd = fillBlanksData.find(b => b.originalIndex === i);
          const ch = ua[i] != null ? ua[i] : "";
          if (bd) {
            const ok = normalizeText(String(bd.correctChar)) === normalizeText(String(ch));
            userLineHtml += `<span class="${ok ? "fill-pos-correct" : "fill-pos-wrong"}">${escapeHtml(ch)}</span>`;
          } else {
            userLineHtml += escapeHtml(ch);
          }
        }
        userSlotHtml = `<div class="word-blank-container" style="margin: 8px auto 10px; text-align: center;">${userLineHtml}</div>`;
      } else if (answerType === "4choice") {
        userSlotHtml = `<div class="quiz-feedback-mirror-choice">${escapeHtml(userA)}</div>`;
      } else {
        userSlotHtml = `<div class="quiz-feedback-mirror-input">${escapeHtml(userA)}</div>`;
      }

      let judgeHtml = "";
      if (isSkip) judgeHtml = `<div class="quiz-feedback-judge" style="color:#FF9800;">⏭️</div>`;
      else if (isCorrect) judgeHtml = `<div class="quiz-feedback-judge" style="color:#4CAF50;">⭕</div>`;
      else judgeHtml = `<div class="quiz-feedback-judge" style="color:#F44336;">×</div>`;

      const correctBlock = `<div class="quiz-feedback-correct-label">せいかい</div><div class="quiz-feedback-correct-box">${escapeHtml(correctDisp)}</div>`;
      let correctBlockVariant = correctBlock;
      if (isChoiceAnswer) {
        correctBlockVariant = (!isCorrect || isSkip) ? correctBlock : "";
      }

      let hintHtml = "";
      if (maxDeduction > 0) hintHtml = `<span class="quiz-feedback-hint-penalty">ヒントをつかったので -${maxDeduction}Pt されます</span>`;

      const plus = pointsPlusHtml ? String(pointsPlusHtml) : "";
      if (isChoiceAnswer) {
        return `<div class="quiz-feedback-panel">
        <div class="quiz-feedback-your-label">あなたのこたえ</div>
        ${userSlotHtml}
        ${judgeHtml}
        ${correctBlockVariant}
        ${hintHtml}
        ${plus}
      </div>`;
      }
      return `<div class="quiz-feedback-panel">
        <div class="quiz-feedback-your-label">あなたのこたえ</div>
        ${userSlotHtml}
        ${correctBlock}
        ${judgeHtml}
        ${hintHtml}
        ${plus}
      </div>`;
    }
    
    const LS_APP_CACHED_MATERIALS = "app_cached_materials";

    function normalizeMaterialsSig(materials) {
      return JSON.stringify((Array.isArray(materials) ? materials : []).map(m => ({
        i: String(m.modeId || ""),
        n: String(m.modeName || ""),
        c: String(m.category || ""),
        u: (Array.isArray(m.units) ? m.units : []).map(x => String(x || "")).sort()
      })).sort((a, b) => a.i.localeCompare(b.i)));
    }

    function getCachedMaterialsFingerprint() {
      try {
        const raw = localStorage.getItem(LS_APP_CACHED_MATERIALS);
        if (!raw) return null;
        const d = JSON.parse(raw);
        return d.materialsFingerprint || null;
      } catch (e) {
        return null;
      }
    }

    function invalidateKanjiPracticeCatalog() {
      try {
        kpCatalogState.materials = [];
        kpCatalogState.sets = [];
        kpCatalogState.setQuestions = [];
        kpCatalogState.filteredChars = [];
        kpCatalogState.loaded = false;
      } catch (_) {}
    }

    function persistMaterialsPayload(d) {
      if (!d || d.status !== "success") return;
      const mats = Array.isArray(d.materials) ? d.materials : [];
      materialsData = mats;
      invalidateKanjiPracticeCatalog();
      const nextNormSig = normalizeMaterialsSig(mats);
      try {
        localStorage.setItem(LS_APP_CACHED_MATERIALS, JSON.stringify({
          status: "success",
          materials: materialsData,
          materialsFingerprint: d.materialsFingerprint || "",
          materialsNormalizedSig: nextNormSig
        }));
      } catch (e) {}
    }

    function fetchMaterialsListFromServer() {
      return fetch(GAS_API_URL, { method: "POST", body: JSON.stringify({ action: "get_materials_list" }) }).then(r => r.json());
    }

    /** サーバのフィンガープリント（または一覧の正規化署名）が変わったときだけ英語・漢字の教材一覧を差し替える。 */
    function applyMaterialsResponseIfChanged(d) {
      if (!d || d.status !== "success") return false;
      const prevFp = getCachedMaterialsFingerprint();
      let prevNormSig = "";
      try {
        const raw = localStorage.getItem(LS_APP_CACHED_MATERIALS);
        if (raw) {
          const c = JSON.parse(raw);
          prevNormSig = c.materialsNormalizedSig || normalizeMaterialsSig(c.materials);
        }
      } catch (e) {}
      const newFp = d.materialsFingerprint || "";
      const nextNormSig = normalizeMaterialsSig(d.materials);
      let changed = false;
      if (newFp) changed = newFp !== prevFp;
      else changed = nextNormSig !== prevNormSig;
      if (!changed && materialsData.length === 0) changed = true;
      if (!changed) return false;
      persistMaterialsPayload(d);
      return true;
    }

    function syncMaterialsListFromServerIfChanged() {
      const statusEl = document.getElementById("home-materials-sync-status");
      return fetchMaterialsListFromServer()
        .then(d => {
          if (applyMaterialsResponseIfChanged(d) && statusEl) {
            statusEl.textContent = "教材一覧を最新にしました（英語・漢字）。";
            setTimeout(() => { if (statusEl.textContent.indexOf("最新にしました") >= 0) statusEl.textContent = ""; }, 6000);
          }
        })
        .catch(() => {});
    }

    function refreshMaterialsManualFromHome(btn) {
      const orig = toggleBtnLoading(btn, true);
      const statusEl = document.getElementById("home-materials-sync-status");
      if (statusEl) statusEl.textContent = "";
      fetchMaterialsListFromServer()
        .then(d => {
          if (!d || d.status !== "success") throw new Error((d && d.message) || "取得に失敗しました");
          persistMaterialsPayload(d);
          if (statusEl) statusEl.textContent = "英語・漢字の教材一覧を更新しました。";
          toggleBtnLoading(btn, false, orig);
        })
        .catch(() => {
          if (statusEl) statusEl.textContent = "更新に失敗しました。通信を確認してください。";
          toggleBtnLoading(btn, false, orig);
        });
    }

    /** 教材一覧（英語・漢字共通）。メモリ → localStorage → GAS の順で取得。 */
    function ensureMaterialsListLoaded() {
      if (materialsData.length > 0) return Promise.resolve(materialsData);
      const cached = localStorage.getItem(LS_APP_CACHED_MATERIALS);
      if (cached) {
        try {
          const d = JSON.parse(cached);
          if (d.status === "success" && Array.isArray(d.materials)) {
            materialsData = d.materials;
            return Promise.resolve(materialsData);
          }
        } catch (e) {}
      }
      return fetchMaterialsListFromServer().then(d => {
        if (d.status === "success") persistMaterialsPayload(d);
        return materialsData;
      });
    }
    function prefetchMaterials() {
      return ensureMaterialsListLoaded().catch(() => {});
    }

    function switchSection(id) {
      abandonKanjiQuizPlayIfLeavingSection(id);
      document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
      document.getElementById(id).classList.add('active');
      document.getElementById('message').innerText="";
      const kanjiFlowIds = [
        "section-kanji-learning",
        "section-kanji-quiz-sets",
        "section-kanji-quiz-play",
        "section-kanji-nigate",
        "section-kanji-practice"
      ];
      let isKanjiFlow = kanjiFlowIds.includes(id) || ((id === 'section-materials' || id === 'section-settings') && currentMaterialsCategory === 'kanji');
      if (id === 'section-result' && document.body.classList.contains('kanji-study-mode')) {
        isKanjiFlow = true;
      }
      document.body.classList.toggle('kanji-study-mode', isKanjiFlow);
      if (id === 'section-kanji-quiz-sets') {
        syncKanjiQuizFormatSelectFromStorage();
      }
      setTimeout(() => syncKanjiVerticalSelects(isKanjiFlow), 10);
    }

    function syncKanjiVerticalSelects(enable) {
      if (!enable) {
        document.querySelectorAll('.cvs-wrapper').forEach(w => w.remove());
        document.querySelectorAll('select.cvs-hidden').forEach(s => {
          s.classList.remove('cvs-hidden');
          s.style.display = '';
          if (s._cvsObserver) {
            s._cvsObserver.disconnect();
            s._cvsObserver = null;
          }
        });
        return;
      }
      const targetIds = [
        'kanji-quiz-format-select', 'kp-book-select', 'kp-sheet-select', 'kp-set-select',
        'kn-book-select', 'kn-sheet-select', 'kn-set-select', 'kn-nigate-axis',
        'setting-play-mode', 'setting-format', 'setting-answer-type', 'setting-order', 'setting-blank-count'
      ];
      targetIds.forEach(id => {
        const select = document.getElementById(id);
        if (!select || select.classList.contains('cvs-hidden')) return;
        select.classList.add('cvs-hidden');
        select.style.display = 'none';
        const wrapper = document.createElement('div');
        wrapper.className = 'cvs-wrapper';
        wrapper.tabIndex = 0;
        const selectedDiv = document.createElement('div');
        selectedDiv.className = 'cvs-selected';
        const menuDiv = document.createElement('div');
        menuDiv.className = 'cvs-menu';
        const updateSelectedText = () => {
          const opt = select.options[select.selectedIndex];
          selectedDiv.innerHTML = (opt ? escapeHtml(opt.text) : "") + '<span class="cvs-arrow">▼</span>';
        };
        updateSelectedText();
        const buildMenu = () => {
          menuDiv.innerHTML = "";
          Array.from(select.options).forEach((opt, i) => {
            const item = document.createElement('div');
            item.className = 'cvs-item';
            if (select.selectedIndex === i) item.classList.add('selected');
            item.innerText = opt.text;
            item.onclick = (e) => {
              e.stopPropagation();
              select.selectedIndex = i;
              updateSelectedText();
              wrapper.classList.remove('open');
              select.dispatchEvent(new Event('change'));
              buildMenu();
            };
            menuDiv.appendChild(item);
          });
        };
        buildMenu();
        wrapper.appendChild(selectedDiv);
        wrapper.appendChild(menuDiv);
        wrapper.onclick = (e) => {
          e.stopPropagation();
          const wasOpen = wrapper.classList.contains('open');
          document.querySelectorAll('.cvs-wrapper').forEach(w => w.classList.remove('open'));
          if (!wasOpen) {
            buildMenu();
            wrapper.classList.add('open');
          }
        };
        select.parentNode.insertBefore(wrapper, select.nextSibling);
        select._cvsObserver = new MutationObserver(() => {
          updateSelectedText();
          buildMenu();
        });
        select._cvsObserver.observe(select, { childList: true, attributes: true, attributeFilter: ['value'] });
        select.addEventListener('change', () => {
          updateSelectedText();
          buildMenu();
        });
      });
    }
    document.addEventListener('click', () => {
      document.querySelectorAll('.cvs-wrapper.open').forEach(w => w.classList.remove('open'));
    });

    function getQuizRecoveryDraft() {
      try {
        const raw = localStorage.getItem(LS_QUIZ_RECOVERY_DRAFT);
        if (!raw) return null;
        return JSON.parse(raw);
      } catch (_) {
        return null;
      }
    }

    function isRecoveryDraftValidForUser(draft, userId) {
      if (!draft || !userId) return false;
      if (String(draft.userId) !== String(userId)) return false;
      if (!Array.isArray(draft.filteredQuestions) || draft.filteredQuestions.length === 0) return false;
      if (!Array.isArray(draft.currentQuestions) || draft.currentQuestions.length === 0) return false;
      if (!Array.isArray(draft.quizResults)) return false;
      if (typeof draft.currentQuestionIndex !== 'number') return false;
      if (draft.currentQuestionIndex < 0 || draft.currentQuestionIndex > draft.filteredQuestions.length) return false;
      return true;
    }

    function saveQuizRecoveryDraft(nextQuestionIndex) {
      try {
        const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
        if (!user || !user.id) return;
        if (!Array.isArray(filteredQuestions) || filteredQuestions.length === 0) return;
        const idx = typeof nextQuestionIndex === 'number' ? nextQuestionIndex : currentQuestionIndex;
        const draft = {
          version: 1,
          savedAt: Date.now(),
          userId: user.id,
          currentModeId,
          currentModeName,
          currentUnitName,
          currentQuestionIndex: idx,
          quizResults: Array.isArray(quizResults) ? quizResults : [],
          currentQuestions: Array.isArray(currentQuestions) ? currentQuestions : [],
          filteredQuestions: Array.isArray(filteredQuestions) ? filteredQuestions : [],
          currentIsReviewMode: !!currentIsReviewMode,
          isTrainingMode: !!isTrainingMode,
          currentTrainingStepIndex,
          currentTrainingMenuId,
          inputMethodMode,
          settings: {
            format: String((document.getElementById('setting-format') || {}).value || ''),
            answerType: String((document.getElementById('setting-answer-type') || {}).value || ''),
            order: String((document.getElementById('setting-order') || {}).value || 'normal'),
            playMode: String((document.getElementById('setting-play-mode') || {}).value || 'normal'),
            blankCount: String((document.getElementById('setting-blank-count') || {}).value || '')
          }
        };
        localStorage.setItem(LS_QUIZ_RECOVERY_DRAFT, JSON.stringify(draft));
        renderHomeResumePanel();
      } catch (_) {}
    }

    function clearQuizRecoveryDraft() {
      localStorage.removeItem(LS_QUIZ_RECOVERY_DRAFT);
      renderHomeResumePanel();
    }

    function discardQuizRecoveryDraft() {
      if (!confirm("保存された途中の学習データを消します。よろしいですか？")) return;
      clearQuizRecoveryDraft();
      alert("復帰データを消しました。");
    }

    function renderHomeResumePanel() {
      const panel = document.getElementById('home-quiz-resume-panel');
      const textEl = document.getElementById('home-quiz-resume-text');
      if (!panel || !textEl) return;
      const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
      const draft = getQuizRecoveryDraft();
      if (!user || !isRecoveryDraftValidForUser(draft, user.id)) {
        panel.style.display = 'none';
        textEl.innerText = "";
        return;
      }
      const total = draft.filteredQuestions.length;
      const done = Math.max(0, Math.min(total, draft.currentQuestionIndex));
      const savedAt = new Date(draft.savedAt || Date.now());
      const stamp = isNaN(savedAt.getTime()) ? "さきほど" : savedAt.toLocaleString('ja-JP');
      textEl.innerText = `${formatUnitSheetDisplayLabel(draft.currentUnitName || "")} / ${done}問目まで保存（全${total}問）。最終保存: ${stamp}`;
      panel.style.display = 'block';
    }

    function resumeQuizFromDraft(opts = {}) {
      const askConfirm = opts && opts.askConfirm === false ? false : true;
      const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
      const draft = getQuizRecoveryDraft();
      if (!user || !isRecoveryDraftValidForUser(draft, user.id)) {
        alert("復帰できる学習データがありません。");
        renderHomeResumePanel();
        return false;
      }
      if (askConfirm && !confirm("途中の学習がありました。再開しますか？")) return false;

      currentModeId = draft.currentModeId || currentModeId;
      currentModeName = draft.currentModeName || currentModeName;
      currentUnitName = draft.currentUnitName || currentUnitName;
      currentQuestions = Array.isArray(draft.currentQuestions) ? draft.currentQuestions : [];
      filteredQuestions = Array.isArray(draft.filteredQuestions) ? draft.filteredQuestions : [];
      quizResults = Array.isArray(draft.quizResults) ? draft.quizResults : [];
      currentQuestionIndex = Math.max(0, Math.min(filteredQuestions.length, Number(draft.currentQuestionIndex) || 0));
      currentIsReviewMode = !!draft.currentIsReviewMode;
      isTrainingMode = !!draft.isTrainingMode;
      currentTrainingStepIndex = draft.currentTrainingStepIndex == null ? null : draft.currentTrainingStepIndex;
      currentTrainingMenuId = draft.currentTrainingMenuId || currentTrainingMenuId;
      if (draft.inputMethodMode === 'pen' || draft.inputMethodMode === 'keyboard') {
        inputMethodMode = draft.inputMethodMode;
      }

      const safeFormat = String((draft.settings && draft.settings.format) || 'ja_to_en');
      const safeAnsType = String((draft.settings && draft.settings.answerType) || 'typing');
      const safeOrder = String((draft.settings && draft.settings.order) || 'normal');
      const safePlayMode = String((draft.settings && draft.settings.playMode) || 'normal');
      const safeBlank = String((draft.settings && draft.settings.blankCount) || '1');

      const formatEl = document.getElementById('setting-format');
      const ansTypeEl = document.getElementById('setting-answer-type');
      const orderEl = document.getElementById('setting-order');
      const playEl = document.getElementById('setting-play-mode');
      const blankEl = document.getElementById('setting-blank-count');
      if (formatEl) formatEl.innerHTML = `<option value="${escapeHtml(safeFormat)}">${escapeHtml(safeFormat)}</option>`;
      if (ansTypeEl) ansTypeEl.innerHTML = `<option value="${escapeHtml(safeAnsType)}">${escapeHtml(safeAnsType)}</option>`;
      if (orderEl) orderEl.value = safeOrder;
      if (playEl) playEl.value = safePlayMode;
      if (blankEl) blankEl.value = safeBlank;

      switchSection('section-quiz');
      showQuestion();
      return true;
    }

    function promptQuizResumeIfNeeded(user) {
      if (!user || !user.id) return;
      const draft = getQuizRecoveryDraft();
      if (!isRecoveryDraftValidForUser(draft, user.id)) return;
      const token = `${user.id}_${draft.savedAt || 0}_${draft.currentQuestionIndex || 0}`;
      if (resumePromptShownToken === token) return;
      resumePromptShownToken = token;
      if (confirm("途中の学習がありました。再開しますか？")) {
        resumeQuizFromDraft({ askConfirm: false });
      }
    }

    function getKanjiQuizRecoveryDraft() {
      try {
        const raw = localStorage.getItem(LS_KANJI_QUIZ_RECOVERY_DRAFT);
        if (!raw) return null;
        return JSON.parse(raw);
      } catch (_) {
        return null;
      }
    }

    function isKanjiRecoveryDraftValidForUser(draft, userId) {
      if (!draft || !userId) return false;
      if (String(draft.userId) !== String(userId)) return false;
      if (!draft.session || typeof draft.session !== "object") return false;
      const s = draft.session;
      if (!Array.isArray(s.questions) || s.questions.length === 0) return false;
      if (!Number.isFinite(Number(s.index))) return false;
      const idx = Number(s.index);
      if (idx < 0 || idx >= s.questions.length) return false;
      return true;
    }

    function buildKanjiSessionDraftPayload() {
      if (!kanjiQuizSession) return null;
      return {
        modeId: kanjiQuizSession.modeId,
        modeName: kanjiQuizSession.modeName,
        unitName: kanjiQuizSession.unitName,
        setId: kanjiQuizSession.setId,
        isTrainingMode: !!kanjiQuizSession.isTrainingMode,
        trainingStepIndex: kanjiQuizSession.trainingStepIndex,
        trainingMenuId: kanjiQuizSession.trainingMenuId,
        questions: Array.isArray(kanjiQuizSession.questions) ? kanjiQuizSession.questions.slice() : [],
        index: Number(kanjiQuizSession.index || 0),
        totalEarned: Number(kanjiQuizSession.totalEarned || 0),
        newTotalPoints: kanjiQuizSession.newTotalPoints == null ? null : Number(kanjiQuizSession.newTotalPoints),
        logs: Array.isArray(kanjiQuizSession.logs) ? kanjiQuizSession.logs.slice() : [],
        nigateTraining: !!kanjiQuizSession.nigateTraining,
        nigateAxis: kanjiQuizSession.nigateAxis || null,
        nigateFeedback: kanjiQuizSession.nigateFeedback ? JSON.parse(JSON.stringify(kanjiQuizSession.nigateFeedback)) : null
      };
    }

    function saveKanjiQuizRecoveryDraft() {
      try {
        const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
        if (!user || !user.id || !kanjiQuizSession) return;
        const draft = {
          version: 1,
          savedAt: Date.now(),
          userId: user.id,
          session: buildKanjiSessionDraftPayload(),
          lastContext: lastKanjiQuizContext ? JSON.parse(JSON.stringify(lastKanjiQuizContext)) : null
        };
        localStorage.setItem(LS_KANJI_QUIZ_RECOVERY_DRAFT, JSON.stringify(draft));
        renderKanjiResumePanel();
      } catch (_) {}
    }

    function clearKanjiQuizRecoveryDraft() {
      localStorage.removeItem(LS_KANJI_QUIZ_RECOVERY_DRAFT);
      renderKanjiResumePanel();
    }

    function discardKanjiQuizRecoveryDraft() {
      if (!confirm("保存された途中の漢字学習データを消します。よろしいですか？")) return;
      clearKanjiQuizRecoveryDraft();
      alert("漢字の復帰データを消しました。");
    }

    function renderKanjiResumePanel() {
      const panel = document.getElementById('kanji-quiz-resume-panel');
      const textEl = document.getElementById('kanji-quiz-resume-text');
      if (!panel || !textEl) return;
      const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
      const draft = getKanjiQuizRecoveryDraft();
      if (!user || !isKanjiRecoveryDraftValidForUser(draft, user.id)) {
        panel.style.display = 'none';
        textEl.innerText = "";
        return;
      }
      const s = draft.session || {};
      const total = Array.isArray(s.questions) ? s.questions.length : 0;
      const done = Math.max(0, Math.min(total, Number(s.index) || 0));
      const savedAt = new Date(draft.savedAt || Date.now());
      const stamp = isNaN(savedAt.getTime()) ? "さきほど" : savedAt.toLocaleString('ja-JP');
      textEl.innerText = `${String(s.modeName || "")} / ${formatUnitSheetDisplayLabel(String(s.unitName || ""))} / セット${String(s.setId || "")} を ${done}問目まで保存（全${total}問）。最終保存: ${stamp}`;
      panel.style.display = 'block';
    }

    function resumeKanjiQuizFromDraft(opts = {}) {
      const askConfirm = opts && opts.askConfirm === false ? false : true;
      const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
      const draft = getKanjiQuizRecoveryDraft();
      if (!user || !isKanjiRecoveryDraftValidForUser(draft, user.id)) {
        alert("復帰できる漢字学習データがありません。");
        renderKanjiResumePanel();
        return false;
      }
      if (askConfirm && !confirm("途中の漢字学習がありました。再開しますか？")) return false;
      const s = draft.session;
      kanjiQuizSession = {
        modeId: s.modeId,
        modeName: s.modeName,
        unitName: s.unitName,
        setId: s.setId,
        isTrainingMode: !!s.isTrainingMode,
        trainingStepIndex: s.trainingStepIndex,
        trainingMenuId: s.trainingMenuId,
        questions: Array.isArray(s.questions) ? s.questions.slice() : [],
        index: Math.max(0, Math.min((Array.isArray(s.questions) ? s.questions.length : 1) - 1, Number(s.index) || 0)),
        totalEarned: Number(s.totalEarned || 0),
        newTotalPoints: s.newTotalPoints == null ? null : Number(s.newTotalPoints),
        logs: Array.isArray(s.logs) ? s.logs.slice() : [],
        selectedChoice: null,
        nigateTraining: !!s.nigateTraining,
        nigateAxis: s.nigateAxis || null,
        nigateFeedback: s.nigateFeedback ? JSON.parse(JSON.stringify(s.nigateFeedback)) : (s.nigateTraining ? { strokeOrderClean: true, brushAllClear: true } : null)
      };
      lastKanjiQuizContext = draft.lastContext && typeof draft.lastContext === "object"
        ? JSON.parse(JSON.stringify(draft.lastContext))
        : {
            modeId: s.modeId,
            modeName: s.modeName,
            unitName: s.unitName,
            setId: s.setId,
            questions: Array.isArray(s.questions) ? s.questions.slice() : [],
            allQuestions: Array.isArray(s.questions) ? s.questions.slice() : [],
            isTrainingMode: !!s.isTrainingMode,
            trainingStepIndex: s.trainingStepIndex,
            trainingMenuId: s.trainingMenuId,
            nigateTraining: !!s.nigateTraining,
            nigateAxis: s.nigateAxis || null,
            formatMode: getKanjiQuizFormatMode()
          };
      switchSection("section-kanji-quiz-play");
      renderKanjiQuizQuestion();
      return true;
    }

    function promptKanjiQuizResumeIfNeeded(user) {
      if (!user || !user.id) return;
      const draft = getKanjiQuizRecoveryDraft();
      if (!isKanjiRecoveryDraftValidForUser(draft, user.id)) return;
      const token = `${user.id}_${draft.savedAt || 0}_${(draft.session || {}).index || 0}`;
      if (kanjiResumePromptShownToken === token) return;
      kanjiResumePromptShownToken = token;
      if (confirm("途中の漢字学習がありました。再開しますか？")) {
        resumeKanjiQuizFromDraft({ askConfirm: false });
      }
    }

    function fetchUsers() { fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_child_users" }) }).then(r=>r.json()).then(d=>{ if(d.status==="success") renderUsers(d.users); }); }
    function renderUsers(users) { const c = document.getElementById('user-container'); c.innerHTML = ""; users.forEach(u => { const div = document.createElement('div'); div.className = 'user-card'; div.onclick = () => showPinScreen(u.id, u.name, false); div.innerHTML = `<div class="user-icon">${u.name.charAt(0)}</div><div>${u.name}</div>`; c.appendChild(div); }); }
    function showPinScreen(id, name, isReset) { selectedUserId = id; isPinResetMode = isReset; currentPin = ""; updatePinDisplay(); document.getElementById('selected-user-name').innerText = name; document.getElementById('pin-instruction').innerText = isReset ? "あたらしい暗証番号（4ケタ）" : "あんしょうばんごう（4ケタ）"; switchSection('section-pin'); }
    function addPin(num) { if(currentPin.length<4){ currentPin+=num; updatePinDisplay(); if(currentPin.length===4) { if(isPinResetMode) executePinReset(); else verifyPin(); } } }
    function deletePin() { currentPin = currentPin.slice(0,-1); updatePinDisplay(); }
    function clearPin() { currentPin = ""; updatePinDisplay(); }
    function updatePinDisplay() { let d=""; for(let i=0;i<4;i++) d+=(i<currentPin.length)?"● ":"_ "; document.getElementById('pin-dots').innerText=d.trim(); }
    function cancelPin() { if(isPinResetMode) showHome(JSON.parse(localStorage.getItem('app_kid_user'))); else switchSection('section-users'); }
    function verifyPin() { document.getElementById('message').innerText = "かくにん中..."; fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "verify_kid_pin", userId: selectedUserId, pin: currentPin }) }).then(r=>r.json()).then(d=>{ if(d.status==="success"){ localStorage.setItem('app_kid_user', JSON.stringify(d.user)); showHome(d.user); } else { document.getElementById('message').innerText = d.message; currentPin=""; updatePinDisplay(); } }); }
    function preparePinReset() { const user = JSON.parse(localStorage.getItem('app_kid_user')); showPinScreen(user.id, user.name, true); }
    function executePinReset() { document.getElementById('message').innerText = "へんこう中..."; fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "change_pin", userId: selectedUserId, newPin: currentPin }) }).then(r=>r.json()).then(d=>{ if(d.status==="success") { alert(d.message); showHome(JSON.parse(localStorage.getItem('app_kid_user'))); } else { document.getElementById('message').innerText = d.message; currentPin=""; updatePinDisplay(); } }); }
    function logout() { localStorage.removeItem('app_kid_user'); document.body.classList.remove('kanji-study-mode'); document.getElementById('global-header').style.display = 'none'; switchSection('section-users'); }

    // ★ 特訓ルートの読み込み処理を追加
    function showHome(user) {
      document.body.classList.remove('kanji-study-mode');
      switchSection('section-home');
      document.getElementById('header-user-info').innerText = `ID: ${user.id} / ${user.name}`; document.getElementById('global-header').style.display = 'block';
      document.getElementById('welcome-message').innerText = `${user.name} さん`; document.getElementById('user-points').innerText = user.points;
      renderHomeResumePanel();
      promptQuizResumeIfNeeded(user);
      applyKanjiHwDominantHandToBody();
      syncKanjiHwHandSwitchUI();
      syncMaterialsListFromServerIfChanged();
    }

    function toggleBtnLoading(btn, isLoading, originalText = "") { if(!btn) return ""; if(isLoading) { const text = btn.innerHTML; btn.innerHTML = "⏳ よみこみ中..."; btn.classList.add("btn-loading"); return text; } else { btn.innerHTML = originalText; btn.classList.remove("btn-loading"); } }

    // ====== ★ 特訓ルートの取得と表示 ======
    function fetchTrainingRoute(userId) {
        const homeRouteArea = document.getElementById('home-route-area');
        if (homeRouteArea) homeRouteArea.style.display = 'block';
        const routeContainer = document.getElementById('route-container');
        if (routeContainer) routeContainer.innerHTML = "<p>ルートを確認中...</p>";
        
        const cacheKey = `app_cached_training_route_${currentTrainingMenuId}`;
        const cached = localStorage.getItem(cacheKey);

        const processData = (d) => {
            if(d.status === "success") {
                dailyRouteData = d.route;
                currentProgressData = d.progress;
                renderTrainingRoute();
            } else if (routeContainer) {
                routeContainer.innerHTML = "<p>ルートの取得に失敗しました。</p>";
            }
        };

        if (cached) {
           try {
               const d = JSON.parse(cached);
               processData(d);
               return;
           } catch(e) {}
        }

        fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_training_route", userId: userId, trainingMenuId: currentTrainingMenuId }) })
        .then(r=>r.json()).then(d=>{
            if(d.status === "success") {
               localStorage.setItem(cacheKey, JSON.stringify(d));
            }
            processData(d);
        }).catch(e => {
            if (routeContainer) routeContainer.innerHTML = "<p>通信エラーが発生しました。</p>";
        });
    }

    function renderTrainingRoute() {
        const container = document.getElementById('route-container');
        container.innerHTML = "";
        
        if (dailyRouteData.length === 0) {
            container.innerHTML = "<p style='text-align:center;'>今日のミッションはないみたい！<br>いつもの学習をがんばろう。</p>";
            document.getElementById('next-route-btn').style.display = 'none';
            return;
        }

        let isLocked = false;
        let nextAvailableFound = false;

        dailyRouteData.forEach((route, index) => {
            const isCleared = currentProgressData[route.stepIndex] === true;
            
            const div = document.createElement('div');
            div.className = `route-item ${isCleared ? 'cleared' : ''} ${isLocked ? 'locked' : ''}`;
            
            const checkboxHtml = `<div class="checkbox">${isCleared ? '✔' : ''}</div>`;
            const infoHtml = `<div style="flex-grow:1; margin-left:10px;">
                                <div style="font-weight:bold;">${escapeHtml(formatUnitSheetDisplayLabel(route.unitName))}</div>
                                <div style="font-size:12px; color:#aaa;">${escapeHtml(route.qFormat)} / ${escapeHtml(route.aFormat)}</div>
                              </div>`;
            
            div.innerHTML = checkboxHtml + infoHtml;

            // クリックで直接挑戦（復習 or アンロック済みなら可能）
            if (isCleared || (!isCleared && !isLocked)) {
                div.style.cursor = "pointer";
                div.onclick = () => startRouteStep(route);
            }

            container.appendChild(div);

            // 未クリアを見つけたら、それ以降はロックする
            if (!isCleared) {
                isLocked = true;
                if (!nextAvailableFound) {
                    nextAvailableFound = true;
                }
            }
        });

        // 「次の内容に取り組む」ボタンの制御
        const nextBtn = document.getElementById('next-route-btn');
        if (!nextAvailableFound) {
            nextBtn.style.display = 'none'; // 全てクリア済み
        } else {
            nextBtn.style.display = 'block';
        }
    }

    // 「次の内容に取り組む」ボタンを押したとき
    function startNextRoute() {
        // 未クリアで一番上にあるものを探す
        const nextRoute = dailyRouteData.find(r => !currentProgressData[r.stepIndex]);
        if (nextRoute) {
            startRouteStep(nextRoute);
        }
    }

    function openTrainingMenu() {
      const user = JSON.parse(localStorage.getItem('app_kid_user'));
      if (!user) { alert("ログインし直してください"); return; }
      switchSection('section-training');
      document.getElementById('training-menu-picker').style.display = 'block';
      document.getElementById('route-container').style.display = 'none';
      document.getElementById('next-route-btn').style.display = 'none';
      renderTrainingMenuPicker();
    }

    function backToTrainingMenuPicker() {
      document.getElementById('training-menu-picker').style.display = 'block';
      document.getElementById('route-container').style.display = 'none';
      document.getElementById('next-route-btn').style.display = 'none';
      renderTrainingMenuPicker();
    }

    function renderTrainingMenuPicker() {
      const box = document.getElementById('training-menu-picker');
      if (!box) return;
      box.innerHTML = "<p style='text-align:center;color:#888;'>よみこみ中...</p>";
      fetchAppSettings().then(d => {
        const settings = (d && d.status === 'success' && d.settings) ? d.settings : {};
        let html = "<div class='training-menu-picker-inner'>";
        for (let m = 1; m <= 12; m++) {
          const label = String(settings['特訓メニュー' + m + '_表示名'] || '').trim() || ('特訓メニュー' + m);
          html += `<button type="button" class="submit-btn btn-purple" style="min-height:48px;padding:10px;font-size:15px;" onclick="selectTrainingMenu(${m})">${escapeHtml(label)}</button>`;
        }
        html += "</div>";
        box.innerHTML = html;
      }).catch(() => { box.innerHTML = "<p style='color:#f88;'>読み込みに失敗しました</p>"; });
    }

    function selectTrainingMenu(menuId) {
      currentTrainingMenuId = menuId;
      const user = JSON.parse(localStorage.getItem('app_kid_user'));
      document.getElementById('training-menu-picker').style.display = 'none';
      document.getElementById('route-container').style.display = 'block';
      document.getElementById('route-container').innerHTML = "<p>ルートを確認中...</p>";
      fetchTrainingRoute(user.id);
    }

    // 指定されたルートステップを開始する準備
    function startRouteStep(route) {
        document.getElementById('route-container').innerHTML = "<p>問題データを準備中...</p>";
        isTrainingMode = true;
        currentTrainingStepIndex = route.stepIndex;

        let mappedFormat = "en_to_ja";
        let mappedAnsType = "4choice";
        let mappedOrder = route.mode ? (route.mode === "ランダム" ? "random" : "normal") : "random";
        let mappedBlank = route.blankCount || "";
        
        if (route.qFormat === "日本語→英単語") mappedFormat = "ja_to_en";
        else if (route.qFormat === "英単語→日本語") mappedFormat = "en_to_ja";
        else if (route.qFormat === "音声→日本語") mappedFormat = "en_audio_to_ja";
        else if (route.qFormat === "音声→英単語") mappedFormat = "qaudio_to_en";
        else if (route.qFormat === "英語→英語" || route.qFormat === "英単語→英単語") mappedFormat = "en_to_en";
        
        if (route.aFormat === "タイピング") mappedAnsType = "typing";
        else if (route.aFormat === "音声") mappedAnsType = "voice";
        else if (route.aFormat === "穴埋め4択") mappedAnsType = "fill_4choice";
        else if (route.aFormat === "穴埋めタイピング") mappedAnsType = "fill_typing";

        prefetchMaterials().then(() => {
            let foundModeId = null;
            let foundModeName = "";
            let resolvedUnitName = "";
            const isKanjiRoute = /漢字/.test(String(route.qFormat || "")) || /採点/.test(String(route.aFormat || ""));
            const routeUnitRaw = String(route.unitName || "");
            const routeUnitNorm = normalizeUnitNameForCompare(routeUnitRaw);
            for (let m of materialsData) {
                const units = Array.isArray(m.units) ? m.units : [];
                const exact = units.find(u => String(u) === routeUnitRaw);
                if (exact) {
                  foundModeId = m.modeId;
                  foundModeName = m.modeName;
                  resolvedUnitName = String(exact);
                  break;
                }
                const fuzzy = units.find(u => normalizeUnitNameForCompare(u) === routeUnitNorm);
                if (fuzzy) {
                  foundModeId = m.modeId;
                  foundModeName = m.modeName;
                  resolvedUnitName = String(fuzzy);
                  break;
                }
            }
            if(!foundModeId) {
                alert(`「${formatUnitSheetDisplayLabel(route.unitName)}」のデータが見つかりませんでした。`);
                fetchTrainingRoute(JSON.parse(localStorage.getItem('app_kid_user')).id);
                return;
            }
            currentUnitName = resolvedUnitName || route.unitName;
            currentModeName = foundModeName;

            if (isKanjiRoute) {
              fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_kanji_quiz_sets", modeId: foundModeId, unitName: currentUnitName }) })
              .then(r => r.json()).then(d => {
                if (d.status !== "success") throw new Error(d.message || "漢字セット取得失敗");
                const sets = Array.isArray(d.sets) ? d.sets : [];
                if (!sets.length) throw new Error("セットが見つかりません");
                const first = sets[0];
                return fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_kanji_quiz_questions", modeId: foundModeId, unitName: currentUnitName, setId: String(first.setId || "") }) });
              })
              .then(r => r.json()).then(q => {
                if (q.status !== "success") throw new Error(q.message || "漢字問題取得失敗");
                var raw = Array.isArray(q.questions) ? q.questions : [];
                var prep = prepareKanjiQuizQuestionsForPlay(raw);
                if (!prep) throw new Error("この形式の問題がありません");
                startKanjiQuizPlay({
                  modeId: foundModeId,
                  modeName: foundModeName,
                  unitName: currentUnitName,
                  setId: String((q.setId != null) ? q.setId : ""),
                  allQuestions: raw,
                  formatMode: prep.formatMode,
                  isTrainingMode: true,
                  trainingStepIndex: currentTrainingStepIndex,
                  trainingMenuId: currentTrainingMenuId
                });
              })
              .catch(e => {
                alert("特訓用の漢字問題取得に失敗しました: " + (e.message || e));
                fetchTrainingRoute(JSON.parse(localStorage.getItem('app_kid_user')).id);
              });
              return;
            }

            const cacheKey = `app_cached_questions_${foundModeId}_${currentUnitName}`;
            const cached = localStorage.getItem(cacheKey);

            const processData = (d) => {
                if(d.status==="success"){ 
                    currentQuestions = d.questions;
                    
                    document.getElementById('setting-format').innerHTML = `<option value="${mappedFormat}">${route.qFormat}</option>`;
                    document.getElementById('setting-answer-type').innerHTML = `<option value="${mappedAnsType}">${route.aFormat}</option>`;
                    document.getElementById('setting-order').value = mappedOrder;
                    document.getElementById('setting-play-mode').value = "normal";
                    if (mappedAnsType.startsWith("fill_") && mappedBlank) {
                      const blankSelect = document.getElementById('setting-blank-count');
                      if (blankSelect) {
                        const exists = Array.from(blankSelect.options).some(o => o.value === String(mappedBlank));
                        if (!exists) blankSelect.insertAdjacentHTML('beforeend', `<option value="${mappedBlank}">${mappedBlank} 文字 かくす</option>`);
                        blankSelect.value = String(mappedBlank);
                      }
                    }
                    
                    prepareQuiz();
                } else {
                    alert("取得失敗: " + d.message);
                    fetchTrainingRoute(JSON.parse(localStorage.getItem('app_kid_user')).id);
                }
            };

            if (cached) {
                try {
                    const d = JSON.parse(cached);
                    processData(d);
                    return;
                } catch(e) {}
            }

            fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_questions", modeId: foundModeId, unitName: currentUnitName }) })
            .then(r=>r.json()).then(d=>{ 
                if (d.status === "success") localStorage.setItem(cacheKey, JSON.stringify(d));
                processData(d);
            }).catch(e => {
                alert("通信エラーが発生しました。");
                fetchTrainingRoute(JSON.parse(localStorage.getItem('app_kid_user')).id);
            });
        });
    }


    // ====== 景品・もちものバグ修正（クリックイベントの復元） ======
    function loadRewards(btn) { 
        const origText = toggleBtnLoading(btn, true); 
        document.getElementById('rewards-container').innerHTML = "<p>よみこみ中...</p>"; 
        fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_rewards" }) })
        .then(r=>r.json()).then(d=>{ 
            toggleBtnLoading(btn, false, origText); 
            if(d.status==="success"){ 
                switchSection('section-rewards'); 
                const c = document.getElementById('rewards-container'); 
                c.innerHTML = ""; 
                d.rewards.forEach(r => { 
                    const div = document.createElement('div'); div.className = "item-card"; 
                    div.innerHTML = `<div class="item-title">🎁 ${r.name}</div><div style="color: gold; font-weight: bold; margin-bottom: 10px;">必要ポイント: ${r.points} Pt</div><div style="font-size: 14px; color: #aaa; margin-bottom: 15px;">${r.desc}</div><button class="submit-btn btn-orange" style="width: 100%; padding: 10px; font-size: 18px;" onclick="exchangeReward('${r.id}', '${r.name}', ${r.points}, this)">交換する</button>`; 
                    c.appendChild(div); 
                }); 
            } 
        }).catch(e => toggleBtnLoading(btn, false, origText)); 
    }

    function exchangeReward(rewardId, rewardName, points, btn) { 
        const user = JSON.parse(localStorage.getItem('app_kid_user')); 
        if(user.points < points) { alert("ポイントが足りないよ！もっと勉強してポイントをためよう！"); return; } 
        if(!confirm(`「${rewardName}」と交換しますか？\n（${points} Ptへります）`)) return; 
        const origText = toggleBtnLoading(btn, true); 
        fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "exchange_reward", userId: user.id, rewardId: rewardId }) })
        .then(r=>r.json()).then(d=>{ 
            toggleBtnLoading(btn, false, origText); 
            if(d.status==="success") { 
                alert(d.message); 
                user.points = d.newPoints; localStorage.setItem('app_kid_user', JSON.stringify(user)); 
                showHome(user); 
            } else { alert(d.message); } 
        }).catch(e => { alert("通信エラーが発生しました。"); toggleBtnLoading(btn, false, origText); }); 
    }

    function loadInventory(btn) { 
        const origText = toggleBtnLoading(btn, true); 
        document.getElementById('inventory-container').innerHTML = "<p>よみこみ中...</p>"; 
        const user = JSON.parse(localStorage.getItem('app_kid_user')); 
        fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_inventory", userId: user.id }) })
        .then(r=>r.json()).then(d=>{ 
            toggleBtnLoading(btn, false, origText); 
            if(d.status==="success") { 
                userInventoryData = d.inventory; 
                showAllInventory = false; 
                switchSection('section-inventory'); 
                renderInventory(); 
            } 
        }).catch(e => toggleBtnLoading(btn, false, origText)); 
    }

    function renderInventory() { 
        const c = document.getElementById('inventory-container'); 
        c.innerHTML = ""; 
        const btn = document.getElementById('toggle-inventory-btn'); 
        btn.innerText = showAllInventory ? "未消化だけ表示する" : "すべて表示する"; 
        let filtered = userInventoryData; 
        if(!showAllInventory) filtered = filtered.filter(i => i.status !== "使用済み"); 
        if(filtered.length === 0) { c.innerHTML = "<p>表示できるアイテムがありません。</p>"; return; } 
        filtered.forEach(item => { 
            const div = document.createElement('div'); div.className = "item-card"; 
            const isUsed = item.status === "使用済み"; 
            const btnHtml = isUsed ? `<button class="submit-btn btn-gray" style="width: 100%; padding: 10px; font-size: 18px;" disabled>使用済み</button>` : `<button class="submit-btn btn-green" style="width: 100%; padding: 10px; font-size: 18px;" onclick="consumeReward(${item.rowIdx}, this)">使う！</button>`; 
            div.innerHTML = `<div class="item-title">🎒 ${item.rewardName}</div><div style="font-size: 12px; color: #aaa; margin-bottom: 10px;">交換日: ${new Date(item.date).toLocaleString()}</div>${btnHtml}`; 
            c.appendChild(div); 
        }); 
    }

    function toggleInventoryView() { showAllInventory = !showAllInventory; renderInventory(); }

    function consumeReward(rowIdx, btn) { 
        if(!confirm("本当にこの景品を使いますか？\n※おうちの人に確認してもらってから押してね！")) return; 
        const origText = toggleBtnLoading(btn, true); 
        fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "consume_reward", rowIdx: rowIdx }) })
        .then(r=>r.json()).then(d=>{ 
            if(d.status==="success") { 
                alert(d.message); 
                userInventoryData.find(i => i.rowIdx === rowIdx).status = "使用済み"; 
                renderInventory(); 
            } 
        }).catch(e => toggleBtnLoading(btn, false, origText)); 
    }

    // ====== 第2弾：外部学習（申請 → 管理者PIN承認） ======
    function loadExternalLearning(btn) {
      const origText = toggleBtnLoading(btn, true);
      document.getElementById('external-container').innerHTML = "<p>さがしています...</p>";
      document.getElementById('external-my-requests').innerHTML = "";
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_external_learning" }) })
      .then(r=>r.json()).then(d=>{
        toggleBtnLoading(btn, false, origText);
        if(d.status === "success") {
          switchSection('section-external');
          externalMenus = d.list || [];
          renderExternalLearningForm();
          refreshMyExternalRequests();
          resetStopwatch('external');
        }
      }).catch(e => toggleBtnLoading(btn, false, origText));
    }

    function renderExternalLearningForm() {
      const c = document.getElementById('external-container'); 
      c.innerHTML = "";
      if (externalMenus.length === 0) {
        c.innerHTML = "<p>登録されているメニューがありません。</p>";
        return;
      }
      const categories = [...new Set(externalMenus.map(m => m.category || ""))].filter(v => v);
      if (categories.length === 0) { c.innerHTML = "<p>カテゴリが設定されていません。</p>"; return; }
      c.innerHTML = `
        <div class="setting-group">
          <h3>カテゴリ</h3>
          <select id="external-category" class="large-select"></select>
        </div>
        <div class="setting-group">
          <h3>分量</h3>
          <select id="external-volume" class="large-select"></select>
          <p style="margin:6px 0 0; color:gold;">獲得ポイント: <span id="external-points-display">0</span> Pt</p>
        </div>
        <div class="setting-group">
          <h3>こどもメモ（自由記述）</h3>
          <textarea id="external-child-memo" rows="3" style="width:100%; box-sizing:border-box; border-radius:10px; padding:10px; background:#222; color:#fff; border:1px solid #555;" placeholder="がんばった内容やメモを書こう（空欄でもOK）"></textarea>
          <button type="button" class="submit-btn btn-blue" style="margin-top:8px; width:100%;" onclick="startJaSpeechToField('external-child-memo')">🎙️ 日本語で入力（マイク）</button>
        </div>
        <button type="button" class="submit-btn btn-green" style="width:100%;" onclick="submitExternalLearningRequestSelected()">申請する</button>
      `;
      const catSelect = document.getElementById('external-category');
      categories.forEach(cat => { const opt = document.createElement('option'); opt.value = cat; opt.innerText = cat; catSelect.appendChild(opt); });
      catSelect.onchange = () => updateExternalVolumeOptions();
      updateExternalVolumeOptions();
    }

    function updateExternalVolumeOptions() {
      const cat = document.getElementById('external-category').value;
      const volSelect = document.getElementById('external-volume');
      volSelect.innerHTML = "";
      const vols = externalMenus.filter(m => m.category === cat);
      vols.forEach(v => { const opt = document.createElement('option'); opt.value = v.volume; opt.innerText = v.volume; opt.dataset.points = v.points; volSelect.appendChild(opt); });
      const pts = document.getElementById('external-points-display');
      const first = vols[0];
      pts.innerText = first ? first.points : 0;
      volSelect.onchange = () => {
        const selected = vols.find(v => v.volume === volSelect.value);
        pts.innerText = selected ? selected.points : 0;
      };
    }

    function getSelectedExternalMenu() {
      const cat = document.getElementById('external-category').value;
      const vol = document.getElementById('external-volume').value;
      const match = externalMenus.find(m => m.category === cat && m.volume === vol);
      const points = match ? match.points : 0;
      return { category: cat, volume: vol, points };
    }

    function refreshMyExternalRequests() {
      const user = JSON.parse(localStorage.getItem('app_kid_user'));
      if (!user) return;
      const box = document.getElementById('external-my-requests');
      box.innerHTML = "<p style='color:#888;'>あなたの申請を読み込み中...</p>";
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_my_external_learning_requests", userId: user.id }) })
      .then(r=>r.json()).then(d=>{
        if(d.status !== "success" || !d.list || d.list.length === 0) {
          box.innerHTML = "<p style='color:#888;'>※まだ申請履歴はありません。</p>";
          return;
        }
        const label = { "申請中": "⏳ 申請中", "承認済み": "✅ 承認済み", "却下": "❌ 却下" };
        let html = "<strong style='color:#fff;'>あなたの申請</strong><ul style='margin:8px 0 0;padding-left:18px;'>";
        d.list.forEach(r => {
          const st = label[r.status] || r.status;
          const memoHtml = r.childMemo ? `<br><span style='font-size:12px;color:#ddd;'>メモ: ${r.childMemo}</span>` : "";
          html += `<li style='margin-bottom:8px;'>${r.category} / ${r.volume}（+${r.points} Pt）<br><span style='font-size:12px;color:#aaa;'>${r.requestedAt} — ${st}</span>${memoHtml}</li>`;
        });
        html += "</ul>";
        box.innerHTML = html;
      }).catch(() => { box.innerHTML = ""; });
    }

    function submitExternalLearningRequestSelected() {
      const menu = getSelectedExternalMenu();
      if(!menu.category || !menu.volume) { alert("カテゴリと分量を選んでください。"); return; }
      if(!confirm(`「${menu.category} / ${menu.volume}」を申請する？\n（おうちの人が管理者PINで承認するまでポイントは付きません）`)) return;
      const user = JSON.parse(localStorage.getItem('app_kid_user'));
      const childMemo = document.getElementById('external-child-memo').value || "";
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "submit_external_learning_request", userId: user.id, category: menu.category, volume: menu.volume, points: menu.points, childMemo: childMemo }) })
      .then(r=>r.json()).then(d=>{
        if(d.status === "success") {
          alert(d.message);
          refreshMyExternalRequests();
        } else {
          alert(d.message || "申請に失敗しました");
        }
      }).catch(() => alert("通信エラーが発生しました。"));
    }

    function openExternalAdminApproval() {
      document.getElementById('external-admin-pin-input').value = "";
      document.getElementById('external-admin-message').innerText = "";
      document.getElementById('external-admin-list').style.display = "none";
      document.getElementById('external-admin-list').innerHTML = "";
      document.getElementById('external-admin-pin-panel').style.display = "block";
      switchSection('section-external-admin');
    }

    function getExternalAdminPin() {
      return (document.getElementById('external-admin-pin-input') && document.getElementById('external-admin-pin-input').value) || "";
    }

    function loadExternalAdminPending() {
      const pin = getExternalAdminPin();
      const msg = document.getElementById('external-admin-message');
      const listEl = document.getElementById('external-admin-list');
      msg.innerText = "";
      if (!pin) {
        msg.innerText = "PINを入力してください。";
        return;
      }
      listEl.style.display = "none";
      listEl.innerHTML = "<p>よみこみ中...</p>";
      listEl.style.display = "flex";
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_pending_external_requests", adminPin: pin }) })
      .then(r=>r.json()).then(d=>{
        if(d.status !== "success") {
          listEl.style.display = "none";
          msg.innerText = d.message || "エラー";
          return;
        }
        document.getElementById('external-admin-pin-panel').style.display = "none";
        msg.innerText = "";
        renderExternalAdminList(d.list || [], pin);
      }).catch(() => {
        listEl.style.display = "none";
        msg.innerText = "通信エラーが発生しました。";
      });
    }

    function renderExternalAdminList(list, pin) {
      const listEl = document.getElementById('external-admin-list');
      listEl.style.display = "flex";
      if (list.length === 0) {
        listEl.innerHTML = "<p style='text-align:center;color:#888;'>申請中のものはありません。</p>";
        const retry = document.createElement('button');
        retry.type = "button";
        retry.className = "submit-btn btn-gray";
        retry.innerText = "PINを入れ直す";
        retry.onclick = () => {
          document.getElementById('external-admin-pin-panel').style.display = "block";
          listEl.style.display = "none";
          listEl.innerHTML = "";
        };
        listEl.appendChild(retry);
        return;
      }
      listEl.innerHTML = "";
      list.forEach(item => {
        const card = document.createElement('div');
        card.style.cssText = "background:#2a2a2a;border-radius:12px;padding:14px;border:1px solid #444;";
        card.innerHTML = `<div style="font-weight:bold;margin-bottom:6px;">${item.userName} <span style="font-size:12px;color:#aaa;">(${item.userId})</span></div>
          <div style="font-size:15px;margin-bottom:4px;">${item.category || ""} / ${item.volume || ""}</div>
          <div style="color:gold;margin-bottom:8px;">+${item.points} Pt</div>
          <div style="font-size:12px;color:#888;margin-bottom:6px;">${item.requestedAt}</div>
          ${item.childMemo ? `<div style="font-size:13px;color:#ddd;margin-bottom:8px;">こどもメモ: ${item.childMemo}</div>` : ""}
          <textarea class="external-admin-memo" rows="2" placeholder="おとなメモ（任意）" style="width:100%;box-sizing:border-box;border-radius:8px;padding:8px;background:#222;color:#fff;border:1px solid #555;margin-bottom:8px;"></textarea>
          <div style="display:flex;gap:10px;flex-wrap:wrap;">
            <button type="button" class="submit-btn btn-green" style="flex:1;min-width:100px;">承認する</button>
            <button type="button" class="cancel-btn" style="flex:1;min-width:100px;background:#663333;color:#fcc;">却下</button>
          </div>`;
        const btns = card.querySelectorAll('button');
        const memoField = card.querySelector('.external-admin-memo');
        btns[0].onclick = () => decideExternalRequest(item.rowIdx, pin, true, memoField ? memoField.value : "");
        btns[1].onclick = () => decideExternalRequest(item.rowIdx, pin, false, memoField ? memoField.value : "");
        listEl.appendChild(card);
      });
      const back = document.createElement('button');
      back.type = "button";
      back.className = "submit-btn btn-gray";
      back.style.marginTop = "8px";
      back.innerText = "PINを入れ直して再読み込み";
      back.onclick = () => {
        document.getElementById('external-admin-pin-panel').style.display = "block";
        listEl.style.display = "none";
        listEl.innerHTML = "";
      };
      listEl.appendChild(back);
    }

    function decideExternalRequest(rowIdx, pin, approve, memo = "") {
      if (!confirm(approve ? "この申請を承認してポイントを付与しますか？" : "この申請を却下しますか？")) return;
      const action = approve ? "approve_external_request" : "reject_external_request";
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: action, adminPin: pin, rowIdx: rowIdx, adminMemo: memo }) })
      .then(r=>r.json()).then(d=>{
        if(d.status !== "success") {
          alert(d.message || "エラー");
          return;
        }
        alert(d.message);
        const user = JSON.parse(localStorage.getItem('app_kid_user') || "{}");
        if (approve && d.userId && user.id === d.userId && typeof d.newTotal === "number") {
          user.points = d.newTotal;
          localStorage.setItem('app_kid_user', JSON.stringify(user));
          const pts = document.getElementById('user-points');
          if (pts) pts.innerText = user.points;
        }
        loadExternalAdminPendingAfterPin(pin);
      }).catch(() => alert("通信エラーが発生しました。"));
    }

    function loadExternalAdminPendingAfterPin(pin) {
      const msg = document.getElementById('external-admin-message');
      const listEl = document.getElementById('external-admin-list');
      msg.innerText = "";
      listEl.innerHTML = "<p>更新中...</p>";
      listEl.style.display = "flex";
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_pending_external_requests", adminPin: pin }) })
      .then(r=>r.json()).then(d=>{
        if(d.status !== "success") {
          msg.innerText = d.message || "エラー";
          listEl.innerHTML = "";
          return;
        }
        renderExternalAdminList(d.list || [], pin);
      }).catch(() => { msg.innerText = "通信エラー"; });
    }

    const USER_PREF_KANJI_HW_DOMINANT = 'kanji_quiz_hw_dominant_hand';
    function getKanjiHwDominantHand() {
      const v = getUserPref(USER_PREF_KANJI_HW_DOMINANT, null);
      return v === 'left' ? 'left' : 'right';
    }
    function applyKanjiHwDominantHandToBody() {
      document.body.classList.toggle('kanji-hw-dominant-left', getKanjiHwDominantHand() === 'left');
    }
    function syncKanjiHwHandSwitchUI() {
      const lefty = getKanjiHwDominantHand() === 'left';
      const bR = document.getElementById('kanji-hw-hand-right');
      const bL = document.getElementById('kanji-hw-hand-left');
      if (bR) {
        bR.classList.toggle('active', !lefty);
        bR.setAttribute('aria-pressed', !lefty ? 'true' : 'false');
      }
      if (bL) {
        bL.classList.toggle('active', lefty);
        bL.setAttribute('aria-pressed', lefty ? 'true' : 'false');
      }
    }
    function setKanjiHwDominantHand(which) {
      setUserPref(USER_PREF_KANJI_HW_DOMINANT, which === 'left' ? 'left' : 'right');
      applyKanjiHwDominantHandToBody();
      syncKanjiHwHandSwitchUI();
    }

    // （いつもの学習のロード関連）
    function openKanjiLearningMenu() {
      switchSection('section-kanji-learning');
      applyKanjiHwDominantHandToBody();
      syncKanjiHwHandSwitchUI();
      renderKanjiResumePanel();
      const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
      promptKanjiQuizResumeIfNeeded(user);
    }
    function loadEnglishMaterials(btn) {
      currentMaterialsCategory = "english";
      loadMaterialsByFilter(btn, "english");
    }
    function loadKanjiMaterials(btn) {
      currentMaterialsCategory = "kanji";
      loadMaterialsByFilter(btn, "kanji");
    }
    function backFromMaterials() {
      if (currentMaterialsCategory === "kanji") {
        openKanjiLearningMenu();
        return;
      }
      showHome(JSON.parse(localStorage.getItem('app_kid_user')));
    }
    let kpCatalogState = {
      materials: [],
      sets: [],
      setQuestions: [],
      filteredChars: [],
      loaded: false
    };
    function kpCacheKeySets(modeId, unitName) {
      return `app_cached_kp_sets_${modeId}_${unitName}`;
    }
    function kpCacheKeyQuestions(modeId, unitName, setId) {
      return `app_cached_kp_questions_${modeId}_${unitName}_${setId}`;
    }
    function kpGetCachedJson(key) {
      try {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        return JSON.parse(raw);
      } catch (_) {
        return null;
      }
    }
    function kpSetCachedJson(key, value) {
      try {
        localStorage.setItem(key, JSON.stringify(value));
      } catch (_) {}
    }
    let __knKpSyncLock = false;
    function syncKnSelectorsFromKp() {
      if (__knKpSyncLock) return;
      const knb = document.getElementById("kn-book-select");
      if (!knb) return;
      __knKpSyncLock = true;
      try {
        const kpb = document.getElementById("kp-book-select");
        const kpsh = document.getElementById("kp-sheet-select");
        const kpst = document.getElementById("kp-set-select");
        const knsh = document.getElementById("kn-sheet-select");
        const knst = document.getElementById("kn-set-select");
        if (kpb && knb) {
          knb.innerHTML = kpb.innerHTML;
          knb.value = kpb.value;
        }
        if (kpsh && knsh) {
          knsh.innerHTML = kpsh.innerHTML;
          knsh.value = kpsh.value;
        }
        if (kpst && knst) {
          const prev = knst.value === "__ALL__" ? "__ALL__" : knst.value;
          knst.innerHTML =
            '<option value="__ALL__">このシートの全セット</option>' + kpst.innerHTML;
          if (prev === "__ALL__") knst.value = "__ALL__";
          else if (Array.from(knst.options).some(function (o) { return o.value === prev; })) knst.value = prev;
          else knst.value = kpst.value || "__ALL__";
        }
      } finally {
        __knKpSyncLock = false;
      }
    }
    function knOnBookChange() {
      const bsel = document.getElementById("kn-book-select");
      const ssel = document.getElementById("kn-sheet-select");
      if (!bsel || !ssel) return;
      const modeId = bsel.value;
      const mat = kpCatalogState.materials.find(function (m) { return String(m.modeId) === String(modeId); });
      const units = mat && Array.isArray(mat.units) ? mat.units : [];
      if (!units.length) {
        ssel.innerHTML = '<option value="">シートなし</option>';
        return;
      }
      ssel.innerHTML = units.map(function (u) {
        return `<option value="${escapeHtml(u)}">${escapeHtml(formatUnitSheetDisplayLabel(u))}</option>`;
      }).join("");
      knOnSheetChange();
    }
    function knOnSheetChange() {
      const bsel = document.getElementById("kn-book-select");
      const ssel = document.getElementById("kn-sheet-select");
      const setSel = document.getElementById("kn-set-select");
      if (!bsel || !ssel || !setSel) return;
      const modeId = bsel.value;
      const unitName = ssel.value;
      setSel.innerHTML = '<option value="">セット読み込み中...</option>';
      const setsKey = kpCacheKeySets(modeId, unitName);
      const cachedSets = kpGetCachedJson(setsKey);
      if (cachedSets && cachedSets.status === "success" && Array.isArray(cachedSets.sets)) {
        if (!cachedSets.sets.length) {
          setSel.innerHTML = '<option value="">セットなし</option>';
          return;
        }
        setSel.innerHTML =
          '<option value="__ALL__">このシートの全セット</option>' +
          cachedSets.sets.map(function (s) {
            return `<option value="${escapeHtml(String(s.setId || ""))}">セット ${escapeHtml(String(s.setId || ""))}（${escapeHtml(String(s.count || 0))}字）</option>`;
          }).join("");
        setSel.value = "__ALL__";
        return;
      }
      fetch(GAS_API_URL, { method: "POST", body: JSON.stringify({ action: "get_kanji_quiz_sets", modeId: modeId, unitName: unitName }) })
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (d.status !== "success") throw new Error(d.message || "セット取得失敗");
          const sets = Array.isArray(d.sets) ? d.sets : [];
          if (!sets.length) {
            setSel.innerHTML = '<option value="">セットなし</option>';
            return;
          }
          setSel.innerHTML =
            '<option value="__ALL__">このシートの全セット</option>' +
            sets.map(function (s) {
              return `<option value="${escapeHtml(String(s.setId || ""))}">セット ${escapeHtml(String(s.setId || ""))}（${escapeHtml(String(s.count || 0))}字）</option>`;
            }).join("");
          setSel.value = "__ALL__";
        })
        .catch(function () {
          setSel.innerHTML = '<option value="">セット取得失敗</option>';
        });
    }
    function knOnSetChange() {}
    function openKanjiNigateSection() {
      initKanjiPracticeCatalog();
      initKanjiHandAnalyticsBridge();
      initKanjiParentKanjiQuizScoredBridge();
      setTimeout(function () {
        syncKnSelectorsFromKp();
        const kna = document.getElementById("kn-nigate-axis");
        if (kna) kna.value = "stroke_order";
        switchSection("section-kanji-nigate");
      }, 120);
    }
    function startKanjiNigateReviewFromUi() {
      const user = JSON.parse(localStorage.getItem("app_kid_user") || "null");
      if (!user || !user.id) {
        alert("ログインしてください。");
        return;
      }
      const b = document.getElementById("kn-book-select");
      const u = document.getElementById("kn-sheet-select");
      const s = document.getElementById("kn-set-select");
      const ax = document.getElementById("kn-nigate-axis");
      if (!b || !u || !s || !ax) return;
      const modeId = b.value;
      const unitName = u.value;
      if (!modeId || !unitName) {
        alert("ブックとシートをえらんでください。");
        return;
      }
      const setVal = s.value;
      if (!setVal) {
        alert("セットをえらんでください。");
        return;
      }
      var setIds = setVal === "__ALL__" ? [] : [setVal];
      const nigateAxis = ax.value;
      fetch(GAS_API_URL, {
        method: "POST",
        body: JSON.stringify({
          action: "get_kanji_weak_review_plan",
          userId: user.id,
          modeId: modeId,
          unitName: unitName,
          setIds: setIds,
          nigateAxis: nigateAxis,
          limit: 12
        })
      })
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (!d || d.status !== "success") {
            alert((d && d.message) || "取得に失敗しました。");
            return;
          }
          if (!d.questions || !d.questions.length) {
            alert(
              d.message ||
                "この条件ではもんだいがありません。先に通常のクイズ・練習で学習して、よわみデータをためましょう。"
            );
            return;
          }
          const mat = kpCatalogState.materials.find(function (m) { return String(m.modeId) === String(modeId); });
          startKanjiQuizPlay({
            modeId: modeId,
            modeName: mat ? mat.modeName || modeId : modeId,
            unitName: unitName,
            setId: setVal === "__ALL__" ? "ALL" : setVal,
            questions: d.questions,
            nigateBypassFilter: true,
            nigateTraining: true,
            nigateAxis: nigateAxis,
            formatMode: "mixed"
          });
        })
        .catch(function () {
          alert("通信エラーが発生しました。");
        });
    }
    function initKanjiHandAnalyticsBridge() {
      if (window.__kanjiHandAnalyticsBridgeBound) return;
      window.__kanjiHandAnalyticsBridgeBound = true;
      var t = null;
      var pending = null;
      function flush() {
        if (!pending) return;
        var sig = pending;
        pending = null;
        var kid = JSON.parse(localStorage.getItem("app_kid_user") || "null");
        if (!kid || !kid.id) return;
        var q =
          kanjiQuizSession && kanjiQuizSession.questions
            ? kanjiQuizSession.questions[kanjiQuizSession.index]
            : null;
        var quizSec = document.getElementById("section-kanji-quiz-play");
        if (quizSec && quizSec.classList.contains("active") && kanjiQuizSession) {
          sig.modeId = kanjiQuizSession.modeId;
          sig.unitName = kanjiQuizSession.unitName;
          sig.setId = q && q.nigateSourceSetId ? String(q.nigateSourceSetId) : String(kanjiQuizSession.setId || "");
          if (sig.setId === "ALL") sig.setId = q && q.nigateSourceSetId ? String(q.nigateSourceSetId) : "";
          sig.questionId = q && q.questionId ? q.questionId : undefined;
        } else {
          var prac = document.getElementById("section-kanji-practice");
          if (prac && prac.classList.contains("active")) {
            var bEl = document.getElementById("kp-book-select");
            var uEl = document.getElementById("kp-sheet-select");
            var sEl = document.getElementById("kp-set-select");
            sig.modeId = (bEl && bEl.value) || "";
            sig.unitName = (uEl && uEl.value) || "";
            sig.setId = (sEl && sEl.value) || "";
            sig.questionId = undefined;
          }
        }
        if (!sig.modeId || !sig.unitName || !sig.setId || !sig.kanjiChar) return;
        if (kanjiQuizSession && kanjiQuizSession.nigateTraining) {
          if (!kanjiQuizSession.nigateFeedback) {
            kanjiQuizSession.nigateFeedback = { strokeOrderClean: true, brushAllClear: true };
          }
          if (sig.hasStrokeOrderIssue) kanjiQuizSession.nigateFeedback.strokeOrderClean = false;
          if (sig.brushEndingAllOk === false) kanjiQuizSession.nigateFeedback.brushAllClear = false;
        }
        fetch(GAS_API_URL, {
          method: "POST",
          body: JSON.stringify({ action: "append_kanji_weak_signals", userId: kid.id, signals: [sig] })
        }).catch(function () {});
      }
      window.addEventListener("message", function (ev) {
        if (!ev || !ev.data || ev.data.type !== "kanjiQuizHandAnalytics") return;
        var d = ev.data;
        pending = {
          at: new Date().toISOString(),
          kanjiChar: String(d.kanjiChar || ""),
          hasStrokeOrderIssue: !!d.hasStrokeOrderIssue,
          brushEndingAllOk: d.brushEndingAllOk !== false,
          strokeCountMismatch: !!d.strokeCountMismatch,
          readingMistake: false,
          handScore: typeof d.handScore === "number" ? d.handScore : undefined,
          passedThreshold: !!d.passedThreshold
        };
        if (t) clearTimeout(t);
        t = setTimeout(flush, 420);
      });
    }
    function queueKanjiStrokeCountWeakSignal() {
      if (!kanjiQuizSession) return;
      var kid = JSON.parse(localStorage.getItem("app_kid_user") || "null");
      if (!kid || !kid.id) return;
      var q = kanjiQuizSession.questions[kanjiQuizSession.index];
      if (!q) return;
      var setId = q.nigateSourceSetId ? String(q.nigateSourceSetId) : String(kanjiQuizSession.setId || "");
      if (!setId || setId === "ALL") return;
      fetch(GAS_API_URL, {
        method: "POST",
        body: JSON.stringify({
          action: "append_kanji_weak_signals",
          userId: kid.id,
          signals: [
            {
              at: new Date().toISOString(),
              modeId: kanjiQuizSession.modeId,
              unitName: kanjiQuizSession.unitName,
              setId: setId,
              kanjiChar: String(q.kanji || ""),
              questionId: q.questionId,
              strokeCountQuizWrong: true,
              hasStrokeOrderIssue: false,
              brushEndingAllOk: true,
              strokeCountMismatch: false,
              readingMistake: false
            }
          ]
        })
      }).catch(function () {});
    }
    function queueKanjiReadingWeakSignal() {
      if (!kanjiQuizSession) return;
      var kid = JSON.parse(localStorage.getItem("app_kid_user") || "null");
      if (!kid || !kid.id) return;
      var q = kanjiQuizSession.questions[kanjiQuizSession.index];
      if (!q) return;
      var setId = q.nigateSourceSetId ? String(q.nigateSourceSetId) : String(kanjiQuizSession.setId || "");
      if (!setId || setId === "ALL") return;
      fetch(GAS_API_URL, {
        method: "POST",
        body: JSON.stringify({
          action: "append_kanji_weak_signals",
          userId: kid.id,
          signals: [
            {
              at: new Date().toISOString(),
              modeId: kanjiQuizSession.modeId,
              unitName: kanjiQuizSession.unitName,
              setId: setId,
              kanjiChar: String(q.kanji || ""),
              questionId: q.questionId,
              readingMistake: true,
              hasStrokeOrderIssue: false,
              brushEndingAllOk: true,
              strokeCountMismatch: false
            }
          ]
        })
      }).catch(function () {});
    }
    function openKanjiPractice() {
      switchSection('section-kanji-practice');
      openKanjiPracticePro();
      initKanjiPracticeScoreListener();
      initKanjiPracticeCatalog();
    }
    /** 漢字練習画面用: iframe 採点メッセージの受け口は initKanjiParentKanjiQuizScoredBridge と共有 */
    function initKanjiPracticeScoreListener() {
      initKanjiParentKanjiQuizScoredBridge();
      initKanjiHandAnalyticsBridge();
    }
    function openKanjiPracticePro() {
      const frame = document.getElementById('kp-pro-frame');
      if (!frame) return Promise.resolve();
      const KP_EMBED_VER = "12";
      if (frame.dataset.kpEmbedVer !== KP_EMBED_VER) {
        frame.dataset.kpLoaded = "";
        frame.dataset.kpEmbedVer = KP_EMBED_VER;
      }
      if (frame.dataset.kpLoaded === "1") {
        kpResizeFrameToContent();
        return Promise.resolve();
      }
      const kpHtml = typeof KP_IFRAME_HTML === "string" ? KP_IFRAME_HTML : "";
      if (!kpHtml) {
        alert("漢字ドリル画面の HTML が code.js 内に未定義です（KP_IFRAME_HTML）。");
        return Promise.reject(new Error("KP_IFRAME_HTML"));
      }
      frame.srcdoc = kpHtml;
      frame.dataset.kpLoaded = "1";
      delete frame.dataset.kanjiQuizPatched;
      setTimeout(kpResizeFrameToContent, 120);
      setTimeout(kpResizeFrameToContent, 700);
      return Promise.resolve();
    }
    function kpFrameInWrongModelWrap_(frame) {
      return frame && frame.parentElement && frame.parentElement.id === "kanji-quiz-wrong-model-wrap";
    }
    function kpComputeFrameHeight_(frame, bodyH, htmlH) {
      if (kpFrameInWrongModelWrap_(frame)) {
        const cap = Math.min(window.innerHeight * 0.72, 760);
        return Math.min(Math.max(bodyH, htmlH, 260), cap) + 8;
      }
      return Math.max(bodyH, htmlH, 620) + 8;
    }
    function kpResizeFrameToContent() {
      const frame = document.getElementById('kp-pro-frame');
      if (!frame || !frame.contentWindow || !frame.contentWindow.document) return;
      try {
        const doc = frame.contentWindow.document;
        const bodyH = doc.body ? doc.body.scrollHeight : 0;
        const htmlH = doc.documentElement ? doc.documentElement.scrollHeight : 0;
        const h = kpComputeFrameHeight_(frame, bodyH, htmlH);
        frame.style.height = `${h}px`;
        if (!frame.dataset.kpObserved) {
          const ro = new ResizeObserver(() => {
            const bH = doc.body ? doc.body.scrollHeight : 0;
            const dH = doc.documentElement ? doc.documentElement.scrollHeight : 0;
            const nextH = kpComputeFrameHeight_(frame, bH, dH);
            frame.style.height = `${nextH}px`;
          });
          if (doc.body) ro.observe(doc.body);
          if (doc.documentElement) ro.observe(doc.documentElement);
          frame.dataset.kpObserved = "1";
        }
      } catch (_) {}
    }
    function setKpStatus(text) {
      const el = document.getElementById('kp-practice-status');
      if (el) el.innerText = text || "";
    }
    function initKanjiPracticeCatalog() {
      if (kpCatalogState.loaded && kpCatalogState.materials.length) return;
      setKpStatus("教材を取得中...");
      const fromMaterialsCache = (() => {
        const d = kpGetCachedJson(LS_APP_CACHED_MATERIALS);
        if (!d || d.status !== "success" || !Array.isArray(d.materials)) return null;
        return d.materials;
      })();
      const pickKanjiMaterials = (all) => {
        return (Array.isArray(all) ? all : []).filter(m => {
          const cat = String(m.category || "").toLowerCase();
          const modeName = String(m.modeName || "");
          const modeId = String(m.modeId || "");
          const joined = `${modeName} ${modeId}`;
          return (cat === "kanji") || /漢字|かんじ|kanji/i.test(joined);
        });
      };
      const applyMaterials = (all) => {
        kpCatalogState.materials = pickKanjiMaterials(all);
        kpCatalogState.loaded = true;
        kpRenderBookSelect();
      };
      const fetchMaterialsFresh = () => {
        fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_materials_list" }) })
        .then(r=>r.json()).then(d=>{
          if (d.status !== "success") throw new Error(d.message || "教材取得失敗");
          persistMaterialsPayload(d);
          applyMaterials(materialsData);
          if (!kpCatalogState.materials.length) {
            setKpStatus("漢字教材が0件です（設定または分類を確認してください）。");
          }
        }).catch(_ => {
          setKpStatus("教材の取得に失敗しました。");
        });
      };
      if (fromMaterialsCache) {
        const cachedKanji = pickKanjiMaterials(fromMaterialsCache);
        if (cachedKanji.length) {
          applyMaterials(fromMaterialsCache);
          setKpStatus("教材キャッシュを使用しました。");
          return;
        }
        // キャッシュが古く漢字教材が含まれない場合は自動で再取得
        setKpStatus("漢字教材キャッシュを更新中...");
        fetchMaterialsFresh();
        return;
      }
      fetchMaterialsFresh();
    }
    function kpRenderBookSelect() {
      const sel = document.getElementById('kp-book-select');
      if (!sel) return;
      const mats = kpCatalogState.materials;
      if (!mats.length) {
        sel.innerHTML = '<option value="">ブックなし</option>';
        setKpStatus("漢字教材が見つかりません。");
        syncKnSelectorsFromKp();
        return;
      }
      sel.innerHTML = mats.map(m => `<option value="${escapeHtml(m.modeId)}">${escapeHtml(m.modeName || m.modeId)}</option>`).join('');
      kpOnBookChange();
    }
    function kpOnBookChange() {
      const bsel = document.getElementById('kp-book-select');
      const ssel = document.getElementById('kp-sheet-select');
      if (!bsel || !ssel) return;
      const modeId = bsel.value;
      const mat = kpCatalogState.materials.find(m => String(m.modeId) === String(modeId));
      const units = mat && Array.isArray(mat.units) ? mat.units : [];
      if (!units.length) {
        ssel.innerHTML = '<option value="">シートなし</option>';
        setKpStatus("シートがありません。");
        syncKnSelectorsFromKp();
        return;
      }
      ssel.innerHTML = units.map(u => `<option value="${escapeHtml(u)}">${escapeHtml(formatUnitSheetDisplayLabel(u))}</option>`).join('');
      kpOnSheetChange();
    }
    function kpOnSheetChange() {
      const bsel = document.getElementById('kp-book-select');
      const ssel = document.getElementById('kp-sheet-select');
      const setSel = document.getElementById('kp-set-select');
      if (!bsel || !ssel || !setSel) return;
      const modeId = bsel.value;
      const unitName = ssel.value;
      kpCatalogState.sets = [];
      kpCatalogState.setQuestions = [];
      setSel.innerHTML = '<option value="">セット読み込み中...</option>';
      setKpStatus(`セットを取得中: ${formatUnitSheetDisplayLabel(unitName)}`);
      const setsKey = kpCacheKeySets(modeId, unitName);
      const cachedSets = kpGetCachedJson(setsKey);
      if (cachedSets && cachedSets.status === "success" && Array.isArray(cachedSets.sets)) {
        kpCatalogState.sets = cachedSets.sets;
        if (!kpCatalogState.sets.length) {
          setSel.innerHTML = '<option value="">セットなし</option>';
          setKpStatus("このシートにセットがありません。");
          kpRenderTiles([]);
          syncKnSelectorsFromKp();
          return;
        }
        setSel.innerHTML = kpCatalogState.sets.map(s => `<option value="${escapeHtml(String(s.setId || ''))}">セット ${escapeHtml(String(s.setId || ''))}（${escapeHtml(String(s.count || 0))}字）</option>`).join('');
        setKpStatus(`セットをキャッシュから読み込みました: ${formatUnitSheetDisplayLabel(unitName)}`);
        kpOnSetChange();
        syncKnSelectorsFromKp();
        return;
      }
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_kanji_quiz_sets", modeId, unitName }) })
      .then(r=>r.json()).then(d=>{
        if (d.status !== "success") throw new Error(d.message || "セット取得失敗");
        kpCatalogState.sets = Array.isArray(d.sets) ? d.sets : [];
        kpSetCachedJson(setsKey, { status: "success", sets: kpCatalogState.sets });
        if (!kpCatalogState.sets.length) {
          setSel.innerHTML = '<option value="">セットなし</option>';
          setKpStatus("このシートにセットがありません。");
          kpRenderTiles([]);
          syncKnSelectorsFromKp();
          return;
        }
        setSel.innerHTML = kpCatalogState.sets.map(s => `<option value="${escapeHtml(String(s.setId || ''))}">セット ${escapeHtml(String(s.setId || ''))}（${escapeHtml(String(s.count || 0))}字）</option>`).join('');
        // 初回のみシート内セット問題を先読みして、以後の切替を高速化
        kpCatalogState.sets.forEach(s => {
          const sid = String(s.setId || "");
          if (!sid) return;
          const qKey = kpCacheKeyQuestions(modeId, unitName, sid);
          if (kpGetCachedJson(qKey)) return;
          fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_kanji_quiz_questions", modeId, unitName, setId: sid }) })
            .then(r => r.json())
            .then(q => {
              if (q && q.status === "success" && Array.isArray(q.questions)) {
                kpSetCachedJson(qKey, { status: "success", questions: q.questions });
              }
            })
            .catch(() => {});
        });
        kpOnSetChange();
        syncKnSelectorsFromKp();
      }).catch(_ => {
        setSel.innerHTML = '<option value="">セット取得失敗</option>';
        setKpStatus("セットの取得に失敗しました。");
        kpRenderTiles([]);
        syncKnSelectorsFromKp();
      });
    }
    function kpOnSetChange() {
      const bsel = document.getElementById('kp-book-select');
      const ssel = document.getElementById('kp-sheet-select');
      const setSel = document.getElementById('kp-set-select');
      if (!bsel || !ssel || !setSel) return;
      const modeId = bsel.value;
      const unitName = ssel.value;
      const setId = setSel.value;
      if (!setId) {
        kpCatalogState.setQuestions = [];
        kpRenderTiles([]);
        syncKnSelectorsFromKp();
        return;
      }
      setKpStatus(`セット ${setId} を読み込み中...`);
      const qKey = kpCacheKeyQuestions(modeId, unitName, setId);
      const cachedQ = kpGetCachedJson(qKey);
      if (cachedQ && cachedQ.status === "success" && Array.isArray(cachedQ.questions)) {
        kpCatalogState.setQuestions = cachedQ.questions;
        kpApplyFilterAndRender();
        setKpStatus(`セット ${setId} をキャッシュから表示中`);
        syncKnSelectorsFromKp();
        return;
      }
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_kanji_quiz_questions", modeId, unitName, setId }) })
      .then(r=>r.json()).then(d=>{
        if (d.status !== "success") throw new Error(d.message || "問題取得失敗");
        kpCatalogState.setQuestions = Array.isArray(d.questions) ? d.questions : [];
        kpSetCachedJson(qKey, { status: "success", questions: kpCatalogState.setQuestions });
        kpApplyFilterAndRender();
        syncKnSelectorsFromKp();
      }).catch(_ => {
        kpCatalogState.setQuestions = [];
        kpRenderTiles([]);
        setKpStatus("セット問題の取得に失敗しました。");
        syncKnSelectorsFromKp();
      });
    }
    function kpApplyFilterAndRender() {
      const q = (document.getElementById('kp-search-input')?.value || "").trim().toLowerCase();
      const list = (kpCatalogState.setQuestions || []).filter(item => {
        const k = String(item.kanji || "");
        const fromNew = String(item.searchText || "").toLowerCase();
        const readings = (Array.isArray(item.readings) ? item.readings : [])
          .map(r => `${r.reading || ""} ${(Array.isArray(r.examples) ? r.examples.join(" ") : "")}`)
          .join(" ")
          .toLowerCase();
        const hay = fromNew || `${k} ${readings}`;
        if (!q) return true;
        return k.toLowerCase().includes(q) || hay.includes(q);
      }).map(item => String(item.kanji || "")).filter(Boolean);
      kpCatalogState.filteredChars = list;
      kpRenderTiles(list);
      setKpStatus(`${list.length} 件表示中 / セット内 ${(kpCatalogState.setQuestions || []).length} 問`);
    }
    function kpRenderTiles(chars) {
      const grid = document.getElementById('kp-tile-grid');
      if (!grid) return;
      grid.innerHTML = "";
      if (!chars || !chars.length) {
        grid.innerHTML = '<div style="grid-column:1/-1;color:#999;text-align:center;padding:18px;">候補なし</div>';
        return;
      }
      const uniq = Array.from(new Set(chars));
      uniq.forEach(ch => {
        const b = document.createElement('button');
        b.type = "button";
        b.className = "menu-btn btn-gray";
        b.style.margin = "0";
        b.style.width = "100%";
        b.style.padding = "14px 8px";
        b.style.fontSize = "52px";
        b.style.lineHeight = "1";
        b.style.color = "#111";
        b.style.border = "1px solid #cfd8dc";
        b.style.background = "#ffffff";
        b.style.boxShadow = "none";
        b.innerText = ch;
        b.title = `${ch} を練習`;
        b.onclick = () => kpSelectPracticeChar(ch);
        grid.appendChild(b);
      });
    }
    function kpSelectPracticeChar(ch) {
      const frame = document.getElementById('kp-pro-frame');
      if (!frame || !frame.contentWindow) return;
      try {
        const doc = frame.contentWindow.document;
        const sel = doc.getElementById('target-kanji');
        if (!sel) return;
        const has = Array.from(sel.options || []).some(o => String(o.value) === String(ch));
        if (!has) {
          setKpStatus(`「${ch}」は筆順データ（KanjiVG）未登録です。`);
          return;
        }
        sel.value = ch;
        if (typeof frame.contentWindow.initTargetKanji === "function") frame.contentWindow.initTargetKanji();
        if (typeof frame.contentWindow.switchMode === "function") frame.contentWindow.switchMode("score");
        kpResizeFrameToContent();
        setKpStatus(`練習対象を「${ch}」に切替えました。`);
      } catch (_) {
        setKpStatus("練習対象の切替に失敗しました。");
      }
    }
    function kpShowModelDemo() {
      const frame = document.getElementById('kp-pro-frame');
      if (!frame || !frame.contentWindow) return;
      try {
        if (typeof frame.contentWindow.switchMode === "function") {
          frame.contentWindow.switchMode("demo");
          kpResizeFrameToContent();
          setKpStatus("お手本表示モードに切り替えました。");
        } else {
          setKpStatus("お手本表示機能の呼び出しに失敗しました。");
        }
      } catch (_) {
        setKpStatus("お手本表示機能の呼び出しに失敗しました。");
      }
    }
    function kpReplayModelDemo() {
      const frame = document.getElementById('kp-pro-frame');
      if (!frame || !frame.contentWindow) return;
      try {
        if (typeof frame.contentWindow.switchMode === "function") frame.contentWindow.switchMode("demo");
        if (typeof frame.contentWindow.playAnimation === "function") {
          frame.contentWindow.playAnimation();
          kpResizeFrameToContent();
          setKpStatus("書き順デモを再生しました。");
        } else {
          setKpStatus("デモ再生機能の呼び出しに失敗しました。");
        }
      } catch (_) {
        setKpStatus("デモ再生機能の呼び出しに失敗しました。");
      }
    }
    function loadMaterialsByFilter(btn, category) {
      const origText = toggleBtnLoading(btn, true);
      document.getElementById('materials-container').innerHTML = "<p>よみこみ中...</p>";
      ensureMaterialsListLoaded()
        .then(() => {
          toggleBtnLoading(btn, false, origText);
          switchSection('section-materials');
          const title = document.getElementById('materials-title');
          if (title) title.innerText = category === "kanji" ? "漢字クイズに挑戦" : "英語の自由学習";
          const c = document.getElementById('materials-container');
          c.innerHTML = "";
          const all = Array.isArray(materialsData) ? materialsData : [];
          const list = all.filter(m => {
            const cat = String(m.category || "").toLowerCase();
            const name = String(m.modeName || "");
            const isKanji = (cat === "kanji") || /漢字|かんじ|kanji/i.test(name);
            return category === "kanji" ? isKanji : !isKanji;
          });
          if (list.length === 0) {
            c.innerHTML = "<p>表示できる教材がありません。</p>";
            return;
          }
          list.forEach(m => {
            const t = document.createElement('h2');
            t.innerText = `📁 ${m.modeName}`;
            c.appendChild(t);
            m.units.forEach(u => {
              const b = document.createElement('button');
              b.className="menu-btn btn-gray";
              b.innerText = `📄 ${formatUnitSheetDisplayLabel(u)}`;
              b.onclick = () => { loadQuestionsForSettings(b, m.modeId, m.modeName, u, category); };
              c.appendChild(b);
            });
          });
        })
        .catch(() => {
          toggleBtnLoading(btn, false, origText);
          document.getElementById('materials-container').innerHTML = "<p>よみこみに失敗しました。</p>";
        });
    }
    const LS_KANJI_QUIZ_FORMAT = 'app_kanji_quiz_format_v1';
    function getKanjiQuizFormatMode() {
      const sel = document.getElementById('kanji-quiz-format-select');
      if (sel && sel.value) return sel.value;
      try {
        const v = localStorage.getItem(LS_KANJI_QUIZ_FORMAT);
        if (v && ['mixed', 'write_kanji', 'select_kana', 'type_yomi'].indexOf(v) >= 0) return v;
      } catch (e) {}
      return 'write_kanji';
    }
    function onKanjiQuizFormatChange() {
      const sel = document.getElementById('kanji-quiz-format-select');
      if (sel) {
        try { localStorage.setItem(LS_KANJI_QUIZ_FORMAT, sel.value); } catch (e) {}
      }
    }
    function syncKanjiQuizFormatSelectFromStorage() {
      const sel = document.getElementById('kanji-quiz-format-select');
      if (!sel) return;
      const v = getKanjiQuizFormatMode();
      sel.value = v;
    }
    function filterKanjiQuizQuestionsByFormat(questions, mode) {
      const arr = Array.isArray(questions) ? questions : [];
      if (!mode || mode === 'mixed') return arr.slice();
      var typeMap = { write_kanji: 'ruby_to_kanji', select_kana: 'okurigana_shift', type_yomi: 'sentence_to_ruby' };
      var t = typeMap[mode];
      if (!t) return arr.slice();
      return arr.filter(function (q) { return q.type === t; });
    }
    function shuffleKanjiQuizQuestionsArray(arr) {
      const a = Array.isArray(arr) ? arr.slice() : [];
      for (let i = a.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const t = a[i];
        a[i] = a[j];
        a[j] = t;
      }
      return a;
    }
    function shuffleKanjiQuizChoicesArray(arr) {
      const a = Array.isArray(arr) ? arr.slice() : [];
      for (let i = a.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const t = a[i];
        a[i] = a[j];
        a[j] = t;
      }
      return a;
    }
    function isHiraganaChar_(ch) {
      if (!ch) return false;
      const cp = ch.codePointAt(0);
      if (cp >= 0x3041 && cp <= 0x3096) return true;
      if (cp === 0x3099 || cp === 0x309a) return true;
      return false;
    }
    function shuffleInPlaceWithRng_(arr, rng) {
      const randomFn = typeof rng === "function" ? rng : Math.random;
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(randomFn() * (i + 1));
        const t = arr[i];
        arr[i] = arr[j];
        arr[j] = t;
      }
      return arr;
    }
    /**
     * 送り仮名境界クイズを自動生成する。
     * 入力: { kanji, reading, sentence }
     * 出力: { question_sentence, target_kanji, options, correct_option } | null
     */
    function generate_okurigana_quiz(input) {
      try {
        const src = input && typeof input === "object" ? input : {};
        const kanji = String(src.kanji || "").trim();
        const reading = String(src.reading || "").trim();
        const sentence = String(src.sentence || "");
        if (!kanji || !reading || !sentence) return null;

        // Step1: 読み2文字以上、かつ例文中に対象漢字があるもののみ対象
        const readingChars = Array.from(reading);
        if (readingChars.length < 2) return null;

        // Step1/2: 文章内で「漢字の直後に連続ひらがながある」出現箇所を探す
        // （最初の indexOf が条件を満たさないケースもあるため、全出現から候補を抽出する）
        const candidates = [];
        let searchFrom = 0;
        while (true) {
          const idx = String(sentence).indexOf(kanji, searchFrom);
          if (idx < 0) break;
          const afterIdx = idx + kanji.length;
          if (afterIdx < String(sentence).length) {
            let okurigana = "";
            for (let i = afterIdx; i < String(sentence).length; i++) {
              const ch = String(sentence).charAt(i);
              if (!isHiraganaChar_(ch)) break;
              okurigana += ch;
            }
            if (okurigana) candidates.push({ idx: idx, okurigana: okurigana });
          }
          searchFrom = idx + 1; // 次の出現へ
        }
        if (!candidates.length) return null;

        // 候補のどれか1つを採用
        const picked = candidates[Math.floor(Math.random() * candidates.length)];
        const idx = picked.idx;
        const okurigana = picked.okurigana;
        const afterIdx = idx + kanji.length;

        // Step2: 抽出結果（例: 下る）
        const extractedBlock = kanji + okurigana;
        if (!extractedBlock) return null;

        // Step3: フルカナ = 読み + 送り仮名
        const fullKana = reading + okurigana;
        const fullChars = Array.from(fullKana);
        const fullLen = fullChars.length;
        if (fullLen < 2) return null;

        // Step4: 分割位置 k=1..N-1 で選択肢生成
        const readingLen = readingChars.length;
        const optionsRaw = [];
        let correctOption = "";
        for (let k = 1; k <= fullLen - 1; k++) {
          const opt = kanji + fullChars.slice(k).join("");
          if (!opt) continue;
          optionsRaw.push(opt);
          if (k === readingLen) correctOption = opt;
        }
        if (!correctOption) return null;

        // Step5: 重複除去 + 最大5択（正解は必ず残す）+ シャッフル
        let uniqueOptions = Array.from(new Set(optionsRaw));
        if (!uniqueOptions.includes(correctOption)) uniqueOptions.push(correctOption);
        if (uniqueOptions.length < 2) return null;
        if (uniqueOptions.length > 5) {
          const wrongs = uniqueOptions.filter(o => o !== correctOption);
          shuffleInPlaceWithRng_(wrongs);
          uniqueOptions = [correctOption].concat(wrongs.slice(0, 4));
        }
        shuffleInPlaceWithRng_(uniqueOptions);

        const qSentence =
          sentence.slice(0, idx) + "【" + kanji + "】" + sentence.slice(afterIdx);
        return {
          question_sentence: qSentence,
          target_kanji: kanji,
          options: uniqueOptions,
          correct_option: correctOption
        };
      } catch (e) {
        console.warn("generate_okurigana_quiz failed:", e);
        return null;
      }
    }
    function rebuildOkuriganaQuestionByAlgorithm_(q) {
      const src = q && typeof q === "object" ? q : {};
      const generated = generate_okurigana_quiz({
        kanji: src.kanji,
        reading: src.readingDisplay || src.reading || "",
        sentence: src.sentence || src.exampleSentence || ""
      });
      if (!generated) return src;
      const out = Object.assign({}, src);
      out.kanji = generated.target_kanji;
      out.correctAnswer = generated.correct_option;
      out.choices = generated.options.slice();
      out.questionSentence = generated.question_sentence;
      return out;
    }
    function prepareKanjiQuizQuestionsForPlay(rawList) {
      var mode = getKanjiQuizFormatMode();
      var normalized = (Array.isArray(rawList) ? rawList : []).map(function (q) {
        if (!q || q.type !== "okurigana_shift") return q;
        return rebuildOkuriganaQuestionByAlgorithm_(q);
      });
      var filtered = filterKanjiQuizQuestionsByFormat(normalized, mode);
      if (!filtered.length) {
        alert('この しかた では もんだいがありません。\nほかの しかたを えらぶか、混合にしてください。');
        return null;
      }
      // 出題順のシャッフルは startKanjiQuizPlay 内で1回だけ行う（二重シャッフル防止）
      return { questions: filtered, formatMode: mode };
    }
    function kanjiQuizDrillCacheKeySets(modeId, unitName) {
      return "app_cached_kanji_quiz_sets_" + String(modeId || "") + "_" + String(unitName || "");
    }
    function kanjiQuizDrillCacheKeyQuestions(modeId, unitName, setId) {
      return (
        "app_cached_kanji_quiz_questions_" +
        String(modeId || "") +
        "_" +
        String(unitName || "") +
        "_" +
        String(setId || "")
      );
    }
    const __kanjiQuizSetsSessionCache = Object.create(null);
    function openKanjiQuizSets(modeId, modeName, unitName, btn, origText) {
      const title = document.getElementById('kanji-quiz-title');
      if (title) title.innerText = `【${modeName}】${formatUnitSheetDisplayLabel(unitName)}`;
      const box = document.getElementById('kanji-quiz-sets-container');
      if (box) box.innerHTML = "<p>セットを読み込み中...</p>";
      const setsCacheKey = kanjiQuizDrillCacheKeySets(modeId, unitName);
      const renderSetsButtons = function (d) {
        toggleBtnLoading(btn, false, origText);
        if (d.status !== "success") {
          alert("取得失敗: " + (d.message || "エラー"));
          return;
        }
        switchSection('section-kanji-quiz-sets');
        syncKanjiQuizFormatSelectFromStorage();
        if (!box) return;
        box.innerHTML = "";
        const sets = Array.isArray(d.sets) ? d.sets : [];
        if (!sets.length) {
          box.innerHTML = "<p>セットが見つかりません。</p>";
          return;
        }
        const startFromQuestionsPayload = function (q, sid) {
          if (q.status !== "success") {
            alert("取得失敗: " + (q.message || "エラー"));
            return;
          }
          var raw = Array.isArray(q.questions) ? q.questions : [];
          var prep = prepareKanjiQuizQuestionsForPlay(raw);
          if (!prep) return;
          startKanjiQuizPlay({
            modeId: modeId,
            modeName: modeName,
            unitName: unitName,
            setId: String(sid != null ? sid : ""),
            allQuestions: raw,
            formatMode: prep.formatMode
          });
        };
        sets.forEach(s => {
          const b = document.createElement('button');
          b.className = "menu-btn btn-gray";
          b.innerText = `セット ${s.setId}（${s.count}字）`;
          b.onclick = () => {
            const qKey = kanjiQuizDrillCacheKeyQuestions(modeId, unitName, s.setId);
            const cachedQ = localStorage.getItem(qKey);
            if (cachedQ) {
              try {
                const q = JSON.parse(cachedQ);
                if (q.status === "success") {
                  startFromQuestionsPayload(q, q.setId != null ? q.setId : s.setId);
                  return;
                }
              } catch (e) {}
            }
            fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_kanji_quiz_questions", modeId: modeId, unitName: unitName, setId: s.setId }) })
              .then(r => r.json()).then(q => {
                if (q.status === "success") {
                  try { localStorage.setItem(qKey, JSON.stringify(q)); } catch (e) {}
                }
                startFromQuestionsPayload(q, s.setId);
              })
              .catch(() => alert("通信エラーが発生しました。"));
          };
          box.appendChild(b);
        });
        try {
          if (typeof requestIdleCallback === "function") {
            requestIdleCallback(function () {
              ensureKanjiPracticeFrameReady();
            }, { timeout: 2500 });
          } else {
            setTimeout(function () {
              ensureKanjiPracticeFrameReady();
            }, 120);
          }
        } catch (e) {}
      };
      if (__kanjiQuizSetsSessionCache[setsCacheKey]) {
        renderSetsButtons(__kanjiQuizSetsSessionCache[setsCacheKey]);
        return;
      }
      const cachedSets = localStorage.getItem(setsCacheKey);
      if (cachedSets) {
        try {
          const d = JSON.parse(cachedSets);
          if (d.status === "success") {
            __kanjiQuizSetsSessionCache[setsCacheKey] = d;
            renderSetsButtons(d);
            return;
          }
        } catch (e) {}
      }
      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_kanji_quiz_sets", modeId: modeId, unitName: unitName }) })
      .then(r => r.json()).then(d => {
        if (d.status === "success") {
          try { localStorage.setItem(setsCacheKey, JSON.stringify(d)); } catch (e) {}
          __kanjiQuizSetsSessionCache[setsCacheKey] = d;
        }
        renderSetsButtons(d);
      }).catch(e => {
        toggleBtnLoading(btn, false, origText);
        alert("通信エラーが発生しました。");
      });
    }
    const KANJI_QUIZ_HAND_PASS = 60;
    const kanjiQuizWriteLogicalSize = 300;
    let kanjiQuizParentStrokes = [];
    let kanjiQuizCurrentStrokePoints = [];
    let kanjiQuizIsDrawing = false;
    let kanjiQuizPenW = 8;
    let kanjiQuizStrokeStartTime = 0;
    let kanjiQuizScrollLockUntil = 0;
    let __kanjiQuizWriteCanvasBound = false;
    let __kanjiQuizWriteReflowListenersBound = false;
    let __kanjiQuizHwFrameReadyP = null;
    let __kanjiQuizScrollGuardBound = false;
    let __kanjiQuizTouchLockLastBump = 0;
    let __kanjiPracticeLastSubmitAt = 0;
    let __kanjiPracticeLastSubmitKey = "";
    function clearKanjiHwFrameReadyCache() {
      __kanjiQuizHwFrameReadyP = null;
    }
    function kanjiQuizTouchScrollLockActive() {
      return Date.now() < kanjiQuizScrollLockUntil;
    }
    function markKanjiQuizTouchDrawActivity(lockMs) {
      const dur = Math.max(0, Number(lockMs) || 0);
      kanjiQuizScrollLockUntil = Date.now() + dur;
    }
    function bindKanjiQuizScrollGuard() {
      if (__kanjiQuizScrollGuardBound) return;
      __kanjiQuizScrollGuardBound = true;
      const stopIfLocked = function (e) {
        if (!kanjiQuizTouchScrollLockActive()) return;
        const sec = document.getElementById("section-kanji-quiz-play");
        if (!sec || !sec.classList.contains("active")) return;
        if (e && typeof e.preventDefault === "function") e.preventDefault();
      };
      document.addEventListener("touchmove", stopIfLocked, { passive: false });
      document.addEventListener("wheel", stopIfLocked, { passive: false });
    }
    function ensureKanjiHwFrameReadyOnce() {
      if (!__kanjiQuizHwFrameReadyP) {
        __kanjiQuizHwFrameReadyP = ensureKanjiPracticeFrameReady();
      }
      return __kanjiQuizHwFrameReadyP;
    }
    function abandonKanjiQuizPlayIfLeavingSection(newSectionId) {
      const quizPlayEl = document.getElementById("section-kanji-quiz-play");
      const wasPlay = quizPlayEl && quizPlayEl.classList.contains("active");
      if (!wasPlay || newSectionId === "section-kanji-quiz-play" || !kanjiQuizSession) return;
      kanjiQuizSession = null;
      lastKanjiQuizContext = null;
      clearKanjiHwFrameReadyCache();
      try {
        kanjiQuizClearWritePad(true);
      } catch (e) {}
      restoreKanjiPracticeFrameIfMoved();
      resetKanjiQuizDrillPlayShell();
    }
    function parseMaskedForDrill(s) {
      const str = String(s || "");
      const idx = str.search(/[＿_〼]/);
      if (idx < 0) return { before: str, after: "" };
      let rest = str.slice(idx + 1);
      rest = rest.replace(/^[＿_〼]+/, "");
      return { before: str.slice(0, idx), after: rest };
    }
    function kanjiQuizCircledQuestionNum(n) {
      const i = Math.max(1, Number(n) || 1);
      if (i <= 20) return String.fromCodePoint(0x245f + i);
      return "(" + i + ")";
    }
    function kanjiQuizCanvasXY(e, canvas) {
      const r = canvas.getBoundingClientRect();
      const L = kanjiQuizWriteLogicalSize;
      return { x: (e.clientX - r.left) * (L / Math.max(1, r.width)), y: (e.clientY - r.top) * (L / Math.max(1, r.height)) };
    }
    function getKanjiQuizStrokeParams() {
      const DEFAULT_PARAMS = {
        baseWidth: 11.0,
        tomeStart: 0.9,
        tomeScale: 1.3,
        haneStart: 0.9,
        haneSharp: 0.55,
        haraiStart: 0.65,
        haraiSharp: 0.8
      };
      try {
        const saved = localStorage.getItem("kanjiStrokeParamsV2");
        if (saved) {
          const o = JSON.parse(saved);
          return Object.assign({}, DEFAULT_PARAMS, o);
        }
      } catch (e) {}
      return Object.assign({}, DEFAULT_PARAMS);
    }
    function kanjiQuizEffectiveBaseWidth() {
      const p = getKanjiQuizStrokeParams();
      const bw = typeof p.baseWidth === "number" ? p.baseWidth : 11;
      return Math.max(4, Math.min(18, bw));
    }
    function kanjiQuizFillBetween(ctx, start, end, width) {
      const dist = Math.hypot(end.x - start.x, end.y - start.y);
      const steps = Math.max(1, Math.ceil(dist * 2.5));
      for (let i = 0; i <= steps; i++) {
        const t = i / steps;
        ctx.beginPath();
        ctx.arc(start.x + (end.x - start.x) * t, start.y + (end.y - start.y) * t, width / 2, 0, Math.PI * 2);
        ctx.fill();
      }
    }
    function kanjiQuizRenderStrokeWithEffect(ctx, points, type, color) {
      const P = getKanjiQuizStrokeParams();
      const bW = typeof P.baseWidth === "number" ? Math.max(4, Math.min(18, P.baseWidth)) : 11;
      const strokeType = type === "none" ? "tome" : type;
      ctx.fillStyle = color || "#333";
      if (!points || !points.length) return;
      if (points.length === 1) {
        ctx.beginPath();
        ctx.arc(points[0].x, points[0].y, Math.max(0.5, bW / 2), 0, Math.PI * 2);
        ctx.fill();
        return;
      }
      const n = points.length;
      for (let i = 0; i < n - 1; i++) {
        const ratio = i / (n - 1);
        let w = bW;
        if (strokeType === "harai" && ratio > P.haraiStart) {
          const denom = 1 - P.haraiStart || 0.001;
          w *= Math.max(0.1, 1 - P.haraiSharp * ((ratio - P.haraiStart) / denom));
        } else if (strokeType === "hane" && ratio > P.haneStart) {
          const denom = 1 - P.haneStart || 0.001;
          w *= Math.max(0.1, 1 - P.haneSharp * ((ratio - P.haneStart) / denom));
        } else if (strokeType === "tome" && ratio > P.tomeStart) {
          const denom = 1 - P.tomeStart || 0.001;
          w *= 1 + (P.tomeScale - 1) * ((ratio - P.tomeStart) / denom);
        }
        kanjiQuizFillBetween(ctx, points[i], points[i + 1], Math.max(0.5, w));
      }
    }
    function kanjiQuizDrawLiveStrokeSegments(ctx, points) {
      const w = kanjiQuizEffectiveBaseWidth();
      ctx.fillStyle = "#333";
      if (!points || !points.length) return;
      if (points.length === 1) {
        ctx.beginPath();
        ctx.arc(points[0].x, points[0].y, Math.max(0.5, w / 2), 0, Math.PI * 2);
        ctx.fill();
        return;
      }
      for (let i = 0; i < points.length - 1; i++) {
        kanjiQuizFillBetween(ctx, points[i], points[i + 1], w);
      }
    }
    function kanjiQuizWebKitCanvasPaintNudge_(canvas) {
      if (!canvas) return;
      try {
        canvas.style.transform = "translateZ(0)";
        canvas.style.webkitTransform = "translateZ(0)";
        void canvas.offsetWidth;
      } catch (e) {}
    }
    function kanjiQuizDrillHandwritingDisplayed() {
      const el = document.getElementById("kanji-quiz-drill-handwriting");
      if (!el) return false;
      try {
        return window.getComputedStyle(el).display !== "none";
      } catch (e) {
        return el.style.display !== "none";
      }
    }
    function kanjiQuizRedrawParentCanvas(livePts) {
      const canvas = document.getElementById("kanji-quiz-write-canvas");
      if (!canvas || !canvas.getContext) return;
      const L = kanjiQuizWriteLogicalSize;
      const dpr = window.devicePixelRatio || 1;
      var ctx;
      try {
        ctx = canvas.getContext("2d", { alpha: true });
      } catch (eCtx) {
        ctx = canvas.getContext("2d");
      }
      if (!ctx) return;
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.scale(dpr, dpr);
      /* 背景のグリッドは CSS（漢字練習 iframe と同じ）に任せ、白矩形は塗らない */
      kanjiQuizParentStrokes.forEach(function (s) {
        if (!s.points || !s.points.length) return;
        kanjiQuizRenderStrokeWithEffect(ctx, s.points, s.type, "#333");
      });
      if (livePts && livePts.length) {
        kanjiQuizDrawLiveStrokeSegments(ctx, livePts);
      }
      kanjiQuizWebKitCanvasPaintNudge_(canvas);
    }
    function kanjiQuizSetupWriteCanvas() {
      const canvas = document.getElementById("kanji-quiz-write-canvas");
      if (!canvas) return;
      const L = kanjiQuizWriteLogicalSize;
      const dpr = window.devicePixelRatio || 1;
      canvas.width = L * dpr;
      canvas.height = L * dpr;
      canvas.style.width = L + "px";
      canvas.style.height = L + "px";
      if (!__kanjiQuizWriteCanvasBound) {
        __kanjiQuizWriteCanvasBound = true;
        bindKanjiQuizScrollGuard();
        function canStartKanjiQuizDraw() {
          if (!kanjiQuizSession || !document.getElementById("section-kanji-quiz-play") || !document.getElementById("section-kanji-quiz-play").classList.contains("active")) return;
          const q = kanjiQuizSession.questions[kanjiQuizSession.index];
          if (!q || q.type !== "ruby_to_kanji") return;
          return true;
        }
        function beginStrokeAt(clientX, clientY) {
          const r = canvas.getBoundingClientRect();
          const px = (clientX - r.left) * (kanjiQuizWriteLogicalSize / Math.max(1, r.width));
          const py = (clientY - r.top) * (kanjiQuizWriteLogicalSize / Math.max(1, r.height));
          kanjiQuizIsDrawing = true;
          kanjiQuizStrokeStartTime = Date.now();
          __kanjiQuizTouchLockLastBump = Date.now();
          markKanjiQuizTouchDrawActivity(700);
          kanjiQuizCurrentStrokePoints = [];
          kanjiQuizCurrentStrokePoints.push({ x: px, y: py });
          kanjiQuizRedrawParentCanvas(kanjiQuizCurrentStrokePoints);
        }
        function moveStrokeAt(clientX, clientY) {
          if (!kanjiQuizIsDrawing) return;
          const r = canvas.getBoundingClientRect();
          const px = (clientX - r.left) * (kanjiQuizWriteLogicalSize / Math.max(1, r.width));
          const py = (clientY - r.top) * (kanjiQuizWriteLogicalSize / Math.max(1, r.height));
          kanjiQuizCurrentStrokePoints.push({ x: px, y: py });
          const now = Date.now();
          if (now - __kanjiQuizTouchLockLastBump >= 260) {
            __kanjiQuizTouchLockLastBump = now;
            markKanjiQuizTouchDrawActivity(700);
          }
          kanjiQuizRedrawParentCanvas(kanjiQuizCurrentStrokePoints);
        }
        function endStroke() {
          if (!kanjiQuizIsDrawing) return;
          kanjiQuizIsDrawing = false;
          markKanjiQuizTouchDrawActivity(420);
          const pts = kanjiQuizCurrentStrokePoints.map(function (p) {
            return { x: p.x, y: p.y };
          });
          if (pts.length) {
            const typ = kanjiQuizClassifyStroke(pts, kanjiQuizStrokeStartTime);
            kanjiQuizParentStrokes.push({ type: typ, points: pts });
          }
          kanjiQuizCurrentStrokePoints = [];
          kanjiQuizRedrawParentCanvas();
        }
        canvas.addEventListener("pointerdown", function (e) {
          if (!canStartKanjiQuizDraw()) return;
          e.preventDefault();
          canvas.setPointerCapture(e.pointerId);
          beginStrokeAt(e.clientX, e.clientY);
        });
        canvas.addEventListener("pointermove", function (e) {
          if (kanjiQuizIsDrawing) e.preventDefault();
          moveStrokeAt(e.clientX, e.clientY);
        });
        canvas.addEventListener("pointerup", function (e) {
          if (kanjiQuizIsDrawing) e.preventDefault();
          endStroke();
        });
        canvas.addEventListener("pointercancel", function (e) {
          if (kanjiQuizIsDrawing) e.preventDefault();
          kanjiQuizIsDrawing = false;
          kanjiQuizCurrentStrokePoints = [];
          kanjiQuizRedrawParentCanvas();
        });
        canvas.addEventListener("mousedown", function (e) {
          if (!canStartKanjiQuizDraw()) return;
          e.preventDefault();
          beginStrokeAt(e.clientX, e.clientY);
        });
        window.addEventListener("mousemove", function (e) {
          if (!kanjiQuizIsDrawing) return;
          moveStrokeAt(e.clientX, e.clientY);
        });
        window.addEventListener("mouseup", function () {
          endStroke();
        });
        canvas.addEventListener("touchstart", function (e) {
          if (!canStartKanjiQuizDraw()) return;
          if (!e.touches || !e.touches.length) return;
          e.preventDefault();
          beginStrokeAt(e.touches[0].clientX, e.touches[0].clientY);
        }, { passive: false });
        canvas.addEventListener("touchmove", function (e) {
          if (!kanjiQuizIsDrawing || !e.touches || !e.touches.length) return;
          e.preventDefault();
          moveStrokeAt(e.touches[0].clientX, e.touches[0].clientY);
        }, { passive: false });
        canvas.addEventListener("touchend", function () {
          endStroke();
        });
      }
      bindKanjiQuizWriteCanvasReflowListeners();
      kanjiQuizRedrawParentCanvas();
    }
    /** レイアウト確定後にキャンバス解像度・スケールを取り直す（回転・リサイズと同じ経路） */
    function kanjiQuizReflowWriteCanvasForCurrentLayout() {
      const sec = document.getElementById("section-kanji-quiz-play");
      if (!sec || !sec.classList.contains("active")) return;
      if (!kanjiQuizSession) return;
      const q = kanjiQuizSession.questions[kanjiQuizSession.index];
      if (!q || q.type !== "ruby_to_kanji") return;
      if (!kanjiQuizDrillHandwritingDisplayed()) return;
      kanjiQuizSetupWriteCanvas();
    }
    function kanjiQuizScheduleWriteCanvasReflow() {
      requestAnimationFrame(function () {
        requestAnimationFrame(function () {
          kanjiQuizReflowWriteCanvasForCurrentLayout();
        });
      });
    }
    /** 手書きクイズ：iframe の高さ・親キャンバスを取り直す（表示ずれ対策の手動トリガー） */
    function kanjiQuizRefreshHandwritingLayout() {
      try {
        window.scrollBy(0, 0);
      } catch (eScroll) {}
      try {
        kanjiQuizReflowWriteCanvasForCurrentLayout();
      } catch (e0) {}
      try {
        kpResizeFrameToContent();
      } catch (e1) {}
      try {
        kanjiQuizScheduleWriteCanvasReflow();
      } catch (e2) {}
      try {
        var c = document.getElementById("kanji-quiz-write-canvas");
        kanjiQuizWebKitCanvasPaintNudge_(c);
      } catch (eNudge) {}
      setTimeout(function () {
        try {
          kpResizeFrameToContent();
        } catch (e3) {}
      }, 120);
      setTimeout(function () {
        try {
          kanjiQuizReflowWriteCanvasForCurrentLayout();
        } catch (e4) {}
      }, 240);
      setTimeout(function () {
        try {
          var cv = document.getElementById("kanji-quiz-write-canvas");
          kanjiQuizWebKitCanvasPaintNudge_(cv);
        } catch (e5) {}
      }, 400);
    }
    function bindKanjiQuizWriteCanvasReflowListeners() {
      if (__kanjiQuizWriteReflowListenersBound) return;
      __kanjiQuizWriteReflowListenersBound = true;
      let debounceTimer = 0;
      function onLayoutChange() {
        if (debounceTimer) clearTimeout(debounceTimer);
        debounceTimer = setTimeout(function () {
          debounceTimer = 0;
          kanjiQuizReflowWriteCanvasForCurrentLayout();
        }, 100);
      }
      window.addEventListener("resize", onLayoutChange);
      window.addEventListener("orientationchange", onLayoutChange);
      try {
        if (window.visualViewport) {
          window.visualViewport.addEventListener("resize", onLayoutChange);
          window.visualViewport.addEventListener("scroll", onLayoutChange);
        }
      } catch (eVv) {}
    }
    function getKanjiQuizStrokeThresholds() {
      const def = { minVelocity: 1.2, hookAngleDiff: 0.6 };
      try {
        const raw = localStorage.getItem("kanjiPenProfiles");
        if (!raw) return def;
        const parsed = JSON.parse(raw);
        const id = parsed.activeId !== undefined ? parsed.activeId : 0;
        const prof = (parsed.profiles || [])[id];
        if (prof && prof.thresholds && typeof prof.thresholds === "object") {
          return {
            minVelocity: Number(prof.thresholds.minVelocity) || def.minVelocity,
            hookAngleDiff: Number(prof.thresholds.hookAngleDiff) || def.hookAngleDiff
          };
        }
      } catch (e) {}
      return def;
    }
    /* iframe メインキャンバス: pointerup で 3 点未満は即「とめ」で積む。detectStrokeType(..., true, terminalInfo) は 5 点未満を「とめ」固定。親も同じ順序・同じ式で揃える。 */
    const KANJI_QUIZ_MIN_STROKE_POINTS = 5;
    function kanjiQuizClassifyStroke(points, drawStartTime) {
      if (!points || points.length < 3) return "tome";
      if (points.length < KANJI_QUIZ_MIN_STROKE_POINTS) return "tome";
      const n = points.length;
      const pStart = points[Math.max(0, n - 6)], pEnd = points[n - 1];
      const tStart = Number.isFinite(drawStartTime) ? drawStartTime : Date.now();
      const velocity =
        Math.hypot(pEnd.x - pStart.x, pEnd.y - pStart.y) / ((Date.now() - tStart) || 1) * 10;
      const dy = pEnd.y - pStart.y;
      const pMid = points[Math.floor(n * 0.5)], pPre = points[Math.floor(n * 0.85)];
      const angleMain = Math.atan2(pPre.y - pMid.y, pPre.x - pMid.x);
      const angleTip = Math.atan2(pEnd.y - pPre.y, pEnd.x - pPre.x);
      let angleDiff = Math.abs(angleTip - angleMain);
      if (angleDiff > Math.PI) angleDiff = 2 * Math.PI - angleDiff;
      const thr = getKanjiQuizStrokeThresholds();
      if (velocity < thr.minVelocity) return "tome";
      if (dy < 0 || angleDiff > thr.hookAngleDiff) return "hane";
      return "harai";
    }
    function kanjiQuizMergeKanjiStrokeParams(partial) {
      const def = {
        baseWidth: 11.0,
        tomeStart: 0.9,
        tomeScale: 1.3,
        haneStart: 0.9,
        haneSharp: 0.55,
        haraiStart: 0.65,
        haraiSharp: 0.8
      };
      let cur = { ...def };
      try {
        const saved = localStorage.getItem("kanjiStrokeParamsV2");
        if (saved) cur = { ...def, ...JSON.parse(saved) };
      } catch (e) {}
      cur = { ...cur, ...partial };
      try {
        localStorage.setItem("kanjiStrokeParamsV2", JSON.stringify(cur));
      } catch (e) {}
      return cur;
    }
    /** 漢字練習の「🖋 設定」と共有する baseWidth（中間 11 を基準に 4/8/14 でスケール） */
    function kanjiQuizSetPenWidth(w) {
      kanjiQuizPenW = w;
      const midRef = 11.0;
      kanjiQuizMergeKanjiStrokeParams({
        baseWidth: Math.max(4, Math.min(18, midRef * (w / 8)))
      });
      [["kanji-pen-fine", 4], ["kanji-pen-mid", 8], ["kanji-pen-thick", 14]].forEach(function (t) {
        const el = document.getElementById(t[0]);
        if (el) el.classList.toggle("is-active-prior", t[1] === w);
      });
      kanjiQuizRedrawParentCanvas();
    }
    /** 保存済み kanjiStrokeParamsV2 に合わせてボタン表示のみ同期（設定を上書きしない） */
    function kanjiQuizSyncPenUiFromStrokeParams() {
      const midRef = 11.0;
      let baseW = midRef;
      try {
        const saved = localStorage.getItem("kanjiStrokeParamsV2");
        if (saved) {
          const o = JSON.parse(saved);
          if (typeof o.baseWidth === "number") baseW = o.baseWidth;
        }
      } catch (e) {}
      const presets = [4, 8, 14];
      let bestI = 1;
      let bestD = Infinity;
      for (let i = 0; i < presets.length; i++) {
        const expected = midRef * (presets[i] / 8);
        const d = Math.abs(baseW - expected);
        if (d < bestD) {
          bestD = d;
          bestI = i;
        }
      }
      kanjiQuizPenW = presets[bestI];
      [["kanji-pen-fine", 4], ["kanji-pen-mid", 8], ["kanji-pen-thick", 14]].forEach(function (t) {
        const el = document.getElementById(t[0]);
        if (el) el.classList.toggle("is-active-prior", t[1] === kanjiQuizPenW);
      });
    }
    function kanjiQuizClearWritePad(clearIframe) {
      kanjiQuizParentStrokes = [];
      kanjiQuizCurrentStrokePoints = [];
      kanjiQuizIsDrawing = false;
      kanjiQuizRedrawParentCanvas();
      if (clearIframe) {
        try {
          const frame = document.getElementById("kp-pro-frame");
          if (frame && frame.contentWindow) frame.contentWindow.postMessage({ type: "quizClearDrawing" }, "*");
        } catch (e) {}
      }
    }
    function kanjiQuizRemoveLastStrokeDecoration() {
      if (!kanjiQuizParentStrokes.length) return;
      kanjiQuizParentStrokes[kanjiQuizParentStrokes.length - 1].type = "none";
      kanjiQuizRedrawParentCanvas();
    }
    let kanjiQuizHandSubmitBusy = false;
    let kanjiQuizHandSubmitAnimTimer = null;
    function setKanjiQuizHandSubmitBusy(isBusy) {
      const btn = document.getElementById("kanji-hw-submit-btn");
      if (!btn) return;
      kanjiQuizHandSubmitBusy = !!isBusy;
      btn.disabled = !!isBusy;
      btn.classList.toggle("is-loading", !!isBusy);
      btn.setAttribute("aria-busy", isBusy ? "true" : "false");
      if (kanjiQuizHandSubmitAnimTimer) {
        clearInterval(kanjiQuizHandSubmitAnimTimer);
        kanjiQuizHandSubmitAnimTimer = null;
      }
      if (isBusy) {
        let tick = 0;
        btn.textContent = "さいてん中.";
        kanjiQuizHandSubmitAnimTimer = setInterval(function () {
          tick = (tick + 1) % 3;
          btn.textContent = "さいてん中" + ".".repeat(tick + 1);
        }, 280);
      } else {
        btn.textContent = "これでかいとう";
      }
    }
    function kanjiQuizRunHandwritingAnswer() {
      if (kanjiQuizHandSubmitBusy) return;
      if (!kanjiQuizParentStrokes.length) {
        alert("かんじを かいてください。");
        return;
      }
      if (!kanjiQuizSession) return;
      setKanjiQuizHandSubmitBusy(true);
      function postEvalToFrame() {
        return (function () {
          const frame = document.getElementById("kp-pro-frame");
          if (!frame || !frame.contentWindow) return false;
          const targets = kanjiQuizSession ? kanjiQuizSession.rubyHandTargets || [] : [];
          const slot = kanjiQuizSession ? kanjiQuizSession.rubyHandSlot || 0 : 0;
          const q = kanjiQuizSession ? kanjiQuizSession.questions[kanjiQuizSession.index] : null;
          const fallbackTargets = kanjiQuizHanOnlyChars((q && (q.correctAnswer || q.kanji)) || "");
          const qKanjiTargets = kanjiQuizHanOnlyChars((q && q.kanji) || "");
          const expectedChar =
            targets[slot] ||
            qKanjiTargets[slot] ||
            qKanjiTargets[0] ||
            fallbackTargets[slot] ||
            fallbackTargets[0] ||
            "";
          if (expectedChar && (!targets.length || targets[slot] !== expectedChar)) {
            kanjiQuizSession.rubyHandTargets = qKanjiTargets.length
              ? qKanjiTargets
              : fallbackTargets.length
                ? fallbackTargets
                : [expectedChar];
            kanjiQuizSession.rubyHandSlot = Math.max(0, Math.min(slot, kanjiQuizSession.rubyHandTargets.length - 1));
          }
          if (expectedChar) {
            selectKanjiCharInQuizFrame(expectedChar);
          }
          const payload = {
            type: "quizEvalParentStrokes",
            expectedChar: expectedChar,
            strokes: kanjiQuizParentStrokes.map(function (s) {
              return { type: s.type, points: s.points.map(function (p) { return { x: +p.x, y: +p.y }; }) };
            })
          };
          setTimeout(function () {
            try {
              frame.contentWindow.postMessage(payload, "*");
            } catch (e) {}
          }, 90);
          return true;
        })();
      }
      ensureKanjiHwFrameReadyOnce().then(function () {
        const frame = document.getElementById("kp-pro-frame");
        if (frame) patchKanjiFrameForQuizPostMessage(frame);
        ensureKanjiFrameForQuizEval();
        if (!postEvalToFrame()) {
          setKanjiQuizHandSubmitBusy(false);
          alert("さいてんようのびゅうが みつかりません。");
        }
      }).catch(function () {
        setKanjiQuizHandSubmitBusy(false);
        alert("さいてんのじゅんびに しっぱいしました。もういちどためしてください。");
      });
    }
    function kanjiQuizOnHandwritingScored(sc) {
      setKanjiQuizHandSubmitBusy(false);
      if (!kanjiQuizSession) return;
      const secHand = document.getElementById("section-kanji-quiz-play");
      if (!secHand || !secHand.classList.contains("active")) return;
      const q = kanjiQuizSession.questions[kanjiQuizSession.index];
      if (!q || q.type !== "ruby_to_kanji") return;
      if (sc == null || sc < KANJI_QUIZ_HAND_PASS) {
        kanjiQuizShowHandwritingWrongFeedback(sc);
        return;
      }
      kanjiQuizHideWrongFeedback();
      const scNum = Number(sc);
      if (!isNaN(scNum)) {
        const pm = kanjiQuizSession.rubyHandMinScore;
        kanjiQuizSession.rubyHandMinScore = pm == null || isNaN(pm) ? scNum : Math.min(pm, scNum);
      }
      const targets = kanjiQuizSession.rubyHandTargets || [];
      const slot = kanjiQuizSession.rubyHandSlot || 0;
      if (slot + 1 >= targets.length) {
        kanjiQuizSession.rubyHandComplete = true;
        kanjiQuizSession.rubyHandKanjiVgPass = true;
        submitKanjiQuizScore();
        return;
      }
      kanjiQuizSession.rubyHandSlot = slot + 1;
      kanjiQuizSession.lastHandScore = null;
      const nextCh = targets[slot + 1];
      if (!selectKanjiCharInQuizFrame(nextCh)) return;
      const sum = document.getElementById("kanji-play-summary");
      if (sum) sum.innerHTML = "";
      kanjiQuizClearWritePad(true);
      kanjiQuizSetupWriteCanvas();
      kanjiQuizScheduleWriteCanvasReflow();
    }
    function kanjiQuizSkipHandwritingQuestion() {
      if (!kanjiQuizSession) return;
      const secHand = document.getElementById("section-kanji-quiz-play");
      if (!secHand || !secHand.classList.contains("active")) return;
      const q = kanjiQuizSession.questions[kanjiQuizSession.index];
      if (!q || q.type !== "ruby_to_kanji") return;

      // 不正解として確定し、保存して次の問題へ進む。
      kanjiQuizSession.rubyHandComplete = true;
      kanjiQuizSession.rubyHandKanjiVgPass = false;
      kanjiQuizSession.lastHandScore = 0;
      kanjiQuizSession.rubyHandMinScore = 0;
      kanjiQuizHideWrongFeedback();
      submitKanjiQuizScore();
    }
    function isKanjiQuizHanChar(ch) {
      if (!ch || ch.length === 0) return false;
      const cp = ch.codePointAt(0);
      if (cp >= 0x4e00 && cp <= 0x9fff) return true;
      if (cp >= 0x3400 && cp <= 0x4dbf) return true;
      if (cp >= 0xf900 && cp <= 0xfaff) return true;
      return false;
    }
    function kanjiQuizHanOnlyChars(str) {
      return Array.from(String(str || "")).filter(isKanjiQuizHanChar);
    }
    function patchKanjiFrameForQuizPostMessage(frame) {
      try {
        const win = frame.contentWindow;
        if (!win) return;
        const doc = frame.contentDocument;
        if (!doc || !doc.documentElement) return;
        if (doc.documentElement.dataset.kanjiQuizParentHook === "1") return;
        const s = doc.createElement("script");
        s.textContent =
          "(function(){" +
          "if(window.__kjQPatchInner)return;" +
          "window.__kjQPatchInner=true;" +
          "window.addEventListener(\"message\",function(ev){" +
          "if(!ev||!ev.data)return;var d=ev.data;" +
          "try{" +
          "if(d.type===\"quizEvalParentStrokes\"&&Array.isArray(d.strokes)){" +
          "if(d.expectedChar){" +
          "var _ksel=document.getElementById(\"target-kanji\");" +
          "if(!_ksel||String(_ksel.value)!==String(d.expectedChar)){" +
          "if(window.parent)window.parent.postMessage({type:\"kanjiQuizScored\",score:0},\"*\");return;" +
          "}" +
          "}" +
          "userStrokes=d.strokes.map(function(s){return{type:(s&&s.type)||\"tome\",points:Array.isArray(s&&s.points)?s.points.map(function(p){return{x:+p.x,y:+p.y}}):[]};});" +
          "if(typeof redrawAllUserStrokes===\"function\")redrawAllUserStrokes();" +
          "if(typeof evaluateKanji===\"function\")evaluateKanji(false);" +
          "}" +
          "if(d.type===\"quizClearDrawing\"&&typeof clearCanvas===\"function\")clearCanvas();" +
          "if(d.type===\"quizRemoveLastDecoration\"&&typeof removeLastEffect===\"function\")removeLastEffect();" +
          "if(d.type===\"quizPlayStrokeOrderDemo\"){" +
          "if(typeof switchMode===\"function\")switchMode(\"demo\");" +
          "setTimeout(function(){if(typeof playAnimation===\"function\")playAnimation();},450);" +
          "}" +
          "}catch(e){}" +
          "});" +
          "function _kjWrapEval(){" +
          "if(typeof window.evaluateKanji!==\"function\"){requestAnimationFrame(_kjWrapEval);return;}" +
          "if(window.__kjEvalWrapped)return;" +
          "window.__kjEvalWrapped=true;" +
          "var _o=window.evaluateKanji;" +
          "window.evaluateKanji=function(t){" +
          "_o.apply(this,arguments);" +
          "setTimeout(function(){try{var e=document.getElementById(\"score\"),m=e&&e.innerText&&e.innerText.match(/(\\d+)/),n=m?parseInt(m[1],10):0;if(isNaN(n))n=0;var s=document.getElementById(\"target-kanji\"),k=s&&s.value?String(s.value):\"\";if(window.parent)window.parent.postMessage({type:\"kanjiQuizScored\",score:n,kanjiChar:k},\"*\");}catch(x){}},120);" +
          "};" +
          "}" +
          "_kjWrapEval();" +
          "})();";
        doc.documentElement.appendChild(s);
        doc.documentElement.dataset.kanjiQuizParentHook = "1";
      } catch (e) {}
    }
    function ensureKanjiPracticeFrameReady() {
      return new Promise(function (resolve) {
        Promise.resolve(openKanjiPracticePro())
          .catch(function () {})
          .then(function () {
        kanjiQuizEnsureScoreListener();
        const frame = document.getElementById("kp-pro-frame");
        if (!frame) {
          resolve();
          return;
        }
        let settled = false;
        let pollId = null;

        function cleanup() {
          window.removeEventListener("message", onKpKanjiReady);
          window.removeEventListener("message", onKpIframeBoot);
          if (pollId != null) {
            clearInterval(pollId);
            pollId = null;
          }
        }

        function finish() {
          if (settled) return;
          settled = true;
          cleanup();
          resolve();
        }

        function onKpKanjiReady(ev) {
          if (!ev.data || ev.data.type !== "kpKanjiDataReady") return;
          if (ev.source !== frame.contentWindow) return;
          finish();
        }

        function onKpIframeBoot(ev) {
          if (!ev.data || ev.data.type !== "kpIframeBootReady") return;
          if (ev.source !== frame.contentWindow) return;
          finish();
        }

        window.addEventListener("message", onKpKanjiReady);
        window.addEventListener("message", onKpIframeBoot);

        function tryPoll() {
          patchKanjiFrameForQuizPostMessage(frame);
          const win = frame.contentWindow;
          if (!win) {
            finish();
            return;
          }

          function check() {
            patchKanjiFrameForQuizPostMessage(frame);
            if (typeof win.applyQuizKanjiData === "function") return true;
            const kd = win.KANJI_DATA;
            return !!(kd && typeof kd === "object" && Object.keys(kd).length > 0);
          }

          if (check()) {
            finish();
            return;
          }

          let n = 0;
          pollId = setInterval(function () {
            n++;
            if (check()) {
              finish();
              return;
            }
            if (n >= 60) finish();
          }, 80);
        }

        if (frame.dataset.kpLoaded === "1" && frame.contentDocument && frame.contentDocument.readyState === "complete") {
          setTimeout(tryPoll, 0);
        } else {
          frame.addEventListener("load", function onLf() {
            frame.removeEventListener("load", onLf);
            setTimeout(tryPoll, 0);
          });
        }
        });
      });
    }
    function ensureKanjiFrameForQuizEval() {
      const hid = document.getElementById("kp-pro-frame-quiz-hidden");
      const frame = document.getElementById("kp-pro-frame");
      if (hid && frame && frame.parentElement !== hid) {
        hid.appendChild(frame);
      }
    }
    function restoreKanjiPracticeFrameIfMoved() {
      const slot = document.getElementById("kp-iframe-slot-practice");
      const frame = document.getElementById("kp-pro-frame");
      if (!frame || !slot) return;
      if (frame.parentElement !== slot) {
        slot.appendChild(frame);
      }
      setTimeout(kpResizeFrameToContent, 80);
    }
    var __kanjiHandScoreToastTimer = null;
    var __kanjiEarnedPtsToastTimer = null;
    /**
     * 今回の問題／採点で加算されたポイントを画面上部に約2秒表示。
     * pointer-events: none・固定レイヤのため操作や fetch の続行を阻害しない。
     */
    function showKanjiEarnedPointsToast(earnedRaw) {
      var earned = Number(earnedRaw);
      if (isNaN(earned)) earned = 0;
      var el = document.getElementById("kanji-earned-pts-toast");
      if (!el) {
        el = document.createElement("div");
        el.id = "kanji-earned-pts-toast";
        el.setAttribute("role", "status");
        el.setAttribute("aria-live", "polite");
        el.style.cssText =
          "position:fixed;left:50%;top:max(12px,env(safe-area-inset-top,12px));transform:translateX(-50%);" +
          "z-index:10051;padding:10px 20px;border-radius:14px;background:rgba(20,35,25,0.92);color:#e8f5e9;" +
          "font-size:clamp(15px,3.8vw,19px);font-weight:800;box-shadow:0 4px 20px rgba(0,0,0,0.25);" +
          "pointer-events:none;opacity:0;transition:opacity 0.22s ease;text-align:center;line-height:1.35;" +
          "max-width:min(92vw,320px);border:1px solid rgba(129,199,132,0.45);";
        document.body.appendChild(el);
      }
      var showNum = (earned >= 0 ? "+" : "") + earned.toFixed(2);
      el.textContent = "ポイント " + showNum + " Pt";
      el.style.opacity = "1";
      if (__kanjiEarnedPtsToastTimer) clearTimeout(__kanjiEarnedPtsToastTimer);
      __kanjiEarnedPtsToastTimer = setTimeout(function () {
        el.style.opacity = "0";
        __kanjiEarnedPtsToastTimer = null;
      }, 2000);
    }
    /** 手書き採点の点数を画面下部に約2秒表示（pointer-events なし。他 UI を塞がない） */
    function showKanjiHandScoreToast(scoreRaw, kanjiOpt) {
      var n = Math.max(0, Math.min(100, Math.round(Number(scoreRaw) || 0)));
      var ch = kanjiOpt != null ? String(kanjiOpt).trim() : "";
      var el = document.getElementById("kanji-hand-score-toast");
      if (!el) {
        el = document.createElement("div");
        el.id = "kanji-hand-score-toast";
        el.setAttribute("role", "status");
        el.setAttribute("aria-live", "polite");
        el.style.cssText =
          "position:fixed;left:50%;bottom:max(20px,env(safe-area-inset-bottom,12px));transform:translateX(-50%);" +
          "z-index:10050;padding:10px 18px;border-radius:12px;background:rgba(30,30,30,0.92);color:#fff;" +
          "font-size:clamp(16px,4vw,20px);font-weight:700;box-shadow:0 4px 20px rgba(0,0,0,0.22);" +
          "pointer-events:none;opacity:0;transition:opacity 0.2s ease;text-align:center;line-height:1.35;" +
          "max-width:min(92vw,300px);";
        document.body.appendChild(el);
      }
      el.textContent = ch ? "「" + ch + "」 " + n + " てん" : n + " てん";
      el.style.opacity = "1";
      if (__kanjiHandScoreToastTimer) clearTimeout(__kanjiHandScoreToastTimer);
      __kanjiHandScoreToastTimer = setTimeout(function () {
        el.style.opacity = "0";
        __kanjiHandScoreToastTimer = null;
      }, 2000);
    }
    /** KP iframe の kanjiQuizScored を1リスナで処理（クイズ優先、その後のみ練習の GAS 保存） */
    function initKanjiParentKanjiQuizScoredBridge() {
      if (window.__kanjiParentKanjiQuizScoredBridgeBound) return;
      window.__kanjiParentKanjiQuizScoredBridgeBound = true;
      window.addEventListener("message", function (ev) {
        if (!ev || !ev.data || ev.data.type !== "kanjiQuizScored") return;

        var quizSec = document.getElementById("section-kanji-quiz-play");
        if (quizSec && quizSec.classList.contains("active") && kanjiQuizSession) {
          kanjiQuizSession.lastHandScore = ev.data.score;
          var qNow = kanjiQuizSession.questions[kanjiQuizSession.index];
          var kToast = String(ev.data.kanjiChar || (qNow && qNow.kanji) || "").trim();
          showKanjiHandScoreToast(ev.data.score, kToast);
          var quizSum = document.getElementById("kanji-play-summary");
          if (quizSum) {
            quizSum.innerHTML =
              "<span style=\"color:#b8860b;font-weight:700;\">さいてん: " + ev.data.score + " てん</span>";
          }
          kanjiQuizOnHandwritingScored(ev.data.score);
          return;
        }

        var pracSec = document.getElementById("section-kanji-practice");
        if (!pracSec || !pracSec.classList.contains("active")) return;

        var score = Math.max(0, Math.min(100, Number(ev.data.score) || 0));
        var kanjiChar = String(ev.data.kanjiChar || "");
        showKanjiHandScoreToast(score, kanjiChar);
        if (!kanjiChar) return;

        var nowMs = Date.now();
        var dedupeKey = kanjiChar + ":" + score;
        if (dedupeKey === __kanjiPracticeLastSubmitKey && nowMs - __kanjiPracticeLastSubmitAt < 950) return;
        __kanjiPracticeLastSubmitKey = dedupeKey;
        __kanjiPracticeLastSubmitAt = nowMs;

        var kidUser = JSON.parse(localStorage.getItem("app_kid_user") || "null");
        if (!kidUser || !kidUser.id) return;

        var bEl = document.getElementById("kp-book-select");
        var uEl = document.getElementById("kp-sheet-select");
        var sEl = document.getElementById("kp-set-select");
        var mid = (bEl && bEl.value) || "KANJI_PRACTICE";
        var unm = (uEl && uEl.value) || "KANJI_PRACTICE";
        var sid = (sEl && sEl.value) || "PRACTICE";
        var unitScope = "KP_" + mid + "_" + unm + "_" + sid;

        fetch(GAS_API_URL, {
          method: "POST",
          body: JSON.stringify({
            action: "save_learning_session",
            userId: kidUser.id,
            unitId: unitScope + "_" + kanjiChar + "_" + nowMs,
            unitSheetName: unm,
            isReviewMode: false,
            isRandom: true,
            results: [],
            learningCategory: "kanji",
            challengeType: "score",
            kanjiChar: kanjiChar,
            score: score,
            questionId: unitScope + "_" + kanjiChar,
            questionCorrect: score >= KANJI_QUIZ_HAND_PASS,
            kanjiSetScopeId: unitScope,
            kanjiSetContinuation: true
          })
        })
          .then(function (r) {
            return r.json();
          })
          .then(function (d) {
            if (!d || d.status !== "success") return;
            try {
              showKanjiEarnedPointsToast(d.earnedPoints);
            } catch (ePt) {}
            if (typeof d.newTotal === "number") {
              kidUser.points = d.newTotal;
              localStorage.setItem("app_kid_user", JSON.stringify(kidUser));
              var pts = document.getElementById("user-points");
              if (pts) pts.innerText = String(d.newTotal);
            }
          })
          .catch(function () {});
      });
    }
    function kanjiQuizEnsureScoreListener() {
      initKanjiParentKanjiQuizScoredBridge();
      initKanjiHandAnalyticsBridge();
      initKanjiQuizWrongModelKpReselectBridge();
    }
    /** KanjiVG 再読込で target-kanji が先頭に戻る場合に、不正解モデル用の字を取り直す */
    function initKanjiQuizWrongModelKpReselectBridge() {
      if (window.__kanjiQuizWrongModelKpReselectBound) return;
      window.__kanjiQuizWrongModelKpReselectBound = true;
      window.addEventListener("message", function (ev) {
        if (!ev || !ev.data || ev.data.type !== "kpKanjiDataReady") return;
        const fr = document.getElementById("kp-pro-frame");
        if (!fr || ev.source !== fr.contentWindow) return;
        if (!kanjiQuizSession || !kanjiQuizSession.wrongModelKanjiChar) return;
        const panel = document.getElementById("kanji-quiz-hw-wrong-panel");
        if (!panel || panel.style.display === "none") return;
        selectKanjiCharInQuizFrame(kanjiQuizSession.wrongModelKanjiChar);
      });
    }
    function selectKanjiCharInQuizFrame(ch) {
      const frame = document.getElementById("kp-pro-frame");
      const win = frame && frame.contentWindow;
      if (!win) return false;
      try {
        const doc = win.document;
        const sel = doc.getElementById("target-kanji");
        if (!sel) return false;
        const target = String(ch || "");
        if (!target) return false;
        try {
          if (kanjiQuizSession && kanjiQuizSession.wrongModelKanjiChar === target) {
            win.__kpPendingKanjiSelect = target;
          }
        } catch (e0) {}
        let ok = Array.from(sel.options || []).some(function (o) { return String(o.value) === target; });
        if (!ok && win.KANJI_DATA && typeof win.KANJI_DATA === "object" && win.KANJI_DATA[target]) {
          const op = doc.createElement("option");
          op.value = target;
          op.textContent = target;
          sel.appendChild(op);
          ok = true;
        }
        if (!ok) {
          alert("「" + target + "」は KanjiVG.txt にないため、手書きできません。");
          return false;
        }
        sel.value = target;
        if (String(sel.value) !== target) {
          return false;
        }
        if (typeof win.initTargetKanji === "function") win.initTargetKanji();
        if (typeof win.switchMode === "function") win.switchMode("score");
        return true;
      } catch (e) {
        return false;
      }
    }
    /**
     * ①問題から確定した漢字 → ② KanjiVG.txt（TSV）を親が fetch し、その文字のストロークだけ iframe に渡す。
     */
    function applyKanjiQuizTargetsFromKanjiVg(targets, activeChar) {
      const chars = Array.isArray(targets) ? targets.filter(Boolean) : [];
      if (!chars.length) return Promise.reject(new Error("手書き対象がありません"));
      if (!window.KanjiVg || typeof KanjiVg.pathsForChar !== "function") {
        return Promise.reject(new Error("KanjiVg が読み込まれていません"));
      }
      const frame = document.getElementById("kp-pro-frame");
      const win = frame && frame.contentWindow;
      const url =
        typeof KanjiVg.resolveTxtUrl === "function"
          ? KanjiVg.resolveTxtUrl()
          : new URL("KanjiVG.txt", window.location.href).href;
      return fetch(url, { cache: "no-store" })
        .then(function (r) {
          if (!r.ok) throw new Error("KanjiVG.txt を取得できません (" + r.status + ")");
          return r.text();
        })
        .then(function (text) {
          const entries = {};
          for (let i = 0; i < chars.length; i++) {
            const ch = chars[i];
            const paths = KanjiVg.pathsForChar(text, ch);
            if (!paths || !paths.length) {
              throw new Error("「" + ch + "」は KanjiVG.txt にありません");
            }
            entries[ch] = paths;
          }
          if (!win || typeof win.applyQuizKanjiData !== "function") {
            throw new Error("採点パネルの準備ができていません");
          }
          win.applyQuizKanjiData(entries);
          const ac = activeChar || chars[0];
          if (!selectKanjiCharInQuizFrame(ac)) throw new Error("select");
          return true;
        });
    }
    function kanjiQuizHideWrongFeedback() {
      const panel = document.getElementById("kanji-quiz-hw-wrong-panel");
      if (panel) panel.style.display = "none";
      try {
        if (kanjiQuizSession) delete kanjiQuizSession.wrongModelKanjiChar;
        const frClear = document.getElementById("kp-pro-frame");
        const wClear = frClear && frClear.contentWindow;
        if (wClear) {
          try {
            delete wClear.__kpPendingKanjiSelect;
          } catch (eClr) {
            wClear.__kpPendingKanjiSelect = "";
          }
        }
      } catch (eH) {}
      const wrap = document.getElementById("kanji-quiz-wrong-model-wrap");
      const frame = document.getElementById("kp-pro-frame");
      const hid = document.getElementById("kp-pro-frame-quiz-hidden");
      if (frame && hid && wrap && frame.parentElement === wrap) {
        hid.appendChild(frame);
        frame.style.width = "";
        frame.style.maxWidth = "";
        frame.style.height = "";
        frame.style.minHeight = "";
      }
    }
    function kanjiQuizMountWrongModelFrameForChar(ch) {
      const modelWrap = document.getElementById("kanji-quiz-wrong-model-wrap");
      const frame = document.getElementById("kp-pro-frame");
      if (!modelWrap || !frame || !ch) return;
      function mountAndPlayDemo() {
        if (frame.parentElement !== modelWrap) {
          modelWrap.innerHTML = "";
          modelWrap.appendChild(frame);
        }
        frame.style.width = "100%";
        frame.style.maxWidth = "100%";
        frame.style.minHeight = "320px";
        frame.style.height = "360px";
        setTimeout(function () {
          try {
            kpResizeFrameToContent();
          } catch (e) {}
        }, 80);
        setTimeout(function () {
          try {
            kpResizeFrameToContent();
          } catch (e) {}
        }, 400);
        setTimeout(function () {
          try {
            if (frame.contentWindow) {
              frame.contentWindow.postMessage({ type: "quizPlayStrokeOrderDemo" }, "*");
            }
          } catch (e) {}
        }, 150);
        setTimeout(function () {
          try {
            kpResizeFrameToContent();
          } catch (e) {}
        }, 550);
      }
      function trySelectWithRetry() {
        let mountDone = false;
        function trySelectOnce() {
          if (mountDone) return;
          if (!selectKanjiCharInQuizFrame(ch)) return;
          mountDone = true;
          mountAndPlayDemo();
        }
        trySelectOnce();
        if (mountDone) return;
        // iframe側の候補再構築タイミングと競合するため、短い間隔で再試行する。
        setTimeout(trySelectOnce, 140);
        setTimeout(trySelectOnce, 360);
      }
      ensureKanjiHwFrameReadyOnce().then(trySelectWithRetry);
    }
    function kanjiQuizShowHandwritingWrongFeedback(sc) {
      const panel = document.getElementById("kanji-quiz-hw-wrong-panel");
      const redWrap = document.getElementById("kanji-quiz-wrong-red-chars");
      if (!kanjiQuizSession || !panel || !redWrap) return;
      const headEl = document.getElementById("kanji-quiz-hw-wrong-heading");
      if (headEl) {
        headEl.textContent =
          sc != null && !isNaN(Number(sc))
            ? "ざんねん（" + sc + "てん・60点みまん）"
            : "ざんねん（60点みまん）";
      }
      const targets = kanjiQuizSession.rubyHandTargets || [];
      const slot = kanjiQuizSession.rubyHandSlot || 0;
      const q = kanjiQuizSession.questions[kanjiQuizSession.index];
      const wordKanjiTargets = kanjiQuizHanOnlyChars((q && (q.correctAnswer || q.kanji)) || "");
      const ch =
        targets[slot] || wordKanjiTargets[slot] || wordKanjiTargets[0] || "";
      if (ch) {
        kanjiQuizSession.rubyHandTargets = wordKanjiTargets.length
          ? wordKanjiTargets
          : targets.length
            ? targets
            : [ch];
        kanjiQuizSession.rubyHandSlot = Math.max(
          0,
          Math.min(slot, kanjiQuizSession.rubyHandTargets.length - 1)
        );
      }
      if (ch) kanjiQuizSession.wrongModelKanjiChar = ch;
      else delete kanjiQuizSession.wrongModelKanjiChar;
      const fullStr = q && q.correctAnswer != null ? String(q.correctAnswer) : "";
      const fullKanji = kanjiQuizHanOnlyChars(fullStr).join("");
      const tlen = (kanjiQuizSession.rubyHandTargets || []).length;
      let html = "";
      if (ch) {
        html += "<span class=\"kanji-quiz-wrong-char-main\">" + escapeHtml(ch) + "</span>";
        if (tlen > 1) {
          html += "<span class=\"kanji-quiz-wrong-sub\">いまのマスにかくかんじ</span>";
        }
      }
      if (fullKanji && tlen > 1) {
        html += "<div class=\"kanji-quiz-wrong-full-line\">ぜんたい：" + escapeHtml(fullKanji) + "</div>";
      } else if (fullKanji && !ch) {
        html += "<span class=\"kanji-quiz-wrong-char-main\">" + escapeHtml(fullKanji) + "</span>";
      }
      redWrap.innerHTML = html || "<span class=\"kanji-quiz-wrong-char-main\">（データなし）</span>";
      panel.style.display = "block";
      if (ch) kanjiQuizMountWrongModelFrameForChar(ch);
      try {
        if (panel) panel.scrollIntoView({ block: "nearest", behavior: "smooth" });
      } catch (e) {}
    }
    function normalizeKanjiQuizInput(str) {
      try {
        return String(str || "").trim().normalize("NFKC");
      } catch (e) {
        return String(str || "").trim();
      }
    }
    /** れいぶん→よみ：伏字にせず原文の漢字をそのまま表示。対象字のみ赤で強調。 */
    function kanjiYomiFormatSentenceBlockHtml(q) {
      if (!q) return "";
      let s = String(q.sentence || "");
      const k = String(q.kanji || "");
      if (!s && q.maskedSentence && k) {
        const ms = String(q.maskedSentence);
        const i = ms.search(/[〼＿_]/);
        if (i >= 0) {
          s = ms.slice(0, i) + k + ms.slice(i + 1).replace(/^[〼＿_]+/, "");
        } else {
          s = ms;
        }
      }
      if (!s) return "";
      const idx = k && s.indexOf(k) >= 0 ? s.indexOf(k) : -1;
      if (idx < 0) {
        return '<div class="kanji-yomi-sentence-block">' + escapeHtml(s) + "</div>";
      }
      const before = escapeHtml(s.slice(0, idx));
      const mid = '<span class="kanji-yomi-target-char">' + escapeHtml(k) + "</span>";
      const after = escapeHtml(s.slice(idx + k.length));
      return '<div class="kanji-yomi-sentence-block">' + before + mid + after + "</div>";
    }
    function kanjiYomiBindTypingInputOnce() {
      const el = document.getElementById("kanji-play-yomi-input");
      if (!el || el.dataset.kanjiYomiBound === "1") return;
      el.dataset.kanjiYomiBound = "1";
      el.addEventListener("input", function () {
        const v = el.value;
        try {
          if (kanjiQuizSession) kanjiQuizSession.sentenceYomiRecognized = v;
        } catch (e) {}
        const hid = document.getElementById("kanji-play-input");
        if (hid) hid.value = v;
      });
    }
    function kanjiQuizIsAllHiraganaScript_(str) {
      const s = normalizeKanjiQuizInput(str);
      if (!s) return false;
      for (const ch of s) {
        const c = ch.codePointAt(0);
        if (c >= 0x3041 && c <= 0x3096) continue;
        if (c === 0x3099 || c === 0x309a) continue;
        if (c === 0x30fc) continue;
        return false;
      }
      return true;
    }
    function kanjiQuizIsAllKatakanaScript_(str) {
      const s = normalizeKanjiQuizInput(str);
      if (!s) return false;
      for (const ch of s) {
        const c = ch.codePointAt(0);
        if (c >= 0x30a1 && c <= 0x30f6) continue;
        if (c === 0x30fc || c === 0x30fd || c === 0x30fe) continue;
        return false;
      }
      return true;
    }
    function kanjiQuizSentenceYomiScriptBonusMultiplier(userAnswerNorm, readingKind) {
      const k = String(readingKind || "");
      if (k === "kun") return kanjiQuizIsAllHiraganaScript_(userAnswerNorm) ? 2 : 1;
      if (k === "on") return kanjiQuizIsAllKatakanaScript_(userAnswerNorm) ? 2 : 1;
      return 1;
    }
    function kanjiQuizTypeBadgeText(type) {
      if (type === "okurigana_shift") return "問題タイプ: おくりがな（選択）";
      if (type === "ruby_to_kanji") return "問題タイプ: よみ→かんじ（しゅどく）";
      if (type === "sentence_to_ruby") return "問題タイプ: れいぶん→よみ（タイピング）";
      if (type === "stroke_count") return "問題タイプ: かくすう";
      return "";
    }
    function resetKanjiQuizDrillPlayShell() {
      setKanjiQuizHandSubmitBusy(false);
      const charEl = document.getElementById("kanji-play-char");
      if (charEl) {
        charEl.style.display = "";
        charEl.style.fontSize = "";
        charEl.innerText = "";
      }
      const promptEl = document.getElementById("kanji-play-prompt");
      if (promptEl) promptEl.innerText = "";
      const detail = document.getElementById("kanji-play-detail");
      if (detail) detail.innerHTML = "";
      const choicesBox = document.getElementById("kanji-play-choices");
      if (choicesBox) {
        choicesBox.innerHTML = "";
        choicesBox.style.display = "none";
      }
      const typingWrap = document.getElementById("kanji-play-typing-wrap");
      if (typingWrap) typingWrap.style.display = "none";
      const inp = document.getElementById("kanji-play-input");
      if (inp) inp.value = "";
      const yomiInp = document.getElementById("kanji-play-yomi-input");
      if (yomiInp) yomiInp.value = "";
      try {
        kanjiQuizClearWritePad(true);
      } catch (eHw) {}
      try {
        if (kanjiQuizSession) kanjiQuizSession.sentenceYomiRecognized = "";
      } catch (eY2) {}
      const subBtn = document.getElementById("kanji-play-submit-btn");
      if (subBtn) subBtn.style.display = "none";
      const summary = document.getElementById("kanji-play-summary");
      if (summary) summary.innerHTML = "";
      try {
        kanjiQuizHideWrongFeedback();
      } catch (e) {}
      const drillHand = document.getElementById("kanji-quiz-drill-handwriting");
      if (drillHand) drillHand.style.display = "none";
      const refreshLayoutBtn = document.getElementById("kanji-quiz-layout-refresh-btn");
      if (refreshLayoutBtn) refreshLayoutBtn.style.display = "none";
      const penCtrls = document.getElementById("kanji-drill-pen-controls");
      if (penCtrls) penCtrls.style.display = "none";
      const cvsActions = document.getElementById("kanji-hw-canvas-actions");
      if (cvsActions) cvsActions.style.display = "none";
      const markerEl = document.getElementById("kanji-drill-q-marker");
      if (markerEl) markerEl.textContent = "";
      const prefixEl = document.getElementById("kanji-drill-ctx-prefix");
      if (prefixEl) prefixEl.textContent = "";
      const suffixEl = document.getElementById("kanji-drill-ctx-suffix");
      if (suffixEl) suffixEl.textContent = "";
      const charHand = document.getElementById("kanji-play-char-handwriting");
      if (charHand) charHand.textContent = "";
      const badge = document.getElementById("kanji-play-type-badge");
      if (badge) badge.innerText = "";
      try {
        kanjiQuizSyncPenUiFromStrokeParams();
      } catch (e) {}
    }
    function renderKanjiQuizQuestion() {
      if (!kanjiQuizSession) return;
      const q = kanjiQuizSession.questions[kanjiQuizSession.index];
      if (!q) return;
      if (!q.type) {
        alert("問題形式が不明です。やり直してください。");
        return;
      }
      if (
        (q.correctAnswer === undefined || q.correctAnswer === null || q.correctAnswer === "") &&
        q.type !== "ruby_to_kanji"
      ) {
        alert("問題データの形式が古い可能性があります。セット一覧から開き直してください。");
        return;
      }
      kanjiQuizSession.selectedChoice = null;
      resetKanjiQuizDrillPlayShell();
      const drillHand = document.getElementById("kanji-quiz-drill-handwriting");
      const markerEl = document.getElementById("kanji-drill-q-marker");
      const prefixEl = document.getElementById("kanji-drill-ctx-prefix");
      const suffixEl = document.getElementById("kanji-drill-ctx-suffix");
      const charHand = document.getElementById("kanji-play-char-handwriting");
      const total = kanjiQuizSession.questions.length;
      const progress = document.getElementById('kanji-play-progress');
      if (progress) progress.innerText = `${kanjiQuizSession.index + 1} / ${total}`;
      const title = document.getElementById('kanji-play-title');
      if (title) title.innerText = `【${kanjiQuizSession.modeName}】${formatUnitSheetDisplayLabel(kanjiQuizSession.unitName)} / セット ${kanjiQuizSession.setId}`;
      const badge = document.getElementById('kanji-play-type-badge');
      if (badge) badge.innerText = kanjiQuizTypeBadgeText(q.type);
      const charEl = document.getElementById('kanji-play-char');
      if (charEl) {
        charEl.style.fontSize = "";
        charEl.style.display = "";
        charEl.style.fontWeight = "";
        charEl.style.color = "";
        if (q.type === "okurigana_shift" || q.type === "stroke_count") {
          charEl.innerText = q.kanji || "？";
          if (q.type === "stroke_count") {
            charEl.style.fontSize = "clamp(72px,18vw,120px)";
          }
        } else if (q.type === "ruby_to_kanji") {
          charEl.innerText = "";
          charEl.style.display = "none";
        } else if (q.type === "sentence_to_ruby") {
          charEl.style.display = "";
          charEl.style.fontSize = "clamp(20px, 4.5vw, 30px)";
          charEl.style.fontWeight = "800";
          charEl.style.color = "#c62828";
          charEl.innerText = q.kanji ? "「" + q.kanji + "」の よみ" : q.readingLabel ? String(q.readingLabel) : "このかんじの よみ";
        } else {
          charEl.innerText = "かな";
        }
      }
      const promptEl = document.getElementById('kanji-play-prompt');
      if (promptEl) {
        if (q.type === "ruby_to_kanji") promptEl.innerText = "";
        else promptEl.innerText = q.prompt || "";
      }
      const detail = document.getElementById('kanji-play-detail');
      const choicesBox = document.getElementById('kanji-play-choices');
      const typingWrap = document.getElementById('kanji-play-typing-wrap');
      const subBtn = document.getElementById('kanji-play-submit-btn');
      const summary = document.getElementById('kanji-play-summary');
      if (summary) summary.innerHTML = "";
      if (q.type === "ruby_to_kanji") {
        try {
          kanjiQuizHideWrongFeedback();
        } catch (e) {}
        if (drillHand) drillHand.style.display = "block";
        const refreshLayoutBtn = document.getElementById("kanji-quiz-layout-refresh-btn");
        if (refreshLayoutBtn) refreshLayoutBtn.style.display = "inline-block";
        const penCtrls = document.getElementById("kanji-drill-pen-controls");
        if (penCtrls) penCtrls.style.display = "none";
        const cvsActions = document.getElementById("kanji-hw-canvas-actions");
        if (cvsActions) cvsActions.style.display = "flex";
        kanjiQuizEnsureScoreListener();
        const idxAtRender = kanjiQuizSession.index;
        const targets = kanjiQuizHanOnlyChars(q.correctAnswer || q.kanji);
        kanjiQuizSession.rubyHandTargets = targets;
        kanjiQuizSession.rubyHandSlot = 0;
        kanjiQuizSession.rubyHandComplete = false;
        kanjiQuizSession.rubyHandKanjiVgPass = false;
        kanjiQuizSession.lastHandScore = null;
        kanjiQuizSession.rubyHandMinScore = null;
        if (!targets.length) {
          alert("手書き対象の漢字が見つかりません。");
          return;
        }
        const ctxDrill = parseMaskedForDrill(q.maskedSentence);
        const qn = kanjiQuizCircledQuestionNum(kanjiQuizSession.index + 1);
        if (markerEl) markerEl.textContent = qn;
        if (prefixEl) prefixEl.textContent = ctxDrill.before || "";
        if (suffixEl) suffixEl.textContent = ctxDrill.after || "";
        if (charHand) charHand.textContent = q.readingDisplay || "";
        const t0 = targets[0];
        kanjiQuizSyncPenUiFromStrokeParams();
        kanjiQuizClearWritePad(false);
        kanjiQuizSetupWriteCanvas();
        kanjiQuizScheduleWriteCanvasReflow();
        if (summary) {
          summary.innerHTML =
            "<span style=\"color:#607d8b;font-size:clamp(12px,3vw,14px);\">さいてんようの データを じゅんびしています…</span>";
        }
        ensureKanjiHwFrameReadyOnce().then(function () {
          const sec = document.getElementById("section-kanji-quiz-play");
          if (!sec || !sec.classList.contains("active")) {
            if (summary) summary.innerHTML = "";
            return;
          }
          if (!kanjiQuizSession || kanjiQuizSession.index !== idxAtRender) {
            if (summary) summary.innerHTML = "";
            return;
          }
          const qNow = kanjiQuizSession.questions[kanjiQuizSession.index];
          if (!qNow || qNow.type !== "ruby_to_kanji") {
            if (summary) summary.innerHTML = "";
            return;
          }
          ensureKanjiFrameForQuizEval();
          applyKanjiQuizTargetsFromKanjiVg(targets, t0)
            .then(function () {
              const sec2 = document.getElementById("section-kanji-quiz-play");
              if (!sec2 || !sec2.classList.contains("active")) {
                if (summary) summary.innerHTML = "";
                return;
              }
              if (!kanjiQuizSession || kanjiQuizSession.index !== idxAtRender) {
                if (summary) summary.innerHTML = "";
                return;
              }
              const qNow2 = kanjiQuizSession.questions[kanjiQuizSession.index];
              if (!qNow2 || qNow2.type !== "ruby_to_kanji") {
                if (summary) summary.innerHTML = "";
                return;
              }
              const reassertCh =
                (kanjiQuizSession.rubyHandTargets || [])[kanjiQuizSession.rubyHandSlot || 0] || t0;
              [200, 520].forEach(function (delayMs) {
                setTimeout(function () {
                  if (!kanjiQuizSession || kanjiQuizSession.index !== idxAtRender) return;
                  const qLate = kanjiQuizSession.questions[kanjiQuizSession.index];
                  if (!qLate || qLate.type !== "ruby_to_kanji") return;
                  const frLate = document.getElementById("kp-pro-frame");
                  if (frLate) patchKanjiFrameForQuizPostMessage(frLate);
                  selectKanjiCharInQuizFrame(reassertCh);
                }, delayMs);
              });
              if (charHand) charHand.textContent = qNow2.readingDisplay || "";
              kanjiQuizSyncPenUiFromStrokeParams();
              kanjiQuizClearWritePad(true);
              kanjiQuizSetupWriteCanvas();
              kanjiQuizScheduleWriteCanvasReflow();
              if (summary) summary.innerHTML = "";
            })
            .catch(function (err) {
              if (summary) summary.innerHTML = "";
              alert(
                "漢字の筆順データを読み込めませんでした: " +
                  (err && err.message ? err.message : String(err))
              );
            });
        });
      } else {
        restoreKanjiPracticeFrameIfMoved();
        if (drillHand) drillHand.style.display = "none";
        if (markerEl) markerEl.textContent = "";
        if (prefixEl) prefixEl.textContent = "";
        if (suffixEl) suffixEl.textContent = "";
        if (charHand) charHand.textContent = "";
      }
      if (detail) {
        if (q.type === "okurigana_shift") {
          const sentenceHtml = kanjiYomiFormatSentenceBlockHtml(q);
          detail.innerHTML =
            `<div style="color:#333;font-size:clamp(16px,3.5vw,22px);line-height:1.55;"><strong>上</strong>に おいた かんじについて、<strong>おくりがなのつながり</strong>として もっとも ふさわしい ひょうきを、<strong>タップ</strong>して えらんでください。</div>
            <div style="margin-top:10px;font-size:clamp(14px,3.2vw,18px);color:#555;">※ まるごとの よみは だしていません。<strong>かたち</strong>だけを みて えらびましょう。</div>` +
            (sentenceHtml
              ? `<div style="margin-top:14px;">${sentenceHtml}</div>`
              : "");
        } else if (q.type === "ruby_to_kanji") {
          detail.innerHTML = "";
        } else if (q.type === "stroke_count") {
          detail.innerHTML =
            '<div style="color:#333;font-size:clamp(15px,3.2vw,20px);line-height:1.5;">教材の字体に基づく<strong>画数</strong>です。教育字体とちがう場合があります。</div>';
        } else if (q.type === "sentence_to_ruby") {
          detail.innerHTML = "";
        } else {
          detail.innerHTML = q.sentence
            ? `<div style="color:#222;font-size:clamp(20px,4.5vw,28px);line-height:1.5;">${escapeHtml(q.sentence || "")}</div>`
            : "";
        }
      }
      if (choicesBox) {
        choicesBox.innerHTML = "";
        if (q.type === "okurigana_shift" || q.type === "stroke_count") {
          choicesBox.style.display = "flex";
          const arr = shuffleKanjiQuizChoicesArray(q.choices);
          arr.forEach(function (c) {
            const b = document.createElement("button");
            b.type = "button";
            b.className = "kanji-drill-choice-btn";
            b.innerText = c;
            b.setAttribute("aria-label", "選択肢 " + c);
            b.onclick = function (e) {
              e.preventDefault();
              if (!kanjiQuizSession) return;
              kanjiQuizSession.selectedChoice = c;
              submitKanjiQuizScore();
            };
            choicesBox.appendChild(b);
          });
        } else {
          choicesBox.style.display = "none";
        }
      }
      if (typingWrap) {
        if (q.type === "sentence_to_ruby") {
          typingWrap.style.display = "flex";
          typingWrap.style.maxWidth = "min(980px, 98vw)";
          kanjiYomiBindTypingInputOnce();
          kanjiQuizSession.sentenceYomiRecognized = "";
          const hidInp = document.getElementById("kanji-play-input");
          if (hidInp) hidInp.value = "";
          const yomiInp = document.getElementById("kanji-play-yomi-input");
          if (yomiInp) yomiInp.value = "";
          const sentenceBox = document.getElementById("kanji-yomi-sentence-box");
          if (sentenceBox) sentenceBox.innerHTML = kanjiYomiFormatSentenceBlockHtml(q);
          renderCustomKeyboard(false, false, "kanji-yomi-keyboard-container", "kanji-play-yomi-input", function () {
            if (!kanjiQuizSession) return;
            const yomiInputNow = document.getElementById("kanji-play-yomi-input");
            const val = yomiInputNow ? String(yomiInputNow.value || "") : "";
            kanjiQuizSession.sentenceYomiRecognized = val;
            const hidNow = document.getElementById("kanji-play-input");
            if (hidNow) hidNow.value = val;
            submitKanjiQuizScore();
          });
        } else {
          typingWrap.style.display = "none";
        }
      }
      if (subBtn) {
        subBtn.style.display = q.type === "sentence_to_ruby" ? "inline-block" : "none";
      }
    }
    function startKanjiQuizPlay(ctx) {
      const mode = ctx.formatMode != null ? ctx.formatMode : getKanjiQuizFormatMode();
      let questions;
      let allQuestionsStored;
      if (ctx.nigateBypassFilter && Array.isArray(ctx.questions) && ctx.questions.length) {
        questions = shuffleKanjiQuizQuestionsArray(ctx.questions.slice());
        allQuestionsStored = questions.slice();
      } else if (Array.isArray(ctx.allQuestions) && ctx.allQuestions.length) {
        const filtered = filterKanjiQuizQuestionsByFormat(ctx.allQuestions, mode);
        if (!filtered.length) {
          alert("この しかた では もんだいがありません。\nほかの しかたを えらぶか、混合にしてください。");
          return;
        }
        allQuestionsStored = filtered.slice();
        questions = shuffleKanjiQuizQuestionsArray(filtered.slice());
      } else {
        const base = Array.isArray(ctx.questions) ? ctx.questions.slice() : [];
        if (!base.length) {
          alert("このセットには問題がありません。");
          return;
        }
        questions = shuffleKanjiQuizQuestionsArray(base);
        allQuestionsStored = questions.slice();
      }
      if (!questions.length) {
        alert("このセットには問題がありません。");
        return;
      }
      const proceed = function () {
        kanjiQuizSession = {
          modeId: ctx.modeId,
          modeName: ctx.modeName,
          unitName: ctx.unitName,
          setId: ctx.setId,
          isTrainingMode: !!ctx.isTrainingMode,
          trainingStepIndex: ctx.trainingStepIndex,
          trainingMenuId: ctx.trainingMenuId,
          questions,
          index: 0,
          totalEarned: 0,
          newTotalPoints: null,
          logs: [],
          selectedChoice: null,
          nigateTraining: !!ctx.nigateTraining,
          nigateAxis: ctx.nigateAxis || null,
          nigateFeedback: ctx.nigateTraining ? { strokeOrderClean: true, brushAllClear: true } : null
        };
        lastKanjiQuizContext = {
          modeId: ctx.modeId,
          modeName: ctx.modeName,
          unitName: ctx.unitName,
          setId: ctx.setId,
          allQuestions: allQuestionsStored,
          questions: questions,
          isTrainingMode: !!ctx.isTrainingMode,
          trainingStepIndex: ctx.trainingStepIndex,
          trainingMenuId: ctx.trainingMenuId,
          formatMode: mode,
          nigateBypassFilter: !!ctx.nigateBypassFilter,
          nigateTraining: !!ctx.nigateTraining,
          nigateAxis: ctx.nigateAxis || null
        };
        switchSection("section-kanji-quiz-play");
        saveKanjiQuizRecoveryDraft();
        renderKanjiQuizQuestion();
      };
      proceed();
    }
    function restartLastKanjiQuizSet() {
      const ctx = lastKanjiQuizContext;
      if (!ctx) {
        openKanjiLearningMenu();
        return;
      }
      if (ctx.nigateBypassFilter && Array.isArray(ctx.questions) && ctx.questions.length) {
        startKanjiQuizPlay({
          modeId: ctx.modeId,
          modeName: ctx.modeName,
          unitName: ctx.unitName,
          setId: ctx.setId,
          questions: ctx.questions,
          nigateBypassFilter: true,
          nigateTraining: !!ctx.nigateTraining,
          nigateAxis: ctx.nigateAxis,
          formatMode: ctx.formatMode || "mixed"
        });
        return;
      }
      const mode = ctx.formatMode != null ? ctx.formatMode : getKanjiQuizFormatMode();
      if (Array.isArray(ctx.questions) && ctx.questions.length) {
        startKanjiQuizPlay({
          modeId: ctx.modeId,
          modeName: ctx.modeName,
          unitName: ctx.unitName,
          setId: ctx.setId,
          questions: ctx.questions.slice(),
          formatMode: mode,
          isTrainingMode: !!ctx.isTrainingMode,
          trainingStepIndex: ctx.trainingStepIndex,
          trainingMenuId: ctx.trainingMenuId
        });
        return;
      }
      if (Array.isArray(ctx.allQuestions) && ctx.allQuestions.length) {
        const filtered = filterKanjiQuizQuestionsByFormat(ctx.allQuestions, mode);
        if (!filtered.length) {
          alert(
            "この しかた では もんだいがありません。\nほかの しかたを えらぶか、混合にしてください。"
          );
          return;
        }
        startKanjiQuizPlay({
          modeId: ctx.modeId,
          modeName: ctx.modeName,
          unitName: ctx.unitName,
          setId: ctx.setId,
          allQuestions: ctx.allQuestions,
          formatMode: mode,
          isTrainingMode: !!ctx.isTrainingMode,
          trainingStepIndex: ctx.trainingStepIndex,
          trainingMenuId: ctx.trainingMenuId
        });
        return;
      }
      openKanjiLearningMenu();
    }
    function showKanjiQuizResult(totalEarned, logs, isTrainingMode, newTotalPoints) {
      restoreKanjiPracticeFrameIfMoved();
      switchSection('section-result');
      const rc = document.getElementById('result-content');
      const retryBtn = document.getElementById('result-retry-btn');
      const settingsBtn = document.getElementById('result-settings-btn');
      const homeBtn = document.getElementById('result-home-btn');
      if (!rc || !retryBtn || !settingsBtn || !homeBtn) return;
      const rows = (Array.isArray(logs) ? logs : []).map(function (v) {
        const mark = v.isCorrect ? "○" : "×";
        return `「${escapeHtml(v.kanji || "")}」 ${mark} ${Number(v.score || 0)}点 → +${Number(v.earned || 0).toFixed(2)}Pt`;
      }).join("<br>");
      const totalLine = (typeof newTotalPoints === "number" && !isNaN(newTotalPoints))
        ? `<p style="margin-top:8px;">合計ポイント: ${Number(newTotalPoints).toFixed(2)}</p>`
        : "";
      var nigateBlock = "";
      var ctxR = lastKanjiQuizContext;
      if (ctxR && ctxR.nigateFeedbackSnapshot && ctxR.nigateTraining) {
        var f = ctxR.nigateFeedbackSnapshot;
        var p1 = f.strokeOrderClean
          ? "手書きでは、かきじゅんの指摘はありませんでした。"
          : "手書きで、かきじゅんに注意が出た回があります。";
        var p2 = f.brushAllClear
          ? "とめ・はね・はらいは、すべて問題なしと判定された手書きがありました（厳密モード時のみ有効です）。"
          : "とめ・はね・はらいに注意が出た手書きがあります（厳密モード時）。";
        nigateBlock =
          '<div style="margin-top:14px;padding:10px;background:#f3e5f5;border-radius:8px;text-align:left;font-size:14px;line-height:1.5;color:#4a148c;"><strong>ニガテ特訓のまとめ</strong><br>' +
          escapeHtml(p1) +
          "<br>" +
          escapeHtml(p2) +
          "</div>";
      }
      rc.innerHTML = `<h2 class="kanji-result-title">✨ 漢字セット完了！ ✨</h2>
        <p class="kanji-result-pts">かくとくポイント: <span class="kanji-result-pts-val">+${Number(totalEarned || 0).toFixed(2)}</span></p>
        <hr class="kanji-result-hr">
        <div class="kanji-result-logs">${rows || "結果なし"}</div>
        ${totalLine}` + nigateBlock;
      retryBtn.style.display = "block";
      retryBtn.innerText = "🔄 同じセットでもう一度";
      retryBtn.onclick = restartLastKanjiQuizSet;
      if (isTrainingMode) {
        settingsBtn.style.display = "block";
        settingsBtn.innerText = "🎯 特訓メニューにもどる";
        settingsBtn.onclick = () => openTrainingMenu();
      } else if (lastKanjiQuizContext && lastKanjiQuizContext.nigateTraining) {
        settingsBtn.style.display = "block";
        settingsBtn.innerText = "🎯 ニガテ特訓にもどる";
        settingsBtn.onclick = function () { switchSection("section-kanji-nigate"); };
      } else {
        settingsBtn.style.display = "block";
        settingsBtn.innerText = "📚 セット一覧にもどる";
        settingsBtn.onclick = () => switchSection('section-kanji-quiz-sets');
      }
      homeBtn.style.display = "block";
      homeBtn.innerText = "🏠 ホームにもどる";
      homeBtn.onclick = () => showHome(JSON.parse(localStorage.getItem('app_kid_user')));
    }
    function cancelKanjiQuizPlay() {
      var goNigate = !!(lastKanjiQuizContext && lastKanjiQuizContext.nigateTraining);
      if (!kanjiQuizSession) {
        restoreKanjiPracticeFrameIfMoved();
        resetKanjiQuizDrillPlayShell();
        switchSection(goNigate ? "section-kanji-nigate" : "section-kanji-quiz-sets");
        return;
      }
      if (!confirm("このセットを中断しますか？")) return;
      kanjiQuizSession = null;
      lastKanjiQuizContext = null;
      clearKanjiHwFrameReadyCache();
      try {
        kanjiQuizClearWritePad(true);
      } catch (e) {}
      restoreKanjiPracticeFrameIfMoved();
      resetKanjiQuizDrillPlayShell();
      switchSection(goNigate ? "section-kanji-nigate" : "section-kanji-quiz-sets");
    }
    function submitKanjiQuizScore() {
      if (!kanjiQuizSession) return;
      const q = kanjiQuizSession.questions[kanjiQuizSession.index];
      if (!q) return;
      if (!q.type) {
        alert("問題形式が不明です。やり直してください。");
        return;
      }
      let userRaw = "";
      let isCorrect = false;
      var scriptBonusMult = 1;
      if (q.type === "okurigana_shift") {
        userRaw = kanjiQuizSession.selectedChoice != null ? String(kanjiQuizSession.selectedChoice) : "";
        if (!userRaw) {
          alert("選択肢を選んでください。");
          return;
        }
        isCorrect = normalizeKanjiQuizInput(userRaw) === normalizeKanjiQuizInput(q.correctAnswer);
      } else if (q.type === "ruby_to_kanji") {
        if (!kanjiQuizSession.rubyHandComplete) {
          alert("かくかんじを マスにかき、「これでかいとう」で こたえてください。");
          return;
        }
        isCorrect = !!kanjiQuizSession.rubyHandKanjiVgPass;
      } else if (q.type === "sentence_to_ruby") {
        userRaw =
          kanjiQuizSession && kanjiQuizSession.sentenceYomiRecognized != null
            ? String(kanjiQuizSession.sentenceYomiRecognized || "")
            : "";
        const hidInp = document.getElementById("kanji-play-input");
        if (!normalizeKanjiQuizInput(userRaw) && hidInp) {
          userRaw = String(hidInp.value || "");
        }
        if (!normalizeKanjiQuizInput(userRaw)) {
          alert("よみを入力してから、「こたえを決定」を おしてください。");
          return;
        }
        isCorrect = normalizeKanjiQuizInput(userRaw) === normalizeKanjiQuizInput(q.correctAnswer);
        if (isCorrect) {
          scriptBonusMult = kanjiQuizSentenceYomiScriptBonusMultiplier(
            normalizeKanjiQuizInput(userRaw),
            q.readingKind
          );
        }
      } else if (q.type === "stroke_count") {
        userRaw = kanjiQuizSession.selectedChoice != null ? String(kanjiQuizSession.selectedChoice) : "";
        if (!userRaw) {
          alert("かくすうを選んでください。");
          return;
        }
        isCorrect = userRaw === String(q.correctAnswer);
      }
      if (
        !isCorrect &&
        kanjiQuizSession &&
        (q.type === "sentence_to_ruby" || q.type === "okurigana_shift")
      ) {
        queueKanjiReadingWeakSignal();
      }
      if (!isCorrect && kanjiQuizSession && q.type === "stroke_count") {
        queueKanjiStrokeCountWeakSignal();
      }
      var scoreForServer = 0;
      if (isCorrect) {
        if (q.type === "ruby_to_kanji") {
          const m = kanjiQuizSession.rubyHandMinScore;
          const fall = Number(kanjiQuizSession.lastHandScore);
          var v = m != null && !isNaN(m) ? m : (!isNaN(fall) && fall > 0 ? fall : 100);
          scoreForServer = Math.max(0, Math.min(100, Math.round(v)));
        } else {
          scoreForServer = 100;
        }
      }
      kanjiQuizSession.selectedChoice = null;
      const user = JSON.parse(localStorage.getItem('app_kid_user') || 'null');
      if (!user || !user.id) {
        alert("ログイン情報が見つかりません。");
        return;
      }
      if (__kanjiQuizSubmitInFlight) return;
      __kanjiQuizSubmitInFlight = true;
      const qid = q.questionId || (`${q.kanji}_${kanjiQuizSession.index}_${q.type}`);
      const unitId = `KANJI_${kanjiQuizSession.modeName}_${kanjiQuizSession.unitName}_SET${kanjiQuizSession.setId}_${qid}`;
      const kanjiSetScopeId = `KANJI_${kanjiQuizSession.modeName}_${kanjiQuizSession.unitName}_SET${kanjiQuizSession.setId}`;
      const payload = {
        action: "save_learning_session",
        userId: user.id,
        unitId: unitId,
        unitSheetName: kanjiQuizSession.unitName,
        isReviewMode: false,
        isRandom: false,
        results: [],
        learningCategory: "kanji",
        challengeType: "score",
        kanjiChar: q.kanji,
        score: scoreForServer,
        questionId: qid,
        questionCorrect: isCorrect,
        kanjiSetScopeId: kanjiSetScopeId,
        kanjiSetContinuation: kanjiQuizSession.index > 0,
        kanjiScriptBonusMult: scriptBonusMult
      };
      if (kanjiQuizSession.isTrainingMode) {
        payload.trainingStepIndex = kanjiQuizSession.trainingStepIndex;
        payload.trainingMenuId = kanjiQuizSession.trainingMenuId;
      }
      const summary = document.getElementById('kanji-play-summary');
      if (summary) {
        if (isCorrect && q.type === "sentence_to_ruby" && scriptBonusMult >= 2) {
          summary.innerHTML =
            '<span style="color:#69F0AE;">せいかい！（ポイント2ばい）</span>';
        } else if (isCorrect) {
          summary.innerHTML = '<span style="color:#69F0AE;">せいかい！</span>';
        } else {
          summary.innerHTML =
            '<span style="color:#FF8A80;">ざんねん… 次はがんばろう</span>';
        }
      }
      function postKanjiScoreWithRetry(retryCount) {
        return fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify(payload) })
          .then(function (r) { return r.json(); })
          .catch(function (e) {
            const msg = String((e && (e.message || e)) || "");
            const canRetry = retryCount > 0 && /Load failed|Failed to fetch|NetworkError|fetch/i.test(msg);
            if (!canRetry) throw e;
            return new Promise(function (resolve) {
              setTimeout(resolve, 450);
            }).then(function () {
              return postKanjiScoreWithRetry(retryCount - 1);
            });
          });
      }
      postKanjiScoreWithRetry(1).then(d => {
        if (d.status !== "success") throw new Error(d.message || "保存に失敗");
        try {
          showKanjiEarnedPointsToast(d.earnedPoints);
        } catch (ePt) {}
        kanjiQuizSession.totalEarned += Number(d.earnedPoints || 0);
        kanjiQuizSession.newTotalPoints = Number(d.newTotal);
        kanjiQuizSession.logs.push({
          kanji: q.kanji,
          score: scoreForServer,
          earned: Number(d.earnedPoints || 0),
          isCorrect: isCorrect,
          qType: q.type
        });
        user.points = d.newTotal;
        localStorage.setItem('app_kid_user', JSON.stringify(user));
        const ptsEl = document.getElementById("user-points");
        if (ptsEl) ptsEl.innerText = String(d.newTotal);
        /* 次の問題へ進む直前：保存完了後でも render 前は旧ドリルが残り得るため、親キャンバス＋採点 iframe を先に空にする */
        try {
          kanjiQuizClearWritePad(true);
          kanjiQuizSetupWriteCanvas();
        } catch (ePreAdv) {}
        kanjiQuizSession.index += 1;
        if (kanjiQuizSession.index >= kanjiQuizSession.questions.length) {
          const finished = {
            totalEarned: kanjiQuizSession.totalEarned,
            logs: kanjiQuizSession.logs.slice(),
            isTrainingMode: !!kanjiQuizSession.isTrainingMode,
            newTotalPoints: Number(kanjiQuizSession.newTotalPoints)
          };
          const nFb = kanjiQuizSession.nigateTraining ? kanjiQuizSession.nigateFeedback : null;
          if (lastKanjiQuizContext && nFb) {
            lastKanjiQuizContext.nigateFeedbackSnapshot = nFb;
          }
          if (kanjiQuizSession.isTrainingMode) {
            fetchTrainingRoute(user.id);
          }
          kanjiQuizSession = null;
          clearKanjiQuizRecoveryDraft();
          showKanjiQuizResult(finished.totalEarned, finished.logs, finished.isTrainingMode, finished.newTotalPoints);
          return;
        }
        saveKanjiQuizRecoveryDraft();
        renderKanjiQuizQuestion();
      }).catch(e => {
        console.warn("kanji save failed, continue quiz:", e);
        const summaryEl = document.getElementById('kanji-play-summary');
        if (summaryEl) {
          summaryEl.innerHTML = '<span style="color:#FF9800;">通信が不安定です。保存はあとで再試行し、次の問題へ進みます。</span>';
        }
        kanjiQuizSession.logs.push({
          kanji: q.kanji,
          score: scoreForServer,
          earned: 0,
          isCorrect: isCorrect,
          qType: q.type,
          saveFailed: true
        });
        try {
          kanjiQuizClearWritePad(true);
          kanjiQuizSetupWriteCanvas();
        } catch (ePreAdv2) {}
        kanjiQuizSession.index += 1;
        if (kanjiQuizSession.index >= kanjiQuizSession.questions.length) {
          const finished = {
            totalEarned: kanjiQuizSession.totalEarned,
            logs: kanjiQuizSession.logs.slice(),
            isTrainingMode: !!kanjiQuizSession.isTrainingMode,
            newTotalPoints: Number(kanjiQuizSession.newTotalPoints)
          };
          const nFb = kanjiQuizSession.nigateTraining ? kanjiQuizSession.nigateFeedback : null;
          if (lastKanjiQuizContext && nFb) {
            lastKanjiQuizContext.nigateFeedbackSnapshot = nFb;
          }
          if (kanjiQuizSession.isTrainingMode) {
            fetchTrainingRoute(user.id);
          }
          kanjiQuizSession = null;
          clearKanjiQuizRecoveryDraft();
          showKanjiQuizResult(finished.totalEarned, finished.logs, finished.isTrainingMode, finished.newTotalPoints);
          return;
        }
        saveKanjiQuizRecoveryDraft();
        renderKanjiQuizQuestion();
      }).finally(function () {
        __kanjiQuizSubmitInFlight = false;
      });
    }
    function loadQuestionsForSettings(btn, mId, mName, uName, categoryHint) { 
      const origText = toggleBtnLoading(btn, true); 
      const isKanjiMaterial = String(categoryHint || "").toLowerCase() === "kanji" || /漢字|かんじ|kanji/i.test(String(mName || ""));
      if (isKanjiMaterial) {
        openKanjiQuizSets(mId, mName, uName, btn, origText);
        return;
      }
      const cacheKey = `app_cached_questions_${mId}_${uName}`;
      const cached = localStorage.getItem(cacheKey);
      
      const processData = (d) => {
        toggleBtnLoading(btn, false, origText); 
        if(d.status==="success"){ 
          currentQuestions = d.questions; currentModeId = mId; currentModeName = mName; currentUnitName = uName; 
          // 特訓モードフラグをリセットしておく（いつもの学習から入った場合）
          isTrainingMode = false;
          openSettingsScreen(); 
        } else {
          alert("取得失敗: " + (d.message || "エラー"));
        }
      };

      if (cached) {
        try {
          const d = JSON.parse(cached);
          processData(d);
          return;
        } catch(e) {}
      }

      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify({ action: "get_questions", modeId: mId, unitName: uName }) })
      .then(r=>r.json()).then(d=>{ 
        if(d.status === "success") {
          localStorage.setItem(cacheKey, JSON.stringify(d));
        }
        processData(d);
      }).catch(e => toggleBtnLoading(btn, false, origText)); 
    }

    function openSettingsScreen() {
      initKeyboardAndSoundSettings();
      switchSection('section-settings'); 
      document.getElementById('settings-title').innerText=`【${currentModeName}】${formatUnitSheetDisplayLabel(currentUnitName)}`;
      const fSelect = document.getElementById('setting-format');
      if (currentModeName.includes("単語")) { fSelect.innerHTML = `<option value="ja_to_en">日本語 ➔ 英語にする</option><option value="en_to_ja">英語 ➔ 日本語にする</option><option value="en_audio_to_ja">英語読み上げ ➔ 日本語にする</option><option value="en_audio_to_en">英語読み上げ ➔ 英語にする</option><option value="en_to_en">英語 ➔ 英語にする</option>`; } 
      else { fSelect.innerHTML = `<option value="ja_to_en">日本語 ➔ 英語にする</option><option value="ja_to_en_sort">並び替え（日本語付き）</option><option value="en_to_ja">英語 ➔ 日本語にする</option><option value="qtext_to_en">疑問文 ➔ 英語にする</option><option value="qaudio_to_en">疑問文(読み上げ) ➔ 英語にする</option><option value="en_audio_to_en">英語読み上げ ➔ 英語にする</option><option value="en_to_en">英語 ➔ 英語にする</option>`; }
      updateAnswerTypeUI();
    }

    function updateAnswerTypeUI() {
      const format = document.getElementById('setting-format').value;
      const aSelect = document.getElementById('setting-answer-type');
      aSelect.innerHTML = "";
      const isWord = currentModeName.includes("単語");

      if (format === "ja_to_en_sort") {
        aSelect.innerHTML += `<option value="sort_all" selected>すべて用いる</option>`;
        aSelect.innerHTML += `<option value="sort_dummy">不要語混入</option>`;
        aSelect.innerHTML += `<option value="sort_missing">不足語補足</option>`;
        toggleBlankSetting();
        return;
      }

      if (format === "ja_to_en") {
        aSelect.innerHTML += `<option value="4choice">４択（えらぶ）</option>`;
        aSelect.innerHTML += `<option value="typing" selected>タイピング（入力する）</option>`;
        aSelect.innerHTML += `<option value="voice">🎙️ 音声入力（マイク）</option>`;
        if (isWord) {
          aSelect.innerHTML += `<option value="fill_4choice">🔠 穴埋め（４択）</option>`;
          aSelect.innerHTML += `<option value="fill_typing">⌨️ 穴埋め（タイピング）</option>`;
        }
      } else if (format === "qtext_to_en" || format === "qaudio_to_en") {
        if(format !== "qaudio_to_en") aSelect.innerHTML += `<option value="4choice">４択（えらぶ）</option>`;
        aSelect.innerHTML += `<option value="typing" selected>タイピング（入力する）</option><option value="voice">🎙️ 音声入力（マイク）</option>`;
      } else if (format === "en_to_ja" || format === "en_audio_to_ja") {
        aSelect.innerHTML += `<option value="4choice" selected>４択（えらぶ）</option>`;
      } else if (format === "en_audio_to_en") {
        aSelect.innerHTML += `<option value="4choice">４択（えらぶ）</option>`;
        aSelect.innerHTML += `<option value="typing" selected>タイピング（入力する）</option>`;
        aSelect.innerHTML += `<option value="voice">🎙️ 音声入力（マイク）</option>`;
        if (isWord) {
          aSelect.innerHTML += `<option value="fill_4choice">🔠 穴埋め（４択）</option>`;
          aSelect.innerHTML += `<option value="fill_typing">⌨️ 穴埋め（タイピング）</option>`;
        }
      } else if (format === "en_to_en") {
        aSelect.innerHTML += `<option value="typing" selected>タイピング（入力する）</option>`;
        aSelect.innerHTML += `<option value="voice">🎙️ 音声入力（マイク）</option>`;
      }
      toggleBlankSetting();
    }

    function toggleBlankSetting() {
      const aSelect = document.getElementById('setting-answer-type').value;
      const blankGroup = document.getElementById('setting-group-blanks');
      const blankSelect = document.getElementById('setting-blank-count');
      
      if (aSelect.startsWith('fill_')) {
        let maxLen = 0;
        currentQuestions.forEach(q => {
          let word = (q["英単語"] || "").trim();
          if(word.length > maxLen) maxLen = word.length;
        });
        
        let maxBlanks = maxLen > 1 ? maxLen - 1 : 1; 
        
        blankSelect.innerHTML = "";
        for(let i = 1; i <= maxBlanks; i++) {
          blankSelect.innerHTML += `<option value="${i}">${i} 文字 かくす</option>`;
        }
        blankGroup.style.display = "block";
      } else {
        blankGroup.style.display = "none";
      }
    }
    function normalizeAnswerTypeForFormat(format, answerType) {
      const f = String(format || "").trim();
      const a = String(answerType || "").trim();
      if (f === "en_to_ja" || f === "en_audio_to_ja") return "4choice";
      return a;
    }
    function syncAnswerTypeWithFormat(format) {
      const ansTypeEl = document.getElementById('setting-answer-type');
      if (!ansTypeEl) return "";
      const current = String(ansTypeEl.value || "").trim();
      const normalized = normalizeAnswerTypeForFormat(format, current);
      if (normalized !== current) {
        const hasOption = Array.from(ansTypeEl.options || []).some(opt => String(opt.value) === normalized);
        if (!hasOption) {
          if (normalized === "4choice") {
            ansTypeEl.innerHTML = `<option value="4choice" selected>４択（えらぶ）</option>`;
          } else {
            ansTypeEl.innerHTML = `<option value="${normalized}" selected>${normalized}</option>`;
          }
        }
        ansTypeEl.value = normalized;
      }
      return String(ansTypeEl.value || "").trim();
    }

    let shiftMode = 0; let isShiftHoldMode = false; let shiftPressStartTime = 0; let shiftTimer1s = null; let shiftTimer2s = null;
    function toKatakanaFromHiragana(text) {
      return String(text || "").replace(/[ぁ-ゖ]/g, function (ch) {
        return String.fromCharCode(ch.charCodeAt(0) + 0x60);
      });
    }
    function convertRomajiChunkToHiragana(romaji) {
      const src = String(romaji || "").toLowerCase();
      const table = {
        kya: "きゃ", kyu: "きゅ", kyo: "きょ", sha: "しゃ", shu: "しゅ", sho: "しょ", cha: "ちゃ", chu: "ちゅ", cho: "ちょ",
        nya: "にゃ", nyu: "にゅ", nyo: "にょ", hya: "ひゃ", hyu: "ひゅ", hyo: "ひょ", mya: "みゃ", myu: "みゅ", myo: "みょ",
        rya: "りゃ", ryu: "りゅ", ryo: "りょ", gya: "ぎゃ", gyu: "ぎゅ", gyo: "ぎょ", ja: "じゃ", ju: "じゅ", jo: "じょ",
        bya: "びゃ", byu: "びゅ", byo: "びょ", pya: "ぴゃ", pyu: "ぴゅ", pyo: "ぴょ",
        fa: "ふぁ", fi: "ふぃ", fe: "ふぇ", fo: "ふぉ",
        a: "あ", i: "い", u: "う", e: "え", o: "お",
        ka: "か", ki: "き", ku: "く", ke: "け", ko: "こ",
        sa: "さ", shi: "し", su: "す", se: "せ", so: "そ",
        ta: "た", chi: "ち", tsu: "つ", te: "て", to: "と",
        na: "な", ni: "に", nu: "ぬ", ne: "ね", no: "の",
        ha: "は", hi: "ひ", fu: "ふ", he: "へ", ho: "ほ",
        ma: "ま", mi: "み", mu: "む", me: "め", mo: "も",
        ya: "や", yu: "ゆ", yo: "よ",
        ra: "ら", ri: "り", ru: "る", re: "れ", ro: "ろ",
        wa: "わ", wo: "を",
        ga: "が", gi: "ぎ", gu: "ぐ", ge: "げ", go: "ご",
        za: "ざ", ji: "じ", zu: "ず", ze: "ぜ", zo: "ぞ",
        da: "だ", de: "で", do: "ど", di: "ぢ", du: "づ",
        ba: "ば", bi: "び", bu: "ぶ", be: "べ", bo: "ぼ",
        pa: "ぱ", pi: "ぴ", pu: "ぷ", pe: "ぺ", po: "ぽ",
        nn: "ん", n: "ん",
        va: "ゔぁ", vi: "ゔぃ", vu: "ゔ", ve: "ゔぇ", vo: "ゔぉ",
        "-": "ー"
      };
      let i = 0;
      let out = "";
      while (i < src.length) {
        const a = src[i];
        const b = src[i + 1] || "";
        if (a && b && a === b && /[bcdfghjklmpqrstvwxyz]/.test(a) && a !== "n") {
          out += "っ";
          i += 1;
          continue;
        }
        const tri = src.slice(i, i + 3);
        const bi = src.slice(i, i + 2);
        if (table[tri]) { out += table[tri]; i += 3; continue; }
        if (table[bi]) { out += table[bi]; i += 2; continue; }
        if (table[a]) { out += table[a]; i += 1; continue; }
        i += 1;
      }
      return out;
    }
    function convertRomajiToKana(romaji, asKatakana) {
      const hira = convertRomajiChunkToHiragana(romaji);
      return asKatakana ? toKatakanaFromHiragana(hira) : hira;
    }
    function updateKeyboardDisplay() { const isUpper = shiftMode > 0 || isShiftHoldMode; const shiftBtn = document.getElementById('shift-btn'); if (shiftBtn) { if (shiftMode === 2) { shiftBtn.innerText = '⇪'; shiftBtn.style.background = '#FF9800'; } else if (shiftMode === 1 || isShiftHoldMode) { shiftBtn.innerText = '⇧'; shiftBtn.style.background = '#e50914'; } else { shiftBtn.innerText = '⇧'; shiftBtn.style.background = '#444'; } } document.querySelectorAll('.key-char').forEach(btn => { const baseChar = btn.getAttribute('data-char'); if (/[a-z]/.test(baseChar)) btn.innerText = isUpper ? baseChar.toUpperCase() : baseChar; }); }
    function renderCustomKeyboard(isFillMode = false, sortMissingMode = false, containerId = 'keyboard-container', targetInputId = 'type-answer', onEnterSubmit = null) {
      const container = document.getElementById(containerId); if(!container) return; container.innerHTML = "";
      const isKanaTypingMode = !sortMissingMode && !isFillMode && targetInputId === 'kanji-play-yomi-input';
      let kanaCommitted = "";
      let kanaRomajiBuffer = "";
      let kanaIsKatakana = false;
      window.__sortMissingModeActive = !!sortMissingMode;
      const scalePct = parseInt(getUserPref(LS_KBD_SCALE, '100'), 10);
      const fontPx = parseInt(getUserPref(LS_KBD_FONT, '18'), 10);
      const padPx = Math.max(8, Math.round(fontPx * 0.72));
      const wrap = document.createElement('div');
      wrap.className = 'keyboard-scale-wrap';
      wrap.style.setProperty('--kb-scale', String(Math.max(0.5, Math.min(2.5, scalePct / 100))));
      wrap.style.setProperty('--vk-font-px', String(fontPx));
      wrap.style.setProperty('--vk-pad-px', String(padPx));
      const rows = [ ['!', '"', '$', '%', '&', "'", '(', ')', '-', '¥'], ['q','w','e','r','t','y','u','i','o','p'], ['a','s','d','f','g','h','j','k','l'], ['z','x','c','v','b','n','m',',','.','?'] ];
      const board = document.createElement('div'); board.className = 'keyboard';
      
      rows.forEach((row, rIdx) => {
        const rowDiv = document.createElement('div'); rowDiv.className = 'key-row';
        if(rIdx === 3) {
          const shiftBtn = document.createElement('button'); shiftBtn.className = 'key key-wide'; shiftBtn.id = 'shift-btn'; shiftBtn.style.touchAction = 'none'; 
          if (isKanaTypingMode) {
            const refreshKanaShiftLabel = function () {
              shiftBtn.innerText = "ひら⇔カタ";
              shiftBtn.style.background = kanaIsKatakana ? "#e50914" : "#444";
            };
            shiftBtn.onclick = function (e) {
              e.preventDefault();
              kanaIsKatakana = !kanaIsKatakana;
              refreshKanaShiftLabel();
            };
            refreshKanaShiftLabel();
          } else {
            shiftBtn.onpointerdown = (e) => { e.preventDefault(); if(e.button !== 0 && e.pointerType === 'mouse') return; if (shiftPressStartTime > 0) return; shiftPressStartTime = Date.now(); shiftTimer1s = setTimeout(() => { if (shiftBtn && !isShiftHoldMode) { shiftBtn.style.background = '#FF9800'; shiftBtn.innerText = '⇪'; } }, 1000); shiftTimer2s = setTimeout(() => { isShiftHoldMode = true; shiftMode = 0; updateKeyboardDisplay(); }, 2000); };
            const handlePointerUp = (e) => { e.preventDefault(); if (shiftPressStartTime === 0) return; clearTimeout(shiftTimer1s); clearTimeout(shiftTimer2s); const duration = Date.now() - shiftPressStartTime; shiftPressStartTime = 0; if (isShiftHoldMode) { isShiftHoldMode = false; shiftMode = 0; } else if (duration >= 1000) { shiftMode = (shiftMode === 2) ? 0 : 2; } else if (duration > 0 && duration < 1000) { shiftMode = (shiftMode === 1 || shiftMode === 2) ? 0 : 1; } updateKeyboardDisplay(); };
            shiftBtn.onpointerup = handlePointerUp; shiftBtn.onpointerleave = handlePointerUp; shiftBtn.onpointercancel = handlePointerUp; shiftBtn.oncontextmenu = (e) => { e.preventDefault(); return false; };
          }
          rowDiv.appendChild(shiftBtn);
        }
        
        row.forEach(char => {
          const keyBtn = document.createElement('button'); keyBtn.innerText = char;
          if (/[a-z]/.test(char)) { keyBtn.className = 'key key-char'; keyBtn.setAttribute('data-char', char); } else { keyBtn.className = 'key'; }
          keyBtn.onclick = () => {
            const isUpper = shiftMode > 0 || isShiftHoldMode; let inputChar = char;
            if (/[a-z]/.test(char) && isUpper) inputChar = char.toUpperCase();
            
            if (isKanaTypingMode) {
              const targetInput = document.getElementById(targetInputId);
              if (targetInput) {
                kanaRomajiBuffer += String(inputChar || "").toLowerCase();
                const matched = convertRomajiToKana(kanaRomajiBuffer, kanaIsKatakana);
                if (matched) {
                  kanaCommitted += matched;
                  kanaRomajiBuffer = "";
                }
                targetInput.value = kanaCommitted;
              }
            } else if (sortMissingMode) {
              const inp = document.getElementById('sort-missing-input');
              if (inp) inp.value += inputChar;
            } else if (isFillMode) handleFillInput(inputChar);
            else {
              const targetInput = document.getElementById(targetInputId);
              if (targetInput) targetInput.value += inputChar;
            }
            
            if (!isKanaTypingMode && shiftMode === 1) { shiftMode = 0; updateKeyboardDisplay(); }
          };
          rowDiv.appendChild(keyBtn);
        });
        
        if(rIdx === 3) {
          const bsBtn = document.createElement('button'); bsBtn.className = 'key key-wide'; bsBtn.innerText = '⌫';
          bsBtn.onclick = () => { 
            if (isKanaTypingMode) {
              const targetInput = document.getElementById(targetInputId);
              if (targetInput) {
                if (kanaRomajiBuffer.length > 0) kanaRomajiBuffer = kanaRomajiBuffer.slice(0, -1);
                else kanaCommitted = kanaCommitted.slice(0, -1);
                targetInput.value = kanaCommitted;
              }
            } else if (sortMissingMode) {
              const inp = document.getElementById('sort-missing-input');
              if (inp) inp.value = inp.value.slice(0, -1);
            } else if(isFillMode) handleFillBackspace();
            else {
              const targetInput = document.getElementById(targetInputId);
              if (targetInput) {
                let val = targetInput.value;
                targetInput.value = val.slice(0, -1);
              }
            }
          };
          rowDiv.appendChild(bsBtn);
        }
        board.appendChild(rowDiv);
      });
      
      const bottomRow = document.createElement('div'); bottomRow.className = 'key-row';
      const spaceBtn = document.createElement('button'); spaceBtn.className = 'key'; spaceBtn.style.flex = "3"; spaceBtn.innerText = isKanaTypingMode ? "へんかん" : "Space";
      spaceBtn.onclick = () => {
        if (isKanaTypingMode) {
          const targetInput = document.getElementById(targetInputId);
          if (targetInput && kanaRomajiBuffer) {
            kanaCommitted += convertRomajiToKana(kanaRomajiBuffer, kanaIsKatakana);
            kanaRomajiBuffer = "";
            targetInput.value = kanaCommitted;
          }
        } else if (sortMissingMode) {
          const inp = document.getElementById('sort-missing-input');
          if (inp) inp.value += " ";
        } else if(isFillMode) handleFillInput(" ");
        else {
          const targetInput = document.getElementById(targetInputId);
          if (targetInput) targetInput.value += " ";
        }
      };
      
      if (sortMissingMode) {
        const cardBtn = document.createElement('button'); cardBtn.className = 'key key-action'; cardBtn.innerText = "カードにする";
        cardBtn.onclick = () => addSortMissingWordCard();
        bottomRow.appendChild(spaceBtn); bottomRow.appendChild(cardBtn);
      } else if (!isFillMode) {
        const enterBtn = document.createElement('button'); enterBtn.className = 'key key-action'; enterBtn.innerText = "けってい";
        enterBtn.onclick = () => {
          if (isKanaTypingMode) {
            const targetInput = document.getElementById(targetInputId);
            if (targetInput && kanaRomajiBuffer) {
              kanaCommitted += convertRomajiToKana(kanaRomajiBuffer, kanaIsKatakana);
              kanaRomajiBuffer = "";
              targetInput.value = kanaCommitted;
            }
          }
          if (typeof onEnterSubmit === "function") {
            onEnterSubmit();
            return;
          }
          const targetInput = document.getElementById(targetInputId);
          const userA = targetInput ? targetInput.value : "";
          const q = filteredQuestions[currentQuestionIndex];
          const format = document.getElementById('setting-format').value;
          let correctA = (format.includes("qtext") || format.includes("qaudio")) ? q["英文"] : (q["英単語"] || q["英文"]);
          checkAnswer(userA, correctA, q);
        };
        bottomRow.appendChild(spaceBtn); bottomRow.appendChild(enterBtn);
      } else {
        const fillSubmitBtn = document.createElement('button'); fillSubmitBtn.className = 'key key-action'; fillSubmitBtn.id = "fill-submit-btn"; fillSubmitBtn.innerText = "✨ これで回答する"; fillSubmitBtn.disabled = true;
        fillSubmitBtn.onclick = () => submitFillAnswer();
        bottomRow.appendChild(spaceBtn); bottomRow.appendChild(fillSubmitBtn);
      }
      
      board.appendChild(bottomRow);
      wrap.appendChild(board);
      container.appendChild(wrap);
      if (!isKanaTypingMode) updateKeyboardDisplay();
    }

    function getPenTargetInput() {
      const ansTypeEl = document.getElementById('setting-answer-type');
      const ansType = ansTypeEl ? ansTypeEl.value : '';
      if (ansType === "typing") {
        return document.getElementById('type-answer');
      }
      if (ansType === "fill_typing") {
        return null;
      }
      return null;
    }

    function composeWithAutoSpace(baseText, appendText) {
      const base = String(baseText || "");
      const add = String(appendText || "");
      if (!add) return base;
      if (!base) return add;
      const last = base.slice(-1);
      const first = add.charAt(0);
      if (last !== ' ' && first !== ' ') return base + ' ' + add;
      return base + add;
    }

    function appendTypeAnswerText(text) {
      const target = getPenTargetInput();
      if (target) {
        target.value = composeWithAutoSpace(target.value, text);
        target.classList.remove('type-answer-temp');
        handwritingState.pendingTempText = "";
        return;
      }
      const ansTypeEl = document.getElementById('setting-answer-type');
      const ansType = ansTypeEl ? String(ansTypeEl.value || '').trim() : '';
      if (ansType === 'fill_typing') {
        const chars = String(text || '').replace(/\s/g, '').split('');
        chars.forEach(ch => {
          if (/^[a-zA-Z]$/.test(ch)) handleFillInput(ch);
        });
      }
    }

    function setTypeAnswerTemporary(text) {
      const target = getPenTargetInput();
      if (!target) return;
      const before = target.value;
      target.value = composeWithAutoSpace(before, text);
      target.classList.add('type-answer-temp');
      handwritingState.pendingTempText = target.value.slice(before.length);
    }

    function confirmTemporaryTypingInput() {
      const target = getPenTargetInput();
      if (!target || !handwritingState.pendingTempText) return;
      target.classList.remove('type-answer-temp');
      handwritingState.pendingTempText = "";
      handwritingState.pendingCandidates = [];
      clearPenCanvas();
      renderPenConfirmBox();
    }

    function removeTemporaryTypingInput() {
      const target = getPenTargetInput();
      if (!target || !handwritingState.pendingTempText) return;
      if (target.value.endsWith(handwritingState.pendingTempText)) {
        target.value = target.value.slice(0, -handwritingState.pendingTempText.length);
      }
      target.classList.remove('type-answer-temp');
      handwritingState.pendingTempText = "";
    }

    function renderPenConfirmBox() {
      const wrap = document.getElementById('pen-confirm-box');
      if (!wrap) return;
      const hasTemp = !!handwritingState.pendingTempText;
      const cands = handwritingState.pendingCandidates || [];
      if (!hasTemp && cands.length === 0) {
        wrap.style.display = 'none';
        wrap.innerHTML = '';
        return;
      }
      wrap.style.display = 'block';
      if (hasTemp) {
        wrap.innerHTML = `
          <div style="font-size:14px;color:#ddd;">うまく変換されましたか？</div>
          <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;">
            <button type="button" class="submit-btn btn-green" style="margin-top:0;padding:8px 14px;font-size:14px;" onclick="confirmTemporaryTypingInput()">はい</button>
            <button type="button" class="submit-btn btn-gray" style="margin-top:0;padding:8px 14px;font-size:14px;" onclick="showAlternativeCandidates()">いいえ</button>
          </div>
        `;
        return;
      }
      const htmlCand = cands.map(c => `<button type="button" class="pen-candidate-btn" onclick="chooseAlternativeCandidate('${String(c).replace(/'/g, "\\'")}')">${escapeHtml(String(c))}</button>`).join('');
      wrap.innerHTML = `
        <div style="font-size:14px;color:#ddd;">次の候補にありますか？</div>
        <div class="pen-candidates">${htmlCand || '<span style="color:#999;">候補なし</span>'}</div>
        <div style="margin-top:8px;">
          <button type="button" class="submit-btn btn-gray" style="margin-top:0;padding:8px 14px;font-size:14px;" onclick="retryPenInput()">入力をやり直す</button>
        </div>
      `;
    }

    function showAlternativeCandidates() {
      removeTemporaryTypingInput();
      renderPenConfirmBox();
    }

    function chooseAlternativeCandidate(text) {
      appendTypeAnswerText(text);
      handwritingState.pendingCandidates = [];
      clearPenCanvas();
      renderPenConfirmBox();
    }

    function retryPenInput() {
      handwritingState.pendingCandidates = [];
      removeTemporaryTypingInput();
      clearPenCanvas();
      renderPenConfirmBox();
      setPenStatus('手書きをやり直してください。');
    }

    function clearTypingInput() {
      const target = getPenTargetInput();
      if (!target) return;
      target.value = "";
      target.classList.remove('type-answer-temp');
      handwritingState.pendingTempText = "";
      handwritingState.pendingCandidates = [];
      renderPenConfirmBox();
    }

    function resetHandwritingInputState() {
      handwritingState.pendingCandidates = [];
      handwritingState.pendingTempText = "";
      handwritingState.currentStroke = [];
      handwritingState.allStrokes = [];
      handwritingState.strokesBackupBeforeClear = null;
      handwritingState.isDrawing = false;
      handwritingState.pointerId = null;
      clearTypingInput();
      clearPenCanvas();
    }

    function setPenMode(mode) {
      handwritingState.penMode = mode;
      setUserPref('pen_mode', mode);
      const btnPen = document.getElementById('btn-pen-mode');
      const btnEraser = document.getElementById('btn-eraser-mode');
      if (btnPen) btnPen.classList.toggle('active', mode === 'pen');
      if (btnEraser) btnEraser.classList.toggle('active', mode === 'eraser');
    }

    function setPenWidth(width) {
      const w = parseInt(width, 10);
      handwritingState.penWidth = w;
      setUserPref('pen_width', w);
      const label = document.getElementById('pen-width-label');
      if (label) label.innerText = w;
      const canvas = document.getElementById('pen-canvas');
      if (canvas && canvas.dataset.ready === "1") {
        const ctx = canvas.getContext('2d');
        ctx.lineWidth = w;
      }
    }

    function setPenGuideSpread(val) {
      const n = Math.max(0, Math.min(100, parseInt(val, 10) || 0));
      handwritingState.guideLineSpread = n;
      setUserPref(LS_PEN_GUIDE_SPREAD, String(n));
      const label = document.getElementById('pen-guide-spread-label');
      if (label) label.textContent = n;
      redrawAllStrokes();
    }

    function setPenGuideShow(checked) {
      handwritingState.showGuideLines = !!checked;
      setUserPref(LS_PEN_GUIDE_SHOW, checked ? '1' : '0');
      redrawAllStrokes();
    }

    function scalePenStrokes(oldW, oldH, newW, newH) {
      if (!oldW || !oldH) return;
      if (oldW === newW && oldH === newH) return;
      const sx = newW / oldW, sy = newH / oldH;
      const mapStroke = (st) => { st.forEach(p => { p.x *= sx; p.y *= sy; }); };
      handwritingState.allStrokes.forEach(mapStroke);
      if (handwritingState.strokesBackupBeforeClear) {
        handwritingState.strokesBackupBeforeClear.forEach(mapStroke);
      }
      if (handwritingState.currentStroke && handwritingState.currentStroke.length) {
        mapStroke(handwritingState.currentStroke);
      }
    }

    function relayoutPenCanvas() {
      const canvas = document.getElementById('pen-canvas');
      if (!canvas || canvas.dataset.ready !== '1') return;
      const dpr = window.devicePixelRatio || 1;
      const rect = canvas.getBoundingClientRect();
      const w = rect.width, h = rect.height;
      canvas.width = w * dpr;
      canvas.height = h * dpr;
      const ctx = canvas.getContext('2d');
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.scale(dpr, dpr);
      ctx.lineWidth = handwritingState.penWidth;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = '#202124';
      canvas.dataset.logW = String(w);
      canvas.dataset.logH = String(h);
      redrawAllStrokes();
    }

    function setPenCanvasMaxWidth(val) {
      const px = Math.max(400, Math.min(1200, parseInt(val, 10) || 1000));
      handwritingState.canvasMaxWidthPx = px;
      setUserPref(LS_PEN_CANVAS_MAX_WIDTH, String(px));
      const wLabel = document.getElementById('pen-canvas-width-label');
      if (wLabel) wLabel.textContent = px;
      const wrap = document.getElementById('pen-canvas-wrap');
      if (wrap) wrap.style.maxWidth = px + 'px';
      const canvas = document.getElementById('pen-canvas');
      if (!canvas) return;
      const oldW = parseFloat(canvas.dataset.logW) || 0;
      const oldH = parseFloat(canvas.dataset.logH) || 0;
      if (canvas.dataset.ready === '1' && oldW > 0 && oldH > 0) {
        const rect = canvas.getBoundingClientRect();
        scalePenStrokes(oldW, oldH, rect.width, rect.height);
      }
      relayoutPenCanvas();
    }

    function undoPenStroke() {
      if (handwritingState.strokesBackupBeforeClear != null) {
        const bak = handwritingState.strokesBackupBeforeClear;
        handwritingState.strokesBackupBeforeClear = null;
        handwritingState.allStrokes = Array.isArray(bak) ? bak.map(s => s.map(p => ({ x: p.x, y: p.y, t: p.t }))) : [];
        redrawAllStrokes();
        return;
      }
      if (handwritingState.allStrokes.length > 0) {
        handwritingState.allStrokes.pop();
        redrawAllStrokes();
      }
    }

    function redrawAllStrokes() {
      const canvas = document.getElementById('pen-canvas');
      if (!canvas) return;
      const ctx = canvas.getContext('2d');
      const dpr = window.devicePixelRatio || 1;
      ctx.clearRect(0, 0, canvas.width / dpr, canvas.height / dpr);
      drawEnglishGuideLines();
      
      handwritingState.allStrokes.forEach(stroke => {
        if (stroke.length < 2) return;
        ctx.beginPath();
        ctx.moveTo(stroke[0].x, stroke[0].y);
        for (let i = 1; i < stroke.length; i++) {
          ctx.lineTo(stroke[i].x, stroke[i].y);
        }
        ctx.stroke();
      });
    }

    function getGuideLineYs() {
      const canvas = document.getElementById('pen-canvas');
      if (!canvas) return null;
      const dpr = window.devicePixelRatio || 1;
      const w = canvas.width / dpr;
      const h = canvas.height / dpr;
      const s = typeof handwritingState.guideLineSpread === 'number' ? handwritingState.guideLineSpread : 50;
      /* 最下位（0）付近で狭い間隔、最上位（100）で広い帯。従来の最低幅 0.45 相当はスライダー約 55 付近 */
      const span = 0.14 + (s / 100) * 0.56;
      const mid = 0.5;
      const y0 = h * (mid - span / 2);
      const y3 = h * (mid + span / 2);
      const step = (y3 - y0) / 3;
      return { w, ys: [y0, y0 + step, y0 + 2 * step, y3] };
    }

    function drawEnglishGuideLines() {
      if (!handwritingState.showGuideLines) return;
      const canvas = document.getElementById('pen-canvas');
      if (!canvas) return;
      const ctx = canvas.getContext('2d');
      const pack = getGuideLineYs();
      if (!pack) return;
      const { w, ys } = pack;
      ctx.save();
      ys.forEach((y, idx) => {
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(w, y);
        ctx.lineWidth = idx === 2 ? 2.2 : 1;
        ctx.strokeStyle = idx === 2 ? 'rgba(120,120,120,0.55)' : 'rgba(160,160,160,0.3)';
        ctx.stroke();
      });
      ctx.restore();
    }

    function distSqToSegment(p1, p2, pt) {
      const l2 = (p2.x - p1.x)**2 + (p2.y - p1.y)**2;
      if (l2 === 0) return (pt.x - p1.x)**2 + (pt.y - p1.y)**2;
      let t = ((pt.x - p1.x) * (p2.x - p1.x) + (pt.y - p1.y) * (p2.y - p1.y)) / l2;
      t = Math.max(0, Math.min(1, t));
      return (pt.x - (p1.x + t * (p2.x - p1.x)))**2 + (pt.y - (p1.y + t * (p2.y - p1.y)))**2;
    }

    function checkEraserCollision(pt) {
      const ERASER_RADIUS = 20;
      const ERASER_RADIUS_SQ = ERASER_RADIUS * ERASER_RADIUS;
      let hitIdx = -1;
      
      for (let i = handwritingState.allStrokes.length - 1; i >= 0; i--) {
        const stroke = handwritingState.allStrokes[i];
        let hit = false;
        for (let j = 0; j < stroke.length - 1; j++) {
          if (distSqToSegment(stroke[j], stroke[j+1], pt) <= ERASER_RADIUS_SQ) {
            hit = true;
            break;
          }
        }
        if (!hit && stroke.length === 1) {
          if ((stroke[0].x - pt.x)**2 + (stroke[0].y - pt.y)**2 <= ERASER_RADIUS_SQ) {
            hit = true;
          }
        }
        if (hit) {
          hitIdx = i;
          break;
        }
      }
      
      if (hitIdx !== -1) {
        handwritingState.allStrokes.splice(hitIdx, 1);
        redrawAllStrokes();
      }
    }

    function clearPenCanvas() {
      const canvas = document.getElementById('pen-canvas');
      if (!canvas) return;
      if (handwritingState.allStrokes.length > 0) {
        handwritingState.strokesBackupBeforeClear = handwritingState.allStrokes.map(s => s.map(p => ({ x: p.x, y: p.y, t: p.t })));
      }
      const ctx = canvas.getContext('2d');
      const dpr = window.devicePixelRatio || 1;
      ctx.clearRect(0, 0, canvas.width / dpr, canvas.height / dpr);
      drawEnglishGuideLines();
      handwritingState.currentStroke = [];
      handwritingState.allStrokes = [];
      handwritingState.isDrawing = false;
      handwritingState.pointerId = null;
      const status = document.getElementById('pen-status');
      if (status) status.innerText = "";
    }

    function setPenStatus(msg) {
      const status = document.getElementById('pen-status');
      if (status) status.innerText = msg || "";
    }

    function refreshPenStatusHint() {
      const status = document.getElementById('pen-status');
      if (!status) return;
      if (status.innerText && status.innerText !== '手書き入力の準備OK') return;
      status.innerText = isIpadStylusOptimizationEnabled()
        ? "iPad書きやすさ優先モード（誤タッチ防止を弱めています）"
        : "手書きペンのみ入力可（パームリジェクション有効）";
    }

    function shouldAcceptPenPointer(e) {
      if (!e) return false;
      if (e.pointerType === 'pen') return true;
      if (!isIpadStylusOptimizationEnabled()) return false;
      if (e.pointerType !== 'touch') return false;
      if (e.isPrimary === false) return false;
      const w = Number(e.width) || 0;
      const h = Number(e.height) || 0;
      const area = w * h;
      const pressure = Number(e.pressure) || 0;
      // iPad書き味優先: 細い接触面・圧力あり・サイズ不明は許容
      if (area === 0) return true;
      if (area <= 260) return true;
      if (pressure >= 0.15) return true;
      return false;
    }

    function getPenCoords(e, canvas) {
      const rect = handwritingState.activePenRect || canvas.getBoundingClientRect();
      return {
        x: e.clientX - rect.left,
        y: e.clientY - rect.top,
        t: Date.now()
      };
    }

    function initializePenCanvas() {
      const canvas = document.getElementById('pen-canvas');
      if (!canvas || canvas.dataset.ready === "1") return;
      
      const dpr = window.devicePixelRatio || 1;
      const rect = canvas.getBoundingClientRect();
      canvas.width = rect.width * dpr;
      canvas.height = rect.height * dpr;
      
      const ctx = canvas.getContext('2d');
      ctx.scale(dpr, dpr);
      
      ctx.lineWidth = handwritingState.penWidth;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = '#202124';
      drawEnglishGuideLines();

      canvas.addEventListener('pointerdown', (e) => {
        if (!shouldAcceptPenPointer(e)) {
          setPenStatus(isIpadStylusOptimizationEnabled() ? 'iPad書きやすさ優先モード: 大きい接触の誤タッチは無効です' : '手書きペンで入力してください（手・マウスは無効）');
          return;
        }
        e.preventDefault();
        const useIpadOpt = isIpadStylusOptimizationEnabled();
        handwritingState.activePenRect = useIpadOpt ? canvas.getBoundingClientRect() : null;
        if (canvas.setPointerCapture) {
          try { canvas.setPointerCapture(e.pointerId); } catch (_) {}
        }
        const pt = getPenCoords(e, canvas);

        if (handwritingState.penMode === 'eraser') {
          checkEraserCollision(pt);
          handwritingState.isDrawing = true;
          handwritingState.pointerId = e.pointerId;
          return;
        }

        handwritingState.isDrawing = true;
        handwritingState.currentStroke = [];
        handwritingState.pointerId = e.pointerId;
        handwritingState.pointerType = e.pointerType;
        handwritingState.currentStroke.push(pt);
        ctx.beginPath();
        ctx.moveTo(pt.x, pt.y);
      });

      canvas.addEventListener('pointermove', (e) => {
        if (!handwritingState.isDrawing || handwritingState.pointerId !== e.pointerId) return;
        if (!shouldAcceptPenPointer(e)) return;
        e.preventDefault();
        const pt = getPenCoords(e, canvas);

        if (handwritingState.penMode === 'eraser') {
          checkEraserCollision(pt);
          return;
        }

        handwritingState.currentStroke.push(pt);
        ctx.lineTo(pt.x, pt.y);
        ctx.stroke();
      });

      const stopPenStroke = (e) => {
        if (!handwritingState.isDrawing || handwritingState.pointerId !== e.pointerId) return;
        if (handwritingState.penMode === 'pen' && handwritingState.currentStroke.length > 1) {
          handwritingState.allStrokes.push(handwritingState.currentStroke);
          handwritingState.strokesBackupBeforeClear = null;
        }
        if (canvas.releasePointerCapture) {
          try { canvas.releasePointerCapture(e.pointerId); } catch (_) {}
        }
        handwritingState.isDrawing = false;
        handwritingState.currentStroke = [];
        handwritingState.pointerId = null;
        handwritingState.activePenRect = null;
      };
      canvas.addEventListener('pointerup', stopPenStroke);
      canvas.addEventListener('pointercancel', stopPenStroke);
      canvas.addEventListener('pointerleave', (e) => {
        if (!handwritingState.isDrawing || handwritingState.pointerId !== e.pointerId) return;
        // iPad では高速筆記時に pointerleave が先に飛ぶことがあるため、
        // ペンがまだ接地中（buttons !== 0）は描画を継続する。
        if (e.buttons !== 0) return;
        stopPenStroke(e);
      });
      canvas.dataset.ready = "1";
      const r0 = canvas.getBoundingClientRect();
      canvas.dataset.logW = String(r0.width);
      canvas.dataset.logH = String(r0.height);
      setPenStatus('手書き入力の準備OK');
      refreshPenStatusHint();
    }

    function renderInputModeUi(targetMode) {
      const area = document.getElementById('quiz-answer-area');
      if (!area) return;
      const panel = area.querySelector('#answer-input-mode-switch');
      if (!panel) return;
      panel.querySelectorAll('.answer-mode-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.mode === targetMode);
      });
    }

    function buildTypingInputAreaMarkup(useFillMode) {
      return `
        <div id="answer-input-mode-switch" class="answer-input-mode-switch">
          <button type="button" class="answer-mode-btn" data-mode="keyboard" onclick="switchInputMethod('keyboard')">⌨️ キーボード</button>
          <button type="button" class="answer-mode-btn" data-mode="pen" onclick="switchInputMethod('pen')">✍️ 手書き</button>
        </div>
        <div id="input-method-body"></div>
      `;
    }

    function mountTypingMethodBody(useFillMode) {
      const body = document.getElementById('input-method-body');
      if (!body) return;
      if (inputMethodMode === 'pen') {
        const submitBtnHtml = useFillMode
          ? `<button type="button" id="fill-submit-btn" class="submit-btn btn-green" style="margin-top:0;padding:10px 20px;font-size:16px;" onclick="submitFillAnswer()" disabled>これで回答</button>`
          : '';
        const typingPreviewHtml = useFillMode
          ? `
            <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">
              <button type="button" id="pen-recognize-btn" class="submit-btn btn-blue pen-recognize-btn" style="margin-top:0;padding:10px 14px;font-size:14px;" onclick="recognizePenStrokes()">文字起こし</button>
              <span style="color:#bbb;font-size:13px;line-height:1.4;">かいたあと「文字起こし」でマスに入ります（サーバーへ送ります）</span>
            </div>
          `
          : `
            <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">
              <button type="button" id="pen-recognize-btn" class="submit-btn btn-blue pen-recognize-btn" style="margin-top:0;padding:10px 14px;font-size:14px;" onclick="recognizePenStrokes()">文字起こし</button>
              <button type="button" class="submit-btn btn-gray" style="margin-top:0;padding:10px 14px;font-size:14px;" onclick="appendTypeAnswerText(' ')">間をあける</button>
              <input type="text" id="type-answer" class="large-input" readonly placeholder="認識結果が入ります" style="margin:0;flex:1;min-width:260px;">
              <button type="button" id="pen-typing-submit-btn" class="submit-btn btn-green" style="margin-top:0;padding:10px 14px;font-size:14px;" onclick="submitTypingFromPen()">答えを送信</button>
            </div>
          `;
        const activePen = handwritingState.penMode === 'pen' ? 'active' : '';
        const activeEraser = handwritingState.penMode === 'eraser' ? 'active' : '';
        const wMax = handwritingState.canvasMaxWidthPx;
        body.innerHTML = `
          <div class="pen-panel">
            <div class="pen-settings pen-tool-row">
              <div class="pen-toggle-switch">
                <button type="button" id="btn-pen-mode" class="${activePen}" onclick="setPenMode('pen')">手書き</button>
                <button type="button" id="btn-eraser-mode" class="${activeEraser}" onclick="setPenMode('eraser')">線消しゴム</button>
              </div>
              <button type="button" id="pen-advanced-toggle-btn" class="cancel-btn" onclick="togglePenAdvancedSettings()" style="padding:4px 10px; font-size:14px; background:#1E88E5; color:#fff; border-radius:20px; border:1px solid #1565C0;">詳細設定ボタンを表示</button>
              <button type="button" class="cancel-btn" onclick="undoPenStroke()" style="padding:4px 10px; font-size:14px; background:#444; color:#fff; border-radius:20px; border:1px solid #555;">↩️ 1つ戻す</button>
              <button type="button" class="cancel-btn" onclick="clearPenCanvas()" style="padding:4px 10px; font-size:14px; background:#444; color:#fff; border-radius:20px; border:1px solid #555;">手書きを全部消す</button>
              <button type="button" class="cancel-btn" onclick="clearTypingInput()" style="padding:4px 10px; font-size:14px; background:#444; color:#fff; border-radius:20px; border:1px solid #555;">テキスト入力を全部消す</button>
            </div>
            <div id="pen-advanced-controls" class="pen-canvas-width-row" style="display:none; justify-content:center; align-items:center; flex-wrap:wrap; gap:10px;">
              <label style="color:#ccc; font-size:14px; display:flex; align-items:center; gap:5px; background:#222; padding:4px 8px; border-radius:20px;">
                太さ: <input type="range" id="pen-width-range" min="1" max="10" value="${handwritingState.penWidth}" onchange="setPenWidth(this.value)">
                <span id="pen-width-label" style="display:inline-block; width:16px; text-align:center;">${handwritingState.penWidth}</span>px
              </label>
              <label style="display:inline-flex;align-items:center;gap:8px;flex-wrap:wrap;">手書きの幅
                <input type="range" id="pen-canvas-width-range" min="400" max="1200" step="20" value="${wMax}" oninput="setPenCanvasMaxWidth(this.value)" style="width:min(240px,50vw);">
                <span id="pen-canvas-width-label">${wMax}</span> px
              </label>
              <label style="color:#ccc;font-size:13px;display:inline-flex;align-items:center;gap:6px;flex-wrap:wrap;max-width:100%;">目印の間隔
                <input type="range" id="pen-guide-spread" class="pen-guide-hslider" min="0" max="100" value="${handwritingState.guideLineSpread}" oninput="setPenGuideSpread(this.value)" title="4本の目印の線の間隔">
                <span id="pen-guide-spread-label" style="min-width:22px;">${handwritingState.guideLineSpread}</span>
              </label>
              <label style="color:#ccc; font-size:13px; display:flex; align-items:center; gap:6px; background:#222; padding:4px 10px; border-radius:20px; cursor:pointer;">
                <input type="checkbox" id="pen-guide-show" ${handwritingState.showGuideLines ? 'checked' : ''} onchange="setPenGuideShow(this.checked)"> 目印の線を表示する
              </label>
              <label style="color:#ccc; font-size:13px; display:flex; align-items:center; gap:6px; background:#222; padding:4px 10px; border-radius:20px; cursor:pointer;">
                <input type="checkbox" id="pen-ipad-stylus-opt" ${isIpadStylusOptimizationEnabled() ? 'checked' : ''} onchange="syncIpadStylusSettingsFromCheckbox(this)"> iPad用のスタイラスペン設定にする
              </label>
            </div>
            ${typingPreviewHtml}
            <div id="pen-confirm-box" class="pen-confirm-box" style="display:none;"></div>
            <div class="pen-canvas-wrap" id="pen-canvas-wrap" style="max-width:${wMax}px;margin:0 auto;">
              <canvas id="pen-canvas" class="pen-canvas" oncontextmenu="event.preventDefault();return false" style="width:100%;height:min(50vh, 520px);max-height:85vh;touch-action:none;"></canvas>
            </div>
            ${useFillMode ? `<div class="pen-panel-controls" style="display:flex;align-items:center;flex-wrap:wrap;gap:12px;justify-content:center;">${submitBtnHtml}</div>` : ``}
            <div id="pen-status" class="pen-status">手書きペンのみ入力可（パームリジェクション有効）</div>
          </div>
        `;
        initializePenCanvas();
        renderPenConfirmBox();
        syncPenAdvancedVisibility();
      } else {
        if (useFillMode) {
          body.innerHTML = `<div id="keyboard-container"></div>`;
          renderCustomKeyboard(true);
        } else {
          body.innerHTML = `<input type="text" id="type-answer" class="large-input" readonly placeholder="キーボードで入力"><div id="keyboard-container"></div>`;
          renderCustomKeyboard(false);
        }
      }
      renderInputModeUi(inputMethodMode);
      syncQuizAnswerAreaPenModeClass();
    }

    function syncQuizAnswerAreaPenModeClass() {
      const a = document.getElementById('quiz-answer-area');
      const sec = document.getElementById('section-quiz');
      const isPen = inputMethodMode === 'pen';
      if (a) a.classList.toggle('pen-mode-active', isPen);
      if (sec) sec.classList.toggle('pen-mode-active', isPen);
    }

    let isPenTypingSubmitInFlight = false;
    function setPenTypingSubmitBusy(isBusy) {
      isPenTypingSubmitInFlight = !!isBusy;
      const btn = document.getElementById('pen-typing-submit-btn');
      if (!btn) return;
      btn.disabled = !!isBusy;
      btn.setAttribute('aria-busy', isBusy ? 'true' : 'false');
      btn.innerText = isBusy ? '送信中...' : '答えを送信';
    }
    function submitTypingFromPen() {
      if (isPenTypingSubmitInFlight) return;
      if (handwritingState.pendingTempText) confirmTemporaryTypingInput();
      const target = document.getElementById('type-answer');
      const userA = target ? String(target.value || "") : "";
      if (!normalizeText(userA)) {
        setPenStatus('先に文字起こしして入力してください。');
        return;
      }
      if (!Array.isArray(filteredQuestions) || currentQuestionIndex < 0 || currentQuestionIndex >= filteredQuestions.length) {
        setPenStatus('問題データの読み込み待ちです。少し待って再度お試しください。');
        return;
      }
      const q = filteredQuestions[currentQuestionIndex];
      if (!q) {
        setPenStatus('問題の取得に失敗しました。次の問題に進んでください。');
        return;
      }
      const formatEl = document.getElementById('setting-format');
      const format = formatEl ? String(formatEl.value || "") : "";
      const correctA = (format.includes("qtext") || format.includes("qaudio")) ? q["英文"] : (q["英単語"] || q["英文"]);
      setPenTypingSubmitBusy(true);
      checkAnswer(userA, correctA, q);
    }

    function switchInputMethod(mode) {
      inputMethodMode = mode === 'pen' ? 'pen' : 'keyboard';
      setUserPref(LS_INPUT_MODE, inputMethodMode);
      const ansTypeEl = document.getElementById('setting-answer-type');
      const ansType = ansTypeEl ? String(ansTypeEl.value || '').trim() : '';
      if (ansType !== "typing" && ansType !== "fill_typing") return;
      mountTypingMethodBody(ansType === "fill_typing");
      const answerArea = document.getElementById('quiz-answer-area');
      if (answerArea) placeQuizFeedbackAboveKeyboard(answerArea);
    }

    function getPenWritingGuideForApi() {
      const canvas = document.getElementById('pen-canvas');
      if (!canvas) return { width: 1000, height: 400 };
      const w = parseFloat(canvas.dataset.logW) || canvas.getBoundingClientRect().width || 1000;
      const h = parseFloat(canvas.dataset.logH) || canvas.getBoundingClientRect().height || 400;
      return {
        width: Math.round(Math.max(200, Math.min(2400, w))),
        height: Math.round(Math.max(200, Math.min(2400, h)))
      };
    }

    function collectInkForApi() {
      const apiInk = [];
      const MAX_POINTS_PER_STROKE = 180;
      handwritingState.allStrokes.forEach(stroke => {
        if (!stroke || stroke.length < 1) return;
        let strokeForApi = stroke;
        if (stroke.length === 1) {
          const p0 = stroke[0];
          strokeForApi = [p0, { x: p0.x, y: p0.y, t: (p0.t != null ? p0.t : Date.now()) + 1 }];
        }
        let sampled = strokeForApi;
        if (strokeForApi.length > MAX_POINTS_PER_STROKE) {
          const step = Math.ceil(strokeForApi.length / MAX_POINTS_PER_STROKE);
          sampled = strokeForApi.filter((_, idx) => idx % step === 0);
          if (sampled[sampled.length - 1] !== strokeForApi[strokeForApi.length - 1]) {
            sampled.push(strokeForApi[strokeForApi.length - 1]);
          }
        }
        const xArr = [];
        const yArr = [];
        const tArr = [];
        const t0 = (typeof sampled[0].t === 'number' && !isNaN(sampled[0].t)) ? sampled[0].t : 0;
        sampled.forEach((p, i) => {
          xArr.push(p.x);
          yArr.push(p.y);
          const t = (typeof p.t === 'number' && !isNaN(p.t)) ? p.t : t0 + i * 16;
          tArr.push(Math.max(0, t - t0));
        });
        apiInk.push([xArr, yArr, tArr]);
      });
      return apiInk;
    }

    function setPenRecognizeButtonState(isBusy) {
      const btn = document.getElementById('pen-recognize-btn');
      if (!btn) return;
      btn.classList.toggle('is-busy', !!isBusy);
      btn.disabled = !!isBusy;
      btn.setAttribute('aria-busy', isBusy ? 'true' : 'false');
      btn.innerText = isBusy ? '⏳ 文字起こし中...' : '文字起こし';
    }

    function recognizePenStrokes() {
      if (inputMethodMode !== 'pen') return;
      const ansTypeEl = document.getElementById('setting-answer-type');
      const ansType = ansTypeEl ? String(ansTypeEl.value || '').trim() : '';
      if (ansType !== "typing" && ansType !== "fill_typing") return;
      if (isPenRecognitionInFlight) {
        setPenStatus('文字起こし中です。少し待ってください。');
        return;
      }
      const apiInk = collectInkForApi();
      if (apiInk.length === 0) {
        setPenStatus('先に手書きで文字を書いてください。');
        return;
      }
      const writingGuide = getPenWritingGuideForApi();
      isPenRecognitionInFlight = true;
      setPenRecognizeButtonState(true);
      setPenStatus('文字起こし中（サーバーと通信）… 画数: ' + String(apiInk.length));
      const ac = new AbortController();
      const PEN_RECOGNIZE_TIMEOUT_MS = 38000;
      const timeoutId = setTimeout(() => ac.abort(), PEN_RECOGNIZE_TIMEOUT_MS);
      function postRecognizeWithRetry(retryCount) {
        return fetch(GAS_API_URL, {
          method: 'POST',
          body: JSON.stringify({ action: "recognize_handwriting", ink: apiInk, writingGuide: writingGuide }),
          signal: ac.signal
        })
        .then(r => {
          if (!r.ok) return Promise.reject(new Error('HTTP ' + r.status));
          return r.json();
        })
        .catch(err => {
          if (err && err.name === 'AbortError') throw err;
          const msg = String((err && (err.message || err)) || "");
          const canRetry = retryCount > 0 && /Load failed|Failed to fetch|NetworkError|fetch/i.test(msg);
          if (!canRetry) throw err;
          setPenStatus('通信が不安定です。文字起こしを再試行中…');
          return new Promise(resolve => setTimeout(resolve, 450)).then(() => postRecognizeWithRetry(retryCount - 1));
        });
      }
      postRecognizeWithRetry(1)
      .then(d => {
        if (d.status !== "success") {
          setPenStatus(d.message || "認識に失敗しました。");
          return;
        }
        const text = String(
          d.text != null ? d.text :
          d.result != null ? d.result :
          d.recognizedText != null ? d.recognizedText :
          d.bestCandidate != null ? d.bestCandidate :
          d.candidate != null ? d.candidate :
          ""
        ).trim();
        if (!text) {
          const keys = d && typeof d === "object" ? Object.keys(d).slice(0, 8).join(", ") : "";
          setPenStatus("認識結果なし。もう一度書いてください。返却キー: " + (keys || "なし"));
          return;
        }
        const rawCandidates =
          Array.isArray(d.candidates) ? d.candidates :
          Array.isArray(d.alternatives) ? d.alternatives :
          Array.isArray(d.results) ? d.results :
          [];
        const candidates = rawCandidates.map(v => String(v || '').trim()).filter(v => v);
        handwritingState.pendingCandidates = candidates.filter(v => v !== text);
        const target = getPenTargetInput();
        if (target) {
          removeTemporaryTypingInput();
          setTypeAnswerTemporary(text);
          renderPenConfirmBox();
        } else {
          const chars = text.replace(/\s/g, '').split('');
          chars.forEach(ch => {
            if (/^[a-zA-Z]$/.test(ch)) handleFillInput(ch);
          });
        }
        setPenStatus(`仮入力: ${text}`);
      })
      .catch(err => {
        if (err && err.name === 'AbortError') {
          setPenStatus('時間がかかりすぎて止めました。通信やGASの状態を確認し、もう一度「文字起こし」を押してください。');
        } else {
          const msg = String((err && (err.message || err)) || "");
          setPenStatus('通信エラー（' + (msg || '応答を読めませんでした') + '）。ネットワークを確認して再度お試しください。');
        }
      })
      .finally(() => {
        clearTimeout(timeoutId);
        isPenRecognitionInFlight = false;
        setPenRecognizeButtonState(false);
      });
    }

    function startVoiceInput() {
      if(!recognition) { alert("お使いのブラウザは音声入力に対応していません。SafariかChromeを使ってね！"); return; }
      const feedback = document.getElementById('voice-feedback'); const btn = document.getElementById('voice-btn'); const displayField = document.getElementById('voice-recognized-text');
      displayField.value = ""; btn.innerText = "🎙️ 聞き取り中..."; btn.style.background = "#e50914"; feedback.innerHTML = "英語ではなしてください...";
      recognition.onresult = (event) => {
        let interimTranscript = ''; let finalTranscript = '';
        for (let i = event.resultIndex; i < event.results.length; ++i) { if (event.results[i].isFinal) finalTranscript += event.results[i][0].transcript; else interimTranscript += event.results[i][0].transcript; }
        let rawTranscript = finalTranscript || interimTranscript; 
        const displayText = convertDigitsToWordsInText(rawTranscript);
        displayField.value = displayText;
        if (displayText) {
          const q = filteredQuestions[currentQuestionIndex]; const format = document.getElementById('setting-format').value; let correctA = (format.includes("qtext") || format.includes("qaudio")) ? q["英文"] : (q["英単語"] || q["英文"]);
          if(normalizeText(displayText) === normalizeText(correctA)) { 
            recognition.stop(); displayField.value = displayText; checkAnswer(displayText, correctA, q); 
          } 
          else if (finalTranscript) { 
            voiceFailCount++;
            if (voiceFailCount >= 5) {
              recognition.stop();
              displayField.value = displayText;
              checkAnswer(displayText, correctA, q);
            } else {
              displayField.value = displayText; 
              feedback.innerHTML = `聞き取った言葉：<br><b style='color:#FF9800;'>「${displayText}」</b><br>ちがうみたい！もう一度🎙️をおしてね。<br><span style="font-size:14px; color:#F44336;">（まちがい: ${voiceFailCount}回 / 5回で不正解）</span>`; 
              btn.innerText = "🎙️ マイクでこたえる"; btn.style.background = "#2196F3"; 
            }
          }
        }
      };
      recognition.onerror = () => { feedback.innerText = "聞き取れませんでした。もう一度おしてね。"; btn.innerText = "🎙️ マイクでこたえる"; btn.style.background = "#2196F3"; };
      recognition.onend = () => { if(btn.innerText.includes("聞き取り中")) { btn.innerText = "🎙️ マイクでこたえる"; btn.style.background = "#2196F3"; } };
      recognition.start();
    }

    function startJaSpeechToField(fieldId) {
      const target = document.getElementById(fieldId);
      if (!target) return;
      if (!('SpeechRecognition' in window || 'webkitSpeechRecognition' in window)) { alert("お使いのブラウザは日本語音声入力に対応していません。Chrome/Edgeを使ってね。"); return; }
      if (!recognitionJa) {
        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
        recognitionJa = new SpeechRecognition();
        recognitionJa.lang = 'ja-JP';
        recognitionJa.interimResults = false;
        recognitionJa.continuous = false;
      }
      const btnText = target.dataset && target.dataset.listeningLabel ? target.dataset.listeningLabel : "🎙️ 日本語で入力";
      recognitionJa.onresult = (event) => {
        let finalTranscript = '';
        for (let i = event.resultIndex; i < event.results.length; ++i) { if (event.results[i].isFinal) finalTranscript += event.results[i][0].transcript; }
        if (finalTranscript) {
          target.value = (target.value ? target.value + " " : "") + finalTranscript;
        }
      };
      recognitionJa.onerror = () => { alert("聞き取れませんでした。もう一度試してください。"); };
      recognitionJa.onend = () => {};
      recognitionJa.start();
    }

    function prepareQuiz() {
      currentIsReviewMode = document.getElementById('setting-play-mode').value === 'review';
      const order = document.getElementById('setting-order').value; 
      const format = document.getElementById('setting-format').value;
      const ansType = syncAnswerTypeWithFormat(format);

      filteredQuestions = [...currentQuestions];

      if (format === "ja_to_en_sort") {
        filteredQuestions = filteredQuestions.filter(q => {
          const toks = parseSortPhraseTokens(q);
          const slot = String(q["並び替え箇所"] || "").trim();
          if (toks.length < 2 || !slot) return false;
          if (ansType === "sort_dummy") {
            const d = getDummyTokenForSort(q);
            if (!d) return false;
          }
          return true;
        });
        if (filteredQuestions.length === 0) {
          alert("並び替えのデータが足りません。\n「並び替え箇所」と「並び替え語句1」…を確認してください。\n不要語混入の場合は「並び替え語句ダミー」も必要です。");
          return;
        }
      }

      if (ansType.startsWith("fill_")) {
        const numBlanks = parseInt(document.getElementById('setting-blank-count').value) || 1;
        filteredQuestions = filteredQuestions.filter(q => {
          let w = (q["英単語"] || q["英文"] || "").trim();
          return w.length > numBlanks;
        });
        if (filteredQuestions.length === 0) {
          alert("選択した文字数を隠せる単語がありません！隠す数を減らしてください。");
          return;
        }
      }

      if (currentIsReviewMode) {
        const user = JSON.parse(localStorage.getItem('app_kid_user')); 
        const unitKey = getDetailedUnitId(); 
        const history = user.historyJson[unitKey] || {};
        filteredQuestions = filteredQuestions.filter(q => checkIsNigate(history[q["通し番号"]]));
        if (filteredQuestions.length === 0) { alert("ニガテな問題はありません！通常モードでスタートします。"); currentIsReviewMode = false; document.getElementById('setting-play-mode').value = 'normal'; filteredQuestions = [...currentQuestions]; }
      }
      if (order === 'random') filteredQuestions.sort(() => Math.random() - 0.5);
      currentQuestionIndex = 0; quizResults = []; switchSection('section-quiz'); saveQuizRecoveryDraft(0); showQuestion();
    }

    /** 正誤表示を「つぎへ」ボタンの直前の定位置へ戻す（answerArea を空にする前に必須） */
    function restoreQuizFeedbackLocation() {
      const fb = document.getElementById('quiz-feedback');
      const nextBtn = document.getElementById('quiz-next-btn');
      if (!fb || !nextBtn || !nextBtn.parentNode) return;
      if (fb.nextElementSibling !== nextBtn) {
        nextBtn.parentNode.insertBefore(fb, nextBtn);
      }
    }

    /** 正誤表示は常に「入力欄の直下」優先で配置する（手書き/キーボード両対応） */
    function placeQuizFeedbackAboveKeyboard(answerAreaEl) {
      const fb = document.getElementById('quiz-feedback');
      if (!fb || !answerAreaEl) return;

      const typeInput = answerAreaEl.querySelector('#type-answer');
      if (typeInput) {
        const hostRow = typeInput.closest('div');
        const hostParent = hostRow && hostRow.parentNode ? hostRow.parentNode : typeInput.parentNode;
        if (hostParent) {
          hostParent.insertBefore(fb, hostRow ? hostRow.nextSibling : typeInput.nextSibling);
          return;
        }
      }

      const kbd = answerAreaEl.querySelector('#keyboard-container');
      if (kbd) {
        answerAreaEl.insertBefore(fb, kbd);
        return;
      }

      answerAreaEl.appendChild(fb);
    }

    function showQuestion() {
      if (currentQuestionIndex >= filteredQuestions.length) { finishQuiz(); return; }
      setPenTypingSubmitBusy(false);
      maxDeduction = 0;
      voiceFailCount = 0;
      clearTimeout(autoNextTimer); 
      destroySortQuizIfAny();
      
      const q = filteredQuestions[currentQuestionIndex]; 
      const format = document.getElementById('setting-format').value; 
      const answerType = syncAnswerTypeWithFormat(format);
      const isWord = currentModeName.includes("単語");

      const kbdPanel = document.getElementById('quiz-keyboard-settings-panel');
      if (kbdPanel) {
        kbdPanel.style.display = (answerType === 'typing' || answerType === 'fill_typing' || (format === 'ja_to_en_sort' && answerType === 'sort_missing')) ? 'block' : 'none';
        const kbdBody = document.getElementById('quiz-keyboard-settings-body');
        if (kbdBody) kbdBody.style.display = 'none';
      }
      
      document.getElementById('quiz-progress').innerText = (isTrainingMode?"🎯 特訓ルート: ":"") + (currentIsReviewMode?"🔥特訓中! ":"") + `第 ${currentQuestionIndex + 1} 問 / 全 ${filteredQuestions.length} 問`;
      
      let qText = "", correctA = "", engTextForAudio = "";
      if (format === "ja_to_en") { qText = q["日本語"]; correctA = (q["英単語"] || q["英文"]).trim(); } 
      else if (format === "en_to_ja") { qText = (q["英単語"] || q["英文"]).trim(); correctA = q["日本語"]; }
      else if (format === "qtext_to_en") { qText = q["疑問文"]; correctA = q["英文"].trim(); }
      else if (format === "en_audio_to_ja") { engTextForAudio = q["英単語"] || q["英文"]; correctA = q["日本語"]; }
      else if (format === "qaudio_to_en") { engTextForAudio = q["疑問文"]; correctA = q["英文"]; }
      else if (format === "en_audio_to_en") { engTextForAudio = q["英単語"] || q["英文"]; correctA = (q["英単語"] || q["英文"]).trim(); }
      else if (format === "en_to_en") { qText = (q["英単語"] || q["英文"]).trim(); engTextForAudio = qText; correctA = qText; }

      if (format === "ja_to_en_sort") {
        const ja = String(q["日本語"] || "");
        const refEn = String(q["並び替え用英文"] || "").trim();
        document.getElementById('quiz-q-text').innerHTML = `<div style="font-size:36px;font-weight:bold;margin-bottom:8px;">${escapeHtml(ja)}</div>${refEn ? `<div style="font-size:22px;color:#aaa;line-height:1.4;">${escapeHtml(refEn)}</div>` : ''}`;
      } else if (format.includes("audio")) {
        document.getElementById('quiz-q-text').innerHTML = `<button onclick="speakText('${engTextForAudio.replace(/'/g, "\\'")}')" class="submit-btn btn-blue" style="border-radius:50px;">🔊 英語をきく</button>`;
        setTimeout(() => speakText(engTextForAudio), 500); 
      } else if (format === "en_to_en") {
        const safeText = qText || "";
        document.getElementById('quiz-q-text').innerHTML = `<div style="margin-bottom:8px;">${safeText}</div><button onclick="speakText('${safeText.replace(/'/g, "\\'")}')" class="submit-btn btn-blue" style="border-radius:50px;">🔊 英語をきく</button>`;
      } else {
        document.getElementById('quiz-q-text').innerText = qText; 
      }

      document.getElementById('quiz-feedback').innerHTML = ""; 
      document.getElementById('quiz-next-btn').style.display = "none";
      
      const hintArea = document.getElementById('quiz-hint-area');
      const hintTextDisplay = document.getElementById('hint-display-text');
      const wordBlankArea = document.getElementById('quiz-word-blank-area');
      hintTextDisplay.style.display = "none"; hintTextDisplay.innerText = "";
      wordBlankArea.style.display = "none"; wordBlankArea.innerHTML = "";
      
      document.getElementById('hint-btn-2').innerText = isWord ? "💡 イニシャルと文字数を見る（-7 Pt）" : "💡 ヒントを見る（-7 Pt）";

      if (answerType === "4choice" || format.includes("en_to_ja") || format === "en_audio_to_ja" || answerType.startsWith("fill_")) {
        hintArea.style.display = "none";
      } else {
        hintArea.style.display = "flex";
        document.getElementById('hint-btn-1').disabled = false;
        document.getElementById('hint-btn-2').disabled = false;
        document.getElementById('hint-btn-3').disabled = false;
      }

      restoreQuizFeedbackLocation();
      const answerArea = document.getElementById('quiz-answer-area');
      answerArea.classList.remove('pen-mode-active');
      const quizSec = document.getElementById('section-quiz');
      if (quizSec) quizSec.classList.remove('pen-mode-active');
      answerArea.innerHTML = ""; questionStartTime = Date.now();
      resetHandwritingInputState();
      saveQuizRecoveryDraft(currentQuestionIndex);
      
      if (format === "ja_to_en_sort" && (answerType === "sort_all" || answerType === "sort_dummy" || answerType === "sort_missing")) {
        setupSortQuiz(q, answerType);
      }
      else if (answerType === "4choice") {
        let choices = [correctA]; 
        let others = currentQuestions.filter(item => { if(format === "ja_to_en" || format === "qtext_to_en" || format === "qaudio_to_en" || format === "en_audio_to_en" || format === "en_to_en") return (item["英単語"] || item["英文"]) !== correctA; else return item["日本語"] !== correctA; }).sort(() => Math.random() - 0.5);
        for (let i = 0; i < 3 && i < others.length; i++) { choices.push((format === "ja_to_en" || format.includes("to_en")) ? (others[i]["英単語"] || others[i]["英文"]) : others[i]["日本語"]); }
        choices.sort(() => Math.random() - 0.5);
        choices.forEach(c => { const btn = document.createElement('button'); btn.innerText = c; btn.className = "choice-btn"; btn.onclick = () => checkAnswer(c, correctA, q); answerArea.appendChild(btn); });
      } 
      else if (answerType === "typing") { 
        shiftMode = 0; isShiftHoldMode = false;
        answerArea.innerHTML = buildTypingInputAreaMarkup(false);
        if (inputMethodMode === 'pen') setPenMode('pen');
        mountTypingMethodBody(false);
        placeQuizFeedbackAboveKeyboard(answerArea);
      }
      else if (answerType === "voice") {
        answerArea.innerHTML = `<input type="text" id="voice-recognized-text" class="large-input" readonly placeholder="ここに聞き取った言葉が出ます"><button id="voice-btn" class="submit-btn btn-blue" style="padding: 20px 40px; font-size: 24px; border-radius: 50px; display: block; margin: 0 auto;" onclick="startVoiceInput()">🎙️ マイクでこたえる</button><div id="voice-feedback" style="margin-top:15px; font-size:18px; color:#ccc;">おうちの静かな場所でやってみてね。</div>`;
      }
      else if (answerType.startsWith("fill_")) {
        setupFillBlankQuiz(correctA, answerType);
      }
      updateSessionScoreDisplay();
    }

    function setupFillBlankQuiz(word, ansType) {
      const numBlanks = parseInt(document.getElementById('setting-blank-count').value) || 1;
      const wordBlankArea = document.getElementById('quiz-word-blank-area');
      const answerArea = document.getElementById('quiz-answer-area');
      
      wordBlankArea.style.display = "block";
      fillBlanksData = [];
      activeFillBlankIndex = 0;

      let availableIndices = [];
      for(let i=1; i<word.length; i++) {
        if(/[a-zA-Z]/.test(word[i])) availableIndices.push(i);
      }
      availableIndices.sort(() => Math.random() - 0.5);
      let hiddenIndices = availableIndices.slice(0, numBlanks).sort((a,b)=>a-b);

      let displayHtml = "";
      let blankCount = 0;

      for (let i=0; i<word.length; i++) {
        if (hiddenIndices.includes(i)) {
          displayHtml += `<span class="blank-box ${blankCount === 0 ? 'active' : ''}" id="fill-blank-${blankCount}" onclick="selectFillBlank(${blankCount})">_</span>`;
          fillBlanksData.push({ originalIndex: i, correctChar: word[i], userInput: '', uiIndex: blankCount });
          blankCount++;
        } else {
          displayHtml += `<span>${word[i]}</span>`;
        }
      }
      wordBlankArea.innerHTML = displayHtml;

      if (ansType === "fill_typing") {
        shiftMode = 0; isShiftHoldMode = false;
        answerArea.innerHTML = buildTypingInputAreaMarkup(true);
        if (inputMethodMode === 'pen') setPenMode('pen');
        mountTypingMethodBody(true);
        placeQuizFeedbackAboveKeyboard(answerArea);
      } else if (ansType === "fill_4choice") {
        answerArea.innerHTML = `
          <div id="fill-4choice-container" class="fill-choices"></div>
          <button id="fill-submit-btn" class="submit-btn btn-green" style="display:none; margin: 20px auto 0;" onclick="submitFillAnswer()">✨ これで回答する</button>
        `;
        renderFill4ChoiceButtons();
      }
    }

    function selectFillBlank(index) {
      document.querySelectorAll('.blank-box').forEach(b => b.classList.remove('active'));
      activeFillBlankIndex = index;
      document.getElementById(`fill-blank-${index}`).classList.add('active');
      
      const ansType = document.getElementById('setting-answer-type').value;
      if (ansType === "fill_4choice") {
        renderFill4ChoiceButtons();
      }
    }

    function renderFill4ChoiceButtons() {
      const container = document.getElementById('fill-4choice-container');
      container.innerHTML = "";
      
      let correctChar = fillBlanksData[activeFillBlankIndex].correctChar;
      let letters = "abcdefghijklmnopqrstuvwxyz";
      if (correctChar === correctChar.toUpperCase()) letters = letters.toUpperCase();
      
      let choices = [correctChar];
      while(choices.length < 4) {
        let r = letters[Math.floor(Math.random() * letters.length)];
        if(!choices.includes(r)) choices.push(r);
      }
      choices.sort(() => Math.random() - 0.5);

      choices.forEach(c => {
        const btn = document.createElement('button');
        btn.className = "fill-choice-btn";
        btn.innerText = c;
        btn.onclick = () => { handleFillInput(c); };
        container.appendChild(btn);
      });
    }

    function handleFillInput(char) {
      fillBlanksData[activeFillBlankIndex].userInput = char;
      document.getElementById(`fill-blank-${activeFillBlankIndex}`).innerText = char;
      
      let nextIndex = fillBlanksData.findIndex(b => b.userInput === '');
      if (nextIndex !== -1) {
        selectFillBlank(nextIndex);
      } else {
        checkFillCompletion();
      }
    }

    function handleFillBackspace() {
      fillBlanksData[activeFillBlankIndex].userInput = '';
      document.getElementById(`fill-blank-${activeFillBlankIndex}`).innerText = "_";
      checkFillCompletion();
    }

    function checkFillCompletion() {
      const allFilled = fillBlanksData.every(b => b.userInput !== '');
      const submitBtn = document.getElementById('fill-submit-btn');
      if (submitBtn) {
        if (allFilled) {
          submitBtn.disabled = false;
          submitBtn.style.display = "block";
          submitBtn.classList.remove('btn-gray');
        } else {
          submitBtn.disabled = true;
          if (document.getElementById('setting-answer-type').value === "fill_4choice") {
            submitBtn.style.display = "none";
          }
        }
      }
    }

    function submitFillAnswer() {
      const q = filteredQuestions[currentQuestionIndex];
      let correctA = (q["英単語"] || q["英文"]).trim();
      
      let constructedAnswer = "";
      let fillIdx = 0;
      for(let i=0; i<correctA.length; i++) {
        let blankData = fillBlanksData.find(b => b.originalIndex === i);
        if (blankData) constructedAnswer += blankData.userInput;
        else constructedAnswer += correctA[i];
      }
      
      checkAnswer(constructedAnswer, correctA, q);
    }

    function useHint(type, penalty) {
      const q = filteredQuestions[currentQuestionIndex];
      const hintTextDisplay = document.getElementById('hint-display-text');
      const format = document.getElementById('setting-format').value;
      const isWord = currentModeName.includes("単語");
      maxDeduction = Math.max(maxDeduction, penalty);
      if (type !== 3) document.getElementById(`hint-btn-${type}`).disabled = true;
      if (type === 1 || type === 2) {
        let hintStr = "";
        if (type === 1) hintStr = q["イニシャル"] || "ヒントなし";
        else if (type === 2) hintStr = isWord ? (q["イニシャルと文字数"] || "ヒントなし") : (q["ヒント"] || q["イニシャルと文字数"] || "ヒントなし");
        hintTextDisplay.innerText = hintStr;
        hintTextDisplay.style.display = "block";
      } else if (type === 3) {
        let speakSrc = "";
        if (format === "ja_to_en_sort") speakSrc = q["英文"] || q["並び替え箇所"] || "";
        else if (format.includes("qtext") || format.includes("qaudio")) speakSrc = q["英文"];
        else speakSrc = q["英単語"] || q["英文"];
        speakText(speakSrc, 0.8);
      }
    }
    function skipQuestion() {
      const q = filteredQuestions[currentQuestionIndex];
      const format = document.getElementById('setting-format').value;
      checkAnswer("[スキップしました]", getCorrectAnswerForQuestion(q, format), q);
    }

    function checkAnswer(userA, correctA, q) {
      document.querySelectorAll('#section-quiz button').forEach(b => b.disabled = true);
      let advanceScheduled = false;
      try {
        const format = document.getElementById('setting-format').value;
        const ansType = document.getElementById('setting-answer-type').value;
        const isWord = currentModeName.includes("単語");

        let resolvedCorrect = correctA;
        if (resolvedCorrect == null || (typeof resolvedCorrect === "string" && resolvedCorrect.trim() === "")) {
          resolvedCorrect = getCorrectAnswerForQuestion(q, format);
        }

        let isCorrect = false;
        if (format === "ja_to_en_sort") {
          resolvedCorrect = getSortPrimaryCorrectDisplay(q);
          isCorrect = isSortAnswerCorrect(userA, q);
        } else {
          const userVariants = expandTextVariants(userA).map(t => normalizeText(t));
          const correctVariants = expandTextVariants(resolvedCorrect).map(t => normalizeText(t));
          const correctSet = new Set(correctVariants);
          isCorrect = userVariants.some(u => correctSet.has(u));
        }
        const basePoint = computeQuizBasePoint(format, ansType, isWord);

        quizResults.push({ questionId: q["通し番号"], isCorrect: isCorrect, timeSec: Math.round((Date.now() - questionStartTime) / 1000), basePoint: basePoint, maxDeduction: maxDeduction });
        saveQuizRecoveryDraft(currentQuestionIndex + 1);
        const feedback = document.getElementById('quiz-feedback');

        document.getElementById('quiz-word-blank-area').style.display = "none";

        const isSkip = userA === "[スキップしました]";
        const fillFeedback = ansType.startsWith("fill_");
        if (!isSkip) playAnswerSound(isCorrect);
        let plusLine = "";
        if (isCorrect && !isSkip) {
          const mult = computePointsMultiplierClient();
          const sheetPct = parseUnitSheetPointPercentClient(currentUnitName);
          const rawBefore = rawPointsFromQuizResults(quizResults.slice(0, -1));
          const rawAfter = rawPointsFromQuizResults(quizResults);
          const eBefore = applySessionEarnedFromRaw(rawBefore, mult, sheetPct);
          const eAfter = applySessionEarnedFromRaw(rawAfter, mult, sheetPct);
          const d = Math.round((eAfter - eBefore) * 100) / 100;
          if (d > 0) plusLine = `<p class="quiz-feedback-points-plus">＋${formatPointDisplayNum(d)}点</p>`;
        }
        feedback.innerHTML = buildFeedbackContentHtml(userA, resolvedCorrect, isCorrect, isSkip, maxDeduction, fillFeedback && !isSkip, ansType, plusLine);
        updateSessionScoreDisplay();
        if (ansType === "typing" || ansType === "fill_typing" || format === "ja_to_en_sort") {
          requestAnimationFrame(() => {
            try { feedback.scrollIntoView({ block: "nearest", behavior: "smooth" }); } catch (_) {}
          });
        }

        if (format === "ja_to_en_sort") destroySortQuizIfAny();
        const speakLine = (format === "ja_to_en_sort" && q["英文"]) ? q["英文"] : String(resolvedCorrect ?? "");
        speakText(String(speakLine));
        advanceScheduled = true;
      } catch (e) {
        console.error(e);
        try { if (document.getElementById('setting-format') && document.getElementById('setting-format').value === 'ja_to_en_sort') destroySortQuizIfAny(); } catch (_) {}
        const feedback = document.getElementById('quiz-feedback');
        try { document.getElementById('quiz-word-blank-area').style.display = "none"; } catch (_) {}
        feedback.innerHTML = `<span style='color:#F44336;font-size:30px;font-weight:bold;'>⚠️ エラー</span><br><span style="font-size:18px;">「つぎへ」でつづけてね。</span>`;
        try {
          const format = document.getElementById('setting-format').value;
          const ansType = document.getElementById('setting-answer-type').value;
          const isWord = currentModeName.includes("単語");
          const basePoint = computeQuizBasePoint(format, ansType, isWord);
          quizResults.push({ questionId: q["通し番号"], isCorrect: false, timeSec: Math.round((Date.now() - questionStartTime) / 1000), basePoint: basePoint, maxDeduction: maxDeduction });
          saveQuizRecoveryDraft(currentQuestionIndex + 1);
        } catch (_) {
          quizResults.push({ questionId: q["通し番号"] || 0, isCorrect: false, timeSec: 0, basePoint: 0, maxDeduction: 0 });
          saveQuizRecoveryDraft(currentQuestionIndex + 1);
        }
        updateSessionScoreDisplay();
        advanceScheduled = true;
      } finally {
        setPenTypingSubmitBusy(false);
        const nextBtn = document.getElementById('quiz-next-btn');
        nextBtn.disabled = false;
        nextBtn.style.display = "block";
        document.querySelectorAll('#section-quiz .cancel-btn').forEach(btn => btn.disabled = false);
        if (advanceScheduled) {
          clearTimeout(autoNextTimer);
          autoNextTimer = setTimeout(() => { nextQuestion(); }, 3000);
        }
      }
    }

    function nextQuestion() { clearTimeout(autoNextTimer); window.speechSynthesis.cancel(); resetHandwritingInputState(); currentQuestionIndex++; showQuestion(); }
    function quitQuiz() { clearTimeout(autoNextTimer); window.speechSynthesis.cancel(); resetHandwritingInputState(); document.getElementById('section-quiz')?.classList.remove('pen-mode-active'); document.getElementById('quiz-answer-area')?.classList.remove('pen-mode-active'); openSettingsScreen(); }
    
    function finishQuiz() {
      document.getElementById('section-quiz')?.classList.remove('pen-mode-active');
      document.getElementById('quiz-answer-area')?.classList.remove('pen-mode-active');
      switchSection('section-result'); 
      document.getElementById('result-content').innerHTML = "<p>おくっています...</p>"; 
      document.getElementById('result-retry-btn').style.display = "none"; document.getElementById('result-settings-btn').style.display = "none"; document.getElementById('result-home-btn').style.display = "none";

      const user = JSON.parse(localStorage.getItem('app_kid_user'));
      const isRandom = document.getElementById('setting-order').value === 'random';
      const detailedUnitId = getDetailedUnitId(); 

      let payload = { action: "save_learning_session", userId: user.id, unitId: detailedUnitId, unitSheetName: currentUnitName, isReviewMode: currentIsReviewMode, isRandom: isRandom, results: quizResults };
      if (isTrainingMode) {
        payload.trainingStepIndex = currentTrainingStepIndex;
        payload.trainingMenuId = currentTrainingMenuId;
      }

      fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify(payload) })
      .then(r => r.json()).then(d => {
        if(d.status === "success") {
          let bonusMsg = "";
          if (currentIsReviewMode) bonusMsg += `<p style="color:#FF9800;font-weight:bold;margin:5px 0;">🔥 ニガテ特訓ボーナス適用</p>`;
          if (d.bonusApplied) bonusMsg += `<p style="color:#e50914;font-weight:bold;margin:5px 0;">🎲 ランダム出題ボーナス（10%UP）</p>`;
          if (isTrainingMode) bonusMsg += `<p style="color:#9C27B0;font-weight:bold;margin:5px 0;">🎯 特訓ルートクリア！</p>`;
          if (d.sheetPointPercent != null && Number(d.sheetPointPercent) < 100) {
            bonusMsg += `<p style="color:#90CAF9;font-size:15px;margin:6px 0;">📎 ${escapeHtml(formatUnitSheetDisplayLabel(currentUnitName))} のため、かくとくポイントは <strong>${d.sheetPointPercent}%</strong> になっています。</p>`;
          }

          document.getElementById('result-content').innerHTML = `<h2 style="color:#4CAF50;">✨ おつかれさま！ ✨</h2>` + bonusMsg + `<p>かくとくポイント: <span style="color:gold;font-size:30px;">+${d.earnedPoints}</span></p><hr style="border-color:#444;"><p>合計ポイント: ${d.newTotal}</p>`;
          user.points = d.newTotal; user.historyJson = d.historyJson; user.dailyPointsJson = d.dailyPointsJson; 
          if(d.trainingProgressJson) {
            user.trainingProgressJson = d.trainingProgressJson;
          }
          localStorage.setItem('app_kid_user', JSON.stringify(user));
          clearQuizRecoveryDraft();
          
          if (!isTrainingMode) {
              document.getElementById('result-retry-btn').style.display = "block"; 
              document.getElementById('result-settings-btn').style.display = "block"; 
          }
          document.getElementById('result-home-btn').style.display = "block";
        }
      }).catch(() => {
        document.getElementById('result-content').innerHTML = `<h2 style="color:#F44336;">通信エラー</h2><p>結果の送信に失敗しました。<br>「ホームにもどる」で復帰データを使って再開できます。</p>`;
        document.getElementById('result-home-btn').style.display = "block";
      });
    }

    document.addEventListener('keydown', (e) => {
      const answerInput = document.getElementById('type-answer');
      const quizSection = document.getElementById('section-quiz');
      if (!quizSection.classList.contains('active')) return;

      const ansType = document.getElementById('setting-answer-type').value;

      if (ansType === "typing" && answerInput) {
        if (inputMethodMode === 'pen') return;
        if (['Enter', 'Backspace', ' ', 'Shift'].includes(e.key) || (e.key.length === 1 && /^[\x20-\x7E¥]$/.test(e.key))) { e.preventDefault(); }
        if (e.key === 'Enter') { const enterBtn = Array.from(document.querySelectorAll('.key-action')).find(btn => btn.innerText === "けってい"); if (enterBtn && !enterBtn.disabled) enterBtn.click(); } 
        else if (e.key === 'Backspace') { answerInput.value = answerInput.value.slice(0, -1); } 
        else if (e.key === ' ') { answerInput.value += ' '; } 
        else if (e.key.length === 1 && /^[\x20-\x7E¥]$/.test(e.key)) { answerInput.value += e.key; }
      }
      else if (ansType === "fill_typing") {
        if (inputMethodMode === 'pen') return;
        if (['Enter', 'Backspace', ' ', 'Shift'].includes(e.key) || (e.key.length === 1 && /^[a-zA-Z]$/.test(e.key))) { e.preventDefault(); }
        if (e.key === 'Enter') { const enterBtn = document.getElementById('fill-submit-btn'); if (enterBtn && !enterBtn.disabled) enterBtn.click(); }
        else if (e.key === 'Backspace') { handleFillBackspace(); }
        else if (e.key.length === 1 && /^[a-zA-Z]$/.test(e.key)) { handleFillInput(e.key); }
      }
    });
