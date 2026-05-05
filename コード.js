function setupSystem() {
  const scriptId = ScriptApp.getScriptId();
  const gasFile = DriveApp.getFileById(scriptId);
  const parents = gasFile.getParents();
  let parentFolder = DriveApp.getRootFolder();
  if (parents.hasNext()) parentFolder = parents.next();

  const props = PropertiesService.getScriptProperties();
  let adminSsId = props.getProperty('ADMIN_SS_ID');
  let materialsFolderId = props.getProperty('MATERIALS_FOLDER_ID');
  let kanjiMaterialsFolderId = props.getProperty('KANJI_MATERIALS_FOLDER_ID');
  let logMessage = "【セットアップログ】\n";

  if (!materialsFolderId) {
    const materialsFolder = parentFolder.createFolder("materials");
    props.setProperty('MATERIALS_FOLDER_ID', materialsFolder.getId());
    logMessage += "✅ materialsフォルダを作成しました。\n";

    const wordModeSs = SpreadsheetApp.create("単語練習モード");
    wordModeSs.getSheets()[0].setName("単元A").appendRow(["通し番号", "英単語", "日本語", "イニシャル", "イニシャルと文字数", "ヒント"]);
    wordModeSs.getSheets()[0].appendRow([1, "twin", "ふたご", "t", "t _ _ _", "tw", "カタカナではツインだよ。"]);
    DriveApp.getFileById(wordModeSs.getId()).moveTo(materialsFolder);

    const phraseModeSs = SpreadsheetApp.create("表現練習モード");
    const phraseHeaders = [
      "通し番号", "日本語", "英文", "別解１", "別解２", "別解３", "別解４", "別解５", "疑問文", "穴埋め１", "穴埋め２", "イニシャル", "イニシャルと文字数", "ヒント",
      "並び替え用英文", "並び替え箇所", "並び替え語句ダミー",
      "並び替え語句1", "並び替え語句2", "並び替え語句3", "並び替え語句4", "並び替え語句5", "並び替え語句6", "並び替え語句7", "並び替え語句8"
    ];
    phraseModeSs.getSheets()[0].setName("単元C").appendRow(phraseHeaders);
    phraseModeSs.getSheets()[0].appendRow([
      1, "私は8歳です。", "I'm eight years old.", "I am eight years old.", "", "", "", "", "How old are you?", "I'm (     ) years old.", "", "I", "I'm _ _ _ _ _", "年齢を聞かれた時の答え方だよ。",
      "", "", "",
      "", "", "", "", "", "", "", ""
    ]);
    phraseModeSs.getSheets()[0].appendRow([
      2, "これは私が昨日買った本です。", "This is the book I bought yesterday.", "", "", "", "", "", "", "", "", "", "", "本を指さしながら言う表現の練習だよ。",
      "This is the book I bought yesterday.", "This is the book I bought yesterday.", "a",
      "This", "is", "the", "book", "I", "bought", "yesterday.", ""
    ]);
    DriveApp.getFileById(phraseModeSs.getId()).moveTo(materialsFolder);
    
    logMessage += "✅ サンプル教材を作成しました。\n";
  }
  let kanjiFolder = null;
  if (!kanjiMaterialsFolderId) {
    kanjiFolder = parentFolder.createFolder("教材");
    props.setProperty('KANJI_MATERIALS_FOLDER_ID', kanjiFolder.getId());
    logMessage += "✅ 教材フォルダ（漢字用）を作成しました。\n";
  } else {
    try { kanjiFolder = DriveApp.getFolderById(kanjiMaterialsFolderId); } catch (_) {}
  }
  if (kanjiFolder) {
    const info = ensureKanjiSampleBook_(kanjiFolder);
    if (info.created) logMessage += "✅ 漢字学習サンプルブックを作成しました。\n";
    if (info.sheetId && !props.getProperty('KANJI_SHEET_ID')) {
      props.setProperty('KANJI_SHEET_ID', info.sheetId);
      logMessage += "✅ KANJI_SHEET_ID をサンプルブックに設定しました。\n";
    }
  }

  let adminSs;
  if (!adminSsId) {
    adminSs = SpreadsheetApp.create("学習アプリ_管理ブック");
    props.setProperty('ADMIN_SS_ID', adminSs.getId());
    DriveApp.getFileById(adminSs.getId()).moveTo(parentFolder); 

    const usersSheet = adminSs.getSheets()[0];
    usersSheet.setName("users");
    // 新しく「特訓進捗_JSON」列を追加しました！
    usersSheet.appendRow(["ID", "名前", "PIN", "合計ポイント", "最終学習日時_JSON", "履歴_JSON", "日別ポイント_JSON", "特訓進捗_JSON"]);
    usersSheet.appendRow(["user_1", "テスト太郎", "1234", 100, "{}", "{}", "{}", "{}"]);

    const rewardsSheet = adminSs.insertSheet("rewards");
    rewardsSheet.appendRow(["ID", "名前", "必要ポイント", "説明"]);
    rewardsSheet.appendRow(["r_1", "YouTube視聴1時間延長券", 50, "管理者に提示して使ってね。"]);

    const inventorySheet = adminSs.insertSheet("inventory");
    inventorySheet.appendRow(["交換日時", "ユーザーID", "景品ID", "景品名", "状態"]);

    logMessage += "✅ 管理ブックと基本シートを作成しました。\n";
  } else {
    adminSs = SpreadsheetApp.openById(adminSsId);
  }

  let appSettingsSheet = adminSs.getSheetByName("アプリ設定");
  if (!appSettingsSheet) {
    appSettingsSheet = adminSs.insertSheet("アプリ設定");
    appSettingsSheet.appendRow(["設定名", "値"]);
    appSettingsSheet.appendRow(["基本ポイント_4択", 2]);
    appSettingsSheet.appendRow(["基本ポイント_タイピング", 20]);
    appSettingsSheet.appendRow(["基本ポイント_穴埋め", 5]);
    appSettingsSheet.appendRow(["基本ポイント_音声", 20]);
    appSettingsSheet.appendRow(["ヒント減点_イニシャル", 5]);
    appSettingsSheet.appendRow(["ヒント減点_文字数", 7]);
    appSettingsSheet.appendRow(["ヒント減点_音声", 10]);
    logMessage += "✅ 「アプリ設定」を追加しました。\n";
  }

  // 形式×答え方ごとの基本ポイントキー（足りないものだけ追加）
  if (appSettingsSheet) {
    const settingsData = appSettingsSheet.getDataRange().getValues();
    const existingKeys = new Set(settingsData.slice(1).map(r => String(r[0] || "")));
    const defaultPointRows = [
      ["基本Pt_ja_to_en_4choice", 3],
      ["基本Pt_ja_to_en_typing", 20],
      ["基本Pt_ja_to_en_voice", 20],
      ["基本Pt_ja_to_en_fill_4choice", 5],
      ["基本Pt_ja_to_en_fill_typing", 5],
      ["基本Pt_en_to_ja_4choice", 2],
      ["基本Pt_en_to_ja_typing", 20],
      ["基本Pt_en_to_ja_voice", 20],
      ["基本Pt_en_audio_to_ja_4choice", 2],
      ["基本Pt_en_audio_to_ja_typing", 20],
      ["基本Pt_en_audio_to_ja_voice", 20],
      ["基本Pt_en_audio_to_en_4choice", 3],
      ["基本Pt_en_audio_to_en_typing", 20],
      ["基本Pt_en_audio_to_en_voice", 20],
      ["基本Pt_en_audio_to_en_fill_4choice", 5],
      ["基本Pt_en_audio_to_en_fill_typing", 5],
      ["基本Pt_en_to_en_typing", 20],
      ["基本Pt_en_to_en_voice", 20],
      ["基本Pt_qtext_to_en_4choice", 3],
      ["基本Pt_qtext_to_en_typing", 30],
      ["基本Pt_qtext_to_en_voice", 30],
      ["基本Pt_qaudio_to_en_typing", 30],
      ["基本Pt_qaudio_to_en_voice", 30],
      ["基本Pt_ja_to_en_sort_sort_all", 25],
      ["基本Pt_ja_to_en_sort_sort_dummy", 28],
      ["基本Pt_ja_to_en_sort_sort_missing", 30],
      ["漢字基本Pt_採点_90以上", 10],
      ["漢字基本Pt_採点_80以上", 5],
      ["漢字基本Pt_採点_70以上", 4],
      ["漢字基本Pt_採点_60以上", 3],
      ["漢字基本Pt_採点_50以上", 1],
      ["漢字採点_高得点回数上限_週", 3],
      ["漢字採点_回数上限後倍率", 0.1],
      ["漢字採点_回復率_日", 0.15],
      ["漢字採点_完全回復日数", 7]
    ];
    defaultPointRows.forEach(row => {
      if (!existingKeys.has(row[0])) {
        appSettingsSheet.appendRow(row);
      }
    });
  }

  let extLearningSheet = adminSs.getSheetByName("外部学習");
  if (!extLearningSheet) {
    extLearningSheet = adminSs.insertSheet("外部学習");
    extLearningSheet.appendRow(["カテゴリ", "分量", "獲得ポイント"]);
    extLearningSheet.appendRow(["ピアノ練習", "30分", 50]);
    extLearningSheet.appendRow(["読書", "30分", 30]);
    logMessage += "✅ 「外部学習」を追加しました。\n";
  }

  let extReqSheet = adminSs.getSheetByName("外部学習申請");
  if (!extReqSheet) {
    extReqSheet = adminSs.insertSheet("外部学習申請");
    extReqSheet.appendRow(["申請日時", "ユーザーID", "ユーザー名", "カテゴリ", "分量", "ポイント", "こどもメモ", "状態", "処理日時", "おとなメモ"]);
    logMessage += "✅ 「外部学習申請」を追加しました。\n";
  }

  const appSettingsForExtPin = adminSs.getSheetByName("アプリ設定");
  if (appSettingsForExtPin) {
    const extPinRows = appSettingsForExtPin.getDataRange().getValues();
    let hasExtAdminPin = false;
    for (let i = 1; i < extPinRows.length; i++) {
      if (String(extPinRows[i][0]) === "外部学習_管理者PIN") {
        hasExtAdminPin = true;
        break;
      }
    }
    if (!hasExtAdminPin) {
      appSettingsForExtPin.appendRow(["外部学習_管理者PIN", "1234"]);
      logMessage += "✅ 「外部学習_管理者PIN」を追加しました（アプリ設定シートで変更してください）。\n";
    }
  }

  // ★ 特訓メニュー1（従来名「特訓メニュー」互換）＋ 特訓メニュー2～12（ヘッダーのみ）
  const trainingHeader = ["対象ユーザー", "単元", "問題の形式", "こたえ方", "出し方", "隠す文字数"];
  const trainingPatterns = [
    ["全員", "", "英単語→日本語", "4択", "ランダム", ""],
    ["全員", "", "英単語→日本語", "タイピング", "ランダム", ""],
    ["全員", "", "英単語→日本語", "音声", "ランダム", ""],
    ["全員", "", "日本語→英単語", "4択", "ランダム", ""],
    ["全員", "", "日本語→英単語", "タイピング", "ランダム", ""],
    ["全員", "", "日本語→英単語", "音声", "ランダム", ""],
    ["全員", "", "日本語→英単語", "穴埋め4択", "ランダム", 1],
    ["全員", "", "日本語→英単語", "穴埋めタイピング", "ランダム", 1],
    ["全員", "", "音声→日本語", "4択", "ランダム", ""],
    ["全員", "", "音声→日本語", "タイピング", "ランダム", ""],
    ["全員", "", "音声→日本語", "音声", "ランダム", ""],
    ["全員", "", "音声→英単語", "タイピング", "ランダム", ""],
    ["全員", "", "音声→英単語", "音声", "ランダム", ""],
    ["全員", "", "英語→英語", "タイピング", "ランダム", ""],
    ["全員", "", "英語→英語", "音声", "ランダム", ""],
    ["全員", "", "漢字→採点チャレンジ", "採点", "ランダム", ""]
  ];
  let trainingSheet = adminSs.getSheetByName("特訓メニュー");
  if (!trainingSheet) {
    trainingSheet = adminSs.insertSheet("特訓メニュー");
    trainingSheet.appendRow(trainingHeader);
    trainingPatterns.forEach(p => trainingSheet.appendRow(p));
    logMessage += "✅ 新しい「特訓メニュー（学習ルート）」を追加しました。\n";
  }
  if (!adminSs.getSheetByName("特訓メニュー1")) {
    const s1 = adminSs.insertSheet("特訓メニュー1");
    const src = adminSs.getSheetByName("特訓メニュー");
    if (src) {
      const data = src.getDataRange().getValues();
      if (data.length) s1.getRange(1, 1, data.length, data[0].length).setValues(data);
      else s1.appendRow(trainingHeader);
    } else {
      s1.appendRow(trainingHeader);
      trainingPatterns.forEach(p => s1.appendRow(p));
    }
    logMessage += "✅ 「特訓メニュー1」を追加しました（従来の「特訓メニュー」と同じ内容）。\n";
  }
  for (let m = 2; m <= 12; m++) {
    const nm = "特訓メニュー" + m;
    if (!adminSs.getSheetByName(nm)) {
      const sh = adminSs.insertSheet(nm);
      sh.appendRow(trainingHeader);
      logMessage += "✅ 「" + nm + "」を追加しました。\n";
    }
  }
  const appSettingsTrain = adminSs.getSheetByName("アプリ設定");
  if (appSettingsTrain) {
    const asData = appSettingsTrain.getDataRange().getValues();
    const existing = {};
    for (let ri = 0; ri < asData.length; ri++) existing[String(asData[ri][0] || "")] = true;
    for (let m = 1; m <= 12; m++) {
      const k = "特訓メニュー" + m + "_表示名";
      if (!existing[k]) appSettingsTrain.appendRow([k, ""]);
    }
  }

  console.log(logMessage);
  return logMessage;
}

function ensureKanjiSampleBook_(kanjiFolder) {
  const out = { created: false, sheetId: "" };
  const sampleName = "漢字学習サンプル";
  let sampleFile = null;
  const files = kanjiFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getName()) === sampleName) {
      sampleFile = f;
      break;
    }
  }
  let ss;
  if (!sampleFile) {
    ss = SpreadsheetApp.create(sampleName);
    sampleFile = DriveApp.getFileById(ss.getId());
    sampleFile.moveTo(kanjiFolder);
    out.created = true;
  } else {
    ss = SpreadsheetApp.openById(sampleFile.getId());
  }
  out.sheetId = ss.getId();

  let sheet = ss.getSheetByName("小１");
  if (!sheet) {
    sheet = ss.getSheets()[0] || ss.insertSheet("小１");
    sheet.setName("小１");
  }
  const header = [
    "セット", "漢字",
    "訓読みA_読み", "訓A_例文1", "訓A_例文2",
    "訓読みB_読み", "訓B_例文1", "訓B_例文2",
    "訓読みC_読み", "訓C_例文1", "訓C_例文2",
    "訓読みD_読み", "訓D_例文1", "訓D_例文2",
    "音読みA_読み", "音A_例文1", "音A_例文2",
    "音読みB_読み", "音B_例文1", "音B_例文2",
    "音読みC_読み", "音C_例文1", "音C_例文2",
    "音読みD_読み", "音D_例文1", "音D_例文2"
  ];
  const data = sheet.getDataRange().getValues();
  const headNow = (data[0] || []).map(v => String(v || "").trim());
  const sameHeader = headNow.length >= header.length && header.every((h, i) => headNow[i] === h);
  if (!sameHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
  const hasRows = sheet.getLastRow() > 1;
  if (!hasRows) {
    const rows = [
      [1, "一", "ひと", "一つある。", "もう一つだ。", "いっ", "けしゴムを一こかした。", "×", "×", "×", "×", "×", "×", "×", "いち", "一ばんだよ。", "一月はふゆだ。", "いつ", "きん一にまぜる。", "とう一する。", "×", "×", "×", "×", "×", "×"],
      [1, "右", "みぎ", "右をむく。", "右手を見る。", "×", "×", "×", "×", "×", "×", "×", "×", "×", "う", "うせつする。", "さゆうを見る。", "×", "×", "×", "×", "×", "×", "×", "×", "×"],
      [1, "雨", "あめ", "雨がふる。", "大雨がふる。", "あま", "雨水がでる。", "雨ぐもをみる。", "×", "×", "×", "×", "×", "×", "う", "雨てんちゅうしだ。", "ごう雨になる。", "×", "×", "×", "×", "×", "×", "×", "×", "×"],
      [1, "円", "まる", "円い形。", "円くかく。", "×", "×", "×", "×", "×", "×", "×", "×", "×", "えん", "百円玉をもつ。", "いちまん円だ。", "×", "×", "×", "×", "×", "×", "×", "×", "×"],
      [1, "王", "×", "×", "×", "×", "×", "×", "×", "×", "×", "×", "×", "×", "おう", "王さまにあう。", "王女さまをみる。", "×", "×", "×", "×", "×", "×", "×", "×", "×"],
      [1, "音", "おと", "音がなる。", "足音をきく。", "ね", "虫の音をきく。", "本音をいう。", "×", "×", "×", "×", "×", "×", "おん", "はつ音がよい。", "おん楽をきく。", "×", "×", "×", "×", "×", "×", "×", "×", "×"],
      [1, "下", "した", "下をむく。", "くつ下をはく。", "さ", "あたまを下げる。", "手を下げる。", "くだ", "さかを下る。", "川を下る。", "お", "山を下りる。", "木から下りる。", "か", "上下する。", "下きゅう生だ。", "げ", "下こうする。", "下しゃする。", "×", "×", "×", "×", "×", "×"]
    ];
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  return out;
}

const sendResponse = (responseObject) => {
  return ContentService.createTextOutput(JSON.stringify(responseObject)).setMimeType(ContentService.MimeType.JSON);
};

function doGet() {
  return ContentService
    .createTextOutput("OK: GAS endpoint is running")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;

    if (action === "save_learning_session") return handleSaveLearningSession(requestData);
    else if (action === "get_child_users") return handleGetChildUsers(requestData);
    else if (action === "verify_kid_pin") return handleVerifyKidPin(requestData);
    else if (action === "get_materials_list") return handleGetMaterialsList(requestData);
    else if (action === "get_questions") return handleGetQuestions(requestData);
    else if (action === "get_rewards") return handleGetRewards(requestData);
    else if (action === "exchange_reward") return handleExchangeReward(requestData);
    else if (action === "change_pin") return handleChangePin(requestData);
    else if (action === "get_inventory") return handleGetInventory(requestData);
    else if (action === "consume_reward") return handleConsumeReward(requestData);
    else if (action === "get_app_settings") return handleGetAppSettings(requestData);
    else if (action === "get_points_multiplier") return handleGetPointsMultiplier(requestData);
    else if (action === "get_external_learning") return handleGetExternalLearning(requestData);
    else if (action === "submit_external_learning_request") return handleSubmitExternalLearningRequest(requestData);
    else if (action === "get_pending_external_requests") return handleGetPendingExternalRequests(requestData);
    else if (action === "approve_external_request") return handleApproveExternalRequest(requestData);
    else if (action === "reject_external_request") return handleRejectExternalRequest(requestData);
    else if (action === "get_my_external_learning_requests") return handleGetMyExternalLearningRequests(requestData);
    else if (action === "recognize_handwriting") return recognizeSentence(requestData.ink || []);
    else if (action === "get_kanji_init_data") return handleGetKanjiInitData(requestData);
    else if (action === "get_kanji_data_from_sheet") return handleGetKanjiDataFromSheet(requestData);
    else if (action === "get_kanji_quiz_sets") return handleGetKanjiQuizSets(requestData);
    else if (action === "get_kanji_quiz_questions") return handleGetKanjiQuizQuestions(requestData);
    else if (action === "append_kanji_weak_signals") return handleAppendKanjiWeakSignals(requestData);
    else if (action === "get_kanji_weak_review_plan") return handleGetKanjiWeakReviewPlan(requestData);
    
    // ★ 特訓ルート用のAPI
    else if (action === "get_training_route") return handleGetTrainingRoute(requestData);
    
    else return sendResponse({ status: "error", message: "無効なactionです" });
  } catch (error) {
    return sendResponse({ status: "error", message: error.toString() });
  }
}

function doOptions(e) { return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT); }

function recognizeSentence(allStrokes) {
  if (!Array.isArray(allStrokes) || allStrokes.length === 0) {
    return sendResponse({ status: "error", message: "ストロークが空です。" });
  }

  const endpoints = [
    "https://inputtools.google.com/request?ime=handwriting&app=mobilesearch&cs=1&oe=UTF-8",
    "https://www.google.com.hk/inputtools/request?ime=handwriting&app=mobilesearch&cs=1&oe=UTF-8"
  ];
  const languages = ["en", "en-US"];
  let lastError = "";

  for (let li = 0; li < languages.length; li++) {
    const lang = languages[li];
    const payload = {
      options: "enable_pre_space",
      requests: [{
        writing_guide: { writing_area_width: 1000, writing_area_height: 400 },
        ink: allStrokes,
        language: lang
      }]
    };

    for (let ei = 0; ei < endpoints.length; ei++) {
      const url = endpoints[ei];
      try {
        const response = UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        const code = response.getResponseCode();
        const body = response.getContentText();
        if (code !== 200) {
          lastError = "HTTP " + code + " @ " + url;
          continue;
        }
        const result = JSON.parse(body);
        if (result[0] === "SUCCESS" && result[1] && result[1][0] && result[1][0][1] && result[1][0][1][0]) {
          const candidates = result[1][0][1]
            .filter(v => typeof v === "string" && v.trim() !== "")
            .slice(0, 10);
          return sendResponse({ status: "success", text: result[1][0][1][0], candidates: candidates });
        }
        lastError = "認識候補なし (" + lang + " @ " + url + ")";
      } catch (e) {
        lastError = String(e);
      }
    }
  }

  return sendResponse({ status: "error", message: "認識エラー: " + lastError });
}

// ==========================================
// ★ 特訓ルート（メニュー1～12）
// ==========================================
function getTrainingMenuSheet_(adminSs, menuId) {
  const name = "特訓メニュー" + menuId;
  let sheet = adminSs.getSheetByName(name);
  if (!sheet && menuId === 1) sheet = adminSs.getSheetByName("特訓メニュー");
  return sheet;
}

/** 旧形式（今日のキー直下に stepIndex: true）を { "1": { stepIndex: true } } に寄せる */
function migrateTrainingProgressIfNeeded_(trainingProgressJson) {
  if (!trainingProgressJson) return;
  const todayStr = new Date().toISOString().split('T')[0];
  const t = trainingProgressJson[todayStr];
  if (!t || typeof t !== "object") return;
  var hasNestedMenu = false;
  for (var k in t) {
    if (["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"].indexOf(String(k)) >= 0) {
      if (t[k] && typeof t[k] === "object" && !Array.isArray(t[k])) {
        hasNestedMenu = true;
        break;
      }
    }
  }
  if (hasNestedMenu) return;
  var hasFlatTrue = false;
  for (var k2 in t) {
    if (t[k2] === true) {
      hasFlatTrue = true;
      break;
    }
  }
  if (!hasFlatTrue) return;
  var nested = {};
  for (var k3 in t) {
    if (t[k3] === true) nested[k3] = true;
  }
  trainingProgressJson[todayStr] = { "1": nested };
}

function normalizeProgressForMenu_(todayBlock, menuId) {
  const mid = String(menuId);
  if (!todayBlock || typeof todayBlock !== "object") return {};
  if (todayBlock[mid] && typeof todayBlock[mid] === "object" && !Array.isArray(todayBlock[mid])) {
    return todayBlock[mid];
  }
  if (mid === "1") {
    var hasLegacy = false;
    for (var k in todayBlock) {
      if (todayBlock[k] === true) {
        hasLegacy = true;
        break;
      }
    }
    if (hasLegacy) {
      var out = {};
      for (var k2 in todayBlock) {
        if (todayBlock[k2] === true) out[k2] = true;
      }
      return out;
    }
  }
  return {};
}

function handleGetTrainingRoute(req) {
  const props = PropertiesService.getScriptProperties();
  const adminSs = SpreadsheetApp.openById(props.getProperty('ADMIN_SS_ID'));
  let menuId = parseInt(req.trainingMenuId, 10);
  if (isNaN(menuId) || menuId < 1 || menuId > 12) menuId = 1;

  const sheet = getTrainingMenuSheet_(adminSs, menuId);
  const route = [];
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    const header = data[0] || [];
    const blankIdx = header.indexOf("隠す文字数");
    for (let i = 1; i < data.length; i++) {
      let target = String(data[i][0]);
      if (target === "全員" || target.includes(req.userId)) {
        route.push({
          stepIndex: i,
          unitName: data[i][1],
          qFormat: data[i][2],
          aFormat: data[i][3],
          mode: data[i][4],
          blankCount: blankIdx >= 0 ? data[i][blankIdx] : ""
        });
      }
    }
  }

  const usersSheet = adminSs.getSheetByName("users");
  const userData = usersSheet.getDataRange().getValues();
  let progressData = {};
  const todayStr = new Date().toISOString().split('T')[0];

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === req.userId) {
      const rawProgress = JSON.parse(userData[i][7] || "{}");
      const todayBlock = rawProgress[todayStr] || {};
      progressData = normalizeProgressForMenu_(todayBlock, menuId);
      break;
    }
  }

  return sendResponse({ status: "success", route: route, progress: progressData, menuId: menuId });
}

/** materials ブックのシート名が「単語A_40」のように末尾 _数字 なら、その数字を得点％として最後に乗算（未指定は100） */
function parseUnitSheetPointPercent_(sheetName) {
  const s = String(sheetName || "");
  const m = s.match(/_(\d+)$/);
  if (!m) return 100;
  let p = parseInt(m[1], 10);
  if (isNaN(p)) return 100;
  if (p < 0) p = 0;
  if (p > 100) p = 100;
  return p;
}

function getAppSettingsMap_(adminSs) {
  const sheet = adminSs.getSheetByName("アプリ設定");
  const out = {};
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) out[String(data[i][0])] = data[i][1];
  }
  return out;
}

function getKanjiBasePointsByScore_(score, settings) {
  const s = Number(score) || 0;
  if (s >= 90) return Number(settings["漢字基本Pt_採点_90以上"] || 10);
  if (s >= 80) return Number(settings["漢字基本Pt_採点_80以上"] || 5);
  if (s >= 70) return Number(settings["漢字基本Pt_採点_70以上"] || 4);
  if (s >= 60) return Number(settings["漢字基本Pt_採点_60以上"] || 3);
  if (s >= 50) return Number(settings["漢字基本Pt_採点_50以上"] || 1);
  return 0;
}

function toDateOnly_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function calcKanjiCharRecoveryRate_(highScoreDates, now, settings) {
  const maxHighTimes = Math.max(1, Number(settings["漢字採点_高得点回数上限_週"] || 3));
  const overRate = Number(settings["漢字採点_回数上限後倍率"] || 0.1);
  const recoverPerDay = Number(settings["漢字採点_回復率_日"] || 0.15);
  const maxDays = Math.max(1, Number(settings["漢字採点_完全回復日数"] || 7));
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const recent = (Array.isArray(highScoreDates) ? highScoreDates : [])
    .map(v => new Date(v))
    .filter(d => !isNaN(d.getTime()) && d >= weekAgo)
    .sort((a, b) => a.getTime() - b.getTime());
  if (recent.length < maxHighTimes) return 1.0;
  const triggerDate = recent[maxHighTimes - 1];
  const days = Math.max(0, Math.floor((toDateOnly_(now) - toDateOnly_(triggerDate)) / (24 * 60 * 60 * 1000)));
  if (days >= maxDays) return 1.0;
  return Math.min(1, overRate + recoverPerDay * days);
}

// 学習結果の保存（進捗チェックの更新を追加）
function handleSaveLearningSession(req) {
  const props = PropertiesService.getScriptProperties();
  const adminSs = SpreadsheetApp.openById(props.getProperty('ADMIN_SS_ID'));
  const usersSheet = adminSs.getSheetByName("users");
  const data = usersSheet.getDataRange().getValues();
  
  let targetRowIdx = -1;
  let userData = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === req.userId) {
      targetRowIdx = i + 1;
      userData = {
        points: Number(data[i][3]) || 0,
        lastStudyJson: JSON.parse(data[i][4] || "{}"),
        historyJson: JSON.parse(data[i][5] || "{}"),
        dailyPointsJson: JSON.parse(data[i][6] || "{}"),
        trainingProgressJson: JSON.parse(data[i][7] || "{}") // ★ 進捗データ
      };
      break;
    }
  }
  if (targetRowIdx === -1) return sendResponse({ status: "error", message: "ユーザーが見つかりません" });

  const now = new Date();
  const todayStr = now.toISOString().split('T')[0];
  let multiplier = 1.0;
  // 漢字セットは教材＋セット単位で前回学習を参照。同一プレイ中の2問目以降は短間隔扱いにせず、直前の1問内の減衰を防ぐ
  const lastStudyKey = String(req.kanjiSetScopeId || req.unitId || "");
  let lastStudyTimeStr = lastStudyKey ? userData.lastStudyJson[lastStudyKey] : undefined;
  if (req.kanjiSetContinuation) lastStudyTimeStr = null;

  // 時間経過による緩和（通常ポイント計算）
  if (lastStudyTimeStr) {
    const lastTime = new Date(lastStudyTimeStr);
    const diffHours = (now - lastTime) / (1000 * 60 * 60);
    let basePercent = 10 + Math.floor(diffHours / 2) * 10;
    if (basePercent > 100) basePercent = 100;
    multiplier = basePercent / 100;
  }
  
  // ボーナスの適用
  if (req.isReviewMode) multiplier += 0.4;
  if (req.isRandom) multiplier += 0.1;

  let sessionRawPoints = 0;
  if (!userData.historyJson[req.unitId]) userData.historyJson[req.unitId] = {};
  const unitHistory = userData.historyJson[req.unitId];

  // 漢字採点チャレンジ専用（score と char を受け取り、文字単位の回数制限＋日次回復を適用）
  if (req.learningCategory === "kanji" && req.challengeType === "score" && req.kanjiChar) {
    const settings = getAppSettingsMap_(adminSs);
    const charKey = String(req.kanjiChar);
    const score = Number(req.score) || 0;
    const basePt = getKanjiBasePointsByScore_(score, settings);
    if (!userData.historyJson.__kanjiChallenge) userData.historyJson.__kanjiChallenge = {};
    if (!userData.historyJson.__kanjiChallenge[charKey]) userData.historyJson.__kanjiChallenge[charKey] = { highScoreDates: [] };
    const cHist = userData.historyJson.__kanjiChallenge[charKey];
    if (!Array.isArray(cHist.highScoreDates)) cHist.highScoreDates = [];
    const recoveryRate = calcKanjiCharRecoveryRate_(cHist.highScoreDates, now, settings);
    sessionRawPoints = Math.round(basePt * recoveryRate * 100) / 100;
    var scriptBonusMult = Number(req.kanjiScriptBonusMult);
    if (
      req.questionCorrect === true &&
      !isNaN(scriptBonusMult) &&
      scriptBonusMult > 1 &&
      scriptBonusMult <= 2
    ) {
      sessionRawPoints = Math.round(sessionRawPoints * scriptBonusMult * 100) / 100;
    }
    // 手書きの実スコアなどで60未満なら高得点カウントしない（アプリ設定の「合格」相当）
    if (score >= 60) cHist.highScoreDates.push(now.toISOString());
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    cHist.highScoreDates = cHist.highScoreDates
      .map(v => new Date(v))
      .filter(d => !isNaN(d.getTime()) && d >= weekAgo)
      .map(d => d.toISOString());
    const qHistId = req.questionId;
    if (qHistId) {
      if (!unitHistory[qHistId]) unitHistory[qHistId] = { results: [], times: [] };
      unitHistory[qHistId].results.push(req.questionCorrect === true ? 1 : 0);
      if (unitHistory[qHistId].results.length > 10) unitHistory[qHistId].results.shift();
      unitHistory[qHistId].times.push(typeof req.timeSec === "number" ? req.timeSec : 0);
      if (unitHistory[qHistId].times.length > 10) unitHistory[qHistId].times.shift();
    }
  } else {
    const resultsList = Array.isArray(req.results) ? req.results : [];
    resultsList.forEach(res => {
      if (res.isCorrect) {
        let qPoint = Math.max(1, (Number(res.basePoint) || 2) - (Number(res.maxDeduction) || 0));
        sessionRawPoints += qPoint;
      }
      const qId = res.questionId;
      if (!unitHistory[qId]) unitHistory[qId] = { results: [], times: [] };
      
      unitHistory[qId].results.push(res.isCorrect ? 1 : 0);
      if (unitHistory[qId].results.length > 10) unitHistory[qId].results.shift();
      unitHistory[qId].times.push(res.timeSec);
      if (unitHistory[qId].times.length > 10) unitHistory[qId].times.shift();
    });
  }

  let earnedPoints = Math.round((sessionRawPoints * multiplier) * 100) / 100;
  const sheetPointPercent = parseUnitSheetPointPercent_(req.unitSheetName);
  if (sheetPointPercent !== 100) {
    earnedPoints = Math.round(earnedPoints * (sheetPointPercent / 100) * 100) / 100;
  }
  const newTotalPoints = Math.round((userData.points + earnedPoints) * 100) / 100;
  
  userData.dailyPointsJson[todayStr] = (userData.dailyPointsJson[todayStr] || 0) + earnedPoints;
  userData.dailyPointsJson[todayStr] = Math.round(userData.dailyPointsJson[todayStr] * 100) / 100;
  if (lastStudyKey) userData.lastStudyJson[lastStudyKey] = now.toISOString();

  // ★ 特訓ルートのステップをクリアした場合は、今日の進捗にチェックを入れる（メニューID別）
  if (req.trainingStepIndex) {
    migrateTrainingProgressIfNeeded_(userData.trainingProgressJson);
    let mid = parseInt(req.trainingMenuId, 10);
    if (isNaN(mid) || mid < 1 || mid > 12) mid = 1;
    const midStr = String(mid);
    if (!userData.trainingProgressJson[todayStr]) userData.trainingProgressJson[todayStr] = {};
    if (!userData.trainingProgressJson[todayStr][midStr]) userData.trainingProgressJson[todayStr][midStr] = {};
    userData.trainingProgressJson[todayStr][midStr][req.trainingStepIndex] = true;
  }

  usersSheet.getRange(targetRowIdx, 4).setValue(newTotalPoints);
  usersSheet.getRange(targetRowIdx, 5).setValue(JSON.stringify(userData.lastStudyJson));
  usersSheet.getRange(targetRowIdx, 6).setValue(JSON.stringify(userData.historyJson));
  usersSheet.getRange(targetRowIdx, 7).setValue(JSON.stringify(userData.dailyPointsJson));
  usersSheet.getRange(targetRowIdx, 8).setValue(JSON.stringify(userData.trainingProgressJson)); // ★ 保存

  return sendResponse({ status: "success", earnedPoints: earnedPoints, newTotal: newTotalPoints, historyJson: userData.historyJson, dailyPointsJson: userData.dailyPointsJson, bonusApplied: req.isRandom, trainingProgressJson: userData.trainingProgressJson, sheetPointPercent: sheetPointPercent });
}

/** 漢字ニガテ：historyJson.__kanjiWeak へ薄いシグナルだけマージ（キー数・recent 上限あり） */
var KANJI_WEAK_MAX_KEYS_ = 200;
var KANJI_WEAK_RECENT_MAX_ = 12;

function kanjiWeakMakeKey_(modeId, unitName, setId, kanji) {
  return [String(modeId || ""), String(unitName || ""), String(setId || ""), String(kanji || "")].join("\x1f");
}

function pruneKanjiWeakIfNeeded_(weakRoot) {
  const keys = Object.keys(weakRoot);
  if (keys.length <= KANJI_WEAK_MAX_KEYS_) return;
  const scored = keys.map(function (k) {
    const r = weakRoot[k] || {};
    return { k: k, t: String(r.updatedAt || "") };
  });
  scored.sort(function (a, b) {
    return a.t.localeCompare(b.t);
  });
  const drop = scored.length - KANJI_WEAK_MAX_KEYS_;
  for (let i = 0; i < drop; i++) delete weakRoot[scored[i].k];
}

function mergeKanjiWeakFromRequest_(userData, req, nowIso) {
  const kanji = String(req.kanji || "").trim();
  if (!kanji) return { ok: false, message: "kanji が空です" };
  const modeId = String(req.modeId || "").trim();
  const unitName = String(req.unitName || "").trim();
  const setId = String(req.setId || "").trim();
  const signal = String(req.signal || "hand_analytics");
  if (!userData.historyJson.__kanjiWeak) userData.historyJson.__kanjiWeak = {};
  const weakRoot = userData.historyJson.__kanjiWeak;
  const key = kanjiWeakMakeKey_(modeId, unitName, setId, kanji);
  if (!weakRoot[key]) {
    weakRoot[key] = {
      modeId: modeId,
      unitName: unitName,
      setId: setId,
      kanji: kanji,
      strokeOrderFails: 0,
      brushFails: 0,
      strokeCountFails: 0,
      readingFails: 0,
      lastRefStrokeCount: null,
      recent: [],
      updatedAt: nowIso
    };
  }
  const row = weakRoot[key];
  if (signal === "hand_analytics") {
    if (req.strokeCountMismatch === true) row.strokeCountFails++;
    if (req.hasStrokeOrderIssue === true) row.strokeOrderFails++;
    if (req.brushEndingAllOk === false) row.brushFails++;
    if (req.referenceStrokeCount != null && !isNaN(Number(req.referenceStrokeCount))) {
      row.lastRefStrokeCount = Number(req.referenceStrokeCount);
    }
  } else if (signal === "reading_mistake") {
    row.readingFails++;
  }
  const ev = {
    at: nowIso,
    signal: signal,
    q: String(req.questionId || "")
  };
  if (req.hasStrokeOrderIssue != null) ev.hso = !!req.hasStrokeOrderIssue;
  if (req.brushEndingAllOk != null) ev.bok = !!req.brushEndingAllOk;
  if (req.strokeCountMismatch != null) ev.scm = !!req.strokeCountMismatch;
  if (!Array.isArray(row.recent)) row.recent = [];
  row.recent.push(ev);
  while (row.recent.length > KANJI_WEAK_RECENT_MAX_) row.recent.shift();
  row.updatedAt = nowIso;
  pruneKanjiWeakIfNeeded_(weakRoot);
  return { ok: true };
}

function handleAppendKanjiWeakSignals(req) {
  const userId = req.userId;
  if (!userId) return sendResponse({ status: "error", message: "userId が必要です" });
  const props = PropertiesService.getScriptProperties();
  const adminSs = SpreadsheetApp.openById(props.getProperty("ADMIN_SS_ID"));
  const usersSheet = adminSs.getSheetByName("users");
  const data = usersSheet.getDataRange().getValues();
  let targetRowIdx = -1;
  let userData = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      targetRowIdx = i + 1;
      userData = {
        points: Number(data[i][3]) || 0,
        lastStudyJson: JSON.parse(data[i][4] || "{}"),
        historyJson: JSON.parse(data[i][5] || "{}"),
        dailyPointsJson: JSON.parse(data[i][6] || "{}"),
        trainingProgressJson: JSON.parse(data[i][7] || "{}")
      };
      break;
    }
  }
  if (targetRowIdx === -1) return sendResponse({ status: "error", message: "ユーザーが見つかりません" });
  const nowIso = new Date().toISOString();
  const r = mergeKanjiWeakFromRequest_(userData, req, nowIso);
  if (!r.ok) return sendResponse({ status: "error", message: r.message || "マージ失敗" });
  usersSheet.getRange(targetRowIdx, 6).setValue(JSON.stringify(userData.historyJson));
  return sendResponse({ status: "success", historyJson: userData.historyJson });
}

function handleGetKanjiWeakReviewPlan(req) {
  const userId = req.userId;
  if (!userId) return sendResponse({ status: "error", message: "userId が必要です" });
  const props = PropertiesService.getScriptProperties();
  const adminSs = SpreadsheetApp.openById(props.getProperty("ADMIN_SS_ID"));
  const usersSheet = adminSs.getSheetByName("users");
  const data = usersSheet.getDataRange().getValues();
  let historyJson = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      historyJson = JSON.parse(data[i][5] || "{}");
      break;
    }
  }
  if (!historyJson) return sendResponse({ status: "error", message: "ユーザーが見つかりません" });
  const weakRoot = historyJson.__kanjiWeak || {};
  const modeId = String(req.modeId || "");
  const unitNames = Array.isArray(req.unitNames) ? req.unitNames.map(function (x) { return String(x); }) : [];
  const setIds = Array.isArray(req.setIds) ? req.setIds.map(function (x) { return String(x); }) : [];
  const setScopes = Array.isArray(req.setScopes)
    ? req.setScopes.map(function (x) {
        return { unitName: String((x && x.unitName) || ""), setId: String((x && x.setId) || "") };
      })
      .filter(function (x) { return x.unitName && x.setId; })
    : [];
  const trainMode = String(req.trainMode || "stroke_order");
  const rows = [];
  Object.keys(weakRoot).forEach(function (k) {
    const r = weakRoot[k];
    if (!r || !r.kanji) return;
    if (modeId && String(r.modeId) !== modeId) return;
    if (unitNames.length && unitNames.indexOf(String(r.unitName)) < 0) return;
    if (setScopes.length) {
      var inScope = setScopes.some(function (sc) {
        return String(r.unitName) === sc.unitName && String(r.setId) === sc.setId;
      });
      if (!inScope) return;
    } else if (setIds.length && setIds.indexOf(String(r.setId)) < 0) return;
    let w = 0;
    if (trainMode === "reading") w = Number(r.readingFails) || 0;
    else if (trainMode === "brush") w = Number(r.brushFails) || 0;
    else if (trainMode === "stroke_count") w = Number(r.strokeCountFails) || 0;
    else w = Number(r.strokeOrderFails) || 0;
    rows.push({
      modeId: r.modeId,
      unitName: r.unitName,
      setId: r.setId,
      kanji: r.kanji,
      w: w,
      strokeOrderFails: r.strokeOrderFails || 0,
      brushFails: r.brushFails || 0,
      strokeCountFails: r.strokeCountFails || 0,
      readingFails: r.readingFails || 0,
      lastRefStrokeCount: r.lastRefStrokeCount
    });
  });
  rows.sort(function (a, b) { return b.w - a.w; });
  return sendResponse({ status: "success", rows: rows.slice(0, 300) });
}

// ==========================================
// 既存のハンドラー（変更なし）
// ==========================================
function handleGetAppSettings(req) { const settingsSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("アプリ設定"); if (!settingsSheet) return sendResponse({ status: "success", settings: {} }); const data = settingsSheet.getDataRange().getValues(); const settings = {}; for (let i = 1; i < data.length; i++) { if (data[i][0]) settings[data[i][0]] = data[i][1]; } return sendResponse({ status: "success", settings: settings }); }
function handleGetExternalLearning(req) { const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("外部学習"); const list = []; if (sheet) { const data = sheet.getDataRange().getValues(); if (data.length > 0) { const headers = data[0].map(String); const isNew = headers[0] === "カテゴリ"; for (let i = 1; i < data.length; i++) { if (!data[i][0]) continue; if (isNew) { list.push({ category: data[i][0], volume: data[i][1], points: Number(data[i][2]) }); } else { list.push({ category: data[i][0], volume: "", points: Number(data[i][1]) }); } } } } return sendResponse({ status: "success", list: list }); }

function ensureExternalLearningRequestSheet_(adminSs) {
  let sheet = adminSs.getSheetByName("外部学習申請");
  if (!sheet) {
    sheet = adminSs.insertSheet("外部学習申請");
    sheet.appendRow(["申請日時", "ユーザーID", "ユーザー名", "カテゴリ", "分量", "ポイント", "こどもメモ", "状態", "処理日時", "おとなメモ"]);
  }
  return sheet;
}

function getExternalLearningAdminPin_(adminSs) {
  const sheet = adminSs.getSheetByName("アプリ設定");
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === "外部学習_管理者PIN") return String(data[i][1] || "");
  }
  return "";
}

function verifyExternalAdminPin_(adminSs, adminPin) {
  const expected = getExternalLearningAdminPin_(adminSs);
  if (!expected) return { ok: false, message: "管理者PINがアプリ設定に登録されていません。「外部学習_管理者PIN」を追加してください。" };
  if (String(adminPin) !== String(expected)) return { ok: false, message: "管理者PINが一致しません" };
  return { ok: true };
}

function validateExternalMenu_(adminSs, menuName, points) {
  const sheet = adminSs.getSheetByName("外部学習");
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return false;
  const headers = data[0].map(String);
  const isNew = headers[0] === "カテゴリ";
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (isNew) {
      if (String(data[i][0]) === String(menuName.category) && String(data[i][1]) === String(menuName.volume) && Number(data[i][2]) === Number(points)) return true;
    } else {
      if (String(data[i][0]) === String(menuName.category || menuName) && Number(data[i][1]) === Number(points)) return true;
    }
  }
  return false;
}

function ensureExternalRequestHeaderMap_(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const map = {};
  ["申請日時", "ユーザーID", "ユーザー名", "カテゴリ", "分量", "ポイント", "こどもメモ", "状態", "処理日時", "おとなメモ"].forEach(key => {
    map[key] = header.indexOf(key) + 1; // 1-based col
  });
  const required = ["申請日時", "ユーザーID", "ユーザー名", "ポイント", "状態"];
  for (const k of required) {
    if (map[k] === 0) return { ok: false, map: null, message: "「外部学習申請」シートのヘッダーが最新ではありません。カテゴリ/分量/こどもメモ/おとなメモ列を追加してください。" };
  }
  return { ok: true, map };
}

function handleSubmitExternalLearningRequest(req) {
  const adminSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID'));
  const usersSheet = adminSs.getSheetByName("users");
  const data = usersSheet.getDataRange().getValues();
  let userName = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === req.userId) {
      userName = String(data[i][1] || "");
      break;
    }
  }
  if (!userName) return sendResponse({ status: "error", message: "ユーザーが見つかりません" });
  if (!validateExternalMenu_(adminSs, { category: req.category, volume: req.volume }, req.points)) return sendResponse({ status: "error", message: "メニューが不正です" });

  const sheet = ensureExternalLearningRequestSheet_(adminSs);
  const { ok, map, message } = ensureExternalRequestHeaderMap_(sheet);
  if (!ok) return sendResponse({ status: "error", message });

  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const row = [];
  row[map["申請日時"] - 1] = nowStr;
  row[map["ユーザーID"] - 1] = req.userId;
  row[map["ユーザー名"] - 1] = userName;
  if (map["カテゴリ"]) row[map["カテゴリ"] - 1] = req.category;
  if (map["分量"]) row[map["分量"] - 1] = req.volume;
  row[map["ポイント"] - 1] = Number(req.points);
  if (map["こどもメモ"]) row[map["こどもメモ"] - 1] = req.childMemo || "";
  row[map["状態"] - 1] = "申請中";
  if (map["処理日時"]) row[map["処理日時"] - 1] = "";
  if (map["おとなメモ"]) row[map["おとなメモ"] - 1] = "";
  sheet.appendRow(row);
  const rowIdx = sheet.getLastRow();

  return sendResponse({ status: "success", message: "申請を受け付けました。おうちの人に承認してもらってね。", rowIdx: rowIdx });
}

function handleGetPendingExternalRequests(req) {
  const adminSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID'));
  const v = verifyExternalAdminPin_(adminSs, req.adminPin);
  if (!v.ok) return sendResponse({ status: "error", message: v.message });

  const sheet = ensureExternalLearningRequestSheet_(adminSs);
  const { ok, map, message } = ensureExternalRequestHeaderMap_(sheet);
  if (!ok) return sendResponse({ status: "error", message });
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map["状態"] - 1]) === "申請中") {
      list.push({
        rowIdx: i + 1,
        requestedAt: String(data[i][map["申請日時"] - 1] || ""),
        userId: String(data[i][map["ユーザーID"] - 1] || ""),
        userName: String(data[i][map["ユーザー名"] - 1] || ""),
        category: map["カテゴリ"] ? String(data[i][map["カテゴリ"] - 1] || "") : "",
        volume: map["分量"] ? String(data[i][map["分量"] - 1] || "") : "",
        points: Number(data[i][map["ポイント"] - 1]) || 0,
        childMemo: map["こどもメモ"] ? String(data[i][map["こどもメモ"] - 1] || "") : ""
      });
    }
  }
  return sendResponse({ status: "success", list: list });
}

function handleApproveExternalRequest(req) {
  const adminSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID'));
  const v = verifyExternalAdminPin_(adminSs, req.adminPin);
  if (!v.ok) return sendResponse({ status: "error", message: v.message });

  const sheet = ensureExternalLearningRequestSheet_(adminSs);
  const rowIdx = Number(req.rowIdx);
  const lastRow = sheet.getLastRow();
  if (rowIdx < 2 || rowIdx > lastRow) return sendResponse({ status: "error", message: "申請が見つかりません" });

  const { ok, map, message } = ensureExternalRequestHeaderMap_(sheet);
  if (!ok) return sendResponse({ status: "error", message });
  const row = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (String(row[map["状態"] - 1]) !== "申請中") return sendResponse({ status: "error", message: "この申請はすでに処理済みです" });

  const userId = String(row[map["ユーザーID"] - 1]);
  const points = Number(row[map["ポイント"] - 1]) || 0;
  const menuName = map["カテゴリ"] ? String(row[map["カテゴリ"] - 1]) : "";
  const userName = String(row[map["ユーザー名"] - 1] || "");

  const usersSheet = adminSs.getSheetByName("users");
  const udata = usersSheet.getDataRange().getValues();
  let found = false;
  let newTotal = 0;
  for (let i = 1; i < udata.length; i++) {
    if (udata[i][0] === userId) {
      found = true;
      newTotal = Math.round(((Number(udata[i][3]) || 0) + points) * 100) / 100;
      usersSheet.getRange(i + 1, 4).setValue(newTotal);
      break;
    }
  }
  if (!found) return sendResponse({ status: "error", message: "ユーザーが見つかりません" });

  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(rowIdx, map["状態"]).setValue("承認済み");
  if (map["処理日時"]) sheet.getRange(rowIdx, map["処理日時"]).setValue(nowStr);
  if (map["おとなメモ"]) sheet.getRange(rowIdx, map["おとなメモ"]).setValue(req.adminMemo || "");

  return sendResponse({ status: "success", message: "承認してポイントを付与しました。", newTotal: newTotal, userId: userId, userName: userName, category: menuName, volume: map["分量"] ? String(row[map["分量"] - 1]) : "", points: points });
}

function handleRejectExternalRequest(req) {
  const adminSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID'));
  const v = verifyExternalAdminPin_(adminSs, req.adminPin);
  if (!v.ok) return sendResponse({ status: "error", message: v.message });

  const sheet = ensureExternalLearningRequestSheet_(adminSs);
  const rowIdx = Number(req.rowIdx);
  const lastRow = sheet.getLastRow();
  if (rowIdx < 2 || rowIdx > lastRow) return sendResponse({ status: "error", message: "申請が見つかりません" });

  const { ok, map, message } = ensureExternalRequestHeaderMap_(sheet);
  if (!ok) return sendResponse({ status: "error", message });

  const status = String(sheet.getRange(rowIdx, map["状態"]).getValue());
  if (status !== "申請中") return sendResponse({ status: "error", message: "この申請はすでに処理済みです" });

  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(rowIdx, map["状態"]).setValue("却下");
  if (map["処理日時"]) sheet.getRange(rowIdx, map["処理日時"]).setValue(nowStr);
  if (map["おとなメモ"]) sheet.getRange(rowIdx, map["おとなメモ"]).setValue(req.adminMemo || "");

  return sendResponse({ status: "success", message: "却下しました。" });
}

function handleGetMyExternalLearningRequests(req) {
  const adminSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID'));
  const sheet = ensureExternalLearningRequestSheet_(adminSs);
  const { ok, map, message } = ensureExternalRequestHeaderMap_(sheet);
  if (!ok) return sendResponse({ status: "error", message });
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = data.length - 1; i >= 1 && list.length < 30; i--) {
    if (String(data[i][map["ユーザーID"] - 1]) === String(req.userId)) {
      list.push({
        rowIdx: i + 1,
        requestedAt: String(data[i][map["申請日時"] - 1] || ""),
        category: map["カテゴリ"] ? String(data[i][map["カテゴリ"] - 1] || "") : "",
        volume: map["分量"] ? String(data[i][map["分量"] - 1] || "") : "",
        points: Number(data[i][map["ポイント"] - 1]) || 0,
        status: String(data[i][map["状態"] - 1] || ""),
        childMemo: map["こどもメモ"] ? String(data[i][map["こどもメモ"] - 1] || "") : ""
      });
    }
  }
  return sendResponse({ status: "success", list: list });
}
function handleGetPointsMultiplier(req) { const data = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("users").getDataRange().getValues(); let multiplier = 1.0; for (let i = 1; i < data.length; i++) { if (data[i][0] === req.userId) { const lastStudyTimeStr = JSON.parse(data[i][4] || "{}")[req.unitId]; if (lastStudyTimeStr) { const diffHours = (new Date() - new Date(lastStudyTimeStr)) / (1000 * 60 * 60); let basePercent = 10 + Math.floor(diffHours / 2) * 10; if (basePercent > 100) basePercent = 100; multiplier = basePercent / 100; } break; } } return sendResponse({ status: "success", multiplier: multiplier }); }
function handleGetChildUsers(req) { const data = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("users").getDataRange().getValues(); const users = []; for (let i = 1; i < data.length; i++) { if (data[i][0] && i > 0) users.push({ id: data[i][0], name: data[i][1] }); } return sendResponse({ status: "success", users: users }); }
function handleVerifyKidPin(req) { const data = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("users").getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === req.userId) { if (String(data[i][2]) === String(req.pin)) { return sendResponse({ status: "success", user: { id: data[i][0], name: data[i][1], points: data[i][3], lastStudyJson: JSON.parse(data[i][4] || "{}"), historyJson: JSON.parse(data[i][5] || "{}"), dailyPointsJson: JSON.parse(data[i][6] || "{}") }, message: "ログイン成功" }); } else return sendResponse({ status: "error", message: "PINがちがいます" }); } } return sendResponse({ status: "error", message: "ユーザーが見つかりません" }); }
function handleChangePin(req) { const usersSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("users"); const data = usersSheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === req.userId) { usersSheet.getRange(i + 1, 3).setValue(req.newPin); return sendResponse({ status: "success", message: "新しいPINをセットしました！" }); } } return sendResponse({ status: "error", message: "ユーザーが見つかりません" }); }
function handleGetMaterialsList(req) {
  const props = PropertiesService.getScriptProperties();
  const materials = [];
  const pushFolderFiles = (folderId, category) => {
    if (!folderId) return;
    let folder;
    try { folder = DriveApp.getFolderById(folderId); } catch (_) { return; }
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const file = files.next();
      materials.push({
        modeId: file.getId(),
        modeName: file.getName(),
        category: category,
        units: SpreadsheetApp.open(file).getSheets().map(s => s.getName())
      });
    }
  };
  pushFolderFiles(props.getProperty('MATERIALS_FOLDER_ID'), "english");
  pushFolderFiles(props.getProperty('KANJI_MATERIALS_FOLDER_ID'), "kanji");
  return sendResponse({ status: "success", materials: materials });
}
function handleGetQuestions(req) { const data = SpreadsheetApp.openById(req.modeId).getSheetByName(req.unitName).getDataRange().getValues(); const headers = data[0]; const questions = []; for (let i = 1; i < data.length; i++) { let qObj = {}; for (let j = 0; j < headers.length; j++) qObj[headers[j]] = data[i][j]; if (qObj["通し番号"]) questions.push(qObj); } return sendResponse({ status: "success", questions: questions }); }
function handleGetRewards(req) { const data = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("rewards").getDataRange().getValues(); const rewards = []; for (let i = 1; i < data.length; i++) { if (data[i][0] && i > 0) rewards.push({ id: data[i][0], name: data[i][1], points: Number(data[i][2]), desc: data[i][3] }); } return sendResponse({ status: "success", rewards: rewards }); }
function handleExchangeReward(req) { const adminSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')); const usersSheet = adminSs.getSheetByName("users"); const usersData = usersSheet.getDataRange().getValues(); let userRow = -1; let currentPoints = 0; for (let i = 1; i < usersData.length; i++) { if (usersData[i][0] === req.userId) { userRow = i + 1; currentPoints = Number(usersData[i][3]) || 0; break; } } if (userRow === -1) return sendResponse({ status: "error", message: "ユーザーが見つかりません" }); const rewardsData = adminSs.getSheetByName("rewards").getDataRange().getValues(); let rewardData = null; for (let i = 1; i < rewardsData.length; i++) { if (rewardsData[i][0] === req.rewardId) { rewardData = { name: rewardsData[i][1], points: Number(rewardsData[i][2]) }; break; } } if (currentPoints < rewardData.points) return sendResponse({ status: "error", message: "ポイントが足りません" }); const newPoints = Math.round((currentPoints - rewardData.points) * 100) / 100; usersSheet.getRange(userRow, 4).setValue(newPoints); adminSs.getSheetByName("inventory").appendRow([new Date().toLocaleString(), req.userId, req.rewardId, rewardData.name, "未消化"]); return sendResponse({ status: "success", newPoints: newPoints, message: `${rewardData.name} をゲットしました！` }); }
function handleGetInventory(req) { const data = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("inventory").getDataRange().getValues(); const inventory = []; for (let i = 1; i < data.length; i++) { if (data[i][1] === req.userId) { inventory.push({ rowIdx: i + 1, date: data[i][0], rewardName: data[i][3], status: data[i][4] }); } } inventory.sort((a, b) => b.rowIdx - a.rowIdx); return sendResponse({ status: "success", inventory: inventory }); }
function handleConsumeReward(req) { SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("inventory").getRange(req.rowIdx, 5).setValue("使用済み"); return sendResponse({ status: "success", message: "景品をつかいました！" }); }

// ★設定を消去する緊急用関数（不要なら後で消してもOK）
function resetProperties() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log("古い設定をすべて消去しました！");
}

function handleGetKanjiInitData(req) {
  const prop = PropertiesService.getScriptProperties();
  const sheetId = prop.getProperty('KANJI_SHEET_ID');
  let targetSheetId = sheetId;
  if (!targetSheetId) {
    const folderId = prop.getProperty('KANJI_MATERIALS_FOLDER_ID');
    if (folderId) {
      try {
        const files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_SHEETS);
        if (files.hasNext()) targetSheetId = files.next().getId();
      } catch (_) {}
    }
  }
  if (!targetSheetId) return sendResponse({ status: "error", message: "漢字教材が未設定です（KANJI_SHEET_ID または 教材フォルダを確認）" });
  try {
    const ss = SpreadsheetApp.openById(targetSheetId);
    const sheets = ss.getSheets().map(s => s.getName());
    return sendResponse({ status: "success", bookName: ss.getName(), sheets: sheets });
  } catch (e) {
    return sendResponse({ status: "error", message: "漢字データにアクセスできません: " + e.message });
  }
}

function handleGetKanjiDataFromSheet(req) {
  const sheetName = String(req.sheetName || "");
  if (!sheetName) return sendResponse({ status: "error", message: "sheetName が未指定です" });
  const prop = PropertiesService.getScriptProperties();
  const sheetId = prop.getProperty('KANJI_SHEET_ID');
  let targetSheetId = sheetId;
  if (!targetSheetId) {
    const folderId = prop.getProperty('KANJI_MATERIALS_FOLDER_ID');
    if (folderId) {
      try {
        const files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_SHEETS);
        if (files.hasNext()) targetSheetId = files.next().getId();
      } catch (_) {}
    }
  }
  if (!targetSheetId) return sendResponse({ status: "error", message: "漢字教材が未設定です（KANJI_SHEET_ID または 教材フォルダを確認）" });
  try {
    const ss = SpreadsheetApp.openById(targetSheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return sendResponse({ status: "success", data: {} });
    const values = sheet.getDataRange().getValues();
    const kanjiMap = {};
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const kanji = row[0];
      if (!kanji || String(kanji).trim().length !== 1) continue;
      const paths = [];
      for (let j = 2; j < row.length; j++) {
        const cellValue = row[j];
        if (!cellValue) continue;
        const strVal = String(cellValue).trim();
        if (!strVal) continue;
        if (strVal.indexOf('|') >= 0) {
          strVal.split('|').forEach(p => {
            const cleaned = String(p || "").trim();
            if (cleaned && (cleaned.charAt(0) === 'M' || cleaned.charAt(0) === 'm')) paths.push(cleaned);
          });
        } else if (strVal.charAt(0) === 'M' || strVal.charAt(0) === 'm') {
          paths.push(strVal);
        }
      }
      if (paths.length > 0) kanjiMap[String(kanji).trim()] = paths;
    }
    return sendResponse({ status: "success", data: kanjiMap });
  } catch (e) {
    return sendResponse({ status: "error", message: "漢字データ取得に失敗しました: " + e.message });
  }
}

function parseKanjiQuizSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return { groups: [] };
  const headers = values[0].map(v => String(v || "").trim());
  const cellText = (v) => {
    const s = String(v || "").trim();
    if (!s || s === "×" || s === "x" || s === "X") return "";
    return s;
  };
  const idxSet = headers.indexOf("セット");
  const idxKanji = headers.indexOf("漢字");
  if (idxSet < 0 || idxKanji < 0) {
    throw new Error("漢字クイズシートの見出しに「セット」「漢字」が必要です。");
  }

  const readingDefs = [];
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    const m = h.match(/^(訓読み|音読み)([A-ZＡ-Ｚ])_読み$/);
    if (!m) continue;
    const kind = m[1] === "訓読み" ? "訓" : "音";
    const label = m[2].toUpperCase();
    const exIdx = [];
    for (let j = 0; j < headers.length; j++) {
      const ex = headers[j];
      if (ex.indexOf(kind + label + "_例文") === 0) exIdx.push(j);
    }
    readingDefs.push({ label: kind + label, readingIdx: i, exampleIdx: exIdx });
  }
  if (readingDefs.length === 0) {
    throw new Error("見出しが新形式ではありません。訓読みA_読み / 音読みA_読み の列が必要です。");
  }

  const groupsMap = {};
  const order = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const setRaw = cellText(row[idxSet]);
    const kanji = cellText(row[idxKanji]);
    if (!setRaw || !kanji) continue;
    const setId = setRaw;
    if (!groupsMap[setId]) {
      groupsMap[setId] = { setId, items: [] };
      order.push(setId);
    }
    const readings = [];
    readingDefs.forEach(def => {
      const reading = cellText(row[def.readingIdx]);
      if (!reading) return;
      const examples = def.exampleIdx
        .map(i => cellText(row[i]))
        .filter(Boolean);
      const rk = def.label.indexOf("音") === 0 ? "on" : "kun";
      readings.push({ label: def.label, kind: rk, reading, examples });
    });
    groupsMap[setId].items.push({
      rowIndex: r + 1,
      kanji,
      readings
    });
  }
  const groups = order.map(setId => groupsMap[setId]).filter(g => g.items.length > 0);
  return { groups };
}

/**
 * 漢字クイズ3形式: 配列シャッフル（非破壊）
 */
function shuffleKanjiQuizArray_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = a[i];
    a[i] = a[j];
    a[j] = t;
  }
  return a;
}

/** ひらがな → カタカナ（音読み表示・正解用） */
function hiraganaToKatakanaKanjiQuiz_(s) {
  return Array.from(String(s || ""))
    .map(function (ch) {
      const c = ch.charCodeAt(0);
      if (c >= 0x3041 && c <= 0x3096) return String.fromCharCode(c + 0x60);
      return ch;
    })
    .join("");
}

/** 音=カタカナ・訓=シートのまま（ひらがな想定） */
function readingDisplayForQuiz_(reading, kind) {
  const r = String(reading || "");
  if (kind === "on") return hiraganaToKatakanaKanjiQuiz_(r);
  return r;
}

/** 正解読み（完全一致判定用）。音はカタカナ、訓は原文 */
function normalizedCorrectReadingAnswer_(reading, kind) {
  return readingDisplayForQuiz_(reading, kind);
}

/** 例文中で最初の kanji 列文字列をマスク（×等は素材に入らない前提） */
function maskKanjiInExampleOnce_(sentence, kanjiCol) {
  const s = String(sentence || "");
  const k = String(kanjiCol || "");
  if (!s || !k) return { ok: false, masked: s };
  const idx = s.indexOf(k);
  if (idx < 0) return { ok: false, masked: s };
  return { ok: true, masked: s.slice(0, idx) + "＿" + s.slice(idx + k.length) };
}

/** 送り仮名クイズ用: ひらがな・カタカナ・長音など（漢字直後の「かな連続」判定） */
function kanjiQuizIsKanaCharForOkurigana_(ch) {
  if (!ch || ch.length === 0) return false;
  const c = ch.codePointAt(0);
  return (
    (c >= 0x3040 && c <= 0x309f) ||
    (c >= 0x30a0 && c <= 0x30ff) ||
    (c >= 0xff65 && c <= 0xff9f) ||
    c === 0x30fc
  );
}

/** 次の語に続く漢字（CJK統合） */
function kanjiQuizIsHanForOkurigana_(ch) {
  if (!ch || ch.length === 0) return false;
  const c = ch.codePointAt(0);
  return (
    (c >= 0x4e00 && c <= 0x9fff) ||
    (c >= 0x3400 && c <= 0x4dbf) ||
    (c >= 0xf900 && c <= 0xfaff)
  );
}

/**
 * 次の漢字の直前までを「この語の表記レンジ」とみなす。
 * 例: 「幸せな時」で k=幸 → 表面は「幸せな」までかなが続くが、直後が漢字「時」のとき末尾の助詞ひらがな（な 等）をはずし「幸せ」に丸める。
 */
var KANJI_OKURIGANA_TRAILING_HIRAGANA_PARTICLE_ = {
  の: 1,
  な: 1,
  に: 1,
  が: 1,
  を: 1,
  は: 1,
  も: 1,
  と: 1,
  で: 1,
  へ: 1,
  や: 1,
  か: 1,
  さ: 1,
  よ: 1,
  ね: 1,
  ぞ: 1,
  わ: 1
};

function kanjiQuizIsTrailingHiraganaParticle_(ch) {
  return !!ch && ch.length === 1 && !!KANJI_OKURIGANA_TRAILING_HIRAGANA_PARTICLE_[ch];
}

/**
 * 例文 ex 内の最初の k 位置から、「次の漢字の手前まで」で切った表記ウィンドウ（先頭は k）。
 * @returns {{ i: number, end: number, win: string } | null}
 */
function kanjiOkuriganaSurfaceWindow_(ex, k) {
  const s = String(ex || "");
  const K = String(k || "");
  if (!K || K.length !== 1) return null;
  const i = s.indexOf(K);
  if (i < 0) return null;
  const afterK0 = i + K.length;
  let j = afterK0;
  while (j < s.length && kanjiQuizIsKanaCharForOkurigana_(s.charAt(j))) j++;
  let end = j;
  if (j < s.length && kanjiQuizIsHanForOkurigana_(s.charAt(j))) {
    let t = j;
    while (t > afterK0 && kanjiQuizIsTrailingHiraganaParticle_(s.charAt(t - 1))) t--;
    end = t;
  }
  return { i: i, end: end, win: s.slice(i, end) };
}

/** 候補表記が「語ウィンドウ」の先頭にアンカーされているか（全文 indexOf では な時 まで誤一致しない） */
function okuriganaCandAnchoredInExample_(ex, k, cand) {
  const a = kanjiOkuriganaSurfaceWindow_(ex, k);
  if (!a || !cand) return false;
  const win = a.win;
  return win.length >= cand.length && win.indexOf(cand) === 0;
}

function buildOkuriganaShiftQuizQuestion_(item) {
  const k = String(item.kanji || "");
  if (k.length !== 1) return null;
  const readings = (Array.isArray(item.readings) ? item.readings : []).filter(function (r) {
    return r.kind === "kun" && String(r.reading || "").length >= 2;
  });
  if (!readings.length) return null;
  const r = readings[Math.floor(Math.random() * readings.length)];
  const reading = String(r.reading || "");
  const examples = Array.isArray(r.examples) ? r.examples : [];
  let bestSplitPos = -1;
  for (let s = 1; s <= reading.length; s++) {
    const cand = k + reading.substring(s);
    if (examples.some(function (ex) { return okuriganaCandAnchoredInExample_(ex, k, cand); })) {
      bestSplitPos = s;
      break;
    }
  }
  if (bestSplitPos < 0) {
    if (!examples.length) bestSplitPos = 1;
    else return null;
  }
  const correct = k + reading.substring(bestSplitPos);
  /** 訓よみの各分割: ① かんじのみ（よみはルビ想定）… k+reading.substring(reading.length)==k 、②〜 かんじ+よみの後ろからの切り落とし */
  const orderedChoices = [];
  for (let sp = 1; sp <= reading.length; sp++) {
    orderedChoices.push(k + reading.substring(sp));
  }
  const uniq = shuffleKanjiQuizArray_(orderedChoices.slice()).filter(function (c) {
    return !!c;
  });
  if (uniq.length < 2) return null;
  const searchParts = [k, reading, r.label].concat(uniq).join(" ");
  return {
    type: "okurigana_shift",
    kanji: k,
    rowIndex: item.rowIndex,
    readingLabel: r.label,
    readingKind: "kun",
    readingHint: reading,
    prompt:
      "「" +
      reading +
      "」の ただしい ひょうきを えらびましょう。（かんじ＋おくりがなの つながりの パターン ぜんぶ です）",
    choices: uniq,
    correctAnswer: correct,
    searchText: searchParts
  };
}

function buildRubyToKanjiQuizQuestion_(item) {
  const k = String(item.kanji || "");
  if (!k) return null;
  const pairs = [];
  (Array.isArray(item.readings) ? item.readings : []).forEach(function (r) {
    const examples = Array.isArray(r.examples) ? r.examples : [];
    examples.forEach(function (ex) {
      const exs = String(ex || "");
      if (exs.indexOf(k) >= 0) pairs.push({ r: r, ex: exs });
    });
  });
  if (!pairs.length) return null;
  const pick = pairs[Math.floor(Math.random() * pairs.length)];
  const masked = maskKanjiInExampleOnce_(pick.ex, k);
  if (!masked.ok) return null;
  const readingDisp = readingDisplayForQuiz_(pick.r.reading, pick.r.kind);
  const searchParts = [k, readingDisp, pick.r.label, masked.masked, pick.ex].join(" ");
  return {
    type: "ruby_to_kanji",
    kanji: k,
    rowIndex: item.rowIndex,
    readingKind: pick.r.kind,
    readingLabel: pick.r.label,
    readingDisplay: readingDisp,
    maskedSentence: masked.masked,
    prompt: "読みと例文の空欄の漢字を、筆順どおりに手書きしましょう（各字とも60点以上）。",
    correctAnswer: k,
    searchText: searchParts
  };
}

function buildSentenceToRubyQuizQuestion_(item) {
  const k = String(item.kanji || "");
  if (!k) return null;
  const pairs = [];
  (Array.isArray(item.readings) ? item.readings : []).forEach(function (r) {
    const examples = Array.isArray(r.examples) ? r.examples : [];
    examples.forEach(function (ex) {
      const exs = String(ex || "");
      if (exs.indexOf(k) >= 0) pairs.push({ r: r, ex: exs });
    });
  });
  if (!pairs.length) return null;
  const pick = pairs[Math.floor(Math.random() * pairs.length)];
  const masked = maskKanjiInExampleOnce_(pick.ex, k);
  if (!masked.ok) return null;
  const ans = normalizedCorrectReadingAnswer_(pick.r.reading, pick.r.kind);
  var hintOn = pick.r.kind === "on" ? "音読みはカタカナ" : "訓読みはひらがな";
  const fullEx = String(pick.ex || "");
  const searchParts = [k, ans, pick.r.label, masked.masked, fullEx].join(" ");
  return {
    type: "sentence_to_ruby",
    kanji: k,
    rowIndex: item.rowIndex,
    readingKind: pick.r.kind,
    readingLabel: pick.r.label,
    readingDisplay: readingDisplayForQuiz_(pick.r.reading, pick.r.kind),
    sentence: fullEx,
    maskedSentence: masked.masked,
    prompt:
      "下のれいぶんのうち、赤いかんじ「" +
      k +
      "」のよみを、マスに手書きし、「文字起こし」→「こたえを決定」で答えましょう。（" +
      hintOn +
      "で書けたとき、せいかいならポイント2ばい）",
    correctAnswer: ans,
    searchText: searchParts
  };
}

/**
 * 3タイプを偏りなく混在（各バケットをシャッフル後ラウンドロビン）
 */
function mergeKanjiQuizBucketsBalanced_(buckets) {
  const order = ["okurigana_shift", "ruby_to_kanji", "sentence_to_ruby"];
  const queues = order.map(function (key) {
    return shuffleKanjiQuizArray_(buckets[key] || []).slice();
  });
  const out = [];
  var keepGoing = true;
  while (keepGoing) {
    keepGoing = false;
    queues.forEach(function (q) {
      if (q.length) {
        out.push(q.shift());
        keepGoing = true;
      }
    });
  }
  return out;
}

function buildKanjiQuizProblemList_(group) {
  const items = group.items || [];
  if (!items.length) return [];
  const buckets = { okurigana_shift: [], ruby_to_kanji: [], sentence_to_ruby: [] };
  items.forEach(function (item) {
    const o = buildOkuriganaShiftQuizQuestion_(item);
    if (o) buckets.okurigana_shift.push(o);
    const r2 = buildRubyToKanjiQuizQuestion_(item);
    if (r2) buckets.ruby_to_kanji.push(r2);
    const r3 = buildSentenceToRubyQuizQuestion_(item);
    if (r3) buckets.sentence_to_ruby.push(r3);
  });
  const merged = shuffleKanjiQuizArray_(mergeKanjiQuizBucketsBalanced_(buckets));
  return merged.map(function (q, i) {
    const base = Object.assign({}, q);
    base.questionIndex = i;
    base.questionId =
      "KANJI_Q_" + String(group.setId) + "_" + q.rowIndex + "_" + q.type + "_" + i;
    return base;
  });
}

/** 同一シートの解析結果を短時間キャッシュし、get_kanji_quiz_sets → get_kanji_quiz_questions の連続で Spreadsheet 再読みを避ける */
function kanjiQuizSheetParsedCacheKey_(modeId, unitName) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    String(modeId) + "\x1f" + String(unitName),
    Utilities.Charset.UTF_8
  );
  return "kq_sh_" + Utilities.base64EncodeWebSafe(digest).slice(0, 36);
}

function getKanjiQuizParsedFromSpreadsheet_(modeId, unitName) {
  const cache = CacheService.getScriptCache();
  const key = kanjiQuizSheetParsedCacheKey_(modeId, unitName);
  const hit = cache.get(key);
  if (hit) {
    try {
      return { parsed: JSON.parse(hit), sheetMissing: false };
    } catch (e) {}
  }
  const sheet = SpreadsheetApp.openById(modeId).getSheetByName(unitName);
  if (!sheet) return { parsed: null, sheetMissing: true };
  const parsed = parseKanjiQuizSheet_(sheet);
  try {
    cache.put(key, JSON.stringify(parsed), 300);
  } catch (e) {}
  return { parsed: parsed, sheetMissing: false };
}

function handleGetKanjiQuizSets(req) {
  const modeId = String(req.modeId || "").trim();
  const unitName = String(req.unitName || "").trim();
  if (!modeId || !unitName) return sendResponse({ status: "error", message: "modeId と unitName が必要です。" });
  try {
    const got = getKanjiQuizParsedFromSpreadsheet_(modeId, unitName);
    if (got.sheetMissing) return sendResponse({ status: "error", message: "指定シートが見つかりません。" });
    const parsed = got.parsed;
    const sets = parsed.groups.map(g => ({
      setId: g.setId,
      count: g.items.length,
      kanjiList: g.items.map(it => it.kanji)
    }));
    return sendResponse({ status: "success", sets });
  } catch (e) {
    return sendResponse({ status: "error", message: "漢字セット取得に失敗しました: " + e.message });
  }
}

function handleGetKanjiQuizQuestions(req) {
  const modeId = String(req.modeId || "").trim();
  const unitName = String(req.unitName || "").trim();
  const setId = String(req.setId || "").trim();
  if (!modeId || !unitName || !setId) return sendResponse({ status: "error", message: "modeId / unitName / setId が必要です。" });
  try {
    const got = getKanjiQuizParsedFromSpreadsheet_(modeId, unitName);
    if (got.sheetMissing) return sendResponse({ status: "error", message: "指定シートが見つかりません。" });
    const parsed = got.parsed;
    const group = parsed.groups.find(g => String(g.setId) === setId);
    if (!group) return sendResponse({ status: "error", message: "指定セットが見つかりません。" });
    const questions = buildKanjiQuizProblemList_(group);
    return sendResponse({
      status: "success",
      setId,
      questions: questions
    });
  } catch (e) {
    return sendResponse({ status: "error", message: "漢字問題取得に失敗しました: " + e.message });
  }
}

// ==========================================
// 漢字ニガテ（弱みシグナル・復習プラン）
// ==========================================
var KANJI_WEAK_MAX_KEYS_ = 120;
var KANJI_WEAK_RECENT_MAX_ = 6;

function kanjiWeakEntryKey_(modeId, unitName, setId, kanji) {
  return String(modeId) + "\x1f" + String(unitName) + "\x1f" + String(setId) + "\x1f" + String(kanji);
}

function trimKanjiWeakEntries_(weakRoot) {
  if (!weakRoot || typeof weakRoot !== "object") return;
  var ent = weakRoot.entries;
  if (!ent || typeof ent !== "object") return;
  var keys = Object.keys(ent);
  if (keys.length <= KANJI_WEAK_MAX_KEYS_) return;
  var scored = keys.map(function (k) {
    var e = ent[k] || {};
    var sum =
      (Number(e.w_strokeOrder) || 0) +
      (Number(e.w_brush) || 0) +
      (Number(e.w_strokeCount) || 0) +
      (Number(e.w_reading) || 0);
    return { k: k, s: sum };
  });
  scored.sort(function (a, b) {
    return a.s - b.s;
  });
  var drop = keys.length - KANJI_WEAK_MAX_KEYS_;
  for (var i = 0; i < drop && i < scored.length; i++) {
    delete ent[scored[i].k];
  }
}

/** 中央 KanjiVG シートから画数マップ（1文字→画数） */
function loadKanjiVgStrokeCounts_() {
  var prop = PropertiesService.getScriptProperties();
  var sheetId = prop.getProperty("KANJI_SHEET_ID");
  if (!sheetId) {
    var folderId = prop.getProperty("KANJI_MATERIALS_FOLDER_ID");
    if (folderId) {
      try {
        var files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_SHEETS);
        if (files.hasNext()) sheetId = files.next().getId();
      } catch (_) {}
    }
  }
  var out = {};
  if (!sheetId) return out;
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName("KanjiVG.txt");
    if (!sheet) return out;
    var values = sheet.getDataRange().getValues();
    for (var i = 0; i < values.length; i++) {
      var kanji = values[i][0];
      if (!kanji || String(kanji).trim().length !== 1) continue;
      var paths = [];
      for (var j = 2; j < values[i].length; j++) {
        var cellValue = values[i][j];
        if (!cellValue) continue;
        var strVal = String(cellValue).trim();
        if (!strVal) continue;
        if (strVal.indexOf("|") >= 0) {
          strVal.split("|").forEach(function (p) {
            var cleaned = String(p || "").trim();
            if (cleaned && (cleaned.charAt(0) === "M" || cleaned.charAt(0) === "m")) paths.push(cleaned);
          });
        } else if (strVal.charAt(0) === "M" || strVal.charAt(0) === "m") {
          paths.push(strVal);
        }
      }
      if (paths.length > 0) out[String(kanji).trim()] = paths.length;
    }
  } catch (_) {}
  return out;
}

function buildStrokeCountQuizQuestion_(item, strokeN, qIndex, setId) {
  var k = String(item.kanji || "");
  if (!k || strokeN < 1) return null;
  var pool = [];
  for (var d = -4; d <= 4; d++) {
    if (d === 0) continue;
    var v = strokeN + d;
    if (v >= 1 && v <= 40) pool.push(v);
  }
  pool = shuffleKanjiQuizArray_(pool).slice(0, 4);
  var choices = shuffleKanjiQuizArray_([strokeN].concat(pool));
  var uniq = [];
  var seen = {};
  choices.forEach(function (c) {
    var s = String(c);
    if (!seen[s]) {
      seen[s] = true;
      uniq.push(s);
    }
  });
  if (uniq.indexOf(String(strokeN)) < 0) uniq[0] = String(strokeN);
  return {
    type: "stroke_count",
    kanji: k,
    rowIndex: item.rowIndex,
    prompt: "この漢字は何画ですか？",
    choices: uniq,
    correctAnswer: String(strokeN),
    searchText: k + " " + strokeN,
    questionIndex: qIndex,
    questionId: "KANJI_Q_" + String(setId) + "_" + item.rowIndex + "_stroke_count_" + qIndex
  };
}

function handleAppendKanjiWeakSignals(req) {
  var userId = req.userId;
  var signals = req.signals;
  if (!userId) return sendResponse({ status: "error", message: "userId が必要です。" });
  if (!Array.isArray(signals) || signals.length === 0) {
    return sendResponse({ status: "success", merged: 0 });
  }
  if (signals.length > 24) signals = signals.slice(0, 24);

  var props = PropertiesService.getScriptProperties();
  var adminSs = SpreadsheetApp.openById(props.getProperty("ADMIN_SS_ID"));
  var usersSheet = adminSs.getSheetByName("users");
  var data = usersSheet.getDataRange().getValues();
  var targetRowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      targetRowIdx = i + 1;
      break;
    }
  }
  if (targetRowIdx === -1) return sendResponse({ status: "error", message: "ユーザーが見つかりません" });

  var historyJson = JSON.parse(data[targetRowIdx - 1][5] || "{}");
  if (!historyJson.__kanjiWeak || typeof historyJson.__kanjiWeak !== "object") {
    historyJson.__kanjiWeak = { v: 1, entries: {} };
  }
  var root = historyJson.__kanjiWeak;
  if (!root.entries || typeof root.entries !== "object") root.entries = {};

  var nowIso = new Date().toISOString();
  var merged = 0;

  signals.forEach(function (sig) {
    var modeId = String(sig.modeId || "").trim();
    var unitName = String(sig.unitName || "").trim();
    var setId = String(sig.setId != null ? sig.setId : "").trim();
    var kanji = String(sig.kanjiChar || sig.kanji || "").trim();
    if (!modeId || !unitName || !setId || !kanji || kanji.length !== 1) return;
    var key = kanjiWeakEntryKey_(modeId, unitName, setId, kanji);
    var e = root.entries[key];
    if (!e) {
      e = {
        modeId: modeId,
        unitName: unitName,
        setId: setId,
        kanji: kanji,
        w_strokeOrder: 0,
        w_brush: 0,
        w_strokeCount: 0,
        w_reading: 0,
        recent: []
      };
      root.entries[key] = e;
    }
    if (sig.hasStrokeOrderIssue === true) e.w_strokeOrder += 1;
    if (sig.brushEndingAllOk === false) e.w_brush += 1;
    if (sig.strokeCountMismatch === true) e.w_strokeCount += 1;
    if (sig.readingMistake === true) e.w_reading += 1;
    if (sig.strokeCountQuizWrong === true) e.w_strokeCount += 1;
    e.updatedAt = nowIso;
    var rec = {
      at: sig.at || nowIso,
      hso: !!sig.hasStrokeOrderIssue,
      bak: sig.brushEndingAllOk !== false,
      scm: !!sig.strokeCountMismatch,
      rm: !!sig.readingMistake
    };
    if (!Array.isArray(e.recent)) e.recent = [];
    e.recent.push(rec);
    while (e.recent.length > KANJI_WEAK_RECENT_MAX_) e.recent.shift();
    merged++;
  });

  trimKanjiWeakEntries_(root);
  usersSheet.getRange(targetRowIdx, 6).setValue(JSON.stringify(historyJson));
  return sendResponse({ status: "success", merged: merged, historyJson: historyJson });
}

function handleGetKanjiWeakReviewPlan(req) {
  var userId = req.userId;
  var modeId = String(req.modeId || "").trim();
  var unitName = String(req.unitName || "").trim();
  var setIds = Array.isArray(req.setIds) ? req.setIds.map(function (x) { return String(x); }) : [];
  var axis = String(req.nigateAxis || "stroke_order").trim();
  var limit = Math.max(1, Math.min(24, parseInt(req.limit, 10) || 12));

  if (!userId || !modeId || !unitName) {
    return sendResponse({ status: "error", message: "userId / modeId / unitName が必要です。" });
  }

  var props = PropertiesService.getScriptProperties();
  var adminSs = SpreadsheetApp.openById(props.getProperty("ADMIN_SS_ID"));
  var usersSheet = adminSs.getSheetByName("users");
  var data = usersSheet.getDataRange().getValues();
  var historyJson = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      historyJson = JSON.parse(data[i][5] || "{}");
      break;
    }
  }

  var weakRoot = historyJson.__kanjiWeak;
  var entries = weakRoot && weakRoot.entries ? weakRoot.entries : {};
  var axisField =
    axis === "brush"
      ? "w_brush"
      : axis === "stroke_count"
        ? "w_strokeCount"
        : axis === "reading"
          ? "w_reading"
          : "w_strokeOrder";

  var scored = [];
  Object.keys(entries).forEach(function (k) {
    var e = entries[k];
    if (!e) return;
    if (String(e.modeId) !== modeId) return;
    if (String(e.unitName) !== unitName) return;
    if (setIds.length && setIds.indexOf(String(e.setId)) < 0) return;
    var w = Number(e[axisField]) || 0;
    if (w < 1) return;
    scored.push({ e: e, w: w, kanji: String(e.kanji || "") });
  });
  scored.sort(function (a, b) {
    return b.w - a.w;
  });

  if (!scored.length) {
    return sendResponse({ status: "success", questions: [], message: "このスコープ・軸では弱みデータがありません。" });
  }

  var got = getKanjiQuizParsedFromSpreadsheet_(modeId, unitName);
  if (got.sheetMissing) return sendResponse({ status: "error", message: "指定シートが見つかりません。" });
  var parsed = got.parsed;
  var groups = (parsed.groups || []).filter(function (g) {
    return !setIds.length || setIds.indexOf(String(g.setId)) >= 0;
  });
  if (!groups.length) return sendResponse({ status: "error", message: "セットが見つかりません。" });

  var strokeMap = axis === "stroke_count" ? loadKanjiVgStrokeCounts_() : {};

  function findItemForKanji(kanji) {
    for (var gi = 0; gi < groups.length; gi++) {
      var g = groups[gi];
      for (var ii = 0; ii < (g.items || []).length; ii++) {
        if (String(g.items[ii].kanji) === String(kanji)) return { group: g, item: g.items[ii] };
      }
    }
    return null;
  }

  var questions = [];
  var qCounter = 0;
  for (var si = 0; si < scored.length && questions.length < limit; si++) {
    var kanji = scored[si].kanji;
    if (!kanji || kanji.length !== 1) continue;
    var found = findItemForKanji(kanji);
    if (!found) continue;
    var item = found.item;
    var setId = String(found.group.setId);
    var q = null;
    if (axis === "stroke_order" || axis === "brush") {
      q = buildRubyToKanjiQuizQuestion_(item);
    } else if (axis === "reading") {
      if (Math.random() < 0.5) q = buildSentenceToRubyQuizQuestion_(item);
      else q = buildOkuriganaShiftQuizQuestion_(item);
    } else if (axis === "stroke_count") {
      var sn = strokeMap[kanji];
      if (!sn) continue;
      q = buildStrokeCountQuizQuestion_(item, sn, qCounter, setId);
    }
    if (!q) continue;
    q.questionIndex = qCounter;
    q.questionId =
      q.questionId ||
      "KANJI_NIGATE_" + axis + "_" + setId + "_" + item.rowIndex + "_" + qCounter;
    q.nigateSourceSetId = setId;
    qCounter++;
    questions.push(q);
  }

  questions = shuffleKanjiQuizArray_(questions).slice(0, limit);
  return sendResponse({ status: "success", questions: questions, nigateAxis: axis });
}