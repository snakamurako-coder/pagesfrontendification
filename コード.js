function setupSystem() {
  const scriptId = ScriptApp.getScriptId();
  const gasFile = DriveApp.getFileById(scriptId);
  const parents = gasFile.getParents();
  let parentFolder = DriveApp.getRootFolder();
  if (parents.hasNext()) parentFolder = parents.next();

  const props = PropertiesService.getScriptProperties();
  let adminSsId = props.getProperty('ADMIN_SS_ID');
  let materialsFolderId = props.getProperty('MATERIALS_FOLDER_ID');
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
      ["基本Pt_ja_to_en_sort_sort_missing", 30]
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
    ["全員", "", "英語→英語", "音声", "ランダム", ""]
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
  const lastStudyTimeStr = userData.lastStudyJson[req.unitId];

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

  req.results.forEach(res => {
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

  let earnedPoints = Math.round((sessionRawPoints * multiplier) * 100) / 100;
  const sheetPointPercent = parseUnitSheetPointPercent_(req.unitSheetName);
  if (sheetPointPercent !== 100) {
    earnedPoints = Math.round(earnedPoints * (sheetPointPercent / 100) * 100) / 100;
  }
  const newTotalPoints = Math.round((userData.points + earnedPoints) * 100) / 100;
  
  userData.dailyPointsJson[todayStr] = (userData.dailyPointsJson[todayStr] || 0) + earnedPoints;
  userData.dailyPointsJson[todayStr] = Math.round(userData.dailyPointsJson[todayStr] * 100) / 100;
  userData.lastStudyJson[req.unitId] = now.toISOString();

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
function handleGetMaterialsList(req) { const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('MATERIALS_FOLDER_ID')); const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS); const materials = []; while (files.hasNext()) { const file = files.next(); materials.push({ modeId: file.getId(), modeName: file.getName(), units: SpreadsheetApp.open(file).getSheets().map(s => s.getName()) }); } return sendResponse({ status: "success", materials: materials }); }
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