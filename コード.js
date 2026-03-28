// ==========================================
// メインの通信処理 (フロントエンドからのリクエスト窓口)
// ==========================================
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const action = p.action;

    // --- ユーザー・設定関連 ---
    if (action === "get_child_users") {
      return createJsonResponse(getChildUsers());
    }
    if (action === "verify_kid_pin") {
      return createJsonResponse(verifyKidPin(p.userId, p.pin));
    }
    if (action === "change_pin") {
      return createJsonResponse(changePin(p.userId, p.newPin));
    }
    if (action === "get_app_settings") {
      return createJsonResponse(getAppSettings());
    }

    // --- 学習問題関連 ---
    if (action === "get_materials_list") {
      return createJsonResponse(getMaterialsList());
    }
    if (action === "get_questions") {
      return createJsonResponse(getQuestions(p.modeId, p.unitName));
    }

    // --- 特訓ルート関連 ---
    if (action === "get_training_route") {
      return createJsonResponse(getTrainingRoute(p.userId));
    }
    if (action === "update_route_progress") {
      return createJsonResponse(updateRouteProgress(p.userId, p.stepIndex));
    }

    // --- 景品・もちもの関連 ---
    if (action === "get_rewards") {
      return createJsonResponse(getRewards());
    }
    if (action === "exchange_reward") {
      return createJsonResponse(exchangeReward(p.userId, p.rewardId));
    }
    if (action === "get_inventory") {
      return createJsonResponse(getInventory(p.userId));
    }
    if (action === "consume_reward") {
      return createJsonResponse(consumeReward(p.rowIdx));
    }

    // --- 外部学習関連 ---
    if (action === "get_external_learning") {
      return createJsonResponse(getExternalLearning());
    }
    if (action === "report_external_learning") {
      return createJsonResponse(reportExternalLearning(p.userId, p.menuName, p.points));
    }

    return createJsonResponse({ status: "error", message: "不明なアクションです" });

  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

// JSONレスポンスを作成する共通関数
function createJsonResponse(responseData) {
  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// ① ユーザー・設定関連の処理
// ==========================================
function getChildUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ユーザー");
  if (!sheet) return { status: "error", message: "「ユーザー」シートがありません" };
  
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      users.push({ id: data[i][0], name: data[i][1] });
    }
  }
  return { status: "success", users: users };
}

function verifyKidPin(userId, pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ユーザー");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userId) {
      if (String(data[i][2]) === String(pin)) {
        return { 
          status: "success", 
          user: { id: data[i][0], name: data[i][1], points: data[i][3] || 0 } 
        };
      } else {
        return { status: "error", message: "暗証番号がちがいます" };
      }
    }
  }
  return { status: "error", message: "ユーザーが見つかりません" };
}

function changePin(userId, newPin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ユーザー");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userId) {
      sheet.getRange(i + 1, 3).setValue(newPin);
      return { status: "success", message: "暗証番号をへんこうしました！" };
    }
  }
  return { status: "error", message: "ユーザーが見つかりません" };
}

function getAppSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("設定");
  if (!sheet) return { status: "success", settings: {} };
  
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  return { status: "success", settings: settings };
}

// ==========================================
// ② 学習問題関連の処理
// ==========================================
function getMaterialsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const materialsMap = {};
  
  // 「単元_」で始まるシートを学習用シートとみなす
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.indexOf("単元_") === 0) {
      const parts = name.split("_");
      const modeName = parts[1] || "その他";
      const unitName = parts[2] || name;
      
      if (!materialsMap[modeName]) {
        materialsMap[modeName] = { modeId: modeName, modeName: modeName, units: [] };
      }
      materialsMap[modeName].units.push(unitName);
    }
  });
  
  const materials = Object.keys(materialsMap).map(key => materialsMap[key]);
  return { status: "success", materials: materials };
}

function getQuestions(modeId, unitName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;

  // 特訓ルートから呼ばれた場合（modeIdが空）は、単元名を含むシートを探す
  if (!modeId && unitName) {
    const sheets = ss.getSheets();
    sheet = sheets.find(s => s.getName().includes(unitName));
  } else {
    // 通常の学習から呼ばれた場合は完全一致で探す
    const targetSheetName = `単元_${modeId}_${unitName}`;
    sheet = ss.getSheetByName(targetSheetName) || ss.getSheetByName(unitName);
  }

  if (!sheet) return { status: "error", message: "問題シートが見つかりません" };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const questions = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    let q = {};
    for (let j = 0; j < headers.length; j++) {
      q[headers[j]] = data[i][j];
    }
    questions.push(q);
  }

  return { status: "success", questions: questions };
}

// ==========================================
// ③ 特訓ルート（今日のミッション）関連の処理
// ==========================================
function getTrainingRoute(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const routeSheet = ss.getSheetByName("特訓メニュー"); 
  if (!routeSheet) return { status: "error", message: "「特訓メニュー」シートが見つかりません" };

  const routeData = routeSheet.getDataRange().getValues();
  const routeHeaders = routeData[0];
  const routeList = [];

  for (let i = 1; i < routeData.length; i++) {
    if (!routeData[i][0]) continue;
    let r = {};
    for (let j = 0; j < routeHeaders.length; j++) {
      r[routeHeaders[j]] = routeData[i][j];
    }
    routeList.push(r);
  }

  const progress = {}; 
  const progressSheet = ss.getSheetByName("特訓進捗");
  
  if (progressSheet) {
    const pData = progressSheet.getDataRange().getValues();
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

    for (let i = 1; i < pData.length; i++) {
      const pUserId = pData[i][0];
      const pDate = pData[i][1];
      const pStep = pData[i][2];
      
      // スプレッドシートの日付オブジェクトを文字列に変換して比較
      let recordDateStr = "";
      if (pDate instanceof Date) {
        recordDateStr = Utilities.formatDate(pDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else {
        recordDateStr = String(pDate).split(" ")[0]; // 簡易的な文字列処理
      }

      if (pUserId == userId && recordDateStr === todayStr) {
         progress[pStep] = true;
      }
    }
  }

  return { status: "success", route: routeList, progress: progress };
}

function updateRouteProgress(userId, stepIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let progressSheet = ss.getSheetByName("特訓進捗");
  
  if (!progressSheet) {
    progressSheet = ss.insertSheet("特訓進捗");
    progressSheet.appendRow(["userId", "date", "stepIndex"]);
    progressSheet.setColumnWidth(2, 120);
  }

  const today = new Date();
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy/MM/dd");
  progressSheet.appendRow([userId, dateStr, stepIndex]);

  return { status: "success" };
}

// ==========================================
// ④ 景品・もちもの関連の処理
// ==========================================
function getRewards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("景品");
  if (!sheet) return { status: "error", message: "「景品」シートがありません" };
  
  const data = sheet.getDataRange().getValues();
  const rewards = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      rewards.push({ id: data[i][0], name: data[i][1], points: data[i][2], desc: data[i][3] });
    }
  }
  return { status: "success", rewards: rewards };
}

function exchangeReward(userId, rewardId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("ユーザー");
  const rewardSheet = ss.getSheetByName("景品");
  let invSheet = ss.getSheetByName("もちもの");
  
  if (!invSheet) {
    invSheet = ss.insertSheet("もちもの");
    invSheet.appendRow(["userId", "rewardId", "rewardName", "date", "status"]);
  }

  const uData = userSheet.getDataRange().getValues();
  let userRow = -1;
  let currentPoints = 0;
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][0] == userId) { userRow = i + 1; currentPoints = Number(uData[i][3] || 0); break; }
  }

  const rData = rewardSheet.getDataRange().getValues();
  let reqPoints = 0;
  let rName = "";
  for (let i = 1; i < rData.length; i++) {
    if (rData[i][0] == rewardId) { reqPoints = Number(rData[i][2]); rName = rData[i][1]; break; }
  }

  if (userRow === -1 || !rName) return { status: "error", message: "データが見つかりません" };
  if (currentPoints < reqPoints) return { status: "error", message: "ポイントが足りません" };

  const newPoints = currentPoints - reqPoints;
  userSheet.getRange(userRow, 4).setValue(newPoints);
  invSheet.appendRow([userId, rewardId, rName, new Date(), "未使用"]);

  return { status: "success", message: `「${rName}」と交換しました！`, newPoints: newPoints };
}

function getInventory(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("もちもの");
  if (!invSheet) return { status: "success", inventory: [] };

  const data = invSheet.getDataRange().getValues();
  const inventory = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userId) {
      inventory.push({
        rowIdx: i + 1,
        rewardId: data[i][1],
        rewardName: data[i][2],
        date: data[i][3],
        status: data[i][4]
      });
    }
  }
  // 最新のものが上に来るように逆順にする
  return { status: "success", inventory: inventory.reverse() };
}

function consumeReward(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("もちもの");
  if (!invSheet) return { status: "error", message: "もちものシートがありません" };
  
  invSheet.getRange(rowIdx, 5).setValue("使用済み");
  return { status: "success", message: "景品を使いました！" };
}

// ==========================================
// ⑤ 外部学習の記録関連
// ==========================================
function getExternalLearning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("外部学習");
  if (!sheet) return { status: "success", list: [] };

  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      list.push({ name: data[i][0], points: Number(data[i][1]) });
    }
  }
  return { status: "success", list: list };
}

function reportExternalLearning(userId, menuName, points) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("ユーザー");
  let logSheet = ss.getSheetByName("学習履歴"); // 学習履歴にまとめる場合

  if (!logSheet) {
    logSheet = ss.insertSheet("学習履歴");
    logSheet.appendRow(["日時", "userId", "学習内容", "獲得ポイント"]);
  }

  const uData = userSheet.getDataRange().getValues();
  let userRow = -1;
  let currentPoints = 0;
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][0] == userId) {
      userRow = i + 1;
      currentPoints = Number(uData[i][3] || 0);
      break;
    }
  }

  if (userRow === -1) return { status: "error", message: "ユーザーが見つかりません" };

  const newTotal = currentPoints + points;
  userSheet.getRange(userRow, 4).setValue(newTotal);

  logSheet.appendRow([new Date(), userId, `[外部学習] ${menuName}`, points]);

  return { status: "success", message: `${points} Ptゲットしました！`, newTotal: newTotal };
}