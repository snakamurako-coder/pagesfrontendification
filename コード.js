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
    phraseModeSs.getSheets()[0].setName("単元C").appendRow(["通し番号", "日本語", "英文", "別解１", "別解２", "別解３", "別解４", "別解５", "疑問文", "穴埋め１", "穴埋め２", "イニシャル", "イニシャルと文字数", "ヒント"]);
    phraseModeSs.getSheets()[0].appendRow([1, "私は8歳です。", "I'm eight years old.", "I am eight years old.", "", "", "", "", "How old are you?", "I'm (     ) years old.", "", "I", "I'm _ _ _ _ _", "年齢を聞かれた時の答え方だよ。"]);
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

  let extLearningSheet = adminSs.getSheetByName("外部学習");
  if (!extLearningSheet) {
    extLearningSheet = adminSs.insertSheet("外部学習");
    extLearningSheet.appendRow(["メニュー名", "獲得ポイント"]);
    extLearningSheet.appendRow(["学校の宿題（全教科）", 50]);
    extLearningSheet.appendRow(["読書（30分）", 30]);
    logMessage += "✅ 「外部学習」を追加しました。\n";
  }

  // ★ 新しい「特訓メニュー」の構造（学習ルート）
  let trainingSheet = adminSs.getSheetByName("特訓メニュー");
  if (!trainingSheet) {
    trainingSheet = adminSs.insertSheet("特訓メニュー");
    trainingSheet.appendRow(["対象ユーザー", "単元", "問題の形式", "こたえ方", "出し方"]);
    trainingSheet.appendRow(["全員", "単元A", "英単語→日本語", "4択", "順番通り"]);
    trainingSheet.appendRow(["全員", "単元A", "日本語→英単語", "タイピング", "ランダム"]);
    trainingSheet.appendRow(["全員", "単元C", "日本語→英文", "穴埋め", "順番通り"]);
    logMessage += "✅ 新しい「特訓メニュー（学習ルート）」を追加しました。\n";
  }

  console.log(logMessage);
  return logMessage;
}

const sendResponse = (responseObject) => {
  return ContentService.createTextOutput(JSON.stringify(responseObject)).setMimeType(ContentService.MimeType.JSON);
};

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
    else if (action === "report_external_learning") return handleReportExternalLearning(requestData);
    
    // ★ 特訓ルート用のAPI
    else if (action === "get_training_route") return handleGetTrainingRoute(requestData);
    
    else return sendResponse({ status: "error", message: "無効なactionです" });
  } catch (error) {
    return sendResponse({ status: "error", message: error.toString() });
  }
}

function doOptions(e) { return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT); }

// ==========================================
// ★ 新しい特訓ルート機能
// ==========================================
function handleGetTrainingRoute(req) {
  const props = PropertiesService.getScriptProperties();
  const adminSs = SpreadsheetApp.openById(props.getProperty('ADMIN_SS_ID'));
  
  // 1. ルート（メニュー）の取得
  const sheet = adminSs.getSheetByName("特訓メニュー");
  const route = [];
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    // 行を上から順番に読み込み、対象ユーザーならリストに追加
    for (let i = 1; i < data.length; i++) {
      let target = String(data[i][0]);
      if (target === "全員" || target.includes(req.userId)) {
        route.push({
          stepIndex: i, // 行番号を固有IDとして使用
          unitName: data[i][1],
          qFormat: data[i][2],
          aFormat: data[i][3],
          mode: data[i][4]
        });
      }
    }
  }

  // 2. 今日の進捗状況の取得
  const usersSheet = adminSs.getSheetByName("users");
  const userData = usersSheet.getDataRange().getValues();
  let progressData = {};
  const todayStr = new Date().toISOString().split('T')[0]; // "2023-10-27" のような形式

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === req.userId) {
      // 8列目(インデックス7)が進捗JSON
      const rawProgress = JSON.parse(userData[i][7] || "{}");
      // 今日の日付のデータだけを返す（日付が変わったら空っぽとして扱う）
      if (rawProgress[todayStr]) {
        progressData = rawProgress[todayStr];
      }
      break;
    }
  }

  return sendResponse({ status: "success", route: route, progress: progressData });
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

  const earnedPoints = Math.round((sessionRawPoints * multiplier) * 100) / 100;
  const newTotalPoints = Math.round((userData.points + earnedPoints) * 100) / 100;
  
  userData.dailyPointsJson[todayStr] = (userData.dailyPointsJson[todayStr] || 0) + earnedPoints;
  userData.dailyPointsJson[todayStr] = Math.round(userData.dailyPointsJson[todayStr] * 100) / 100;
  userData.lastStudyJson[req.unitId] = now.toISOString();

  // ★ 特訓ルートのステップをクリアした場合は、今日の進捗にチェックを入れる
  if (req.trainingStepIndex) {
    if (!userData.trainingProgressJson[todayStr]) {
      userData.trainingProgressJson = {}; // 過去のゴミを消して今日からスタート
      userData.trainingProgressJson[todayStr] = {};
    }
    userData.trainingProgressJson[todayStr][req.trainingStepIndex] = true;
  }

  usersSheet.getRange(targetRowIdx, 4).setValue(newTotalPoints);
  usersSheet.getRange(targetRowIdx, 5).setValue(JSON.stringify(userData.lastStudyJson));
  usersSheet.getRange(targetRowIdx, 6).setValue(JSON.stringify(userData.historyJson));
  usersSheet.getRange(targetRowIdx, 7).setValue(JSON.stringify(userData.dailyPointsJson));
  usersSheet.getRange(targetRowIdx, 8).setValue(JSON.stringify(userData.trainingProgressJson)); // ★ 保存

  return sendResponse({ status: "success", earnedPoints: earnedPoints, newTotal: newTotalPoints, historyJson: userData.historyJson, dailyPointsJson: userData.dailyPointsJson, bonusApplied: req.isRandom });
}

// ==========================================
// 既存のハンドラー（変更なし）
// ==========================================
function handleGetAppSettings(req) { const settingsSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("アプリ設定"); if (!settingsSheet) return sendResponse({ status: "success", settings: {} }); const data = settingsSheet.getDataRange().getValues(); const settings = {}; for (let i = 1; i < data.length; i++) { if (data[i][0]) settings[data[i][0]] = data[i][1]; } return sendResponse({ status: "success", settings: settings }); }
function handleGetExternalLearning(req) { const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("外部学習"); const list = []; if (sheet) { const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0]) list.push({ name: data[i][0], points: Number(data[i][1]) }); } } return sendResponse({ status: "success", list: list }); }
function handleReportExternalLearning(req) { const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_SS_ID')).getSheetByName("users"); const data = sheet.getDataRange().getValues(); let newTotal = 0; for (let i = 1; i < data.length; i++) { if (data[i][0] === req.userId) { newTotal = Number(data[i][3]) + Number(req.points); sheet.getRange(i + 1, 4).setValue(newTotal); break; } } return sendResponse({ status: "success", newTotal: newTotal, message: req.menuName + " のポイントをゲットしたよ！" }); }
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