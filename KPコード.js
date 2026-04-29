/**
 * 厳密漢字判定ツール (KanjiVG) - サーバー側スクリプト
 * * 【🚀 必須の事前準備】
 * 1. 読み込みたいスプレッドシートのIDをコピー
 * 2. GASエディタの「プロジェクトの設定（歯車）」>「スクリプト プロパティ」に以下を追加
 * - プロパティ: KANJI_SHEET_ID
 * - 値: コピーしたスプレッドシートのID
 */

/**
 * ウェブアプリとしてのエントリポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('漢字書き順・美文字ドリル PRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * アプリ起動時の初期設定情報を取得
 */
function getAppInitData() {
  const prop = PropertiesService.getScriptProperties();
  const sheetId = prop.getProperty('KANJI_SHEET_ID');
  
  if (!sheetId) {
    throw new Error("❌ スクリプトプロパティ 'KANJI_SHEET_ID' が未設定です。");
  }
  
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets().map(s => s.getName());
    
    // スプレッドシート名とシート一覧を返す
    return {
      bookName: ss.getName(),
      sheets: sheets
    };
  } catch (e) {
    throw new Error("❌ スプレッドシートにアクセスできません。IDと共有権限を確認してください。\n" + e.message);
  }
}

/**
 * 特定のシートから漢字データ（SVGパス）を抽出して整形する
 * @param {string} sheetName 読み込むシート名
 */
function getKanjiDataFromSheet(sheetName) {
  const prop = PropertiesService.getScriptProperties();
  const sheetId = prop.getProperty('KANJI_SHEET_ID');
  
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {};

    // 全データを一括取得して処理を高速化
    const values = sheet.getDataRange().getValues();
    const kanjiMap = {};

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const kanji = row[0]; // A列：漢字

      // A列が「1文字」の場合のみ処理（ヘッダーや空行を無視）
      if (kanji && String(kanji).trim().length === 1) {
        const paths = [];

        // B列(index 1)はUnicode表記（4.00E+00などバグりやすいため）を無視
        // C列(index 2)以降のパスデータを走査
        for (let j = 2; j < row.length; j++) {
          const cellValue = row[j];
          if (!cellValue) continue;

          const strVal = String(cellValue).trim();
          if (strVal === "") continue;

          // 「|」区切りで複数画が入っている場合の分割処理
          if (strVal.includes('|')) {
            const parts = strVal.split('|');
            parts.forEach(p => {
              const cleaned = p.trim();
              if (cleaned.startsWith('M') || cleaned.startsWith('m')) {
                paths.push(cleaned);
              }
            });
          } else {
            // 1セル1画、かつMから始まるSVGパスであれば採用
            if (strVal.startsWith('M') || strVal.startsWith('m')) {
              paths.push(strVal);
            }
          }
        }

        // 有効なパスが見つかれば登録
        if (paths.length > 0) {
          kanjiMap[kanji] = paths;
        }
      }
    }
    
    return kanjiMap;

  } catch (e) {
    console.error("Data Fetch Error: " + e.message);
    return {};
  }
}