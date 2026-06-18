// ============================================================
// AIミニアプリ構築講座 アクセスログ記録スクリプト v2
// 受講者リスト照合機能付き
// ============================================================

const SHEET_NAME = "アクセスログ";
const MEMBER_SHEET = "受講者リスト";
const SPREADSHEET_ID = "1k-s06---LcTz7SDFIuSdalCo3nYwyLn4Lxqd84puPa0";

// ── 受講者リスト（メルアド → 氏名の対応表）──
const MEMBERS = {
  "m-sasano@fuyou-nagasaki.co.jp":         "佐々野　真",
  "tanaka@smifulhome.jp":                   "田中　由里子",
  "ma-bull109mikko@docomo.ne.jp":           "山下　美都子",
  "ybs-tt0123@outlook.jp":                  "津浪　勉",
  "mken_ta@yahoo.co.jp":                    "松永　あゆみ",
  "shimomura@n-kensou.com":                 "下村　慈",
  "komatu_setubi_yuuko@xqg.biglobe.ne.jp":  "野上　裕子",
  "ogatasw.with.office@gmail.com":          "小方　優子",
  "naosiro2@gmail.com":                     "前田　直",
  "hamasaki.ryunosuke@think-nagasaki.or.jp":"濱崎",
};

// ── POSTリクエストを受け取る（ログ記録）──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    writeLog(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GETリクエスト（テスト用）──
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "ログ記録サーバー稼働中" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── スプレッドシートに書き込む ──
function writeLog(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupHeaders(sheet);
  }

  // メルアドから氏名を取得
  const detail = data.detail || "";
  let email = "";
  let name = "（受講者外）";

  // 詳細列からメルアドを抽出（「eラーニング｜xxx@xxx.com」形式）
  const parts = detail.split("｜");
  if (parts.length >= 2) {
    email = parts[parts.length - 1].trim().toLowerCase();
    name = MEMBERS[email] || "（リスト外）";
  }

  const row = [
    data.datetime || new Date().toLocaleString("ja-JP", { timeZone: "Asia/Tokyo" }),
    data.type || "不明",
    parts[0] || detail,   // 種別詳細（eラーニング / セミナー等）
    email,                 // メールアドレス
    name,                  // 氏名（受講者リストから照合）
    data.ua ? data.ua.substring(0, 60) : ""
  ];

  sheet.appendRow(row);

  // ログイン失敗が続いたらメール通知
  checkSuspiciousAccess(sheet);
}

// ── ヘッダー設定 ──
function setupHeaders(sheet) {
  const headers = ["日時（JST）", "種別", "サイト", "メールアドレス", "氏名", "ブラウザ情報"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1D9E75");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 200);
}

// ── 不審アクセス検知 ──
function checkSuspiciousAccess(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return;

  const startRow = Math.max(2, lastRow - 9);
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 2).getValues();
  const failCount = rows.filter(r => r[1] === "ログイン失敗").length;

  if (failCount >= 5) {
    const email = Session.getActiveUser().getEmail();
    GmailApp.sendEmail(
      email,
      "⚠️ 【AIミニアプリ講座】不審なアクセスを検知",
      `直近${rows.length}件中${failCount}件のログイン失敗が検出されました。\n\n` +
      `スプレッドシートを確認してください：\n` +
      `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit\n\n` +
      `※このメールは自動送信です。`
    );
  }
}

// ── 初回セットアップ ──
function setupSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // アクセスログシート
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  setupHeaders(sheet);

  // 受講者リストシート
  let mSheet = ss.getSheetByName(MEMBER_SHEET);
  if (!mSheet) mSheet = ss.insertSheet(MEMBER_SHEET);
  const mHeaders = ["メールアドレス", "氏名", "ログイン回数", "最終ログイン"];
  mSheet.getRange(1, 1, 1, mHeaders.length).setValues([mHeaders]);
  const mHeaderRange = mSheet.getRange(1, 1, 1, mHeaders.length);
  mHeaderRange.setBackground("#1A2E23");
  mHeaderRange.setFontColor("#FFFFFF");
  mHeaderRange.setFontWeight("bold");
  mSheet.setFrozenRows(1);
  mSheet.setColumnWidth(1, 240);
  mSheet.setColumnWidth(2, 120);
  mSheet.setColumnWidth(3, 100);
  mSheet.setColumnWidth(4, 160);

  // 受講者データを入力
  const members = Object.entries(MEMBERS).map(([email, name]) => [email, name, 0, ""]);
  mSheet.getRange(2, 1, members.length, 4).setValues(members);

  SpreadsheetApp.flush();
  Logger.log("セットアップ完了");
}
