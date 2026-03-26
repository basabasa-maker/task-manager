/**
 * Google Apps Script - タスク管理アプリ同期用API
 * GitHub Pages の PWA からタスクデータを受け取り、Google Sheets に保存する
 *
 * セットアップ:
 *   1. setupSheet() を実行してシートを初期化
 *   2. ウェブアプリとしてデプロイ（全員がアクセス可能に設定）
 */
 
const SHEET_NAME = "Tasks";
const HEADERS = ["id", "title", "quadrant", "due", "category", "note", "done", "created", "completedDate", "progress"];
 
/**
 * GET リクエスト: タスク一覧を JSON で返す
 */
function doGet(e) {
  try {
    const tasks = getTasks();
    const callback = e.parameter.callback;
    const result = JSON.stringify({
      success: true,
      data: tasks,
      count: tasks.length,
      timestamp: new Date().toISOString()
    });
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + result + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const callback = (e.parameter || {}).callback;
    const result = JSON.stringify({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    });
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + result + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  }
}
 
/**
 * POST リクエスト: タスクデータを受け取りシートに書き込む
 */
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      throw new Error("リクエストボディが空です");
    }
 
    const payload = JSON.parse(e.postData.contents);
    let tasks = Array.isArray(payload) ? payload : (payload.data || []);
 
    // バリデーションとサニタイズ
    tasks = tasks.map(task => sanitizeTask(task));
 
    // シートに書き込み
    writeTasks(tasks);
 
    // 最終同期タイムスタンプを更新
    updateLastSyncTimestamp();
 
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: tasks.length + " 件のタスクを同期しました",
        count: tasks.length,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
 
/**
 * シートを初期化する（最初に1回実行）
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
 
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
 
  // ヘッダー行を書き込み
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4285F4");
  headerRange.setFontColor("white");
 
  // 最終同期列
  sheet.getRange(1, HEADERS.length + 2).setValue("lastSync");
  sheet.getRange(1, HEADERS.length + 2).setFontWeight("bold");
 
  // ヘッダー行を固定
  sheet.setFrozenRows(1);
 
  // 列幅を自動調整
  for (let i = 1; i <= HEADERS.length; i++) {
    try { sheet.autoResizeColumn(i); } catch(e) {}
  }
 
  Logger.log("シートを初期化しました: " + SHEET_NAME);
}
 
/**
 * シートからタスクを読み取る
 */
function getTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
 
  if (!sheet) return [];
 
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
 
  const data = sheet.getRange(1, 1, lastRow, HEADERS.length).getValues();
  const headers = data[0];
  const tasks = [];
 
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // ID列が空なら空行としてスキップ
    if (!row[0] && row[0] !== 0) continue;
 
    const task = {};
    for (let j = 0; j < headers.length; j++) {
      const key = headers[j];
      let value = row[j];
 
      if (key === "done") {
        task[key] = (value === true || value === "TRUE" || value === "true");
      } else if (key === "progress" || key === "quadrant") {
        task[key] = isNaN(Number(value)) ? 0 : Number(value);
      } else if (key === "id") {
        task[key] = value;
      } else {
        task[key] = (value === null || value === undefined) ? "" : String(value);
      }
    }
    tasks.push(task);
  }
 
  return tasks;
}
 
/**
 * タスクをシートに書き込む（全行上書き）
 */
function writeTasks(tasks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
 
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
 
  // ヘッダー以外をクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  }
 
  if (tasks.length === 0) return;
 
  // データ行を作成
  const rows = tasks.map(task =>
    HEADERS.map(h => {
      const v = task[h];
      if (h === "done") return v === true || v === "true" ? "TRUE" : "FALSE";
      if (h === "progress") return isNaN(Number(v)) ? 0 : Number(v);
      if (h === "quadrant") return isNaN(Number(v)) ? 1 : Number(v);
      return (v === null || v === undefined) ? "" : String(v);
    })
  );
 
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
}
 
/**
 * 最終同期タイムスタンプを更新
 */
function updateLastSyncTimestamp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;
 
  const col = HEADERS.length + 2; // lastSync列
  sheet.getRange(2, col).setValue(new Date().toISOString());
}
 
/**
 * タスクデータをサニタイズ
 */
function sanitizeTask(task) {
  const sanitized = {};
 
  HEADERS.forEach(h => { sanitized[h] = ""; });
 
  if (typeof task === "object" && task !== null) {
    Object.keys(task).forEach(key => {
      if (!HEADERS.includes(key)) return;
      let v = task[key];
      if (v === null || v === undefined) v = "";
 
      if (key === "done") {
        sanitized[key] = (v === true || v === "true" || v === "TRUE");
      } else if (key === "progress") {
        sanitized[key] = Math.max(0, Math.min(100, isNaN(Number(v)) ? 0 : Number(v)));
      } else if (key === "quadrant") {
        sanitized[key] = isNaN(Number(v)) ? 1 : Number(v);
      } else {
        sanitized[key] = String(v).trim();
      }
    });
  }
 
  if (!sanitized.id) {
    throw new Error("タスクにIDがありません");
  }
 
  return sanitized;
}
 
Task app cloud sync and progress tracking - Claude
