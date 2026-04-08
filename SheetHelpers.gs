// ============================================================
// 共有ヘルパー — 複数モジュールから使用
// ============================================================

function getExistingCwTaskIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  if (!sheet || sheet.getLastRow() <= 1) return new Set();
  const ids = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).getValues();
  return new Set(ids.map(r => String(r[0])));
}

function getUniqueMemberNames() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MEMBERS);
  if (!s || s.getLastRow() <= 1) return [];
  return [...new Set(s.getRange(2,2,s.getLastRow()-1,1).getValues().map(r=>String(r[0])).filter(n=>n))].sort();
}

function getUniqueProjectNames() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROJECTS);
  if (!s || s.getLastRow() <= 1) return [];
  return [...new Set(s.getRange(2,2,s.getLastRow()-1,1).getValues().map(r=>String(r[0])).filter(n=>n))].sort();
}

function getProjectMap() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROJECTS);
  if (!s || s.getLastRow() <= 1) return {};
  const map = {};
  s.getRange(2,1,s.getLastRow()-1,2).getValues().forEach(r => { map[String(r[0])] = String(r[1]); });
  return map;
}

function getRoomName(roomId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROOMS);
  if (!sheet || sheet.getLastRow() <= 1) return roomId;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(roomId)) return String(data[i][1]);
  }
  return roomId;
}

function addLog(taskId, oldS, newS, who) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!sheet) return;
  sheet.appendRow([new Date(), taskId, oldS, newS, who]);

  // 古いログを自動削除
  const total = sheet.getLastRow() - 1;
  if (total > LOG_MAX_ROWS) {
    sheet.deleteRows(2, total - LOG_MAX_ROWS);
  }
}
