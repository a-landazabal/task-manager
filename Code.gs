// ============================================================
// 設定
// ============================================================
const CHATWORK_API_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
const CHATWORK_API_BASE  = 'https://api.chatwork.com/v2';

const SHEET_TASKS     = 'タスク';
const SHEET_MEMBERS   = 'メンバー';
const SHEET_ROOMS     = 'ルーム';
const SHEET_LOG       = 'ログ';
const SHEET_DASHBOARD = 'ダッシュボード';
const SHEET_PROJECTS  = 'プロジェクト';

const STATUS_PENDING  = '未完了';
const STATUS_FINISHED = '完了';

// ============================================================
// セットアップ — 一度だけ実行
// ============================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets();

  // ルームとメンバーが既にある場合はスキップ
  if (!ss.getSheetByName(SHEET_ROOMS) || ss.getSheetByName(SHEET_ROOMS).getLastRow() <= 1) syncRooms();
  if (!ss.getSheetByName(SHEET_MEMBERS) || ss.getSheetByName(SHEET_MEMBERS).getLastRow() <= 1) syncMembers();

  setupDashboard();
  createTrigger();
  createEditTrigger();

  Logger.log('セットアップ完了！ルームシートで監視したいルームの D列 を「YES」に変更してください。');
}

// ============================================================
// シート初期化
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── タスク（12列）──
  let s = ss.getSheetByName(SHEET_TASKS);
  if (!s) s = ss.insertSheet(SHEET_TASKS);
  if (s.getLastRow() === 0) {
    s.appendRow(['タスクID','作成日時','更新日時','内容','担当者ID','担当者名','作成者ID','作成者名','ルームID','CWタスクID','ステータス','プロジェクト']);
    s.getRange(1,1,1,12).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }

  // ── メンバー（4列）──
  s = ss.getSheetByName(SHEET_MEMBERS);
  if (!s) s = ss.insertSheet(SHEET_MEMBERS);
  if (s.getLastRow() === 0) {
    s.appendRow(['アカウントID','名前','ルームID','ルーム名']);
    s.getRange(1,1,1,4).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }

  // ── ルーム（4列：ID, 名前, 最終確認, 監視フラグ）──
  s = ss.getSheetByName(SHEET_ROOMS);
  if (!s) s = ss.insertSheet(SHEET_ROOMS);
  if (s.getLastRow() === 0) {
    s.appendRow(['ルームID','ルーム名','最終確認日時','タスク監視']);
    s.getRange(1,1,1,4).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }

  // ── プロジェクト（2列）──
  s = ss.getSheetByName(SHEET_PROJECTS);
  if (!s) s = ss.insertSheet(SHEET_PROJECTS);
  if (s.getLastRow() === 0) {
    s.appendRow(['アカウントID','プロジェクト名']);
    s.getRange(1,1,1,2).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }

  // ── ログ（5列）──
  s = ss.getSheetByName(SHEET_LOG);
  if (!s) s = ss.insertSheet(SHEET_LOG);
  if (s.getLastRow() === 0) {
    s.appendRow(['日時','タスクID','変更前','変更後','変更者']);
    s.getRange(1,1,1,5).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }

  Logger.log('シート初期化完了。');
}

function createTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    const name = t.getHandlerFunction();
    if (name === 'pollChatWork' || name === 'dailySync') ScriptApp.deleteTrigger(t);
  });

  // 1分ごとのタスクポーリング
  ScriptApp.newTrigger('pollChatWork').timeBased().everyMinutes(1).create();

  // 毎日午前6時にルーム・メンバー自動同期
  ScriptApp.newTrigger('dailySync').timeBased().atHour(6).everyDays(1).create();

  Logger.log('トリガー設定完了（ポーリング + 毎日同期）。');
}

/**
 * 毎日自動実行 — 新しいルームとメンバーを同期
 */
function dailySync() {
  syncRooms();
  Utilities.sleep(3000); // API制限回避
  syncMembers();
  Logger.log('日次同期完了。');
}

function createEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  Logger.log('編集トリガー設定完了。');
}

// ============================================================
// メインポーリング — 監視フラグ=YESのルームだけをポーリング
// ============================================================
function pollChatWork() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const roomsSheet = ss.getSheetByName(SHEET_ROOMS);
  if (!roomsSheet || roomsSheet.getLastRow() <= 1) return;

  const rows = roomsSheet.getRange(2, 1, roomsSheet.getLastRow() - 1, 4).getValues();

  for (let i = 0; i < rows.length; i++) {
    const roomId      = String(rows[i][0]).trim();
    const roomName    = String(rows[i][1]).trim();
    const lastChecked = rows[i][2] ? Number(rows[i][2]) : 0;
    const monitor     = String(rows[i][3]).trim().toUpperCase();

    if (!roomId || monitor !== 'YES') continue;

    try {
      // ── Chatworkタスク同期 ──
      syncChatworkTasks(roomId, roomName);

      // ── メッセージ（【プロジェクト追加】コマンドのみ） ──
      if (lastChecked === 0) {
        roomsSheet.getRange(i + 2, 3).setValue(Math.floor(Date.now() / 1000));
        continue;
      }

      const messages = fetchMessages(roomId, lastChecked);
      if (messages.length > 0) {
        processCommands(messages);
        roomsSheet.getRange(i + 2, 3).setValue(Math.max(...messages.map(m => m.send_time)));
      }
    } catch (e) {
      Logger.log(`ルーム ${roomId} エラー: ${e.message}`);
    }
  }
}

// ============================================================
// ChatWork API
// ============================================================
function fetchMessages(roomId, since) {
  const res = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/messages?force=0`, {
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return [];
  const msgs = JSON.parse(res.getContentText());
  return since > 0 ? msgs.filter(m => m.send_time > since) : msgs;
}

function fetchChatworkTasks(roomId, status) {
  const res = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/tasks?status=${status}`, {
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code === 429) {
    Utilities.sleep(5000);
    const retry = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/tasks?status=${status}`, {
      headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
    });
    return retry.getResponseCode() === 200 ? JSON.parse(retry.getContentText()) : [];
  }
  return code === 200 ? JSON.parse(res.getContentText()) : [];
}

// ============================================================
// Chatworkタスク同期 — プロジェクト検出 + 重複防止 + 完了同期
// ============================================================
function syncChatworkTasks(roomId, roomName) {
  const projectMap = getProjectMap();

  // ── 未完了タスクを取得 ──
  const openTasks = fetchChatworkTasks(roomId, 'open');

  // パス1: 全エントリからプロジェクト名を検出
  // 同じタスク内容+作成者のグループで、プロジェクトアカウントがいればそれがプロジェクト名
  const projectByGroup = {};
  openTasks.forEach(t => {
    const aid = String(t.account.account_id);
    const cid = String(t.assigned_by_account.account_id);
    const key = cid + '|' + (t.body || '');
    if (projectMap[aid]) projectByGroup[key] = projectMap[aid];
    if (projectMap[cid]) projectByGroup[key] = projectMap[cid];
  });

  // パス2: 人のタスクだけ追加
  openTasks.forEach(t => {
    const cwTaskId   = String(t.task_id);
    const aid        = String(t.account.account_id);
    const cid        = String(t.assigned_by_account.account_id);
    const key        = cid + '|' + (t.body || '');

    // プロジェクトアカウントはスキップ（ラベルのみ）
    if (projectMap[aid]) return;

    // 既にシートにあればスキップ
    if (taskExistsByCwId(cwTaskId)) return;

    createTask({
      description   : t.body || '',
      assignedToId  : aid,
      assignedToName: t.account.name,
      createdById   : cid,
      createdByName : t.assigned_by_account.name,
      roomId        : roomId,
      cwTaskId      : cwTaskId,
      project       : projectByGroup[key] || ''
    });
  });

  // ── 完了タスクを同期 ──
  const doneTasks = fetchChatworkTasks(roomId, 'done');
  doneTasks.forEach(t => {
    if (projectMap[String(t.account.account_id)]) return;
    markTaskDoneByCwId(String(t.task_id));
  });
}

// ============================================================
// タスク管理
// ============================================================
function createTask(d) {
  if (taskExistsByCwId(d.cwTaskId)) return;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TASKS);
  const now   = new Date();

  sheet.appendRow([
    'TASK-' + now.getTime(), now, now,
    d.description, d.assignedToId, d.assignedToName,
    d.createdById, d.createdByName, d.roomId,
    d.cwTaskId, STATUS_PENDING, d.project || ''
  ]);

  const row = sheet.getLastRow();
  sheet.getRange(row, 11).setBackground('#FFF9C4');

  addLog('TASK-' + now.getTime(), '', STATUS_PENDING, d.createdByName);
  Logger.log(`タスク作成: ${d.assignedToName} — ${d.description.substring(0,30)}`);
  refreshDashboard();
}

function taskExistsByCwId(cwTaskId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const ids = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).getValues();
  return ids.some(r => String(r[0]) === cwTaskId);
}

function markTaskDoneByCwId(cwTaskId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][9]) === cwTaskId && String(data[i][10]) === STATUS_PENDING) {
      const row = i + 2;
      sheet.getRange(row, 3).setValue(new Date());
      sheet.getRange(row, 11).setValue(STATUS_FINISHED);
      sheet.getRange(row, 11).setBackground('#C8E6C9');
      addLog(data[i][0], STATUS_PENDING, STATUS_FINISHED, 'Chatwork');
      Logger.log(`タスク完了: ${data[i][0]}`);
      refreshDashboard();
    }
  }
}

// ============================================================
// コマンド処理（【プロジェクト追加】）
// ============================================================
function processCommands(messages) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROJECTS);
  messages.forEach(m => {
    const match = (m.body || '').match(/【プロジェクト追加】\s*(.+?)\s+(\d+)/);
    if (!match) return;
    const name = match[1].trim();
    const id   = match[2].trim();
    const existing = getProjectMap();
    if (existing[id]) return;
    sheet.appendRow([id, name]);
    Logger.log(`プロジェクト追加: ${name} (${id})`);
  });
}

// ============================================================
// ログ
// ============================================================
function addLog(taskId, oldS, newS, who) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (sheet) sheet.appendRow([new Date(), taskId, oldS, newS, who]);
}

// ============================================================
// ルーム自動取得
// ============================================================
function syncRooms() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ROOMS);
  const existing = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().forEach(r => existing.add(String(r[0])));
  }

  try {
    const res = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms`, {
      headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return;
    const rooms = JSON.parse(res.getContentText());
    const newRows = [];
    rooms.forEach(r => {
      const id = String(r.room_id);
      if (!existing.has(id)) { newRows.push([id, r.name, '', '']); existing.add(id); }
    });
    if (newRows.length > 0) sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, 4).setValues(newRows);
    Logger.log(`ルーム同期: ${newRows.length}件追加（合計${existing.size}件）`);
  } catch(e) { Logger.log(`ルーム取得エラー: ${e.message}`); }
}

// ============================================================
// メンバー自動取得
// ============================================================
function syncMembers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName(SHEET_ROOMS);
  const mSheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!rSheet || rSheet.getLastRow() <= 1) return;

  const existing = new Set();
  if (mSheet.getLastRow() > 1) {
    mSheet.getRange(2,1,mSheet.getLastRow()-1,4).getValues()
      .forEach(r => existing.add(String(r[0])+'_'+String(r[2])));
  }

  const rooms = rSheet.getRange(2,1,rSheet.getLastRow()-1,2).getValues();
  const newRows = [];

  for (let i = 0; i < rooms.length; i++) {
    const roomId = String(rooms[i][0]).trim();
    const roomName = String(rooms[i][1]).trim();
    if (!roomId) continue;

    try {
      if (i > 0) Utilities.sleep(1000);
      const res = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/members`, {
        headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code === 429) { Utilities.sleep(10000); continue; }
      if (code !== 200) continue;
      JSON.parse(res.getContentText()).forEach(m => {
        const key = String(m.account_id)+'_'+roomId;
        if (!existing.has(key)) { newRows.push([String(m.account_id), m.name, roomId, roomName]); existing.add(key); }
      });
    } catch(e) { Logger.log(`メンバー取得エラー [${roomId}]: ${e.message}`); }
  }

  if (newRows.length > 0) mSheet.getRange(mSheet.getLastRow()+1, 1, newRows.length, 4).setValues(newRows);
  if (mSheet.getLastRow() > 2) mSheet.getRange(2,1,mSheet.getLastRow()-1,4).sort({column:4,ascending:true});
  Logger.log(`メンバー同期: ${newRows.length}名追加`);
}

// ============================================================
// ダッシュボード
// ============================================================
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(SHEET_DASHBOARD);
  if (dash) ss.deleteSheet(dash);
  dash = ss.insertSheet(SHEET_DASHBOARD, 0);

  dash.getRange('A:H').setFontFamily('Arial').setFontSize(10);

  // タイトル
  dash.getRange('A1:H1').merge().setValue('タスク管理ダッシュボード')
    .setFontSize(16).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#2C3E50').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(1, 45);

  // フィルター
  dash.getRange('A2').setValue('担当者').setFontWeight('bold').setFontSize(11)
    .setBackground('#34495E').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  dash.getRange('B2:C2').merge().setValue('すべて').setFontSize(11)
    .setBackground('#FFFFFF').setFontColor('#2C3E50')
    .setBorder(true,true,true,true,false,false,'#BDC3C7',SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center');
  dash.getRange('D2').setValue('').setBackground('#34495E');
  dash.getRange('E2').setValue('プロジェクト').setFontWeight('bold').setFontSize(11)
    .setBackground('#34495E').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  dash.getRange('F2:H2').merge().setValue('すべて').setFontSize(11)
    .setBackground('#FFFFFF').setFontColor('#2C3E50')
    .setBorder(true,true,true,true,false,false,'#BDC3C7',SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center');
  dash.setRowHeight(2, 35);

  // ドロップダウン
  const mRule = SpreadsheetApp.newDataValidation().requireValueInList(['すべて'].concat(getUniqueMemberNames())).setAllowInvalid(false).build();
  dash.getRange('B2').setDataValidation(mRule);
  const pRule = SpreadsheetApp.newDataValidation().requireValueInList(['すべて'].concat(getUniqueProjectNames())).setAllowInvalid(false).build();
  dash.getRange('F2').setDataValidation(pRule);

  // サマリー
  dash.getRange('A3').setValue('合計').setFontWeight('bold').setBackground('#ECF0F1').setHorizontalAlignment('center');
  dash.getRange('B3').setValue('0').setFontWeight('bold').setFontSize(14).setBackground('#ECF0F1').setHorizontalAlignment('center');
  dash.getRange('C3').setValue('未完了').setFontWeight('bold').setBackground('#FFF9C4').setHorizontalAlignment('center');
  dash.getRange('D3').setValue('0').setFontWeight('bold').setFontSize(14).setBackground('#FFF9C4').setHorizontalAlignment('center');
  dash.getRange('E3').setValue('完了').setFontWeight('bold').setBackground('#C8E6C9').setHorizontalAlignment('center');
  dash.getRange('F3').setValue('0').setFontWeight('bold').setFontSize(14).setBackground('#C8E6C9').setHorizontalAlignment('center');
  dash.getRange('G3').setValue('完了率').setFontWeight('bold').setBackground('#D5F5E3').setHorizontalAlignment('center');
  dash.getRange('H3').setValue('0%').setFontWeight('bold').setFontSize(14).setBackground('#D5F5E3').setHorizontalAlignment('center');
  dash.setRowHeight(3, 32);

  // スペーサー
  dash.setRowHeight(4, 6);
  dash.getRange('A4:H4').setBackground('#2C3E50');

  // テーブルヘッダー
  dash.getRange(5,1,1,8).setValues([['No.','作成日時','内容','担当者','作成者','プロジェクト','ステータス','更新日時']])
    .setFontWeight('bold').setFontSize(10).setBackground('#2C3E50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  dash.setRowHeight(5, 28);
  dash.setFrozenRows(5);

  dash.setColumnWidth(1,50); dash.setColumnWidth(2,140); dash.setColumnWidth(3,280);
  dash.setColumnWidth(4,130); dash.setColumnWidth(5,130); dash.setColumnWidth(6,130);
  dash.setColumnWidth(7,90); dash.setColumnWidth(8,140);

  refreshDashboard();
  Logger.log('ダッシュボード作成完了。');
}

function refreshDashboard() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SHEET_DASHBOARD);
  const task = ss.getSheetByName(SHEET_TASKS);
  if (!dash) return;

  const startRow = 6, cols = 8;

  if (!task || task.getLastRow() <= 1) {
    if (dash.getLastRow() >= startRow) { const r=dash.getRange(startRow,1,dash.getLastRow()-startRow+1,cols); r.breakApart(); r.clearContent().clearFormat(); }
    dash.getRange('B3').setValue('0'); dash.getRange('D3').setValue('0');
    dash.getRange('F3').setValue('0'); dash.getRange('H3').setValue('0%');
    return;
  }

  const fPerson  = dash.getRange('B2').getValue();
  const fProject = dash.getRange('F2').getValue();
  const lastCol  = Math.max(task.getLastColumn(), 12);
  const data     = task.getRange(2, 1, task.getLastRow()-1, lastCol).getValues();

  const filtered = data.filter(r => {
    if (fPerson !== 'すべて' && String(r[5]) !== fPerson && String(r[7]) !== fPerson) return false;
    if (fProject !== 'すべて' && String(r[11]||'') !== fProject) return false;
    return true;
  });

  const total = filtered.length;
  const pending = filtered.filter(r => String(r[10]) === STATUS_PENDING).length;
  const done = filtered.filter(r => String(r[10]) === STATUS_FINISHED).length;
  const rate = total > 0 ? Math.round(done/total*100) : 0;

  dash.getRange('B3').setValue(total).setFontWeight('bold').setFontSize(14).setBackground('#ECF0F1').setHorizontalAlignment('center');
  dash.getRange('D3').setValue(pending).setFontWeight('bold').setFontSize(14).setBackground('#FFF9C4').setHorizontalAlignment('center');
  dash.getRange('F3').setValue(done).setFontWeight('bold').setFontSize(14).setBackground('#C8E6C9').setHorizontalAlignment('center');
  dash.getRange('H3').setValue(rate+'%').setFontWeight('bold').setFontSize(14).setBackground('#D5F5E3').setHorizontalAlignment('center');

  if (dash.getLastRow() >= startRow) { const r=dash.getRange(startRow,1,dash.getLastRow()-startRow+1,cols); r.breakApart(); r.clearContent().clearFormat(); }

  if (total === 0) { dash.getRange(startRow,1).setValue('該当するタスクはありません。').setFontColor('#95A5A6').setFontStyle('italic'); return; }

  const out = filtered.map((r,i) => [i+1, r[1], r[3], r[5], r[7], String(r[11]||''), r[10], r[2]]);
  dash.getRange(startRow, 1, out.length, cols).setValues(out);

  for (let i = 0; i < out.length; i++) {
    const row = startRow + i;
    const status = String(out[i][6]);
    dash.getRange(row,1,1,cols).setBackground(i%2===0?'#FFFFFF':'#F8F9FA')
      .setBorder(false,false,true,false,false,false,'#ECF0F1',SpreadsheetApp.BorderStyle.SOLID);
    const sc = dash.getRange(row, 7);
    if (status === STATUS_PENDING)  sc.setBackground('#FFF9C4').setFontColor('#F39C12').setFontWeight('bold');
    if (status === STATUS_FINISHED) sc.setBackground('#C8E6C9').setFontColor('#27AE60').setFontWeight('bold');
    dash.getRange(row, 1).setHorizontalAlignment('center');
  }
  Logger.log(`ダッシュボード更新: ${total}件`);
}

// ============================================================
// ヘルパー
// ============================================================
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

// ============================================================
// シート編集検知
// ============================================================
function onEdit(e) {
  const sheet = e.range.getSheet();
  const name  = sheet.getName();

  // ダッシュボードフィルター変更
  if (name === SHEET_DASHBOARD && e.range.getRow() === 2 && (e.range.getColumn() === 2 || e.range.getColumn() === 6)) {
    refreshDashboard(); return;
  }

  // タスクシートのステータス変更
  if (name !== SHEET_TASKS || e.range.getColumn() !== 11 || e.range.getRow() === 1) return;

  const val = String(e.value || '').trim();
  const row = e.range.getRow();

  if (val !== STATUS_PENDING && val !== STATUS_FINISHED) {
    e.range.setValue(e.oldValue || STATUS_PENDING);
    SpreadsheetApp.getUi().alert('無効なステータスです。\n使用できる値：未完了、完了');
    return;
  }

  sheet.getRange(row, 3).setValue(new Date());
  sheet.getRange(row, 11).setBackground(val === STATUS_FINISHED ? '#C8E6C9' : '#FFF9C4');
  addLog(sheet.getRange(row,1).getValue(), e.oldValue||'', val, Session.getActiveUser().getEmail());
  refreshDashboard();
}
