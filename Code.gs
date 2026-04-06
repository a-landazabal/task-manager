// ============================================================
// 設定 — setup() を実行する前に入力してください
// ============================================================
const CHATWORK_API_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
const CHATWORK_API_BASE  = 'https://api.chatwork.com/v2';

// シート名
const SHEET_TASKS   = 'タスク';
const SHEET_MEMBERS = 'メンバー';
const SHEET_ROOMS   = 'ルーム';
const SHEET_LOG     = 'ログ';

// ステータス
const STATUS_PENDING  = '未完了';
const STATUS_FINISHED = '完了';

// ============================================================
// セットアップ — スクリプトを貼り付けたら一度だけ実行してください
// ============================================================
function setup() {
  setupSheets();
  createTrigger();
  createEditTrigger();
  Logger.log('セットアップ完了！ルームシートにルームIDを追加すると自動的にポーリングが始まります。');
}

function createEditTrigger() {
  // 既存の onEdit トリガーを削除して重複を防ぐ
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEdit') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log('編集トリガー設定完了：ステータス変更を自動検知します。');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── タスク ─────────────────────────────────────────────────
  let tasks = ss.getSheetByName(SHEET_TASKS);
  if (!tasks) tasks = ss.insertSheet(SHEET_TASKS);
  if (tasks.getLastRow() === 0) {
    tasks.appendRow([
      'タスクID', '作成日時', '更新日時', '内容',
      '担当者ID（aid）', '担当者名',
      '作成者ID（aid）', '作成者名',
      'ルームID', 'メッセージID', 'ステータス'
    ]);
    tasks.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    tasks.setFrozenRows(1);
  }

  // ── メンバー ───────────────────────────────────────────────
  let members = ss.getSheetByName(SHEET_MEMBERS);
  if (!members) members = ss.insertSheet(SHEET_MEMBERS);
  if (members.getLastRow() === 0) {
    members.appendRow(['アカウントID（aid）', '名前']);
    members.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    members.setFrozenRows(1);
  }

  // ── ルーム ─────────────────────────────────────────────────
  let rooms = ss.getSheetByName(SHEET_ROOMS);
  if (!rooms) rooms = ss.insertSheet(SHEET_ROOMS);
  if (rooms.getLastRow() === 0) {
    rooms.appendRow(['ルームID', 'ルーム名', '最終確認日時（Unixタイムスタンプ）']);
    rooms.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    rooms.setFrozenRows(1);
  }

  // ── ログ ───────────────────────────────────────────────────
  let log = ss.getSheetByName(SHEET_LOG);
  if (!log) log = ss.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(['日時', 'タスクID', '変更前ステータス', '変更後ステータス', '変更者']);
    log.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    log.setFrozenRows(1);
  }

  Logger.log('全シートの初期化が完了しました。');
}

function createTrigger() {
  // 既存の pollChatWork トリガーを削除して重複を防ぐ
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'pollChatWork') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 1分ごとにポーリング
  ScriptApp.newTrigger('pollChatWork')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('トリガー設定完了：pollChatWork が1分ごとに実行されます。');
}

// ============================================================
// メインポーリング — 1分ごとに自動実行されます
// ============================================================
function pollChatWork() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const roomsSheet = ss.getSheetByName(SHEET_ROOMS);

  if (!roomsSheet || roomsSheet.getLastRow() <= 1) {
    Logger.log('ルームシートにルームが登録されていません。ルームIDとルーム名を追加してください。');
    return;
  }

  const rows = roomsSheet.getRange(2, 1, roomsSheet.getLastRow() - 1, 3).getValues();

  rows.forEach((row, i) => {
    const roomId      = String(row[0]).trim();
    const roomName    = String(row[1]).trim();
    const lastChecked = row[2] ? Number(row[2]) : 0;

    if (!roomId) return;

    try {
      const messages = fetchMessages(roomId, lastChecked);

      if (messages.length > 0) {
        processMessages(messages, roomId, roomName);

        // 次回のポーリングで新しいメッセージだけ取得するため最新の送信時刻を保存
        const latestTime = Math.max(...messages.map(m => m.send_time));
        roomsSheet.getRange(i + 2, 3).setValue(latestTime);
      }
    } catch (e) {
      Logger.log(`ルーム ${roomId} のポーリング中にエラーが発生しました: ${e.message}`);
    }
  });
}

// ============================================================
// ChatWork API
// ============================================================
function fetchMessages(roomId, since) {
  const url = `${CHATWORK_API_BASE}/rooms/${roomId}/messages?force=0`;

  const res = UrlFetchApp.fetch(url, {
    method : 'GET',
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log(`ChatWork APIエラー [ルーム ${roomId}]: ${res.getResponseCode()} — ${res.getContentText()}`);
    return [];
  }

  const messages = JSON.parse(res.getContentText());

  // 最終確認日時より新しいメッセージのみ返す
  return since > 0 ? messages.filter(m => m.send_time > since) : messages;
}

// ============================================================
// メッセージ処理
// ============================================================
function processMessages(messages, roomId, roomName) {
  messages.forEach(message => {
    const body = message.body || '';

    // 【タスク】キーワードがないメッセージは無視
    if (!body.includes('【タスク】')) return;

    // メッセージ内の [To:アカウントID] タグをすべて取得
    const toTags = body.match(/\[To:(\d+)\]/g);
    if (!toTags) return; // タグなし → スキップ

    const senderId   = String(message.account.account_id);
    const senderName = message.account.name;

    // タグごとにタスクを作成
    toTags.forEach(tag => {
      const assigneeId   = tag.match(/\[To:(\d+)\]/)[1];
      const assigneeName = lookupMemberName(assigneeId);

      // 説明文から [To:X] タグとその他のタグを除去
      const description = body
        .replace(/\[To:\d+\][^\[]*/g, '')
        .replace(/\[.*?\]/g, '')
        .trim() || body;

      createTask({
        description   : description,
        assignedToId  : assigneeId,
        assignedToName: assigneeName,
        createdById   : senderId,
        createdByName : senderName,
        roomId        : roomId,
        messageId     : String(message.message_id)
      });
    });
  });
}

// ============================================================
// タスク管理
// ============================================================
function createTask(data) {
  // 同じメッセージ × 同じ担当者の重複タスクを防ぐ
  if (taskExistsForMessage(data.messageId, data.assignedToId)) {
    Logger.log(`重複スキップ — メッセージ ${data.messageId}（担当者: ${data.assignedToId}）のタスクはすでに存在します。`);
    return;
  }

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(SHEET_TASKS);
  const taskId     = generateTaskId();
  const now        = new Date();

  tasksSheet.appendRow([
    taskId,
    now,
    now,
    data.description,
    data.assignedToId,
    data.assignedToName,
    data.createdById,
    data.createdByName,
    data.roomId,
    data.messageId,
    STATUS_PENDING
  ]);

  // ステータスセルに色を付ける
  const lastRow = tasksSheet.getLastRow();
  applyStatusColor(tasksSheet, lastRow, STATUS_PENDING);

  addLog(taskId, '', STATUS_PENDING, data.createdByName);
  Logger.log(`タスク作成: ${taskId} → 担当者: ${data.assignedToName}`);
}

/**
 * タスクのステータスを更新します。
 * newStatus は '未完了' または '完了' のいずれか。
 * changedBy は変更者の名前またはアカウントID。
 */
function updateTaskStatus(taskId, newStatus, changedBy) {
  const validStatuses = [STATUS_PENDING, STATUS_FINISHED];
  if (!validStatuses.includes(newStatus)) {
    Logger.log(`無効なステータス: ${newStatus}。使用できる値: 未完了、完了`);
    return false;
  }

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(SHEET_TASKS);

  if (tasksSheet.getLastRow() <= 1) return false;

  const data = tasksSheet.getRange(2, 1, tasksSheet.getLastRow() - 1, 11).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(taskId)) {
      const oldStatus = data[i][10];
      const rowNumber = i + 2;

      tasksSheet.getRange(rowNumber, 3).setValue(new Date());  // 更新日時
      tasksSheet.getRange(rowNumber, 11).setValue(newStatus);  // ステータス
      applyStatusColor(tasksSheet, rowNumber, newStatus);

      addLog(taskId, oldStatus, newStatus, changedBy);
      Logger.log(`タスク ${taskId}: ${oldStatus} → ${newStatus}`);
      return true;
    }
  }

  Logger.log(`タスクが見つかりません: ${taskId}`);
  return false;
}

// ============================================================
// ログ
// ============================================================
function addLog(taskId, oldStatus, newStatus, changedBy) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(SHEET_LOG);

  if (!log) {
    Logger.log('ログシートが見つかりません。setup() を実行してください。');
    return;
  }

  log.appendRow([new Date(), taskId, oldStatus, newStatus, changedBy]);
}

// ============================================================
// ヘルパー関数
// ============================================================
function generateTaskId() {
  return 'TASK-' + new Date().getTime();
}

function taskExistsForMessage(messageId, assignedToId) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(SHEET_TASKS);

  if (!tasksSheet || tasksSheet.getLastRow() <= 1) return false;

  // J列（10番目）= メッセージID、E列（5番目）= 担当者ID
  // 同じメッセージでも担当者が違えば別タスクとして作成する
  const data = tasksSheet.getRange(2, 5, tasksSheet.getLastRow() - 1, 6).getValues();
  return data.some(row => String(row[5]) === String(messageId) && String(row[0]) === String(assignedToId));
}

function lookupMemberName(accountId) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const members = ss.getSheetByName(SHEET_MEMBERS);

  if (!members || members.getLastRow() <= 1) return accountId;

  const data  = members.getRange(2, 1, members.getLastRow() - 1, 2).getValues();
  const found = data.find(row => String(row[0]) === String(accountId));
  return found ? found[1] : accountId; // メンバーシートにない場合はIDをそのまま使用
}

function applyStatusColor(sheet, rowNumber, status) {
  const cell = sheet.getRange(rowNumber, 11); // K列 = ステータス
  if (status === STATUS_PENDING)  cell.setBackground('#FFF9C4'); // 黄色
  if (status === STATUS_FINISHED) cell.setBackground('#C8E6C9'); // 緑
}

// ============================================================
// シート直接編集の検知 — ステータス変更時に自動で色・日時・ログを更新
// ============================================================
function onEdit(e) {
  const sheet = e.range.getSheet();

  // タスクシート以外は無視
  if (sheet.getName() !== SHEET_TASKS) return;

  // K列（11列目）= ステータス列以外は無視
  if (e.range.getColumn() !== 11) return;

  // ヘッダー行は無視
  if (e.range.getRow() === 1) return;

  const newStatus = String(e.value || '').trim();
  const rowNumber = e.range.getRow();
  const validStatuses = [STATUS_PENDING, STATUS_FINISHED];

  // 無効なステータスが入力された場合は元に戻す
  if (!validStatuses.includes(newStatus)) {
    e.range.setValue(e.oldValue || STATUS_PENDING);
    SpreadsheetApp.getUi().alert(
      '無効なステータスです。\n使用できる値：未完了、完了'
    );
    return;
  }

  // 更新日時を自動セット（C列 = 3列目）
  sheet.getRange(rowNumber, 3).setValue(new Date());

  // 色を更新
  applyStatusColor(sheet, rowNumber, newStatus);

  // タスクIDと変更者を取得してログに記録
  const taskId    = sheet.getRange(rowNumber, 1).getValue();
  const oldStatus = e.oldValue || '';
  const changedBy = Session.getActiveUser().getEmail();

  addLog(taskId, oldStatus, newStatus, changedBy);
}
