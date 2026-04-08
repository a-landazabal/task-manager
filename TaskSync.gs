// ============================================================
// メインポーリング — 全ルームをバッチローテーションでポーリング
// ============================================================

function pollChatWork() {
  const props = PropertiesService.getScriptProperties();

  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const roomsSheet = ss.getSheetByName(SHEET_ROOMS);
    if (!roomsSheet || roomsSheet.getLastRow() <= 1) return;

    const rows      = roomsSheet.getRange(2, 1, roomsSheet.getLastRow() - 1, 3).getValues();
    const totalRooms = rows.length;
    if (totalRooms === 0) return;

    const startIdx = Number(props.getProperty('POLL_INDEX') || 0) % totalRooms;
    let processed  = 0;

    for (let n = 0; n < ROOMS_PER_BATCH && n < totalRooms; n++) {
      const i           = (startIdx + n) % totalRooms;
      const roomId      = String(rows[i][0]).trim();
      const roomName    = String(rows[i][1]).trim();
      const lastChecked = rows[i][2] ? Number(rows[i][2]) : 0;

      if (!roomId) continue;

      try {
        syncChatworkTasks(roomId, roomName);

        if (lastChecked === 0) {
          roomsSheet.getRange(i + 2, 3).setValue(Math.floor(Date.now() / 1000));
          processed++;
          continue;
        }

        const messages = fetchMessages(roomId, lastChecked);
        if (messages.length > 0) {
          processCommands(messages);
          roomsSheet.getRange(i + 2, 3).setValue(Math.max(...messages.map(m => m.send_time)));
        }
        processed++;
      } catch (e) {
        Logger.log(`ルーム ${roomId} エラー: ${e.message}`);
      }
    }

    props.setProperty('POLL_INDEX', String((startIdx + ROOMS_PER_BATCH) % totalRooms));
    refreshDashboard();

    props.setProperty('POLL_FAIL_COUNT', '0');
    Logger.log(`ポーリング完了: ${processed}/${ROOMS_PER_BATCH}件（位置 ${startIdx}〜）`);

  } catch (e) {
    const failCount = Number(props.getProperty('POLL_FAIL_COUNT') || 0) + 1;
    props.setProperty('POLL_FAIL_COUNT', String(failCount));
    Logger.log(`ポーリング全体エラー (${failCount}回目): ${e.message}`);

    if (failCount === FAIL_NOTIFY_THRESHOLD) {
      const owner = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
      if (owner) {
        MailApp.sendEmail(owner,
          '【タスク管理ツール】ポーリングエラー通知',
          'タスク管理ツールのポーリングが ' + FAIL_NOTIFY_THRESHOLD + ' 回連続で失敗しました。\n\n' +
          'エラー内容: ' + e.message + '\n\n' +
          'スプレッドシートを確認してください。'
        );
        Logger.log('エラー通知メール送信済み: ' + owner);
      }
    }
  }
}

// ============================================================
// Chatworkタスク同期 — プロジェクト検出 + 重複防止 + 完了同期
// ============================================================

function syncChatworkTasks(roomId, roomName) {
  const projectMap   = getProjectMap();
  const existingCwIds = getExistingCwTaskIds();

  const openTasks = fetchChatworkTasks(roomId, 'open');

  const projectByGroup = {};
  openTasks.forEach(t => {
    const aid = String(t.account.account_id);
    const cid = String(t.assigned_by_account.account_id);
    const key = cid + '|' + (t.body || '');
    if (projectMap[aid]) projectByGroup[key] = projectMap[aid];
    if (projectMap[cid]) projectByGroup[key] = projectMap[cid];
  });

  openTasks.forEach(t => {
    const cwTaskId   = String(t.task_id);
    const aid        = String(t.account.account_id);
    const cid        = String(t.assigned_by_account.account_id);
    const key        = cid + '|' + (t.body || '');

    if (projectMap[aid]) return;
    if (existingCwIds.has(cwTaskId)) return;

    createTask({
      description   : t.body || '',
      assignedToId  : aid,
      assignedToName: t.account.name,
      createdById   : cid,
      createdByName : t.assigned_by_account.name,
      roomId        : roomId,
      cwTaskId      : cwTaskId,
      project       : projectByGroup[key] || '',
      dueDate       : t.limit_time ? new Date(t.limit_time * 1000) : ''
    });
    existingCwIds.add(cwTaskId);
  });

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
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TASKS);
  const now   = new Date();

  sheet.appendRow([
    'TASK-' + now.getTime(), now, now,
    d.description, d.assignedToId, d.assignedToName,
    d.createdById, d.createdByName, d.roomId,
    d.cwTaskId, STATUS_PENDING, d.project || '', d.dueDate || ''
  ]);

  const row = sheet.getLastRow();
  sheet.getRange(row, 11).setBackground('#FFF9C4');

  addLog('TASK-' + now.getTime(), '', STATUS_PENDING, d.createdByName);
  Logger.log(`タスク作成: ${d.assignedToName} — ${d.description.substring(0,30)}`);
}

function markTaskDoneByCwId(cwTaskId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][9]) === cwTaskId && String(data[i][10]) === STATUS_PENDING) {
      const row = i + 2;
      sheet.getRange(row, 3).setValue(new Date());
      sheet.getRange(row, 11).setValue(STATUS_FINISHED);
      sheet.getRange(row, 11).setBackground('#C8E6C9');
      addLog(data[i][0], STATUS_PENDING, STATUS_FINISHED, 'Chatwork');
      Logger.log(`タスク完了: ${data[i][0]}`);
    }
  }
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
      if (!existing.has(id)) { newRows.push([id, r.name, '']); existing.add(id); }
    });
    if (newRows.length > 0) sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, 3).setValues(newRows);
    Logger.log(`ルーム同期: ${newRows.length}件追加（合計${existing.size}件）`);
  } catch(e) { Logger.log(`ルーム取得エラー: ${e.message}`); }
}

// ============================================================
// メンバー自動取得（バッチ処理）
// ============================================================

function syncMembers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName(SHEET_ROOMS);
  const mSheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!rSheet || rSheet.getLastRow() <= 1) return;

  const props  = PropertiesService.getScriptProperties();
  const rooms  = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 2).getValues();
  const total  = rooms.length;
  const startIdx = Number(props.getProperty('MEMBER_SYNC_INDEX') || 0) % total;

  const existing = new Set();
  if (mSheet.getLastRow() > 1) {
    mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 4).getValues()
      .forEach(r => existing.add(String(r[0]) + '_' + String(r[2])));
  }

  const newRows = [];
  const batchSize = Math.min(MEMBERS_BATCH_SIZE, total);

  for (let n = 0; n < batchSize; n++) {
    const i = (startIdx + n) % total;
    const roomId   = String(rooms[i][0]).trim();
    const roomName = String(rooms[i][1]).trim();
    if (!roomId) continue;

    try {
      if (n > 0) Utilities.sleep(1000);
      const res = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/members`, {
        headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code === 429) { Utilities.sleep(10000); continue; }
      if (code !== 200) continue;
      JSON.parse(res.getContentText()).forEach(m => {
        const key = String(m.account_id) + '_' + roomId;
        if (!existing.has(key)) { newRows.push([String(m.account_id), m.name, roomId, roomName]); existing.add(key); }
      });
    } catch (e) { Logger.log(`メンバー取得エラー [${roomId}]: ${e.message}`); }
  }

  if (newRows.length > 0) mSheet.getRange(mSheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  if (mSheet.getLastRow() > 2) mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 4).sort({ column: 4, ascending: true });

  props.setProperty('MEMBER_SYNC_INDEX', String((startIdx + batchSize) % total));
  Logger.log(`メンバー同期: ${newRows.length}名追加（ルーム ${startIdx}〜${startIdx + batchSize - 1}）`);
}
