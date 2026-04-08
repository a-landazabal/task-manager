// ============================================================
// セットアップ — 一度だけ実行
// ============================================================

function setup() {
  // ── APIトークン検証 ──
  if (!CHATWORK_API_TOKEN) {
    SpreadsheetApp.getUi().alert(
      'セットアップエラー',
      'Chatwork APIトークンが設定されていません。\n\n' +
      '設定方法:\n' +
      '1. Apps Script エディタを開く\n' +
      '2.「プロジェクトの設定」→「スクリプト プロパティ」\n' +
      '3. プロパティ名: CHATWORK_API_TOKEN\n' +
      '4. 値: ChatWorkの設定 → API で取得したトークン',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // ── APIトークンの接続テスト ──
  const testRes = UrlFetchApp.fetch(CHATWORK_API_BASE + '/me', {
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
  });
  if (testRes.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert(
      '接続エラー',
      'Chatwork APIに接続できませんでした。\n' +
      'トークンが正しいか確認してください。\n\n' +
      'エラーコード: ' + testRes.getResponseCode(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets();

  // ルームとメンバーが既にある場合はスキップ
  if (!ss.getSheetByName(SHEET_ROOMS) || ss.getSheetByName(SHEET_ROOMS).getLastRow() <= 1) syncRooms();
  if (!ss.getSheetByName(SHEET_MEMBERS) || ss.getSheetByName(SHEET_MEMBERS).getLastRow() <= 1) syncMembers();

  setupDashboard();
  setupGuide();
  createTrigger();
  createEditTrigger();

  // ご案内シートを最初に表示
  const guide = ss.getSheetByName(SHEET_GUIDE);
  if (guide) ss.setActiveSheet(guide);

  Logger.log('セットアップ完了！全ルームを自動的にバッチポーリングします。');
}

// ============================================================
// シート初期化
// ============================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── タスク（13列）──
  let s = ss.getSheetByName(SHEET_TASKS);
  if (!s) s = ss.insertSheet(SHEET_TASKS);
  if (s.getLastRow() === 0) {
    s.appendRow(['タスクID','作成日時','更新日時','内容','担当者ID','担当者名','作成者ID','作成者名','ルームID','CWタスクID','ステータス','プロジェクト','期限']);
    s.getRange(1,1,1,13).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }

  // ステータス列にドロップダウン（K列、2行目以降）
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS_PENDING, STATUS_FINISHED])
    .setAllowInvalid(false).build();
  s.getRange(2, 11, s.getMaxRows() - 1, 1).setDataValidation(statusRule);

  // シート保護
  const taskProtections = s.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  taskProtections.forEach(p => p.remove());
  const tp = s.protect().setDescription('タスクシート保護');
  tp.setWarningOnly(true);

  // ── メンバー（4列）──
  s = ss.getSheetByName(SHEET_MEMBERS);
  if (!s) s = ss.insertSheet(SHEET_MEMBERS);
  if (s.getLastRow() === 0) {
    s.appendRow(['アカウントID','名前','ルームID','ルーム名']);
    s.getRange(1,1,1,4).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }
  s.protect().setDescription('メンバーシート（自動管理）').setWarningOnly(true);

  // ── ルーム（3列）──
  s = ss.getSheetByName(SHEET_ROOMS);
  if (!s) s = ss.insertSheet(SHEET_ROOMS);
  if (s.getLastRow() === 0) {
    s.appendRow(['ルームID','ルーム名','最終確認日時']);
    s.getRange(1,1,1,3).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }
  s.protect().setDescription('ルームシート（自動管理）').setWarningOnly(true);

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
  s.protect().setDescription('ログシート（自動管理）').setWarningOnly(true);

  Logger.log('シート初期化完了。');
}

// ============================================================
// トリガー管理
// ============================================================

function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const name = t.getHandlerFunction();
    if (name === 'pollChatWork' || name === 'dailySync' || name === 'memberSync' || name === 'notionSync' || name === 'calendarSync') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('pollChatWork').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('dailySync').timeBased().atHour(6).everyDays(1).create();
  ScriptApp.newTrigger('memberSync').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('notionSync').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('calendarSync').timeBased().everyMinutes(10).create();

  Logger.log('トリガー設定完了（ポーリング + 日次ルーム同期 + 毎時メンバー同期 + Notion同期 + カレンダー同期）。');
}

function dailySync() {
  syncRooms();
  Logger.log('日次ルーム同期完了。');
}

function memberSync() {
  syncMembers();
}

function notionSync() {
  syncTasksToNotion();
}

function calendarSync() {
  syncNotionToCalendar();
}

function createEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  Logger.log('編集トリガー設定完了。');
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
