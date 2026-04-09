// ============================================================
// コマンド処理（【プロジェクト追加】）
// ============================================================

function processCommands(messages) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROJECTS);
  messages.forEach(m => {
    const body = (m.body || '').trim();

    // フォーマット1: 【プロジェクト追加】 名前 ID
    // フォーマット2: 【プロジェクト追加】 ID 名前
    const match1 = body.match(/【プロジェクト追加】\s*(.+?)[\s　]+(\d+)\s*$/);
    const match2 = body.match(/【プロジェクト追加】\s*(\d+)[\s　]+(.+)/);

    let name, id;
    if (match1) {
      name = match1[1].trim();
      id   = match1[2].trim();
    } else if (match2) {
      id   = match2[1].trim();
      name = match2[2].trim();
    } else {
      return;
    }

    const existing = getProjectMap();
    if (existing[id]) return;
    sheet.appendRow([id, name]);
    Logger.log(`プロジェクト追加: ${name} (${id})`);
  });
}
