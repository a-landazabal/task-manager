// ============================================================
// ご案内シート
// ============================================================

function setupGuide() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let g = ss.getSheetByName(SHEET_GUIDE);
  if (g) ss.deleteSheet(g);
  g = ss.insertSheet(SHEET_GUIDE, 0);

  g.setColumnWidth(1, 40);
  g.setColumnWidth(2, 700);

  g.getRange('A:B').setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(10).setWrap(true);

  let row = 1;

  // ── タイトル ──
  g.getRange(row, 1, 1, 2).merge().setValue('Chatwork タスク管理ツール — ご案内')
    .setFontSize(18).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#2C3E50').setHorizontalAlignment('center').setVerticalAlignment('middle');
  g.setRowHeight(row, 55);

  // ── このツールについて ──
  row = 3;
  g.getRange(row, 1, 1, 2).merge().setValue('このツールについて')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#34495E');
  g.setRowHeight(row, 32);
  row++;
  g.getRange(row, 1, 1, 2).merge().setValue(
    'Chatworkのタスクを自動的にこのスプレッドシートに取り込み、Notion・Googleカレンダーと連携するツールです。\n' +
    'Chatwork → スプレッドシート（毎1分） → Notion（毎5分） → Googleカレンダー（毎10分）で自動同期します。'
  ).setFontSize(10).setVerticalAlignment('top');
  g.setRowHeight(row, 50);

  // ── シートの説明 ──
  row = 6;
  g.getRange(row, 1, 1, 2).merge().setValue('各シートの説明')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#34495E');
  g.setRowHeight(row, 32);

  const sheets = [
    ['ダッシュボード', 'タスクの全体像を確認できます。担当者・プロジェクトでフィルターが可能です。'],
    ['タスク',         'Chatworkから取り込んだ全タスクの一覧です。ステータス列（K列）を「完了」に変更するとタスクを完了にできます。'],
    ['メンバー',       'Chatworkのルームに参加しているメンバーの一覧です。毎日自動で更新されます。'],
    ['ルーム',         'Chatworkのルーム一覧です。全ルームが自動的にポーリング対象になります。'],
    ['プロジェクト',    'タスクをプロジェクト別に分類するための設定シートです。'],
    ['ログ',           'タスクのステータス変更履歴です。最新1000件を保持します。'],
  ];

  sheets.forEach((s, i) => {
    const r = row + 1 + i;
    g.getRange(r, 1, 1, 2).merge().setValue('  ' + s[0] + '  —  ' + s[1])
      .setFontSize(10).setBackground(i % 2 === 0 ? '#F8F9FA' : '#FFFFFF');
    g.setRowHeight(r, 28);
  });

  // ── 使い方 ──
  row = row + 1 + sheets.length + 1;
  g.getRange(row, 1, 1, 2).merge().setValue('使い方')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#34495E');
  g.setRowHeight(row, 32);
  row++;
  g.getRange(row, 1, 1, 2).merge().setValue(
    '1. セットアップは完了済みです。自動でChatworkからタスクが取り込まれます。\n' +
    '2.「ダッシュボード」シートでタスクの全体像を確認できます。\n' +
    '3.「タスク」シートのK列（ステータス）を「完了」に変えるとタスクを完了にできます。\n' +
    '4. ダッシュボードの担当者・プロジェクトフィルターで絞り込みができます。'
  ).setFontSize(10).setVerticalAlignment('top');
  g.setRowHeight(row, 75);

  // ── プロジェクト追加方法 ──
  row += 2;
  g.getRange(row, 1, 1, 2).merge().setValue('プロジェクトの追加方法')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#34495E');
  g.setRowHeight(row, 32);
  row++;
  g.getRange(row, 1, 1, 2).merge().setValue(
    'Chatworkのメッセージで以下のコマンドを送信すると、プロジェクトが自動追加されます：\n\n' +
    '    【プロジェクト追加】 プロジェクト名 アカウントID\n' +
    '    【プロジェクト追加】 アカウントID プロジェクト名\n\n' +
    '例:  【プロジェクト追加】 営業管理 12345678\n' +
    '例:  【プロジェクト追加】 12345678 営業管理\n\n' +
    'どちらの順番でもOKです。または「プロジェクト」シートに直接追加することもできます。'
  ).setFontSize(10).setVerticalAlignment('top');
  g.setRowHeight(row, 120);

  // ── Notion連携 ──
  row += 2;
  g.getRange(row, 1, 1, 2).merge().setValue('Notion連携')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#34495E');
  g.setRowHeight(row, 32);
  row++;
  g.getRange(row, 1, 1, 2).merge().setValue(
    'スプレッドシートのタスクは5分ごとにNotionデータベースへ自動同期されます。\n' +
    'Notionでは以下の情報が表示されます：タスク名、担当者、作成者、プロジェクト、期限、ステータス、ルーム名\n\n' +
    'Notionのカンバンビューやカレンダービューで、タスクを視覚的に管理できます。'
  ).setFontSize(10).setVerticalAlignment('top');
  g.setRowHeight(row, 75);

  // ── Googleカレンダー連携 ──
  row += 2;
  g.getRange(row, 1, 1, 2).merge().setValue('Googleカレンダー連携')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#34495E');
  g.setRowHeight(row, 32);
  row++;
  g.getRange(row, 1, 1, 2).merge().setValue(
    '期限があるタスクは10分ごとにGoogleカレンダー（Chatworkタスク）へ自動同期されます。\n\n' +
    '・未完了タスク → 赤色の終日イベント [未完了] 担当者 — タスク内容\n' +
    '・完了タスク → 緑色の終日イベント [完了] 担当者 — タスク内容\n\n' +
    'イベントをクリックすると、担当者・作成者・プロジェクト・ルーム名が確認できます。'
  ).setFontSize(10).setVerticalAlignment('top');
  g.setRowHeight(row, 100);

  // ── 注意事項 ──
  row += 2;
  g.getRange(row, 1, 1, 2).merge().setValue('注意事項')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#E74C3C');
  g.setRowHeight(row, 32);
  row++;
  g.getRange(row, 1, 1, 2).merge().setValue(
    '・ヘッダー行（1行目）は編集しないでください。\n' +
    '・「ダッシュボード」シートはフィルター以外を編集しないでください（自動更新されます）。\n' +
    '・「ルーム」「メンバー」シートは自動管理されています。手動での変更は不要です。\n' +
    '・タスクの変更はK列（ステータス）のみ操作してください。'
  ).setFontSize(10).setVerticalAlignment('top');
  g.setRowHeight(row, 75);

  // ── フッター ──
  row += 2;
  g.getRange(row, 1, 1, 2).merge().setValue('問題が発生した場合は管理者にお問い合わせください。')
    .setFontSize(9).setFontColor('#95A5A6').setHorizontalAlignment('center');

  const protection = g.protect().setDescription('ご案内シート');
  protection.setWarningOnly(true);

  Logger.log('ご案内シート作成完了。');
}
