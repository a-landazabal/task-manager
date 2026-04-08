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
  dash.getRange('A1:I1').merge().setValue('タスク管理ダッシュボード')
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
  dash.getRange('A4:I4').setBackground('#2C3E50');

  // テーブルヘッダー
  dash.getRange(5,1,1,9).setValues([['No.','作成日時','内容','担当者','作成者','プロジェクト','期限','ステータス','更新日時']])
    .setFontWeight('bold').setFontSize(10).setBackground('#2C3E50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  dash.setRowHeight(5, 28);
  dash.setFrozenRows(5);

  dash.setColumnWidth(1,50); dash.setColumnWidth(2,140); dash.setColumnWidth(3,280);
  dash.setColumnWidth(4,130); dash.setColumnWidth(5,130); dash.setColumnWidth(6,130);
  dash.setColumnWidth(7,110); dash.setColumnWidth(8,90); dash.setColumnWidth(9,140);

  refreshDashboard();
  Logger.log('ダッシュボード作成完了。');
}

function refreshDashboard() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SHEET_DASHBOARD);
  const task = ss.getSheetByName(SHEET_TASKS);
  if (!dash) return;

  // ── フィルタードロップダウンを最新に更新 ──
  const currentPerson  = dash.getRange('B2').getValue();
  const currentProject = dash.getRange('F2').getValue();
  const mRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['すべて'].concat(getUniqueMemberNames())).setAllowInvalid(false).build();
  dash.getRange('B2').setDataValidation(mRule);
  if (currentPerson) dash.getRange('B2').setValue(currentPerson);
  const pRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['すべて'].concat(getUniqueProjectNames())).setAllowInvalid(false).build();
  dash.getRange('F2').setDataValidation(pRule);
  if (currentProject) dash.getRange('F2').setValue(currentProject);

  const startRow = 6, cols = 9;

  if (!task || task.getLastRow() <= 1) {
    if (dash.getLastRow() >= startRow) { const r=dash.getRange(startRow,1,dash.getLastRow()-startRow+1,cols); r.breakApart(); r.clearContent().clearFormat(); }
    dash.getRange('B3').setValue('0'); dash.getRange('D3').setValue('0');
    dash.getRange('F3').setValue('0'); dash.getRange('H3').setValue('0%');
    return;
  }

  const fPerson  = dash.getRange('B2').getValue();
  const fProject = dash.getRange('F2').getValue();
  const lastCol  = Math.max(task.getLastColumn(), 13);
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

  const out = filtered.map((r,i) => [i+1, r[1], r[3], r[5], r[7], String(r[11]||''), r[12] || '', r[10], r[2]]);
  dash.getRange(startRow, 1, out.length, cols).setValues(out);

  for (let i = 0; i < out.length; i++) {
    const row = startRow + i;
    const status = String(out[i][7]);
    dash.getRange(row,1,1,cols).setBackground(i%2===0?'#FFFFFF':'#F8F9FA')
      .setBorder(false,false,true,false,false,false,'#ECF0F1',SpreadsheetApp.BorderStyle.SOLID);
    const sc = dash.getRange(row, 8);
    if (status === STATUS_PENDING)  sc.setBackground('#FFF9C4').setFontColor('#F39C12').setFontWeight('bold');
    if (status === STATUS_FINISHED) sc.setBackground('#C8E6C9').setFontColor('#27AE60').setFontWeight('bold');
    dash.getRange(row, 1).setHorizontalAlignment('center');
    dash.getRange(row, 7).setHorizontalAlignment('center');
  }
  Logger.log(`ダッシュボード更新: ${total}件`);
}
