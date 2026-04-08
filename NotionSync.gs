// ============================================================
// Notion同期 — Google Sheet → Notion Database
// ============================================================

/**
 * タスクシートからNotionデータベースへ同期
 * - 新しいタスクをNotionに追加
 * - ステータス変更を同期
 * - CWタスクIDで重複防止
 */
function syncTasksToNotion() {
  if (!NOTION_API_TOKEN || !NOTION_DATABASE_ID) {
    Logger.log('Notion APIトークンまたはデータベースIDが未設定です。');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const notionPages = getNotionPages();

  let created = 0, updated = 0;

  for (let i = 0; i < data.length; i++) {
    const taskId      = String(data[i][0]);       // A: タスクID
    const description = String(data[i][3]);        // D: 内容
    const assignee    = String(data[i][5]);        // F: 担当者名
    const createdBy   = String(data[i][7]);        // H: 作成者名
    const roomId      = String(data[i][8]);        // I: ルームID
    const cwTaskId    = String(data[i][9]);        // J: CWタスクID
    const status      = String(data[i][10]);       // K: ステータス
    const project     = String(data[i][11] || ''); // L: プロジェクト
    const dueDate     = data[i][12];               // M: 期限

    // ルーム名を取得
    const roomName    = getRoomName(roomId);

    if (!taskId || !cwTaskId) continue;

    const existingPage = notionPages[taskId];

    if (existingPage) {
      // ── 既存ページのステータスを同期 ──
      if (existingPage.status !== status) {
        updateNotionPageStatus(existingPage.id, status);
        updated++;
      }
    } else {
      // ── 新しいページを作成 ──
      createNotionPage({
        taskId      : taskId,
        description : description,
        assignee    : assignee,
        createdBy   : createdBy,
        status      : status,
        dueDate     : dueDate,
        roomName    : roomName,
        project     : project
      });
      created++;
    }

    // API制限回避
    if ((created + updated) % 3 === 0) Utilities.sleep(500);
  }

  Logger.log(`Notion同期完了: 新規${created}件、更新${updated}件`);
}

// ============================================================
// Notion API — ページ作成
// ============================================================

function createNotionPage(task) {
  const properties = {
    'タスク名': {
      title: [{ text: { content: task.description.substring(0, 2000) || task.taskId } }]
    },
    'ステータス': {
      status: { name: task.status === STATUS_FINISHED ? '完了' : '未着手' }
    },
    'タスクID': {
      rich_text: [{ text: { content: task.taskId } }]
    },
    '担当者': {
      rich_text: [{ text: { content: task.assignee } }]
    },
    '作成者': {
      rich_text: [{ text: { content: task.createdBy || '' } }]
    },
    'チャットルーム': {
      rich_text: [{ text: { content: task.roomName || '' } }]
    },
    'プロジェクト': {
      rich_text: [{ text: { content: task.project || '' } }]
    }
  };

  // 期限がある場合のみ追加
  if (task.dueDate && task.dueDate instanceof Date && !isNaN(task.dueDate)) {
    properties['期限'] = {
      date: { start: Utilities.formatDate(task.dueDate, 'Asia/Tokyo', 'yyyy-MM-dd') }
    };
  }

  const payload = {
    parent: { database_id: NOTION_DATABASE_ID },
    properties: properties
  };

  const res = UrlFetchApp.fetch(NOTION_API_BASE + '/pages', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + NOTION_API_TOKEN,
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200) {
    Logger.log(`Notionページ作成エラー [${task.taskId}]: ${code} ${res.getContentText().substring(0, 200)}`);
  }
  return code === 200;
}

// ============================================================
// Notion API — ステータス更新
// ============================================================

function updateNotionPageStatus(pageId, status) {
  const payload = {
    properties: {
      'ステータス': {
        status: { name: status === STATUS_FINISHED ? '完了' : '未着手' }
      }
    }
  };

  const res = UrlFetchApp.fetch(NOTION_API_BASE + '/pages/' + pageId, {
    method: 'PATCH',
    headers: {
      'Authorization': 'Bearer ' + NOTION_API_TOKEN,
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log(`Notionステータス更新エラー [${pageId}]: ${res.getContentText().substring(0, 200)}`);
  }
}

// ============================================================
// Notion API — 既存ページ取得（タスクIDでマッピング）
// ============================================================

function getNotionPages() {
  const pages = {};
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const payload = {
      page_size: 100
    };
    if (startCursor) payload.start_cursor = startCursor;

    const res = UrlFetchApp.fetch(NOTION_API_BASE + '/databases/' + NOTION_DATABASE_ID + '/query', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + NOTION_API_TOKEN,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log(`Notionデータベース取得エラー: ${res.getContentText().substring(0, 200)}`);
      break;
    }

    const result = JSON.parse(res.getContentText());

    result.results.forEach(page => {
      const taskIdProp = page.properties['タスクID'];
      if (taskIdProp && taskIdProp.rich_text && taskIdProp.rich_text.length > 0) {
        const taskId = taskIdProp.rich_text[0].plain_text;
        const statusProp = page.properties['ステータス'];
        const statusName = statusProp && statusProp.status ? statusProp.status.name : '';

        pages[taskId] = {
          id: page.id,
          status: statusName === '完了' ? STATUS_FINISHED : STATUS_PENDING
        };
      }
    });

    hasMore = result.has_more;
    startCursor = result.next_cursor;
  }

  Logger.log(`Notion既存ページ: ${Object.keys(pages).length}件`);
  return pages;
}
