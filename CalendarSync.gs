// ============================================================
// カレンダー同期 — Notion → Google Calendar
// ============================================================

/**
 * Notionデータベースからタスクを取得し、Google Calendarに終日イベントを作成
 * - 期限があるタスクのみ対象
 * - タスクIDで重複防止
 * - ステータス変更を同期（完了→イベント削除 or 色変更）
 */
function syncNotionToCalendar() {
  if (isOutsideWorkingHours()) return;

  if (!NOTION_API_TOKEN || !NOTION_DATABASE_ID || !CALENDAR_ID) {
    Logger.log('Notion/Calendar設定が未完了です。');
    return;
  }

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    Logger.log('カレンダーが見つかりません: ' + CALENDAR_ID);
    return;
  }

  // ── Notionから期限ありのタスクを取得 ──
  const notionTasks = getNotionTasksWithDueDate();
  if (notionTasks.length === 0) {
    Logger.log('カレンダー同期: 期限付きタスクなし');
    return;
  }

  // ── 既存カレンダーイベントを取得（タスクIDで照合） ──
  const existingEvents = getExistingCalendarEvents(calendar);

  let created = 0,
    updated = 0,
    removed = 0;

  notionTasks.forEach((task) => {
    const existing = existingEvents[task.taskId];

    const statusLabel = task.status === '完了' ? '[完了]' : '[未完了]';
    const title = `${statusLabel} ${task.assignee} — ${task.title}`;
    const description = [
      '━━━━━━━━━━━━━━━━',
      '担当者: ' + task.assignee,
      '作成者: ' + (task.createdBy || 'なし'),
      'プロジェクト: ' + (task.project || 'なし'),
      'ステータス: ' + task.status,
      'ルーム: ' + (task.roomName || 'なし'),
      'タスクID: ' + task.taskId,
      '━━━━━━━━━━━━━━━━',
    ].join('\n');

    // 色: 完了=緑、期限超過=赤、未完了=黄
    var color;
    if (task.status === '完了') {
      color = CalendarApp.EventColor.GREEN;
    } else {
      var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      color = task.dueDate < today ? CalendarApp.EventColor.RED : CalendarApp.EventColor.YELLOW;
    }

    var eventDate = new Date(task.dueDate + 'T00:00:00+09:00');

    if (existing) {
      // ── 既存イベントを更新 ──
      var event = existing.event;
      var needsUpdate = event.getTitle() !== title;
      var currentColor = event.getColor();
      var newColorId = color === CalendarApp.EventColor.GREEN ? '10'
        : color === CalendarApp.EventColor.RED ? '11' : '5';

      if (needsUpdate || currentColor !== newColorId) {
        event.deleteEvent();
        var newEvent = calendar.createAllDayEvent(title, eventDate);
        newEvent.setDescription(description);
        newEvent.setColor(color);
        newEvent.setTag('taskId', task.taskId);
        updated++;
      }
    } else {
      // ── 新規イベント作成 ──
      var event = calendar.createAllDayEvent(title, eventDate);
      event.setDescription(description);
      event.setColor(color);
      event.setTag('taskId', task.taskId);
      created++;
    }
  });

  Logger.log(`カレンダー同期完了: 新規${created}件、更新${updated}件、削除${removed}件`);
}

// ============================================================
// Notion API — 期限ありタスクを取得
// ============================================================

function getNotionTasksWithDueDate() {
  const tasks = [];
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const payload = {
      page_size: 100,
      filter: {
        property: '期限',
        date: { is_not_empty: true },
      },
    };
    if (startCursor) payload.start_cursor = startCursor;

    const res = UrlFetchApp.fetch(NOTION_API_BASE + '/databases/' + NOTION_DATABASE_ID + '/query', {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + NOTION_API_TOKEN,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() !== 200) {
      Logger.log(`Notionタスク取得エラー: ${res.getContentText().substring(0, 200)}`);
      break;
    }

    const result = JSON.parse(res.getContentText());

    result.results.forEach((page) => {
      const props = page.properties;

      // タスク名
      const titleProp = props['タスク名'];
      const title =
        titleProp && titleProp.title && titleProp.title.length > 0
          ? titleProp.title[0].plain_text
          : '';

      // タスクID
      const taskIdProp = props['タスクID'];
      const taskId =
        taskIdProp && taskIdProp.rich_text && taskIdProp.rich_text.length > 0
          ? taskIdProp.rich_text[0].plain_text
          : '';

      // 担当者
      const assigneeProp = props['担当者'];
      const assignee =
        assigneeProp && assigneeProp.rich_text && assigneeProp.rich_text.length > 0
          ? assigneeProp.rich_text[0].plain_text
          : '';

      // 作成者
      const createdByProp = props['作成者'];
      const createdBy =
        createdByProp && createdByProp.rich_text && createdByProp.rich_text.length > 0
          ? createdByProp.rich_text[0].plain_text
          : '';

      // チャットルーム
      const roomProp = props['チャットルーム'];
      const roomName =
        roomProp && roomProp.rich_text && roomProp.rich_text.length > 0
          ? roomProp.rich_text[0].plain_text
          : '';

      // ステータス
      const statusProp = props['ステータス'];
      const status = statusProp && statusProp.status ? statusProp.status.name : '';

      // 期限
      const dateProp = props['期限'];
      const dueDate = dateProp && dateProp.date ? dateProp.date.start : '';

      // プロジェクト
      const projectProp = props['プロジェクト'];
      const project =
        projectProp && projectProp.rich_text && projectProp.rich_text.length > 0
          ? projectProp.rich_text[0].plain_text
          : '';

      if (taskId && dueDate) {
        tasks.push({
          title: title,
          taskId: taskId,
          assignee: assignee,
          createdBy: createdBy,
          roomName: roomName,
          status: status,
          dueDate: dueDate,
          project: project,
        });
      }
    });

    hasMore = result.has_more;
    startCursor = result.next_cursor;
  }

  Logger.log(`Notionタスク取得: ${tasks.length}件（期限あり）`);
  return tasks;
}

// ============================================================
// Google Calendar — 既存イベント取得（タスクIDで照合）
// ============================================================

function getExistingCalendarEvents(calendar) {
  const events = {};

  // 過去30日〜未来90日のイベントを検索
  const now = new Date();
  const start = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const end = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);

  const calEvents = calendar.getEvents(start, end);

  calEvents.forEach((event) => {
    // descriptionからタスクIDを抽出
    const desc = event.getDescription() || '';
    const match = desc.match(/タスクID:\s*(TASK-\d+)/);
    if (match) {
      events[match[1]] = { event: event };
    }
  });

  Logger.log(`既存カレンダーイベント: ${Object.keys(events).length}件`);
  return events;
}
