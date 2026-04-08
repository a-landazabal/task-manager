// ============================================================
// 設定 — 全グローバル定数
// ============================================================

// ── Chatwork API ──
const CHATWORK_API_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
const CHATWORK_API_BASE  = 'https://api.chatwork.com/v2';

// ── シート名 ──
const SHEET_TASKS     = 'タスク';
const SHEET_MEMBERS   = 'メンバー';
const SHEET_ROOMS     = 'ルーム';
const SHEET_LOG       = 'ログ';
const SHEET_DASHBOARD = 'ダッシュボード';
const SHEET_PROJECTS  = 'プロジェクト';
const SHEET_GUIDE     = 'ご案内';

// ── ステータス ──
const STATUS_PENDING  = '未完了';
const STATUS_FINISHED = '完了';

// ── Notion API ──
const NOTION_API_TOKEN  = PropertiesService.getScriptProperties().getProperty('NOTION_API_TOKEN');
const NOTION_DATABASE_ID = PropertiesService.getScriptProperties().getProperty('NOTION_DATABASE_ID');
const NOTION_API_BASE   = 'https://api.notion.com/v1';
const NOTION_VERSION    = '2022-06-28';

// ── 通知メール ──
const NOTIFY_EMAIL = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');

// ── Google Calendar ──
const CALENDAR_ID = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');

// ── バッチ設定 ──
const ROOMS_PER_BATCH       = 5;    // ポーリング: 1回あたりのルーム数
const MEMBERS_BATCH_SIZE    = 20;   // メンバー同期: 1回あたりのルーム数
const FAIL_NOTIFY_THRESHOLD = 10;   // 連続失敗回数でメール通知
const LOG_MAX_ROWS          = 1000; // ログ最大行数
