# Task Manager — Chatwork × Google Sheets × Notion × Google Calendar

## Overview

Google Apps Script (GAS) project that syncs Chatwork tasks into a Google Spreadsheet, pushes them to a Notion database, and creates Google Calendar events. Modular file structure — each `.gs` file handles a specific concern.

## Data Flow

```
Chatwork API → Google Sheet (every 1 min) → Notion DB (every 5 min) → Google Calendar (every 10 min)
```

## File Structure

| File | Purpose |
|---|---|
| `Config.gs` | All global constants (API tokens, sheet names, batch sizes, calendar ID) |
| `ChatworkApi.gs` | Chatwork REST API wrappers (messages, tasks) |
| `SheetHelpers.gs` | Shared spreadsheet utilities (project map, member names, logging) |
| `TaskSync.gs` | Chatwork polling, task create/complete, room/member sync |
| `Commands.gs` | Chat command processing (【プロジェクト追加】, supports both name/ID order) |
| `Dashboard.gs` | Dashboard setup and refresh with filters (担当者, プロジェクト) |
| `Guide.gs` | ご案内 (user guide) sheet |
| `Triggers.gs` | setup(), onEdit(), trigger creation, sheet initialization |
| `NotionSync.gs` | Google Sheet → Notion database sync (creates/updates pages) |
| `CalendarSync.gs` | Notion → Google Calendar sync (creates/updates all-day events) |

## Architecture

- **Runtime**: Google Apps Script (V8), deployed inside a Google Spreadsheet
- **External APIs**: Chatwork REST API v2, Notion API v2022-06-28, Google Calendar (built-in CalendarApp)
- **Auth**: API tokens stored in Apps Script Script Properties (not in code)
- **GAS module system**: All `.gs` files share a single global scope — no imports/exports needed
- **No build step** — files are copy-pasted or clasp-pushed directly to GAS

## Script Properties Required

| Property | Value |
|---|---|
| `CHATWORK_API_TOKEN` | Chatwork API token (from ChatWork設定 → API) |
| `NOTION_API_TOKEN` | Notion internal integration secret (starts with `ntn_`) |
| `NOTION_DATABASE_ID` | Notion database ID (from the page URL) |
| `CALENDAR_ID` | Google Calendar ID (from calendar settings → カレンダーの統合) |

## Spreadsheet Sheets

| Sheet Name | Purpose |
|---|---|
| ご案内 | User guide (protected, warning-only) |
| ダッシュボード | Dashboard with filters (担当者, プロジェクト) and summary stats (9 columns) |
| タスク | Main task registry (13 columns: タスクID〜期限) |
| メンバー | Chatwork members per room (4 columns, auto-managed) |
| ルーム | Chatwork rooms + last checked timestamp (3 columns) |
| プロジェクト | Account ID → project name mapping (2 columns) |
| ログ | Change audit trail (5 columns, capped at 1000 rows) |

## Notion Database Properties

| Property | Type | Mapped from |
|---|---|---|
| タスク名 | Title | タスク sheet column D (内容) |
| ステータス | Status | タスク sheet column K (未着手/完了) |
| 期限 | Date | タスク sheet column M |
| 担当者名 | Text | タスク sheet column F (担当者名) |
| チャットルーム | Text | タスク sheet column I (ルームID) |
| タスクID | Text | タスク sheet column A (TASK-timestamp) |
| 担当者 | Person | (unused by sync — Notion template default) |

## Google Calendar Events

- Calendar: `Chatworkタスク` (separate from team calendar IT事業部)
- Event type: all-day events on the 期限 date
- Title format: `[未完了] 担当者名 — タスク内容` or `[完了] 担当者名 — タスク内容`
- Color: Red = 未完了, Green = 完了
- Description includes: 担当者, プロジェクト, ステータス, ルーム, タスクID
- Duplicate prevention: matches events by タスクID in description
- Completed tasks stay on calendar (color changes to green, title updates)

## Triggers

| Trigger | Interval | Function | File |
|---|---|---|---|
| pollChatWork | Every 5 min | Batch poll Chatwork rooms for tasks/messages | TaskSync.gs |
| dailySync | Daily at 6:00 | Sync new Chatwork rooms | Triggers.gs |
| memberSync | Every 1 hour | Batch sync members (20 rooms/run) | Triggers.gs |
| notionSync | Every 5 min | Sync tasks from Sheet → Notion | Triggers.gs |
| calendarSync | Every 10 min | Sync tasks from Notion → Calendar | Triggers.gs |
| onEdit | On spreadsheet edit | Dashboard filter + status change handling | Triggers.gs |

## Key Flows

1. **Setup** (`setup()` in Triggers.gs) — validates Chatwork API token, creates sheets, syncs rooms/members, sets up all triggers
2. **Chatwork Polling** (`pollChatWork()` in TaskSync.gs) — processes 5 rooms per batch in round-robin rotation, syncs open/done tasks
3. **Task Sync** (`syncChatworkTasks()` in TaskSync.gs) — fetches tasks, deduplicates via batched Set, detects project labels, includes 期限
4. **Commands** (`processCommands()` in Commands.gs) — `【プロジェクト追加】` accepts both `name ID` and `ID name` formats
5. **Dashboard** (`refreshDashboard()` in Dashboard.gs) — auto-updates filter dropdowns, 9-column display with 期限
6. **Notion Sync** (`syncTasksToNotion()` in NotionSync.gs) — creates new pages, updates status changes, uses タスクID for dedup
7. **Calendar Sync** (`syncNotionToCalendar()` in CalendarSync.gs) — creates all-day events for tasks with 期限, updates color on completion
8. **Error notification** — emails owner every 10 consecutive polling failures (10, 20, 30, etc.)
9. **Stale cleanup** — `syncRooms()` removes rooms no longer in Chatwork, `syncMembers()` removes members who left rooms

## API Rate Limit Handling

- `fetchChatworkTasks()` and `fetchMessages()` retry 3 times with exponential backoff (5s, 10s, 20s) on HTTP 429
- `syncMembers()` sleeps 1s between rooms, skips on 429 (10s wait)
- Batch rotation keeps Chatwork polling to ~10 API calls per run
- Member sync: 20 rooms/hour, full rotation scales with room count
- Notion sync: 500ms sleep every 3 operations
- GAS daily quota: 20,000 URL fetches/day (free), 100,000 (Workspace)

## Daily API Usage Estimate (8:30-20:00)

| Function | Interval | Calls/run | Runs/day | Total |
|---|---|---|---|---|
| pollChatWork | 5 min | ~10 | 138 | ~1,380 |
| notionSync | 5 min | ~5 | 138 | ~690 |
| calendarSync | 10 min | ~5 | 69 | ~345 |
| memberSync | 1 hour | ~20 | 11 | ~220 |
| dailySync | 1x at 6 AM | ~2 | 1 | ~2 |
| **Total** | | | | **~2,600** |

Approximately 13% of the 20,000 GAS free daily quota.

## Working Hours

- All triggers except `dailySync` and `onEdit` are guarded by `isOutsideWorkingHours()` (8:30-20:00)
- `dailySync` runs at 6:00 AM (before working hours) to sync rooms before polling starts
- `onEdit` always responds (user-initiated)

## Conventions

- All UI text and sheet names are in Japanese
- Status values: `未完了` (pending) / `完了` (done)
- Notion status mapping: `未着手` = 未完了, `完了` = 完了
- Task IDs: `TASK-{timestamp}` format
- Color coding (Sheet): `#FFF9C4` = pending (yellow), `#C8E6C9` = done (green)
- Color coding (Calendar): Yellow = pending, Red = overdue (past due date), Green = done
- Notion emoji status: 🟡 = pending, 🟢 = done (prefixed to task title)
- Sheet protection: warning-only (editors can override, viewers cannot interact)
- Project detection: Chatwork accounts in プロジェクト sheet are treated as labels, not assignees

## Development

- Edit `.gs` files locally, push to GAS via [clasp](https://github.com/google/clasp) or paste manually
- API tokens go in Script Properties, not in code — `.env` is for local reference only
- Test by running `setup()` in the Apps Script editor
- All files share GAS global scope — functions can call across files without imports
- When updating: only paste changed files into Apps Script, unchanged files don't need to be touched
