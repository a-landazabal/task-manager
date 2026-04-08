// ============================================================
// Chatwork API ラッパー
// ============================================================

function fetchMessages(roomId, since) {
  const res = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/messages?force=1`, {
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return [];
  const msgs = JSON.parse(res.getContentText());
  return since > 0 ? msgs.filter(m => m.send_time > since) : msgs;
}

function fetchChatworkTasks(roomId, status) {
  const res = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/tasks?status=${status}`, {
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code === 429) {
    Utilities.sleep(5000);
    const retry = UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/tasks?status=${status}`, {
      headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, muteHttpExceptions: true
    });
    return retry.getResponseCode() === 200 ? JSON.parse(retry.getContentText()) : [];
  }
  return code === 200 ? JSON.parse(res.getContentText()) : [];
}
