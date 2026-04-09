// ============================================================
// Chatwork API ラッパー
// ============================================================

function fetchMessages(roomId, since) {
  var url = CHATWORK_API_BASE + '/rooms/' + roomId + '/messages?force=1';
  var options = {
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN },
    muteHttpExceptions: true,
  };

  for (var attempt = 0; attempt < 3; attempt++) {
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    if (code === 200) {
      var msgs = JSON.parse(res.getContentText());
      return since > 0 ? msgs.filter(function (m) { return m.send_time > since; }) : msgs;
    }
    if (code !== 429) return [];
    var wait = 5000 * Math.pow(2, attempt);
    Logger.log('Chatwork 429 - attempt ' + (attempt + 1) + ', waiting ' + (wait / 1000) + 's');
    Utilities.sleep(wait);
  }
  Logger.log('Chatwork messages failed after 3 retries: room ' + roomId);
  return [];
}

function fetchChatworkTasks(roomId, status) {
  var url = CHATWORK_API_BASE + '/rooms/' + roomId + '/tasks?status=' + status;
  var options = {
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN },
    muteHttpExceptions: true,
  };

  for (var attempt = 0; attempt < 3; attempt++) {
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText());
    if (code !== 429) return [];
    var wait = 5000 * Math.pow(2, attempt);
    Logger.log('Chatwork 429 - attempt ' + (attempt + 1) + ', waiting ' + (wait / 1000) + 's');
    Utilities.sleep(wait);
  }
  Logger.log('Chatwork tasks failed after 3 retries: room ' + roomId);
  return [];
}
