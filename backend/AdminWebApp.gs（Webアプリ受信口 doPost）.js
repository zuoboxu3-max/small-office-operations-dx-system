/****************************************************
 * AdminWebApp.gs（統合版）
 * - doPost で log / trigger_status を受信しシート更新
 ****************************************************/

function doPost(e) {
  try {
    const contents = (e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const body = JSON.parse(contents);

    const expectedToken = PropertiesService.getScriptProperties().getProperty(ADMIN_CONFIG.TOKEN_KEY);
    if (!body.token || body.token !== expectedToken) {
      return _json_({ ok: false, error: "unauthorized" });
    }

    const action = body.action;
    const payload = body.payload;

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      if (action === "log") {
        _appendLogUnified_(payload);
      } else if (action === "trigger_status") {
        _writeTriggerStatusUnified_(payload);
      }
      return _json_({ ok: true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return _json_({ ok: false, error: String(err && err.stack ? err.stack : err) });
  }
}

function _json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Log 追記：正式フォーマットに統一
 * ["日時","種類","項目","内容","実行ID","実行者","詳細データ"]
 */
function _appendLogUnified_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Log") || _ensureSheet_(ss, "Log", LOG_HEADER);

  const ts = p.timestamp ? new Date(p.timestamp) : new Date();
  const level = p.level || "INFO";
  const item = p.action || "";
  const content = p.message || "";
  const runId = p.runId || "";
  const executor = p.executor || "";
  const extra = p.context ? JSON.stringify(p.context) : "";

  sh.appendRow([ts, level, item, content, runId, executor, extra]);
}

/**
 * TriggerStatus 更新：日本語ヘッダで統一
 * ["日時","スクリプトID","処理名","種別","補足"]
 */
function _writeTriggerStatusUnified_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("TriggerStatus") || _ensureSheet_(ss, "TriggerStatus", TRIGGER_HEADER);

  sh.clearContents();
  sh.getRange(1, 1, 1, TRIGGER_HEADER.length).setValues([TRIGGER_HEADER]);

  const rows = (p.triggers || []).map(t => [
    new Date(),
    p.scriptId || "",
    t.handler || "",
    t.type || "",
    ""
  ]);

  if (rows.length > 0) sh.getRange(2, 1, rows.length, TRIGGER_HEADER.length).setValues(rows);
}
