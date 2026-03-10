/**
 * LoggerClient.gs（台帳側：管理者Webアプリへ送信するための関数定義）
 * これがプロジェクト内に存在しないと "is not defined" エラーになります。
 */

// 前回の画像で取得したURLと、adminSetupで出たトークンをここに貼る//
const ADMIN_WEBAPP_URL = 'YOUR_ADMIN_WEBAPP_URL';
const ADMIN_WEBAPP_TOKEN = 'YOUR_ADMIN_WEBAPP_TOKEN';

/** 管理者へINFOログを送信 */
function logInfoToAdmin_(runId, action, message, context) {
  _postToAdmin_("log", {
    timestamp: new Date().toISOString(),
    level: "INFO",
    runId,
    action,
    message,
    executor: _safeExecutorEmail_(),
    context: context || null
  });
}

/** 管理者へERRORログを送信 */
function logErrorToAdmin_(runId, action, err, context) {
  const msg = (err && err.stack) ? err.stack : String(err);
  _postToAdmin_("log", {
    timestamp: new Date().toISOString(),
    level: "ERROR",
    runId,
    action,
    message: msg,
    executor: _safeExecutorEmail_(),
    context: context || null
  });
}

/** トリガー状態を送信 */
function postTriggerStatusToAdmin_() {
  try {
    const triggers = ScriptApp.getProjectTriggers().map(t => ({
      handler: t.getHandlerFunction(),
      type: String(t.getEventType())
    }));
    _postToAdmin_("trigger_status", {
      scriptId: ScriptApp.getScriptId(),
      triggers
    });
  } catch (e) {
    console.warn("Trigger sync skipped", e);
  }
}

/** 実際にWebアプリへPOSTする共通関数 */
function _postToAdmin_(action, payload) {
  try {
    if (!ADMIN_WEBAPP_URL || ADMIN_WEBAPP_URL.indexOf("http") !== 0) return;

    const body = {
      token: ADMIN_WEBAPP_TOKEN,
      action: action,
      payload: payload
    };

    UrlFetchApp.fetch(ADMIN_WEBAPP_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.warn("Admin post failed", e);
  }
}

/** 実行者のメアドを取得 */
function _safeExecutorEmail_() {
  try {
    return Session.getEffectiveUser().getEmail() || "unknown";
  } catch (e) {
    return "unknown";
  }
}

function testAdminLog() {
  const runId = "test-" + Utilities.formatDate(new Date(), "JST", "HHmm");
  
  try {
    logInfoToAdmin_(
      runId, 
      "CONNECTION_TEST", 
      "台帳側からの通信テストです。管理者SSのLogシートにこれが表示されれば成功です！"
    );
    
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert("テスト送信コマンドを送りました。管理者SSを確認してください。");
    }
  } catch (e) {
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert("送信エラー: " + e.message);
    }
  }
}

function debugPost() {
  const url = ADMIN_WEBAPP_URL;
  const payload = {
    token: ADMIN_WEBAPP_TOKEN,
    action: "log",
    payload: { message: "デバッグテスト" }
  };
  
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  console.log("ステータスコード: " + response.getResponseCode());
  console.log("レスポンス内容: " + response.getContentText());
}
