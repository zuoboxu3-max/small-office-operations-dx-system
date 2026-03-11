/****************************************************
 * AdminSetup.gs（統合版）
 ****************************************************/

// スクリプトプロパティのキーを一括管理
const ADMIN_CONFIG = {
  URL_KEY: "ADMIN_WEBAPP_URL",
  TOKEN_KEY: "ADMIN_WEBAPP_TOKEN",
  DEST_ID_KEY: "ORGANIZE_DEST_PARENT_ID",
  SRC_ID_KEY: "ORGANIZE_SOURCE_FOLDER_ID",
  SIZE_ROOT_KEY: "SIZECHECK_ROOT_FOLDER_ID",
  WIZARD_RUN_KEY: "ADMIN_WIZARD_LAST_RUN"
};

// Logシートの正式ヘッダ（全機能で統一）
const LOG_HEADER = ["日時","種類","項目","内容","実行ID","実行者","詳細データ"];
const TRIGGER_HEADER = ["日時","スクリプトID","処理名","種別","補足"];
const SIZE_HEADER = ["日時","フォルダID","フォルダ名","容量（byte）","容量（見やすく）","実行ID","実行者"];

/**
 * onOpen：管理者メニュー
 */
function onOpen(e) {
  try {
    buildAdminMenu_();
  } catch (err) {
    Logger.log("onOpen failed: " + String(err && err.stack ? err.stack : err));
  }
}

function buildAdminMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("⚙ 管理者ツール")
    .addItem("✅ 初期導入ウィザード（1クリック）", "adminBootstrapWizard")
    .addSeparator()
    .addItem("🔍 設定確認（Debug）", "debugAdminWebAppConfig_")
    .addSeparator()
    .addItem("容量チェック（テスト実行）", "testFolderSizeCheckNow")
    .addItem("月末チェック用トリガー作成", "createDailyMonthEndSizeCheckTrigger")
    .addToUi();
}

/**
 * 初期シート作成とトークン発行（Wizardから呼ぶ）
 */
function adminSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSheet_(ss, "Log", LOG_HEADER);
  _ensureSheet_(ss, "TriggerStatus", TRIGGER_HEADER);
  _ensureSheet_(ss, "FolderSizeMonthly", SIZE_HEADER);

  const props = PropertiesService.getScriptProperties();
  let token = (props.getProperty(ADMIN_CONFIG.TOKEN_KEY) || "").trim();
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, "");
    props.setProperty(ADMIN_CONFIG.TOKEN_KEY, token);
  }

  ss.toast("初期設定が完了しました。");
  return token;
}

/**
 * シートの存在確認とヘッダー設定
 */
function _ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && header) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sh;
}

/**
 * プロパティ優先で設定取得
 */
function _getAdminWebAppConfig_() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty(ADMIN_CONFIG.URL_KEY) || "";
  const token = props.getProperty(ADMIN_CONFIG.TOKEN_KEY) || "";
  return { url: String(url || "").trim(), token: String(token || "").trim() };
}

function debugAdminWebAppConfig_() {
  const cfg = _getAdminWebAppConfig_();
  Logger.log("ADMIN_WEBAPP_URL=" + cfg.url);
  Logger.log("ADMIN_WEBAPP_TOKEN exists? " + (!!cfg.token));
  if (cfg.token) Logger.log("TOKEN head=" + String(cfg.token).slice(0, 6) + "...");
}

/**
 * ローカルLogシートへ追記（正式フォーマット）
 */
function appendLocalLog_(level, item, content, runId, extra) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Log") || _ensureSheet_(ss, "Log", LOG_HEADER);

    const now = new Date();
    const executor = "masked_user";
    const extraStr = extra ? JSON.stringify(extra) : "";

    // ["日時","種類","項目","内容","実行ID","実行者","詳細データ"]
    sh.appendRow([now, level, item, content, runId || "", executor, extraStr]);
  } catch (e) {
    Logger.log("Local log failed: " + e.message);
  }
}

/**
 * 管理者WebアプリへPOST（doPost受信口に投げる）
 * ※URL未設定時は何もしない（呼び出し側でローカルに落とす）
 */
function _postToAdmin_(action, payload) {
  const cfg = _getAdminWebAppConfig_();
  if (!cfg.url || !cfg.url.startsWith("http") || !cfg.token) return false;

  try {
    // サイズ制限（簡易ガード）
    const safePayload = JSON.parse(JSON.stringify(payload || {}));
    if (safePayload.message) safePayload.message = String(safePayload.message).slice(0, 8000);

    const body = { token: cfg.token, action: action, payload: safePayload };

    const res = UrlFetchApp.fetch(cfg.url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (res.getResponseCode() >= 400) {
      Logger.log("[AdminPost Error] " + res.getContentText());
      return false;
    }
    return true;
  } catch (e) {
    Logger.log("Admin log post failed: " + e.message);
    return false;
  }
}

/**
 * INFOログ（優先：WebApp送信、失敗/未設定ならローカル）
 */
function logInfo_(runId, action, message, context) {
  const payload = {
    timestamp: new Date().toISOString(),
    level: "INFO",
    runId: runId || "",
    action: action || "",
    message: message || "",
    executor: _safeExecutorEmail_(),
    context: context || null
  };
  const ok = _postToAdmin_("log", payload);
  if (!ok) appendLocalLog_("INFO", action, message, runId, context || null);
}

/**
 * ERRORログ（優先：WebApp送信、失敗/未設定ならローカル）
 */
function logError_(runId, action, err, context) {
  const msg = (err && err.stack) ? err.stack : String(err);
  const payload = {
    timestamp: new Date().toISOString(),
    level: "ERROR",
    runId: runId || "",
    action: action || "",
    message: msg,
    executor: _safeExecutorEmail_(),
    context: context || null
  };
  const ok = _postToAdmin_("log", payload);
  if (!ok) appendLocalLog_("ERROR", action, msg, runId, context || null);
}

/**
 * Trigger状態をWebAppへ送信（受信側でTriggerStatusに反映）
 */
function postTriggerStatusToAdmin_() {
  try {
    const triggers = ScriptApp.getProjectTriggers().map(t => ({
      handler: t.getHandlerFunction(),
      type: String(t.getEventType())
    }));
    _postToAdmin_("trigger_status", { scriptId: ScriptApp.getScriptId(), triggers });
  } catch (e) {
    Logger.log("status post failed: " + String(e && e.stack ? e.stack : e));
  }
}

/**
 * 実行ユーザー（取れなければ unknown）
 */
function _safeExecutorEmail_() {
  return "masked_user";
}

/**
 * Drive URL/ID から ID抽出（失敗時 null）
 */
function extractDriveId_(text) {
  const m = String(text || "").match(/[-\w]{25,}/);
  return m ? m[0] : null;
}
