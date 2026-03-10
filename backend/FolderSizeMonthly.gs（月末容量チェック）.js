/****************************************************
 * FolderSizeMonthly.gs（統合版）
 ****************************************************/

const SIZECHECK_ROOT_FOLDER_ID_PROP = ADMIN_CONFIG.SIZE_ROOT_KEY;

/**
 * トリガー作成：毎日21時（関数側で月末判定）
 */
function createDailyMonthEndSizeCheckTrigger() {
  const handler = "monthEndFolderSizeCheckAdmin";

  // 既存削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    try {
      if (t.getHandlerFunction && t.getHandlerFunction() === handler) {
        ScriptApp.deleteTrigger(t);
      }
    } catch (e) {}
  });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();

  try { postTriggerStatusToAdmin_(); } catch (e) {}

  SpreadsheetApp.getActive().toast("トリガー作成完了：毎日実行（※月末のみ記録）");
}

/**
 * 月末チェック本体：月末のみ FolderSizeMonthly に記録（当月重複はスキップ）
 */
function monthEndFolderSizeCheckAdmin() {
  const runId = Utilities.getUuid().replace(/-/g, "");
  const executor = _safeExecutorEmail_();

  const props = PropertiesService.getScriptProperties();
  const folderId = (props.getProperty(SIZECHECK_ROOT_FOLDER_ID_PROP) || "").trim();

  if (!folderId) {
    appendLocalLog_("ERROR", "FOLDER_SIZE", "親フォルダIDが未設定です（SIZECHECK_ROOT_FOLDER_ID）", runId, { executor });
    return;
  }

  const now = new Date();
  if (!isLastDayOfMonth_(now)) return;

  try {
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();

    const bytes = calculateFolderSize_(folder);
    const human = formatBytes_(bytes);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("FolderSizeMonthly") || _ensureSheet_(ss, "FolderSizeMonthly", SIZE_HEADER);

    if (hasMonthEntry_(sh, folderId, now)) {
      appendLocalLog_("INFO", "FOLDER_SIZE", "当月分が既に記録されているためスキップしました", runId, { folderId });
      logInfo_(runId, "FOLDER_SIZE", "当月分が既に記録されているためスキップしました", { folderId });
      return;
    }

    // SIZE_HEADER: ["日時","フォルダID","フォルダ名","容量（byte）","容量（見やすく）","実行ID","実行者"]
    sh.appendRow([now, folderId, folderName, bytes, human, runId, executor]);

    appendLocalLog_("INFO", "FOLDER_SIZE", `【月末】${folderName} 合計: ${human}`, runId, { folderId, bytes });
    logInfo_(runId, "FOLDER_SIZE", `【月末】${folderName} 合計: ${human}`, { folderId, bytes });
  } catch (err) {
    appendLocalLog_("ERROR", "FOLDER_SIZE", String(err && err.stack ? err.stack : err), runId, { folderId });
    logError_(runId, "FOLDER_SIZE", err, { folderId });
  }
}

/**
 * 手動テスト：月末でなくてもサイズ記録（動作確認用）
 */
function testFolderSizeCheckNow() {
  const runId = Utilities.getUuid().replace(/-/g, "");
  const executor = _safeExecutorEmail_();

  const props = PropertiesService.getScriptProperties();
  const folderId = (props.getProperty(SIZECHECK_ROOT_FOLDER_ID_PROP) || "").trim();

  if (!folderId) {
    SpreadsheetApp.getActive().toast("親フォルダIDが未設定です。ウィザードで保存先を設定してください。");
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    const bytes = calculateFolderSize_(folder);
    const human = formatBytes_(bytes);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("FolderSizeMonthly") || _ensureSheet_(ss, "FolderSizeMonthly", SIZE_HEADER);

    sh.appendRow([new Date(), folderId, folderName, bytes, human, runId, executor]);

    appendLocalLog_("INFO", "FOLDER_SIZE_TEST", `【テスト】${folderName} 合計: ${human}`, runId, { folderId, bytes });
    SpreadsheetApp.getActive().toast(`【テスト】合計サイズ: ${human}`);
  } catch (err) {
    appendLocalLog_("ERROR", "FOLDER_SIZE_TEST", String(err && err.stack ? err.stack : err), runId, { folderId });
    SpreadsheetApp.getActive().toast("テスト失敗：権限やフォルダIDを確認してください");
  }
}

// ---------- 内部ユーティリティ ----------

function calculateFolderSize_(folder) {
  let size = 0;

  const files = folder.getFiles();
  while (files.hasNext()) size += files.next().getSize();

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) size += calculateFolderSize_(subFolders.next());

  return size;
}

function formatBytes_(bytes) {
  const units = ["B","KB","MB","GB","TB"];
  let i = 0;
  let n = bytes;
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i++;
  }
  return n.toFixed(2) + " " + units[i];
}

function isLastDayOfMonth_(date) {
  const d = new Date(date);
  d.setDate(d.getDate() + 1);
  return d.getDate() === 1;
}

/**
 * 当月分が既に記録済みかチェック
 * A:日時 / B:フォルダID を参照
 */
function hasMonthEntry_(sh, folderId, dateObj) {
  const tz = Session.getScriptTimeZone();
  const monthKey = Utilities.formatDate(dateObj, tz, "yyyy-MM");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const ts = values[i][0];
    const fid = values[i][1];
    if (!ts) continue;

    const mk = Utilities.formatDate(new Date(ts), tz, "yyyy-MM");
    if (mk !== monthKey) {
      if (mk < monthKey) break;
      continue;
    }
    if (String(fid) === String(folderId)) return true;
  }
  return false;
}
