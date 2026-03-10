//設定//
const MASTER_SHEET_NAME = "台帳";
const PARENT_FOLDER_ID = 'YOUR_PARENT_FOLDER_ID';

const TRIGGER_HEADER = "ドライブリンク"; 
const DONE_HEADER = "転記済み"; 
const DONE_AT_HEADER = "転記日時"; 

const CAL_TRIGGER_HEADER = "カレンダー登録"; 
const CAL_DONE_HEADER = "カレンダー済み"; 
const CALENDAR_ID = 'your-calendar@example.com';

// 転記先のテンプレートスプレッドシートIDリスト
const TARGET_SPREADSHEET_IDS = [
  'YOUR_TEMPLATE_SPREADSHEET_ID_1',
  'YOUR_TEMPLATE_SPREADSHEET_ID_2'
  'YOUR_TEMPLATE_SPREADSHEET_ID_3'
  'YOUR_TEMPLATE_SPREADSHEET_ID_4'
];

// 五十音フォルダ分けマップ
const gojuonMap = {
  '01_あ行': ['あ', 'い', 'う', 'え', 'お', 'ぁ', 'ぃ', 'ぅ', 'ぇ', 'ぉ'],
  '02_か行': ['か', 'き', 'く', 'け', 'こ', 'が', 'ぎ', 'ぐ', 'げ', 'ご'],
  '03_さ行': ['さ', 'し', 'す', 'せ', 'そ', 'ざ', 'じ', 'ず', 'ぜ', 'ぞ'],
  '04_た行': ['た', 'ち', 'つ', 'て', 'と', 'だ', 'ぢ', 'づ', 'で', 'ど'],
  '05_な行': ['な', 'に', 'ぬ', 'ね', 'の'],
  '06_は行': ['は', 'ひ', 'ふ', 'へ', 'ほ', 'ば', 'び', 'ぶ', 'べ', 'ぼ', 'ぱ', 'ぴ', 'ぷ', 'ぺ', 'ぽ'],
  '07_ま行': ['ま', 'み', 'む', 'め', 'も'],
  '08_や行': ['や', 'ゆ', 'よ', 'ゃ', 'ゅ', 'ょ'],
  '09_ら行': ['ら', 'り', 'る', 'れ', 'ろ'],
  '10_わ行': ['わ', 'を', 'ん']
};

// 転記項目のマッピング設定
const FIELD_MAPPINGS = [
  { label: "利用者ID", altLabels: ["利用者ID","ID"], keys: ["利用者ID"], offset: 1 },
  { label: "氏名", altLabels: ["利用者氏名","氏名","児童氏名"], keys: ["氏名（必須）"], offset: 1 },
  { label: "性別", altLabels: ["性別"], keys: ["性別"], offset: 1 },
  { label: "年齢", altLabels: ["年齢"], keys: ["年齢"], offset: 1 },
  { label: "生年月日", altLabels: ["生年月日"], keys: ["生年月日"], offset: 1 },
  { label: "支援区分", altLabels: ["支援区分","障害支援区分"], keys: ["支援区分"], offset: 1 },
  { label: "受給者証番号", altLabels: ["受給者証番号"], keys: ["障害福祉サービス受給者証番号"], offset: 1 },
  { label: "上限額", altLabels: ["利用者負担上限額","負担上限額"], keys: ["利用者負担上限額"], offset: 1 },
  { label: "相談支援事業者", altLabels: ["相談支援事業者名"], keys: ["相談支援事業者名"], offset: 1 },
  { label: "担当者", altLabels: ["計画作成担当者"], keys: ["計画作成担当者"], offset: 1 },
  { label: "計画作成日", altLabels: ["計画作成日"], keys: ["計画作成日"], offset: 1 },
  { label: "モニタリング日", altLabels: ["モニタリング実施日"], keys: ["モニタリング実施日"], offset: 1 },
  { label: "住所", altLabels: ["住所"], keys: ["住所"], offset: 1 },
  { label: "電話", altLabels: ["電話番号","電話"], keys: ["電話"], offset: 1 }
];

// =================================================================
// 🛠️ UI・共通実行エンジン
// =================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙ 業務一括処理 ⚙')
    .addItem('チェックした行をカレンダー＆転記実行', 'processAllSelected')
    .addItem('カレンダー登録のみ実行', 'processOnlyCalendar')
    .addSeparator()
    .addItem('【手動】チェックボックスをすべて解除', 'resetAllCheckboxes_')
    .addToUi();
}

/**
 * チェックボックスのリセット（利便性向上）
 */
function resetAllCheckboxes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const colMap = getHeaderToColMap_(sheet);
  
  const targets = [TRIGGER_HEADER, CAL_TRIGGER_HEADER];
  targets.forEach(header => {
    const colIndex = colMap[header];
    if (colIndex) {
      sheet.getRange(2, colIndex, lastRow - 1, 1).uncheck();
    }
  });
  ss.toast("準備完了です。", "システム", 3);
}

function processAllSelected() { executeBatch_({ doDrive: true, doCalendar: true }); }
function processOnlyCalendar() { executeBatch_({ doDrive: false, doCalendar: true }); }

/**
 * 一括処理のメインエンジン
 */
/**
 * 一括処理のメインエンジン（ログ送信機能付き）
 */
function executeBatch_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) { ss.toast("シート「" + MASTER_SHEET_NAME + "」が見つかりません。"); return; }
  
  const runId = Utilities.getUuid(); // 今回の実行を一意に特定するID
  logInfoToAdmin_(runId, "BATCH_START", "一括処理を開始しました", { options });

  const data = sheet.getDataRange().getValues();
  const colMap = getHeaderToColMap_(sheet);

  let targetRows = [];
  for (let i = 1; i < data.length; i++) {
    const isDriveCheck = options.doDrive && data[i][colMap[TRIGGER_HEADER]-1] === true;
    const isCalCheck = options.doCalendar && data[i][colMap[CAL_TRIGGER_HEADER]-1] === true;
    if (isDriveCheck || isCalCheck) {
      targetRows.push({rowNum: i + 1, doDrive: isDriveCheck, doCal: isCalCheck});
    }
  }

  if (targetRows.length === 0) { 
    ss.toast("対象が選択されていません。", "中断"); 
    logInfoToAdmin_(runId, "BATCH_CANCEL", "対象選択なしで終了");
    return; 
  }

  targetRows.forEach((item, index) => {
    const rowNum = item.rowNum;
    try {
      const rowObj = getRowObjectByHeader_(sheet, rowNum);
      const userName = rowObj["氏名（必須）"] || "不明";

      // 1. カレンダー登録
      if (item.doCal) {
        createCalendarEventsInternal_(rowNum, rowObj);
        sheet.getRange(rowNum, colMap[CAL_TRIGGER_HEADER]).setValue(false);
        logInfoToAdmin_(runId, "CALENDAR_SUCCESS", `${userName}様のカレンダー登録完了`, { rowNum });
      }

      // 2. ドライブ転記
      if (item.doDrive) {
        pushAllTargets(rowNum);
        sheet.getRange(rowNum, colMap[TRIGGER_HEADER]).setValue(false);
        logInfoToAdmin_(runId, "DRIVE_SUCCESS", `${userName}様のドライブ転記完了`, { rowNum });
      }
      
      SpreadsheetApp.flush();
      ss.toast(`${targetRows.length}件中 ${index + 1}件完了`, "進捗");

    } catch (e) {
      const errorCol = colMap[DONE_HEADER] || 1;
      sheet.getRange(rowNum, errorCol).setValue("❌エラー: " + e.message).setBackground("#f4cccc");
      
      // 管理者へ詳細なエラーを送信
      logErrorToAdmin_(runId, "ROW_FATAL", e, { rowNum, options });
    }
  });

  logInfoToAdmin_(runId, "BATCH_END", "一括処理を正常に終了しました");
  postTriggerStatusToAdmin_(); // 最後にトリガー状態を管理者に同期
  ss.toast("全工程が終了しました。");
}

// =================================================================
// ⚙️ 転記メインロジック
// =================================================================

function pushAllTargets(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const colMap = getHeaderToColMap_(sheet);
  const rowObj = getRowObjectByHeader_(sheet, row);
  
  // 1. 氏名と頭文字をrowObjから確実に取得
  const userName = String(rowObj["氏名（必須）"] || "").trim();
  const initial = String(rowObj["頭文字（必須）"] || "").trim();
  
  if (!userName || userName === "undefined") throw new Error("氏名が空です");
  if (!initial) throw new Error("頭文字が入力されていません（ア、カ、サ等）");

  // 2. フォルダの取得（頭文字フォルダ > 個人名フォルダ）
  const userFolder = getTargetFolderByHierarchy_(userName, initial);

  TARGET_SPREADSHEET_IDS.forEach(id => {
    const templateFile = DriveApp.getFileById(id);
    const targetFileName = `${templateFile.getName()}_${userName}`;
    
    // 同名ファイルがあるか確認
    const existingFiles = userFolder.getFilesByName(targetFileName);
    let targetSSFile = existingFiles.hasNext() ? existingFiles.next() : templateFile.makeCopy(targetFileName, userFolder);
    let targetSS = SpreadsheetApp.openById(targetSSFile.getId());
    
    // 3. コピー元シートを固定（1枚目のシートが原本という前提、または名前指定）
    // 常に「原本」という名前のシートをコピーするようにするとミスがありません
    const templateSheet = targetSS.getSheets().find(s => s.getName().includes("原本")) || targetSS.getSheets()[0];
    
    const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmm");
    const newSheetName = `${userName}_${nowStr}`; // シート名はシンプルに
    
    // シートコピーと転記実行
    const newSheet = templateSheet.copyTo(targetSS).setName(newSheetName);
    applyMappingsByLabel_(newSheet, rowObj, FIELD_MAPPINGS);

    // 4. 新しいシートを一番左（先頭）へ移動し、アクティブにする
    newSheet.activate();
    targetSS.moveActiveSheet(1); 
  });

  // 完了記録
  sheet.getRange(row, colMap[DONE_HEADER]).setValue("✅完了(履歴保存)").setBackground("#d9ead3");
  if (colMap[DONE_AT_HEADER]) sheet.getRange(row, colMap[DONE_AT_HEADER]).setValue(new Date());
}

// =================================================================
// 📅 カレンダー登録
// =================================================================

function createCalendarEventsInternal_(rowNum, rowObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const colMap = getHeaderToColMap_(sheet);
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) throw new Error("カレンダーIDが見つかりません: " + CALENDAR_ID);

  const userName = rowObj["氏名（必須）"] || "不明";
  
  const dateFields = [
    { name: "モニタリング", val: rowObj["モニタリング日"] },
    { name: "計画作成", val: rowObj["計画作成日"] }
  ];
  
  dateFields.forEach(field => {
    let d = field.val;
    if (d && !(d instanceof Date)) d = new Date(d);
    
    if (d instanceof Date && !isNaN(d.getTime())) {
      const title = `【${field.name}】${userName}様`;
      // 同日の同一タイトル重複チェック
      const exist = calendar.getEventsForDay(d, {search: title});
      if (exist.length === 0) {
        calendar.createAllDayEvent(title, d);
      }
    }
  });

  const statusCol = colMap[CAL_DONE_HEADER];
  if (statusCol) {
    sheet.getRange(rowNum, statusCol).setValue("✅完了").setBackground("#d9ead3");
  }
}

// =================================================================
// 🛠 ヘルパー関数群
// =================================================================

function getTargetFolderByHierarchy_(userName, firstChar) {
  const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  let gyoName = "11_その他";
  
  for (let key in gojuonMap) {
    if (gojuonMap[key].includes(firstChar)) { gyoName = key; break; }
  }
  
  const gyoFolder = getOrCreateSubFolder_(parentFolder, gyoName);
  const charFolder = getOrCreateSubFolder_(gyoFolder, firstChar);
  return getOrCreateSubFolder_(charFolder, userName);
}

function getOrCreateSubFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function applyMappingsByLabel_(sheet, data, mappings) {
  const cache = new Map();
  mappings.forEach(m => {
    const val = normalizeValue_(pickFirst_(data, m.keys));
    if (val === "") return;

    for (const label of m.altLabels) {
      const cell = findLabelCell_(sheet, label, cache);
      if (cell) {
        let targetCol = cell.getColumn();
        const merges = cell.getMergedRanges();
        if (merges.length > 0) targetCol = merges[0].getLastColumn();
        
        sheet.getRange(cell.getRow(), targetCol + (m.offset || 1)).setValue(val);
        break; 
      }
    }
  });
}

function findLabelCell_(sheet, label, cache) {
  const key = `${sheet.getSheetId()}_${label}`;
  if (cache.has(key)) return cache.get(key);
  
  const cell = sheet.createTextFinder(label).matchEntireCell(false).findNext();
  cache.set(key, cell);
  return cell;
}

function toHiragana_(str) {
  if (!str) return "";
  return str.replace(/[\u30a1-\u30f6]/g, m => String.fromCharCode(m.charCodeAt(0) - 0x60));
}

function getHeaderToColMap_(sheet) {
  const map = {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

function getRowObjectByHeader_(sheet, row) {
  const colMap = getHeaderToColMap_(sheet);
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const obj = {};
  Object.keys(colMap).forEach(h => { obj[h] = values[colMap[h] - 1]; });
  return obj;
}

function pickFirst_(data, keys) {
  for (const k of keys) { 
    if (data[k] !== undefined && data[k] !== "") return data[k]; 
  }
  return "";
}

function normalizeValue_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Tokyo", "yyyy/MM/dd");
  return v !== null && v !== undefined ? String(v).trim() : "";
}

