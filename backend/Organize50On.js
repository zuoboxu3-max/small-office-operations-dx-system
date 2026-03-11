/****************************************************
 * Organize50On.gs（公開版・抽象化版）
 ****************************************************/

/**
 * 命名規則に基づいて、移行元フォルダ配下の各フォルダを
 * 階層構造へ「コピー」して整理する
 */
function copyAndOrganizeByRule_FromProps_() {
  const runId = Utilities.getUuid().slice(0, 8);
  const props = PropertiesService.getScriptProperties();

  const destParentId = (props.getProperty(ADMIN_CONFIG.DEST_ID_KEY) || "").trim();
  const sourceFolderId = (props.getProperty(ADMIN_CONFIG.SRC_ID_KEY) || "").trim();

  if (!destParentId || !sourceFolderId) {
    throw new Error("整理先または移行元のフォルダIDが設定されていません。設定を再確認してください。");
  }

  const destParent = DriveApp.getFolderById(destParentId);
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);

  logInfo_(runId, "ORGANIZE_START", `整理開始: ${sourceFolder.getName()}`, {
    destParentId,
    sourceFolderId
  });

  const start = Date.now();
  const LIMIT_MS = 5.5 * 60 * 1000;

  const targetFolders = sourceFolder.getFolders();
  while (targetFolders.hasNext()) {
    if (Date.now() - start > LIMIT_MS) {
      logInfo_(
        runId,
        "ORGANIZE_PARTIAL",
        "タイムアウト回避のため中断しました。再実行で続けてください。",
        {}
      );
      break;
    }

    const sourceItemFolder = targetFolders.next();
    const recordKey = (sourceItemFolder.getName() || "").trim();
    if (!recordKey) continue;

    const categoryKey = getNormalizedHeadKey_(recordKey);
    const groupName = getGroupName_(categoryKey);

    if (!groupName) {
      logInfo_(runId, "ORGANIZE_SKIP", "分類判定不可のためスキップしました", {
        categoryKey
      });
      continue;
    }

    const groupFolder = getOrCreateSubFolder(destParent, groupName);

    const subKey = getSubCategoryKey_(categoryKey);
    if (!subKey) {
      logInfo_(runId, "ORGANIZE_SKIP", "補助分類判定不可のためスキップしました", {
        categoryKey
      });
      continue;
    }

    const subFolder = getOrCreateSubFolder(groupFolder, subKey);
    const detailFolder = getOrCreateSubFolder(subFolder, categoryKey);

    if (detailFolder.getFoldersByName(recordKey).hasNext()) {
      logInfo_(runId, "ORGANIZE_SKIP", "対象フォルダは既に存在します", {
        recordKey
      });
      continue;
    }

    const newRecordFolder = detailFolder.createFolder(recordKey);
    copyFolderContents(sourceItemFolder, newRecordFolder);

    logInfo_(runId, "ORGANIZE_SUCCESS", "対象レコードのコピー完了", {
      categoryKey,
      groupName
    });
  }

  logInfo_(runId, "ORGANIZE_END", "整理処理が完了しました", {});
}

/**
 * 階層構造のひな型フォルダを作成する
 * 必要に応じて配下の補助分類フォルダまで生成する
 */
function ensureFolderSkeleton_FromProps_() {
  const props = PropertiesService.getScriptProperties();
  const destParentId = (props.getProperty(ADMIN_CONFIG.DEST_ID_KEY) || "").trim();

  if (!destParentId) {
    throw new Error("整理先IDがありません。");
  }

  ensureFolderSkeleton_(destParentId);
}

/**
 * グループ → 補助分類 → 詳細分類 の3階層でひな型を作成
 */
function ensureFolderSkeleton_(destParentId) {
  const parent = DriveApp.getFolderById(destParentId);

  const groups = [
    { name: "01_グループA", base: { "あ": "あ", "い": "い", "う": "う", "え": "え", "お": "お" } },
    { name: "02_グループB", base: { "あ": "か", "い": "き", "う": "く", "え": "け", "お": "こ" } },
    { name: "03_グループC", base: { "あ": "さ", "い": "し", "う": "す", "え": "せ", "お": "そ" } },
    { name: "04_グループD", base: { "あ": "た", "い": "ち", "う": "つ", "え": "て", "お": "と" } },
    { name: "05_グループE", base: { "あ": "な", "い": "に", "う": "ぬ", "え": "ね", "お": "の" } },
    { name: "06_グループF", base: { "あ": "は", "い": "ひ", "う": "ふ", "え": "へ", "お": "ほ" } },
    { name: "07_グループG", base: { "あ": "ま", "い": "み", "う": "む", "え": "め", "お": "も" } },
    { name: "08_グループH", base: { "あ": "や", "う": "ゆ", "お": "よ" } },
    { name: "09_グループI", base: { "あ": "ら", "い": "り", "う": "る", "え": "れ", "お": "ろ" } },
    { name: "10_グループJ", base: { "あ": "わ", "い": "ゐ", "え": "ゑ", "お": "を" } }
  ];

  const subKeys = ["あ", "い", "う", "え", "お"];

  groups.forEach(g => {
    const groupFolder = getOrCreateSubFolder(parent, g.name);

    subKeys.forEach(subKey => {
      const subFolder = getOrCreateSubFolder(groupFolder, subKey);
      const detailKey = g.base[subKey];

      if (detailKey) {
        getOrCreateSubFolder(subFolder, detailKey);
      }
    });
  });
}

/** フォルダ配下の内容を再帰的にコピー */
function copyFolderContents(source, destination) {
  const files = source.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.makeCopy(file.getName(), destination);
  }

  const subFolders = source.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    const newSubFolder = destination.createFolder(subFolder.getName());
    copyFolderContents(subFolder, newSubFolder);
  }
}

/** フォルダの取得・作成 */
function getOrCreateSubFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

/**
 * 文字列から先頭の分類用キーを取得
 * 同時実行対策として Lock + 作業シートを使用
 */
function getNormalizedHeadKey_(text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("_KanaWork") || ss.insertSheet("_KanaWork");
    try {
      sh.hideSheet();
    } catch (e) {}

    sh.getRange("A1").setValue(text);
    sh.getRange("B1").setFormula(
      '=IF(A1="","",REGEXREPLACE(HIRAGANA(PHONETIC(A1)),"[^ぁ-んー]",""))'
    );
    SpreadsheetApp.flush();

    const normalized = String(sh.getRange("B1").getDisplayValue() || "").trim();
    sh.getRange("A1:B1").clearContent();

    return normalized ? normalized.charAt(0) : "";
  } finally {
    lock.releaseLock();
  }
}

/** カタカナ → ひらがな */
function _kataToHira_(s) {
  return String(s).replace(/[ァ-ン]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) - 0x60)
  );
}

/** 分類キーからグループ名を返す */
function getGroupName_(char) {
  const kana = _kataToHira_(char);

  if (/[あ-お]/.test(kana)) return "01_グループA";
  if (/[か-ご]/.test(kana)) return "02_グループB";
  if (/[さ-ぞ]/.test(kana)) return "03_グループC";
  if (/[た-ど]/.test(kana)) return "04_グループD";
  if (/[な-の]/.test(kana)) return "05_グループE";
  if (/[は-ぼぱ-ぽ]/.test(kana)) return "06_グループF";
  if (/[ま-も]/.test(kana)) return "07_グループG";
  if (/[や-よ]/.test(kana)) return "08_グループH";
  if (/[ら-ろ]/.test(kana)) return "09_グループI";
  if (/[わ-ん]/.test(kana)) return "10_グループJ";

  return null;
}

/**
 * 分類キーから補助分類キーを返す
 * 判定できない場合は "" を返す
 */
function getSubCategoryKey_(char) {
  const k = _kataToHira_(char);

  if (!k) return "";

  if (/[ぁあかがさざたただなはばぱまゃやらわゎ]/.test(k)) return "あ";
  if (/[ぃいきぎしじちぢにひびぴみりゐ]/.test(k)) return "い";
  if (/[ぅうくぐすずつづぬふぶぷむゅゆる]/.test(k)) return "う";
  if (/[ぇえけげせぜてでねへべぺめれゑ]/.test(k)) return "え";
  if (/[ぉおこごそぞとどのほぼぽもょよろを]/.test(k)) return "お";

  return "";
}
