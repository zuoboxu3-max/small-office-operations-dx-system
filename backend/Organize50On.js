/****************************************************
 * Organize50On.gs（統合版）
 ****************************************************/

/**
 * ✅（移行あり）移行元フォルダ配下の各フォルダを
 * 50音階層（行＞頭文字＞個人名）に「コピー」して整理する
 */
function copyAndOrganizeTo50On_FromProps_() {
  const runId = Utilities.getUuid().slice(0, 8);
  const props = PropertiesService.getScriptProperties();

  const destParentId = (props.getProperty(ADMIN_CONFIG.DEST_ID_KEY) || "").trim();
  const sourceFolderId = (props.getProperty(ADMIN_CONFIG.SRC_ID_KEY) || "").trim();

  if (!destParentId || !sourceFolderId) {
    throw new Error("整理先または移行元のフォルダIDが設定されていません。ウィザードを再実行してください。");
  }

  const destParent = DriveApp.getFolderById(destParentId);
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);

  logInfo_(runId, "ORGANIZE_START", `整理開始: ${sourceFolder.getName()}`, { destParentId, sourceFolderId });

  const start = Date.now();
  const LIMIT_MS = 5.5 * 60 * 1000;

  const targetFolders = sourceFolder.getFolders();
  while (targetFolders.hasNext()) {
    if (Date.now() - start > LIMIT_MS) {
      logInfo_(runId, "ORGANIZE_PARTIAL", "タイムアウト回避のため中断しました。再実行で続けてください。", {});
      break;
    }

    const userFolder = targetFolders.next();
    const userName = (userFolder.getName() || "").trim();
    if (!userName) continue;

    const kanaHead = getHiraganaHeadFromKanji_(userName); // ひらがな1文字
    const gyoName = getGyoName(kanaHead);

    if (!gyoName) {
      logInfo_(runId, "ORGANIZE_SKIP", `行判定不可のためスキップ: ${userName}`, { kanaHead });
      continue;
    }

    const gyoFolder = getOrCreateSubFolder(destParent, gyoName);

    const vowel = getVowel_(kanaHead);
    if (!vowel) {
  logInfo_(runId, "ORGANIZE_SKIP", `母音判定不可のためスキップ: ${userName}`, { kanaHead });
  continue;
}

    const vowelFolder = getOrCreateSubFolder(gyoFolder, vowel);
    const kanaFolder  = getOrCreateSubFolder(vowelFolder, kanaHead);


    if (kanaFolder.getFoldersByName(userName).hasNext()) {
      logInfo_(runId, "ORGANIZE_SKIP", `既に存在: ${userName}`, {});
      continue;
    }

    const newUserFolder = kanaFolder.createFolder(userName);
    copyFolderContents(userFolder, newUserFolder);

    logInfo_(runId, "ORGANIZE_SUCCESS", `コピー完了: ${userName}`, { kanaHead, gyoName });
  }

  logInfo_(runId, "ORGANIZE_END", "整理処理が完了しました", {});
}

/**
 * ✅（新規運用）50音の“箱（行フォルダ）”だけ作る
 * 必要ならこの関数内で「頭文字フォルダ」まで作る拡張が可能
 */
function ensure50OnSkeleton_FromProps_() {
  const props = PropertiesService.getScriptProperties();
  const destParentId = (props.getProperty(ADMIN_CONFIG.DEST_ID_KEY) || "").trim();
  if (!destParentId) throw new Error("整理先IDがありません。");

  ensure50OnSkeleton_(destParentId);
}

/**
 * 行フォルダ → 母音 → かな の3階層で「箱」を作成
 * ※や行/わ行など欠ける段は、母音フォルダは作るが中のかなは一部のみ作成
 */
function ensure50OnSkeleton_(destParentId) {
  const parent = DriveApp.getFolderById(destParentId);

  const gyos = [
    { name: "01_あ行", base: { "あ":"あ","い":"い","う":"う","え":"え","お":"お" } },
    { name: "02_か行", base: { "あ":"か","い":"き","う":"く","え":"け","お":"こ" } },
    { name: "03_さ行", base: { "あ":"さ","い":"し","う":"す","え":"せ","お":"そ" } },
    { name: "04_た行", base: { "あ":"た","い":"ち","う":"つ","え":"て","お":"と" } },
    { name: "05_な行", base: { "あ":"な","い":"に","う":"ぬ","え":"ね","お":"の" } },
    { name: "06_は行", base: { "あ":"は","い":"ひ","う":"ふ","え":"へ","お":"ほ" } },
    { name: "07_ま行", base: { "あ":"ま","い":"み","う":"む","え":"め","お":"も" } },
    // 欠ける段あり：い/えなし
    { name: "08_や行", base: { "あ":"や","う":"ゆ","お":"よ" } },
    { name: "09_ら行", base: { "あ":"ら","い":"り","う":"る","え":"れ","お":"ろ" } },
    // 欠ける段あり：うなし（ゐゑは古仮名。必要なら残す/消す選択可能）
    { name: "10_わ行", base: { "あ":"わ","い":"ゐ","え":"ゑ","お":"を" } }
  ];

  const vowels = ["あ","い","う","え","お"];

  gyos.forEach(g => {
    const gyoFolder = getOrCreateSubFolder(parent, g.name);

    vowels.forEach(v => {
      const vowelFolder = getOrCreateSubFolder(gyoFolder, v);

      // 3階層目（かな）を必要分だけ作る
      const kana = g.base[v];
      if (kana) {
        getOrCreateSubFolder(vowelFolder, kana);
      }
    });
  });
}

/** フォルダの中身を再帰的にコピー */
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

/** フォルダの取得・作成（なければ作る） */
function getOrCreateSubFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

/**
 * 漢字→ひらがな先頭（同時実行対策：Lock + _KanaWork）
 */
function getHiraganaHeadFromKanji_(text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("_KanaWork") || ss.insertSheet("_KanaWork");
    try { sh.hideSheet(); } catch (e) {}

    sh.getRange("A1").setValue(text);
    sh.getRange("B1").setFormula(
      '=IF(A1="","",REGEXREPLACE(HIRAGANA(PHONETIC(A1)),"[^ぁ-んー]",""))'
    );
    SpreadsheetApp.flush();

    const hira = String(sh.getRange("B1").getDisplayValue() || "").trim();
    sh.getRange("A1:B1").clearContent();

    // 取得できない場合は空（=スキップ側に流す）
    return hira ? hira.charAt(0) : "";
  } finally {
    lock.releaseLock();
  }
}

/** カタカナ→ひらがな */
function _kataToHira_(s) {
  return String(s).replace(/[ァ-ン]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0x60));
}

/** 行判定（ひらがな・カタカナ対応） */
function getGyoName(char) {
  const kana = _kataToHira_(char);

  if (/[あ-お]/.test(kana)) return "01_あ行";
  if (/[か-ご]/.test(kana)) return "02_か行";
  if (/[さ-ぞ]/.test(kana)) return "03_さ行";
  if (/[た-ど]/.test(kana)) return "04_た行";
  if (/[な-の]/.test(kana)) return "05_な行";
  if (/[は-ぼぱ-ぽ]/.test(kana)) return "06_は行";
  if (/[ま-も]/.test(kana)) return "07_ま行";
  if (/[や-よ]/.test(kana)) return "08_や行";
  if (/[ら-ろ]/.test(kana)) return "09_ら行";
  if (/[わ-ん]/.test(kana)) return "10_わ行";

  return null;
}
/**
 * ひらがな1文字から母音（あ/い/う/え/お）を返す
 * 判定できない場合は "" を返す
 */
function getVowel_(kana1) {
  const k = _kataToHira_(kana1);

  if (!k) return "";

  // あ段
  if (/[ぁあかがさざたただなはばぱまゃやらわゎ]/.test(k)) return "あ";
  // い段
  if (/[ぃいきぎしじちぢにひびぴみりゐ]/.test(k)) return "い";
  // う段
  if (/[ぅうくぐすずつづぬふぶぷむゅゆる]/.test(k)) return "う";
  // え段
  if (/[ぇえけげせぜてでねへべぺめれゑ]/.test(k)) return "え";
  // お段
  if (/[ぉおこごそぞとどのほぼぽもょよろを]/.test(k)) return "お";

  // ん、記号など
  return "";
}
