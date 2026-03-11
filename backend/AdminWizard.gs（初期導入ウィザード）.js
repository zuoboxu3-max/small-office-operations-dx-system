/****************************************************
 * AdminWizard.gs（統合版）
 ****************************************************/

function adminBootstrapWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const runId = Utilities.getUuid().replace(/-/g, "").slice(0, 12);

  // 1) シート作成・トークン発行
  adminSetup();

  // 2) WebアプリURL（ログ送信先）の設定
  // 未設定なら必須入力、設定済みなら空で維持OK
  const currentUrl = (props.getProperty(ADMIN_CONFIG.URL_KEY) || "").trim();
  const resUrl = ui.prompt(
    "【1/3】管理者WebアプリURLを入力してください",
    currentUrl
      ? "現在： " + currentUrl + "\n\n※空欄でOKすると現状維持"
      : "未設定です。WebアプリURLを入力してください（キャンセルで中止）",
    ui.ButtonSet.OK_CANCEL
  );
  if (resUrl.getSelectedButton() !== ui.Button.OK) return;
  const urlInput = (resUrl.getResponseText() || "").trim();
  if (!currentUrl && !urlInput) {
    ui.alert("WebアプリURLが未設定のため中止します。");
    return;
  }
  if (urlInput) props.setProperty(ADMIN_CONFIG.URL_KEY, urlInput);

  // 3) 保存先（整理後）フォルダ
  let destParentId = (props.getProperty(ADMIN_CONFIG.DEST_ID_KEY) || "").trim();
  const resDest = ui.prompt(
    "【2/3】保存先（整理後）フォルダのURLを入力してください",
    destParentId
      ? "現在： " + destParentId + "\n\n※空欄でOKすると現状維持"
      : "未設定です。保存先フォルダURLを入力してください（キャンセルで中止）",
    ui.ButtonSet.OK_CANCEL
  );
  if (resDest.getSelectedButton() !== ui.Button.OK) return;
  const destInput = (resDest.getResponseText() || "").trim();
  if (!destParentId && !destInput) {
    ui.alert("保存先が未設定のため中止します。");
    return;
  }
  if (destInput) {
    const newId = extractDriveId_(destInput);
    if (!newId) {
      ui.alert("保存先フォルダIDを取得できませんでした。URLを確認してください。");
      return;
    }
    destParentId = newId;
    props.setProperty(ADMIN_CONFIG.DEST_ID_KEY, destParentId);
    // 容量チェック対象も同じ親を使う運用（必要なら後で変更可）
    props.setProperty(ADMIN_CONFIG.SIZE_ROOT_KEY, destParentId);
  }

  // 4) 移行元（整理前）フォルダ：空欄なら箱作りモード
  let sourceFolderId = (props.getProperty(ADMIN_CONFIG.SRC_ID_KEY) || "").trim();
  const resSrc = ui.prompt(
    "【3/3】移行元（整理前）フォルダのURLを入力してください",
    sourceFolderId
      ? "現在： " + sourceFolderId + "\n\n※空欄でOKすると『箱作りモード』"
      : "未設定です。\n※空欄でOKすると『箱作りモード』",
    ui.ButtonSet.OK_CANCEL
  );
  if (resSrc.getSelectedButton() !== ui.Button.OK) return;

  const srcInput = (resSrc.getResponseText() || "").trim();
  if (srcInput) {
    const newSrcId = extractDriveId_(srcInput);
    if (!newSrcId) {
      ui.alert("移行元フォルダIDを取得できませんでした。URLを確認してください。");
      return;
    }
    sourceFolderId = newSrcId;
    props.setProperty(ADMIN_CONFIG.SRC_ID_KEY, sourceFolderId);
  } else {
    // 空欄なら解除して箱作りモード
    props.deleteProperty(ADMIN_CONFIG.SRC_ID_KEY);
    sourceFolderId = "";
  }

  // 5) テスト：受信口へ自己POST（ログ動作確認）
  try {
    postTestLogToSelfWebApp_(runId);
  } catch (e) {
    // ここで止める必要はないが、利用者に伝える
    appendLocalLog_("ERROR", "WIZARD", "自己POSTテストに失敗しました（WebアプリURL/デプロイを確認）", runId, { error: String(e) });
  }

  // 6) 整理実行 or 箱作り
  if (!sourceFolderId) {
    const res = ui.alert("確認", "移行元が未設定です。50音の『箱（あ〜わ行）』だけ作成しますか？", ui.ButtonSet.YES_NO);
    if (res === ui.Button.YES) {
      try {
        ensure50OnSkeleton_FromProps_();
        appendLocalLog_("INFO", "WIZARD", "箱作り（行フォルダ）を作成しました", runId, {});
      } catch (e) {
        appendLocalLog_("ERROR", "WIZARD", "箱作りに失敗しました", runId, { error: String(e) });
        ui.alert("箱作りに失敗しました。ログをご確認ください。");
        return;
      }
    } else {
      ui.alert("処理を終了しました。");
      return;
    }
  } else {
    try {
     ss.toast("階層整理を実行中です...", "実行中", -1);
copyAndOrganizeByRule_FromProps_();
ss.toast("整理が完了しました。", "完了", 5);
appendLocalLog_("INFO", "WIZARD", "階層整理が完了しました", runId, {});
    } catch (e) {
      appendLocalLog_("ERROR", "WIZARD", "50音整理でエラーが発生しました", runId, { error: String(e && e.stack ? e.stack : e) });
      ui.alert("整理中にエラーが発生しました。ログを確認してください。");
      return;
    }
  }

  // 7) 月末容量チェック用トリガー作成
  createDailyMonthEndSizeCheckTrigger();

  // 8) 最終実行記録
  props.setProperty(ADMIN_CONFIG.WIZARD_RUN_KEY, new Date().toISOString());

  ui.alert("✅ セットアップが完了しました！");
}

/**
 * 自分自身のWebアプリへテストログ送信して受信確認
 */
function postTestLogToSelfWebApp_(runId) {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty(ADMIN_CONFIG.URL_KEY);
  const token = props.getProperty(ADMIN_CONFIG.TOKEN_KEY);
  if (!url || !token) throw new Error("ADMIN_WEBAPP_URL または TOKEN が未設定です");

  const payload = {
    timestamp: new Date().toISOString(),
    level: "INFO",
    runId: runId,
    action: "WIZARD_TEST",
    message: "初期導入ウィザードのテストログです（doPost受信確認）",
    executor: _safeExecutorEmail_(),
    context: { note: "self-post test", time: new Date().toString() }
  };

  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ token: token, action: "log", payload: payload }),
    muteHttpExceptions: true
  });
}
