/**
 * 排程寄送：管理時間觸發器，並由觸發器定時掃描「已預約區」寄出到期的草稿。
 *
 * 為何要這樣設計：
 *   - Gmail API 沒有開放排程寄送欄位（scheduled send 是 Gmail UI 專有功能），
 *     所以必須靠 Apps Script 的時間觸發器在指定時間後檢查並送出。
 *   - 觸發器以「設定者」的身份執行（getEffectiveUser），所以 processScheduledDrafts
 *     只會處理自己當初建立的列（用「建立者」欄做身份比對）。
 *     避免多人共用試算表時把別人的草稿弄丟、或拿著對方 draftId 在自己信箱找不到。
 */

/**
 * 安裝定時觸發器（先移除既有的同名觸發器，避免重複）。
 * 由選單呼叫。
 */
function installScheduledTrigger() {
  removeScheduledTrigger();
  ScriptApp.newTrigger(TRIGGER_HANDLER)
    .timeBased()
    .everyHours(TRIGGER_INTERVAL_HOURS)
    .create();
  Browser.msgBox(
    `已啟用排程觸發器（每 ${TRIGGER_INTERVAL_HOURS} 小時掃描一次）。\n` +
      `執行身份：${Session.getEffectiveUser().getEmail()}`,
  );
}

/**
 * 移除這個專案內所有指向 TRIGGER_HANDLER 的觸發器。
 */
function removeScheduledTrigger() {
  const triggers = ScriptApp.getProjectTriggers().filter(
    (t) => t.getHandlerFunction() === TRIGGER_HANDLER,
  );
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));
  if (triggers.length > 0) {
    SpreadsheetApp.getActive().toast(`已移除 ${triggers.length} 個排程觸發器`);
  }
}

/**
 * 觸發器入口。掃「已預約區」，挑出：
 *   - 建立者欄等於當前執行身份 (避免處理到別人的列)
 *   - 預定寄送時間 ≤ 現在
 * 然後逐列 Gmail.Users.Drafts.send，再把該列搬到「已寄出區」。
 *
 * 失敗也搬到已寄出區並寫錯誤訊息，避免下一輪重試同一封造成重寄。
 */
function processScheduledDrafts() {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduledSheet = ss.getSheetByName(SHEET_SCHEDULED);
  const sentSheet = ss.getSheetByName(SHEET_SENT);

  const lastRow = scheduledSheet.getLastRow();
  if (lastRow < 2) return;

  // getValues 而非 getDisplayValues，這樣日期回來是 Date 物件可直接比對。
  const lastCol = scheduledSheet.getLastColumn();
  const all = scheduledSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const heads = all.shift();
  const idx = (name) => heads.indexOf(name);
  const iRecipient = idx(SCH_COL_RECIPIENT);
  const iSubject = idx(SCH_COL_SUBJECT);
  const iDraftId = idx(SCH_COL_DRAFT_ID);
  const iScheduledAt = idx(SCH_COL_SCHEDULED_AT);
  const iOwner = idx(SCH_COL_OWNER);

  const owner = getSenderIdentity_();
  const now = new Date();
  const rowsToDelete = [];

  all.forEach((rowArr, i) => {
    if (String(rowArr[iOwner]) !== owner) return;

    const scheduledAt = parseDate_(rowArr[iScheduledAt]);
    if (!scheduledAt) return;
    if (scheduledAt.getTime() > now.getTime()) return;

    const draftId = String(rowArr[iDraftId]);
    let result = "成功";
    try {
      Gmail.Users.Drafts.send({ id: draftId }, "me");
    } catch (e) {
      result = `失敗：${e.message}`;
    }

    sentSheet.appendRow(
      buildRowForSheet_(sentSheet, {
        [SENT_COL_RECIPIENT]: rowArr[iRecipient],
        [SENT_COL_SUBJECT]: rowArr[iSubject],
        [SENT_COL_SCHEDULED_AT]: scheduledAt,
        [SENT_COL_SENT_AT]: new Date(),
        [SENT_COL_RESULT]: result,
      }),
    );
    rowsToDelete.push(i + 2);
  });

  rowsToDelete
    .sort((a, b) => b - a)
    .forEach((r) => scheduledSheet.deleteRow(r));
}

/**
 * 取得（或建立）「已預約寄送」標籤。
 */
function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/**
 * 把輸入解析為 Date。接受：
 *   - Date 物件本身（從 sheet getValues 拿到的）
 *   - "YYYY-MM-DD HH:mm" / "YYYY-MM-DD HH:mm:ss" / ISO 8601
 *   - 試算表上以日期格式顯示出來的字串
 *
 * 解析失敗回 null。
 */
function parseDate_(input) {
  if (input instanceof Date) return isNaN(input.getTime()) ? null : input;
  if (input == null || input === "") return null;
  const s = String(input).trim();
  // 把 "YYYY-MM-DD HH:mm" 補成 ISO 友善的 "YYYY-MM-DDTHH:mm" 提高跨平台一致性。
  const iso = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?$/.test(s)
    ? s.replace(" ", "T")
    : s;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? null : d;
}
