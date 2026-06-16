/**
 * 為主要操作區的每一列建立 Gmail 草稿、貼上「已預約寄送」標籤，
 * 並把預約資訊寫到「已預約區」分頁。實際寄出由排程觸發器在預定時間到了之後執行。
 *
 * UI 入口在「預約寄送對話框.js」(HtmlService)，這裡只負責邏輯，方便測試 / 重用。
 *
 * 流程：
 *   1. 讀「主要操作區」每一列做變數替換、組 MIME、Gmail.Users.Drafts.create
 *   2. 把該封 draft 對應的 message 加上「已預約寄送」標籤
 *   3. 在「已預約區」寫一列（含 draftId 連結、預定時間、建立者）
 *   4. 刪掉「主要操作區」上已處理的那一列
 *
 * 每列預定時間優先序：該列的「預定寄送時間」欄 > 對話框輸入的統一時間。
 * 兩者都沒有 → 該列略過並列入 errors。
 *
 * @param {string} subjectLine        Gmail 草稿主旨
 * @param {string} unifiedTimeStr     對話框輸入的統一時間（"YYYY-MM-DDTHH:mm"），可空
 * @return {{successCount:number, errors:string[]}}
 */
function runCreateScheduledDrafts_(subjectLine, unifiedTimeStr) {
  if (!subjectLine) throw new Error("請輸入草稿主旨");
  assertSenderConfigured_();

  ensureSheets();

  const unifiedTime = unifiedTimeStr ? parseDate_(unifiedTimeStr) : null;
  if (unifiedTimeStr && !unifiedTime) {
    throw new Error("統一預定寄送時間格式無法解析");
  }

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  const label = getOrCreateLabel_(LABEL_SCHEDULED);
  const owner = getSenderIdentity_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const scheduledSheet = ss.getSheetByName(SHEET_SCHEDULED);

  const dataRange = mainSheet.getDataRange();
  const data = dataRange.getDisplayValues();
  if (data.length < 2) throw new Error("「主要操作區」沒有資料列");

  const heads = data.shift();
  if (heads.indexOf(RECIPIENT_COL) === -1) {
    throw new Error(`「主要操作區」找不到「${RECIPIENT_COL}」欄`);
  }

  const rowsToDelete = [];
  const errors = [];
  // 草稿用到、但「主要操作區」沒有對應欄位的變數名（跨所有列取聯集，
  // 只提醒一次，避免每列重複洗版）。
  const missingKeysAll = {};

  data.forEach((rowArr, i) => {
    const row = heads.reduce((o, k, j) => {
      o[k] = rowArr[j] || "";
      return o;
    }, {});
    if (!row[RECIPIENT_COL]) return;

    const rowTime = row[MAIN_COL_SCHEDULED_AT]
      ? parseDate_(row[MAIN_COL_SCHEDULED_AT])
      : null;
    const scheduledAt = rowTime || unifiedTime;
    if (!scheduledAt) {
      errors.push(`第 ${i + 2} 列：沒有預定寄送時間`);
      return;
    }

    try {
      const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
      msgObj.missingKeys.forEach((k) => (missingKeysAll[k] = true));

      // assertSenderConfigured_ 已確保 alias / name 都非空且 alias 被 Gmail 授權，
      // 所以這裡可以直接組「Name <alias>」而不再 fallback。
      const fromHeader = `${encodeRfc2047_(SENDER_NAME)} <${SENDER_ALIAS}>`;

      const raw = buildMimeMessage_({
        from: fromHeader,
        to: row[RECIPIENT_COL],
        cc: emailTemplate.message.cc,
        bcc: emailTemplate.message.bcc,
        subject: msgObj.subject,
        text: msgObj.text,
        html: msgObj.html,
        inlineImages: emailTemplate.inlineImages,
        attachments: emailTemplate.attachments,
      });

      const draftResp = Gmail.Users.Drafts.create(
        { message: { raw: base64UrlEncode_(raw) } },
        "me",
      );

      // 標籤套用在 draft 的底層 message 上（draft 本身沒有 label 概念）。
      const thread = GmailApp.getMessageById(draftResp.message.id).getThread();
      thread.addLabel(label);

      // draftId 欄塞成 HYPERLINK 公式，點擊就跳 Gmail。
      // getValues 讀回來是公式結果（label 字串 = draftId），所以
      // processScheduledDrafts 拿到的還是純 draftId 不受影響。
      const draftLink = `=HYPERLINK("${thread.getPermalink()}","${draftResp.id}")`;

      scheduledSheet.appendRow(
        buildRowForSheet_(scheduledSheet, {
          [SCH_COL_RECIPIENT]: row[RECIPIENT_COL],
          [SCH_COL_SUBJECT]: msgObj.subject,
          [SCH_COL_DRAFT_ID]: draftLink,
          [SCH_COL_SCHEDULED_AT]: scheduledAt,
          [SCH_COL_CREATED_AT]: new Date(),
          [SCH_COL_OWNER]: owner,
        }),
      );

      rowsToDelete.push(i + 2);
    } catch (e) {
      errors.push(`第 ${i + 2} 列：${e.message}`);
    }
  });

  // 由下往上刪，避免行號偏移。
  rowsToDelete.sort((a, b) => b - a).forEach((r) => mainSheet.deleteRow(r));

  const missingKeys = Object.keys(missingKeysAll);
  const warnings = missingKeys.length
    ? [
        `草稿用到這些變數，但「主要操作區」沒有同名欄位（已當空白處理）：` +
          missingKeys.map((k) => `{{${k}}}`).join("、") +
          `。請確認欄位標題與變數名完全相同（含全形 / 半形、空白）。`,
      ]
    : [];

  return { successCount: rowsToDelete.length, errors: errors, warnings: warnings };
}

/**
 * 「現在寄出」：把「已預約區」中由自己建立的草稿全部立刻送出（忽略預定時間）。
 * 觸發器同樣會做這件事但只挑時間到的列；這裡就是手動觸發、不理時間。
 */
function sendEmails() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "現在寄出",
    "把「已預約區」中由自己建立的草稿全部立刻送出（忽略預定時間）。確定要繼續？",
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp !== ui.Button.OK) return;
  processScheduledDrafts(true);
}
