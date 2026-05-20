/**
 * 為主要操作區的每一列建立 Gmail 草稿、貼上「已預約寄送」標籤，
 * 並把預約資訊寫到「已預約區」分頁。實際寄出由排程觸發器在預定時間到了之後執行。
 *
 * 流程：
 *   1. 向使用者詢問 Gmail 草稿主旨（作為模板）與「統一預定寄送時間」（可留空）
 *   2. 讀「主要操作區」每一列做變數替換、組 MIME、Gmail.Users.Drafts.create
 *   3. 把該封 draft 對應的 message 加上「已預約寄送」標籤
 *   4. 在「已預約區」寫一列（包含 draftId、預定時間、建立者）
 *   5. 刪掉「主要操作區」上已處理的那一列
 *
 * 每列預定時間優先序：該列的「預定寄送時間」欄 > 對話框輸入的統一時間。
 * 兩者都沒有 → 該列略過並提示。
 */
function createScheduledDrafts() {
  ensureSheets();

  const subjectLine = Browser.inputBox(
    "建立預約寄送（1/2）",
    "輸入或貼上草稿欄的主旨名稱（包含大括號）",
    Browser.Buttons.OK_CANCEL,
  );
  if (subjectLine === "cancel" || subjectLine === "") return;

  const unifiedTimeStr = Browser.inputBox(
    "建立預約寄送（2/2）",
    "統一預定寄送時間（格式 YYYY-MM-DD HH:mm，留空則使用每列「預定寄送時間」欄）",
    Browser.Buttons.OK_CANCEL,
  );
  if (unifiedTimeStr === "cancel") return;

  const unifiedTime = unifiedTimeStr ? parseDate_(unifiedTimeStr) : null;
  if (unifiedTimeStr && !unifiedTime) {
    Browser.msgBox("時間格式無法解析，請改用 YYYY-MM-DD HH:mm");
    return;
  }

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  const label = getOrCreateLabel_(LABEL_SCHEDULED);
  const owner = getSenderIdentity_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const scheduledSheet = ss.getSheetByName(SHEET_SCHEDULED);

  const dataRange = mainSheet.getDataRange();
  const data = dataRange.getDisplayValues();
  if (data.length < 2) {
    Browser.msgBox("「主要操作區」沒有資料列");
    return;
  }
  const heads = data.shift();
  const recipientIdx = heads.indexOf(RECIPIENT_COL);
  if (recipientIdx === -1) {
    Browser.msgBox(`找不到「${RECIPIENT_COL}」欄`);
    return;
  }

  const rowsToDelete = [];
  const errors = [];

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

      let fromHeader = null;
      if (SENDER_ALIAS || SENDER_NAME) {
        const addr = SENDER_ALIAS || Session.getEffectiveUser().getEmail();
        if (addr) {
          fromHeader = SENDER_NAME
            ? `${encodeRfc2047_(SENDER_NAME)} <${addr}>`
            : addr;
        }
      }

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
      // 用 GmailApp 的 thread.addLabel 比直接呼 Gmail.Users.Messages.modify
      // 簡單，且不需要額外的 OAuth scope。
      const thread = GmailApp.getMessageById(draftResp.message.id).getThread();
      thread.addLabel(label);

      scheduledSheet.appendRow(
        buildRowForSheet_(scheduledSheet, {
          [SCH_COL_RECIPIENT]: row[RECIPIENT_COL],
          [SCH_COL_SUBJECT]: msgObj.subject,
          [SCH_COL_DRAFT_ID]: draftResp.id,
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
  rowsToDelete
    .sort((a, b) => b - a)
    .forEach((r) => mainSheet.deleteRow(r));

  const summary =
    `已建立 ${rowsToDelete.length} 封預約草稿` +
    (errors.length > 0 ? `\n\n錯誤：\n${errors.join("\n")}` : "");
  Browser.msgBox(summary);
}
