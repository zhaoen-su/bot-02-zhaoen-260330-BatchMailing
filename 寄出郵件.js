/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 */
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  let processedSubjectLine = subjectLine;
  if (!processedSubjectLine) {
    processedSubjectLine = Browser.inputBox(
      "批次寄信小工具",
      "輸入或貼上草稿欄的主旨名稱（包含大括號）",
      Browser.Buttons.OK_CANCEL,
    );

    if (processedSubjectLine === "cancel" || processedSubjectLine === "") {
      return;
    }
  }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(processedSubjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift();

  // Gets the index of the column named '已寄出'
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  // Converts 2d array into an object array
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => {
      o[k] = r[i] || "";
      return o;
    }, {}),
  );

  const out = [];

  obj.forEach((row) => {
    if (row[EMAIL_SENT_COL] === "") {
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        let fromHeader = null;
        if (SENDER_ALIAS || SENDER_NAME) {
          const addr = SENDER_ALIAS || Session.getActiveUser().getEmail();
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

        Gmail.Users.Messages.send({ raw: base64UrlEncode_(raw) }, "me");
        out.push([new Date()]);
      } catch (e) {
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
}
