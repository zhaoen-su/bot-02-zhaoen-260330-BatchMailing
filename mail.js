/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 */
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  let processedSubjectLine = subjectLine;
  if (!processedSubjectLine) {
    processedSubjectLine = Browser.inputBox(
      "Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
        "draft message you would like to mail merge with:",
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

        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages,
        });
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
