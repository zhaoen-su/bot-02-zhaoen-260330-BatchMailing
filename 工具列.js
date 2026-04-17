/**
 * Creates the menu item "批次寄信" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("批次寄信")
    .addItem("送出", "sendEmails")
    .addItem("送出（自訂 HTML 模板）", "sendEmailsHTML")
    .addToUi();
}
