/**
 * Creates the menu item "批次寄信" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("批次寄信").addItem("送出", "sendEmails").addToUi();
}
