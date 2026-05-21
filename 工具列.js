/**
 * 試算表開啟時建立「批次寄信」選單。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("預約寄信")
    .addItem("建立預約寄信", "createScheduledDrafts")
    .addItem("現在寄出", "sendEmails")
    .addToUi();
}
