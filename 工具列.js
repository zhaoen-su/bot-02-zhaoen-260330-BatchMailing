/**
 * 試算表開啟時建立「批次寄信」選單。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("批次寄信")
    .addItem("建立預約寄送", "createScheduledDrafts")
    .addSeparator()
    .addItem("啟用排程觸發器", "installScheduledTrigger")
    .addItem("移除排程觸發器", "removeScheduledTrigger")
    .addSeparator()
    .addItem("初始化分頁結構", "ensureSheets")
    .addToUi();
}
