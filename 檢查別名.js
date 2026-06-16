function checkAliases() {
  const aliases = GmailApp.getAliases();
  Logger.log("登入帳號(getActiveUser)：" + Session.getActiveUser().getEmail());
  Logger.log("執行身份(getEffectiveUser)：" + Session.getEffectiveUser().getEmail());
  Logger.log("SENDER_ALIAS 設定值：" + JSON.stringify(SENDER_ALIAS));
  Logger.log("Gmail 已驗證別名清單：" + JSON.stringify(aliases));
}
