function checkAliases() {
  const aliases = GmailApp.getAliases();
  Logger.log("可用別名：" + JSON.stringify(aliases));
  Logger.log("預設寄件地址：" + Session.getActiveUser().getEmail());
}