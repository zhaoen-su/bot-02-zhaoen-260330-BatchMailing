/**
 * Change these to match the column names you are using for email
 * recipient addresses and email sent column.
 */
const RECIPIENT_COL = "收件者";
const EMAIL_SENT_COL = "已寄出";

// 以別名身份寄出；留空字串代表用預設帳號寄出。
// 必須是 GmailApp.getAliases() 回傳清單內的地址（執行 checkAliases 驗證）。
const SENDER_ALIAS = "";
const SENDER_NAME = "";
