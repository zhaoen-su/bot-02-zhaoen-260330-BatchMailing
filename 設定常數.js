/**
 * 全域設定。修改前請先確認試算表上的欄位名稱與這裡一致。
 */

// 主要操作區 ──────────────────────────────────────────────
const RECIPIENT_COL = "收件者";
// 主要操作區可選欄位：每列指定的預定寄送時間。空白者會 fallback
// 到「建立預約寄送」對話框輸入的統一時間。
const MAIN_COL_SCHEDULED_AT = "預定寄送時間";

// 寄件人 ──────────────────────────────────────────────────
// 以別名身份寄出；留空字串代表用預設帳號寄出。
// 必須是 GmailApp.getAliases() 回傳清單內的地址（執行 checkAliases 驗證）。
const SENDER_ALIAS = "";
const SENDER_NAME = "";

// 分頁名稱 ────────────────────────────────────────────────
const SHEET_MAIN = "主要操作區";
const SHEET_SCHEDULED = "已預約區";
const SHEET_SENT = "已寄出區";

// 「已預約區」欄位 ────────────────────────────────────────
const SCH_COL_RECIPIENT = "收件者";
const SCH_COL_SUBJECT = "主旨";
const SCH_COL_DRAFT_ID = "draftId";
const SCH_COL_SCHEDULED_AT = "預定寄送時間";
const SCH_COL_CREATED_AT = "建立時間";
const SCH_COL_OWNER = "建立者";

// 「已寄出區」欄位 ────────────────────────────────────────
const SENT_COL_RECIPIENT = "收件者";
const SENT_COL_SUBJECT = "主旨";
const SENT_COL_SCHEDULED_AT = "預定寄送時間";
const SENT_COL_SENT_AT = "實際寄出時間";
const SENT_COL_RESULT = "結果";

// Gmail 標籤名稱 ──────────────────────────────────────────
const LABEL_SCHEDULED = "已預約寄送";

// 觸發器 ──────────────────────────────────────────────────
const TRIGGER_HANDLER = "processScheduledDrafts";
const TRIGGER_INTERVAL_HOURS = 1;

/**
 * 寄件人「身份識別字串」，用來在「已預約區」標記列的擁有者。
 * 有設別名就用別名；沒有就退回實際帳號 email。
 *
 * 觸發器執行時，getEffectiveUser() 回的是觸發器擁有者，所以這個值會跟
 * 當初「建立預約寄送」時寫入的值相符 → 觸發器只會處理自己建立的列。
 */
function getSenderIdentity_() {
  return SENDER_ALIAS || Session.getEffectiveUser().getEmail();
}
