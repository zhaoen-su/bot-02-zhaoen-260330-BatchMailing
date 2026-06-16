/**
 * 全域設定。修改前請先確認試算表上的欄位名稱與這裡一致。
 */

// 主要操作區 ──────────────────────────────────────────────
const RECIPIENT_COL = "收件者";
// 主要操作區可選欄位：每列指定的預定寄送時間。空白者會 fallback
// 到「建立預約寄送」對話框輸入的統一時間。
const MAIN_COL_SCHEDULED_AT = "預定寄送時間";

// 寄件人 ──────────────────────────────────────────────────
// 以別名身份寄出；從「專案設定 → 指令碼屬性」讀取，方便不同部署
// 環境用不同值，不必改程式。沒設定就 fallback 成空字串 = 用預設帳號寄出。
// 必須是 GmailApp.getAliases() 回傳清單內的地址（執行 checkAliases 驗證）。
//
// 設定 key：SENDER_ALIAS（例：hr@yourdomain.com）、SENDER_NAME（例：HR 部門）
const SENDER_ALIAS = (
  PropertiesService.getScriptProperties().getProperty("SENDER_ALIAS") || ""
).trim();
const SENDER_NAME = (
  PropertiesService.getScriptProperties().getProperty("SENDER_NAME") || ""
).trim();

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
// 觸發器本身設定為每小時掃一次「已預約區」，不依賴這個值。
// 這個常數只給「建立預約寄送」對話框算「下一個未到的週五 ?:00」作為預設時間，
// 例如想改成「下一個週五 14:00」就把 14 寫在這裡。
const TRIGGER_HOUR = 10;

/**
 * 寄件人「身份識別字串」，用來在「已預約區」標記列的擁有者。
 * 觸發器執行時，getEffectiveUser() 回的是觸發器擁有者，所以這個值會跟
 * 當初「建立預約寄送」時寫入的值相符 → 觸發器只會處理自己建立的列。
 *
 * 呼叫端應先跑 assertSenderConfigured_() 確認 SENDER_ALIAS 已設定。
 */
function getSenderIdentity_() {
  return SENDER_ALIAS;
}

/**
 * 驗證寄件人 alias / name 已設定且 alias 真的被 Gmail 授權。
 * 設計目的：避免「Properties 沒設」或「Gmail 沒把該地址列為 send-as」時
 * Gmail 後端默默把 From 改回執行者本人帳號，造成寄件人對外身份不一致。
 *
 * 三種失敗都會直接拋錯：
 *   1. SENDER_ALIAS 為空字串
 *   2. SENDER_NAME 為空字串
 *   3. SENDER_ALIAS 既不是執行身份本人的地址，也不在 GmailApp.getAliases() 清單內
 */
function assertSenderConfigured_() {
  if (!SENDER_ALIAS) {
    throw new Error(
      "SENDER_ALIAS 未設定。請到「專案設定 → 指令碼屬性」新增 key=SENDER_ALIAS。",
    );
  }
  if (!SENDER_NAME) {
    throw new Error(
      "SENDER_NAME 未設定。請到「專案設定 → 指令碼屬性」新增 key=SENDER_NAME。",
    );
  }

  // 寄件地址只要符合下列兩者之一，Gmail 就會照用、不會默默改回本人帳號：
  //   (a) 等於目前執行身份本人的地址（用自己帳號寄）
  //   (b) 是已在 Gmail「以這個地址寄送郵件」驗證過的 send-as 別名
  //
  // 關鍵陷阱：GmailApp.getAliases() 只列「額外的 send-as 別名」，不含本人主要
  // 地址。所以當 SENDER_ALIAS 就是執行者自己的主要地址時，getAliases() 會是空的，
  // 若只比對別名清單就會把「用自己帳號寄」這個完全合法的情況誤擋。
  // email 比對去空白、忽略大小寫，避免指令碼屬性裡多打空白或大小寫不同被誤擋。
  const sameAddr = (a, b) =>
    String(a).trim().toLowerCase() === String(b).trim().toLowerCase();

  const me = Session.getEffectiveUser().getEmail();
  if (sameAddr(SENDER_ALIAS, me)) return;

  const aliases = GmailApp.getAliases();
  if (!aliases.some((a) => sameAddr(a, SENDER_ALIAS))) {
    throw new Error(
      `SENDER_ALIAS "${SENDER_ALIAS}" 既不是目前執行身份（${me}）的地址，` +
        `也不在 Gmail 已驗證的別名清單內。\n\n` +
        `若想用 ${me} 寄出：把指令碼屬性 SENDER_ALIAS 改成 ${me}。\n` +
        `若想用 ${SENDER_ALIAS} 寄出：請先到 Gmail 設定 →「帳戶和匯入」→` +
        `「以這個地址寄送郵件」新增並完成驗證。\n\n` +
        `目前可用別名：${aliases.join("、") || "（無，代表這個帳號沒有額外別名）"}`,
    );
  }
}
