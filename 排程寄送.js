/**
 * 排程寄送：管理時間觸發器，並由觸發器定時掃描「已預約區」寄出到期的草稿。
 *
 * 為何要這樣設計：
 *   - Gmail API 沒有開放排程寄送欄位（scheduled send 是 Gmail UI 專有功能），
 *     所以必須靠 Apps Script 的時間觸發器在指定時間後檢查並送出。
 *   - 觸發器以「設定者」的身份執行（getEffectiveUser），所以 processScheduledDrafts
 *     只會處理自己當初建立的列（用「建立者」欄做身份比對）。
 *     避免多人共用試算表時把別人的草稿弄丟、或拿著對方 draftId 在自己信箱找不到。
 */

/**
 * 安裝定時觸發器（先移除既有的同名觸發器，避免重複）。
 * 由選單呼叫。
 */
function installScheduledTrigger() {
  removeScheduledTrigger();
  ScriptApp.newTrigger(TRIGGER_HANDLER)
    .timeBased()
    .everyHours(1)
    .create();
  Browser.msgBox(
    `已啟用排程觸發器（每小時掃描一次「已預約區」）。\n` +
      `執行身份：${Session.getEffectiveUser().getEmail()}`,
  );
}

/**
 * 移除這個專案內所有指向 TRIGGER_HANDLER 的觸發器。
 * 注意：每個 Apps Script 使用者只看得到自己安裝的觸發器，
 * 所以這裡也只會移除「目前執行者自己」安裝的那些。
 */
function removeScheduledTrigger() {
  const triggers = ScriptApp.getProjectTriggers().filter(
    (t) => t.getHandlerFunction() === TRIGGER_HANDLER,
  );
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));
  if (triggers.length > 0) {
    SpreadsheetApp.getActive().toast(`已移除 ${triggers.length} 個排程觸發器`);
  }
  return triggers.length;
}

/**
 * 給選單用的「停用排程觸發器」入口。
 * 跟 removeScheduledTrigger 的差別：不管有沒有移除到都會跳訊息給使用者看，
 * 避免從選單按了沒反應的困惑。
 */
function removeScheduledTriggerFromMenu() {
  const n = removeScheduledTrigger();
  Browser.msgBox(
    n > 0
      ? `已移除 ${n} 個排程觸發器（執行身份：${Session.getEffectiveUser().getEmail()}）。`
      : `目前沒有由你（${Session.getEffectiveUser().getEmail()}）安裝的排程觸發器。`,
  );
}

/**
 * 給選單用的「檢查排程觸發器狀態」入口。
 * 開啟 HTML 對話框，呈現：
 *   - 目前執行身份 + SENDER_ALIAS/NAME 的設定與授權狀態
 *   - 屬於這個身份的觸發器清單（handler / event type / source / id）
 *   - 「已預約區」屬於這個身份的待寄列數與最近一筆預定時間
 *
 * 目的：讓使用者一眼看出「是否真的有觸發器在跑」「為什麼到時間沒寄」。
 */
function checkTriggerStatus() {
  const data = collectTriggerStatus_();
  const html = HtmlService.createHtmlOutput(buildTriggerStatusHtml_(data))
    .setWidth(520)
    .setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(html, "排程觸發器狀態");
}

/**
 * 蒐集要顯示的觸發器狀態。
 * 注意：ScriptApp.getProjectTriggers() 只看得到「呼叫者自己」安裝的觸發器，
 * 所以這裡看到的數量就是「會幫目前使用者處理草稿的觸發器」數量。
 */
function collectTriggerStatus_() {
  const me = Session.getEffectiveUser().getEmail();

  const triggers = ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === TRIGGER_HANDLER)
    .map((t) => ({
      id: t.getUniqueId(),
      handler: t.getHandlerFunction(),
      eventType: String(t.getEventType()),
      source: String(t.getTriggerSource()),
    }));

  // SENDER_ALIAS 設定與授權狀態 — 觸發器到時間真的會跑，但如果 alias
  // 沒授權，processScheduledDrafts 進去就會被 assertSenderConfigured_ 擋掉。
  const aliasAuthorized = SENDER_ALIAS
    ? safeIsAliasAuthorized_(SENDER_ALIAS)
    : null;

  // 「已預約區」中屬於目前 owner 的列數 + 最近一筆預定時間
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const owner = SENDER_ALIAS || me;
  let pendingCount = 0;
  let nextScheduledAt = null;
  const sheet = ss.getSheetByName(SHEET_SCHEDULED);
  if (sheet && sheet.getLastRow() >= 2) {
    const all = sheet
      .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
      .getValues();
    const heads = all.shift();
    const iOwner = heads.indexOf(SCH_COL_OWNER);
    const iAt = heads.indexOf(SCH_COL_SCHEDULED_AT);
    all.forEach((r) => {
      if (String(r[iOwner]) !== owner) return;
      pendingCount++;
      const d = parseDate_(r[iAt]);
      if (d && (!nextScheduledAt || d < nextScheduledAt)) nextScheduledAt = d;
    });
  }

  return {
    user: me,
    senderAlias: SENDER_ALIAS,
    senderName: SENDER_NAME,
    aliasAuthorized,
    triggers,
    pendingCount,
    nextScheduledAt: nextScheduledAt
      ? Utilities.formatDate(nextScheduledAt, tz, "yyyy-MM-dd HH:mm")
      : null,
  };
}

/**
 * 包一層 try/catch — GmailApp.getAliases 在沒授權或被佔用時會丟錯，
 * 不該讓整個狀態對話框因此打不開。失敗回 null 表示「無法判斷」。
 */
function safeIsAliasAuthorized_(alias) {
  try {
    return GmailApp.getAliases().indexOf(alias) !== -1;
  } catch (e) {
    return null;
  }
}

/**
 * 組裝狀態對話框的 HTML。沿用「預約寄送對話框」的 Material 風格。
 * 所有動態值都走 escapeHtml_ 避免 XSS（draftId / 試算表內容可能含 < > & ）。
 */
function buildTriggerStatusHtml_(data) {
  const esc = escapeHtml_;

  const aliasBadge = !data.senderAlias
    ? `<span class="badge warn">未設定</span>`
    : data.aliasAuthorized === false
      ? `<span class="badge error">未被 Gmail 授權</span>`
      : data.aliasAuthorized === true
        ? `<span class="badge ok">已授權</span>`
        : `<span class="badge">無法驗證</span>`;

  const nameBadge = data.senderName
    ? esc(data.senderName)
    : `<span class="badge warn">未設定</span>`;

  const triggerBlock =
    data.triggers.length === 0
      ? `<div class="empty">
           目前由你（${esc(data.user)}）安裝的排程觸發器：0 個。<br/>
           從選單「啟用排程觸發器」可以新增。
         </div>`
      : data.triggers
          .map(
            (t, i) => `
            <div class="trigger-item">
              <div class="trigger-title">觸發器 ${i + 1}</div>
              <dl>
                <dt>Handler</dt><dd><code>${esc(t.handler)}</code></dd>
                <dt>事件類型</dt><dd>${esc(t.eventType)}</dd>
                <dt>來源</dt><dd>${esc(t.source)}</dd>
                <dt>Unique ID</dt><dd><code>${esc(t.id)}</code></dd>
              </dl>
            </div>`,
          )
          .join("");

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet"
        href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500&family=Roboto:wght@400;500&family=Roboto+Mono&display=swap">
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Roboto','Google Sans','PingFang TC','Microsoft JhengHei',-apple-system,sans-serif;
      margin: 0; padding: 24px; background: #fff;
      color: #202124; font-size: 14px; line-height: 1.5;
      -webkit-font-smoothing: antialiased;
    }
    h2 {
      font-family: 'Google Sans', sans-serif;
      font-size: 16px; font-weight: 500; margin: 0 0 16px;
    }
    .section { margin-bottom: 20px; }
    .section-title {
      font-size: 11px; font-weight: 500; color: #5f6368;
      letter-spacing: 0.5px; margin-bottom: 6px;
      text-transform: uppercase;
    }
    .row {
      display: flex; justify-content: space-between; gap: 12px;
      padding: 8px 0; border-bottom: 1px solid #f1f3f4;
      font-size: 13px;
    }
    .row:last-child { border-bottom: none; }
    .row .k { color: #5f6368; flex-shrink: 0; }
    .row .v {
      color: #202124; font-weight: 500;
      text-align: right; word-break: break-all;
    }
    code {
      font-family: 'Roboto Mono', monospace;
      background: #f1f3f4; color: #202124;
      padding: 1px 5px; border-radius: 3px; font-size: 12px;
    }
    .badge {
      display: inline-block; font-size: 11px; font-weight: 500;
      padding: 2px 8px; border-radius: 10px;
      background: #f1f3f4; color: #5f6368;
      margin-left: 4px;
    }
    .badge.ok    { background: #e6f4ea; color: #137333; }
    .badge.warn  { background: #fef7e0; color: #b06000; }
    .badge.error { background: #fce8e6; color: #c5221f; }

    .trigger-item {
      border: 1px solid #dadce0; border-radius: 6px;
      padding: 12px 14px; margin-bottom: 8px;
    }
    .trigger-title {
      font-size: 12px; font-weight: 500; color: #5f6368;
      margin-bottom: 8px; letter-spacing: 0.3px;
    }
    dl { margin: 0; }
    dt {
      float: left; clear: left; width: 80px;
      color: #5f6368; font-size: 12px; padding: 3px 0;
    }
    dd {
      margin: 0 0 0 90px; padding: 3px 0;
      font-size: 13px; word-break: break-all;
    }
    .empty {
      padding: 16px; background: #f8f9fa;
      border-radius: 6px; color: #5f6368;
      font-size: 13px; text-align: center; line-height: 1.6;
    }

    .actions {
      display: flex; justify-content: flex-end; gap: 8px;
      margin-top: 16px;
    }
    button {
      font-family: 'Google Sans','Roboto',sans-serif;
      font-size: 14px; font-weight: 500; letter-spacing: 0.25px;
      height: 36px; padding: 0 24px;
      border: none; border-radius: 4px; cursor: pointer;
      transition: background .15s, box-shadow .15s;
    }
    button.text { background: transparent; color: #1a73e8; padding: 0 16px; }
    button.text:hover { background: rgba(26,115,232,.08); }
    button.primary { background: #1a73e8; color: #fff; }
    button.primary:hover {
      background: #1765cc;
      box-shadow: 0 1px 2px 0 rgba(60,64,67,.3),
                  0 1px 3px 1px rgba(60,64,67,.15);
    }
  </style>
</head>
<body>
  <h2>排程觸發器狀態</h2>

  <div class="section">
    <div class="section-title">執行身份</div>
    <div class="row">
      <span class="k">目前帳號</span>
      <span class="v">${esc(data.user)}</span>
    </div>
    <div class="row">
      <span class="k">SENDER_ALIAS</span>
      <span class="v">${data.senderAlias ? esc(data.senderAlias) : ""}${aliasBadge}</span>
    </div>
    <div class="row">
      <span class="k">SENDER_NAME</span>
      <span class="v">${nameBadge}</span>
    </div>
  </div>

  <div class="section">
    <div class="section-title">觸發器（共 ${data.triggers.length} 個）</div>
    ${triggerBlock}
  </div>

  <div class="section">
    <div class="section-title">已預約區待處理</div>
    <div class="row">
      <span class="k">待寄草稿數</span>
      <span class="v">${data.pendingCount}</span>
    </div>
    <div class="row">
      <span class="k">最近一筆預定時間</span>
      <span class="v">${data.nextScheduledAt ? esc(data.nextScheduledAt) : "—"}</span>
    </div>
  </div>

  <div class="actions">
    <button type="button" class="primary" onclick="google.script.host.close()">關閉</button>
  </div>
</body>
</html>`;
}

/**
 * 簡易 HTML escape，避免試算表內容（draftId、別名、名稱）
 * 帶 < > & " 時破壞 HTML 結構或造成 XSS。
 */
function escapeHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * 觸發器入口。掃「已預約區」，挑出：
 *   - 建立者欄等於當前執行身份 (避免處理到別人的列)
 *   - 預定寄送時間 ≤ 現在
 * 然後逐列 Gmail.Users.Drafts.send，再把該列搬到「已寄出區」。
 *
 * 失敗也搬到已寄出區並寫錯誤訊息，避免下一輪重試同一封造成重寄。
 *
 * @param {boolean=} ignoreScheduledTime
 *   true 時略過「預定時間 ≤ 現在」的檢查，已預約區裡所有 owner 相符
 *   的列都會立刻送出。給「現在寄出」用。
 *   注意：觸發器呼叫時會把 event 物件當第一個引數傳進來，所以這裡
 *   用嚴格相等 `=== true` 比對，避免被誤觸發。
 */
function processScheduledDrafts(ignoreScheduledTime) {
  const force = ignoreScheduledTime === true;
  assertSenderConfigured_();
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduledSheet = ss.getSheetByName(SHEET_SCHEDULED);
  const sentSheet = ss.getSheetByName(SHEET_SENT);

  const lastRow = scheduledSheet.getLastRow();
  if (lastRow < 2) return;

  // getValues 而非 getDisplayValues，這樣日期回來是 Date 物件可直接比對。
  const lastCol = scheduledSheet.getLastColumn();
  const all = scheduledSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const heads = all.shift();
  const idx = (name) => heads.indexOf(name);
  const iRecipient = idx(SCH_COL_RECIPIENT);
  const iSubject = idx(SCH_COL_SUBJECT);
  const iDraftId = idx(SCH_COL_DRAFT_ID);
  const iScheduledAt = idx(SCH_COL_SCHEDULED_AT);
  const iOwner = idx(SCH_COL_OWNER);

  const owner = getSenderIdentity_();
  const now = new Date();
  const rowsToDelete = [];

  all.forEach((rowArr, i) => {
    if (String(rowArr[iOwner]) !== owner) return;

    const scheduledAt = parseDate_(rowArr[iScheduledAt]);
    if (!scheduledAt) return;
    if (!force && scheduledAt.getTime() > now.getTime()) return;

    const draftId = String(rowArr[iDraftId]);
    let result = "成功";
    try {
      Gmail.Users.Drafts.send({ id: draftId }, "me");
    } catch (e) {
      result = `失敗：${e.message}`;
    }

    sentSheet.appendRow(
      buildRowForSheet_(sentSheet, {
        [SENT_COL_RECIPIENT]: rowArr[iRecipient],
        [SENT_COL_SUBJECT]: rowArr[iSubject],
        [SENT_COL_SCHEDULED_AT]: scheduledAt,
        [SENT_COL_SENT_AT]: new Date(),
        [SENT_COL_RESULT]: result,
      }),
    );
    rowsToDelete.push(i + 2);
  });

  rowsToDelete
    .sort((a, b) => b - a)
    .forEach((r) => scheduledSheet.deleteRow(r));
}

/**
 * 取得（或建立）「已預約寄送」標籤。
 */
function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/**
 * 把輸入解析為 Date。接受：
 *   - Date 物件本身（從 sheet getValues 拿到的）
 *   - "YYYY-MM-DD HH:mm" / "YYYY-MM-DD HH:mm:ss" / ISO 8601
 *   - 試算表上以日期格式顯示出來的字串
 *
 * 解析失敗回 null。
 */
function parseDate_(input) {
  if (input instanceof Date) return isNaN(input.getTime()) ? null : input;
  if (input == null || input === "") return null;
  const s = String(input).trim();
  // 把 "YYYY-MM-DD HH:mm" 補成 ISO 友善的 "YYYY-MM-DDTHH:mm" 提高跨平台一致性。
  const iso = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?$/.test(s)
    ? s.replace(" ", "T")
    : s;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? null : d;
}
