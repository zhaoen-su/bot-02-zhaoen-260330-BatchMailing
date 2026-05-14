/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
 */
function getGmailTemplateFromDrafts_(subject_line) {
  const drafts = GmailApp.getDrafts();
  const draft = drafts.filter(subjectFilter_(subject_line))[0];
  if (!draft) {
    throw new Error("找不到相符的主旨");
  }

  const msg = draft.getMessage();
  const htmlBody = msg.getBody();
  const attachments = msg.getAttachments({ includeInlineImages: false });

  // Inline image 用 Content-ID 標頭來對應 cid → blob，不靠檔名 / alt / 順序。
  // 草稿經多次編輯後，MIME 中 attachment 的儲存順序未必跟 HTML 中 <img> 出現
  // 順序一致；用順序配對會讓圖張冠李戴，且因尺寸不同造成比例變形。
  const inlineImagesObj = collectInlineImagesByContentId_(msg.getId());

  // 直接沿用草稿上的 Cc / Bcc，讓使用者改副本時只需要改 Gmail 草稿，
  // 不用碰程式碼或試算表欄位。getCc()/getBcc() 沒填會回空字串。
  return {
    message: {
      subject: subject_line,
      text: msg.getPlainBody(),
      html: htmlBody,
      cc: msg.getCc(),
      bcc: msg.getBcc(),
    },
    attachments: attachments,
    inlineImages: inlineImagesObj,
  };
}

/**
 * 走訪訊息的 MIME tree，依各 part 的 Content-ID 標頭組成 cid → Blob 的對應表。
 * 需要 appsscript.json 啟用 Advanced Gmail Service。
 */
function collectInlineImagesByContentId_(messageId) {
  const fullMsg = Gmail.Users.Messages.get("me", messageId);
  const result = {};

  const walk = (parts) => {
    if (!parts) return;
    for (const part of parts) {
      if (part.parts) walk(part.parts);

      const headers = part.headers || [];
      const cidHeader = headers.find(
        (h) => h.name.toLowerCase() === "content-id",
      );
      if (!cidHeader) continue;

      // Content-ID 標頭值通常包成 <ii_xxx>，HTML 中 src 是 cid:ii_xxx，
      // 所以兩邊比對前要先去掉外層的角括號。
      const cid = cidHeader.value.replace(/^<|>$/g, "");

      const body = part.body || {};
      let data;
      if (body.attachmentId) {
        data = Gmail.Users.Messages.Attachments.get(
          "me",
          messageId,
          body.attachmentId,
        ).data;
      } else if (body.data) {
        data = body.data;
      } else {
        continue;
      }
      if (data == null) continue;

      // Apps Script 的 Advanced Gmail Service 已經幫我們把 base64url 解成
      // Byte[]，型別是陣列不是字串。只有少數情境會回字串形式，這裡兩種都吃。
      const bytes = typeof data === "string" ? decodeBase64Url_(data) : data;
      result[cid] = Utilities.newBlob(bytes, part.mimeType, part.filename);
    }
  };

  walk(fullMsg.payload && fullMsg.payload.parts);
  return result;
}

/**
 * Gmail API 回傳的是 base64url（用 - / _ 取代 + / /，且常省略 = 結尾 padding）。
 * Utilities.base64DecodeWebSafe 對沒有 padding 的輸入有時會丟「無法解碼字串」，
 * 這裡先轉成標準 base64 並補齊 padding，再用較穩定的 base64Decode。
 */
function decodeBase64Url_(s) {
  if (!s) return [];
  // 先濾掉所有非 base64url 字元（空白、換行、tab 等）。
  // Gmail API 的回應在長字串中偶爾會夾換行或不可見字元，沒清掉的話
  // base64Decode 會直接吐「無法解碼字串」。
  const cleaned = String(s).replace(/[^A-Za-z0-9\-_+/=]/g, "");
  const standard = cleaned.replace(/-/g, "+").replace(/_/g, "/");
  const padded = standard + "=".repeat((4 - (standard.length % 4)) % 4);
  try {
    return Utilities.base64Decode(padded);
  } catch (e) {
    // 印出前後段協助 debug，再把錯誤往外丟
    Logger.log(
      "decodeBase64Url_ 失敗，長度=%s 前 60 字元=%s 後 60 字元=%s",
      padded.length,
      padded.slice(0, 60),
      padded.slice(-60),
    );
    throw e;
  }
}

/**
 * Filter draft objects with the matching subject line.
 * @param {string} subject_line to search for draft message
 * @return {Function} filter callback
 */
function subjectFilter_(subject_line) {
  return (element) => {
    if (element.getMessage().getSubject() === subject_line) {
      return element;
    }
  };
}

/**
 * Fill template string with data object.
 * @param {object} template object containing subject, text, and html
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
 */
function fillInTemplateFromObject_(template, data) {
  const replace = (str) =>
    String(str).replace(
      /{{([^{}]+)}}/g,
      (_, key) => data[key.trim()] ?? "",
    );
  return {
    subject: replace(template.subject),
    text: replace(template.text),
    html: replace(template.html),
  };
}
