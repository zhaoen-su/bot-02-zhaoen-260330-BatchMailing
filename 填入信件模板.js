/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
 */
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    const allInlineImages = draft.getMessage().getAttachments({
      includeInlineImages: true,
      includeAttachments: false,
    });
    const attachments = draft
      .getMessage()
      .getAttachments({ includeInlineImages: false });
    const htmlBody = msg.getBody();

    // Creates an inline image object with the image name as key
    const img_obj = allInlineImages.reduce((obj, i) => {
      obj[i.getName()] = i;
      return obj;
    }, {});

    // Regexp searches for all img string positions with cid
    const imgexp = /<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>/g;
    const matches = [...htmlBody.matchAll(imgexp)];

    const inlineImagesObj = {};
    for (const match of matches) {
      inlineImagesObj[match[1]] = img_obj[match[2]];
    }

    return {
      message: {
        subject: subject_line,
        text: msg.getPlainBody(),
        html: htmlBody,
      },
      attachments: attachments,
      inlineImages: inlineImagesObj,
    };
  } catch (e) {
    throw new Error("找不到相符的主旨");
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
    html: encodeEmoji_(replace(template.html)),
  };
}

/**
 * GmailApp.sendEmail 的 MIME encoder 無法正確處理 surrogate pair（如 emoji），
 * 會變成 ? 之類的 placeholder。送出前把所有 U+10000 以上的字元轉成
 * HTML 數字實體（例如 🔥 → &#128293;），email client 會照常渲染。
 */
function encodeEmoji_(str) {
  return String(str).replace(
    /[\u{10000}-\u{10FFFF}]/gu,
    (ch) => `&#${ch.codePointAt(0)};`,
  );
}
