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
    throw new Error("Oops - can't find Gmail draft");
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
  let template_string = JSON.stringify(template);

  template_string = template_string.replace(/{{[^{}]+}}/g, (key) => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe.
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
 */
function escapeData_(str) {
  return str
    .replace(/[\\]/g, "\\\\")
    .replace(/[\"]/g, '\\"')
    .replace(/[\/]/g, "\\/")
    .replace(/[\b]/g, "\\b")
    .replace(/[\f]/g, "\\f")
    .replace(/[\n]/g, "\\n")
    .replace(/[\r]/g, "\\r")
    .replace(/[\t]/g, "\\t");
}
