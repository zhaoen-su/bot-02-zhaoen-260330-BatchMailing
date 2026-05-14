/**
 * 自己組 RFC 822 raw MIME 訊息，搭配 Gmail.Users.Messages.send 寄出。
 *
 * 為什麼不用 GmailApp.sendEmail：它的 MIME encoder 對 surrogate pair（U+10000
 * 以上的字元，如 emoji、數學雙線字 𝕏𝕦𝕒𝕟）會壞掉，主旨 / 純文字 / HTML 都
 * 可能變亂碼。直接走 base64 + UTF-8，所有 Unicode 都能正確送達。
 */

/**
 * 組整封 RFC 822 訊息，回傳 string（後續再交給 base64UrlEncode_ 餵給 Gmail API）。
 *
 * 結構：
 *   multipart/mixed                ← 有附件才包這層
 *     multipart/related            ← 有 inline image 才包這層
 *       multipart/alternative
 *         text/plain;  base64; UTF-8
 *         text/html;   base64; UTF-8
 *       image/...; cid:xxx
 *     application/...; attachment
 */
function buildMimeMessage_(opts) {
  const {
    from,
    to,
    cc,
    bcc,
    subject,
    text,
    html,
    inlineImages,
    attachments,
  } = opts;

  const inlineKeys = Object.keys(inlineImages || {});
  const hasInline = inlineKeys.length > 0;
  const hasAttachments = (attachments || []).length > 0;

  // 最內層：text + html 的 alternative
  const altBoundary = `alt_${Utilities.getUuid()}`;
  let bodyContentType = `multipart/alternative; boundary="${altBoundary}"`;
  let body = assembleParts_(altBoundary, [
    buildTextPart_(text || "", "text/plain"),
    buildTextPart_(html || "", "text/html"),
  ]);

  // 有 inline image → 包成 multipart/related
  if (hasInline) {
    const relBoundary = `rel_${Utilities.getUuid()}`;
    const altPart = `Content-Type: ${bodyContentType}\r\n\r\n${body}`;
    const imageParts = inlineKeys.map((cid) =>
      buildInlineImagePart_(cid, inlineImages[cid]),
    );
    body = assembleParts_(relBoundary, [altPart, ...imageParts]);
    bodyContentType = `multipart/related; boundary="${relBoundary}"`;
  }

  // 有附件 → 再包成 multipart/mixed
  if (hasAttachments) {
    const mixedBoundary = `mixed_${Utilities.getUuid()}`;
    const innerPart = `Content-Type: ${bodyContentType}\r\n\r\n${body}`;
    const attachmentParts = attachments.map((blob) => buildAttachmentPart_(blob));
    body = assembleParts_(mixedBoundary, [innerPart, ...attachmentParts]);
    bodyContentType = `multipart/mixed; boundary="${mixedBoundary}"`;
  }

  const headers = [];
  if (from) headers.push(`From: ${from}`);
  headers.push(`To: ${to}`);
  if (cc) headers.push(`Cc: ${cc}`);
  if (bcc) headers.push(`Bcc: ${bcc}`);
  headers.push(`Subject: ${encodeRfc2047_(subject)}`);
  headers.push("MIME-Version: 1.0");
  headers.push(`Content-Type: ${bodyContentType}`);

  return headers.join("\r\n") + "\r\n\r\n" + body;
}

/**
 * 把多個 part 用 boundary 串成 multipart body。
 */
function assembleParts_(boundary, parts) {
  let result = "";
  for (const part of parts) {
    result += `--${boundary}\r\n${part}\r\n`;
  }
  result += `--${boundary}--\r\n`;
  return result;
}

/**
 * text/plain 或 text/html 的 part：base64 + UTF-8。
 * surrogate pair 會經由 newBlob → UTF-8 bytes 正確編碼。
 */
function buildTextPart_(content, mimeType) {
  const bytes = Utilities.newBlob(content).getBytes();
  const base64 = chunkBase64_(Utilities.base64Encode(bytes));
  return [
    `Content-Type: ${mimeType}; charset=UTF-8`,
    "Content-Transfer-Encoding: base64",
    "",
    base64,
  ].join("\r\n");
}

function buildInlineImagePart_(cid, blob) {
  const base64 = chunkBase64_(Utilities.base64Encode(blob.getBytes()));
  return [
    `Content-Type: ${blob.getContentType() || "application/octet-stream"}`,
    "Content-Transfer-Encoding: base64",
    `Content-ID: <${cid}>`,
    "Content-Disposition: inline",
    "",
    base64,
  ].join("\r\n");
}

function buildAttachmentPart_(blob) {
  const base64 = chunkBase64_(Utilities.base64Encode(blob.getBytes()));
  const filename = blob.getName() || "attachment";
  const mimeType = blob.getContentType() || "application/octet-stream";
  return [
    `Content-Type: ${mimeType}`,
    "Content-Transfer-Encoding: base64",
    `Content-Disposition: attachment; ${formatFilename_(filename)}`,
    "",
    base64,
  ].join("\r\n");
}

/**
 * RFC 5322 規定每行 ≤ 998 octets；base64 慣例切 76 字元一行。
 */
function chunkBase64_(s) {
  if (!s) return "";
  return s.match(/.{1,76}/g).join("\r\n");
}

/**
 * RFC 2231 編碼非 ASCII 檔名，避免中文附件名亂碼。ASCII 檔名走簡單形式。
 */
function formatFilename_(name) {
  if (/^[\x20-\x7E]*$/.test(name) && !/[";\\]/.test(name)) {
    return `filename="${name}"`;
  }
  return `filename*=UTF-8''${encodeURIComponent(name)}`;
}

/**
 * RFC 2047 encoded-word：=?UTF-8?B?<base64>?=
 *
 * 為了避開 75 字元上限，按 code point（用 for...of 自動處理 surrogate pair）
 * 切成小段，每段獨立做一個 encoded-word，連續 encoded-word 間用 CRLF + 空白
 * 接續（folding）。10 個 code unit 是保守上限：每字最多 4 bytes UTF-8、base64
 * 膨脹 4/3，加 wrapper 約 12 字，10 × 4 × 4/3 + 12 ≈ 65，在 75 之內。
 */
function encodeRfc2047_(str) {
  if (!str) return "";
  if (/^[\x20-\x7E]*$/.test(str)) return str;

  const chunks = [];
  let buf = "";
  for (const ch of str) {
    if (buf.length >= 10) {
      chunks.push(buf);
      buf = "";
    }
    buf += ch;
  }
  if (buf) chunks.push(buf);

  return chunks
    .map(
      (c) =>
        `=?UTF-8?B?${Utilities.base64Encode(Utilities.newBlob(c).getBytes())}?=`,
    )
    .join("\r\n ");
}

/**
 * Gmail.Users.Messages.send 要的 raw 格式：base64url（- _ 取代 + /）。
 */
function base64UrlEncode_(str) {
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(str).getBytes());
}
