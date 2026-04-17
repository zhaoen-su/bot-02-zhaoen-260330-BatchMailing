/**
 * 內建 HTML 信件模板
 * ------------------------------------------------------------
 * 不再從 Gmail 草稿讀取，改為直接寫死在這個檔案。
 * 如需調整：
 *   1. 改 EMAIL_SUBJECT 設定主旨
 *   2. 在 EMAIL_HTML_TEMPLATE 內用 {{欄位名稱}} 插入 Sheet 欄位
 *      （欄位名稱需與試算表第一列的標題完全相同）
 */

const EMAIL_SUBJECT = "【CAPSULE】甄選感謝函 - {{姓名}}";

// Drive 上 logo 圖片的檔案 ID（會以 inline image 形式內嵌進信件）
const INLINE_LOGO_FILE_ID = "1cR75kwNoAapopNax7Ud6LH8uQtjcFP67";

const EMAIL_HTML_TEMPLATE = `
<!DOCTYPE html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{{姓名}}</title>
  </head>
  <body style="margin:0; padding:0; background-color:#f1f5f9; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang TC','Microsoft JhengHei',Roboto,sans-serif; color:#0f172a;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f1f5f9; padding:40px 16px;">
      <tr>
        <td align="center">
          <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px; width:100%; background-color:#ffffff; border-radius:16px; box-shadow:0 1px 3px rgba(15,23,42,0.08); overflow:hidden;">

            <!-- 頂部色帶 -->
            <tr>
              <td style="background:linear-gradient(135deg,#6366f1 0%,#8b5cf6 100%); height:6px; font-size:0; line-height:0;">&nbsp;</td>
            </tr>

            <!-- Logo -->
            <tr>
              <td align="center" style="padding:40px 48px 0 48px;">
                <img src="cid:logo" alt="CAPSULE" width="220" style="display:block; border:0; outline:none; text-decoration:none; max-width:220px; height:auto;" />
              </td>
            </tr>

            <!-- 主要內容 -->
            <tr>
              <td style="padding:32px 48px 32px 48px;">
                <h1 style="margin:0 0 24px 0; font-size:24px; line-height:1.3; font-weight:700; color:#0f172a;">
                  {{姓名}} 您好：
                </h1>
                <p style="margin:0 0 16px 0; font-size:16px; line-height:1.7; color:#334155;">
                  非常感謝您日前特地撥冗參與本公司面試。
                  基於多項綜合考量，並經過整體評估與審慎討論後，
                  我們很遺憾的通知您<b>未能錄取</b>。
                </p>
                <p style="margin:0 0 32px 0; font-size:16px; line-height:1.7; color:#334155;">
                  再次感謝您對於我們公司的愛護與支持，
                  未來若有其他適合您的職缺，仍歡迎您再次投遞，
                  祝福您未來一切順利，謝謝您！
                </p>
              </td>
            </tr>

            <!-- 分隔線 -->
            <tr>
              <td style="padding:0 48px;">
                <div style="border-top:1px solid #e2e8f0; height:1px; line-height:1px; font-size:0;">&nbsp;</div>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="padding:24px 48px 40px 48px;">
                <p style="margin:0 0 4px 0; font-size:13px; line-height:1.6; color:#64748b;">
                  CAPSULE CXO TEAM
                </p>
                <p style="margin:0 0 12px 0; font-size:12px; line-height:1.6; color:#94a3b8;">
                  這封信是由批次寄信小工具自動寄出。
                </p>
                <p style="margin:0; font-size:12px; line-height:1.6; color:#94a3b8;">
                  若不希望再收到此類信件，請
                  <a href="mailto:unsubscribe@example.com?subject=取消訂閱&body=請將我從寄送名單中移除，謝謝。"
                     style="color:#6366f1; text-decoration:underline;">
                    取消訂閱
                  </a>
                  。
                </p>
              </td>
            </tr>

          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
`;

/**
 * 取得內建 HTML 模板，結構與 getGmailTemplateFromDrafts_() 相同。
 * @return {object} 包含 message / attachments / inlineImages
 */
function getBuiltInTemplate_() {
  // 由 HTML 自動產生純文字版本（去掉 tag、壓縮空白），fallback 用
  const plainText = EMAIL_HTML_TEMPLATE.replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/\s+\n/g, "\n")
    .replace(/\n\s+/g, "\n")
    .replace(/\n{2,}/g, "\n\n")
    .trim();

  // 從 Drive 載入 logo 圖片，作為 inline image 嵌入信件
  // key "logo" 對應 HTML 內的 <img src="cid:logo">
  const logoBlob = DriveApp.getFileById(INLINE_LOGO_FILE_ID).getBlob();

  return {
    message: {
      subject: EMAIL_SUBJECT,
      text: plainText,
      html: EMAIL_HTML_TEMPLATE,
    },
    attachments: [],
    inlineImages: {
      logo: logoBlob,
    },
  };
}

/**
 * 測試用獨立寄信入口：使用本檔案內建的 HTML 模板，
 * 不從 Gmail 草稿讀取主旨。所有邏輯（讀 Sheet、變數替換、寄信、回寫時間）
 * 都封裝在這個檔案內，與原本的 sendEmails() 互不干擾。
 *
 * 在 Apps Script 編輯器或工具列選單呼叫 sendEmailsHTML 即可。
 * @param {Sheet} sheet 預設為目前作用中的工作表
 */
function sendEmailsHTML(sheet = SpreadsheetApp.getActiveSheet()) {
  const emailTemplate = getBuiltInTemplate_();

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  // 第一列作為欄位標題
  const heads = data.shift();
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  // 將二維陣列轉成以欄位名為 key 的物件陣列
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => {
      o[k] = r[i] || "";
      return o;
    }, {}),
  );

  const out = [];

  obj.forEach((row) => {
    if (row[EMAIL_SENT_COL] === "") {
      try {
        const msgObj = fillInHTMLTemplate_(emailTemplate.message, row);

        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages,
        });
        out.push([new Date()]);
      } catch (e) {
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
}

/**
 * 將 {{欄位名}} 標記替換為 row 物件對應的值。
 * 與 fillInTemplateFromObject_ 邏輯相同，刻意改名避免與另一檔案衝突，
 * 讓本檔案能完全自包含。
 */
function fillInHTMLTemplate_(template, data) {
  const replace = (str) =>
    String(str).replace(
      /{{([^{}]+)}}/g,
      (_, key) => data[key.trim()] ?? "",
    );
  return {
    subject: replace(template.subject),
    text: replace(template.text),
    html: encodeEmojiHTML_(replace(template.html)),
  };
}

/**
 * GmailApp.sendEmail 的 MIME encoder 無法正確處理 surrogate pair（如 emoji），
 * 會變成 ? 之類的 placeholder。送出前把所有 U+10000 以上的字元轉成
 * HTML 數字實體（例如 🔥 → &#128293;），email client 會照常渲染。
 */
function encodeEmojiHTML_(str) {
  return String(str).replace(
    /[\u{10000}-\u{10FFFF}]/gu,
    (ch) => `&#${ch.codePointAt(0)};`,
  );
}
