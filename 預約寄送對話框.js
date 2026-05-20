/**
 * 「建立預約寄送」HtmlService 對話框：把「草稿主旨」+「統一預定寄送時間」
 * 合併在同一個視窗。比連續兩個 Browser.inputBox 順手，使用者也能看到
 * 即時結果 / 錯誤而不用再彈一個 msgBox。
 *
 * 流程：
 *   1. 選單呼叫 createScheduledDrafts → showModalDialog
 *   2. 對話框送出表單 → google.script.run.runScheduledDraftCreation(subject, time)
 *   3. server 端跑 runCreateScheduledDrafts_，回傳 { successCount, errors }
 *   4. 對話框顯示結果（不自動關閉，讓使用者讀完錯誤再手動關）
 */

/**
 * 選單入口：開啟對話框。
 */
function createScheduledDrafts() {
  const html = HtmlService.createHtmlOutput(buildScheduleDialogHtml_())
    .setWidth(460)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, "建立預約寄送");
}

/**
 * 對話框透過 google.script.run 呼叫的 server 端入口。
 * 把實際邏輯薄薄包一層，方便日後想加額外的權限檢查或記錄。
 *
 * @param {string} subjectLine
 * @param {string} scheduledTimeStr 來自 <input type="datetime-local">，
 *                                  格式 "YYYY-MM-DDTHH:mm"；可為空字串。
 * @return {{successCount:number, errors:string[]}}
 */
function runScheduledDraftCreation(subjectLine, scheduledTimeStr) {
  return runCreateScheduledDrafts_(subjectLine, scheduledTimeStr);
}

function buildScheduleDialogHtml_() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI',
                   'PingFang TC', 'Microsoft JhengHei', sans-serif;
      margin: 0; padding: 20px; color: #1f2937;
    }
    label { display: block; margin-bottom: 16px; font-size: 13px; font-weight: 500; }
    .hint { color: #6b7280; font-size: 11px; margin-top: 4px; font-weight: normal; }
    input[type="text"], input[type="datetime-local"] {
      display: block; width: 100%; box-sizing: border-box;
      padding: 8px 10px; margin-top: 6px;
      border: 1px solid #d1d5db; border-radius: 6px; font-size: 13px;
    }
    input:focus { outline: none; border-color: #2563eb; }
    .actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; }
    button {
      padding: 8px 16px; border-radius: 6px; font-size: 13px;
      border: none; cursor: pointer;
    }
    button.primary { background: #2563eb; color: #fff; }
    button.primary:disabled { background: #93c5fd; cursor: progress; }
    button.cancel { background: #f3f4f6; color: #374151; }
    #result {
      margin-top: 16px; font-size: 12px; white-space: pre-line;
      padding: 10px 12px; border-radius: 6px;
    }
    #result.ok    { background: #ecfdf5; color: #065f46; }
    #result.error { background: #fef2f2; color: #991b1b; }
    #result:empty { display: none; }
  </style>
</head>
<body>
  <form id="form">
    <label>
      草稿主旨（與 Gmail 草稿欄相符，可含 {{欄位名}}）
      <input name="subject" type="text" required autofocus />
    </label>
    <label>
      統一預定寄送時間
      <input name="time" type="datetime-local" />
      <div class="hint">
        留空則改用每列「預定寄送時間」欄；兩者都空白的列會被略過。
      </div>
    </label>
    <div class="actions">
      <button type="button" class="cancel" onclick="google.script.host.close()">關閉</button>
      <button type="submit" class="primary" id="submit">建立預約</button>
    </div>
  </form>
  <div id="result"></div>
  <script>
    const form = document.getElementById('form');
    const submitBtn = document.getElementById('submit');
    const result = document.getElementById('result');

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(form);
      submitBtn.disabled = true;
      submitBtn.textContent = '建立中…';
      result.className = '';
      result.textContent = '';

      google.script.run
        .withSuccessHandler((r) => {
          result.className = r.errors.length ? 'error' : 'ok';
          result.textContent =
            '已建立 ' + r.successCount + ' 封預約草稿' +
            (r.errors.length ? '\\n\\n錯誤：\\n' + r.errors.join('\\n') : '');
          submitBtn.disabled = false;
          submitBtn.textContent = '建立預約';
        })
        .withFailureHandler((err) => {
          result.className = 'error';
          result.textContent = '失敗：' + err.message;
          submitBtn.disabled = false;
          submitBtn.textContent = '建立預約';
        })
        .runScheduledDraftCreation(fd.get('subject'), fd.get('time'));
    });
  </script>
</body>
</html>
  `;
}
