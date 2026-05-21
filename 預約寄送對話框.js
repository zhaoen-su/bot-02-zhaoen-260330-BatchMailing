/**
 * 「建立預約寄送」HtmlService 對話框：只問「草稿主旨」，預定寄送時間
 * 自動使用「下一個未到的週五 TRIGGER_HOUR:00」。
 *
 * 沿用以前的彈性：使用者若想覆寫個別列的時間，在「主要操作區」的
 * 「預定寄送時間」欄填日期即可，runCreateScheduledDrafts_ 仍會優先
 * 採用列上的時間。
 *
 * 流程：
 *   1. 選單呼叫 createScheduledDrafts → showModalDialog
 *   2. server 端先算好預設時間，連同 HTML 一起回給 client
 *   3. 對話框送出表單 → google.script.run.runScheduledDraftCreation(subject, time)
 *   4. server 端跑 runCreateScheduledDrafts_，回傳 { successCount, errors }
 */

/**
 * 選單入口：開啟對話框。
 */
function createScheduledDrafts() {
  const html = HtmlService.createHtmlOutput(buildScheduleDialogHtml_())
    .setWidth(460)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, "建立預約寄送");
}

/**
 * 對話框透過 google.script.run 呼叫的 server 端入口。
 *
 * @param {string} subjectLine
 * @param {string} scheduledTimeStr "YYYY-MM-DDTHH:mm"，由對話框 hidden input 傳回
 * @return {{successCount:number, errors:string[]}}
 */
function runScheduledDraftCreation(subjectLine, scheduledTimeStr) {
  return runCreateScheduledDrafts_(subjectLine, scheduledTimeStr);
}

/**
 * 算出「下一個未到的週五 TRIGGER_HOUR:00」。
 * 用「未到」而非「當週」，避免今天已過週五 10:00 時排到過去時間
 * 觸發器掃到就馬上送，違反使用者「下次週五一起送」的預期。
 */
function nextFridayTenAm_() {
  const now = new Date();
  const d = new Date(now);
  d.setHours(TRIGGER_HOUR, 0, 0, 0);
  // getDay: 0=Sun, 5=Fri
  let daysUntilFriday = (5 - d.getDay() + 7) % 7;
  if (daysUntilFriday === 0 && now.getTime() >= d.getTime()) {
    daysUntilFriday = 7;
  }
  d.setDate(d.getDate() + daysUntilFriday);
  return d;
}

function buildScheduleDialogHtml_() {
  const tz = Session.getScriptTimeZone();
  const scheduledAt = nextFridayTenAm_();
  const isoLocal = Utilities.formatDate(scheduledAt, tz, "yyyy-MM-dd'T'HH:mm");
  const weekdayNames = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"];
  const human =
    Utilities.formatDate(scheduledAt, tz, "yyyy/MM/dd") +
    `（${weekdayNames[scheduledAt.getDay()]}）` +
    Utilities.formatDate(scheduledAt, tz, " HH:mm");

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
    input[type="text"] {
      display: block; width: 100%; box-sizing: border-box;
      padding: 8px 10px; margin-top: 6px;
      border: 1px solid #d1d5db; border-radius: 6px; font-size: 13px;
    }
    input:focus { outline: none; border-color: #2563eb; }
    .scheduled-box {
      background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px;
      padding: 10px 12px; margin-bottom: 16px; font-size: 13px;
    }
    .scheduled-box .label { color: #6b7280; font-size: 11px; }
    .scheduled-box .value { font-weight: 600; margin-top: 2px; }
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
    <div class="scheduled-box">
      <div class="label">統一預定寄送時間</div>
      <div class="value">${human}</div>
      <div class="hint">
        如要改個別列的時間，請在「主要操作區」的「預定寄送時間」欄填寫；
        該欄空白者套用上方時間。
      </div>
    </div>
    <input type="hidden" name="time" value="${isoLocal}" />
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
