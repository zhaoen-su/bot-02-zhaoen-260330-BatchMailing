/**
 * 「建立預約寄送」HtmlService 對話框：問「草稿主旨」與「統一預定寄送時間」。
 * 時間欄預填「下一個未到的週五 TRIGGER_HOUR:00」，但可隨意改。
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
    .setWidth(480)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "建立預約寄送");
}

/**
 * 對話框透過 google.script.run 呼叫的 server 端入口。
 *
 * @param {string} subjectLine
 * @param {string} scheduledTimeStr "YYYY-MM-DDTHH:mm"，由對話框 datetime-local 欄位傳回
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
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet"
        href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500&family=Roboto:wght@400;500&family=Roboto+Mono&display=swap">
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Roboto', 'Google Sans', 'PingFang TC',
                   'Microsoft JhengHei', -apple-system, sans-serif;
      margin: 0; padding: 24px; background: #fff;
      color: #202124; font-size: 14px; line-height: 1.5;
      -webkit-font-smoothing: antialiased;
    }

    .field { margin-bottom: 20px; }
    .field-label {
      display: block; font-size: 12px; font-weight: 500;
      color: #5f6368; margin-bottom: 6px; letter-spacing: 0.3px;
    }
    .field input {
      display: block; width: 100%; height: 40px;
      padding: 8px 12px; font-size: 14px; font-family: inherit;
      color: #202124; background: #fff;
      border: 1px solid #dadce0; border-radius: 4px;
      outline: none; transition: border-color .15s, box-shadow .15s;
    }
    .field input:hover { border-color: #80868b; }
    .field input:focus {
      border-color: #1a73e8;
      box-shadow: inset 0 0 0 1px #1a73e8;
    }
    .field-hint {
      font-size: 12px; color: #5f6368; margin-top: 6px; line-height: 1.45;
    }
    .field-hint code {
      font-family: 'Roboto Mono', monospace;
      background: #f1f3f4; color: #202124;
      padding: 1px 5px; border-radius: 3px; font-size: 11px;
    }

    .actions {
      display: flex; justify-content: flex-end; gap: 8px;
      margin-top: 8px;
    }
    button {
      font-family: 'Google Sans', 'Roboto', sans-serif;
      font-size: 14px; font-weight: 500; letter-spacing: 0.25px;
      height: 36px; padding: 0 24px;
      border: none; border-radius: 4px; cursor: pointer;
      transition: background .15s, box-shadow .15s;
    }
    button.text { background: transparent; color: #1a73e8; padding: 0 16px; }
    button.text:hover { background: rgba(26,115,232,.08); }
    button.text:active { background: rgba(26,115,232,.16); }

    button.primary { background: #1a73e8; color: #fff; }
    button.primary:hover {
      background: #1765cc;
      box-shadow: 0 1px 2px 0 rgba(60,64,67,.3),
                  0 1px 3px 1px rgba(60,64,67,.15);
    }
    button.primary:disabled {
      background: #f1f3f4; color: #80868b;
      cursor: not-allowed; box-shadow: none;
    }

    .alert {
      margin-top: 16px; padding: 12px 16px;
      font-size: 13px; line-height: 1.5; white-space: pre-line;
      border-radius: 4px;
    }
    .alert.ok    { background: #e6f4ea; color: #137333; }
    .alert.error { background: #fce8e6; color: #c5221f; }
    .alert:empty { display: none; }
  </style>
</head>
<body>
  <form id="form">
    <div class="field">
      <label class="field-label" for="subject">草稿主旨</label>
      <input id="subject" name="subject" type="text" required autofocus />
      <div class="field-hint">
        與 Gmail 草稿欄相符，可含 <code>{{欄位名}}</code> 變數。
      </div>
    </div>
    <div class="field">
      <label class="field-label" for="time">統一預定寄送時間</label>
      <input id="time" name="time" type="datetime-local"
             value="${isoLocal}" required />
      <div class="field-hint">
        預設為下一個未到的週五 ${TRIGGER_HOUR}:00（${human}），可改成任何時間。
      </div>
    </div>
    <div class="actions">
      <button type="button" class="text"
              onclick="google.script.host.close()">關閉</button>
      <button type="submit" class="primary" id="submit">建立預約</button>
    </div>
  </form>
  <div id="result" class="alert"></div>
  <script>
    const form = document.getElementById('form');
    const submitBtn = document.getElementById('submit');
    const result = document.getElementById('result');

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(form);
      submitBtn.disabled = true;
      submitBtn.textContent = '建立中…';
      result.className = 'alert';
      result.textContent = '';

      google.script.run
        .withSuccessHandler((r) => {
          result.className = 'alert ' + (r.errors.length ? 'error' : 'ok');
          result.textContent =
            '已建立 ' + r.successCount + ' 封預約草稿' +
            (r.errors.length ? '\\n\\n錯誤：\\n' + r.errors.join('\\n') : '');
          submitBtn.disabled = false;
          submitBtn.textContent = '建立預約';
        })
        .withFailureHandler((err) => {
          result.className = 'alert error';
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
