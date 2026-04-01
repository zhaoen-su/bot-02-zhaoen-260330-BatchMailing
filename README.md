# BatchMailing — 批次寄信工具

Google Sheets 批次寄信腳本，透過 Google Apps Script 讀取試算表中的收件人資料，搭配 Gmail 草稿作為信件模板，一鍵完成大量郵件寄送。

> 雖然 repo 名稱有 "bot"，但這個工具不是自動化機器人，操作皆需由使用者手動觸發。

## 功能特色

- 從 Gmail 草稿取得信件模板（支援 HTML 格式、附件、內嵌圖片）
- 使用 `{{欄位名稱}}` 語法在模板中插入個人化內容
- 自動跳過已寄出的列，避免重複寄信
- 寄送完成後在試算表中記錄寄出時間，寄送失敗則記錄錯誤訊息

## 使用方式

### 1. 準備試算表

在 Google Sheets 中建立資料表，表頭須包含以下欄位：

| 欄位名稱 | 說明 |
|---------|------|
| `收件人` | 收件者的 Email 地址 |
| `已寄出` | 寄送狀態（留空表示尚未寄出，腳本會自動填入寄出時間） |

其餘欄位可自由新增，欄位名稱即可作為模板中的變數名稱。

### 2. 撰寫 Gmail 草稿

在 Gmail 中建立一封草稿，信件主旨與內文皆可使用 `{{欄位名稱}}` 語法插入變數。例如：

```
您好 {{姓名}}，

感謝您參加 {{活動名稱}}，以下是您的報名資訊...
```

草稿可包含 HTML 格式、附件及內嵌圖片，皆會完整保留。

### 3. 執行寄信

1. 開啟試算表，上方選單會出現 **批次寄信** 選項
2. 點選 **批次寄信 → 送出**
3. 在彈出的對話框中輸入（或貼上）Gmail 草稿的**主旨**
4. 腳本會自動比對草稿、逐列寄出郵件，並在「已寄出」欄記錄時間

## 部署方式

### 方法一：Apps Script 編輯器

1. 開啟目標 Google Sheet
2. 選擇 **擴充功能 → Apps Script**
3. 將程式碼內容貼入編輯器，儲存即可

### 方法二：使用 clasp

```bash
# 安裝 clasp
npm install -g @google/clasp

# 登入 Google 帳號
clasp login

# 推送至已綁定的 Apps Script 專案
clasp push
```

## 注意事項

- 本腳本使用 `@OnlyCurrentDoc` 註解，OAuth 權限僅限於綁定的試算表
- Gmail 每日寄信數量有上限（一般帳號 500 封、Workspace 帳號 2000 封），請留意用量
- 如需寄送含 Unicode/Emoji 的郵件，可將程式碼中的 `GmailApp.sendEmail` 改為 `MailApp.sendEmail`

## 授權來源

Original code by Martin Hawksey, licensed under Apache License 2.0
https://developers.google.com/apps-script/samples/automations/mail-merge
Modified by CAPSULE CXO TEAM