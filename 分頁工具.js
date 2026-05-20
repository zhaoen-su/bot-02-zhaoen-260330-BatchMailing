/**
 * 確保三個分頁都存在，且第一列是預期的欄位標題。
 *
 * - 不存在的分頁會新建並寫入標頭。
 * - 已存在但缺欄位的分頁，會把缺的欄位補在最後一欄（不刪除既有資料）。
 * - 主要操作區只確保有「收件者」欄；其他自訂欄位（姓名、公司等）由
 *   使用者自行加入，這裡不動。
 */
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss, SHEET_MAIN, [RECIPIENT_COL]);
  ensureSheetWithHeaders_(ss, SHEET_SCHEDULED, [
    SCH_COL_RECIPIENT,
    SCH_COL_SUBJECT,
    SCH_COL_DRAFT_ID,
    SCH_COL_SCHEDULED_AT,
    SCH_COL_CREATED_AT,
    SCH_COL_OWNER,
  ]);
  ensureSheetWithHeaders_(ss, SHEET_SENT, [
    SENT_COL_RECIPIENT,
    SENT_COL_SUBJECT,
    SENT_COL_SCHEDULED_AT,
    SENT_COL_SENT_AT,
    SENT_COL_RESULT,
  ]);
}

function ensureSheetWithHeaders_(ss, name, requiredHeaders) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet
      .getRange(1, 1, 1, requiredHeaders.length)
      .setValues([requiredHeaders])
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }
  const lastCol = sheet.getLastColumn();
  const existing =
    lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const missing = requiredHeaders.filter((h) => existing.indexOf(h) === -1);
  if (missing.length > 0) {
    sheet
      .getRange(1, lastCol + 1, 1, missing.length)
      .setValues([missing])
      .setFontWeight("bold");
    if (sheet.getFrozenRows() === 0) sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 把分頁第一列讀成 { 欄位名: columnIndex(從1起算) } 的 map。
 * 對於 appendRow 之外的精準寫入比較好用。
 */
function getHeaderIndexMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h !== "" && map[h] === undefined) map[h] = i + 1;
  });
  return map;
}

/**
 * 從一個 row 物件（key 是欄位名）依照 sheet 的標頭順序組成 array，
 * 沒有對應 key 的欄位寫空字串。給 appendRow 用。
 */
function buildRowForSheet_(sheet, rowObj) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headers.map((h) => (rowObj[h] !== undefined ? rowObj[h] : ""));
}
