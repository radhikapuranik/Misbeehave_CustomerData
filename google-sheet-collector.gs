/**
 * Misbeehave survey → Google Sheet
 *
 * 1. Standalone Apps Script project (script.google.com) that writes to SHEET_ID below.
 * 2. Delete what's there, paste this whole file, click Save.
 * 3. Deploy → New deployment → type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, approve the permissions, copy the Web app URL.
 * 4. In index.html, paste that URL into RESPONSES_ENDPOINT = '...'
 *
 * Each finished survey becomes one row. Headers are created automatically.
 */
var SHEET_ID = '1UMbkkVFPL-ZlbD9Y2fKVr9kN5XBhMtIO9Gc6UtpotbA';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var keys = Object.keys(data);

    var headers = sheet.getLastRow() === 0 ? [] : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    keys.forEach(function (k) { if (headers.indexOf(k) === -1) headers.push(k); });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    var row = headers.map(function (h) { return data[h] !== undefined ? data[h] : ''; });
    sheet.appendRow(row);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
