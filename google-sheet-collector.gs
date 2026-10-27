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
 * Each finished survey becomes one row on the "Responses" tab, in the fixed
 * column order below. If the survey questions change, update COLUMNS and run
 * rebuild() once from the editor to re-lay-out existing rows (the old layout
 * is copied to a dated "Raw backup" tab first).
 */
var SHEET_ID = '1UMbkkVFPL-ZlbD9Y2fKVr9kN5XBhMtIO9Gc6UtpotbA';
var TAB = 'Responses';
var DATE_FORMAT = 'd mmm yyyy, h:mm am/pm';

// [payload key, column heading, column width]
var COLUMNS = [
  ['submitted_at', 'Submitted at', 150],
  ['email', 'Email', 220],
  ['q1_last_used', 'Q1 · Last put on lip balm', 170],
  ['q1_last_used_other', 'Q1 · Other (typed)', 130],
  ['q2_brand_on_hand', 'Q2 · Lip balm on them right now', 220],
  ['q2_brand_on_hand_other', 'Q2 · Other (typed)', 130],
  ['q3_where_buy', 'Q3 · Where they usually buy', 220],
  ['q3_where_buy_other', 'Q3 · Other (typed)', 130],
  ['q4_last_paid', 'Q4 · Paid for last one', 140],
  ['q4_last_paid_other', 'Q4 · Other (typed)', 130],
  ['q5_why_stopped', 'Q5 · Why they stopped the last one', 240],
  ['q5_why_stopped_other', 'Q5 · Other (typed)', 130],
  ['q6_how_long_lasts', 'Q6 · How long one lasts', 170],
  ['q6_how_long_lasts_other', 'Q6 · Other (typed)', 130],
  ['q7_when_most', 'Q7 · When they use it most', 200],
  ['q7_when_most_other', 'Q7 · Other (typed)', 130],
  ['q8_bought_6mo', 'Q8 · Bought in last 6 months', 220],
  ['q8_bought_6mo_other', 'Q8 · Other (typed)', 130],
  ['q9_what_mattered', 'Q9 · What mattered most', 260],
  ['q9_what_mattered_other', 'Q9 · Other (typed)', 130],
  ['q10_age', 'Age', 80],
  ['q10_gender', 'Gender', 110],
  ['q10_city', 'City', 130],
  ['duration_sec', 'Time taken (sec)', 100],
  ['utm_source', 'UTM source', 110],
  ['utm_medium', 'UTM medium', 110],
  ['utm_campaign', 'UTM campaign', 120],
  ['referrer', 'Referrer', 180],
  ['device', 'Device', 90],
  ['response_id', 'Response ID', 140]
];

// Keys used by earlier versions of the survey → their current key.
var RENAMED = {
  q3_last_bought_where: 'q3_where_buy', q3_last_bought_where_other: 'q3_where_buy_other',
  q6_bought_6mo: 'q8_bought_6mo', q6_bought_6mo_other: 'q8_bought_6mo_other',
  q9_age: 'q10_age', q9_gender: 'q10_gender', q9_city: 'q10_city'
};

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = normalise(JSON.parse(e.postData.contents));
    var sheet = responsesSheet();
    var keys = headerKeys(sheet);
    Object.keys(data).forEach(function (k) {
      if (keys.indexOf(k) === -1) { keys.push(k); sheet.getRange(1, keys.length).setValue(k); }
    });
    sheet.appendRow(keys.map(function (k) { return cell(k, data[k]); }));
    var r = sheet.getLastRow();
    sheet.getRange(r, 1, 1, keys.length).setWrap(true).setVerticalAlignment('top');
    sheet.getRange(r, keys.indexOf('submitted_at') + 1).setNumberFormat(DATE_FORMAT);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

/** Run once from the editor: backs up the current tab, then rewrites it in the COLUMNS layout. */
function rebuild() {
  var sheet = responsesSheet();
  var values = sheet.getDataRange().getValues();
  var ss = sheet.getParent();
  var name = 'Raw backup ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'd MMM HH:mm');
  var backup = ss.insertSheet(name, ss.getSheets().length);
  backup.getRange(1, 1, values.length, values[0].length).setValues(values);

  var head = values[0].map(keyFor);
  var rows = values.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    var o = {};
    head.forEach(function (k, i) { if (r[i] !== '') o[k] = r[i]; });
    return normalise(o);
  });

  sheet.clear();
  sheet.clearFormats();
  var keys = headerKeys(sheet);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, keys.length).setValues(rows.map(function (o) {
      return keys.map(function (k) { return cell(k, o[k]); });
    }));
  }
}

function responsesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB);
  if (!sheet) { sheet = ss.getSheets()[0]; sheet.setName(TAB); }
  return sheet;
}

// Returns the payload key for each column, writing and styling the header row if the tab is empty.
function headerKeys(sheet) {
  if (sheet.getLastRow() > 0) return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(keyFor);
  var n = COLUMNS.length, rows = sheet.getMaxRows();
  sheet.getRange(1, 1, 1, n).setValues([COLUMNS.map(function (c) { return c[1]; })])
    .setFontWeight('bold').setBackground('#2b1a14').setFontColor('#f6eedf')
    .setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 44);
  sheet.getRange(2, 1, rows - 1, n).setWrap(true).setVerticalAlignment('top');
  sheet.getRange(2, 1, rows - 1, 1).setNumberFormat(DATE_FORMAT);
  COLUMNS.forEach(function (c, i) { sheet.setColumnWidth(i + 1, c[2]); });
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  return COLUMNS.map(function (c) { return c[0]; });
}

function keyFor(heading) {
  for (var i = 0; i < COLUMNS.length; i++) if (COLUMNS[i][1] === heading) return COLUMNS[i][0];
  return RENAMED[heading] || heading;
}

function normalise(data) {
  var out = {};
  Object.keys(data).forEach(function (k) {
    var nk = RENAMED[k] || k;
    if (out[nk] === undefined || out[nk] === '') out[nk] = data[k];
  });
  return out;
}

function cell(key, v) {
  if (v === undefined || v === null) return '';
  if (key === 'submitted_at' && !(v instanceof Date) && v !== '') { var d = new Date(v); return isNaN(d) ? v : d; }
  return v;
}
