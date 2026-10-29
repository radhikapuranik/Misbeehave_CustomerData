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
 * column order below. A typed "Other" answer is folded into its question's
 * column (e.g. "Nykaa; Other: Airport shop") rather than getting its own.
 * If the survey questions change, update COLUMNS and run rebuild() once from
 * the editor to re-lay-out existing rows (the old layout is copied to a dated
 * "Raw backup" tab first).
 */
var SHEET_ID = '1UMbkkVFPL-ZlbD9Y2fKVr9kN5XBhMtIO9Gc6UtpotbA';
var TAB = 'Responses';
var DATE_FORMAT = 'd mmm yyyy, h:mm am/pm';
var HEADER_BG = '#2b1a14', ROW_BG = '#ffffff', ALT_ROW_BG = '#f1f1f1';

// [payload key, column heading, column width]
var COLUMNS = [
  ['date', 'Date', 150],
  ['name', 'Name', 160],
  ['email', 'Email', 220],
  ['phone', 'Phone', 130],
  ['q1_last_used', 'Q1 · Last put on lip balm', 170],
  ['q2_brand_on_hand', 'Q2 · Lip balm on them right now', 220],
  ['q3_where_buy', 'Q3 · Where they usually buy', 220],
  ['q4_last_paid', 'Q4 · Paid for last one', 140],
  ['q5_why_stopped', 'Q5 · Why they stopped the last one', 240],
  ['q6_how_long_lasts', 'Q6 · How long one lasts', 170],
  ['q7_when_most', 'Q7 · When they use it most', 200],
  ['q8_bought_6mo', 'Q8 · Bought in last 6 months', 220],
  ['q9_what_mattered', 'Q9 · What mattered most', 260],
  ['q10_age', 'Age', 80],
  ['q10_gender', 'Gender', 110],
  ['q10_city', 'City', 130],
  ['utm_source', 'UTM source', 110],
  ['utm_medium', 'UTM medium', 110],
  ['utm_campaign', 'UTM campaign', 120],
  ['referrer', 'Referrer', 180],
  ['device', 'Device', 90],
  ['duration_sec', 'Time taken (sec)', 100],
  ['response_id', 'Response ID', 140]
];

// Keys used by earlier versions of the survey → their current key.
var RENAMED = {
  submitted_at: 'date', 'Submitted at': 'date',
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
    sheet.getRange(r, keys.indexOf('date') + 1).setNumberFormat(DATE_FORMAT);
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
    .setFontWeight('bold').setFontColor('#f6eedf')
    .setWrap(true).setVerticalAlignment('middle');
  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(1, 1, rows, n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    .setHeaderRowColor(HEADER_BG).setFirstRowColor(ROW_BG).setSecondRowColor(ALT_ROW_BG);
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
  // Older layouts had a separate "Q3 · Other (typed)" column per question.
  var m = /^(Q\d+) · Other \(typed\)$/.exec(heading);
  if (m) for (var j = 0; j < COLUMNS.length; j++) if (COLUMNS[j][1].indexOf(m[1] + ' · ') === 0) return COLUMNS[j][0] + '_other';
  return RENAMED[heading] || heading;
}

function normalise(data) {
  var out = {};
  Object.keys(data).forEach(function (k) {
    var nk = RENAMED[k] || k;
    if (out[nk] === undefined || out[nk] === '') out[nk] = data[k];
  });
  return mergeOther(out);
}

// Folds each "<question>_other" answer into its question, replacing the "Other" choice with "Other: <typed text>".
function mergeOther(data) {
  Object.keys(data).forEach(function (k) {
    var m = /^(.*)_other$/.exec(k);
    if (!m) return;
    var typed = String(data[k] || '').trim();
    if (typed) {
      var parts = String(data[m[1]] || '').split('; ').filter(Boolean);
      var i = parts.indexOf('Other');
      if (i > -1) parts[i] = 'Other: ' + typed; else parts.push('Other: ' + typed);
      data[m[1]] = parts.join('; ');
    }
    delete data[k];
  });
  return data;
}

function cell(key, v) {
  if (v === undefined || v === null) return '';
  if (key === 'date' && !(v instanceof Date) && v !== '') {
    // The survey sends Indian time as "YYYY-MM-DD HH:mm"; older rows used ISO timestamps.
    var s = String(v), d = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s) ? new Date(s.replace(' ', 'T') + ':00+05:30') : new Date(s);
    return isNaN(d) ? v : d;
  }
  // Leading apostrophe keeps "+91…" as text instead of a number.
  if (key === 'phone' && v !== '') return "'" + v;
  return v;
}
