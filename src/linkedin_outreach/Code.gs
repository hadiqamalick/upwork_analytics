/**
 * LinkedIn Outreach Dashboard — server side.
 * Strictly read-only: reads the clean_leads tab, never writes to the sheet.
 * Files: Code.gs (server), Index.html (shell), css.html (styles), JavaScript.html (client logic).
 */

var SHEET_ID = '1cCpQCph1Dk0lW-HyZgbyO-SivpqApUgBbcn9pYYCg2g';
var TAB_NAME = 'clean_leads';

/** Serves the dashboard shell. */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('LinkedIn Outreach Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Standard include helper so Index.html can pull in css.html and JavaScript.html. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Reads and normalizes clean_leads.
 * - Rows with blank prospect_name are padding and excluded.
 * - Date columns may hold real Date cells or "yyyy-MM-dd" strings; both are
 *   normalized to "yyyy-MM-dd" strings (spreadsheet timezone).
 * - All text values are trimmed. moved_to_hubspot is normalized to boolean.
 */
function getLeads() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME);
  if (!sh) throw new Error('Tab "' + TAB_NAME + '" not found in the spreadsheet.');

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { rows: [], fetchedAt: new Date().toISOString() };

  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var DATE_COLS = ['date_added', 'request_sent_date', 'accepted_date',
                   'first_message_date', 'first_reply_date'];
  var TEXT_COLS = ['prospect_name', 'linkedin_url', 'company', 'title_designation',
                   'email', 'email_status', 'country', 'bd_name', 'profile_used',
                   'status', 'reply_status', 'notes'];

var tz = ss.getSpreadsheetTimeZone() || 'Etc/UTC';

  function normDate(v) {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    }
    var s = String(v == null ? '' : v).trim();
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[1] + '-' + m[2] + '-' + m[3] : '';
  }

  function normText(v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return String(v == null ? '' : v).trim();
  }

  function normBool(v) {
    if (v === true) return true;
    var s = String(v == null ? '' : v).trim().toLowerCase();
    return s === 'true' || s === 'yes' || s === 'y' || s === '1';
  }

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var raw = values[r];
    var name = normText(idx.prospect_name != null ? raw[idx.prospect_name] : '');
    if (!name) continue; // padding row

    var row = {};
    TEXT_COLS.forEach(function (c) {
      row[c] = idx[c] != null ? normText(raw[idx[c]]) : '';
    });
    DATE_COLS.forEach(function (c) {
      row[c] = idx[c] != null ? normDate(raw[idx[c]]) : '';
    });
    row.moved_to_hubspot = idx.moved_to_hubspot != null ? normBool(raw[idx.moved_to_hubspot]) : false;
    rows.push(row);
  }

  return { rows: rows, fetchedAt: new Date().toISOString() };
}