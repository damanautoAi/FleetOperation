/***************************************************************************************
 * FLEET OPERATIONS — Google Sheets connector (Web App)
 * ------------------------------------------------------------------------------------
 * HOW TO DEPLOY (one time):
 *   1. Open your Fleet Google Sheet in the browser.
 *   2. Menu:  Extensions  ->  Apps Script.
 *   3. Delete any sample code, paste THIS whole file, click the disk icon (Save).
 *   4. Click  Deploy  ->  New deployment.
 *   5. Gear icon -> choose type  "Web app".
 *   6. Description: Fleet.  Execute as:  Me.   Who has access:  Anyone.
 *   7. Click Deploy, allow/authorize the permissions it asks for.
 *   8. Copy the "Web app URL" (ends with /exec).
 *   9. In the website click the gear (Settings) and paste that URL.  Done.
 *
 *  When you change this script later, Deploy -> Manage deployments -> edit -> Version:
 *  "New version" -> Deploy (the /exec URL stays the same).
 ***************************************************************************************/

var APP_VERSION = 6;   // v3 head, v4 colors, v5 fast list, v6 dashboard date edits

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  var p  = (e && e.parameter) ? e.parameter : {};
  var cb = p.callback;
  // POST body (text/plain JSON) overrides query params when present
  if (e && e.postData && e.postData.contents) {
    try { var body = JSON.parse(e.postData.contents); for (var k in body) p[k] = body[k]; } catch (ignore) {}
  }
  var out;
  try {
    var action = p.action || 'list';
    if      (action === 'ping')   out = { version: APP_VERSION };
    else if (action === 'list')   out = { version: APP_VERSION, sheets: listSheets() };
    else if (action === 'meta')   out = meta();
    else if (action === 'sheet')  out = getSheet(p.name, p.limit ? Number(p.limit) : 0, p.head ? Number(p.head) : 0, p.fmt === '1' || p.fmt === 1);
    else if (action === 'add')    out = addRow(p.sheet, parseRow(p.row));
    else if (action === 'update') out = updateCell(p.sheet, Number(p.row), Number(p.col), p.val);
    else if (action === 'delete') out = deleteRow(p.sheet, Number(p.row));
    else throw 'Unknown action: ' + action;
    return reply({ ok: true, data: out }, cb);
  } catch (err) {
    return reply({ ok: false, error: String(err) }, cb);
  }
}

function reply(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function parseRow(r) { return (typeof r === 'string') ? JSON.parse(r) : r; }

/* Sheets that must never be written to (formula / summary tabs). */
var PROTECTED = ['Dashboard'];

function listSheets() {
  // Lightweight & fast: only names/flags. Row counts are read lazily when a sheet
  // is opened (getSheet returns totalRows). Calling getLastRow/getLastColumn for
  // every sheet here was very slow on formula-heavy workbooks and caused timeouts.
  return ss().getSheets().map(function (s) {
    return { name: s.getName(), hidden: s.isSheetHidden(), locked: PROTECTED.indexOf(s.getName()) >= 0 };
  });
}

function meta() {
  // cheap fingerprint per sheet so the website knows when to refresh
  return ss().getSheets().map(function (s) {
    return { name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn() };
  });
}

function getSheet(name, limit, head, fmt) {
  var sh = ss().getSheetByName(name);
  if (!sh) throw 'No sheet named: ' + name;
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return { name: name, values: [], startRow: 1 };

  // tail mode (logs) — contiguous-from-top not required, no formatting payload.
  if (!head && !fmt && limit && lastRow > limit) {
    var header = sh.getRange(1, 1, 1, lastCol).getDisplayValues();
    var tail   = sh.getRange(lastRow - limit + 1, 1, limit, lastCol).getDisplayValues();
    return { name: name, values: header.concat(tail), startRow: 1, truncated: true, totalRows: lastRow };
  }

  // contiguous-from-top block (full, or first `head` rows)
  var numRows = (head && head > 0) ? Math.min(head, lastRow) : lastRow;
  var rng = sh.getRange(1, 1, numRows, lastCol);
  var res = { name: name, values: rng.getDisplayValues(), startRow: 1, totalRows: lastRow, truncated: numRows < lastRow };

  if (fmt) {
    try { res.bg = rng.getBackgrounds(); } catch (e) { res.bg = []; }
    try { res.fc = rng.getFontColors(); } catch (e) { res.fc = []; }
    try { res.fw = rng.getFontWeights(); } catch (e) { res.fw = []; }
    res.merges = [];
    try {
      var mr = sh.getMergedRanges();
      for (var i = 0; i < mr.length; i++) {
        var m = mr[i], r0 = m.getRow(), c0 = m.getColumn();
        if (r0 <= numRows && c0 <= lastCol) {
          res.merges.push({ r: r0 - 1, c: c0 - 1, nr: m.getNumRows(), nc: m.getNumColumns() });
        }
      }
    } catch (e) {}
    res.dv = {};
    try {
      var dvs = rng.getDataValidations();
      var rangeCache = {};   // cache range-based lists so we read each source range only ONCE
      for (var rr = 0; rr < dvs.length; rr++) {
        for (var cc = 0; cc < dvs[rr].length; cc++) {
          var dv = dvs[rr][cc]; if (!dv) continue;
          var t = dv.getCriteriaType(), cv = dv.getCriteriaValues(), list = null;
          if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) { list = cv[0]; }
          else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && cv[0]) {
            var srcR = cv[0];
            var keyR = srcR.getSheet().getName() + '!' + srcR.getA1Notation();
            if (rangeCache[keyR]) { list = rangeCache[keyR]; }
            else {
              list = []; var rv = srcR.getDisplayValues();
              for (var a = 0; a < rv.length; a++) for (var b = 0; b < rv[a].length; b++) { var x = rv[a][b]; if (String(x).trim() !== '') list.push(x); }
              rangeCache[keyR] = list;
            }
          }
          if (list && list.length) res.dv[rr + '_' + cc] = list;
        }
      }
    } catch (e) {}
  }
  return res;
}

function addRow(name, rowArr) {
  guard(name);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ss().getSheetByName(name);
    if (!sh) throw 'No sheet: ' + name;
    sh.appendRow(rowArr);
    SpreadsheetApp.flush();
    return { row: sh.getLastRow() };
  } finally { lock.releaseLock(); }
}

function updateCell(name, row, col, val) {
  // single-cell edits are allowed on any sheet (incl. Dashboard) — e.g. the
  // E-Waybill date picker. (add/delete rows are still guarded.)
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ss().getSheetByName(name);
    if (!sh) throw 'No sheet: ' + name;
    // ISO date string ("YYYY-MM-DD") -> real Date so the cell stays a date
    if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(val)) {
      var p = val.split('-');
      sh.getRange(row, col).setValue(new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
    } else {
      sh.getRange(row, col).setValue(val);
    }
    SpreadsheetApp.flush();
    return { row: row, col: col };
  } finally { lock.releaseLock(); }
}

function deleteRow(name, row) {
  guard(name);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ss().getSheetByName(name);
    if (!sh) throw 'No sheet: ' + name;
    sh.deleteRow(row);
    SpreadsheetApp.flush();
    return { deleted: row };
  } finally { lock.releaseLock(); }
}

function guard(name) {
  if (PROTECTED.indexOf(name) >= 0) throw 'Sheet "' + name + '" is read-only (auto-calculated).';
}
