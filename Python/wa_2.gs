/** Google Apps Script: Web App for BOQ uploads.
 * Deploy: Deploy > Manage deployments > New deployment > "Web app"
 * - Execute as: Me
 * - Who has access: Anyone
 */
function doPost(e) {
  try {
    var body = e.postData && e.postData.contents ? e.postData.contents : "{}";
    var p = JSON.parse(body);

    var ss    = SpreadsheetApp.openById(p.sheetId);
    var tab   = String(p.tab || 'Detail');
    var mode  = String(p.mode || 'replace').toLowerCase(); // 'replace'|'append'
    var sh    = ss.getSheetByName(tab) || ss.insertSheet(tab);

    var headers = Array.isArray(p.headers) ? p.headers : [];
    var rows    = Array.isArray(p.rows)    ? p.rows    : [];
    var vAlign  = String(p.vAlign || "");           // "middle" or ""
    var sparse  = String(p.sparseAnchor || "last"); // "first"|"last"|"middle"

    if (mode === 'replace') {
      sh.clearContents();
      if (headers.length) {
        sh.getRange(1,1,1,headers.length).setValues([headers]);
      }
    }
    // Append rows below last row (taking into account header if present)
    var startRow = sh.getLastRow() + 1;
    if (rows.length) {
      sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }

    // --- Display polish: center & merge contiguous categories in column E ---
    var hasHeader = (headers.length > 0) && (sh.getRange(1,1,1,headers.length).getDisplayValue().trim() !== "");
    var firstDataRow = hasHeader ? 2 : 1;
    var lastRow = sh.getLastRow();
    if (lastRow >= firstDataRow) {
      // Center everything horizontally; vertical center if requested
      var rngAll = sh.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, sh.getLastColumn());
      rngAll.setHorizontalAlignment('center');
      if (vAlign === 'middle') rngAll.setVerticalAlignment('middle');

      // Merge contiguous identical values in Category column (E = index 5)
      mergeContiguousColumn_(sh, 5, firstDataRow, lastRow, sparse);
    }

    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

/** Merge contiguous identical cells in a given column.
 * @param {Sheet} sh
 * @param {number} colIndex 1-based column index (5 = E)
 * @param {number} r1 first data row
 * @param {number} rN last data row
 * @param {string} anchor 'first'|'last'|'middle' (which cell keeps the label)
 */
function mergeContiguousColumn_(sh, colIndex, r1, rN, anchor) {
  if (rN < r1) return;
  var values = sh.getRange(r1, colIndex, rN - r1 + 1, 1).getValues().map(function(r){ return sanitize_(r[0]); });

  var runStart = 0; // index within values
  while (runStart < values.length) {
    var v = values[runStart];
    if (!v) { runStart++; continue; }
    var runEnd = runStart + 1;
    while (runEnd < values.length && values[runEnd] === v) runEnd++;

    // only merge if run length >= 2
    var len = runEnd - runStart;
    if (len >= 2) {
      var a = r1 + runStart;
      var b = r1 + runEnd - 1;
      // set just one anchor label
      if (anchor === 'first') {
        sh.getRange(a, colIndex, len, 1).clearContent();
        sh.getRange(a, colIndex).setValue(v);
      } else if (anchor === 'middle') {
        var mid = Math.floor((a + b) / 2);
        sh.getRange(a, colIndex, len, 1).clearContent();
        sh.getRange(mid, colIndex).setValue(v);
      } else { // 'last' (default)
        sh.getRange(a, colIndex, len, 1).clearContent();
        sh.getRange(b, colIndex).setValue(v);
      }
      // merge & center
      sh.getRange(a, colIndex, len, 1).merge();
      sh.getRange(a, colIndex).setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
    runStart = runEnd;
  }
}

function sanitize_(s) {
  if (s == null) return "";
  s = String(s).replace(/\u00A0/g, ' '); // NBSP → space
  s = s.replace(/\s+/g, ' ').trim();
  return s.toUpperCase(); // case-insensitive grouping
}
