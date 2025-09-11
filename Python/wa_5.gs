/** Google Apps Script: Web App for BOQ uploads + image previews.
 * Deploy: Deploy > Manage deployments > New deployment > "Web app"
 * - Execute as: Me
 * - Who has access: Anyone
 */

function doPost(e) {
  try {
    var body = e.postData && e.postData.contents ? e.postData.contents : "{}";
    var p = JSON.parse(body);

    var ss   = SpreadsheetApp.openById(p.sheetId);
    var tab  = String(p.tab || 'Detail');
    var mode = String(p.mode || 'replace').toLowerCase(); // 'replace'|'append'
    var sh   = ss.getSheetByName(tab) || ss.insertSheet(tab);

    var headers = Array.isArray(p.headers) ? p.headers : [];
    var rows    = Array.isArray(p.rows)    ? p.rows    : [];
    var images  = Array.isArray(p.images)  ? p.images  : []; // base64 PNGs aligned to rows
    var vAlign  = String(p.vAlign || "");           // "middle" or ""
    var sparse  = String(p.sparseAnchor || "last"); // "first"|"last"|"middle"
    var runId   = String(p.runId || "run");
    var driveFolderId = String(p.driveFolderId || ""); // optional

    if (mode === 'replace') {
      sh.clearContents();
      if (headers.length) {
        sh.getRange(1,1,1,headers.length).setValues([headers]);
      }
    }
    // Append rows
    var startRow = sh.getLastRow() + 1;
    if (rows.length) {
      sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }

    // ----- Display polish + category merges -----
    var hasHeader = (sh.getLastRow() > 0 && headers.length > 0);
    var firstDataRow = hasHeader ? 2 : 1;
    var lastRow = sh.getLastRow();
    if (lastRow >= firstDataRow) {
      var rngAll = sh.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, sh.getLastColumn());
      rngAll.setHorizontalAlignment('center');
      if (vAlign === 'middle') rngAll.setVerticalAlignment('middle');

      // Merge contiguous identical values in Category column (E = 5)
      mergeContiguousColumn_(sh, 5, firstDataRow, lastRow, sparse);
    }

    // ----- Preview images -----
    // Find the "Preview" column index from headers (fallback to 11)
    var previewCol = 11;
    if (headers && headers.length) {
      var idx = headers.indexOf("Preview");
      if (idx >= 0) previewCol = idx + 1;
    }

    if (rows.length && previewCol) {
      var folder = getOrCreateFolder_(driveFolderId);
      for (var i = 0; i < rows.length; i++) {
        var b64 = images[i] || "";
        if (!b64) continue;

        // Create file in Drive
        var fileName = (runId + "_" + (startRow + i)).replace(/[^\w\-\.]/g, "_") + ".png";
        var blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
        var file = folder.createFile(blob);
        // Shareable for IMAGE() formula
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (err) {
          // ignore if restricted domain; IMAGE will still show for permitted viewers
        }
        var fileId = file.getId();
        var publicUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

        var cell = sh.getRange(startRow + i, previewCol);
        cell.setFormula('=IMAGE("' + publicUrl + '")');
        cell.setHorizontalAlignment('center').setVerticalAlignment('middle');
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
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
      } else { // 'last' default
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

/** Get folder by ID or create/find "DXF-Previews" at root. */
function getOrCreateFolder_(folderId) {
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  var name = "DXF-Previews";
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}
