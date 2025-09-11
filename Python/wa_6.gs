/** Web App: BOQ uploads + preview images (Drive URLs or direct embed)
 * Payload JSON (POST):
 * {
 *   sheetId: string,
 *   tab: string,                 // e.g. "Sheet9"
 *   mode: "replace"|"append",    // default: "replace"
 *   headers: string[],           // header row (sent on replace)
 *   rows: any[][],               // data rows
 *   images: string[],            // base64 PNG per row; "" for no image
 *   vAlign: ""|"middle",         // optional vertical alignment
 *   sparseAnchor: "first"|"last"|"middle", // label placement in merged Category col
 *   runId: string,               // used for file naming
 *   driveFolderId: string,       // optional; else auto "DXF-Previews"
 *   embedImages: boolean,        // OPTIONAL: true = insert PNG objects; false = =IMAGE(url)
 *   imageMode: 1|2|3|4,          // OPTIONAL for =IMAGE; default 1
 *   imageW: number,              // OPTIONAL for mode=3|4; default 28
 *   imageH: number               // OPTIONAL for mode=3|4; default 28
 * }
 */

function doPost(e) {
  try {
    var body = e.postData && e.postData.contents ? e.postData.contents : "{}";
    var p = JSON.parse(body);

    var ss  = SpreadsheetApp.openById(p.sheetId);
    var tab = String(p.tab || "Detail");
    var sh  = ss.getSheetByName(tab) || ss.insertSheet(tab);

    var mode   = String(p.mode || "replace").toLowerCase();
    var headers = Array.isArray(p.headers) ? p.headers : [];
    var rows    = Array.isArray(p.rows)    ? p.rows    : [];
    var images  = Array.isArray(p.images)  ? p.images  : [];
    var vAlign  = String(p.vAlign || "");
    var sparse  = String(p.sparseAnchor || "last");
    var runId   = String(p.runId || "run");
    var folderId = String(p.driveFolderId || "");
    var embed   = !!p.embedImages;

    // Optional =IMAGE controls (used only when embedImages === false)
    var IMG_MODE = Number(p.imageMode || 1);  // 1=fit, 2=stretch, 3=custom, 4=original size (custom dims allowed)
    var IMG_W    = Number(p.imageW || 28);
    var IMG_H    = Number(p.imageH || 28);

    if (mode === "replace") {
      sh.clearContents();
      if (headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // Write rows
    var startRow = sh.getLastRow() + 1;
    if (rows.length) {
      sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }

    // Sheet polish & category merges (Category column = E = 5)
    var hasHeader = (sh.getLastRow() > 0 && headers.length > 0);
    var firstDataRow = hasHeader ? 2 : 1;
    var lastRow = sh.getLastRow();
    if (lastRow >= firstDataRow) {
      var rngAll = sh.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, sh.getLastColumn());
      rngAll.setHorizontalAlignment("center");
      if (vAlign === "middle") rngAll.setVerticalAlignment("middle");
      mergeContiguousColumn_(sh, 5, firstDataRow, lastRow, sparse); // Category
    }

    // ---------- Preview images ----------
    // Determine the Preview column robustly.
    var previewCol = detectPreviewColumn_(sh, headers);

    if (rows.length && previewCol) {
      var folder = getOrCreateFolder_(folderId);

      for (var i = 0; i < rows.length; i++) {
        var b64 = images[i] || "";
        if (!b64) continue;

        var fileName = (runId + "_" + (startRow + i)).replace(/[^\w\-\.]/g, "_") + ".png";
        var blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);

        if (embed) {
          // Directly embed PNG into the sheet (no URL & sharing needed)
          var img = sh.insertImage(blob, previewCol, startRow + i);
          img.setAltTextDescription(fileName);
          // Optionally resize to IMG_W x IMG_H (px)
          try { img.setWidth(IMG_W).setHeight(IMG_H); } catch (e2) {}
          continue;
        }

        // Otherwise upload to Drive and insert =IMAGE(url)
        var file = folder.createFile(blob);
        try {
          // Broadest sharing first, then domain fallback
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (err1) {
          try {
            file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (err2) {
            // If both fail, the URL may not be publicly fetchable by IMAGE()
            // Consider using embedImages:true in payload for a guaranteed inline image.
          }
        }
        var fileId = file.getId();
        var publicUrl = "https://drive.google.com/uc?export=view&id=" + fileId;

        var cell = sh.getRange(startRow + i, previewCol);
        // =IMAGE(url[, mode, height, width]) — we prefer mode with dimensions
        if (IMG_MODE === 3 || IMG_MODE === 4) {
          cell.setFormula('=IMAGE("' + publicUrl + '",' + IMG_MODE + ',' + IMG_H + ',' + IMG_W + ')');
        } else {
          cell.setFormula('=IMAGE("' + publicUrl + '",' + IMG_MODE + ')');
        }
        cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Detect the "Preview" column:
 * 1) If headers are provided in payload, use them.
 * 2) Else read the first row of the live sheet.
 * 3) Fallback to column 13 (your current header order).
 */
function detectPreviewColumn_(sh, headersFromPayload) {
  var DEFAULT_PREVIEW_COL = 13;

  if (headersFromPayload && headersFromPayload.length) {
    var idx = headersFromPayload.indexOf("Preview");
    if (idx >= 0) return idx + 1;
  }
  if (sh.getLastRow() >= 1) {
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h).trim(); });
    var i2 = hdr.indexOf("Preview");
    if (i2 >= 0) return i2 + 1;
  }
  return DEFAULT_PREVIEW_COL;
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
  var values = sh.getRange(r1, colIndex, rN - r1 + 1, 1)
    .getValues()
    .map(function (r) { return sanitize_(r[0]); });

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
      sh.getRange(a, colIndex, len, 1).clearContent();
      if (anchor === "first") {
        sh.getRange(a, colIndex).setValue(v);
      } else if (anchor === "middle") {
        var mid = Math.floor((a + b) / 2);
        sh.getRange(mid, colIndex).setValue(v);
      } else { // 'last'
        sh.getRange(b, colIndex).setValue(v);
      }
      // merge & center
      sh.getRange(a, colIndex, len, 1).merge();
      sh.getRange(a, colIndex).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }
    runStart = runEnd;
  }
}

function sanitize_(s) {
  if (s == null) return "";
  s = String(s).replace(/\u00A0/g, " "); // NBSP → space
  s = s.replace(/\s+/g, " ").trim();
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
