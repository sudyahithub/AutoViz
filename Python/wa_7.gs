/** Web App: BOQ uploads + preview images (Drive URLs or direct embed)
 * Deploy: Deploy → Manage deployments → Web app (Execute as: Me, Access: Anyone)
 *
 * Payload JSON (POST):
 * {
 *   sheetId: string,
 *   tab: string,
 *   mode: "replace"|"append",
 *   headers: string[],
 *   rows: any[][],
 *   images: string[],          // base64 PNG per row; "" for none
 *   vAlign: ""|"middle",
 *   sparseAnchor: "first"|"last"|"middle",
 *   runId: string,
 *   driveFolderId: string,
 *   embedImages: boolean,      // true → insert PNG objects (no URL)
 *   imageMode: 1|2|3|4,        // used only when embedImages === false
 *   imageW: number,            // for mode 3/4 or embed sizing
 *   imageH: number
 * }
 */

function doPost(e) {
  try {
    var p = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");

    var ss  = SpreadsheetApp.openById(p.sheetId);
    var sh  = ss.getSheetByName(String(p.tab || "Detail")) || ss.insertSheet(String(p.tab || "Detail"));
    var mode   = String(p.mode || "replace").toLowerCase();
    var headers = Array.isArray(p.headers) ? p.headers : [];
    var rows    = Array.isArray(p.rows)    ? p.rows    : [];
    var images  = Array.isArray(p.images)  ? p.images  : [];
    var vAlign  = String(p.vAlign || "");
    var sparse  = String(p.sparseAnchor || "last");
    var runId   = String(p.runId || "run");
    var folderId = String(p.driveFolderId || "");
    var embed   = !!p.embedImages;

    var IMG_MODE = Number(p.imageMode || 1);
    var IMG_W    = Number(p.imageW || 28);
    var IMG_H    = Number(p.imageH || 28);

    if (mode === "replace") {
      sh.clearContents();
      if (headers.length) sh.getRange(1,1,1,headers.length).setValues([headers]);
    }

    var startRow = sh.getLastRow() + 1;
    if (rows.length) sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    // Formatting & Category merge (E)
    var firstDataRow = (headers.length ? 2 : 1);
    var lastRow = sh.getLastRow();
    if (lastRow >= firstDataRow) {
      var rngAll = sh.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, sh.getLastColumn());
      rngAll.setHorizontalAlignment("center");
      if (vAlign === "middle") rngAll.setVerticalAlignment("middle");
      mergeContiguousColumn_(sh, 5, firstDataRow, lastRow, sparse);
    }

    // Previews
    var previewCol = detectPreviewColumn_(sh, headers);
    if (rows.length && previewCol) {
      var folder = getOrCreateFolder_(folderId);
      for (var i = 0; i < rows.length; i++) {
        var b64 = images[i] || "";
        if (!b64) continue;

        var fileName = (runId + "_" + (startRow + i)).replace(/[^\w\-\.]/g, "_") + ".png";
        var blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);

        if (embed) {
          var img = sh.insertImage(blob, previewCol, startRow + i);
          img.setAltTextDescription(fileName);
          try { img.setWidth(IMG_W).setHeight(IMG_H); } catch (e2) {}
          continue;
        }

        var file = folder.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
        catch (err1) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (err2) {} }
        var publicUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();

        var cell = sh.getRange(startRow + i, previewCol);
        if (IMG_MODE === 3 || IMG_MODE === 4) {
          cell.setFormula('=IMAGE("' + publicUrl + '",' + IMG_MODE + ',' + IMG_H + ',' + IMG_W + ')');
        } else {
          cell.setFormula('=IMAGE("' + publicUrl + '",' + IMG_MODE + ')');
        }
        cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function detectPreviewColumn_(sh, headersFromPayload) {
  var DEFAULT_PREVIEW_COL = 13; // M (matches your CSV header order)
  if (headersFromPayload && headersFromPayload.length) {
    var idx = headersFromPayload.indexOf("Preview");
    if (idx >= 0) return idx + 1;
  }
  if (sh.getLastRow() >= 1) {
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
    var i2 = hdr.indexOf("Preview");
    if (i2 >= 0) return i2 + 1;
  }
  return DEFAULT_PREVIEW_COL;
}

function mergeContiguousColumn_(sh, colIndex, r1, rN, anchor) {
  if (rN < r1) return;
  var values = sh.getRange(r1, colIndex, rN - r1 + 1, 1).getValues().map(function(r){ return sanitize_(r[0]); });
  var runStart = 0;
  while (runStart < values.length) {
    var v = values[runStart]; if (!v) { runStart++; continue; }
    var runEnd = runStart + 1; while (runEnd < values.length && values[runEnd] === v) runEnd++;
    var len = runEnd - runStart;
    if (len >= 2) {
      var a = r1 + runStart, b = r1 + runEnd - 1;
      sh.getRange(a, colIndex, len, 1).clearContent();
      if (anchor === "first") sh.getRange(a, colIndex).setValue(v);
      else if (anchor === "middle") sh.getRange(Math.floor((a+b)/2), colIndex).setValue(v);
      else sh.getRange(b, colIndex).setValue(v);
      sh.getRange(a, colIndex, len, 1).merge();
      sh.getRange(a, colIndex).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }
    runStart = runEnd;
  }
}
function sanitize_(s){ if (s==null) return ""; s=String(s).replace(/\u00A0/g," "); s=s.replace(/\s+/g," ").trim(); return s.toUpperCase(); }
function getOrCreateFolder_(folderId){
  if (folderId) { try { return DriveApp.getFolderById(folderId); } catch (e) {} }
  var it = DriveApp.getFoldersByName("DXF-Previews");
  return it.hasNext() ? it.next() : DriveApp.createFolder("DXF-Previews");
}
