/** Web App: BOQ uploads + previews (FULL CODE, with header-skip & hiding A..C)
 * - Detail/Count sheet: stores image in Drive + writes =IMAGE(public URL) in "Preview" column
 * - ByLayer sheet: sets Preview cell background color only (no image)
 * - Always treats row 1 as header (skips formatting/merging on row 1)
 * - Automatically hides columns A..C (run_id, source_file, handle)
 * Deploy > Manage deployments > Web app (Execute as: Me, Access: Anyone)
 */
function doPost(e) {
  try {
    const p = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");

    const ss    = SpreadsheetApp.openById(String(p.sheetId));
    const tab   = String(p.tab || "Detail");
    const sh    = ss.getSheetByName(tab) || ss.insertSheet(tab);

    const mode      = String(p.mode || "replace").toLowerCase();      // 'replace' | 'append'
    const headers   = Array.isArray(p.headers)  ? p.headers  : [];
    const rows      = Array.isArray(p.rows)     ? p.rows     : [];
    const images    = Array.isArray(p.images)   ? p.images   : [];    // base64 PNGs aligned to rows
    const colors    = Array.isArray(p.bgColors) ? p.bgColors : [];    // hex colors per row (ByLayer)
    const vAlign    = String(p.vAlign || "");
    const sparse    = String(p.sparseAnchor || "last");               // 'first'|'last'|'middle'
    const runId     = String(p.runId || "run");
    const folderId  = String(p.driveFolderId || "");
    const colorOnly = !!p.colorOnly;                                  // true for ByLayer; false for Detail

    // We still size cells, but IMAGE() will be 1-arg and scale to fit
    const IMG_W = Number(p.imageW || 42);
    const IMG_H = Number(p.imageH || 42);
    const PAD_W = 8, PAD_H = 8;

    // 1) Replace mode: clear + write header
    if (mode === "replace") {
      sh.clearContents();
      if (headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }

    // 2) Append data rows (never includes headers)
    const startRow = sh.getLastRow() + 1;
    if (rows.length) {
      sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }

    // 3) Basic formatting & merging (skip header row)
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    // Detect if row 1 is a header even when we appended without headers in this request
    let looksLikeHeader = false;
    if (lastRow >= 1 && lastCol >= 1) {
      const row1 = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      const a1 = sanitize_(row1[0] || "");
      const headerTokens = row1.map(s => sanitize_(s));
      looksLikeHeader = a1 === "RUN_ID" ||
                        headerTokens.includes("SOURCE_FILE") ||
                        headerTokens.includes("ENTITY_TYPE");
    }

    const firstDataRow = looksLikeHeader ? 2 : 1;
    if (lastRow >= firstDataRow) {
      const rngAll = sh.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol);
      rngAll.setHorizontalAlignment("center");
      if (vAlign === "middle") rngAll.setVerticalAlignment("middle");
      // Merge the Category column only for data rows (E = 5)
      if (lastCol >= 5) mergeContiguousColumn_(sh, 5, firstDataRow, lastRow, sparse);
    }

    // 4) Ensure "Preview" column exists; get 1-based index
    const previewCol = ensurePreviewColumn_(sh);

    // 5) Size preview column & rows for just-written rows
    if (rows.length) {
      sh.setColumnWidth(previewCol, IMG_W + PAD_W);
      sh.setRowHeights(startRow, rows.length, IMG_H + PAD_H);
    }

    // 6) Apply color or images (only newly added rows)
    if (rows.length && previewCol) {
      const folder = getOrCreateFolder_(folderId);

      for (let i = 0; i < rows.length; i++) {
        const r = startRow + i;
        const cell = sh.getRange(r, previewCol);

        // If a color is provided, always apply as background (used by ByLayer)
        const hex = (colors[i] || "").toString().trim();
        if (hex) cell.setBackground(hex);

        if (colorOnly) continue;  // ByLayer: stop here (no images)

        // Detail/Count row image
        const b64 = images[i] || "";
        if (!b64) continue;

        const fileName = (runId + "_" + r).replace(/[^\w\-\.]/g, "_") + ".png";
        const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
        const file = folder.createFile(blob);

        // Make link-viewable so IMAGE() works
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (err) {
          // Fallback if domain blocks "anyone with link"
          try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {}
        }

        const url = "https://drive.google.com/uc?export=view&id=" + file.getId();

        // Locale-safe single-argument IMAGE()
        cell.setFormula('=IMAGE("' + url + '")');
        cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
      }
    }

    // 7) Hide bookkeeping columns A..C (run_id, source_file, handle)
    try {
      const FIRST_N = 3; // A..C
      if (sh.getMaxColumns() >= FIRST_N) {
        sh.hideColumns(1, FIRST_N);
      }
    } catch (e) {}

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Guarantee a "Preview" column exists; return its 1-based index. */
function ensurePreviewColumn_(sh) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() >= 1 && lastCol > 0) {
    const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    let idx = hdr.findIndex(h => (h || "").toString().trim().toLowerCase() === "preview");
    if (idx >= 0) return idx + 1;

    sh.insertColumnAfter(lastCol);
    sh.getRange(1, lastCol + 1).setValue("Preview");
    return lastCol + 1;
  }
  // No header yet
  sh.getRange(1, 1).setValue("Preview");
  return 1;
}

/** Merge contiguous identical cells in a given column (data rows only). */
function mergeContiguousColumn_(sh, colIndex, r1, rN, anchor) {
  if (rN < r1) return;
  const values = sh.getRange(r1, colIndex, rN - r1 + 1, 1)
    .getValues()
    .map(r => sanitize_(r[0]));

  let runStart = 0;
  while (runStart < values.length) {
    const v = values[runStart];
    if (!v) { runStart++; continue; }
    let runEnd = runStart + 1;
    while (runEnd < values.length && values[runEnd] === v) runEnd++;

    const len = runEnd - runStart;
    if (len >= 2) {
      const a = r1 + runStart, b = r1 + runEnd - 1;
      sh.getRange(a, colIndex, len, 1).clearContent();
      if (anchor === "first") sh.getRange(a, colIndex).setValue(v);
      else if (anchor === "middle") sh.getRange(Math.floor((a + b) / 2), colIndex).setValue(v);
      else sh.getRange(b, colIndex).setValue(v); // default: 'last'
      sh.getRange(a, colIndex, len, 1).merge()
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    }
    runStart = runEnd;
  }
}

function sanitize_(s) {
  if (s == null) return "";
  s = String(s).replace(/\u00A0/g, " ");
  return s.replace(/\s+/g, " ").trim().toUpperCase();
}

function getOrCreateFolder_(folderId) {
  if (folderId) { try { return DriveApp.getFolderById(folderId); } catch (e) {} }
  const it = DriveApp.getFoldersByName("DXF-Previews");
  return it.hasNext() ? it.next() : DriveApp.createFolder("DXF-Previews");
}
