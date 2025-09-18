/** Web App: BOQ uploads + previews (no hidden columns, header-aware, with merging) */
function doPost(e) {
  try {
    const p = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");

    const ss  = SpreadsheetApp.openById(String(p.sheetId));
    const tab = String(p.tab || "Detail");
    const sh  = ss.getSheetByName(tab) || ss.insertSheet(tab);

    const mode       = String(p.mode || "replace").toLowerCase();     // 'replace' | 'append'
    const headers    = Array.isArray(p.headers)  ? p.headers  : [];
    const rows       = Array.isArray(p.rows)     ? p.rows     : [];
    const images     = Array.isArray(p.images)   ? p.images   : [];   // base64 PNGs per row
    const colors     = Array.isArray(p.bgColors) ? p.bgColors : [];   // hex colors per row (ByLayer)
    const vAlign     = String(p.vAlign || "");                         // "", "middle"
    const runId      = String(p.runId || "run");
    const folderId   = String(p.driveFolderId || "");
    const colorOnly  = !!p.colorOnly;                                  // true for ByLayer; false for Detail

    // Preview sizing
    const IMG_W = Number(p.imageW || 42);
    const IMG_H = Number(p.imageH || 42);
    const PAD_W = 8, PAD_H = 8;

    // 1) Replace: clear + headers
    let startRow;
    if (mode === "replace") {
      sh.clearContents();
      if (headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        startRow = 2;
      } else {
        startRow = 1;
      }
    } else {
      const last = sh.getLastRow();
      startRow = last ? last + 1 : (headers.length ? 2 : 1);
    }

    // 2) Write rows
    if (rows.length) {
      const nCols = Math.max(...rows.map(r => r.length));
      if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
      if (sh.getMaxRows() < startRow - 1 + rows.length) {
        sh.insertRowsAfter(sh.getMaxRows(), startRow - 1 + rows.length - sh.getMaxRows());
      }
      sh.getRange(startRow, 1, rows.length, nCols).setValues(rows);
    }

    // 3) Format newly written rows
    const lastCol = sh.getLastColumn();
    if (rows.length) {
      const rng = sh.getRange(startRow, 1, rows.length, lastCol);
      rng.setHorizontalAlignment("center");
      if (vAlign === "middle") rng.setVerticalAlignment("middle");
    }

    // 4) Ensure "Preview" column
    const previewCol = ensurePreviewColumn_(sh);

    // 5) Size preview column & rows
    if (rows.length) {
      sh.setColumnWidth(previewCol, IMG_W + PAD_W);
      sh.setRowHeights(startRow, rows.length, IMG_H + PAD_H);
    }

    // 6) Color or images
    if (rows.length && previewCol) {
      if (colorOnly && colors.length) {
        for (let i = 0; i < rows.length; i++) {
          const hex = (colors[i] || "").toString().trim();
          if (hex) sh.getRange(startRow + i, previewCol).setBackground(hex);
        }
      } else if (!colorOnly && images.length) {
        const folder = getOrCreateFolder_(folderId);
        for (let i = 0; i < rows.length; i++) {
          const b64 = images[i] || "";
          if (!b64) continue;

          const r = startRow + i;
          const cell = sh.getRange(r, previewCol);

          const fileName = (runId + "_" + r).replace(/[^\w\-\.]/g, "_") + ".png";
          const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
          const file = folder.createFile(blob);
          try {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (_) {
            try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {}
          }
          const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
          cell.setFormula('=IMAGE("' + url + '")')
              .setHorizontalAlignment("center")
              .setVerticalAlignment("middle");
        }
      }
    }

    // 7) Merge bands for CATEGORY and ZONE (treat blanks as continuation; also merge equal repeats)
    mergeBandsByHeaders_(sh, ["category", "zone"], "first");

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, startRow, rows: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Ensure "Preview" column exists; return its index (1-based). */
function ensurePreviewColumn_(sh) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() >= 1 && lastCol > 0) {
    const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idx = hdr.findIndex(h => (h || "").toString().trim().toLowerCase() === "preview");
    if (idx >= 0) return idx + 1;
    sh.insertColumnAfter(lastCol);
    sh.getRange(1, lastCol + 1).setValue("Preview");
    return lastCol + 1;
  }
  sh.getRange(1, 1).setValue("Preview");
  return 1;
}

/**
 * Merge contiguous bands in the given header-named columns.
 * Rules:
 *   - Start a band on any NON-EMPTY value.
 *   - Extend the band while next rows are EITHER BLANK OR EQUAL (case-insensitive).
 *   - Close the band when a DIFFERENT non-empty value appears.
 *   - `anchor` = "first" | "middle" | "last" controls where the visible value is placed.
 */
function mergeBandsByHeaders_(sheet, headerNames, anchor) {
  if (!sheet || !headerNames || !headerNames.length) return;

  const rng = sheet.getDataRange();
  const values = rng.getValues();                 // includes header row
  if (values.length < 2) return;                  // nothing to merge

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const r1 = 2;                                   // data starts at row 2
  const rN = values.length;

  headerNames.forEach(name => {
    const idx = headers.indexOf(String(name).trim().toLowerCase());
    if (idx < 0) return;                          // header not found
    const col = idx + 1;

    // Break existing merges in this column (data rows only)
    sheet.getRange(r1, col, rN - r1 + 1, 1).breakApart();

    let s = r1;           // start of current band
    let v = "";           // band value (original, for display)
    let vn = "";          // normalized (uppercased) for compare

    let r = r1;
    while (r <= rN) {
      const raw = String(values[r - 1][idx] || "");
      const norm = raw.trim().toUpperCase();

      if (!v) {
        // No active band; start when we see non-empty
        if (norm) { s = r; v = raw; vn = norm; }
        r++;
        continue;
      }

      // We have an active band with vn; continue while blank OR equal to vn
      const isBlank = !norm;
      const isEqual = (norm === vn);
      if (isBlank || isEqual) {
        r++;
        continue;
      }

      // Different non-empty => close previous band [s .. r-1]
      const e = r - 1;
      if (e > s) applyMerge_(sheet, col, s, e, v, anchor);
      // start new band
      s = r; v = raw; vn = norm;
      r++;
    }

    // Flush last band
    if (v && rN >= s) {
      const e = rN;
      if (e > s) applyMerge_(sheet, col, s, e, v, anchor);
    }
  });
}

function applyMerge_(sheet, col, s, e, value, anchor) {
  const len = e - s + 1;
  const band = sheet.getRange(s, col, len, 1);
  band.clearContent();
  const anchorRow = (anchor === "first") ? s :
                    (anchor === "middle") ? Math.floor((s + e) / 2) : e;
  sheet.getRange(anchorRow, col).setValue(value);
  band.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
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
