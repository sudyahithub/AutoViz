



/** Web App: Detail vs ByLayer handling (keeps entity_type/category only for ByLayer) */
function doPost(e) {
  try {
    const p = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");

    const ss  = SpreadsheetApp.openById(String(p.sheetId));
    const tab = String(p.tab || "Detail");
    const sh  = ss.getSheetByName(tab) || ss.insertSheet(tab);

    const mode       = String(p.mode || "replace").toLowerCase();     // 'replace' | 'append'
    const headersIn  = Array.isArray(p.headers)  ? p.headers  : null;
    const rowsIn     = Array.isArray(p.rows)     ? p.rows     : [];
    const images     = Array.isArray(p.images)   ? p.images   : [];
    const colors     = Array.isArray(p.bgColors) ? p.bgColors : [];
    const colorOnly  = !!p.colorOnly;                                  // true → ByLayer
    const vAlign     = String(p.vAlign || "");
    const runId      = String(p.runId || "run");
    const folderId   = String(p.driveFolderId || "");

    // Preview sizing
    const IMG_W = Number(p.imageW || 42);
    const IMG_H = Number(p.imageH || 42);
    const PAD_W = 8, PAD_H = 8;

    // ---------- 0) Prepare headers/rows (branch on colorOnly) ----------
    let headers = headersIn ? headersIn.slice() : null;
    let rows    = rowsIn.slice();

    if (!colorOnly && headers) {
      // DETAIL: drop entity_type, category
      const kill = new Set(["entity_type", "category"]);
      const keepIdx = headers.map((h, i) => ({ i, keep: !kill.has(String(h).trim().toLowerCase()) }))
                             .filter(x => x.keep)
                             .map(x => x.i);
      headers = keepIdx.map(i => headersIn[i]);
      rows    = rows.map(r => keepIdx.map(i => r[i]));
    }
    // BYLAYER: keep columns as-is

    // ---------- 1) Replace mode: clear & write header ----------
    let startRow;
    if (mode === "replace") {
      sh.clearContents();
      if (headers && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        startRow = 2;
      } else {
        startRow = 1;
      }
    } else {
      const last = sh.getLastRow();
      startRow = last ? last + 1 : (headers ? 2 : 1);
    }

    // ---------- 2) Write rows ----------
    if (rows.length) {
      const nCols = Math.max(...rows.map(r => r.length));
      if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
      if (sh.getMaxRows() < startRow - 1 + rows.length) {
        sh.insertRowsAfter(sh.getMaxRows(), startRow - 1 + rows.length - sh.getMaxRows());
      }
      sh.getRange(startRow, 1, rows.length, nCols).setValues(rows);
    }

    // ---------- 3) Basic formatting ----------
    if (rows.length) {
      const rng = sh.getRange(startRow, 1, rows.length, sh.getLastColumn());
      rng.setHorizontalAlignment("center");
      if (vAlign === "middle") rng.setVerticalAlignment("middle");
    }

    // ---------- 4) Ensure Preview column ----------
    const previewCol = ensurePreviewColumn_(sh);

    // ---------- 5) Size preview column / new rows ----------
    if (rows.length) {
      sh.setColumnWidth(previewCol, IMG_W + PAD_W);
      sh.setRowHeights(startRow, rows.length, IMG_H + PAD_H);
    }

    // ---------- 6) Previews / Color swatches ----------
    if (rows.length && previewCol) {
      if (colorOnly) {
        // BYLAYER: set background color only
        for (let i = 0; i < rows.length; i++) {
          const hex = (colors[i] || "").toString().trim();
          if (hex) sh.getRange(startRow + i, previewCol).setBackground(hex);
        }
      } else {
        // DETAIL: upload PNGs & write =IMAGE(url)
        if (images.length) {
          const folder = getOrCreateFolder_(folderId);
          for (let i = 0; i < rows.length; i++) {
            const b64 = images[i] || "";
            if (!b64) continue;
            const r = startRow + i;
            const fileName = (runId + "_" + r).replace(/[^\w\-\.]/g, "_") + ".png";
            const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
            const file = folder.createFile(blob);
            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
            catch (_) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {} }
            const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
            sh.getRange(r, previewCol).setFormula('=IMAGE("' + url + '")')
              .setHorizontalAlignment("center").setVerticalAlignment("middle");
          }
        }
      }
    }

    // ---------- 7) Post-write shaping ----------
    if (colorOnly) {
      // BYLAYER: DO NOTHING with zone. Keep entity_type/category. No sorts/merges needed.
      // This keeps the layout like your image-1 (Preview far right, zone empty).
    } else {
      // DETAIL: normalize + sort zone, merges, etc.
      const colZone = colIndexByHeader_(sh, "zone");
      if (colZone > 0) {
        // 7a) blanks → "misc"
        const r1 = 2, rN = sh.getLastRow();
        if (rN >= r1) {
          const zoneRng = sh.getRange(r1, colZone, rN - r1 + 1, 1);
          const Z = zoneRng.getValues();
          let changed = false;
          for (let i = 0; i < Z.length; i++) {
            const s = String(Z[i][0] || "").trim();
            if (!s) { Z[i][0] = "misc"; changed = true; }
          }
          if (changed) zoneRng.setValues(Z);
        }

        // 7b) sort by zone, forcing "misc" to last using temp key
        const lastCol = sh.getLastColumn();
        const r1s = 2, rNs = sh.getLastRow();
        if (rNs >= r1s) {
          sh.insertColumnAfter(lastCol);
          const skCol = lastCol + 1;
          sh.getRange(1, skCol).setValue("__sort_zone__");
          const vals = sh.getRange(r1s, colZone, rNs - r1s + 1, 1).getValues();
          const keys = vals.map(v => {
            const s = String(v[0] || "").toLowerCase().trim();
            return (s === "misc") ? "zzzzzz" : s;
          }).map(k => [k]);
          sh.getRange(r1s, skCol, keys.length, 1).setValues(keys);
          sh.getRange(r1s, 1, rNs - r1s + 1, skCol).sort([{ column: skCol, ascending: true }]);
          sh.deleteColumn(skCol);
        }

        // 7c) merges
        mergeBandsByHeaders_(sh, ["zone"], "first");
        mergeColumnByGroup_(sh, "category1", "zone", "first");
      }

      // DETAIL: also remove any lingering entity_type/category columns if present
      removeColumnsByHeader_(sh, ["entity_type", "category"]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, wrote: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ===== Helpers ===== */

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

function colIndexByHeader_(sh, name) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() < 1 || lastCol < 1) return 0;
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = hdr.findIndex(h => String(h).trim().toLowerCase() === String(name).toLowerCase());
  return idx >= 0 ? idx + 1 : 0;
}

function removeColumnsByHeader_(sh, names) {
  if (!names || !names.length) return;
  const set = new Set(names.map(n => String(n).toLowerCase()));
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() < 1 || lastCol < 1) return;
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  for (let c = hdr.length - 1; c >= 0; c--) {
    if (set.has(hdr[c])) sh.deleteColumn(c + 1);
  }
}

/** Merge vertical bands for the given header names (blank or equal → same band). */
function mergeBandsByHeaders_(sheet, headerNames, anchor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const r1 = 2, rN = values.length;

  headerNames.forEach(name => {
    const idx = headers.indexOf(String(name).toLowerCase());
    if (idx < 0) return;
    const col = idx + 1;
    sheet.getRange(r1, col, rN - r1 + 1, 1).breakApart();
    let s = r1, v = "", vn = "";
    for (let r = r1; r <= rN + 1; r++) {
      const raw = r <= rN ? String(values[r - 1][idx] || "") : "\u0000__END__";
      const norm = raw.trim().toUpperCase();
      if (!v) { if (norm) { s = r; v = raw; vn = norm; } continue; }
      const cont = r <= rN && (!norm || norm === vn);
      if (cont) continue;
      const e = r - 1;
      if (e > s) applyMerge_(sheet, col, s, e, v, anchor);
      v = ""; vn = "";
      if (r <= rN && norm) { s = r; v = raw; vn = norm; }
    }
  });
}

/** Merge contiguous duplicates in target column, within the same group column. */
function mergeColumnByGroup_(sheet, targetHeader, groupHeader, anchor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const tIdx = headers.indexOf(String(targetHeader).toLowerCase());
  const gIdx = headers.indexOf(String(groupHeader).toLowerCase());
  if (tIdx < 0 || gIdx < 0) return;

  const tCol = tIdx + 1;
  const r1 = 2, rN = values.length;
  sheet.getRange(r1, tCol, rN - r1 + 1, 1).breakApart();

  let runStart = r1;
  let lastGroup = String(values[r1 - 1][gIdx] || "").trim().toUpperCase();
  let lastVal   = String(values[r1 - 1][tIdx] || "").trim().toUpperCase();

  for (let r = r1 + 1; r <= rN + 1; r++) {
    const cg = (r <= rN) ? String(values[r - 1][gIdx] || "").trim().toUpperCase() : "\u0000__END__";
    const cv = (r <= rN) ? String(values[r - 1][tIdx] || "").trim().toUpperCase() : "\u0000__END__";
    if (r <= rN && cg === lastGroup && cv === lastVal) continue;
    const runEnd = r - 1;
    if (lastVal && runEnd > runStart) applyMerge_(sheet, tCol, runStart, runEnd, values[runStart - 1][tIdx], anchor);
    if (r <= rN) { runStart = r; lastGroup = cg; lastVal = cv; }
  }
}

function applyMerge_(sheet, col, s, e, value, anchor) {
  const band = sheet.getRange(s, col, e - s + 1, 1);
  band.clearContent();
  const anchorRow = (anchor === "first") ? s : (anchor === "middle" ? Math.floor((s + e) / 2) : e);
  sheet.getRange(anchorRow, col).setValue(value);
  band.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
}

function getOrCreateFolder_(folderId) {
  if (folderId) { try { return DriveApp.getFolderById(folderId); } catch (e) {} }
  const it = DriveApp.getFoldersByName("DXF-Previews");
  return it.hasNext() ? it.next() : DriveApp.createFolder("DXF-Previews");
}
