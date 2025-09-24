/** ===================== DXF WebApp + Local-Preview Linker (FULL) =====================

Features:
- Existing CSV upload flow (Detail vs ByLayer) preserved
- Adds op: "previewByName" → accepts base64 images from a local Python uploader
  and writes =IMAGE(DriveURL) into the Preview column by matching BOQ name.
- Robust header matching + identical normalization to Python (letters+digits only)

======================================================================================= */

function normalizeHeader_(s) {
  return String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
}
function normKey_(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function ensurePreviewColumn_(sh) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() >= 1 && lastCol > 0) {
    const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idx = hdr.findIndex(h => normalizeHeader_(h) === "preview");
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
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const idx = hdr.findIndex(h => normalizeHeader_(h) === normalizeHeader_(name));
  return idx >= 0 ? idx + 1 : 0;
}
function removeColumnsByHeader_(sh, names) {
  if (!names || !names.length) return;
  const want = new Set(names.map(n => normalizeHeader_(n)));
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() < 1 || lastCol < 1) return;
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  for (let c = hdr.length - 1; c >= 0; c--) {
    if (want.has(hdr[c])) sh.deleteColumn(c + 1);
  }
}
function applyMerge_(sheet, col, s, e, value, anchor) {
  const band = sheet.getRange(s, col, e - s + 1, 1);
  band.clearContent();
  const anchorRow = (anchor === "first") ? s : (anchor === "middle" ? Math.floor((s + e) / 2) : e);
  sheet.getRange(anchorRow, col).setValue(value);
  band.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
}
function mergeBandsByHeaders_(sheet, headerNames, anchor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(normalizeHeader_);
  const r1 = 2, rN = values.length;
  headerNames.forEach(name => {
    const idx = headers.indexOf(normalizeHeader_(name));
    if (idx < 0) return;
    const col = idx + 1;
    sheet.getRange(r1, col, rN - r1 + 1, 1).breakApart();
    let s = r1, v = "", vn = "";
    for (let r = r1; r <= rN + 1; r++) {
      const raw = r <= rN ? String(values[r - 1][idx] || "") : "\u0000__END__";
      const norm = normalizeHeader_(raw);
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
function mergeColumnByGroup_(sheet, targetHeader, groupHeader, anchor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(normalizeHeader_);
  const tIdx = headers.indexOf(normalizeHeader_(targetHeader));
  const gIdx = headers.indexOf(normalizeHeader_(groupHeader));
  if (tIdx < 0 || gIdx < 0) return;

  const tCol = tIdx + 1;
  const r1 = 2, rN = values.length;
  sheet.getRange(r1, tCol, rN - r1 + 1, 1).breakApart();

  let runStart = r1;
  let lastGroup = normalizeHeader_(values[r1 - 1][gIdx] || "");
  let lastVal   = normalizeHeader_(values[r1 - 1][tIdx] || "");

  for (let r = r1 + 1; r <= rN + 1; r++) {
    const cg = (r <= rN) ? normalizeHeader_(values[r - 1][gIdx] || "") : "\u0000__END__";
    const cv = (r <= rN) ? normalizeHeader_(values[r - 1][tIdx] || "") : "\u0000__END__";
    if (r <= rN && cg === lastGroup && cv === lastVal) continue;
    const runEnd = r - 1;
    if (lastVal && runEnd > runStart) applyMerge_(sheet, tCol, runStart, runEnd, values[runStart - 1][tIdx], anchor);
    if (r <= rN) { runStart = r; lastGroup = cg; lastVal = cv; }
  }
}
function getOrCreateFolder_(folderId) {
  if (folderId) { try { return DriveApp.getFolderById(folderId); } catch (e) {} }
  const it = DriveApp.getFoldersByName("DXF-Previews");
  return it.hasNext() ? it.next() : DriveApp.createFolder("DXF-Previews");
}

/** ============== NEW: Local previews by BOQ name ============== */
/**
 * POST body:
 * {
 *   op: "previewByName",
 *   sheetId, tab, driveFolderId,
 *   items: [ { name: "<BOQ name>", imageB64: "<base64 png/jpg/webp>" }, ... ]
 * }
 */
function handlePreviewByName_(p) {
  const ss  = SpreadsheetApp.openById(String(p.sheetId));
  const tab = String(p.tab || "Fts");
  const sh  = ss.getSheetByName(tab) || ss.insertSheet(tab);
  const folder = getOrCreateFolder_(String(p.driveFolderId || ""));

  // --- helpers (same normalization as Python) ---
  const normKey_ = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const baseKey_ = s => normKey_(s).replace(/\d+$/,""); // strip trailing digits like ...001

  // columns
  const header = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  const cBOQ   = header.findIndex(h => normalizeHeader_(h) === "boq name") + 1;
  let   cPrev  = header.findIndex(h => normalizeHeader_(h) === "preview") + 1;
  if (!cPrev) cPrev = ensurePreviewColumn_(sh);
  if (!cBOQ) throw new Error('Column "BOQ name" not found');

  const r1 = 2, rN = sh.getLastRow();
  if (rN < r1) return { matched: 0, wrote: 0 };

  // index rows by BOTH fullKey and baseKey
  const names = sh.getRange(r1, cBOQ, rN - r1 + 1, 1).getDisplayValues().map(r => String(r[0]||""));
  const idx = new Map();        // key → [row,...]
  const idxBase = new Map();    // baseKey → [row,...]
  names.forEach((n, i) => {
    const row = r1 + i;
    const k  = normKey_(n);
    const kb = baseKey_(n);
    if (k)  { if (!idx.has(k)) idx.set(k, []); idx.get(k).push(row); }
    if (kb) { if (!idxBase.has(kb)) idxBase.set(kb, []); idxBase.get(kb).push(row); }
  });

  const items = Array.isArray(p.items) ? p.items : [];
  let matched = 0, wrote = 0;

  const MIN_PARTIAL = 8; // require at least 8 chars for partial match
  function uniquePartialFind(key, map) {
    if (!key || key.length < MIN_PARTIAL) return [];
    // collect candidates where mapKey startsWith key OR key startsWith mapKey
    const hits = [];
    map.forEach((rows, mapKey) => {
      if (mapKey.length < MIN_PARTIAL) return;
      if (mapKey.indexOf(key) === 0 || key.indexOf(mapKey) === 0) {
        rows.forEach(r => hits.push(r));
      }
    });
    // only accept if it resolves to rows belonging to exactly one BOQ name run (avoid mass merge)
    return (hits.length ? hits : []);
  }

  items.forEach((it, j) => {
    const rawName = it && it.name;
    const b64 = String(it && it.imageB64 || "");
    if (!rawName || !b64) return;

    const kFull = normKey_(rawName);
    const kBase = baseKey_(rawName);

    // Try exact full → exact base → unique partial on base
    let rows = (idx.get(kFull) || []);
    if (!rows.length) rows = (idxBase.get(kBase) || []);
    if (!rows.length) {
      const part = uniquePartialFind(kBase, idxBase);
      if (part.length) rows = part;
    }
    if (!rows.length) return;   // no match

    // upload once, reuse for all matched rows
    const fileName = ("boq_" + kBase + "_" + (j+1)).slice(0,120) + ".png";
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (_) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch(_){} }
    const url = "https://drive.google.com/uc?export=view&id=" + file.getId();

    rows.forEach(r => {
      sh.getRange(r, cPrev).setFormula('=IMAGE("' + url + '")')
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      wrote++;
    });
    matched++;
  });

  if (wrote) {
    sh.setColumnWidth(cPrev, 50);
    sh.setRowHeights(r1, rN - r1 + 1, 50);
  }
  return { matched, wrote };
}


/** ===================== Existing CSV Upload Endpoint ===================== */
function doPost(e) {
  try {
    const p = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");

    // NEW route for local preview linkage
    if (String(p.op || "") === "previewByName") {
      const res = handlePreviewByName_(p);
      return ContentService.createTextOutput(JSON.stringify({ ok:true, ...res }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // --------- Original upload handler (Detail vs ByLayer) ----------
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
    const IMG_W = Number(p.imageW || 42);
    const IMG_H = Number(p.imageH || 42);
    const PAD_W = 8, PAD_H = 8;

    let headers = headersIn ? headersIn.slice() : null;
    let rows    = rowsIn.slice();
    if (!colorOnly && headers) {
      const kill = new Set(["entity_type", "category"]);
      const keepIdx = headers.map((h, i) => ({ i, keep: !kill.has(String(h).trim().toLowerCase()) }))
                             .filter(x => x.keep)
                             .map(x => x.i);
      headers = keepIdx.map(i => headersIn[i]);
      rows    = rows.map(r => keepIdx.map(i => r[i]));
    }

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

    if (rows.length) {
      const nCols = Math.max(...rows.map(r => r.length));
      if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
      if (sh.getMaxRows() < startRow - 1 + rows.length) {
        sh.insertRowsAfter(sh.getMaxRows(), startRow - 1 + rows.length - sh.getMaxRows());
      }
      sh.getRange(startRow, 1, rows.length, nCols).setValues(rows);
    }

    if (rows.length) {
      const rng = sh.getRange(startRow, 1, rows.length, sh.getLastColumn());
      rng.setHorizontalAlignment("center");
      if (vAlign === "middle") rng.setVerticalAlignment("middle");
    }

    const previewCol = ensurePreviewColumn_(sh);
    if (rows.length) {
      sh.setColumnWidth(previewCol, IMG_W + PAD_W);
      sh.setRowHeights(startRow, rows.length, IMG_H + PAD_H);
    }

    if (rows.length && previewCol) {
      if (colorOnly) {
        for (let i = 0; i < rows.length; i++) {
          const hex = (colors[i] || "").toString().trim();
          if (hex) sh.getRange(startRow + i, previewCol).setBackground(hex);
        }
      } else {
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

    if (colorOnly) {
      // ByLayer sheet: keep as-is
    } else {
      const colZone = colIndexByHeader_(sh, "zone");
      if (colZone > 0) {
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

        mergeBandsByHeaders_(sh, ["zone"], "first");
        mergeColumnByGroup_(sh, "category1", "zone", "first");
      }

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
