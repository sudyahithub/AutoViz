/**
 * POST payload example:
 * {
 *   "sheetId": "<<<SPREADSHEET_ID>>>",
 *   "tab": "Detail",
 *   "mode": "replace",            // or "append"
 *   "headers": ["qty_type","qty_value","BOQ name","category","handle","remarks","length","width"],
 *   "rows": [...],
 *   "categoryHeaderName": "category", // header text (case-insensitive)
 *   "fallbackCategoryCol": 7,         // if header not found; 1-based (G=7)
 *   "sortBeforeMerge": true,          // <-- NEW: sort data rows by category before merging
 *   "secondarySortCols": [3]          // optional: also sort by BOQ name (example), 1-based indices
 * }
 */

const DEFAULT_CATEGORY_HEADER = "category";
const DEFAULT_FALLBACK_COL_INDEX = 7; // G

function doPost(e) {
  try {
    const body = e.postData && e.postData.contents ? e.postData.contents : "{}";
    const p = JSON.parse(body);

    const ss = SpreadsheetApp.openById(p.sheetId);
    const tabName = String(p.tab || "Detail");
    const mode = String(p.mode || "replace").toLowerCase();
    const headers = p.headers || [];
    const rows = p.rows || [];

    const headerName = (p.categoryHeaderName || DEFAULT_CATEGORY_HEADER);
    const fallbackIdx = Number(p.fallbackCategoryCol || DEFAULT_FALLBACK_COL_INDEX);
    const sortBeforeMerge = (p.sortBeforeMerge !== false); // default true
    const secondarySortCols = Array.isArray(p.secondarySortCols) ? p.secondarySortCols : [];

    let sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);

    if (mode === "replace") {
      sh.clearContents();
      if (headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (rows.length) sh.getRange(2, 1, rows.length, headers.length || rows[0].length).setValues(rows);
    } else if (mode === "append") {
      if (sh.getLastRow() === 0 && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
      if (rows.length) {
        const startRow = sh.getLastRow() + 1;
        const width = headers.length || rows[0].length;
        sh.getRange(startRow, 1, rows.length, width).setValues(rows);
      }
    } else {
      throw new Error("Unknown mode: " + mode);
    }

    // Find category column (by header, else fallback)
    let categoryCol = findHeaderCol_(sh, headerName);
    if (categoryCol < 1) categoryCol = Math.max(1, Number(fallbackIdx) || DEFAULT_FALLBACK_COL_INDEX);

    // 1) SORT so identical categories become contiguous
    if (sortBeforeMerge) {
      sortDataRowsBy_(sh, categoryCol, secondarySortCols);
    }

    // 2) MERGE consecutive duplicates and 3) center alignment on that column
    const result = mergeConsecutiveDuplicatesAndCenter_(sh, categoryCol);

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        target_col: result.colIndex,
        merged_groups: result.groups,
        sorted: !!sortBeforeMerge
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Find column by header text (row 1), case-insensitive, trimmed; 1-based index. */
function findHeaderCol_(sh, headerName) {
  if (!headerName) return -1;
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return -1;
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const target = String(headerName).trim().toUpperCase();
  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] == null ? "" : headerRow[c]).trim().toUpperCase();
    if (h === target) return c + 1;
  }
  return -1;
}

/** Sort data rows (2..lastRow) by primaryCol asc, optionally then by others (1-based indices). */
function sortDataRowsBy_(sh, primaryCol, secondaryCols) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);

  const specs = [{ column: primaryCol, ascending: true }];
  (secondaryCols || []).forEach(c => {
    const ci = Number(c);
    if (ci >= 1) specs.push({ column: ci, ascending: true });
  });

  dataRange.sort(specs);
}

/** Merge runs of identical, non-empty values; center them vertically & horizontally. */
function mergeConsecutiveDuplicatesAndCenter_(sh, col) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { groups: 0, colIndex: col };

  const dataRows = lastRow - 1;
  const rng = sh.getRange(2, col, dataRows, 1);
  const values = rng.getValues();
  const norm = v => String(v == null ? "" : v).trim().toUpperCase();

  // Clear any existing merges in that column (data rows) to avoid overlap errors.
  sh.getRange(2, col, dataRows, 1).breakApart();

  let groups = 0;
  let startAbs = 2;                  // sheet row where current run starts
  let prevNorm = norm(values[0][0]); // normalized first value

  for (let i = 1; i < values.length; i++) {
    const curNorm = norm(values[i][0]);
    if (curNorm !== prevNorm) {
      if (prevNorm !== "" && (i + 1) > (startAbs + 1)) {
        sh.getRange(startAbs, col, i - startAbs + 0, 1).merge();
        groups++;
      }
      startAbs = i + 2;  // convert 0-based index to absolute sheet row
      prevNorm = curNorm;
    }
  }
  // close last run
  if (prevNorm !== "" && lastRow >= startAbs + 1) {
    sh.getRange(startAbs, col, lastRow - startAbs + 1, 1).merge();
    groups++;
  }

  // Center alignment for data rows in this column
  sh.getRange(2, col, dataRows, 1)
    .setVerticalAlignment("MIDDLE")
    .setHorizontalAlignment("CENTER");

  return { groups, colIndex: col };
}
