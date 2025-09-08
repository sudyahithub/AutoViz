/** ===================== CONFIG ===================== */
const HEADER_ROWS  = 1;                   // number of header rows
const CATEGORY_HEADER_NAME = "category";  // header text to find (case-insensitive)
const SKIP_VALUES = new Set([""]);   // don't merge these values
const TARGET_SHEET = null;                // e.g. "Sheet2" to restrict; null = any sheet
/** ================================================== */

/** Menu */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("BOQ Tools")
    .addItem("Sort by Category & Merge (now)", "sortByCategoryAndMerge")
    .addItem("Enable AUTO (on edit)", "setupAutoEdit")
    .addItem("Disable AUTO", "disableAuto")
    .addToUi();
}

/** Enable/disable auto (optional) */
function setupAutoEdit() {
  disableAuto();
  ScriptApp.newTrigger("autoSortMergeOnEdit")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert("✅ Auto enabled: runs after each edit/paste.");
}
function disableAuto() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction && t.getHandlerFunction();
    if (fn === "autoSortMergeOnEdit") ScriptApp.deleteTrigger(t);
  });
}
function autoSortMergeOnEdit(e) {
  try { sortByCategoryAndMerge(); } catch (err) { console.error(err); }
}

/** ===== Main: NO helper columns, preserves column structure ===== */
function sortByCategoryAndMerge() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (TARGET_SHEET && sh.getName() !== TARGET_SHEET) return;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= HEADER_ROWS || lastCol < 1) return;

  // Remove basic filter (merges & sorts are blocked by active filter)
  const filter = sh.getFilter();
  if (filter) filter.remove();

  // UNMERGE the entire data body BEFORE sorting
  const startRow = HEADER_ROWS + 1;
  const numRows  = lastRow - HEADER_ROWS;
  const dataBody = sh.getRange(startRow, 1, numRows, lastCol);
  dataBody.breakApart();

  // Find the category column by header name
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const normHeader = s => String(s || "")
    .replace(/\uFEFF/g, "").replace(/\u00A0/g, " ").replace(/[\u2000-\u200D\u2060]/g, " ")
    .replace(/\s+/g, " ").trim().toLowerCase();
  let catCol = -1;
  for (let c = 0; c < header.length; c++) {
    if (normHeader(header[c]) === CATEGORY_HEADER_NAME) { catCol = c + 1; break; }
  }
  if (catCol === -1) return; // no "category" header

  // Normalize category text IN PLACE (so values compare equal)
  const catRange = sh.getRange(startRow, catCol, numRows, 1);
  const cleaned = catRange.getValues().map(r => [normalizeCategory(r[0])]);
  catRange.setNumberFormat('@');  // force plain text
  catRange.setValues(cleaned);

  // Sort the data body BY THE CATEGORY COLUMN ITSELF (no helper col)
  sh.getRange(startRow, 1, numRows, lastCol)
    .sort([{ column: catCol, ascending: true }]);

  // Merge contiguous identical categories
  const catSorted = sh.getRange(startRow, catCol, numRows, 1);
  const vals = catSorted.getValues().map(r => r[0]);
  let blockStart = startRow, prev = vals[0];

  for (let i = 1; i <= numRows; i++) {
    const atEnd = i === numRows;
    const cur = atEnd ? "__END__" : vals[i];
    if (cur !== prev) {
      const len = (startRow + i) - blockStart;
      if (!SKIP_VALUES.has(prev) && len > 1) {
        const r = sh.getRange(blockStart, catCol, len, 1);
        r.merge();  // vertical merge for single column
        r.setVerticalAlignment("middle").setHorizontalAlignment("center");
      }
      blockStart = startRow + i;
      prev = cur;
    }
  }

  SpreadsheetApp.flush();
}

/** Normalization for stable equality */
function normalizeCategory(val) {
  if (val === null || val === undefined) return "";
  let s = String(val);
  s = s.replace(/\uFEFF/g, "")
       .replace(/\u00A0/g, " ")
       .replace(/[\u2000-\u200D\u2060]/g, " ")
       .replace(/[\u2010-\u2015\u2212]/g, "-")
       .replace(/\s+/g, " ")
       .trim()
       .toUpperCase();
  return s;
}
