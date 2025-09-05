/** ===================== CONFIG ===================== */
const HEADER_ROWS  = 1;                  // number of header rows
const CATEGORY_HEADER = "category";      // header to find (case-insensitive)
const SKIP_VALUES = new Set(["", "0"]);  // don't merge these labels
const TARGET_SHEET = null;               // set a tab name to restrict; null = current sheet
const TMP_SHEET_NAME = ".__tmp_sort_boq"; // temp sheet used for safe sort
/** ================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("BOQ Tools")
    .addItem("1) Dry-run: color groups", "dryRunColorGroups")
    .addItem("2) Merge contiguous (no sort)", "mergeContiguousNoSort")
    .addItem("3) Stable sort by Category + merge", "stableSortAndMerge")
    .addToUi();
}

/* ---------- utilities ---------- */

function _getContext_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (TARGET_SHEET && sh.getName() !== TARGET_SHEET) return {ok:false, msg:"Wrong sheet", sh:null};
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= HEADER_ROWS || lastCol < 1) return {ok:false, msg:"No data", sh:null};

  // find category column
  const header = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  const n = (s)=>String(s||"").replace(/\uFEFF/g,"").replace(/\u00A0/g," ")
     .replace(/[\u2000-\u200D\u2060]/g," ").replace(/\s+/g," ").trim().toLowerCase();
  const catCol = header.findIndex(h => n(h) === n(CATEGORY_HEADER)) + 1;
  if (!catCol) return {ok:false, msg:`Header "${CATEGORY_HEADER}" not found`, sh:null};

  return {ok:true, ss, sh, lastRow, lastCol, catCol, startRow: HEADER_ROWS+1, numRows: lastRow-HEADER_ROWS};
}

function _normalizeCategoryCell(val) {
  if (val === null || val === undefined) return "";
  let s = String(val);
  s = s.replace(/\uFEFF/g,"")
       .replace(/\u00A0/g," ")
       .replace(/[\u2000-\u200D\u2060]/g," ")
       .replace(/[\u2010-\u2015\u2212]/g,"-")
       .replace(/\s+/g," ")
       .trim()
       .toUpperCase();
  return s;
}

function _removeBasicFilterAndUnmerge(sh, startRow, numRows, lastCol) {
  const f = sh.getFilter(); if (f) f.remove();
  sh.getRange(startRow, 1, numRows, lastCol).breakApart();
}

/* ---------- 1) Dry-run color groups ---------- */
function dryRunColorGroups() {
  const ctx = _getContext_();
  if (!ctx.ok) { SpreadsheetApp.getUi().alert(ctx.msg); return; }
  const {sh, lastCol, catCol, startRow, numRows} = ctx;

  // normalize Category cells only (so groups are accurate)
  const catR = sh.getRange(startRow, catCol, numRows, 1);
  const cats = catR.getValues().map(r => [_normalizeCategoryCell(r[0])]);
  catR.setNumberFormat('@').setValues(cats);

  // clear existing backgrounds in data region
  sh.getRange(startRow, 1, numRows, lastCol).setBackground(null);

  // color each contiguous block (no merges)
  const flat = cats.map(r => r[0]);
  let i0 = 0;
  for (let i = 1; i <= flat.length; i++) {
    const atEnd = i === flat.length;
    const diff = atEnd ? true : (flat[i] !== flat[i0]);
    if (diff) {
      const label = flat[i0];
      const len = i - i0;
      if (!SKIP_VALUES.has(label) && len > 0) {
        const rng = sh.getRange(startRow + i0, 1, len, lastCol);
        rng.setBackground(_pastel());
      }
      i0 = i;
    }
  }
  _toast("Dry-run colored groups. If colors look right, run #2 or #3.");
}

/* ---------- 2) Merge contiguous (no sort) ---------- */
function mergeContiguousNoSort() {
  const ctx = _getContext_();
  if (!ctx.ok) { SpreadsheetApp.getUi().alert(ctx.msg); return; }
  const {sh, lastCol, catCol, startRow, numRows} = ctx;

  _removeBasicFilterAndUnmerge(sh, startRow, numRows, lastCol);

  // normalize Category cells only
  const catR = sh.getRange(startRow, catCol, numRows, 1);
  const cats = catR.getValues().map(r => [_normalizeCategoryCell(r[0])]);
  catR.setNumberFormat('@').setValues(cats);

  // merge contiguous equal labels in the category column
  _mergeRuns(sh, catCol, startRow, cats.map(r => r[0]));

  _toast("Merged contiguous groups (no sorting).");
}

/* ---------- 3) Stable sort + merge (preserve formatting/formulas) ---------- */
function stableSortAndMerge() {
  const ctx = _getContext_();
  if (!ctx.ok) { SpreadsheetApp.getUi().alert(ctx.msg); return; }
  const {ss, sh, lastCol, catCol, startRow, numRows} = ctx;

  _removeBasicFilterAndUnmerge(sh, startRow, numRows, lastCol);

  // normalize Category column values ONLY
  const catR = sh.getRange(startRow, catCol, numRows, 1);
  const cats = catR.getValues().map(r => [_normalizeCategoryCell(r[0])]);
  catR.setNumberFormat('@').setValues(cats);

  // build stable order (category, then original row index)
  const keys = cats.map((r, i) => ({cat:r[0], orig:i, from:startRow+i}));
  keys.sort((a,b)=> (a.cat<b.cat?-1:a.cat>b.cat?1:a.orig-b.orig));

  // create hidden temp sheet
  let tmp = ss.getSheetByName(TMP_SHEET_NAME);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(TMP_SHEET_NAME);
  tmp.hideSheet();

  // ensure enough columns
  if (tmp.getMaxColumns() < lastCol) tmp.insertColumnsAfter(tmp.getMaxColumns(), lastCol - tmp.getMaxColumns());

  // copy header with formats/formulas
  sh.getRange(1, 1, HEADER_ROWS, lastCol).copyTo(tmp.getRange(1, 1, HEADER_ROWS, lastCol), {contentsOnly:false});

  // copy rows in stable order, preserving formulas/formatting
  keys.forEach((k, i) => {
    sh.getRange(k.from, 1, 1, lastCol).copyTo(tmp.getRange(startRow + i, 1, 1, lastCol), {contentsOnly:false});
  });

  // write sorted block back
  tmp.getRange(startRow, 1, numRows, lastCol).copyTo(sh.getRange(startRow, 1, numRows, lastCol), {contentsOnly:false});
  ss.deleteSheet(tmp);

  // merge contiguous equal labels
  const sortedCats = sh.getRange(startRow, catCol, numRows, 1).getValues().map(r => r[0]);
  _mergeRuns(sh, catCol, startRow, sortedCats);

  _toast("Stable-sorted by Category and merged.");
}

/* ---------- helpers ---------- */

function _mergeRuns(sh, col, startRow, labels) {
  // clear previous merges in the category column slice
  sh.getRange(startRow, col, labels.length, 1).breakApart();

  let i0 = 0;
  for (let i = 1; i <= labels.length; i++) {
    const atEnd = i === labels.length;
    const diff = atEnd ? true : (labels[i] !== labels[i0]);
    if (diff) {
      const lab = labels[i0];
      const len = i - i0;
      if (!SKIP_VALUES.has(lab) && len > 1) {
        const r = sh.getRange(startRow + i0, col, len, 1);
        r.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
      }
      i0 = i;
    }
  }
}

function _pastel() {
  const ch = ()=> Math.floor(180 + Math.random()*60).toString(16).padStart(2,"0");
  return `#${ch()}${ch()}${ch()}`;
}

function _toast(msg) {
  SpreadsheetApp.getActive().toast(msg, "BOQ Tools", 5);
}
