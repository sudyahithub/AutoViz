#!/usr/bin/env python3
# dxf_to_boq_csv.py
#
# DXF → CSV with:
# - INSERT bbox (length/width), aggregation by (block_name, layer)
# - Per-layer totals: length (all line-like) & area (all closed shapes)
# - "category" is layer; blank layer -> "misc"
# - Optional push to Google Sheets via Apps Script Web App (no local Google auth)
#
# Example:
#   python dxf_to_boq_csv.py --name VizmapDemoV1 --target-units mm \
#     --gs-webapp "https://script.google.com/macros/s/AKfycbzuE6DQMzRPWs46h24qbX1e-Fs-vcU386zY5PAqkS2VNLiu-wwU8WGQZl6hgzsuALfZ/exec" \
#     --gsheet-id 1TbmU6vYevnhhYGFP-j91g752iFj_O0nkofAZn86akU0 --gsheet-tab Sheet2
#
# CSV columns (final):
#   run_id, source_file, entity_type, qty_type, qty_value,
#   BOQ name, category, handle, remarks, length, width

from __future__ import annotations

import argparse, csv, uuid, time, math, logging, json
from pathlib import Path
from typing import List, Iterable, Tuple, Optional
import urllib.request

import ezdxf

# ===== DEFAULT PATHS (CLI can override) =====
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"       # folder or a single .dxf file
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"   # output folder for CSVs
# ===========================================

# ===== GOOGLE WEB APP DEFAULTS (can be overridden via CLI) =====
# If GS_WEBAPP_URL is non-empty, the script will POST results to it (no local auth).
GS_WEBAPP_URL       = "https://script.google.com/macros/s/AKfycbzuE6DQMzRPWs46h24qbX1e-Fs-vcU386zY5PAqkS2VNLiu-wwU8WGQZl6hgzsuALfZ/exec"
GSHEET_ID           = "1TbmU6vYevnhhYGFP-j91g752iFj_O0nkofAZn86akU0"
GSHEET_TAB          = "Sheet2"
GSHEET_SUMMARY_TAB  = ""                  # optional; blank to skip
GSHEET_MODE         = "replace"           # "replace" or "append"
# ===============================================================

CSV_HEADERS = [
    "run_id","source_file","entity_type","qty_type","qty_value",
    "BOQ name","category","handle","remarks","length","width"
]

# ---------------------------
# Utilities
# ---------------------------
def make_run_id() -> str:
    ts = time.strftime("%Y%m%d-%H%M"); rnd = uuid.uuid4().hex[:6]
    return f"r{ts}-{rnd}"

def layer_or_misc(name: str) -> str:
    s = (name or "").strip()
    return s if s else "misc"

def units_scale_to_meters(doc) -> float:
    try: code = int(doc.header.get("$INSUNITS", 0))
    except Exception: code = 0
    mapping = {0:0.001,1:0.0254,2:0.3048,4:0.001,5:0.01,6:1.0}
    scale = mapping.get(code, 1.0)
    if code not in mapping:
        logging.warning("Unrecognized $INSUNITS=%s; assuming meters.", code)
    else:
        logging.info("Detected $INSUNITS=%s → %s m/unit", code, scale)
    return scale

def to_target_units(v_m: float, target: str, kind: str) -> float:
    t = (target or "m").lower().strip()
    if kind=="length":
        return {"m":v_m,"mm":v_m*1000,"cm":v_m*100,"ft":v_m/0.3048}.get(t, v_m)
    return {"m":v_m,"mm":v_m*1_000_000,"cm":v_m*10_000,"ft":v_m/(0.3048**2)}.get(t, v_m)

def dist2d(p1, p2) -> float: return math.hypot(p2[0]-p1[0], p2[1]-p1[1])

def polyline_length_xy(pts: list[tuple[float,float]], closed: bool) -> float:
    if len(pts)<2: return 0.0
    L = sum(dist2d(pts[i], pts[i+1]) for i in range(len(pts)-1))
    if closed: L += dist2d(pts[-1], pts[0])
    return L

def polygon_area_xy(pts: list[tuple[float,float]]) -> float:
    n=len(pts); 
    if n<3: return 0.0
    s=0.0
    for i in range(n):
        x1,y1=pts[i]; x2,y2=pts[(i+1)%n]; s += x1*y2 - x2*y1
    return abs(s)*0.5

def _sample_arc_pts(cx, cy, r, start_deg: Optional[float], end_deg: Optional[float]):
    if r<=0: return []
    if start_deg is None or end_deg is None: start_deg, end_deg = 0.0, 360.0
    sweep = (end_deg - start_deg) % 360.0
    steps = max(8, int(sweep/7.5)+1)
    for i in range(steps+1):
        a=math.radians(start_deg + sweep*(i/steps))
        yield (cx + r*math.cos(a), cy + r*math.sin(a))

def _collect_points_from_entity(e):
    et = e.dxftype()
    if et=="LINE":
        yield (float(e.dxf.start.x), float(e.dxf.start.y))
        yield (float(e.dxf.end.x),   float(e.dxf.end.y))
    elif et=="LWPOLYLINE":
        for v in e: yield (float(v[0]), float(v[1]))
    elif et=="POLYLINE":
        for v in e.vertices():
            loc=getattr(v.dxf,"location",None)
            yield (float(loc.x),float(loc.y)) if loc is not None else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0)))
    elif et=="CIRCLE":
        cx,cy=float(e.dxf.center.x), float(e.dxf.center.y); r=float(e.dxf.radius)
        yield from _sample_arc_pts(cx,cy,r,None,None)
    elif et=="ARC":
        cx,cy=float(e.dxf.center.x), float(e.dxf.center.y); r=float(e.dxf.radius)
        sa,ea=float(e.dxf.start_angle), float(e.dxf.end_angle)
        yield from _sample_arc_pts(cx,cy,r,sa,ea)
    elif et=="HATCH":
        paths=getattr(e,"paths",None)
        if paths:
            for path in paths:
                verts=getattr(path,"polyline_path",None)
                if verts:
                    for v in verts:
                        x=float(getattr(v,"x",v[0])); y=float(getattr(v,"y",v[1]))
                        yield (x,y)

def _bbox_of_insert_xy(ins) -> Optional[Tuple[float,float]]:
    try:
        minx=miny=float("inf"); maxx=maxy=float("-inf")
        for ve in ins.virtual_entities():
            for (x,y) in _collect_points_from_entity(ve) or []:
                if x<minx: minx=x
                if y<miny: miny=y
                if x>maxx: maxx=x
                if y>maxy: maxy=y
        if minx==float("inf"): return None
        return (max(0.0, maxx-minx), max(0.0, maxy-miny))
    except Exception as ex:
        logging.debug("bbox for INSERT failed: %s", ex)
        return None

def make_row(run_id, source_file, entity_type, qty_type, qty_value,
             block_name="", layer="", handle="", remarks="",
             bbox_length=None, bbox_width=None) -> dict:
    return {
        "run_id": run_id, "source_file": source_file,
        "entity_type": entity_type, "qty_type": qty_type,
        "qty_value": f"{float(qty_value):.6f}" if isinstance(qty_value,(int,float,str)) else "",
        "block_name": block_name or "", "layer": layer_or_misc(layer),
        "handle": handle or "", "remarks": remarks or "",
        "bbox_length": "" if bbox_length is None else f"{float(bbox_length):.6f}",
        "bbox_width":  "" if bbox_width  is None else f"{float(bbox_width):.6f}",
    }

# ---------------------------
# Detailed rows
# ---------------------------
def iter_block_rows(msp, run_id, source_file, include_xrefs: bool,
                    scale_to_m: float, target_units: str) -> list[dict]:
    out=[]
    for ins in msp.query("INSERT"):
        try:
            name = getattr(ins, "effective_name", None) or getattr(ins, "block_name", None) or getattr(ins.dxf,"name","")
            if not include_xrefs and ("|" in (name or "")):  # skip xref|block unless asked
                continue
            bbox_du = _bbox_of_insert_xy(ins)
            if bbox_du:
                dx_m = bbox_du[0] * scale_to_m
                dy_m = bbox_du[1] * scale_to_m
                dx_out = to_target_units(dx_m, target_units, "length")
                dy_out = to_target_units(dy_m, target_units, "length")
            else:
                dx_out = dy_out = None
            out.append(make_row(
                run_id, source_file, "INSERT", "count", 1.0,
                block_name=name,
                layer=getattr(ins.dxf,"layer",""),
                handle=getattr(ins.dxf,"handle",""),
                remarks="", bbox_length=dx_out, bbox_width=dy_out
            ))
        except Exception as ex:
            logging.exception("INSERT failed: %s", ex)
    return out

# ---------------------------
# Per-layer metrics
# ---------------------------
def compute_layer_metrics(msp, scale_to_m: float, target_units: str):
    length_by_layer = {}
    area_by_layer   = {}

    def add_len(layer, L_du):
        L_m = L_du * scale_to_m
        L_out = to_target_units(L_m, target_units, "length")
        k = layer_or_misc(layer)
        length_by_layer[k] = length_by_layer.get(k, 0.0) + L_out

    def add_area(layer, A_du):
        A_m2 = A_du * (scale_to_m**2)
        A_out = to_target_units(A_m2, target_units, "area")
        k = layer_or_misc(layer)
        area_by_layer[k] = area_by_layer.get(k, 0.0) + A_out

    # LINE
    for e in msp.query("LINE"):
        try:
            p1=(e.dxf.start.x,e.dxf.start.y); p2=(e.dxf.end.x,e.dxf.end.y)
            add_len(e.dxf.layer, dist2d(p1,p2))
        except Exception: pass

    # LWPOLYLINE
    for e in msp.query("LWPOLYLINE"):
        try:
            pts=[(float(v[0]),float(v[1])) for v in e]
            closed=bool(getattr(e,"closed",False))
            add_len(e.dxf.layer, polyline_length_xy(pts, closed))
            if closed and len(pts)>=3:
                # bulges ignored
                add_area(e.dxf.layer, polygon_area_xy(pts))
        except Exception: pass

    # POLYLINE
    for e in msp.query("POLYLINE"):
        try:
            pts=[]
            for v in e.vertices():
                loc=getattr(v.dxf,"location",None)
                if loc is not None: pts.append((float(loc.x),float(loc.y)))
                else: pts.append((float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
            closed = bool(getattr(e,"is_closed",getattr(e,"closed",False)))
            add_len(e.dxf.layer, polyline_length_xy(pts, closed))
            if closed and len(pts)>=3:
                add_area(e.dxf.layer, polygon_area_xy(pts))
        except Exception: pass

    # ARC
    for e in msp.query("ARC"):
        try:
            r=float(e.dxf.radius)
            sweep=(float(e.dxf.end_angle)-float(e.dxf.start_angle))%360.0
            add_len(e.dxf.layer, (2.0*math.pi*r)*(sweep/360.0))
        except Exception: pass

    # CIRCLE
    for e in msp.query("CIRCLE"):
        try:
            r=float(e.dxf.radius)
            add_len(e.dxf.layer, (2.0*math.pi*r))
            add_area(e.dxf.layer, math.pi*(r**2))
        except Exception: pass

    # HATCH (area only; try direct area or polyline boundaries)
    for e in msp.query("HATCH"):
        try:
            A_du=None
            if hasattr(e,"get_filled_area"):
                try:
                    area_du=e.get_filled_area()
                    if area_du and area_du>0: A_du=float(area_du)
                except Exception: A_du=None
            if A_du is None:
                total=0.0; used=False
                paths=getattr(e,"paths",None)
                if paths:
                    for path in paths:
                        verts=getattr(path,"polyline_path",None)
                        if verts:
                            pts=[]
                            for v in verts:
                                x=float(getattr(v,"x",v[0])); y=float(getattr(v,"y",v[1]))
                                pts.append((x,y))
                            if len(pts)>=3:
                                total += polygon_area_xy(pts); used=True
                if used: A_du=total
            if A_du and A_du>0: add_area(e.dxf.layer, A_du)
        except Exception: pass

    return length_by_layer, area_by_layer

# ---------------------------
# CSV / Web App I/O
# ---------------------------
def _map_to_csv_headers(row: dict) -> dict:
    return {
        "run_id": row.get("run_id",""),
        "source_file": row.get("source_file",""),
        "entity_type": row.get("entity_type",""),
        "qty_type": row.get("qty_type",""),
        "qty_value": row.get("qty_value",""),
        "BOQ name": row.get("block_name",""),
        "category": row.get("layer",""),
        "handle": row.get("handle",""),
        "remarks": row.get("remarks",""),
        "length": row.get("bbox_length",""),
        "width": row.get("bbox_width",""),
    }

def write_csv(rows: list[dict], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        writer.writeheader()
        for r in rows:
            writer.writerow(_map_to_csv_headers(r))

def _row_to_csv_order(row: dict) -> list:
    return [
        row.get("run_id",""),
        row.get("source_file",""),
        row.get("entity_type",""),
        row.get("qty_type",""),
        row.get("qty_value",""),
        row.get("block_name",""),   # BOQ name
        row.get("layer",""),        # category
        row.get("handle",""),
        row.get("remarks",""),
        row.get("bbox_length",""),  # length
        row.get("bbox_width",""),   # width
    ]

def _make_insert_summary(detail_rows: list[dict]) -> list[list]:
    bucket={}
    for r in detail_rows:
        if r.get("entity_type")!="INSERT" or r.get("qty_type")!="count": continue
        key=(r.get("layer",""), r.get("block_name",""))
        qty=float(r.get("qty_value","0") or 0)
        bucket[key]=bucket.get(key,0.0)+qty
    out=[["category","BOQ name","qty"]]
    for (cat,name),qty in sorted(bucket.items()):
        out.append([cat,name,qty])
    return out

def push_rows_to_webapp(rows: list[dict], webapp_url: str, spreadsheet_id: str,
                        tab: str, mode: str = "replace",
                        summary_tab: str = "", insert_summary: bool = True) -> None:
    """POST rows to the Apps Script Web App (no local auth needed)."""
    if not webapp_url or not spreadsheet_id or not tab:
        logging.info("WebApp push not configured (missing url/id/tab). Skipping upload.")
        return

    headers = CSV_HEADERS[:]
    data_rows = [_row_to_csv_order(r) for r in rows]
    payload = {
        "sheetId": spreadsheet_id,
        "tab": tab,
        "mode": (mode or "replace").lower(),
        "headers": headers,
        "rows": data_rows,
    }
    if summary_tab:
        payload["summaryTab"] = summary_tab
        payload["summaryRows"] = _make_insert_summary(rows) if insert_summary else []

    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        webapp_url, data=body,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        resp_body = resp.read().decode("utf-8", errors="ignore")
        logging.info("WebApp response: %s", resp_body)

def print_summary(rows: list[dict], out_path: Path) -> None:
    total_insert_groups = sum(1 for r in rows if r["entity_type"]=="INSERT" and r["qty_type"]=="count")
    logging.info("----- SUMMARY for %s -----", out_path.name)
    logging.info("INSERT groups (after aggregation): %d", total_insert_groups)
    logging.info("CSV written to: %s", out_path)

# ---------------------------
# Batch helpers
# ---------------------------
def collect_dxf_files(path: Path, recursive: bool) -> List[Path]:
    if path.is_file():
        if path.suffix.lower() == ".dxf": return [path]
        logging.error("Provided file is not a .dxf: %s", path); return []
    if path.is_dir():
        pattern="**/*.dxf" if recursive else "*.dxf"
        files=sorted(path.glob(pattern))
        if not files: logging.warning("No DXF files found in %s (recursive=%s)", path, recursive)
        return files
    logging.error("Path does not exist: %s", path); return []

def derive_out_path(dxf_path: Path, out_dir: Path | None) -> Path:
    return (out_dir / f"{dxf_path.stem}_raw_extract.csv") if out_dir else dxf_path.with_name(f"{dxf_path.stem}_raw_extract.csv")

# ---------------------------
# Main processing
# ---------------------------
def process_one_dxf(dxf_path: Path, out_dir: Path | None,
                    target_units: str, include_xrefs: bool,
                    layer_metrics: bool, aggregate_inserts: bool) -> list[dict]:
    logging.info("Processing DXF: %s", dxf_path)
    run_id = make_run_id()
    source_file = dxf_path.name

    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    scale_to_m = units_scale_to_meters(doc)

    rows: list[dict] = []

    # INSERT rows (then aggregate by name+layer)
    insert_rows = iter_block_rows(msp, run_id, source_file, include_xrefs, scale_to_m, target_units)
    if aggregate_inserts:
        # aggregate
        groups={}
        for r in insert_rows:
            key=(r["block_name"], r["layer"])
            g=groups.setdefault(key, {"count":0,"xs":[],"ys":[]})
            g["count"] += 1
            try:
                if r["bbox_length"] and r["bbox_width"]:
                    g["xs"].append(float(r["bbox_length"]))
                    g["ys"].append(float(r["bbox_width"]))
            except Exception: pass
        for (name, layer), g in groups.items():
            # robust representative size = median
            xs=sorted(g["xs"]); ys=sorted(g["ys"])
            bx = xs[len(xs)//2] if xs else None
            by = ys[len(ys)//2] if ys else None
            rows.append(make_row(
                run_id, source_file, "INSERT", "count", float(g["count"]),
                block_name=name, layer=layer, handle="",
                remarks=f"aggregated {g['count']} inserts", bbox_length=bx, bbox_width=by
            ))
    else:
        rows.extend(insert_rows)

    # Per-layer metrics
    # Per-layer metrics
    if layer_metrics:
        L_by, A_by = compute_layer_metrics(msp, scale_to_m, target_units)

        # length totals per layer → also fill the 'length' column
        for ly, val in sorted(L_by.items()):
            rows.append(make_row(
                run_id, source_file,
                entity_type="LAYER_SUMMARY",
                qty_type="length",
                qty_value=val,
                block_name="",
                layer=ly,
                remarks="total length per layer",
                bbox_length=val,     # <-- fill Length column
                bbox_width=None,     # leave Width blank
            ))

        # area totals per layer → put the value into the 'width' column
        for ly, val in sorted(A_by.items()):
            rows.append(make_row(
                run_id, source_file,
                entity_type="LAYER_SUMMARY",
                qty_type="area",
                qty_value=val,
                block_name="",
                layer=ly,
                remarks="total area per layer",
                bbox_length=None,    # leave Length blank
                bbox_width=val,      # <-- fill Width column
            ))


    out_path = derive_out_path(dxf_path, out_dir)
    write_csv(rows, out_path)
    print_summary(rows, out_path)
    return rows

# ---------------------------
# CLI
# ---------------------------
def main():
    ap = argparse.ArgumentParser(description="DXF → CSV (INSERT bbox + per-layer length/area) + optional Google Web App upload.")
    ap.add_argument("--dxf", help="Path to a DXF file OR a folder containing DXFs. Default: DXF_FOLDER.")
    ap.add_argument("--name", help="Base filename (no extension) as <DXF_FOLDER>/<NAME>.dxf.")
    ap.add_argument("--out-dir", help="Directory to write CSVs (single or batch). Default: OUT_ROOT.")
    ap.add_argument("--out", help="(Single-file only) explicit CSV path.")
    ap.add_argument("--recursive", action="store_true", help="If input is a folder, search subfolders for *.dxf.")
    ap.add_argument("--target-units", default="m", help="m, mm, cm, ft. Default: m.")
    ap.add_argument("--include-xrefs", action="store_true", help="Include INSERTs with '|' in name.")
    ap.add_argument("--no-layer-metrics", action="store_true", help="Disable layer length/area summary rows.")
    ap.add_argument("--no-aggregate-inserts", action="store_true", help="Write one row per INSERT instead of grouping.")
    # Web App options
    ap.add_argument("--gs-webapp", default=None, help="Apps Script Web App URL (default: GS_WEBAPP_URL).")
    ap.add_argument("--gsheet-id", default=None, help="Spreadsheet ID (default: GSHEET_ID).")
    ap.add_argument("--gsheet-tab", default=None, help="Worksheet/tab for detail rows (default: GSHEET_TAB).")
    ap.add_argument("--gsheet-summary-tab", default=None, help="Optional summary tab (default: GSHEET_SUMMARY_TAB).")
    ap.add_argument("--gsheet-mode", choices=["replace","append"], default=None, help="Write mode (default: GSHEET_MODE).")
    ap.add_argument("--verbose", action="store_true", help="Enable debug logging.")
    args = ap.parse_args()

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    # Resolve input/output
    if args.dxf:
        dxf_input = Path(args.dxf)
    elif args.name:
        dxf_input = Path(DXF_FOLDER) / f"{args.name}.dxf"
    else:
        dxf_input = Path(DXF_FOLDER)

    out_dir   = Path(args.out_dir) if args.out_dir else Path(OUT_ROOT)
    explicit_out = Path(args.out) if args.out else None
    layer_metrics = not args.no_layer_metrics
    aggregate_inserts = not args.no_aggregate_inserts

    # Resolve Web App settings
    gs_webapp = (args.gs_webapp if args.gs_webapp is not None else GS_WEBAPP_URL).strip()
    gsheet_id = (args.gsheet_id if args.gsheet_id is not None else GSHEET_ID).strip()
    gsheet_tab = (args.gsheet_tab if args.gsheet_tab is not None else GSHEET_TAB).strip()
    gsheet_summary_tab = (args.gsheet_summary_tab if args.gsheet_summary_tab is not None else GSHEET_SUMMARY_TAB).strip()
    gsheet_mode = (args.gsheet_mode if args.gsheet_mode is not None else GSHEET_MODE).strip().lower()

    if not dxf_input.exists():
        logging.error("DXF input path not found: %s", dxf_input); return

    files = collect_dxf_files(dxf_input, recursive=args.recursive)
    if not files: return

    # Single-file explicit output
    if explicit_out:
        if len(files) != 1:
            logging.error("--out is for a single file. For folders, use --out-dir."); return
        f = files[0]
        rows = process_one_dxf(
            dxf_path=f, out_dir=explicit_out.parent,
            target_units=args.target_units, include_xrefs=args.include_xrefs,
            layer_metrics=layer_metrics, aggregate_inserts=aggregate_inserts,
        )
        explicit_out.parent.mkdir(parents=True, exist_ok=True)
        write_csv(rows, explicit_out)
        print_summary(rows, explicit_out)
        if gs_webapp and gsheet_id:
            push_rows_to_webapp(rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, gsheet_summary_tab)
        return

    # Batch
    out_dir = out_dir if str(out_dir).strip() else None
    all_rows: list[dict] = []
    for f in files:
        try:
            rows = process_one_dxf(
                dxf_path=f, out_dir=out_dir,
                target_units=args.target_units, include_xrefs=args.include_xrefs,
                layer_metrics=layer_metrics, aggregate_inserts=aggregate_inserts,
            )
            all_rows.extend(rows or [])
        except Exception as ex:
            logging.exception("Failed processing %s: %s", f, ex)

    if all_rows and gs_webapp and gsheet_id:
        push_rows_to_webapp(all_rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, gsheet_summary_tab)

if __name__ == "__main__":
    main()
