#!/usr/bin/env python3
# dxf_to_boq_csv.py
#
# DXF → CSV (counts, lengths, areas) with:
# - INSERT footprint bbox (length/width) measured in placed orientation
# - INSERT aggregation by (block_name, layer)
# - Friendly headers: BOQ name, category, length, width
# - Optional push to Google Sheets via a Web App (no local Google auth)
#
# Usage examples:
#   python dxf_to_boq_csv.py --name VizmapDemoV1 --target-units mm
#   python dxf_to_boq_csv.py --dxf "C:\\in\\file.dxf" --out "C:\\out\\file.csv"
#   python dxf_to_boq_csv.py --dxf "C:\\in\\DXFs" --recursive --out-dir "C:\\out"
#   python dxf_to_boq_csv.py --name VizmapDemoV1 --target-units mm \
#       --gs-webapp "https://script.google.com/macros/s/XXXX/exec" \
#       --gsheet-id 1AbcDEFghiJKLmnopQRstuVWxyz0123456789 --gsheet-tab Detail
#
# CSV columns (final):
#   run_id, source_file, entity_type, qty_type, qty_value,
#   BOQ name, category, handle, remarks, length, width

from __future__ import annotations

import argparse
import csv
import uuid
import time
import math
import logging
from pathlib import Path
from typing import List, Iterable, Tuple, Optional
import json
import urllib.request

import ezdxf

# ===== EDIT THESE DEFAULT PATHS (can be overridden via CLI) =====
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"       # folder or a single .dxf file
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"   # output folder for CSVs
# ===============================================================

# ===== GOOGLE WEB APP DEFAULTS (can be overridden via CLI) =====
# If GS_WEBAPP_URL is non-empty, the script will POST results to it (no local auth).
GS_WEBAPP_URL       = "https://script.google.com/macros/s/AKfycbzuE6DQMzRPWs46h24qbX1e-Fs-vcU386zY5PAqkS2VNLiu-wwU8WGQZl6hgzsuALfZ/exec"   # e.g. "https://script.google.com/macros/s/XXXX/exec"
GSHEET_ID           = "1TbmU6vYevnhhYGFP-j91g752iFj_O0nkofAZn86akU0"   # Spreadsheet ID (the long string in the sheet URL)
GSHEET_TAB          = "Sheet2"           # detail tab name
GSHEET_SUMMARY_TAB  = ""                 # optional; blank to skip
GSHEET_MODE         = "replace"          # "replace" or "append"
# ===============================================================

# Internal field keys we populate in rows (do not change)
INTERNAL_FIELDS = [
    "run_id","source_file","entity_type","qty_type","qty_value",
    "block_name","layer","handle","remarks","bbox_length","bbox_width"
]

# Final CSV headers you wanted
CSV_HEADERS = [
    "run_id","source_file","entity_type","qty_type","qty_value",
    "BOQ name","category","handle","remarks","length","width"
]

# ---------------------------
# Helpers: run id & unit math
# ---------------------------
def make_run_id() -> str:
    ts = time.strftime("%Y%m%d-%H%M")
    rnd = uuid.uuid4().hex[:6]
    return f"r{ts}-{rnd}"

def units_scale_to_meters(doc) -> float:
    """Map DXF $INSUNITS to a scale factor from drawing units -> meters."""
    try:
        code = int(doc.header.get("$INSUNITS", 0))
    except Exception:
        code = 0
    mapping = {
        0: 0.001,  # unitless → assume mm
        1: 0.0254, # inches
        2: 0.3048, # feet
        4: 0.001,  # millimeters
        5: 0.01,   # centimeters
        6: 1.0,    # meters
    }
    scale = mapping.get(code, 1.0)
    if code not in mapping:
        logging.warning("Unrecognized $INSUNITS=%s; assuming meters (scale=1.0).", code)
    else:
        logging.info("Detected $INSUNITS=%s → scale to meters = %s", code, scale)
    return scale

def to_target_units(value_m: float, target: str, kind: str) -> float:
    t = (target or "m").lower().strip()
    if kind == "length":
        return {"m":value_m,"mm":value_m*1000,"cm":value_m*100,"ft":value_m/0.3048}.get(t, value_m)
    return {"m":value_m,"mm":value_m*1_000_000,"cm":value_m*10_000,"ft":value_m/(0.3048**2)}.get(t, value_m)

# ---------------------------
# Geometry helpers
# ---------------------------
def dist2d(p1, p2) -> float:
    return math.hypot(p2[0] - p1[0], p2[1] - p1[1])

def polyline_length_xy(points: list[tuple[float,float]], closed: bool) -> float:
    if len(points) < 2:
        return 0.0
    total = sum(dist2d(points[i], points[i+1]) for i in range(len(points)-1))
    if closed:
        total += dist2d(points[-1], points[0])
    return total

def polygon_area_xy(points: list[tuple[float,float]]) -> float:
    n = len(points)
    if n < 3: return 0.0
    s = 0.0
    for i in range(n):
        x1, y1 = points[i]
        x2, y2 = points[(i+1) % n]
        s += x1*y2 - x2*y1
    return abs(s) * 0.5

# Bounding box utilities for virtual entities
def _sample_arc_pts(cx, cy, r, start_deg: Optional[float], end_deg: Optional[float]) -> Iterable[Tuple[float,float]]:
    if r <= 0:
        return []
    if start_deg is None or end_deg is None:
        start_deg, end_deg = 0.0, 360.0
    sweep = (end_deg - start_deg) % 360.0
    steps = max(8, int(sweep / 7.5) + 1)
    for i in range(steps + 1):
        a = math.radians(start_deg + sweep * (i/steps))
        yield (cx + r * math.cos(a), cy + r * math.sin(a))

def _collect_points_from_entity(e) -> Iterable[Tuple[float,float]]:
    et = e.dxftype()
    if et == "LINE":
        yield (float(e.dxf.start.x), float(e.dxf.start.y))
        yield (float(e.dxf.end.x),   float(e.dxf.end.y))
    elif et == "LWPOLYLINE":
        for v in e:
            yield (float(v[0]), float(v[1]))
    elif et == "POLYLINE":
        for v in e.vertices():
            loc = getattr(v.dxf, "location", None)
            if loc is not None:
                yield (float(loc.x), float(loc.y))
            else:
                yield (float(getattr(v.dxf, "x", 0.0)), float(getattr(v.dxf, "y", 0.0)))
    elif et == "CIRCLE":
        cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
        r = float(e.dxf.radius)
        for p in _sample_arc_pts(cx, cy, r, None, None):
            yield p
    elif et == "ARC":
        cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
        r = float(e.dxf.radius)
        sa = float(e.dxf.start_angle)
        ea = float(e.dxf.end_angle)
        for p in _sample_arc_pts(cx, cy, r, sa, ea):
            yield p
    elif et == "HATCH":
        paths = getattr(e, "paths", None)
        if paths:
            for path in paths:
                verts = getattr(path, "polyline_path", None)
                if verts:
                    for v in verts:
                        x = float(getattr(v, "x", v[0]))
                        y = float(getattr(v, "y", v[1]))
                        yield (x, y)

def _bbox_of_insert_xy(ins) -> Optional[Tuple[float,float]]:
    try:
        minx = miny = float("inf")
        maxx = maxy = float("-inf")
        for ve in ins.virtual_entities():
            for (x, y) in _collect_points_from_entity(ve) or []:
                if x < minx: minx = x
                if y < miny: miny = y
                if x > maxx: maxx = x
                if y > maxy: maxy = y
        if minx == float("inf"):
            return None
        return (max(0.0, maxx - minx), max(0.0, maxy - miny))
    except Exception as ex:
        logging.debug("bbox for INSERT failed: %s", ex)
        return None

# ---------------------------
# CSV row builder
# ---------------------------
def make_row(run_id, source_file, entity_type, qty_type, qty_value,
             block_name="", layer="", handle="", remarks="",
             bbox_length=None, bbox_width=None) -> dict:
    return {
        "run_id": run_id, "source_file": source_file,
        "entity_type": entity_type, "qty_type": qty_type,
        "qty_value": f"{float(qty_value):.6f}" if isinstance(qty_value,(int,float,str)) else "",
        "block_name": block_name or "", "layer": layer or "",
        "handle": handle or "", "remarks": remarks or "",
        "bbox_length": "" if bbox_length is None else f"{float(bbox_length):.6f}",
        "bbox_width":  "" if bbox_width  is None else f"{float(bbox_width):.6f}",
    }

# ---------------------------
# Iterators
# ---------------------------
def iter_block_rows(msp, run_id, source_file, include_xrefs: bool,
                    scale_to_m: float, target_units: str) -> list[dict]:
    rows: list[dict] = []
    for ins in msp.query("INSERT"):
        try:
            name = getattr(ins, "effective_name", None) or getattr(ins, "block_name", None)
            if not name:
                name = ins.dxf.name if hasattr(ins.dxf, "name") else ""
            if not include_xrefs and ("|" in (name or "")):
                continue

            bbox_du = _bbox_of_insert_xy(ins)
            if bbox_du:
                dx_m = bbox_du[0] * scale_to_m
                dy_m = bbox_du[1] * scale_to_m
                dx_out = to_target_units(dx_m, target_units, "length")
                dy_out = to_target_units(dy_m, target_units, "length")
            else:
                dx_out = dy_out = None

            rows.append(make_row(
                run_id, source_file, "INSERT", "count", 1.0,
                block_name=name or "",
                layer=getattr(ins.dxf, "layer", "") or "",
                handle=getattr(ins.dxf, "handle", "") or "",
                remarks="",
                bbox_length=dx_out, bbox_width=dy_out,
            ))
        except Exception as ex:
            logging.exception("Failed INSERT: %s", ex)
    return rows

def _layer_match(layer_name: str, wanted_upper: set[str]) -> bool:
    return bool(wanted_upper) and (layer_name or "").upper() in wanted_upper

def iter_length_rows(msp, run_id, source_file, layers, scale_to_m, target_units) -> list[dict]:
    rows: list[dict] = []
    wanted = {s.upper() for s in (layers or []) if s}
    if not wanted: return rows

    for e in msp.query("LINE"):
        try:
            if not _layer_match(e.dxf.layer, wanted): continue
            p1 = (e.dxf.start.x, e.dxf.start.y)
            p2 = (e.dxf.end.x,  e.dxf.end.y)
            L_m = dist2d(p1, p2) * scale_to_m
            rows.append(make_row(run_id, source_file, "LINE", "length",
                                 to_target_units(L_m, target_units, "length"),
                                 layer=e.dxf.layer, handle=e.dxf.handle))
        except Exception as ex:
            logging.exception("LINE length: %s", ex)

    for e in msp.query("LWPOLYLINE"):
        try:
            if not _layer_match(e.dxf.layer, wanted): continue
            pts = [(float(v[0]), float(v[1])) for v in e]
            closed = bool(getattr(e, "closed", False))
            L_m = polyline_length_xy(pts, closed) * scale_to_m
            rows.append(make_row(run_id, source_file, "LWPOLYLINE", "length",
                                 to_target_units(L_m, target_units, "length"),
                                 layer=e.dxf.layer, handle=e.dxf.handle))
        except Exception as ex:
            logging.exception("LWPOLYLINE length: %s", ex)

    for e in msp.query("POLYLINE"):
        try:
            if not _layer_match(e.dxf.layer, wanted): continue
            pts = []
            for v in e.vertices():
                loc = getattr(v.dxf, "location", None)
                if loc is not None:
                    pts.append((float(loc.x), float(loc.y)))
                else:
                    pts.append((float(getattr(v.dxf, "x", 0.0)),
                                float(getattr(v.dxf, "y", 0.0))))
            closed = bool(getattr(e, "is_closed", getattr(e, "closed", False)))
            L_m = polyline_length_xy(pts, closed) * scale_to_m
            rows.append(make_row(run_id, source_file, "POLYLINE", "length",
                                 to_target_units(L_m, target_units, "length"),
                                 layer=e.dxf.layer, handle=e.dxf.handle))
        except Exception as ex:
            logging.exception("POLYLINE length: %s", ex)

    for e in msp.query("ARC"):
        try:
            if not _layer_match(e.dxf.layer, wanted): continue
            r = float(e.dxf.radius)
            sweep = (float(e.dxf.end_angle) - float(e.dxf.start_angle)) % 360.0
            L_m = (2.0 * math.pi * r) * (sweep / 360.0) * scale_to_m
            rows.append(make_row(run_id, source_file, "ARC", "length",
                                 to_target_units(L_m, target_units, "length"),
                                 layer=e.dxf.layer, handle=e.dxf.handle))
        except Exception as ex:
            logging.exception("ARC length: %s", ex)

    for e in msp.query("CIRCLE"):
        try:
            if not _layer_match(e.dxf.layer, wanted): continue
            L_m = (2.0 * math.pi * float(e.dxf.radius)) * scale_to_m
            rows.append(make_row(run_id, source_file, "CIRCLE", "length",
                                 to_target_units(L_m, target_units, "length"),
                                 layer=e.dxf.layer, handle=e.dxf.handle))
        except Exception as ex:
            logging.exception("CIRCLE length: %s", ex)

    return rows

def _lwpolyline_area_row(e, run_id, source_file, scale_to_m, target_units):
    if not bool(getattr(e, "closed", False)): return None
    pts = [(float(v[0]), float(v[1])) for v in e]
    had_bulge = any(abs((float(v[-1]) if len(v) >= 5 else 0.0)) > 1e-12 for v in e)
    A_m2 = polygon_area_xy(pts) * (scale_to_m ** 2)
    return make_row(run_id, source_file, "LWPOLYLINE", "area",
                    to_target_units(A_m2, target_units, "area"),
                    layer=e.dxf.layer, handle=e.dxf.handle,
                    remarks=("polyline area approx; bulges ignored" if had_bulge else ""))

def _polyline_area_row(e, run_id, source_file, scale_to_m, target_units):
    closed = bool(getattr(e, "is_closed", getattr(e, "closed", False)))
    if not closed: return None
    pts = []
    had_bulge = False
    for v in e.vertices():
        loc = getattr(v.dxf, "location", None)
        if loc is not None:
            pts.append((float(loc.x), float(loc.y)))
        else:
            pts.append((float(getattr(v.dxf, "x", 0.0)),
                        float(getattr(v.dxf, "y", 0.0))))
        try:
            if abs(float(getattr(v.dxf, "bulge", 0.0))) > 1e-12:
                had_bulge = True
        except Exception:
            pass
    A_m2 = polygon_area_xy(pts) * (scale_to_m ** 2)
    return make_row(run_id, source_file, "POLYLINE", "area",
                    to_target_units(A_m2, target_units, "area"),
                    layer=e.dxf.layer, handle=e.dxf.handle,
                    remarks=("polyline area approx; bulges ignored" if had_bulge else ""))

def _hatch_area_rows(e, run_id, source_file, scale_to_m, target_units) -> list[dict]:
    out: list[dict] = []
    A_m2 = None
    remarks = ""
    try:
        if hasattr(e, "get_filled_area"):
            area_du = e.get_filled_area()
            if area_du and area_du > 0:
                A_m2 = float(area_du) * (scale_to_m ** 2)
                remarks = "hatch area via API"
    except Exception:
        logging.debug("HATCH get_filled_area() failed; skipping direct area.")
    if A_m2 and A_m2 > 0.0:
        out.append(make_row(run_id, source_file, "HATCH", "area",
                            to_target_units(A_m2, target_units, "area"),
                            layer=e.dxf.layer, handle=e.dxf.handle, remarks=remarks))
        return out

    total_du = 0.0
    used_polyline_boundary = False
    try:
        paths = getattr(e, "paths", None)
        if paths:
            for path in paths:
                verts = getattr(path, "polyline_path", None)
                if verts:
                    pts = []
                    for v in verts:
                        try:
                            x = float(getattr(v, "x", v[0]))
                            y = float(getattr(v, "y", v[1]))
                            pts.append((x, y))
                        except Exception:
                            continue
                    if len(pts) >= 3:
                        total_du += polygon_area_xy(pts)
                        used_polyline_boundary = True
        if used_polyline_boundary and total_du > 0.0:
            A_m2 = total_du * (scale_to_m ** 2)
    except Exception:
        logging.debug("HATCH polyline-boundary fallback not available.")
    if A_m2 and A_m2 > 0.0:
        out.append(make_row(run_id, source_file, "HATCH", "area",
                            to_target_units(A_m2, target_units, "area"),
                            layer=e.dxf.layer, handle=e.dxf.handle,
                            remarks="hatch area approx via polyline boundary; bulges ignored"))
    return out

def iter_area_rows(msp, run_id, source_file, layers, scale_to_m, target_units) -> list[dict]:
    rows: list[dict] = []
    wanted = {s.upper() for s in (layers or []) if s}
    if not wanted: return rows

    for e in msp.query("HATCH"):
        try:
            if _layer_match(e.dxf.layer, wanted):
                rows.extend(_hatch_area_rows(e, run_id, source_file, scale_to_m, target_units))
        except Exception as ex:
            logging.exception("HATCH area: %s", ex)

    for e in msp.query("LWPOLYLINE"):
        try:
            if _layer_match(e.dxf.layer, wanted):
                row = _lwpolyline_area_row(e, run_id, source_file, scale_to_m, target_units)
                if row: rows.append(row)
        except Exception as ex:
            logging.exception("LWPOLYLINE area: %s", ex)

    for e in msp.query("POLYLINE"):
        try:
            if _layer_match(e.dxf.layer, wanted):
                row = _polyline_area_row(e, run_id, source_file, scale_to_m, target_units)
                if row: rows.append(row)
        except Exception as ex:
            logging.exception("POLYLINE area: %s", ex)

    return rows

# ---------------------------
# Aggregation for INSERTs
# ---------------------------
def _median(vals: list[float]) -> float:
    vals = sorted(vals)
    n = len(vals)
    if n == 0: return float("nan")
    return vals[n//2] if n % 2 == 1 else 0.5 * (vals[n//2 - 1] + vals[n//2])

def aggregate_insert_rows(insert_rows: list[dict], run_id: str, source_file: str) -> list[dict]:
    groups: dict[tuple[str,str], dict] = {}
    for r in insert_rows:
        key = (r.get("block_name",""), r.get("layer",""))
        g = groups.setdefault(key, {"count":0, "xs":[], "ys":[]})
        try:
            g["count"] += int(float(r.get("qty_value", "1") or 1))
        except Exception:
            g["count"] += 1
        try:
            bx = r.get("bbox_length"); by = r.get("bbox_width")
            if bx not in ("", None) and by not in ("", None):
                g["xs"].append(float(bx))
                g["ys"].append(float(by))
        except Exception:
            pass

    out: list[dict] = []
    for (name, layer), g in groups.items():
        bx = _median(g["xs"]) if g["xs"] else None
        by = _median(g["ys"]) if g["ys"] else None
        out.append(make_row(
            run_id, source_file, "INSERT", "count", float(g["count"]),
            block_name=name, layer=layer, handle="",
            remarks=f"aggregated {g['count']} inserts", bbox_length=bx, bbox_width=by
        ))
    return out

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

def _make_summary_rows(detail_rows: list[dict]) -> list[list]:
    bucket = {}
    for r in detail_rows:
        if r.get("entity_type") != "INSERT" or r.get("qty_type") != "count":
            continue
        key = (r.get("layer",""), r.get("block_name",""))
        try:
            qty = float(r.get("qty_value","0") or 0)
        except Exception:
            qty = 0.0
        bucket[key] = bucket.get(key, 0.0) + qty
    out = [["category","BOQ name","qty"]]
    for (cat, name), qty in sorted(bucket.items()):
        out.append([cat, name, qty])
    return out

def push_rows_to_webapp(rows: list[dict], webapp_url: str, spreadsheet_id: str,
                        tab: str, mode: str = "replace",
                        summary_tab: str = "") -> None:
    """POST rows to the Apps Script Web App (no local auth needed)."""
    if not webapp_url or not spreadsheet_id or not tab:
        logging.info("WebApp push not configured (missing url/id/tab). Skipping upload.")
        return

    headers = CSV_HEADERS[:]  # human-friendly headers
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
        payload["summaryRows"] = _make_summary_rows(rows)

    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        webapp_url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        resp_body = resp.read().decode("utf-8", errors="ignore")
        logging.info("WebApp response: %s", resp_body)

def print_summary(rows: list[dict], out_path: Path) -> None:
    total_inserts = sum(1 for r in rows if r["entity_type"] == "INSERT" and r["qty_type"] == "count")
    logging.info("----- SUMMARY for %s -----", out_path.name)
    logging.info("INSERT groups (after aggregation): %d", total_inserts)
    logging.info("CSV written to: %s", out_path)

# ---------------------------
# Batch helpers
# ---------------------------
def collect_dxf_files(path: Path, recursive: bool) -> List[Path]:
    if path.is_file():
        if path.suffix.lower() == ".dxf":
            return [path]
        logging.error("Provided file is not a .dxf: %s", path)
        return []
    if path.is_dir():
        pattern = "**/*.dxf" if recursive else "*.dxf"
        files = sorted(path.glob(pattern))
        if not files:
            logging.warning("No DXF files found in %s (recursive=%s)", path, recursive)
        return files
    logging.error("Path does not exist: %s", path)
    return []

def derive_out_path(dxf_path: Path, out_dir: Path | None) -> Path:
    return (out_dir / f"{dxf_path.stem}_raw_extract.csv") if out_dir else dxf_path.with_name(f"{dxf_path.stem}_raw_extract.csv")

def process_one_dxf(dxf_path: Path, out_dir: Path | None,
                    length_layers, area_layers, target_units, include_xrefs,
                    aggregate_inserts: bool) -> list[dict]:
    logging.info("Processing DXF: %s", dxf_path)
    run_id = make_run_id()
    source_file = dxf_path.name

    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    scale_to_m = units_scale_to_meters(doc)

    rows: list[dict] = []
    insert_rows = iter_block_rows(msp, run_id, source_file, include_xrefs=include_xrefs,
                                  scale_to_m=scale_to_m, target_units=target_units)
    if aggregate_inserts:
        insert_rows = aggregate_insert_rows(insert_rows, run_id, source_file)
    rows.extend(insert_rows)

    if length_layers:
        rows.extend(iter_length_rows(msp, run_id, source_file, length_layers, scale_to_m, target_units))
    if area_layers:
        rows.extend(iter_area_rows(msp, run_id, source_file, area_layers, scale_to_m, target_units))

    if not rows:
        logging.warning("No rows generated for %s. Check layer filters or DXF content.", dxf_path)

    out_path = derive_out_path(dxf_path, out_dir)
    write_csv(rows, out_path)
    print_summary(rows, out_path)
    return rows

# ---------------------------
# Main
# ---------------------------
def main():
    ap = argparse.ArgumentParser(description="DXF → CSV (counts, lengths, areas) + bbox + aggregation + optional Google Web App upload.")
    ap.add_argument("--dxf", help="Path to a DXF file OR a folder containing DXFs. Default: DXF_FOLDER.")
    ap.add_argument("--name", help="Base filename (no extension) as <DXF_FOLDER>/<NAME>.dxf.")
    ap.add_argument("--out-dir", help="Directory to write CSVs (single or batch). Default: OUT_ROOT.")
    ap.add_argument("--out", help="(Single-file only) explicit CSV path.")
    ap.add_argument("--recursive", action="store_true", help="If input is a folder, search subfolders for *.dxf.")
    ap.add_argument("--length-layers", action="append", default=[], help="Repeatable. Example: --length-layers WALL")
    ap.add_argument("--area-layers", action="append", default=[], help="Repeatable. Example: --area-layers FLR")
    ap.add_argument("--target-units", default="m", help="m, mm, cm, ft. Default: m.")
    ap.add_argument("--include-xrefs", action="store_true", help="Include INSERTs with '|' in name.")
    ap.add_argument("--no-aggregate-inserts", action="store_true", help="Write one row per INSERT instead of grouping.")
    # Web App options
    ap.add_argument("--gs-webapp", default=None, help="Apps Script Web App URL; when set, upload to Sheets via POST (no local auth).")
    ap.add_argument("--gsheet-id", default=None, help="Spreadsheet ID (default: GSHEET_ID).")
    ap.add_argument("--gsheet-tab", default=None, help="Worksheet/tab for detail rows (default: GSHEET_TAB).")
    ap.add_argument("--gsheet-summary-tab", default=None, help="Optional summary tab (default: GSHEET_SUMMARY_TAB).")
    ap.add_argument("--gsheet-mode", choices=["replace","append"], default=None, help="Write mode (default: GSHEET_MODE).")
    ap.add_argument("--verbose", action="store_true", help="Enable debug logging.")
    args = ap.parse_args()

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    # Resolve paths
    if args.dxf:
        dxf_input = Path(args.dxf)
    elif args.name:
        dxf_input = Path(DXF_FOLDER) / f"{args.name}.dxf"
    else:
        dxf_input = Path(DXF_FOLDER)

    out_dir   = Path(args.out_dir) if args.out_dir else Path(OUT_ROOT)
    explicit_out = Path(args.out) if args.out else None
    aggregate_inserts = not args.no_aggregate_inserts

    # Resolve web app settings
    gs_webapp = (args.gs_webapp if args.gs_webapp is not None else GS_WEBAPP_URL).strip()
    gsheet_id = (args.gsheet_id if args.gsheet_id is not None else GSHEET_ID).strip()
    gsheet_tab = (args.gsheet_tab if args.gsheet_tab is not None else GSHEET_TAB).strip()
    gsheet_summary_tab = (args.gsheet_summary_tab if args.gsheet_summary_tab is not None else GSHEET_SUMMARY_TAB).strip()
    gsheet_mode = (args.gsheet_mode if args.gsheet_mode is not None else GSHEET_MODE).strip().lower()

    if not dxf_input.exists():
        logging.error("DXF input path not found: %s", dxf_input)
        return

    files = collect_dxf_files(dxf_input, recursive=args.recursive)
    if not files:
        return

    # Single-file explicit CSV path
    if explicit_out:
        if len(files) != 1:
            logging.error("--out is for a single file. For folders, use --out-dir.")
            return
        f = files[0]
        rows = process_one_dxf(
            dxf_path=f,
            out_dir=explicit_out.parent,
            length_layers=args.length_layers,
            area_layers=args.area_layers,
            target_units=args.target_units,
            include_xrefs=args.include_xrefs,
            aggregate_inserts=aggregate_inserts,
        )
        explicit_out.parent.mkdir(parents=True, exist_ok=True)
        write_csv(rows, explicit_out)
        print_summary(rows, explicit_out)

        if gs_webapp and gsheet_id:
            push_rows_to_webapp(rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, gsheet_summary_tab)
        return

    # Normal (single or batch)
    out_dir = out_dir if str(out_dir).strip() else None
    all_rows: list[dict] = []
    for f in files:
        try:
            rows = process_one_dxf(
                dxf_path=f,
                out_dir=out_dir,
                length_layers=args.length_layers,
                area_layers=args.area_layers,
                target_units=args.target_units,
                include_xrefs=args.include_xrefs,
                aggregate_inserts=aggregate_inserts,
            )
            all_rows.extend(rows or [])
        except Exception as ex:
            logging.exception("Failed processing %s: %s", f, ex)

    if all_rows and gs_webapp and gsheet_id:
        push_rows_to_webapp(all_rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, gsheet_summary_tab)

if __name__ == "__main__":
    main()
