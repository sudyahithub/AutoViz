# python tr.py --name VizmapDemoV1 --length-layers WALL --area-layers FLR --target-units m --verbose


#!/usr/bin/env python3
# dxf_to_boq_csv.py
#
# Quick DXF → CSV raw quantities extractor (counts, lengths, areas) + INSERT bbox size.
# - Python 3.9+
# - Only ezdxf + stdlib (argparse, csv, uuid, time, math, pathlib, logging)
# - Windows & Linux
#
# CSV header:
#   run_id,source_file,entity_type,qty_type,qty_value,block_name,layer,handle,remarks,bbox_length,bbox_width
#
# ── CLI QUICK START ─────────────────────────────────────────────────────────────
# 1) Process a single file by NAME from DXF_FOLDER (writes CSV to OUT_ROOT):
#      python dxf_to_boq_csv.py --name VizmapDemoV1
# 2) Add layers and set units:
#      python dxf_to_boq_csv.py --name VizmapDemoV1 \
#          --length-layers WALL --area-layers FLR --target-units m
# 3) Explicit DXF and CSV:
#      python dxf_to_boq_csv.py --dxf "C:\path\in.dxf" --out "C:\path\out.csv"
# ────────────────────────────────────────────────────────────────────────────────

from __future__ import annotations

import argparse, csv, uuid, time, math, logging
from pathlib import Path
from typing import List, Iterable, Tuple, Optional

import ezdxf

# ===== EDIT THESE DEFAULT PATHS (can be overridden via CLI) =====
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"       # folder or a single .dxf file
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"   # output folder for CSVs
# ===============================================================

CSV_FIELDS = [
    "run_id", "source_file", "entity_type", "qty_type", "qty_value",
    "block_name", "layer", "handle", "remarks",
    "bbox_length", "bbox_width",  # NEW
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
    """
    Convert a value from meters (length) or m^2 (area) to target units.
    target in {"m","mm","cm","ft"}, kind in {"length","area"}.
    """
    t = (target or "m").lower().strip()
    if kind == "length":
        if t == "m":  return value_m
        if t == "mm": return value_m * 1000.0
        if t == "cm": return value_m * 100.0
        if t == "ft": return value_m / 0.3048
        logging.warning("Unknown target length unit '%s', defaulting to meters.", target)
        return value_m
    # area
    if t == "m":  return value_m
    if t == "mm": return value_m * 1_000_000.0
    if t == "cm": return value_m * 10_000.0
    if t == "ft": return value_m / (0.3048 ** 2)
    logging.warning("Unknown target area unit '%s', defaulting to m^2.", target)
    return value_m

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
    # step ~7.5° → 48 samples max
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
        # use polyline boundaries when present
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
    """
    Return (dx, dy) span in drawing units for the INSERT footprint in WCS.
    Uses virtual_entities() transformed by the insert to gather points.
    """
    try:
        pts_min_x = pts_min_y = float("inf")
        pts_max_x = pts_max_y = float("-inf")

        # virtual_entities() returns geometry already transformed by INSERT
        for ve in ins.virtual_entities():  # type: ignore[attr-defined]
            for (x, y) in _collect_points_from_entity(ve) or []:
                if x < pts_min_x: pts_min_x = x
                if y < pts_min_y: pts_min_y = y
                if x > pts_max_x: pts_max_x = x
                if y > pts_max_y: pts_max_y = y

        if pts_min_x == float("inf") or pts_min_y == float("inf"):
            return None  # nothing collected

        dx = max(0.0, pts_max_x - pts_min_x)
        dy = max(0.0, pts_max_y - pts_min_y)
        return (dx, dy)
    except Exception as ex:
        logging.debug("bbox for INSERT failed: %s", ex)
        return None

# ---------------------------
# CSV row builder
# ---------------------------
def make_row(run_id, source_file, entity_type, qty_type, qty_value,
             block_name="", layer="", handle="", remarks="",
             bbox_length=None, bbox_width=None) -> dict:
    row = {
        "run_id": run_id, "source_file": source_file,
        "entity_type": entity_type, "qty_type": qty_type,
        "qty_value": f"{qty_value:.6f}" if isinstance(qty_value, (int,float)) else (qty_value or ""),
        "block_name": block_name or "", "layer": layer or "",
        "handle": handle or "", "remarks": remarks or "",
        "bbox_length": "" if bbox_length is None else f"{bbox_length:.6f}",
        "bbox_width":  "" if bbox_width  is None else f"{bbox_width:.6f}",
    }
    return row

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
            if not include_xrefs and ("|" in (name or "")):  # skip xref|block unless asked
                continue

            # compute bbox in drawing units → meters → target length units
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

    # LINE
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

    # LWPOLYLINE
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

    # POLYLINE (2D/3D)
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

    # ARC
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

    # CIRCLE
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
    pts, had_bulge = [], False
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

    # Polyline boundary fallback
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
# IO & summary
# ---------------------------
def load_doc(path: Path):
    try:
        return ezdxf.readfile(str(path))
    except Exception as ex:
        logging.error("Failed to read DXF '%s': %s", path, ex)
        raise

def write_csv(rows: list[dict], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)

def print_summary(rows: list[dict], out_path: Path) -> None:
    total_inserts = sum(1 for r in rows if r["entity_type"] == "INSERT" and r["qty_type"] == "count")
    length_by_layer, area_by_layer = {}, {}
    for r in rows:
        qtype = r.get("qty_type")
        layer = r.get("layer", "")
        try:
            val = float(r.get("qty_value", "0") or 0)
        except Exception:
            val = 0.0
        if qtype == "length":
            length_by_layer[layer] = length_by_layer.get(layer, 0.0) + val
        elif qtype == "area":
            area_by_layer[layer] = area_by_layer.get(layer, 0.0) + val

    logging.info("----- SUMMARY for %s -----", out_path.name)
    logging.info("INSERTs counted: %d", total_inserts)
    if length_by_layer:
        logging.info("Total linear length per layer (in target units):")
        for ly, v in sorted(length_by_layer.items()):
            logging.info("  %-20s : %.6f", ly, v)
    else:
        logging.info("No linear length layers processed.")

    if area_by_layer:
        logging.info("Total area per layer (in target units^2):")
        for ly, v in sorted(area_by_layer.items()):
            logging.info("  %-20s : %.6f", ly, v)
    else:
        logging.info("No area layers processed.")
    logging.info("CSV written to: %s", out_path)

# ---------------------------
# Batch helpers
# ---------------------------
def collect_dxf_files(path: Path, recursive: bool) -> List[Path]:
    """Return a list of .dxf files from a single file or a directory."""
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
                    length_layers, area_layers, target_units, include_xrefs) -> None:
    logging.info("Processing DXF: %s", dxf_path)
    run_id = make_run_id()
    source_file = dxf_path.name

    doc = load_doc(dxf_path)
    msp = doc.modelspace()
    scale_to_m = units_scale_to_meters(doc)

    rows: list[dict] = []
    rows.extend(iter_block_rows(msp, run_id, source_file, include_xrefs=include_xrefs,
                                scale_to_m=scale_to_m, target_units=target_units))
    if length_layers:
        rows.extend(iter_length_rows(msp, run_id, source_file, length_layers, scale_to_m, target_units))
    if area_layers:
        rows.extend(iter_area_rows(msp, run_id, source_file, area_layers, scale_to_m, target_units))

    if not rows:
        logging.warning("No rows generated for %s. Check layer filters or DXF content.", dxf_path)

    out_path = derive_out_path(dxf_path, out_dir)
    write_csv(rows, out_path)
    print_summary(rows, out_path)

# ---------------------------
# Main
# ---------------------------
def main():
    ap = argparse.ArgumentParser(description="DXF → CSV raw quantities extractor (counts, lengths, areas) + INSERT bbox size.")
    ap.add_argument("--dxf",
                    help="Path to a DXF file OR a folder containing DXFs. "
                         "Default: value from DXF_FOLDER at top of script.")
    ap.add_argument("--name",
                    help="Base filename (no extension) to resolve as <DXF_FOLDER>/<NAME>.dxf.")
    ap.add_argument("--out-dir",
                    help="Directory to write CSVs (single or batch). "
                         "Default: value from OUT_ROOT at top of script.")
    ap.add_argument("--out",
                    help="(Single-file only) explicit CSV path. If provided, --out-dir is ignored.")
    ap.add_argument("--recursive", action="store_true",
                    help="When --dxf (or DXF_FOLDER) is a folder, search subfolders for *.dxf.")
    ap.add_argument("--length-layers", action="append", default=[],
                    help="Repeatable. Example: --length-layers WALL")
    ap.add_argument("--area-layers", action="append", default=[],
                    help="Repeatable. Example: --area-layers FLR")
    ap.add_argument("--target-units", default="m",
                    help="m, mm, cm, ft. Default: m.")
    ap.add_argument("--include-xrefs", action="store_true",
                    help="Include INSERTs with '|' in name. Default: skip them.")
    ap.add_argument("--verbose", action="store_true", help="Enable debug logging.")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    # Resolve input
    if args.dxf:
        dxf_input = Path(args.dxf)
    elif args.name:
        dxf_input = Path(DXF_FOLDER) / f"{args.name}.dxf"
    else:
        dxf_input = Path(DXF_FOLDER)

    out_dir   = Path(args.out_dir) if args.out_dir else Path(OUT_ROOT)
    explicit_out = Path(args.out) if args.out else None

    if not dxf_input.exists():
        logging.error("DXF input path not found: %s", dxf_input)
        return

    files = collect_dxf_files(dxf_input, recursive=args.recursive)
    if not files:
        return

    # Single-file explicit CSV
    if explicit_out:
        if len(files) != 1:
            logging.error("--out can only be used with a single DXF file. For folders, use --out-dir.")
            return
        f = files[0]
        run_id = make_run_id()
        source_file = f.name
        doc = load_doc(f)
        msp = doc.modelspace()
        scale_to_m = units_scale_to_meters(doc)
        rows: list[dict] = []
        rows.extend(iter_block_rows(msp, run_id, source_file, include_xrefs=args.include_xrefs,
                                    scale_to_m=scale_to_m, target_units=args.target_units))
        if args.length_layers:
            rows.extend(iter_length_rows(msp, run_id, source_file, args.length_layers, scale_to_m, args.target_units))
        if args.area_layers:
            rows.extend(iter_area_rows(msp, run_id, source_file, args.area_layers, scale_to_m, args.target_units))
        explicit_out.parent.mkdir(parents=True, exist_ok=True)
        write_csv(rows, explicit_out)
        print_summary(rows, explicit_out)
        return

    # Normal (single or batch)
    out_dir = out_dir if str(out_dir).strip() else None
    for f in files:
        try:
            process_one_dxf(
                dxf_path=f,
                out_dir=out_dir,
                length_layers=args.length_layers,
                area_layers=args.area_layers,
                target_units=args.target_units,
                include_xrefs=args.include_xrefs,
            )
        except Exception as ex:
            logging.exception("Failed processing %s: %s", f, ex)

if __name__ == "__main__":
    main()
