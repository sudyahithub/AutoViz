#!/usr/bin/env python3
# dxf_to_boq_csv.py — DXF → CSV + Google Sheets (Apps Script Web App) + Preview images
#
# - INSERT aggregation by (block_name, layer) using median bbox (length/width)
# - Per-layer metrics with OPEN length separated from CLOSED perimeter + area
# - Adds a Preview column: tiny PNGs rendered from block geometry (first instance)
# - Safe numeric formatting
# - Optional upload to Google Sheets via Apps Script Web App
# - Header order (as requested):
#   run_id, source_file, handle, entity_type, category,
#   BOQ name, qty_type, qty_value, length, width, perimeter, area, Preview, remarks
#
# Install deps:  pip install ezdxf requests matplotlib

from __future__ import annotations

import argparse, csv, uuid, time, math, logging, io, base64
from pathlib import Path
from typing import List, Tuple, Optional, Dict

import requests
import ezdxf

# Headless matplotlib for PNG previews
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ===== DEFAULT PATHS (CLI can override) =====
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"
# ===========================================

# ===== GOOGLE WEB APP DEFAULTS (can be overridden via CLI) =====
GS_WEBAPP_URL       = "https://script.google.com/macros/s/AKfycbz8imGgNtDdQks_v9xjJjZ2F4YgNkjUaawgHAYZRRceEjZKL3N11UsPb5_cbr7Q30iY/exec"
GSHEET_ID           = "1TbmU6vYevnhhYGFP-j91g752iFj_O0nkofAZn86akU0"
GSHEET_TAB          = "Sheet8"
GSHEET_SUMMARY_TAB  = ""          # optional; blank to skip
GSHEET_MODE         = "replace"   # "replace" or "append"
# Optional: Drive folder to store previews. Leave "" to auto-create "DXF-Previews".
GS_DRIVE_FOLDER_ID  = ""
# ===============================================================

CSV_HEADERS = [
    "run_id","source_file","handle","entity_type","category",
    "BOQ name","qty_type","qty_value","length","width","perimeter","area","Preview","remarks"
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
    try:
        code = int(doc.header.get("$INSUNITS", 0))
    except Exception:
        code = 0
    # 0: unitless (assume mm), 1: inches, 2: feet, 4:mm, 5:cm, 6:m
    mapping = {0:0.001, 1:0.0254, 2:0.3048, 4:0.001, 5:0.01, 6:1.0}
    scale = mapping.get(code, 1.0)
    if code not in mapping:
        logging.warning("Unrecognized $INSUNITS=%s; assuming meters.", code)
    else:
        logging.info("Detected $INSUNITS=%s → %s m/unit", code, scale)
    return scale

def to_target_units(v_m: float, target: str, kind: str) -> float:
    t = (target or "m").lower().strip()
    if kind == "length":
        return {"m":v_m, "mm":v_m*1000, "cm":v_m*100, "ft":v_m/0.3048}.get(t, v_m)
    # area
    return {"m":v_m, "mm":v_m*1_000_000, "cm":v_m*10_000, "ft":v_m/(0.3048**2)}.get(t, v_m)

def dist2d(p1, p2) -> float:
    return math.hypot(p2[0]-p1[0], p2[1]-p1[1])

def polyline_length_xy(pts: list[tuple[float,float]], closed: bool) -> float:
    if len(pts) < 2: return 0.0
    L = sum(dist2d(pts[i], pts[i+1]) for i in range(len(pts)-1))
    if closed: L += dist2d(pts[-1], pts[0])
    return L

def polygon_area_xy(pts: list[tuple[float,float]]) -> float:
    n = len(pts)
    if n < 3: return 0.0
    s = 0.0
    for i in range(n):
        x1,y1 = pts[i]; x2,y2 = pts[(i+1)%n]
        s += x1*y2 - x2*y1
    return abs(s)*0.5

def _sample_arc_pts(cx, cy, r, start_deg: Optional[float], end_deg: Optional[float]):
    if r <= 0: return []
    if start_deg is None or end_deg is None:
        start_deg, end_deg = 0.0, 360.0
    sweep = (end_deg - start_deg) % 360.0
    steps = max(16, int(max(16, sweep/6.0)))
    for i in range(steps+1):
        a = math.radians(start_deg + sweep*(i/steps))
        yield (cx + r*math.cos(a), cy + r*math.sin(a))

# Bulge-aware arc sampling for polylines
def _bulge_arc_points(p1: tuple[float,float], p2: tuple[float,float], bulge: float, min_steps: int = 8) -> list[tuple[float,float]]:
    if abs(bulge) < 1e-12:
        return [p1, p2]
    x1,y1 = p1; x2,y2 = p2
    dx, dy = x2 - x1, y2 - y1
    c = math.hypot(dx, dy)
    if c < 1e-12:
        return [p1]
    theta = 4.0 * math.atan(bulge)              # signed sweep
    s_half = 2*bulge / (1 + bulge*bulge)
    if abs(s_half) < 1e-12:
        return [p1, p2]
    R = c / (2.0 * s_half)
    nx, ny = (-dy / c, dx / c)
    cos_half = (1 - bulge*bulge) / (1 + bulge*bulge)
    d = R * cos_half
    mx, my = (x1 + x2) * 0.5, (y1 + y2) * 0.5
    cx, cy = mx + nx * d, my + ny * d

    a1 = math.atan2(y1 - cy, x1 - cx)
    a2 = math.atan2(y2 - cy, x2 - cx)
    raw_ccw = (a2 - a1) % (2*math.pi)
    sweep = raw_ccw if theta >= 0 else raw_ccw - 2*math.pi

    steps = max(min_steps, int(abs(sweep) / (6*math.pi/180)))  # ~6°
    pts = []
    for i in range(steps+1):
        t = i / steps
        ang = a1 + sweep * t
        pts.append((cx + R * math.cos(ang), cy + R * math.sin(ang)))
    return pts

def _collect_points_from_entity(e):
    et = e.dxftype()
    if et == "LINE":
        yield (float(e.dxf.start.x), float(e.dxf.start.y))
        yield (float(e.dxf.end.x), float(e.dxf.end.y))
    elif et == "LWPOLYLINE":
        verts = list(e); n = len(verts)
        if n == 0: return
        closed = bool(getattr(e, "closed", False))
        for i in range(n if closed else n-1):
            j = (i + 1) % n
            x1,y1 = float(verts[i][0]), float(verts[i][1])
            x2,y2 = float(verts[j][0]), float(verts[j][1])
            b = 0.0
            try: b = float(verts[i][4])
            except Exception: b = 0.0
            for p in _bulge_arc_points((x1,y1),(x2,y2),b)[:-1]:
                yield p
        yield (float(verts[-1][0]), float(verts[-1][1]))
        if closed:
            yield (float(verts[0][0]), float(verts[0][1]))
    elif et == "POLYLINE":
        vs = list(e.vertices())
        if not vs: return
        pts = []
        for v in vs:
            loc = getattr(v.dxf, "location", None)
            if loc is not None:
                pts.append((float(loc.x), float(loc.y)))
            else:
                pts.append((float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
        closed = bool(getattr(e,"is_closed",getattr(e,"closed",False)))
        n = len(pts)
        for i in range(n - (0 if closed else 1)):
            j = (i + 1) % n
            b = 0.0
            try: b = float(vs[i].dxf.bulge)
            except Exception: b = 0.0
            for p in _bulge_arc_points(pts[i], pts[j], b)[:-1]:
                yield p
        yield pts[-1]
        if closed:
            yield pts[0]
    elif et == "CIRCLE":
        cx, cy = float(e.dxf.center.x), float(e.dxf.center.y); r = float(e.dxf.radius)
        yield from _sample_arc_pts(cx, cy, r, None, None)
    elif et == "ARC":
        cx, cy = float(e.dxf.center.x), float(e.dxf.center.y); r = float(e.dxf.radius)
        sa, ea = float(e.dxf.start_angle), float(e.dxf.end_angle)
        yield from _sample_arc_pts(cx, cy, r, sa, ea)
    elif et == "HATCH":
        paths = getattr(e, "paths", None)
        if paths:
            for path in paths:
                verts = getattr(path, "polyline_path", None)
                if verts:
                    for v in verts:
                        x = float(getattr(v, "x", v[0])); y = float(getattr(v, "y", v[1]))
                        yield (x, y)

def _bbox_of_insert_xy(ins) -> Optional[Tuple[float,float]]:
    try:
        minx=miny=float("inf"); maxx=maxy=float("-inf")
        for ve in ins.virtual_entities():
            got_any = False
            for (x, y) in _collect_points_from_entity(ve) or []:
                got_any = True
                if x < minx: minx = x
                if y < miny: miny = y
                if x > maxx: maxx = x
                if y > maxy: maxy = y
            if not got_any:
                continue
        if minx == float("inf"):
            return None
        dx = max(0.0, maxx - minx)
        dy = max(0.0, maxy - miny)
        L = max(dx, dy)  # normalize
        W = min(dx, dy)
        return (L, W)
    except Exception as ex:
        logging.debug("bbox for INSERT failed: %s", ex)
        return None

def _fmt_num(val, places: int = 6) -> str:
    if val is None: return ""
    try:
        return f"{float(str(val).strip()):.{places}f}"
    except Exception:
        return ""

def make_row(run_id, source_file, entity_type, qty_type, qty_value,
             block_name="", layer="", handle="", remarks="",
             bbox_length=None, bbox_width=None, preview_b64:str="",
             perimeter=None, area=None) -> dict:
    return {
        "run_id": run_id, "source_file": source_file,
        "entity_type": entity_type, "qty_type": qty_type,
        "qty_value": _fmt_num(qty_value),
        "block_name": block_name or "", "layer": layer_or_misc(layer),
        "handle": handle or "", "remarks": remarks or "",
        "bbox_length": _fmt_num(bbox_length),
        "bbox_width":  _fmt_num(bbox_width),
        "preview_b64": preview_b64 or "",
        "perimeter": _fmt_num(perimeter),
        "area": _fmt_num(area),
    }

# ---------------------------
# Tiny renderer (from an INSERT's virtual geometry)
# ---------------------------
def _render_preview_from_insert(ins, size_px:int=192, pad_ratio:float=0.06) -> str:
    try:
        polylines: list[list[tuple[float,float]]] = []
        minx=miny=float("inf"); maxx=maxy=float("-inf")

        for ve in ins.virtual_entities():
            et = ve.dxftype()
            pts = []
            if et == "LINE":
                pts = [(float(ve.dxf.start.x), float(ve.dxf.start.y)),
                       (float(ve.dxf.end.x),   float(ve.dxf.end.y))]
            elif et == "LWPOLYLINE":
                verts = list(ve); n = len(verts)
                if n:
                    closed = bool(getattr(ve, "closed", False))
                    for i in range(n if closed else n-1):
                        j = (i + 1) % n
                        b = 0.0
                        try: b = float(verts[i][4])
                        except Exception: b = 0.0
                        seg = _bulge_arc_points((float(verts[i][0]), float(verts[i][1])),
                                                (float(verts[j][0]), float(verts[j][1])), b)
                        if pts and seg:
                            if pts[-1] == seg[0]: pts.extend(seg[1:])
                            else: pts.extend(seg)
                        else:
                            pts.extend(seg)
                    if closed and pts and pts[0] != pts[-1]:
                        pts.append(pts[0])
            elif et == "POLYLINE":
                vs = list(ve.vertices())
                if vs:
                    tmp = []
                    coords = []
                    for v in vs:
                        loc=getattr(v.dxf,"location",None)
                        if loc is not None: coords.append((float(loc.x),float(loc.y)))
                        else: coords.append((float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
                    closed = bool(getattr(ve,"is_closed",getattr(ve,"closed",False)))
                    n = len(coords)
                    for i in range(n - (0 if closed else 1)):
                        j = (i + 1) % n
                        b = 0.0
                        try: b = float(vs[i].dxf.bulge)
                        except Exception: b = 0.0
                        seg = _bulge_arc_points(coords[i], coords[j], b)
                        if tmp and seg:
                            if tmp[-1] == seg[0]: tmp.extend(seg[1:])
                            else: tmp.extend(seg)
                        else:
                            tmp.extend(seg)
                    if closed and tmp and tmp[0] != tmp[-1]:
                        tmp.append(tmp[0])
                    pts = tmp
            elif et == "CIRCLE":
                cx, cy = float(ve.dxf.center.x), float(ve.dxf.center.y); r = float(ve.dxf.radius)
                if r > 0:
                    pts = list(_sample_arc_pts(cx, cy, r, None, None))
                    if pts and pts[0] != pts[-1]:
                        pts.append(pts[0])
            elif et == "ARC":
                cx, cy = float(ve.dxf.center.x), float(ve.dxf.center.y); r = float(ve.dxf.radius)
                sa, ea = float(ve.dxf.start_angle), float(ve.dxf.end_angle)
                if r > 0:
                    pts = list(_sample_arc_pts(cx, cy, r, sa, ea))

            if len(pts) >= 2 and any((pts[i] != pts[i+1]) for i in range(len(pts)-1)):
                polylines.append(pts)
                for (x,y) in pts:
                    if x < minx: minx = x
                    if y < miny: miny = y
                    if x > maxx: maxx = x
                    if y > maxy: maxy = y

        if minx == float("inf") or not polylines:
            return ""

        w = max(maxx - minx, 1.0)
        h = max(maxy - miny, 1.0)
        size = max(w, h)
        pad = max(size * pad_ratio, 0.5)
        cx = (minx + maxx) * 0.5
        cy = (miny + maxy) * 0.5
        half = size * 0.5 + pad
        xmin, xmax = cx - half, cx + half
        ymin, ymax = cy - half, cy + half

        fig = plt.figure(figsize=(size_px/100, size_px/100), dpi=100)
        ax = fig.add_subplot(111)
        ax.axis("off"); ax.set_aspect("equal")

        for pts in polylines:
            xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
            ax.plot(xs, ys, linewidth=1)

        ax.set_xlim([xmin, xmax]); ax.set_ylim([ymin, ymax])

        buf = io.BytesIO()
        plt.subplots_adjust(0,0,1,1)
        fig.savefig(buf, format="png", transparent=True, bbox_inches="tight", pad_inches=0)
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception as ex:
        logging.debug("preview render failed: %s", ex)
        return ""

def _build_preview_cache(msp) -> Dict[str, str]:
    cache: Dict[str, str] = {}
    for ins in msp.query("INSERT"):
        try:
            name = getattr(ins, "effective_name", None) or getattr(ins, "block_name", None) or getattr(ins.dxf, "name", "")
            if not name or name in cache:
                continue
            b64 = _render_preview_from_insert(ins)
            cache[name] = b64 or ""
        except Exception:
            pass
    return cache

# ---------------------------
# Detailed INSERT rows
# ---------------------------
def iter_block_rows(msp, run_id, source_file, include_xrefs: bool,
                    scale_to_m: float, target_units: str,
                    preview_cache: Dict[str,str] | None = None) -> list[dict]:
    out = []
    preview_cache = preview_cache or {}
    for ins in msp.query("INSERT"):
        try:
            name = getattr(ins, "effective_name", None) or getattr(ins, "block_name", None) or getattr(ins.dxf, "name", "")
            if not include_xrefs and ("|" in (name or "")):
                continue
            bbox_du = _bbox_of_insert_xy(ins)
            if bbox_du:
                L_m = bbox_du[0] * scale_to_m
                W_m = bbox_du[1] * scale_to_m
                L_out = to_target_units(L_m, target_units, "length")
                W_out = to_target_units(W_m, target_units, "length")
            else:
                L_out = W_out = None
            out.append(make_row(
                run_id, source_file, "INSERT", "count", 1.0,
                block_name=name,
                layer=getattr(ins.dxf, "layer", ""),
                handle=getattr(ins.dxf, "handle", ""),
                remarks="",
                bbox_length=L_out, bbox_width=W_out,
                preview_b64=preview_cache.get(name, "")
            ))
        except Exception as ex:
            logging.exception("INSERT failed: %s", ex)
    return out

# ---------------------------
# Per-layer metrics (SPLIT OPEN vs CLOSED)
# ---------------------------
def compute_layer_metrics(msp, scale_to_m: float, target_units: str):
    """
    Returns three dicts keyed by normalized layer:
      open_len_by_layer → LINE, ARC, open POLYLINE/LWPOLYLINE
      peri_by_layer     → perimeter of closed POLYLINE/LWPOLYLINE & CIRCLE
      area_by_layer     → area of closed POLYLINE/LWPOLYLINE, HATCH, CIRCLE
    Values are already converted to target units.
    """
    open_len_by_layer: Dict[str, float] = {}
    peri_by_layer:     Dict[str, float] = {}
    area_by_layer:     Dict[str, float] = {}

    def add_open_len(layer, L_du):
        if L_du <= 0: return
        L_out = to_target_units(L_du * scale_to_m, target_units, "length")
        k = layer_or_misc(layer)
        open_len_by_layer[k] = open_len_by_layer.get(k, 0.0) + L_out

    def add_perimeter(layer, P_du):
        if P_du <= 0: return
        P_out = to_target_units(P_du * scale_to_m, target_units, "length")
        k = layer_or_misc(layer)
        peri_by_layer[k] = peri_by_layer.get(k, 0.0) + P_out

    def add_area(layer, A_du):
        if A_du <= 0: return
        A_out = to_target_units(A_du * (scale_to_m**2), target_units, "area")
        k = layer_or_misc(layer)
        area_by_layer[k] = area_by_layer.get(k, 0.0) + A_out

    # LINE
    for e in msp.query("LINE"):
        try:
            p1=(e.dxf.start.x,e.dxf.start.y); p2=(e.dxf.end.x,e.dxf.end.y)
            add_open_len(e.dxf.layer, dist2d(p1,p2))
        except Exception: pass

    # LWPOLYLINE
    for e in msp.query("LWPOLYLINE"):
        try:
            verts = list(e)
            if not verts: 
                continue
            closed = bool(getattr(e,"closed",False))
            dense: list[tuple[float,float]] = []
            n = len(verts)
            for i in range(n if closed else n-1):
                j = (i + 1) % n
                b = 0.0
                try: b = float(verts[i][4])
                except Exception: b = 0.0
                seg = _bulge_arc_points((float(verts[i][0]), float(verts[i][1])),
                                        (float(verts[j][0]), float(verts[j][1])), b)
                dense.extend(seg[:-1])
            dense.append((float(verts[-1][0]), float(verts[-1][1])))
            if closed: dense.append((float(verts[0][0]), float(verts[0][1])))

            L = polyline_length_xy(dense, closed=False)
            if closed:
                add_perimeter(e.dxf.layer, L)
                if len(dense) >= 3:
                    add_area(e.dxf.layer, polygon_area_xy(dense[:-1]))
            else:
                add_open_len(e.dxf.layer, L)
        except Exception: pass

    # 2D/3D POLYLINE
    for e in msp.query("POLYLINE"):
        try:
            vs = list(e.vertices())
            if not vs: 
                continue
            coords = []
            for v in vs:
                loc=getattr(v.dxf,"location",None)
                if loc is not None: coords.append((float(loc.x),float(loc.y)))
                else: coords.append((float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
            closed = bool(getattr(e,"is_closed",getattr(e,"closed",False)))
            n = len(coords)
            dense=[]
            for i in range(n - (0 if closed else 1)):
                j = (i + 1) % n
                b = 0.0
                try: b = float(vs[i].dxf.bulge)
                except Exception: b = 0.0
                seg = _bulge_arc_points(coords[i], coords[j], b)
                dense.extend(seg[:-1])
            dense.append(coords[-1])
            if closed: dense.append(coords[0])

            L = polyline_length_xy(dense, closed=False)
            if closed:
                add_perimeter(e.dxf.layer, L)
                if len(dense) >= 3:
                    add_area(e.dxf.layer, polygon_area_xy(dense[:-1]))
            else:
                add_open_len(e.dxf.layer, L)
        except Exception: pass

    # ARC (open)
    for e in msp.query("ARC"):
        try:
            r=float(e.dxf.radius)
            sweep=(float(e.dxf.end_angle)-float(e.dxf.start_angle))%360.0
            add_open_len(e.dxf.layer, (2.0*math.pi*r)*(sweep/360.0))
        except Exception: pass

    # CIRCLE (closed)
    for e in msp.query("CIRCLE"):
        try:
            r=float(e.dxf.radius)
            add_perimeter(e.dxf.layer, (2.0*math.pi*r))
            add_area(e.dxf.layer, math.pi*(r**2))
        except Exception: pass

    # HATCH (area only, best-effort)
    for e in msp.query("HATCH"):
        try:
            A_du=None
            if hasattr(e,"get_filled_area"):
                try:
                    v=e.get_filled_area()
                    if v and v>0: A_du=float(v)
                except Exception:
                    A_du=None
            if A_du is None:
                total=0.0; used=False
                paths=getattr(e,"paths",None)
                if paths:
                    for path in paths:
                        verts=getattr(path,"polyline_path",None)
                        if verts:
                            pts=[(float(getattr(v,"x",v[0])), float(getattr(v,"y",v[1]))) for v in verts]
                            if len(pts)>=3:
                                total += polygon_area_xy(pts); used=True
                if used: A_du=total
            if A_du and A_du>0: add_area(e.dxf.layer, A_du)
        except Exception: pass

    return open_len_by_layer, peri_by_layer, area_by_layer

# --- New: recover rectangle dimensions from perimeter & area (if rectangular) ---
def solve_rect_dims_from_perimeter_area(P: float, A: float) -> Tuple[Optional[float], Optional[float]]:
    """Given perimeter P and area A, return (L, W) with L>=W if real, else (None,None).
    Works in whatever units P and A are already in (must be consistent)."""
    try:
        if P is None or A is None or P <= 0 or A <= 0:
            return (None, None)
        S = P / 2.0                  # a + b
        D = S*S - 4.0*A              # (a-b)^2
        if D < -1e-9:
            return (None, None)
        D = max(D, 0.0)
        root = math.sqrt(D)
        a = 0.5 * (S + root)
        b = 0.5 * (S - root)
        if a <= 0 or b <= 0:
            return (None, None)
        L, W = (a, b) if a >= b else (b, a)
        return (L, W)
    except Exception:
        return (None, None)

def make_layer_total_rows(open_by, peri_by, area_by, run_id, source_file,
                          mode: str = "split"):
    """
    mode:
      'split'    → up to two rows per layer:
                    - OPEN length only
                    - CLOSED rectangle dims + (perimeter & area in new columns)
      'combined' → one row per layer (open+closed length; area from closed)
    """
    rows = []
    all_layers = sorted(set(open_by.keys()) | set(peri_by.keys()) | set(area_by.keys()))

    if mode == "combined":
        for ly in all_layers:
            L_tot = open_by.get(ly, 0.0) + peri_by.get(ly, 0.0)
            A_tot = area_by.get(ly, 0.0)
            rows.append(make_row(
                run_id, source_file, "LAYER_SUMMARY", "layer", None,
                layer=ly, remarks="totals per layer (open+closed length; area from closed)",
                bbox_length=(L_tot if L_tot>0 else None),
                bbox_width=(A_tot if A_tot>0 else None),
            ))
        return rows

    # split rows
    for ly in all_layers:
        if open_by.get(ly, 0.0) > 0:
            rows.append(make_row(
                run_id, source_file, "LAYER_SUMMARY", "layer", None,
                layer=ly, remarks="OPEN length only",
                bbox_length=open_by[ly], bbox_width=None,
            ))
        if peri_by.get(ly, 0.0) > 0 or area_by.get(ly, 0.0) > 0:
            P = peri_by.get(ly, None)
            A = area_by.get(ly, None)
            # Derive L & W from P and A (only meaningful for rectangles)
            L_rec, W_rec = solve_rect_dims_from_perimeter_area(P, A)
            rows.append(make_row(
                run_id, source_file, "LAYER_SUMMARY", "layer", None,
                layer=ly, remarks="CLOSED (rectangle): length/width + perimeter & area",
                bbox_length=L_rec, bbox_width=W_rec,
                perimeter=P, area=A
            ))
    return rows

# ---------------------------
# CSV / Web App I/O
# ---------------------------
def _map_to_csv_headers(row: dict) -> dict:
    return {
        "run_id":     row.get("run_id",""),
        "source_file":row.get("source_file",""),
        "handle":     row.get("handle",""),
        "entity_type":row.get("entity_type",""),
        "category":   row.get("layer",""),
        "BOQ name":   row.get("block_name",""),
        "qty_type":   row.get("qty_type",""),
        "qty_value":  row.get("qty_value",""),
        "length":     row.get("bbox_length",""),
        "width":      row.get("bbox_width",""),
        "perimeter":  row.get("perimeter",""),
        "area":       row.get("area",""),
        "Preview":    "",  # filled by Apps Script using uploaded image
        "remarks":    row.get("remarks",""),
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
        row.get("handle",""),
        row.get("entity_type",""),
        row.get("layer",""),
        row.get("block_name",""),
        row.get("qty_type",""),
        row.get("qty_value",""),
        row.get("bbox_length",""),
        row.get("bbox_width",""),
        row.get("perimeter",""),
        row.get("area",""),
        "",  # Preview placeholder
        row.get("remarks",""),
    ]

def _chunk(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i+n]

def push_rows_to_webapp(rows: list[dict], webapp_url: str, spreadsheet_id: str,
                        tab: str, mode: str = "replace", summary_tab: str = "",
                        batch_rows: int = 25, timeout: int = 300,
                        valign_middle: bool = False, sparse_anchor: str = "last",
                        drive_folder_id: str = "") -> None:
    """POST rows + base64 previews to the Apps Script Web App (batched and retried)."""
    if not webapp_url or not spreadsheet_id or not tab:
        logging.info("WebApp push not configured (missing url/id/tab). Skipping upload.")
        return

    headers = CSV_HEADERS[:]
    data_rows = [[
        r.get("run_id",""), r.get("source_file",""), r.get("handle",""),
        r.get("entity_type",""), r.get("layer",""), r.get("block_name",""),
        r.get("qty_type",""), r.get("qty_value",""), r.get("bbox_length",""),
        r.get("bbox_width",""), r.get("perimeter",""), r.get("area",""),
        "", r.get("remarks",""),
    ] for r in rows]
    images = [r.get("preview_b64","") for r in rows]

    first_mode = (mode or "replace").lower()
    sess = requests.Session()

    def post_with_retries(payload, tries=4, backoff=2.0):
        for attempt in range(1, tries+1):
            try:
                return sess.post(webapp_url, json=payload, timeout=timeout, allow_redirects=True)
            except requests.exceptions.ReadTimeout:
                if attempt == tries:
                    raise
                wait = backoff ** attempt
                logging.warning("Upload timeout; retrying in %.1fs (attempt %d/%d)...", wait, attempt+1, tries)
                time.sleep(wait)

    total = len(data_rows)
    sent = 0
    for idx, i in enumerate(range(0, total, batch_rows), start=1):
        chunk_rows = data_rows[i:i+batch_rows]
        chunk_imgs = images[i:i+batch_rows]

        payload = {
            "sheetId": spreadsheet_id,
            "tab": tab,
            "mode": "replace" if (i == 0 and first_mode == "replace") else "append",
            "headers": headers if (i == 0 and first_mode == "replace") else [],
            "rows": chunk_rows,
            "images": chunk_imgs,
            "driveFolderId": (drive_folder_id or ""),
            "vAlign": "middle" if valign_middle else "",
            "sparseAnchor": (sparse_anchor or "last"),
            "runId": rows[0].get("run_id","") if rows else ""
        }
        if summary_tab and i == 0:
            payload["summaryTab"] = summary_tab
            payload["summaryRows"] = []

        r = post_with_retries(payload)
        if not r.ok:
            raise RuntimeError(f"WebApp upload failed (batch {idx}): HTTP {r.status_code} {r.text}")
        sent += len(chunk_rows)
        logging.info("WebApp batch %d: uploaded %d/%d rows", idx, sent, total)

# ---------------------------
# Reporting
# ---------------------------
def print_summary(rows: list[dict], out_path: Path) -> None:
    total_insert_groups = sum(1 for r in rows if r["entity_type"]=="INSERT" and r["qty_type"]=="count")
    logging.info("----- SUMMARY for %s -----", out_path.name)
    logging.info("INSERT groups (after aggregation): %d", total_insert_groups)
    logging.info("CSV written to: %s", out_path)

# ---------------------------
# Sorting (to club identical categories together)
# ---------------------------
def _norm_cat(s: str) -> str:
    s = (s or "").strip()
    s = " ".join(s.split())
    return s.upper()

def sort_rows_for_category_blocks(rows: list[dict]) -> None:
    def _key(r):
        cat = _norm_cat(r.get("layer", ""))
        et  = r.get("entity_type", "")
        et_rank = 0 if et == "INSERT" else 1
        return (cat, et_rank, r.get("block_name", ""))
    rows.sort(key=_key)

# ---------------------------
# Batch helpers
# ---------------------------
def collect_dxf_files(path: Path, recursive: bool) -> List[Path]:
    if path.is_file():
        if path.suffix.lower() == ".dxf":
            return [path]
        logging.error("Provided file is not a .dxf: %s", path)
        return []
    if path is None or not path.exists():
        logging.error("Path does not exist: %s", path)
        return []
    if path.is_dir():
        pattern = "**/*.dxf" if recursive else "*.dxf"
        files = sorted(path.glob(pattern))
        if not files:
            logging.warning("No DXF files found in %s (recursive=%s)", path, recursive)
        return files

def derive_out_path(dxf_path: Path, out_dir: Path | None) -> Path:
    return (out_dir / f"{dxf_path.stem}_raw_extract.csv") if out_dir else dxf_path.with_name(f"{dxf_path.stem}_raw_extract.csv")

# ---------------------------
# Main processing
# ---------------------------
def process_one_dxf(dxf_path: Path, out_dir: Path | None,
                    target_units: str, include_xrefs: bool,
                    layer_metrics: bool, aggregate_inserts: bool,
                    layer_metrics_mode: str) -> list[dict]:
    logging.info("Processing DXF: %s", dxf_path)
    run_id = make_run_id()
    source_file = dxf_path.name

    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    scale_to_m = units_scale_to_meters(doc)

    # Build previews once per distinct block
    preview_cache = _build_preview_cache(msp)

    rows: list[dict] = []

    # INSERTs (then aggregate)
    insert_rows = iter_block_rows(msp, run_id, source_file, include_xrefs, scale_to_m, target_units, preview_cache)
    if aggregate_inserts:
        groups: Dict[tuple[str,str], dict] = {}
        for r in insert_rows:
            key = (r["block_name"], r["layer"])
            g = groups.setdefault(key, {"count":0,"xs":[],"ys":[], "preview": r.get("preview_b64","")})
            g["count"] += 1
            try:
                if r["bbox_length"] and r["bbox_width"]:
                    g["xs"].append(float(r["bbox_length"]))
                    g["ys"].append(float(r["bbox_width"]))
            except Exception:
                pass
        for (name, layer), g in groups.items():
            xs = sorted(g["xs"]); ys = sorted(g["ys"])
            bx = xs[len(xs)//2] if xs else None
            by = ys[len(ys)//2] if ys else None
            rows.append(make_row(
                run_id, source_file, "INSERT", "count", float(g["count"]),
                block_name=name, layer=layer, handle="",
                remarks=f"aggregated {g['count']} inserts", bbox_length=bx, bbox_width=by,
                preview_b64=g.get("preview","")
            ))
    else:
        rows.extend(insert_rows)

    # Per-layer metrics
    if layer_metrics:
        open_by, peri_by, area_by = compute_layer_metrics(msp, scale_to_m, target_units)
        rows.extend(make_layer_total_rows(open_by, peri_by, area_by, run_id, source_file,
                                          mode=layer_metrics_mode))

    # Sort so identical categories are contiguous (for merging later)
    sort_rows_for_category_blocks(rows)

    out_path = derive_out_path(dxf_path, out_dir)
    write_csv(rows, out_path)
    print_summary(rows, out_path)
    return rows

# ---------------------------
# CLI
# ---------------------------
def main():
    ap = argparse.ArgumentParser(description="DXF → CSV (INSERT bbox + per-layer totals) + Google Sheets Web App upload (+ previews).")
    ap.add_argument("--dxf", help="Path to a DXF file OR a folder containing DXFs. Default: DXF_FOLDER.")
    ap.add_argument("--name", help="Base filename (no extension) as <DXF_FOLDER>/<NAME>.dxf.")
    ap.add_argument("--out-dir", help="Directory to write CSVs (single or batch). Default: OUT_ROOT.")
    ap.add_argument("--out", help="(Single-file only) explicit CSV path.")
    ap.add_argument("--recursive", action="store_true", help="If input is a folder, search subfolders for *.dxf.")
    ap.add_argument("--target-units", default="m", help="m, mm, cm, ft. Default: m.")
    ap.add_argument("--include-xrefs", action="store_true", help="Include INSERTs with '|' in name.")
    ap.add_argument("--no-layer-metrics", action="store_true", help="Disable layer length/area summary rows.")
    ap.add_argument("--no-aggregate-inserts", action="store_true", help="Write one row per INSERT instead of grouping.")
    # Layer metrics reporting mode
    ap.add_argument("--layer-metrics-mode", choices=["combined","split"], default="split",
                    help="How to report layer totals. 'split' (default) outputs OPEN and CLOSED rows.")
    # Web App options
    ap.add_argument("--gs-webapp", default=None, help="Apps Script Web App URL (default: GS_WEBAPP_URL).")
    ap.add_argument("--gsheet-id", default=None, help="Spreadsheet ID (default: GSHEET_ID).")
    ap.add_argument("--gsheet-tab", default=None, help="Worksheet/tab for detail rows (default: GSHEET_TAB).")
    ap.add_argument("--gsheet-summary-tab", default=None, help="Optional summary tab (default: GSHEET_SUMMARY_TAB).")
    ap.add_argument("--gsheet-mode", choices=["replace","append"], default=None, help="Write mode (default: GSHEET_MODE).")
    ap.add_argument("--batch-rows", type=int, default=3000, help="Rows per WebApp request (default 3000).")
    # Display options for Google Sheet
    ap.add_argument("--align-middle", action="store_true",
                    help="Sheets upload only: ask Web App to vertically center all cells.")
    ap.add_argument("--sparse-anchor", choices=["first","last","middle"], default="last",
                    help="Where to show the 'category' label within each contiguous block (default: last).")
    # Drive folder for previews
    ap.add_argument("--drive-folder-id", default=None, help="Google Drive folder ID to store preview PNGs. If empty, Web App creates 'DXF-Previews'.")
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

    # Web App settings
    gs_webapp = (args.gs_webapp if args.gs_webapp is not None else GS_WEBAPP_URL).strip()
    gsheet_id = (args.gsheet_id if args.gsheet_id is not None else GSHEET_ID).strip()
    gsheet_tab = (args.gsheet_tab if args.gsheet_tab is not None else GSHEET_TAB).strip()
    gsheet_summary_tab = (args.gsheet_summary_tab if args.gsheet_summary_tab is not None else GSHEET_SUMMARY_TAB).strip()
    gsheet_mode = (args.gsheet_mode if args.gsheet_mode is not None else GSHEET_MODE).strip().lower()
    batch_rows = args.batch_rows
    align_middle = args.align_middle
    sparse_anchor = args.sparse_anchor
    drive_folder_id = (args.drive_folder_id if args.drive_folder_id is not None else GS_DRIVE_FOLDER_ID).strip()

    if not dxf_input.exists():
        logging.error("DXF input path not found: %s", dxf_input)
        return

    files = collect_dxf_files(dxf_input, recursive=args.recursive)
    if not files:
        return

    # Single-file explicit output
    if explicit_out:
        if len(files) != 1:
            logging.error("--out is for a single file. For folders, use --out-dir.")
            return
        f = files[0]
        rows = process_one_dxf(
            dxf_path=f, out_dir=explicit_out.parent,
            target_units=args.target_units, include_xrefs=args.include_xrefs,
            layer_metrics=layer_metrics, aggregate_inserts=aggregate_inserts,
            layer_metrics_mode=args.layer_metrics_mode
        )
        explicit_out.parent.mkdir(parents=True, exist_ok=True)
        write_csv(rows, explicit_out)
        print_summary(rows, explicit_out)
        if gs_webapp and gsheet_id:
            push_rows_to_webapp(rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, gsheet_summary_tab,
                                batch_rows=batch_rows, valign_middle=align_middle,
                                sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)
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
                layer_metrics_mode=args.layer_metrics_mode
            )
            all_rows.extend(rows or [])
        except Exception as ex:
            logging.exception("Failed processing %s: %s", f, ex)

    if all_rows and gs_webapp and gsheet_id:
        push_rows_to_webapp(all_rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, gsheet_summary_tab,
                            batch_rows=batch_rows, valign_middle=align_middle,
                            sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)

if __name__ == "__main__":
    main()
