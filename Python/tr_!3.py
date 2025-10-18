####2####


#!/usr/bin/env python3
# DXF → CSV + Google Sheets (Apps Script Web App)
# - Aggregates INSERTs by (block_name, PLANNER, zone) using median bbox (L/W)
# - Adds category1 = original DWG layer (kept in CSV & Detail; NOT sent to ByLayer)
# - Zone detection from PLANNER (INSERT or closed LWPOLYLINE + best label)
# - Layer totals with dominant color vote → ByLayer tab
# - YOUR ASK:
#   * Detail sheet removes perimeter & area
#   * ByLayer sheet removes zone, category1, BOQ name, qty_value

from __future__ import annotations
import argparse, csv, uuid, time, math, logging, io, base64
from pathlib import Path
from typing import List, Tuple, Optional, Dict
from dataclasses import dataclass

import requests
import ezdxf
from ezdxf import colors as ezcolors

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ===== Defaults you can edit =====
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"

GS_WEBAPP_URL       = "https://script.google.com/macros/s/AKfycbzdajkMohJJnWwbCKLQIp6imQe8VYCkkLQD4fB1sa0_2MfN7yhPONo8j3IacxIWna8u/exec"
GSHEET_ID           = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM"
GSHEET_TAB          = "sdasdssds"
GSHEET_SUMMARY_TAB  = ""       # blank → auto "<GSHEET_TAB>_ByLayer"
GSHEET_MODE         = "replace"
GS_DRIVE_FOLDER_ID  = "" 

# ========== Headers ==========
# Master CSV headers (for on-disk CSV/debug; keeps everything)
CSV_HEADERS = [
    "entity_type","category","zone","category1",
    "BOQ name","qty_type","qty_value","length (ft)","width (ft)","perimeter","area (ft2)","Preview","remarks"
]

# Detail (blocks) sheet → perimeter & area REMOVED
DETAIL_HEADERS = [
    "entity_type","category","zone","category1",
    "BOQ name","qty_type","qty_value","length (ft)","width (ft)","Preview","remarks"
]

# ByLayer sheet → zone, category1, BOQ name, qty_value REMOVED
LAYER_HEADERS = [
    "entity_type","category","qty_type","length (ft)","width (ft)","perimeter","area (ft2)","Preview","remarks"
]

# ===== Formatting & switches =====
DEC_PLACES = 2
FORCE_PLANNER_CATEGORY = True

# ===== Utilities =====
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

def _bulge_arc_points(p1, p2, bulge: float, min_steps: int = 8):
    if abs(bulge) < 1e-12: return [p1, p2]
    x1,y1 = p1; x2,y2 = p2
    dx, dy = x2-x1, y2-y1
    c = math.hypot(dx, dy)
    if c < 1e-12: return [p1]
    theta = 4.0 * math.atan(bulge)
    s_half = 2*bulge / (1 + bulge*bulge)
    if abs(s_half) < 1e-12: return [p1, p2]
    R = c / (2.0 * s_half)
    nx, ny = (-dy/c, dx/c)
    cos_half = (1 - bulge*bulge) / (1 + bulge*bulge)
    d = R * cos_half
    mx, my = (x1+x2)/2.0, (y1+y2)/2.0
    cx, cy = mx + nx*d, my + ny*d
    a1 = math.atan2(y1 - cy, x1 - cx)
    a2 = math.atan2(y2 - cy, x2 - cx)
    raw_ccw = (a2 - a1) % (2*math.pi)
    sweep = raw_ccw if theta >= 0 else raw_ccw - 2*math.pi
    steps = max(min_steps, int(abs(sweep) / (6*math.pi/180)))
    return [(cx + R*math.cos(a1 + sweep*(i/steps)),
             cy + R*math.sin(a1 + sweep*(i/steps))) for i in range(steps+1)]

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
            pts.append((float(loc.x), float(loc.y)) if loc is not None
                       else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
        closed = bool(getattr(e,"is_closed",getattr(e,"closed",False)))
        n = len(pts)
        for i in range(n - (0 if closed else 1)):
            j = (i + 1) % n
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
                minx = min(minx, x); miny = min(miny, y)
                maxx = max(maxx, x); maxy = max(maxy, y)
            if not got_any:
                continue
        if minx == float("inf"):
            return None
        dx = max(0.0, maxx - minx)
        dy = max(0.0, maxy - miny)
        L = max(dx, dy)
        W = min(dx, dy)
        return (L, W)
    except Exception:
        return None

def _fmt_num(val, places: int | None = None) -> str:
    if val is None or val == "": return ""
    try:
        p = DEC_PLACES if places is None else places
        num = float(str(val).strip())
        return f"{num:.{p}f}"
    except Exception:
        return ""

def _rgb_to_hex(rgb: tuple[int,int,int]) -> str:
    r, g, b = rgb
    return "#{:02X}{:02X}{:02X}".format(r, g, b)

# ===== ZONES =====
def _insert_bbox(ins) -> Optional[tuple[float,float,float,float]]:
    try:
        minx=miny=float("inf"); maxx=maxy=float("-inf"); anyp=False
        for ve in ins.virtual_entities():
            for (x,y) in _collect_points_from_entity(ve) or []:
                anyp=True
                minx=min(minx,x); miny=min(miny,y)
                maxx=max(maxx,x); maxy=max(maxy,y)
        if not anyp: return None
        return (minx,miny,maxx,maxy)
    except Exception:
        return None

def _bbox_center(b: tuple[float,float,float,float]) -> tuple[float,float]:
    minx, miny, maxx, maxy = b
    return ((minx+maxx)*0.5, (miny+maxy)*0.5)

def point_in_polygon(pt: tuple[float,float], poly: list[tuple[float,float]]) -> bool:
    x, y = pt; inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i+1) % n]
        if ((y1 > y) != (y2 > y)) and (x < (x2 - x1) * (y - y1) / (y2 - y1 + 1e-12) + x1):
            inside = not inside
    return inside

def _poly_from_lwpoly(e) -> list[tuple[float,float]]:
    verts = list(e)
    if not verts: return []
    pts = []
    n = len(verts)
    closed = bool(getattr(e, "closed", False))
    for i in range(n if closed else n-1):
        j = (i + 1) % n
        x1,y1 = float(verts[i][0]), float(verts[i][1])
        x2,y2 = float(verts[j][0]), float(verts[j][1])
        try: b = float(verts[i][4])
        except Exception: b = 0.0
        seg = _bulge_arc_points((x1,y1),(x2,y2),b)
        if pts and seg:
            pts.extend(seg[1:])
        else:
            pts.extend(seg)
    if closed and pts and pts[0] != pts[-1]:
        pts.append(pts[0])
    return pts

@dataclass
class Zone:
    name: str
    poly: list  # list[(x,y)]

def _collect_planner_zones(msp) -> list[Zone]:
    zones: list[Zone] = []
    # A) PLANNER INSERTs: bbox polygon & name from ATTRIB or block name
    for ins in msp.query('INSERT[layer=="PLANNER"]'):
        try:
            b = _insert_bbox(ins)
            if not b: continue
            minx, miny, maxx, maxy = b
            poly = [(minx, miny), (maxx, miny), (maxx, maxy), (minx, maxy), (minx, miny)]
            zname = None
            try:
                cand_tags = {"NAME","ROOM","ZONE","LABEL","TITLE"}
                for att in getattr(ins, "attribs", lambda: [])() or []:
                    tag = (getattr(att.dxf, "tag", "") or "").upper()
                    if tag in cand_tags:
                        txt = (getattr(att.dxf, "text", "") or "").strip()
                        if txt: zname = txt; break
            except Exception:
                pass
            if not zname:
                zname = (getattr(ins, "effective_name", None) or getattr(ins, "block_name", None) or getattr(ins.dxf, "name", "")).strip()
            if not zname: zname = "Zone"
            zones.append(Zone(name=zname, poly=poly))
        except Exception:
            pass

    if zones:
        seen = set(); out=[]
        for z in zones:
            key=(z.name, tuple(z.poly))
            if key not in seen:
                out.append(z); seen.add(key)
        return out

    # B) Closed PLANNER LWPOLYLINEs + nearest label
    tmp: list[Zone] = []
    for e in msp.query('LWPOLYLINE[layer=="PLANNER"]'):
        try:
            if not bool(getattr(e, "closed", False)): continue
            poly = _poly_from_lwpoly(e)
            if len(poly) >= 3: tmp.append(Zone(name="", poly=poly))
        except Exception: pass
    if not tmp: return []

    labels: list[tuple[str,tuple[float,float]]] = []
    for t in msp.query('TEXT'):
        try:
            labels.append(((t.dxf.text or "").strip(), (float(t.dxf.insert.x), float(t.dxf.insert.y))))
        except Exception: pass
    for mt in msp.query('MTEXT'):
        try:
            raw=(mt.text or "").strip()
            labels.append((raw.split("\n",1)[0].strip(), (float(mt.dxf.insert.x), float(mt.dxf.insert.y))))
        except Exception: pass

    def _centroid(poly):
        xs=[p[0] for p in poly]; ys=[p[1] for p in poly]
        return ((sum(xs)/len(xs)) if xs else 0.0, (sum(ys)/len(ys)) if ys else 0.0)

    used=set(); zones_out=[]
    for i,z in enumerate(tmp, start=1):
        zname=None
        for idx,(txt,pt) in enumerate(labels):
            if idx in used or not txt: continue
            if point_in_polygon(pt,z.poly):
                zname=txt; used.add(idx); break
        if not zname and labels:
            cx,cy=_centroid(z.poly); best=None
            for idx,(txt,(x,y)) in enumerate(labels):
                if idx in used or not txt: continue
                d=(x-cx)*(x-cx)+(y-cy)*(y-cy)
                if (best is None) or (d<best[0]): best=(d,idx,txt)
            if best: _,idx,txt=best; zname=txt; used.add(idx)
        if not zname: zname=f"Zone {i:02d}"
        zones_out.append(Zone(name=zname, poly=z.poly))
    return zones_out

def _zone_for_point(pt: tuple[float,float], zones: list[Zone]) -> Optional[str]:
    for z in zones:
        if point_in_polygon(pt, z.poly):
            return z.name
    return None

# ===== Row builder =====
def make_row(entity_type, qty_type, qty_value,
             block_name="", layer="", handle="", remarks="",
             bbox_length=None, bbox_width=None,
             preview_b64:str="", preview_hex:str="",
             perimeter=None, area=None, zone:str="", category1:str="") -> dict:
    return {
        "entity_type": entity_type, "qty_type": qty_type,
        "qty_value": _fmt_num(qty_value),
        "block_name": block_name or "",
        "layer": layer_or_misc(layer),
        "zone": (zone or ""),
        "category1": category1 or "",
        "handle": handle or "", "remarks": remarks or "",
        "bbox_length": _fmt_num(bbox_length),
        "bbox_width":  _fmt_num(bbox_width),
        "preview_b64": preview_b64 or "",
        "preview_hex": preview_hex or "",
        "perimeter": _fmt_num(perimeter),
        "area": _fmt_num(area),
    }

# ===== Previews (Detail) =====
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
                        try: b = float(verts[i][4])
                        except Exception: b = 0.0
                        seg = _bulge_arc_points((float(verts[i][0]), float(verts[i][1])),
                                                (float(verts[j][0]), float(verts[j][1])), b)
                        pts.extend(seg if not pts or pts[-1] != seg[0] else seg[1:])
                    if closed and pts and pts[0] != pts[-1]:
                        pts.append(pts[0])
            elif et == "POLYLINE":
                vs = list(ve.vertices())
                if vs:
                    tmp, coords = [], []
                    for v in vs:
                        loc=getattr(v.dxf,"location",None)
                        coords.append((float(loc.x),float(loc.y)) if loc is not None
                                      else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
                    closed = bool(getattr(ve,"is_closed",getattr(ve,"closed",False)))
                    n = len(coords)
                    for i in range(n - (0 if closed else 1)):
                        j = (i + 1) % n
                        try: b = float(vs[i].dxf.bulge)
                        except Exception: b = 0.0
                        seg = _bulge_arc_points(coords[i], coords[j], b)
                        tmp.extend(seg if not tmp or tmp[-1] != seg[0] else seg[1:])
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
                    minx = min(minx, x); miny = min(miny, y)
                    maxx = max(maxx, x); maxy = max(maxy, y)

        if minx == float("inf") or not polylines:
            return ""

        w = max(maxx - minx, 1.0); h = max(maxy - miny, 1.0)
        size = max(w, h); pad = max(size * pad_ratio, 0.5)
        cx = (minx + maxx) * 0.5; cy = (miny + maxy) * 0.5
        half = size * 0.5 + pad
        xmin, xmax = cx - half, cx + half
        ymin, ymax = cy - half, cy + half

        fig = plt.figure(figsize=(size_px/100, size_px/100), dpi=100)
        ax = fig.add_subplot(111)
        ax.axis("off"); ax.set_aspect("equal")
        for pts in polylines:
            xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
            ax.plot(xs, ys, linewidth=1.25)
        ax.set_xlim([xmin, xmax]); ax.set_ylim([ymin, ymax])

        buf = io.BytesIO()
        plt.subplots_adjust(0,0,1,1)
        fig.savefig(buf, format="png", transparent=True, bbox_inches="tight", pad_inches=0)
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        return ""

def _build_preview_cache(msp) -> Dict[str, str]:
    cache: Dict[str, str] = {}
    for ins in msp.query("INSERT"):
        try:
            name = getattr(ins, "effective_name", None) or getattr(ins, "block_name", None) or getattr(ins.dxf, "name", "")
            if not name or name in cache: continue
            cache[name] = _render_preview_from_insert(ins) or ""
        except Exception:
            pass
    return cache

def _layer_rgb_map(doc) -> Dict[str, tuple[int,int,int]]:
    m: Dict[str, tuple[int,int,int]] = {}
    try:
        for layer in doc.layers:
            name = layer.dxf.name or ""
            key = layer_or_misc(name)
            tc = getattr(layer.dxf, "true_color", 0)
            if tc:
                rgb = ezcolors.int2rgb(tc)
            else:
                aci = int(getattr(layer.dxf, "color", 7) or 7)
                rgb = ezcolors.aci2rgb(aci if 0 <= aci <= 256 else 7)
            m[key] = rgb
    except Exception:
        pass
    return m

def _resolve_entity_rgb(e, layer_rgb_map: Dict[str, tuple[int,int,int]]) -> tuple[int,int,int]:
    ly = layer_or_misc(getattr(e.dxf, "layer", ""))
    try:
        tc = int(getattr(e.dxf, "true_color", 0) or 0)
        if tc: return ezcolors.int2rgb(tc)
    except Exception:
        pass
    try:
        aci = int(getattr(e.dxf, "color", 256) or 256)
    except Exception:
        aci = 256
    if 1 <= aci <= 255:
        return ezcolors.aci2rgb(aci)
    return layer_rgb_map.get(ly, (200,200,200))

def _entity_weight_for_colorvote(e) -> float:
    et = e.dxftype()
    try:
        if et == "LINE":
            p1=(e.dxf.start.x,e.dxf.start.y); p2=(e.dxf.end.x,e.dxf.end.y)
            return dist2d(p1,p2)
        if et == "ARC":
            r=float(e.dxf.radius)
            sweep=(float(e.dxf.end_angle)-float(e.dxf.start_angle))%360.0
            return (2.0*math.pi*r)*(sweep/360.0)
        if et == "CIRCLE":
            r=float(e.dxf.radius); return 2.0*math.pi*r
        if et == "LWPOLYLINE":
            verts=list(e)
            if not verts: return 0.0
            closed=bool(getattr(e,"closed",False))
            dense=[]; n=len(verts)
            for i in range(n if closed else n-1):
                j=(i+1)%n
                try: b=float(verts[i][4])
                except Exception: b=0.0
                seg=_bulge_arc_points((float(verts[i][0]),float(verts[i][1])),
                                      (float(verts[j][0]),float(verts[j][1])), b)
                dense.extend(seg[:-1])
            dense.append((float(verts[-1][0]),float(verts[-1][1])))
            return polyline_length_xy(dense, closed=False)
        if et == "POLYLINE":
            vs=list(e.vertices())
            if not vs: return 0.0
            coords=[]
            for v in vs:
                loc=getattr(v.dxf,"location",None)
                coords.append((float(loc.x),float(loc.y)) if loc is not None
                              else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
            n=len(coords); closed=bool(getattr(e,"is_closed",getattr(e,"closed",False)))
            dense=[]
            for i in range(n - (0 if closed else 1)):
                j=(i+1)%n
                try: b=float(vs[i].dxf.bulge)
                except Exception: b=0.0
                seg=_bulge_arc_points(coords[i], coords[j], b)
                dense.extend(seg[:-1])
            dense.append(coords[-1])
            return polyline_length_xy(dense, closed=False)
        if et == "HATCH":
            if hasattr(e,"get_filled_area"):
                try: return float(e.get_filled_area()) or 0.0
                except Exception: pass
            return 0.0
    except Exception:
        pass
    return 0.0

def _dominant_layer_rgb_map(msp, base_layer_rgb: Dict[str, tuple[int,int,int]], scale_to_m: float) -> Dict[str, tuple[int,int,int]]:
    votes: Dict[str, Dict[tuple[int,int,int], float]] = {}
    def _acc(e):
        ly = layer_or_misc(getattr(e.dxf,"layer",""))
        rgb = _resolve_entity_rgb(e, base_layer_rgb)
        w   = _entity_weight_for_colorvote(e)
        w *= (scale_to_m**2) if e.dxftype()=="HATCH" else scale_to_m
        if w <= 0: w = 1.0
        d = votes.setdefault(ly, {})
        d[rgb] = d.get(rgb, 0.0) + w
    for et in ("LINE","LWPOLYLINE","POLYLINE","ARC","CIRCLE","HATCH"):
        for e in msp.query(et):
            try: _acc(e)
            except Exception: pass
    out = dict(base_layer_rgb)
    for ly, hist in votes.items():
        if hist:
            out[ly] = max(hist.items(), key=lambda kv: kv[1])[0]
    return out

# ===== Rows: INSERT detail =====
def iter_block_rows(msp, include_xrefs: bool,
                    scale_to_m: float, target_units: str,
                    preview_cache: Dict[str,str] | None = None,
                    zones: list[Zone] | None = None) -> list[dict]:
    out = []; preview_cache = preview_cache or {}; zones = zones or []
    for ins in msp.query("INSERT"):
        try:
            ly = getattr(ins.dxf, "layer", "")
            if layer_or_misc(ly).upper() == "PLANNER":  # skip zone rectangles
                continue
            name = getattr(ins, "effective_name", None) or getattr(ins, "block_name", None) or getattr(ins.dxf, "name", "")
            if not include_xrefs and ("|" in (name or "")):  # skip xrefs
                continue
            bbox_du = _bbox_of_insert_xy(ins)
            if bbox_du:
                L_m = bbox_du[0] * scale_to_m; W_m = bbox_du[1] * scale_to_m
                L_out = to_target_units(L_m, target_units, "length")
                W_out = to_target_units(W_m, target_units, "length")
            else:
                L_out = W_out = None

            center_zone = ""
            b = _insert_bbox(ins)
            if b:
                cx, cy = _bbox_center(b)
                zname = _zone_for_point((cx,cy), zones)
                if zname:
                    center_zone = zname

            upload_layer = "PLANNER" if (FORCE_PLANNER_CATEGORY and center_zone) else ly
            remarks_txt  = f"dwg_layer={ly}; aggregated 1 inserts"

            out.append(make_row(
                "INSERT", "count", 1.0,
                block_name=name,
                layer=upload_layer,
                handle=getattr(ins.dxf, "handle", ""),
                bbox_length=L_out, bbox_width=W_out,
                preview_b64=preview_cache.get(name, ""),
                zone=center_zone,
                category1=(ly or "").strip().lower(),
                remarks=remarks_txt
            ))
        except Exception as ex:
            logging.exception("INSERT failed: %s", ex)
    return out

# ===== Layer metrics =====
def compute_layer_metrics(msp, scale_to_m: float, target_units: str):
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

    for e in msp.query("LINE"):
        try:
            p1=(e.dxf.start.x,e.dxf.start.y); p2=(e.dxf.end.x,e.dxf.end.y)
            add_open_len(e.dxf.layer, dist2d(p1,p2))
        except Exception: pass

    for e in msp.query("LWPOLYLINE"):
        try:
            verts = list(e); 
            if not verts: continue
            closed = bool(getattr(e,"closed",False))
            dense=[]; n=len(verts)
            for i in range(n if closed else n-1):
                j=(i+1)%n
                try: b=float(verts[i][4])
                except Exception: b=0.0
                seg=_bulge_arc_points((float(verts[i][0]), float(verts[i][1])),
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

    for e in msp.query("POLYLINE"):
        try:
            vs = list(e.vertices())
            if not vs: continue
            coords=[]
            for v in vs:
                loc=getattr(v.dxf,"location",None)
                coords.append((float(loc.x),float(loc.y)) if loc is not None
                              else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
            closed = bool(getattr(e,"is_closed",getattr(e,"closed",False)))
            n=len(coords); dense=[]
            for i in range(n - (0 if closed else 1)):
                j=(i+1)%n
                try: b=float(vs[i].dxf.bulge)
                except Exception: b=0.0
                seg=_bulge_arc_points(coords[i], coords[j], b)
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

    for e in msp.query("ARC"):
        try:
            r=float(e.dxf.radius)
            sweep=(float(e.dxf.end_angle)-float(e.dxf.start_angle))%360.0
            add_open_len(e.dxf.layer, (2.0*math.pi*r)*(sweep/360.0))
        except Exception: pass

    for e in msp.query("CIRCLE"):
        try:
            r=float(e.dxf.radius)
            add_perimeter(e.dxf.layer, (2.0*math.pi*r))
            add_area(e.dxf.layer, math.pi*(r**2))
        except Exception: pass

    for e in msp.query("HATCH"):
        try:
            A_du=None
            if hasattr(e,"get_filled_area"):
                try:
                    v=e.get_filled_area()
                    if v and v>0: A_du=float(v)
                except Exception:
                    A_du=None
            if A_du and A_du>0: add_area(e.dxf.layer, A_du)
        except Exception: pass

    return open_len_by_layer, peri_by_layer, area_by_layer

def solve_rect_dims_from_perimeter_area(P: float, A: float) -> Tuple[Optional[float], Optional[float]]:
    try:
        if P is None or A is None or P <= 0 or A <= 0:
            return (None, None)
        S = P / 2.0
        D = S*S - 4.0*A
        if D < -1e-9: return (None, None)
        D = max(D, 0.0)
        root = math.sqrt(D)
        a = 0.5*(S + root)
        b = 0.5*(S - root)
        if a <= 0 or b <= 0: return (None, None)
        return (a, b) if a >= b else (b, a)
    except Exception:
        return (None, None)

def make_layer_total_rows(open_by, peri_by, area_by, layer_rgb: Dict[str, tuple[int,int,int]] | None = None,
                          mode: str = "split"):
    rows = []
    all_layers = sorted(set(open_by.keys()) | set(peri_by.keys()) | set(area_by.keys()))
    layer_rgb = layer_rgb or {}

    def hex_for(ly: str) -> str:
        rgb = layer_rgb.get(ly)
        return _rgb_to_hex(rgb) if rgb else ""

    if mode == "combined":
        for ly in all_layers:
            L_tot = open_by.get(ly, 0.0) + peri_by.get(ly, 0.0)
            A_tot = area_by.get(ly, 0.0)
            rows.append(make_row(
                "LAYER_SUMMARY","layer",None,
                layer=ly, remarks="totals per layer (open+closed length; area from closed)",
                bbox_length=(L_tot if L_tot>0 else None),
                bbox_width=(A_tot if A_tot>0 else None),
                preview_hex=hex_for(ly),
            ))
        return rows

    for ly in all_layers:
        if open_by.get(ly, 0.0) > 0:
            rows.append(make_row(
                "LAYER_SUMMARY","layer",None,
                layer=ly, remarks="OPEN length only",
                bbox_length=open_by[ly], bbox_width=None,
                preview_hex=hex_for(ly),
            ))
        if peri_by.get(ly, 0.0) > 0 or area_by.get(ly, 0.0) > 0:
            P = peri_by.get(ly, None); A = area_by.get(ly, None)
            L_rec, W_rec = solve_rect_dims_from_perimeter_area(P, A)
            rows.append(make_row(
                "LAYER_SUMMARY","layer",None,
                layer=ly, remarks="CLOSED (rectangle): length/width + perimeter & area",
                bbox_length=L_rec, bbox_width=W_rec,
                perimeter=P, area=A,
                preview_hex=hex_for(ly),
            ))
    return rows

# ===== I/O =====
def write_csv(rows: list[dict], out_path: Path) -> None:
    # Find the exact header labels for length/width/area from CSV_HEADERS
    def _find(hint: str) -> str:
        # pick the first header containing the hint (case-insensitive)
        for h in CSV_HEADERS:
            if hint.lower() in h.lower():
                return h
        raise KeyError(f"CSV_HEADERS missing a column containing: {hint}")

    LENGTH_COL = _find("length")
    WIDTH_COL  = _find("width")
    AREA_COL   = _find("area")
    PERI_COL   = _find("perimeter")
    PREV_COL   = _find("preview")
    BOQ_COL    = _find("boq name")
    QTY_T_COL  = _find("qty_type")
    QTY_V_COL  = _find("qty_value")
    ENT_COL    = _find("entity_type")
    CAT_COL    = _find("category")
    ZONE_COL   = _find("zone")
    CAT1_COL   = _find("category1")
    REM_COL    = _find("remarks")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        writer.writeheader()
        for r in rows:
            writer.writerow({
                ENT_COL:   r.get("entity_type",""),
                CAT_COL:   r.get("layer",""),
                ZONE_COL:  r.get("zone",""),
                CAT1_COL:  r.get("category1",""),
                BOQ_COL:   r.get("block_name",""),
                QTY_T_COL: r.get("qty_type",""),
                QTY_V_COL: r.get("qty_value",""),
                LENGTH_COL: r.get("bbox_length",""),
                WIDTH_COL:  r.get("bbox_width",""),
                PERI_COL:   r.get("perimeter",""),
                AREA_COL:   r.get("area",""),
                PREV_COL:   "",  # preview images handled by Web App
                REM_COL:     r.get("remarks",""),
            })

def push_rows_to_webapp(rows: list[dict], webapp_url: str, spreadsheet_id: str,
                        tab: str, mode: str = "replace", summary_tab: str = "",
                        batch_rows: int = 25, timeout: int = 300,
                        valign_middle: bool = False, sparse_anchor: str = "last",
                        drive_folder_id: str = "") -> None:
    if not webapp_url or not spreadsheet_id or not tab:
        logging.info("WebApp push not configured (missing url/id/tab). Skipping upload.")
        return

    sess = requests.Session()
    first_mode = (mode or "replace").lower()

    def post_with_retries(payload, tries=4, backoff=2.0):
        for attempt in range(1, tries+1):
            try:
                return sess.post(webapp_url, json=payload, timeout=timeout, allow_redirects=True)
            except requests.exceptions.ReadTimeout:
                if attempt == tries: raise
                time.sleep(backoff ** attempt)

    total = len(rows); sent = 0
    for idx, i in enumerate(range(0, total, batch_rows), start=1):
        chunk = rows[i:i+batch_rows]
        if not chunk: break

        is_layer = (chunk[0].get("entity_type") == "LAYER_SUMMARY")

        if is_layer:
            # ===== ByLayer sheet: zone, category1, BOQ name, qty_value REMOVED
            headers = LAYER_HEADERS
            data_rows = [[
                r.get("entity_type",""),
                r.get("layer",""),           # category
                r.get("qty_type",""),
                r.get("bbox_length",""),
                r.get("bbox_width",""),
                r.get("perimeter",""),
                r.get("area",""),
                "",                          # Preview (color cell handled separately)
                r.get("remarks",""),
            ] for r in chunk]
            images     = [""] * len(chunk)                      # no images
            bg_colors  = [r.get("preview_hex","") for r in chunk]
            color_only = True
        else:
            # ===== Detail sheet: perimeter & area REMOVED
            headers = DETAIL_HEADERS
            data_rows = [[
                r.get("entity_type",""),
                r.get("layer",""),          # category
                r.get("zone",""),
                r.get("category1",""),
                r.get("block_name",""),
                r.get("qty_type",""),
                r.get("qty_value",""),
                r.get("bbox_length",""),
                r.get("bbox_width",""),
                "",                         # Preview (image written by Web App)
                r.get("remarks",""),
            ] for r in chunk]
            images     = [r.get("preview_b64","") for r in chunk]
            bg_colors  = [""] * len(chunk)
            color_only = False

        payload = {
            "sheetId": spreadsheet_id,
            "tab": tab,
            "mode": "replace" if (i == 0 and first_mode == "replace") else "append",
            "headers": headers if (i == 0 and first_mode == "replace") else [],
            "rows": data_rows,
            "images": images,
            "bgColors": bg_colors,
            "colorOnly": color_only,
            "embedImages": False,
            "driveFolderId": (drive_folder_id or ""),
            "vAlign": "middle" if valign_middle else "",
            "sparseAnchor": (sparse_anchor or "last"),
        }

        if summary_tab and i == 0:
            payload["summaryTab"] = summary_tab
            payload["summaryRows"] = []

        r = post_with_retries(payload)
        if not r.ok:
            raise RuntimeError(f"WebApp upload failed (batch {idx}): HTTP {r.status_code} {r.text}")
        sent += len(data_rows)
        logging.info("WebApp batch %d: uploaded %d/%d rows", idx, sent, total)

# ===== Misc helpers =====
def print_summary(rows: list[dict], out_path: Path) -> None:
    total_insert_groups = sum(1 for r in rows if r["entity_type"]=="INSERT" and r["qty_type"]=="count")
    logging.info("----- SUMMARY for %s -----", out_path.name)
    logging.info("INSERT groups (after aggregation): %d", total_insert_groups)
    logging.info("CSV written to: %s", out_path)

def _norm_cat(s: str) -> str:
    s = (s or "").strip()
    s = " ".join(s.split())
    return s.upper()

def sort_rows_for_category_blocks(rows: list[dict]) -> None:
    def _key(r):
        cat  = _norm_cat(r.get("layer", ""))
        zone = (r.get("zone","") or "").lower()
        cat1 = (r.get("category1","") or "").lower()
        et   = r.get("entity_type", "")
        et_rank = 0 if et == "INSERT" else 1
        return (cat, zone, cat1, et_rank, r.get("block_name",""))
    rows.sort(key=_key)

def collect_dxf_files(path: Path, recursive: bool) -> List[Path]:
    if path.is_file():
        if path.suffix.lower() == ".dxf": return [path]
        logging.error("Provided file is not a .dxf: %s", path); return []
    if path is None or not path.exists():
        logging.error("Path does not exist: %s", path); return []
    pattern = "**/*.dxf" if recursive else "*.dxf"
    files = sorted(path.glob(pattern))
    if not files: logging.warning("No DXF files found in %s (recursive=%s)", path, recursive)
    return files

def derive_out_path(dxf_path: Path, out_dir: Path | None) -> Path:
    return (out_dir / f"{dxf_path.stem}_raw_extract.csv") if out_dir else dxf_path.with_name(f"{dxf_path.stem}_raw_extract.csv")

def split_rows_for_upload(rows: list[dict]) -> tuple[list[dict], list[dict]]:
    detail, layer = [], []
    for r in rows:
        if (r.get("entity_type") == "LAYER_SUMMARY") and (r.get("qty_type") == "layer"):
            layer.append(r)
        else:
            detail.append(r)
    return detail, layer

# ===== Main pipeline =====
def process_one_dxf(dxf_path: Path, out_dir: Path | None,
                    target_units: str, include_xrefs: bool,
                    layer_metrics: bool, aggregate_inserts: bool,
                    layer_metrics_mode: str) -> list[dict]:
    logging.info("Processing DXF: %s", dxf_path)

    doc = ezdxf.readfile(str(dxf_path)); msp = doc.modelspace()
    scale_to_m = units_scale_to_meters(doc)

    preview_cache = _build_preview_cache(msp)
    zones = _collect_planner_zones(msp)

    rows: list[dict] = []

    insert_rows = iter_block_rows(msp, include_xrefs, scale_to_m, target_units, preview_cache, zones)

    if aggregate_inserts:
        groups: Dict[tuple[str,str,str], dict] = {}
        for r in insert_rows:
            key = (r["block_name"], "PLANNER" if FORCE_PLANNER_CATEGORY else r["layer"], r.get("zone",""))
            g = groups.setdefault(key, {"count":0,"xs":[],"ys":[], "preview": r.get("preview_b64",""),
                                        "category1": r.get("category1","")})
            g["count"] += 1
            try:
                if r["bbox_length"] and r["bbox_width"]:
                    g["xs"].append(float(r["bbox_length"]))
                    g["ys"].append(float(r["bbox_width"]))
            except Exception:
                pass
        for (name, layer, zone_name), g in groups.items():
            xs = sorted(g["xs"]); ys = sorted(g["ys"])
            bx = xs[len(xs)//2] if xs else None
            by = ys[len(ys)//2] if ys else None
            rows.append(make_row(
                "INSERT", "count", float(g["count"]),
                block_name=name, layer=layer, handle="",
                remarks=f"aggregated {g['count']} inserts", bbox_length=bx, bbox_width=by,
                preview_b64=g.get("preview",""), zone=zone_name, category1=g.get("category1","")
            ))
    else:
        rows.extend(insert_rows)

    if layer_metrics:
        open_by, peri_by, area_by = compute_layer_metrics(msp, scale_to_m, target_units)
        base_layer_rgb = _layer_rgb_map(doc)
        dom_layer_rgb  = _dominant_layer_rgb_map(msp, base_layer_rgb, scale_to_m)
        rows.extend(make_layer_total_rows(open_by, peri_by, area_by, layer_rgb=dom_layer_rgb, mode=layer_metrics_mode))

    sort_rows_for_category_blocks(rows)

    out_path = derive_out_path(dxf_path, out_dir)
    write_csv(rows, out_path)
    logging.info("CSV written to: %s", out_path)
    return rows

def main():
    ap = argparse.ArgumentParser(description="DXF → CSV + Sheets upload (previews + PLANNER zones + category1).")
    ap.add_argument("--dxf"); ap.add_argument("--name")
    ap.add_argument("--decimals", type=int, default=None)
    ap.add_argument("--out-dir"); ap.add_argument("--out")
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--target-units", default="ft")
    ap.add_argument("--include-xrefs", action="store_true")
    ap.add_argument("--no-layer-metrics", action="store_true")
    ap.add_argument("--no-aggregate-inserts", action="store_true")
    ap.add_argument("--layer-metrics-mode", choices=["combined","split"], default="split")
    ap.add_argument("--gs-webapp", default=None); ap.add_argument("--gsheet-id", default=None)
    ap.add_argument("--gsheet-tab", default=None); ap.add_argument("--gsheet-summary-tab", default=None)
    ap.add_argument("--gsheet-mode", choices=["replace","append"], default=None)
    ap.add_argument("--batch-rows", type=int, default=3000)
    ap.add_argument("--align-middle", action="store_true")
    ap.add_argument("--sparse-anchor", choices=["first","last","middle"], default="last")
    ap.add_argument("--drive-folder-id", default=None)
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    global DEC_PLACES
    if args.decimals is not None:
        DEC_PLACES = max(0, min(10, args.decimals))

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    dxf_input = Path(args.dxf) if args.dxf else (Path(DXF_FOLDER)/f"{args.name}.dxf" if args.name else Path(DXF_FOLDER))
    out_dir   = Path(args.out_dir) if args.out_dir else Path(OUT_ROOT)
    explicit_out = Path(args.out) if args.out else None

    layer_metrics = not args.no_layer_metrics
    aggregate_inserts = not args.no_aggregate_inserts

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
        logging.error("DXF input path not found: %s", dxf_input); return

    files = collect_dxf_files(dxf_input, recursive=args.recursive)
    if not files: return

    def _summary_tab_name():
        return gsheet_summary_tab if gsheet_summary_tab else (gsheet_tab + "_ByLayer")

    if explicit_out:
        if len(files) != 1:
            logging.error("--out is for a single file. For folders, use --out-dir."); return
        f = files[0]
        rows = process_one_dxf(f, explicit_out.parent, args.target_units, args.include_xrefs,
                               layer_metrics, aggregate_inserts, args.layer_metrics_mode)
        explicit_out.parent.mkdir(parents=True, exist_ok=True)
        write_csv(rows, explicit_out)

        if gs_webapp and gsheet_id:
            detail_rows, layer_rows = split_rows_for_upload(rows)
            if detail_rows:
                push_rows_to_webapp(detail_rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, "",
                                    batch_rows=batch_rows, valign_middle=align_middle,
                                    sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)
            if layer_rows:
                push_rows_to_webapp(layer_rows, gs_webapp, gsheet_id, _summary_tab_name(),
                                    "replace" if gsheet_mode=="replace" else "append", "",
                                    batch_rows=batch_rows, valign_middle=align_middle,
                                    sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)
        return

    out_dir = out_dir if str(out_dir).strip() else None
    all_rows: list[dict] = []
    for f in files:
        try:
            rows = process_one_dxf(f, out_dir, args.target_units, args.include_xrefs,
                                   layer_metrics, aggregate_inserts, args.layer_metrics_mode)
            all_rows.extend(rows or [])
        except Exception as ex:
            logging.exception("Failed processing %s: %s", f, ex)

    if all_rows and gs_webapp and gsheet_id:
        detail_rows, layer_rows = split_rows_for_upload(all_rows)
        if detail_rows:
            push_rows_to_webapp(detail_rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, "", 
                                batch_rows=batch_rows, valign_middle=align_middle,
                                sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)
        if layer_rows:
            push_rows_to_webapp(layer_rows, gs_webapp, gsheet_id, _summary_tab_name(),
                                "replace" if gsheet_mode=="replace" else "append", "",
                                batch_rows=batch_rows, valign_middle=align_middle,
                                sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)

if __name__ == "__main__":
    main()
