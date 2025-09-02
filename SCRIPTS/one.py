# export_instances_and_clusters_anchor_darken.py
# One image per block INSTANCE + anchor/star AUTO-CLUSTERS.
# Keeps white background and simply DARKENS the drawn linework (no inversion).
# Adds optional slight thickening so anti-aliased edges darken too.

import os
import math
import numpy as np
import cv2

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import ezdxf
from collections import defaultdict
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

# ====== EDIT THESE PATHS ======
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"
# ==============================

# What to export
EXPORT_BLOCK_INSTANCES = True
EXPORT_AUTO_CLUSTERS   = True
EXPORT_GROUPS          = False

# Rendering & framing
DPI            = 240
PAD_PCT        = 0.04     # pre-save padding so nothing clips
MARGIN_PCT     = 0.10     # visible border after trim
OUTPUT_EXT     = ".jpg"

# Style
WHITE_ON_BLACK = False    # keep False (we're not inverting)

# Line enhancement
THICKEN_PX     = 1        # 0=off; 1–3 recommended (elliptical kernel)
INK_THRESHOLD  = 245      # 0–255: what counts as "ink" vs white
DARKEN_ALPHA   = 0.65     # 0..1: pull ink toward black (higher = darker)

# Skip labels/dims etc. while rendering
SKIP_LAYERS    = {"DEFPOINTS", "DIM", "DIMENSIONS", "ANNOTATION", "TEXT", "NOTES"}
SKIP_DXF_TYPES = {"TEXT", "MTEXT", "DIMENSION"}

# Subfolders
INSTANCES_DIRNAME = "INSTANCES"
GROUPS_DIRNAME    = "GROUPS"
AUTOCLUSTERS_DIR  = "AUTO_CLUSTERS"

# ---------- Anchor/star clustering tuning ----------
AUTOCLUSTER_EXPAND_BY   = None
AUTOCLUSTER_EXPAND_K    = 0.55
ANCHOR_AREA_PERCENTILE  = 70
CENTER_EPS_K            = 0.45
AUTOCLUSTER_MIN_SIZE    = 2
# ---------------------------------------------------

BACKGROUND = (1, 1, 1, 1)  # render on white


def make_safe(name: str) -> str:
    for c in '<>:"/\\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unnamed"


def build_frontend(doc, ax):
    ctx     = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    try:
        from ezdxf.addons.drawing.config import Configuration
        try:
            cfg = Configuration.defaults()
        except Exception:
            cfg = Configuration()
        return Frontend(ctx, backend, cfg)
    except Exception:
        return Frontend(ctx, backend)


def render_entities_to_file(doc, entities, out_path, dpi=DPI, pad_pct=PAD_PCT):
    fig = plt.figure()
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BACKGROUND)
    fig.set_dpi(dpi)

    frontend = build_frontend(doc, ax)
    msp = doc.modelspace()
    for e in entities:
        if SKIP_LAYERS and hasattr(e.dxf, "layer") and e.dxf.layer in SKIP_LAYERS:
            continue
        if SKIP_DXF_TYPES and e.dxftype() in SKIP_DXF_TYPES:
            continue
        try:
            frontend.draw_entity(e, msp)
        except Exception:
            pass

    ax.set_aspect("equal", adjustable="box")
    ax.autoscale(True, axis="both", tight=True)
    x0, x1 = ax.get_xlim(); y0, y1 = ax.get_ylim()
    w = max(x1 - x0, 1e-6); h = max(y1 - y0, 1e-6)
    ax.set_xlim(x0 - w*pad_pct, x1 + w*pad_pct)
    ax.set_ylim(y0 - h*pad_pct, y1 + h*pad_pct)
    ax.axis("off")
    fig.savefig(out_path, bbox_inches="tight", pad_inches=0)
    plt.close(fig)


def trim_margin_bw_and_thicken(path_in: str):
    """
    Trim, add white margin, then DARKEN only the linework (non-white pixels).
    Optional: thicken the mask slightly so anti-aliased edges also darken.
    """
    img = cv2.imread(path_in)
    if img is None:
        return

    # 1) Trim the renderer's outer white frame
    gray0 = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, th = cv2.threshold(gray0, 250, 255, cv2.THRESH_BINARY_INV)
    coords = cv2.findNonZero(th)
    if coords is not None:
        x, y, w, h = cv2.boundingRect(coords)
        img = img[y:y+h, x:x+w]

    # 2) Add uniform white margin
    h, w = img.shape[:2]
    size = max(h, w)
    margin = int(round(size * MARGIN_PCT))
    img = cv2.copyMakeBorder(img, margin, margin, margin, margin,
                             cv2.BORDER_CONSTANT, value=[255, 255, 255])

    # 3) Build an ink mask: anything not near-white is "linework"
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, ink = cv2.threshold(gray, INK_THRESHOLD, 255, cv2.THRESH_BINARY_INV)

    # 4) Optional slight thickening so edge pixels are included
    if THICKEN_PX > 0:
        k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (THICKEN_PX, THICKEN_PX))
        ink = cv2.dilate(ink, k, iterations=1)

    # 5) Darken only where mask says there is ink (blend toward black)
    mask = ink.astype(bool)
    if mask.any():
        src = img.astype(np.float32)
        src[mask] = src[mask] * (1.0 - DARKEN_ALPHA)  # multiply toward 0 (black)
        img = np.clip(src, 0, 255).astype(np.uint8)

    # 6) Save
    params = []
    if OUTPUT_EXT.lower() == ".jpg":
        params = [int(cv2.IMWRITE_JPEG_QUALITY), 95]
    cv2.imwrite(path_in, img, params)


# ---------------- Instances ----------------
def export_instances(doc, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    msp = doc.modelspace()
    per_name_counter = defaultdict(int)
    exported = 0

    for insert in msp.query("INSERT"):
        if SKIP_LAYERS and insert.dxf.layer in SKIP_LAYERS:
            continue
        name = getattr(insert.dxf, "name", "")
        if not name:
            continue
        base = make_safe(name)
        per_name_counter[base] += 1
        seq = per_name_counter[base]
        out_file = os.path.join(out_dir, f"{base}__{seq:03d}{OUTPUT_EXT}")
        try:
            render_entities_to_file(doc, [insert], out_file)
            trim_margin_bw_and_thicken(out_file)
            exported += 1
            print(f"  ✓ Instance {name} -> {out_file}")
        except Exception as e:
            print(f" X ! Instance {name} failed: {e}")
    print(f"  Instances done: {exported} images")


# ---------- helpers for anchor clustering ----------
def bbox_of_insert(doc, insert):
    fig = plt.figure()
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BACKGROUND)
    fig.set_dpi(72)
    frontend = build_frontend(doc, ax)
    msp = doc.modelspace()
    try:
        frontend.draw_entity(insert, msp)
    except Exception:
        plt.close(fig)
        return None
    ax.set_aspect("equal", adjustable="box")
    ax.autoscale(True, axis="both", tight=True)
    x0, x1 = ax.get_xlim(); y0, y1 = ax.get_ylim()
    plt.close(fig)
    return (float(x0), float(y0), float(x1), float(y1))

def bbox_expand(bb, expand):
    x0,y0,x1,y1 = bb
    return (x0-expand, y0-expand, x1+expand, y1+expand)

def boxes_overlap(a, b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return not (ax1 < bx0 or bx1 < ax0 or ay1 < by0 or by1 < ay0)

def bbox_center(bb):
    x0,y0,x1,y1 = bb
    return ((x0+x1)/2.0, (y0+y1)/2.0)

# ------------- AUTO CLUSTERS (ANCHOR / STAR) -------------
def export_auto_clusters(doc, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    msp = doc.modelspace()

    inserts = [e for e in msp.query("INSERT") if not (SKIP_LAYERS and e.dxf.layer in SKIP_LAYERS)]
    if not inserts:
        print("  (No INSERTs to cluster)")
        return

    boxes, areas, diags = [], [], []
    for ins in inserts:
        bb = bbox_of_insert(doc, ins)
        if bb is None:
            boxes.append(None); areas.append(0); diags.append(0)
            continue
        x0,y0,x1,y1 = bb
        boxes.append(bb)
        areas.append((x1-x0)*(y1-y0))
        diags.append(math.hypot(x1-x0, y1-y0))

    nonzero_diags = [d for d in diags if d > 0]
    median_diag   = sorted(nonzero_diags)[len(nonzero_diags)//2] if nonzero_diags else 1.0
    expand = AUTOCLUSTER_EXPAND_BY if AUTOCLUSTER_EXPAND_BY is not None else median_diag * AUTOCLUSTER_EXPAND_K
    center_eps = median_diag * CENTER_EPS_K

    valid_idx   = [i for i,b in enumerate(boxes) if b is not None]
    valid_areas = [areas[i] for i in valid_idx]
    if not valid_areas:
        print("  (No valid bboxes)")
        return
    thresh_area = np.percentile(valid_areas, ANCHOR_AREA_PERCENTILE)
    anchor_ids  = [i for i in valid_idx if areas[i] >= thresh_area]
    anchor_ids.sort(key=lambda i: areas[i], reverse=True)

    assigned = set()
    exported = 0
    cluster_id = 0

    for ai in anchor_ids:
        if ai in assigned: 
            continue
        anchor = inserts[ai]
        abb    = boxes[ai]
        if abb is None:
            continue

        # Build star cluster around this anchor (NO transitive chaining)
        Aexp   = bbox_expand(abb, expand)
        Acx, Acy = bbox_center(abb)
        members = [anchor]
        member_idx = [ai]

        for i in valid_idx:
            if i == ai or i in assigned:
                continue
            bb = boxes[i]
            if bb is None:
                continue
            # condition 1: overlaps expanded anchor bbox
            if boxes_overlap(bb, Aexp):
                members.append(inserts[i]); member_idx.append(i); continue
            # condition 2: within center radius of anchor
            cx, cy = bbox_center(bb)
            if math.hypot(cx - Acx, cy - Acy) <= center_eps:
                members.append(inserts[i]); member_idx.append(i)

        if len(members) >= AUTOCLUSTER_MIN_SIZE:
            cluster_id += 1
            for i in member_idx:
                assigned.add(i)

            # Name by anchor’s block name
            anchor_name = make_safe(getattr(anchor.dxf, "name", "Cluster"))
            out_file = os.path.join(out_dir, f"{anchor_name}__cluster{cluster_id:03d}{OUTPUT_EXT}")
            try:
                render_entities_to_file(doc, members, out_file)
                trim_margin_bw_and_thicken(out_file)
                exported += 1
                print(f"  ✓ Cluster {cluster_id} ({anchor_name}) -> {out_file}  [{len(members)} inserts]")
            except Exception as e:
                print(f"  ! Cluster {cluster_id} failed: {e}")

    print(f"  Auto-clusters done: {exported} images (expand={expand:.2f}, center_eps={center_eps:.2f})")


# ---------------- Groups (optional; unchanged) ----------------
def export_groups(doc, out_dir):
    if not hasattr(doc, "groups") or len(doc.groups) == 0:
        print("  (No GROUPs found)")
        return
    os.makedirs(out_dir, exist_ok=True)
    per_name_counter = defaultdict(int)
    exported = 0
    for gname in doc.groups.group_names():
        try:
            grp = doc.groups.get(gname)
        except Exception:
            continue
        ents = []
        for h in getattr(grp, "entity_handles", []):
            e = doc.entitydb.get(h)
            if e is None: continue
            if SKIP_LAYERS and hasattr(e.dxf, "layer") and e.dxf.layer in SKIP_LAYERS: continue
            if SKIP_DXF_TYPES and e.dxftype() in SKIP_DXF_TYPES: continue
            ents.append(e)
        if not ents:
            continue
        base = make_safe(gname)
        per_name_counter[base] += 1
        seq = per_name_counter[base]
        out_file = os.path.join(out_dir, f"{base}__{seq:03d}{OUTPUT_EXT}")
        try:
            render_entities_to_file(doc, ents, out_file)
            trim_margin_bw_and_thicken(out_file)
            exported += 1
            print(f"  ✓ Group {gname} -> {out_file}")
        except Exception as e:
            print(f"  ! Group {gname} failed: {e}")
    print(f"  Groups done: {exported} images")


# ---------------- Driver ----------------
def export_one_dxf(dxf_path: str, out_root_for_drawing: str):
    os.makedirs(out_root_for_drawing, exist_ok=True)
    doc = ezdxf.readfile(dxf_path)
    if EXPORT_BLOCK_INSTANCES:
        export_instances(doc, os.path.join(out_root_for_drawing, INSTANCES_DIRNAME))
    if EXPORT_AUTO_CLUSTERS:
        export_auto_clusters(doc, os.path.join(out_root_for_drawing, AUTOCLUSTERS_DIR))
    if EXPORT_GROUPS:
        export_groups(doc, os.path.join(out_root_for_drawing, GROUPS_DIRNAME))

def main():
    print(f"Scanning: {DXF_FOLDER}")
    for name in os.listdir(DXF_FOLDER):
        if not name.lower().endswith(".dxf"):
            continue
        dxf_path = os.path.join(DXF_FOLDER, name)
        draw_name = os.path.splitext(name)[0]
        out_dir   = os.path.join(OUT_ROOT, draw_name)
        print(f"\nProcessing {name} …")
        export_one_dxf(dxf_path, out_dir)

if __name__ == "__main__":
    main()
