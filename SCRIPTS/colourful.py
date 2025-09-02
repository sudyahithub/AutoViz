# export_instances_and_clusters_anchor_darken.py
# One image per block INSTANCE + anchor/star AUTO-CLUSTERS.
# Output: transparent PNG (RGBA) at exactly N×N (configurable with --size).
# Linework = white (opaque), background = fully transparent.
# Robust transparency: transparent matplotlib render + edge-connected background,
# gap-closing, pre/post thickening, optional glow, with SAFE mask shaping.

import os
import math
import argparse
import numpy as np
import cv2

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import ezdxf
from collections import defaultdict
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

# ====== EDIT THESE DEFAULT PATHS (can be overridden via CLI) ======
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"
# =================================================================

# What to export
EXPORT_BLOCK_INSTANCES = True
EXPORT_AUTO_CLUSTERS   = True
EXPORT_GROUPS          = False

# Rendering & framing
DPI            = 240
PAD_PCT        = 0.04     # pre-save padding so nothing clips
MARGIN_PCT     = 0.10     # margin after trim (before resize)

# ---------- Output settings ----------
OUTPUT_EXT     = ".png"   # transparent PNG
TARGET_SIZE    = 128      # default final width & height (can be overridden by --size)
# -------------------------------------

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

# Matplotlib background (transparent while saving)
BACKGROUND = (0, 0, 0, 0)

# ===== Stronger visibility controls =====
# Ink detection & pre-resize enhancement
INK_THRESHOLD     = 252   # higher captures faint anti-aliased pixels
THICKEN_PX        = 3     # pre-resize dilation kernel size (pixels)
THICKEN_ITER      = 2     # pre-resize dilation iterations
CLOSE_GAPS_KSIZE  = 2     # morphological close kernel to bridge dash gaps

# Post-resize stroke boost (scales with final size; base tuned for 128)
POST_THICKEN_PX_BASE   = 1     # at 128; scales with size/128
POST_THICKEN_ITER      = 1

# Optional soft glow to improve contrast on dark UIs
ADD_GLOW          = True
GLOW_BLUR_PX      = 2     # Gaussian blur sigma/pixel
GLOW_ALPHA        = 90    # 0..255 opacity of halo

# Background extraction & cleanup
EDGE_BG_THRESH     = 245  # near-white considered background when connected to edges
MIN_VISIBLE_ALPHA  = 8    # alpha below this becomes fully transparent
# =======================================


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
    """Render selected entities with a transparent canvas."""
    fig = plt.figure()
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor((0, 0, 0, 0))
    fig.patch.set_alpha(0.0)
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

    fig.savefig(out_path, bbox_inches="tight", pad_inches=0, transparent=True)
    plt.close(fig)


# ---------- Helpers: shapes & RGBA safety ----------

def _ensure_rgba(img):
    """Return (H,W,4) uint8 RGBA image from BGR/BGRA/GRAY (handles weird cases)."""
    if img is None:
        return None
    if img.ndim == 2:
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
    elif img.ndim == 3:
        c = img.shape[2]
        if c == 4:
            pass
        elif c == 3:
            img = np.dstack([img, np.full(img.shape[:2], 255, np.uint8)])
        elif c == 1:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
        else:
            img = img[..., :3]
            img = np.dstack([img, np.full(img.shape[:2], 255, np.uint8)])
    return np.ascontiguousarray(img, dtype=np.uint8)


def _as_single_channel(mask):
    """Return a contiguous 2D uint8 array from mask that may be HxW, HxW x1, HxW x3, or bool."""
    if mask is None:
        return None
    if mask.ndim == 3 and mask.shape[2] == 1:
        mask = mask[..., 0]
    elif mask.ndim == 3 and mask.shape[2] == 3:
        mask = cv2.cvtColor(mask, cv2.COLOR_BGR2GRAY)
    if mask.dtype != np.uint8:
        mask = (mask > 0).astype(np.uint8)
    return np.ascontiguousarray(mask)


def _resize_to_square_rgba(img_rgba, size: int):
    """Fit RGBA image onto a size×size transparent canvas, preserving aspect ratio."""
    img_rgba = _ensure_rgba(img_rgba)
    h, w = img_rgba.shape[:2]
    if h == 0 or w == 0:
        return np.zeros((size, size, 4), np.uint8)

    scale = min(size / w, size / h)
    new_w = max(1, int(round(w * scale)))
    new_h = max(1, int(round(h * scale)))

    interp = cv2.INTER_AREA if (new_w < w or new_h < h) else cv2.INTER_CUBIC
    resized = cv2.resize(img_rgba, (new_w, new_h), interpolation=interp)
    resized = _ensure_rgba(resized)

    canvas = np.zeros((size, size, 4), dtype=np.uint8)
    x0 = (size - new_w) // 2
    y0 = (size - new_h) // 2
    canvas[y0:y0+new_h, x0:x0+new_w] = resized
    return canvas


def _add_transparent_margin_rgba(img_rgba, margin):
    img_rgba = _ensure_rgba(img_rgba)
    if margin <= 0:
        return img_rgba
    h, w = img_rgba.shape[:2]
    canvas = np.zeros((h + 2*margin, w + 2*margin, 4), dtype=np.uint8)
    canvas[margin:margin+h, margin:margin+w] = img_rgba
    return canvas


def _dilate_mask(mask, ksize, iters):
    m = _as_single_channel(mask)
    if m is None or ksize <= 0 or iters <= 0:
        return m
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (max(1, ksize), max(1, ksize)))
    out = cv2.dilate(m, k, iterations=max(1, iters))
    return _as_single_channel(out)


def _close_small_gaps(mask, ksize=2):
    m = _as_single_channel(mask)
    if m is None or ksize <= 0:
        return m
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (ksize, ksize))
    out = cv2.morphologyEx(m, cv2.MORPH_CLOSE, k, iterations=1)
    return _as_single_channel(out)


def _apply_glow(out_rgba, alpha_mask, blur_px=2, glow_alpha=90):
    out_rgba = _ensure_rgba(out_rgba)
    a_in = _as_single_channel(alpha_mask)
    if blur_px <= 0 or glow_alpha <= 0:
        return out_rgba
    glow = cv2.GaussianBlur(a_in, (0, 0), blur_px)
    glow = np.clip(glow, 0, 255).astype(np.uint8)

    # Compose: semi-opaque white glow under crisp white stroke
    h, w = glow.shape
    base = np.zeros((h, w, 4), dtype=np.uint8)
    base[..., :3] = 255
    base[..., 3]  = (glow.astype(np.uint16) * glow_alpha // 255).astype(np.uint8)

    # Alpha-over compositing: base UNDER out_rgba
    a0 = base[..., 3:4].astype(np.float32) / 255.0
    a1 = out_rgba[..., 3:4].astype(np.float32) / 255.0
    rgb0 = base[..., :3].astype(np.float32)
    rgb1 = out_rgba[..., :3].astype(np.float32)

    a = a1 + a0 * (1 - a1)
    rgb = (rgb1 * a1 + rgb0 * a0 * (1 - a1)) / np.clip(a, 1e-6, 1.0)

    out = np.zeros_like(out_rgba)
    out[..., :3] = np.clip(rgb, 0, 255).astype(np.uint8)
    out[..., 3]  = np.clip(a * 255, 0, 255).astype(np.uint8)
    return out


def _labels_touching_border(labels):
    h, w = labels.shape
    border_labels = set(np.unique(np.r_[labels[0,:], labels[-1,:], labels[:,0], labels[:,-1]]))
    return border_labels


# ---------- Main conversion ----------

def trim_and_to_rgba_white_lines(path_in: str, target_size: int):
    """
    Robust pipeline:
      - Read RGBA from Matplotlib (transparent).
      - Trim to content using alpha OR non-white RGB.
      - Add transparent margin.
      - Build background mask as edge-connected near-white/transparent.
      - Ink mask = not background AND (alpha >= MIN_VISIBLE_ALPHA OR gray < INK_THRESHOLD).
      - Close gaps + pre-resize thicken.
      - White strokes on transparent.
      - Resize to target_size.
      - Post-resize thicken (scaled) + optional glow.
      - Save PNG (RGBA).
    """
    img = cv2.imread(path_in, cv2.IMREAD_UNCHANGED)
    img = _ensure_rgba(img)
    if img is None:
        return

    # 1) Trim: any pixel that has alpha>0 OR is not near-white is content
    rgb  = img[..., :3]
    a    = img[..., 3]
    gray = cv2.cvtColor(rgb, cv2.COLOR_BGR2GRAY)
    non_white = (gray < 250).astype(np.uint8) * 255
    trim_mask = cv2.bitwise_or((a > 0).astype(np.uint8) * 255, non_white)

    coords = cv2.findNonZero(trim_mask)
    if coords is not None:
        x, y, w, h = cv2.boundingRect(coords)
        if w > 0 and h > 0:
            img = img[y:y+h, x:x+w]
            img = _ensure_rgba(img)
            rgb  = img[..., :3]
            a    = img[..., 3]
            gray = cv2.cvtColor(rgb, cv2.COLOR_BGR2GRAY)

    # 2) Transparent margin before scaling
    size_max = max(img.shape[0], img.shape[1])
    margin = int(round(size_max * MARGIN_PCT))
    img = _add_transparent_margin_rgba(img, margin)
    rgb  = img[..., :3]
    a    = img[..., 3]
    gray = cv2.cvtColor(rgb, cv2.COLOR_BGR2GRAY)

    # 3) Background mask via edge-connected near-white OR low-alpha
    bg_candidates = _as_single_channel(((gray >= EDGE_BG_THRESH) | (a <= MIN_VISIBLE_ALPHA)).astype(np.uint8))
    _, labels = cv2.connectedComponents(bg_candidates, connectivity=8)
    border_lbls = _labels_touching_border(labels)
    bg_mask = np.isin(labels, list(border_lbls)).astype(np.uint8) * 255  # 0/255

    # 4) Ink mask (robust)
    ink = ((bg_mask == 0) & ((a >= MIN_VISIBLE_ALPHA) | (gray < INK_THRESHOLD))).astype(np.uint8) * 255
    ink = _close_small_gaps(ink, ksize=CLOSE_GAPS_KSIZE)
    ink = _dilate_mask(ink, THICKEN_PX, THICKEN_ITER)

    # 5) RGBA output with white strokes
    h, w = img.shape[:2]
    out = np.zeros((h, w, 4), dtype=np.uint8)      # transparent by default
    mask = _as_single_channel(ink) > 0
    out[mask, :3] = 255                            # white RGB
    out[mask,  3] = 255                            # opaque alpha

    # 6) Resize to target_size transparent canvas
    out = _resize_to_square_rgba(out, target_size)

    # 7) Post-resize thicken scaled for size (baseline is 128)
    post_px = max(1, int(round(POST_THICKEN_PX_BASE * (target_size / 128.0))))
    alpha = _as_single_channel(out[..., 3])
    alpha = _dilate_mask(alpha, post_px, POST_THICKEN_ITER)
    vis = alpha > 0
    out[..., :3][vis] = 255
    # in-place maximum with guaranteed 2D alpha
    out_alpha = out[..., 3]
    np.maximum(out_alpha, alpha, out=out_alpha)

    # 8) Optional glow/halo (helps on dark UIs)
    if ADD_GLOW:
        out = _apply_glow(out, out[..., 3], blur_px=GLOW_BLUR_PX, glow_alpha=GLOW_ALPHA)

    # 9) Save as PNG (RGBA preserved)
    cv2.imwrite(path_in, out, [int(cv2.IMWRITE_PNG_COMPRESSION), 3])


# ---------------- Instances ----------------
def export_instances(doc, out_dir, target_size: int):
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
            trim_and_to_rgba_white_lines(out_file, target_size)
            exported += 1
            print(f"  ✓ Instance {name} -> {out_file}")
        except Exception as e:
            print(f" X ! Instance {name} failed: {e}")
    print(f"  Instances done: {exported} images")


# ---------- helpers for anchor clustering ----------
def bbox_of_insert(doc, insert):
    fig = plt.figure()
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor((0,0,0,0))
    fig.patch.set_alpha(0.0)
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
def export_auto_clusters(doc, out_dir, target_size: int):
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
            if boxes_overlap(bb, Aexp):
                members.append(inserts[i]); member_idx.append(i); continue
            cx, cy = bbox_center(bb)
            if math.hypot(cx - Acx, cy - Acy) <= center_eps:
                members.append(inserts[i]); member_idx.append(i)

        if len(members) >= AUTOCLUSTER_MIN_SIZE:
            cluster_id += 1
            for i in member_idx:
                assigned.add(i)

            anchor_name = make_safe(getattr(anchor.dxf, "name", "Cluster"))
            out_file = os.path.join(out_dir, f"{anchor_name}__cluster{cluster_id:03d}{OUTPUT_EXT}")
            try:
                render_entities_to_file(doc, members, out_file)
                trim_and_to_rgba_white_lines(out_file, target_size)
                exported += 1
                print(f"  ✓ Cluster {cluster_id} ({anchor_name}) -> {out_file}  [{len(members)} inserts]")
            except Exception as e:
                print(f"  ! Cluster {cluster_id} failed: {e}")

    print(f"  Auto-clusters done: {exported} images (expand={expand:.2f}, center_eps={center_eps:.2f})")


# ---------------- Groups (optional; unchanged) ----------------
def export_groups(doc, out_dir, target_size: int):
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
            trim_and_to_rgba_white_lines(out_file, target_size)
            exported += 1
            print(f"  ✓ Group {gname} -> {out_file}")
        except Exception as e:
            print(f"  ! Group {gname} failed: {e}")
    print(f"  Groups done: {exported} images")


# ---------------- Driver ----------------
def export_one_dxf(dxf_path: str, out_root_for_drawing: str, target_size: int):
    os.makedirs(out_root_for_drawing, exist_ok=True)
    doc = ezdxf.readfile(dxf_path)
    if EXPORT_BLOCK_INSTANCES:
        export_instances(doc, os.path.join(out_root_for_drawing, INSTANCES_DIRNAME), target_size)
    if EXPORT_AUTO_CLUSTERS:
        export_auto_clusters(doc, os.path.join(out_root_for_drawing, AUTOCLUSTERS_DIR), target_size)
    if EXPORT_GROUPS:
        export_groups(doc, os.path.join(out_root_for_drawing, GROUPS_DIRNAME), target_size)


def main():
    ap = argparse.ArgumentParser(description="Export DXF instances/clusters to transparent PNG icons.")
    ap.add_argument("--dxf", default=DXF_FOLDER, help="Folder containing DXF files")
    ap.add_argument("--out", default=OUT_ROOT, help="Output root folder")
    ap.add_argument("--size", type=int, default=TARGET_SIZE, help="Final square icon size in pixels (e.g., 128)")
    args = ap.parse_args()

    dxf_folder = args.dxf
    out_root   = args.out
    size       = max(8, int(args.size))  # guardrail

    print(f"Scanning: {dxf_folder}  |  Output: {out_root}  |  Size: {size}x{size}")
    for name in os.listdir(dxf_folder):
        if not name.lower().endswith(".dxf"):
            continue
        dxf_path = os.path.join(dxf_folder, name)
        draw_name = os.path.splitext(name)[0]
        out_dir   = os.path.join(out_root, draw_name)
        print(f"\nProcessing {name} …")
        export_one_dxf(dxf_path, out_dir, size)

if __name__ == "__main__":
    main()
