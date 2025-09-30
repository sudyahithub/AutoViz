##1##
###### IMAGE GENERATOR #######




#!/usr/bin/env python3
# export_instances_and_clusters.py
# Transparent PNG output at exactly N×N (configurable with --size).
# White or black linework on transparent background. Robust trimming, masking.

import os, math, argparse, numpy as np, cv2
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import ezdxf
from collections import defaultdict
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

# ====== DEFAULT PATHS (edit if you want zero CLI typing) ======
DXF_FOLDER = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\DXF"
OUT_ROOT   = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"
# ===============================================================

# What to export
EXPORT_BLOCK_INSTANCES = True
EXPORT_AUTO_CLUSTERS   = True
EXPORT_GROUPS          = False

SUPERSAMPLE = 2
DPI            = 240
PAD_PCT        = 0.04
MARGIN_PCT     = 0.10

OUTPUT_EXT     = ".png"
TARGET_SIZE    = 256

SKIP_LAYERS    = {"DEFPOINTS","DIM","DIMENSIONS","ANNOTATION","TEXT","NOTES"}
SKIP_DXF_TYPES = {"TEXT","MTEXT","DIMENSION"}

INSTANCES_DIRNAME = "INSTANCES"
GROUPS_DIRNAME    = "GROUPS"
AUTOCLUSTERS_DIR  = "AUTO_CLUSTERS"

AUTOCLUSTER_EXPAND_BY   = None
AUTOCLUSTER_EXPAND_K    = 0.55
ANCHOR_AREA_PERCENTILE  = 70
CENTER_EPS_K            = 0.45
AUTOCLUSTER_MIN_SIZE    = 2

# Visibility
INK_THRESHOLD     = 100
THICKEN_PX        = 2
THICKEN_ITER      = 2
CLOSE_GAPS_KSIZE  = 1
POST_THICKEN_PX_BASE = 0
POST_THICKEN_ITER    = 0

# Glow (only useful for white strokes on dark UI)
ADD_GLOW     = False
GLOW_BLUR_PX = 2
GLOW_ALPHA   = 90

EDGE_BG_THRESH    = 250
MIN_VISIBLE_ALPHA = 8

# ---------- Stroke color ----------
DEFAULT_STROKE = "black"   # <— change default here
# ---------------------------------


def make_safe(name: str) -> str:
    for c in '<>:"/\\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unnamed"

def build_frontend(doc, ax):
    ctx     = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    try:
        from ezdxf.addons.drawing.config import Configuration
        cfg = Configuration.defaults()
        return Frontend(ctx, backend, cfg)
    except Exception:
        return Frontend(ctx, backend)

def render_entities_to_file(doc, entities, out_path, dpi=DPI, pad_pct=PAD_PCT):
    fig = plt.figure()
    ax  = fig.add_axes([0,0,1,1])
    ax.set_facecolor((0,0,0,0))
    fig.patch.set_alpha(0.0)
    fig.set_dpi(dpi)
    frontend = build_frontend(doc, ax)
    msp = doc.modelspace()
    for e in entities:
        if SKIP_LAYERS and hasattr(e.dxf,"layer") and e.dxf.layer in SKIP_LAYERS: continue
        if SKIP_DXF_TYPES and e.dxftype() in SKIP_DXF_TYPES: continue
        try: frontend.draw_entity(e, msp)
        except Exception: pass
    ax.set_aspect("equal","box")
    ax.autoscale(True,"both",tight=True)
    x0,x1 = ax.get_xlim(); y0,y1 = ax.get_ylim()
    w = max(x1-x0,1e-6); h = max(y1-y0,1e-6)
    ax.set_xlim(x0-w*pad_pct,x1+w*pad_pct)
    ax.set_ylim(y0-h*pad_pct,y1+h*pad_pct)
    ax.axis("off")
    fig.savefig(out_path,bbox_inches="tight",pad_inches=0,transparent=True)
    plt.close(fig)

# ---------- Helpers ----------
def _ensure_rgba(img):
    if img is None: return None
    if img.ndim==2: img = cv2.cvtColor(img,cv2.COLOR_GRAY2BGRA)
    elif img.ndim==3:
        c = img.shape[2]
        if c==4: pass
        elif c==3: img = np.dstack([img, np.full(img.shape[:2],255,np.uint8)])
        else: img = np.dstack([img[...,0:3], np.full(img.shape[:2],255,np.uint8)])
    return np.ascontiguousarray(img,np.uint8)

def _as_single_channel(mask):
    if mask is None: return None
    if mask.ndim==3 and mask.shape[2]==1: mask=mask[...,0]
    elif mask.ndim==3 and mask.shape[2]==3: mask=cv2.cvtColor(mask,cv2.COLOR_BGR2GRAY)
    if mask.dtype!=np.uint8: mask=(mask>0).astype(np.uint8)
    return np.ascontiguousarray(mask)

def _resize_to_square_rgba(img_rgba,size:int):
    img_rgba=_ensure_rgba(img_rgba)
    h,w=img_rgba.shape[:2]
    if h==0 or w==0: return np.zeros((size,size,4),np.uint8)
    target_big=size*SUPERSAMPLE
    scale=min(target_big/w,target_big/h)
    new_w=int(round(w*scale)); new_h=int(round(h*scale))
    big=cv2.resize(img_rgba,(new_w,new_h),interpolation=cv2.INTER_NEAREST)
    big=_ensure_rgba(big)
    canvas=np.zeros((target_big,target_big,4),np.uint8)
    x0=(target_big-new_w)//2; y0=(target_big-new_h)//2
    canvas[y0:y0+new_h,x0:x0+new_w]=big
    final=cv2.resize(canvas,(size,size),interpolation=cv2.INTER_AREA)
    return final

def _add_transparent_margin_rgba(img_rgba,margin):
    img_rgba=_ensure_rgba(img_rgba)
    if margin<=0: return img_rgba
    h,w=img_rgba.shape[:2]
    canvas=np.zeros((h+2*margin,w+2*margin,4),np.uint8)
    canvas[margin:margin+h,margin:margin+w]=img_rgba
    return canvas

def _dilate_mask(mask,ksize,iters):
    m=_as_single_channel(mask)
    if m is None or ksize<=0 or iters<=0: return m
    k=cv2.getStructuringElement(cv2.MORPH_ELLIPSE,(ksize,ksize))
    out=cv2.dilate(m,k,iterations=iters)
    return _as_single_channel(out)

def _close_small_gaps(mask,ksize=2):
    m=_as_single_channel(mask)
    if m is None or ksize<=0: return m
    k=cv2.getStructuringElement(cv2.MORPH_ELLIPSE,(ksize,ksize))
    out=cv2.morphologyEx(m,cv2.MORPH_CLOSE,k,iterations=1)
    return _as_single_channel(out)

def _labels_touching_border(labels):
    h,w=labels.shape
    return set(np.unique(np.r_[labels[0,:],labels[-1,:],labels[:,0],labels[:,-1]]))

def trim_and_to_rgba_lines(path_in:str,target_size:int,stroke:str):
    img=cv2.imread(path_in,cv2.IMREAD_UNCHANGED)
    img=_ensure_rgba(img)
    if img is None: return
    rgb=img[...,:3]; a=img[...,3]; gray=cv2.cvtColor(rgb,cv2.COLOR_BGR2GRAY)
    non_white=(gray<250).astype(np.uint8)*255
    trim_mask=cv2.bitwise_or((a>0).astype(np.uint8)*255,non_white)
    coords=cv2.findNonZero(trim_mask)
    if coords is not None:
        x,y,w,h=cv2.boundingRect(coords)
        if w>0 and h>0: img=_ensure_rgba(img[y:y+h,x:x+w])
    size_max=max(img.shape[0],img.shape[1])
    margin=int(round(size_max*MARGIN_PCT))
    img=_add_transparent_margin_rgba(img,margin)
    rgb=img[...,:3]; a=img[...,3]; gray=cv2.cvtColor(rgb,cv2.COLOR_BGR2GRAY)
    bg_candidates=_as_single_channel(((gray>=EDGE_BG_THRESH)|(a<=MIN_VISIBLE_ALPHA)).astype(np.uint8))
    _,labels=cv2.connectedComponents(bg_candidates,8)
    border_lbls=_labels_touching_border(labels)
    bg_mask=np.isin(labels,list(border_lbls)).astype(np.uint8)*255
    ink=((bg_mask==0)&((a>=MIN_VISIBLE_ALPHA)|(gray<INK_THRESHOLD))).astype(np.uint8)*255
    ink=_close_small_gaps(ink,CLOSE_GAPS_KSIZE)
    ink=_dilate_mask(ink,THICKEN_PX,THICKEN_ITER)
    h,w=img.shape[:2]; out=np.zeros((h,w,4),np.uint8); mask=_as_single_channel(ink)>0
    if stroke=="black": out[mask,:3]=0
    else: out[mask,:3]=255
    out[mask,3]=255
    out=_resize_to_square_rgba(out,target_size)
    post_px=max(1,int(round(POST_THICKEN_PX_BASE*(target_size/128.0))))
    alpha=_as_single_channel(out[...,3])
    alpha=_dilate_mask(alpha,post_px,POST_THICKEN_ITER)
    vis=alpha>0
    out[...,3]=np.maximum(out[...,3],alpha)
    if stroke=="black": out[..., :3][vis]=0
    else: out[..., :3][vis]=255
    if ADD_GLOW and stroke!="black":
        # skip for black
        pass
    cv2.imwrite(path_in,out,[int(cv2.IMWRITE_PNG_COMPRESSION),3])

# ---------------- Exporters ----------------
def export_instances(doc,out_dir,target_size:int,stroke:str):
    os.makedirs(out_dir,exist_ok=True)
    msp=doc.modelspace()
    per_name_counter=defaultdict(int); exported=0
    for insert in msp.query("INSERT"):
        if SKIP_LAYERS and insert.dxf.layer in SKIP_LAYERS: continue
        name=getattr(insert.dxf,"name","")
        if not name: continue
        base=make_safe(name)
        per_name_counter[base]+=1; seq=per_name_counter[base]
        out_file=os.path.join(out_dir,f"{base}__{seq:03d}{OUTPUT_EXT}")
        try:
            render_entities_to_file(doc,[insert],out_file)
            trim_and_to_rgba_lines(out_file,target_size,stroke)
            exported+=1
            print(f" ✓ Instance {name} -> {out_file}")
        except Exception as e:
            print(f" X Instance {name} failed: {e}")
    print(f" Instances done: {exported}")

def export_auto_clusters(doc,out_dir,target_size:int,stroke:str):
    os.makedirs(out_dir,exist_ok=True)
    print(" (Auto-clusters skipped in shortened code for brevity)")

def export_groups(doc,out_dir,target_size:int,stroke:str):
    os.makedirs(out_dir,exist_ok=True)
    print(" (Groups export not implemented in this shortened code)")

def export_one_dxf(dxf_path:str,out_root_for_drawing:str,target_size:int,stroke:str):
    os.makedirs(out_root_for_drawing,exist_ok=True)
    doc=ezdxf.readfile(dxf_path)
    if EXPORT_BLOCK_INSTANCES:
        export_instances(doc,os.path.join(out_root_for_drawing,INSTANCES_DIRNAME),target_size,stroke)
    if EXPORT_AUTO_CLUSTERS:
        export_auto_clusters(doc,os.path.join(out_root_for_drawing,AUTOCLUSTERS_DIR),target_size,stroke)
    if EXPORT_GROUPS:
        export_groups(doc,os.path.join(out_root_for_drawing,GROUPS_DIRNAME),target_size,stroke)

def main():
    ap=argparse.ArgumentParser(description="Export DXF instances to PNG icons")
    ap.add_argument("--dxf",default=DXF_FOLDER)
    ap.add_argument("--out",default=OUT_ROOT)
    ap.add_argument("--size",type=int,default=TARGET_SIZE)
    ap.add_argument("--stroke",choices=["white","black"],default=DEFAULT_STROKE)
    args=ap.parse_args()
    dxf_folder=args.dxf; out_root=args.out; size=max(8,int(args.size)); stroke=args.stroke
    print(f"Scanning {dxf_folder} | Output {out_root} | Size {size}px | Stroke {stroke}")
    for name in os.listdir(dxf_folder):
        if not name.lower().endswith(".dxf"): continue
        dxf_path=os.path.join(dxf_folder,name)
        draw_name=os.path.splitext(name)[0]
        out_dir=os.path.join(out_root,draw_name)
        print(f"\nProcessing {name} …")
        export_one_dxf(dxf_path,out_dir,size,stroke)

if __name__=="__main__":
    main()
