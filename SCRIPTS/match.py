# match_projects_buckets_pro.py
# WORK images -> best MASTER match.
# Outputs buckets, CSVs, and two JSONs for AutoCAD (by handle & by "<BlockName>__NNN").
# Includes: stronger cleaning, mirror-aware pHash, chamfer distance,
# and FAST AI prefilter (resnet18, batched, cached). Writes "rot" for AutoCAD.

import os, cv2, json, shutil, re
import numpy as np
import pandas as pd
from pathlib import Path
from typing import Optional, List

from ai_prefilter import DeepEmbedder, build_deep_index, deep_prefilter

# ========= EDIT THESE =========
EXPORTS_ROOT = r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS"
MASTER_PROJ  = "M1"                    # library/master
WORK_PROJ    = "POONAWALA OP 2.1"     # work file
SUBFOLDERS   = ["AUTO_CLUSTERS", "INSTANCES"]
OUT_DIR      = os.path.join(EXPORTS_ROOT, "_MATCH_RESULTS")
TOPK         = 5
SAVE_PREVIEW = True
PHASH_PREFILTER_MAX = 26
AI_TOP_M     = 80
AI_USE_MIRROR = False                  # speed: 4 rotations only (set True for 8 with mirror)
AI_BATCH_SIZE = 64
# =================================

JSON_OUT_HANDLE = os.path.join(OUT_DIR, "work_to_master.json")
JSON_OUT_STEM   = os.path.join(OUT_DIR, "work_to_master_by_stem.json")
JSON_MAP_HANDLE = {}
JSON_MAP_STEM   = {}
IMG_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}

HANDLE_ANY = re.compile(r'H([0-9A-Fa-f]{4,})')
HANDLE_TRAIL_HEX = re.compile(r'([0-9A-Fa-f]{4,})$')
def extract_handle_from_path(p: str) -> Optional[str]:
    stem = Path(p).stem
    m = HANDLE_ANY.search(stem)
    if m: return m.group(1).upper()
    m = HANDLE_TRAIL_HEX.search(stem)
    return m.group(1).upper() if m else None

def ensure_dir(p): os.makedirs(p, exist_ok=True)
def safe_name(s: str) -> str: return re.sub(r'[<>:"/\\|?*]+', "_", s).strip() or "Unnamed"

# ---------- preprocessing ----------
def remove_long_lines(bin_img: np.ndarray) -> np.ndarray:
    edges = cv2.Canny(bin_img, 60, 140)
    h, w = bin_img.shape
    min_len = int(0.58 * min(h, w))
    lines = cv2.HoughLinesP(edges, 1, np.pi/180, threshold=90,
                            minLineLength=min_len, maxLineGap=8)
    out = bin_img.copy()
    if lines is not None:
        for x1, y1, x2, y2 in lines[:,0]:
            cv2.line(out, (x1, y1), (x2, y2), 255, thickness=7)
    return out

def remove_border_touching(bin_img: np.ndarray) -> np.ndarray:
    inv = 255 - bin_img
    num, labels, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
    h, w = bin_img.shape
    out = np.full_like(bin_img, 255)
    for i in range(1, num):
        x,y,ww,hh,area = stats[i]
        on_border = (x==0 or y==0 or x+ww>=w-1 or y+hh>=h-1)
        if on_border and (ww>0.5*w or hh>0.5*h):  # frame-ish
            continue
        mask = (labels==i).astype(np.uint8)*255
        out = cv2.bitwise_and(out, 255-mask)
    return out

def skeletonize(bin_img: np.ndarray) -> np.ndarray:
    img = (255-bin_img)//255
    skel = np.zeros_like(img, np.uint8)
    kernel = cv2.getStructuringElement(cv2.MORPH_CROSS, (3,3))
    tmp = np.zeros_like(img); eroded = np.zeros_like(img)
    while True:
        cv2.erode(img, kernel, eroded)
        cv2.dilate(eroded, kernel, tmp)
        cv2.subtract(img, tmp, tmp)
        cv2.bitwise_or(skel, tmp, skel)
        img[:] = eroded
        if cv2.countNonZero(img) == 0: break
    skel = 255 - (skel*255)
    return skel

def load_and_clean(path: str):
    img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if img is None: return None
    _, th = cv2.threshold(img, 250, 255, cv2.THRESH_BINARY_INV)
    coords = cv2.findNonZero(th)
    if coords is not None:
        x, y, w, h = cv2.boundingRect(coords)
        img = img[y:y+h, x:x+w]
    blur = cv2.GaussianBlur(img, (3, 3), 0)
    binr = cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                 cv2.THRESH_BINARY, 25, 5)
    inv = 255 - binr
    cnts, _ = cv2.findContours(inv, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    crop = binr
    if cnts:
        x, y, w, h = cv2.boundingRect(max(cnts, key=cv2.contourArea))
        crop = binr[y:y+h, x:x+w]
    side = max(int(round(max(crop.shape)/0.82)), 64)
    canvas = np.full((side, side), 255, np.uint8)
    r = min(side/crop.shape[0], side/crop.shape[1])
    resized = cv2.resize(crop, (int(crop.shape[1]*r), int(crop.shape[0]*r)), interpolation=cv2.INTER_AREA)
    H, W = resized.shape
    canvas[(side-H)//2:(side-H)//2+H, (side-W)//2:(side-W)//2+W] = resized
    canvas = remove_long_lines(canvas)
    canvas = remove_border_touching(canvas)
    canvas = skeletonize(canvas)
    return canvas

# ---------- descriptors & scoring ----------
def phash(image: np.ndarray, hash_size=8, highfreq=4):
    img = cv2.resize(image, (hash_size*highfreq, hash_size*highfreq), interpolation=cv2.INTER_AREA).astype(np.float32)
    dct = cv2.dct(img)[:hash_size, :hash_size]
    med = np.median(dct[1:, 1:])
    return (dct > med).astype(np.uint8).flatten()

def hamming(a, b): return int(np.count_nonzero(a != b))

def phash_best_rotflip(qimg: np.ndarray, cimg: np.ndarray):
    qh = phash(qimg)
    best = (10**9, 0, False)
    for flip in (False, True):
        cur0 = cv2.flip(cimg, 1) if flip else cimg
        cur = cur0; rot = 0
        for _ in range(4):
            ch = phash(cur)
            dist = hamming(qh, ch)
            if dist < best[0]:
                best = (int(dist), rot, flip)
            cur = np.rot90(cur); rot = (rot + 90) % 360
    return best

def feats_akaze(img):
    ak = cv2.AKAZE_create()
    kp, des = ak.detectAndCompute(img, None)
    return kp, des

def ratio_match(des1, des2, norm=cv2.NORM_HAMMING, ratio=0.75):
    if des1 is None or des2 is None: return []
    if len(des1) < 2 or len(des2) < 2: return []
    bf = cv2.BFMatcher(norm, crossCheck=False)
    try:
        knn = bf.knnMatch(des1, des2, k=2)
    except cv2.error:
        return []
    good = []
    for neigh in knn:
        if len(neigh) < 2: continue
        a, b = neigh[0], neigh[1]
        if a.distance < ratio * b.distance:
            good.append(a)
    return good

def homography_inliers(kp1, kp2, good, ransac_reproj=3.0):
    if good is None or len(good) < 8: return 0
    pts1 = np.float32([kp1[m.queryIdx].pt for m in good]).reshape(-1,1,2)
    pts2 = np.float32([kp2[m.trainIdx].pt for m in good]).reshape(-1,1,2)
    if len(pts1) < 8 or len(pts2) < 8: return 0
    H, mask = cv2.findHomography(pts1, pts2, cv2.RANSAC, ransac_reproj)
    return int(mask.sum()) if mask is not None else 0

def hu_vec(img):
    edges = cv2.Canny(img, 50, 150)
    m = cv2.moments(edges)
    hu = cv2.HuMoments(m).flatten()
    return np.sign(hu) * np.log10(np.abs(hu)+1e-12)

def aspect_ratio(img):
    h,w = img.shape[:2]
    return w / float(h)

def ar_penalty(ar_q, ar_c):
    return abs(np.log((ar_q+1e-6)/(ar_c+1e-6)))

def chamfer_distance(qbin: np.ndarray, cbin: np.ndarray, size: int = 512) -> float:
    q = cv2.resize(qbin, (size, size), interpolation=cv2.INTER_NEAREST)
    c = cv2.resize(cbin, (size, size), interpolation=cv2.INTER_NEAREST)
    q = (255 - q); c = (255 - c)
    qmask = (q > 0); cmask = (c > 0)
    if not qmask.any() or not cmask.any(): return 999.0
    dt_q = cv2.distanceTransform((~qmask).astype(np.uint8)*255, cv2.DIST_L2, 3)
    dt_c = cv2.distanceTransform((~cmask).astype(np.uint8)*255, cv2.DIST_L2, 3)
    cd1 = float(dt_c[qmask].mean())
    cd2 = float(dt_q[cmask].mean())
    d = 0.5*(cd1+cd2)
    diag = (2**0.5) * size
    return float(d / diag * 100.0)

def composite_score(ph_d, good, inl, hu_d, chamf, ar_pen):
    score  = 1.00 * ph_d
    score += 2.20 * min(chamf, 8.0)
    score += 6.00 * min(ar_pen, 0.8)
    score += 0.60 * min(hu_d, 5.0)
    score -= 0.10 * min(good, 40)
    score -= 0.25 * min(inl,  40)
    return float(score)

def accept_policy(ph_d, good, inl, hu_d):
    if ph_d <= 12 and (good >= 12 or inl >= 10): return True
    if ph_d <= 18 and inl >= 14: return True
    if hu_d <= 0.55 and good >= 16: return True
    return False

# ---------- IO ----------
def list_images(folder: str) -> List[str]:
    return sorted([str(p) for p in Path(folder).rglob("*") if p.suffix.lower() in IMG_EXTS])

def build_index(paths: List[str]):
    idx = []
    for p in paths:
        img = load_and_clean(p)
        if img is None: continue
        idx.append({"path": p, "img": img, "hu": hu_vec(img)})
    return idx

def montage(imgA: np.ndarray, imgB: np.ndarray, w=380) -> np.ndarray:
    a = cv2.resize(imgA, (w, w), interpolation=cv2.INTER_AREA)
    b = cv2.resize(imgB, (w, w), interpolation=cv2.INTER_AREA)
    return cv2.hconcat([a, b])

# ---------- main ----------
def process_subfolder(sub: str):
    work_dir   = os.path.join(EXPORTS_ROOT, WORK_PROJ, sub)
    master_dir = os.path.join(EXPORTS_ROOT, MASTER_PROJ, sub)

    q_paths = list_images(work_dir)
    c_paths = list_images(master_dir)
    print(f"[{sub}] WORK images: {len(q_paths)} | MASTER images: {len(c_paths)}")
    if not q_paths or not c_paths:
        print(f"[{sub}] Skipped (no images).")
        return pd.DataFrame()

    print(f"[{sub}] Indexing MASTER candidates…")
    cand_idx = build_index(c_paths)

    # AI prefilter index (fast, cached)
    print(f"[{sub}] Building AI index…")
    embedder = DeepEmbedder(backbone="resnet18")
    cache_dir = os.path.join(OUT_DIR, "_EMB_CACHE")
    deep_idx = build_deep_index(
        cand_idx, embedder,
        use_mirror=AI_USE_MIRROR,
        batch_size=AI_BATCH_SIZE,
        cache_dir=cache_dir,
        cache_tag=f"{MASTER_PROJ}_{sub}"
    )

    rows = []
    buckets_root = os.path.join(OUT_DIR, f"buckets_{WORK_PROJ}_{sub}")
    ensure_dir(buckets_root)

    for i, qp in enumerate(q_paths, 1):
        qimg = load_and_clean(qp)
        if qimg is None: continue
        qhu  = hu_vec(qimg)
        qkp, qdes = feats_akaze(qimg)
        ar_q = aspect_ratio(qimg)

        keep_ids = deep_prefilter(qimg, embedder, deep_idx, top_m=AI_TOP_M)
        c_subset = [cand_idx[j] for j in keep_ids] if keep_ids else cand_idx

        scores = []
        for c in c_subset:
            ar_c = aspect_ratio(c["img"])
            arp  = ar_penalty(ar_q, ar_c)
            if arp > 0.45:  # AR gate
                continue

            ph_d, rot_deg, flipped = phash_best_rotflip(qimg, c["img"])
            if ph_d > PHASH_PREFILTER_MAX:
                continue

            timg = c["img"]
            if flipped: timg = cv2.flip(timg, 1)
            for _ in range(rot_deg//90):
                timg = np.rot90(timg)

            ckp, cdes = feats_akaze(c["img"])
            good = ratio_match(qdes, cdes)
            inl  = homography_inliers(qkp, ckp, good)
            hud  = float(np.linalg.norm(qhu - c["hu"]))
            chamf = chamfer_distance(qimg, timg)

            score = composite_score(ph_d, len(good), inl, hud, chamf, arp)
            scores.append((score, ph_d, len(good), inl, hud, c["path"], c["img"], rot_deg, flipped, arp, chamf))

        if not scores:
            for c in cand_idx:
                ph_d, rot_deg, flipped = phash_best_rotflip(qimg, c["img"])
                hud  = float(np.linalg.norm(qhu - c["hu"]))
                score = composite_score(ph_d, 0, 0, hud, chamfer_distance(qimg, c["img"]), 0.0)
                scores.append((score, ph_d, 0, 0, hud, c["path"], c["img"], rot_deg, flipped, 0.0, 999.0))

        scores.sort(key=lambda x: x[0])
        bestk = scores[:TOPK]
        best  = bestk[0]

        q_stem = Path(qp).stem
        bucket_dir = os.path.join(buckets_root, safe_name(q_stem))
        ensure_dir(bucket_dir)

        shutil.copy2(qp,      os.path.join(bucket_dir, f"00_query{Path(qp).suffix.lower()}"))
        shutil.copy2(best[5], os.path.join(bucket_dir, f"01_match_from_{safe_name(MASTER_PROJ)}{Path(best[5]).suffix.lower()}"))

        if SAVE_PREVIEW:
            try:
                prev = montage(qimg, best[6])
                cv2.imwrite(os.path.join(bucket_dir, "preview_side_by_side.png"), prev)
            except Exception:
                pass

        with open(os.path.join(bucket_dir, "info.txt"), "w", encoding="utf-8") as f:
            f.write(
                "query: {}\nmatch: {}\nscore: {:.4f}\nphash_dist: {}\norb_good: {}\n"
                "inliers: {}\nhu_dist: {:.4f}\nrot: {}\nflip: {}\nchamfer: {:.3f}\n"
                .format(qp, best[5], best[0], best[1], best[2], best[3], best[4], best[7], best[8], best[10])
            )

        for rank, (score, ph_d, good, inl, hud, cpath, _cimg, rot_deg, flipped, arp, chamf) in enumerate(bestk, 1):
            rows.append({
                "subfolder": sub,
                "bucket": bucket_dir,
                "query_path": qp,
                "cand_path": cpath,
                "rank": rank,
                "score": round(score, 4),
                "phash_dist": ph_d,
                "orb_good": good,
                "inliers": inl,
                "hu_dist": round(hud, 4),
                "rot": int(rot_deg),
                "flip": bool(flipped),
                "ar_penalty": round(arp, 4),
                "chamfer": round(chamf, 4),
                "accepted": accept_policy(ph_d, good, inl, hud)
            })

        def topk_items():
            items = []
            for (score, ph_d, good, inl, hud, cpath, _cimg, rot_deg, flipped, arp, chamf) in bestk:
                stem = Path(cpath).stem
                base = re.sub(r'__\d{2,}$', '', stem)
                items.append({
                    "master_block": base,
                    "score": float(score),
                    "preview": cpath,
                    "rot": int(rot_deg)
                })
            return items

        h = extract_handle_from_path(qp)
        if h:
            tk = topk_items()
            if (h not in JSON_MAP_HANDLE) or (tk and tk[0]["score"] < JSON_MAP_HANDLE[h]["topk"][0]["score"]):
                JSON_MAP_HANDLE[h] = {"work_image": qp, "topk": tk}
        else:
            print(f"[WARN] No HANDLE in filename: {qp}")

        stem_key = Path(qp).stem
        tk2 = topk_items()
        if (stem_key not in JSON_MAP_STEM) or (tk2 and tk2[0]["score"] < JSON_MAP_STEM[stem_key]["topk"][0]["score"]):
            JSON_MAP_STEM[stem_key] = {"work_image": qp, "topk": tk2}

        print(f"[{sub}] [{i}/{len(q_paths)}] {Path(qp).name} -> {Path(best[5]).name} "
              f"(score={best[0]:.2f}, pH={best[1]}, good={best[2]}, inl={best[3]}, rot={best[7]}, flip={best[8]})")

    df = pd.DataFrame(rows)
    ensure_dir(OUT_DIR)
    csv_path = os.path.join(OUT_DIR, f"matches_{WORK_PROJ}_vs_{MASTER_PROJ}_{sub}.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"[{sub}] Saved CSV → {csv_path}")
    return df

def main():
    ensure_dir(OUT_DIR)
    dfs = []
    for sub in SUBFOLDERS:
        dfs.append(process_subfolder(sub))
    dfs = [d for d in dfs if not d.empty]
    if dfs:
        df_all = pd.concat(dfs, ignore_index=True)
        combined_csv = os.path.join(OUT_DIR, f"matches_{WORK_PROJ}_vs_{MASTER_PROJ}_COMBINED.csv")
        df_all.to_csv(combined_csv, index=False, encoding="utf-8-sig")
        print(f"Saved COMBINED CSV → {combined_csv}")

    with open(JSON_OUT_HANDLE, "w", encoding="utf-8") as f:
        json.dump(JSON_MAP_HANDLE, f, ensure_ascii=False, indent=2)
    with open(JSON_OUT_STEM, "w", encoding="utf-8") as f:
        json.dump(JSON_MAP_STEM, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON (handle) → {JSON_OUT_HANDLE}  (entries: {len(JSON_MAP_HANDLE)})")
    print(f"Saved JSON (stem)   → {JSON_OUT_STEM}    (entries: {len(JSON_MAP_STEM)})")

if __name__ == "__main__":
    main()
