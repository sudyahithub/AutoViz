# ai_prefilter.py
# Fast deep-embedding prefilter for CAD linework similarity.
# - resnet18 backbone (lightweight)
# - batched embedding (low RAM, faster on CPU)
# - optional mirror invariance
# - on-disk caching of candidate embeddings

from typing import List, Dict, Tuple
import os, hashlib, json
import numpy as np
from PIL import Image
import torch
import torch.nn as nn
import torchvision.models as models
import torchvision.transforms as T

try:
    import faiss  # optional
    HAVE_FAISS = True
except Exception:
    HAVE_FAISS = False

_IMAGENET = {"mean": [0.485, 0.456, 0.406], "std": [0.229, 0.224, 0.225]}
_to_tensor = T.Compose([
    T.Resize(224, interpolation=T.InterpolationMode.BILINEAR),
    T.CenterCrop(224),
    T.ToTensor(),
    T.Normalize(_IMAGENET["mean"], _IMAGENET["std"]),
])

def npgray_to_tensor(img_gray: np.ndarray) -> torch.Tensor:
    pil = Image.fromarray(img_gray).convert("RGB")
    return _to_tensor(pil)

def rot_variants(img: np.ndarray) -> List[np.ndarray]:
    cur = img
    vs = []
    for _ in range(4):
        vs.append(cur)
        cur = np.rot90(cur)
    return vs  # 4 rotations

def rotflip_variants(img: np.ndarray) -> List[np.ndarray]:
    vs = []
    for flip in (False, True):
        base = np.fliplr(img) if flip else img
        cur = base
        for _ in range(4):
            vs.append(cur)
            cur = np.rot90(cur)
    return vs  # 8 rotations x mirror

class DeepEmbedder:
    def __init__(self, device: str = None, backbone: str = "resnet18"):
        torch.set_grad_enabled(False)
        try:
            torch.set_num_threads(max(1, min(4, os.cpu_count() or 1)))
        except Exception:
            pass

        self.device = device or ("cuda" if torch.cuda.is_available() else "cpu")

        if backbone == "resnet18":
            m = models.resnet18(weights=models.ResNet18_Weights.IMAGENET1K_V1)
            feat_dim = 512
        elif backbone == "mobilenet_v3_small":
            m = models.mobilenet_v3_small(weights=models.MobileNet_V3_Small_Weights.IMAGENET1K_V1)
            m = nn.Sequential(m.features, nn.AdaptiveAvgPool2d(1))
            feat_dim = 576
        else:  # fallback
            m = models.resnet18(weights=models.ResNet18_Weights.IMAGENET1K_V1)
            feat_dim = 512

        # cut classifier, keep global pooled features
        if isinstance(m, models.ResNet):
            self.backbone = nn.Sequential(*(list(m.children())[:-1]))
        else:
            self.backbone = m

        self.backbone.eval().to(self.device)
        self.feat_dim = feat_dim

    @torch.inference_mode()
    def embed_batch(self, imgs: List[np.ndarray]) -> np.ndarray:
        if not imgs:
            return np.zeros((0, self.feat_dim), dtype="float32")
        batch = torch.stack([npgray_to_tensor(x) for x in imgs], dim=0).to(self.device)
        feats = self.backbone(batch).squeeze(-1).squeeze(-1)
        feats = torch.nn.functional.normalize(feats, p=2, dim=1)
        return feats.cpu().numpy().astype("float32")

def _cache_key(paths: List[str], use_mirror: bool, backbone: str) -> str:
    meta = []
    for p in paths:
        try:
            meta.append((p, os.path.getmtime(p), os.path.getsize(p)))
        except Exception:
            meta.append((p, 0, 0))
    blob = json.dumps({"meta": meta, "mirror": use_mirror, "bb": backbone}, separators=(",", ":"))
    return hashlib.md5(blob.encode("utf-8")).hexdigest()

def build_deep_index(
    cand_idx: List[Dict],
    embedder: DeepEmbedder,
    use_mirror: bool = False,
    batch_size: int = 64,
    cache_dir: str = None,
    cache_tag: str = ""
):
    """
    Returns dict:
      'emb':   (N*K, D) float32, L2-normalized
      'owner': (N*K,) int32, row -> candidate index
      'paths': list of candidate paths
      'k_orients': 4 or 8
      'faiss': optional index
    Uses caching if cache_dir is provided.
    """
    paths = [c["path"] for c in cand_idx]
    k_orients = 8 if use_mirror else 4

    cache = None
    if cache_dir:
        os.makedirs(cache_dir, exist_ok=True)
        key = _cache_key(paths, use_mirror, "resnet18") + ("_" + cache_tag if cache_tag else "")
        cache = os.path.join(cache_dir, f"deepidx_{key}.npz")
        if os.path.exists(cache):
            data = np.load(cache, allow_pickle=False)
            emb = data["emb"]; owner = data["owner"]; paths_saved = list(data["paths"])
            if len(paths_saved) == len(paths) and paths_saved == paths:
                out = {"emb": emb, "owner": owner, "paths": paths, "k_orients": int(k_orients)}
                if HAVE_FAISS and emb.shape[0] > 0:
                    idx = faiss.IndexFlatIP(emb.shape[1]); idx.add(emb)
                    out["faiss"] = idx
                else:
                    out["faiss"] = None
                return out  # cache hit

    # Build fresh
    emb_list: List[np.ndarray] = []
    owner_list: List[int] = []
    imgs_chunk: List[np.ndarray] = []
    owners_chunk: List[int] = []

    def flush_chunk():
        nonlocal emb_list, imgs_chunk, owners_chunk, owner_list
        if not imgs_chunk: return
        em = embedder.embed_batch(imgs_chunk)
        emb_list.append(em)
        owner_list.extend(owners_chunk)
        imgs_chunk.clear(); owners_chunk.clear()

    for i, c in enumerate(cand_idx):
        vs = rotflip_variants(c["img"]) if use_mirror else rot_variants(c["img"])
        for v in vs:
            imgs_chunk.append(v)
            owners_chunk.append(i)
            if len(imgs_chunk) >= batch_size:
                flush_chunk()
        # flush occasionally to keep UI responsive
        if (i+1) % 200 == 0:
            flush_chunk()

    flush_chunk()

    emb = np.vstack(emb_list) if emb_list else np.zeros((0, embedder.feat_dim), dtype="float32")
    owner = np.array(owner_list, dtype=np.int32)
    out = {"emb": emb, "owner": owner, "paths": paths, "k_orients": int(k_orients)}

    if cache:
        try:
            np.savez_compressed(cache, emb=emb, owner=owner, paths=np.array(paths, dtype=object))
        except Exception:
            pass

    if HAVE_FAISS and emb.shape[0] > 0:
        index = faiss.IndexFlatIP(emb.shape[1]); index.add(emb)
        out["faiss"] = index
    else:
        out["faiss"] = None
    return out

def deep_prefilter(
    qimg: np.ndarray,
    embedder: DeepEmbedder,
    deep_idx,
    top_m: int = 80,
):
    """Return candidate indices (into cand_idx) ranked by cosine similarity."""
    k = deep_idx.get("k_orients", 4)
    qvars = rotflip_variants(qimg) if k == 8 else rot_variants(qimg)
    qemb = embedder.embed_batch(qvars)                   # (k,D)
    E = deep_idx["emb"]                                  # (N*k,D)
    owner = deep_idx["owner"]

    if deep_idx["faiss"] is not None:
        owners = []
        K = min(max(top_m, 50), E.shape[0])
        for i in range(qemb.shape[0]):
            D, I = deep_idx["faiss"].search(qemb[i:i+1], K)
            owners.extend(owner[I[0]])
        uniq = list(dict.fromkeys(owners))
        sims = qemb @ E.T
        max_per_owner = {}
        for o in uniq:
            mask = (owner == o)
            max_per_owner[o] = float(np.max(sims[:, mask]))
        ranked = sorted(max_per_owner.items(), key=lambda kv: -kv[1])
        return [i for (i, _) in ranked[:top_m]]

    sims = qemb @ E.T
    max_per_owner = {}
    for o in range(owner.max()+1):
        mask = (owner == o)
        if not np.any(mask): continue
        max_per_owner[o] = float(np.max(sims[:, mask]))
    ranked = sorted(max_per_owner.items(), key=lambda kv: -kv[1])
    return [i for (i, _) in ranked[:top_m]]
