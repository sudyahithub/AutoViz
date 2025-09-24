#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Local → Sheets Preview linker (FULL)
- Scans LOCAL folder for images (png/jpg/jpeg/webp).
- Matches file basename to "BOQ name" in Sheet (same normalization as Apps Script).
- Sends images as base64 to your Web App (op=previewByName).
"""

import os, sys, base64
from pathlib import Path
import argparse
import requests

CONFIG = {
    "WEBAPP_URL": "https://script.google.com/macros/s/AKfycbwTTg9SzLo70ICTbpr2a5zNw84CG6kylNulVONenq4BADQIuCq7GuJqtDq7H_QfV0pe/exec",
    "SHEET_ID":   "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
    "TAB":        "Fts",
    "DRIVE_FOLDER_ID": "",  # optional; leave "" to let Web App create/use "DXF-Previews"
    "LOCAL_DIR":  r"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS\dem2\INSTANCES",
    "BATCH":      100,
    "TIMEOUT":    300,
}

IMG_EXTS = {".png", ".jpg", ".jpeg", ".webp"}

def norm_key(s: str) -> str:
    """Letters+digits only (must match Apps Script normKey_)."""
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())

def collect_items_from_local(local_dir: Path):
    items = []
    for root, _, files in os.walk(local_dir):
        for fn in files:
            ext = Path(fn).suffix.lower()
            if ext not in IMG_EXTS:
                continue
            name_no_ext = Path(fn).stem
            full = Path(root) / fn
            try:
                with open(full, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("ascii")
                items.append({"name": name_no_ext, "imageB64": b64})
            except Exception as ex:
                print(f"[WARN] Could not read {full}: {ex}", file=sys.stderr)
    return items

def send_batch(sess: requests.Session, items_batch):
    payload = {
        "op": "previewByName",
        "sheetId": CONFIG["SHEET_ID"],
        "tab": CONFIG["TAB"],
        "driveFolderId": CONFIG["DRIVE_FOLDER_ID"],
        "items": items_batch,
    }
    r = sess.post(CONFIG["WEBAPP_URL"], json=payload, timeout=CONFIG["TIMEOUT"])
    if not r.ok:
        raise RuntimeError(f"WebApp HTTP {r.status_code}: {r.text}")
    return r.json()

def main():
    ap = argparse.ArgumentParser(description="Upload local preview images to Google Sheet by BOQ name")
    ap.add_argument("--dir", default=CONFIG["LOCAL_DIR"], help="Local images directory")
    ap.add_argument("--batch", type=int, default=CONFIG["BATCH"])
    args = ap.parse_args()

    local_dir = Path(args.dir)
    if not local_dir.exists():
        print(f"[ERROR] Local folder not found: {local_dir}", file=sys.stderr)
        sys.exit(1)

    items = collect_items_from_local(local_dir)
    if not items:
        print("[INFO] No images found to upload.")
        return

    print(f"[INFO] Found {len(items)} candidate images. Uploading in batches of {args.batch}...")
    sess = requests.Session()
    total_matched = 0
    total_wrote = 0

    # (Optional) quick preview of first few normalized keys
    sample = items[:5]
    if sample:
        print("[DEBUG] Example normalized keys:")
        for it in sample:
            print(f"  '{it['name']}' -> '{norm_key(it['name'])}'")

    for i in range(0, len(items), args.batch):
        batch = items[i:i+args.batch]
        try:
            res = send_batch(sess, batch)
            matched = int(res.get("matched", 0))
            wrote   = int(res.get("wrote", 0))
            total_matched += matched
            total_wrote   += wrote
            print(f"Batch {i//args.batch + 1}: matched {matched}, wrote {wrote}")
        except Exception as ex:
            print(f"[ERROR] Batch {i//args.batch + 1} failed: {ex}", file=sys.stderr)

    print(f"Done. Total matched names: {total_matched}, previews written: {total_wrote}")

if __name__ == "__main__":
    main()
