#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把「商品圖片/」資料夾裡的照片，批次壓縮後嵌進 POS 的商品匯入檔。

用法（在這個資料夾裡執行）：
    python3 更新商品圖片.py

規則：
  - 依「貨號」配對。檔名只要以貨號開頭即可，例如 MC2401_平安.jpg、MC2401.png、MC2401_新照片.jpeg 都會被認出來
  - 換照片時，請保留檔名最前面的貨號（後面叫什麼都沒差）
  - 圖片會自動縮到最長邊 400px、轉 JPEG 品質 0.7（避免塞爆 POS 的本機儲存空間）
  - 產出：商品匯入_文博.json（直接拿去 POS 的「商品管理 → 匯入商品 JSON」）
"""
import json, os, base64, io, sys
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(HERE, "商品匯入_文博.json")
IMG_DIR = os.path.join(HERE, "商品圖片")
MAX_SIDE = 400
QUALITY = 70
EXTS = (".jpg", ".jpeg", ".png", ".webp", ".heic", ".HEIC")


def to_data_url(path):
    """開圖 → 轉正 → 縮圖 → JPEG → base64 data URL"""
    img = Image.open(path)
    # 依 EXIF 轉正（手機直拍常見）
    try:
        from PIL import ImageOps
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    if max(w, h) > MAX_SIDE:
        if w >= h:
            img = img.resize((MAX_SIDE, round(h * MAX_SIDE / w)), Image.LANCZOS)
        else:
            img = img.resize((round(w * MAX_SIDE / h), MAX_SIDE), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=QUALITY, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def main():
    if not os.path.isdir(IMG_DIR):
        print("找不到「商品圖片」資料夾"); sys.exit(1)

    data = json.load(open(JSON_PATH, encoding="utf-8"))
    files = [f for f in os.listdir(IMG_DIR) if f.lower().endswith(tuple(e.lower() for e in EXTS))]

    matched, missing, total_kb = 0, [], 0
    for p in data["products"]:
        pid = p["id"]
        # 以貨號開頭配對（優先完全等於貨號的檔名）
        cands = [f for f in files if os.path.splitext(f)[0] == pid] or \
                [f for f in files if f.startswith(pid)]
        if not cands:
            missing.append(f"{pid} {p['name']}")
            p.pop("img", None)
            continue
        url = to_data_url(os.path.join(IMG_DIR, sorted(cands)[0]))
        p["img"] = url
        total_kb += len(url) / 1024
        matched += 1

    json.dump(data, open(JSON_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    print(f"✅ 已嵌入 {matched} / {len(data['products'])} 張圖，總大小約 {total_kb/1024:.1f} MB")
    print(f"   產出：{os.path.basename(JSON_PATH)}")
    if missing:
        print(f"\n⚠️ 這 {len(missing)} 款沒找到對應圖片（檔名要以貨號開頭）：")
        for m in missing:
            print("   -", m)
    if total_kb / 1024 > 4:
        print("\n⚠️ 超過 4MB，POS 的本機儲存可能塞不下。把 MAX_SIDE 調小（例如 300）再跑一次。")


if __name__ == "__main__":
    main()
