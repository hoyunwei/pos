#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
幫「還沒有照片」的商品產生佔位圖，放進 商品圖片/。

用法（在「文博會POS」資料夾裡執行）：
    python3 產生佔位圖.py            # 只補缺圖的商品
    python3 產生佔位圖.py --all      # 全部重產（會蓋掉現有真照片，慎用）

佔位圖長相：淺色底 + 白框，上面小字寫分類，中間大字寫品名，
下面寫貨號和價格。之後拍到真照片，直接覆蓋同貨號開頭的檔案即可。
"""
import json, os, sys, re
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(HERE, "商品匯入_文博.json")
IMG_DIR = os.path.join(HERE, "商品圖片")
SIZE = 600
FONT_PATH = "/System/Library/Fonts/PingFang.ttc"

# 分類底色（水晶貼、貼紙沿用原本那兩張的顏色）
CAT_BG = {
    "水晶貼": (206, 233, 226),
    "貼紙":   (240, 224, 190),
    "海報":   (222, 226, 240),
    "冰箱貼": (240, 219, 219),
    "明信片": (233, 228, 214),
    "凸版":   (215, 231, 219),
    "造型卡": (240, 228, 236),
    "書籤":   (216, 229, 236),
    "L夾":    (232, 232, 224),
    "帆布袋": (226, 222, 213),
}
DEFAULT_BG = (230, 230, 230)
INK = (51, 51, 51)          # 品名
SUB = (85, 85, 85)          # 分類
CODE = (128, 128, 128)      # 貨號
PRICE = (168, 82, 64)       # 價格


def font(size, index=2):
    return ImageFont.truetype(FONT_PATH, size, index=index)


def text_w(draw, s, f):
    return draw.textbbox((0, 0), s, font=f)[2]


def wrap(draw, s, f, max_w):
    """逐字斷行，避免把（）拆在行尾行首。"""
    lines, cur = [], ""
    for ch in s:
        trial = cur + ch
        if cur and text_w(draw, trial, f) > max_w:
            # 收尾字元不要留在行首
            if ch in "）)】」":
                cur = trial
                continue
            # 開頭字元不要留在行尾
            if cur[-1] in "（(【「":
                cur, trial = cur[:-1], cur[-1] + ch
                lines.append(cur)
                cur = trial
                continue
            lines.append(cur)
            cur = ch
        else:
            cur = trial
    if cur:
        lines.append(cur)
    return lines


def make(pid, name, category, price, out_path):
    bg = CAT_BG.get(category, DEFAULT_BG)
    im = Image.new("RGB", (SIZE, SIZE), bg)
    d = ImageDraw.Draw(im)

    # 白色內框
    d.rectangle([18, 18, SIZE - 19, SIZE - 19], outline=(255, 255, 255), width=5)

    # 分類（上方小字）
    f_cat = font(26)
    d.text((SIZE / 2, 68), category, font=f_cat, fill=SUB, anchor="mm")

    # 品名（中間大字，最多 3 行，放不下就縮字級）
    for size in (72, 64, 56, 48, 42, 36, 30):
        f = font(size)
        lines = wrap(d, name, f, SIZE - 130)
        if len(lines) <= 3:
            break
    lh = size * 1.32
    y = SIZE / 2 - (len(lines) - 1) * lh / 2
    for ln in lines:
        d.text((SIZE / 2, y), ln, font=f, fill=INK, anchor="mm")
        y += lh

    # 貨號 + 價格
    f_small = font(26)
    d.text((SIZE / 2, 520), pid, font=f_small, fill=CODE, anchor="mm")
    d.text((SIZE / 2, 558), f"${price}" if price else "未定價",
           font=f_small, fill=PRICE, anchor="mm")

    im.save(out_path, "JPEG", quality=88, optimize=True)


def main():
    force = "--all" in sys.argv
    data = json.load(open(JSON_PATH, encoding="utf-8"))
    files = os.listdir(IMG_DIR)

    made, skipped = [], 0
    for p in data["products"]:
        pid = p["id"]
        # 檔名以「貨號 + _ 」開頭，或整個檔名就是貨號，才算有圖
        has = any(re.match(re.escape(pid) + r"(_|\.)", f) for f in files)
        if has and not force:
            skipped += 1
            continue
        safe = re.sub(r'[/\\:]', '_', p["name"])
        out = os.path.join(IMG_DIR, f"{pid}_{safe}.jpg")
        make(pid, p["name"], p["category"], p.get("price"), out)
        made.append(os.path.basename(out))

    print(f"✅ 產出 {len(made)} 張佔位圖（已有照片的 {skipped} 款跳過）")
    for m in made:
        print("   +", m)
    if made:
        print("\n接著跑： python3 更新商品圖片.py")


if __name__ == "__main__":
    main()
