#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prep_image.py — Xử lý ảnh sản phẩm (kể cả ẢNH CHỤP MÀN HÌNH) trước khi chèn báo giá.
"""

import sys
from PIL import Image, ImageChops


def _bg_color(im):
    im = im.convert("RGB")
    w, h = im.size
    pts = [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1),
           (w // 2, 0), (w // 2, h - 1), (0, h // 2), (w - 1, h // 2)]
    from collections import Counter
    c = Counter(im.getpixel(p) for p in pts)
    return c.most_common(1)[0][0]


def autocrop(im, tol=18, pad=6):
    rgb = im.convert("RGB")
    bg = Image.new("RGB", rgb.size, _bg_color(rgb))
    diff = ImageChops.difference(rgb, bg).convert("L")
    mask = diff.point(lambda v: 255 if v > tol else 0)
    bbox = mask.getbbox()
    if not bbox:
        return im
    l, t, r, b = bbox
    l = max(0, l - pad); t = max(0, t - pad)
    r = min(im.size[0], r + pad); b = min(im.size[1], b + pad)
    return im.crop((l, t, r, b))


def whiten_bg(im, tol=28):
    from collections import deque
    rgb = im.convert("RGB")
    w, h = rgb.size
    px = rgb.load()
    bg = _bg_color(rgb)
    def close(a, b):
        return all(abs(a[i] - b[i]) <= tol for i in range(3))
    seen = [[False] * w for _ in range(h)]
    dq = deque()
    for x in range(w):
        for y in (0, h - 1):
            dq.append((x, y))
    for y in range(h):
        for x in (0, w - 1):
            dq.append((x, y))
    WHITE = (255, 255, 255)
    while dq:
        x, y = dq.popleft()
        if x < 0 or y < 0 or x >= w or y >= h or seen[y][x]:
            continue
        seen[y][x] = True
        if close(px[x, y], bg):
            px[x, y] = WHITE
            dq.extend([(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)])
    return rgb


def remove_bg_transparent(im):
    try:
        from rembg import remove
        return remove(im.convert("RGBA"))
    except Exception:
        pass
    from collections import deque
    rgba = im.convert("RGBA")
    w, h = rgba.size
    px = rgba.load()
    bg = _bg_color(im)
    tol = 28
    def close(a, b):
        return all(abs(a[i] - b[i]) <= tol for i in range(3))
    seen = [[False] * w for _ in range(h)]
    dq = deque([(x, y) for x in range(w) for y in (0, h - 1)] +
               [(x, y) for y in range(h) for x in (0, w - 1)])
    while dq:
        x, y = dq.popleft()
        if x < 0 or y < 0 or x >= w or y >= h or seen[y][x]:
            continue
        seen[y][x] = True
        r, g, b, a = px[x, y]
        if close((r, g, b), bg):
            px[x, y] = (r, g, b, 0)
            dq.extend([(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)])
    return rgba


def prep_image(src, dst, bg="white", do_crop=True, max_side=900):
    """bg: 'white' | 'transparent' | 'keep'. Trả về đường dẫn dst."""
    im = Image.open(src)
    if do_crop:
        im = autocrop(im)
    if bg == "white":
        im = whiten_bg(im)
    elif bg == "transparent":
        im = remove_bg_transparent(im)
        im = autocrop(im) if do_crop else im
    if max(im.size) > max_side:
        r = max_side / max(im.size)
        im = im.resize((int(im.size[0] * r), int(im.size[1] * r)), Image.LANCZOS)
    if dst.lower().endswith((".jpg", ".jpeg")) and im.mode == "RGBA":
        im = im.convert("RGB")
    im.save(dst)
    return dst


if __name__ == "__main__":
    a = sys.argv[1:]
    if len(a) < 2:
        print(__doc__); sys.exit(1)
    src, dst = a[0], a[1]
    bg = "white"
    if "--transparent" in a:
        bg = "transparent"
    if "--no-bg" in a:
        bg = "keep"
    out = prep_image(src, dst, bg=bg)
    print("Đã xử lý ->", out)
