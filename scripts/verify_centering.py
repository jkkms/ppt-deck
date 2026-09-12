#!/usr/bin/env python3
"""도형 안 글자의 중앙 정렬 전수 측정.

눈대중이나 한두 장 확인으로는 못 잡는다. 렌더한 픽셀에서 글자 잉크의 bbox 를 찾아
도형 중심과의 어긋남을 pt 단위로 잰다. 통과 기준은 벗어난 도형 0개.

사용:  verify_centering.py out/deck.manifest.json [--tol 0.6] [--scale 4]
"""
from __future__ import annotations
import argparse, json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import svgpreview

KINDS = ("badge", "chip", "node")


def main(mpath, tol, scale):
    from PIL import Image
    man = json.load(open(mpath, encoding="utf-8"))
    shapes = [s for s in man.get("shapes", []) if s.get("kind") in KINDS]
    by_slide = {}
    for s in shapes:
        by_slide.setdefault(s["slide"], []).append(s)

    bad = tot = 0
    for n, shs in sorted(by_slide.items()):
        img = svgpreview.raster(man, n, scale).convert("RGB")
        for sh in shs:
            fill = tuple(int(sh["color"][i:i + 2], 16) for i in (0, 2, 4))
            x0, y0 = int(sh["x"] * scale), int(sh["y"] * scale)
            x1, y1 = int((sh["x"] + sh["w"]) * scale), int((sh["y"] + sh["h"]) * scale)
            crop = img.crop((x0, y0, x1, y1))
            px = crop.load()
            W, H = crop.size
            # 채움색과 충분히 다른 픽셀 = 글자 잉크 (테두리 안쪽만 본다)
            inset = max(2, int(min(W, H) * 0.14))
            xs, ys = [], []
            for yy in range(inset, H - inset):
                for xx in range(inset, W - inset):
                    r, g, b = px[xx, yy]
                    if abs(r - fill[0]) + abs(g - fill[1]) + abs(b - fill[2]) > 90:
                        xs.append(xx); ys.append(yy)
            if not xs:
                continue                      # 글자 없는 마디(node)는 건너뛴다
            tot += 1
            cx_ink = (min(xs) + max(xs) + 1) / 2 / scale
            cy_ink = (min(ys) + max(ys) + 1) / 2 / scale
            dx = cx_ink - sh["w"] / 2
            dy = cy_ink - sh["h"] / 2
            if abs(dx) > tol or abs(dy) > tol:
                bad += 1
                print(f"  ✗ s{n:02d} {sh['kind']:6} ({sh['x']:.0f},{sh['y']:.0f}) "
                      f"{sh['w']:.0f}x{sh['h']:.0f}  좌우 {dx:+.2f}pt  상하 {dy:+.2f}pt")
    print(f"\n  중앙 정렬 {tot - bad}/{tot} 통과 · 벗어난 도형 {bad}개 (허용 {tol}pt)")
    return 1 if bad else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest"); ap.add_argument("--tol", type=float, default=0.6)
    ap.add_argument("--scale", type=float, default=4.0)
    a = ap.parse_args()
    sys.exit(main(a.manifest, a.tol, a.scale))
