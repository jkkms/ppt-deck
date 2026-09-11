#!/usr/bin/env python3
"""팔레트 비교 렌더러 — "말로 고르지 말고 보고 고른다".

같은 아웃라인을 팔레트 3종으로 각각 빌드해 대표 슬라이드만 PNG로 뽑는다.
사용:  preview.py outline.yaml [--slides 1,5] [--palettes letterpress,blueprint,graphite]
"""
from __future__ import annotations
import argparse, os, shutil, sys, tempfile

import yaml
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import load_spec
import build as B
import render as R


def main(outline, slides, palettes):
    spec = load_spec()
    o = yaml.safe_load(open(outline, encoding="utf-8"))
    pals = palettes or list(spec["palettes"])
    outdir = os.path.join(os.path.dirname(os.path.abspath(outline)), "preview")
    os.makedirs(outdir, exist_ok=True)
    made = []
    for name in pals:
        with tempfile.TemporaryDirectory() as td:
            o.setdefault("meta", {})["palette"] = name
            tmp_outline = os.path.join(td, "o.yaml")
            yaml.safe_dump(o, open(tmp_outline, "w", encoding="utf-8"), allow_unicode=True)
            B.WARN.clear()
            pptx = B.build(tmp_outline, os.path.join(td, f"{name}.pptx"))
            pdf = R.to_pdf(pptx, td)
            import pymupdf
            doc = pymupdf.open(pdf)
            for i in slides:
                if i <= len(doc):
                    png = os.path.join(outdir, f"palette-{name}-{i:02d}.png")
                    doc[i - 1].get_pixmap(dpi=100).save(png)
                    made.append(png)
            doc.close()
    print(f"\n✓ 팔레트 비교 {len(made)}장 -> {outdir}")
    for p in made: print("  " + p)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("outline")
    ap.add_argument("--slides", default="1,5")
    ap.add_argument("--palettes", default="")
    a = ap.parse_args()
    main(a.outline, [int(x) for x in a.slides.split(",") if x.strip()],
         [p for p in a.palettes.split(",") if p.strip()])
