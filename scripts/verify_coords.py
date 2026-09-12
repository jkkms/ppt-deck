#!/usr/bin/env python3
"""좌표 대조 — 명세서 §3(좌표 표)의 실값과 빌드 결과를 맞춰본다. 허용 오차 0pt.

사용:  verify_coords.py out/deck.manifest.json
"""
from __future__ import annotations
import argparse, json, os, sys
from collections import defaultdict

# (레이아웃 태그, 순번) -> (x, y, w).  None 은 검사하지 않음.
EXPECT = {
    "cover":      [("eyebrow", 72, 56, 606), ("cover-year", 632, 56, 256),
                   ("cover-title", 72, 252, 606), ("cover-meta", 72, 441, 326)],
    "section":    [("section-num", 72, 56, 466), ("runner", 632, 60, 256),
                   ("section-title", 72, 319, 606)],
    "statement-":  [("eyebrow", 72, 56, 466), ("runner", 632, 60, 256),
                    ("statement", 72, 319, 676)],
    "statement+":  [("eyebrow", 72, 56, 466), ("statement", 72, 120, 676),
                    ("statement-body", 72, 400, 466)],
    "two_col+":   [("eyebrow", 72, 56, 256), ("col-title", 72, 108, 256),
                   ("small", 72, 262, 256), ("bullets", 422, 108, 466),
                   ("h2", 422, 363, 466), ("body", 422, 400, 466)],
    "two_col-":   [("eyebrow", 72, 56, 256), ("col-title", 72, 398, 256),
                   ("bullets", 422, 372, 466)],
    "pair":       [("eyebrow", 72, 56, 466), ("cards-head", 72, 86, 606),
                   ("card-index", 72, 177, 326), ("card-title", 72, 209, 326),
                   ("card-body", 72, 287, 326), ("card-index", 492, 261, 396),
                   ("card-title", 492, 293, 396), ("card-body", 492, 371, 396)],
    "ledger":     [("ledger-index", 72, 172, 70), ("ledger-title", 142, 178, 256),
                   ("ledger-body", 492, 176, 396),
                   ("ledger-index", 72, 281, 70), ("ledger-index", 72, 390, 70)],
    "quad":       [("card-index", 72, 177, 326), ("card-title", 72, 207, 326),
                   ("card-body", 72, 247, 326), ("card-index", 492, 177, 396),
                   ("card-index", 72, 338, 326), ("card-index", 492, 338, 396)],
    "dense":      [("card-index", 72, 177, 256), ("card-index", 352, 177, 256),
                   ("card-index", 632, 177, 256), ("card-index", 72, 338, 256),
                   ("card-title", 72, 207, 256), ("card-body", 72, 277, 256)],
    "data+":      [("eyebrow", 72, 56, 466), ("hero", 72, 140, 396),
                   ("hero-cap", 72, 308, 326),
                   ("sub-value", 562, 140, 326), ("sub-value", 562, 244, 326),
                   ("sub-value", 562, 348, 326), ("sub-label", 562, 194, 326),
                   ("sub-label", 562, 298, 326), ("sub-label", 562, 402, 326)],
    "data-":      [("eyebrow", 72, 56, 466), ("runner", 632, 60, 256),
                   ("hero", 72, 260, 606), ("hero-cap", 72, 428, 466)],
    "quote":      [("quote", 142, 232, 606), ("quote-source", 142, 433, 326),
                   ("small", 142, 461, 326)],
    "table":      [("eyebrow", 72, 56, 466), ("table-head", 72, 86, 606),
                   ("th", 88, 174, 380), ("th", 492, 174, 100),
                   ("td", 88, 214, 380), ("td", 88, 254, 380),
                   ("td", 88, 294, 380), ("td", 88, 334, 380), ("td", 88, 374, 380),
                   ("small", 72, 424, 466)],
    "closing":    [("eyebrow", 72, 56, 326), ("ledger-label", 562, 196, 326),
                   ("ledger-value", 562, 220, 326), ("ledger-label", 562, 266, 326),
                   ("ledger-label", 562, 336, 326), ("closing-title", 72, 402, 466)],
}
SHAPES = {   # (kind, x, y, w, h)
    "data+":  [("hero_rule", 72, 292, 396, 2)],
    "data-":  [("hero_rule", 72, 412, 606, 2)],
    "table":  [("plate", 72, 168, 816, 34),
               ("table_rule", 72, 241.5, 816, 0.75), ("table_rule", 72, 281.5, 816, 0.75),
               ("table_rule", 72, 321.5, 816, 0.75), ("table_rule", 72, 361.5, 816, 0.75)],
}


def main(mpath, mapping):
    man = json.load(open(mpath, encoding="utf-8"))
    boxes, shapes = defaultdict(list), defaultdict(list)
    for b in man["boxes"]:
        boxes[b["slide"]].append(b)
    for sh in man.get("shapes", []):
        shapes[sh["slide"]].append(sh)

    bad = tot = 0
    for slide_no, key in mapping.items():
        got = boxes.get(slide_no, [])
        for tag, ex, ey, ew in EXPECT.get(key, []):
            tot += 1
            hit = [b for b in got if b["tag"] == tag
                   and abs(b["x"] - ex) < 1e-6 and abs(b["y"] - ey) < 1e-6
                   and abs(b["w"] - ew) < 1e-6]
            if not hit:
                bad += 1
                near = [f"{b['tag']}({b['x']:.0f},{b['y']:.0f},{b['w']:.0f})"
                        for b in got if b["tag"] == tag]
                print(f"  ✗ s{slide_no:02d} {key:11} {tag:14} 기대 ({ex},{ey},{ew}) "
                      f"실제 {near or '없음'}")
        for kind, ex, ey, ew, eh in SHAPES.get(key, []):
            tot += 1
            hit = [s for s in got and shapes.get(slide_no, []) if s.get("kind") == kind
                   and abs(s["x"] - ex) < 1e-6 and abs(s["y"] - ey) < 1e-6
                   and abs(s["w"] - ew) < 1e-6 and abs(s["h"] - eh) < 1e-6]
            if not hit:
                bad += 1
                near = [f"{s.get('kind')}({s['x']:.0f},{s['y']:.1f},{s['w']:.0f}x{s['h']})"
                        for s in shapes.get(slide_no, [])]
                print(f"  ✗ s{slide_no:02d} {key:11} {kind:14} 기대 ({ex},{ey},{ew}x{eh}) "
                      f"실제 {near or '없음'}")
    print(f"\n  좌표 대조 {tot - bad}/{tot} 일치 · 불일치 {bad}건 (허용 오차 0pt)")
    return 1 if bad else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest")
    ap.add_argument("--map", default="")
    a = ap.parse_args()
    # 예제 아웃라인의 슬라이드 순서 -> 대조 키
    default = {1: "cover", 2: "section", 3: "statement-", 4: "two_col+", 5: "pair",
               6: "statement+", 7: "ledger", 8: "data+", 9: "quad", 10: "quote",
               11: "dense", 12: "two_col-", 13: "data-", 14: "table", 15: "closing"}
    sys.exit(main(a.manifest, default))
