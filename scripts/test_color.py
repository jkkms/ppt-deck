#!/usr/bin/env python3
"""색 계산 단위 테스트 — 명세서 §2 토큰 표의 기대값을 고정한다.

이 값이 안 맞으면 구현이 틀린 것이다. 기대값을 고치지 마라.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import load_spec, derive, invert_pal, contrast

EXPECT = {              # 팔레트: (muted, ground 대비, figure 대비, hairline)  §5.2
    "letterpress": ("656464", 5.24, 2.91, "B3B1AD"),
    "graphite":    ("6B6D70", 4.96, 2.96, "B9BABA"),
    "blueprint":   ("ABABA5", 7.12, 1.85, "545B64"),
    "dive":        ("B7BABF", 8.99, 1.95, "565F6B"),
}
GROUND_MIN, FIGURE_MIN = 4.5, 1.7


def main():
    spec = load_spec()
    bad = 0
    for name, (hx, cg, cf, hl) in EXPECT.items():
        pal = spec["palettes"][name]
        col = derive(pal, spec)
        got_g = contrast(col["muted"], col["ground"])
        got_f = contrast(col["muted"], col["figure"])
        ok_hex = col["muted"] == hx and col["hairline"] == hl
        ok_g = abs(got_g - cg) < 0.015
        ok_f = abs(got_f - cf) < 0.015
        gate = got_g >= GROUND_MIN and got_f >= FIGURE_MIN
        mark = "✓" if (ok_hex and ok_g and ok_f and gate) else "✗"
        bad += 0 if mark == "✓" else 1
        print(f"  {mark} {name:12} muted #{col['muted']}(기대 {hx}) "
              f"hairline #{col['hairline']}(기대 {hl}) · "
              f"ground {got_g:.2f}(기대 {cg}) · figure {got_f:.2f}(기대 {cf})"
              f"{'' if gate else '  ← 판정 미달'}")
    # §5.3 반전 필드도 derive를 다시 통과해야 한다
    inv = derive(invert_pal(spec["palettes"]["letterpress"]), spec)
    ok = inv["muted"] == "B3B1AD"
    bad += 0 if ok else 1
    print(f"  {'✓' if ok else '✗'} letterpress 반전 muted #{inv['muted']} (기대 B3B1AD)")
    print(f"\n  실패 {bad} / {len(EXPECT)}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
