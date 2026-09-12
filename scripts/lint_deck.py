#!/usr/bin/env python3
"""§13 구조 검사 — 빌드된 .pptx 와 그 매니페스트를 본다. 전부 에러(경고 아님).

사용:  lint_deck.py out/deck.pptx
"""
from __future__ import annotations
import argparse, json, os, re, sys, zipfile
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import load_spec, col_x, span_w, contrast, derive, invert_pal, content_r

FAIL = []


def fail(code, msg): FAIL.append(f"[{code}] {msg}")


def main(path):
    spec = load_spec()
    man_p = os.path.splitext(path)[0] + ".manifest.json"
    man = json.load(open(man_p, encoding="utf-8")) if os.path.exists(man_p) else None
    z = zipfile.ZipFile(path)
    slides = sorted((n for n in z.namelist()
                     if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                    key=lambda n: int(re.search(r"\d+", os.path.basename(n)).group()))

    ST = spec["styles"]
    sizes = {s["size"] for s in ST.values()}
    track_of = {s["size"]: s["tracking"] for s in ST.values()}
    COLS = {col_x(n) for n in range(1, 13)}

    # ---- SCALE · TRACK · ALIGN · LINES · NO_SHADOW (XML) --------------------
    for i, sl in enumerate(slides, 1):
        x = z.read(sl).decode()
        for m in re.finditer(r'<a:rPr\b([^>]*)>', x):
            attrs = m.group(1)
            sz = re.search(r'sz="(\d+)"', attrs)
            if not sz:
                continue
            pt = int(sz.group(1)) / 100
            if pt not in sizes:
                fail("SCALE", f"s{i}: 크기 {pt}pt 는 스케일 14종에 없다")
            spc = re.search(r'spc="(-?\d+)"', attrs)
            want = track_of.get(pt)
            got = int(spc.group(1)) if spc else 0
            if want is not None and got != want:
                fail("TRACK", f"s{i}: {pt}pt 의 자간 {got} != 스타일 값 {want}")
        if 'algn="ctr"' in x:
            fail("ALIGN", f"s{i}: 가운데 정렬 단락이 있다")
        if "<a:gradFill" in x or "roundRect" in x:
            fail("HARD", f"s{i}: 그라데이션 또는 둥근 모서리가 있다")
        for sp in re.findall(r"<p:sp>.*?</p:sp>", x, re.S):
            if "<a:prstGeom" in sp and "<a:effectLst/>" not in sp and "<a:effectLst>" in sp:
                fail("NO_SHADOW", f"s{i}: 효과가 걸린 도형이 있다")
            if re.search(r"<a:ln\b(?![^>]*/>)", sp) and "<a:noFill/>" not in sp:
                fail("LINES", f"s{i}: 외곽선이 채워진 도형이 있다")

    # ---- CONTRAST (스펙) ----------------------------------------------------
    c = spec["colors"]
    for name, pal in spec["palettes"].items():
        for label, p in ((name, pal), (name + " 반전", invert_pal(pal))):
            try:
                col = derive(p, spec)
            except AssertionError as e:
                fail("CONTRAST", f"{label}: {e}")
                continue
            rg = contrast(col["muted"], col["ground"])
            lim = c.get("min_ratio_caption", c.get("muted_min_ratio_ground", 4.5))
            if rg < lim:
                fail("CONTRAST", f"{label}: muted vs ground {rg:.2f} < {lim}")
            if "ink2" in col and col["ink2"] != col["figure"]:
                ri = contrast(col["ink2"], col["ground"])
                if ri < c.get("min_ratio_body", 4.5):
                    fail("CONTRAST", f"{label}: ink2 vs ground {ri:.2f}")

    # ---- GRID · EYEBROW_Y · ANCHOR · PLATE (매니페스트) ---------------------
    if man:
        by_slide = defaultdict(list)
        for b in man["boxes"]:
            by_slide[b["slide"]].append(b)
        for n, boxes in sorted(by_slide.items()):
            for b in boxes:
                # image_full 의 글자는 판 내부 좌표라 그리드 밖을 허용한다
                if b["tag"] in ("full-title", "caption") or (b["tag"] == "eyebrow" and b["y"] > 300):
                    continue
                if b["x"] not in COLS and round(b["x"]) not in {round(col_x(n)) + spec["table"]["pad_x"] for n in range(1, 13)} | {round(col_x(n) + span_w(sp) - spec["table"]["pad_x"]) for n in range(1, 13) for sp in (2, 3, 4, 6)} | {58, 72}:
                    fail("GRID", f"s{n}: '{b['tag']}' x={b['x']} 가 컬럼 좌표가 아니다")
                if b["x"] + b["w"] > content_r() + 1e-6:
                    fail("GRID", f"s{n}: '{b['tag']}' 오른변 {b['x']+b['w']:.0f} > {content_r():.0f}")
                # style_usage 상 eyebrow 스타일은 카드 색인·인용 출처에도 쓰인다.
                # 검사 대상은 '눈썹 라벨 역할'(tag == eyebrow)뿐이다.
                # image_full 만 예외 — 눈썹이 판 내부(y 340)에 놓인다 (§8.15).
                ok_y = {(spec.get("rhythm_y") or {}).get("eyebrow_y", 56), 340, 354}
                if b["tag"] == "eyebrow" and b["y"] not in ok_y:
                    fail("EYEBROW_Y", f"s{n}: 눈썹 라벨 y={b['y']} (56 고정)")
            R = spec.get("rhythm_y")
            tops = [b["y"] for b in boxes]
            bottoms = [b["y"] + b["h"] for b in boxes]
            if R:
                # house 프로파일: 눈썹·제목이 고정 행에 있고, 어떤 요소도
                # 상단 마진 위나 본문 하한 아래로 나가지 않으면 된다.
                if min(tops) < R["eyebrow_y"] - 0.5:
                    fail("ANCHOR", f"s{n}: 상단 마진 {R['eyebrow_y']} 위로 나간 요소가 있다")
                if max(bottoms) > R["content_bottom"] + 0.5:
                    fail("ANCHOR", f"s{n}: 본문 하한 {R['content_bottom']} 아래로 "
                                   f"{max(bottoms):.0f} 까지 내려갔다")
            elif 56 not in tops and max(bottoms) < 483.5:
                fail("ANCHOR", f"s{n}: 상단 56 도 하단 484 도 잡히지 않았다 "
                               f"(최하단 {max(bottoms):.0f})")
        for sh in man.get("shapes", []):
            k = sh.get("kind")
            if k == "plate":
                if min(sh["w"], sh["h"]) < spec["plates"]["plate_min_side"]:
                    fail("PLATE", f"s{sh['slide']}: plate 짧은 변 {min(sh['w'], sh['h'])} < 32")
            elif k == "hero_rule":
                if sh["h"] != spec["rules"]["hero_rule_w"]:
                    fail("PLATE", f"s{sh['slide']}: hero_rule 두께 {sh['h']}")
            elif k == "table_rule":
                if sh["h"] != spec["rules"]["table_rule_w"]:
                    fail("PLATE", f"s{sh['slide']}: table_rule 두께 {sh['h']}")
            else:
                fail("PLATE", f"s{sh['slide']}: 허용되지 않은 사각형 '{k}'")
    else:
        fail("PLATE", "매니페스트가 없어 사각형·그리드 검사를 못 했다")

    print(f"ppt-deck lint_deck — {os.path.basename(path)} (슬라이드 {len(slides)}장)")
    for f in FAIL:
        print("  ✗ " + f)
    if not FAIL:
        print("  ✓ §13 검사 전부 통과")
    print(f"\n  FAIL {len(FAIL)}")
    return 1 if FAIL else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("pptx")
    sys.exit(main(ap.parse_args().pptx))
