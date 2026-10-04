#!/usr/bin/env python3
"""진짜 렌더로 재는 상자 정렬 전수 감사 — 어떤 도구로 만든 .pptx 든 잰다.

verify_centering.py 는 빌더 매니페스트를 우리 래스터로 그려 잰다. 이건 PowerPoint·Keynote·
LibreOffice 가 실제로 그린 PDF 를 잰다 — 커닝·줄바꿈·글꼴 대체까지 들어간 결과다.
두 프로젝트(Dive 덱 audit.py, 면담·면접 덱 audit_real_all.py)의 감사를 하나로 합쳤다.

  ① 채움 둥근 사각형(카드·패널·띠) 안 내용의 위/아래 여백 차이
  ② 번호 배지와 옆 글 첫 줄의 세로 중심 (또는 배지가 카드 세로 가운데)

"0개"를 믿기 전에 **잰 상자 수**가 슬라이드의 상자 수와 맞는지 본다 — 폭 3" 넘는 카드만 재던
감사가 4·5칸 줄을 통째로 빼놓고 "벗어난 상자 0개"를 보고한 일이 있다(2026-09-25).

  audit_render.py deck.pptx                        # LibreOffice 로 렌더
  audit_render.py deck.pptx --app keynote          # Keynote (사용자 허락 후)
  audit_render.py deck.pptx --pdf 이미뽑은.pdf      # 렌더 생략
  --tol 0.03   허용 차(인치). 카드 안 제목·설명 위가 0.05" 더 큰 배치(원칙 §4-2)는 --top-extra 0.055
"""
from __future__ import annotations
import argparse, os, sys, tempfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main(pptx, pdf=None, app=None, tol=0.03, top_extra=0.0, dpi=200, cell_label_w=2.6):
    from pptx import Presentation
    from pptx.util import Emu
    from PIL import Image
    import pymupdf
    I = lambda v: Emu(v).inches if v is not None else 0.0
    if not pdf:
        import render
        td = tempfile.mkdtemp()
        pdf = render.to_pdf(pptx, td, app)
    doc = pymupdf.open(pdf)
    prs = Presentation(pptx)
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    def prst(sh):
        g = sh._element.find(".//" + A + "prstGeom")
        return g.get("prst") if g is not None else None

    cards_n = cards_bad = badges_n = badges_bad = skipped = 0
    for si, s in enumerate(prs.slides):
        if si >= len(doc):
            break
        page = doc[si]
        pm = page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72))
        im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
        px_all = im.load()

        def ink_rows(x, y, w, h, thr=40, skip=None):
            x0, y0 = int(x * dpi), int(y * dpi)
            x1, y1 = int((x + w) * dpi), int((y + h) * dpi)
            if x1 - x0 < 6 or y1 - y0 < 6:
                return None, 0
            c = Counter()
            for xx in range(x0 + 2, x1 - 2, 2):
                c[px_all[xx, y0 + 1]] += 1
                c[px_all[xx, y1 - 2]] += 1
            bg = c.most_common(1)[0][0]           # 배경은 가장자리 최빈값 (고정값 가정 금지)
            d = lambda p: abs(p[0] - bg[0]) + abs(p[1] - bg[1]) + abs(p[2] - bg[2])
            rows = []
            for yy in range(y0, y1):
                hit = 0
                for xx in range(x0, x1, 2):
                    if skip and skip(xx / dpi, yy / dpi):
                        continue
                    if d(px_all[xx, yy]) > thr:
                        hit += 1
                        if hit >= 2:
                            rows.append(yy - y0)
                            break
            return rows, y1 - y0

        cards, pics, texts, dots = [], [], [], []
        for sh in s.shapes:
            x, y, w, h = I(sh.left), I(sh.top), I(sh.width), I(sh.height)
            kind = prst(sh)
            try:
                solid = sh.fill.type == 1
            except Exception:
                solid = False
            if kind == "roundRect" and solid and h >= 0.55 and w >= 0.6:
                cards.append((x, y, w, h))
            elif kind == "ellipse" and solid and 0.25 <= w <= 0.8 and abs(w - h) < 0.02:
                dots.append((x, y, w, h))
            elif sh.shape_type is not None and "PICTURE" in str(sh.shape_type):
                pics.append((x, y, w, h))
            if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
                texts.append((x, y, w, h, sh.text_frame.text))

        for (x, y, w, h) in cards:
            # 카드 안에 다른 카드가 통째로 들어 있으면(바깥 판) 안쪽을 따로 재므로 건너뛴다
            inner = [c for c in cards if c != (x, y, w, h) and c[0] >= x - .01 and c[1] >= y - .01
                     and c[0] + c[2] <= x + w + .01 and c[1] + c[3] <= y + h + .01]
            labels = [t for t in texts if t[4].startswith("Colab ") and x <= t[0] <= x + w
                      and y <= t[1] <= y + 0.4]
            skip = (lambda X, Y, L=labels: any(lx - .05 <= X <= lx + lw + .05 and ly - .05 <= Y <= ly + lh + .05
                                               for lx, ly, lw, lh, _ in L)) if labels else None
            ins = 0.06
            rows, ch = ink_rows(x + ins, y + ins, w - 2 * ins, h - 2 * ins, skip=skip)
            if not rows:
                skipped += 1
                continue
            cards_n += 1
            t = rows[0] / dpi + ins
            b = (ch - 1 - rows[-1]) / dpi + ins
            for (qx, qy, qw, qh) in pics + inner:        # 안의 사진·작은 판은 선언 사각형 기준
                if qx >= x - .01 and qy >= y - .01 and qx + qw <= x + w + .01 and qy + qh <= y + h + .01:
                    t, b = min(t, qy - y), min(b, y + h - (qy + qh))
            diff = t - b
            if round(abs(diff), 3) > tol and not (top_extra and abs(diff - top_extra) <= tol):
                cards_bad += 1
                print(f"  ✗ {si + 1:2d}쪽 상자 ({x:.2f},{y:.2f}) {w:.2f}x{h:.2f}  "
                      f"위 {t:.3f}\" 아래 {b:.3f}\"  차 {diff:+.3f}\"")

        for (bx, by, bw, bh) in dots:
            host = [c for c in cards if c[0] <= bx and bx + bw <= c[0] + c[2] and c[1] <= by and by + bh <= c[1] + c[3]]
            # 짝이 되는 글 = 배지 바로 오른쪽(0.8" 안), 같은 카드 안. 옆 카드의 글과 짝짓지 않는다
            cand = [q for q in texts if bx + bw - .01 < q[0] < bx + bw + .8
                    and abs(q[1] + min(q[3], .3) / 2 - (by + bh / 2)) < .3
                    and (not host or host[0][0] <= q[0] <= host[0][0] + host[0][2])]
            # 옆에 짝이 되는 글이 없으면 배지가 글 위에 얹힌 배치다(Dive 위 배지 카드 · 면접 덱 번호 카드).
            # 배지와 글을 한 덩어리로 세우므로 ①의 카드 위/아래 여백 검사가 이미 잰다
            if not cand:
                continue
            brows, _ = ink_rows(bx + .03, by - .02, bw - .06, bh + .04)
            bcen = by - .02 + (brows[0] + brows[-1]) / 2 / dpi if brows else by + bh / 2
            ok = False
            tcen = None
            if cand:
                q = min(cand, key=lambda q: q[0])
                trows, _ = ink_rows(q[0], q[1], min(q[2], 3.0), min(q[3], .32))
                if trows:
                    tcen = q[1] + (trows[0] + trows[-1]) / 2 / dpi
                    ok = abs(bcen - tcen) <= tol
            if host and not ok:
                ok = abs(bcen - (host[0][1] + host[0][3] / 2)) <= tol
            if tcen is None and not host:
                continue
            badges_n += 1
            if not ok:
                badges_bad += 1
                print(f"  ✗ {si + 1:2d}쪽 배지 ({bx:.2f},{by:.2f}) 중심 {bcen:.3f}"
                      + (f" · 글 {tcen:.3f}" if tcen else "") + "  — 글 첫 줄이나 카드 가운데와 어긋난다")

    print(f"\n  상자 {cards_n}개 잼 · 어긋남 {cards_bad}  |  배지 {badges_n}개 잼 · 어긋남 {badges_bad}"
          f"  |  잉크 없어 건너뜀 {skipped}  (허용 {tol}\", {dpi}dpi, 렌더 {app or '파일/LibreOffice'})")
    return 1 if (cards_bad or badges_bad) else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--pdf")
    ap.add_argument("--app", choices=["keynote", "powerpoint"])
    ap.add_argument("--tol", type=float, default=0.03)
    ap.add_argument("--top-extra", type=float, default=0.0)
    ap.add_argument("--dpi", type=int, default=200)
    a = ap.parse_args()
    sys.exit(main(a.pptx, a.pdf, a.app, a.tol, a.top_extra, a.dpi))
