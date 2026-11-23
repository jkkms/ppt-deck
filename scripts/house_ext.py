"""house 시리즈 확장 1 — 부제 · 위 쪽 번호 · 서명 · 결론 띠 · rows · columns (house 프로파일).

처음엔 면담·면접 강의자료 00 덱 두 개(2026-10-17)의 강의 폴더 `ext.py` 였다. 같은 확장을 열세 덱에서 쓰게 되어
2026-11-23 스킬로 올렸다(「덱 전용 레이아웃 얹기」 — 다른 덱에서 또 쓰면 스킬로). 판 카드는 house_panels.py.
빌더(build.py)를 불러와 LAYOUTS 에 더하고 그대로 돌린다. 크기·색·좌표는 전부 house 스펙과 그리드에서 나온다.

이 시리즈에만 있는 것
- subtitle: 제목 아래 작은 회색 한 줄. 사용자 덱은 모든 장에 부제가 있다(chain·panel_list 의 lead 와 같은 자리·크기).
  파트 표지에서는 짧은 선 아래에 lead 크기로.
- 쪽 번호: 이 시리즈 실측대로 오른쪽 위 눈썹 줄(Dive 는 오른쪽 아래). 표지 제외. meta.page_numbers_top: true
- rows: 줄마다 「제목 + 두 문장 설명」. 배지에는 번호(순서일 때만 numbered) 또는 선 아이콘(icon).
  emphasis 로 한 줄만 들어올린다(옅은 판). stack: true 면 제목 아래에 설명.
- columns: 같은 내용을 세로 칸으로. 제목 h2(card_grid 는 h1 이라 문장 제목이 세 줄로 접힌다).
  emphasis 한 칸만 진초록 면 + 흰 글씨.
- band / band_top / foot: 결론 띠는 진초록 면 + 흰 글씨(2026-10-02 기출 덱 결론 띠와 같다).
  house 팔레트에서 badge_row 의 띠(dark 면 + highlight 글씨)는 161C18 위 1F5D42 라 대비가 약하다.
- cover.signature: 표지 오른쪽 아래 사용자 서명.

사용:  python scripts/house_ext.py outline.yaml -o out/deck.pptx   (판 카드까지 쓰면 house_panels.py)
"""
from __future__ import annotations

import argparse
import os
import sys

SKILL = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SKILL)
import build as B  # noqa: E402
from deckkit import col_x, span_w, content_r, advance  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
# 레이아웃마다 제목 폭(g.head 의 span). 부제는 제목이 실제로 끝난 자리 아래에 둔다
TITLE_SPAN = {"two_col": 7, "data": 7, "statement": 9, "cards": 9, "table": 9}


def _lines(d, text, w, style):
    """줄 수 — 어절 단위로 접는다. 파워포인트·키노트는 한글을 띄어쓰기에서 끊는데
    deckkit.line_count 는 글자 단위라 좁은 칸에서 한 줄씩 모자랐다(제목이 설명을 덮음, 10/17)."""
    if isinstance(text, list):
        text = "\n".join(str(t) for t in text)
    st = d.spec["styles"][style]
    font = d.spec["fonts"][st["font"]]
    adv = lambda t: sum(advance(ch, st["size"], st["tracking"], font) for ch in t)
    sp, total = adv(" "), 0
    w = w - max(4.0, w * 0.03)                 # 여유 — 153.2pt 제목이 153.5pt 칸에서 실제로는 접혔다
    for para in str(text).split("\n"):
        used, lines = 0.0, 1
        for word in para.split(" "):
            ww = adv(word)
            if used and used + sp + ww > w:
                lines += 1
                used = ww
            else:
                used += (sp if used else 0) + ww
            while used > w:                    # 한 어절이 칸보다 길면 글자로 접힌다
                lines += 1
                used -= w
        total += lines
    return total


def _h(d, text, w, style):
    """높이. 설명을 줄 목록으로 주면(원래 덱이 일부러 끊어 둔 짧은 구) 한 줄씩 센다."""
    st = d.spec["styles"][style]
    return st["size"] * st["leading"] * _lines(d, text, w, style)


def _balance(d, text, w, style):
    """두세 줄로 접히는 제목은 줄 길이를 고르게 — 마지막 줄에 「것」「가지」만 남지 않게.
    어절 사이 모든 끊는 자리를 보고 가장 긴 줄이 가장 짧아지는 것을 고른다(줄 수는 그대로)."""
    from itertools import combinations
    t = str(text)
    n = _lines(d, t, w, style)
    if "\n" in t or n not in (2, 3):
        return t.split("\n")
    words = t.split(" ")
    if len(words) > 12:
        return [t]
    st = d.spec["styles"][style]
    font = d.spec["fonts"][st["font"]]
    adv = lambda x: sum(advance(ch, st["size"], st["tracking"], font) for ch in x)
    best, score = None, 1e9
    for cut in combinations(range(1, len(words)), n - 1):
        parts = [" ".join(words[a:b]) for a, b in zip((0,) + cut, cut + (len(words),))]
        if any(_lines(d, p_, w, style) > 1 for p_ in parts):
            continue
        sc = max(adv(p_) for p_ in parts) - min(adv(p_) for p_ in parts) * 0.5
        if sc < score:
            best, score = parts, sc
    return best or [t]


def _title_bottom(d, sl, lay):
    """제목이 실제로 끝난 자리. 부제는 그 아래에 둔다(제목이 두 줄이면 함께 내려간다)."""
    g = B.L(d)
    if not sl.get("title"):
        return g.title_y
    w = span_w(TITLE_SPAN.get(lay, 12))
    t = g.balance(sl["title"], w, "h1")
    return g.title_y + _h(d, t, w, "h1")


def band(d, s, x, y, w, text, *, style="lead"):
    """결론 띠 — 진초록 면 + 흰 글씨, 한 줄 높이 = 줄 높이 + 여백(원칙 §4-3)."""
    return B._band(d, s, x, y, w, text, fill="accent", color="ground", style=style)


def band_h(d, style="lead"):
    st = d.spec["styles"][style]
    return round(st["size"] * st["leading"] + d.spec["shapes"]["panel_pad_h"])


def _extras_h(d, sl, W):
    """band_top · band · foot 의 높이 합과 간격."""
    vgap = 24
    top = band_h(d) + vgap if sl.get("band_top") else 0
    bot = (vgap + band_h(d)) if sl.get("band") else 0
    foot = (vgap + _h(d, sl["foot"], W, "body")) if sl.get("foot") else 0
    return top, bot, foot, vgap


def _draw_extras_top(d, s, sl, y, W):
    if sl.get("band_top"):
        h = band(d, s, col_x(1), y, W, sl["band_top"])
        return y + h + 24
    return y


def _draw_extras_bottom(d, s, sl, y, W):
    if sl.get("band"):
        y += 24
        band(d, s, col_x(1), y, W, sl["band"])
        y += band_h(d)
    if sl.get("foot"):
        fh = _h(d, sl["foot"], W, "body")
        d.text(s, "body", col_x(1), y + 24, W, fh, sl["foot"], color="muted", tag="foot")


def _badge(d, s, it, i, cx, cy, sl, lift=False):
    """번호(순서일 때만) · × 같은 글리프 · 선 아이콘. 들어올린 칸(진초록 면)과 어두운 장에서는
    밝은 원 + 진초록 기호 — 진초록 원은 진초록 면·161C18 바탕 위에서 묻힌다."""
    SH = d.spec["shapes"]
    bd = SH["badge_d"]
    inv = getattr(d, "field_inverted", False)
    if lift or inv:
        circ, ink = ("figure" if inv else "ground"), "accent"
    else:
        circ, ink = "accent", "ground"
    if sl.get("numbered") or it.get("mark"):
        g = str(it.get("index", i + 1)) if sl.get("numbered") else it["mark"]
        d.badge(s, cx, cy, bd, circ, glyph=g, glyph_color=ink, style="small")
    else:
        d.badge(s, cx, cy, bd, circ)
        d.icon(s, it.get("icon", sl.get("icon", "check")), cx, cy, bd * SH["icon_ratio"], ink)


# ================================================================ rows
def rows(d, s, m, sl):
    """줄마다 배지 + 제목(왼쪽 4칸) + 설명(오른쪽 6칸). 배지와 글을 같은 왼쪽 안여백으로 넣어
    emphasis 한 줄에 옅은 판을 깔아도 기준선이 흔들리지 않는다."""
    g = B.L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"rows 는 2~5줄 (받은 값 {n})")
    W = span_w(12)
    pad = SH["panel_pad_x"]
    bd = SH["badge_d"]
    stack = bool(sl.get("stack"))
    show_badge = sl.get("badges", True)
    if not show_badge:                          # 배지 없는 표 같은 목록(분야 · 개념)
        pad = 0
    tx = col_x(2) if not stack else col_x(1) + pad + bd + SH["badge_text_gap"] + 3
    if not show_badge:
        tx = col_x(1)
    tw = span_w(4) if not stack else content_r() - pad - tx
    bx, bw = col_x(7), span_w(6) - pad
    if not show_badge:                          # 표처럼 짧은 이름 + 설명 — 설명을 4칸째로 당긴다
        tw = span_w(3)
        bx, bw = col_x(4), content_r() - col_x(4)
    tstyle = sl.get("title_style", "h2")
    bstyle = sl.get("body_style", "body")
    gv = {2: 34, 3: 28, 4: 20, 5: 14}[n]
    ipad = 16                                   # 들어올린 줄의 판 위아래 안여백

    def row_h(it):
        th = _h(d, _balance(d, it.get("title", ""), tw, tstyle), tw, tstyle)
        if stack:
            return th + ((6 + _h(d, it["body"], tw, bstyle)) if it.get("body") else 0)
        return max(th, _h(d, it["body"], bw, bstyle) if it.get("body") else 0)

    hs = [row_h(it) for it in items]
    emph = sl.get("emphasis")
    top_h, bot_h, foot_h, vgap = _extras_h(d, sl, W)
    total = top_h + sum(hs) + gv * (n - 1) + bot_h + foot_h + (2 * ipad if emph is not None else 0)
    y = B.resolve_y(total, d.spec, top=g.block)
    y = _draw_extras_top(d, s, sl, y, W)
    line1 = d.spec["styles"][tstyle]["size"] * d.spec["styles"][tstyle]["leading"]
    for i, (it, h) in enumerate(zip(items, hs)):
        lift = emph is not None and i == emph
        if lift:
            d.panel(s, col_x(1), y, W, h + 2 * ipad, color="accent_tint")
            y += ipad
        if show_badge:
            _badge(d, s, it, i, col_x(1) + pad + bd / 2, y + line1 / 2, sl)
        tl = _balance(d, it.get("title", ""), tw, tstyle)
        th = _h(d, tl, tw, tstyle)
        d.text(s, tstyle, tx, y, tw, th, tl, tag="row-title")
        if it.get("body"):
            if stack:
                bh_ = _h(d, it["body"], tw, bstyle)
                d.text(s, bstyle, tx, y + th + 6, tw, bh_, it["body"], color="ink2", tag="row-body")
            else:
                bh_ = _h(d, it["body"], bw, bstyle)
                d.text(s, bstyle, bx, y + 2, bw, bh_, it["body"], color="ink2", tag="row-body")
        y += h + (ipad if lift else 0) + gv
    _draw_extras_bottom(d, s, sl, y - gv, W)


# ================================================================ columns
def columns(d, s, m, sl):
    """세로 칸. 배지 → 제목(h2) → 설명. 채움은 emphasis 한 칸에만(진초록 면 + 흰 글씨) —
    모든 칸에 같은 채움을 찍으면 위계가 평평해진다(디자인 원칙 2)."""
    g = B.L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"columns 는 2~5칸 (받은 값 {n})")
    W = span_w(12)
    gap = 28 if n <= 3 else 18
    cw = (W - gap * (n - 1)) / n
    pad = 22
    tw = cw - 2 * pad
    bd = SH["badge_d"]
    show_badge = sl.get("badges", True)
    bstyle = sl.get("body_style", "body" if n <= 3 else "small")

    def blk(it):
        h = (bd + 16) if show_badge else 0
        h += _h(d, _balance(d, it.get("title", ""), tw, "h2"), tw, "h2")
        if it.get("body"):
            h += 10 + _h(d, it["body"], tw, bstyle)
        return h
    ch = max(blk(it) for it in items) + 2 * pad
    top_h, bot_h, foot_h, vgap = _extras_h(d, sl, W)
    y = B.resolve_y(top_h + ch + bot_h + foot_h, d.spec, top=g.block)
    y = _draw_extras_top(d, s, sl, y, W)
    emph = sl.get("emphasis")
    for i, it in enumerate(items):
        x = col_x(1) + i * (cw + gap)
        lift = emph is not None and i == emph
        if lift:
            # 판은 그 칸 내용에 맞춘다(가장 긴 칸에 맞추면 아래가 빈다). 키노트는 본문 줄 간격을 계산보다
            # 조금 넓게 그려 잉크가 몇 pt 내려온다 — 실제 렌더 실측 감사로 잰 값을 lift_pad 로 더한다(10/17)
            d.panel(s, x, y, cw, blk(it) + 2 * pad + float(sl.get("lift_pad", 0)), color="accent")
        cy = y + pad
        if show_badge:
            _badge(d, s, it, i, x + pad + bd / 2, cy + bd / 2, sl, lift=lift)
            cy += bd + 16
        tl = _balance(d, it.get("title", ""), tw, "h2")
        th = _h(d, tl, tw, "h2")
        d.text(s, "h2", x + pad, cy, tw, th, tl,
               color="ground" if lift else "figure", tag="card-title")
        cy += th + 10
        if it.get("body"):
            bh_ = _h(d, it["body"], tw, bstyle)
            d.text(s, bstyle, x + pad, cy, tw, bh_, it["body"],
                   color="ground" if lift else "ink2", tag="card-body")
    _draw_extras_bottom(d, s, sl, y + ch, W)


# ================================================================ 모든 장에 얹는 것
def wrap(lay, fn):
    def run(d, s, m, sl):
        r = fn(d, s, m, sl)
        g = B.L(d)
        SH = d.spec["shapes"]
        sub = sl.get("subtitle")
        if sub and lay == "section" and sl.get("part"):
            th = _h(d, sl.get("title", ""), span_w(9), "cover")
            y = SH["part_title_y"] + th + SH["part_rule_gap"] + 24
            h = _h(d, sub, span_w(9), "lead")
            d.text(s, "lead", col_x(1), y, span_w(9), h, sub, color="muted", tag="part-subtitle")
        elif sub:
            y = max(g.top, _title_bottom(d, sl, lay) + 10)
            h = _h(d, sub, span_w(10), "small")
            d.text(s, "small", col_x(1), y, span_w(10), h, sub, color="muted", tag="subtitle")
        if m.get("page_numbers_top") and lay != "cover" and not sl.get("bare"):
            d.text(s, "micro", content_r() - 60, g.eyebrow_y, 60, 15, str(d._slide_i),
                   align="right", color="faint", tag="page-number")
        if lay == "cover" and sl.get("signature"):
            p = sl["signature"]
            p = p if os.path.isabs(p) else os.path.abspath(p)      # build() 가 아웃라인 폴더 기준으로 이미 바꿔 둔다
            from PIL import Image
            iw, ih = Image.open(p).size
            w = 100.0
            h = w * ih / iw
            d.picture(s, p, content_r() - w, 474, w, h, "contain", tag="signature")
        return r
    return run


def install():
    B.LAYOUTS.update({"rows": rows, "columns": columns})
    for k in list(B.LAYOUTS):
        B.LAYOUTS[k] = wrap(k, B.LAYOUTS[k])


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("outline")
    ap.add_argument("-o", "--out")
    ap.add_argument("--spec")
    a = ap.parse_args(argv)
    install()
    B.build(a.outline, a.out, a.spec)


if __name__ == "__main__":
    main()
