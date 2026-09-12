#!/usr/bin/env python3
"""ppt-deck 빌더 — outline.yaml -> .pptx

레이아웃 8종. 모든 좌표는 deck-spec.yaml의 그리드에서 파생된다.
사용:  build.py outline.yaml [-o out/deck.pptx] [--spec other-spec.yaml]
"""
from __future__ import annotations
import argparse, os, sys

import yaml
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import Deck, load_spec, est_height, est_lines, est_adv

WARN: list[str] = []


def warn(msg): WARN.append(msg)


# ================================================================= 레이아웃 8종
def _head(d: Deck, s, sl, span=9):
    """눈썹 라벨 + 제목. 라벨은 지금 어느 장(章)인지 알려주는 표지지 장식이 아니다."""
    g = d.g
    y = g.margin_top
    if sl.get("label"):
        d.text(s, "label", g.margin_x, y, g.w(span), 16, sl["label"], color="accent")
        y += 28
    if sl.get("title"):
        SP = d.spec["styles"]["h1"]
        th = est_height([sl["title"]], SP["size"], SP["leading"], g.w(span),
                        font=d.spec["fonts"][SP["font"]])
        d.text(s, "h1", g.margin_x, y, g.w(span), th + 8, sl["title"], tag="head")
        y += th + 8
    return y


def cover(d: Deck, s, m, sl):
    g = d.g
    meta = " · ".join(x for x in [m.get("author"), m.get("org"), m.get("date")] if x)
    if meta:
        d.text(s, "label", g.margin_x, g.margin_top, g.w(6), 16, meta, color="muted")
    d.rule(s, g.margin_x, 176, g.w(3), 4)
    d.text(s, "display", g.margin_x, 200, g.w(9), 232,
           sl.get("title", m.get("title", "")), anchor="bottom", tag="cover-title")
    sub = sl.get("subtitle", m.get("subtitle"))
    if sub:
        d.text(s, "h2", g.margin_x, 448, g.w(8), 36, sub, color="muted")


def section(d: Deck, s, m, sl):
    g = d.g
    d.text(s, "label", g.margin_x, g.margin_top, g.w(4), 16,
           sl.get("label", "SECTION"), color="ground")
    d.text(s, "mega", g.x(7), 40, g.w(5), 190, str(sl.get("number", "")),
           color="figure", align="right", tag="section-num")
    d.text(s, "display", g.margin_x, 250, g.w(8), 190, sl.get("title", ""),
           color="ground", anchor="bottom", tag="section-title")


def statement(d: Deck, s, m, sl):
    g = d.g
    d.text(s, "display", g.margin_x, 154, g.w(10), 272, sl.get("text", ""),
           anchor="middle", tag="statement")
    if sl.get("note"):
        d.text(s, "label", g.margin_x, 440, g.w(8), 36, sl["note"], color="muted")


def two_col(d: Deck, s, m, sl):
    """두 컬럼을 같은 기준선에 세운다. 한쪽만 가운데로 띄우면 강조선과 제목의 관계가 끊긴다.
    내용이 적을 때는 컬럼을 따로 움직이지 않고 블록 전체를 내려 광학 중심을 맞춘다."""
    g = d.g
    SP = d.spec["styles"]
    # 발표용은 본문도 크게 간다. 18pt를 강당 뒤에서 읽을 수 있다는 건 착각이다.
    bstyle, gap = ("h2", 26) if d.density == "speaker" else ("body", 14)
    bw, bs = g.w(7), SP[bstyle]
    lw = g.w(4)

    bullets = [str(b) for b in (sl.get("bullets") or [])]
    lim = d.limits["bullets"]
    if len(bullets) > lim:
        warn(f"slide {d._slide_i}: 불릿 {len(bullets)}개 > {d.density} 한도 {lim}개 — 슬라이드를 쪼개라")
    for b in bullets:                           # 3줄로 접히는 불릿은 불릿이 아니라 문단이다
        n = est_lines("— " + b, bs["size"], bw, font=d.spec["fonts"][bs["font"]])
        if n > 2:
            warn(f"slide {d._slide_i}: 불릿이 {n}줄로 접힌다 — {b[:24]!r}… 명사구로 잘라라")

    right = bullets or ([sl["body"]] if sl.get("body") else [])
    FN = d.spec["fonts"]
    rh = est_height(right, bs["size"], bs["leading"], bw, gap, bool(bullets),
                    font=FN[bs["font"]]) if right else 0
    th = est_height([sl.get("title", "")], SP["h1"]["size"], SP["h1"]["leading"], lw,
                    font=FN[SP["h1"]["font"]])
    lh = est_height([sl["lead"]], SP["h2"]["size"], SP["h2"]["leading"], lw,
                    font=FN[SP["h2"]["font"]]) if sl.get("lead") else 0
    left = th + (lh + 36 if lh else 0)

    avail = g.bottom - 80
    top = 80 + max(0, min(120, (avail - max(rh, left)) / 2.4))
    if top + max(rh, left) > g.bottom:
        warn(f"slide {d._slide_i}: 본문이 하단 마진을 넘는다 — 슬라이드를 쪼개라")

    if sl.get("label"):
        d.text(s, "label", g.margin_x, top - 28, lw, 16, sl["label"], color="accent")
    d.text(s, "h1", g.margin_x, top, lw, th + 8, sl.get("title", ""), tag="col-title")
    if lh:
        d.text(s, "h2", g.margin_x, top + th + 36, lw, lh + 8, sl["lead"], color="muted")
    if right:
        d.text(s, bstyle, g.x(5), top, bw, rh + 8, right,
               accent_lead=bool(bullets), space_after=gap, tag="bullets")


def _fit(d: Deck, text, width, chain, suffix=None, suffix_scale=0.4):
    """한 줄에 들어가는 가장 큰 스케일 단계를 고른다.
    타입스케일을 벗어난 임의 크기를 만들지 않는 선에서만 줄인다 — 스케일 이탈이 AI 티의 본체다."""
    from deckkit import est_adv
    for name in chain:
        st = d.spec["styles"][name]
        fnt = d.spec["fonts"][st["font"]]
        extra = est_adv("\u202f" + str(suffix), fnt) * st["size"] * suffix_scale if suffix else 0.0
        if est_lines(str(text), st["size"], width, extra, fnt) <= 1:
            return name
    return chain[-1]


def data(d: Deck, s, m, sl):
    """숫자 3개를 똑같은 크기로 나란히 놓으면 그건 대시보드 위젯이지 슬라이드가 아니다.
    하나를 히어로로 세우고 나머지를 종속시킨다. 위계가 곧 주장이다."""
    g = d.g
    _head(d, s, sl)
    items = sl.get("items") or []
    if not items:
        return
    hero, subs = items[0], items[1:3]
    hw = g.w(7) if subs else g.w(10)
    hstyle = _fit(d, hero.get("value", ""), hw,
                  ["stat1", "stat2", "stat3", "stat_sub"], hero.get("unit"))
    d.text(s, hstyle, g.margin_x, 150, hw, 190,
           str(hero.get("value", "")), anchor="bottom", suffix=hero.get("unit"),
           tag="stat-hero")
    d.rule(s, g.margin_x, 358, hw, 4)
    if hero.get("caption"):
        d.text(s, "small", g.margin_x, 374, hw, 76, hero["caption"], color="muted")

    x, w = g.x(8), g.w(4)
    for k, it in enumerate(subs):
        top = 150 + k * 160
        d.rule(s, x, top, w, 1, color="muted")
        sstyle = _fit(d, it.get("value", ""), w, ["stat_sub", "stat_sub2", "h1"],
                      it.get("unit"), 0.45)
        d.text(s, sstyle, x, top + 14, w, 76, str(it.get("value", "")),
               anchor="bottom", suffix=it.get("unit"), suffix_scale=0.45, tag="stat-sub")
        if it.get("caption"):
            d.text(s, "small", x, top + 98, w, 56, it["caption"], color="muted")


def quote(d: Deck, s, m, sl):
    g = d.g
    d.text(s, "display", g.margin_x, 118, 120, 110, "\u201c", color="accent", tag="quotemark")
    d.text(s, "quote", g.margin_x, 208, g.w(9), 190, sl.get("text", ""), tag="quote")
    if sl.get("source"):
        d.text(s, "label", g.margin_x, 418, g.w(6), 36, "\u2014 " + sl["source"], color="muted")


def table(d: Deck, s, m, sl):
    """네이티브 표를 쓰지 않는다. 테마 표스타일(줄무늬 채움·테두리)이 곧 AI 티다."""
    g = d.g
    _head(d, s, sl)
    heads, rows = sl.get("headers") or [], sl.get("rows") or []
    lim = d.limits["table_rows"]
    if len(rows) > lim:
        warn(f"slide {d._slide_i}: 표 {len(rows)}행 > {d.density} 한도 {lim}행 — 이어지는 슬라이드로 나눠라")
    rows = rows[:lim]
    n = max(1, len(heads) or (len(rows[0]) if rows else 1))
    cw = (g.content_w - (n - 1) * 24) / n
    xs = [g.margin_x + i * (cw + 24) for i in range(n)]
    for i, h in enumerate(heads[:n]):
        d.text(s, "label", xs[i], 150, cw, 16, str(h), color="accent")
    d.rule(s, g.margin_x, 174, g.content_w, 2)
    y0 = 186
    rh = min(110, max(30, (g.bottom - 24 - y0) / max(1, len(rows))))  # 행수에 따라 숨통을 조절
    for r_i, r in enumerate(rows):
        y = y0 + r_i * rh
        for i, cell in enumerate(r[:n]):
            d.text(s, "small", xs[i], y, cw, rh - 8, str(cell), tag="cell")
        d.rule(s, g.margin_x, y + rh - 8, g.content_w, 0.75, color="muted")


def cards(d: Deck, s, m, sl):
    """번호 카드 나열. 상자·둥근모서리·그림자·모서리 액센트 줄을 쓰지 않는다.
    테두리 친 카드나 한쪽 모서리 색줄은 AI 슬라이드의 대표적 지문이라 번호와 여백만으로 구획한다.
    카드 높이를 실제로 재서 블록 전체를 광학 중심에 앉힌다 — 위로 몰리면 아래가 통째로 빈다."""
    g = d.g
    head_y = _head(d, s, sl)
    items = (sl.get("items") or [])[:6]
    n = len(items)
    if not n:
        return
    if len(sl.get("items") or []) > 6:
        warn(f"slide {d._slide_i}: 카드 6개 초과 — 두 장으로 나눠라")

    cols = 2 if n in (2, 4) else 3
    span = {2: 6, 3: 4}[cols]
    w = g.w(span)
    SP, FN = d.spec["styles"], d.spec["fonts"]

    def part(text, style):
        if not text:
            return 0.0
        st = SP[style]
        return est_height([str(text)], st["size"], st["leading"], w, font=FN[st["font"]])

    blocks = []
    for it in items:
        nh = part(it.get("number"), "h2")
        th = part(it.get("title"), "h2")
        bh = part(it.get("body"), "small")
        blocks.append((nh, th, bh,
                       nh + (10 if nh else 0) + th + (12 if bh else 0) + bh))

    rows = (n + cols - 1) // cols
    THUMB = 108 + 16 if any(it.get("image") for it in items) else 0
    rowh = [THUMB + max(blocks[r * cols + c][3] for c in range(cols) if r * cols + c < n)
            for r in range(rows)]
    RGAP = 40
    total = sum(rowh) + RGAP * (rows - 1)
    avail = g.bottom - head_y
    top0 = head_y + max(28, min((avail - total) / 2, 120))
    if total > avail:
        warn(f"slide {d._slide_i}: 카드 블록 {total:.0f}pt > 남은 높이 {avail:.0f}pt "
             f"— 본문을 줄이거나 두 장으로 나눠라")

    thumb = 108 if any(it.get("image") for it in items) else 0
    for k, it in enumerate(items):
        r, c = divmod(k, cols)
        x = g.x(c * span)
        y = top0 + sum(rowh[:r]) + RGAP * r
        if thumb:
            if it.get("image"):
                d.picture(s, it["image"], x, y, w, thumb, it.get("fit", "cover"),
                          tag="card-img", focus=it.get("focus", "top"))
            y += thumb + 16
        nh, th, bh, _ = blocks[k]
        if nh:
            d.text(s, "h2", x, y, w, nh + 4, str(it["number"]), color="accent", tag="card-num")
            y += nh + 10
        if th:
            d.text(s, "h2", x, y, w, th + 4, it["title"], tag="card-title")
            y += th + (12 if bh else 0)
        if bh:
            d.text(s, "small", x, y, w, bh + 4, it["body"], color="muted", tag="card-body")


def image_split(d: Deck, s, m, sl):
    """반출혈 이미지 + 텍스트. 이미지가 화면 절반을 끝까지 밀고 나간다.
    액자에 넣지 않는 이유 — 테두리 친 이미지 카드가 템플릿 냄새의 출처다."""
    g = d.g
    side = sl.get("side", "right")
    IW = 420
    if side == "left":
        d.picture(s, sl["image"], 0, 0, IW, g.height, sl.get("fit", "cover"),
                  focus=sl.get("focus", "center"))
        tx, tw = g.x(6), g.w(6)
    else:
        d.picture(s, sl["image"], g.width - IW, 0, IW, g.height, sl.get("fit", "cover"),
                  focus=sl.get("focus", "center"))
        tx, tw = g.margin_x, g.w(6)

    SP, FN = d.spec["styles"], d.spec["fonts"]
    bstyle, gap = ("h2", 26) if d.density == "speaker" else ("body", 14)
    bs = SP[bstyle]
    bullets = [str(b) for b in (sl.get("bullets") or [])]
    lim = d.limits["bullets"]
    if len(bullets) > lim:
        warn(f"slide {d._slide_i}: 불릿 {len(bullets)}개 > {d.density} 한도 {lim}개")

    th = est_height([sl.get("title", "")], SP["h1"]["size"], SP["h1"]["leading"], tw,
                    font=FN[SP["h1"]["font"]])
    lh = est_height([sl["lead"]], SP["h2"]["size"], SP["h2"]["leading"], tw,
                    font=FN[SP["h2"]["font"]]) if sl.get("lead") else 0
    bh = est_height(bullets, bs["size"], bs["leading"], tw, gap, True,
                    font=FN[bs["font"]]) if bullets else 0
    total = th + (lh + 24 if lh else 0) + (bh + 30 if bh else 0)
    y = max(g.margin_top + 24, (g.height - total) / 2)
    if sl.get("label"):
        d.text(s, "label", tx, y - 28, tw, 16, sl["label"], color="accent")
    d.text(s, "h1", tx, y, tw, th + 6, sl.get("title", ""), tag="split-title")
    y += th
    if lh:
        y += 24
        d.text(s, "h2", tx, y, tw, lh + 6, sl["lead"], color="muted")
        y += lh
    if bh:
        y += 30
        d.text(s, bstyle, tx, y, tw, bh + 6, bullets, accent_lead=True,
               space_after=gap, tag="bullets")
    if sl.get("caption"):
        d.text(s, "label", tx, g.bottom - 16, tw, 16, sl["caption"], color="muted")


def image_full(d: Deck, s, m, sl):
    """위는 전출혈 이미지, 아래는 바탕색 판에 제목·캡션.
    사진 위에 글자를 얹지 않는다 — 어떤 사진이 올지 모르는 채로 가독성을 도박하지 않는다.
    판 높이는 글 분량에 맞춰 늘린다. 고정 높이로 두면 두 줄짜리 캡션이 화면 밖으로 나간다."""
    g = d.g
    SP, FN = d.spec["styles"], d.spec["fonts"]
    tw = g.w(9)
    th = est_height([sl.get("title", "")], SP["h1"]["size"], SP["h1"]["leading"], tw,
                    font=FN[SP["h1"]["font"]]) if sl.get("title") else 0
    ch = est_height([sl["caption"]], SP["small"]["size"], SP["small"]["leading"], tw,
                    font=FN[SP["small"]["font"]]) if sl.get("caption") else 0
    lh = 26 if sl.get("label") else 0
    plate = 30 + lh + th + (12 if ch else 0) + ch + 32

    d.picture(s, sl["image"], 0, 0, g.width, g.height - plate, sl.get("fit", "cover"),
              focus=sl.get("focus", "center"))
    y = g.height - plate + 30
    if lh:
        d.text(s, "label", g.margin_x, y, tw, 16, sl["label"], color="accent")
        y += lh
    if th:
        d.text(s, "h1", g.margin_x, y, tw, th + 6, sl["title"], tag="full-title")
        y += th + 12
    if ch:
        d.text(s, "small", g.margin_x, y, tw, ch + 6, sl["caption"], color="muted")


def closing(d: Deck, s, m, sl):
    g = d.g
    d.rule(s, g.right - g.w(3), 176, g.w(3), 4)
    d.text(s, "display", g.x(3), 200, g.w(9), 180, sl.get("title", "감사합니다"),
           anchor="bottom", align="right", tag="closing-title")
    lines = sl.get("lines") or [x for x in [m.get("author"), m.get("contact")] if x]
    if lines:
        d.text(s, "body", g.x(5), 404, g.w(7), 80, lines, color="muted", align="right")


LAYOUTS = {"cover": cover, "section": section, "statement": statement,
           "two_col": two_col, "cards": cards, "data": data, "quote": quote,
           "table": table, "image_split": image_split, "image_full": image_full,
           "closing": closing}
INVERTED = {"section"}            # accent 바탕으로 반전되는 레이아웃
NO_PAGENUM = {"cover", "section", "closing", "image_full"}


# ================================================================= 빌드
def build(outline_path, out_path=None, spec_path=None, embed=False):
    with open(outline_path, encoding="utf-8") as f:
        o = yaml.safe_load(f)
    spec = load_spec(spec_path)
    m = o.get("meta") or {}
    d = Deck(spec, palette=m.get("palette"), density=m.get("density"))

    slides = o.get("slides") or []
    run, prev = 0, None
    for i, sl in enumerate(slides, 1):
        lay = sl.get("layout")
        if lay not in LAYOUTS:
            raise SystemExit(f"slide {i}: 알 수 없는 layout '{lay}' — {sorted(LAYOUTS)} 중 하나여야 한다")
        run = run + 1 if lay == prev else 1
        if run > spec["rhythm"]["max_same_layout_run"]:
            warn(f"slide {i}: '{lay}' {run}연속 — 리듬이 죽는다. 다른 레이아웃을 끼워라")
        prev = lay
        s = d.slide(bg="accent" if lay in INVERTED else "ground",
                    invert=bool(sl.get("invert")))
        LAYOUTS[lay](d, s, m, sl)
        if lay not in NO_PAGENUM and i >= spec["rhythm"]["page_number_from"]:
            d.page_number(s, i)

    out = out_path or m.get("output") or "out/deck.pptx"
    pptx, man = d.save(out, embed=embed or bool(m.get('embed_fonts')))

    # --- 기하 자기검증: 넘침 / 캔버스 이탈 -------------------------------
    for b in d.manifest:
        need = est_height(b["text"], b["size"], b["leading"], b["w"],
                          b["space_after"], b["accent_lead"], b.get("extra_pt", 0.0),
                          b.get("font"))
        if need > b["h"] * 1.02:
            warn(f"slide {b['slide']}: '{b['tag']}' 넘침 추정 {need:.0f}pt > 상자 {b['h']:.0f}pt "
                 f"— 글자를 줄이지 말고 내용을 쪼개라")
        if b["x"] + b["w"] > d.g.width + 0.5 or b["y"] + b["h"] > d.g.height + 0.5:
            warn(f"slide {b['slide']}: '{b['tag']}' 캔버스 이탈")

    for im_ in d.images:                      # 이미지가 본문 글자를 덮지 않는지
        for b in d.manifest:
            if b["slide"] != im_["slide"] or b["tag"] == "pagenum":
                continue
            ox = min(im_["x"]+im_["w"], b["x"]+b["w"]) - max(im_["x"], b["x"])
            oy = min(im_["y"]+im_["h"], b["y"]+b["h"]) - max(im_["y"], b["y"])
            if ox > 2 and oy > 2:
                warn(f"slide {b['slide']}: 이미지가 '{b['tag']}' 를 덮는다 "
                     f"({ox:.0f}x{oy:.0f}pt) — 글자를 이미지 위에 얹지 마라")

    print(f"✓ {pptx}  ({len(slides)} slides, palette={d.pal.__dict__['ground']}/"
          f"{d.pal.accent}, density={d.density})")
    print(f"  manifest: {man}")
    if getattr(d, "embedded", 0):
        print(f"  글꼴 {d.embedded}종 임베드 ({os.path.getsize(pptx)/1048576:.1f} MB)")
    if WARN:
        print(f"\n⚠ {len(WARN)}건:", file=sys.stderr)
        for w in WARN:
            print("  - " + w, file=sys.stderr)
    return pptx


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("outline")
    ap.add_argument("-o", "--out")
    ap.add_argument("--spec")
    ap.add_argument("--embed-fonts", action="store_true",
                    help="Pretendard TTF를 파일에 박는다 (굵기당 1~2MB)")
    a = ap.parse_args()
    build(a.outline, a.out, a.spec, a.embed_fonts)
