#!/usr/bin/env python3
"""ppt-deck 빌더 — outline.yaml -> .pptx

좌표 정본: 'ppt-deck 재설계' 명세서 §3(좌표 표) / HANDOFF §8.
x·w 는 전부 col_x()/span_w() 로만 계산한다. 명세서의 pt 실값은 검증용이다.

사용:  build.py outline.yaml [-o out/deck.pptx] [--spec other.yaml] [--embed-fonts]
"""
from __future__ import annotations
import argparse, os, sys

import yaml
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import (Deck, load_spec, col_x, span_w, block_h, line_count,
                     resolve_y, CONTENT_R)

WARN: list[str] = []


def warn(msg): WARN.append(msg)


def _txt(sl, *keys):
    for k in keys:
        if sl.get(k):
            return str(sl[k])
    return ""


def _eyebrow(d, s, sl, span=7):
    """모든 레이아웃 공통. 정렬과 무관하게 항상 y 56 (§7)."""
    if sl.get("eyebrow"):
        d.text(s, "eyebrow", col_x(1), 56, span_w(span), 24, sl["eyebrow"],
               color="accent", tag="eyebrow")


def _runner(d, s, sl, y=60):
    """우상단 러너. small muted 우정렬."""
    if sl.get("runner"):
        d.text(s, "small", col_x(9), y, span_w(4), 22, sl["runner"],
               color="muted", align="right", tag="runner")


# ================================================================= §8.1 cover
def cover(d: Deck, s, m, sl):
    _eyebrow(d, s, sl, span=9)
    if sl.get("year"):
        d.text(s, "stat2", col_x(9), 56, span_w(4), 116, str(sl["year"]),
               color="accent", align="right", tag="cover-year")
    d.text(s, "display", col_x(1), 252, span_w(9), 165,
           _txt(sl, "title") or m.get("title", ""), tag="cover-title")
    meta = sl.get("meta") or [x for x in (m.get("org"), m.get("date")) if x]
    if meta:
        d.text(s, "small", col_x(1), 441, span_w(5), 44, meta,
               color="muted", tag="cover-meta")


# ================================================================= §8.2 closing
def closing(d: Deck, s, m, sl):
    """반전 필드. cover와 네 축이 다르다 — 명도·제목 줄수·측정·보조 요소 (§3 #5)."""
    _eyebrow(d, s, sl, span=5)
    for i, row in enumerate((sl.get("ledger") or [])[:3]):
        y = 196 + i * 70                      # 행 step 70
        d.text(s, "small", col_x(8), y, span_w(5), 22, row.get("label", ""),
               color="muted", tag="ledger-label")
        d.text(s, "body", col_x(8), y + 24, span_w(5), 28, row.get("value", ""),
               font_key="head", tag="ledger-value")   # body 크기 + SemiBold
    title = _txt(sl, "title") or "감사합니다"
    if line_count(title, span_w(7), d.spec["styles"]["display"],
                  d.spec["fonts"]["display"]) > 1:
        warn(f"slide {d._slide_i}: closing 제목이 2줄이다 — 1줄만 허용된다"
             f"(2줄이면 cover와 구별이 사라진다). 문구를 줄여라")
    d.text(s, "display", col_x(1), 402, span_w(7), 82, title, tag="closing-title")


# ================================================================= §8.3 section
def section(d: Deck, s, m, sl):
    """반전 필드. 상단 164pt 덩어리와 하단 165pt 덩어리 사이 99pt 공백이 앵커다."""
    num = str(sl.get("number", ""))
    if len(num) > 2:
        warn(f"slide {d._slide_i}: section 번호 '{num}' 는 세 자리 — 두 자리만 지원한다")
    d.text(s, "mega", col_x(1), 56, span_w(7), 164, num, color="accent", tag="section-num")
    _runner(d, s, sl, y=60)
    d.text(s, "display", col_x(1), 319, span_w(9), 165, _txt(sl, "title"),
           tag="section-title")


# ================================================================= §8.4 statement
def statement(d: Deck, s, m, sl):
    text = _txt(sl, "text")
    ST, FN = d.spec["styles"], d.spec["fonts"]
    n = line_count(text, span_w(10), ST["display"], FN["display"])
    full = bool(sl.get("body"))
    _eyebrow(d, s, sl, span=7)
    _runner(d, s, sl, y=60)
    # 강조는 선이 아니라 색 — 문단(행) 단위로만
    paras = text.split("\n")
    acc = sl.get("accent_lines") or ([len(paras) - 1] if len(paras) > 1 else [])
    if full:
        d.text(s, "display", col_x(1), 120, span_w(10), 247, paras,
               accent_paras=tuple(acc), tag="statement")
        d.text(s, "body", col_x(1), 400, span_w(7), 84, sl["body"], tag="statement-body")
    else:
        if n > 2:
            warn(f"slide {d._slide_i}: statement 본문이 {n}줄 — 보조 body 없이 3줄이면 "
                 f"body를 넣어 '가득' 상태로 쓰거나 문구를 줄여라")
        d.text(s, "display", col_x(1), 319, span_w(10), 165, paras,
               accent_paras=tuple(acc), tag="statement")


# ================================================================= §8.5 two_col
def two_col(d: Deck, s, m, sl):
    """비대칭 4:7. col5를 통째로 비워 94pt 실공백. 하단 정렬이면 양쪽 모두 484."""
    ST, FN = d.spec["styles"], d.spec["fonts"]
    _eyebrow(d, s, sl, span=4)
    title = _txt(sl, "title")
    paras = [str(b) for b in (sl.get("bullets") or [])] or \
            ([str(sl["body"])] if sl.get("body") else [])
    lim = d.limits["bullets"]
    if len(paras) > lim:
        warn(f"slide {d._slide_i}: 본문 단락 {len(paras)}개 > {d.density} 한도 {lim}개")

    body_h = block_h("\n".join(paras), span_w(7), ST["body"], FN["body"]) if paras else 0
    full = body_h >= 150 or bool(sl.get("subhead"))

    if full:
        d.text(s, "h1", col_x(1), 108, span_w(4), 130, title, tag="col-title")
        if sl.get("lead"):
            d.text(s, "small", col_x(1), 262, span_w(4), 87, sl["lead"], color="muted")
        if paras:
            d.text(s, "body", col_x(6), 108, span_w(7), 223, paras, tag="bullets")
        if sl.get("subhead"):
            d.text(s, "h2", col_x(6), 363, span_w(7), 31, sl["subhead"])
            d.text(s, "body", col_x(6), 400, span_w(7), 84, sl.get("subbody", ""))
    else:
        d.text(s, "h1", col_x(1), 398, span_w(4), 86, title, tag="col-title")
        if paras:
            d.text(s, "body", col_x(6), 372, span_w(7), 112, paras, tag="bullets")


# ================================================================= §8.6–8.10 cards
CARD_MODE = {2: "pair", 3: "ledger", 4: "quad", 5: "dense", 6: "dense"}


def cards_mode(items: list) -> str:
    n = len(items)
    if n not in CARD_MODE:
        raise SystemExit(f"cards는 항목 2~6개만 지원한다 (받은 값: {n}). "
                         f"7개 이상은 슬라이드를 쪼개라. 1개는 statement/data를 써라.")
    return CARD_MODE[n]


def cards(d: Deck, s, m, sl):
    items = sl.get("items") or []
    mode = cards_mode(items)
    _eyebrow(d, s, sl, span=7)
    d.text(s, "h1", col_x(1), 86, span_w(9), 43, _txt(sl, "title"), tag="cards-head")
    {"pair": _cards_pair, "ledger": _cards_ledger,
     "quad": _cards_quad, "dense": _cards_dense}[mode](d, s, items)
    return mode


def _card_body(d, s, it, x, w, iy, ty, by, t_style, b_style, b_color="muted"):
    d.text(s, "eyebrow", x, iy, w, 24, str(it.get("index", "")),
           color="accent", tag="card-index")
    d.text(s, "h2", x, ty, w, 62, it.get("title", ""), tag="card-title")
    if it.get("body"):
        d.text(s, b_style, x, by, w, 84, it["body"], color=b_color, tag="card-body")


def _cards_pair(d, s, items):
    """불균등 5:6 + 수직 엇단 84pt. 테두리·배경 채움 없음."""
    _card_body(d, s, items[0], col_x(1), span_w(5), 177, 209, 287, "h2", "body")
    _card_body(d, s, items[1], col_x(7), span_w(6), 261, 293, 371, "h2", "body")


def _cards_ledger(d, s, items):
    """전폭 3행 원장. 행 구분선 없음 — 46pt 색인 숫자가 행 시작점을 잡는다."""
    for i, it in enumerate(items[:3]):
        y = 172 + i * 109                     # 행 높이 94, step 109
        d.text(s, "stat_sub2", col_x(1), y, 70, 46, str(it.get("index", "")),
               color="accent", tag="ledger-index")
        d.text(s, "h2", col_x(2), y + 6, span_w(4), 62, it.get("title", ""),
               tag="ledger-title")
        if it.get("body"):
            d.text(s, "body", col_x(7), y + 4, span_w(6), 84, it["body"],
                   color="muted", tag="ledger-body")


def _cards_quad(d, s, items):
    """불균등 2x2. 열폭 326 / 396. 균등 4분할이 아니다."""
    slots = [(col_x(1), span_w(5), 177), (col_x(7), span_w(6), 177),
             (col_x(1), span_w(5), 338), (col_x(7), span_w(6), 338)]
    for it, (x, w, y) in zip(items[:4], slots):
        _card_body(d, s, it, x, w, y, y + 30, y + 70, "h2", "small")


def _cards_dense(d, s, items):
    """3x2. 항목 5개면 첫 행을 두 칸으로 바꾼다 (균등 배치 회피)."""
    n = len(items)
    if n == 5:
        slots = [(col_x(1), span_w(7), 177), (col_x(9), span_w(4), 177),
                 (col_x(1), span_w(4), 338), (col_x(5), span_w(4), 338),
                 (col_x(9), span_w(4), 338)]
    else:
        slots = [(col_x(1), span_w(4), 177), (col_x(5), span_w(4), 177),
                 (col_x(9), span_w(4), 177), (col_x(1), span_w(4), 338),
                 (col_x(5), span_w(4), 338), (col_x(9), span_w(4), 338)]
    for it, (x, w, y) in zip(items[:6], slots):
        _card_body(d, s, it, x, w, y, y + 30, y + 100, "h2", "small")


# ================================================================= §8.11 data
def data(d: Deck, s, m, sl):
    items = sl.get("items") or []
    if not items:
        return
    hero, subs = items[0], items[1:4]
    _eyebrow(d, s, sl, span=7)
    if subs:
        d.text(s, "stat1", col_x(1), 140, span_w(6), 136, str(hero.get("value", "")),
               suffix=hero.get("unit"), suffix_style="stat_sub", tag="hero")
        d.hero_rule(s, col_x(1), 292, span_w(6))
        if hero.get("caption"):      # 히어로 캡션은 muted가 아니라 figure
            d.text(s, "body", col_x(1), 308, span_w(5), 84, hero["caption"], tag="hero-cap")
        for i, it in enumerate(subs):
            y = 140 + i * 104                 # 종속 행 step 104
            d.text(s, "stat_sub2", col_x(8), y, span_w(5), 46, str(it.get("value", "")),
                   suffix=it.get("unit"), suffix_style="h2", tag="sub-value")
            if it.get("caption"):
                d.text(s, "small", col_x(8), y + 54, span_w(5), 44, it["caption"],
                       color="muted", tag="sub-label")
    else:
        _runner(d, s, sl, y=60)
        d.text(s, "stat1", col_x(1), 260, span_w(9), 136, str(hero.get("value", "")),
               suffix=hero.get("unit"), suffix_style="stat_sub", tag="hero")
        d.hero_rule(s, col_x(1), 412, span_w(9))
        if hero.get("caption"):
            d.text(s, "body", col_x(1), 428, span_w(7), 56, hero["caption"], tag="hero-cap")


# ================================================================= §8.12 quote
def quote(d: Deck, s, m, sl):
    """항상 하단 정렬. col1을 통째로 비운 70pt 들여쓰기가 앵커다."""
    ST, FN = d.spec["styles"], d.spec["fonts"]
    text = _txt(sl, "text")
    h = block_h(text, span_w(9), ST["quote"], FN["head"])
    # §14.5 — 반올림으로 1pt 차가 나는 계산값은 좌표 표의 표기값을 쓴다 (3줄 232 / 2줄 288 / 4줄 176)
    d.text(s, "quote", col_x(2), round(401 - h), span_w(9), round(h), text, tag="quote")
    if sl.get("source"):
        d.text(s, "eyebrow", col_x(2), 433, span_w(5), 24, sl["source"],
               color="accent", tag="quote-source")
    if sl.get("note"):
        d.text(s, "small", col_x(2), 461, span_w(5), 23, sl["note"], color="muted")


# ================================================================= §8.13 table
def table(d: Deck, s, m, sl):
    """네이티브 표를 쓰지 않는다. 열은 그리드 스냅 + 24pt 공백 + 우정렬이 가른다."""
    T = d.spec["table"]
    heads, rows = sl.get("headers") or [], sl.get("rows") or []
    n = len(heads) or (len(rows[0]) if rows else 0)
    if n not in T["col_pattern"]:
        raise SystemExit(f"table은 2~5열만 지원한다 (받은 값: {n}). 열을 줄이거나 슬라이드를 쪼개라.")
    lim = d.limits["table_rows"]
    if len(rows) > lim:
        warn(f"slide {d._slide_i}: 표 {len(rows)}행 > {d.density} 한도 {lim}행 "
             f"— 행 높이를 줄이지 말고 슬라이드를 나눠라")
    rows = rows[:lim]
    cols = table_columns(d, n)

    _eyebrow(d, s, sl, span=7)
    d.text(s, "h1", col_x(1), 86, span_w(9), 43, _txt(sl, "title"), tag="table-head")
    d.plate(s, col_x(1), T["header_y"], span_w(12), T["header_h"], color="figure")
    for (x, w, align, tx), h in zip(cols, heads[:n]):
        d.text(s, "small", tx if align == "left" else x, T["header_y"] + 6,
               w - T["pad_x"], 22, str(h), color="ground", align=align, tag="th")
    for r_i, row in enumerate(rows):
        y = T["first_row_y"] + r_i * T["row_h"]
        for (x, w, align, tx), cell in zip(cols, row[:n]):
            d.text(s, "body", tx if align == "left" else x, y + 6,
                   w - T["pad_x"], 28, str(cell), align=align,
                   font_key="head" if align == "right" else None, tag="td")
        if r_i < len(rows) - 1:               # 마지막 행 뒤에는 선을 넣지 않는다
            d.table_rule(s, col_x(1), y + 33.5, span_w(12))
    if sl.get("footnote"):
        d.text(s, "small", col_x(1), 424, span_w(7), 44, sl["footnote"], color="muted")


def table_columns(d: Deck, n: int):
    """[(x, w, align, text_x)]. 첫 열만 좌정렬, 나머지는 우정렬 (§11)."""
    spans = d.spec["table"]["col_pattern"][n]
    pad = d.spec["table"]["pad_x"]
    out, col = [], 1
    for i, sp in enumerate(spans):
        x, w = col_x(col), span_w(sp)
        out.append((x, w, "left", x + pad) if i == 0 else (x, w, "right", x + w - pad))
        col += sp
    return out


# ================================================================= §8.14–8.15 image
def image_split(d: Deck, s, m, sl):
    """이미지는 우·상·하 3변 재단. 좌우를 뒤집는 변형은 만들지 않는다 (대칭 반복 금지)."""
    d.picture(s, sl["image"], col_x(7), 0, 468, 540, sl.get("fit", "cover"),
              focus=sl.get("focus", "center"))
    _eyebrow(d, s, sl, span=5)
    d.text(s, "h1", col_x(1), 108, span_w(5), 130, _txt(sl, "title"), tag="split-title")
    body = [str(b) for b in (sl.get("bullets") or [])] or \
           ([str(sl["body"])] if sl.get("body") else [])
    if body:
        d.text(s, "body", col_x(1), 262, span_w(5), 195, body, tag="split-body")


def image_full(d: Deck, s, m, sl):
    """이미지 전출혈 + ground 단색 판. 판이 문제 3의 앵커 장치다."""
    d.picture(s, sl["image"], 0, 0, 960, 540, sl.get("fit", "cover"),
              focus=sl.get("focus", "center"))
    d.plate(s, 0, 300, 560, 240, color="ground")
    if sl.get("eyebrow"):
        d.text(s, "eyebrow", 72, 340, 440, 24, sl["eyebrow"], color="accent", tag="eyebrow")
    d.text(s, "h1", 72, 372, 440, 87, _txt(sl, "title"), tag="full-title")
    if sl.get("caption"):
        d.text(s, "small", 72, 471, 440, 22, sl["caption"], color="muted")


LAYOUTS = {"cover": cover, "closing": closing, "section": section,
           "statement": statement, "two_col": two_col, "cards": cards,
           "data": data, "quote": quote, "table": table,
           "image_split": image_split, "image_full": image_full}
INVERTED = {"section", "closing"}          # 반전 필드 — 배경색으로 처리 (§12)


# ================================================================= 빌드
def build(outline_path, out_path=None, spec_path=None, embed=False):
    with open(outline_path, encoding="utf-8") as f:
        o = yaml.safe_load(f)
    spec = load_spec(spec_path)
    m = o.get("meta") or {}

    base = os.path.dirname(os.path.abspath(outline_path))
    def _abs(v):
        v = os.path.expanduser(str(v))
        return v if os.path.isabs(v) else os.path.normpath(os.path.join(base, v))
    for sl in (o.get("slides") or []):
        if sl.get("image"):
            sl["image"] = _abs(sl["image"])
        for it in (sl.get("items") or []):
            if it.get("image"):
                it["image"] = _abs(it["image"])

    d = Deck(spec, palette=m.get("palette"), density=m.get("density"))
    slides = o.get("slides") or []
    prev, run, prev_mode = None, 0, None
    for i, sl in enumerate(slides, 1):
        lay = sl.get("layout")
        if lay not in LAYOUTS:
            raise SystemExit(f"slide {i}: 알 수 없는 layout '{lay}' — {sorted(LAYOUTS)}")
        run = run + 1 if lay == prev else 1
        if run > spec["rhythm"]["max_same_layout_run"]:
            warn(f"slide {i}: '{lay}' {run}연속 — 리듬이 죽는다")
        prev = lay
        s = d.slide(invert=lay in INVERTED or bool(sl.get("invert")))
        mode = LAYOUTS[lay](d, s, m, sl)
        if lay == "cards":
            if mode == prev_mode:
                warn(f"slide {i}: cards 모드 '{mode}' 가 직전 장과 같다 — "
                     f"항목 수를 조정해 구성을 바꿔라 (max_same_cards_mode_run 1)")
            prev_mode = mode
        else:
            prev_mode = None

    out = out_path or m.get("output") or "out/deck.pptx"
    pptx, man = d.save(out, embed=embed or bool(m.get("embed_fonts")))

    for b in d.manifest:                      # 기하 자기검증
        if b["x"] + b["w"] > CONTENT_R + 0.5 and b["tag"] not in ("full-title", "eyebrow"):
            warn(f"slide {b['slide']}: '{b['tag']}' 오른변 {b['x']+b['w']:.0f} > 888 (그리드 밖)")
        if b["y"] + b["h"] > 540.5:
            warn(f"slide {b['slide']}: '{b['tag']}' 캔버스 하단 이탈")

    print(f"✓ {pptx}  ({len(slides)} slides, palette={d.pal_name}, density={d.density})")
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
    ap.add_argument("outline"); ap.add_argument("-o", "--out"); ap.add_argument("--spec")
    ap.add_argument("--embed-fonts", action="store_true")
    a = ap.parse_args()
    build(a.outline, a.out, a.spec, a.embed_fonts)
