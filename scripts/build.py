#!/usr/bin/env python3
"""ppt-deck 빌더 — outline.yaml -> .pptx

좌표는 스펙의 그리드와 수직 리듬에서만 나온다. house 프로파일의 리듬은
사용자 덱 4종(137장) 실측값이다 — 눈썹 45 / 제목 68 / 본문 117..495, 좌우 마진 58.

사용:  build.py outline.yaml [-o out/deck.pptx] [--spec other.yaml] [--embed-fonts]
"""
from __future__ import annotations
import argparse, os, sys

import yaml
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import (Deck, load_spec, col_x, span_w, content_r,
                     block_h, block_w, line_count, resolve_y)

WARN: list[str] = []


def warn(msg): WARN.append(msg)


def _t(sl, *keys):
    for k in keys:
        if sl.get(k):
            return str(sl[k])
    return ""


def _paras(sl, *keys):
    for k in keys:
        v = sl.get(k)
        if isinstance(v, list) and v:
            return [str(x) for x in v]
        if isinstance(v, str) and v:
            return [v]
    return []


class L:
    """레이아웃이 공유하는 리듬. 스펙의 rhythm_y 를 그대로 읽는다."""
    def __init__(self, d):
        r = d.spec.get("rhythm_y") or {"eyebrow_y": 56, "title_y": 86,
                                       "content_y": 117, "content_bottom": 484}
        self.eyebrow_y, self.title_y = r["eyebrow_y"], r["title_y"]
        self.top, self.bottom = r["content_y"], r["content_bottom"]
        self.h = self.bottom - self.top
        self.gap = d.spec["anchors"]["gap_block"]
        self.d = d

    def head(self, s, sl, title_style="h1", span=9):
        """눈썹 + 제목. 두 좌표는 모든 장에서 고정이다 — 그래서 훑을 때 눈이 안 흔들린다."""
        d = self.d
        if sl.get("eyebrow"):
            if sl.get("eyebrow_chip") or d.spec.get("eyebrow_style") == "chip":
                SH = d.spec["shapes"]
                cw = self.bh(sl["eyebrow"], 400, "small") * 0 + \
                    line_count(sl["eyebrow"], 400, d.spec["styles"]["small"],
                               d.spec["fonts"]["body"]) * 0
                w = block_w(sl["eyebrow"], d.spec["styles"]["small"],
                            d.spec["fonts"]["body"]) + 36
                d.chip(s, col_x(1), self.eyebrow_y - 6, w, SH["eyebrow_chip_h"],
                       sl["eyebrow"], color="accent", text_color="ground", style="small")
            else:
                d.text(s, "micro", col_x(1), self.eyebrow_y, span_w(span), 15,
                       sl["eyebrow"], color="accent", tag="eyebrow")
        if sl.get("runner"):
            d.text(s, "micro", col_x(9), self.eyebrow_y, span_w(4), 15,
                   sl["runner"], color="muted", align="right", tag="runner")
        if sl.get("title"):
            st = d.spec["styles"][title_style]
            h = block_h(sl["title"], span_w(span), st, d.spec["fonts"][st["font"]])
            d.text(s, title_style, col_x(1), self.title_y, span_w(span), h,
                   str(sl["title"]).split("\n"), tag="title")

    def bh(self, text, w, style):
        st = self.d.spec["styles"][style]
        return block_h(text, w, st, self.d.spec["fonts"][st["font"]])


# ================================================================= 레이아웃
def cover(d, s, m, sl):
    g = L(d)
    if sl.get("eyebrow"):
        d.text(s, "micro", col_x(1), g.eyebrow_y, span_w(8), 15, sl["eyebrow"],
               color="accent", tag="eyebrow")
    meta = _paras(sl, "meta")
    mh = g.bh("\n".join(meta), span_w(6), "small") if meta else 0
    title = _t(sl, "title") or m.get("title", "")
    th = g.bh(title, span_w(9), "cover")
    y_meta = g.bottom - mh
    y_title = y_meta - g.gap - th
    d.text(s, "cover", col_x(1), y_title, span_w(9), th, title.split("\n"), tag="title")
    if meta:
        d.text(s, "small", col_x(1), y_meta, span_w(6), mh, meta,
               color="muted", tag="meta")


def closing(d, s, m, sl):
    """반전 필드. 표지와 다른 축 — 정보 원장이 우측에 서고 제목은 1줄이다."""
    g = L(d)
    if sl.get("eyebrow"):
        d.text(s, "micro", col_x(1), g.eyebrow_y, span_w(5), 15, sl["eyebrow"],
               color="accent", tag="eyebrow")
    rows = (sl.get("ledger") or [])[:3]
    for i, row in enumerate(rows):
        y = g.top + 20 + i * 62
        d.text(s, "micro", col_x(8), y, span_w(5), 15, row.get("label", ""),
               color="muted", tag="ledger-label")
        d.text(s, "h2", col_x(8), y + 18, span_w(5), 28, row.get("value", ""),
               tag="ledger-value")
    title = _t(sl, "title") or "감사합니다"
    th = g.bh(title, span_w(7), "cover")
    d.text(s, "cover", col_x(1), g.bottom - th, span_w(7), th, title, tag="title")


def section(d, s, m, sl):
    """반전 필드. 큰 번호가 아니라 눈썹 + 제목의 위치만으로 장을 가른다."""
    g = L(d)
    if sl.get("number"):
        d.text(s, "display", col_x(1), g.eyebrow_y, span_w(3), 46,
               str(sl["number"]), color="accent", tag="section-num")
    if sl.get("runner"):
        d.text(s, "micro", col_x(9), g.eyebrow_y, span_w(4), 15, sl["runner"],
               color="muted", align="right", tag="runner")
    title = _t(sl, "title")
    th = g.bh(title, span_w(9), "cover")
    d.text(s, "cover", col_x(1), g.bottom - th, span_w(9), th, title.split("\n"),
           tag="title")


def statement(d, s, m, sl):
    g = L(d)
    g.head(s, sl, span=9)
    text = _t(sl, "text")
    th = g.bh(text, span_w(10), "display")
    body = _paras(sl, "body")
    bh_ = g.bh("\n".join(body), span_w(7), "lead") if body else 0
    total = th + (g.gap + bh_ if bh_ else 0)
    y = resolve_y(total, d.spec)
    acc = sl.get("accent_lines") or []
    d.text(s, "display", col_x(1), y, span_w(10), th, text.split("\n"),
           accent_paras=tuple(acc), tag="statement")
    if bh_:
        d.text(s, "lead", col_x(1), y + th + g.gap, span_w(7), bh_, body,
               color="ink2", tag="statement-body")


def two_col(d, s, m, sl):
    """비대칭 4:7. 좌우 실공백 92pt."""
    g = L(d)
    g.head(s, sl, span=7)
    lead = _t(sl, "lead")
    body = _paras(sl, "bullets", "body")
    lw, rw = span_w(4), span_w(7)
    lh = g.bh(lead, lw, "lead") if lead else 0
    rh = g.bh("\n".join(body), rw, "body") if body else 0
    sub, subb = _t(sl, "subhead"), _t(sl, "subbody")
    sh_ = g.bh(sub, rw, "h2") if sub else 0
    sbh = g.bh(subb, rw, "body") if subb else 0
    right_total = rh + (g.gap + sh_ if sh_ else 0) + (8 + sbh if sbh else 0)
    y = resolve_y(max(lh, right_total), d.spec)
    if lh:
        d.text(s, "lead", col_x(1), y, lw, lh, lead, color="muted", tag="lead")
    if rh:
        d.text(s, "body", col_x(6), y, rw, rh, body, color="ink2", tag="body")
    if sh_:
        d.text(s, "h2", col_x(6), y + rh + g.gap, rw, sh_, sub, tag="subhead")
        if sbh:
            d.text(s, "body", col_x(6), y + rh + g.gap + sh_ + 8, rw, sbh, subb,
                   color="ink2", tag="subbody")


CARD_MODE = {2: "pair", 3: "ledger", 4: "quad", 5: "dense", 6: "dense"}


def cards(d, s, m, sl):
    items = sl.get("items") or []
    n = len(items)
    if n not in CARD_MODE:
        raise SystemExit(f"cards는 항목 2~6개만 지원한다 (받은 값: {n}).")
    g = L(d)
    g.head(s, sl, span=9)
    mode = CARD_MODE[n]
    gv = d.spec["cards"]["gutter_v"]
    BIAS = 0.38      # 남는 높이의 38%만 위에 둔다 — 정중앙보다 살짝 위가 안정적이다

    def cell_h(it, w):
        h = d.spec["shapes"]["badge_d_sub"] + 10 + g.bh(it.get("title", ""), w, "h2")
        if it.get("body"):
            h += 8 + g.bh(it["body"], w, "small")
        return h

    SH = d.spec["shapes"]

    def cell(it, x, w, y):
        dd = SH["badge_d_sub"]
        d.badge(s, x + dd / 2, y + dd / 2, dd, "accent",
                glyph=str(it.get("index", "")), glyph_color="ground", style="small")
        th = g.bh(it.get("title", ""), w, "h2")
        d.text(s, "h2", x, y + dd + 10, w, th, it.get("title", ""), tag="card-title")
        if it.get("body"):
            bh_ = g.bh(it["body"], w, "small")
            d.text(s, "small", x, y + SH["badge_d_sub"] + 10 + th + 8, w, bh_,
                   it["body"], color="muted", tag="card-body")

    # 행은 실제 높이로 위에서부터 쌓는다. 남은 높이에 균등 분배하면 흩어져 보인다.
    def place(total):
        return g.top + max(0.0, (g.h - total) * BIAS)

    if mode == "pair":
        cols = ((1, 5), (7, 6))
        y = place(max(cell_h(it, span_w(sp)) for it, (c, sp) in zip(items, cols)))
        for it, (c, sp) in zip(items, cols):
            cell(it, col_x(c), span_w(sp), y)
    elif mode == "ledger":
        rows_h = []
        for it in items[:3]:
            rows_h.append(max(g.bh(it.get("title", ""), span_w(4), "h2"),
                              g.bh(it.get("body", ""), span_w(6), "body") if it.get("body") else 0))
        y = place(sum(rows_h) + gv * (len(rows_h) - 1))
        for it in items[:3]:
            th = g.bh(it.get("title", ""), span_w(4), "h2")
            bh_ = g.bh(it.get("body", ""), span_w(6), "body") if it.get("body") else 0
            SH = d.spec["shapes"]
            d.badge(s, col_x(1) + SH["badge_d"] / 2, y + 13, SH["badge_d"], "accent",
                    glyph=str(it.get("index", "")), glyph_color="ground", style="small")
            d.text(s, "h2", col_x(2), y, span_w(4), th, it.get("title", ""),
                   tag="ledger-title")
            if bh_:
                d.text(s, "body", col_x(7), y + 2, span_w(6), bh_, it["body"],
                       color="ink2", tag="ledger-body")
            y += max(th, bh_) + gv
    else:
        slots = {"quad": [(1, 5, 0), (7, 6, 0), (1, 5, 1), (7, 6, 1)],
                 "dense5": [(1, 7, 0), (9, 4, 0), (1, 4, 1), (5, 4, 1), (9, 4, 1)],
                 "dense6": [(1, 4, 0), (5, 4, 0), (9, 4, 0),
                            (1, 4, 1), (5, 4, 1), (9, 4, 1)]}[
            "quad" if mode == "quad" else ("dense5" if n == 5 else "dense6")]
        row0 = max(cell_h(it, span_w(sp)) for it, (c, sp, r) in zip(items, slots) if r == 0)
        row1 = max([cell_h(it, span_w(sp)) for it, (c, sp, r) in zip(items, slots) if r == 1]
                   or [0])
        y0 = place(row0 + (gv + row1 if row1 else 0))
        for it, (c, sp, r) in zip(items, slots):
            cell(it, col_x(c), span_w(sp), y0 + r * (row0 + gv))
    return mode


def data(d, s, m, sl):
    g = L(d)
    g.head(s, sl, span=7)
    items = sl.get("items") or []
    if not items:
        return
    hero, subs = items[0], items[1:4]
    hw = span_w(6) if subs else span_w(9)
    hh = g.bh(str(hero.get("value", "")), hw, "hero")
    cap = hero.get("caption", "")
    ch = g.bh(cap, span_w(5) if subs else span_w(7), "body") if cap else 0
    hero_total = hh + 14 + (2 + 12 + ch if ch else 0)
    subs_h = (len(subs) - 1) * 66 + 58 if subs else 0
    y = resolve_y(max(hero_total, subs_h), d.spec)   # 두 컬럼을 같은 기준선에 세운다
    d.text(s, "hero", col_x(1), y, hw, hh, str(hero.get("value", "")),
           suffix=hero.get("unit"), suffix_style="h1", tag="hero")
    d.hero_rule(s, col_x(1), y + hh + 14, hw)
    if ch:
        d.text(s, "body", col_x(1), y + hh + 14 + 12, span_w(5) if subs else span_w(7),
               ch, cap, color="ink2", tag="hero-cap")
    step = max(66.0, (max(hero_total, subs_h)) / max(1, len(subs)))
    for i, it in enumerate(subs):
        sy = y + i * step
        d.text(s, "h1", col_x(8), sy, span_w(5), 38, str(it.get("value", "")),
               suffix=it.get("unit"), suffix_style="small", tag="sub-value")
        if it.get("caption"):
            d.text(s, "small", col_x(8), sy + 38, span_w(5), 20, it["caption"],
                   color="muted", tag="sub-label")


def quote(d, s, m, sl):
    """col1 을 비운 들여쓰기가 앵커다. 인용부호 도형·세로선 없음."""
    g = L(d)
    text = _t(sl, "text")
    th = g.bh(text, span_w(8), "h1")
    # 아래에서부터 쌓는다 — 부기 하단이 본문 하한에 정확히 닿는다
    y_note = g.bottom - 18 if sl.get("note") else g.bottom
    y_src = y_note - (18 if sl.get("source") else 0)
    y = y_src - 20 - th
    d.text(s, "h1", col_x(2), y, span_w(8), th, text, tag="quote")
    if sl.get("source"):
        d.text(s, "micro", col_x(2), y_src, span_w(5), 15, sl["source"],
               color="accent", tag="quote-source")
    if sl.get("note"):
        d.text(s, "small", col_x(2), y_note, span_w(5), 18, sl["note"],
               color="muted", tag="quote-note")


def chain(d, s, m, sl):
    """전제 패널 + 꼬리를 무는 질문. 사용자 덱 12쪽의 구조를 좌표까지 그대로 옮긴 것이다.
    패널 = 묶음, 배지 = 행의 시작점, 연결선 = 파고드는 관계. 셋 다 장식이 아니다."""
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    y = g.top
    if sl.get("lead"):
        lh = g.bh(sl["lead"], span_w(10), "lead")
        d.text(s, "lead", col_x(1), y, span_w(10), lh, sl["lead"], color="muted", tag="lead")
        y += lh + 30

    bx = col_x(1) + 40                      # 배지 중심 x (실측 98)
    if sl.get("panel"):
        ph = SH["panel_h"]
        d.panel(s, col_x(1), y, span_w(12), ph, radius=SH["panel_radius"])
        d.badge(s, bx, y + ph / 2, SH["badge_d"], "accent",
                glyph=sl.get("panel_mark", "\u21b3"), glyph_color="ground", style="h2")
        d.text(s, "body", col_x(1) + 73, y + ph / 2 - 11, span_w(12) - 90, 22,
               sl["panel"], font_key="head", tag="panel-text")
        y += ph

    steps = sl.get("steps") or []
    dsub, ind, step = SH["badge_d_sub"], SH["chain_indent"], SH["chain_step"]
    if len(steps) > 4:
        warn(f"slide {d._slide_i}: chain 단계 {len(steps)}개 — 4개까지만 들여쓰기가 화면에 든다")
    for i, tx in enumerate(steps[:4]):
        cx = bx + i * ind
        cy = y + 39 + i * step
        d.connector(s, cx, (y + 4.6) if i == 0 else (cy - step + dsub / 2 + 5),
                    (cy - dsub / 2) - ((y + 4.6) if i == 0 else (cy - step + dsub / 2 + 5)))
        d.badge(s, cx, cy, dsub, "accent_soft", glyph="\u21b3",
                glyph_color="ground", style="micro")
        d.text(s, "small", cx + 17, cy - 9, 22, 18, str(i + 1),
               color="accent", font_key="head", tag="chain-num")
        d.text(s, "body", cx + 43, cy - 11, content_r() - (cx + 43), 22, str(tx),
               font_key="head", tag="chain-text")


def panel_list(d, s, m, sl):
    """전폭 tint 패널을 쌓는다. 실측 — 면담01 7쪽. 패널 높이는 내용이 정하고,
    패널 사이는 27.5pt 로 고정. 각 패널은 번호 칩 + 주문장 + 하위 행으로 구성된다."""
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    y = g.top
    if sl.get("lead"):
        lh = g.bh(sl["lead"], span_w(10), "small")
        d.text(s, "small", col_x(1), y, span_w(10), lh, sl["lead"],
               color="muted", tag="lead")
        y += lh + 22

    px, pw = col_x(1), span_w(12)
    tx = px + SH["panel_text_x"]
    for it in (sl.get("items") or [])[:4]:
        lead_h = g.bh(it.get("text", ""), pw - SH["panel_text_x"] - SH["panel_pad_x"], "lead")
        subs = it.get("subs") or []
        ph = SH["panel_pad_y"] + lead_h + 8 + len(subs) * SH["subrow_step"] + SH["panel_pad_y"]
        d.panel(s, px, y, pw, ph, radius=SH["panel_radius"])
        d.chip(s, px + SH["panel_pad_x"] + 0.5, y + SH["panel_pad_y"],
               SH["chip_w"], SH["chip_h"], str(it.get("index", "")),
               color="accent", text_color="ground", style="small")
        d.text(s, "lead", tx, y + SH["panel_pad_y"] + 1,
               pw - SH["panel_text_x"] - SH["panel_pad_x"], lead_h,
               it.get("text", ""), font_key="head", tag="panel-text")
        sy = y + SH["panel_pad_y"] + lead_h + 8
        for k, sub in enumerate(subs):
            d.text(s, "body", tx, sy + k * SH["subrow_step"], 40, 18,
                   f"{it.get('index', '')}-{k + 1}", color="accent", tag="panel-subnum")
            d.text(s, "body", tx + SH["subrow_indent"], sy + k * SH["subrow_step"],
                   pw - SH["panel_text_x"] - SH["subrow_indent"] - SH["panel_pad_x"], 18,
                   str(sub), color="ink2", tag="panel-sub")
        if y + ph > g.bottom + 0.5:
            warn(f"slide {d._slide_i}: 패널이 본문 하한을 넘는다 — 항목을 줄여라")
        y += ph + SH["panel_gap"]


def card_grid(d, s, m, sl):
    """세로 카드 그리드. 실측 — Dive 14쪽. 카드 156x245, 간격 14.4,
    배지는 카드 가로 중앙, 제목·본문은 좌 13pt 안여백."""
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = (sl.get("items") or [])[:5]
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"card_grid 는 항목 2~5개만 지원한다 (받은 값: {n}).")
    total = span_w(12)
    cw = (total - SH["card_gap"] * (n - 1)) / n
    ch = SH["card_h"]
    y = g.top + max(0.0, (g.h - ch) * 0.35)
    pad, bd = SH["card_pad"], SH["card_badge_d"]
    for i, it in enumerate(items):
        x = col_x(1) + i * (cw + SH["card_gap"])
        d.panel(s, x, y, cw, ch, color=sl.get("card_color", "accent_tint"),
                radius=SH["panel_radius"])
        d.badge(s, x + cw / 2, y + 58, bd, "accent",
                glyph=str(it.get("index", "")), glyph_color="ground", style="h2")
        th = g.bh(it.get("title", ""), cw - pad * 2, "h1")
        d.text(s, "h1", x + pad, y + 105, cw - pad * 2, th, it.get("title", ""),
               tag="card-title")
        if it.get("body"):
            bh_ = g.bh(it["body"], cw - pad * 2, "small")
            d.text(s, "small", x + pad, y + 172, cw - pad * 2, bh_, it["body"],
                   color="ink2", tag="card-body")


def table(d, s, m, sl):
    T = d.spec["table"]
    g = L(d)
    g.head(s, sl, span=9)
    heads, rows = sl.get("headers") or [], sl.get("rows") or []
    n = len(heads) or (len(rows[0]) if rows else 0)
    if n not in T["col_pattern"]:
        raise SystemExit(f"table 은 2~5열만 지원한다 (받은 값: {n}).")
    lim = d.limits["table_rows"]
    if len(rows) > lim:
        warn(f"slide {d._slide_i}: 표 {len(rows)}행 > 한도 {lim}행 — 슬라이드를 나눠라")
    rows = rows[:lim]
    cols, pad = table_columns(d, n), T["pad_x"]
    hy = g.top
    d.plate(s, col_x(1), hy, span_w(12), T["header_h"], color="figure")
    for (x, w, align, tx), h in zip(cols, heads[:n]):
        d.text(s, "small", tx if align == "left" else x, hy + 7, w - pad, 18,
               str(h), color="ground", align=align, tag="th")
    for i, row in enumerate(rows):
        y = hy + T["header_h"] + i * T["row_h"]
        for (x, w, align, tx), cell in zip(cols, row[:n]):
            d.text(s, "body", tx if align == "left" else x, y + 8, w - pad, 20,
                   str(cell), align=align, color="figure" if align == "left" else "ink2",
                   font_key="head" if align == "right" else None, tag="td")
        if i < len(rows) - 1:
            d.table_rule(s, col_x(1), y + T["row_h"] - 0.5, span_w(12))
    if sl.get("footnote"):
        d.text(s, "micro", col_x(1), g.bottom - 16, span_w(8), 15, sl["footnote"],
               color="muted", tag="footnote")


def table_columns(d, n):
    spans = d.spec["table"]["col_pattern"][n]
    pad = d.spec["table"]["pad_x"]
    out, col = [], 1
    for i, sp in enumerate(spans):
        x, w = col_x(col), span_w(sp)
        out.append((x, w, "left", x + pad) if i == 0 else (x, w, "right", x + w - pad))
        col += sp
    return out


def image_split(d, s, m, sl):
    g = L(d)
    IW = 392
    d.picture(s, sl["image"], 960 - IW, 0, IW, 540, sl.get("fit", "cover"),
              focus=sl.get("focus", "center"))
    g.head(s, sl, span=6)
    body = _paras(sl, "bullets", "body")
    if body:
        bh_ = g.bh("\n".join(body), span_w(6), "body")
        d.text(s, "body", col_x(1), resolve_y(bh_, d.spec), span_w(6), bh_, body,
               color="ink2", tag="body")


def image_full(d, s, m, sl):
    d.picture(s, sl["image"], 0, 0, 960, 540, sl.get("fit", "cover"),
              focus=sl.get("focus", "center"))
    PW, PH = 470, 186
    d.plate(s, 0, 540 - PH, PW, PH, color="ground")
    if sl.get("eyebrow"):
        d.text(s, "micro", 58, 540 - PH + 26, PW - 116, 15, sl["eyebrow"],
               color="accent", tag="eyebrow")
    title = _t(sl, "title")
    th = block_h(title, PW - 116, d.spec["styles"]["h1"], d.spec["fonts"]["head"])
    d.text(s, "h1", 58, 540 - PH + 48, PW - 116, th, title, tag="title")
    if sl.get("caption"):
        d.text(s, "small", 58, 540 - 44, PW - 116, 18, sl["caption"],
               color="muted", tag="caption")


LAYOUTS = {"cover": cover, "closing": closing, "section": section, "chain": chain,
           "statement": statement, "two_col": two_col, "cards": cards,
           "panel_list": panel_list, "card_grid": card_grid,
           "data": data, "quote": quote, "table": table,
           "image_split": image_split, "image_full": image_full}
INVERTED = {"section", "closing"}


def build(outline_path, out_path=None, spec_path=None, embed=False):
    o = yaml.safe_load(open(outline_path, encoding="utf-8"))
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
            raise SystemExit(f"slide {i}: 알 수 없는 layout '{lay}'")
        run = run + 1 if lay == prev else 1
        if run > spec["rhythm"]["max_same_layout_run"]:
            warn(f"slide {i}: '{lay}' {run}연속 — 리듬이 죽는다")
        prev = lay
        s = d.slide(invert=lay in INVERTED or bool(sl.get("invert")))
        mode = LAYOUTS[lay](d, s, m, sl)
        if lay == "cards":
            if mode == prev_mode:
                warn(f"slide {i}: cards 모드 '{mode}' 가 직전 장과 같다")
            prev_mode = mode
        else:
            prev_mode = None

    out = out_path or m.get("output") or "out/deck.pptx"
    pptx, man = d.save(out, embed=embed or bool(m.get("embed_fonts")))
    for b in d.manifest:
        if b["x"] + b["w"] > content_r() + 0.5 and not b["tag"] in ("title", "caption", "eyebrow"):
            warn(f"slide {b['slide']}: '{b['tag']}' 오른변 {b['x']+b['w']:.0f} > {content_r():.0f}")
        if b["y"] + b["h"] > 540.5:
            warn(f"slide {b['slide']}: '{b['tag']}' 캔버스 하단 이탈 ({b['y']+b['h']:.0f})")
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
