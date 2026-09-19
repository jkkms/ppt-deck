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
                     block_h, block_w, line_count, resolve_y, contrast)

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
        self.block = r.get("block_y", r["content_y"])   # 도형 블록 시작 (실측 176)
        self.h = self.bottom - self.top
        self.gap = d.spec["anchors"]["gap_block"]
        self.title_bottom = self.title_y
        self.d = d

    def balance(self, text, w, style):
        """두 줄로 접히는 제목의 줄 길이를 고르게 맞춘다 (CSS text-wrap: balance 의 대응).
        한 줄이 길고 다음 줄에 두 글자만 남는 제목은 그 자체로 조잡해 보인다."""
        t = str(text)
        if "\n" in t:
            return t
        st = self.d.spec["styles"][style]
        fn = self.d.spec["fonts"][st["font"]]
        if line_count(t, w, st, fn) != 2:
            return t
        words = t.split(" ")
        if len(words) < 2:
            return t
        best, gap = None, 1e9
        for k in range(1, len(words)):
            a, b = " ".join(words[:k]), " ".join(words[k:])
            if line_count(a, w, st, fn) > 1 or line_count(b, w, st, fn) > 1:
                continue
            g = abs(block_w(a, st, fn) - block_w(b, st, fn))
            if g < gap:
                best, gap = a + "\n" + b, g
        return best or t

    def head(self, s, sl, title_style="h1", span=9, x=None):
        """눈썹 + 제목. 두 좌표는 모든 장에서 고정이다 — 그래서 훑을 때 눈이 안 흔들린다.
        x 는 사진이 왼쪽에 설 때(image_split)만 옮긴다 — 세로 좌표는 그대로 둔다."""
        d = self.d
        X = col_x(1) if x is None else x
        if sl.get("eyebrow"):
            if sl.get("eyebrow_chip") or d.spec.get("eyebrow_style") == "chip":
                SH = d.spec["shapes"]
                cw = self.bh(sl["eyebrow"], 400, "small") * 0 + \
                    line_count(sl["eyebrow"], 400, d.spec["styles"]["small"],
                               d.spec["fonts"]["body"]) * 0
                w = block_w(sl["eyebrow"], d.spec["styles"]["small"],
                            d.spec["fonts"]["body"]) + 36
                d.chip(s, X, self.eyebrow_y - 6, w, SH["eyebrow_chip_h"],
                       sl["eyebrow"], color="accent", text_color="ground", style="small")
            else:
                d.text(s, "micro", X, self.eyebrow_y, span_w(span), 15,
                       sl["eyebrow"], color="accent", tag="eyebrow")
        if sl.get("runner"):
            d.text(s, "micro", col_x(9), self.eyebrow_y, span_w(4), 15,
                   sl["runner"], color="muted", align="right", tag="runner")
        if sl.get("kind"):
            # 실습·활동 표시. 눈썹이 '어느 장인지'라면 이건 '무엇을 하는 장인지'다
            SH = d.spec["shapes"]
            w = block_w(sl["kind"], d.spec["styles"]["small"],
                        d.spec["fonts"]["body"]) + 34
            d.chip(s, content_r() - w, self.title_y + 4, w, SH["eyebrow_chip_h"],
                   sl["kind"], color="accent", text_color="ground", style="small")
        if sl.get("title"):
            st = d.spec["styles"][title_style]
            t = self.balance(sl["title"], span_w(span), title_style)
            h = block_h(t, span_w(span), st, d.spec["fonts"][st["font"]])
            d.text(s, title_style, X, self.title_y, span_w(span), h,
                   t.split("\n"), tag="title")
            self.title_bottom = self.title_y + h

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
    total = th + (g.gap + mh if mh else 0)
    y_title = resolve_y(total, d.spec, top=g.top)
    y_meta = y_title + th + g.gap
    d.text(s, "cover", col_x(1), y_title, span_w(9), th, title.split("\n"), tag="title")
    if meta:
        d.text(s, "small", col_x(1), y_meta, span_w(6), mh, meta,
               color="muted", tag="meta")


def closing(d, s, m, sl):
    """반전 필드. 두 갈래다.

    ledger 가 없으면 질문 시간(Q&A) 장 — 사용자 원칙 「마무리(Q&A) 슬라이드는 단순하게」.
      알약 라벨 · 큰 제목 · 맺음말 한 줄. 부제·예시 질문 카드는 없다. 글 덩어리는 왼쪽, 세로 가운데.
      그림은 **선택**이다 — 넣을 거리가 있을 때만 오른쪽에 크게 두고 뒤에 조명 원 하나.
      없다고 억지로 채우지 않는다 (사용자 지시, 2026-09-13). 실측 2차시 39쪽.
    ledger 가 있으면 정보 원장이 우측에 서고 제목은 1줄이다.
    """
    g = L(d)
    if not sl.get("ledger"):
        return _closing_qa(d, s, g, sl)
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
    d.text(s, "cover", col_x(1), resolve_y(th, d.spec, top=g.top), span_w(7), th,
           title, tag="title")



def _closing_qa(d, s, g, sl):
    SH = d.spec["shapes"]
    mid = d.spec["canvas"]["height_pt"] / 2
    # 오른쪽 — 캐릭터와 조명 원. 둘은 중심을 같이 쓰고, 그림 오른쪽 끝이 여백선이다
    if sl.get("image"):
        iw, ih = SH["qa_img_w"], SH["qa_img_h"]
        ix = content_r() - iw
        d.glow(s, ix + iw / 2, mid, SH["qa_glow_d"], SH["qa_glow_contrast"])
        d.picture(s, sl["image"], ix, mid - ih / 2, iw, ih, sl.get("fit", "contain"),
                  tag="qa-image")
    else:
        ix = content_r() + SH["qa_text_gap"]   # 그림이 없으면 글이 오른쪽 여백선까지 쓴다
    # 왼쪽 — 알약 + 제목이 한 덩어리로 세로 가운데
    tw = ix - col_x(1) - SH["qa_text_gap"]
    title = _t(sl, "title") or "무엇이든 물어보세요"
    th = g.bh(title, tw, "cover")
    label = sl.get("label") or ""
    st_s = d.spec["styles"]["small"]
    ph = SH["eyebrow_chip_h"] if label else 0
    gap = SH["qa_label_gap"] if label else 0
    top = round(mid - (ph + gap + th) / 2)
    if label:
        pw = block_w(label, st_s, d.spec["fonts"]["body"]) + 2 * SH["qa_label_pad_x"]
        # 알약 글자색은 대비로 고른다 — 반전 필드에서 ground 는 어둡고 accent 도 어두울 수 있다
        tc = max(("ground", "figure"), key=lambda k: contrast(d.c(k), d.c("accent")))
        # 알약 라벨도 같은 반경이다 — 실측 108x30.2 의 반경이 11.5 로, 진짜 알약(15.1)이 아니었다
        d.chip(s, col_x(1), top, pw, ph, label, color="accent", text_color=tc,
               style="small", radius=SH["chip_radius"])
    d.text(s, "cover", col_x(1), top + ph + gap, tw, th, title, tag="title")
    # 맺음말 한 줄은 본문 바닥선에 — 덩어리에 붙이지 않는다
    if sl.get("line"):
        lh = g.bh(sl["line"], span_w(8), "body")
        bottom = (d.spec.get("rhythm_y") or {}).get(
            "content_bottom", d.spec["canvas"]["height_pt"] - d.spec["grid"]["margin_bottom"])
        d.text(s, "body", col_x(1), bottom - lh, span_w(8), lh,
               sl["line"], color="muted", tag="qa-line")
    for k in ("subtitle", "lead", "questions", "items"):
        if sl.get(k):
            warn(f"closing(Q&A): '{k}' 는 넣지 않는다 — 사용자가 직접 지운 요소다. 무시함")

def section(d, s, m, sl):
    """반전 필드. 큰 번호가 아니라 눈썹 + 제목의 위치만으로 장을 가른다.

    `part:` 를 주면 파트 표지가 된다 — 사용자 원칙 §10(배서위 덱에서 배운 점):
    라벨 + 큰 제목 + 짧은 가로선 하나를 **모든 파트에서 같은 좌표**에 둔다.
    좌표가 같아야 장을 넘길 때 "지금 어디쯤인지"가 저절로 읽힌다.
    그래서 여기서만 광학 중심을 쓰지 않고 고정 y 를 쓴다.
    """
    g = L(d)
    if sl.get("part"):
        SH = d.spec["shapes"]
        py = SH["part_label_y"]
        d.text(s, "micro", col_x(1), py, span_w(6), 15, str(sl["part"]),
               color="accent", tag="part-label")
        title = _t(sl, "title")
        th = g.bh(title, span_w(9), "cover")
        ty = SH["part_title_y"]
        d.text(s, "cover", col_x(1), ty, span_w(9), th, title.split("\n"), tag="title")
        d.hero_rule(s, col_x(1), ty + th + SH["part_rule_gap"], SH["part_rule_w"])
        return
    if sl.get("number"):
        d.text(s, "display", col_x(1), g.eyebrow_y, span_w(3), 46,
               str(sl["number"]), color="accent", tag="section-num")
    if sl.get("runner"):
        d.text(s, "micro", col_x(9), g.eyebrow_y, span_w(4), 15, sl["runner"],
               color="muted", align="right", tag="runner")
    title = _t(sl, "title")
    th = g.bh(title, span_w(9), "cover")
    d.text(s, "cover", col_x(1), resolve_y(th, d.spec, top=g.top), span_w(9), th,
           title.split("\n"), tag="title")


def statement(d, s, m, sl):
    g = L(d)
    g.head(s, sl, span=9)
    text = _t(sl, "text")
    th = g.bh(text, span_w(11), "cover")
    body = _paras(sl, "body")
    bh_ = g.bh("\n".join(body), span_w(7), "lead") if body else 0
    total = th + (g.gap + bh_ if bh_ else 0)
    y = resolve_y(total, d.spec, top=g.top)
    acc = sl.get("accent_lines") or []
    d.text(s, "cover", col_x(1), y, span_w(11), th, text.split("\n"),
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
    # 분량이 적을 때 13pt 로 두면 화면이 비어 보인다. 한 단계 키운다.
    bstyle = "lead" if sum(len(b) for b in body) < 190 else "body"
    lh = g.bh(lead, lw, "lead") if lead else 0
    rh = g.bh("\n".join(body), rw, bstyle) if body else 0
    sub, subb = _t(sl, "subhead"), _t(sl, "subbody")
    sh_ = g.bh(sub, rw, "h2") if sub else 0
    sbh = g.bh(subb, rw, bstyle) if subb else 0
    right_total = rh + (g.gap + sh_ if sh_ else 0) + (8 + sbh if sbh else 0)
    # 좌·우 컬럼을 같은 기준선에 세우고, 남는 높이의 38%만 위에 둔다
    # 왼쪽에 리드가 없으면 2단이 아니다 — 빈 컬럼을 남기지 말고 한 단으로 넓게 쓴다
    if not lead:
        rw = span_w(9)
        rh = g.bh("\n".join(body), rw, bstyle) if body else 0
        sh_ = g.bh(sub, rw, "h2") if sub else 0
        sbh = g.bh(subb, rw, bstyle) if subb else 0
        right_total = rh + (g.gap + sh_ if sh_ else 0) + (8 + sbh if sbh else 0)
        col = 1
    else:
        col = 6
    # 리드는 제목 바로 아래(content_y)에 붙인다. 본문과 같은 줄에 두면 허공에 뜬다.
    if lh:
        d.text(s, "lead", col_x(1), max(g.top, g.title_bottom + 14), lw, lh,
               lead, color="muted", tag="lead")
    y = resolve_y(right_total, d.spec, top=g.block)
    if rh:
        d.text(s, bstyle, col_x(col), y, rw, rh, body, color="figure", tag="body")
    if sh_:
        d.text(s, "h2", col_x(col), y + rh + g.gap, rw, sh_, sub, tag="subhead")
        if sbh:
            d.text(s, bstyle, col_x(col), y + rh + g.gap + sh_ + 8, rw, sbh, subb,
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
        h = d.spec["shapes"]["badge_d"] + 12 + g.bh(it.get("title", ""), w, tstyle)
        if it.get("body"):
            h += 8 + g.bh(it["body"], w, "small")
        return h

    SH = d.spec["shapes"]

    tstyle = "h1" if n <= 4 else "h2"

    def cell(it, x, w, y):
        dd = SH["badge_d"]
        d.badge(s, x + dd / 2, y + dd / 2, dd, "accent",
                glyph=str(it.get("index", "")), glyph_color="ground", style="small")
        th = g.bh(it.get("title", ""), w, tstyle)
        d.text(s, tstyle, x, y + dd + 12, w, th, it.get("title", ""), tag="card-title")
        if it.get("body"):
            bh_ = g.bh(it["body"], w, "small")
            d.text(s, "small", x, y + SH["badge_d"] + 12 + th + 8, w, bh_,
                   it["body"], color="muted", tag="card-body")

    # 행은 실제 높이로 위에서부터 쌓는다. 남은 높이에 균등 분배하면 흩어져 보인다.
    def place(total):
        return g.block + max(0.0, (g.bottom - g.block - total) * BIAS)

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
    y = resolve_y(max(hero_total, subs_h), d.spec, top=g.block)
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
    src_h = (18 if sl.get("source") else 0) + (18 if sl.get("note") else 0)
    y = resolve_y(th + 20 + src_h, d.spec, top=g.top)
    y_src = y + th + 20
    y_note = y_src + (18 if sl.get("source") else 0)
    d.text(s, "h1", col_x(2), y, span_w(8), th, text, tag="quote")
    if sl.get("source"):
        d.text(s, "micro", col_x(2), y_src, span_w(5), 15, sl["source"],
               color="accent", tag="quote-source")
    if sl.get("note"):
        d.text(s, "small", col_x(2), y_note, span_w(5), 18, sl["note"],
               color="muted", tag="quote-note")


def chain(d, s, m, sl):
    """전제 패널 + 꼬리를 무는 질문. 소집면담00 12쪽.

    원본과 나란히 그려 보고 다시 옮겼다 (2026-09-13, 사용자: "원본이 훨씬 좋아").
      - 전제와 질문 글은 15.5pt 굵게 — 13pt 로는 사슬이 힘없이 흩어졌다
      - 배지 안은 글리프가 아니라 선 아이콘 그림. 전제는 내용에 맞는 아이콘(panel_icon)
      - 연결선은 accent_pale 가는 선 (원본 E1E7E1 과 같은 값). 진한 선은 도식처럼 튄다
      - 연결선은 늘 같은 길이로 다음 배지 윗변에서 끝난다
      - 질문 배지는 왼쪽 끝을 계단으로 맞춘다 (첫 배지 왼쪽 끝 = 전제 배지 왼쪽 끝)
      - 제목 아래 설명은 11.5pt — 사슬 글보다 작아야 위계가 선다
    """
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    y = g.top
    if sl.get("lead"):
        lh = g.bh(sl["lead"], span_w(10), "small")
        d.text(s, "small", col_x(1), y, span_w(10), lh, sl["lead"], color="muted", tag="lead")
    y = max(y + 30, g.block)

    D, dsub = SH["badge_d"], SH["badge_d_sub"]
    bl = col_x(1) + SH["chain_badge_dx"]      # 배지 왼쪽 끝 (실측 82.1)
    ist = d.spec["styles"]["lead"]
    row_h = ist["size"] * ist["leading"]
    if sl.get("panel"):
        # 한 줄짜리 전제 띠는 얇게 — 줄 높이 + 여백 (원칙 §4-3). 두꺼운 띠는 글이 뜬다
        ph = round(row_h + SH["panel_pad_h"])
        d.panel(s, col_x(1), y, span_w(12), ph, radius=SH["panel_radius"])
        d.badge(s, bl + D / 2, y + ph / 2, D, "accent")
        d.icon(s, sl.get("panel_icon", "quote"), bl + D / 2, y + ph / 2,
               D * SH["icon_ratio"], "ground")
        tx = bl + SH["chain_panel_text_dx"]
        d.text(s, "lead", tx, y + ph / 2 - row_h / 2, col_x(1) + span_w(12) - tx - 16, row_h,
               sl["panel"], font_key="head", anchor="middle", tag="panel-text")
        y += ph

    steps = sl.get("steps") or []
    ind, step = SH["chain_indent"], SH["chain_step"]
    if len(steps) > 4:
        warn(f"slide {d._slide_i}: chain 단계 {len(steps)}개 — 4개까지만 들여쓰기가 화면에 든다")
    for i, tx in enumerate(steps[:4]):
        left = bl + i * ind
        top = y + SH["chain_first_gap"] + i * step
        cx, cy = left + dsub / 2, top + dsub / 2
        line_h = SH["chain_line_h"]
        if i or sl.get("panel"):              # 이을 앞 행이 있을 때만
            d.connector(s, cx, top - line_h, line_h, w=SH["chain_line_w"], color="accent_pale")
        d.badge(s, cx, cy, dsub, "accent_soft")
        d.icon(s, "arrow", cx, cy, dsub * SH["icon_ratio"], "ground")
        d.text(s, "small", left + SH["chain_num_dx"], cy - 9, 22, 18, str(i + 1),
               color="accent_soft", font_key="head", anchor="middle", tag="chain-num")
        x2 = left + SH["chain_text_dx"]
        d.text(s, "lead", x2, cy - row_h / 2, content_r() - x2, row_h, str(tx),
               font_key="head", anchor="middle", tag="chain-text")


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
    y = max(y, g.block)

    px, pw = col_x(1), span_w(12)
    tx = px + SH["panel_text_x"]
    items_ = (sl.get("items") or [])[:4]
    tw_ = pw - SH["panel_text_x"] - SH["panel_pad_x"]
    # 패널 높이를 가장 큰 것에 맞춘다 — 높이가 제각각이면 목록이 흐트러져 보인다
    ph = max(SH["panel_pad_y"] + g.bh(it.get("text", ""), tw_, "lead") + 8
             + len(it.get("subs") or []) * SH["subrow_step"] + SH["panel_pad_y"]
             for it in items_) if items_ else 0
    for it in items_:
        lead_h = g.bh(it.get("text", ""), tw_, "lead")
        subs = it.get("subs") or []
        d.panel(s, px, y, pw, ph, radius=SH["panel_radius"])
        # 칩을 주문장 첫 줄의 세로 중앙에 맞춘다
        line_h = d.spec["styles"]["lead"]["size"] * d.spec["styles"]["lead"]["leading"]
        d.chip(s, px + SH["panel_pad_x"], y + SH["panel_pad_y"] + (line_h - SH["chip_h"]) / 2,
               SH["chip_w"], SH["chip_h"], str(it.get("index", "")),
               color="accent", text_color="ground", style="small")
        d.text(s, "lead", tx, y + SH["panel_pad_y"], tw_, lead_h,
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
    """가로로 늘어선 카드. 높이는 내용이 정하고, 채움은 강조할 카드 하나에만 준다.

    모든 블록에 같은 테두리·채움·반경을 찍으면 위계가 평평해진다 — 카드는
    '따로 떨어진 물체'라는 뜻이라, 전부에 쓰면 아무것도 구별되지 않는다.
    기본은 채움 없이 배지와 여백만으로 가르고, emphasis 로 지정한 하나만 들어올린다.
    """
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = (sl.get("items") or [])[:5]
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"card_grid 는 항목 2~5개만 지원한다 (받은 값: {n}).")
    total = span_w(12)
    cw = (total - SH["card_gap"] * (n - 1)) / n
    pad, bd = SH["card_pad"], SH["badge_d"]
    emph = sl.get("emphasis")
    numbered = any(it.get("index") for it in items)

    inner = cw - pad * 2
    tstyle = "h1" if n <= 4 else "h2"     # 실측 — 5열 카드도 제목 32pt 였다
    th_max = max(g.bh(it.get("title", ""), inner, tstyle) for it in items)
    bh_max = max((g.bh(it["body"], inner, "small") for it in items if it.get("body")),
                 default=0)
    ch = pad + (bd + 14 if numbered else 0) + th_max + (10 + bh_max if bh_max else 0) + pad
    y = g.block + max(0.0, (g.bottom - g.block - ch) * 0.30)

    for i, it in enumerate(items):
        x = col_x(1) + i * (cw + SH["card_gap"])
        lift = (emph is not None and i == emph)
        if lift:
            d.panel(s, x, y, cw, ch, color="figure", radius=SH["panel_radius"])
        yy = y + pad
        if numbered:
            d.badge(s, x + pad + bd / 2, yy + bd / 2, bd,
                    "accent_soft" if lift else "accent",
                    glyph=str(it.get("index", "")),
                    glyph_color="figure" if lift else "ground", style="small")
            yy += bd + 14
        d.text(s, tstyle, x + pad, yy, inner, th_max, it.get("title", ""),
               color="ground" if lift else "figure", tag="card-title")
        yy += th_max + 10
        if it.get("body"):
            d.text(s, "small", x + pad, yy, inner, bh_max, it["body"],
                   color="ground" if lift else "muted", tag="card-body")


def timeline(d, s, m, sl):
    """가로 타임라인. 실측 — Dive 10쪽. 축선 위에 시기, 아래에 제목·설명.
    지나간 마디는 faint, 최근 마디는 accent 로 칠해 현재 위치를 표시한다."""
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    nodes = (sl.get("nodes") or [])[:5]
    n = len(nodes)
    if not 3 <= n <= 5:
        raise SystemExit(f"timeline 은 마디 3~5개만 지원한다 (받은 값: {n}).")
    total = span_w(12)
    nw = total / n * 0.85
    step = (total - nw) / (n - 1)
    body_h = 50 + SH["node_desc_dy"] + 58
    axis_y = g.block + 62
    d.divider(s, col_x(1) + SH["node_d"] / 2, axis_y,
              step * (n - 1), color="faint")
    live = sl.get("live_from", max(1, n - 1))       # 이 마디부터 accent
    for i, nd in enumerate(nodes):
        x = col_x(1) + i * step
        d.badge(s, x + SH["node_d"] / 2, axis_y + 0.5, SH["node_d"],
                "accent" if i + 1 >= live else "faint", kind="node")
        d.text(s, "h2", x, axis_y + SH["node_year_dy"], nw, 28,
               str(nd.get("when", "")), tag="node-year")
        d.text(s, "lead", x, axis_y + SH["node_title_dy"], nw, 24,
               str(nd.get("title", "")), font_key="head", tag="node-title")
        if nd.get("body"):
            bh_ = g.bh(nd["body"], nw, "small")
            d.text(s, "small", x, axis_y + SH["node_desc_dy"], nw, bh_,
                   nd["body"], color="muted", tag="node-body")
    if sl.get("footnote"):
        d.text(s, "micro", col_x(1), g.bottom - 16, span_w(8), 15,
               sl["footnote"], color="muted", tag="footnote")


def stair(d, s, m, sl):
    """단계가 올라갈수록 패널이 넓어지는 계단. 실측 — Dive 26쪽.
    마지막 단만 반전 채움으로 도착점을 표시한다. 우측 라벨은 패널 밖에 고정."""
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = (sl.get("items") or [])[:5]
    n = len(items)
    if not 3 <= n <= 5:
        raise SystemExit(f"stair 는 단계 3~5개만 지원한다 (받은 값: {n}).")
    total = span_w(12)
    ph, step = SH["stair_h"], SH["stair_step"]
    y0 = g.block
    has_label = any(it.get("label") for it in items)
    lab_x = col_x(1) + total * SH["stair_w_max"] + 28 if has_label else None
    for i, it in enumerate(items):
        frac = SH["stair_w_min"] + (SH["stair_w_max"] - SH["stair_w_min"]) * i / (n - 1)
        w = total * frac
        y = y0 + i * step
        last = (i == n - 1)
        d.panel(s, col_x(1), y, w, ph,
                color="figure" if last else "accent_tint", radius=SH["panel_radius"])
        d.badge(s, col_x(1) + SH["stair_pad_x"] + SH["stair_badge_d"] / 2, y + ph / 2,
                SH["stair_badge_d"], "accent_soft" if last else "accent",
                glyph=str(it.get("index", i + 1)),
                glyph_color="figure" if last else "ground", style="lead")
        tc = "ground" if last else "figure"
        bc = "ground" if last else "muted"
        tx_off = SH["stair_pad_x"] + SH["stair_badge_d"] + SH["badge_text_gap"]
        tw = w - tx_off - SH["stair_text_pad_r"]
        d.text(s, "lead", col_x(1) + tx_off, y + 11, tw, 22,
               it.get("title", ""), color=tc, font_key="head", tag="stair-title")
        if it.get("body"):
            d.text(s, "small", col_x(1) + tx_off, y + 38, tw, 18,
                   it["body"], color=bc, tag="stair-body")
        if lab_x and it.get("label"):
            d.text(s, "h2", lab_x, y + ph / 2 - 14,
                   col_x(1) + total - lab_x, 28, it["label"],
                   color="ink2", tag="stair-label")
    if sl.get("footnote"):
        d.text(s, "small", col_x(1), g.bottom - 20, span_w(12), 19,
               sl["footnote"], color="muted", tag="footnote")


def compare(d, s, m, sl):
    """좌우 대비 두 패널. 실측 — Dive 17쪽. 한쪽은 tint, 다른 쪽은 반전 채움으로
    어느 쪽이 답인지 색이 먼저 말한다. 배지 기호(x / o)가 그 판단을 반복한다."""
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    pair = (sl.get("pair") or [])[:2]
    if len(pair) != 2:
        raise SystemExit("compare 는 항목이 정확히 2개여야 한다.")
    ph, gap = SH["compare_h"], SH["compare_gap"]
    pw = (span_w(12) - gap) / 2
    y = g.block
    for i, it in enumerate(pair):
        x = col_x(1) + i * (pw + gap)
        good = bool(it.get("good"))
        d.panel(s, x, y, pw, ph, color="figure" if good else "accent_tint",
                radius=SH["panel_radius"])
        d.badge(s, x + SH["compare_pad_x"] + SH["stair_badge_d"] / 2, y + ph / 2,
                SH["stair_badge_d"], "accent" if good else "accent_soft",
                glyph=it.get("mark", "\u25cb" if good else "\u00d7"),
                glyph_color="figure" if good else "ground", style="lead")
        tx_off = SH["compare_pad_x"] + SH["stair_badge_d"] + SH["badge_text_gap"]
        tx = x + tx_off
        tw = pw - tx_off - 24
        tc = "ground" if good else "figure"
        bc = "ground" if good else "ink2"
        th = g.bh(it.get("title", ""), tw, "h2")
        bh_ = g.bh(it["body"], tw, "body") if it.get("body") else 0
        blk = th + (14 + bh_ if bh_ else 0)
        ty = y + (ph - blk) / 2                 # 제목+본문을 패널 세로 중앙에
        d.text(s, "h2", tx, ty, tw, th, it.get("title", ""), color=tc,
               tag="compare-title")
        if bh_:
            d.text(s, "body", tx, ty + th + 14, tw, bh_, it["body"],
                   color=bc, tag="compare-body")


def nest(d, s, m, sl):
    """포함관계. 실측 — Dive 27쪽. 왼쪽 동심 타원 + 오른쪽 설명 패널 + 결론 칩.
    타원은 바깥에서 안으로 진해진다."""
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    rings = (sl.get("rings") or [])[:3]
    if len(rings) != 3:
        raise SystemExit("nest 는 고리가 정확히 3개여야 한다.")
    lw = span_w(6)
    ox, oy = col_x(1) + 18, g.block
    ow, oh = lw - 36, 277.0
    tones = ["accent_pale", "accent_mid", "accent"]
    for k in range(3):
        f = k / 3.0
        w, h = ow * (1 - f * 0.58), oh * (1 - f * 0.61)
        x = ox + (ow - w) / 2
        y = oy + k * 42
        d.oval(s, x, y, w, h, color=tones[k])
        # 가장 안쪽 고리는 라벨을 세로 중앙에 — 바깥 고리는 위쪽에 얹어 겹침을 피한다
        ly = y + (h - 26) / 2 if k == 2 else y + SH["ring_label_dy"]
        d.text(s, "lead", x, ly, w, 26, str(rings[k].get("name", "")),
               align="center", color="ground" if k == 2 else "figure", tag="ring-label")

    rx, rw = col_x(7), span_w(6)
    phh = 80
    stack = 3 * phh + 2 * 9 + (12 + 40 if sl.get("conclusion") else 0)
    diag_h = oy + 2 * 42 + oh * 0.39 - g.block          # 왼쪽 그림의 실제 높이
    pyy = g.block + max(0.0, (max(diag_h, oh) - stack) / 2)
    for k, r in enumerate(rings):
        y = pyy + k * (phh + 9)
        d.panel(s, rx, y, rw, phh, color="accent_tint", radius=SH["panel_radius"])
        d.text(s, "h2", rx + SH["nest_pad_x"], y + 8, rw - 60, 26,
               str(r.get("name", "")), tag="nest-title")
        if r.get("body"):
            d.text(s, "small", rx + SH["nest_pad_x"], y + 40, rw - 60, 34,
                   r["body"], color="ink2", tag="nest-body")
    if sl.get("conclusion"):
        cy = pyy + 3 * (phh + 9) + 12
        d.panel(s, rx, cy, rw, 40, color="figure", radius=SH["panel_radius"])
        d.text(s, "lead", rx, cy + 10, rw, 22, sl["conclusion"], align="center",
               color="ground", font_key="head", tag="nest-conclusion")


def dw(t: str) -> int:
    """표시 폭. 한글·전각은 두 칸으로 센다 — 주석 칸을 맞추려면 이게 있어야 한다."""
    n = 0
    for ch in str(t):
        n += 2 if ("\u1100" <= ch <= "\u11ff" or "\u3000" <= ch <= "\u303f"
                   or "\u3130" <= ch <= "\u318f" or "\uac00" <= ch <= "\ud7af"
                   or "\uff00" <= ch <= "\uff60") else 1
    return n


def code(d, s, m, sl):
    """코드 블록. 고정폭은 Consolas(Office 동봉이라 Mac·Windows 모두 있다).

    주석은 칸을 맞춘다 — 한글을 두 칸으로 세지 않으면 어긋난다.
    마지막 줄 뒤에 줄바꿈을 붙이지 않는다(빈 문단이 생겨 글이 위로 붙는다).
    """
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    y = g.top
    if sl.get("lead"):
        lh = g.bh(sl["lead"], span_w(10), "lead")
        d.text(s, "lead", col_x(1), y, span_w(10), lh, sl["lead"],
               color="muted", tag="lead")
    y = g.block

    raw = [str(l) for l in (sl.get("code") or [])]
    # 주석 칸은 가장 긴 코드 줄에 맞춘다. 고정 칸보다 긴 줄이 있으면 그 줄만 어긋난다
    longest = max((dw(l.split("#", 1)[0].rstrip()) for l in raw
                   if "#" in l and not l.lstrip().startswith("#")), default=0)
    col = sl.get("comment_col", max(SH["code_comment_col"], longest + 3))
    lines = []
    for l in raw:
        if "#" in l and not l.lstrip().startswith("#"):
            head_, note = l.split("#", 1)
            head_ = head_.rstrip()
            lines.append(head_ + " " * max(1, col - dw(head_)) + "#" + note)
        else:
            lines.append(l)

    st = d.spec["styles"]["code"]
    pw = span_w(12)
    inner = pw - SH["code_pad_x"] * 2
    # 넘침은 눈으로 찾지 않는다 — 빌드할 때 잰다
    widest = max((block_w(l, st, d.spec["fonts"]["mono"]) for l in lines), default=0)
    if widest > inner:
        warn(f"slide {d._slide_i}: 코드가 패널을 넘는다 ({widest:.0f}pt > {inner:.0f}pt) "
             f"— 주석 칸을 줄이거나 줄을 의도적으로 나눠라")
    ch = len(lines) * st["size"] * st["leading"] + SH["code_pad_y"] * 2
    d.panel(s, col_x(1), y, pw, ch, color="figure", radius=SH["panel_radius"])
    d.text(s, "code", col_x(1) + SH["code_pad_x"], y + SH["code_pad_y"], inner,
           ch - SH["code_pad_y"] * 2, lines, color="ground", anchor="middle",
           accent_paras=tuple(i for i, l in enumerate(lines) if l.lstrip().startswith("#")),
           tag="code")
    if sl.get("caption"):
        d.text(s, "small", col_x(1), y + ch + SH["code_caption_gap"], span_w(9), 20,
               sl["caption"], color="muted", tag="caption")


def code_explain(d, s, m, sl):
    """코드 왼쪽, 설명 오른쪽. 2차시 덱의 주력 구성(16~27쪽)이다.

    코드는 어두운 판에 고정폭으로, 설명은 tint 판에 사람 말로. 둘을 같은 높이로 맞춘다.
    아래에 배열 칸(cells)을 두면 '무엇이 어떻게 바뀌는지'가 눈에 보인다.
    """
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    cw, ew = span_w(7), span_w(5)
    cx, ex = col_x(1), col_x(8)

    raw = [str(l) for l in (sl.get("code") or [])]
    # 주석 칸은 가장 긴 코드 줄에 맞춘다. 고정 칸보다 긴 줄이 있으면 그 줄만 어긋난다
    longest = max((dw(l.split("#", 1)[0].rstrip()) for l in raw
                   if "#" in l and not l.lstrip().startswith("#")), default=0)
    col = sl.get("comment_col", max(SH["code_comment_col"], longest + 3))
    lines = []
    for l in raw:
        if "#" in l and not l.lstrip().startswith("#"):
            h_, n_ = l.split("#", 1)
            h_ = h_.rstrip()
            lines.append(h_ + " " * max(1, col - dw(h_)) + "#" + n_)
        else:
            lines.append(l)
    cst = d.spec["styles"]["code"]
    widest = max((block_w(l, cst, d.spec["fonts"]["mono"]) for l in lines), default=0)
    if widest > cw - SH["code_pad_x"] * 2:
        warn(f"slide {d._slide_i}: 코드가 판을 넘는다 — 주석 칸을 줄이거나 줄을 나눠라")

    eh_t = g.bh(sl.get("point", ""), ew - 60, "h2") if sl.get("point") else 0
    eh_b = g.bh(sl.get("explain", ""), ew - 60, "lead") if sl.get("explain") else 0
    ph = max(len(lines) * cst["size"] * cst["leading"] + SH["code_pad_y"] * 2,
             eh_t + (12 + eh_b if eh_b else 0) + 44)
    y = g.block
    d.panel(s, cx, y, cw, ph, color="figure", radius=SH["panel_radius"])
    d.text(s, "code", cx + SH["code_pad_x"], y + (ph - len(lines) * cst["size"]
           * cst["leading"]) / 2, cw - SH["code_pad_x"] * 2,
           len(lines) * cst["size"] * cst["leading"], lines, color="ground",
           accent_paras=tuple(i for i, l in enumerate(lines) if l.lstrip().startswith("#")),
           tag="code")
    d.panel(s, ex, y, ew, ph, radius=SH["panel_radius"])
    ty = y + (ph - (eh_t + (12 + eh_b if eh_b else 0))) / 2
    if eh_t:
        d.text(s, "h2", ex + 30, ty, ew - 60, eh_t, sl["point"], tag="explain-point")
    if eh_b:
        d.text(s, "lead", ex + 30, ty + eh_t + 12, ew - 60, eh_b, sl["explain"],
               color="ink2", tag="explain-body")

    cells = sl.get("cells") or []
    if cells:
        cy = y + ph + 34
        if sl.get("cells_label"):
            d.text(s, "small", cx, cy - 26, span_w(6), 20, sl["cells_label"],
                   color="muted", tag="cells-label")
        cs, gp = SH["chipcell"], SH["chipcell_gap"]
        for i, v in enumerate(cells[:14]):
            d.chip(s, cx + i * (cs + gp), cy, cs, cs, str(v),
                   color="accent_tint", text_color="figure", style="lead")


def cellgrid(d, s, m, sl):
    """자료구조를 칸으로 그려 가르친다. 2차시 23쪽 — 시리즈와 데이터프레임.

    표를 '보여 주는' table 과 다르다. 이건 구조를 '설명하는' 그림이라
    머리 칸만 진한 톤으로 띄우고 나머지는 tint 로 둔다.

    카드 안 배치는 사용자 원칙 §4-2 를 따른다 (사용자가 직접 고쳐 보여 준 배치):
      - 제목·설명은 왼쪽 정렬, 둘을 바짝 붙여 한 덩어리
      - 글 덩어리 위 여백은 표 아래 여백보다 조금 더 (info_pad_top_extra) — 의도다
      - 표는 내용 폭으로 두고 카드 가로 가운데. 글과 같은 기준선에 억지로 맞추지 않는다
      - 나란한 카드는 같은 세로 배치 — 가장 긴 제목·설명·표에 맞춘다
    """
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    blocks = (sl.get("blocks") or [])[:2]
    if not blocks:
        return
    n = len(blocks)
    pw = (span_w(12) - 20) / n if n > 1 else span_w(12)
    ch, gp = SH["cell_h"], SH["cell_gap"]
    pad_x = SH["panel_pad_x"]
    tw_max = pw - 2 * pad_x
    st_s = d.spec["styles"]["small"]
    fn_s = d.spec["fonts"][st_s["font"]]

    # 형제 카드가 같은 배치를 쓰도록 세로 치수는 전부 최댓값으로 맞춘다
    title_h = max(g.bh(b.get("title", ""), tw_max, "h2") for b in blocks)
    desc_h = max((g.bh(b["desc"], tw_max, "small") for b in blocks if b.get("desc")),
                 default=0)
    maxrows = max(len(b.get("rows") or []) for b in blocks)
    grid_h = (maxrows + 1) * ch + maxrows * gp
    pad_b = SH["info_pad_bottom"]
    pad_t = pad_b + SH["info_pad_top_extra"]
    desc_dy = title_h + SH["info_title_gap"]
    table_dy = desc_dy + (desc_h + SH["info_table_gap"] if desc_h else SH["info_table_gap"])
    ph = pad_t + table_dy + grid_h + pad_b

    y = g.block
    for k, b in enumerate(blocks):
        x = col_x(1) + k * (pw + 20)
        d.panel(s, x, y, pw, ph, radius=SH["panel_radius"])
        d.text(s, "h2", x + pad_x, y + pad_t, tw_max, title_h, b.get("title", ""),
               tag="cg-title")
        if b.get("desc"):
            d.text(s, "small", x + pad_x, y + pad_t + desc_dy, tw_max, desc_h, b["desc"],
                   color="muted", tag="cg-desc")
        heads = b.get("headers") or []
        rows = b.get("rows") or []
        ncol = max(len(heads), max((len(r) for r in rows), default=0))
        if not ncol:
            continue
        widths = (b.get("widths") or [1] * ncol)[:ncol]
        widths += [1] * (ncol - len(widths))
        # 열 폭 = 비율 x 단위, 단 가장 긴 칸 글자는 들어가야 한다
        ws = []
        for j in range(ncol):
            cells = [str(heads[j])] if j < len(heads) else []
            cells += [str(r[j]) for r in rows if j < len(r)]
            need = max((block_w(c, st_s, fn_s) for c in cells), default=0) + 20
            ws.append(round(max(widths[j] * SH["info_col_unit"], need)))
        tw = sum(ws) + gp * (ncol - 1)
        if tw > tw_max:                       # 넘치면 카드 폭 안으로 비례 축소
            k_ = (tw_max - gp * (ncol - 1)) / sum(ws)
            ws = [w_ * k_ for w_ in ws]
            tw = tw_max
        gx0 = round(x + (pw - tw) / 2)        # 표는 카드 가로 가운데 (정수 pt — 칸 글자 중앙 오차 방지)
        gy = y + pad_t + table_dy
        for j, hcell in enumerate(heads[:ncol]):
            gx = gx0 + sum(ws[:j]) + gp * j
            d.chip(s, gx, gy, ws[j], ch, str(hcell), color="accent_mid",
                   text_color="figure", style="small", radius=SH["chip_radius"])
        for i, row in enumerate(rows):
            gy2 = gy + (i + 1) * (ch + gp)
            for j, cell in enumerate(row[:ncol]):
                gx = gx0 + sum(ws[:j]) + gp * j
                d.chip(s, gx, gy2, ws[j], ch, str(cell), color="accent_tint",
                       text_color="figure", style="small", radius=SH["chip_radius"])


def gallery(d, s, m, sl):
    """같은 것을 여러 방식으로 보여 주는 장. 2차시 30쪽 — 네 가지 그래프.

    한 칸 = 그림 하나 + 이름 + 코드 + 한 줄 설명. 개념 하나에 시각 자료 하나라는
    원칙을 n 칸으로 반복한 것이다.
    """
    g = L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = (sl.get("items") or [])[:4]
    n = len(items)
    if not 2 <= n <= 4:
        raise SystemExit(f"gallery 는 항목 2~4개만 지원한다 (받은 값: {n}).")
    gap = SH["gallery_gap"]
    pw = (span_w(12) - gap * (n - 1)) / n
    imh = SH["gallery_img_h"]
    ph = 36 + imh + 30 + 26 + 22 + 22 + 20
    y = g.block
    for i, it in enumerate(items):
        x = col_x(1) + i * (pw + gap)
        d.panel(s, x, y, pw, ph, radius=SH["panel_radius"])
        if it.get("image"):
            d.picture(s, it["image"], x + 14, y + 36, pw - 28, imh,
                      it.get("fit", "contain"), tag="gallery-img")
        ty = y + 36 + imh + 30
        d.text(s, "h2", x + 30, ty, pw - 60, 26, it.get("title", ""), tag="gallery-title")
        if it.get("code"):
            d.text(s, "small", x + 30, ty + 34, pw - 60, 20, it["code"],
                   font_key="mono", color="accent", tag="gallery-code")
        if it.get("body"):
            d.text(s, "small", x + 30, ty + 64, pw - 60, 20, it["body"],
                   color="muted", tag="gallery-body")


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
    hy = T.get("header_y", g.block)
    avail = g.bottom - hy - T["header_h"] - (26 if sl.get("footnote") else 0)
    row_h = min(52.0, max(T["row_h"], avail / max(1, len(rows))))
    d.plate(s, col_x(1), hy, span_w(12), T["header_h"], color="figure")
    for (x, w, align, tx), h in zip(cols, heads[:n]):
        # 머리 라벨도 판 세로 중앙에
        d.text(s, "small", tx if align == "left" else x, hy, w - pad, T["header_h"],
               str(h), color="ground", align=align, anchor="middle",
               exact_center=True, tag="th")
    for i, row in enumerate(rows):
        y = hy + T["header_h"] + i * row_h
        for (x, w, align, tx), cell in zip(cols, row[:n]):
            # 셀 글씨를 행 높이 안에서 세로 중앙에 — 행마다 눈높이가 흔들리지 않게
            d.text(s, "body", tx if align == "left" else x, y, w - pad, row_h,
                   str(cell), align=align, anchor="middle", exact_center=True,
                   color="figure" if align == "left" else "ink2",
                   font_key="head" if align == "right" else None, tag="td")
        if i < len(rows) - 1:
            d.table_rule(s, col_x(1), y + row_h - 0.5, span_w(12))
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
    """사진 반쪽 + 글 반쪽. 크기와 위치는 고정하고 **좌우만 번갈아** 쓴다 (원칙 §10).
    규칙은 유지한 채 화면이 지루해지지 않는다. `side: left|right` 로 직접 정할 수도 있다."""
    g = L(d)
    IW = 392
    side = sl.get("side", "auto")
    if side == "auto":
        side = "left" if getattr(d, "_img_side", "left") == "right" else "right"
    d._img_side = side
    ix = 0 if side == "left" else 960 - IW
    tx = col_x(1) + (IW if side == "left" else 0)
    d.picture(s, sl["image"], ix, 0, IW, 540, sl.get("fit", "cover"),
              focus=sl.get("focus", "center"))
    g.head(s, sl, span=6, x=tx)
    body = _paras(sl, "bullets", "body")
    if body:
        bh_ = g.bh("\n".join(body), span_w(6), "body")
        d.text(s, "body", tx, resolve_y(bh_, d.spec, top=g.top), span_w(6), bh_, body,
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
           "timeline": timeline, "stair": stair, "compare": compare, "nest": nest,
           "code": code, "code_explain": code_explain,
           "cellgrid": cellgrid, "gallery": gallery,
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
        s = d.slide(invert=lay in INVERTED or bool(sl.get("invert")),
                    tone=sl.get("tone"))
        mode = LAYOUTS[lay](d, s, m, sl)
        d.notes(s, sl.get("notes"))     # 자세한 설명은 노트로
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
