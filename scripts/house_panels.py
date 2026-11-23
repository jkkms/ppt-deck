"""house 시리즈 확장 2 — 판 카드 레이아웃(pcards · prows · plist · ppanel · pgrid · pcol · pfig · orig · ask · ctimeline).

처음엔 모의 면접 시리즈(2026-10-18 ~ 11-18, 열한 덱)의 강의 폴더 `ext_mock.py` 였다. 2026-11-23 스킬로 올렸다.
house_ext.py 의 subtitle · 쪽 번호 · 서명 · 결론 띠를 그대로 쓰고, 판 카드 레이아웃을 더한다.
사용자 지적(10/18): 「디자인이 구린 것이 많다」 → 옛 기출 덱처럼 옅은 초록 판 안에 글을 담는다.

- pcards : 가로로 나란한 판 카드 2~5개. 배지(번호·아이콘·기호) → 제목 → 설명. 카드 높이는 가장 긴 내용에 맞춘다.
           덩어리 전체(띠 포함)를 제목 아래 영역의 세로 가운데에 둔다.
- prows  : 세로로 쌓은 판 줄 2~5개. 배지 · 제목(왼쪽) · 설명(오른쪽). 글은 줄 판 높이 그대로의 상자에 넣고 **앱이 세로 가운데**를 잡는다.
- ask    : 질문 장 — 판 없이(31쪽 「처음과 끝」처럼) 「질문 N · 출처」 한 줄 + 큰 질문. 묶음을 **슬라이드 세로 정중앙**에.
           문장마다 새 줄로 시작하고 문장 안은 자연스럽게 넘긴다(사용자이 고친 방식, 10/18).
- ctimeline : 타임라인을 본문 영역 세로 가운데에(엔진 timeline 은 블록 위에 붙는다).
- pfig   : 풀이 그림 + 결론 줄 · 판 없는 목록(그림 옆) 또는 결론 줄 + 아래 그림(fig_side: below) — 04부터.

판 안 글의 줄 간격(10/18 실측으로 정함)
  파워포인트·키노트는 줄 간격 배수를 글꼴 고유의 한 줄 높이(Pretendard 1.193em)에 곱한다. 엔진 모형(크기 × 배수)과 달라
  줄이 많을수록 판 아래로 처졌다. 그래서 판 안 글은 **행간 1.0(한 줄 상자 1.193em) + 문단 사이 간격**으로 쌓는다.
  행간 1.0 이면 한글 잉크가 줄 상자 정가운데에 앉으므로(deckkit 배지와 같은 원리) 앱마다 달라지지 않는다.
  줄 사이 간격 = 크기 × 배수 − 1.193 × 크기 → 줄 간격은 모형과 같고, 덩어리 높이 = 1.193 × 크기 + (줄 수 − 1) × 크기 × 배수.

사용:  python scripts/house_panels.py outline.yaml -o out/deck.pptx
"""
from __future__ import annotations

import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import house_ext as E  # noqa: E402

B = E.B
from deckkit import col_x, span_w, content_r, wrap_lines, WORD_WRAP, _apply_font, _end_para  # noqa: E402
from pptx.util import Pt  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN  # noqa: E402

LB = (1950 + 494) / 2048      # Pretendard hhea 오름 + 내림 / em — 행간 1.0 의 한 줄 상자
INK_EM = 0.155                # 행간 1.0 줄 상자에서 한글 잉크 위·아래로 남는 몫(각각, em)
DIG_EM = (1950 - 1468) / 2048  # 숫자는 한글보다 낮게 시작한다 — 줄 상자 위에서 숫자 잉크까지(0.235em, Pretendard 실측)


def _bottom(d):
    return (d.spec.get("rhythm_y") or {}).get("content_bottom", 495)


def _mid_y(d, top, total):
    """본문 영역(top ~ 하한)의 세로 가운데."""
    return top + max(0.0, (_bottom(d) - top - total) / 2)


def _top(d, g, sl):
    """덩어리를 놓기 시작할 수 있는 가장 위 — 제목·부제 아래."""
    t = E._title_bottom(d, sl, sl.get("layout"))
    if sl.get("subtitle"):
        t = max(g.top, t + 10) + E._h(d, sl["subtitle"], span_w(10), "small")
    return max(g.block - 20, t + 28)


# ---------------------------------------------------------------- 판 안 글 덩어리
def _st(d, style):
    return d.spec["styles"][style]


def _lb(d, style):
    return _st(d, style)["size"] * LB


def _pitch(d, style, leading=None):
    st = _st(d, style)
    return st["size"] * (leading or st["leading"])


def _gap(d, style, leading=None):
    return max(0.0, _pitch(d, style, leading) - _lb(d, style))


def _wrap(d, text, w, style):
    """문단(문장)마다 어절 단위로 접는다 — 키노트가 한글을 글자 단위로 끊는 것을 막으려고 줄을 미리 나눈다."""
    st = _st(d, style)
    font = d.spec["fonts"][st["font"]]
    paras = text if isinstance(text, list) else str(text).split("\n")
    out = []
    for p_ in paras:
        p_ = str(p_)
        if ("[" in p_ and "]" in p_) or "**" in p_:    # 강조 표시([강조색] · **굵게**)가 있는 줄 — 한 줄에 들어가면 그대로(표시는 폭에 안 셈)
            plain = re.sub(r"\*\*|[\[\]]", "", p_)
            if len(wrap_lines(plain, w, st, font)) == 1:
                out.append(p_)
                continue
            p_ = plain
        out.extend(wrap_lines(p_, w, st, font) if WORD_WRAP["on"] else [p_])
    return out


def _wrap_no_orphan(d, text, w, style, short=0.4, sep=False):
    """문장마다 자연스럽게 접되(사용자이 고친 방식), 마지막 줄이 폭의 1/3 보다 짧으면 그 문장만 줄 길이를 고르게 다시 나눈다.
    「…바뀌었다고 / 썼습니다.」처럼 한 어절만 넘어가는 것을 막는다. 문장 안을 억지로 반반 가르지는 않는다."""
    from deckkit import advance
    import functools
    st = _st(d, style)
    font = d.spec["fonts"][st["font"]]
    adv = lambda x: sum(advance(ch, st["size"], st["tracking"], font) for ch in x)
    lim = w - max(4.0, w * 0.03)
    out = []
    for sent in str(text).split("\n"):
        # 「한 권」「두 가지」처럼 수와 단위는 갈라 놓지 않는다(붙임 공백)
        sent = re.sub(r"(한|두|세|네|다섯) (권|개|번|가지|명|문장|곳)", "\\1\u00a0\\2", sent)
        lines = wrap_lines(sent, w, st, font)
        # sep=True(판 안 설명) — 두 줄이면 고아 줄이 없어도 구분 기호 자리를 먼저 본다(「무엇이 / 달라졌는지가」 대신 「했고, / 그래서 …」)
        if len(lines) < 2 or (adv(lines[-1]) >= short * lim and not (sep and len(lines) == 2)):
            out.extend(lines)
            continue
        # 두 줄이면 먼저 구분 기호( · , — ) 자리에서 끊어 본다 — 구 한가운데를 가르지 않으려고
        if len(lines) == 2:
            cand = []
            for mt in re.finditer(r"( · |, | — )", sent):
                a_ = sent[:mt.start()] + mt.group(1).rstrip()
                b_ = sent[mt.end():]
                if adv(a_) <= lim and adv(b_) <= lim:
                    cand.append((max(adv(a_), adv(b_)), a_, b_))
            if cand:
                _, a_, b_ = min(cand)
                out.extend([a_, b_])
                continue
            # 구분 기호가 없으면 이음 어미(~면 · ~고 · ~데 · ~서 · ~며 · ~지만) 뒤에서 끊어 본다 — 「음수가 / 되면」 대신 「되면 / 그 …」
            ws = sent.split(" ")
            for k in range(1, len(ws)):
                if re.search(r"(면|고|데|서|며|지만|데도|어도|아도|다가|어|아|여|워|해|혀|쳐|져|려|켜)$", ws[k - 1]):   # ~어 · ~아 꼴 이음(식혀 · 녹여 · 바꿔)
                    a_, b_ = " ".join(ws[:k]), " ".join(ws[k:])
                    if adv(a_) <= lim and adv(b_) <= lim and min(adv(a_), adv(b_)) >= 0.35 * lim:
                        cand.append((max(adv(a_), adv(b_)), a_, b_))
            if cand:
                _, a_, b_ = min(cand)
                out.extend([a_, b_])
                continue
            # 그래도 없으면 조사(을 · 를 · 은 · 는 · 이 · 가 · 에 · 에서 · 로 · 와 · 과) 뒤 — 구와 구 사이(10/25 03:
            # 「시스템 / 엔지니어로」 대신 「엔지니어로 / 진로가 …」, 「어떻게 / 보였는지」 대신 「것을 / 어떻게 …」)
            for k in range(1, len(ws)):
                if re.search(r"(을|를|은|는|이|가|에|에서|에게|로|와|과)[’'」』]?$", ws[k - 1]):
                    a_, b_ = " ".join(ws[:k]), " ".join(ws[k:])
                    if adv(a_) <= lim and adv(b_) <= lim and min(adv(a_), adv(b_)) >= 0.35 * lim:
                        cand.append((max(adv(a_), adv(b_)), a_, b_))
            if cand:
                _, a_, b_ = min(cand)
                out.extend([a_, b_])
                continue
        words = sent.split(" ")
        n, W_ = len(lines), len(words)

        @functools.lru_cache(None)
        def best(i, k):
            if k == 1:
                ww = adv(" ".join(words[i:]))
                return (ww, ()) if ww <= lim else (1e9, ())
            res = (1e9, ())
            for j in range(i + 1, W_ - k + 2):
                ww = adv(" ".join(words[i:j]))
                if ww > lim:
                    break
                m_, cuts = best(j, k - 1)
                if max(ww, m_) < res[0]:
                    res = (max(ww, m_), (j,) + cuts)
            return res
        m_, cuts = best(0, n)
        if m_ >= 1e9:
            out.extend(lines)
            continue
        idx = (0,) + cuts + (W_,)
        out.extend(" ".join(words[a:b]) for a, b in zip(idx, idx[1:]))
    return out


def block_h(d, groups, w):
    h = 0.0
    for i, g_ in enumerate(groups):
        n = len(_wrap(d, g_["text"], w, g_["style"]))
        h += (g_.get("before", 0) if i else 0) + _lb(d, g_["style"]) + (n - 1) * _pitch(d, g_["style"], g_.get("leading"))
    return h


def block(d, s, x, y, w, h, groups, *, anchor="middle", tag="card"):
    """여러 스타일 문단을 한 글상자에 — 행간 1.0, 줄 사이는 문단 앞 간격. anchor="middle" 이면 앱이 세로 가운데를 잡는다.
    groups: [{style, text(str|list), color, before(앞 그룹과의 간격 pt), leading(줄 간격 배수, 선택)}]
    미리보기·검사용 매니페스트에는 그룹마다 모형 위치를 적는다."""
    box = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}.get(anchor, MSO_ANCHOR.TOP)
    H = block_h(d, groups, w)
    cur = y + {"middle": (h - H) / 2, "bottom": h - H}.get(anchor, 0)
    first = True
    for gi, g_ in enumerate(groups):
        style = g_["style"]
        st = _st(d, style)
        font = d.spec["fonts"][st["font"]]
        col = d.c(g_.get("color", "figure"))
        lines = _wrap(d, g_["text"], w, style)
        if gi:
            cur += g_.get("before", 0)
        gy = cur
        for li, ln in enumerate(lines):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.0
            if not first:
                p.space_before = Pt(g_.get("before", 0) if li == 0 else _gap(d, style, g_.get("leading")))
            # [강조색] · **굵게**(2026-11-18 사용자가 마지막 정리 덱 「입실 시간보다 30분 일찍 도착」의 「30분」을 직접 굵게 함)
            segs = re.split(r"(\[[^\]]+\]|\*\*[^*]+\*\*)", ln) if (("[" in ln and "]" in ln) or "**" in ln) else [ln]
            for sg in segs:
                if not sg:
                    continue
                r = p.add_run()
                lit = sg.startswith("[") and sg.endswith("]")
                bold = len(sg) > 4 and sg.startswith("**") and sg.endswith("**")
                r.text = sg[1:-1] if lit else sg[2:-2] if bold else sg
                _apply_font(r, font, st["size"], d.c("accent") if lit else col, bold=bold, tracking=st.get("tracking", 0))
            _end_para(p, font, st["size"], col)
            first = False
        pitch = _pitch(d, style, g_.get("leading"))
        gh = _lb(d, style) + (len(lines) - 1) * pitch
        cur += gh
        d.manifest.append(dict(slide=d._slide_i, tag=f"{tag}-{gi}", style=style, x=x, y=gy, w=w, h=gh,
                               size=st["size"], leading=pitch / st["size"], space_after=0, extra_pt=0.0,
                               exact_center=False, font=font, tracking=st.get("tracking", 0), color=col,
                               align="left", anchor="top", suffix=None, suffix_style=None, suffix_color=col,
                               text=[re.sub(r"\*\*|[\[\]]", "", x) for x in lines]))
    return box


def _num_label(it, i):
    """번호는 원 배지 대신 큰 글자 「1.」(사용자이 17쪽에서 직접 바꿈 — 「원 안 숫자는 답답하고 같은 모양이 이어지면 짜친다」, 10/18).
    index 를 따로 준 것(01, 02 …)은 그대로 쓴다."""
    idx = it.get("index")
    return str(idx) if idx not in (None, "") else f"{i + 1}."


# ================================================================ pcards
def pcards(d, s, m, sl):
    """판 카드 — 같은 줄 카드는 같은 높이(사용자이 직접 맞춤, 10/18), 제목과 설명은 따로 글상자(고치기 쉽게)."""
    g = B.L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"pcards 는 2~5칸 (받은 값 {n})")
    W = span_w(12)
    gap = 24 if n <= 3 else 16
    cw = (W - gap * (n - 1)) / n
    padx = sl.get("pad_x", 32 if n <= 3 else 24)     # 여백 넉넉하게(사용자 「박스 안에 여백이 너무 없어」)
    pady = sl.get("pad_y", 30 if n <= 3 else 26)
    tw = cw - 2 * padx
    bd = SH["badge_d"]
    show_badge = sl.get("badges", True)
    tstyle = sl.get("title_style", "h2")
    bstyle = sl.get("body_style", "lead" if n <= 3 else ("body" if n == 4 else "small"))
    emph = sl.get("emphasis")
    tb_gap = 14

    def tg(it, lift=False):
        return [dict(style=tstyle, text=E._balance(d, it.get("title", ""), tw, tstyle),
                     color="ground" if lift else "figure")]

    def bg(it, lift=False):
        return [dict(style=bstyle, text=it["body"], color="ground" if lift else "ink2",
                     leading=sl.get("body_leading"))]

    numbered = bool(sl.get("numbered"))
    num_h = _lb(d, "h1")

    num_up = DIG_EM * _st(d, "h1")["size"]

    def head_h():
        if numbered:
            return num_h + 10 - num_up
        return (bd + 16) if show_badge else 0

    def content_h(it):
        h = head_h() + block_h(d, tg(it), tw)
        if it.get("body"):
            h += tb_gap + block_h(d, bg(it), tw)
        return h

    last = bstyle if any(it.get("body") for it in items) else tstyle
    pad_b = pady - INK_EM * _st(d, last)["size"]
    top_up = 0.0 if (numbered or show_badge) else INK_EM * _st(d, tstyle)["size"]
    ch = max(content_h(it) for it in items) + pady + pad_b - top_up
    top_h, bot_h, foot_h, vgap = E._extras_h(d, sl, W)
    y = _mid_y(d, _top(d, g, sl), top_h + ch + bot_h + foot_h)
    y = E._draw_extras_top(d, s, sl, y, W)
    for i, it in enumerate(items):
        x = col_x(1) + i * (cw + gap)
        lift = emph is not None and i == emph
        d.panel(s, x, y, cw, ch, color="accent" if lift else "accent_tint")
        cy = y + pady - top_up
        if numbered:
            cy -= num_up
            block(d, s, x + padx, cy, tw, num_h, [dict(style="h1", text=_num_label(it, i),
                  color="ground" if lift else "figure")], anchor="top", tag="card")
            cy += num_h + 10
        elif show_badge:
            E._badge(d, s, it, i, x + padx + bd / 2, cy + bd / 2, sl, lift=lift)
            cy += bd + 16
        th = block_h(d, tg(it), tw)
        block(d, s, x + padx, cy, tw, th, tg(it, lift), anchor="top", tag="card")
        if it.get("body"):
            by = cy + th + tb_gap
            # 가장 긴 칸은 설명을 아래 끝 기준으로 — 키노트가 문단 간격을 덜 줘도 아래 여백이 정확하다(차이는 제목 아래 간격이 흡수).
            # 짧은 칸은 위 기준(제목 바로 아래) — 같은 높이 카드라 아래가 남는 것은 사용자이 고른 방식
            full = abs(content_h(it) - (ch - pady - pad_b + top_up)) < 1.0
            block(d, s, x + padx, by, tw, y + ch - pad_b - by, bg(it, lift),
                  anchor="bottom" if full else "top", tag="card")
    E._draw_extras_bottom(d, s, sl, y + ch, W)


# ================================================================ prows
def prows(d, s, m, sl):
    g = B.L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"prows 는 2~5줄 (받은 값 {n})")
    W = span_w(12)
    x0 = col_x(1)
    padx = 24
    pady = {2: 22, 3: 20, 4: 16, 5: 12}[n]
    gv = {2: 24, 3: 22, 4: 18, 5: 15}[n]      # 줄 판 사이(사용자 「박스 간 간격이 좁아 답답」, 10/18)
    bd = SH["badge_d"]
    show_badge = sl.get("badges", True)
    stack = bool(sl.get("stack"))
    tstyle = sl.get("title_style", "h2")
    bstyle = sl.get("body_style", "lead" if n <= 3 else "body")
    tx = x0 + padx + ((bd + SH["badge_text_gap"]) if show_badge else 0)
    split = col_x(int(sl.get("body_col", 6)))          # 설명이 시작하는 열
    if stack:
        tw = content_r() - padx - tx
        bx, bw = tx, tw
    else:
        tw = split - 20 - tx
        bx, bw = split, content_r() - padx - split
    emph = sl.get("emphasis")

    def tg(it, lift=False):
        return [dict(style=tstyle, text=E._balance(d, it.get("title", ""), tw, tstyle),
                     color="ground" if lift else "figure")]

    def bg(it, lift=False, before=0):
        return [dict(style=bstyle, text=it["body"], color="ground" if lift else "ink2",
                     before=before, leading=sl.get("body_leading"))] if it.get("body") else []

    def rh(it):
        if stack:
            inner = block_h(d, tg(it) + bg(it, before=6), tw)
        else:
            inner = max(block_h(d, tg(it), tw), block_h(d, bg(it), bw) if it.get("body") else 0)
        return max(inner, bd if show_badge else 0) + 2 * pady

    hs = [rh(it) for it in items]
    top_h, bot_h, foot_h, vgap = E._extras_h(d, sl, W)
    total = top_h + sum(hs) + gv * (n - 1) + bot_h + foot_h
    y = _mid_y(d, _top(d, g, sl), total)
    y = E._draw_extras_top(d, s, sl, y, W)
    for i, (it, h) in enumerate(zip(items, hs)):
        lift = emph is not None and i == emph
        d.panel(s, x0, y, W, h, color="accent" if lift else "accent_tint")
        if show_badge and sl.get("numbered"):
            block(d, s, x0 + padx, y, bd + SH["badge_text_gap"] - 4, h,
                  [dict(style="h1", text=_num_label(it, i), color="ground" if lift else "figure")], tag="row")
        elif show_badge:
            E._badge(d, s, it, i, x0 + padx + bd / 2, y + h / 2, sl, lift=lift)
        if stack:
            th_ = block_h(d, tg(it), tw)
            inner = block_h(d, tg(it) + bg(it, before=6), tw)
            ty = y + (h - inner) / 2
            block(d, s, tx, ty, tw, th_, tg(it, lift), anchor="top", tag="row")
            if it.get("body"):
                block(d, s, bx, ty + th_ + 6, bw, inner - th_ - 6, bg(it, lift), anchor="top", tag="row")
        else:
            block(d, s, tx, y, tw, h, tg(it, lift), tag="row")
            if it.get("body"):
                block(d, s, bx, y, bw, h, bg(it, lift), tag="row")
        y += h + gv
    E._draw_extras_bottom(d, s, sl, y - gv, W)


# ================================================================ ask
def ask(d, s, m, sl):
    """질문 장 — 판 없이 「질문 N · 출처」 + 큰 질문. 묶음을 슬라이드 세로 정중앙에(사용자 지시 10/18)."""
    g = B.L(d)
    g.head(s, sl, span=12)
    H = d.spec["canvas"]["height_pt"]
    src = str(sl.get("source", ""))
    num = sl.get("number")
    mt = re.match(r"질문\s*(\d+)\s*·\s*(.*)", src)
    if mt and num is None:
        num, src = mt.group(1), mt.group(2)
    label = f"질문 {num} · {src}" if num else src
    W = span_w(12)
    style = sl.get("text_style", "h1")
    groups = [dict(style="small", text=label, color="accent"),
              dict(style=style, text=_wrap_no_orphan(d, sl["text"], W, style), color="figure", before=18)]
    bh = max(300.0, block_h(d, groups, W) + 20)
    block(d, s, col_x(1), H / 2 - bh / 2, W, bh, groups, anchor="middle", tag="ask")


# ================================================================ ctimeline
def ctimeline(d, s, m, sl):
    """엔진 timeline 과 같은 마디·글 배치. 축 높이만 본문 영역 세로 가운데로."""
    g = B.L(d)
    SH = d.spec["shapes"]
    g.head(s, sl, span=12)
    nodes = (sl.get("nodes") or [])[:5]
    n = len(nodes)
    if not 3 <= n <= 5:
        raise SystemExit(f"ctimeline 은 마디 3~5개 (받은 값 {n})")
    total = span_w(12)
    nw = total / n * 0.85
    step = (total - nw) / (n - 1)
    body_h = max((E._h(d, nd.get("body", ""), nw, "small") for nd in nodes), default=0)
    up, down = -SH["node_year_dy"], SH["node_desc_dy"] + body_h
    foot = 40 if sl.get("footnote") else 0
    top = _mid_y(d, _top(d, g, sl), up + down + foot)
    axis_y = top + up
    d.divider(s, col_x(1) + SH["node_d"] / 2, axis_y, step * (n - 1), color="faint")
    live = sl.get("live_from", max(1, n - 1))
    for i, nd in enumerate(nodes):
        xx = col_x(1) + i * step
        d.badge(s, xx + SH["node_d"] / 2, axis_y + 0.5, SH["node_d"],
                "accent" if i + 1 >= live else "faint", kind="node")
        d.text(s, "h2", xx, axis_y + SH["node_year_dy"], nw, 28, str(nd.get("when", "")), tag="node-year")
        d.text(s, "lead", xx, axis_y + SH["node_title_dy"], nw, 24, str(nd.get("title", "")),
               font_key="head", tag="node-title")
        if nd.get("body"):
            bh = E._h(d, nd["body"], nw, "small")
            d.text(s, "small", xx, axis_y + SH["node_desc_dy"], nw, bh, nd["body"],
                   color="muted", tag="node-body")
    if sl.get("footnote"):
        d.text(s, "small", col_x(1), axis_y + down + 24, span_w(10), 18, sl["footnote"],
               color="muted", tag="foot")


# ================================================================ 변주(10/18 사용자 「같은 디자인을 계속 반복하면 짜쳐 — 스타일은 동일하되 디자인에 변주」)
# 색 · 글꼴 · 크기 · 반경 · 여백은 그대로 두고 짜임만 바꾼다.
#   plist  : 판 없는 목록 — 번호(초록 「1.」)·아이콘 + 제목 + 설명. 줄 사이는 여백으로만 가른다
#   ppanel : 큰 판 하나 안에 여러 줄(같은 열의 상자는 한 상자로 — 원칙 §4)
#   pgrid  : 2 × 2 판 — 머리(번호·아이콘)는 왼쪽, 제목과 설명은 오른쪽에 쌓는다
#   pcol   : 판 없는 단 — 번호·아이콘 → 제목 → 설명을 위에서 아래로
#   pside  : 왼쪽 진초록 판에 핵심 한 문장, 오른쪽에 판 없는 목록
def _head_kind(sl, items):
    if sl.get("numbered"):
        return "num"
    if any(it.get("label") for it in items):
        return "label"
    if sl.get("badges") is False:
        return None
    if sl.get("icon") or any(it.get("icon") or it.get("mark") for it in items):
        return "badge"
    return None


def _head_w(d, kind, items):
    SH = d.spec["shapes"]
    if kind == "num":
        from deckkit import advance
        st = _st(d, "h1")
        font = d.spec["fonts"][st["font"]]
        w = max(sum(advance(ch, st["size"], st["tracking"], font) for ch in _num_label(it, i))
                for i, it in enumerate(items))
        return w + 18
    if kind == "badge":
        return SH["badge_d"] + SH["badge_text_gap"]
    if kind == "label":
        from deckkit import advance
        st = _st(d, "h2")
        font = d.spec["fonts"][st["font"]]
        return max(sum(advance(ch, st["size"], st["tracking"], font) for ch in str(it.get("label", ""))) for it in items) + 20
    return 0.0


def _head(d, s, kind, it, i, sl, x, cy, *, color="accent", lift=False):
    """머리 하나 — 중심 높이 cy 에. 번호는 판 없는 큰 글자, 아이콘·기호는 원 배지."""
    SH = d.spec["shapes"]
    if kind == "num":
        h = _lb(d, "h1")
        block(d, s, x, cy - h / 2, _head_w(d, kind, [it]) + 8, h,
              [dict(style="h1", text=_num_label(it, i), color="ground" if lift else color)], tag="row")
    elif kind == "badge":
        E._badge(d, s, it, i, x + SH["badge_d"] / 2, cy, dict(sl, numbered=False), lift=lift)
    elif kind == "label":
        h = _lb(d, "h2")
        block(d, s, x, cy - h / 2, _head_w(d, kind, [it]) + 8, h,
              [dict(style="h2", text=str(it.get("label", "")), color="ground" if lift else color)], tag="row")


def _tgroup(d, it, w, style, lift=False):
    return [dict(style=style, text=E._balance(d, it.get("title", ""), w, style),
                 color="ground" if lift else "figure")]


def _bgroup(d, it, style, leading=None, lift=False, w=None):
    """설명 — 폭(w)을 주면 한 어절만 넘어가는 줄을 막아 미리 접는다."""
    if not it.get("body"):
        return []
    body = it["body"]
    if w:
        body = [ln for p_ in (body if isinstance(body, list) else [body])
                for ln in _wrap_no_orphan(d, str(p_), w, style, sep=True)]
    return [dict(style=style, text=body, color="ground" if lift else "ink2", leading=leading)]


def _list(d, s, sl, items, x0, x1, y, *, stack, split=None, gap, kind, draw=True, tb_gap=8):
    """판 없는 줄 목록을 (x0~x1, y 부터) 쌓는다. 줄 높이 목록을 돌려준다.
    stack=True 면 제목 아래 설명, 아니면 제목(왼쪽) · 설명(split 열부터)."""
    tstyle = sl.get("title_style", "h2")
    n = len(items)
    bstyle = sl.get("body_style", "lead" if n <= 3 else "body")
    lead_ = sl.get("body_leading")
    hw = _head_w(d, kind, items)
    tx = x0 + hw
    if stack:
        tw = x1 - tx
        bx, bw = tx, tw
    else:
        tw = split - 24 - tx
        bx, bw = split, x1 - split
    # 이름표(label)도 줄 높이는 번호 줄과 같게 — 같은 덱의 판들이 같은 줄 간격으로 보이게(사용자 판 26 · 32쪽, 10/25)
    head_h = {"num": _lb(d, "h1"), "badge": d.spec["shapes"]["badge_d"], "label": _lb(d, "h1")}.get(kind, 0)
    hs = []
    for i, it in enumerate(items):
        T = _tgroup(d, it, tw, tstyle)
        Bd = _bgroup(d, it, bstyle, lead_, w=bw)
        th = block_h(d, T, tw)
        bh = block_h(d, Bd, bw) if Bd else 0.0
        if stack:
            # 머리는 제목 첫 줄 가운데에 맞춘다 — 머리가 더 크면 위아래로 넘친 만큼 줄을 키운다
            over = max(0.0, (head_h - _lb(d, tstyle)) / 2)
            h = over + th + ((tb_gap + bh) if Bd else 0)
            h = max(h, over * 2 + _lb(d, tstyle))
        else:
            h = max(th, bh, head_h)
        hs.append(h)
        if draw:
            if stack:
                ty = y + over
                if kind:
                    _head(d, s, kind, it, i, sl, x0, ty + _lb(d, tstyle) / 2)
                block(d, s, tx, ty, tw, th, T, anchor="top", tag="row")
                if Bd:
                    block(d, s, bx, ty + th + tb_gap, bw, bh, Bd, anchor="top", tag="row")
            else:
                if kind:
                    _head(d, s, kind, it, i, sl, x0, y + h / 2)
                block(d, s, tx, y, tw, h, T, tag="row")
                if Bd:
                    block(d, s, bx, y, bw, h, Bd, tag="row")
        y += h + gap
    return hs


def plist(d, s, m, sl):
    """판 없는 목록 — 번호는 초록 큰 글자. 줄 사이는 여백으로만 가른다(선을 긋지 않는다 — 원칙 11)."""
    g = B.L(d)
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"plist 는 2~5줄 (받은 값 {n})")
    W = span_w(12)
    x0, x1 = col_x(1), content_r()
    kind = _head_kind(sl, items)
    stack = bool(sl.get("stack"))
    split = col_x(int(sl.get("body_col", 5)))
    gap = sl.get("gap", {2: 48, 3: 40, 4: 32, 5: 26}[n])
    hs = _list(d, s, sl, items, x0, x1, 0, stack=stack, split=split, gap=gap, kind=kind, draw=False)
    top_h, bot_h, foot_h, vgap = E._extras_h(d, sl, W)
    total = top_h + sum(hs) + gap * (n - 1) + bot_h + foot_h
    y = _mid_y(d, _top(d, g, sl), total)
    y = E._draw_extras_top(d, s, sl, y, W)
    _list(d, s, sl, items, x0, x1, y, stack=stack, split=split, gap=gap, kind=kind)
    E._draw_extras_bottom(d, s, sl, y + sum(hs) + gap * (n - 1), W)


def ppanel(d, s, m, sl):
    """큰 판 하나에 여러 줄 — 줄마다 판을 따로 깔지 않는다."""
    g = B.L(d)
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 5:
        raise SystemExit(f"ppanel 은 2~5줄 (받은 값 {n})")
    W = span_w(12)
    padx = sl.get("pad_x", 44)
    pady = sl.get("pad_y", {2: 40, 3: 34, 4: 32, 5: 30}[n])     # 3줄 34 · 28 은 사용자이 고친 판(10/25)
    x0, x1 = col_x(1) + padx, content_r() - padx
    kind = _head_kind(sl, items)
    stack = bool(sl.get("stack"))
    split = col_x(int(sl.get("body_col", 5)))
    gap = sl.get("gap", {2: 36, 3: 28, 4: 24, 5: 20}[n])
    hs = _list(d, s, sl, items, x0, x1, 0, stack=stack, split=split, gap=gap, kind=kind, draw=False)
    inner = sum(hs) + gap * (n - 1)
    ph = inner + 2 * pady
    top_h, bot_h, foot_h, vgap = E._extras_h(d, sl, W)
    # 결론 줄(h2, [ ] 안은 강조색) + 한 줄 설명(lead) — 판 위에. 결론을 먼저 읽고 판에서 단계를 본다
    C = []
    if sl.get("conclusion"):
        C.append(dict(style="h2", text=str(sl["conclusion"]), color="figure"))
        if sl.get("conclusion_sub"):
            C.append(dict(style="lead", text=str(sl["conclusion_sub"]), color="ink2", before=8))
    ch_ = (block_h(d, C, W) + 22) if C else 0.0
    y = _mid_y(d, _top(d, g, sl), top_h + ch_ + ph + bot_h + foot_h)
    y = E._draw_extras_top(d, s, sl, y, W)
    if C:
        block(d, s, col_x(1), y, W, ch_ - 22, C, anchor="top", tag="row")
        y += ch_
    d.panel(s, col_x(1), y, W, ph, color="accent_tint")
    _list(d, s, sl, items, x0, x1, y + pady, stack=stack, split=split, gap=gap, kind=kind)
    E._draw_extras_bottom(d, s, sl, y + ph, W)


def pgrid(d, s, m, sl):
    """2 × 2 판(2 · 4 · 6칸). 같은 높이, 판 안 글 덩어리는 세로 가운데."""
    g = B.L(d)
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if n not in (2, 4, 6):
        raise SystemExit(f"pgrid 는 2 · 4 · 6칸 (받은 값 {n})")
    W = span_w(12)
    gx = gy = sl.get("gap", 20)
    cw = (W - gx) / 2
    padx = sl.get("pad_x", 30)
    pady = sl.get("pad_y", 28)
    kind = _head_kind(sl, items)
    hw = _head_w(d, kind, items)
    tw = cw - 2 * padx - hw
    tstyle = sl.get("title_style", "h2")
    bstyle = sl.get("body_style", "lead")
    tb_gap = 10

    def ch_(it):
        h = block_h(d, _tgroup(d, it, tw, tstyle), tw)
        if it.get("body"):
            h += tb_gap + block_h(d, _bgroup(d, it, bstyle, sl.get("body_leading"), w=tw), tw)
        return h

    lb_t = _lb(d, tstyle)
    if kind == "badge":
        top_ink = min(INK_EM * _st(d, tstyle)["size"], lb_t / 2 - d.spec["shapes"]["badge_d"] / 2)
    elif kind == "num":
        top_ink = min(INK_EM * _st(d, tstyle)["size"], lb_t / 2 - _lb(d, "h1") / 2 + DIG_EM * _st(d, "h1")["size"])
    else:
        top_ink = INK_EM * _st(d, tstyle)["size"]

    def ink_h(it):
        last = bstyle if it.get("body") else tstyle
        return ch_(it) - INK_EM * _st(d, last)["size"] - top_ink

    ch = max(ink_h(it) for it in items) + 2 * pady
    rows_n = (n + 1) // 2
    top_h, bot_h, foot_h, vgap = E._extras_h(d, sl, W)
    total = top_h + rows_n * ch + (rows_n - 1) * gy + bot_h + foot_h
    y0 = _mid_y(d, _top(d, g, sl), total)
    y0 = E._draw_extras_top(d, s, sl, y0, W)
    for i, it in enumerate(items):
        x = col_x(1) + (i % 2) * (cw + gx)
        y = y0 + (i // 2) * (ch + gy)
        lift = sl.get("emphasis") is not None and i == sl.get("emphasis")   # 대비 장의 「대신 할 것」 한 칸만 진초록(00 26쪽과 같은 색)
        d.panel(s, x, y, cw, ch, color="accent" if lift else "accent_tint")
        row = items[(i // 2) * 2:(i // 2) * 2 + 2]
        cy = y + (ch - max(ink_h(r) for r in row)) / 2 - top_ink
        T = _tgroup(d, it, tw, tstyle, lift)
        th = block_h(d, T, tw)
        tx = x + padx + hw
        if kind:
            _head(d, s, kind, it, i, sl, x + padx, cy + _lb(d, tstyle) / 2, lift=lift)
        block(d, s, tx, cy, tw, th, T, anchor="top", tag="card")
        if it.get("body"):
            Bd = _bgroup(d, it, bstyle, sl.get("body_leading"), lift, w=tw)
            full = abs(ink_h(it) - max(ink_h(r) for r in row)) < 1.0
            block(d, s, tx, cy + th + tb_gap, tw, block_h(d, Bd, tw), Bd, anchor="bottom" if full else "top", tag="card")
    E._draw_extras_bottom(d, s, sl, y0 + rows_n * ch + (rows_n - 1) * gy, W)


def pcol(d, s, m, sl):
    """판 없는 단 2~4개 — 머리(초록 번호 · 아이콘) → 제목 → 설명. 단의 위를 맞추고 덩어리는 세로 가운데."""
    g = B.L(d)
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 4:
        raise SystemExit(f"pcol 은 2~4단 (받은 값 {n})")
    W = span_w(12)
    per = {2: 6, 3: 4, 4: 3}[n]
    cw = span_w(per) - {2: 28, 3: 24, 4: 16}[n]      # 단 사이 = 홈 20 + 이 몫
    kind = _head_kind(sl, items)
    tstyle = sl.get("title_style", "h2")
    bstyle = sl.get("body_style", "lead" if n <= 3 else "body")
    head_h = _lb(d, "h1") if kind == "num" else (d.spec["shapes"]["badge_d"] if kind == "badge" else 0)
    hgap = 16 if kind else 0
    tb_gap = 12

    def col_h(it):
        h = head_h + hgap + block_h(d, _tgroup(d, it, cw, tstyle), cw)
        if it.get("body"):
            h += tb_gap + block_h(d, _bgroup(d, it, bstyle, sl.get("body_leading"), w=cw), cw)
        return h

    hmax = max(col_h(it) for it in items)
    top_h, bot_h, foot_h, vgap = E._extras_h(d, sl, W)
    y0 = _mid_y(d, _top(d, g, sl), top_h + hmax + bot_h + foot_h)
    y0 = E._draw_extras_top(d, s, sl, y0, W)
    for i, it in enumerate(items):
        x = col_x(1 + i * per)
        y = y0
        if kind == "num":
            block(d, s, x, y, cw, head_h, [dict(style="h1", text=_num_label(it, i), color="accent")],
                  anchor="top", tag="col")
        elif kind == "badge":
            _head(d, s, kind, it, i, sl, x, y + head_h / 2)
        y += head_h + hgap
        T = _tgroup(d, it, cw, tstyle)
        th = block_h(d, T, cw)
        block(d, s, x, y, cw, th, T, anchor="top", tag="col")
        if it.get("body"):
            Bd = _bgroup(d, it, bstyle, sl.get("body_leading"), w=cw)
            block(d, s, x, y + th + tb_gap, cw, block_h(d, Bd, cw), Bd, anchor="top", tag="col")
    E._draw_extras_bottom(d, s, sl, y0 + hmax, W)


def pside(d, s, m, sl):
    """왼쪽 진초록 판 — 핵심 한 문장(side) · 오른쪽 — 판 없는 목록(제목 아래 설명). 판 높이 = 목록 높이.
    ★ 2026-10-25 사용자: 「왼쪽 상자에만 눈길이 가고 강조가 싸 보인다, 색감도 안 맞는다 — 이 레이아웃은 쓰지 말 것」.
      새 덱에서는 ppanel 의 conclusion(결론 줄, [ ] 강조) + conclusion_sub 로 바꾼다. 이미 배포한 00 · 01 을 다시 짤 때도 바꾼다."""
    E.B.warn(f"slide {d._slide_i}: pside 는 쓰지 않는다(10/25) — ppanel 의 conclusion 으로")
    g = B.L(d)
    g.head(s, sl, span=12)
    items = sl.get("items") or []
    n = len(items)
    if not 2 <= n <= 4:
        raise SystemExit(f"pside 는 2~4줄 (받은 값 {n})")
    W = span_w(12)
    lw = span_w(int(sl.get("side_span", 6)))
    padx, pady = 36, 36
    sstyle = sl.get("side_style", "h1")
    stext = _wrap_no_orphan(d, sl["side"], lw - 2 * padx, sstyle)
    S_ = [dict(style=sstyle, text=stext, color="ground")]
    if sl.get("side_label"):
        S_ = [dict(style="small", text=sl["side_label"], color="accent_tint")] + [dict(S_[0], before=14)]
    sh = block_h(d, S_, lw - 2 * padx)
    x1 = col_x(1) + lw + 48
    kind = _head_kind(sl, items)
    gap = sl.get("gap", {2: 36, 3: 30, 4: 24}[n])
    hs = _list(d, s, sl, items, x1, content_r(), 0, stack=True, gap=gap, kind=kind, draw=False)
    rh = sum(hs) + gap * (n - 1)
    ph = max(rh + 2 * 20, sh + 2 * pady)
    top_h, bot_h, foot_h, vgap = E._extras_h(d, sl, W)
    y = _mid_y(d, _top(d, g, sl), top_h + ph + bot_h + foot_h)
    y = E._draw_extras_top(d, s, sl, y, W)
    d.panel(s, col_x(1), y, lw, ph, color="accent")
    block(d, s, col_x(1) + padx, y, lw - 2 * padx, ph, S_, tag="card")
    _list(d, s, sl, items, x1, content_r(), y + (ph - rh) / 2, stack=True, gap=gap, kind=kind)
    E._draw_extras_bottom(d, s, sl, y + ph, W)


# ================================================================ orig — 원본 문항지 조각(모의 면접 02~)
def orig(d, s, m, sl):
    """원본 문항지 조각 — 눈썹(문항 · 조각 · 출처)과 그림만. 한 문항의 조각은 같은 배율(scale, pt/px)로 넣는다
    (9/27 「문제와도 동일」 — 조각마다 칸을 꽉 채우면 글씨 크기가 들쭉날쭉하다). 배율은 아웃라인에서 정한다.
    그림은 본문 폭 가운데, 눈썹 아래 ~ 본문 하한 사이 세로 가운데."""
    from PIL import Image
    g = B.L(d)
    g.head(s, sl, span=12)
    path = sl["image"] if os.path.isabs(sl["image"]) else os.path.join(os.getcwd(), sl["image"])
    iw, ih = Image.open(path).size
    sc = float(sl["scale"])
    w, h = iw * sc, ih * sc
    top, bot = 80.0, _bottom(d)
    if w > span_w(12) + 0.5 or h > bot - top + 0.5:
        raise SystemExit(f"orig: {sl['image']} 가 영역을 넘는다 ({w:.0f}×{h:.0f})")
    d.picture(s, path, col_x(1) + (span_w(12) - w) / 2, top + (bot - top - h) / 2, w, h, fit="contain", tag="orig")


# ================================================================ pfig — 풀이 그림 + 글(모의 면접 04~)
def _conclusion(sl):
    """결론 줄(h2, [ ] 안은 강조색) + 한 줄 설명(lead) — ppanel 과 같은 모양."""
    C = []
    if sl.get("conclusion"):
        C.append(dict(style="h2", text=str(sl["conclusion"]), color="figure"))
        if sl.get("conclusion_sub"):
            C.append(dict(style="lead", text=str(sl["conclusion_sub"]), color="ink2", before=8))
    return C


def pfig(d, s, m, sl):
    """풀이 그림 한 장 — 판 없이 바탕에 바로 둔다(그림이 주인공, 판은 글에만 쓰던 것).
    fig_side: left · right — 그림 칸(fig_span 칸, 기본 6) 옆에 결론 줄 + 판 없는 목록(제목 아래 설명). 둘은 서로 세로 가운데.
              below — 결론 줄을 위에, 그림을 아래 본문 폭 가운데에(가로로 긴 그림). 그림은 칸 안에 비율을 지켜 넣는다."""
    from PIL import Image
    g = B.L(d)
    g.head(s, sl, span=12)
    path = sl["image"] if os.path.isabs(sl["image"]) else os.path.join(os.getcwd(), sl["image"])
    iw, ih = Image.open(path).size
    W = span_w(12)
    top, bot = _top(d, g, sl), _bottom(d)
    items = sl.get("items") or []
    side = sl.get("fig_side", "left")
    if side == "below":
        C = _conclusion(sl)
        ch_ = (block_h(d, C, W) + 30) if C else 0.0
        fw_max = span_w(int(sl.get("fig_span", 12)))
        sc = min(fw_max / iw, float(sl.get("fig_h") or (bot - top - ch_)) / ih)
        fw, fh = iw * sc, ih * sc
        y = _mid_y(d, top, ch_ + fh)
        if C:
            block(d, s, col_x(1), y, W, ch_ - 30, C, anchor="top", tag="row")
        d.picture(s, path, col_x(1) + (W - fw) / 2, y + ch_, fw, fh, fit="contain", tag="fig")
        return
    fs = int(sl.get("fig_span", 6))
    fw_max = span_w(fs)
    if side == "left":
        fx, tx0, tx1 = col_x(1), col_x(fs + 1), content_r()
    else:
        fx, tx0, tx1 = content_r() - fw_max, col_x(1), col_x(1) + span_w(12 - fs)
    tw = tx1 - tx0
    sc = min(fw_max / iw, float(sl.get("fig_h") or (bot - top)) / ih)
    fw, fh = iw * sc, ih * sc
    n = len(items)
    kind = _head_kind(sl, items)
    gap = sl.get("gap", {1: 0, 2: 30, 3: 24, 4: 20, 5: 16}.get(n, 16))
    hs = _list(d, s, sl, items, tx0, tx1, 0, stack=True, gap=gap, kind=kind, draw=False) if items else []
    lh = sum(hs) + gap * (n - 1) if items else 0.0
    C = _conclusion(sl)
    cb = block_h(d, C, tw) if C else 0.0
    ch_ = cb + (28 if items else 0) if C else 0.0
    th = ch_ + lh
    H = max(fh, th)
    y = _mid_y(d, top, H)
    d.picture(s, path, fx + (fw_max - fw) / 2, y + (H - fh) / 2, fw, fh, fit="contain", tag="fig")
    ty = y + (H - th) / 2
    if C:
        block(d, s, tx0, ty, tw, cb, C, anchor="top", tag="row")
        ty += ch_
    if items:
        _list(d, s, sl, items, tx0, tx1, ty, stack=True, gap=gap, kind=kind)


def install():
    E.install()
    for k, fn in (("pcards", pcards), ("prows", prows), ("ask", ask), ("ctimeline", ctimeline),
                  ("plist", plist), ("ppanel", ppanel), ("pgrid", pgrid), ("pcol", pcol), ("pside", pside),
                  ("orig", orig), ("pfig", pfig)):
        B.LAYOUTS[k] = E.wrap(k, fn)


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
