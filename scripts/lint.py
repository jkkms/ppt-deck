#!/usr/bin/env python3
"""ppt-deck 콘텐츠 린터 — outline.yaml에 남은 'AI 티'를 잡는다.

기하(넘침·이탈)는 build.py가 자기검증한다. 여기서는 내용과 리듬만 본다.
사용:  lint.py outline.yaml [--spec other-spec.yaml]
종료코드: ERROR 있으면 1
"""
from __future__ import annotations
import argparse, os, re, sys
from collections import Counter

import yaml
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import load_spec, EMOJI

SENTENCE_END = re.compile(r"(습니다|합니다|됩니다|입니다|이다|한다|된다|요)[.!?]?$|[.!?]$")
CLICHE = ["살펴보겠습니다", "알아보겠습니다", "중요합니다", "핵심입니다", "혁신적",
          "효율적", "다양한", "다각도", "체계적으로", "극대화", "패러다임",
          "~을 통해 ~할 수 있습니다"]
ERRORS: list[str] = []
WARNS: list[str] = []


def err(i, code, msg): ERRORS.append(f"[{code}] slide {i}: {msg}")
def wrn(i, code, msg): WARNS.append(f"[{code}] slide {i}: {msg}")


def texts_of(sl) -> list[str]:
    out = []
    for k in ("title", "subtitle", "lead", "text", "note", "source", "body"):
        if sl.get(k): out.append(str(sl[k]))
    out += [str(b) for b in (sl.get("bullets") or [])]
    out += [str(x) for x in (sl.get("lines") or [])]
    for it in (sl.get("items") or []):
        out += [str(it.get(k, "")) for k in
                ("value", "unit", "caption", "number", "title", "body") if it.get(k)]
    for r in (sl.get("rows") or []):
        out += [str(c) for c in r]
    out += [str(h) for h in (sl.get("headers") or [])]
    return out


def glyph_gaps(spec, texts):
    """폰트에 없는 문자를 찾는다. PowerPoint는 이걸 두부(□)로 그린다.

    Pretendard에는 ⊃ ⊂ ✕ ∴ ∵ ⊙ 가 없다 — 수학·논리 기호를 쓸 때 반드시 걸린다.
    렌더해 보기 전에는 눈치채기 어렵고, 미리보기에서도 대체 글꼴로 그려져 속기 쉽다.
    """
    import metrics as _m
    tbls = _m.load() or {}
    if not tbls:
        return {}
    gaps = {}
    for t in texts:
        for ch in str(t):
            if ch.isspace() or "가" <= ch <= "힣" or ord(ch) < 0x80:
                continue
            missing = [n for n, tb in tbls.items() if str(ord(ch)) not in tb["adv"]]
            if len(missing) == len(tbls):
                gaps.setdefault(ch, 0)
                gaps[ch] += 1
    return gaps


def main(path, spec_path=None):
    spec = load_spec(spec_path)
    o = yaml.safe_load(open(path, encoding="utf-8"))
    base = os.path.dirname(os.path.abspath(path))
    m = o.get("meta") or {}
    slides = o.get("slides") or []
    dens = m.get("density") or spec["density"]
    lim = spec["density_limits"][dens]
    n = len(slides)

    # --- 슬라이드별 ---------------------------------------------------
    prev, run = None, 0
    for i, sl in enumerate(slides, 1):
        lay = sl.get("layout")
        run = run + 1 if lay == prev else 1
        prev = lay
        if run > spec["rhythm"]["max_same_layout_run"]:
            err(i, "RHYTHM", f"'{lay}' {run}연속. 모든 장이 같은 골격이면 그게 AI PPT다 — "
                             f"statement/data/quote를 끼워 호흡을 끊어라")

        for t in texts_of(sl):
            if EMOJI.search(t):
                err(i, "EMOJI", f"이모지 발견: {t[:30]!r}")
            for c in CLICHE:
                if c in t:
                    wrn(i, "CLICHE", f"상투어 '{c}' — 구체적 명사로 바꿔라: {t[:34]!r}")

        bl = [str(b) for b in (sl.get("bullets") or [])]
        if len(bl) > lim["bullets"]:
            err(i, "DENSITY", f"불릿 {len(bl)}개 > {dens} 한도 {lim['bullets']}개. "
                              f"글자를 줄이지 말고 슬라이드를 쪼개라")
        for b in bl:
            if SENTENCE_END.search(b.strip()):
                wrn(i, "SENTENCE", f"불릿이 완결 문장이다 — 명사구로 잘라라: {b[:34]!r}")
            if len(b) > lim["bullet_chars"]:
                wrn(i, "LONG", f"불릿 {len(b)}자 > {lim['bullet_chars']}자: {b[:34]!r}")
        if len(bl) >= 3:
            L = [len(b) for b in bl]
            if max(L) - min(L) <= 2:
                wrn(i, "UNIFORM", f"불릿 길이가 {L}로 균일하다. 사람은 이렇게 안 쓴다 — "
                                  f"하나는 짧게, 하나는 길게")
            heads = Counter(b[:2] for b in bl)
            if heads.most_common(1)[0][1] >= 3:
                wrn(i, "UNIFORM", "불릿이 전부 같은 두 글자로 시작한다 — 기계적 병렬")

        if lay in ("image_split", "image_full") and not sl.get("image"):
            err(i, "IMAGE", f"'{lay}' 인데 image 경로가 없다")
        for ip in ([sl.get("image")] + [it.get("image") for it in (sl.get("items") or [])]):
            if not ip:
                continue
            q = os.path.expanduser(str(ip))
            q = q if os.path.isabs(q) else os.path.join(base, q)   # 아웃라인 파일 기준
            if not os.path.exists(q):
                err(i, "IMAGE", f"이미지 파일이 없다: {ip}")
        if lay == "cards" and len(sl.get("items") or []) > 6:
            err(i, "DENSITY", f"카드 {len(sl['items'])}개 > 6개. 한 화면에 여섯 덩이 넘게 놓으면 아무도 안 읽는다")
        if lay == "two_col" and not sl.get("lead") and len(bl) <= 2:
            wrn(i, "BALANCE", "lead 없이 불릿 2개 이하 — 왼쪽 컬럼이 비어 화면이 한쪽으로 쏠린다. "
                              "lead를 넣거나 statement로 바꿔라")
        if lay == "table" and len(sl.get("rows") or []) > lim["table_rows"]:
            err(i, "DENSITY", f"표 {len(sl['rows'])}행 > 한도 {lim['table_rows']}행")

    # --- 덱 전체 ------------------------------------------------------
    lays = [s.get("layout") for s in slides]
    if n >= 6:
        if lays[0] != "cover":
            wrn(0, "STRUCTURE", "첫 장이 cover가 아니다")
        if lays[-1] != "closing":
            wrn(0, "STRUCTURE", "마지막 장이 closing이 아니다")
        if not ({"statement", "quote", "data"} & set(lays)):
            err(0, "RHYTHM", "statement/quote/data가 하나도 없다. 본문 레이아웃만 반복되는 덱은 "
                             "내용과 무관하게 AI가 찍어낸 티가 난다")
        has_img = any(s_.get("image") or any(it.get("image") for it in (s_.get("items") or []))
                      for s_ in slides)
        # 시각적 닻이 하나도 없을 때만 경고한다. 큰 숫자 한 장(data)도 닻으로 친다 —
        # 순수 타이포그래피 덱은 정당한 선택이고, 문제는 '불릿만 30장'이다.
        if not has_img and not ({"data", "image_split", "image_full"} & set(lays)) and n >= 8:
            wrn(0, "VISUAL", "이미지도 큰 숫자 슬라이드도 없다. 글자만 이어지면 기억에 남지 않는다 — "
                             "image_split / image_full / cards 썸네일 / data 를 섞어라")
        if len(set(lays)) < 4:
            wrn(0, "RHYTHM", f"레이아웃이 {len(set(lays))}종뿐 — 최소 4종 섞어라")

    titles = [str(s.get("title", "")) for s in slides if s.get("title")]
    if len(titles) >= 3:
        tail = Counter(t[-2:] for t in titles if len(t) >= 2)
        w, c = tail.most_common(1)[0]
        if c >= 3:
            wrn(0, "PARALLEL", f"제목 {c}개가 '{w}'로 끝난다 — '~의 이해/활용/전망'식 병렬은 "
                               f"AI 목차의 지문이다. 최소 하나는 깨라")
        head = Counter(t[:2] for t in titles if len(t) >= 2)
        hw, hc = head.most_common(1)[0]
        if hc >= 3:
            wrn(0, "PARALLEL", f"제목 {hc}개가 '{hw}'로 시작한다 — 'X의 이해/활용/전망'식 목차")
        Lt = [len(t) for t in titles]
        if len(titles) >= 4 and max(Lt) - min(Lt) <= 1:
            wrn(0, "PARALLEL", f"제목 길이가 {Lt}로 전부 같다")

    gaps = glyph_gaps(spec, [t for sl in slides for t in texts_of(sl)]
                      + [str(v) for v in m.values()])
    for ch, cnt in sorted(gaps.items(), key=lambda kv: -kv[1]):
        err(0, "GLYPH", f"'{ch}' (U+{ord(ch):04X}) 가 폰트에 없다 — PowerPoint가 두부(□)로 그린다. "
                        f"{cnt}곳에서 쓰였다")

    # --- 출력 ---------------------------------------------------------
    print(f"ppt-deck lint — {path}  ({n} slides, density={dens})")
    for e in ERRORS: print("  ✗ " + e)
    for w in WARNS:  print("  · " + w)
    if not ERRORS and not WARNS:
        print("  ✓ 걸린 것 없음")
    print(f"\n  ERROR {len(ERRORS)} / WARN {len(WARNS)}")
    return 1 if ERRORS else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("outline"); ap.add_argument("--spec")
    a = ap.parse_args()
    sys.exit(main(a.outline, a.spec))
