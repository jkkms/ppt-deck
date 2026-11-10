#!/usr/bin/env python3
"""발표 분량 예산 검사 — 청중에게 보여 주는 양은 발표 시간에 맞춘다(2026-11-10).

사용자 지시: 「내가 아는 건 지금처럼 많아야 하는데, 사람들에게는 지금의 반 정도만 발표」.
그래서 둘을 나눠 센다.
  청중 몫(예산 안)  본편 장 수 · 노트의 「말할 것」(맨 앞 ~ 「알아 둘 것」·「예상 질문」 앞)
  발표자 몫(제한 없음)  노트의 「알아 둘 것」·「예상 질문」 · 부록 장

예산(근거: GDGoC 11월 세미나 서빙 본편 = 15분 자리에 43장 · 말할 것 8,319자 → 「반으로」)
  본편 장 수   ≤ 발표 분 × 1.5        (15분 → 22장)
  말할 것      ≤ 발표 분 × 280자      (공백 뺀 글자 수, 15분 → 4,200자)

본편 범위: --pages 로 주지 않으면 2쪽부터, 「부록」·「감사합니다」만 있는 장 앞까지.
발표 분은 말하는 시간이다. 실습·활동·데모 재생 시간은 뺀다.

사용:  budget.py deck.pptx --minutes 15 [--pages 18-60] [--each]
종료코드: 예산을 넘으면 1
"""
from __future__ import annotations
import argparse, re, sys
from pptx import Presentation

SLIDES_PER_MIN = 1.5
CHARS_PER_MIN = 280
KNOW = re.compile(r"알아\s*둘\s*것|예상\s*질문")
END = {"부록", "감사합니다", "감사합니다.", "Q&A", "질문"}


def notes(s):
    return s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""


def screen_texts(s):
    out = []
    for sh in s.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t and not re.fullmatch(r"(서빙\s*)?\d+", t):   # 쪽 번호는 빼고
                out.append(t)
        if getattr(sh, "has_table", False) and sh.has_table:
            out += [c.text for r in sh.table.rows for c in r.cells if c.text.strip()]
    return out


def nchar(t):
    return len(re.sub(r"\s", "", t))


def body_range(slides, spec):
    if spec:
        a, b = (int(x) for x in spec.split("-"))
        return a, b
    for i, s in enumerate(slides, 1):
        if i > 1 and set(screen_texts(s)) and set(screen_texts(s)) <= END:
            return 2, i - 1
    return 2, len(slides)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--minutes", type=float, required=True, help="말하는 시간(분)")
    ap.add_argument("--pages", help="본편 쪽 범위, 예: 18-60")
    ap.add_argument("--each", action="store_true", help="장마다 글자 수")
    a = ap.parse_args()

    slides = list(Presentation(a.pptx).slides)
    lo, hi = body_range(slides, a.pages)
    say = know = scr = 0
    rows = []
    for i in range(lo, hi + 1):
        n = notes(slides[i - 1])
        m = KNOW.search(n)
        s_say, s_know = (n[:m.start()], n[m.start():]) if m else (n, "")
        c = (nchar(s_say), nchar(s_know), sum(nchar(t) for t in screen_texts(slides[i - 1])))
        say, know, scr = say + c[0], know + c[1], scr + c[2]
        rows.append((i, *c))

    k = hi - lo + 1
    k_max = int(a.minutes * SLIDES_PER_MIN)
    c_max = int(a.minutes * CHARS_PER_MIN)
    print(f"본편 {lo}~{hi}쪽 · 발표 {a.minutes:g}분")
    print(f"  장 수        {k:>6}  / 예산 {k_max:>5}  {'넘침' if k > k_max else '통과'}")
    print(f"  말할 것      {say:>6}자 / 예산 {c_max:>5}자 {'넘침' if say > c_max else '통과'}")
    print(f"  화면 글      {scr:>6}자 (참고)")
    print(f"  알아 둘 것   {know:>6}자 (발표자 몫, 제한 없음)")
    if a.each:
        print("\n  쪽  말할것  알아둘것  화면")
        for r in rows:
            print(f"  {r[0]:>3} {r[1]:>6} {r[2]:>8} {r[3]:>5}")
    over = k > k_max or say > c_max
    if over:
        print("\n넘친 만큼 청중 몫에서 덜어 「알아 둘 것」·부록으로 옮긴다. 장을 더 쪼개서 맞추지 않는다.")
    return 1 if over else 0


if __name__ == "__main__":
    sys.exit(main())
