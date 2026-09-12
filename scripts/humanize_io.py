#!/usr/bin/env python3
"""아웃라인 ↔ 윤문 텍스트 왕복.

humanize-korean 은 CLI 가 아니라 모델이 도는 스킬이므로, 이 스크립트는 기계적인 부분만
맡는다 — 뽑아내기(extract)와 되돌려 넣기(apply). 윤문 자체는 humanize-korean 이 한다.

  1) extract  outline.yaml -o text.txt   (+ text.txt.map.json)
  2) humanize-korean 을 standard 경로로 돌린다. 아래 구조 계약을 반드시 프롬프트에 넣는다.
  3) apply    outline.yaml text.txt      (원본 YAML 을 제자리에서 고친다)

뽑지 않는 것 — code 블록, 숫자만 있는 값, 이미지 경로, index/number 같은 식별자.
슬라이드 안 줄바꿈은 " / " 로 표시한다(구조 계약과 같은 규약).
"""
from __future__ import annotations
import argparse, json, os, re, sys

import yaml

CONTRACT = """구조 계약(반드시 지킬 것):
- [n쪽] 표시와 줄 수를 그대로 유지한다. 줄을 합치거나 쪼개지 않는다. 한 줄 입력 → 한 줄 출력.
- " / "는 슬라이드 안의 줄바꿈이다. 그대로 둔다.
- 코드·함수명·수치·슬라이스 표기·URL·고유명사·출처 표기는 건드리지 않는다.
- 경어체 유지.
- 제목 줄은 짧은 명사구 쪽이 좋다. 단 서술어가 내용을 담은 제목은 그대로 둔다.
- 설명을 위한 구체적 서술은 손대지 않는다.
산문 규칙(문단 구성, 접속 보강, 장문 삽입)은 적용 금지."""

# 뽑을 필드 — 식별자(index/number)·경로·수치는 제외한다
# 눈썹·러너·시기는 구조 라벨이다. 윤문 대상이 아니고, " / " 가 들어 있어
# 줄바꿈 표시와 충돌한다.
FIELDS = ("title", "subtitle", "lead", "text", "body", "note", "notes", "source",
          "caption", "footnote", "subhead", "subbody", "panel", "conclusion")
SKIP_LAYOUT_FIELDS = {"code": {"code", "caption"}}
NUMERIC = re.compile(r"^[\d\s.,:%+\-/()가-힣]{0,4}$")


def _walk(sl, idx):
    """(경로, 문자열) 쌍을 낸다. 경로는 되돌려 넣을 때 쓰는 식별자다."""
    skip = SKIP_LAYOUT_FIELDS.get(sl.get("layout"), set())
    for k in FIELDS:
        if k in skip or not sl.get(k):
            continue
        v = sl[k]
        if isinstance(v, list):
            for i, x in enumerate(v):
                if isinstance(x, str) and x.strip():
                    yield f"{idx}.{k}[{i}]", x
        elif isinstance(v, str) and v.strip():
            yield f"{idx}.{k}", v
    for grp in ("items", "steps", "nodes", "rings", "pair", "ledger"):
        for i, it in enumerate(sl.get(grp) or []):
            if isinstance(it, str):
                if it.strip():
                    yield f"{idx}.{grp}[{i}]", it
                continue
            for k in ("title", "body", "text", "label", "caption", "name", "subs"):
                v = it.get(k)
                if isinstance(v, list):
                    for j, x in enumerate(v):
                        if isinstance(x, str) and x.strip():
                            yield f"{idx}.{grp}[{i}].{k}[{j}]", x
                elif isinstance(v, str) and v.strip():
                    if k == "value" and NUMERIC.match(v):
                        continue          # 수치는 손대지 않는다
                    yield f"{idx}.{grp}[{i}].{k}", v


def extract(outline, out):
    o = yaml.safe_load(open(outline, encoding="utf-8"))
    lines, mapping = [], []
    for idx, sl in enumerate(o.get("slides") or [], 1):
        got = [(p, t) for p, t in _walk(sl, idx)]
        if not got:
            continue
        lines.append(f"[{idx}쪽]")
        for path, t in got:
            if " / " in t and "\n" not in t:
                print(f"  ! {path}: 본문에 ' / ' 가 있어 줄바꿈 표시와 겹친다 — "
                      f"윤문 후 확인하라", file=sys.stderr)
            lines.append(t.replace("\n", " / "))
            mapping.append({"path": path, "orig": t})
    open(out, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    json.dump(mapping, open(out + ".map.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print(f"✓ {out}  ({len(mapping)}줄 / {sum(1 for l in lines if l.startswith('['))}장)")
    print(f"  대응표: {out}.map.json")
    print("\n--- humanize-korean 프롬프트에 이 계약을 그대로 넣어라 ---")
    print(CONTRACT)


def apply(outline, text, out=None):
    mapping = json.load(open(text + ".map.json", encoding="utf-8"))
    body = [l for l in open(text, encoding="utf-8").read().split("\n")
            if l.strip() and not l.strip().startswith("[")]
    if len(body) != len(mapping):
        raise SystemExit(f"줄 수가 다르다 — 원본 {len(mapping)} / 윤문 {len(body)}. "
                         f"구조 계약(줄 합치기·쪼개기 금지)이 깨졌다. 다시 윤문하라.")
    src = open(outline, encoding="utf-8").read()
    changed = 0
    for m, new in zip(mapping, body):
        old = m["orig"]
        new = new.replace(" / ", "\n")
        if new == old:
            continue
        # YAML 안에서 원문을 정확히 찾아 바꾼다 (주석·서식 보존)
        for q in ('"', "'"):
            a, b = q + old.replace("\n", "\\n") + q, q + new.replace("\n", "\\n") + q
            if src.count(a) == 1:
                src = src.replace(a, b); changed += 1; break
            a2, b2 = q + old + q, q + new + q
            if src.count(a2) == 1:
                src = src.replace(a2, b2); changed += 1; break
        else:
            print(f"  ! 찾지 못했거나 중복이다 ({m['path']}): {old[:34]!r}", file=sys.stderr)
    open(out or outline, "w", encoding="utf-8").write(src)
    print(f"✓ {changed}줄 반영 -> {out or outline}")
    print("  결과를 그대로 받지 말고 검토하라. 사용자가 지정한 문구가 있으면 그쪽이 우선이다.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("extract"); e.add_argument("outline"); e.add_argument("-o", required=True)
    a = sub.add_parser("apply"); a.add_argument("outline"); a.add_argument("text")
    a.add_argument("-o")
    n = ap.parse_args()
    if n.cmd == "extract":
        extract(n.outline, n.o)
    else:
        apply(n.outline, n.text, n.o)
