#!/usr/bin/env python3
"""배포본 대조 — 덮어쓰기 전에 사용자가 직접 고친 것이 있는지 확인한다.

원칙: 자료를 고치기 **전에 반드시** 확인한다. 세션 시작 때 한 번이 아니라
**쓰기 직전마다**. md5 는 빌드마다 달라지므로(타임스탬프) 판단 기준이 아니다 —
도형 단위 대조가 기준이다.

  reconcile.py dump    deck.pptx
  reconcile.py diff    배포본.pptx 내빌드.pptx
  reconcile.py deploy  내빌드.pptx 배포본.pptx [--first]

deploy 는 배포본이 마지막 배포 이후 바뀌었으면 **거부하고 차이를 보여 준다.**
그 차이가 사용자의 수정이다. 되돌리지 말고 생성 스크립트에 반영한 뒤 다시 배포한다.
"""
from __future__ import annotations
import argparse, hashlib, json, os, re, shutil, sys, zipfile

from pptx import Presentation
from pptx.util import Emu

HERE = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.join(HERE, "..", ".baselines")
P = lambda v: round(Emu(v).pt, 1)


def _key(path):
    return hashlib.sha1(os.path.abspath(path).encode()).hexdigest()[:16]


def dump(path):
    """도형 단위 지문. (쪽, x, y, w, h, 크기, 굵기, 글) — 위치·서식·문구를 모두 본다."""
    prs = Presentation(path)
    out = []
    for i, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            geo = (i, P(sh.left), P(sh.top), P(sh.width), P(sh.height))
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                prst = re.search(r'prstGeom prst="([^"]+)"', sh._element.xml)
                out.append(geo + (None, None, "<shape:%s>" % (prst.group(1) if prst else "?")))
                continue
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    out.append(geo + (r.font.size.pt if r.font.size else None,
                                      bool(r.font.bold), r.text))
    return out


def _edit_signals(path):
    """PowerPoint 가 남기는 편집 흔적. 결정적 증거는 아니지만 참고가 된다."""
    try:
        x = zipfile.ZipFile(path).read("docProps/app.xml").decode()
    except Exception:
        return {}
    g = lambda t: (re.search(rf"<{t}>(.*?)</{t}>", x) or [None, ""])[1]
    return {"Application": g("Application"), "TotalTime": g("TotalTime")}


def diff(a, b, label_a="배포본", label_b="내 빌드"):
    da, db = set(map(tuple, dump(a))), set(map(tuple, dump(b)))
    only_a, only_b = sorted(da - db), sorted(db - da)
    if not only_a and not only_b:
        print("  ✓ 도형 단위로 완전히 같다")
        return 0
    print(f"  {label_a}에만 있는 것 {len(only_a)}개 · {label_b}에만 있는 것 {len(only_b)}개\n")
    for tag, rows in ((f"{label_a} (사용자 수정일 가능성)", only_a), (f"{label_b}", only_b)):
        if not rows:
            continue
        print(f"  --- {tag}")
        for r in rows[:40]:
            i, x, y, w, h, sz, bold, t = r
            print(f"    s{i:02d} ({x:6},{y:6}) {w:6}x{h:5} "
                  f"{('%.1fpt' % sz) if sz else '   -  '} {'B' if bold else ' '} {t[:44]!r}")
        if len(rows) > 40:
            print(f"    … 외 {len(rows) - 40}개")
        print()
    return 1


def deploy(src, dest, first=False):
    os.makedirs(os.path.abspath(BASE_DIR), exist_ok=True)
    bpath = os.path.join(BASE_DIR, _key(dest) + ".json")
    if os.path.exists(dest):
        if not os.path.exists(bpath):
            if not first:
                raise SystemExit(
                    f"✗ 기준 덤프가 없다. 지금 덮어쓰면 사용자의 수정을 확인할 방법이 없다.\n"
                    f"  배포본을 먼저 검토하고, 내 빌드와 같다고 판단되면 --first 로 기준을 세워라.")
        else:
            base = [tuple(r) for r in json.load(open(bpath, encoding="utf-8"))]
            cur = [tuple(r) for r in dump(dest)]
            if set(base) != set(cur):
                print(f"✗ 배포본이 마지막 배포 이후 바뀌었다 — 덮어쓰지 않았다.\n")
                only_cur = sorted(set(cur) - set(base))
                for r in only_cur[:40]:
                    i, x, y, w, h, sz, bold, t = r
                    print(f"    s{i:02d} ({x:6},{y:6}) {w:6}x{h:5} "
                          f"{('%.1fpt' % sz) if sz else '   -  '} {'B' if bold else ' '} {t[:44]!r}")
                if len(only_cur) > 40:
                    print(f"    … 외 {len(only_cur) - 40}개")
                sig = _edit_signals(dest)
                if sig:
                    print(f"\n  편집 흔적: {sig}")
                print("\n  이 차이가 사용자의 수정이다. 되돌리지 말고 아웃라인·빌더에 반영한 뒤")
                print("  다시 빌드해서 배포하라. 의도가 불분명하면 물어봐라.")
                return 1
    shutil.copyfile(src, dest)
    json.dump(dump(dest), open(bpath, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"✓ 배포 완료 -> {dest}")
    print(f"  기준 덤프 갱신: {os.path.normpath(bpath)}")
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("dump"); d.add_argument("pptx")
    f = sub.add_parser("diff"); f.add_argument("a"); f.add_argument("b")
    p_ = sub.add_parser("deploy"); p_.add_argument("src"); p_.add_argument("dest")
    p_.add_argument("--first", action="store_true")
    n = ap.parse_args()
    if n.cmd == "dump":
        for r in dump(n.pptx):
            print(r)
        sys.exit(0)
    sys.exit(diff(n.a, n.b) if n.cmd == "diff" else deploy(n.src, n.dest, n.first))
