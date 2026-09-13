#!/usr/bin/env python3
"""사용자의 「PPT 제작 원칙」 문서 추적.

사용자는 실제 자료를 고치면서 이 문서를 계속 고친다. 스킬은 그 문서를 따라가야 한다.
마지막으로 반영한 판을 스냅숏으로 두고, 바뀐 절을 찾아 보여 준다.

  principles.py check     바뀐 절이 있으면 보여 주고 종료코드 1
  principles.py ack       현재 판을 '반영 완료'로 기록 (빌더·스펙에 반영한 뒤에만)
  principles.py log       반영 이력

문서 경로와 스냅숏은 저장소 밖(.principles/, gitignore)에 둔다 —
사용자 개인 원칙 문서라 공개 저장소에 올리지 않는다.
"""
from __future__ import annotations
import argparse, datetime, difflib, hashlib, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(HERE, "..", ".principles")
CONF = os.path.join(STATE, "config.json")
SNAP = os.path.join(STATE, "snapshot.md")
LOG = os.path.join(STATE, "log.jsonl")
SEEN = os.path.join(STATE, "seen.sha")   # check 가 마지막으로 보여 준 판
DEFAULT = os.path.expanduser(
    "~/Library/Mobile Documents/com~apple~CloudDocs/2026/주니어카이스트/PPT 제작 원칙.md")


def doc_path():
    if os.path.exists(CONF):
        return json.load(open(CONF, encoding="utf-8")).get("path", DEFAULT)
    return DEFAULT


def sections(text):
    """'## ' · '### ' 제목 단위로 자른다. 번호가 밀려도(4-2 -> 4-3) 제목 글로 짝짓는다."""
    out, cur, buf = {}, "(머리말)", []
    for line in text.split("\n"):
        m = re.match(r"^(#{2,3})\s+(.*)$", line)
        if m:
            out[cur] = "\n".join(buf).strip()
            cur, buf = line.strip(), []
        else:
            buf.append(line)
    out[cur] = "\n".join(buf).strip()
    return out


def _key(h):
    """번호를 떼고 제목 글만 남긴다. '### 4-2. 코드 블록' -> '코드 블록'"""
    return re.sub(r"^#{2,3}\s*[\d\-.]+\s*", "", h).strip()


def _sha(text):
    return hashlib.sha1(text.encode()).hexdigest()[:8]


def check():
    p = doc_path()
    if not os.path.exists(p):
        print(f"원칙 문서가 없다: {p}")
        return 2
    cur = open(p, encoding="utf-8").read()
    if not os.path.exists(SNAP):
        os.makedirs(STATE, exist_ok=True)
        open(SEEN, "w").write(_sha(cur))
        print("스냅숏이 없다 — 처음 추적을 시작한다. 문서를 읽고 반영한 뒤 ack 하라.")
        return 1
    old = open(SNAP, encoding="utf-8").read()
    os.makedirs(STATE, exist_ok=True)
    open(SEEN, "w").write(_sha(cur))
    if old == cur:
        print(f"✓ 원칙 문서 변경 없음 ({_sha(cur)})")
        return 0

    so, sn = sections(old), sections(cur)
    ko = {_key(h): (h, b) for h, b in so.items()}
    kn = {_key(h): (h, b) for h, b in sn.items()}
    added = [k for k in kn if k not in ko]
    removed = [k for k in ko if k not in kn]
    renum = [k for k in kn if k in ko and kn[k][0] != ko[k][0]]
    changed = [k for k in kn if k in ko and kn[k][1] != ko[k][1]]

    print(f"✗ 원칙 문서가 바뀌었다 — {p}\n")
    for k in added:
        print(f"  [새 절]   {kn[k][0]}")
    for k in removed:
        print(f"  [삭제]    {ko[k][0]}")
    for k in renum:
        print(f"  [번호]    {ko[k][0]}  ->  {kn[k][0]}")
    for k in changed:
        print(f"  [수정]    {kn[k][0]}")
    print()
    for k in added + changed:
        print(f"───── {kn[k][0]}")
        a = ko[k][1].split("\n") if k in ko else []
        b = kn[k][1].split("\n")
        for line in difflib.unified_diff(a, b, lineterm="", n=1):
            if line.startswith(("---", "+++")):
                continue
            print("  " + line)
        print()
    print("  반영 절차: 이 차이를 스펙·빌더·SKILL.md 에 반영 -> 빌드·검사 통과 -> principles.py ack")
    return 1


def ack(note=""):
    p = doc_path()
    cur = open(p, encoding="utf-8").read()
    # 사용자는 문서를 편집하는 중일 수 있다. check 로 보여 준 판과 다르면
    # 아무도 읽지 않은 절을 '반영 완료'로 찍게 된다 (2026-09-13 실제로 6초 차이로 일어났다)
    seen = open(SEEN).read().strip() if os.path.exists(SEEN) else ""
    if seen != _sha(cur):
        print(f"✗ ack 거부 — check 이후 문서가 또 바뀌었다 (본 판 {seen or '없음'} / 현재 {_sha(cur)}).")
        print("  check 를 다시 돌려 새로 바뀐 절을 읽고 반영한 뒤 ack 하라.")
        sys.exit(1)
    os.makedirs(STATE, exist_ok=True)
    prev = open(SNAP, encoding="utf-8").read() if os.path.exists(SNAP) else ""
    open(SNAP, "w", encoding="utf-8").write(cur)
    so, sn = sections(prev), sections(cur)
    entry = {"at": datetime.datetime.now().isoformat(timespec="seconds"),
             "sha": _sha(cur),
             "sections": len(sn),
             "new": [h for h in sn if _key(h) not in {_key(x) for x in so}],
             "note": note}
    open(LOG, "a", encoding="utf-8").write(json.dumps(entry, ensure_ascii=False) + "\n")
    print(f"✓ 반영 완료로 기록 — {entry['sha']} · 절 {entry['sections']}개")


def log():
    if not os.path.exists(LOG):
        print("이력 없음"); return
    for l in open(LOG, encoding="utf-8"):
        e = json.loads(l)
        print(f"  {e['at']}  {e['sha']}  새 절 {len(e['new'])}  {e.get('note','')}")
        for h in e["new"]:
            print(f"      + {h}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("check")
    a = sub.add_parser("ack"); a.add_argument("--note", default="")
    sub.add_parser("log")
    c = sub.add_parser("set-path"); c.add_argument("path")
    n = ap.parse_args()
    if n.cmd == "check":
        sys.exit(check())
    elif n.cmd == "ack":
        ack(n.note)
    elif n.cmd == "log":
        log()
    else:
        os.makedirs(STATE, exist_ok=True)
        json.dump({"path": n.path}, open(CONF, "w", encoding="utf-8"), ensure_ascii=False)
        print("경로 설정:", n.path)
