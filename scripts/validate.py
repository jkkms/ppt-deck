#!/usr/bin/env python3
"""생성된 .pptx 무결성 검사. Anthropic pptx 스킬의 validate 단계를 이 빌더에 맞게 옮긴 것.

검사 항목
  1. ZIP·XML 형식      — 모든 파트가 well-formed 인가
  2. 콘텐츠 타입       — 슬라이드마다 Override 가 있는가
  3. 관계(rels)        — 슬라이드 rels 의 Target 이 실제 존재하는가, 고아 파트는 없는가
  4. 글꼴              — 모든 run 에 latin·ea·cs 가 다 지정됐는가 (한글은 ea 가 없으면 딴 글꼴로 렌더)
  5. endParaRPr        — 문단 끝 서식이 비어 있지 않은가 (글상자 클릭 시 글꼴 튐)
  6. 테마              — Calibri·맑은 고딕 같은 기본값이 남아 있지 않은가
  7. 금지 효과         — 그림자·그라데이션·둥근모서리가 섞이지 않았는가

사용:  validate.py deck.pptx      종료코드: 실패 있으면 1
"""
from __future__ import annotations
import argparse, os, re, sys, zipfile
from xml.dom.minidom import parseString

FAIL, WARN = [], []


def fail(msg, fix): FAIL.append(f"{msg}\n      고치는 법: {fix}")
def warn(msg): WARN.append(msg)


def main(path):
    if not zipfile.is_zipfile(path):
        print("✗ ZIP 이 아니다 — 파일이 깨졌다"); return 1
    z = zipfile.ZipFile(path)
    names = set(z.namelist())

    bad = z.testzip()
    if bad:
        fail(f"ZIP 손상: {bad}", "다시 빌드하라")

    for n in sorted(names):
        if n.endswith((".xml", ".rels")):
            try:
                parseString(z.read(n))
            except Exception as e:
                fail(f"XML 오류 {n}: {e}", "빌더에서 고쳐라. 패킹된 XML을 손으로 만지지 말 것")

    ct = z.read("[Content_Types].xml").decode() if "[Content_Types].xml" in names else ""
    slides = sorted(n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
    if not slides:
        fail("슬라이드가 하나도 없다", "아웃라인에 slides 가 비었는지 확인하라")
    for sl in slides:
        if f'PartName="/{sl}"' not in ct:
            fail(f"콘텐츠 타입 누락: {sl}",
                 "[Content_Types].xml 에 Override 를 추가하라")
        rels = f"ppt/slides/_rels/{os.path.basename(sl)}.rels"
        if rels in names:
            for tgt in re.findall(r'Target="([^"]+)"', z.read(rels).decode()):
                if tgt.startswith(("http://", "https://", "../")) is False:
                    continue
                res = os.path.normpath(os.path.join("ppt/slides", tgt)).replace("\\", "/")
                if tgt.startswith("../") and res not in names:
                    fail(f"{sl} 의 관계가 끊겼다 → {tgt}",
                         "참조하는 파트를 추가하거나 관계를 지워라")

    runs_missing, end_missing, n_runs = 0, 0, 0
    for sl in slides:
        x = z.read(sl).decode()
        for rpr in re.findall(r"<a:rPr\b[^>]*/?>(?:.*?</a:rPr>)?", x, re.S):
            if "<a:latin" not in rpr:
                continue
            n_runs += 1
            if "<a:ea" not in rpr or "<a:cs" not in rpr:
                runs_missing += 1
        for p in re.findall(r"<a:p>.*?</a:p>", x, re.S):
            if "<a:r>" in p and "<a:endParaRPr" not in p:
                end_missing += 1
        for prst, why in (("<a:gradFill", "그라데이션"),
                          ("<a:effectLst><a:outerShdw", "그림자")):
            if prst in x:
                warn(f"{os.path.basename(sl)}: {why} 발견 — 이 덱의 설계 규칙에 없는 효과다")

    if runs_missing:
        fail(f"run {runs_missing}/{n_runs} 개에 ea 또는 cs 타이프페이스가 없다",
             "deckkit._apply_font 가 a:latin·a:ea·a:cs 를 모두 쓰는지 확인하라. "
             "한글은 ea 가 없으면 PowerPoint 가 딴 글꼴로 렌더한다")
    if end_missing:
        fail(f"문단 {end_missing} 개의 endParaRPr 가 비었다",
             "deckkit._end_para 가 문단마다 호출되는지 확인하라. "
             "비면 글상자를 클릭하는 순간 테마 글꼴(Calibri·맑은 고딕)로 튄다")

    # 테마는 아랍어·태국어용 폴백 목록에 Arial 등을 원래 갖고 있다. 그건 정상이다.
    # 우리가 쓰는 자리(major/minorFont 의 latin, 그리고 한글 script="Hang")만 본다.
    theme = "".join(z.read(n).decode() for n in names if n.startswith("ppt/theme/"))
    slots = re.findall(r"<a:(?:major|minor)Font>(.*?)</a:(?:major|minor)Font>", theme, re.S)
    for blk in slots:
        latin = re.search(r"<a:latin typeface=\"([^\"]*)\"", blk)
        hang = re.search(r'script="Hang" typeface="([^"]*)"', blk)
        for got, where in ((latin.group(1) if latin else "", "latin"),
                           (hang.group(1) if hang else "", 'script="Hang"')):
            if not got:
                fail(f"테마 {where} 가 비었다",
                     "deckkit.patch_theme 이 major/minorFont 를 모두 채우는지 확인하라")
            elif got in ("Calibri", "Calibri Light", "맑은 고딕", "Malgun Gothic", "Aptos"):
                fail(f"테마 {where} 에 기본 글꼴 '{got}' 가 남아 있다",
                     "deckkit.patch_theme 을 확인하라. 이게 남으면 새 글상자가 이 글꼴로 생긴다")

    # 폰트를 임베드했다면 그 구조가 성립하는가
    fnt = [n for n in names if n.endswith(".fntdata")]
    if fnt:
        pres = z.read("ppt/presentation.xml").decode()
        prels = z.read("ppt/_rels/presentation.xml.rels").decode()
        if 'Extension="fntdata"' not in ct:
            fail("fntdata 의 Content-Type Default 가 없다",
                 '[Content_Types].xml 에 <Default Extension="fntdata" '
                 'ContentType="application/x-fontdata"/> 를 넣어라')
        ids = set(re.findall(r'Id="(rId\d+)"', prels))
        miss = set(re.findall(r'<p:regular r:id="(rId\d+)"/>', pres)) - ids
        if miss:
            fail(f"임베드 글꼴의 관계가 끊겼다: {sorted(miss)}",
                 "presentation.xml.rels 에 해당 Relationship 을 추가하라")
        kids = re.findall(r"<p:(notesSz|embeddedFontLst|defaultTextStyle)", pres)
        if "embeddedFontLst" in kids and kids.index("embeddedFontLst") == 0:
            fail("embeddedFontLst 가 notesSz 보다 앞에 있다",
                 "CT_Presentation 자식 순서를 지켜라 — 어기면 PowerPoint 가 파일을 못 연다")
        for n in fnt:
            if z.read(n)[:4] not in (b"\x00\x01\x00\x00", b"true", b"ttcf"):
                fail(f"{n} 이 TrueType 이 아니다",
                     "PowerPoint 는 OTF(CFF) 임베딩을 거부한다. Pretendard 는 "
                     "public/static/alternative 의 TTF 빌드를 설치하라")

    print(f"ppt-deck validate — {os.path.basename(path)}  "
          f"(슬라이드 {len(slides)}장 · run {n_runs}개)")
    for f in FAIL: print("  ✗ " + f)
    for w in WARN: print("  · " + w)
    if not FAIL and not WARN:
        print("  ✓ 걸린 것 없음")
    print(f"\n  FAIL {len(FAIL)} / WARN {len(WARN)}")
    return 1 if FAIL else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("pptx")
    sys.exit(main(ap.parse_args().pptx))
