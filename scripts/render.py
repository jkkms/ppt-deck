#!/usr/bin/env python3
"""ppt-deck 검수 렌더러 — .pptx -> PDF -> 슬라이드별 PNG.

눈으로 확인하지 않은 덱은 반드시 어딘가 넘친다. 이 단계를 건너뛰지 마라.
경로: LibreOffice(soffice)가 있으면 그걸 쓰고, 없으면 PowerPoint를 AppleScript로 부른다.
사용:  render.py out/deck.pptx [--dpi 110] [--only 1,3,7]
"""
from __future__ import annotations
import argparse, glob, os, shutil, subprocess, sys, tempfile

SOFFICE_CANDIDATES = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]


def find_soffice():
    for c in SOFFICE_CANDIDATES:
        p = shutil.which(c) or (c if os.path.exists(c) else None)
        if p: return p
    return None


def to_pdf(pptx: str, outdir: str, allow_powerpoint: bool = False) -> str:
    pptx = os.path.abspath(pptx)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    so = find_soffice()
    if so:
        subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx],
                       check=True, capture_output=True, timeout=180)
        return pdf
    if not allow_powerpoint:
        raise SystemExit(
            "LibreOffice(soffice)가 없어 렌더할 수 없다.\n"
            "  brew install --cask libreoffice   ← 헤드리스로 조용히 돌아간다 (권장)\n"
            "PowerPoint로 대신 뽑으려면 --allow-powerpoint 를 붙여라. "
            "단 앱이 뜨고 macOS 자동화 권한을 요구한다.")
    if not os.path.exists("/Applications/Microsoft PowerPoint.app"):
        raise SystemExit("LibreOffice도 PowerPoint도 없다.  brew install --cask libreoffice")
    script = f'''
    tell application "Microsoft PowerPoint"
        activate
        open POSIX file "{pptx}"
        delay 1
        save active presentation in POSIX file "{pdf}" as save as PDF
        delay 1
        close active presentation saving no
    end tell'''
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=300)
    if r.returncode != 0 or not os.path.exists(pdf):
        raise SystemExit("PowerPoint 변환 실패 (자동화 권한을 물었을 수 있다):\n" + r.stderr)
    return pdf


def main(pptx, dpi, only, allow_powerpoint=False):
    outdir = os.path.join(os.path.dirname(os.path.abspath(pptx)), "preview")
    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        pdf = to_pdf(pptx, td, allow_powerpoint)
        import pymupdf as fitz
        doc = fitz.open(pdf)
        want = set(only) if only else None
        made = []
        for i, page in enumerate(doc, 1):
            if want and i not in want: continue
            png = os.path.join(outdir, f"slide-{i:02d}.png")
            page.get_pixmap(dpi=dpi).save(png)
            made.append(png)
        doc.close()
    print(f"✓ {len(made)} png -> {outdir}")
    for p in made: print("  " + p)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx"); ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--only", default="")
    ap.add_argument("--allow-powerpoint", action="store_true",
                    help="LibreOffice가 없을 때 PowerPoint 자동화로 대체 (권한 요구)")
    a = ap.parse_args()
    main(a.pptx, a.dpi, [int(x) for x in a.only.split(",") if x.strip()],
         a.allow_powerpoint)
