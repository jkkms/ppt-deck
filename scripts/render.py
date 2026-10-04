#!/usr/bin/env python3
"""ppt-deck 검수 렌더러 — .pptx -> PDF -> 슬라이드별 PNG.

눈으로 확인하지 않은 덱은 반드시 어딘가 넘친다. 이 단계를 건너뛰지 마라.

경로 (앞에서부터):
  1) LibreOffice(soffice) — 있으면 헤드리스로 조용히 돈다
  2) --app keynote     — Keynote 로 PDF 내보내기 (면담·면접 덱 감사에서 쓰던 길)
  3) --app powerpoint  — PowerPoint 로 PDF 저장 (Dive 덱 감사에서 쓰던 길). 발표를 PowerPoint 로
                         한다면 최종 확인은 이 렌더로 한다
  2)·3)은 앱이 뜨고 macOS 자동화 권한을 물으므로 **사용자가 허락했을 때만** 쓴다.

두 앱은 샌드박스라 ~/Downloads·iCloud Drive·임시 폴더를 못 읽거나 못 쓴다 — 못 써도 "성공"을
돌려주고 파일만 안 만든다. 그래서 ~/Documents/ppqa/ 아래에 복사해 놓고 거기로 내보낸다.

사용:  render.py out/deck.pptx [--dpi 110] [--only 1,3,7] [--app keynote|powerpoint] [--keep-pdf]
"""
from __future__ import annotations
import argparse, os, shutil, subprocess, tempfile, time

SOFFICE_CANDIDATES = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
STAGE = os.path.expanduser("~/Documents/ppqa")      # 샌드박스 앱이 읽고 쓸 수 있는 자리


def find_soffice():
    for c in SOFFICE_CANDIDATES:
        p = shutil.which(c) or (c if os.path.exists(c) else None)
        if p:
            return p
    return None


def _stage(pptx: str) -> tuple[str, str]:
    """~/Documents/ppqa/<이름>/ 에 복사하고 (복사본, 나올 PDF) 경로를 돌려준다."""
    stem = os.path.splitext(os.path.basename(pptx))[0]
    d = os.path.join(STAGE, stem)
    os.makedirs(d, exist_ok=True)
    src = os.path.join(d, stem + ".pptx")
    pdf = os.path.join(d, stem + ".pdf")
    shutil.copy2(pptx, src)
    if os.path.exists(pdf):
        os.remove(pdf)
    return src, pdf


def _wait(pdf: str, secs=30):
    for _ in range(secs * 2):
        if os.path.exists(pdf) and os.path.getsize(pdf) > 0:
            return True
        time.sleep(0.5)
    return False


def to_pdf(pptx: str, outdir: str, app: str | None = None) -> str:
    pptx = os.path.abspath(pptx)
    if app is None:
        so = find_soffice()
        if so:
            pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
            subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx],
                           check=True, capture_output=True, timeout=180)
            return pdf
        raise SystemExit(
            "LibreOffice(soffice)가 없어 렌더할 수 없다.\n"
            "  brew install --cask libreoffice   ← 헤드리스로 조용히 돌아간다\n"
            "  --app keynote | --app powerpoint  ← 앱이 뜨고 자동화 권한을 묻는다 (사용자 허락 후)")
    src, pdf = _stage(pptx)
    if app == "keynote":
        # 먼저 띄워 둔다 — 안 띄운 채 부르면 -600. -1712(시간 초과)는 화면에 대화상자가 떠 있다는 뜻
        subprocess.run(["open", "-a", "Keynote"], check=False)
        time.sleep(2)
        script = f'''
        tell application "Keynote"
            set d to open (POSIX file "{src}")
            export d to (POSIX file "{pdf}") as PDF with properties {{PDF image quality:Best}}
            close d saving no
        end tell'''
    elif app == "powerpoint":
        script = f'''
        tell application "Microsoft PowerPoint"
            activate
            open POSIX file "{src}"
            set d to active presentation
            save d in (POSIX file "{pdf}") as save as PDF
            close d saving no
        end tell'''
    else:
        raise SystemExit(f"--app 은 keynote 또는 powerpoint (받은 값 {app})")
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=300)
    if r.returncode != 0:
        msg = r.stderr.strip()
        if "-1712" in msg:
            msg += "\n  → 앱 화면에 대화상자가 떠 있다. 다시 시도하지 말고 사용자에게 창을 봐 달라고 한다"
        raise SystemExit(f"{app} 변환 실패 (자동화 권한을 물었을 수 있다):\n{msg}")
    if not _wait(pdf):
        raise SystemExit(f"{app} 이 성공을 돌려줬지만 PDF 가 없다 — 샌드박스 밖 경로였을 수 있다: {pdf}")
    return pdf


def main(pptx, dpi, only, app=None, keep_pdf=False):
    outdir = os.path.join(os.path.dirname(os.path.abspath(pptx)), "preview")
    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        pdf = to_pdf(pptx, td, app)
        if keep_pdf:
            shutil.copy2(pdf, os.path.join(outdir, os.path.basename(pdf)))
        import pymupdf as fitz
        doc = fitz.open(pdf)
        want = set(only) if only else None
        made = []
        for i, page in enumerate(doc, 1):
            if want and i not in want:
                continue
            png = os.path.join(outdir, f"slide-{i:02d}.png")
            page.get_pixmap(dpi=dpi).save(png)
            made.append(png)
        doc.close()
    print(f"✓ {len(made)} png -> {outdir}" + (f"  (렌더: {app})" if app else ""))
    for p in made:
        print("  " + p)
    return pdf


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--only", default="")
    ap.add_argument("--app", choices=["keynote", "powerpoint"],
                    help="LibreOffice 대신 앱으로 PDF 를 뽑는다 (자동화 권한 — 사용자 허락 후)")
    ap.add_argument("--allow-powerpoint", action="store_true",
                    help="옛 이름. --app powerpoint 와 같다")
    ap.add_argument("--keep-pdf", action="store_true")
    a = ap.parse_args()
    app = a.app or ("powerpoint" if a.allow_powerpoint else None)
    main(a.pptx, a.dpi, [int(x) for x in a.only.split(",") if x.strip()], app, a.keep_pdf)
