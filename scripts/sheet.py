# -*- coding: utf-8 -*-
"""문항지(A4 PDF)와 덱에 넣을 문항 조각을 같은 HTML 에서 만든다 — house 판 카드 덱의 `orig` 장용.

2026-11-03 ~ 11-18 면접 모의 면접 시리즈(4회 · 5회 · 6회 · 추가 문항)의 강의 폴더 스크립트 다섯 개에 같은 코드가 있어
2026-11-23 스킬로 올렸다. 직접 조판한 문항지(원본 문구를 한 글자씩 옮긴 것)는 HTML → 크롬(화면 없이) PDF 로 찍고,
같은 글을 조각마다 PDF 한 쪽에 찍어 200dpi 로 그린 뒤 잉크까지만 잘라 틀을 두른다 — 문항지와 슬라이드의 글이 한 글자도 다르지 않다.

  import sheet as SH
  SH.sheet_pdf([("모의 면접 5회 · 학업역량 문항 1", BODY1, "출처 · …(문구는 원본 그대로, 서식만 바꿈)"), …], OUT, workdir=HERE)
  man = SH.pieces([("1", HTML1a), ("1", HTML1b), ("2", HTML2)], CROPS, workdir=HERE)   # → {문항: [{file, w, h}]}
  yaml_text = SH.orig_slides(man, "1", "../figs/orig_x", "문항 1 · 문항지 ({k}/{n}) · ○○과학고 2025학년도 소집면접")

- 이미지는 workdir 기준 상대 경로로 적는다(임시 HTML 을 workdir 에 쓴다)
- 같은 문항이 다음 쪽으로 이어지면 아래 오른쪽에 「※ 뒷장에 계속」 — 실제 소집면접 문항지 양식(문항마다 새 쪽)
- 쪽마다 본문 아래 끝을 재어 한도(쪽 아래 글 위)를 넘으면 알린다. 쪽 수가 예상과 다르면 멈춘다
글꼴: Pretendard, 없는 글리프(㉠ ㉮ ∠ ∥ ∽)는 Apple SD Gothic Neo 로 넘어간다. 아래첨자는 <sub> 대신 유니코드(H₂O, x₁)."""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import urllib.parse

CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
DPI = 200
MARGIN_MM = 12.7

CSS = """
@page { size: 210mm 297mm; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: white; }
body { font-family: Pretendard, "Apple SD Gothic Neo", sans-serif; color: #161C18; font-size: 10.6pt; line-height: 1.72; word-break: keep-all; -webkit-print-color-adjust: exact; }
.page { width: 210mm; height: 297mm; position: relative; page-break-after: always; overflow: hidden; }
.head { position: absolute; top: 12mm; left: 12.7mm; right: 12.7mm; text-align: center; font-weight: 600; font-size: 13pt; border-bottom: 0.8pt solid #161C18; padding-bottom: 1.6mm; line-height: 1.3; }
.body { position: absolute; top: 26mm; left: 12.7mm; right: 12.7mm; }
.foot { position: absolute; bottom: 7mm; left: 12.7mm; right: 12.7mm; font-size: 7.5pt; color: #6B776F; display: flex; justify-content: space-between; line-height: 1.2; }
.cont { position: absolute; bottom: 14mm; right: 12.7mm; font-weight: 600; font-size: 10pt; line-height: 1.2; }
.piece-page { width: 210mm; height: 297mm; page-break-after: always; padding: 3mm 12.7mm; overflow: hidden; }
.note { margin: 0 0 2.2mm 0; font-weight: 600; }
.stem { margin: 0 0 1.2mm 0; }
.box { border: 0.7pt solid #161C18; padding: 2.6mm 3.4mm; margin: 0 0 3mm 0; }
.box p { margin: 0 0 1.2mm 0; }
.box p:last-child { margin-bottom: 0; }
.ind { text-indent: 1em; }
.q { margin: 0 0 2.6mm 0; padding-left: 9.5mm; text-indent: -9.5mm; }
.qn { font-weight: 600; }
.bu { font-weight: 600; text-decoration: underline; }
.ov { text-decoration: overline; text-decoration-thickness: 0.6pt; }
.fig { text-align: center; margin: 1.6mm 0 0.6mm 0; }
.fig img { display: inline-block; vertical-align: bottom; }
.cap { text-align: center; font-size: 9.6pt; font-weight: 600; margin: 0 0 1.2mm 0; }
.two { display: flex; gap: 4mm; align-items: center; }
.two .txt { flex: 1; }
table { border-collapse: collapse; margin: 1.4mm auto 1.6mm auto; font-size: 9.8pt; line-height: 1.38; }
th, td { border: 0.6pt solid #161C18; padding: 0.8mm 2.2mm; text-align: center; font-weight: 400; }
th { background: #EDF1EC; font-weight: 600; }
i { font-family: "Times New Roman", serif; font-size: 1.06em; }
.ibox { border: 0.6pt solid #161C18; padding: 1.6mm 3.4mm; margin: 1.6mm 2mm 0.6mm 2mm; }
.ibox ul, .dl { list-style: none; margin: 0; padding: 0; }
.ibox li, .dl li { padding-left: 3.6mm; text-indent: -3.6mm; line-height: 1.75; }
.ibox li::before { content: "•"; display: inline-block; width: 3.6mm; text-indent: 0; }
.dl li::before { content: "·"; display: inline-block; width: 3.6mm; text-indent: 0; font-weight: 600; }
/* 〈보기〉 〈규칙〉 — 이름표가 상자 위 테두리 가운데에 걸친다 */
.bogi { border: 0.6pt solid #161C18; padding: 2.8mm 3.4mm 1.8mm 3.4mm; margin: 3.4mm 0 2.6mm 0; position: relative; }
.bogi .lab { position: absolute; top: -2.6mm; left: 50%; transform: translateX(-50%); background: white; padding: 0 2mm; line-height: 1.2; }
.dh { font-weight: 600; margin: 0 0 0.4mm 0; }
.space { height: 34mm; }
"""


def doc(body, css=CSS):
    return f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>{css}</style></head><body>{body}</body></html>'


def chrome_pdf(html, pdf, workdir):
    """HTML 문자열 → PDF. 임시 HTML 을 workdir 에 써서 상대 경로 그림이 풀리게 한다."""
    with tempfile.TemporaryDirectory() as prof:
        path = os.path.join(workdir, "_tmp_sheet.html")
        open(path, "w", encoding="utf-8").write(html)
        try:
            subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
                            f"--user-data-dir={prof}", "--no-pdf-header-footer", f"--print-to-pdf={pdf}",
                            "file://" + urllib.parse.quote(path)], check=True, capture_output=True, timeout=120)
        finally:
            os.remove(path)


def sheet_pdf(pages, out, workdir, css=CSS):
    """pages = [(머리 글, 본문 HTML, 출처 줄)] — 같은 머리가 이어지면 앞 쪽에 「※ 뒷장에 계속」."""
    import pymupdf
    html = []
    for n, (head, body, src) in enumerate(pages, 1):
        cont = n < len(pages) and pages[n][0] == head
        html.append(f'<div class="page"><div class="head">{head}</div><div class="body">{body}</div>'
                    + ('<div class="cont">※ 뒷장에 계속</div>' if cont else "")
                    + f'<div class="foot"><span>{src}</span><span>{n}</span></div></div>')
    chrome_pdf(doc("".join(html), css), out, workdir)
    d = pymupdf.open(out)
    limit = d[0].rect.height - 14 / 25.4 * 72 - 12
    for i, p in enumerate(d):                    # 본문이 「뒷장에 계속」 · 쪽 아래 글 위에서 끝나는지
        ys = [b[3] for b in p.get_text("blocks") if "출처" not in b[4] and "뒷장에 계속" not in b[4]]
        ims = [p.get_image_bbox(x).y1 for x in p.get_images(full=True)]
        end = max(ys + ims) if ys or ims else 0
        print(f"문항지 {i + 1}쪽 본문 아래 끝 {end:.0f}pt / 한도 {limit:.0f}pt" + ("  ⚠ 넘침" if end > limit else ""))
    if d.page_count != len(pages):
        raise SystemExit(f"문항지 쪽 수 {d.page_count} — 예상 {len(pages)}. 한 쪽이 넘쳤다")
    return out


def pieces(items, out_dir, workdir, css=CSS, dpi=DPI, pad=14, frame=22):
    """items = [(문항 이름, 조각 HTML)] → out_dir/q{문항}_{k}.jpg + manifest.json.
    조각 폭은 모두 본문 폭으로 맞춘다 — 같은 문항 조각이 같은 배율로 들어가게. 글만 있는 조각은 흰 네모가 어색하니 틀을 두른다."""
    import numpy as np
    import pymupdf
    from PIL import Image, ImageDraw
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith(".jpg"):
            os.remove(os.path.join(out_dir, f))
    tmp = os.path.join(out_dir, "_pieces.pdf")
    chrome_pdf(doc("".join(f'<div class="piece-page">{h}</div>' for _, h in items), css), tmp, workdir)
    d = pymupdf.open(tmp)
    if d.page_count != len(items):
        raise SystemExit("조각 쪽 수가 예상과 다르다 — 조각 하나가 A4 한 쪽을 넘었다")
    man = {}
    mpx = int(MARGIN_MM / 25.4 * dpi)
    for i, (q, _) in enumerate(items):
        pm = d[i].get_pixmap(dpi=dpi)
        im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
        g = np.asarray(im.convert("L"))
        ys, xs = np.where(g < 245)
        crop = im.crop((max(0, xs.min() - pad), max(0, ys.min() - pad),
                        min(im.width, xs.max() + pad + 1), min(im.height, ys.max() + pad + 1)))
        W = int((210 - 2 * MARGIN_MM) / 25.4 * dpi) + 2 * pad
        canvas = Image.new("RGB", (W, crop.height), "white")
        canvas.paste(crop, (max(0, xs.min() - pad - mpx + pad), 0))
        framed = Image.new("RGB", (canvas.width + 2 * frame, canvas.height + 2 * frame), "white")
        framed.paste(canvas, (frame, frame))
        ImageDraw.Draw(framed).rectangle((0, 0, framed.width - 1, framed.height - 1), outline=(22, 28, 24), width=2)
        lst = man.setdefault(str(q), [])
        name = f"q{q}_{len(lst) + 1}.jpg"
        framed.save(os.path.join(out_dir, name), quality=92)
        lst.append({"file": name, "w": framed.width, "h": framed.height})
    json.dump(man, open(os.path.join(out_dir, "manifest.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    os.remove(tmp)
    return man


def scale(items, w_avail=844.0, h_avail=495.0 - 80.0, cap=72 / DPI * 1.5):
    """한 문항 조각들의 공통 배율 — 조각마다 들어갈 수 있는 최대 배율 가운데 최솟값.
    상한은 인쇄 크기의 1.5배(학생이 화면으로 10분 동안 읽는 모의 면접 문항). 기출 강의 덱은 1.35배."""
    return round(min(min(w_avail / it["w"], h_avail / it["h"], cap) for it in items) - 0.0005, 4)


def orig_slides(man, q, rel_dir, eyebrow, cap=72 / DPI * 1.5):
    """manifest 의 한 문항 → `orig` 장 YAML 조각. eyebrow 에는 {k} {n} 을 쓴다(「문항 1 · 문항지 ({k}/{n}) · 출처」)."""
    items = man[str(q)]
    sc = scale(items, cap=cap)
    out = []
    for k, it in enumerate(items, 1):
        out.append(f'''  - layout: orig
    eyebrow: "{eyebrow.format(k=k, n=len(items))}"
    image: "{rel_dir.rstrip('/')}/{it["file"]}"
    scale: {sc}
''')
    return "\n".join(out)
