#!/usr/bin/env python3
"""매니페스트 → 슬라이드별 SVG. LibreOffice도 PowerPoint도 쓰지 않는다.

ppt-master가 SVG를 중간 표현으로 두는 방식을 가져왔다. 빌더가 남긴 좌표·색·실측 글꼴 폭을
그대로 그리므로 구성·균형·색·여백은 신뢰할 수 있다.

한계(반드시 알고 볼 것): 줄바꿈을 빌더와 같은 메트릭으로 계산하므로 **빌더가 놓친 넘침은
여기서도 안 보인다.** 커닝·합자·PowerPoint 고유의 줄바꿈 규칙도 재현하지 않는다.
최종 확인은 PowerPoint에서 직접 열어 보는 것이다.

사용:  svgpreview.py deck.manifest.json [--only 1,4] [--sheet]
"""
from __future__ import annotations
import argparse, json, math, os, sys
from html import escape

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import _adv_em, load_spec


def est_adv(text, font=None):
    return sum(_adv_em(c, font) for c in str(text))


def wrap(text, size, width, font, extra_pt=0.0):
    """빌더와 같은 규칙으로 줄을 나눈다. 공백에서 끊고, 한 낱말이 넘치면 글자 단위로."""
    if not text:
        return [""]
    words, lines, cur = text.split(" "), [], ""
    for w in words:
        trial = (cur + " " + w) if cur else w
        if est_adv(trial, font) * size + (extra_pt if w is words[-1] else 0) <= width or not cur:
            cur = trial
        else:
            lines.append(cur); cur = w
    if cur:
        lines.append(cur)
    out = []
    for ln in lines:                      # 낱말 하나가 줄보다 길면 글자 단위로 쪼갠다
        while est_adv(ln, font) * size > width and len(ln) > 1:
            k = len(ln)
            while k > 1 and est_adv(ln[:k], font) * size > width:
                k -= 1
            out.append(ln[:k]); ln = ln[k:]
        out.append(ln)
    return out or [""]


def render(man, idx):
    W, H = man.get("canvas", [960, 540])
    ground = man.get("grounds", {}).get(str(idx), man["palette"]["ground"])
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}">',
             f'<rect width="{W}" height="{H}" fill="#{ground}"/>']

    for im_ in man.get("images", []):
        if im_["slide"] != idx:
            continue
        parts.append(f'<rect x="{im_["x"]:.1f}" y="{im_["y"]:.1f}" width="{im_["w"]:.1f}" '
                     f'height="{im_["h"]:.1f}" fill="#8A8F96"/>'
                     f'<text x="{im_["x"]+12:.0f}" y="{im_["y"]+24:.0f}" font-size="12" '
                     f'fill="#fff" font-family="Pretendard">[image]</text>')

    for sh in man.get("shapes", []):
        if sh["slide"] != idx:
            continue
        parts.append(f'<rect x="{sh["x"]:.1f}" y="{sh["y"]:.1f}" width="{sh["w"]:.1f}" '
                     f'height="{sh["h"]:.2f}" fill="#{sh["color"]}"/>')

    for b in man["boxes"]:
        if b["slide"] != idx:
            continue
        size, lead, font = b["size"], b["leading"], b.get("font", "Pretendard")
        trk = (b.get("tracking") or 0) / 100.0
        lines, meta = [], []
        for i, para in enumerate(b["text"]):
            pre = ""
            extra = b.get("extra_pt", 0) if i == len(b["text"]) - 1 else 0
            for sub in str(para).split("\n"):
                for j, ln in enumerate(wrap(pre + sub, size, b["w"], font, extra)):
                    lines.append(ln)
                    meta.append((i, j == 0))
            if b.get("space_after") and i < len(b["text"]) - 1:
                lines.append(None); meta.append((i, False))

        real = [l for l in lines if l is not None]
        if b.get("exact_center"):
            # 도형 안 글리프: 라인박스를 도형 중앙에 놓고 베이스라인을 메트릭으로 잡는다
            ASC, DESC = 0.952, 0.241
            lb = size * (ASC + DESC)
            y0 = b["y"] + (b["h"] - lb * len(real)) / 2
            y = y0 + size * ASC
            lead = ASC + DESC
        else:
            total = len(real) * size * lead + (b.get("space_after") or 0) * (len(lines) - len(real))
            y0 = {"top": b["y"], "middle": b["y"] + (b["h"] - total) / 2,
                  "bottom": b["y"] + b["h"] - total}[b.get("anchor", "top")]
            y = y0 + size * lead * 0.78            # 대략의 베이스라인

        for k, ln in enumerate(lines):
            if ln is None:
                y += b.get("space_after") or 0
                continue
            align = b.get("align", "left")
            anchor = {"left": "start", "right": "end", "center": "middle"}[align]
            x = {"left": b["x"], "right": b["x"] + b["w"],
                 "center": b["x"] + b["w"] / 2}[align]
            # 대시 머리글은 accent 색, 본문은 지정 색
            if False:
                hang = est_adv("—", font) * size + size * 0.85
                parts.append(f'<text x="{x:.1f}" y="{y:.1f}" font-family="{font}" '
                             f'font-size="{size:.1f}" fill="#{man["palette"]["accent"]}" '
                             f'xml:space="preserve">—</text>')
                parts.append(f'<text x="{x + hang:.1f}" y="{y:.1f}" font-family="{font}" '
                             f'font-size="{size:.1f}" fill="#{b["color"]}" '
                             f'letter-spacing="{trk:.2f}" xml:space="preserve">'
                             f'{escape(ln[2:])}</text>')
            else:
                tail = ""
                if b.get("suffix") and k == len(lines) - 1:
                    ss = size * 0.4
                    tail = (f'<tspan font-size="{ss:.1f}" fill="#{b["suffix_color"]}">'
                            f'{escape(b["suffix"])}</tspan>')
                parts.append(f'<text x="{x:.1f}" y="{y:.1f}" text-anchor="{anchor}" '
                             f'font-family="{font}" font-size="{size:.1f}" fill="#{b["color"]}" '
                             f'letter-spacing="{trk:.2f}" xml:space="preserve">'
                             f'{escape(ln)}{tail}</text>')
            y += size * lead
    parts.append("</svg>")
    return "\n".join(parts)


def _fontfile(name):
    import metrics as _m
    stem = _m.WANT.get(name, "Pretendard-Regular")
    return _m._find(stem)


def _fontfile(name):
    import metrics as _m
    stem = _m.WANT.get(name, "Pretendard-Regular")
    return _m._find(stem)


def raster(man, idx, scale=1.0):
    """SVG와 같은 좌표계를 PIL로 직접 그린다. 실제 Pretendard TTF를 써서
    글리프·줄바꿈·자간이 PowerPoint와 같은 메트릭으로 나온다."""
    from PIL import Image, ImageDraw, ImageFont
    W, H = man.get("canvas", [960, 540])
    W, H = int(W * scale), int(H * scale)
    ground = man.get("grounds", {}).get(str(idx), man["palette"]["ground"])
    img = Image.new("RGB", (W, H), "#" + ground)
    dr = ImageDraw.Draw(img)
    cache = {}

    def F(name, size):
        k = (name, round(size * scale))
        if k not in cache:
            fp = _fontfile(name)
            cache[k] = ImageFont.truetype(fp, max(1, round(size * scale)))
        return cache[k]

    def draw(x, y, text, name, size, color, trk=0.0):
        f = F(name, size)
        if name == "Consolas" and any("\uac00" <= c <= "\ud7a3" for c in text):
            # 고정폭에는 한글이 없다 — 실제 렌더와 같게 한글만 본문 글꼴로
            cx = x * scale
            for ch in text:
                fc = F("Pretendard", size) if "\uac00" <= ch <= "\ud7a3" else f
                dr.text((cx, y * scale), ch, font=fc, fill="#" + color, anchor="ls")
                cx += dr.textlength(ch, font=fc) + trk * scale
            return
        if not trk:
            dr.text((x * scale, y * scale), text, font=f, fill="#" + color, anchor="ls")
            return
        cx = x * scale                      # 자간이 있으면 글자 단위로 찍는다
        for ch in text:
            dr.text((cx, y * scale), ch, font=f, fill="#" + color, anchor="ls")
            cx += dr.textlength(ch, font=f) + trk * scale

    # 도형과 그림을 만든 순서대로 그린다 — 순서를 무시하면 패널이 그림을 덮는다
    layer = [("img", z) for z in man.get("images", []) if z["slide"] == idx] + \
            [("sh", z) for z in man.get("shapes", []) if z["slide"] == idx]
    layer.sort(key=lambda t: t[1].get("seq", 0))
    for kindtag, z in layer:
        if kindtag == "img":
            if not os.path.exists(z["src"]):
                continue
            ph = Image.open(z["src"]).convert("RGB").resize(
                (max(1, round(z["w"] * scale)), max(1, round(z["h"] * scale))), Image.LANCZOS)
            img.paste(ph, (round(z["x"] * scale), round(z["y"] * scale)))
            continue
        box = [z["x"] * scale, z["y"] * scale,
               (z["x"] + z["w"]) * scale, (z["y"] + z["h"]) * scale]
        k = z.get("kind")
        if k in ("badge", "node", "ring"):
            dr.ellipse(box, fill="#" + z["color"])
        elif k in ("panel", "chip"):
            dr.rounded_rectangle(box, radius=z.get("radius", 5) * scale,
                                 fill="#" + z["color"])
        else:
            dr.rectangle(box, fill="#" + z["color"])

    for b in man["boxes"]:
        if b["slide"] != idx:
            continue
        size, lead, font = b["size"], b["leading"], b.get("font", "Pretendard")
        trk = (b.get("tracking") or 0) / 100.0
        lines = []
        for i, para in enumerate(b["text"]):
            pre = ""
            extra = b.get("extra_pt", 0) if i == len(b["text"]) - 1 else 0
            for sub in str(para).split("\n"):
                lines += wrap(pre + sub, size, b["w"], font, extra)
            if b.get("space_after") and i < len(b["text"]) - 1:
                lines.append(None)
        real = [l for l in lines if l is not None]
        if b.get("exact_center"):
            # 도형 안 글리프: 라인박스를 도형 중앙에 놓고 베이스라인을 메트릭으로 잡는다
            ASC, DESC = 0.952, 0.241
            lb = size * (ASC + DESC)
            y0 = b["y"] + (b["h"] - lb * len(real)) / 2
            y = y0 + size * ASC
            lead = ASC + DESC
        else:
            total = len(real) * size * lead + (b.get("space_after") or 0) * (len(lines) - len(real))
            y0 = {"top": b["y"], "middle": b["y"] + (b["h"] - total) / 2,
                  "bottom": b["y"] + b["h"] - total}[b.get("anchor", "top")]
            y = y0 + size * lead * 0.78

        for k, ln in enumerate(lines):
            if ln is None:
                y += b.get("space_after") or 0
                continue
            f = F(font, size)
            adv = dr.textlength(ln, font=f) / scale + trk * max(0, len(ln) - 1)
            if b.get("suffix") and k == len(lines) - 1:
                adv += dr.textlength(b["suffix"], font=F(font, size * 0.4)) / scale
            x = {"left": b["x"], "right": b["x"] + b["w"] - adv,
                 "center": b["x"] + (b["w"] - adv) / 2}[b.get("align", "left")]
            if False:
                draw(x, y, "—", font, size, man["palette"]["accent"])
                hang = est_adv("—", font) * size + size * 0.85
                draw(x + hang, y, ln[2:], font, size, b["color"], trk)
            else:
                draw(x, y, ln, font, size, b["color"], trk)
                if b.get("suffix") and k == len(lines) - 1:
                    sx = x + dr.textlength(ln, font=f) / scale + trk * max(0, len(ln) - 1)
                    draw(sx, y, b["suffix"], font, size * 0.4,
                         b["suffix_color"])
            y += size * lead
    return img


def main(mpath, only, sheet, png=False, scale=1.2):
    man = json.load(open(mpath, encoding="utf-8"))
    n = max(b["slide"] for b in man["boxes"])
    outdir = os.path.join(os.path.dirname(os.path.abspath(mpath)), "svg")
    os.makedirs(outdir, exist_ok=True)
    want = only or range(1, n + 1)
    made = []
    for i in want:
        p = os.path.join(outdir, f"slide-{i:02d}.svg")
        open(p, "w", encoding="utf-8").write(render(man, i))
        made.append(p)
    if png:
        pdir = os.path.join(os.path.dirname(os.path.abspath(mpath)), "png")
        os.makedirs(pdir, exist_ok=True)
        for i in want:
            q = os.path.join(pdir, f"slide-{i:02d}.png")
            raster(man, i, scale).save(q)
            made.append(q)
    if sheet:                                   # 전체를 한 장에 늘어놓은 컨택트시트
        W, H = man.get("canvas", [960, 540])
        cols, gap, sc = 3, 24, 0.42
        cw, ch = W * sc, H * sc
        rows = math.ceil(len(want) / cols)
        body = [f'<svg xmlns="http://www.w3.org/2000/svg" '
                f'width="{cols*(cw+gap)+gap:.0f}" height="{rows*(ch+gap+18)+gap:.0f}" '
                f'viewBox="0 0 {cols*(cw+gap)+gap:.0f} {rows*(ch+gap+18)+gap:.0f}">',
                f'<rect width="100%" height="100%" fill="#9AA0A6"/>']
        for k, i in enumerate(want):
            r, c = divmod(k, cols)
            x, y = gap + c * (cw + gap), gap + r * (ch + gap + 18)
            inner = render(man, i).split("\n", 1)[1].rsplit("</svg>", 1)[0]
            body.append(f'<g transform="translate({x:.0f},{y:.0f}) scale({sc})">{inner}</g>')
            body.append(f'<text x="{x:.0f}" y="{y+ch+13:.0f}" font-family="Pretendard" '
                        f'font-size="11" fill="#fff">slide {i}</text>')
        body.append("</svg>")
        p = os.path.join(outdir, "contact-sheet.svg")
        open(p, "w", encoding="utf-8").write("\n".join(body))
        made.append(p)
    print(f"✓ SVG {len(made)}개 -> {outdir}")
    for p in made:
        print("  " + p)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest"); ap.add_argument("--only", default="")
    ap.add_argument("--sheet", action="store_true")
    ap.add_argument("--png", action="store_true", help="PIL로 실제 TTF를 써서 PNG까지 굽는다")
    ap.add_argument("--scale", type=float, default=1.2)
    a = ap.parse_args()
    main(a.manifest, [int(x) for x in a.only.split(",") if x.strip()], a.sheet,
         a.png, a.scale)
