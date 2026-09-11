"""ppt-deck 저수준 엔진 — 토큰 로딩, 폰트/도형 헬퍼, 그리드 계산.

설계 원칙
  1) 좌표는 grid.col_x()/span_w()만 통해 나온다. 눈대중 배치를 코드 레벨에서 차단.
  2) 글자 크기는 spec.styles에 있는 것만 쓴다. text()가 스타일 이름을 강제한다.
  3) 그림자/둥근모서리/그라데이션/테마 표스타일은 생성 경로 자체가 없다.
"""
from __future__ import annotations
import json, math, os, re
from dataclasses import dataclass, field

import yaml
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
SPEC_PATH = os.path.join(HERE, "..", "references", "deck-spec.yaml")

# 화살표(→), 기하기호(✕ ○ ⊃), 원문자(① ⑤)는 이모지가 아니라 활자다. 잡으면 안 된다.
EMOJI = re.compile(
    "[\U0001F000-\U0001FAFF]"                              # 그림문자 본체
    "|[\u2190-\u21FF\u2600-\u27BF\u2B00-\u2BFF]\uFE0F"   # 변형 선택자로 이모지화된 기호
    "|[\u2705\u274C\u2757\u2B50\u2764\u26A0]"            # 변형자 없이도 이모지로 통용되는 것들
)


# ---------------------------------------------------------------- spec / color
def load_spec(path: str | None = None) -> dict:
    with open(path or SPEC_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f)


def mix(a: str, b: str, t: float) -> str:
    """a를 b쪽으로 t만큼 섞는다. muted 색을 팔레트에서 파생시키기 위한 것."""
    ai = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    bi = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ai, bi))


@dataclass
class Palette:
    ground: str
    figure: str
    accent: str
    muted: str = ""

    @classmethod
    def build(cls, spec, name=None):
        name = name or spec["active_palette"]
        p = spec["palettes"][name]
        return cls(p["ground"], p["figure"], p["accent"],
                   mix(p["figure"], p["ground"], spec["muted_mix"]))


@dataclass
class Grid:
    margin_x: int; margin_top: int; margin_bottom: int
    columns: int; col_w: int; gutter: int
    width: int; height: int

    @classmethod
    def build(cls, spec):
        g, c = spec["grid"], spec["canvas"]
        return cls(g["margin_x"], g["margin_top"], g["margin_bottom"],
                   g["columns"], g["col_w"], g["gutter"],
                   c["width_pt"], c["height_pt"])

    @property
    def step(self): return self.col_w + self.gutter
    @property
    def content_w(self): return self.width - 2 * self.margin_x
    @property
    def right(self): return self.width - self.margin_x
    @property
    def bottom(self): return self.height - self.margin_bottom

    def x(self, col: int) -> float:
        """컬럼 인덱스(0-based)의 왼쪽 좌표."""
        return self.margin_x + col * self.step

    def w(self, span: int) -> float:
        """span개 컬럼의 폭(사이 거터 포함)."""
        return span * self.col_w + (span - 1) * self.gutter


# ---------------------------------------------------------------- 폰트 적용
def _apply_font(run, name, size, color, *, bold=False, tracking=0):
    f = run.font
    f.name = name                      # <a:latin>
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = RGBColor.from_string(color)
    rPr = f._rPr
    for tag in ("a:ea", "a:cs"):       # 한글은 ea를 안 넣으면 PowerPoint가 딴 폰트로 렌더한다
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)
    if tracking:
        rPr.set("spc", str(int(tracking)))


def _end_para(p, name, size, color):
    """<a:endParaRPr>를 마지막 run과 같은 서식으로 채운다."""
    pPr = p._p
    el = pPr.find(qn("a:endParaRPr"))
    if el is None:
        el = pPr.makeelement(qn("a:endParaRPr"), {})
        pPr.append(el)                      # endParaRPr는 <a:p>의 마지막 자식이어야 한다
    el.set("lang", "ko-KR")
    el.set("sz", str(int(round(size * 100))))
    for tag in ("a:latin", "a:ea", "a:cs"):
        c = el.find(qn(tag))
        if c is None:
            c = el.makeelement(qn(tag), {}); el.append(c)
        c.set("typeface", name)
    fill = el.find(qn("a:solidFill"))
    if fill is None:
        fill = el.makeelement(qn("a:solidFill"), {}); el.insert(0, fill)
    srgb = fill.find(qn("a:srgbClr"))
    if srgb is None:
        srgb = fill.makeelement(qn("a:srgbClr"), {}); fill.append(srgb)
    srgb.set("val", color)


def patch_theme(path: str, fonts: dict):
    """저장된 pptx의 theme1.xml 글꼴을 실제 쓰는 글꼴로 바꾼다.

    python-pptx 기본 테마는 latin=Calibri, script="Hang"=맑은 고딕이다. 그대로 두면
    endParaRPr를 채워도 새 글상자를 만들 때 그 테마가 튀어나온다.
    OOXML은 ElementTree로 왕복시키면 네임스페이스 접두사가 재작성되어 깨지므로 정규식으로 다룬다.
    """
    import re, shutil, zipfile
    head, body = fonts["head"], fonts["body"]
    src = zipfile.ZipFile(path)
    names = src.namelist()
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n in names:
            data = src.read(n)
            if n.startswith("ppt/theme/theme") and n.endswith(".xml"):
                x = data.decode("utf-8")
                x = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*(")',
                           lambda m: m.group(1) + head + m.group(2), x)
                x = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*(")',
                           lambda m: m.group(1) + body + m.group(2), x)
                x = re.sub(r'<a:font script="Hang" typeface="[^"]*"/>', "", x)
                x = x.replace("</a:majorFont>",
                              f'<a:font script="Hang" typeface="{head}"/></a:majorFont>')
                x = x.replace("</a:minorFont>",
                              f'<a:font script="Hang" typeface="{body}"/></a:minorFont>')
                data = x.encode("utf-8")
            out.writestr(n, data)
    src.close()
    shutil.move(tmp, path)


def embed_fonts(path: str, fonts: dict) -> int:
    """쓰이는 Pretendard 굵기를 .pptx 안에 박아 넣는다.

    발표 PC에 Pretendard가 없으면 대체 글꼴로 무너지는 것이 이 스킬의 가장 큰 실무 약점이었다.
    주의 — PowerPoint는 **TTF(글리프 아웃라인)만** 임베딩한다. OTF(CFF)는
    "저장할 수 없는 글꼴" 오류로 거부한다. metrics.py 가 찾은 파일이 .otf 면 건너뛴다.
    비용: 굵기당 1~2MB. 반환값은 실제로 박은 개수.
    """
    import re, shutil, zipfile
    import metrics as _m

    used, seen = [], set()
    for key in ("display", "head", "body", "label"):
        name = fonts[key]
        if name in seen:
            continue
        seen.add(name)
        f = _m._find(_m.WANT.get(name, "Pretendard-Regular"))
        if f and f.lower().endswith(".ttf"):
            used.append((name, f))
    if not used:
        return 0

    src = zipfile.ZipFile(path)
    names = src.namelist()
    parts = {n: src.read(n) for n in names}
    src.close()

    rels = parts["ppt/_rels/presentation.xml.rels"].decode()
    next_id = max([int(m) for m in re.findall(r'Id="rId(\d+)"', rels)] or [0]) + 1
    lst, adds = [], []
    for i, (name, file) in enumerate(used, start=1):
        rid, part = f"rId{next_id}", f"ppt/fonts/font{i}.fntdata"
        next_id += 1
        parts[part] = open(file, "rb").read()
        adds.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                    f'officeDocument/2006/relationships/font" Target="fonts/font{i}.fntdata"/>')
        lst.append(f'<p:embeddedFont><p:font typeface="{name}" pitchFamily="34" '
                   f'charset="-127"/><p:regular r:id="{rid}"/></p:embeddedFont>')
    parts["ppt/_rels/presentation.xml.rels"] = rels.replace(
        "</Relationships>", "".join(adds) + "</Relationships>").encode()

    ct = parts["[Content_Types].xml"].decode()
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace("<Types ", "<Types ", 1)
        ct = re.sub(r"(<Types[^>]*>)",
                    r'\1<Default Extension="fntdata" ContentType="application/x-fontdata"/>',
                    ct, count=1)
    parts["[Content_Types].xml"] = ct.encode()

    # CT_Presentation 의 자식 순서상 embeddedFontLst 는 notesSz 다음이다. 순서를 어기면 열리지 않는다.
    pres = parts["ppt/presentation.xml"].decode()
    if "<p:embeddedFontLst>" not in pres:
        block = "<p:embeddedFontLst>" + "".join(lst) + "</p:embeddedFontLst>"
        m = re.search(r"<p:notesSz[^>]*/>", pres)
        if not m:
            return 0
        pres = pres[:m.end()] + block + pres[m.end():]
    parts["ppt/presentation.xml"] = pres.encode()

    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        out.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for n, data in parts.items():
            out.writestr(n, data)
    shutil.move(tmp, path)
    return len(used)


# ---------------------------------------------------------------- Deck
class Deck:
    def __init__(self, spec: dict, palette: str | None = None, density: str | None = None):
        self.spec = spec
        self.pal = Palette.build(spec, palette)
        self.g = Grid.build(spec)
        self.density = density or spec["density"]
        self.limits = spec["density_limits"][self.density]
        self.prs = Presentation()
        self.prs.slide_width = Pt(self.g.width)
        self.prs.slide_height = Pt(self.g.height)
        self._blank = self.prs.slide_layouts[6]
        self.manifest: list[dict] = []
        self.shapes: list[dict] = []
        self.images: list[dict] = []
        self._tmpdir = os.path.join(os.environ.get("TMPDIR", "/tmp"),
                                    f"ppt-deck-{os.getpid()}")
        self.grounds: dict[int, str] = {}
        self._slide_i = 0
        self._invert = False        # True면 ground/figure를 맞바꾼다 (밝은 장/어두운 장 교차)

    # ---- 색 이름 해석 -------------------------------------------------
    def c(self, name: str) -> str:
        if self._invert and name in ("ground", "figure"):
            name = "figure" if name == "ground" else "ground"
        return {"ground": self.pal.ground, "figure": self.pal.figure,
                "accent": self.pal.accent, "muted": self.pal.muted}[name]

    # ---- 슬라이드 ------------------------------------------------------
    def slide(self, bg="ground", invert=False):
        self._invert = invert
        s = self.prs.slides.add_slide(self._blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(self.c(bg))
        self._slide_i += 1
        self.grounds[self._slide_i] = self.c(bg)
        return s

    # ---- 도형 ----------------------------------------------------------
    def block(self, s, x, y, w, h, color="accent"):
        """평면 사각형. 테두리·그림자·둥근모서리 없음."""
        sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(self.c(color))
        sh.line.fill.background()
        sh.shadow.inherit = False
        sh.text_frame.text = ""
        self.shapes.append(dict(slide=self._slide_i, x=x, y=y, w=w, h=h,
                                color=self.c(color)))
        return sh

    def rule(self, s, x, y, w, thickness=3, color="accent"):
        """헤어라인. 장식이 아니라 구조 표시용."""
        return self.block(s, x, y, w, thickness, color)

    # ---- 이미지 --------------------------------------------------------
    def picture(self, s, path, x, y, w, h, fit="cover", tag="image", focus="center"):
        """상자에 이미지를 앉힌다.

        cover  = 상자를 꽉 채우고 넘치는 쪽을 잘라낸다(비율 유지). 반출혈 배치용.
        contain = 상자 안에 통째로 넣고 남는 쪽을 여백으로 둔다. 도표·스크린샷용.

        PowerPoint의 crop 속성에 맡기지 않고 PIL로 미리 잘라 넣는다 — 잘린 원본이
        파일에 남지 않으므로 용량이 줄고, 미리보기와 실제 렌더가 반드시 일치한다.
        """
        from PIL import Image
        path = os.path.expanduser(str(path))
        if not os.path.exists(path):
            raise FileNotFoundError(f"이미지 없음: {path}")
        im = Image.open(path)
        im = im.convert("RGB") if im.mode not in ("RGB", "RGBA") else im
        iw, ih = im.size
        box_ar, img_ar = w / h, iw / ih

        if fit == "cover":
            if img_ar > box_ar:                      # 원본이 더 넓다 → 좌우를 자른다
                nw = int(ih * box_ar)
                im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
            else:                                    # 원본이 더 높다 → 위아래를 자른다
                nh = int(iw / box_ar)
                # 인물 사진을 가운데로 자르면 머리가 날아간다. focus 로 기준선을 고른다.
                top = {"top": 0, "center": (ih - nh) // 2, "bottom": ih - nh}[focus]
                im = im.crop((0, top, iw, top + nh))
            dx, dy, dw, dh = x, y, w, h
        else:
            sc = min(w / iw, h / ih)
            dw, dh = iw * sc, ih * sc
            dx, dy = x + (w - dw) / 2, y + (h - dh) / 2

        target_px = int(max(dw, dh) * 2.2)           # 인쇄까지 버티게 2배 조금 넘게
        if max(im.size) > target_px:
            im.thumbnail((target_px, target_px), Image.LANCZOS)

        os.makedirs(self._tmpdir, exist_ok=True)
        tmp = os.path.join(self._tmpdir, f"img{len(self.images):03d}.png")
        im.save(tmp)
        s.shapes.add_picture(tmp, Pt(dx), Pt(dy), Pt(dw), Pt(dh))
        self.images.append(dict(slide=self._slide_i, x=dx, y=dy, w=dw, h=dh,
                                src=tmp, tag=tag))
        return dx, dy, dw, dh

    # ---- 텍스트 --------------------------------------------------------
    def text(self, s, style: str, x, y, w, h, content, *,
             color="figure", align="left", anchor="top", accent_lead=False,
             space_after=0, tag="", suffix=None, suffix_scale=0.4,
             suffix_color="accent"):
        """content: str 또는 list[str](문단들). style은 spec.styles의 키여야 한다."""
        st = self.spec["styles"][style]
        font = self.spec["fonts"][st["font"]]
        paras = content if isinstance(content, list) else [content]

        box = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}[anchor]

        for i, ptext in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT,
                           "center": PP_ALIGN.CENTER}[align]
            p.line_spacing = st["leading"]
            if space_after and i < len(paras) - 1:
                p.space_after = Pt(space_after)
            if accent_lead:                     # 불릿 대신 짧은 대시 — 기본 불릿 글머리표를 쓰지 않는다
                gap = st["size"] * 0.85          # 대시-본문 간격. 공백문자는 폰트마다 들쭉날쭉해 못 믿는다
                hang = int((st["size"] + gap) * 12700)
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", str(hang)); pPr.set("indent", str(-hang))
                r = p.add_run(); r.text = "\u2014"
                _apply_font(r, font, st["size"], self.c("accent"),
                            tracking=int(gap * 100))
            r = p.add_run(); r.text = str(ptext)
            _apply_font(r, font, st["size"], self.c(color),
                        tracking=st.get("tracking", 0))
            _end_para(p, font, st["size"], self.c(color))
            if suffix and i == len(paras) - 1:
                sfx = str(suffix)
                if sfx[0].isalpha() or "가" <= sfx[0] <= "힣":   # 27% 는 붙이고 62 시간 은 띄운다
                    sfx = "\u202f" + sfx        # 붙임 공백. U+2009는 여기서 줄을 끊어버린다
                r = p.add_run(); r.text = sfx
                _apply_font(r, font, st["size"] * suffix_scale, self.c(suffix_color))

        extra_pt = 0.0
        if suffix:                      # 단위는 본문보다 작게 찍힌다. 본문 크기로 세면 과대추정이다.
            sfx = str(suffix)
            if sfx[0].isalpha() or "가" <= sfx[0] <= "힣":
                sfx = "\u202f" + sfx
            extra_pt = est_adv(sfx, font) * st["size"] * suffix_scale
        self.manifest.append(dict(slide=self._slide_i, tag=tag or style, style=style,
                                  x=x, y=y, w=w, h=h, size=st["size"],
                                  leading=st["leading"], accent_lead=accent_lead,
                                  space_after=space_after, extra_pt=extra_pt,
                                  font=font, tracking=st.get("tracking", 0),
                                  color=self.c(color), align=align, anchor=anchor,
                                  suffix=str(suffix) if suffix else None,
                                  suffix_scale=suffix_scale,
                                  suffix_color=self.c(suffix_color),
                                  text=[str(t) for t in paras]))
        return box

    # ---- 쪽번호 --------------------------------------------------------
    def page_number(self, s, n):
        """기본은 우하단. 거기에 이미지가 깔려 있으면 반대쪽으로 피한다 —
        사진 위의 회색 숫자는 안 보이거나 지저분하다."""
        g = self.g
        x, align = g.right - 60, "right"
        for im in self.images:
            if im["slide"] != self._slide_i:
                continue
            if (im["x"] < g.right and im["x"] + im["w"] > g.right - 60
                    and im["y"] < g.height - 18 and im["y"] + im["h"] > g.height - 34):
                x, align = g.margin_x, "left"
                break
        else:
            pass
        if align == "left":
            for im in self.images:          # 왼쪽도 덮였으면 아예 찍지 않는다
                if im["slide"] == self._slide_i and im["x"] <= g.margin_x \
                        and im["x"] + im["w"] > g.margin_x + 60 \
                        and im["y"] + im["h"] > g.height - 34:
                    return
        self.text(s, "label", x, g.height - 34, 60, 16, f"{n:02d}",
                  color="muted", align=align, tag="pagenum")

    def save(self, path, embed=False):
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        self.prs.save(path)
        patch_theme(path, self.spec["fonts"])
        self.embedded = embed_fonts(path, self.spec["fonts"]) if embed else 0
        mpath = os.path.splitext(path)[0] + ".manifest.json"
        with open(mpath, "w", encoding="utf-8") as f:
            json.dump({"spec_styles": self.spec["styles"], "palette": self.pal.__dict__,
                       "density": self.density, "canvas": [self.g.width, self.g.height],
                       "grounds": self.grounds, "shapes": self.shapes,
                       "images": self.images,
                       "boxes": self.manifest}, f,
                      ensure_ascii=False, indent=1)
        return path, mpath


# ---------------------------------------------------------------- 넘침 추정
def est_adv(text: str, font: str | None = None) -> float:
    """문자열의 가로 폭을 em 단위로 계산.

    설치된 Pretendard에서 뽑은 실측 글리프 폭을 쓴다(scripts/metrics.py).
    캐시가 없을 때만 어림값으로 물러난다 — 어림값은 한글을 1.0em으로 잡아 16% 과대추정한다.
    """
    import metrics as _m
    tbl = (_m.load() or {}).get(font or "")
    if tbl:
        adv, hangul, dflt = tbl["adv"], tbl["hangul"], tbl["default"]
        total = 0.0
        for ch in text:
            cp = str(ord(ch))
            if cp in adv:
                total += adv[cp]
            elif "\uac00" <= ch <= "\ud7a3":
                total += hangul
            else:
                total += dflt
        return total
    out = 0.0
    for ch in text:                     # 폴백: 실측 캐시가 없을 때만
        if ch in " \u2009\u202f":
            out += 0.28 if ch == " " else 0.14
        elif "가" <= ch <= "힣" or "㄰" <= ch <= "㆏" or "一" <= ch <= "鿿":
            out += 0.87
        elif ch.isdigit():
            out += 0.62
        else:
            out += 0.52
    return out


def est_lines(text: str, size: float, width: float, extra_pt: float = 0.0,
              font: str | None = None) -> int:
    if width <= 0:
        return 1
    return max(1, math.ceil((est_adv(text, font) * size + extra_pt) / width))


def est_height(paras, size, leading, width, space_after=0, accent_lead=False,
               extra_pt=0.0, font=None):
    last = len(paras) - 1
    lines = sum(est_lines(("— " if accent_lead else "") + str(p), size, width,
                          extra_pt if i == last else 0.0, font)
                for i, p in enumerate(paras))
    return lines * size * leading + space_after * max(0, last)
