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


# ---------------------------------------------------------------- spec
def load_spec(path: str | None = None) -> dict:
    with open(path or SPEC_PATH, encoding="utf-8") as f:
        spec = yaml.safe_load(f)
    _set_grid(spec)
    return spec


# ---------------------------------------------------------------- 좌표 체계
# 그리드 값은 스펙 파일에서 온다 — 프로파일(house / editorial)마다 다르다.
GRID = {"x0": 58, "step": 72, "col_w": 52, "gutter": 20, "right": 902}
CANVAS_W, CANVAS_H = 960, 540


def _set_grid(spec):
    g = spec["grid"]
    GRID.update(x0=g["margin_x"], step=g["col_w"] + g["gutter"],
                col_w=g["col_w"], gutter=g["gutter"],
                right=g["margin_x"] + 11 * (g["col_w"] + g["gutter"]) + g["col_w"])
    globals()["CANVAS_W"] = spec["canvas"]["width_pt"]
    globals()["CANVAS_H"] = spec["canvas"]["height_pt"]


def col_x(n: int) -> float:
    """1-based 컬럼 인덱스 -> x(pt)."""
    assert 1 <= n <= 12, f"컬럼 인덱스 {n} 범위 밖"
    return GRID["x0"] + (n - 1) * GRID["step"]


def span_w(s: int) -> float:
    """스팬 -> 폭(pt)."""
    assert 1 <= s <= 12, f"스팬 {s} 범위 밖"
    return GRID["step"] * s - GRID["gutter"]


def content_r() -> float:
    return GRID["right"]


# ---------------------------------------------------------------- §5 색 계산
def _hex_to_rgb(h: str):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _rgb_to_hex(r: float, g: float, b: float) -> str:
    return "{:02X}{:02X}{:02X}".format(
        *(max(0, min(255, int(round(v)))) for v in (r, g, b)))


def _rel_luminance(h: str) -> float:
    """WCAG 상대휘도."""
    out = []
    for c in _hex_to_rgb(h):
        s_ = c / 255.0
        out.append(s_ / 12.92 if s_ <= 0.04045 else ((s_ + 0.055) / 1.055) ** 2.4)
    r, g, b = out
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(h1: str, h2: str) -> float:
    l1, l2 = _rel_luminance(h1), _rel_luminance(h2)
    lo, hi = min(l1, l2), max(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def mix(figure: str, ground: str, t: float) -> str:
    """figure를 ground 쪽으로 t 비율 섞는다. sRGB 선형 보간."""
    f, g = _hex_to_rgb(figure), _hex_to_rgb(ground)
    return _rgb_to_hex(*(f[i] * (1 - t) + g[i] * t for i in range(3)))


def is_dark_ground(pal: dict) -> bool:
    """figure가 ground보다 밝으면 어두운 팔레트."""
    return _rel_luminance(pal["figure"]) > _rel_luminance(pal["ground"])


def invert_pal(pal: dict) -> dict:
    """반전 필드. accent는 반전하지 않는다."""
    return {"ground": pal["figure"], "figure": pal["ground"], "accent": pal["accent"]}


def derive(pal: dict, cfg: dict) -> dict:
    """팔레트 3색 -> 사용 색 전체. 새 hex 는 생기지 않는다(전부 figure->ground 파생).

    2단(figure·muted)으로는 작은 활자에서 위계가 안 선다. 사용자 덱의 실측 램프가
    4단(161C18 / 3D4841 / 6B776F / 97A29B)이라 그 구조를 따른다.
    """
    c = cfg["colors"]
    f, g = pal["figure"], pal["ground"]
    if "tone_ink2" in c:                       # house 프로파일
        out = {"ground": g, "figure": f, "accent": pal["accent"],
               "ink2":  mix(f, g, c["tone_ink2"]),
               "muted": mix(f, g, c["tone_muted"]),
               "faint": mix(f, g, c["tone_faint"]),
               "hairline": mix(f, g, c["hairline_mix"]),
               # 실측 — 패널 채움 EDF1EC = accent 를 ground 쪽으로 0.94, 보조 배지 6E9C7F = 0.37
               "accent_tint": mix(pal["accent"], g, c.get("accent_tint_mix", 0.94)),
               "accent_soft": mix(pal["accent"], g, c.get("accent_soft_mix", 0.37)),
               "accent_mid":  mix(pal["accent"], g, c.get("accent_mid_mix", 0.75)),
               "accent_pale": mix(pal["accent"], g, c.get("accent_pale_mix", 0.88))}
        assert contrast(out["ink2"], g) >= c["min_ratio_body"], \
            f"ink2 vs ground {contrast(out['ink2'], g):.2f}"
        assert contrast(out["muted"], g) >= c["min_ratio_caption"], \
            f"muted vs ground {contrast(out['muted'], g):.2f}"
        return out
    # editorial 프로파일 (v2)
    t = c["muted_mix_dark"] if is_dark_ground(pal) else c["muted_mix_light"]
    muted, hairline = mix(f, g, t), mix(f, g, c["hairline_mix"])
    assert contrast(muted, g) >= c["muted_min_ratio_ground"]
    assert contrast(muted, f) >= c["muted_min_ratio_figure"]
    return {"ground": g, "figure": f, "accent": pal["accent"],
            "ink2": f, "muted": muted, "faint": muted, "hairline": hairline,
            "accent_tint": mix(pal["accent"], g, 0.94),
            "accent_soft": mix(pal["accent"], g, 0.37),
            "accent_mid":  mix(pal["accent"], g, c.get("accent_mid_mix", 0.75)),
            "accent_pale": mix(pal["accent"], g, c.get("accent_pale_mix", 0.88))}


# ---------------------------------------------------------------- §6 줄수 계산
HANGUL_EM = 0.8643        # Pretendard 한글 실측 글리프 폭
LATIN_EM = 0.55           # 숫자·라틴 근사치 (실측 캐시가 있으면 그쪽이 우선)


def _adv_em(ch: str, font: str | None) -> float:
    """글자 하나의 em 폭. references/font-metrics.json 이 있으면 실측을 먼저 본다."""
    if font:
        import metrics as _m
        tbl = (_m.load() or {}).get(font)
        if tbl:
            v = tbl["adv"].get(str(ord(ch)))
            if v is not None:
                return v
            if "\uac00" <= ch <= "\ud7a3":
                return tbl["hangul"]
    return HANGUL_EM if ord(ch) >= 0x1100 else LATIN_EM


def advance(ch: str, size: float, tracking: int, font: str | None = None) -> float:
    """글자 하나의 진행폭(pt). tracking 단위는 1/100 pt."""
    return size * _adv_em(ch, font) + tracking / 100.0


def line_count(text: str, w: float, style: dict, font: str | None = None) -> int:
    """폭 w(pt) 안에서 필요한 줄 수. 명시적 줄바꿈을 우선 존중한다."""
    total = 0
    for para in str(text).split("\n"):
        used, lines = 0.0, 1
        for ch in para:
            a = advance(ch, style["size"], style["tracking"], font)
            if used + a > w:
                lines += 1
                used = a
            else:
                used += a
        total += lines
    return total


def block_w(text: str, style: dict, font: str | None = None) -> float:
    """한 줄로 놓았을 때의 폭(pt). 칩처럼 내용에 맞춰 크기가 정해지는 도형에 쓴다."""
    return sum(advance(ch, style["size"], style["tracking"], font) for ch in str(text))


def block_h(text: str, w: float, style: dict, font: str | None = None) -> float:
    """텍스트 박스 높이(pt) = size * leading * 줄수."""
    return style["size"] * style["leading"] * line_count(text, w, style, font)


# ---------------------------------------------------------------- 정렬 엔진
def resolve_y(block_height: float, spec: dict, top: float | None = None) -> float:
    """본문 블록 상단 y. 남는 높이의 block_bias 만큼만 위에 두고 나머지는 아래로.

    하단 정렬(bottom=484)은 폐기했다 — 내용이 적은 장에서 글자가 왼쪽 아래로 몰려
    화면이 버려진 것처럼 보였다. 정중앙보다 살짝 위가 안정적이다.
    """
    r = spec.get("rhythm_y")
    t = top if top is not None else (r["content_y"] if r else 56.0)
    bottom = r["content_bottom"] if r else 484.0
    bias = spec["anchors"].get("block_bias", 0.38)
    return t + max(0.0, (bottom - t - block_height) * bias)


# ---------------------------------------------------------------- §12 사각형 가드
def assert_plate_ok(w: float, h: float, has_text: bool, kind: str, spec=None):
    """렌더러가 사각형을 그릴 수 있는 경우는 정확히 이 셋뿐이다."""
    lim = (spec or {}).get("plates", {}).get("plate_min_side", 32)
    if kind == "plate":
        assert min(w, h) >= lim, f"plate 짧은 변 {min(w, h)} < {lim}"
        assert has_text, "활자를 담지 않는 채움 사각형은 생성 금지"
    elif kind in ("hero_rule", "table_rule", "connector"):
        pass                      # 두께는 스펙의 rules 그룹이 강제한다
    else:
        raise ValueError(f"허용되지 않은 사각형 종류: {kind}")


# ---------------------------------------------------------------- 폰트 적용
def _apply_font(run, name, size, color, *, bold=False, tracking=0, ea=None):
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
        # 고정폭(Consolas)에는 한글 글리프가 없다 -> ea 만 한글 글꼴로 돌린다
        el.set("typeface", (ea or name) if tag == "a:ea" else name)
    if tracking:
        rPr.set("spc", str(int(tracking)))


def _end_para(p, name, size, color, ea=None):
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
        c.set("typeface", (ea or name) if tag == "a:ea" else name)
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
        self.pal_name = palette or spec["active_palette"]
        self.base_pal = dict(spec["palettes"][self.pal_name])
        self.colors = derive(self.base_pal, spec)      # 현재 필드의 색
        self.field_inverted = False
        self.density = density or spec["density"]
        self.limits = spec["density_limits"][self.density]
        self.prs = Presentation()
        self.prs.slide_width = Pt(CANVAS_W)
        self.prs.slide_height = Pt(CANVAS_H)
        self._blank = self.prs.slide_layouts[6]
        self.manifest: list[dict] = []
        self.shapes: list[dict] = []
        self.images: list[dict] = []
        self._seq = 0
        self._tmpdir = os.path.join(os.environ.get("TMPDIR", "/tmp"),
                                    f"ppt-deck-{os.getpid()}")
        self.grounds: dict[int, str] = {}
        self._slide_i = 0
        self._plates: list[dict] = []

    # ---- 색 이름 해석 -------------------------------------------------
    def c(self, name: str) -> str:
        """현재 필드의 색. 반전 필드는 derive()를 다시 통과한 값이라
        muted 혼합비도 방향에 맞게 재계산돼 있다 (§5.3)."""
        return self.colors[name]

    # ---- 슬라이드 ------------------------------------------------------
    def slide(self, invert=False):
        """반전 필드는 슬라이드 배경색으로 처리한다. 사각형을 깔지 않는다 (§12)."""
        self.field_inverted = invert
        pal = invert_pal(self.base_pal) if invert else self.base_pal
        self.colors = derive(pal, self.spec)
        bg = "ground"
        s = self.prs.slides.add_slide(self._blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(self.c(bg))
        self._slide_i += 1
        self.grounds[self._slide_i] = self.c(bg)
        return s

    # ---- 도형 ----------------------------------------------------------
    def _rect(self, s, x, y, w, h, color, kind, has_text):
        assert_plate_ok(w, h, has_text, kind, self.spec)   # 이 셋 외의 사각형은 없다
        sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(self.c(color))
        sh.line.fill.background()
        sh.shadow.inherit = False
        sh.text_frame.text = ""
        self._seq += 1
        self.shapes.append(dict(slide=self._slide_i, x=x, y=y, w=w, h=h,
                                color=self.c(color), kind=kind, seq=self._seq))
        return sh

    def panel(self, s, x, y, w, h, color="accent_tint", radius=5.0):
        """살짝 둥근 채움 패널. 실측 — 반경 5pt, 채움은 accent 를 ground 쪽으로 0.94 섞은 톤.
        장식이 아니라 한 덩어리를 묶는 그릇이므로 반드시 안에 활자가 들어간다."""
        sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        try:
            sh.adjustments[0] = min(0.5, radius / max(1e-6, min(w, h)))
        except Exception:
            pass
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(self.c(color))
        sh.line.fill.background(); sh.shadow.inherit = False
        sh.text_frame.text = ""
        self._seq += 1
        self.shapes.append(dict(slide=self._slide_i, x=x, y=y, w=w, h=h,
                                color=self.c(color), seq=self._seq, kind="panel", radius=radius))
        return sh

    def chip(self, s, x, y, w, h, text="", color="accent", text_color="ground",
             style="small", radius=4.0):
        """번호 칩 / 눈썹 칩. 실측 47.5x23 (번호) · 147.6x30 (눈썹), accent 채움."""
        sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        try:
            sh.adjustments[0] = min(0.5, radius / max(1e-6, min(w, h)))
        except Exception:
            pass
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(self.c(color))
        sh.line.fill.background(); sh.shadow.inherit = False
        sh.text_frame.text = ""
        self._seq += 1
        self.shapes.append(dict(slide=self._slide_i, x=x, y=y, w=w, h=h,
                                color=self.c(color), seq=self._seq, kind="chip", radius=radius))
        if text:
            self.text(s, style, x, y, w, h, str(text), color=text_color,
                      align="center", anchor="middle", exact_center=True,
                      tag="chip-text")
        return sh

    def divider(self, s, x, y, w, color="hairline"):
        """전폭 구분선. 실측 845 길이로 73회 — 구획을 나누는 구조선이다."""
        return self._rect(s, x, y, w, self.spec["shapes"].get("divider_w", 1),
                          color, "connector", True)

    def badge(self, s, cx, cy, d=32.0, color="accent", glyph="", glyph_color="ground",
              style="small", kind="badge"):
        """원형 배지. 실측 지름 32(주) / 24(보조). 번호·기호를 담아 행의 시작점을 잡는다."""
        x, y = cx - d / 2, cy - d / 2
        sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Pt(x), Pt(y), Pt(d), Pt(d))
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(self.c(color))
        sh.line.fill.background(); sh.shadow.inherit = False
        sh.text_frame.text = ""
        self._seq += 1
        self.shapes.append(dict(slide=self._slide_i, x=x, y=y, w=d, h=d,
                                color=self.c(color), kind=kind, seq=self._seq))
        if glyph:
            # 박스를 원과 정확히 같은 사각형으로 두고 세로 중앙 앵커에 맡긴다.
            # Pretendard 는 라인박스 중심(0.3555em)과 글리프 잉크 중심(0.3535em)이
            # 거의 같아서, 이렇게만 하면 광학 중앙과 0.2% 안에서 일치한다.
            self.text(s, style, x, y, d, d, str(glyph), color=glyph_color,
                      align="center", anchor="middle", exact_center=True,
                      tag="badge-glyph")
        return sh

    def oval(self, s, x, y, w, h, color="accent_pale"):
        """임의 비율 타원. 포함관계(동심원) 다이어그램에만 쓴다."""
        sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Pt(x), Pt(y), Pt(w), Pt(h))
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(self.c(color))
        sh.line.fill.background(); sh.shadow.inherit = False
        sh.text_frame.text = ""
        self._seq += 1
        self.shapes.append(dict(slide=self._slide_i, x=x, y=y, w=w, h=h,
                                color=self.c(color), seq=self._seq, kind="ring"))
        return sh

    ICONS = ("arrow", "pencil", "quote", "question", "check")

    def icon(self, s, name, cx, cy, size, color="ground"):
        """배지 안 선 아이콘. 사용자 덱은 글리프(↳)가 아니라 선 아이콘 그림을 쓴다 —
        글리프는 폰트마다 획 굵기·중심이 달라 배지 안에서 뜬다 (소집면담00 12쪽 대조).
        원본 아이콘 실측 — 잉크는 캔버스의 75%, 획 굵기는 잉크 폭의 11.5%, 끝은 둥글게.
        잉크 bbox 를 캔버스 가운데로 옮겨 광학 중앙을 맞춘다.
        """
        from PIL import Image, ImageDraw
        if name not in self.ICONS:
            raise ValueError(f"아이콘 '{name}' 없음 — {self.ICONS}")
        R = 400
        im = Image.new("RGBA", (R, R), (0, 0, 0, 0))
        dr = ImageDraw.Draw(im)
        rgb = tuple(int(self.c(color)[i:i + 2], 16) for i in (0, 2, 4)) + (255,)
        lw = int(R * 0.075)

        def ln(pts):
            dr.line(pts, fill=rgb, width=lw, joint="curve")
            for p in (pts[0], pts[-1]):
                dr.ellipse([p[0] - lw / 2, p[1] - lw / 2, p[0] + lw / 2, p[1] + lw / 2], fill=rgb)

        u = R / 10
        if name == "arrow":                      # ↳ 아래로 내려와 오른쪽으로
            ln([(3 * u, 1.5 * u), (3 * u, 6 * u), (8 * u, 6 * u)])
            ln([(6 * u, 4 * u), (8 * u, 6 * u), (6 * u, 8 * u)])
        elif name == "pencil":
            ln([(2 * u, 8 * u), (2.4 * u, 6.2 * u), (6.8 * u, 1.8 * u), (8.2 * u, 3.2 * u),
                (3.8 * u, 7.6 * u), (2 * u, 8 * u)])
            ln([(5.6 * u, 3 * u), (7 * u, 4.4 * u)])
            ln([(4.6 * u, 8.2 * u), (8.4 * u, 8.2 * u)])
        elif name == "quote":
            for x0 in (2.2 * u, 5.8 * u):
                dr.ellipse([x0, 3 * u, x0 + 2.2 * u, 5.2 * u], fill=rgb)
                ln([(x0 + 2.0 * u, 4.3 * u), (x0 + 1.2 * u, 7.2 * u)])
        elif name == "question":
            dr.arc([3 * u, 1.5 * u, 7 * u, 5.5 * u], 180, 90, fill=rgb, width=lw)
            ln([(5 * u, 5.5 * u), (5 * u, 6.4 * u)])
            dr.ellipse([5 * u - lw * 0.7, 8.2 * u - lw * 0.7, 5 * u + lw * 0.7, 8.2 * u + lw * 0.7], fill=rgb)
        elif name == "check":
            ln([(2 * u, 5.2 * u), (4.2 * u, 7.4 * u), (8 * u, 2.8 * u)])
        bb = im.getbbox()
        ink = im.crop(bb)
        side = int(max(ink.size) / 0.75)
        sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        sq.paste(ink, ((side - ink.size[0]) // 2, (side - ink.size[1]) // 2))
        os.makedirs(self._tmpdir, exist_ok=True)
        tmp = os.path.join(self._tmpdir, f"icon-{name}-{self.c(color)}.png")
        sq.save(tmp)
        return self.picture(s, tmp, cx - size / 2, cy - size / 2, size, size, "contain", tag="icon")

    def glow(self, s, cx, cy, d, target=1.28):
        """그림 뒤 둥근 조명. 사용자 원칙 — Q&A 장 캐릭터 뒤 원 하나로 무게를 준다.

        새 hex 를 들이지 않는다: 바탕을 accent 쪽으로 섞는다. 섞는 비율은 고정값이 아니라
        바탕과의 대비로 정한다 — 실측 2차시 39쪽 바탕 0E1A2B 위 17304C 가 대비 1.28.
        고정 비율로는 어두운 accent(forest)에서 원이 사라졌다.
        accent 로 그 대비가 안 나오면 figure 쪽으로 섞는다. 지름 338 (캐릭터 폭의 0.89배).
        그림을 받치지 않는 원은 lint_deck 가 장식으로 잡는다.
        """
        g = self.c("ground")
        color = g
        for toward in (self.c("accent"), self.c("figure")):
            for i in range(1, 101):
                cand = mix(g, toward, i / 100)
                if contrast(cand, g) >= target:
                    color = cand
                    break
            if color != g:
                break
        sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Pt(cx - d / 2), Pt(cy - d / 2), Pt(d), Pt(d))
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(color)
        sh.line.fill.background(); sh.shadow.inherit = False
        sh.text_frame.text = ""
        self._seq += 1
        self.shapes.append(dict(slide=self._slide_i, x=cx - d / 2, y=cy - d / 2, w=d, h=d,
                                color=color, seq=self._seq, kind="glow"))
        return sh

    def connector(self, s, x, y, h, w=1.0, color="accent_soft"):
        """배지와 배지를 잇는 세로선. 관계를 나타내는 선이라 장식선 금지에 걸리지 않는다."""
        return self._rect(s, x - w / 2, y, w, h, color, "connector", True)

    def plate(self, s, x, y, w, h, color="figure"):
        """채움 판. 짧은 변 >= 32 이고 반드시 활자를 담을 때만 허용된다."""
        return self._rect(s, x, y, w, h, color, "plate", True)

    def hero_rule(self, s, x, y, w):
        """data 히어로 밑줄. 허용된 구조선 1 — 2pt figure 단색."""
        return self._rect(s, x, y, w, self.spec["rules"]["hero_rule_w"],
                          "figure", "hero_rule", True)

    def table_rule(self, s, x, y, w):
        """table 행 구분선. 허용된 구조선 2 — 0.75pt hairline."""
        return self._rect(s, x, y, w, self.spec["rules"]["table_rule_w"],
                          "hairline", "table_rule", True)

    def notes(self, s, text):
        """자세한 설명은 슬라이드가 아니라 여기로 내린다.
        슬라이드는 시각 자료이고, 말로 할 것은 발표 노트에 적는다."""
        if not text:
            return
        s.notes_slide.notes_text_frame.text = str(text)
        self.note_count = getattr(self, "note_count", 0) + 1

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
        self._seq += 1
        self.images.append(dict(slide=self._slide_i, x=dx, y=dy, w=dw, h=dh,
                                src=tmp, tag=tag, seq=self._seq))
        return dx, dy, dw, dh

    # ---- 텍스트 --------------------------------------------------------
    def text(self, s, style: str, x, y, w, h, content, *,
             color="figure", align="left", anchor="top",
             space_after=0, tag="", suffix=None, suffix_style=None,
             suffix_color="figure", font_key=None, accent_paras=(),
             exact_center=False):
        """content: str 또는 list[str](문단들). style은 spec.styles의 키여야 한다."""
        st = self.spec["styles"][style]
        fkey = font_key or st["font"]
        font = self.spec["fonts"][fkey]
        ea = self.spec["fonts"]["body"] if fkey == "mono" else None
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
            # 정확 중앙 모드에서는 행간을 1.0 으로 둔다. leading 배수를 주면
            # PowerPoint 가 여분을 라인 위쪽에 몰아 넣어 글자가 아래로 내려간다.
            p.line_spacing = 1.0 if exact_center else st["leading"]
            if space_after and i < len(paras) - 1:
                p.space_after = Pt(space_after)
            if str(ptext) == "":
                # 빈 줄은 빈 문단으로만 둔다. 빈 run 은 PowerPoint 가 저장할 때 지워
                # 사용자 편집본과 도형 단위 대조가 어긋난다
                _end_para(p, font, st["size"], self.c(color), ea=ea)
                continue
            r = p.add_run(); r.text = str(ptext)
            _apply_font(r, font, st["size"],
                        self.c("accent") if i in accent_paras else self.c(color),
                        tracking=st.get("tracking", 0), ea=ea)
            _end_para(p, font, st["size"], self.c(color), ea=ea)
            if suffix and i == len(paras) - 1:
                sfx = str(suffix)
                if sfx[0].isalpha() or "가" <= sfx[0] <= "힣":   # 27% 는 붙이고 62 시간 은 띄운다
                    sfx = "\u202f" + sfx        # 붙임 공백. U+2009는 여기서 줄을 끊어버린다
                sst = self.spec["styles"][suffix_style or "stat_sub"]
                r = p.add_run(); r.text = sfx
                _apply_font(r, self.spec["fonts"][sst["font"]], sst["size"],
                            self.c(suffix_color), tracking=sst.get("tracking", 0))

        extra_pt = 0.0
        if suffix:                      # 단위는 제 스타일 크기로 잰다
            sfx = str(suffix)
            if sfx[0].isalpha() or "가" <= sfx[0] <= "힣":
                sfx = "\u202f" + sfx
            sst = self.spec["styles"][suffix_style or "stat_sub"]
            extra_pt = sum(_adv_em(c, self.spec["fonts"][sst["font"]]) for c in sfx) * sst["size"]
        self.manifest.append(dict(slide=self._slide_i, tag=tag or style, style=style,
                                  x=x, y=y, w=w, h=h, size=st["size"],
                                  leading=st["leading"],                                   space_after=space_after, extra_pt=extra_pt,
                                  exact_center=exact_center,
                                  font=font, tracking=st.get("tracking", 0),
                                  color=self.c(color), align=align, anchor=anchor,
                                  suffix=str(suffix) if suffix else None,
                                  suffix_style=suffix_style,
                                  suffix_color=self.c(suffix_color),
                                  text=[str(t) for t in paras]))
        return box

    # ---- 쪽번호 --------------------------------------------------------
    def save(self, path, embed=False):
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        self.prs.save(path)
        patch_theme(path, self.spec["fonts"])
        self.embedded = embed_fonts(path, self.spec["fonts"]) if embed else 0
        mpath = os.path.splitext(path)[0] + ".manifest.json"
        with open(mpath, "w", encoding="utf-8") as f:
            json.dump({"spec_styles": self.spec["styles"], "palette": self.colors,
                       "density": self.density, "canvas": [CANVAS_W, CANVAS_H],
                       "grounds": self.grounds, "shapes": self.shapes,
                       "images": self.images,
                       "boxes": self.manifest}, f,
                      ensure_ascii=False, indent=1)
        return path, mpath
