> **저장소 보관본.** editorial 프로파일(`references/deck-spec.editorial.yaml`)의 정본이다.
> 본문이 가리키는 `ppt-deck 재설계.dc.html` 명세서는 저장소에 두지 않는다 — 렌더 런타임이 따로
> 필요하고, 값은 이 문서 §4–§12 에 전부 옮겨 적혀 있다(§0.3). §2·§3 '토큰 표·좌표 표'를 인용한
> 곳은 이 문서의 §4(토큰)·§8(좌표)를 보면 된다.

# HANDOFF — ppt-deck 시각 재설계 적용

이 문서는 `ppt-deck 재설계.dc.html`(디자인 명세서)의 내용을 코드로 옮기기 위한 지시서다.
읽는 순서: §0 → §1 → §2 → §3 → 나머지.

---

## §0. 먼저 읽을 것 — 이 작업의 성격

### 0.1 하지 말 것 (가장 중요)

- **`ppt-deck 재설계.dc.html`을 구현하지 마라.** 이 파일은 산출물이 아니라 명세서다.
  HTML/CSS/React를 저장소로 옮기는 작업이 아니다. 웹 페이지를 만들지 않는다.
- **`support.js`를 읽지 마라.** 명세서를 브라우저에 렌더하기 위한 런타임이며 이 작업과 무관하다.
- **아트보드의 사선 무늬(`repeating-linear-gradient`)를 구현하지 마라.** 이미지 자리표시자
  표시일 뿐이다. 실제 렌더에서 그 위치에는 사용자 이미지가 들어간다.
- **문서 섹션(2·3·4·5)의 `border-top`, `background:#FAF9F6` 등을 슬라이드 디자인으로
  착각하지 마라.** 명세서 본문의 조판이며 PPTX와 무관하다.

### 0.2 할 일

`https://github.com/jkkms/ppt-deck` 의 다음 파일을 수정한다.

| 파일 | 작업 |
| --- | --- |
| `references/deck-spec.yaml` | 토큰 추가·삭제·수정 (§4) |
| `scripts/deckkit.py` | 색 계산(§5), 줄수 계산(§6), 정렬 엔진(§7) |
| `scripts/build.py` | 레이아웃 함수 11종 좌표 교체 (§8–§12) |
| `scripts/lint.py` | 신규 검사 6종 추가 (§13) |
| `references/outline.example.yaml` | `cards` 예시가 없으면 항목 2·3·4·6개 예시 추가 |

### 0.3 명세서 안에서 값을 찾는 방법

`ppt-deck 재설계.dc.html`은 다섯 부분이다.

1. **아트보드 17장** — 960×540pt 실치수 렌더. 각 요소가 `position:absolute; left/top/width`
   인라인 스타일로 pt 단위로 박혀 있다. **화면 좌표가 곧 pptx 좌표다.**
2. **각 아트보드 밑 모노스페이스 한 줄** — 그 보드의 좌표 요약.
3. **"2 · 토큰 표"** — 토큰 값의 **정본**.
4. **"3 · 좌표 표"** — 좌표의 **정본**.
5. **"4 · 일곱 가지에 대한 답" / "5 · 금지 항목 자기 점검"** — 각 변경의 의도. 왜 이렇게
   했는지 판단이 필요할 때 여기를 본다.

**충돌 시 우선순위: 3(좌표 표) > 1(아트보드 인라인 스타일) > 2(토큰 표) > 그 외.**
단 토큰 값(색·크기·자간)은 2가 정본이다. 어긋난 곳을 발견하면 고치지 말고 **보고**하라.

아래 §4–§12에 모든 값을 그대로 옮겨 적었으므로, 원칙적으로 HTML을 열지 않아도 작업이
가능하다. HTML은 눈으로 결과를 확인할 때만 쓴다.

---

## §1. 불변 제약 — 이걸 넘는 코드는 전부 되돌려야 한다

렌더 한계가 아니라 **의도적 금지 규칙**이다. "여기엔 필요한데요"는 없다.

```
캔버스        960 × 540 pt 고정 (16:9). 모든 단위는 pt.
그릴 수 있는 것  평면 사각형(autoshape RECTANGLE), 텍스트 박스, 이미지. 그게 전부.
금지          그림자, 둥근모서리(radius 항상 0), 그라데이션, 3D, 투명도(alpha),
             네이티브 표 스타일, 이모지, 아이콘, 일러스트, 이탤릭(합성 이탤릭은 한글에서 깨진다)
폰트          Pretendard 4웨이트만 — Black / SemiBold / Medium / Regular. 다른 서체 금지.
색            팔레트당 hex 3개(ground/figure/accent). 4번째 색이 필요하면
             색이 부족한 게 아니라 레이아웃이 틀린 거다.
정렬          본문은 좌측 정렬. 가운데 정렬은 제목에만 (현 설계에서는 0회).
구성          비대칭. 균등 3분할 같은 대칭 반복 금지.
선            허용된 두 개만 — data 히어로 밑줄(2pt), table 행 구분선(0.75pt).
             제목 밑 강조선 / 카드 모서리 액센트 줄 / 화면을 가로지르는 색 띠 전부 금지.
             (AI 생성 슬라이드의 대표적 지문이라 의도적으로 뺐다)
```

python-pptx 상의 구체적 주의:

- 모양 생성 후 **`shape.shadow.inherit = False`** 를 반드시 호출한다. python-pptx는 기본
  테마 그림자를 상속한다. 기존 코드에 헬퍼가 있으면 그걸 쓴다.
- `shape.line.fill.background()` 로 외곽선을 없앤다. 기본값이 선 있음인 경우가 있다.
- 자간(tracking)은 python-pptx에 API가 없다. `run._r.get_or_add_rPr().set('spc', str(v))`
  로 XML 속성을 직접 쓴다. 단위는 **1/100 pt** (예: `-200` = −2pt). 기존 코드에 이미
  헬퍼가 있다면 그것을 재사용한다.
- 행간(leading)은 `paragraph.line_spacing = <float>` (배수)로 넣는다.
- 1 pt = 12700 EMU. `pptx.util.Pt()`를 쓰고 직접 EMU 산술을 하지 않는다.
- 텍스트 박스는 `tf.word_wrap = True`, `tf.margin_left = tf.margin_right =
  tf.margin_top = tf.margin_bottom = 0`, `tf.auto_size = MSO_AUTO_SIZE.NONE`.
  **좌표 표의 y는 텍스트 박스 상단이며, 내부 여백 0을 전제로 계산된 값이다.**
- 세로 정렬은 `tf.vertical_anchor = MSO_ANCHOR.TOP`. 중앙 정렬을 쓰면 §7의 정렬 엔진이
  무의미해진다.

---

## §2. 좌표 체계

```
캔버스        960 × 540
마진          좌우 72 / 상 56 / 하 56
콘텐츠 박스    x 72 – 888 (816),  y 56 – 484 (428)
컬럼          12개, 컬럼폭 46, 거터 24  → step 70
```

수식 — 코드에 이 두 함수를 두고 **모든 x/w를 이걸로만 계산한다**. 좌표 표의 pt 실값은
검증용이다(두 값이 다르면 버그).

```python
GRID_X0, GRID_STEP, COL_W, GUTTER = 72, 70, 46, 24

def col_x(n: int) -> float:
    """1-based 컬럼 인덱스 → x(pt).  col1=72, col6=422, col12=842"""
    assert 1 <= n <= 12
    return GRID_X0 + (n - 1) * GRID_STEP

def span_w(s: int) -> float:
    """스팬 → 폭(pt).  span1=46, span4=256, span5=326, span6=396,
                      span7=466, span9=606, span10=676, span12=816"""
    assert 1 <= s <= 12
    return GRID_STEP * s - GUTTER
```

검산표 (린터가 이 값으로 검증하게 하라):

| col | x | | span | w |
| --- | --- | --- | --- | --- |
| 1 | 72 | | 1 | 46 |
| 2 | 142 | | 2 | 116 |
| 3 | 212 | | 3 | 186 |
| 4 | 282 | | 4 | 256 |
| 5 | 352 | | 5 | 326 |
| 6 | 422 | | 6 | 396 |
| 7 | 492 | | 7 | 466 |
| 8 | 562 | | 8 | 536 |
| 9 | 632 | | 9 | 606 |
| 10 | 702 | | 10 | 676 |
| 11 | 772 | | 11 | 746 |
| 12 | 842 | | 12 | 816 |

`col_x(n) + span_w(s)` 가 888을 넘으면 그리드 밖이다 → 린트 에러.

---

## §3. 일곱 가지 문제 → 변경 대응표

작업 중 판단이 필요할 때 이 의도를 기준으로 결정하라.

| # | 관찰된 문제 | 해결 방식 | 관련 절 |
| --- | --- | --- | --- |
| 1 | 내용 적은 장의 하단 공백 과다 (기존 보정식 `남은높이/2.4, max 120`) | 보정식 폐기. 상단(y=56) / 하단(bottom=484) **이분법**. 경계 `블록높이/428 ≥ 0.58`. eyebrow는 항상 y 56 고정 → 공백이 눈썹과 본문 **사이**로 이동 | §7 |
| 2 | cards 3열 그리드 반복으로 단조로움 | 항목 수가 구성을 결정. 2→pair, 3→ledger, 4→quad, 5–6→dense. + 같은 모드 연속 금지 | §9 |
| 3 | 장식선 제거 후 시각적 앵커 상실 | 세 장치로 대체 — ① 대형 활자 덩어리(mega 200 / stat1 160 / stat_sub2 46 색인) ② 채움 판(짧은 변 ≥32pt **이고** 활자 포함일 때만) ③ 반전 필드(section·closing 전면) | §8.3, §8.11, §12 |
| 4 | 눈썹 라벨 12pt가 위계에서 사라짐 | 키우지 않고 **스타일을 분리**. `eyebrow 20/1.20/+120` 신설(SemiBold, accent). label 12는 쪽번호·각주 전용으로 축소 | §6 |
| 5 | cover / closing이 좌우 반전만으로는 여전히 유사 | 네 축을 동시에 다르게 — 명도(ground/figure 반전), 제목 줄수(2/1), 측정(span9/span7), 보조 요소(우상단 stat2 연도 / col8 정보 원장 3행) | §8.1, §8.2 |
| 6 | 어두운 팔레트의 muted 대비 부족 | 명/암 혼합비 분리. **이유가 다르다** — 밝은 쪽은 ground 대비 4.5:1이 상한을 정해 0.34가 최대, 어두운 쪽은 ground 대비가 자동 충족이라 figure 분리가 기준이 되어 0.30 | §5 |
| 7 | table 3열 이상에서 열 정렬이 약하게 읽힘 | 세로 괘선을 늘리지 않고 **열을 그리드에 스냅**. 수치열 span2 우정렬(자리수 정렬), 머리행을 816×34 figure 채움 판으로, 행 사이만 0.75pt | §11 |

---

## §4. `references/deck-spec.yaml` — 변경 전체

### 4.1 유지 (한 글자도 바꾸지 말 것)

```yaml
canvas:
  width_pt: 960
  height_pt: 540

grid:
  margin_x: 72
  margin_top: 56
  margin_bottom: 56
  columns: 12
  col_w: 46
  gutter: 24

palettes:            # hex 12개 전부 원본 그대로
  letterpress: {ground: "F4F1EA", figure: "1B1B1F", accent: "C8452D"}
  blueprint:   {ground: "122033", figure: "EDE6D6", accent: "D98E32"}
  graphite:    {ground: "FAFAF8", figure: "22252A", accent: "6E7F32"}
  dive:        {ground: "0E1A2B", figure: "FFFFFF", accent: "1B7FD4"}
active_palette: letterpress

fonts:
  display: "Pretendard Black"
  head:    "Pretendard SemiBold"
  body:    "Pretendard"
  label:   "Pretendard Medium"

density: speaker
density_limits:
  speaker: {bullets: 3, bullet_chars: 34, table_rows: 5}
  reading: {bullets: 6, bullet_chars: 80, table_rows: 10}

rhythm:
  max_same_layout_run: 2
  page_number_from: 2
```

### 4.2 삭제

```yaml
muted_mix: 0.45      # ← 삭제. 아래 colors 그룹의 두 값으로 대체
```

### 4.3 신규·수정 — 그대로 붙여넣을 YAML

```yaml
# --- 색 파생 규칙 ------------------------------------------------------------
# muted는 figure를 ground 쪽으로 섞어 만든다. 혼합비를 명/암에서 다르게 가져간다.
# 밝은 ground: ground 대비 4.5:1 이 상한을 결정 → 0.34가 최대치
# 어두운 ground: ground 대비는 자동 충족, figure와의 분리가 기준 → 0.30
# 판정 규칙: muted는 ground 대비 ≥ 4.5:1  AND  figure 대비 ≥ 1.7:1 을 동시에 만족
colors:
  muted_mix_light: 0.34     # figure가 ground보다 어두운 팔레트 (letterpress, graphite)
  muted_mix_dark:  0.30     # figure가 ground보다 밝은 팔레트 (blueprint, dive)
  hairline_mix:    0.70     # 0.75pt 구조선 전용. 활자에는 절대 쓰지 않는다
  muted_min_ratio_ground: 4.5
  muted_min_ratio_figure: 1.7

# --- 타입 스케일 -------------------------------------------------------------
# tracking 단위 = 1/100 pt (PPTX rPr/@spc). 이 표에 없는 크기는 존재하지 않는다.
styles:
  mega:      {font: display, size: 200, leading: 0.82, tracking: -500}
  stat1:     {font: display, size: 160, leading: 0.85, tracking: -400}
  stat2:     {font: display, size: 132, leading: 0.88, tracking: -350}
  stat3:     {font: display, size: 96,  leading: 0.92, tracking: -250}
  stat_sub:  {font: display, size: 62,  leading: 0.95, tracking: -160}
  stat_sub2: {font: display, size: 46,  leading: 1.00, tracking: -120}
  display:   {font: display, size: 84,  leading: 0.98, tracking: -200}
  quote:     {font: head,    size: 44,  leading: 1.28, tracking: -90}
  h1:        {font: head,    size: 40,  leading: 1.08, tracking: -80}
  h2:        {font: head,    size: 24,  leading: 1.30, tracking: -40}
  eyebrow:   {font: head,    size: 20,  leading: 1.20, tracking: 120}   # 신규
  body:      {font: body,    size: 18,  leading: 1.55, tracking: 0}
  small:     {font: body,    size: 15,  leading: 1.45, tracking: 0}
  label:     {font: label,   size: 12,  leading: 1.20, tracking: 70}

# 스타일 용도 고정 — 아무 데나 쓰지 않는다
style_usage:
  mega:      [section.number]
  stat1:     [data.hero]
  stat2:     [cover.year]
  stat3:     []                       # 현 설계에서 미사용. 스케일에는 남긴다
  stat_sub:  [data.hero_unit]
  stat_sub2: [data.secondary_value, cards.ledger_index]
  display:   [cover.title, section.title, statement.text, closing.title]
  quote:     [quote.text]
  h1:        [two_col.title, cards.head, table.head, image_split.title, image_full.title]
  h2:        [two_col.subhead, cards.item_title, data.unit_small]
  eyebrow:   [*.eyebrow, cards.item_index, quote.source]
  body:      [*.body, table.cell]
  small:     [*.caption, *.runner, table.header_label, table.footnote]
  label:     [page_number, footnote_micro]    # 눈썹 라벨에서 퇴출

# --- 정렬 엔진 ---------------------------------------------------------------
anchors:
  anchor_top:     56      # 상단 정렬 시 블록 상단 y
  anchor_bottom:  484     # 하단 정렬 시 블록 하단 y
  content_h:      428     # 484 - 56
  fill_threshold: 0.58    # 블록높이 / 428 >= 0.58 → 상단, 미만 → 하단 (경계 248.24pt)
  eyebrow_y:      56      # 본문 정렬과 무관하게 항상 고정
  gap_eyebrow:    30      # eyebrow 박스 하단 → 다음 요소 상단
  gap_block:      32      # 블록 간 최소 간격

# --- 채움 판 / 선 -----------------------------------------------------------
plates:
  plate_min_side:  32     # 채운 사각형의 짧은 변 하한
  plate_needs_text: true  # 내부에 활자를 담지 않는 채움 사각형은 생성 금지
rules:                    # 허용된 선. 이 목록에 없는 선은 생성하지 않는다
  hero_rule_w:  2         # data 히어로 숫자 밑줄. figure 단색
  table_rule_w: 0.75      # table 행 구분선. hairline 색

# --- cards ------------------------------------------------------------------
cards:
  mode_by_count: {2: pair, 3: ledger, 4: quad, 5: dense, 6: dense}
  head:          {eyebrow_y: 56, h1_y: 86, h1_h: 43}   # 4모드 공통
  card_stagger:  84       # pair 모드의 수직 엇단
  row_h:         94       # ledger
  row_step:      109      # ledger
  cell_h:        146      # quad · dense
  cell_step:     161      # quad · dense

# --- table ------------------------------------------------------------------
table:
  header_h: 34
  row_h:    40
  pad_x:    16
  header_y: 168
  first_row_y: 208
  col_pattern:            # 열 수 → 컬럼 스팬 배열. 합 + 거터가 12컬럼에 맞는다
    2: [6, 6]
    3: [6, 3, 3]
    4: [6, 2, 2, 2]
    5: [4, 2, 2, 2, 2]

hard_rules:
  - no_shadow
  - no_rounded_corner
  - no_gradient
  - no_3d
  - no_theme_table_style
  - no_emoji
  - no_decor_rule          # 신규: 허용 목록(hero_rule, table_rule) 외의 선 금지
  - no_full_width_band     # 신규: 활자를 담지 않는 채움 사각형 금지

rhythm:
  max_same_layout_run: 2
  max_same_cards_mode_run: 1   # 신규: cards 연속 시 같은 모드 반복 금지
  page_number_from: 2
```

### 4.4 요약 — 스케일은 13종 → 14종

`eyebrow` 하나만 추가했고 기존 13종은 size/leading/tracking 모두 그대로다. 결과는 여전히
유한한 목록이다. **이 14종에 없는 크기를 코드가 만들어내면 린트 에러다.**

---

## §5. 색 계산 — `scripts/deckkit.py`

### 5.1 구현

```python
from typing import Tuple

def _hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

def _rgb_to_hex(r: float, g: float, b: float) -> str:
    return "{:02X}{:02X}{:02X}".format(
        *(max(0, min(255, int(round(v)))) for v in (r, g, b)))

def _rel_luminance(h: str) -> float:
    """WCAG 상대휘도."""
    out = []
    for c in _hex_to_rgb(h):
        s = c / 255.0
        out.append(s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4)
    r, g, b = out
    return 0.2126 * r + 0.7152 * g + 0.0722 * b

def contrast(h1: str, h2: str) -> float:
    l1, l2 = _rel_luminance(h1), _rel_luminance(h2)
    lo, hi = min(l1, l2), max(l1, l2)
    return (hi + 0.05) / (lo + 0.05)

def mix(figure: str, ground: str, t: float) -> str:
    """figure를 ground 쪽으로 t 비율 섞는다. sRGB 값에서 선형 보간(기존 동작 유지)."""
    f, g = _hex_to_rgb(figure), _hex_to_rgb(ground)
    return _rgb_to_hex(*(f[i] * (1 - t) + g[i] * t for i in range(3)))

def is_dark_ground(pal: dict) -> bool:
    """figure가 ground보다 밝으면 어두운 팔레트."""
    return _rel_luminance(pal["figure"]) > _rel_luminance(pal["ground"])

def derive(pal: dict, cfg: dict) -> dict:
    """팔레트 3색 → 사용 색 전체. 4번째 hex는 생기지 않는다(전부 파생값)."""
    c = cfg["colors"]
    t_muted = c["muted_mix_dark"] if is_dark_ground(pal) else c["muted_mix_light"]
    muted    = mix(pal["figure"], pal["ground"], t_muted)
    hairline = mix(pal["figure"], pal["ground"], c["hairline_mix"])

    # 자기 검증 — 값을 바꿨을 때 조용히 깨지는 걸 막는다
    r_ground = contrast(muted, pal["ground"])
    r_figure = contrast(muted, pal["figure"])
    assert r_ground >= c["muted_min_ratio_ground"], f"muted vs ground {r_ground:.2f}"
    assert r_figure >= c["muted_min_ratio_figure"], f"muted vs figure {r_figure:.2f}"

    return {
        "ground": pal["ground"], "figure": pal["figure"], "accent": pal["accent"],
        "muted": muted, "hairline": hairline,
    }
```

### 5.2 기대값 — 이 숫자가 안 나오면 구현이 틀렸다

| 팔레트 | 방향 | 혼합비 | muted | ground 대비 | figure 대비 | hairline |
| --- | --- | --- | --- | --- | --- | --- |
| letterpress | light | 0.34 | `656464` | 5.24 : 1 | 2.91 : 1 | `B3B1AD` |
| graphite | light | 0.34 | `6B6D70` | 4.96 : 1 | 2.96 : 1 | `B9BABA` |
| blueprint | dark | 0.30 | `ABABA5` | 7.12 : 1 | 1.85 : 1 | `545B64` |
| dive | dark | 0.30 | `B7BABF` | 8.99 : 1 | 1.95 : 1 | `565F6B` |

참고 — 기존 단일 0.45의 실측: letterpress `7D7B7A` 3.74:1(미달), graphite `82858A`
3.39:1(미달), blueprint `8A8D8D` 4.90:1(여유 0), dive `9298A0` 6.02:1.
반올림 차로 ±1 정도는 허용하되, 대비 비율은 위 값과 소수 둘째 자리까지 맞아야 한다.

### 5.3 반전 필드(invert)의 색

`section`과 `closing`은 ground/figure를 맞바꾼 필드다. 이때 **muted 혼합비도 방향이 바뀌므로
다시 계산한다.**

```python
def invert(pal: dict) -> dict:
    return {"ground": pal["figure"], "figure": pal["ground"], "accent": pal["accent"]}
# 사용: colors = derive(invert(pal), cfg)   ← derive를 다시 통과시킨다.
```

예: letterpress 반전(bg `1B1B1F`, 활자 `F4F1EA`)은 어두운 ground이므로 0.30이 적용되어
muted = `B3B1AD`. accent는 반전하지 않는다(`C8452D` 그대로).

### 5.4 색 사용 규칙

- `accent`는 장당 **한 역할**에만. eyebrow, section 번호, 색인 숫자, statement 강조행,
  quote 출처 중 하나. 한 장에서 두 가지 이상에 쓰지 않는다.
- `muted`는 캡션·러너·각주 전용. 제목·본문·수치에 쓰지 않는다.
  단 **data 히어로 캡션은 muted가 아니라 figure**다(좌표 표 비고 참조).
- `hairline`은 0.75pt 구조선 전용. 활자에 쓰면 린트 에러.
- 투명도는 어디에도 없다. muted는 미리 계산된 불투명 hex다.

---

## §6. 줄수 계산 — `scripts/deckkit.py`

레이아웃이 좌표를 계산하기 전에 **줄수를 먼저 알아야** 블록 높이가 나오고, 블록 높이가
나와야 §7 정렬 엔진이 동작한다.

```python
HANGUL_EM = 0.8643        # Pretendard 한글 실측 글리프 폭
LATIN_EM  = 0.55          # 숫자·라틴 근사치

def advance(ch: str, size: float, tracking: int) -> float:
    """글자 하나의 진행폭(pt). tracking 단위는 1/100 pt."""
    em = HANGUL_EM if ord(ch) >= 0x1100 else LATIN_EM
    return size * em + tracking / 100.0

def line_count(text: str, w: float, style: dict) -> int:
    """폭 w(pt) 안에서 필요한 줄 수. 명시적 줄바꿈(\\n)을 우선 존중한다."""
    total = 0
    for para in text.split("\n"):
        used, lines = 0.0, 1
        for ch in para:
            a = advance(ch, style["size"], style["tracking"])
            if used + a > w:
                lines += 1
                used = a
            else:
                used += a
        total += lines
    return total

def block_h(text: str, w: float, style: dict) -> float:
    """텍스트 박스 높이(pt) = size * leading * 줄수."""
    return style["size"] * style["leading"] * line_count(text, w, style)
```

`references/font-metrics.json`이 이미 있으므로, 글자별 실측 폭이 그 안에 있으면
`advance()`가 그것을 우선 참조하게 하고 위 상수는 폴백으로 둔다.

검산 (좌표 표의 h가 이렇게 나온다):

```
display 84, leading 0.98 → 1줄 82.3 / 2줄 164.6 (표기 165) / 3줄 246.9 (표기 247)
h1      40, leading 1.08 → 1줄 43.2 (43) / 2줄 86.4 (86)  / 3줄 129.6 (130)
h2      24, leading 1.30 → 1줄 31.2 (31) / 2줄 62.4 (62)
body    18, leading 1.55 → 1줄 27.9      / 3줄 83.7 (84)  / 4줄 111.6 (112)
                           7줄 195.3 (195) / 8줄 223.2 (223)
small   15, leading 1.45 → 1줄 21.75 (22) / 2줄 43.5 (44)  / 4줄 87
quote   44, leading 1.28 → 3줄 168.96 (169)
mega   200, leading 0.82 → 1줄 164
stat1  160, leading 0.85 → 1줄 136
stat2  132, leading 0.88 → 1줄 116.16 (116)
stat_sub2 46, leading 1.00 → 1줄 46
eyebrow 20, leading 1.20 → 1줄 24
```

한 줄에 들어가는 한글 글자 수 (검증용):

```
display 84 (tracking -200) → 글자 진행폭 70.6pt
    span9  606pt → 8자 / span10 676pt → 9자
h1 40 (-80) → 33.8pt
    span4 256pt → 7자 / span5 326pt → 9자
h2 24 (-40) → 20.3pt
    span4 256pt → 12자 / span5 326pt → 16자
quote 44 (-90) → 37.1pt
    span9 606pt → 16자
body 18 (0) → 15.56pt
    span5 326pt → 20자 / span6 396pt → 25자 / span7 466pt → 29자
small 15 (0) → 12.96pt
    span4 256pt → 19자 / span5 326pt → 25자
```

**분량이 넘치면 자간을 조이지 말고 분량을 줄인다.** 린터가 `density_limits`로 잡는다.

---

## §7. 정렬 엔진 — 문제 1의 해결

### 7.1 폐기할 기존 코드

"남은 높이 / 2.4, 최대 120pt"로 블록을 내려보내는 보정식을 **완전히 제거**한다. 검색어:
`2.4`, `min(120`, `remaining`, `offset`, `nudge`.

### 7.2 새 규칙

```python
ANCHOR_TOP, ANCHOR_BOTTOM, CONTENT_H, FILL_THRESHOLD = 56.0, 484.0, 428.0, 0.58

def resolve_y(block_height: float) -> float:
    """블록 상단 y를 반환. 중간값은 없다."""
    if block_height / CONTENT_H >= FILL_THRESHOLD:      # 경계 248.24pt
        return ANCHOR_TOP
    return ANCHOR_BOTTOM - block_height
```

- **`block_height`는 본문 블록의 높이다. eyebrow는 포함하지 않는다.**
  eyebrow는 정렬과 무관하게 항상 `y = 56`에 고정된다. 이게 문제 1 해결의 핵심이다.
  빈 공간이 화면 아래가 아니라 **눈썹과 본문 사이**로 옮겨가면서 프레임으로 읽힌다.
- 블록이 여러 요소로 구성되면 `block_height`는 첫 요소 상단부터 마지막 요소 하단까지의
  총 높이다. 하단 정렬 시 각 요소는 이 블록 원점을 기준으로 상대 배치한다.
- 두 컬럼이 있는 레이아웃(`two_col`)에서 하단 정렬이면 **양쪽 컬럼 모두 하단 484에 붙인다**
  (컬럼별로 따로 판정하지 않는다).
- `quote`는 예외적으로 **분량과 무관하게 항상 하단 정렬**이다.
- `image_split`에서 하단 정렬이 발동하면 **텍스트 블록만** 내려가고 이미지는 그대로다.

### 7.3 상태별 실제 적용 결과 (아트보드 대조)

| 보드 | 레이아웃 | 블록높이 | 판정 | 결과 y |
| --- | --- | --- | --- | --- |
| 04 | statement 적음 | 165 | 165/428 = 0.385 < 0.58 | 하단 → 319 |
| 05 | statement 가득 | 364 | 0.850 ≥ 0.58 | 상단 → 120 |
| 07 | two_col 적음 | 112 | 0.262 < 0.58 | 하단 → 좌 398 / 우 372 |
| 06 | two_col 가득 | 376 | 0.878 ≥ 0.58 | 상단 → 108 |
| 13 | data 적음 | 224 | 0.523 < 0.58 | 하단 → 260 |
| 12 | data 가득 | 305 | 0.713 ≥ 0.58 | 상단 → 140 |

보드 05의 상단 y가 56이 아니라 120인 이유: eyebrow(y 56, h 24) + `gap_eyebrow` 30 이후가
본문 시작이므로 `56 + 24 + 30 = 110`이 최소이고, 보조 body와의 `gap_block` 32를 맞춰
하단 484에 정확히 닿도록 120으로 내려 잡았다. **eyebrow가 있는 레이아웃의 상단 정렬 y는
`max(110, 484 - block_height)`가 아니라 좌표 표의 실값을 쓴다** — §8에 레이아웃별로
전부 적어뒀다.

---

## §8. 레이아웃 좌표 — 전체

표기 규칙: `x = col N span S (pt실값)`, `y`, `w`, `h`. h는 §6으로 계산된 블록 높이이며
검산용이다. 모든 텍스트는 좌정렬(명시된 우정렬 제외), 세로 정렬 TOP, 내부 여백 0.

### 8.1 `cover`

필드: ground. 정렬: 하단.

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| eyebrow | eyebrow | accent | col1 span9 (72) | 56 | 606 | 24 |
| 연도 숫자 | stat2 | accent | col9 span4 (632) | 56 | 256 | 116 |
| 제목 2줄 | display | figure | col1 span9 (72) | 252 | 606 | 165 |
| meta 2줄 | small | muted | col1 span5 (72) | 441 | 326 | 44 |

- 연도 숫자는 **우측 정렬**(`PP_ALIGN.RIGHT`). 박스 오른변 = 888.
- 제목 + meta가 한 블록(252–484, 232pt)으로 하단 484에 붙는다.
- 더미: eyebrow `01 / 2026 연간 전략`, 숫자 `26`, 제목 `재고 예측을\n다시 설계한다`,
  meta `제품기획팀 · 2026년 3월 12일\n내부 검토용 초안 v3`.

### 8.2 `closing`

필드: **반전**(bg = figure, 활자 = ground). 정렬: 하단. → cover와 네 축이 다르다(§3 #5).

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| eyebrow | eyebrow | accent | col1 span5 (72) | 56 | 326 | 24 |
| 원장 라벨 ×3 | small | muted | col8 span5 (562) | 196 / 266 / 336 | 326 | 22 |
| 원장 값 ×3 | body(SemiBold) | figure | col8 span5 (562) | 220 / 290 / 360 | 326 | 28 |
| 제목 1줄 | display | figure | col1 span7 (72) | 402 | 466 | 82 |

- 원장 행 step 70. 행 내부: 라벨 y, 값 y+24.
- 원장 값은 body 크기에 **SemiBold 웨이트**(head 폰트). 스케일 외 크기가 아니다.
- 제목은 **1줄만** 허용(2줄이 되면 cover와 구별이 사라진다). 넘치면 문구를 줄인다.
- 더미: eyebrow `04 / 마무리`, 원장 `담당 / 제품기획팀 · 이든`, `회신 기한 / 3월 27일
  금요일까지`, `문서 / 내부 위키 · 재고예측 2026`, 제목 `다음 단계는`.

### 8.3 `section`

필드: **반전**. 정렬: 상하 대립(고정 좌표, 정렬 엔진 미적용).

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| 번호 | mega | accent | col1 span7 (72) | 56 | 466 | 164 |
| 러너 | small | muted | col9 span4 (632) | 60 | 256 | 22 |
| 제목 2줄 | display | figure | col1 span9 (72) | 319 | 606 | 165 |

- 러너는 **우측 정렬**.
- 상단 164pt 덩어리와 하단 165pt 덩어리 사이 99pt 공백이 앵커 역할을 한다(문제 3).
  선을 넣지 않는다.
- 번호는 두 자리(`01`–`99`)를 전제로 폭을 잡았다. 세 자리는 지원하지 않는다.
- 더미: 번호 `03`, 러너 `3개 장 · 7분`, 제목 `주문 주기를\n먼저 쪼갠다`.

### 8.4 `statement`

필드: ground. 정렬: **엔진**.

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| eyebrow | eyebrow | accent | col1 span7 (72) | 56 (고정) | 466 | 24 |
| 주기 small | small | muted | col9 span4 (632) | 60 | 256 | 22 |
| 본문 2줄 (적음) | display | figure | col1 span10 (72) | **319** | 676 | 165 |
| 본문 3줄 (가득) | display | figure | col1 span10 (72) | **120** | 676 | 247 |
| 보조 body 3줄 (가득만) | body | figure | col1 span7 (72) | 400 | 466 | 84 |

- 주기 small은 우측 정렬. 있어도 되고 없어도 된다.
- **강조는 선이 아니라 색**: 본문 중 한 줄(문장 단위) 전체를 accent로 칠한다. 단어 단위
  부분 강조는 하지 않는다(한글에서 산만해진다). 구현: 해당 문단을 별도 run으로 분리하고
  그 run의 color만 accent로.
- 더미(적음) `예측이 틀린 게 아니라\n입력이 늦다` — 2행이 accent.
  더미(가득) `재고는 창고의\n문제가 아니라\n주문서의 문제다` — 3행이 accent.

### 8.5 `two_col`

필드: ground. 비대칭 **4 : 7**. 좌우 실공백 94pt(col5를 통째로 비운다). 정렬: **엔진**.

가득 찰 때:

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| eyebrow | eyebrow | accent | col1 span4 (72) | 56 | 256 | 24 |
| 좌 제목 3줄 | h1 | figure | col1 span4 (72) | 108 | 256 | 130 |
| 좌 리드 4줄 | small | muted | col1 span4 (72) | 262 | 256 | 87 |
| 우 본문 8줄 | body | figure | col6 span7 (422) | 108 | 466 | 223 |
| 우 소제목 | h2 | figure | col6 span7 (422) | 363 | 466 | 31 |
| 우 본문2 3줄 | body | figure | col6 span7 (422) | 400 | 466 | 84 |

내용 적을 때:

| 요소 | x | y | w | h |
| --- | --- | --- | --- | --- |
| eyebrow | col1 span4 (72) | 56 (고정) | 256 | 24 |
| 좌 제목 2줄 | col1 span4 (72) | 398 | 256 | 86 |
| 우 본문 4줄 | col6 span7 (422) | 372 | 466 | 112 |

- 좌우 **모두** 하단 484에 붙는다. 위쪽 292pt 공백이 프레임이 된다.
- 좌 리드(small muted)는 분량이 적으면 생략한다.
- 기존 린터의 `BALANCE` 경고(리드 없이 불릿 2개 이하일 때 왼쪽이 빔)는 이 하단 정렬로
  해소된다 — 경고 조건을 재검토하라.

### 8.6 `cards` — 공통 머리

4모드 전부 동일.

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| eyebrow | eyebrow | accent | col1 span7 (72) | 56 | 466 | 24 |
| 제목 1줄 | h1 | figure | col1 span9 (72) | **86** | 606 | 43 |

`h1_y = 86`은 4모드와 `table` 레이아웃에서 공통이다. 예외 없음.

### 8.7 `cards / pair` — 항목 2개

불균등 5 : 6 + 수직 엇단 84pt. 컬럼 간 실공백 94pt.

| | x | 색인 y | 제목 y | 본문 y | w |
| --- | --- | --- | --- | --- | --- |
| 카드 가 | col1 span5 (72) | 177 | 209 | 287 | 326 |
| 카드 나 | col7 span6 (492) | 261 | 293 | 371 | 396 |

카드 내부 상대 step: 색인 `y`, 제목 `y+32`, 본문 `y+110`.
스타일: 색인 = eyebrow(accent), 제목 = h2 최대 2줄(figure), 본문 = body 최대 4줄(muted).

- **테두리·모서리줄·배경 채움 없음.** 분리는 94pt 실공백 + 84pt 엇단 + accent 색인이 한다.
- 색인은 `가`/`나` 또는 `01`/`02`. 통일만 되면 된다.

### 8.8 `cards / ledger` — 항목 3개

3등분 그리드를 쓰지 않는다. 전폭 3행 원장.

행 y: **172 / 281 / 390** (행 높이 94, step 109). 마지막 행 하단 484.

| 셀 | 스타일 | 색 | x | y | w |
| --- | --- | --- | --- | --- | --- |
| 색인 | stat_sub2 | accent | col1 (72) | 행 y | 70 |
| 제목 | h2 | figure | col2 span4 (142) | 행 y + 6 | 256 |
| 본문 3줄 | body | muted | col7 span6 (492) | 행 y + 4 | 396 |

- 색인 박스 폭은 70(컬럼폭 46 + 거터 24) — 한 자리 숫자를 46pt로 놓기 위한 값.
- **행 구분선 없음.** 46pt 색인 숫자가 행 시작점을 잡는다.
- 제목·본문의 +6 / +4는 큰 숫자와의 시각 정렬 보정이다. 그대로 쓴다.

### 8.9 `cards / quad` — 항목 4개

불균등 2×2. 열폭 326 / 396. 균등 4분할이 아니다.

셀 원점: `(72, 177)` `(492, 177)` `(72, 338)` `(492, 338)` — 셀 높이 146, step 161.

| 셀 내부 | 스타일 | 색 | y 오프셋 | 최대 줄수 |
| --- | --- | --- | --- | --- |
| 색인 | eyebrow | accent | +0 | 1 |
| 제목 | h2 | figure | +30 | 1 |
| 본문 | small | muted | +70 | 3 |

좌열 w 326, 우열 w 396.

### 8.10 `cards / dense` — 항목 5–6개

3×2. 열 x: **72 / 352 / 632** (col1 · col5 · col9, 각 span4, w 256). 행 y: 177 / 338.

| 셀 내부 | 스타일 | 색 | y 오프셋 | 최대 줄수 |
| --- | --- | --- | --- | --- |
| 색인 | eyebrow | accent | +0 | 1 |
| 제목 | h2 | figure | +30 | 2 |
| 본문 | small | muted | +100 | 2 (20자 이내) |

- **항목 5개**면 첫 행을 두 칸으로 바꾼다: `col1 span7 (72, w 466)` + `col9 span4
  (632, w 256)`. 둘째 행은 3칸 그대로. 균등 배치를 피하기 위한 예외다.
- dense는 이 설계에서 유일하게 균등 열폭을 쓰는 모드다. 고밀도 전용이므로
  `max_same_cards_mode_run: 1`로 연속 등장을 막는다.

### 8.11 `data`

필드: 자유(아트보드는 dive). 정렬: **엔진**.

가득 찰 때 (히어로 + 종속 3행):

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| eyebrow | eyebrow | accent | col1 span7 (72) | 56 | 466 | 24 |
| 히어로 값 | stat1 | figure | col1 span6 (72) | 140 | 396 | 136 |
| 히어로 단위 | stat_sub | figure | — 값과 같은 run 내 인라인 — | | | |
| 히어로 밑줄 | 사각형 2pt | figure | col1 span6 (72) | 292 | 396 | 2 |
| 히어로 캡션 3줄 | body | **figure** | col1 span5 (72) | 308 | 326 | 84 |
| 종속 값 ×3 | stat_sub2 | figure | col8 span5 (562) | 140 / 244 / 348 | 326 | 46 |
| 종속 단위 ×3 | h2 | figure | — 값과 같은 run 내 인라인 — | | | |
| 종속 라벨 ×3 | small | muted | col8 span5 (562) | 194 / 298 / 402 | 326 | 44 |

- 종속 행 step 104. 라벨은 값 y + 54.
- **히어로 캡션은 muted가 아니라 figure**다. 히어로의 무게를 캡션까지 끌고 내려간다.
  종속 라벨만 muted.
- 히어로 밑줄은 **허용된 구조선**이다. 두께 2pt, figure 단색, 폭은 히어로 박스 폭과 동일.
  장식선이 아니라 값과 캡션을 묶는 구조선이라 허용된다.
- 단위는 값과 같은 단락 안의 별도 run으로 넣는다(폰트 크기만 다르게). 별도 텍스트 박스로
  만들면 베이스라인이 안 맞는다.

내용 적을 때 (히어로만):

| 요소 | x | y | w | h |
| --- | --- | --- | --- | --- |
| eyebrow | col1 span7 (72) | 56 (고정) | 466 | 24 |
| 주기 small (우정렬) | col9 span4 (632) | 60 | 256 | 22 |
| 히어로 값 | col1 span9 (72) | 260 | 606 | 136 |
| 히어로 밑줄 | col1 span9 (72) | 412 | 606 | 2 |
| 히어로 캡션 2줄 | col1 span7 (72) | 428 | 466 | 56 |

- 종속 행이 없으면 히어로가 **span6 → span9**로 넓어진다. 밑줄 폭도 함께 606.
- 우상단 러너만 y 60에 남아 대각 긴장을 만든다.

### 8.12 `quote`

필드: ground. 정렬: **항상 하단**(엔진 미적용).

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| 인용 3줄 | quote | figure | col2 span9 (142) | 232 | 606 | 169 |
| 출처 | eyebrow | accent | col2 span5 (142) | 433 | 326 | 24 |
| 부기 | small | muted | col2 span5 (142) | 461 | 326 | 23 |

- **col1을 통째로 비운 70pt 들여쓰기가 앵커**다. 인용부호 도형·큰따옴표 장식·세로선 없음.
- 인용문에 eyebrow가 없다. 대신 출처를 eyebrow 스타일로 쓴다.
- 인용이 2줄이면 y 288, 4줄이면 y 176으로 옮겨 하단 401을 유지한다
  (`y = 401 - block_h`).

### 8.13 `table`

필드: 자유(아트보드는 graphite). 고정 좌표.

| 요소 | 스타일 | 색 | x | y | w | h |
| --- | --- | --- | --- | --- | --- | --- |
| eyebrow | eyebrow | accent | col1 span7 (72) | 56 | 466 | 24 |
| 제목 1줄 | h1 | figure | col1 span9 (72) | 86 | 606 | 43 |
| 머리행 채움 판 | 사각형 | **figure 채움** | col1 span12 (72) | 168 | 816 | 34 |
| 머리 라벨 | small | **ground** | 열별 (아래) | 174 | 열별 | 22 |
| 본문행 ×5 | body | figure | 열별 | 208 / 248 / 288 / 328 / 368 | 열별 | 28 |
| 행 구분선 ×4 | 사각형 0.75pt | hairline | col1 span12 (72) | 241.5 / 281.5 / 321.5 / 361.5 | 816 | 0.75 |
| 각주 | small | muted | col1 span7 (72) | 424 | 466 | 44 |

- 머리 라벨 y = 머리판 y + 6 = 174. 본문 셀 y = 행 y + 6 (행 높이 40 안에서 18pt 텍스트
  수직 중앙에 가깝게 오도록 계산된 값).
- 행 구분선 y = 행 y + 33.5. 마지막 행 뒤에는 선을 넣지 않는다.
- **네이티브 표(`shapes.add_table`)를 쓰지 않는다.** 사각형과 텍스트 박스로 직접 그린다
  (`no_theme_table_style`).
- 5행은 `density_limits.speaker.table_rows = 5`에 맞춘 값. reading 모드(10행)는
  행 높이를 줄이지 말고 `row_h` 40 유지 + 슬라이드 분할로 처리한다.

### 8.14 `image_split`

| 요소 | x | y | w | h |
| --- | --- | --- | --- | --- |
| 이미지 | col7 좌변 (492) | **0** | 468 | **540** |
| eyebrow | col1 span5 (72) | 56 | 326 | 24 |
| 제목 3줄 (h1) | col1 span5 (72) | 108 | 326 | 130 |
| 본문 7줄 (body) | col1 span5 (72) | 262 | 326 | 195 |

- 이미지는 **우·상·하 3변 재단**(y 0, h 540, 오른변 960). 액자·테두리·모서리 반경 없음.
- 텍스트와 이미지 사이 실공백 94pt.
- 원본 비율 유지 + 중앙 크롭. PIL로 미리 잘라 넣는다(기존 구현 유지).
- 내용이 적으면 텍스트 블록만 하단 484 정렬. 이미지는 그대로.
- 좌우를 뒤집는 변형(이미지 왼쪽)은 만들지 않는다 — 대칭 반복 금지.

### 8.15 `image_full`

| 요소 | x | y | w | h |
| --- | --- | --- | --- | --- |
| 이미지 | 0 | 0 | 960 | 540 |
| 채움 판 | 0 | 300 | 560 | 240 |
| eyebrow | 72 | 340 | 440 | 24 |
| 제목 2줄 (h1) | 72 | 372 | 440 | 87 |
| 캡션 1줄 (small) | 72 | 471 | 440 | 22 |

- 채움 판은 **ground 단색**, 좌·하 2변 재단. 투명도 없음. 텍스트는 판 위에 얹는다.
- 판 내부 패딩: 좌 72 / 상 40 / 하 47.
- 이 판이 문제 3의 앵커 장치다. `plate_min_side 32` 규칙(짧은 변 240 ≥ 32, 활자 포함)을
  만족한다.

---

## §9. cards 모드 분기 — 구현

```python
CARD_MODE = {2: "pair", 3: "ledger", 4: "quad", 5: "dense", 6: "dense"}

def cards_mode(items: list) -> str:
    n = len(items)
    if n not in CARD_MODE:
        raise ValueError(f"cards는 항목 2~6개만 지원한다 (받은 값: {n}). "
                         f"7개 이상은 슬라이드를 쪼개라.")
    return CARD_MODE[n]
```

- 항목 1개는 `cards`가 아니라 `statement` 또는 `data`를 쓴다.
- 7개 이상은 에러. 자동으로 줄이거나 4열로 늘리지 마라.
- `dense`에서 `n == 5`면 §8.10의 예외 배치를 적용한다.
- 모드는 항목 수만으로 결정된다. YAML에서 사용자가 모드를 직접 지정하는 옵션은 만들지 않는다.

---

## §10. (예약)

---

## §11. table 열 계산 — 문제 7의 해결

```python
COL_PATTERN = {2: [6, 6], 3: [6, 3, 3], 4: [6, 2, 2, 2], 5: [4, 2, 2, 2, 2]}
PAD_X = 16

def table_columns(n: int):
    """[(x, w, align, text_x)] 반환. 첫 열만 좌정렬, 나머지는 우정렬."""
    spans = COL_PATTERN[n]
    out, col = [], 1
    for i, s in enumerate(spans):
        x, w = col_x(col), span_w(s)
        if i == 0:
            out.append((x, w, "left",  x + PAD_X))
        else:
            out.append((x, w, "right", x + w - PAD_X))
        col += s
    return out
```

4열 검산: `[(72,396,left,88), (492,116,right,592), (632,116,right,732),
(772,116,right,872)]`. 마지막 열 오른변 888 = 콘텐츠 우변. 열 사이 실공백 24pt.

- **수치열은 우정렬**이므로 자리수가 세로로 맞는다. 이게 3열 이상에서 눈이 행을 따라가게
  하는 핵심이다.
- 머리 라벨도 해당 열의 정렬을 따른다(수치열 머리는 우정렬).
- **세로 괘선을 추가하지 마라.** 열 구분은 그리드 스냅 + 24pt 공백 + 우정렬이 한다.
- 웨이트로도 열을 구분한다: **수치 = SemiBold(head 폰트, body 크기), 텍스트 = Regular.**
- 6열 이상은 지원하지 않는다. 열을 줄이거나 슬라이드를 쪼갠다.

---

## §12. 채움 판 / 선 — 생성 규칙

렌더러가 사각형을 그릴 수 있는 경우는 **정확히 이 셋뿐**이다.

```python
def assert_plate_ok(w: float, h: float, has_text: bool, kind: str):
    if kind == "plate":
        assert min(w, h) >= 32, f"plate 짧은 변 {min(w,h)} < 32"
        assert has_text,        "활자를 담지 않는 채움 사각형은 생성 금지"
    elif kind == "hero_rule":
        assert h == 2
    elif kind == "table_rule":
        assert h == 0.75
    else:
        raise ValueError(f"허용되지 않은 사각형 종류: {kind}")
```

허용 목록:

| kind | 어디 | 치수 | 색 |
| --- | --- | --- | --- |
| plate | `table` 머리행 | 816 × 34 | figure |
| plate | `image_full` 텍스트 판 | 560 × 240 | ground |
| hero_rule | `data` 히어로 밑줄 | 396 또는 606 × 2 | figure |
| table_rule | `table` 행 구분선 | 816 × 0.75 | hairline |

이 규칙 때문에 "제목 밑 4pt 강조선"은 **정의상 생성할 수 없다**(짧은 변 4 < 32, 활자 없음).
"화면을 가로지르는 색 띠"도 활자를 담지 않으므로 불가능하다. table 머리행은 가로로 길지만
34pt ≥ 32이고 라벨 활자를 담으므로 허용된다 — 이 경계가 의도된 판정이다.

반전 필드(section·closing)는 사각형이 아니라 **슬라이드 배경색**으로 처리한다
(`slide.background.fill`). 960×540 사각형을 깔지 마라.

---

## §13. `scripts/lint.py` — 추가 검사

기존 검사를 유지하고 아래를 추가한다. 전부 **에러**(경고 아님).

| ID | 검사 | 기준 |
| --- | --- | --- |
| `GRID` | 모든 텍스트 박스의 x가 `col_x(n)` 집합에 있고, `x + w ≤ 888` | §2 |
| `SCALE` | 모든 run의 size가 `styles` 14종의 size 집합에 있음 | §4.3 |
| `TRACK` | 모든 run의 `spc`가 해당 스타일의 tracking과 일치 | §4.3 |
| `ANCHOR` | 각 블록의 상단이 56이거나 하단이 484 (그 외 y 금지) | §7 |
| `EYEBROW_Y` | eyebrow 박스의 y가 정확히 56 | §7 |
| `PLATE` | 모든 사각형이 §12 허용 목록 중 하나에 해당 | §12 |
| `CONTRAST` | muted vs ground ≥ 4.5, muted vs figure ≥ 1.7 | §5 |
| `ALIGN` | body·small 단락의 정렬이 LEFT (예외: 러너·표 수치열의 RIGHT). CENTER 0회 | §1 |
| `CARDS_MODE` | 연속한 cards 슬라이드의 모드가 서로 다름 | §4.3 |
| `LAYOUT_RUN` | 같은 레이아웃 3연속 금지 (기존) | 유지 |
| `NO_SHADOW` | 모든 shape의 `shadow.inherit == False` | §1 |
| `LINES` | 도형 `line`이 채워진 경우 0회 (외곽선 전부 제거) | §1 |

`SCALE`과 `PLATE`가 가장 중요하다. 이 둘이 통과하면 "AI 티"의 주요 경로가 막힌다.

---

## §14. 검증 절차

1. `python scripts/build.py references/outline.example.yaml` 로 pptx를 만든다.
2. `python scripts/lint.py out/deck.pptx` — §13 검사 전부 통과해야 한다.
3. `python scripts/svgpreview.py` 또는 `preview.py`로 장별 이미지를 뽑는다.
4. 뽑은 이미지를 `ppt-deck 재설계.dc.html`의 해당 아트보드와 **겹쳐 비교**한다.
   대조 매핑:

| 아트보드 | 레이아웃 · 상태 | 팔레트 |
| --- | --- | --- |
| 01 | cover | letterpress |
| 02 | closing | letterpress 반전 |
| 03 | section | blueprint 반전 |
| 04 | statement 적음 | graphite |
| 05 | statement 가득 | graphite |
| 06 | two_col 가득 | letterpress |
| 07 | two_col 적음 | letterpress |
| 08 | cards pair (2) | blueprint |
| 09 | cards ledger (3) | blueprint |
| 10 | cards quad (4) | graphite |
| 11 | cards dense (6) | graphite |
| 12 | data 가득 | dive |
| 13 | data 적음 | dive |
| 14 | quote | letterpress |
| 15 | table | graphite |
| 16 | image_split | blueprint |
| 17 | image_full | dive |

5. 좌표 오차 허용 범위: **0pt**. 반올림으로 1pt 차가 나는 항목(h 계산값)은 좌표 표의
   표기값을 쓴다.
6. §5.2의 muted/hairline hex와 대비 비율을 단위 테스트로 고정하라.

---

## §15. 커밋 전 최종 체크리스트

- [ ] `muted_mix: 0.45`가 파일에서 사라졌다
- [ ] 기존 하단 보정식(`/2.4`, `max 120`)이 코드에서 사라졌다
- [ ] hex 12개가 한 글자도 안 바뀌었다
- [ ] 타입 스케일이 14종이고, 기존 13종의 size/leading/tracking이 그대로다
- [ ] `eyebrow`가 모든 레이아웃에서 y 56에 있다
- [ ] 제목 아래에 선이 있는 레이아웃이 0개다
- [ ] cards 4모드가 전부 테두리·배경 채움 없이 렌더된다
- [ ] table이 네이티브 표가 아니라 사각형+텍스트로 그려진다
- [ ] table에 세로 괘선이 없다
- [ ] 가운데 정렬 단락이 0개다
- [ ] 그림자·둥근모서리·그라데이션·투명도·이탤릭이 0회다
- [ ] 이모지·아이콘이 0개다
- [ ] `image_split`/`image_full` 이미지가 캔버스 변까지 재단됐다
- [ ] 린트 13종 전부 통과
- [ ] 17장 아트보드 대조에서 좌표 불일치 0건

---

## §16. 판단이 필요할 때

명세서에 답이 없는 상황을 만나면, 다음 순서로 결정한다.

1. §1 불변 제약을 어기는 선택지는 제외한다.
2. §3의 의도(일곱 문제의 해결 방향)에 맞는 쪽을 고른다.
3. 그래도 갈리면 **더 적게 그리는 쪽**을 고른다. 요소를 추가해서 해결하려 하지 마라.
4. 구조적 변경이 필요하다고 판단되면 임의로 하지 말고 **어떤 선택지들이 있었고 왜 막혔는지
   적어 보고**하라.

특히 다음은 절대 자율 판단 대상이 아니다.

- 새 hex 색 추가
- 타입 스케일에 크기 추가
- 새 선·테두리 추가
- 마진·컬럼 수·거터 변경
- 가운데 정렬 도입
