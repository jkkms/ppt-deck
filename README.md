# ppt-deck

AI가 찍어낸 티가 나지 않는 한글 `.pptx`를 만드는 [Claude Code](https://claude.com/claude-code) 스킬.

> A Claude Code skill that generates Korean PowerPoint decks which don't look AI-generated.
> Design tokens first, eight rhythmic layouts, a content linter for "AI tells", and a visual review loop.

HTML 슬라이드가 아니라 **PowerPoint에서 바로 열리는 네이티브 .pptx**를 만듭니다. 발표장 PC에 그대로 꽂아야 하는 사람을 위한 것입니다.

<br>

## 왜 만들었나

AI로 만든 PPT는 내용과 무관하게 알아보입니다. 원인은 대부분 디자인 감각이 아니라 **구조**입니다.

| 티 나는 지점 | ppt-deck의 처리 |
| --- | --- |
| 전 슬라이드가 "제목 + 불릿" 한 골격 | 레이아웃 8종을 리듬으로 배치. 같은 레이아웃 3연속은 린터가 막는다 |
| Office 기본 파랑 `#4472C4`, 맑은 고딕 | 팔레트당 hex 3개만. Pretendard 4단 웨이트 |
| 사방 균등 여백 + 중앙 정렬 대칭 | 비대칭 12컬럼 그리드. 좌측 정렬 고정 |
| 텍스트박스 눈대중 배치 | 좌표가 그리드 함수에서만 나온다. 하드코딩 불가 |
| 폰트 크기 종류 남발 | 타입 스케일 12단. 이 표에 없는 크기는 존재하지 않는다 |
| 그림자·둥근모서리·그라데이션·3D | 생성 경로 자체가 없다 |
| 같은 크기 숫자 3개를 나란히 (대시보드 위젯) | `data`는 히어로 1 + 보조 2의 비대칭 위계 |
| 완결 문장 불릿, 기계적 병렬, 이모지 | 린터가 `SENTENCE` / `UNIFORM` / `PARALLEL` / `EMOJI`로 잡는다 |

**가장 큰 원인은 디자인이 아니라 문장입니다.** 그래서 린터의 절반이 콘텐츠 검사입니다.

<br>

## 설치

```bash
git clone https://github.com/jkkms/ppt-deck.git ~/.claude/skills/ppt-deck
cd ~/.claude/skills/ppt-deck
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
```

Claude Code에서 `/ppt-deck`으로 호출합니다.

### 필요한 것

- **Pretendard** — [설치](https://github.com/orioncactus/pretendard). Black / SemiBold / Medium / Regular 웨이트를 씁니다
- **LibreOffice** (선택) — `brew install --cask libreoffice`. 검수 렌더를 헤드리스로 돌립니다. 없으면 렌더만 건너뛰고 생성은 정상 작동합니다

<br>

## 쓰는 법

```bash
PY=.venv/bin/python

$PY scripts/lint.py       outline.yaml                      # 내용·리듬 검사
$PY scripts/build.py      outline.yaml -o out/deck.pptx     # 빌드 + 기하 자기검증
$PY scripts/svgpreview.py out/deck.manifest.json --png --sheet   # 검수 (외부 앱 불필요)
$PY scripts/validate.py   out/deck.pptx                     # 파일 무결성
```

`svgpreview.py`는 설치된 Pretendard TTF를 직접 읽어 그리므로 **LibreOffice도 PowerPoint도 필요 없습니다.** `--sheet`는 전체 슬라이드를 한 장에 늘어놓은 컨택트시트를 만듭니다.

아웃라인은 YAML 하나입니다. 문법은 [`references/outline.example.yaml`](references/outline.example.yaml) 참조.

```yaml
meta:
  title: "재현 가능한 채점 파이프라인"
  palette: letterpress      # letterpress | blueprint | graphite
  density: speaker          # speaker(발표용) | reading(읽기용)

slides:
  - layout: cover
  - layout: section
    number: "01"
    title: "채점은 왜 매번 무너지는가"
  - layout: data
    title: "재채점 실측"
    items:                  # 순서가 곧 위계. 첫 항목이 히어로
      - {value: "11.3", unit: "점", caption: "조교 간 평균 점수 편차"}
      - {value: "62", unit: "시간", caption: "1회차 총 채점 시간"}
```

<br>

## 레이아웃 8종

| layout | 쓰는 자리 |
| --- | --- |
| `cover` | 표지 |
| `section` | 장 구분. accent 전면 반전 |
| `statement` | 주장 한 줄. 호흡을 끊는다 |
| `cards` | 번호 카드 2~6개 (썸네일 가능) |
| `image_split` | 반출혈 이미지 + 텍스트 |
| `image_full` | 전출혈 이미지 + 바탕색 판 |
| `two_col` | 본문 주력. 비대칭 4:7 |
| `data` | 숫자 1~3개. 첫 항목이 히어로 |
| `quote` | 인용·증언 |
| `table` | 비교·일정. 네이티브 표를 쓰지 않는다 |
| `closing` | 마지막. 표지를 좌우 반전 |

네이티브 PowerPoint 표를 쓰지 않는 이유: 테마 표스타일의 줄무늬 채움과 전체 테두리가 가장 알아보기 쉬운 템플릿 지문이기 때문입니다. 헤어라인으로 직접 그립니다.

<br>

## 팔레트

hex 값은 팔레트당 **3개뿐**입니다. 나머지 이름은 전부 그 3개의 의미론적 별칭입니다. 4번째 색이 필요해지면 색이 부족한 게 아니라 레이아웃이 틀린 것입니다.

| | 바탕 | 본문 | 강조 | 성격 |
| --- | --- | --- | --- | --- |
| `letterpress` | `#F4F1EA` | `#1B1B1F` | `#C8452D` | 편집디자인. 인쇄·배포물 |
| `blueprint` | `#122033` | `#EDE6D6` | `#D98E32` | 어두운 강당, 프로젝터 |
| `graphite` | `#FAFAF8` | `#22252A` | `#6E7F32` | 기술 문서, 고밀도 읽기 |
| `dive` | `#0E1A2B` | `#FFFFFF` | `#1B7FD4` | 강의용. `invert`로 명·암 교차 |

<br>

## 이미지와 폰트 임베딩

```yaml
- layout: image_split          # 반출혈 이미지 + 텍스트
  image: "photos/vision.png"
  side: right                  # right(기본) | left
  fit: cover                   # cover(기본, 잘라 채움) | contain(통째로)
  focus: top                   # 세로 크롭 기준선. 인물 사진은 top
  title: "보고 구별합니다"
  bullets: ["사람과 물체를 알아봅니다"]
```

빌더가 PIL로 미리 잘라 넣어 비율이 찌그러지지 않고, 잘린 원본도 파일에 남지 않습니다. 쪽번호가 이미지에 깔리면 자동으로 피합니다.

```bash
$PY scripts/build.py outline.yaml -o out/deck.pptx --embed-fonts
```

Pretendard를 파일에 박아 넣습니다(굵기당 1~2MB). **PowerPoint는 TTF만 임베드합니다** — Pretendard 공식 릴리스의 `public/static/`은 OTF(CFF)라 거부당하니 `public/static/alternative/`의 TrueType 빌드를 설치하세요.

## 구조

```
SKILL.md                      Claude Code 오케스트레이션 (5단계)
references/deck-spec.yaml     디자인 토큰 단일 원천 — 색·폰트·스케일·그리드·밀도
references/outline.example.yaml
scripts/deckkit.py            엔진 — 그리드, 한글 ea 타이프페이스, 자간, 행잉인덴트
scripts/build.py              레이아웃 8종 + 빌드 후 기하 자기검증
scripts/lint.py               콘텐츠 AI 티 린터
scripts/preview.py            팔레트 비교 렌더
scripts/svgpreview.py         매니페스트 → SVG/PNG (외부 앱 불필요)
scripts/validate.py           파일 무결성
scripts/metrics.py            Pretendard 실측 글리프 폭
scripts/render.py             LibreOffice 실제 렌더 (선택)
```

### 설계 계약

1. 색·크기·좌표는 전부 `deck-spec.yaml`에서 파생된다. 예쁘게 하려고 코드에 숫자를 박지 않는다
2. 넘치면 **글자를 줄이지 않고 내용을 쪼갠다.** 타입스케일 이탈이 AI 티의 최대 원인이다
3. 렌더해서 눈으로 보기 전에는 완료가 아니다

<br>

## 알려진 한계

- 미리보기는 빌더와 같은 메트릭으로 줄을 나눕니다. **빌더가 놓친 넘침은 미리보기에서도 안 보입니다.** 커닝·합자·PowerPoint 고유의 줄바꿈 규칙은 재현하지 않습니다
- **차트가 없습니다.** 숫자는 `data` 레이아웃의 큰 활자로 보여주거나, 차트를 이미지로 만들어 넣어야 합니다
- macOS에서 개발·검증했습니다

<br>

## 계보

- [zarazhangrui/frontend-slides](https://github.com/zarazhangrui/frontend-slides) — 디자인 토큰 선확정, show-don't-tell, 밀도 모드, 안티슬롭 규칙. HTML이 아닌 네이티브 .pptx로 옮겼습니다
- **Anthropic `pptx` 스킬** — 필수 QA 단계(내용→파일→시각)와 검증기 개념. "제목 밑 강조선·모서리 액센트 줄 금지" 지침을 받아들여 장식용 줄을 전부 걷어냈습니다
- [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) — SVG를 중간 표현으로 두는 미리보기 방식. 외부 렌더러 없이 검수하는 경로가 여기서 나왔습니다

[Pretendard](https://github.com/orioncactus/pretendard) by orioncactus (SIL OFL 1.1).

<br>

## 라이선스

MIT
