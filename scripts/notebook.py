#!/usr/bin/env python3
"""칸 목록 하나로 Colab 노트북과 슬라이드용 코드·출력을 함께 만든다 (사용자 원칙 §4-5).

설명은 슬라이드로, 실행은 노트북으로 넘겨 보여 주는 수업에서 둘이 어긋나지 않게 한다.
원본은 칸 목록(cells.py) 하나 — 노트북도 슬라이드 코드 카드도 여기서 나온다.

  1) 노트북을 조립한다
  2) 이 맥에서 실제로 실행한다 (준비 칸만 로컬용으로 바꿔서)
  3) 칸마다 코드·출력·그림을 cells.json 과 img/colab_N.png 로 저장한다
  4) 슬라이드 빌드가 남긴 cellpages.json(칸 → 쪽)이 있으면 칸 제목에 「PPT n쪽」을 넣는다

순서: notebook.py → build.py(meta.cells) → notebook.py 다시 (쪽 번호가 들어간 노트북)

  notebook.py lesson/cells3.py --out lesson/nb3
  notebook.py lesson/cells3.py --out lesson/nb3 --no-run     # 실행 없이 조립만

칸 목록 파일(cells.py):
  TITLE = "3차시 실습 — 머신러닝"
  INTRO = "칸 번호는 슬라이드의 `Colab N번 칸` 표시와 같습니다."   # 선택
  SETUP_COLAB = "..."   # 선택 — 노트북 맨 위 준비 칸 (없으면 한글 그래프 기본값)
  SETUP_LOCAL = "..."   # 선택 — 이 맥에서 실행할 때 준비 칸 대신 쓸 것
  CELLS = [
    {"section": "① 머신러닝 기본"},
    {"no": 1, "title": "공부용과 시험용 나누기", "code": "...", "note": "교재와 다른 점 한 줄"},
  ]

실행은 이 스크립트를 돌린 파이썬의 기본 커널(python3)로 한다 — 스킬 가상환경에
nbformat·nbclient·ipykernel 과 수업에 쓰는 패키지를 깔아 둔다. 전역 커널 등록은 하지 않는다
(다른 세션의 임시 폴더를 가리키던 커널이 세션이 끝나며 깨진 일이 있었다, 2026-10-04).
"""
from __future__ import annotations
import argparse, base64, copy, json, os, runpy, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

SETUP_COLAB = """# 그래프에 한글이 나오도록 나눔고딕을 설치합니다 (수업 전에 한 번만 실행)
!apt-get -qq install -y fonts-nanum > /dev/null
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/nanum/NanumGothic.ttf')
plt.rc('font', family='NanumGothic')
plt.rc('axes', unicode_minus=False)
plt.rc('font', size=13)
%config InlineBackend.figure_format = 'retina'"""


def _setup_local() -> str:
    """이 맥에서 실행할 때의 준비 칸. Colab 은 나눔고딕, 여기서는 설치된 한글 글꼴."""
    import metrics
    font = metrics._find("NanumGothic") or metrics._find("Pretendard-Regular")
    add = (f"fm.fontManager.addfont({font!r})\n"
           f"plt.rc('font', family=fm.FontProperties(fname={font!r}).get_name())\n") if font else ""
    return ("import numpy as np\nimport matplotlib.pyplot as plt\n"
            "import matplotlib.font_manager as fm\n" + add +
            "plt.rc('axes', unicode_minus=False)\nplt.rc('font', size=13)\n"
            "%config InlineBackend.figure_format = 'retina'")


def build(src: str, out: str, run: bool = True, timeout: int = 600):
    try:
        import nbformat
        from nbformat.v4 import new_notebook, new_markdown_cell, new_code_cell
    except ImportError:
        raise SystemExit("nbformat·nbclient·ipykernel 이 없다 — "
                         f"{sys.executable} -m pip install nbformat nbclient ipykernel")
    spec = runpy.run_path(src)
    CELLS, TITLE = spec["CELLS"], spec["TITLE"]
    os.makedirs(out, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    pf = os.path.join(out, "cellpages.json")
    pages = json.load(open(pf, encoding="utf-8")) if os.path.exists(pf) else {}

    nb = new_notebook()
    nb.metadata["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
    nb.metadata["colab"] = {"provenance": []}
    intro = spec.get("INTRO", "칸 번호는 슬라이드 코드 카드의 `Colab N번 칸` 표시와 같습니다.")
    nb.cells.append(new_markdown_cell(f"# {TITLE}\n\n{intro}\n\n수업 전에 맨 위 **준비 칸**을 한 번 실행해 주세요."))
    nb.cells.append(new_markdown_cell("## 준비 칸"))
    nb.cells.append(new_code_cell(spec.get("SETUP_COLAB", SETUP_COLAB) + spec.get("COLAB_EXTRA", "")))
    SETUP_AT = len(nb.cells) - 1
    idx, seen = {}, set()
    for c in CELLS:
        if "section" in c:
            nb.cells.append(new_markdown_cell(f"---\n## {c['section']}"))
            continue
        n = c["no"]
        if n in seen:
            raise SystemExit(f"칸 번호 {n} 이 두 번 나온다")
        seen.add(n)
        pg = pages.get(str(n))
        head = f"### {n}번 칸 · {c['title']}" + (f"  ·  PPT {pg}쪽" if pg else "")
        if c.get("note"):
            head += f"\n\n{c['note']}"          # 교재 코드를 바꿨으면 이유 한 줄 (원칙 §4-5)
        nb.cells.append(new_markdown_cell(head))
        idx[n] = len(nb.cells)
        nb.cells.append(new_code_cell(c["code"].strip("\n")))

    result = {}
    if run:
        from nbclient import NotebookClient
        exe = copy.deepcopy(nb)
        exe.cells[SETUP_AT].source = spec.get("SETUP_LOCAL", _setup_local())
        wd = os.path.join(out, ".run")
        os.makedirs(wd, exist_ok=True)
        NotebookClient(exe, timeout=timeout, kernel_name="python3",
                       resources={"metadata": {"path": wd}}).execute()
        for a, b in zip(nb.cells, exe.cells):
            if a.cell_type == "code":
                a.outputs, a.execution_count = b.outputs, b.execution_count
        os.makedirs(os.path.join(out, "img"), exist_ok=True)
        for c in CELLS:
            if "no" not in c:
                continue
            n = c["no"]
            texts, img = [], None
            for o in exe.cells[idx[n]].outputs:
                if o.output_type == "error":
                    raise SystemExit(f"{n}번 칸 오류: {o.ename}: {o.evalue}")
                if o.output_type == "stream" and o.name == "stderr":
                    # 경고가 학생 화면 출력에 섞이면 안 된다 — 코드를 고친다 (원칙 §4-5)
                    raise SystemExit(f"{n}번 칸 경고 출력 — 코드를 고쳐라: {o.text[:200]}")
                if o.output_type == "stream" and o.name == "stdout":
                    texts.append(o.text)
                elif o.output_type in ("display_data", "execute_result"):
                    if "image/png" in o.data:
                        img = f"img/colab_{n}.png"
                        with open(os.path.join(out, img), "wb") as fh:
                            fh.write(base64.b64decode(o.data["image/png"]))
                    elif "text/plain" in o.data:
                        texts.append(o.data["text/plain"] + "\n")
            result[str(n)] = {"title": c["title"],
                              "code": c["code"].strip("\n").split("\n"),
                              "out": "".join(texts).rstrip("\n").split("\n") if texts else [],
                              "img": img}
        json.dump(result, open(os.path.join(out, "cells.json"), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    nbformat.write(nb, os.path.join(out, stem + ".ipynb"))
    n_img = sum(1 for v in result.values() if v["img"])
    print(f"✓ {stem}.ipynb — 칸 {len(idx)}개" + (f" 실행 · 그림 {n_img}개" if run else " (실행 안 함)") +
          f" · 쪽 번호 {'있음' if pages else '없음'}  → {out}")
    for k, v in result.items():
        print(f"  {k}번 칸 → {' | '.join(v['out'])[:70]}{'  [그림]' if v['img'] else ''}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cells", help="칸 목록 파일 (TITLE · CELLS)")
    ap.add_argument("--out", required=True, help="cells.json · 노트북 · img/ 가 놓일 폴더")
    ap.add_argument("--no-run", action="store_true")
    ap.add_argument("--timeout", type=int, default=600)
    a = ap.parse_args()
    build(a.cells, a.out, run=not a.no_run, timeout=a.timeout)
