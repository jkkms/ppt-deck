#!/usr/bin/env python3
"""설치된 Pretendard TTF에서 실제 글리프 폭을 뽑아 캐시한다.

지금까지 넘침 검사는 '한글 1.0em / 숫자 0.62em' 같은 어림값이었다. 어림값은 두 번 틀렸다
(11.3점 줄바꿈, 히어로 숫자 오탐). 폰트가 손에 있는데 추정할 이유가 없다.

사용:  metrics.py [--rebuild]      -> references/font-metrics.json
"""
from __future__ import annotations
import json, os, sys, glob

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "..", "references", "font-metrics.json")
FONT_DIRS = [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts", "/System/Library/Fonts"]

# deck-spec.yaml의 fonts 값 → 파일명 조각
WANT = {
    "Pretendard Black": "Pretendard-Black",
    "Pretendard SemiBold": "Pretendard-SemiBold",
    "Pretendard Medium": "Pretendard-Medium",
    "Pretendard": "Pretendard-Regular",
}


def _find(stem):
    for d in FONT_DIRS:
        for ext in (".ttf", ".otf"):
            p = os.path.join(d, stem + ext)
            if os.path.exists(p):
                return p
        hits = glob.glob(os.path.join(d, stem + "*.tt[fc]"))
        if hits:
            return sorted(hits)[0]
    return None


def build(verbose=True):
    from fontTools.ttLib import TTFont
    out = {}
    for name, stem in WANT.items():
        path = _find(stem)
        if not path:
            if verbose:
                print(f"  ! {name}: 파일 없음 ({stem})", file=sys.stderr)
            continue
        f = TTFont(path, fontNumber=0, lazy=True)
        upm = f["head"].unitsPerEm
        cmap = f.getBestCmap()
        hmtx = f["hmtx"]
        # 코드포인트 → em 단위 advance. 한글 11,172자를 전부 담으면 파일이 커지므로
        # 실제로 폭이 다른 글리프만 남기고, 한글은 대표값 하나로 접는다(모노스페이스 특성).
        adv, hangul = {}, {}
        for cp, gname in cmap.items():
            if cp > 0x2FFFF:
                continue
            w = hmtx[gname][0] / upm
            if 0xAC00 <= cp <= 0xD7A3:
                hangul[round(w, 4)] = hangul.get(round(w, 4), 0) + 1
                continue
            adv[str(cp)] = round(w, 4)
        out[name] = {
            "upm": upm,
            "hangul": max(hangul, key=hangul.get) if hangul else 1.0,
            "default": round(hmtx[cmap.get(0x20, ".notdef")][0] / upm, 4) if 0x20 in cmap else 0.3,
            "adv": adv,
            "source": os.path.basename(path),
        }
        if verbose:
            print(f"  ✓ {name:22} {os.path.basename(path):28} "
                  f"글리프 {len(adv)}자 + 한글 {out[name]['hangul']}em")
        f.close()
    os.makedirs(os.path.dirname(os.path.abspath(CACHE)), exist_ok=True)
    with open(CACHE, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(",", ":"))
    return out


_CACHE = None


def load():
    """캐시된 메트릭. 없으면 None을 돌려주고 호출부가 어림값으로 물러난다."""
    global _CACHE
    if _CACHE is None:
        try:
            with open(CACHE, encoding="utf-8") as fh:
                _CACHE = json.load(fh)
        except Exception:
            _CACHE = {}
    return _CACHE or None


if __name__ == "__main__":
    build()
    print(f"\n캐시: {os.path.normpath(CACHE)}  "
          f"({os.path.getsize(CACHE)/1024:.0f} KB)")
