# -*- coding: utf-8 -*-
"""Layout QA: bounds, footer collisions, and estimated text overflow per shape."""
import sys, os, math, unicodedata
from pptx import Presentation
from pptx.util import Emu
from PIL import ImageFont

FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
_cache = {}


def font(size, bold):
    key = (round(size * 4), bold)
    if key not in _cache:
        _cache[key] = ImageFont.truetype(FONT_BOLD if bold else FONT_PATH,
                                         max(4, int(round(size * 4))))
    return _cache[key]


def text_width_in(text, size, bold):
    f = font(size, bold)
    return f.getlength(text) / 4.0 / 72.0


def check(path, verbose=True):
    prs = Presentation(path)
    SW = prs.slide_width / 914400
    SH = prs.slide_height / 914400
    problems = []
    for i, slide in enumerate(prs.slides, start=1):
        for shp in slide.shapes:
            x, y = shp.left / 914400, shp.top / 914400
            w, h = shp.width / 914400, shp.height / 914400
            if x < -0.02 or y < -0.02 or x + w > SW + 0.02 or y + h > SH + 0.02:
                problems.append((i, "bounds", f"{shp.shape_type} {w:.2f}x{h:.2f} at ({x:.2f},{y:.2f})"))
            if not shp.has_text_frame:
                continue
            tf = shp.text_frame
            txt = tf.text.strip()
            if not txt:
                continue
            # skip lines without horizontal overflow risk (bars/chips are sized to fit)
            inner_w = w - 0.06
            need = 0.0
            for p in tf.paragraphs:
                ptext = "".join(r.text for r in p.runs)
                if not ptext:
                    need += 0.12
                    continue
                size = max([(r.font.size.pt if r.font.size else 12) for r in p.runs] or [12])
                bold = any(r.font.bold for r in p.runs)
                tw = sum(text_width_in(r.text, r.font.size.pt if r.font.size else 12,
                                       r.font.bold) for r in p.runs)
                lines = max(1, math.ceil(tw / max(0.35, inner_w)))
                ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
                sa = (p.space_after.pt / 72.0) if p.space_after else 0.0
                need += lines * size * ls / 72.0 + sa
            if need > h + 0.10:
                problems.append((i, "overflow",
                                 f"{need:.2f}\" need > {h:.2f}\" box | {txt[:46]!r}"))
            if y + need > 7.02 and 0.5 < y < 6.90:
                problems.append((i, "footer-collide",
                                 f"bottom {y + need:.2f} > 7.02 | {txt[:40]!r}"))
    if verbose:
        for p in problems:
            print(f"slide {p[0]:>2} [{p[1]}] {p[2]}")
        print(f"\n{len(problems)} issue(s) over {len(prs.slides._sldIdLst)} slides")
    return problems


if __name__ == "__main__":
    check(sys.argv[1] if len(sys.argv) > 1 else
          os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "تدريب-المراقبين-الجويين-المسيرات.pptx"))
