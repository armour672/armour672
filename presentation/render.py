# -*- coding: utf-8 -*-
"""Raster preview renderer for the generated pptx (approximate, for visual QA)."""
import os, sys, math
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PPTX = os.path.join(BASE, "تدريب-المراقبين-الجويين-المسيرات.pptx")
OUT = os.path.join(BASE, "presentation", "preview")
SCALE = 120  # px per inch
REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
_fc = {}


def fnt(size_pt, bold):
    k = (round(size_pt * 2), bold)
    if k not in _fc:
        _fc[k] = ImageFont.truetype(BOLD if bold else REG, max(6, int(round(size_pt * SCALE / 72))))
    return _fc[k]


def shape_fill(shp):
    try:
        if shp.fill.type is None:
            return None, None
    except Exception:
        return None, None
    try:
        if shp.fill.type == 1:  # solid
            c = shp.fill.fore_color.rgb
            alpha = 1.0
            srgb = shp.fill._xPr.find(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
            if srgb is not None:
                a = srgb[0].find("{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
                if a is not None:
                    alpha = int(a.get("val")) / 100000
            return (str(c), alpha)
    except Exception:
        pass
    return None, None


def line_style(shp):
    try:
        ln = shp.line
        if ln.color and ln.color.type is not None:
            alpha = 1.0
            xPr = shp.line._get_or_add_ln()
            a = xPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
            if a is not None:
                alpha = int(a.get("val")) / 100000
            return str(ln.color.rgb), max(0.6, (ln.width.pt if ln.width else 0.75)) * SCALE / 72, alpha
    except Exception:
        pass
    return None, 0, 0


def draw_wrapped(draw, text, font, box, color, align, line_h):
    """Simple greedy wrap; box = (x0,y0,x1,y1) in px."""
    x0, y0, x1, y1 = box
    shaped = get_display(arabic_reshaper.reshape(text))
    words = shaped.split(" ")
    maxw = x1 - x0
    lines, cur = [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if draw.textlength(t, font=font) <= maxw or not cur:
            cur = t
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    y = y0
    for ln in lines:
        tw = draw.textlength(ln, font=font)
        x = x1 - tw if align == "r" else (x0 + (maxw - tw) / 2 if align == "c" else x0)
        draw.text((x, y), ln, font=font, fill=color)
        y += line_h
    return y


def render():
    prs = Presentation(PPTX)
    os.makedirs(OUT, exist_ok=True)
    W = int(prs.slide_width / 914400 * SCALE)
    H = int(prs.slide_height / 914400 * SCALE)
    paths = []
    for idx, slide in enumerate(prs.slides, start=1):
        img = Image.new("RGB", (W, H), "#0B1A2E")
        draw = ImageDraw.Draw(img)
        # background gradient (approximate)
        for y in range(H):
            t = y / H
            base = (11, 26, 46)
            col = tuple(int(base[i] * (1 - t) + c * t) for i, c in enumerate((5, 11, 20)))
            draw.line([(0, y), (W, y)], fill=col)
        for shp in slide.shapes:
            x, y = shp.left / 914400 * SCALE, shp.top / 914400 * SCALE
            w, h = shp.width / 914400 * SCALE, shp.height / 914400 * SCALE
            box = [x, y, x + w, y + h]
            if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    im = Image.open(shp.image.blob and __import__("io").BytesIO(shp.image.blob))
                    im = im.convert("RGB")
                    # apply crop
                    l = int(im.width * (shp.crop_left or 0)); r = int(im.width * (shp.crop_right or 0))
                    t_ = int(im.height * (shp.crop_top or 0)); b = int(im.height * (shp.crop_bottom or 0))
                    im = im.crop((l, t_, im.width - r, im.height - b)).resize((int(w), int(h)))
                    img.paste(im, (int(x), int(y)))
                except Exception as e:
                    print("pic fail", e)
                continue
            fill, alpha = shape_fill(shp)
            lcol, lw, lalpha = line_style(shp)
            if lcol:
                ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
                od = ImageDraw.Draw(ov)
                col = tuple(int(lcol[i:i + 2], 16) for i in (0, 2, 4)) + (int(255 * lalpha),)
                if str(shp.shape_type) == "AUTO_SHAPE (9)":
                    pass
                rad = 10
                try:
                    if shp.adjustments and len(shp.adjustments) and shp.auto_shape_type is not None \
                            and "ROUNDED" in str(shp.auto_shape_type):
                        rad = int(min(w, h) * float(shp.adjustments[0]))
                except Exception:
                    rad = 0
                od.rounded_rectangle(box, radius=rad, outline=col, width=max(1, int(lw)))
                img.paste(Image.alpha_composite(img.convert("RGBA"), ov).convert("RGB"), (0, 0))
            if fill:
                ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
                od = ImageDraw.Draw(ov)
                col = tuple(int(fill[i:i + 2], 16) for i in (0, 2, 4)) + (int(255 * alpha),)
                try:
                    rounded = shp.auto_shape_type is not None and "ROUNDED" in str(shp.auto_shape_type)
                    oval = shp.auto_shape_type is not None and "OVAL" in str(shp.auto_shape_type)
                except Exception:
                    rounded = oval = False
                if oval:
                    od.ellipse(box, fill=col)
                elif rounded:
                    rad = 4
                    try:
                        rad = int(min(w, h) * float(shp.adjustments[0]))
                    except Exception:
                        pass
                    od.rounded_rectangle(box, radius=max(2, rad), fill=col)
                else:
                    od.rectangle(box, fill=col)
                img.paste(Image.alpha_composite(img.convert("RGBA"), ov).convert("RGB"), (0, 0))
            if getattr(shp, "has_table", False) and shp.has_table:
                tbl = shp.table
                yy = y
                for r_i, row in enumerate(tbl.rows):
                    xx = x
                    rh_ = row.height / 914400 * SCALE
                    for c_i, cell in enumerate(row.cells):
                        cw_ = tbl.columns[c_i].width / 914400 * SCALE
                        try:
                            fc = str(cell.fill.fore_color.rgb)
                        except Exception:
                            fc = "0E1D30"
                        draw.rectangle([xx, yy, xx + cw_, yy + rh_], fill="#" + fc)
                        txt = cell.text.strip()
                        if txt:
                            runs = [r for p_ in cell.text_frame.paragraphs for r in p_.runs]
                            size = max([(r.font.size.pt if r.font.size else 11) for r in runs] or [11])
                            bold = any(r.font.bold for r in runs)
                            try:
                                col = "#" + str(runs[0].font.color.rgb)
                            except Exception:
                                col = "#C3D2E4"
                            draw_wrapped(draw, txt, fnt(size, bold),
                                         (xx + 8, yy + 6, xx + cw_ - 10, yy + rh_), col, "r",
                                         size * 1.25 * SCALE / 72)
                        xx += cw_
                    yy += rh_
                continue
            if not shp.has_text_frame or not shp.text_frame.text.strip():
                continue
            tf = shp.text_frame
            if "MIDDLE" in str(tf.vertical_anchor):
                tot = 0.0
                iw = w - 0.06
                for p_ in tf.paragraphs:
                    t_ = "".join(r.text for r in p_.runs)
                    if not t_:
                        continue
                    sz = max([(r.font.size.pt if r.font.size else 12) for r in p_.runs] or [12])
                    bd = any(r.font.bold for r in p_.runs)
                    shaped = get_display(arabic_reshaper.reshape(t_))
                    mw = iw * SCALE
                    nlines, cur = 1, ""
                    for wd in shaped.split(" "):
                        tt = (cur + " " + wd).strip()
                        if draw.textlength(tt, font=fnt(sz, bd)) <= mw or not cur:
                            cur = tt
                        else:
                            nlines += 1
                            cur = wd
                    ls = p_.line_spacing if isinstance(p_.line_spacing, float) else 1.2
                    tot += nlines * sz * ls * SCALE / 72 + ((p_.space_after.pt if p_.space_after else 0) * SCALE / 72)
                shp_top_override = y + max(0.0, (h - tot) / 2)
            else:
                shp_top_override = None
            cy = (shp_top_override if shp_top_override is not None else y) + \
                 (tf.margin_top / 914400 * SCALE if tf.margin_top else 0) + 2
            x0 = x + ((tf.margin_left or 0) / 914400 * SCALE) + 2
            x1 = x + w - ((tf.margin_right or 0) / 914400 * SCALE) - 2
            # vertical anchor middle approximation
            if str(tf.vertical_anchor) == "MIDDLE (3)":
                pass
            for p in tf.paragraphs:
                runs = [r for r in p.runs if r.text]
                if not runs:
                    continue
                text = "".join(r.text for r in runs)
                size = max([(r.font.size.pt if r.font.size else 12) for r in runs])
                bold = any(r.font.bold for r in runs)
                try:
                    c = str(runs[0].font.color.rgb)
                except Exception:
                    c = "FFFFFF"
                align = "r"
                try:
                    if str(p.alignment) == "CENTER (2)":
                        align = "c"
                    elif str(p.alignment) == "LEFT (1)":
                        align = "l"
                except Exception:
                    pass
                # skip pure latin-number debug? keep all
                ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
                lh = size * ls * SCALE / 72
                cy = draw_wrapped(draw, text, fnt(size, bold), (x0, cy, x1, y + h),
                                  "#" + c, align, lh)
                if p.space_after:
                    cy += p.space_after.pt * SCALE / 72
        p = os.path.join(OUT, f"slide-{idx:02d}.png")
        img.save(p, quality=90)
        paths.append(p)
    print("rendered", len(paths), "->", OUT)


if __name__ == "__main__":
    render()
