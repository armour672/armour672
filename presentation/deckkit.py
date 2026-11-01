# -*- coding: utf-8 -*-
"""Small design-system layer over python-pptx for a right-to-left Arabic deck."""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

# ---------- palette ----------
INK        = "F2F6FB"   # primary text
BODY       = "C3D2E4"   # body text
MUTED      = "8497AF"   # secondary text
GOLD       = "C9A24B"
GOLD_SOFT  = "E4C87E"
CYAN       = "58C8C0"
NAVY_DEEP  = "060E1A"
NAVY       = "0B1A2E"
CARD       = "112338"
CARD_2     = "0E1D30"
LINE       = "20374F"

FONT = "Arial"
SW, SH = 13.333, 7.5
M = 0.85          # side margin
CW = SW - 2 * M   # content width


AR_DIGITS = "٠١٢٣٤٥٦٧٨٩"


def ar_num(n, pad=2):
    """Arabic-Indic digits, zero padded (Arabic text renders LTR for numbers)."""
    out = "".join(AR_DIGITS[int(d)] for d in str(n))
    while pad and len(out) < pad:
        out = AR_DIGITS[0] + out
    return out


def rgb(h):
    return RGBColor.from_string(h)


# ---------- run / paragraph helpers ----------
def _rpr(run):
    return run._r.get_or_add_rPr()


def style_run(run, size=14, bold=False, color=BODY, font=FONT, italic=False,
              spacing=None, alpha=None, underline=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    if underline:
        f.underline = True
    f.name = font
    if alpha is None:
        f.color.rgb = rgb(color)
    else:
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls
        rPr = _rpr(run)
        for tag in ("a:solidFill", "a:noFill"):
            el = rPr.find(qn(tag))
            if el is not None:
                rPr.remove(el)
        rPr.append(parse_xml(
            '<a:solidFill %s><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:solidFill>'
            % (nsdecls("a"), color, int(alpha * 100000))))
    rPr = _rpr(run)
    for tag in ("a:cs", "a:ea"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", font)
    if spacing is not None:
        rPr.set("spc", str(int(spacing * 100)))
    return run


def _ppr(p):
    return p._p.get_or_add_pPr()


def rtl(p, align=PP_ALIGN.RIGHT):
    """Right-to-left paragraph with Arabic shaping hints."""
    pPr = _ppr(p)
    pPr.set("rtl", "1")
    p.alignment = align
    return p


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, margins=(0, 0, 0, 0),
            wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_right, tf.margin_top, tf.margin_bottom = (
        Inches(margins[0]), Inches(margins[1]), Inches(margins[2]), Inches(margins[3]))
    # kill autofit surprises
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    for tag in ("a:normAutofit", "a:spAutoFit"):
        el = bodyPr.find(qn(tag))
        if el is not None:
            bodyPr.remove(el)
    bodyPr.append(bodyPr.makeelement(qn("a:noAutofit"), {}))
    return tb


def para(tf, first=False):
    if first and len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs:
        return tf.paragraphs[0]
    return tf.add_paragraph()


def set_bullet(p, char="•", color=GOLD, indent=0.20, size_pct=100):
    """Real PowerPoint bullet with a hanging indent (RTL-safe)."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls
    pPr = _ppr(p)
    pPr.set("marL", str(int(indent * 914400)))
    pPr.set("indent", str(int(-indent * 914400)))
    xml = ('<a:buClr %s><a:srgbClr val="%s"/></a:buClr>' % (nsdecls("a"), color))
    xml += ('<a:buSzPct %s val="%d"/>' % (nsdecls("a"), size_pct * 1000))
    xml += ('<a:buChar %s char="%s"/>' % (nsdecls("a"), char))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buSzPct", "a:buFont"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    for frag in xml.split("><"):
        pass
    from pptx.oxml import parse_xml as px
    wrapper = px('<a:root %s>%s</a:root>' % (nsdecls("a"), xml))
    for el in list(wrapper):
        pPr.append(el)


def set_autonum(p, indent=0.24, style="arabicDbPeriod", color=GOLD, size_pct=100):
    """Real PowerPoint auto-numbering (Arabic-Indic digits), RTL-safe."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls
    pPr = _ppr(p)
    pPr.set("marL", str(int(indent * 914400)))
    pPr.set("indent", str(int(-indent * 914400)))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buSzPct", "a:buFont"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    xml = ('<a:buClr %s><a:srgbClr val="%s"/></a:buClr>' % (nsdecls("a"), color))
    xml += ('<a:buSzPct %s val="%d"/>' % (nsdecls("a"), size_pct * 1000))
    xml += ('<a:buAutoNum %s type="%s" startAt="1"/>' % (nsdecls("a"), style))
    wrapper = parse_xml('<a:root %s>%s</a:root>' % (nsdecls("a"), xml))
    for el in list(wrapper):
        pPr.append(el)


def write(tf, lines, size=14, color=BODY, bold=False, align=PP_ALIGN.RIGHT,
          space_after=6, line_spacing=1.32, font=FONT, spacing=None,
          first=True, bullet=None, bullet_color=GOLD, space_before=0, numbered=False):
    """lines: str or list of (text, dict(overrides)) or str."""
    if isinstance(lines, str):
        lines = [lines]
    out = []
    for i, item in enumerate(lines):
        opts = {}
        if isinstance(item, tuple):
            item, opts = item
        p = tf.paragraphs[0] if (i == 0 and first) else tf.add_paragraph()
        rtl(p, opts.get("align", align))
        sb = opts.get("space_before", space_before)
        if sb:
            p.space_before = Pt(sb)
        p.space_after = Pt(opts.get("space_after", space_after))
        p.line_spacing = opts.get("line_spacing", line_spacing)
        if numbered:
            set_autonum(p, color=opts.get("bullet_color", bullet_color))
        elif bullet:
            set_bullet(p, char=bullet, color=opts.get("bullet_color", bullet_color))
        r = p.add_run()
        r.text = item
        _sz = opts.get("size", size)
        if _sz < 9:            # Arial runs small for Arabic: keep captions legible
            opts["size"] = 9
        style_run(r, size=opts.get("size", size), bold=opts.get("bold", bold),
                  color=opts.get("color", color), font=opts.get("font", font),
                  spacing=opts.get("spacing", spacing),
                  underline=opts.get("underline", False))
        out.append(p)
    return out


def set_alpha(shape, alpha):
    """Make a solid-filled shape semi transparent (alpha 0..1)."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls
    spPr = shape.fill._xPr
    srgb = spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    for el in list(srgb):
        srgb.remove(el)
    srgb.append(parse_xml('<a:alpha %s val="%d"/>' % (nsdecls("a"), int(alpha * 100000))))


def set_line_alpha(shape, alpha):
    """Make a shape outline semi transparent."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls
    ln = shape.line._get_or_add_ln()
    srgb = ln.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    for el in list(srgb):
        srgb.remove(el)
    srgb.append(parse_xml('<a:alpha %s val="%d"/>' % (nsdecls("a"), int(alpha * 100000))))


# ---------- shape helpers ----------
def rect(slide, x, y, w, h, fill=CARD, alpha=None, line=None, line_w=0.75,
         radius=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, shadow=False):
    if radius is None:
        shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.adjustments[0] = radius
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
        if alpha is not None:
            set_alpha(shp, alpha)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w)
    if shp.has_text_frame:
        shp.text_frame.word_wrap = True
        shp.text_frame.margin_left = 0
        shp.text_frame.margin_right = 0
        shp.text_frame.margin_top = 0
        shp.text_frame.margin_bottom = 0
    return shp


def hline(slide, x, y, w, color=LINE, weight=0.75):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y),
                                    Inches(x + w), Inches(y))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(weight)
    ln.shadow.inherit = False
    return ln


def vline(slide, x, y, h, color=LINE, weight=0.75):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y),
                                    Inches(x), Inches(y + h))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(weight)
    ln.shadow.inherit = False
    return ln


def picture_cover(slide, path, x, y, w, h):
    """Insert picture cropped (cover) into the given rect."""
    from PIL import Image
    iw, ih = Image.open(path).size
    box_ratio = w / h
    img_ratio = iw / ih
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    if img_ratio > box_ratio:      # crop sides
        frac = (1 - box_ratio / img_ratio) / 2
        pic.crop_left = frac
        pic.crop_right = frac
    else:                          # crop top/bottom
        frac = (1 - img_ratio / box_ratio) / 2
        pic.crop_top = frac
        pic.crop_bottom = frac
    return pic


# ---------- slide chrome ----------
def background(slide, top=NAVY, bottom=NAVY_DEEP, angle=90):
    fill = slide.background.fill
    fill.gradient()
    fill.gradient_angle = angle
    stops = fill.gradient_stops
    stops[0].color.rgb = rgb(top)
    stops[0].position = 0.0
    stops[1].color.rgb = rgb(bottom)
    stops[1].position = 1.0


def watermark_circles(slide, x, y, d, color=GOLD, alpha=0.10, rings=3, weight=0.75):
    for i in range(rings):
        dd = d - i * d / (rings * 2.2)
        shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                     Inches(dd), Inches(dd))
        shp.fill.background()
        shp.line.color.rgb = rgb(color)
        shp.line.width = Pt(weight)
        set_line_alpha(shp, alpha)
        shp.shadow.inherit = False


CLASSIFIED = "محظور"


def chrome(slide, idx, total, footer_txt=None):
    """Centred classification band: -٤- / محظور (top) and محظور (bottom)."""
    # thin rule framing the footer
    hline(slide, M, 7.02, CW, color="1B3049", weight=0.75)

    # top centre: page number between two dashes (LTR paragraph so the two
    # dashes cannot swap sides under the bidi algorithm, e.g. -٠٤-)
    tb = textbox(slide, SW / 2 - 1.6, 0.10, 3.2, 0.22)
    pnum = write(tb.text_frame, [("-" + ar_num(idx, 0) + "-",
                                  {"color": GOLD, "size": 9, "bold": True,
                                   "align": PP_ALIGN.CENTER})],
                 first=True, space_after=0)[0]
    pnum._p.get_or_add_pPr().set("rtl", "0")
    # top centre: classification, underlined
    tb = textbox(slide, SW / 2 - 1.6, 0.31, 3.2, 0.24)
    write(tb.text_frame, [(CLASSIFIED, {"color": GOLD_SOFT, "size": 9.5, "bold": True,
                                        "align": PP_ALIGN.CENTER, "spacing": 2.0,
                                        "underline": True})], first=True, space_after=0)
    # bottom centre: classification, underlined
    tb = textbox(slide, SW / 2 - 1.6, 7.13, 3.2, 0.24)
    write(tb.text_frame, [(CLASSIFIED, {"color": GOLD_SOFT, "size": 9.5, "bold": True,
                                        "align": PP_ALIGN.CENTER, "spacing": 2.0,
                                        "underline": True})], first=True, space_after=0)


def heading(slide, title, kicker=None, sub=None, y=0.55):
    """Right-aligned kicker + title + gold rule + optional subtitle.

    Returns the y coordinate where content may start.
    """
    from qa import text_width_in
    x_right = SW - M
    ty = y
    if kicker:
        tb = textbox(slide, x_right - 7.0, ty, 7.0, 0.26)
        write(tb.text_frame, [(kicker, {"color": GOLD, "size": 9.5, "bold": True,
                                        "spacing": 1.2})], first=True, space_after=0)
        ty += 0.27
    tsize = 25.5
    tw = text_width_in(title, tsize, True) * 1.04
    th = 0.50 if tw <= 10.5 else 0.86
    tb = textbox(slide, x_right - 10.6, ty, 10.6, th + 0.06)
    write(tb.text_frame, [(title, {"color": INK, "size": tsize, "bold": True})],
          first=True, space_after=0, line_spacing=1.1)
    rule_y = ty + th + 0.05
    hline(slide, x_right - 1.35, rule_y, 1.35, color=GOLD, weight=2.0)
    ny = rule_y + 0.42
    if sub:
        tb = textbox(slide, x_right - 10.6, rule_y + 0.14, 10.6, 0.34)
        write(tb.text_frame, [(sub, {"color": MUTED, "size": 11.5})], first=True,
              space_after=0, line_spacing=1.24)
        ny = rule_y + 0.56
    return ny


def card(slide, x, y, w, h, fill=CARD, line=LINE, accent=None, accent_w=0.055,
         radius=0.045):
    """Card with an optional gold/cyan bar on its right (RTL leading) edge."""
    shp = rect(slide, x, y, w, h, fill=fill, line=line, radius=radius)
    if accent:
        bar = rect(slide, x + w - accent_w, y + 0.001, accent_w, h - 0.002,
                   fill=accent, shape=MSO_SHAPE.RECTANGLE)
    return shp


def numbered_badge(slide, x, y, d, n, color=GOLD, txt_color=None, size=12):
    shp = rect(slide, x, y, d, d, fill=None, line=color, line_w=1.0,
               shape=MSO_SHAPE.OVAL)
    tf = shp.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    write(tf, [(n, {"color": txt_color or GOLD_SOFT, "size": size, "bold": True,
                    "align": PP_ALIGN.CENTER})], first=True, space_after=0)
    return shp
