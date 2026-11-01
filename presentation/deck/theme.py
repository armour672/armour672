# -*- coding: utf-8 -*-
"""Design system: colours, typography helpers, drawing primitives for a
16:9 (1280x720 px @96dpi) Arabic military presentation."""
import json
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

PXE = 9525  # English Metric Units per pixel at 96 dpi


def E(px):
    return Emu(int(round(px * PXE)))


def hx(s):
    return RGBColor.from_string(s.replace('#', '').upper())


# ---------------------------------------------------------------- palette
INK      = '080F16'
PANEL    = '0B1723'
PANEL2   = '0E1E2C'
GOLD     = 'C9A34E'
GOLD_L   = 'E6CE93'
GOLD_D   = '8C6C2C'
CREAM    = 'F3EFE6'
WHITE    = 'FFFFFF'
MUTED    = 'A9B7C4'
MUTED2   = '7F8F9E'
RED      = 'B4453D'
RED_L    = 'D9736A'
TEAL     = '5FAFB0'
GREEN    = '4F8A6A'

FONT = 'Arial'

AR_DIGITS = str.maketrans('0123456789', '٠١٢٣٤٥٦٧٨٩')


def has_arabic(t):
    return any('\u0600' <= ch <= '\u06FF' or '\u0750' <= ch <= '\u077F'
               or '\uFB50' <= ch <= '\uFEFF' for ch in t)


def ar(n):
    """Western digits -> Arabic-Indic (Hindi) digits."""
    return str(n).translate(AR_DIGITS)


# ---------------------------------------------------------------- xml utils
def _alpha(parent, value):
    el = etree.SubElement(parent, qn('a:alpha'))
    el.set('val', str(int(round(max(0.0, min(1.0, value)) * 100000))))
    return el


def _set_fill(shape, color, alpha):
    shape.fill.solid()
    shape.fill.fore_color.rgb = hx(color)
    if alpha < 1:
        sf = shape.fill._xPr.find(qn('a:solidFill'))
        _alpha(sf.find(qn('a:srgbClr')), alpha)


def _set_line(shape, color, alpha, width):
    ln = shape.line
    if color is None:
        ln.fill.background()
        return
    ln.color.rgb = hx(color)
    ln.width = Pt(width)
    if alpha < 1:
        sf = shape.line.color._xFill
        _alpha(sf.find(qn('a:srgbClr')), alpha)


def _no_shadow(shape):
    spPr = shape._element.spPr
    if spPr.find(qn('a:effectLst')) is None:
        etree.SubElement(spPr, qn('a:effectLst'))


def _soft_shadow(shape, blur=18, dist=6, direction=270, color='000000', alpha=0.45):
    spPr = shape._element.spPr
    lst = spPr.find(qn('a:effectLst'))
    if lst is None:
        lst = etree.SubElement(spPr, qn('a:effectLst'))
    sh = etree.SubElement(lst, qn('a:outerShdw'))
    sh.set('blurRad', str(int(blur * PXE)))
    sh.set('dist', str(int(dist * PXE)))
    sh.set('dir', str(int(direction * 60000)))
    sh.set('rotWithShape', '0')
    clr = etree.SubElement(sh, qn('a:srgbClr'))
    clr.set('val', color)
    _alpha(clr, alpha)


def _darken(obj, amount=0.0, br=None):
    pass


# ---------------------------------------------------------------- text
_ANCHOR = {'t': MSO_ANCHOR.TOP, 'm': MSO_ANCHOR.MIDDLE, 'b': MSO_ANCHOR.BOTTOM}
_ALIGN = {'r': PP_ALIGN.RIGHT, 'l': PP_ALIGN.LEFT, 'c': PP_ALIGN.CENTER,
          'j': PP_ALIGN.JUSTIFY}


def _font_runs(para, p):
    rtl = p.get('rtl', True)
    for child in para._p.findall(qn('a:r')):
        rPr = child.find(qn('a:rPr'))
        if rPr is None:
            continue
        latin = rPr.find(qn('a:latin'))
        if latin is not None:
            latin.set('typeface', p.get('font', FONT))
        for tag in ('a:ea', 'a:cs'):
            el = rPr.find(qn(tag))
            if el is None:
                el = etree.Element(qn(tag))
                rPr.insert_element_before(el, 'a:sym', 'a:hlinkClick',
                                          'a:hlinkMouseOver', 'a:rtl', 'a:extLst')
            el.set('typeface', p.get('font', FONT))


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = E(1280)
        self.prs.slide_height = E(720)
        self.blank = self.prs.slide_layouts[6]
        self.layout = []

    # ------------------------------------------------------------ slide
    def new(self, bg):
        s = self.prs.slides.add_slide(self.blank)
        s.shapes.add_picture(bg, 0, 0, E(1280), E(720))
        self.slide = s
        self.ops = []
        self.layout.append({'bg': bg, 'ops': self.ops})
        return s

    # ------------------------------------------------------------ shapes
    def rect(self, x, y, w, h, fill=None, alpha=1.0, line=None, lalpha=1.0,
             lw=1.0, radius=0, shadow=False, dash=None):
        shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        shp = self.slide.shapes.add_shape(shp_type, E(x), E(y), E(w), E(h))
        if radius:
            try:
                shp.adjustments[0] = min(0.5, radius / float(min(w, h)))
            except Exception:
                pass
        if fill is None:
            shp.fill.background()
        else:
            _set_fill(shp, fill, alpha)
        _set_line(shp, line, lalpha, lw)
        if dash:
            shp.line.dash_style = dash
        if shadow:
            _soft_shadow(shp)
        else:
            _no_shadow(shp)
        shp.text_frame.word_wrap = False
        self.ops.append({'k': 'rect' if not radius else 'round', 'x': x, 'y': y,
                         'w': w, 'h': h, 'fill': fill, 'alpha': alpha,
                         'line': line, 'lalpha': lalpha, 'lw': lw,
                         'r': radius})
        return shp

    def line(self, x1, y1, x2, y2, color=GOLD, alpha=1.0, w=1.0):
        from pptx.enum.shapes import MSO_CONNECTOR
        cn = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                             E(x1), E(y1), E(x2), E(y2))
        cn.line.color.rgb = hx(color)
        cn.line.width = Pt(w)
        if alpha < 1:
            sfill = cn.line.color._xFill
            _alpha(sfill.find(qn('a:srgbClr')), alpha)
        self.ops.append({'k': 'line', 'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                         'line': color, 'lalpha': alpha, 'lw': w})
        return cn

    def hline(self, x1, x2, y, color=GOLD, alpha=0.4, w=1.0):
        return self.line(x1, y, x2, y, color, alpha, w)

    def pic(self, path, x, y, w, h=None):
        h = h if h is not None else w
        self.slide.shapes.add_picture(path, E(x), E(y), E(w), E(h))
        self.ops.append({'k': 'img', 'path': path, 'x': x, 'y': y,
                         'w': w, 'h': h})
        return None

    def text(self, x, y, w, h, paras, anchor='t', align='r', wrap=True,
             rtl=True, autofit=False):
        """paras: str | list of dict(t=..., size=, bold=, color=, spc=, lh=,
        sa=, sb=, align=, rtl=)"""
        if isinstance(paras, str):
            paras = [{'t': paras}]
        tb = self.slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = _ANCHOR[anchor]
        for i, p in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            al = p.get('align', align)
            para.alignment = _ALIGN[al]
            r = para.add_run()
            r.text = p['t']
            f = r.font
            f.name = p.get('font', FONT)
            f.size = Pt(p.get('size', 14))
            f.bold = bool(p.get('bold', False))
            f.italic = bool(p.get('italic', False))
            f.color.rgb = hx(p.get('color', CREAM))
            rPr = r.font._rPr
            if p.get('spc') and not has_arabic(p['t']):
                rPr.set('spc', str(int(p['spc'] * 100)))
            lh = p.get('lh', 1.28)
            para.line_spacing = lh
            if p.get('sa'):
                para.space_after = Pt(p['sa'])
            if p.get('sb'):
                para.space_before = Pt(p['sb'])
            prtl = p.get('rtl', rtl)
            pPr = para._p.get_or_add_pPr()
            pPr.set('rtl', '1' if prtl else '0')
            pPr.set('algn', {'r': 'r', 'l': 'l', 'c': 'ctr', 'j': 'just'}[al])
            _font_runs(para, p)
        self.ops.append({'k': 'text', 'x': x, 'y': y, 'w': w, 'h': h,
                         'anchor': anchor, 'align': align,
                         'paras': paras, 'rtl': rtl, 'wrap': wrap})
        return tb

    # ------------------------------------------------------------ save
    def save(self, pptx_path, layout_path):
        self.prs.save(pptx_path)
        with open(layout_path, 'w', encoding='utf-8') as f:
            json.dump(self.layout, f, ensure_ascii=False)
