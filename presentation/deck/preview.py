# -*- coding: utf-8 -*-
"""Render layout.json to PNG with PIL so the deck can be visually checked."""
import json, os, sys
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display

REG = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
BOLD = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
W, H = 1280, 720
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, 'preview')
os.makedirs(OUT, exist_ok=True)

reshaper = arabic_reshaper.ArabicReshaper(
    configuration={'delete_harakat': False, 'support_ligatures': False})
FONTS = {}


def font(size, bold):
    key = (int(round(size)), bold)
    if key not in FONTS:
        FONTS[key] = ImageFont.truetype(BOLD if bold else REG, int(round(size)))
    return FONTS[key]


def has_ar(t):
    return any('\u0600' <= c <= '\u06FF' or '\uFB50' <= c <= '\uFEFF' for c in t)


def shape_text(t):
    return get_display(reshaper.reshape(t), base_dir='R')


def hexrgb(c, a=255):
    c = c.replace('#', '')
    return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16), int(a))


def measure(txt, f, spc):
    if not txt:
        return 0
    w = 0
    for ch in txt:
        w += f.getlength(ch) + spc
    return w - spc


def wrap(logical, f, spc, maxw):
    words = logical.split(' ')
    lines, cur = [], ''
    for wd in words:
        cand = (cur + ' ' + wd).strip()
        if measure(shape_text(cand), f, spc) <= maxw or not cur:
            cur = cand
        else:
            lines.append(cur)
            cur = wd
    if cur:
        lines.append(cur)
    return lines


def draw_paras(d, x, y, w, h, op):
    paras = op['paras']
    align = op.get('align', 'r')
    rtl_default = op.get('rtl', True)
    blocks = []
    for p in paras:
        size = p.get('size', 14) * 4.0 / 3.0
        bold = p.get('bold', False)
        f = font(size, bold)
        spc = p.get('spc', 0) * 4.0 / 3.0 if not has_ar(p['t']) else 0.0
        lh = p.get('lh', 1.28) * size
        rtl = p.get('rtl', rtl_default)
        lines = wrap(p['t'], f, spc, w) if op.get('wrap', True) else [p['t']]
        sa = p.get('sa', 0) * 4.0 / 3.0
        sb = p.get('sb', 0) * 4.0 / 3.0
        blocks.append((p, f, spc, lh, lines, rtl, sa, sb))
    total = sum(len(b[4]) * b[3] + b[6] + b[7] for b in blocks)
    if op.get('anchor') == 'm':
        cy = y + (h - total) / 2.0
    elif op.get('anchor') == 'b':
        cy = y + h - total
    else:
        cy = y
    for p, f, spc, lh, lines, rtl, sa, sb in blocks:
        cy += sb
        for ln in lines:
            disp = shape_text(ln) if rtl else ln
            tw = measure(disp, f, spc)
            al = p.get('align', align)
            if al == 'r':
                tx = x + w - tw
            elif al == 'c':
                tx = x + (w - tw) / 2.0
            else:
                tx = x
            col = hexrgb(p.get('color', 'F3EFE6'))
            asc = f.getmetrics()[0]
            yb = cy + (lh - f.size) / 2.0 + asc
            if spc:
                bx = tx
                for ch in disp:
                    d.text((bx, yb), ch, font=f, fill=col, anchor='ls')
                    bx += f.getlength(ch) + spc
            else:
                d.text((tx, yb), disp, font=f, fill=col, anchor='ls')
            cy += lh
        cy += sa
    return cy


def render(slide_idx, s):
    img = Image.open(os.path.join(BASE, 'assets', 'bg', os.path.basename(s['bg'])))
    img = img.convert('RGB').resize((W, H), Image.LANCZOS).convert('RGBA')
    overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    for op in s['ops']:
        k = op['k']
        if k in ('rect', 'round'):
            box = (op['x'], op['y'], op['x'] + op['w'], op['y'] + op['h'])
            fill = hexrgb(op['fill'], op['alpha'] * 255) if op['fill'] else None
            line = hexrgb(op['line'], op['lalpha'] * 255) if op['line'] else None
            lw = max(1, int(round(op['lw'] * 4.0 / 3.0)))
            if k == 'round':
                d.rounded_rectangle(box, radius=op['r'], fill=fill, outline=line, width=lw)
            else:
                d.rectangle(box, fill=fill, outline=line, width=lw)
        elif k == 'line':
            d.line((op['x1'], op['y1'], op['x2'], op['y2']),
                   fill=hexrgb(op['line'], op['lalpha'] * 255),
                   width=max(1, int(round(op['lw'] * 4.0 / 3.0))))
        elif k == 'img':
            p = Image.open(op['path']).convert('RGBA')
            p = p.resize((int(op['w']), int(op['h'])), Image.LANCZOS)
            overlay.alpha_composite(p, (int(op['x']), int(op['y'])))
        elif k == 'text':
            draw_paras(d, op['x'], op['y'], op['w'], op['h'], op)
    out = Image.alpha_composite(img, overlay).convert('RGB')
    out.save(os.path.join(OUT, 's%02d.png' % (slide_idx + 1)))
    return out


if __name__ == '__main__':
    layout = json.load(open(os.path.join(BASE, 'deck', 'layout.json'), encoding='utf-8'))
    outs = [render(i, s) for i, s in enumerate(layout)]
    cols, tw = 3, 426
    rows = (len(outs) + cols - 1) // cols
    th = int(tw * 9 / 16)
    for part in range(0, rows, 3):
        sheet = Image.new('RGB', (cols * tw + 8, 3 * th + 8), '#101010')
        for j in range(3):
            r = part + j
            if r >= rows:
                break
            for c in range(cols):
                i = r * cols + c
                if i >= len(outs):
                    break
                im = outs[i].resize((tw - 6, th - 6), Image.LANCZOS)
                sheet.paste(im, (c * tw + 4, j * th + 4))
        sheet.save('/tmp/sheet_%d.png' % (part // 3 + 1))
    print('rendered', len(outs))
