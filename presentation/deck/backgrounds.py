# -*- coding: utf-8 -*-
"""Prepare slide backgrounds: cover-crop, darken, vignette, cover ornament."""
import os
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, 'assets')
OUT = os.path.join(SRC, 'bg')
W, H = 1920, 1080
os.makedirs(OUT, exist_ok=True)


def cover_crop(path, w=W, h=H):
    im = Image.open(path).convert('RGB')
    sw, sh = im.size
    s = max(w / float(sw), h / float(sh))
    im = im.resize((int(sw * s + 0.5), int(sh * s + 0.5)), Image.LANCZOS)
    x = (im.width - w) // 2
    y = (im.height - h) // 2
    return im.crop((x, y, x + w, y + h))


def vignette(im, strength=0.55, power=2.2):
    w, h = im.size
    mask = Image.new('L', (w, h), 0)
    d = ImageDraw.Draw(mask)
    cx, cy = w / 2.0, h / 2.0
    maxr = (cx ** 2 + cy ** 2) ** 0.5
    steps = 90
    for i in range(steps, 0, -1):
        r = maxr * i / steps
        a = int(strength * 255 * (1 - i / float(steps)) ** power)
        d.ellipse((cx - r, cy - r * 0.98, cx + r, cy + r * 0.98), fill=a)
    mask = mask.filter(ImageFilter.GaussianBlur(w // 26))
    black = Image.new('RGB', (w, h), (0, 0, 0))
    return Image.composite(black, im, mask)


def scrim(im, top=0.0, bottom=0.0, uniform=0.0):
    w, h = im.size
    grad = Image.new('L', (1, h))
    for y in range(h):
        t = y / float(h - 1)
        v = uniform + (top - uniform) * (1 - t) + (bottom - uniform) * t
        grad.putpixel((0, y), int(max(0, min(1, v)) * 255))
    grad = grad.resize((w, h))
    black = Image.new('RGB', (w, h), (4, 8, 12))
    return Image.composite(black, im, grad)


def finish(name, src, bright=1.0, sat=1.0, vig=0.5, uniform=0.42,
           top=0.0, bottom=0.0, contrast=1.0, blur=0.0):
    im = cover_crop(os.path.join(SRC, src))
    if blur:
        im = im.filter(ImageFilter.GaussianBlur(blur))
    im = ImageEnhance.Brightness(im).enhance(bright)
    im = ImageEnhance.Color(im).enhance(sat)
    im = ImageEnhance.Contrast(im).enhance(contrast)
    im = scrim(im, top=top, bottom=bottom, uniform=uniform)
    im = vignette(im, vig)
    p = os.path.join(OUT, name + '.jpg')
    im.save(p, quality=90, subsampling=1)
    return p


def gold_frame(im, inset=13, inset2=23, a1=170, a2=90):
    w, h = im.size
    layer = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    gold = (201, 163, 78)
    gold2 = (230, 206, 147)
    d.rectangle((inset, inset, w - 1 - inset, h - 1 - inset), outline=gold + (a1,), width=3)
    d.rectangle((inset2, inset2, w - 1 - inset2, h - 1 - inset2), outline=gold2 + (a2,), width=1)
    # corner blocks
    s = 13
    for (cx, cy) in [(inset, inset), (w - 1 - inset, inset),
                     (inset, h - 1 - inset), (w - 1 - inset, h - 1 - inset)]:
        d.rectangle((cx - s // 2, cy - s // 2, cx + s // 2, cy + s // 2), fill=gold + (a1,))
    return Image.alpha_composite(im.convert('RGBA'), layer).convert('RGB')


if __name__ == '__main__':
    # cover -------------------------------------------------------------
    cov = finish('cover', 'bg_cover.jpg', bright=0.82, sat=0.95, vig=0.55,
                 uniform=0.30, contrast=1.05)
    im = Image.open(cov)
    im = gold_frame(im)
    im.save(os.path.join(OUT, 'cover.jpg'), quality=92, subsampling=1)

    # section / content backgrounds ------------------------------------
    finish('radar', 'bg_radar.jpg', bright=0.55, sat=0.75, vig=0.45, uniform=0.40)
    finish('intro', 'bg_intro.jpg', bright=0.50, sat=0.60, vig=0.42, uniform=0.56)
    finish('law', 'bg_law.jpg', bright=0.55, sat=0.6, vig=0.45, uniform=0.42)
    finish('human', 'bg_human.jpg', bright=0.5, sat=0.6, vig=0.45, uniform=0.46)
    finish('sim', 'bg_sim.jpg', bright=0.5, sat=0.6, vig=0.45, uniform=0.46)
    finish('blueprint', 'bg_blueprint.jpg', bright=0.42, sat=0.7, vig=0.42, uniform=0.34)
    finish('roadmap', 'bg_roadmap.jpg', bright=0.55, sat=0.7, vig=0.45, uniform=0.38)
    finish('final', 'bg_final.jpg', bright=0.46, sat=0.6, vig=0.42, uniform=0.42)
    finish('agenda', 'bg_blueprint.jpg', bright=0.30, sat=0.5, vig=0.45,
           uniform=0.42, blur=1.2)
    finish('matrix', 'bg_radar.jpg', bright=0.30, sat=0.4, vig=0.5, uniform=0.50,
           blur=1.6)
    finish('texture', 'bg_law.jpg', bright=0.28, sat=0.35, vig=0.5,
           uniform=0.52, blur=2.0)
    print('backgrounds done')
