# -*- coding: utf-8 -*-
"""Build the presentation."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Inches
import deckkit
from slides_a import s01, s02, s04, s05, s06, s07, s08, s09, s10, divider
from slides_b import s12, s13, s14, s15, s16, s17
from slides_c import (s19, s20, s21, s22, s23, s24, s25, s26, s27, s28, s29)
from notes import NOTES

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "تدريب-المراقبين-الجويين-المسيرات.pptx")


def build():
    prs = Presentation()
    prs.slide_width = Inches(deckkit.SW)
    prs.slide_height = Inches(deckkit.SH)

    seq = [
        (s01, None),
        (s02, None),
        (divider, ("١", "المبحث الأول", "الإطار المفاهيمي والتنظيمي",
                   ["المصطلحات والتصنيفات", "الملاحق الدولية", "الأطر الإقليمية والوطنية", "أدوار المراقب"])),
        (s04, None), (s05, None), (s06, None), (s07, None), (s08, None), (s09, None), (s10, None),
        (divider, ("٢", "المبحث الثاني", "الاحتياجات التدريبية والعوامل البشرية",
                   ["فئات المخاطر الخمس", "الفروق التشغيلية", "العوامل البشرية", "الاتصالات"])),
        (s12, None), (s13, None), (s14, None), (s15, None), (s16, None), (s17, None),
        (divider, ("٣", "المبحث الثالث", "الإطار التدريبي المقترح وآليات تنفيذه",
                   ["مبادئ التصميم", "نموذج الكفايات", "المنهج والوحدات", "التقييم والخارطة المرحلية"])),
        (s19, None), (s20, None), (s21, None), (s22, None), (s23, None), (s24, None),
        (s25, None), (s26, None), (s27, None), (s28, None), (s29, None),
    ]
    total = len(seq)
    assert len(NOTES) == total, (len(NOTES), total)
    for i, (fn, arg) in enumerate(seq, start=1):
        if fn is divider:
            slide = fn(prs, arg[0], arg[1], arg[2], arg[3], i, total)
        else:
            slide = fn(prs, i, total)
        slide.notes_slide.notes_text_frame.text = NOTES[i - 1]
    prs.save(OUT)
    print("saved", OUT, "slides:", len(prs.slides.__iter__.__self__._sldIdLst), "total", total)


if __name__ == "__main__":
    build()
