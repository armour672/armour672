# -*- coding: utf-8 -*-
"""Shared table helper."""
from deckkit import *


def table(slide, x, y, w, col_widths, rows, header=True, row_h=None,
          head_h=0.44, head_size=11, body_size=10.5, head_fill="16283E",
          band_a="0E1D30", band_b="112338", head_color=GOLD_SOFT,
          body_color=BODY, aligns=None, bold_first_col=True):
    """rows: list of list of str.

    RTL table: ``col_widths[0]`` and the first value of each row belong to the
    RIGHT-most column (the leading column in Arabic reading order).
    """
    ncols = len(col_widths)
    nrows = len(rows)
    heights = []
    for i in range(nrows):
        heights.append(head_h if (header and i == 0) else (row_h or 0.42))
    gt = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(w),
                                Inches(sum(heights)))
    tbl = gt.table
    # strip the theme table style -> no style, no grid
    tblPr = tbl._tbl.tblPr
    for el in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(el)
    el = tblPr.makeelement(qn("a:tableStyleId"), {})
    el.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # No Style, No Grid
    tblPr.append(el)
    tbl.first_row = False
    tbl.horz_banding = False
    for j, cw in enumerate(col_widths):
        tbl.columns[len(col_widths) - 1 - j].width = Inches(cw)
    for i, hgt in enumerate(heights):
        tbl.rows[i].height = Inches(hgt)
    for i, row in enumerate(rows):
        is_head = header and i == 0
        for j, val in enumerate(row):
            ac = ncols - 1 - j              # actual (left-to-right) column index
            cell = tbl.cell(i, ac)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(head_fill if is_head else
                                          (band_a if (i % 2) else band_b))
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.14)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            al = PP_ALIGN.RIGHT
            if aligns:
                al = aligns[j]
            rtl(p, al)
            p.line_spacing = 1.2
            r = p.add_run()
            r.text = str(val)
            style_run(r, size=head_size if is_head else body_size,
                      bold=is_head or (bold_first_col and j == 0),
                      color=head_color if is_head else body_color)
    # hairline: give header a gold bottom edge via a thin rect over it
    if header:
        hline(slide, x, y + head_h, w, color=GOLD, weight=1.25)
    return tbl
