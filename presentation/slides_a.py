# -*- coding: utf-8 -*-
"""Slides 1-10: cover, agenda, section I (concept + regulation)."""
from deckkit import *
from tablekit import table
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

import os
A = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets") + "/"
TITLE = "تدريب المراقبين الجويين في التعامل مع أطقم المسيرات"


# ---------------------------------------------------------------- 1 · cover
def s01(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, "0A1626", "050B14")
    pic = picture_cover(s, A + "cover.jpg", 0, 0, SW, SH)
    # darkening veils
    veil = rect(s, 0, 0, SW, SH, fill="050C16", shape=MSO_SHAPE.RECTANGLE)
    set_alpha(veil, 0.52)
    veil2 = rect(s, 0, 0, SW, 4.1, fill="040A12", shape=MSO_SHAPE.RECTANGLE)
    set_alpha(veil2, 0.35)
    # gold frame
    frame = rect(s, 0.42, 0.62, SW - 0.84, 6.32, fill=None, line="7C6A3E",
                 line_w=0.75)
    set_line_alpha(frame, 0.55)

    tb = textbox(s, SW - M - 8.6, 0.86, 8.6, 0.32)
    write(tb.text_frame, [("قيادة القوات الجوية   ·   معهد تدريب القوات الجوية",
                           {"color": GOLD_SOFT, "size": 11, "bold": True, "spacing": 1.0})],
          first=True, space_after=0)
    tb = textbox(s, SW - M - 8.6, 1.22, 8.6, 0.32)
    write(tb.text_frame, [("دورة قادة أ  —  فرسان رمق ١٥", {"color": MUTED, "size": 10})],
          first=True, space_after=0)

    tb = textbox(s, SW - M - 10.4, 2.55, 10.4, 0.34)
    write(tb.text_frame, [("ورقة عسكرية", {"color": GOLD, "size": 12, "bold": True})],
          first=True, space_after=0)

    tb = textbox(s, SW - M - 10.6, 2.98, 10.6, 1.9)
    write(tb.text_frame, [
        ("تدريب المراقبين الجويين", {"color": INK, "size": 46, "bold": True}),
        ("في التعامل مع أطقم المسيرات", {"color": INK, "size": 33, "bold": True}),
    ], first=True, space_after=6, line_spacing=1.16)

    hline(s, SW - M - 2.0, 4.86, 2.0, color=GOLD, weight=2.0)

    tb = textbox(s, SW - M - 9.9, 5.02, 9.9, 0.8)
    write(tb.text_frame, [
        ("الإطار المفاهيمي والتنظيمي  ·  الاحتياجات التدريبية والعوامل البشرية  ·  إطار تدريبي قائم على الكفاءات",
         {"color": BODY, "size": 12}),
    ], first=True, space_after=0, line_spacing=1.35)

    # authors
    names = ["مقدم مراقب جوي / أحمد عبد الرحمن أحمد عدلان",
             "مقدم مراقب جوي / معمر خليفة محمد سعيد",
             "مقدم طيار / عامر محمد علي محمد فضل"]
    hline(s, M + 0.02, 6.28, 4.2, color="5E5233", weight=0.75)
    tb = textbox(s, SW - M - 10.0, 6.34, 10.0, 0.28)
    write(tb.text_frame, [("إعداد", {"color": GOLD, "size": 9.5, "bold": True})],
          first=True, space_after=0)
    tb = textbox(s, SW - M - 10.0, 6.62, 10.0, 0.42)
    write(tb.text_frame, ["   ·   ".join(names)], size=10, color=BODY,
          first=True, space_after=0)

    chrome(s, n, total)
    return s


# ---------------------------------------------------------------- 2 · agenda
def s02(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    watermark_circles(s, -1.6, 4.3, 5.4, alpha=0.07)
    chrome(s, n, total)
    y = heading(s, "محاور العرض", kicker="خارطة الطريق",
                sub="أربعة محاور تنتقل من التنظيم الدولي إلى إطار تدريبي قابل للتنفيذ")

    items = [
        ("٠١", "الإطار المفاهيمي والتنظيمي", "المصطلحات، التصنيفات، الملاحق الدولية، والأطر الإقليمية والوطنية"),
        ("٠٢", "الاحتياجات التدريبية", "فئات المخاطر الخمس، الفروق التشغيلية، العوامل البشرية والاتصالات"),
        ("٠٣", "الإطار التدريبي المقترح", "مبادئ التصميم، نموذج ست كفايات، وثماني وحدات تدريبية"),
        ("٠٤", "التنفيذ والتحقق", "منظومة تقييم ثلاثية، خارطة مرحلية، نتائج وتوصيات"),
    ]
    gap = 0.24
    cw = (CW - 3 * gap) / 4
    h = 3.55
    for i, (num, t, d) in enumerate(items):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.12, cw, h, fill=CARD, line=LINE, accent=GOLD)
        tb = textbox(s, x + 0.26, y + 0.34, cw - 0.52, 0.62)
        write(tb.text_frame, [(num, {"color": GOLD, "size": 30, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.26, y + 1.02, cw - 0.52, 0.72)
        write(tb.text_frame, [(t, {"color": INK, "size": 14.5, "bold": True})],
              first=True, space_after=0, line_spacing=1.25)
        hline(s, x + cw - 0.92, y + 1.84, 0.66, color="3A4E68", weight=0.75)
        tb = textbox(s, x + 0.26, y + 1.98, cw - 0.52, 1.42)
        write(tb.text_frame, [(d, {"color": BODY, "size": 10.5})],
              first=True, space_after=0, line_spacing=1.42)

    tb = textbox(s, SW - M - 9.4, 6.42, 9.4, 0.4)
    write(tb.text_frame, [("المرجع الرئيس:  ICAO Doc 10019 (2015)  ·  Circular 328  ·  الملحق ٦ الجزء الرابع  ·  Doc 9868 / 10056",
                           {"color": MUTED, "size": 9.5})], first=True, space_after=0)
    return s


# ---------------------------------------------------------------- divider
def divider(prs, num, kicker, title, lines, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, "08131F", "050B14")
    picture_cover(s, A + "radar.jpg", 0, 0, SW, SH)
    veil = rect(s, 0, 0, SW, SH, fill="050C16", shape=MSO_SHAPE.RECTANGLE)
    set_alpha(veil, 0.55)
    veil = rect(s, 0, 0, 7.2, SH, fill="050C16", shape=MSO_SHAPE.RECTANGLE)
    set_alpha(veil, 0.55)

    tb = textbox(s, SW - M - 6.6, 1.86, 6.6, 1.45)
    write(tb.text_frame, [(num, {"color": GOLD, "size": 74, "bold": True})],
          first=True, space_after=0)
    tb = textbox(s, SW - M - 6.6, 3.12, 6.6, 0.30)
    write(tb.text_frame, [(kicker, {"color": GOLD_SOFT, "size": 11.5, "bold": True,
                                    "spacing": 1.4})], first=True, space_after=0)
    tb = textbox(s, SW - M - 9.7, 3.48, 9.7, 1.1)
    write(tb.text_frame, [(title, {"color": INK, "size": 29, "bold": True})],
          first=True, space_after=0, line_spacing=1.2)
    hline(s, SW - M - 1.7, 4.86, 1.7, color=GOLD, weight=2.0)
    tb = textbox(s, SW - M - 9.0, 5.04, 9.0, 0.9)
    write(tb.text_frame, [("   ·   ".join(lines), {"color": BODY, "size": 11.5})],
          first=True, space_after=0, line_spacing=1.4)
    chrome(s, n, total)
    return s


# ---------------------------------------------------------------- 4 · intro
def s04(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "لم تعد الأجواء حكراً على الطائرات المأهولة",
                kicker="المقدمة",
                sub="العقد الأخير: تحوّل جوهري في بنية الحركة الجوية، ومسيرات أصبحت طرفاً فاعلاً في العمليات العسكرية والأمنية والأعمال الجوية المدنية")

    left_w = 3.55
    right_w = CW - left_w - 0.42
    body = [
        "المسيرات اليوم في التصوير الجوي والمسح ومراقبة المحاصيل ومكافحة الحرائق والنقل الطبي، وفي العمليات العسكرية والأمنية على حدٍّ سواء.",
        "واكبه جهد تنظيمي متسارع قادته منظمة الطيران المدني الدولي (ICAO): تعديلات الملاحق الأول والثاني والحادي عشر، ثم الملحق السادس – الجزء الرابع.",
        "وأُصدر الدليل المرجعي للطائرات المأهولة عن بُعد (Doc 10019)، وإطار إدارة حركة الطائرات غير المأهولة (UTM)، مع لوائح أمريكية (Part 107) وأوروبية (2019/945 و2019/947) ومفهوم U-space.",
    ]
    tb = textbox(s, SW - M - right_w, y + 0.04, right_w, 3.1)
    write(tb.text_frame, body, size=12.5, color=BODY, space_after=13, line_spacing=1.42)

    # stat column (left side)
    stats = [("٣٣", "ساعة", "إغلاق مدرجات مطار غاتويك — ديسمبر ٢٠١٨"),
             ("+١٠٠٠", "رحلة", "تأثرت بالحادثة"),
             ("١٤٠", "ألف مسافر", "تضرروا من توقف التشغيل")]
    cx = M
    cy = y + 0.02
    for num, unit, cap in stats:
        card(s, cx, cy, left_w, 0.96, fill=CARD_2, line=LINE, accent=GOLD, accent_w=0.045)
        tb = textbox(s, cx + 0.14, cy + 0.14, left_w - 0.38, 0.46)
        write(tb.text_frame, [(num + "  ", {"color": GOLD_SOFT, "size": 20, "bold": True}),
                              (unit, {"color": INK, "size": 11, "bold": True})],
              first=True, space_after=0, line_spacing=1.1)
        tb = textbox(s, cx + 0.14, cy + 0.56, left_w - 0.38, 0.32)
        write(tb.text_frame, [(cap, {"color": MUTED, "size": 9})], first=True, space_after=0)
        cy += 1.10

    # closing band
    band = rect(s, M, 5.72, CW, 0.92, fill="0F2136", line="2A4258", radius=0.05)
    tb = textbox(s, M + 0.55, 5.86, CW - 1.1, 0.66, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("والسؤال: هل تعدّ برامج الإعداد الحالية المراقبَ الجوي لهذا التحوّل؟",
                           {"color": INK, "size": 14, "bold": True})],
          first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 5 · problem
def s05(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "مشكلة البحث: فجوة الإنسان",
                kicker="المشكلة",
                sub="إلى جانب فجوات التقنية والتشريع، ظلّ تدريب المراقب موجَّهاً إلى منطق تشغيلي لم يعد قائماً وحده")

    cols = [
        ("المنطق التقليدي", GOLD, [
            "طيار على المتن",
            "رؤية طبيعية مباشرة",
            "مناورة فورية",
            "اتصال لاسلكي مباشر",
        ]),
        ("ما تُدخله المسيرات", CYAN, [
            "طيار بعيد قد يكون على آلاف الكيلومترات",
            "وصلة C2 قابلة للانقطاع",
            "قدرات أداء متفاوتة",
            "محدودية الرؤية والتفادي",
            "احتمال اختراق غير مصرّح به",
        ]),
        ("النتيجة", "B98A5E", [
            "مقررات تقليدية بلا وحدات إلزامية متخصصة",
            "لا سيناريوهات محاكاة واقعية للطوارئ المرتبطة بها",
            "لا معايير تقييم مصمّمة لهذا الغرض",
        ]),
    ]
    gap = 0.3
    cw = (CW - 2 * gap - 0.9) / 3
    for i, (title, col, items) in enumerate(cols):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.08, cw, 3.16, fill=CARD, line=LINE, accent=col)
        tb = textbox(s, x + 0.24, y + 0.28, cw - 0.5, 0.34)
        write(tb.text_frame, [(title, {"color": col, "size": 12.5, "bold": True})],
              first=True, space_after=0)
        hline(s, x + 0.24, y + 0.68, cw - 0.5, color="2A4058", weight=0.75)
        tb = textbox(s, x + 0.24, y + 0.84, cw - 0.5, 2.30)
        write(tb.text_frame, [(it, {"color": BODY, "size": 11}) for it in items], bullet="•",
              first=True, space_after=9, line_spacing=1.36)

    band = rect(s, M, 6.10, CW, 0.76, fill="101F33", line="2A4258", radius=0.05)
    tb = textbox(s, M + 0.5, 6.18, CW - 1.0, 0.60, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("الفجوة تزداد تعقيداً بعدم اكتمال الأطر المعيارية الدولية — ما يوجب تصميم التدريب بصيغة «مرنة قابلة للتحديث» لا بصيغة مقرر ثابت.",
                           {"color": GOLD_SOFT, "size": 12})], first=True, space_after=0,
          line_spacing=1.32)
    return s


# ---------------------------------------------------------------- 6 · terms
def s06(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "المصطلحات والتعريفات الأساسية", kicker="المفاهيم",
                sub="عشر مصطلحات تُشكّل الحدّ الأدنى من اللغة المشتركة للمراقب الجوي")

    terms = [
        ("الطائرة غير المأهولة (UA)", "أي طائرة تُشغَّل أو تُصمَّم للتشغيل دون طيار على متنها"),
        ("الطائرة المأهولة عن بُعد (RPA)", "طائرة غير مأهولة يقودها طيار من محطة بعيدة، تُدمج في إدارة الحركة الجوية على قدم المساواة"),
        ("نظام الطائرات غير المأهولة (UAS)", "الطائرة ومحطات القيادة البعيدة ووصلات القيادة والتحكم والعناصر اللازمة للتشغيل"),
        ("وصلة القيادة والتحكم (C2 Link)", "رابط البيانات بين الطائرة ومحطة الطيار البعيد، داخل خط البصر الراديوي أو خارجه"),
        ("الطيار البعيد (Remote Pilot)", "من يمارس مهام جوهرية في التشغيل ويُرخَّص وفق الملحق الأول للـ ICAO"),
        ("الرؤية والتفادي (DAA)", "وظائف إدراك المخاطر وتفاديها، وتنفرد بها الأنظمة التقنية خارج مدى الرؤية"),
        ("إدارة حركة الطائرات غير المأهولة (UTM)", "خدمات رقمية لإدارة أعداد كبيرة من الطائرات في الفضاء المنخفض الارتفاع"),
        ("منظومة U-space", "منظومة أوروبية لخدمات إدارة الحركة غير المأهولة بالتنسيق مع الحركة الجوية"),
        ("التشغيل خارج مدى الرؤية (BVLOS)", "الأكثر تعقيداً تنظيمياً في تقييم المخاطر"),
        ("التدريب القائم على الكفاءات (CBTA)", "تحديد الكفايات ثم تصميم التدريب والتقييم حول إثباتها بسلوك قابل للملاحظة"),
    ]
    gap_x, gap_y = 0.26, 0.115
    cw = (CW - gap_x) / 2
    ch = (4.86 - 4 * gap_y) / 5
    for i, (t, d) in enumerate(terms):
        col, row = i % 2, i // 2
        x = SW - M - cw - col * (cw + gap_x)
        yy = y + 0.04 + row * (ch + gap_y)
        card(s, x, yy, cw, ch, fill=CARD_2, line="1D3149", accent=GOLD, accent_w=0.04)
        tb = textbox(s, x + 0.16, yy + 0.08, cw - 0.34, 0.24)
        write(tb.text_frame, [(t, {"color": INK, "size": 10.2, "bold": True})],
              first=True, space_after=0, line_spacing=1.1)
        tb = textbox(s, x + 0.16, yy + 0.32, cw - 0.34, ch - 0.36)
        write(tb.text_frame, [(d, {"color": MUTED, "size": 8.2})], first=True,
              space_after=0, line_spacing=1.18)
    return s


# ---------------------------------------------------------------- 7 · classes
def s07(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "تصنيفات المسيرات وأنماط تشغيلها", kicker="التصنيف",
                sub="ثلاث فئات أوروبية تنظّم العمليات غير المأهولة، ومقابلها الأمريكي القائم على الوزن والاستخدام")

    cats = [("الفئة المفتوحة", "دون تصريح مسبق", "ضمن حدود الارتفاع والمسافة والوزن، وخطر منخفض"),
            ("الفئة المحددة", "تقييم مخاطر وإقرار أو تصريح", "منهجية SORA لتقييم المخاطر التشغيلية"),
            ("الفئة المعتمدة", "شهادات وتراخيص", "للعمليات عالية المخاطر، على غرار الطيران المأهول")]
    gap = 0.28
    cw = (CW - 2 * gap) / 3
    for i, (t, tag, d) in enumerate(cats):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.04, cw, 2.16, fill=CARD, line=LINE, accent=CYAN if i == 2 else GOLD)
        tb = textbox(s, x + 0.24, y + 0.24, cw - 0.5, 0.32)
        write(tb.text_frame, [(t, {"color": INK, "size": 14, "bold": True})], first=True,
              space_after=0)
        tb = textbox(s, x + 0.24, y + 0.62, cw - 0.5, 0.28)
        write(tb.text_frame, [(tag, {"color": GOLD_SOFT, "size": 9.5, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.24, y + 0.96, cw - 0.5, 1.05)
        write(tb.text_frame, [(d, {"color": BODY, "size": 10.5})], first=True,
              space_after=0, line_spacing=1.4)

    # bottom two panels
    by = y + 2.44
    pw = (CW - 0.3) / 2
    panels = [
        ("التصنيف الوظيفي الفني", [
            "أنظمة قادرة على الامتثال الكامل لمقومات التجهيز فتُشارك كأي طائرة مأهولة.",
            "أنظمة محدودة القدرات تستلزم ترتيبات تشغيلية خاصة أو فصلاً مكانياً.",
            "التنظيم الأمريكي يقوم على الوزن والاستخدام مع تركيز على الطائرات الصغيرة.",
        ]),
        ("أنماط التشغيل والغلاف الأدائي", [
            "أجنحة ثابتة، ومروحيات، ومركبات متعددة المراوح، ومناطيد.",
            "لكل غلاف أداء أثره في المناورة والمسافات الفاصلة.",
            "لا يُطلب من المراقب خبرة بكل طراز، بل نموذج ذهني واضح للفروق السلوكية الأساسية.",
        ]),
    ]
    for i, (t, items) in enumerate(panels):
        x = SW - M - pw - i * (pw + 0.3)
        rect(s, x, by, pw, 2.02, fill=CARD_2, line="1D3149", radius=0.04)
        tb = textbox(s, x + 0.26, by + 0.16, pw - 0.5, 0.3)
        write(tb.text_frame, [(t, {"color": CYAN, "size": 11.5, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.26, by + 0.52, pw - 0.5, 1.4)
        write(tb.text_frame, [(it, {"color": BODY, "size": 9.8}) for it in items], bullet="•",
              first=True, space_after=5, line_spacing=1.32)
    return s


# ---------------------------------------------------------------- 8 · intl framework
def s08(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "الإطار التنظيمي الدولي", kicker="التنظيم الدولي",
                sub="من المادة الثامنة في اتفاقية شيكاغو إلى معايير العمليات الدولية للطائرات المأهولة عن بُعد")

    nodes = [
        ("١٩٤٤", "اتفاقية شيكاغو — المادة ٨", "تشغيل الطائرة التي تطير بلا طيار وفق تصريح خاص وبشروط تضمن تجنّب أي خطر على الطائرات المدنية"),
        ("٢٠١٥", "Doc 10019 — الدليل المرجعي", "المسيرات «شريك متكافئ» قادر على التفاعل مع المراقبة الجوية في الزمن الحقيقي، وبدائل إدارة فقدان القيادة"),
        ("٢٠٢٢", "الملحق ١ — التعديل ١٧٥", "رخص طياري المسيرات تُصدر عن سلطة الترخيص في دولة المشغّل من ٣ نوفمبر ٢٠٢٢: طرف جديد في الاتصال اللاسلكي"),
        ("٢٠٢٤", "الملحق ٦ — الجزء الرابع", "سارية من ٢٢ يوليو ٢٠٢٤ وواجبة التطبيق في ٢٦ نوفمبر ٢٠٢٦: العمليات والأداء والتجهيز والاتصال والملاحة والصيانة"),
    ]
    gap = 0.26
    cw = (CW - 3 * gap) / 4
    line_y = y + 0.60
    hline(s, M + 0.1, line_y, CW - 0.2, color="2E4A64", weight=1.0)
    for i, (yr, t, d) in enumerate(nodes):
        x = SW - M - cw - i * (cw + gap)
        cx = x + cw / 2
        dot = rect(s, cx - 0.055, line_y - 0.055, 0.11, 0.11, fill=GOLD,
                   shape=MSO_SHAPE.OVAL)
        tb = textbox(s, x, line_y - 0.46, cw, 0.3)
        write(tb.text_frame, [(yr, {"color": GOLD_SOFT, "size": 15, "bold": True,
                                    "align": PP_ALIGN.CENTER})], first=True, space_after=0)
        card(s, x, line_y + 0.30, cw, 2.08, fill=CARD_2, line="1D3149")
        tb = textbox(s, x + 0.2, line_y + 0.46, cw - 0.42, 0.6)
        write(tb.text_frame, [(t, {"color": INK, "size": 10.5, "bold": True})],
              first=True, space_after=0, line_spacing=1.22)
        tb = textbox(s, x + 0.2, line_y + 1.10, cw - 0.42, 1.2)
        write(tb.text_frame, [(d, {"color": MUTED, "size": 8.8})], first=True,
              space_after=0, line_spacing=1.3)

    # doc chips row (full content width, sized to fit)
    dy = line_y + 2.50
    tb = textbox(s, SW - M - 5.0, dy - 0.26, 5.0, 0.26)
    write(tb.text_frame, [("حزمة الوثائق المرجعية", {"color": GOLD, "size": 9, "bold": True})],
          first=True, space_after=0)
    chips = ["Doc 4444 — إدارة الحركة الجوية", "Doc 9868 — PANS-TRG",
             "Doc 10056 — تدريب المراقبين", "Doc 9683 — العوامل البشرية",
             "Doc 9432 — العبارات اللاسلكية"]
    from qa import text_width_in
    gap_c = 0.12
    widths = [text_width_in(c, 8.6, False) * 0.82 + 0.34 for c in chips]
    scale = (CW - gap_c * (len(chips) - 1)) / sum(widths)
    widths = [w * scale for w in widths]
    cx = SW - M
    for c, w in zip(chips, widths):
        cx -= w
        ch = rect(s, cx, dy, w, 0.34, fill="132741", line="2C4763", radius=0.28)
        ch.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        write(ch.text_frame, [(c, {"color": BODY, "size": 8.4, "align": PP_ALIGN.CENTER})],
              first=True, space_after=0)
        cx -= gap_c

    band = rect(s, M, 6.36, CW, 0.56, fill="0F2136", line="2A4258", radius=0.05)
    tb = textbox(s, M + 0.45, 6.42, CW - 0.9, 0.44, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("أربعة متطلبات للدمج: ألّا يترتب أثر جوهري على مستخدمي المجال الجوي الحاليين، والامتثال لأنظمة الطيران المأهول، وعدم الإخلال بمستويات السلامة، وأن تُدار العمليات وتُعتبر مكافئة من قِبل المراقبة الجوية.",
                           {"color": GOLD_SOFT, "size": 9.5})], first=True, space_after=0,
          line_spacing=1.25)
    return s


# ---------------------------------------------------------------- 9 · regional
def s09(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "الأطر الإقليمية والوطنية", kicker="المستوى الإقليمي",
                sub="أربعة نماذج تنظيمية، وما يلزم المراقب معرفته من كل منها")

    items = [
        ("الولايات المتحدة", [
            "تصريح مسبق لدخول الفضاء الجوي الخاضع للمراقبة (14 CFR 107.41).",
            "منظومة LAANC تمنح تصاريح آلية شبه فورية ضمن حدود خرائط المنشآت.",
            "الأثر التدريبي: قراءة بيانات الرحلات وحدود تدخّل المراقب.",
        ]),
        ("الاتحاد الأوروبي", [
            "اللائحتان (EU) 2019/945 و2019/947 وتصنيفاتهما الثلاثة.",
            "منهجية SORA لتقييم المخاطر التي طوّرتها هيئة EASA.",
            "منظومة U-space بخدماتها الرقمية وواجهتها الإجرائية مع المراقبة الجوية.",
        ]),
        ("الأطر العربية", [
            "عُمان والسعودية والإمارات: تسجيل الطائرات وترخيص التشغيل والتصاريح المسبقة.",
            "مسافات حظر حول المطارات والموانئ والمواقع المحمية.",
            "المعالجة الصريحة لمتطلبات تدريب المراقبين محدودة التفصيل.",
        ]),
        ("السودان", [
            "قانون سلامة الطيران المدني — المادة ٧٩ (البند ١/هـ).",
            "سند وطني باشتراطه تصريحاً من السلطة لطيران أي مركبة هوائية بدون طيار.",
            "نقطة انطلاق تحليلية لبناء نموذج الكفايات المحلي.",
        ]),
    ]
    gap = 0.24
    cw = (CW - gap) / 2
    ch = 1.94
    for i, (t, lines) in enumerate(items):
        col, row = i % 2, i // 2
        x = SW - M - cw - col * (cw + gap)
        yy = y + 0.02 + row * (ch + 0.16)
        accent = GOLD if i == 3 else LINE
        card(s, x, yy, cw, ch, fill=CARD if i != 3 else "16273C", line=LINE, accent=accent)
        tb = textbox(s, x + 0.26, yy + 0.18, cw - 0.5, 0.3)
        write(tb.text_frame, [(t, {"color": INK, "size": 13, "bold": True})], first=True,
              space_after=0)
        tb = textbox(s, x + 0.26, yy + 0.56, cw - 0.5, ch - 0.66)
        write(tb.text_frame, [(l, {"color": BODY, "size": 9.6}) for l in lines], bullet="•",
              first=True, space_after=4, line_spacing=1.28)

    band = rect(s, M, 6.24, CW, 0.62, fill="0F2136", line="2A4258", radius=0.05)
    tb = textbox(s, M + 0.45, 6.30, CW - 0.9, 0.5, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("خلاصة الباحثين: معظم الأطر العربية تعالج الجانب التنظيمي للطائرة والمشغّل والتصاريح، بينما يبقى البعد التدريبي للمراقب محدود التفصيل.",
                           {"color": GOLD_SOFT, "size": 10})], first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 10 · role
def s10(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "أدوار المراقب ومسؤولياته القانونية", kicker="الدور القانوني",
                sub="مبدأ «التكافؤ التشغيلي والقانوني»: العبرة ليست في وجود طيار على المتن بل في امتثال الطائرة للقواعد")

    quote = rect(s, SW - M - 6.2, y + 0.02, 6.2, 1.34, fill="132741", line="2C4763", radius=0.04)
    vline(s, SW - M - 0.16, y + 0.10, 1.18, color=GOLD, weight=3.0)
    tb = textbox(s, SW - M - 5.85, y + 0.18, 5.55, 1.1)
    write(tb.text_frame, [
        ("تُطبَّق القواعد ذاتها في تقسيم المسؤوليات، وأُكّد عدم مسؤولية المراقب عن الحوادث الناجمة عن أنظمة غير ممتثلة للأنظمة الدولية أو الوطنية.",
         {"color": INK, "size": 11.5}),
        ("IFATCA, 2024", {"color": MUTED, "size": 9}),
    ], first=True, space_after=6, line_spacing=1.34)

    tb = textbox(s, M, y + 0.1, 4.5, 1.3)
    write(tb.text_frame, [
        ("ومع ذلك، ينشئ هذا المبدأ حاجة إلى تعزيز قدرات المراقب في أربعة جوانب:", {"color": BODY, "size": 10.5}),
    ], first=True, space_after=0, line_spacing=1.35)

    four = [
        ("حدود المسؤولية", "فهم حدود مسؤولية المراقب وتمييزها عن مسؤوليات المشغّل"),
        ("تقييم المخاطر", "إجراءات قائمة على تقييم المخاطر مع الطائرات المصرّح بها وغير المصرّح بها"),
        ("الإبلاغ عن الأحداث", "معرفة متطلبات الإبلاغ ومسارها الزمني"),
        ("التنسيق الفوري", "مع سلطة الطيران المدني والجهات الأمنية والمطارات ومنظومات الكشف والتصدي"),
    ]
    gap = 0.22
    cw = (CW - 3 * gap) / 4
    yy = y + 1.62
    for i, (t, d) in enumerate(four):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, yy, cw, 1.94, fill=CARD, line=LINE, accent=GOLD, accent_w=0.045)
        numbered_badge(s, x + cw - 0.62, yy + 0.20, 0.36, ar_num(i + 1), size=9.5)
        tb = textbox(s, x + 0.22, yy + 0.66, cw - 0.46, 0.3)
        write(tb.text_frame, [(t, {"color": INK, "size": 11.5, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.22, yy + 1.02, cw - 0.46, 0.84)
        write(tb.text_frame, [(d, {"color": MUTED, "size": 9.2})], first=True,
              space_after=0, line_spacing=1.3)

    band = rect(s, M, 6.22, CW, 0.64, fill="101F33", line="2A4258", radius=0.05)
    tb = textbox(s, M + 0.45, 6.28, CW - 0.9, 0.52, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("لم يعد المراقب مسؤولاً عن الفصل بين المسارات فحسب، بل عن إدارة علاقة مركّبة مع طيار ليس على متن الطائرة، ووصلة بيانات قابلة للانقطاع، ومستخدمين جدد بعضهم غير ممتثل.",
                           {"color": GOLD_SOFT, "size": 10.5})], first=True, space_after=0,
          line_spacing=1.3)
    return s
