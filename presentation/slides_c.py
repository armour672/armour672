# -*- coding: utf-8 -*-
"""Slides 19-29: proposed framework, evaluation, roadmap, results, recommendations."""
from deckkit import *
from tablekit import table
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from slides_a import A


# ---------------------------------------------------------------- 19 · principles
def s19(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "أسس نظرية ومبادئ تصميم", kicker="الإطار المقترح",
                sub="غاية التدريب ليست إنجاز ساعات رسمية أو استظهار محتوى، بل بلوغ أداء وظيفي مُثبت وفق معايير محددة مسبقاً")

    quote = rect(s, SW - M - 5.3, y, 5.3, 1.02, fill="132741", line="2C4763", radius=0.04)
    vline(s, SW - M - 0.16, y + 0.06, 0.9, color=GOLD, weight=3.0)
    tb = textbox(s, SW - M - 4.95, y + 0.12, 4.65, 0.8)
    write(tb.text_frame, [
        ("أُدرجت إجراءات التدريب القائم على الكفاءات للمراقبين في Doc 9868، ونموذج الكفايات الموائم في دليل ICAO (2022).",
         {"color": BODY, "size": 9.6}),
    ], first=True, space_after=0, line_spacing=1.32)

    tb = textbox(s, M, y + 0.06, 4.9, 0.9)
    write(tb.text_frame, [("خمسة مبادئ يحكم بها الإطار:", {"color": INK, "size": 11, "bold": True})],
          first=True, space_after=0)

    pr = [
        ("الموائمة المحلية", "بناء النموذج على البيئة المحلية بدل النقل الأعمى للنماذج الجاهزة"),
        ("الملاحظة المتعددة", "تقييم جمعي متعدد الملاحظين بدل الحكم على أداء واحد"),
        ("شفافية المعايير", "إعلان معايير الأداء للمتدرّب مسبقاً قبل التنفيذ"),
        ("امتداد التحقق", "مراقبة الأداء أثناء العمل والمراجعات الدورية بعد التدريب"),
        ("التغذية الراجعة الفورية", "التقييم نشاط تكويني مصاحب، لا حكم نهائي لاحق"),
    ]
    gap = 0.2
    cw = (CW - 4 * gap) / 5
    yy = y + 1.10
    for i, (t, d) in enumerate(pr):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, yy, cw, 2.42, fill=CARD_2, line="1D3149", accent=GOLD, accent_w=0.045)
        numbered_badge(s, x + cw - 0.56, yy + 0.22, 0.34, ar_num(i + 1), size=9)
        tb = textbox(s, x + 0.2, yy + 0.86, cw - 0.42, 0.62)
        write(tb.text_frame, [(t, {"color": INK, "size": 11, "bold": True})],
              first=True, space_after=0, line_spacing=1.2)
        tb = textbox(s, x + 0.2, yy + 1.54, cw - 0.42, 0.74)
        write(tb.text_frame, [(d, {"color": MUTED, "size": 8.8})], first=True,
              space_after=0, line_spacing=1.3)

    band = rect(s, M, 6.02, CW, 0.82, fill="101F33", line="2A4258", radius=0.05)
    tb = textbox(s, M + 0.5, 6.10, CW - 1.0, 0.66, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("النموذج مُصمَّم ليكون «قابلاً للتحقيق» لا مثالياً فقط: ست كفايات، وثماني وحدات تدريبية، وستون ساعة كحد أدنى، ومنظومة تقييم ثلاثية، وخارطة تنفيذ يمكن البدء بها فوراً.",
                           {"color": GOLD_SOFT, "size": 10.5})], first=True, space_after=0,
          line_spacing=1.32)
    return s


# ---------------------------------------------------------------- 20 · job analysis
def s20(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "تحليل الاحتياج والتوصيف الوظيفي", kicker="الخطوة الأولى",
                sub="سبع مهام جديدة تُدخل متطلبات «تفاوضية» و«تحليلية» تختلف عن منطق السيطرة الفورية التقليدي")

    tasks = [
        "تقييم صلاحية قبول تشغيل مستخدم ذي قدرات محدودة",
        "التحقق من بيانات خطة الطيران وكفاية المعلومات عن الطائرة والطيار المسير ومحطة القيادة",
        "مراقبة الالتزام بالحيّز المعتمد وتقدير اتجاهات الخروج قبل وقوعه",
        "إدارة حالة طارئة معلنة مع الحفاظ على الفصل الآمن للحركة المأهولة",
        "إدارة اقتحام غير مصرّح به بتصعيد وتنسيق",
        "إدارة الاتصال المزدوج مع الطيار المسير والطائرات المتأثرة",
        "توثيق الحدث وإبلاغ أنظمة السلامة",
    ]
    # two-column numbered list, RTL: first column at right
    col_w = (CW - 0.36) / 2
    per = 4
    pitch = 1.02
    for i, t in enumerate(tasks):
        col = 0 if i < per else 1
        row = i if i < per else i - per
        x = SW - M - col_w - col * (col_w + 0.36)
        yy = y + 0.02 + row * pitch
        rect(s, x, yy, col_w, 0.86, fill=CARD_2, line="1D3149", radius=0.04)
        num = i + 1
        numbered_badge(s, x + col_w - 0.64, yy + 0.25, 0.36, ar_num(num), size=9)
        tb = textbox(s, x + 0.24, yy + 0.14, col_w - 1.02, 0.62, anchor=MSO_ANCHOR.MIDDLE)
        write(tb.text_frame, [(t, {"color": BODY, "size": 10})], first=True,
              space_after=0, line_spacing=1.3)

    # highlight card in the free slot of the left column
    hx = M
    hy = y + 0.02 + 3 * pitch
    rect(s, hx, hy, col_w, 0.86, fill="16273C", line="2C4763", radius=0.04)
    vline(s, hx + 0.12, hy + 0.12, 0.62, color=CYAN, weight=2.5)
    tb = textbox(s, hx + 0.3, hy + 0.13, col_w - 0.55, 0.62, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("جوهر التغيير: متطلبات «تفاوضية» و«تحليلية» بدل منطق السيطرة الفورية التقليدي.",
                           {"color": GOLD_SOFT, "size": 9.6})], first=True, space_after=0,
          line_spacing=1.3)

    band = rect(s, M, 6.16, CW, 0.68, fill="132741", line="2C4763", radius=0.05)
    tb = textbox(s, M + 0.5, 6.22, CW - 1.0, 0.56, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("مخرج هذه الخطوة: تقرير فجوات، ونموذج كفايات محلي، وخطة تدريب سنوية — وهي مخرجات لا تحتاج انتظار تقنيات متقدمة.",
                           {"color": GOLD_SOFT, "size": 10.5})], first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 21 · competency model
def s21(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "نموذج الكفايات المقترح", kicker="ست كفايات",
                sub="كل كفاية مقترنة بسلوك قابل للملاحظة ومعيار أداء قابل للقياس")

    rows = [
        ["مجال الكفاءة", "الكفاية المطلوبة", "معيار الأداء"],
        ["المعرفة الفنية والتنظيمية", "يميّز فئات العمليات ويستدعي القيد التنظيمي للموقف", "استدعاء صحيح للقيد في أقل من دقيقة"],
        ["الوعي الظرفي الممتد", "يدمج مصادر بيانات متعددة وينتبه لشذوذ الطائرة", "إدراك الانحراف قبل التصعيد ودمج ثلاثة مصادر"],
        ["إدارة الحركة والسلامة", "يختار أداة الفصل الأنسب لمستوى الوصلة والاستجابة", "عدم تجاوز المسافات الفاصلة وفق الإجراء"],
        ["إدارة الطوارئ والانقطاع", "ينفّذ سلسلة الطوارئ في الزمن المحدد", "تنفيذ كامل الخطوات دون نقص خلال الزمن المحدد"],
        ["الاتصال اللاسلكي والتعاون", "يصوغ تعليمات موجزة ذات سلطة ويؤكد الاستلام", "خلو التعليمات من الغموض والخطأ المصطلحي"],
        ["العوامل البشرية والتصرف المهني", "يتبنى استراتيجية لإدارة ضغط العمل ويسجّل الشكوك", "انخفاض ضغط العمل الحرج وارتفاع الوعي الظرفي"],
    ]
    table(s, M, y + 0.06, CW, [2.63, 4.45, 4.55], rows, row_h=0.66, head_h=0.46,
          head_size=10.5, body_size=10)

    tb = textbox(s, SW - M - 9.4, 6.66, 9.4, 0.3)
    write(tb.text_frame, [("كل كفاية تُقاس بثلاثة أدوات: اختبار معرفي، وملاحظة سلوك في المحاكاة، ومؤشر أداء من بيانات الجلسة.",
                           {"color": MUTED, "size": 9})], first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 22 · curriculum
def s22(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "المنهج التدريبي ووحداته", kicker="ثماني وحدات",
                sub="موزّعة على ثلاث مراحل — أساسي، متوسط، عالي — بستين ساعة تدريبية كحد أدنى")

    rows = [
        ["الوحدة", "المرحلة", "المحتوى الأساسي", "الساعات"],
        ["مدخل إلى الطائرات غير المأهولة", "أساسي", "المصطلحات والأنماط والتصنيفات وأمثلة عملية", "٦"],
        ["الإطار التنظيمي والمسؤوليات القانونية", "أساسي", "الملاحق الدولية واللوائح الوطنية وحدود المسؤولية والإبلاغ", "٦"],
        ["الأداء والفروق التشغيلية", "أساسي", "السرعات والارتفاعات والمناورة والاستجابة وحدود التجهيز", "٦"],
        ["الوعي الظرفي ودمج مصادر البيانات", "متوسط", "واجهات UTM ومؤشرات الاتصال وتقييم البلاغات", "٨"],
        ["إدارة التعارض والفصل في بيئة مختلطة", "متوسط", "تدبير النزاعات وأزمنة التعليمات والسيناريوهات المعقدة", "١٠"],
        ["إدارة الطوارئ والاقتحامات", "متوسط", "فقدان الاتصال والعطل والإنهاء القسري وغير الممتثلين", "١٠"],
        ["الاتصال اللاسلكي مع الطيار المسير", "متوسط", "العبارات المنظمة والاستجابة والتواصل الثلاثي", "٦"],
        ["العوامل البشرية وإدارة الضغط الذهني", "عالي", "العبء واليقظة والثقة في الأتمتة والمرونة المعرفية", "٨"],
        ["الإجمالي — حدّ أدنى إلزامي", "٣ مراحل", "", "٦٠"],
    ]
    t = table(s, M, y + 0.04, CW, [4.03, 1.55, 4.9, 1.15], rows, row_h=0.455,
              head_h=0.42, head_size=10, body_size=9.6)
    # emphasise total row
    for j in range(4):
        cell = t.cell(9, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb("1B2E45")
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                style_run(r, size=10.5, bold=True, color=GOLD_SOFT)

    tb = textbox(s, SW - M - 10.2, 6.72, 10.2, 0.3)
    write(tb.text_frame, [("يُضاف إلى الساعات زمن التقييم والمراجعة الفردية — ويُعاد توزيعها محلياً بحسب حجم الحركة وموارد الوحدة.",
                           {"color": MUTED, "size": 9})], first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 23 · tools
def s23(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "الأدوات والوسائل التدريبية", kicker="الوسائل",
                sub="المحاكاة أداة محورية، تُعزَّز بتوليد السيناريوهات والتدريب المشترك وحُزَم منخفضة الكلفة")

    tools = [
        ("المحاكاة القياسية والمتقدمة", "الأداة المحورية", [
            "تقارب في قطاع مزدحم",
            "فقدان التحكم في القيادة",
            "اقتحام مجال جوي",
            "مهام متزامنة مع انصراف الانتباه",
        ]),
        ("الطيار الافتراضي بالذكاء الاصطناعي", "مُعزِّز للسيناريو", [
            "إنتاج السيناريوهات النادرة وتكرارها",
            "تغذية راجعة فورية قابلة للقياس",
            "حوكمة واضحة للأداة وحدودها",
        ]),
        ("التدريب المشترك", "بناء اللغة التشغيلية", [
            "المراقبون والطيارون المسيرون",
            "مشغلو UTM والجهات الأمنية",
            "توقعات متبادلة مكتوبة وموثقة",
        ]),
        ("التمارين المكتبية", "الحالات المؤسسية", [
            "الاقتحامات والتنسيق الأمني",
            "التصعيد مع سلطة الطيران المدني",
            "إدارة الأزمة الإعلامية والتشغيلية",
        ]),
    ]
    gap = 0.22
    cw = (CW - 3 * gap) / 4
    for i, (t, tag, items) in enumerate(tools):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.06, cw, 2.98, fill=CARD, line=LINE,
             accent=CYAN if i in (1, 2) else GOLD)
        tb = textbox(s, x + 0.22, y + 0.24, cw - 0.46, 0.22)
        write(tb.text_frame, [(tag, {"color": GOLD, "size": 8.5, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.22, y + 0.48, cw - 0.46, 0.62)
        write(tb.text_frame, [(t, {"color": INK, "size": 11, "bold": True})],
              first=True, space_after=0, line_spacing=1.2)
        hline(s, x + 0.22, y + 1.16, cw - 0.46, color="2A4058", weight=0.75)
        tb = textbox(s, x + 0.22, y + 1.26, cw - 0.46, 1.66, anchor=MSO_ANCHOR.MIDDLE)
        write(tb.text_frame, [(it, {"color": BODY, "size": 9.2}) for it in items], bullet="•",
              first=True, space_after=6, line_spacing=1.3)

    band = rect(s, M, 5.66, CW, 1.16, fill="132741", line="2C4763", radius=0.05)
    tb = textbox(s, M + 0.5, 5.74, CW - 1.0, 0.4)
    write(tb.text_frame, [("بديل منخفض الكلفة للأطر محدودة الموارد", {"color": GOLD_SOFT,
                                                                      "size": 11, "bold": True})],
          first=True, space_after=2)
    tb = textbox(s, M + 0.5, 6.08, CW - 1.0, 0.68)
    write(tb.text_frame, [("حُزَم التدريب النقالة ووحدات المحاكاة الخفيفة تتيح الوصول إلى الوحدات النائية دون استثمار في منظومات كاملة، وتُستخدم مع سيناريوهات مكتوبة وموحّدة ومعايير تقييم واحدة، بما يحفظ اتساق المخرجات بين المواقع.",
                           {"color": BODY, "size": 10})], first=True, space_after=0,
          line_spacing=1.32)
    return s


# ---------------------------------------------------------------- 24 · evaluation
def s24(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "منظومة التقييم ثلاثية المستويات", kicker="القياس",
                sub="من المعرفة المفاهيمية إلى الأداء الملاحَظ إلى قياس الأثر — مع شروط للصحة والثبات")

    levels = [
        ("المستوى الأول", "المعرفة المفاهيمية", [
            "اختبارات قصيرة متكررة",
            "التركيز على حالات الاستدلال لا الاسترجاع",
            "تغطية القيود التنظيمية وفئات العمليات",
        ], GOLD),
        ("المستوى الثاني", "الأداء في المحاكاة", [
            "معايير سلوك قابلة للملاحظة",
            "تقييم جمعي متعدد الملاحظين",
            "تقليل التباين بين المدربين",
        ], CYAN),
        ("المستوى الثالث", "قياس الأثر", [
            "عبء العمل بمقاييس معيارية (NASA-TLX)",
            "الوعي الظرفي بالاستجواب أو مقاييس مرتبطة بالمهمة",
            "مؤشرات بيانات المحاكاة: التعليمات/الدقيقة وزمن التأخر في القرار",
        ], "B98A5E"),
    ]
    gap = 0.3
    cw = (CW - 2 * gap) / 3
    # staircase: level 1 rightmost, each next lower
    for i, (lv, t, items, col) in enumerate(levels):
        x = SW - M - cw - i * (cw + gap)
        yy = y + 0.06 + i * 0.34
        h = 3.34 - i * 0.34 + 0.34
        card(s, x, yy, cw, h, fill=CARD, line=LINE, accent=col)
        tb = textbox(s, x + 0.24, yy + 0.22, cw - 0.5, 0.24)
        write(tb.text_frame, [(lv, {"color": col, "size": 9, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.24, yy + 0.48, cw - 0.5, 0.34)
        write(tb.text_frame, [(t, {"color": INK, "size": 13.5, "bold": True})],
              first=True, space_after=0)
        hline(s, x + 0.24, yy + 0.94, cw - 0.5, color="2A4058", weight=0.75)
        tb = textbox(s, x + 0.24, yy + 1.06, cw - 0.5, h - 1.2)
        write(tb.text_frame, [(it, {"color": BODY, "size": 9.4}) for it in items], bullet="•",
              first=True, space_after=7, line_spacing=1.3)

    band = rect(s, M, 5.94, CW, 0.9, fill="101F33", line="2A4258", radius=0.05)
    vline(s, SW - M - 0.16, 6.02, 0.74, color=GOLD, weight=3.0)
    tb = textbox(s, M + 0.5, 6.02, CW - 1.2, 0.74, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("شرط الباحثين: صحة الأداة وثبات التقييم وإمكانية المقارنة عبر الزمن — لتحقق التعلّم وتحسين البرنامج، لا لمجرد إصدار حكم على المتدرّب.",
                           {"color": GOLD_SOFT, "size": 10.5})], first=True, space_after=0,
          line_spacing=1.3)
    return s


# ---------------------------------------------------------------- 25 · roadmap
def s25(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "خارطة التنفيذ المرحلية والتحقق", kicker="التنفيذ",
                sub="ثلاث مراحل تبدأ فوراً من النقاط التحليلية والتنظيمية دون انتظار تقنيات متقدمة")

    phases = [
        ("المرحلة الأولى", "٣ أشهر", ["تحليل الفجوات وجرد العمليات", "تحديد الكفايات المحلية وإعداد النموذج الموائم"],
         "المخرج: تقرير فجوات، نموذج كفايات، خطة سنوية"),
        ("المرحلة الثانية", "٦ أشهر", ["إعداد المناهج والسيناريوهات وتدريب المدربين", "تشغيل تجربة محدودة وتكييف أدوات التقييم"],
         "المخرج: مناهج معتمدة وسيناريوهات مكتوبة"),
        ("المرحلة الثالثة", "١٢ شهراً", ["تعميم التدريب وإدماجه في البرامج الأولية والتحويلية", "إنشاء لجنة المراجعة الدورية وربط التدريب بمؤشرات السلامة"],
         "المخرج: برنامج مدمج في المسار المهني"),
    ]
    gap = 0.3
    cw = (CW - 2 * gap) / 3
    for i, (ph, dur, items, out) in enumerate(phases):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.02, cw, 2.5, fill=CARD, line=LINE, accent=GOLD)
        tb = textbox(s, x + 0.24, y + 0.18, cw - 0.5, 0.26)
        write(tb.text_frame, [(ph, {"color": GOLD_SOFT, "size": 10, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.24, y + 0.46, cw - 0.5, 0.42)
        write(tb.text_frame, [(dur, {"color": INK, "size": 19, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.24, y + 0.98, cw - 0.5, 0.92)
        write(tb.text_frame, [(it, {"color": BODY, "size": 9.2}) for it in items], bullet="•",
              first=True, space_after=4, line_spacing=1.28)
        tb = textbox(s, x + 0.24, y + 1.94, cw - 0.5, 0.44)
        write(tb.text_frame, [(out, {"color": CYAN, "size": 8.6})], first=True,
              space_after=0, line_spacing=1.25)

    # validation panel
    py = y + 2.68
    rect(s, M, py, CW, 2.0, fill="0E1D30", line="1D3149", radius=0.04)
    tb = textbox(s, SW - M - 5.6, py + 0.16, 5.35, 0.3)
    write(tb.text_frame, [("التحقق التجريبي — دراسة شبه تجريبية", {"color": CYAN, "size": 11.5,
                                                                    "bold": True})],
          first=True, space_after=0)
    tb = textbox(s, SW - M - 5.6, py + 0.5, 5.35, 1.36)
    write(tb.text_frame, [
        ("على ما لا يقل عن أربعين مراقباً، بمتغير مستقل هو البرنامج المقترح مقابل التقليدي، واختبار قبلي وبعدي فوري، وقياس مؤجل بعد ثلاثة أشهر لثبات الأثر.",
         {"color": BODY, "size": 9.4}),
    ], first=True, space_after=0, line_spacing=1.32)

    deps = [("الأداء التشغيلي", "الأخطاء، زمن القرار، دقة إدارة الطوارئ"),
            ("الحالة الذهنية", "ضغط العمل، الوعي الظرفي، الثقة"),
            ("السلامة متوسطة المدى", "جودة الإبلاغ، تكرار الحالات الحرجة")]
    tb = textbox(s, M + 0.35, py + 0.2, 3.0, 0.26)
    write(tb.text_frame, [("المتغيرات التابعة", {"color": GOLD, "size": 8.5, "bold": True})],
          first=True, space_after=0)
    dw = 4.9
    dy = py + 0.50
    for i, (t, d) in enumerate(deps):
        yy = dy + i * 0.50
        rect(s, M + 0.3, yy, dw, 0.42, fill="132741", line="2C4763", radius=0.05)
        tb = textbox(s, M + 0.45, yy + 0.06, dw - 0.3, 0.3, anchor=MSO_ANCHOR.MIDDLE)
        write(tb.text_frame, [(t + "   ", {"color": INK, "size": 9.4, "bold": True}),
                              ("—  " + d, {"color": MUTED, "size": 8.8})],
              first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 26 · challenges
def s26(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "التحديات وسبل مواجهتها", kicker="المخاطر المؤسسية",
                sub="أربعة تحديات واقعية، لكل منها مقابل تدبيري قابل للتنفيذ داخل الموازنة الحالية")

    rows = [
        ("الكلفة", "تأجيل التدريبات المكلفة مثل المحاكيات الكاملة", "حُزَم تدريب نقالة ووحدات محاكاة خفيفة، وتدرّج في الاستثمار"),
        ("ندرة المدربين", "محدودية الكوادر المؤهلة للتدريب على بيئة المسيرات", "الابتعاث أو استجلاب مدربين، وتأهيل مدربين داخليين بالتوازي"),
        ("الإيقاع التنظيمي", "تغيّر المعايير الدولية والوطنية بوتيرة متسارعة", "حسن إدارة التحديث: مراجعة دورية للمحتوى وربطه بالتعديلات"),
        ("المقاومة المهنية", "تردد الكوادر تجاه مهام ومسؤوليات جديدة", "التحفيز والزيادة الأفقية، وإشراك المراقبين والطيارين المسيرين في التصميم"),
    ]
    hy = y + 0.02
    col_ws = [2.35, 4.3, 4.95]
    xs, cursor = [], SW - M
    for w in col_ws:
        cursor -= w
        xs.append(cursor)
    for x, w, t in zip(xs, col_ws, ["التحدي", "التجلي العملي", "سبل المواجهة"]):
        tb = textbox(s, x + 0.12, hy, w - 0.24, 0.28)
        write(tb.text_frame, [(t, {"color": GOLD_SOFT, "size": 9.5, "bold": True})],
              first=True, space_after=0)
    hline(s, M, hy + 0.30, CW, color=GOLD, weight=1.0)

    ry = hy + 0.44
    rh = 1.10
    for i, (a, b, c) in enumerate(rows):
        rect(s, M, ry - 0.04, CW, rh - 0.06, fill="0E1D30" if i % 2 == 0 else "112338",
             radius=0.03)
        numbered_badge(s, SW - M - 0.62, ry + 0.30, 0.40, ar_num(i + 1), size=9.5)
        tb = textbox(s, xs[0] + 0.12, ry + 0.10, xs[0] and col_ws[0] - 0.92, rh - 0.22,
                     anchor=MSO_ANCHOR.MIDDLE)
        write(tb.text_frame, [(a, {"color": INK, "size": 11.5, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, xs[1] + 0.12, ry + 0.10, col_ws[1] - 0.24, rh - 0.22,
                     anchor=MSO_ANCHOR.MIDDLE)
        write(tb.text_frame, [(b, {"color": BODY, "size": 10})], first=True,
              space_after=0, line_spacing=1.3)
        tb = textbox(s, xs[2] + 0.12, ry + 0.10, col_ws[2] - 0.24, rh - 0.22,
                     anchor=MSO_ANCHOR.MIDDLE)
        write(tb.text_frame, [(c, {"color": CYAN, "size": 10})], first=True,
              space_after=0, line_spacing=1.3)
        ry += rh

    tb = textbox(s, SW - M - 10.4, 6.66, 10.4, 0.3)
    write(tb.text_frame, [("قاعدة عامة: كل تحدٍّ يُدار بثلاثة عناصر — بديل منخفض الكلفة، وشراكة مؤسسية، وجدول زمني معلن.",
                           {"color": MUTED, "size": 9})], first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 27 · results
def s27(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "الخلاصة والنتائج", kicker="النتائج",
                sub="كيف نُعدّ المراقب الجوي ليكون شريكاً آمناً وفاعلاً في بيئة متعددة المستخدمين؟")

    results = [
        ("أولاً", "الدمج واقع تشغيلي", "منظومة معيارية دولية آخذة في الاكتمال (معايير العمليات الدولية واجبة التطبيق ٢٦ نوفمبر ٢٠٢٦)، مع بقاء مكونات فنية كأنظمة الرؤية والتفادي موضع استكمال."),
        ("ثانياً", "التكافؤ لا يُلغي التعقيد", "وجود الطيار المسير ووصلة القيادة والتحكم ينشئ حاجة إلى كفايات جديدة لا تغطيها برامج الإعداد التقليدية."),
        ("ثالثاً", "خمس فئات للمخاطر", "كل فئة تستلزم سيناريو تدريبياً خاصاً ومعايير أداء قابلة للملاحظة ومعياراً زمنياً للاستجابة، لا مجرد شرح نظري."),
        ("رابعاً", "عبء العمل مشروط", "الأدبيات لا تُثبت زيادة غير مقبولة في عبء العمل شريطة أدوات دعم وإجراءات واضحة، والأدلة محدودة تستدعي قياساً محلياً مستمراً."),
        ("خامساً", "الاحتياج الحقيقي", "القدرة على الإدراك والاستدلال واتخاذ القرار في ظروف الغموض وعدم اليقين — والنموذج الملائم هو التدريب القائم على الكفايات."),
        ("سادساً", "إطار قابل للتحقيق", "ست كفايات وثماني وحدات وخارطة ثلاثية المراحل، يمكن تحويله إلى برنامج تحويلي مُدمج في المسار المهني القائم."),
        ("سابعاً", "النجاح مشروط مؤسسياً", "قيادة تعتبر التدريب ركناً من أركان السلامة، وآلية دورية لتحويل الأحداث والتعديلات إلى تحديثات تدريبية، وشراكة مع المشغّلين."),
    ]
    col_w = (CW - 0.34) / 2
    pitch = 1.13
    for i, (num, t, d) in enumerate(results):
        col = 0 if i < 4 else 1
        row = i if i < 4 else i - 4
        x = SW - M - col_w - col * (col_w + 0.34)
        yy = y + 0.02 + row * pitch
        card(s, x, yy, col_w, 1.04, fill=CARD_2, line="1D3149", accent=GOLD, accent_w=0.045)
        tb = textbox(s, x + 0.22, yy + 0.10, col_w - 0.46, 0.24)
        write(tb.text_frame, [(num, {"color": GOLD, "size": 8.8, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.22, yy + 0.32, col_w - 0.46, 0.26)
        write(tb.text_frame, [(t, {"color": INK, "size": 10.8, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.22, yy + 0.58, col_w - 0.46, 0.54)
        write(tb.text_frame, [(d, {"color": MUTED, "size": 8.6})], first=True,
              space_after=0, line_spacing=1.24)

    tb = textbox(s, SW - M - 11.2, 6.62, 11.2, 0.3)
    write(tb.text_frame, [("المحصلة: الفجوة ليست في المعرفة النظرية، بل في تحويل الكفايات إلى سلوك ملاحَظ ومعيار قياس وزمن استجابة.",
                           {"color": GOLD_SOFT, "size": 9.5})], first=True, space_after=0)
    return s


# ---------------------------------------------------------------- 28 · recommendations
def s28(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "التوصيات", kicker="مخرج تنفيذي",
                sub="خمس عشرة توصية موزّعة على ثلاثة مسارات: تنظيمي، تدريبي، تشغيلي")

    groups = [
        ("المسار التنظيمي والمؤسسي", GOLD, [
            "إصدار متطلبات تدريبية إلزامية صريحة لكوادر المراقبة الجوية بحدّ أدنى من الساعات والوحدات.",
            "تعريف مسؤوليات المراقب وحدود تدخله في حالات الطائرات غير الممتثلة.",
            "إنشاء سجل وطني وقاعدة بيانات للأحداث والاقتحامات أساساً لتحليل السلامة وتصميم السيناريوهات.",
            "اعتماد نموذج اللوائح الدولية أو ما يوازيه وطنياً لضمان الاتساق وتبادل الممارسات.",
            "تأسيس لجنة دائمة لتدريب الطائرات غير المأهولة مسؤولة عن النموذج المحلي ومراجعته.",
        ]),
        ("المسار التدريبي والتقييمي", CYAN, [
            "تطوير قواعد الكفايات والاختبارات التأهيلية مع تقييم دوري مستمر لا لمرة واحدة.",
            "إدخال وحدات إلزامية في التدريب الأولي، وتدريب تحويلي للمراقبين الحاليين بمعدل زمني محدد.",
            "توفير سيناريوهات محاكاة معتمدة تغطي الحالات النادرة عالية الخطورة بتدريب متكرر.",
            "تدريب مشترك بين المراقبين والطيارين المسيرين ومشغلي UTM بحدٍّ أدنى دوري سنوي.",
            "تطوير تدريب المدربين شرطاً مسبقاً لأي توسّع في البرامج.",
        ]),
        ("المسار التشغيلي والموارد", "B98A5E", [
            "إجراءات تشغيلية قياسية مكتوبة (SOPs) للاقتحامات والتصعيد والتنسيق الأمني، والتدريب عليها مشتركاً.",
            "قياس عبء العمل والوعي الظرفي بانتظام واستخدام النتائج في ضبط أحجام الحركة.",
            "حُزَم تدريب نقالة ووحدات محاكاة خفيفة للوصول إلى الوحدات النائية.",
            "شراكات أكاديمية لإنتاج مواد محكمة ودراسات قياس أثر بالتعاون مع مقدمي الخدمة.",
            "إدماج مبادئ العوامل البشرية في كل وحدة تدريبية لا في وحدة منفصلة فقط.",
        ]),
    ]
    gap = 0.26
    cw = (CW - 2 * gap) / 3
    for i, (title, col, items) in enumerate(groups):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.02, cw, 4.42, fill=CARD, line=LINE, accent=col)
        tb = textbox(s, x + 0.24, y + 0.2, cw - 0.52, 0.3)
        write(tb.text_frame, [(title, {"color": col, "size": 11.5, "bold": True})],
              first=True, space_after=0)
        hline(s, x + 0.24, y + 0.58, cw - 0.52, color="2A4058", weight=0.75)
        tb = textbox(s, x + 0.24, y + 0.72, cw - 0.52, 3.5)
        for k, it in enumerate(items):
            write(tb.text_frame, [(it, {"color": BODY, "size": 9.2})],
                  first=(k == 0), space_after=9, line_spacing=1.32, numbered=True)
    return s


# ---------------------------------------------------------------- 29 · closing
def s29(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, "0A1626", "050B14")
    pic = picture_cover(s, A + "cover.jpg", 0, 0, SW, SH)
    veil = rect(s, 0, 0, SW, SH, fill="050C16", shape=MSO_SHAPE.RECTANGLE)
    set_alpha(veil, 0.74)
    frame = rect(s, 0.42, 0.62, SW - 0.84, 6.32, fill=None, line="7C6A3E", line_w=0.75)
    set_line_alpha(frame, 0.5)

    tb = textbox(s, SW - M - 9.0, 2.28, 9.0, 0.34)
    write(tb.text_frame, [("خاتمة", {"color": GOLD, "size": 12, "bold": True})],
          first=True, space_after=0)
    tb = textbox(s, SW - M - 9.6, 2.68, 9.6, 1.7)
    write(tb.text_frame, [
        ("الاستعداد للبيئة المختلطة", {"color": INK, "size": 36, "bold": True}),
        ("ليس خياراً تدريبياً، بل شرط سلامة", {"color": INK, "size": 26, "bold": True}),
    ], first=True, space_after=6, line_spacing=1.16)
    hline(s, SW - M - 2.0, 4.42, 2.0, color=GOLD, weight=2.0)
    tb = textbox(s, SW - M - 8.6, 4.60, 8.6, 0.9)
    write(tb.text_frame, [("ثلاثة عوامل مؤسسية تحكم النجاح: قيادة تعتبر التدريب ركناً من أركان السلامة، وآلية دورية تحوّل الأحداث والتعديلات إلى تحديثات تدريبية، وشراكة فاعلة مع مشغّلي الأنظمة غير المأهولة تُبنى على العمليات الحقيقية لا على الافتراضات.",
                           {"color": BODY, "size": 11})], first=True, space_after=0,
          line_spacing=1.4)

    tb = textbox(s, SW - M - 9.0, 6.02, 9.0, 0.34)
    write(tb.text_frame, [("والله ولي التوفيق", {"color": GOLD_SOFT, "size": 15, "bold": True})],
          first=True, space_after=0)

    names = ["مقدم مراقب جوي / أحمد عبد الرحمن أحمد عدلان",
             "مقدم مراقب جوي / معمر خليفة محمد سعيد",
             "مقدم طيار / عامر محمد علي محمد فضل"]
    tb = textbox(s, SW - M - 10.0, 6.62, 10.0, 0.34)
    write(tb.text_frame, ["   ·   ".join(names)], size=9, color=MUTED, first=True, space_after=0)

    chrome(s, n, total)
    return s
