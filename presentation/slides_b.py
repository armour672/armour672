# -*- coding: utf-8 -*-
"""Slides 12-17: risks, operational differences, human factors, comms, lessons, needs matrix."""
from deckkit import *
from tablekit import table
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from slides_a import A


# ---------------------------------------------------------------- 12 · risks
def s12(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "خمس فئات للمخاطر التشغيلية", kicker="تحليل المخاطر",
                sub="لكل فئة سيناريوها التدريبي المقابل، ومعيار أداء قابل للملاحظة، وزمن استجابة مستهدف")

    risks = [
        ("التصادم الجوي", "مع طائرة مأهولة أو أخرى غير مأهولة، وتتضاعف قرب مسارات الاقتراب والمغادرة وخارج مدى الرؤية"),
        ("اقتحام المجال المحظور", "مناطق حظر الطيران حول المطارات والفعاليات الكبرى وعمليات مكافحة الحرائق — أعقد الحالات لأن مصدرها مشغّل غير ممتثل"),
        ("فقدان وصلة القيادة والتحكم", "الأكثر ارتباطاً بالمهمة التدريبية، ويستلزم افتراض دخول الطائرة برنامج طوارئ تلقائي"),
        ("فقدان السيطرة أو الإنهاء القسري", "هروب الطائرة أو إنهاء الرحلة قسرياً، وقد ينتهي بسقوط حطام يهدد السلامة على الأرض"),
        ("المخاطر الأمنية", "حادثة غاتويك جسّدت العواقب التشغيلية الهائلة للتهديد غير المتماثل"),
    ]
    gap = 0.2
    cw = (CW - 4 * gap) / 5
    for i, (t, d) in enumerate(risks):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.06, cw, 3.05, fill=CARD_2, line="1D3149",
             accent=GOLD if i != 2 else CYAN, accent_w=0.045)
        tb = textbox(s, x + 0.2, y + 0.24, cw - 0.4, 0.3)
        write(tb.text_frame, [(ar_num(i + 1), {"color": GOLD, "size": 16, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.2, y + 0.62, cw - 0.4, 0.62)
        write(tb.text_frame, [(t, {"color": INK, "size": 11.5, "bold": True})],
              first=True, space_after=0, line_spacing=1.2)
        hline(s, x + 0.2, y + 1.32, cw - 0.4, color="2A4058", weight=0.75)
        tb = textbox(s, x + 0.2, y + 1.46, cw - 0.4, 1.5)
        write(tb.text_frame, [(d, {"color": MUTED, "size": 9})], first=True,
              space_after=0, line_spacing=1.3)

    band = rect(s, M, 5.62, CW, 1.22, fill="101F33", line="2A4258", radius=0.05)
    vline(s, SW - M - 0.16, 5.72, 1.02, color=GOLD, weight=3.0)
    tb = textbox(s, M + 0.55, 5.74, CW - 1.2, 0.4)
    write(tb.text_frame, [("الإشكال المعرفي الأهم في السياق التدريبي", {"color": GOLD_SOFT,
                                                                        "size": 11, "bold": True})],
          first=True, space_after=2)
    tb = textbox(s, M + 0.55, 6.10, CW - 1.2, 0.66)
    write(tb.text_frame, [("التقييم السريع لجدية التهديد عند ورود بلاغ بصري غير مؤكد قرارٌ بشري مركّب يتراوح بين التقليل من تهديد حقيقي والإفراط في رد الفعل وما يرتب عليه من كلفة غير مبررة — ومن هنا ضرورة إدراج كفاية «تقييم البلاغات ودرجة اليقين» في البرامج التدريبية.",
                           {"color": BODY, "size": 10})], first=True, space_after=0,
          line_spacing=1.32)
    return s


# ---------------------------------------------------------------- 13 · differences
def s13(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "الفروق التشغيلية وأثرها على مهام المراقب", kicker="الأداء والتجهيز",
                sub="خمسة أوجه تمسّ بنية المهمة التدريبية، وتنتهي إلى مفهوم واحد: «التصور الوظيفي للقيد»")

    rows = [
        ("موقع الطيار", "في محطة بعيدة بدلاً من المتن", "تعذّر الاعتماد على التقارير البصرية"),
        ("القدرة على المناورة", "تتوقف على أداء الطائرة وزمن الكمون", "استلزام مسافات فاصلة أكبر"),
        ("الرؤية والتفادي", "قد تنعدم القناة الطبيعية خارج مدى الرؤية", "تنفرد الأنظمة التقنية بتكوين الصورة الجوية"),
        ("التجهيز الملاحي", "متفاوت بحسب الفئة والحجم", "ارتفاع خطر «الهدف غير المرئي»"),
        ("الامتثال للأوامر", "قد يخضع لقيود المشغّل والمنظومة", "التحقق من قابلية التنفيذ قبل إصدار التعليم"),
    ]
    col_w = [2.45, 4.24, 4.94]          # from right to left
    xs = []
    cursor = SW - M
    for w in col_w:
        cursor -= w
        xs.append(cursor)
    hy = y + 0.02
    for x, w, t in zip(xs, col_w, ["البُعد", "الفرق التشغيلي", "الأثر على مهمة المراقب"]):
        tb = textbox(s, x + 0.16, hy, w - 0.3, 0.28)
        write(tb.text_frame, [(t, {"color": GOLD_SOFT, "size": 9.5, "bold": True})],
              first=True, space_after=0)
    hline(s, M, hy + 0.30, CW, color=GOLD, weight=1.0)

    ry = hy + 0.44
    rh = 0.72
    for i, (a, b, c) in enumerate(rows):
        if i % 2 == 0:
            rect(s, M, ry - 0.04, CW, rh - 0.02, fill="0E1D30", radius=0.02)
        tb = textbox(s, xs[0] + 0.16, ry + 0.10, col_w[0] - 0.3, rh - 0.24)
        write(tb.text_frame, [(a, {"color": INK, "size": 10.5, "bold": True})],
              first=True, space_after=0, line_spacing=1.22)
        tb = textbox(s, xs[1] + 0.10, ry + 0.10, col_w[1] - 0.22, rh - 0.24)
        write(tb.text_frame, [(b, {"color": BODY, "size": 10})], first=True, space_after=0,
              line_spacing=1.24)
        tb = textbox(s, xs[2] + 0.10, ry + 0.10, col_w[2] - 0.2, rh - 0.24)
        write(tb.text_frame, [(c, {"color": CYAN, "size": 10})], first=True, space_after=0,
              line_spacing=1.24)
        ry += rh

    band = rect(s, M, 6.02, CW, 0.86, fill="132741", line="2C4763", radius=0.05)
    tb = textbox(s, M + 0.5, 6.10, CW - 1.0, 0.72, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("لا يعني ذلك أن يتحول المراقب إلى مهندس تقني، بل أن يمتلك «تصوراً وظيفياً للقيود»: القدرة على استدعاء قيود الأداء الجوهرية عند لحظة القرار، واستخدامها في تقدير زمن المناورة ومسافات الفصل واختيار نوع التعليمات.",
                           {"color": GOLD_SOFT, "size": 10.5})], first=True, space_after=0,
          line_spacing=1.32)
    return s


# ---------------------------------------------------------------- 14 · human factors
def s14(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "العوامل البشرية في بيئة مختلطة", kicker="العوامل البشرية",
                sub="الوعي الظرفي وعبء العمل واليقظة: ثلاث بوابات يمرّ منها قرار المراقب")

    items = [
        ("الوعي الظرفي", "إدراك العناصر وفهم معناها وتوقع حالتها المستقبلية", [
            "الإدراك: دمج مصادر جديدة — UTM وبلاغات المشغّل ومؤشرات الوصلة.",
            "الفهم: دلالة فقدان الاتصال والطيران الذاتي.",
            "التنبؤ: تقدير سلوك الطائرة وفق برنامج طوارئ قد لا يعرفه المراقب.",
        ], "Endsley, 1995"),
        ("عبء العمل", "النتيجة مشروطة بظروف التجربة ولا تصحّ تعميمها على البيئات العالية الكثافة", [
            "إدخال طائرة في حالة طارئة داخل حيّز مشترك لم يتجاوز الحدود المقبولة من إدراك المراقب وعبء عمله.",
            "إعلان النوايا حسّن الوعي الظرفي وكفاءة الإدارة.",
            "قياس عبء العمل بأدوات معيارية (NASA-TLX) ضروري.",
        ], "Reyes-Muñoz et al., 2023"),
        ("اليقظة والملل", "المراقبة السلبية للبيانات الآلية ترفع احتمال تدهور اليقظة في فترات الانحسار", [
            "أثر حجم المهام في الوعي الظرفي والاستراتيجيات التعويضية.",
            "الحل التدريبي: صمت تشغيلي يعقبه انفجار مفاجئ في الحركة.",
        ], "Friedrich et al., 2019"),
    ]
    gap = 0.30
    cw = (CW - 2 * gap) / 3
    for i, (t, sub, bullets, ref) in enumerate(items):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.06, cw, 3.62, fill=CARD, line=LINE,
             accent=GOLD if i == 0 else (CYAN if i == 1 else "B98A5E"))
        tb = textbox(s, x + 0.26, y + 0.24, cw - 0.62, 0.34)
        write(tb.text_frame, [(t, {"color": INK, "size": 14.5, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.26, y + 0.62, cw - 0.62, 0.62)
        write(tb.text_frame, [(sub, {"color": MUTED, "size": 8.8})], first=True,
              space_after=0, line_spacing=1.28)
        hline(s, x + 0.26, y + 1.28, cw - 0.62, color="2A4058", weight=0.75)
        tb = textbox(s, x + 0.26, y + 1.40, cw - 0.62, 1.9)
        write(tb.text_frame, [(b, {"color": BODY, "size": 9.2}) for b in bullets],
              first=True, space_after=7, line_spacing=1.3, bullet="•")
        tb = textbox(s, x + 0.26, y + 3.18, cw - 0.62, 0.28)
        write(tb.text_frame, [(ref, {"color": GOLD, "size": 8, "bold": True})],
              first=True, space_after=0)

    band = rect(s, M, 5.96, CW, 0.72, fill="132741", line="2C4763", radius=0.05)
    vline(s, SW - M - 0.16, 6.04, 0.56, color=GOLD, weight=3.0)
    tb = textbox(s, M + 0.5, 6.02, CW - 1.2, 0.6, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("الأثر التدريبي: كل سيناريو يجب أن يقيس عبء العمل والوعي الظرفي صراحةً، لا أن يكتفي بالحكم على النتيجة النهائية للأداء.",
                           {"color": GOLD_SOFT, "size": 10.2})], first=True, space_after=0,
          line_spacing=1.3)
    return s


# ---------------------------------------------------------------- 15 · comms
def s15(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "الاتصالات وزمن الاستجابة", kicker="العبارات اللاسلكية",
                sub="التشغيل المشترك يستلزم إتقان العبارات المنظمة مع طرف بعيد لا يرى الطائرة")

    items = [
        ("العبارات المنظمة والتأكيد", "إتقان العبارات مع طرف بعيد والتأكد من آلية تأكيد التعليمات وفق دليل العبارات اللاسلكية (Doc 9432)"),
        ("زمن الاستجابة", "مراعاة زمن الاستجابة في الاتصال والتعليمات، وضبط التوقعات والمسافات الفاصلة على أساسه"),
        ("التعرف والتسجيل الأولي", "الهوية والموقع والارتفاع ونوع العملية ومدى الرؤية، وتعليمات ردّ الاستجواب"),
        ("فقدان الاتصال مع الطيار البعيد", "افتراض احتمال دخول الطائرة برنامج طوارئ تلقائي، وتخطيط صريح لإدارة الموقف لا افتراض أن الحالة وصلت إلى حل"),
    ]
    lw = CW - 4.55
    tb = textbox(s, SW - M - lw, y + 0.06, lw, 3.9)
    for i, (t, d) in enumerate(items):
        write(tb.text_frame, [
            (t, {"color": INK, "size": 11.5, "bold": True, "space_after": 3}),
            (d, {"color": BODY, "size": 9.6, "space_after": 17}),
        ], first=(i == 0) and False, line_spacing=1.32)

    # example exchange panel (left)
    px = M
    pw = 4.15
    py = y + 0.02
    card(s, px, py, pw, 3.72, fill="0E1D30", line="1D3149")
    tb = textbox(s, px + 0.26, py + 0.18, pw - 0.5, 0.3)
    write(tb.text_frame, [("نموذج تبادل لاسلكي — حالة فقدان وصلة", {"color": GOLD_SOFT,
                                                                     "size": 9.5, "bold": True})],
          first=True, space_after=0)
    exch = [
        ("المراقب", "«ألفا ٢٧، البرج، فقدت مؤشرات الوصلة، أكّد الحالة»", 0.62),
        ("الطيار البعيد", "«البرج، ألفا ٢٧، فقدت القيادة، الطائرة في برنامج العودة التلقائي»", 1.14),
        ("المراقب", "«ألفا ٢٧، البرج، روجر، احتفظ بالارتفاع ١٢٠٠ قدم، أبلغ عن أي تغيير»", 2.02),
    ]
    for who, line, dy in exch:
        tb = textbox(s, px + 0.26, py + dy, pw - 0.5, 0.22)
        write(tb.text_frame, [(who, {"color": CYAN, "size": 8.2, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, px + 0.26, py + dy + 0.22, pw - 0.5, 0.66)
        write(tb.text_frame, [(line, {"color": INK, "size": 9.4})], first=True,
              space_after=0, line_spacing=1.28)
    tb = textbox(s, px + 0.26, py + 3.28, pw - 0.5, 0.28)
    write(tb.text_frame, [("نموذج توضيحي للتدريب — لا يُغني عن العبارات المعتمدة",
                           {"color": MUTED, "size": 7.6})], first=True, space_after=0)

    band = rect(s, M, 6.22, CW, 0.64, fill="101F33", line="2A4258", radius=0.05)
    tb = textbox(s, M + 0.45, 6.28, CW - 0.9, 0.52, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("قاعدة تدريبية: كل تعليم يُصدر لطيار مسير يجب أن يُصمَّم على افتراض أنه قد لا يُنفَّذ فوراً أو قد لا يُنفَّذ أصلاً — فالتعليم الجيد يحتمل التأخير والرفض والتفسير الخاطئ.",
                           {"color": GOLD_SOFT, "size": 10})], first=True, space_after=0,
          line_spacing=1.3)
    return s


# ---------------------------------------------------------------- 16 · lessons
def s16(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "الدروس المستخلصة من التجارب", kicker="الأدبيات",
                sub="ثلاثة دروس تشغيلية ودرس قانوني، تتقارب جميعها نحو نموذج تدريبي مختلط")

    items = [
        ("٠١", "المحاكاة القائمة على السيناريو", "فعالية التدريب بالمحاكاة وفق منهجية «السيناريو – الحل – التقييم» مع اعتماد عبء العمل والكفاءة مؤشّرين للتقييم.", "Pinto Neto et al., 2024"),
        ("٠٢", "الطيار الافتراضي بالذكاء الاصطناعي", "توليد السيناريوهات النادرة وتقديم تغذية راجعة فورية قابلة للقياس، مع حفاظ على حوكمة واضحة للأداة وحدود استخدامها.", "Zuluaga-Gomez et al., 2023 · ICAO APAC, 2025"),
        ("٠٣", "حدود المحاكاة وحدها", "تسرّع الإلمام بالمهارات لكنها لا تُغني عن التدريب الميداني في المهارات عالية الترابط، ما يرجّح نموذجاً مختلطاً.", "خلاصة تحليلية"),
        ("٠٤", "الدرس القانوني", "وجوب فهم المراقب حدود مسؤوليته ومبدأ التكافؤ وعدم مسؤوليته عن حوادث الأنظمة غير الممتثلة.", "IFATCA, 2024"),
    ]
    gap = 0.26
    cw = (CW - 3 * gap) / 4
    for i, (num, t, d, ref) in enumerate(items):
        x = SW - M - cw - i * (cw + gap)
        card(s, x, y + 0.08, cw, 3.1, fill=CARD, line=LINE, accent=GOLD, accent_w=0.045)
        tb = textbox(s, x + 0.24, y + 0.26, cw - 0.5, 0.4)
        write(tb.text_frame, [(num, {"color": GOLD, "size": 20, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.24, y + 0.72, cw - 0.5, 0.62)
        write(tb.text_frame, [(t, {"color": INK, "size": 11.5, "bold": True})],
              first=True, space_after=0, line_spacing=1.22)
        hline(s, x + 0.24, y + 1.42, cw - 0.5, color="2A4058", weight=0.75)
        tb = textbox(s, x + 0.24, y + 1.56, cw - 0.5, 1.2)
        write(tb.text_frame, [(d, {"color": BODY, "size": 9.5})], first=True,
              space_after=0, line_spacing=1.32)
        tb = textbox(s, x + 0.24, y + 2.76, cw - 0.5, 0.26)
        write(tb.text_frame, [(ref, {"color": MUTED, "size": 7.8})], first=True, space_after=0)

    # convergence band
    by = 6.12
    arrow = rect(s, M, by, CW, 0.74, fill="132741", line="2C4763", radius=0.05)
    tb = textbox(s, M + 0.5, by + 0.14, CW - 1.0, 0.5, anchor=MSO_ANCHOR.MIDDLE)
    write(tb.text_frame, [("المحصلة:  نموذج مختلط (محاكاة + طيار افتراضي + تدريب ميداني مشترك) محكوم بمعايير أداء قابلة للملاحظة",
                           {"color": GOLD_SOFT, "size": 11.5, "bold": True})], first=True,
          space_after=0)
    return s


# ---------------------------------------------------------------- 17 · needs matrix
def s17(prs, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s)
    chrome(s, n, total)
    y = heading(s, "مصفوفة الاحتياجات التدريبية المستخلصة", kicker="جسر الانتقال",
                sub="ستة مجالات احتياج تتحول إلى كفايات محددة، ثم إلى وحدات تدريبية ومعايير تقييم")

    domains = [
        ("أساسيات الأنظمة غير المأهولة", "المصطلحات والتصنيفات وقدرات الأداء"),
        ("الوعي الظرفي الممتد", "دمج مصادر البيانات المتعددة"),
        ("إدارة التعارض والفصل", "في ضوء قيود الأداء وزمن الاستجابة"),
        ("إدارة الطوارئ والاقتحام", "والتنسيق المؤسسي مع الجهات الخارجية"),
        ("الاتصالات والعبارات اللاسلكية", "مع الطيار المسير والطائرات المتأثرة"),
        ("العوامل البشرية", "والإبلاغ الدقيق وتقييم البلاغات"),
    ]
    gap_x, gap_y = 0.26, 0.18
    cw = (CW - 2 * gap_x) / 3
    ch = 0.86
    for i, (t, d) in enumerate(domains):
        col, row = i % 3, i // 3
        x = SW - M - cw - col * (cw + gap_x)
        yy = y + 0.04 + row * (ch + gap_y)
        card(s, x, yy, cw, ch, fill=CARD_2, line="1D3149", accent=CYAN, accent_w=0.045)
        tb = textbox(s, x + 0.22, yy + 0.14, cw - 0.44, 0.28)
        write(tb.text_frame, [(t, {"color": INK, "size": 10.8, "bold": True})],
              first=True, space_after=0)
        tb = textbox(s, x + 0.22, yy + 0.44, cw - 0.44, 0.34)
        write(tb.text_frame, [(d, {"color": MUTED, "size": 9})], first=True, space_after=0)

    # flow: needs -> competencies -> units
    fy = y + 2.66
    steps = [("٦", "مجالات احتياج"), ("٦", "كفايات محددة"), ("٨", "وحدات تدريبية"), ("٣", "مستويات تقييم")]
    gap = 0.34
    bw = (CW - 3 * gap) / 4
    for i, (num, label) in enumerate(steps):
        x = SW - M - bw - i * (bw + gap)
        rect(s, x, fy, bw, 0.92, fill="132741", line="2C4763", radius=0.06)
        tb = textbox(s, x + 0.2, fy + 0.08, bw - 0.4, 0.5, anchor=MSO_ANCHOR.MIDDLE)
        write(tb.text_frame, [(num + "  ", {"color": GOLD, "size": 17, "bold": True}),
                              (label, {"color": INK, "size": 10, "bold": True})],
              first=True, space_after=0, line_spacing=1.1)
        tb = textbox(s, x + 0.2, fy + 0.58, bw - 0.4, 0.26)
        write(tb.text_frame, [("مخرَج المبحث", {"color": MUTED, "size": 8})], first=True,
              space_after=0)
        if i < 3:
            # left-pointing arrow between boxes (RTL flow goes leftwards)
            from pptx.enum.text import PP_ALIGN as PA
            tb = textbox(s, x - gap, fy + 0.24, gap, 0.4)
            write(tb.text_frame, [("◀", {"color": GOLD, "size": 11, "align": PA.CENTER})],
                  first=True, space_after=0)

    tb = textbox(s, SW - M - 9.6, 5.98, 9.6, 0.5)
    write(tb.text_frame, [("هذه المجالات الستة هي مخرَج المبحث الثاني ومدخَل المبحث الثالث: كل مجال يتحول إلى كفاية باسمها ومعيار أدائها القابل للقياس.",
                           {"color": BODY, "size": 10.5})], first=True, space_after=0,
          line_spacing=1.34)
    return s
