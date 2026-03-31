def build_slide(slide):
    # Colors
    DARK_BLUE = RGBColor(0x1B, 0x36, 0x5D)
    DARK_GRAY = RGBColor(0x55, 0x55, 0x55)
    DARK_RED = RGBColor(0x9E, 0x2A, 0x2B)
    LIGHT_RED = RGBColor(0xC1, 0x3C, 0x3D)
    GRAY = RGBColor(0xB0, 0xB0, 0xB0)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    BLACK = RGBColor(0x00, 0x00, 0x00)
    BG_GRAY = RGBColor(0xEB, 0xEF, 0xF2)

    # 1. Slide Background
    slide_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    slide_bg.fill.solid()
    slide_bg.fill.fore_color.rgb = BG_GRAY
    slide_bg.line.fill.background()

    # 2. Header Banner
    add_header_banner(slide, "SAP ABAP AI 效率测评报告", bg_color=DARK_BLUE)

    # 3. Main White Card
    bg_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.8), Inches(12.533), Inches(6.4))
    bg_card.fill.solid()
    bg_card.fill.fore_color.rgb = WHITE
    bg_card.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)

    # 4. Main Title & Subtitle
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(10), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "第2页 中等复杂度：深陷“虚构字段”泥潭"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE

    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(10), Inches(0.5))
    tf_sub = sub_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "案例B - 采购配额维护（中等复杂度）验证"
    p_sub.font.size = Pt(20)
    p_sub.font.bold = True
    p_sub.font.color.rgb = DARK_GRAY

    # 5. Section 1: 业务理解偏差
    icon1 = slide.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(0.8), Inches(0.8))
    icon1.text_frame.text = "🧠❌"
    icon1.text_frame.paragraphs[0].font.size = Pt(32)

    t1 = slide.shapes.add_textbox(Inches(1.8), Inches(2.3), Inches(4), Inches(0.4))
    t1.text_frame.text = "业务理解偏差"
    t1.text_frame.paragraphs[0].font.size = Pt(20)
    t1.text_frame.paragraphs[0].font.bold = True

    b1 = slide.shapes.add_textbox(Inches(1.8), Inches(2.7), Inches(4.5), Inches(1.0))
    tf1 = b1.text_frame
    tf1.word_wrap = True
    p1_1 = tf1.paragraphs[0]
    p1_1.text = "• Copilot 完全混淆“配额”与“货源”概念。"
    p1_1.font.size = Pt(14)
    p1_2 = tf1.add_paragraph()
    p1_2.text = "• 数据模型从底层开始错误，导致逻辑无法构建。"
    p1_2.font.size = Pt(14)

    # 6. Section 2: 严重的“幻觉”现象
    icon2 = slide.shapes.add_textbox(Inches(0.8), Inches(3.8), Inches(0.8), Inches(0.8))
    icon2.text_frame.text = "🔗💥"
    icon2.text_frame.paragraphs[0].font.size = Pt(32)

    t2 = slide.shapes.add_textbox(Inches(1.8), Inches(3.7), Inches(4), Inches(0.4))
    t2.text_frame.text = "严重的“幻觉”现象"
    t2.text_frame.paragraphs[0].font.size = Pt(20)
    t2.text_frame.paragraphs[0].font.bold = True

    lbl_real = slide.shapes.add_textbox(Inches(1.8), Inches(4.2), Inches(1.2), Inches(0.3))
    lbl_real.text_frame.text = "实际字段"
    lbl_real.text_frame.paragraphs[0].font.size = Pt(12)

    lbl_fake = slide.shapes.add_textbox(Inches(3.0), Inches(4.2), Inches(1.2), Inches(0.3))
    lbl_fake.text_frame.text = "虚构字段"
    lbl_fake.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_fake.text_frame.paragraphs[0].font.color.rgb = DARK_RED

    box_real = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.8), Inches(4.6), Inches(1.2), Inches(1.0))
    box_real.fill.solid()
    box_real.fill.fore_color.rgb = GRAY
    box_real.line.fill.background()

    box_fake = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.0), Inches(4.6), Inches(2.0), Inches(1.0))
    box_fake.fill.solid()
    box_fake.fill.fore_color.rgb = DARK_RED
    box_fake.line.fill.background()
    tf_fake = box_fake.text_frame
    tf_fake.text = "9处关键字段\n(占比 50%)"
    tf_fake.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_fake.paragraphs[0].font.size = Pt(14)
    tf_fake.paragraphs[0].font.color.rgb = WHITE
    if len(tf_fake.paragraphs) > 1:
        tf_fake.paragraphs[1].alignment = PP_ALIGN.CENTER
        tf_fake.paragraphs[1].font.size = Pt(14)
        tf_fake.paragraphs[1].font.color.rgb = WHITE

    arrow1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.1), Inches(4.7), Inches(1.2), Inches(0.8))
    arrow1.fill.solid()
    arrow1.fill.fore_color.rgb = LIGHT_RED
    arrow1.line.fill.background()
    tf_arr1 = arrow1.text_frame
    tf_arr1.text = "连锁关键字段\n(21个)"
    tf_arr1.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_arr1.paragraphs[0].font.size = Pt(11)
    tf_arr1.paragraphs[0].font.color.rgb = WHITE
    if len(tf_arr1.paragraphs) > 1:
        tf_arr1.paragraphs[1].alignment = PP_ALIGN.CENTER
        tf_arr1.paragraphs[1].font.size = Pt(11)
        tf_arr1.paragraphs[1].font.color.rgb = WHITE

    # 7. Section 3: 接口规范缺失
    icon3 = slide.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(0.8), Inches(0.8))
    icon3.text_frame.text = "🔌❌"
    icon3.text_frame.paragraphs[0].font.size = Pt(32)

    t3 = slide.shapes.add_textbox(Inches(1.8), Inches(5.7), Inches(4), Inches(0.4))
    t3.text_frame.text = "接口规范缺失"
    t3.text_frame.paragraphs[0].font.size = Pt(20)
    t3.text_frame.paragraphs[0].font.bold = True

    b3 = slide.shapes.add_textbox(Inches(1.8), Inches(6.1), Inches(4.5), Inches(1.0))
    tf3 = b3.text_frame
    tf3.word_wrap = True
    p3_1 = tf3.paragraphs[0]
    p3_1.text = "• AI 无法识别 SAP 函数模块 (FM) 仅接受 DDIC 类型的硬性规则。"
    p3_1.font.size = Pt(14)
    p3_2 = tf3.add_paragraph()
    p3_2.text = "• 数据类型不匹配导致接口调用必然失败。"
    p3_2.font.size = Pt(14)

    # 8. Section 4: 开发成本倒挂
    icon4 = slide.shapes.add_textbox(Inches(6.2), Inches(5.8), Inches(0.8), Inches(0.8))
    icon4.text_frame.text = "⚖️❌"
    icon4.text_frame.paragraphs[0].font.size = Pt(32)

    t4 = slide.shapes.add_textbox(Inches(7.2), Inches(5.7), Inches(4), Inches(0.4))
    t4.text_frame.text = "开发成本倒挂"
    t4.text_frame.paragraphs[0].font.size = Pt(20)
    t4.text_frame.paragraphs[0].font.bold = True

    b4 = slide.shapes.add_textbox(Inches(7.2), Inches(6.1), Inches(5.0), Inches(1.0))
    tf4 = b4.text_frame
    tf4.word_wrap = True
    p4_1 = tf4.paragraphs[0]
    p4_1.text = "• 修复 AI 错误的代码成本已超过直接重写。"
    p4_1.font.size = Pt(14)
    p4_2 = tf4.add_paragraph()
    p4_2.text = "• AI 辅助在该场景下失去价值，带来负面效率。"
    p4_2.font.size = Pt(14)

    # 9. Large Diagram
    diag_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.4), Inches(2.2), Inches(6.3), Inches(3.3))
    diag_bg.fill.solid()
    diag_bg.fill.fore_color.rgb = RGBColor(0xF9, 0xF9, 0xF9)
    diag_bg.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)

    l_real = slide.shapes.add_textbox(Inches(6.5), Inches(2.3), Inches(1.3), Inches(0.3))
    l_real.text_frame.text = "实际字段"
    l_real.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    l_real.text_frame.paragraphs[0].font.size = Pt(12)

    l_fake = slide.shapes.add_textbox(Inches(7.85), Inches(2.3), Inches(1.8), Inches(0.3))
    l_fake.text_frame.text = "虚构字段"
    l_fake.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    l_fake.text_frame.paragraphs[0].font.size = Pt(12)
    l_fake.text_frame.paragraphs[0].font.color.rgb = DARK_RED

    l_err = slide.shapes.add_textbox(Inches(10.0), Inches(2.3), Inches(2.5), Inches(0.3))
    l_err.text_frame.text = "连锁语法错误"
    l_err.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    l_err.text_frame.paragraphs[0].font.size = Pt(12)
    l_err.text_frame.paragraphs[0].font.color.rgb = DARK_RED

    # Gray Stack
    gray_top = 2.7
    for i in range(4):
        gb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(gray_top + i*0.55), Inches(1.3), Inches(0.5))
        gb.fill.solid()
        gb.fill.fore_color.rgb = GRAY
        gb.line.color.rgb = WHITE
        gb.line.width = Pt(1)

    # Red Stack
    red_texts = ["虚构: QUOTA_MATNR", "虚构: SOURCE_VENDOR", "虚构: VALID_DATE_FROM"]
    red_top = 2.7
    for i, text in enumerate(red_texts):
        rb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.85), Inches(red_top + i*0.733), Inches(1.8), Inches(0.68))
        rb.fill.solid()
        rb.fill.fore_color.rgb = DARK_RED
        rb.line.color.rgb = WHITE
        rb.line.width = Pt(1)
        tf = rb.text_frame
        tf.text = text
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].font.size = Pt(10)
        tf.paragraphs[0].font.color.rgb = WHITE

    # Arrow
    arrow2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.7), Inches(3.6), Inches(0.25), Inches(0.4))
    arrow2.fill.solid()
    arrow2.fill.fore_color.rgb = LIGHT_RED
    arrow2.line.fill.background()

    # Red Grid
    err_left = 10.0
    err_top = 2.7
    err_w = 2.5
    err_h = 2.2
    
    base_red = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(err_left), Inches(err_top), Inches(err_w), Inches(err_h))
    base_red.fill.solid()
    base_red.fill.fore_color.rgb = LIGHT_RED
    base_red.line.fill.background()

    # Grid Lines
    for i in range(1, 4):
        y = err_top + i * (err_h / 4)
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left), Inches(y), Inches(err_left + err_w), Inches(y))
        line.line.color.rgb = WHITE
        line.line.width = Pt(1.5)
        
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.25), Inches(err_top), Inches(err_left + err_w*0.25), Inches(err_top + err_h/4)).line.color.rgb = WHITE
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.5), Inches(err_top), Inches(err_left + err_w*0.5), Inches(err_top + err_h/4)).line.color.rgb = WHITE
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.75), Inches(err_top), Inches(err_left + err_w*0.75), Inches(err_top + err_h/4)).line.color.rgb = WHITE
    
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.33), Inches(err_top + err_h/4), Inches(err_left + err_w*0.33), Inches(err_top + err_h/2)).line.color.rgb = WHITE
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.66), Inches(err_top + err_h/4), Inches(err_left + err_w*0.66), Inches(err_top + err_h/2)).line.color.rgb = WHITE
    
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.25), Inches(err_top + err_h/2), Inches(err_left + err_w*0.25), Inches(err_top + err_h*0.75)).line.color.rgb = WHITE
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.5), Inches(err_top + err_h/2), Inches(err_left + err_w*0.5), Inches(err_top + err_h*0.75)).line.color.rgb = WHITE
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.75), Inches(err_top + err_h/2), Inches(err_left + err_w*0.75), Inches(err_top + err_h*0.75)).line.color.rgb = WHITE
    
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.33), Inches(err_top + err_h*0.75), Inches(err_left + err_w*0.33), Inches(err_top + err_h)).line.color.rgb = WHITE
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(err_left + err_w*0.66), Inches(err_top + err_h*0.75), Inches(err_left + err_w*0.66), Inches(err_top + err_h)).line.color.rgb = WHITE

    # Center Block
    err_center = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.2), Inches(3.1), Inches(2.1), Inches(1.4))
    err_center.fill.solid()
    err_center.fill.fore_color.rgb = DARK_RED
    err_center.line.color.rgb = WHITE
    err_center.line.width = Pt(1.5)
    tf_err = err_center.text_frame
    tf_err.text = "连锁语法错误\n(21个)"
    tf_err.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_err.paragraphs[0].font.size = Pt(16)
    tf_err.paragraphs[0].font.bold = True
    tf_err.paragraphs[0].font.color.rgb = WHITE
    if len(tf_err.paragraphs) > 1:
        tf_err.paragraphs[1].alignment = PP_ALIGN.CENTER
        tf_err.paragraphs[1].font.size = Pt(16)
        tf_err.paragraphs[1].font.bold = True
        tf_err.paragraphs[1].font.color.rgb = WHITE

    # Bottom Text
    bot_text = slide.shapes.add_textbox(Inches(6.4), Inches(5.0), Inches(6.3), Inches(0.4))
    tf_bot = bot_text.text_frame
    tf_bot.text = "Claude Code 虚构了关键字段，引发 21 个连锁语法错误。"
    tf_bot.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_bot.paragraphs[0].font.size = Pt(14)
    tf_bot.paragraphs[0].font.color.rgb = BLACK