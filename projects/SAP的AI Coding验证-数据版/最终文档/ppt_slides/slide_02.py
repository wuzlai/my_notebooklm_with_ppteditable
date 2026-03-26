def build_slide(slide):
    # Define custom colors
    DARK_BLUE_BG = RGBColor(0x1D, 0x3B, 0x6C)
    TAG_BLUE = RGBColor(0x3A, 0x5A, 0x8C)
    LIGHT_BG = RGBColor(0xF8, 0xF9, 0xFA)
    BORDER_GRAY = RGBColor(0xE0, 0xE0, 0xE0)
    FUNNEL_DARK_BLUE = RGBColor(0x28, 0x52, 0x96)
    FUNNEL_MED_BLUE = RGBColor(0x5B, 0x9B, 0xD5)
    FUNNEL_LIGHT_BLUE = RGBColor(0x9D, 0xC3, 0xE6)
    FUNNEL_TEAL = RGBColor(0x45, 0x9E, 0x97)
    CONCL_BG = RGBColor(0xE0, 0xF2, 0xF1)
    CONCL_TEXT = RGBColor(0x00, 0x4D, 0x40)

    # 1. Header Section
    # Background
    header_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, Inches(1.2))
    header_bg.fill.solid()
    header_bg.fill.fore_color.rgb = DARK_BLUE_BG
    header_bg.line.fill.background()

    # "中等" Tag Box
    tag_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.2), Inches(0.8), Inches(0.4))
    tag_box.fill.solid()
    tag_box.fill.fore_color.rgb = TAG_BLUE
    tag_box.line.fill.background()
    tf_tag = tag_box.text_frame
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = "中等"
    p_tag.alignment = PP_ALIGN.CENTER
    p_tag.font.size = Pt(18)
    p_tag.font.color.rgb = WHITE
    p_tag.font.bold = True

    # Title
    title_box = slide.shapes.add_textbox(Inches(1.4), Inches(0.15), Inches(10), Inches(0.5))
    tf_title = title_box.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.text = "中等难度：数据字典“幻觉”导致代码不可用"
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE
    p_title.font.name = FONT_NAME

    # Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(10), Inches(0.4))
    tf_sub = sub_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "案例B —— 采购配额维护 (中等复杂度验证)"
    p_sub.font.size = Pt(14)
    p_sub.font.color.rgb = WHITE
    p_sub.font.name = FONT_NAME

    # 2. Left Column (Bullet Points)
    left_col_x = Inches(0.5)
    start_y = Inches(1.5)
    spacing = Inches(1.4)

    add_bullet_item(slide, left_col_x, start_y, "🧠", "业务理解偏差：", "Copilot 完全混淆采购配额与货源概念，导致核心数据模型错误。", width=Inches(6))
    add_bullet_item(slide, left_col_x, start_y + spacing, "⚠️", "虚构字段危机：", "Claude Code 虚构字段占比高达 50%，直接引发 21 个连锁语法错误。", width=Inches(6))
    
    # Small gauge text for item 2
    gauge_txt = slide.shapes.add_textbox(left_col_x + Inches(4.5), start_y + spacing + Inches(0.3), Inches(1.5), Inches(0.4))
    gauge_txt.text_frame.text = "50%\n虚构字段占比 ⚠️"
    gauge_txt.text_frame.paragraphs[0].font.size = Pt(10)
    gauge_txt.text_frame.paragraphs[0].font.color.rgb = RED
    gauge_txt.text_frame.paragraphs[0].font.bold = True
    gauge_txt.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    add_bullet_item(slide, left_col_x, start_y + spacing*2, "🔗", "接口规范缺失：", "AI 无法正确识别 SAP 函数模块 (FM) 接口仅接受 DDIC 类型的规则。", width=Inches(6))
    add_bullet_item(slide, left_col_x, start_y + spacing*3, "🧰", "修复成本极高：", "由于虚构字段与逻辑错误交织，AI 生成代码的修复成本远超重写成本。", width=Inches(6))

    # 3. Right Column - Top Box (Warning & Chart)
    box1_left = Inches(6.8)
    box1_top = Inches(1.4)
    box1_width = Inches(6.0)
    box1_height = Inches(2.2)

    bg_shape1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, box1_left, box1_top, box1_width, box1_height)
    bg_shape1.fill.solid()
    bg_shape1.fill.fore_color.rgb = WHITE
    bg_shape1.line.color.rgb = BORDER_GRAY
    bg_shape1.line.width = Pt(1)

    # Warning Icon
    warn_icon = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, box1_left + Inches(0.8), box1_top + Inches(0.3), Inches(1.2), Inches(1.0))
    warn_icon.fill.solid()
    warn_icon.fill.fore_color.rgb = RGBColor(0xFF, 0xCD, 0xD2)
    warn_icon.line.color.rgb = RED
    warn_icon.line.width = Pt(3)
    txBox = slide.shapes.add_textbox(box1_left + Inches(0.8), box1_top + Inches(0.4), Inches(1.2), Inches(0.8))
    p = txBox.text_frame.paragraphs[0]
    p.text = "!"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RED

    # Donut Chart
    chart_data = CategoryChartData()
    chart_data.categories = ['虚构字段', '真实字段']
    chart_data.add_series('Series 1', (50, 50))
    x, y, cx, cy = box1_left + Inches(3.0), box1_top + Inches(0.1), Inches(2.0), Inches(1.5)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, chart_data).chart
    chart.has_legend = False
    points = chart.series[0].points
    points[0].format.fill.solid()
    points[0].format.fill.fore_color.rgb = RED
    points[1].format.fill.solid()
    points[1].format.fill.fore_color.rgb = GRAY_BAR

    # Chart Labels
    lbl1 = slide.shapes.add_textbox(box1_left + Inches(3.2), box1_top + Inches(0.6), Inches(1), Inches(0.4))
    lbl1.text_frame.text = "50%\n虚构字段"
    lbl1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    lbl1.text_frame.paragraphs[0].font.size = Pt(10)
    lbl1.text_frame.paragraphs[0].font.color.rgb = RED
    lbl1.text_frame.paragraphs[0].font.bold = True

    lbl2 = slide.shapes.add_textbox(box1_left + Inches(4.2), box1_top + Inches(0.6), Inches(1), Inches(0.4))
    lbl2.text_frame.text = "50%\n真实字段"
    lbl2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    lbl2.text_frame.paragraphs[0].font.size = Pt(10)
    lbl2.text_frame.paragraphs[0].font.color.rgb = GRAY_TEXT

    # Warning Text
    warn_text_box = slide.shapes.add_textbox(box1_left, box1_top + Inches(1.6), box1_width, Inches(0.4))
    p = warn_text_box.text_frame.paragraphs[0]
    p.text = "警告：虚构字段占比过高，引发严重错误"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RED
    p.font.name = FONT_NAME

    # 4. Right Column - Bottom Box (Funnels)
    box2_left = Inches(6.8)
    box2_top = Inches(3.8)
    box2_width = Inches(6.0)
    box2_height = Inches(3.3)

    bg_shape2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, box2_left, box2_top, box2_width, box2_height)
    bg_shape2.fill.solid()
    bg_shape2.fill.fore_color.rgb = WHITE
    bg_shape2.line.color.rgb = BORDER_GRAY
    bg_shape2.line.width = Pt(1)

    # --- Left Funnel ---
    f1_cx = box2_left + Inches(1.5)
    f1_top = box2_top + Inches(0.5)

    t1 = slide.shapes.add_textbox(f1_cx - Inches(1.5), box2_top + Inches(0.05), Inches(3), Inches(0.4))
    t1.text_frame.text = "AI生成过程\n(Claude Code)"
    t1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    t1.text_frame.paragraphs[0].font.size = Pt(10)
    t1.text_frame.paragraphs[0].font.bold = True

    l1 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, f1_cx - Inches(1), f1_top, Inches(2), Inches(0.35))
    l1.rotation = 180; l1.fill.solid(); l1.fill.fore_color.rgb = FUNNEL_DARK_BLUE; l1.line.fill.background()
    l1.text_frame.text = "需求输入"; l1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; l1.text_frame.paragraphs[0].font.size = Pt(10); l1.text_frame.paragraphs[0].font.color.rgb = WHITE

    l2 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, f1_cx - Inches(0.75), f1_top + Inches(0.4), Inches(1.5), Inches(0.35))
    l2.rotation = 180; l2.fill.solid(); l2.fill.fore_color.rgb = FUNNEL_MED_BLUE; l2.line.fill.background()
    l2.text_frame.text = "代码生成"; l2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; l2.text_frame.paragraphs[0].font.size = Pt(10); l2.text_frame.paragraphs[0].font.color.rgb = WHITE

    l3 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, f1_cx - Inches(0.5), f1_top + Inches(0.8), Inches(1.0), Inches(0.35))
    l3.rotation = 180; l3.fill.solid(); l3.fill.fore_color.rgb = FUNNEL_LIGHT_BLUE; l3.line.fill.background()
    l3.text_frame.text = "可用代码"; l3.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; l3.text_frame.paragraphs[0].font.size = Pt(10); l3.text_frame.paragraphs[0].font.color.rgb = BLACK

    # Left Funnel Annotations
    arrow_down = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, f1_cx - Inches(1.4), f1_top, Inches(0.15), Inches(1.2))
    arrow_down.fill.solid(); arrow_down.fill.fore_color.rgb = RED; arrow_down.line.fill.background()

    cross_bg = slide.shapes.add_shape(MSO_SHAPE.OVAL, f1_cx - Inches(1.475), f1_top + Inches(1.25), Inches(0.3), Inches(0.3))
    cross_bg.fill.solid(); cross_bg.fill.fore_color.rgb = RED; cross_bg.line.fill.background()
    cross_txt = slide.shapes.add_textbox(f1_cx - Inches(1.475), f1_top + Inches(1.2), Inches(0.3), Inches(0.3))
    cross_txt.text_frame.text = "X"; cross_txt.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; cross_txt.text_frame.paragraphs[0].font.color.rgb = WHITE; cross_txt.text_frame.paragraphs[0].font.bold = True

    a1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, f1_cx + Inches(0.8), f1_top + Inches(0.5), Inches(0.3), Inches(0.1))
    a1.fill.solid(); a1.fill.fore_color.rgb = RED; a1.line.fill.background()
    t_a1 = slide.shapes.add_textbox(f1_cx + Inches(1.1), f1_top + Inches(0.35), Inches(1.5), Inches(0.4))
    t_a1.text_frame.text = "数据字典“幻觉”\n虚构字段"
    t_a1.text_frame.paragraphs[0].font.size = Pt(8); t_a1.text_frame.paragraphs[0].font.bold = True; t_a1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    a2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, f1_cx + Inches(0.6), f1_top + Inches(0.9), Inches(0.3), Inches(0.1))
    a2.fill.solid(); a2.fill.fore_color.rgb = RED; a2.line.fill.background()
    t_a2 = slide.shapes.add_textbox(f1_cx + Inches(0.9), f1_top + Inches(0.8), Inches(1.5), Inches(0.4))
    t_a2.text_frame.text = "21+ 语法错误\n修复成本高"
    t_a2.text_frame.paragraphs[0].font.size = Pt(8); t_a2.text_frame.paragraphs[0].font.bold = True; t_a2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    block = slide.shapes.add_shape(MSO_SHAPE.NO_SYMBOL, f1_cx + Inches(1.5), f1_top + Inches(0.7), Inches(0.3), Inches(0.3))
    block.fill.solid(); block.fill.fore_color.rgb = DARK_BLUE_BG; block.line.fill.background()
    t_block = slide.shapes.add_textbox(f1_cx + Inches(1.2), f1_top + Inches(1.0), Inches(0.9), Inches(0.4))
    t_block.text_frame.text = "高遗失,\n不可用"
    t_block.text_frame.paragraphs[0].font.size = Pt(8); t_block.text_frame.paragraphs[0].font.bold = True; t_block.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    b1 = slide.shapes.add_textbox(f1_cx - Inches(1), f1_top + Inches(1.25), Inches(2), Inches(0.3))
    b1.text_frame.text = "成功上线"
    b1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; b1.text_frame.paragraphs[0].font.size = Pt(9); b1.text_frame.paragraphs[0].font.color.rgb = GRAY_TEXT

    b_lbl1 = slide.shapes.add_textbox(f1_cx - Inches(1.5), f1_top + Inches(1.55), Inches(3), Inches(0.4))
    b_lbl1.text_frame.text = "AI生成过程\n(Claude Code)"
    b_lbl1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; b_lbl1.text_frame.paragraphs[0].font.size = Pt(9); b_lbl1.text_frame.paragraphs[0].font.bold = True

    # --- Right Funnel ---
    f2_cx = box2_left + Inches(4.5)
    f2_top = box2_top + Inches(0.5)

    t2 = slide.shapes.add_textbox(f2_cx - Inches(1.5), box2_top + Inches(0.05), Inches(3), Inches(0.4))
    t2.text_frame.text = "理想过程 / 人工修复\n(Contrast)"
    t2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    t2.text_frame.paragraphs[0].font.size = Pt(10)
    t2.text_frame.paragraphs[0].font.bold = True

    r1 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, f2_cx - Inches(1), f2_top, Inches(2), Inches(0.35))
    r1.rotation = 180; r1.fill.solid(); r1.fill.fore_color.rgb = FUNNEL_DARK_BLUE; r1.line.fill.background()
    r1.text_frame.text = "需求输入"; r1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; r1.text_frame.paragraphs[0].font.size = Pt(10); r1.text_frame.paragraphs[0].font.color.rgb = WHITE

    r2 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, f2_cx - Inches(0.75), f2_top + Inches(0.4), Inches(1.5), Inches(0.35))
    r2.rotation = 180; r2.fill.solid(); r2.fill.fore_color.rgb = FUNNEL_TEAL; r2.line.fill.background()
    r2.text_frame.text = "人工/高质量代码"; r2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; r2.text_frame.paragraphs[0].font.size = Pt(10); r2.text_frame.paragraphs[0].font.color.rgb = WHITE

    r3 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, f2_cx - Inches(0.5), f2_top + Inches(0.8), Inches(1.0), Inches(0.35))
    r3.rotation = 180; r3.fill.solid(); r3.fill.fore_color.rgb = GREEN; r3.line.fill.background()
    r3.text_frame.text = "成功上线"; r3.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; r3.text_frame.paragraphs[0].font.size = Pt(10); r3.text_frame.paragraphs[0].font.color.rgb = WHITE

    arrow_up = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, f2_cx + Inches(1.2), f2_top, Inches(0.15), Inches(1.2))
    arrow_up.fill.solid(); arrow_up.fill.fore_color.rgb = GREEN; arrow_up.line.fill.background()

    check_bg = slide.shapes.add_shape(MSO_SHAPE.OVAL, f2_cx + Inches(1.125), f2_top + Inches(1.25), Inches(0.3), Inches(0.3))
    check_bg.fill.solid(); check_bg.fill.fore_color.rgb = GREEN; check_bg.line.fill.background()
    check_txt = slide.shapes.add_textbox(f2_cx + Inches(1.125), f2_top + Inches(1.2), Inches(0.3), Inches(0.3))
    check_txt.text_frame.text = "✓"; check_txt.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; check_txt.text_frame.paragraphs[0].font.color.rgb = WHITE; check_txt.text_frame.paragraphs[0].font.bold = True

    b_lbl2 = slide.shapes.add_textbox(f2_cx - Inches(1.5), f2_top + Inches(1.55), Inches(3), Inches(0.4))
    b_lbl2.text_frame.text = "理想过程 / 人工修复\n(Contrast)"
    b_lbl2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; b_lbl2.text_frame.paragraphs[0].font.size = Pt(9); b_lbl2.text_frame.paragraphs[0].font.bold = True

    # --- Conclusion Box ---
    conc_left = box2_left + Inches(0.1)
    conc_top = box2_top + box2_height - Inches(0.75)
    conc_width = box2_width - Inches(0.2)
    
    conc_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, conc_left, conc_top, conc_width, Inches(0.65))
    conc_bg.fill.solid()
    conc_bg.fill.fore_color.rgb = CONCL_BG
    conc_bg.line.fill.background()
    
    conc_txt = slide.shapes.add_textbox(conc_left, conc_top, conc_width, Inches(0.65))
    tf = conc_txt.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "核心结论：AI 在数据字典和接口规范上的“幻觉”导致代码质量极低，修复成本远超预期，当前阶段不可直接用于复杂业务场景。"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = CONCL_TEXT
    p.font.name = FONT_NAME

    # 5. Footer
    footer = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer.text_frame.text = "第2页，共4页"
    footer.text_frame.paragraphs[0].font.size = Pt(12)
    footer.text_frame.paragraphs[0].font.color.rgb = GRAY_TEXT