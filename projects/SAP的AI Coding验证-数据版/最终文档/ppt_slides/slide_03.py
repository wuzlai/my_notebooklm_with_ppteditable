def build_slide(slide):
    # Colors
    DARK_BLUE = RGBColor(0x1A, 0x3B, 0x66)
    GRAY_TEXT = RGBColor(0x55, 0x55, 0x55)
    LIGHT_GRAY_BG = RGBColor(0xF8, 0xF9, 0xFA)
    BORDER_GRAY = RGBColor(0xE0, 0xE0, 0xE0)
    RED = RGBColor(0xD3, 0x2F, 0x2F)
    BLUE = RGBColor(0x19, 0x76, 0xD2)
    GREEN_BG = RGBColor(0xEA, 0xF8, 0xE6)
    GREEN_BORDER = RGBColor(0x8B, 0xF3, 0x69)
    GREEN_ICON = RGBColor(0x4C, 0xAF, 0x50)
    BLACK = RGBColor(0x00, 0x00, 0x00)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    # 1. Title Area
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.0), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "第3页 | 高复杂度：效率反降 60% 的“修复爆炸”"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE

    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12.0), Inches(0.5))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = "案例C - 跨工厂 STO 报表（高复杂度）验证"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(20)
    p.font.color.rgb = GRAY_TEXT

    # 2. Left Panel (Comparison)
    left_panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.8), Inches(4.6), Inches(5.2))
    left_panel.fill.solid()
    left_panel.fill.fore_color.rgb = LIGHT_GRAY_BG
    left_panel.line.color.rgb = BORDER_GRAY
    left_panel.line.width = Pt(1)

    # Vertical Divider in Left Panel
    divider = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(2.9), Inches(2.2), Inches(2.9), Inches(6.6))
    divider.line.color.rgb = BORDER_GRAY
    divider.line.width = Pt(1)

    # Left Column (Claude Code)
    tx_claude = slide.shapes.add_textbox(Inches(0.6), Inches(2.2), Inches(2.3), Inches(0.8))
    tf = tx_claude.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "Claude Code\n"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(16)
    run.font.bold = True
    run = p.add_run()
    run.text = "修复耗时"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(16)

    # Red Stopwatch Icon
    stopwatch_red = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.55), Inches(3.2), Inches(0.4), Inches(0.4))
    stopwatch_red.fill.background()
    stopwatch_red.line.color.rgb = RED
    stopwatch_red.line.width = Pt(3)
    sw_top_red = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.7), Inches(3.1), Inches(0.1), Inches(0.1))
    sw_top_red.fill.solid()
    sw_top_red.fill.fore_color.rgb = RED
    sw_top_red.line.fill.background()
    sw_hand_red = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1.75), Inches(3.4), Inches(1.75), Inches(3.25))
    sw_hand_red.line.color.rgb = RED
    sw_hand_red.line.width = Pt(2)

    # 8 人天
    tx_8 = slide.shapes.add_textbox(Inches(0.6), Inches(3.8), Inches(2.3), Inches(0.8))
    tf = tx_8.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "8 "
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(44)
    run.font.bold = True
    run = p.add_run()
    run.text = "人天"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(20)

    # -60% Efficiency
    arrow_down = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(1.0), Inches(5.1), Inches(0.2), Inches(0.3))
    arrow_down.fill.solid()
    arrow_down.fill.fore_color.rgb = RED
    arrow_down.line.fill.background()

    tx_eff = slide.shapes.add_textbox(Inches(1.2), Inches(4.9), Inches(1.7), Inches(0.8))
    tf = tx_eff.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "-60%\n"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RED
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = "效率"
    run2.font.name = "Microsoft YaHei"
    run2.font.size = Pt(16)

    # Right Column (Manual)
    tx_manual = slide.shapes.add_textbox(Inches(2.9), Inches(2.2), Inches(2.3), Inches(0.8))
    tf = tx_manual.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "\n手写开发耗时"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(16)
    run.font.bold = True

    # Blue Stopwatch Icon
    stopwatch_blue = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.85), Inches(3.2), Inches(0.4), Inches(0.4))
    stopwatch_blue.fill.background()
    stopwatch_blue.line.color.rgb = BLUE
    stopwatch_blue.line.width = Pt(3)
    sw_top_blue = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.0), Inches(3.1), Inches(0.1), Inches(0.1))
    sw_top_blue.fill.solid()
    sw_top_blue.fill.fore_color.rgb = BLUE
    sw_top_blue.line.fill.background()
    sw_hand_blue = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(4.05), Inches(3.4), Inches(4.05), Inches(3.25))
    sw_hand_blue.line.color.rgb = BLUE
    sw_hand_blue.line.width = Pt(2)

    # 5 人天
    tx_5 = slide.shapes.add_textbox(Inches(2.9), Inches(3.8), Inches(2.3), Inches(0.8))
    tf = tx_5.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "5 "
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(44)
    run.font.bold = True
    run = p.add_run()
    run.text = "人天"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(20)

    # 3. Right Top Panel (Trend Chart)
    right_top_panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.5), Inches(1.8), Inches(7.3), Inches(3.6))
    right_top_panel.fill.solid()
    right_top_panel.fill.fore_color.rgb = LIGHT_GRAY_BG
    right_top_panel.line.color.rgb = BORDER_GRAY
    right_top_panel.line.width = Pt(1)

    tx_trend_title = slide.shapes.add_textbox(Inches(5.7), Inches(2.0), Inches(6.0), Inches(0.4))
    tf = tx_trend_title.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "修复-爆炸模式：错误数量非线性增长趋势"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(16)
    run.font.bold = True

    # Trend Line Segments
    points = [
        (6.7, 4.2), (7.9, 4.2), (7.9, 4.5), (9.1, 4.5), 
        (10.2, 3.2), (10.6, 3.8), (11.8, 2.6)
    ]
    for i in range(len(points) - 1):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(points[i][0]), Inches(points[i][1]), Inches(points[i+1][0]), Inches(points[i+1][1]))
        line.line.color.rgb = RED
        line.line.width = Pt(4)
        if i == len(points) - 2:
            line.line.end_arrowhead = 2 # Triangle arrowhead

    # Data Points (Circles)
    def add_data_point(x, y, text1, text2):
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x-0.08), Inches(y-0.08), Inches(0.16), Inches(0.16))
        circle.fill.solid()
        circle.fill.fore_color.rgb = RED
        circle.line.color.rgb = WHITE
        circle.line.width = Pt(2)
        
        tx = slide.shapes.add_textbox(Inches(x-0.6), Inches(y-0.7) if y > 3.5 else Inches(y-0.1), Inches(1.2), Inches(0.6))
        tf = tx.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = text1 + "\n"
        run.font.name = "Microsoft YaHei"
        run.font.size = Pt(14)
        run.font.bold = True
        if "9" in text1:
            run.font.color.rgb = RED
        run2 = p.add_run()
        run2.text = text2
        run2.font.name = "Microsoft YaHei"
        run2.font.size = Pt(12)

    add_data_point(6.7, 4.2, "3 错误", "(起始)")
    add_data_point(9.1, 4.5, "2 错误", "(初次修复)")
    add_data_point(11.2, 3.3, "9 错误", "(引发新矛盾)")

    # Bomb Icon
    bomb_body = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.7), Inches(2.0), Inches(0.45), Inches(0.45))
    bomb_body.fill.solid()
    bomb_body.fill.fore_color.rgb = BLACK
    bomb_body.line.fill.background()
    
    bomb_cap = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.9), Inches(1.9), Inches(0.1), Inches(0.15))
    bomb_cap.fill.solid()
    bomb_cap.fill.fore_color.rgb = GRAY_TEXT
    bomb_cap.line.fill.background()
    
    explosion = slide.shapes.add_shape(MSO_SHAPE.EXPLOSION1, Inches(12.1), Inches(1.8), Inches(0.35), Inches(0.35))
    explosion.fill.solid()
    explosion.fill.fore_color.rgb = RGBColor(0xFF, 0x57, 0x22)
    explosion.line.color.rgb = RGBColor(0xFF, 0xEB, 0x3B)

    # JetBrains Mono Text
    tx_jb = slide.shapes.add_textbox(Inches(8.5), Inches(4.9), Inches(2.0), Inches(0.3))
    tf = tx_jb.text_frame
    p = tf.paragraphs[0]
    p.text = "JetBrains Mono"
    p.font.name = "Consolas"
    p.font.size = Pt(10)
    p.font.color.rgb = GRAY_TEXT

    # 4. Right Bottom Left (Data Dictionary)
    # Database Icon
    for i in range(3):
        cyl = slide.shapes.add_shape(MSO_SHAPE.CAN, Inches(5.5), Inches(5.7 + i*0.12), Inches(0.3), Inches(0.18))
        cyl.fill.solid()
        cyl.fill.fore_color.rgb = RGBColor(0x78, 0x90, 0x9C)
        cyl.line.color.rgb = WHITE
    
    cross = slide.shapes.add_shape(MSO_SHAPE.MATH_MULTIPLY, Inches(5.7), Inches(5.9), Inches(0.15), Inches(0.15))
    cross.fill.solid()
    cross.fill.fore_color.rgb = RED
    cross.line.fill.background()

    tx_dict_title = slide.shapes.add_textbox(Inches(5.9), Inches(5.65), Inches(3.0), Inches(0.4))
    tf = tx_dict_title.text_frame
    p = tf.paragraphs[0]
    p.text = "数据字典重灾区"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(14)
    p.font.bold = True

    tx_dict_desc = slide.shapes.add_textbox(Inches(5.4), Inches(6.1), Inches(3.5), Inches(0.8))
    tf = tx_dict_desc.text_frame
    p1 = tf.paragraphs[0]
    run1 = p1.add_run()
    run1.text = "累计虚构 10+ 项表字段和数据类型\n"
    run1.font.name = "Microsoft YaHei"
    run1.font.size = Pt(12)
    run1.font.bold = True
    run1.font.color.rgb = RED
    
    p2 = tf.add_paragraph()
    run2 = p2.add_run()
    run2.text = "开发者需大量时间查表对数"
    run2.font.name = "Microsoft YaHei"
    run2.font.size = Pt(12)
    run2.font.color.rgb = BLACK

    # 5. Right Bottom Right (Advantage Box)
    adv_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.2), Inches(5.5), Inches(3.6), Inches(1.4))
    adv_box.fill.solid()
    adv_box.fill.fore_color.rgb = GREEN_BG
    adv_box.line.color.rgb = GREEN_BORDER
    adv_box.line.width = Pt(2)

    # Document Icon
    doc = slide.shapes.add_shape(MSO_SHAPE.FOLDED_CORNER, Inches(9.4), Inches(5.7), Inches(0.25), Inches(0.35))
    doc.fill.solid()
    doc.fill.fore_color.rgb = WHITE
    doc.line.color.rgb = GRAY_TEXT
    
    chk = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.55), Inches(5.9), Inches(0.15), Inches(0.15))
    chk.fill.solid()
    chk.fill.fore_color.rgb = GREEN_ICON
    chk.line.fill.background()

    tx_adv_title = slide.shapes.add_textbox(Inches(9.8), Inches(5.65), Inches(2.8), Inches(0.4))
    tf = tx_adv_title.text_frame
    p = tf.paragraphs[0]
    p.text = "文档解析优势 (Claude)"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(14)
    p.font.bold = True

    tx_adv_desc = slide.shapes.add_textbox(Inches(9.3), Inches(6.1), Inches(3.4), Inches(0.7))
    tf = tx_adv_desc.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "● Claude 虽支持文档解析，但生成的代码架构因底层逻辑矛盾而无法运行。"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(11)
    run.font.color.rgb = BLACK