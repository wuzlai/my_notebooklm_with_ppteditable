def build_slide(slide):
    # Colors
    DARK_BLUE = RGBColor(0x1B, 0x2A, 0x49)
    GRAY_TEXT = RGBColor(0x66, 0x66, 0x66)
    LIGHT_GRAY = RGBColor(0x99, 0x99, 0x99)
    GREEN = RGBColor(0x2E, 0xA1, 0x54)
    RED = RGBColor(0xD9, 0x3A, 0x36)
    BLUE_ICON = RGBColor(0x4A, 0x86, 0xC8)
    BORDER_COLOR = RGBColor(0xE5, 0xE5, 0xE5)
    BLACK = RGBColor(0x00, 0x00, 0x00)

    # 1. Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(8.0), Inches(0.6))
    tf_title = title_box.text_frame
    p_title = tf_title.paragraphs[0]
    run_title = p_title.add_run()
    run_title.text = "简单场景验证：Claude Code 胜出"
    run_title.font.size = Pt(28)
    run_title.font.bold = True
    run_title.font.color.rgb = DARK_BLUE
    run_title.font.name = "Microsoft YaHei"

    # 2. Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(8.0), Inches(0.4))
    tf_sub = sub_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    run_sub = p_sub.add_run()
    run_sub.text = "销售订单报表查询（低复杂度）效率对比"
    run_sub.font.size = Pt(16)
    run_sub.font.color.rgb = GRAY_TEXT
    run_sub.font.name = "Microsoft YaHei"

    # 3. Left Panel (White Box)
    left_panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.7), Inches(5.4), Inches(4.8))
    left_panel.fill.solid()
    left_panel.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    left_panel.line.color.rgb = BORDER_COLOR
    left_panel.line.width = Pt(1)

    # Helper for left panel text
    def add_left_text(y, runs):
        tb = slide.shapes.add_textbox(Inches(1.2), y, Inches(4.5), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.line_spacing = 1.3
        for text, is_bold, color in runs:
            r = p.add_run()
            r.text = text
            r.font.name = "Microsoft YaHei"
            r.font.size = Pt(13)
            if is_bold:
                r.font.bold = True
            if color:
                r.font.color.rgb = color

    # Item 1: Up Arrow Icon & Text
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.75), Inches(2.05), Inches(0.28), Inches(0.28))
    circle.fill.background()
    circle.line.color.rgb = GREEN
    circle.line.width = Pt(1.5)
    arrow = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(0.84), Inches(2.1), Inches(0.1), Inches(0.18))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = GREEN
    arrow.line.fill.background()
    
    add_left_text(Inches(1.95), [
        ("效率拐点：", True, BLACK),
        ("Claude Code 耗时 ", False, BLACK),
        ("30 分钟", True, BLACK),
        ("，较手写开发效率", False, BLACK),
        ("提升 50%", True, BLACK),
        ("。", False, BLACK)
    ])

    # Item 2: Warning Icon & Text
    triangle = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(0.75), Inches(3.05), Inches(0.28), Inches(0.28))
    triangle.fill.background()
    triangle.line.color.rgb = RED
    triangle.line.width = Pt(1.5)
    warn_tb = slide.shapes.add_textbox(Inches(0.75), Inches(3.1), Inches(0.28), Inches(0.28))
    warn_p = warn_tb.text_frame.paragraphs[0]
    warn_p.alignment = PP_ALIGN.CENTER
    warn_r = warn_p.add_run()
    warn_r.text = "!"
    warn_r.font.size = Pt(12)
    warn_r.font.bold = True
    warn_r.font.color.rgb = RED

    add_left_text(Inches(2.95), [
        ("稳定性差异：", True, BLACK),
        ("GitHub Copilot 运行出现 ", False, BLACK),
        ("Short Dump", True, RED),
        ("，而 Claude Code ", False, BLACK),
        ("运行正常", True, GREEN),
        ("。", False, BLACK)
    ])

    # Item 3: Chat Icon & Text
    chat1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGULAR_CALLOUT, Inches(0.75), Inches(4.1), Inches(0.22), Inches(0.18))
    chat1.fill.background()
    chat1.line.color.rgb = BLUE_ICON
    chat1.line.width = Pt(1.5)
    chat2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGULAR_CALLOUT, Inches(0.82), Inches(4.18), Inches(0.22), Inches(0.18))
    chat2.fill.background()
    chat2.line.color.rgb = BLUE_ICON
    chat2.line.width = Pt(1.5)

    add_left_text(Inches(3.95), [
        ("交互成本：", True, BLACK),
        ("Claude Code 仅需 ", False, BLACK),
        ("2 轮", True, BLACK),
        ("人工干预，远优于 Copilot 的 ", False, BLACK),
        ("5 轮", True, BLACK),
        ("以上。", False, BLACK)
    ])

    # Item 4: Link Icon & Text
    link1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.78), Inches(5.15), Inches(0.18), Inches(0.1))
    link1.rotation = 45
    link1.fill.background()
    link1.line.color.rgb = LIGHT_GRAY
    link1.line.width = Pt(1.5)
    link2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.86), Inches(5.23), Inches(0.18), Inches(0.1))
    link2.rotation = 45
    link2.fill.background()
    link2.line.color.rgb = LIGHT_GRAY
    link2.line.width = Pt(1.5)

    add_left_text(Inches(4.95), [
        ("核心瓶颈：", True, BLACK),
        ("两者初次生成均不可直接运行，仍需人工修正 ", False, BLACK),
        ("SQL", True, BLACK),
        (" 逻辑。", False, BLACK)
    ])

    # 4. Right Top Panel (Chart)
    chart_title = slide.shapes.add_textbox(Inches(6.4), Inches(1.7), Inches(4.0), Inches(0.4))
    p_ct = chart_title.text_frame.paragraphs[0]
    r_ct = p_ct.add_run()
    r_ct.text = "开发耗时对比（分钟）"
    r_ct.font.size = Pt(14)
    r_ct.font.bold = True
    r_ct.font.name = "Microsoft YaHei"

    chart_data = CategoryChartData()
    chart_data.categories = ['Claude Code', 'GitHub Copilot', '手写开发']
    chart_data.add_series('耗时', (30, 65, 60))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, Inches(6.4), Inches(2.2), Inches(6.4), Inches(2.0), chart_data
    )
    chart = chart_shape.chart
    chart.has_legend = False

    val_axis = chart.value_axis
    val_axis.maximum_scale = 70
    val_axis.major_unit = 15
    val_axis.has_major_gridlines = True
    val_axis.major_gridlines.format.line.color.rgb = BORDER_COLOR
    val_axis.tick_labels.font.size = Pt(10)
    val_axis.tick_labels.font.color.rgb = GRAY_TEXT

    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.size = Pt(11)
    cat_axis.tick_labels.font.color.rgb = GRAY_TEXT
    cat_axis.has_major_gridlines = False

    series = chart.series[0]
    series.has_data_labels = True

    # Claude Code (Green)
    p0 = series.points[0]
    p0.format.fill.solid()
    p0.format.fill.fore_color.rgb = GREEN
    p0.data_label.has_text_frame = True
    p0.data_label.text_frame.text = "30 min ✅"
    p0.data_label.font.size = Pt(10)
    p0.data_label.font.bold = True

    # GitHub Copilot (Red)
    p1 = series.points[1]
    p1.format.fill.solid()
    p1.format.fill.fore_color.rgb = RED
    p1.data_label.has_text_frame = True
    p1.data_label.text_frame.text = "> 60 min ⚠️"
    p1.data_label.font.size = Pt(10)
    p1.data_label.font.bold = True
    p1.data_label.font.color.rgb = RED

    # 手写开发 (Red)
    p2 = series.points[2]
    p2.format.fill.solid()
    p2.format.fill.fore_color.rgb = RED
    p2.data_label.has_text_frame = True
    p2.data_label.text_frame.text = "60 min"
    p2.data_label.font.size = Pt(10)
    p2.data_label.font.bold = True

    # 5. Right Bottom Left Panel (Data Card)
    card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.4), Inches(4.5), Inches(2.6), Inches(2.0))
    card1.fill.solid()
    card1.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    card1.line.color.rgb = BORDER_COLOR

    tb_50 = slide.shapes.add_textbox(Inches(6.5), Inches(4.8), Inches(2.0), Inches(0.6))
    p_50 = tb_50.text_frame.paragraphs[0]
    r_50 = p_50.add_run()
    r_50.text = "50%"
    r_50.font.size = Pt(40)
    r_50.font.bold = True
    r_50.font.color.rgb = GREEN

    tb_eff = slide.shapes.add_textbox(Inches(6.5), Inches(5.5), Inches(2.0), Inches(0.4))
    p_eff = tb_eff.text_frame.paragraphs[0]
    r_eff1 = p_eff.add_run()
    r_eff1.text = "效率提升 "
    r_eff1.font.size = Pt(16)
    r_eff1.font.bold = True
    r_eff1.font.name = "Microsoft YaHei"
    r_eff2 = p_eff.add_run()
    r_eff2.text = "↗"
    r_eff2.font.size = Pt(16)
    r_eff2.font.bold = True
    r_eff2.font.color.rgb = GREEN

    tb_vs = slide.shapes.add_textbox(Inches(6.5), Inches(6.0), Inches(2.4), Inches(0.3))
    p_vs = tb_vs.text_frame.paragraphs[0]
    r_vs = p_vs.add_run()
    r_vs.text = "Claude Code vs 手写开发"
    r_vs.font.size = Pt(10)
    r_vs.font.color.rgb = GRAY_TEXT
    r_vs.font.name = "Microsoft YaHei"

    # 6. Right Bottom Right Panel (Grid)
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.3), Inches(4.5), Inches(3.5), Inches(2.0))
    card2.fill.solid()
    card2.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    card2.line.color.rgb = BORDER_COLOR

    line_v = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(11.05), Inches(4.5), Inches(11.05), Inches(6.5))
    line_v.line.color.rgb = BORDER_COLOR
    line_h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(9.3), Inches(5.5), Inches(12.8), Inches(5.5))
    line_h.line.color.rgb = BORDER_COLOR

    def add_grid_cell(x, y, icon_text, icon_color, text, icon_size=24):
        tb = slide.shapes.add_textbox(x, y, Inches(1.75), Inches(1.0))
        tf = tb.text_frame
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = icon_text + "\n"
        r1.font.size = Pt(icon_size)
        r1.font.color.rgb = icon_color
        r1.font.name = "Segoe UI Emoji"

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = text
        r2.font.size = Pt(11)
        r2.font.color.rgb = BLACK
        r2.font.name = "Microsoft YaHei"

    add_grid_cell(Inches(9.3), Inches(4.6), "❌", RED, "GitHub Copilot")
    add_grid_cell(Inches(11.05), Inches(4.6), "❌", RED, "Short Dump")
    add_grid_cell(Inches(9.3), Inches(5.6), "✅", GREEN, "Copilot: 5+ 轮")
    add_grid_cell(Inches(11.05), Inches(5.6), "👥", BLUE_ICON, "Claude Code: 2 轮")

    # 7. Page Number
    page_num = slide.shapes.add_textbox(Inches(11.8), Inches(6.9), Inches(1.0), Inches(0.3))
    p_page = page_num.text_frame.paragraphs[0]
    p_page.alignment = PP_ALIGN.RIGHT
    r_page = p_page.add_run()
    r_page.text = "Page 1 / 4"
    r_page.font.size = Pt(10)
    r_page.font.color.rgb = LIGHT_GRAY
    r_page.font.name = "Microsoft YaHei"