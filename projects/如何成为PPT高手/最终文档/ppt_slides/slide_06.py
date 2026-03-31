def build_slide(slide):
    # Colors
    BLUE_PRIMARY = RGBColor(0x00, 0x52, 0xCC)
    TEXT_BLACK = RGBColor(0x33, 0x33, 0x33)
    TEXT_GRAY = RGBColor(0x7F, 0x7F, 0x7F)
    LIGHT_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
    CUBE_FILL = RGBColor(0xF4, 0xF6, 0xF9)
    CUBE_LINE = RGBColor(0x2F, 0x45, 0x6A)
    HIGHLIGHT_FILL = RGBColor(0xDE, 0xEA, 0xF6)
    HIGHLIGHT_LINE = RGBColor(0x5B, 0x9B, 0xD5)

    # 1. Top Left Page Indicator
    # Small grey dash
    dash = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.15), Inches(0.04))
    dash.fill.solid()
    dash.fill.fore_color.rgb = LIGHT_GRAY
    dash.line.fill.background()

    # "Page 6"
    tx_page = slide.shapes.add_textbox(Inches(0.35), Inches(0.45), Inches(1), Inches(0.3))
    tf_page = tx_page.text_frame
    tf_page.word_wrap = False
    p_page = tf_page.paragraphs[0]
    p_page.text = "Page 6"
    p_page.font.name = "Microsoft YaHei"
    p_page.font.size = Pt(12)
    p_page.font.bold = True
    p_page.font.color.rgb = TEXT_BLACK

    # "6/11"
    tx_num = slide.shapes.add_textbox(Inches(0.35), Inches(0.7), Inches(1), Inches(0.3))
    tf_num = tx_num.text_frame
    p_num = tf_num.paragraphs[0]
    p_num.text = "6/11"
    p_num.font.name = "Microsoft YaHei"
    p_num.font.size = Pt(10)
    p_num.font.color.rgb = TEXT_GRAY

    # 2. Main Title
    tx_title = slide.shapes.add_textbox(Inches(2), Inches(0.8), Inches(9.333), Inches(0.8))
    tf_title = tx_title.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.alignment = PP_ALIGN.CENTER
    
    run1 = p_title.add_run()
    run1.text = "法则二："
    run1.font.name = "Microsoft YaHei"
    run1.font.size = Pt(36)
    run1.font.bold = True
    run1.font.color.rgb = TEXT_BLACK

    run2 = p_title.add_run()
    run2.text = "设计统一建立专业信任"
    run2.font.name = "Microsoft YaHei"
    run2.font.size = Pt(36)
    run2.font.bold = True
    run2.font.color.rgb = BLUE_PRIMARY

    # 3. Subtitle
    tx_sub = slide.shapes.add_textbox(Inches(2), Inches(1.7), Inches(9.333), Inches(0.5))
    tf_sub = tx_sub.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.text = "从视觉一致性中体现专业度"
    p_sub.font.name = "Microsoft YaHei"
    p_sub.font.size = Pt(20)
    p_sub.font.color.rgb = TEXT_BLACK

    # 4. Central Graphics (Cubes)
    cube_y = Inches(2.8)
    cube_w = Inches(2.2)
    cube_h = Inches(2.2)

    # Cube 1 (Left)
    cube1 = slide.shapes.add_shape(MSO_SHAPE.CUBE, Inches(2.8), cube_y, cube_w, cube_h)
    cube1.fill.solid()
    cube1.fill.fore_color.rgb = CUBE_FILL
    cube1.line.color.rgb = CUBE_LINE
    cube1.line.width = Pt(2)

    # Cube 2 (Middle)
    cube2 = slide.shapes.add_shape(MSO_SHAPE.CUBE, Inches(5.5), cube_y, cube_w, cube_h)
    cube2.fill.solid()
    cube2.fill.fore_color.rgb = CUBE_FILL
    cube2.line.color.rgb = CUBE_LINE
    cube2.line.width = Pt(2)

    # Flying piece for Cube 2
    small_cube = slide.shapes.add_shape(MSO_SHAPE.CUBE, Inches(6.8), Inches(2.5), Inches(0.8), Inches(0.8))
    small_cube.fill.solid()
    small_cube.fill.fore_color.rgb = HIGHLIGHT_FILL
    small_cube.line.color.rgb = HIGHLIGHT_LINE
    small_cube.line.width = Pt(1.5)

    # Arrow for flying piece
    arrow_insert = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(7.1), Inches(3.4), Inches(0.2), Inches(0.3))
    arrow_insert.fill.solid()
    arrow_insert.fill.fore_color.rgb = HIGHLIGHT_LINE
    arrow_insert.line.fill.background()

    # Cube 3 (Right)
    cube3 = slide.shapes.add_shape(MSO_SHAPE.CUBE, Inches(8.2), cube_y, cube_w, cube_h)
    cube3.fill.solid()
    cube3.fill.fore_color.rgb = CUBE_FILL
    cube3.line.color.rgb = CUBE_LINE
    cube3.line.width = Pt(2)

    # Highlighted piece on Cube 3 (Simulated by a smaller cube on top right corner)
    hl_cube = slide.shapes.add_shape(MSO_SHAPE.CUBE, Inches(9.2), Inches(3.2), Inches(0.7), Inches(0.7))
    hl_cube.fill.solid()
    hl_cube.fill.fore_color.rgb = HIGHLIGHT_FILL
    hl_cube.line.color.rgb = HIGHLIGHT_LINE
    hl_cube.line.width = Pt(1.5)

    # 5. Bottom Arrow and Text
    # Long arrow line
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(3.8), Inches(5.4), Inches(9.4), Inches(5.4))
    connector.line.color.rgb = HIGHLIGHT_LINE
    connector.line.width = Pt(1.5)
    # Add arrow head (using standard line properties if possible, or draw a small triangle)
    arrow_head = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(9.3), Inches(5.33), Inches(0.15), Inches(0.15))
    arrow_head.rotation = 90
    arrow_head.fill.solid()
    arrow_head.fill.fore_color.rgb = HIGHLIGHT_LINE
    arrow_head.line.fill.background()

    # Text below arrow
    tx_arrow = slide.shapes.add_textbox(Inches(2), Inches(5.6), Inches(9.333), Inches(0.4))
    tf_arrow = tx_arrow.text_frame
    p_arrow = tf_arrow.paragraphs[0]
    p_arrow.alignment = PP_ALIGN.CENTER
    p_arrow.text = "统一感能降低观众的视觉疲劳"
    p_arrow.font.name = "Microsoft YaHei"
    p_arrow.font.size = Pt(14)
    p_arrow.font.bold = True
    p_arrow.font.color.rgb = TEXT_BLACK

    # 6. Bottom 3 Columns
    col_y = Inches(6.3)
    
    # --- Column 1 ---
    # Icon 1: Compass/Ruler
    compass = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(1.0), col_y + Inches(0.1), Inches(0.3), Inches(0.4))
    compass.rotation = -90
    compass.fill.background()
    compass.line.color.rgb = CUBE_LINE
    compass.line.width = Pt(1.5)
    
    ruler = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.4), col_y, Inches(0.15), Inches(0.5))
    ruler.fill.background()
    ruler.line.color.rgb = HIGHLIGHT_LINE
    ruler.line.width = Pt(1.5)

    # Text 1
    tx_col1_title = slide.shapes.add_textbox(Inches(1.7), col_y - Inches(0.1), Inches(3.0), Inches(0.3))
    p_col1_title = tx_col1_title.text_frame.paragraphs[0]
    r1 = p_col1_title.add_run()
    r1.text = "1. "
    r1.font.bold = True
    r1.font.size = Pt(13)
    r2 = p_col1_title.add_run()
    r2.text = "风格漂移"
    r2.font.bold = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = BLUE_PRIMARY
    r3 = p_col1_title.add_run()
    r3.text = "是PPT的大忌"
    r3.font.bold = True
    r3.font.size = Pt(13)

    tx_col1_desc = slide.shapes.add_textbox(Inches(1.7), col_y + Inches(0.2), Inches(3.0), Inches(0.4))
    p_col1_desc = tx_col1_desc.text_frame.paragraphs[0]
    p_col1_desc.text = "避免混乱，保持整体风格的一致性。"
    p_col1_desc.font.size = Pt(11)
    p_col1_desc.font.color.rgb = TEXT_BLACK

    # --- Column 2 ---
    # Icon 2: Eye
    eye_outer = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.0), col_y + Inches(0.1), Inches(0.5), Inches(0.3))
    eye_outer.fill.background()
    eye_outer.line.color.rgb = CUBE_LINE
    eye_outer.line.width = Pt(1.5)
    
    eye_inner = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.15), col_y + Inches(0.15), Inches(0.2), Inches(0.2))
    eye_inner.fill.background()
    eye_inner.line.color.rgb = HIGHLIGHT_LINE
    eye_inner.line.width = Pt(1.5)
    
    pulse = slide.shapes.add_shape(MSO_SHAPE.ZIG_ZAG, Inches(5.0), col_y + Inches(0.45), Inches(0.5), Inches(0.1))
    pulse.fill.background()
    pulse.line.color.rgb = HIGHLIGHT_LINE
    pulse.line.width = Pt(1.5)

    # Text 2
    tx_col2_title = slide.shapes.add_textbox(Inches(5.7), col_y - Inches(0.1), Inches(3.2), Inches(0.3))
    p_col2_title = tx_col2_title.text_frame.paragraphs[0]
    r1 = p_col2_title.add_run()
    r1.text = "2. 统一感能降低观众的"
    r1.font.bold = True
    r1.font.size = Pt(13)
    r2 = p_col2_title.add_run()
    r2.text = "视觉疲劳"
    r2.font.bold = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = BLUE_PRIMARY

    tx_col2_desc = slide.shapes.add_textbox(Inches(5.7), col_y + Inches(0.2), Inches(3.2), Inches(0.4))
    p_col2_desc = tx_col2_desc.text_frame.paragraphs[0]
    p_col2_desc.text = "视觉流畅，让观众更专注于内容。"
    p_col2_desc.font.size = Pt(11)
    p_col2_desc.font.color.rgb = TEXT_BLACK

    # --- Column 3 ---
    # Icon 3: Browser/Window
    browser = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.0), col_y, Inches(0.5), Inches(0.4))
    browser.fill.background()
    browser.line.color.rgb = CUBE_LINE
    browser.line.width = Pt(1.5)
    
    browser_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.0), col_y + Inches(0.1), Inches(0.5), Inches(0.02))
    browser_line.fill.solid()
    browser_line.fill.fore_color.rgb = CUBE_LINE
    browser_line.line.fill.background()
    
    mag_glass = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.3), col_y + Inches(0.2), Inches(0.2), Inches(0.2))
    mag_glass.fill.background()
    mag_glass.line.color.rgb = HIGHLIGHT_LINE
    mag_glass.line.width = Pt(1.5)

    # Text 3
    tx_col3_title = slide.shapes.add_textbox(Inches(9.7), col_y - Inches(0.1), Inches(3.5), Inches(0.3))
    p_col3_title = tx_col3_title.text_frame.paragraphs[0]
    r1 = p_col3_title.add_run()
    r1.text = "3. "
    r1.font.bold = True
    r1.font.size = Pt(13)
    r2 = p_col3_title.add_run()
    r2.text = "专业感"
    r2.font.bold = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = BLUE_PRIMARY
    r3 = p_col3_title.add_run()
    r3.text = "源于对细节的严苛把控"
    r3.font.bold = True
    r3.font.size = Pt(13)

    tx_col3_desc = slide.shapes.add_textbox(Inches(9.7), col_y + Inches(0.2), Inches(3.5), Inches(0.4))
    p_col3_desc = tx_col3_desc.text_frame.paragraphs[0]
    p_col3_desc.text = "对齐、间距、字体、颜色的精准规范。"
    p_col3_desc.font.size = Pt(11)
    p_col3_desc.font.color.rgb = TEXT_BLACK