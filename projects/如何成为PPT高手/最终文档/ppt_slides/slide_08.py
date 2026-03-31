def build_slide(slide):
    # Colors
    BLUE_TITLE = RGBColor(0x1A, 0x66, 0xCC)
    DARK_TEXT = RGBColor(0x22, 0x22, 0x22)
    GRAY_TEXT = RGBColor(0x66, 0x66, 0x66)
    BLUE_LINE = RGBColor(0x5B, 0x9B, 0xD5)
    LIGHT_BLUE_FILL = RGBColor(0xE6, 0xF0, 0xFA)
    GREEN_OK = RGBColor(0x4C, 0xAF, 0x50)
    RED_ERR = RGBColor(0xF4, 0x43, 0x36)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    
    # 1. Title and Subtitle
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(8.0), Inches(0.8))
    p_title = tb_title.text_frame.paragraphs[0]
    
    run1 = p_title.add_run()
    run1.text = "排版逻辑："
    run1.font.size = Pt(32)
    run1.font.bold = True
    run1.font.color.rgb = BLUE_TITLE
    run1.font.name = "Microsoft YaHei"
    
    run2 = p_title.add_run()
    run2.text = "始终如一的风格表达"
    run2.font.size = Pt(32)
    run2.font.bold = True
    run2.font.color.rgb = DARK_TEXT
    run2.font.name = "Microsoft YaHei"
    
    tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(8.0), Inches(0.5))
    p_sub = tb_sub.text_frame.paragraphs[0]
    p_sub.text = "规范化的布局让阅读更顺畅"
    p_sub.font.size = Pt(18)
    p_sub.font.color.rgb = GRAY_TEXT
    p_sub.font.name = "Microsoft YaHei"

    # 2. Left Section: Alignment & Margins
    # Icon (Grid)
    icon_grid = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.2), Inches(0.6), Inches(0.45))
    icon_grid.fill.solid()
    icon_grid.fill.fore_color.rgb = LIGHT_BLUE_FILL
    icon_grid.line.color.rgb = BLUE_LINE
    icon_grid.line.width = Pt(1.5)
    
    line_v = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.95), Inches(2.1), Inches(0.95), Inches(2.7))
    line_v.line.color.rgb = BLUE_LINE
    line_h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.7), Inches(2.4), Inches(1.5), Inches(2.4))
    line_h.line.color.rgb = BLUE_LINE

    # Heading
    tb_h1 = slide.shapes.add_textbox(Inches(1.6), Inches(2.15), Inches(4.0), Inches(0.5))
    p_h1 = tb_h1.text_frame.paragraphs[0]
    p_h1.text = "建立统一的页边距与对齐基准"
    p_h1.font.size = Pt(18)
    p_h1.font.bold = True
    p_h1.font.name = "Microsoft YaHei"

    # Wireframe Graphic
    # Outer Box
    box_outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.7), Inches(3.1), Inches(3.3), Inches(2.2))
    box_outer.fill.solid()
    box_outer.fill.fore_color.rgb = LIGHT_BLUE_FILL
    box_outer.line.color.rgb = BLUE_LINE
    
    # Inner Dashed Box
    box_inner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.9), Inches(3.3), Inches(2.9), Inches(1.8))
    box_inner.fill.background()
    box_inner.line.color.rgb = BLUE_LINE
    box_inner.line.dash_style = 3 # Dashed
    
    # Content Blocks inside Wireframe
    rect_title = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.95), Inches(3.5), Inches(0.9), Inches(0.2))
    rect_title.fill.solid()
    rect_title.fill.fore_color.rgb = BLUE_TITLE
    rect_title.line.fill.background()
    
    for i in range(3):
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.95), Inches(3.8 + i*0.15), Inches(1.3), Inches(0.05))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
        line.line.fill.background()
        
    rect_img = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.4), Inches(3.5), Inches(1.35), Inches(0.9))
    rect_img.fill.background()
    rect_img.line.color.rgb = BLUE_LINE
    
    # Mountain placeholder inside rect_img
    tri1 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(3.45), Inches(3.9), Inches(0.6), Inches(0.5))
    tri1.fill.solid()
    tri1.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
    tri1.line.fill.background()
    
    tri2 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(3.85), Inches(4.0), Inches(0.5), Inches(0.4))
    tri2.fill.solid()
    tri2.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
    tri2.line.fill.background()
    
    sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.15), Inches(3.65), Inches(0.15), Inches(0.15))
    sun.fill.solid()
    sun.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
    sun.line.fill.background()

    # Alignment Guides
    guide_v = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(3.35), Inches(2.9), Inches(3.35), Inches(5.5))
    guide_v.line.color.rgb = BLUE_TITLE
    guide_h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1.5), Inches(4.05), Inches(5.2), Inches(4.05))
    guide_h.line.color.rgb = BLUE_TITLE

    # Labels for Alignment
    lbl_align_l = slide.shapes.add_textbox(Inches(0.7), Inches(3.9), Inches(1.0), Inches(0.3))
    lbl_align_l.text_frame.text = "对齐基准"
    lbl_align_l.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_align_l.text_frame.paragraphs[0].font.color.rgb = BLUE_TITLE

    lbl_align_b = slide.shapes.add_textbox(Inches(3.0), Inches(5.6), Inches(1.0), Inches(0.3))
    lbl_align_b.text_frame.text = "对齐基准"
    lbl_align_b.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_align_b.text_frame.paragraphs[0].font.color.rgb = BLUE_TITLE

    lbl_margin = slide.shapes.add_textbox(Inches(5.2), Inches(3.9), Inches(1.0), Inches(0.3))
    lbl_margin.text_frame.text = "页边距"
    lbl_margin.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_margin.text_frame.paragraphs[0].font.color.rgb = BLUE_TITLE

    # 3. Top Right Section: Icon Consistency
    # Icon (Four squares/circles)
    for r in range(2):
        for c in range(2):
            sq = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6 + c*0.25), Inches(2.2 + r*0.25), Inches(0.2), Inches(0.2))
            sq.fill.solid()
            sq.fill.fore_color.rgb = LIGHT_BLUE_FILL
            sq.line.color.rgb = BLUE_LINE

    # Heading
    tb_h2 = slide.shapes.add_textbox(Inches(7.2), Inches(2.15), Inches(5.5), Inches(0.5))
    p_h2 = tb_h2.text_frame.paragraphs[0]
    p_h2.text = "保持图标风格一致（全线框或全色块）"
    p_h2.font.size = Pt(18)
    p_h2.font.bold = True
    p_h2.font.name = "Microsoft YaHei"

    # Correct Icons (Outline)
    icon_correct = slide.shapes.add_textbox(Inches(7.2), Inches(2.9), Inches(2.5), Inches(0.8))
    p_ic = icon_correct.text_frame.paragraphs[0]
    p_ic.text = "⚙  💡  📄"
    p_ic.font.size = Pt(36)
    p_ic.font.color.rgb = BLUE_TITLE
    
    # Correct Label
    chk_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.6), Inches(4.0), Inches(0.2), Inches(0.2))
    chk_circle.fill.solid()
    chk_circle.fill.fore_color.rgb = GREEN_OK
    chk_circle.line.fill.background()
    
    lbl_correct = slide.shapes.add_textbox(Inches(7.8), Inches(3.9), Inches(1.5), Inches(0.3))
    lbl_correct.text_frame.text = "正确（一致）"
    lbl_correct.text_frame.paragraphs[0].font.size = Pt(14)

    # Incorrect Icons (Mixed)
    icon_incorrect = slide.shapes.add_textbox(Inches(10.0), Inches(2.9), Inches(2.5), Inches(0.8))
    p_ii = icon_incorrect.text_frame.paragraphs[0]
    p_ii.text = "📢  ✋  ☁"
    p_ii.font.size = Pt(36)
    p_ii.font.color.rgb = BLUE_TITLE

    # Incorrect Label
    err_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.4), Inches(4.0), Inches(0.2), Inches(0.2))
    err_circle.fill.solid()
    err_circle.fill.fore_color.rgb = RED_ERR
    err_circle.line.fill.background()
    
    lbl_incorrect = slide.shapes.add_textbox(Inches(10.6), Inches(3.9), Inches(1.5), Inches(0.3))
    lbl_incorrect.text_frame.text = "错误（混杂）"
    lbl_incorrect.text_frame.paragraphs[0].font.size = Pt(14)

    # 4. Bottom Right Section: Whitespace
    # Icon (Document)
    icon_doc = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.6), Inches(4.7), Inches(0.45), Inches(0.5))
    icon_doc.fill.solid()
    icon_doc.fill.fore_color.rgb = LIGHT_BLUE_FILL
    icon_doc.line.color.rgb = BLUE_LINE
    for i in range(3):
        dl = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.7), Inches(4.85 + i*0.1), Inches(0.25), Inches(0.03))
        dl.fill.solid()
        dl.fill.fore_color.rgb = BLUE_LINE
        dl.line.fill.background()

    # Heading
    tb_h3 = slide.shapes.add_textbox(Inches(7.2), Inches(4.75), Inches(5.0), Inches(0.5))
    p_h3 = tb_h3.text_frame.paragraphs[0]
    p_h3.text = "留白艺术：给内容呼吸的空间"
    p_h3.font.size = Pt(18)
    p_h3.font.bold = True
    p_h3.font.name = "Microsoft YaHei"

    # Good Layout Graphic (Whitespace)
    box_ws_outer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.3), Inches(5.5), Inches(2.6), Inches(1.8))
    box_ws_outer.fill.solid()
    box_ws_outer.fill.fore_color.rgb = RGBColor(0xF0, 0xF8, 0xFF)
    box_ws_outer.line.color.rgb = BLUE_LINE
    box_ws_outer.line.dash_style = 3
    
    # Diagonal lines for whitespace indication
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.3), Inches(5.5), Inches(7.7), Inches(5.9)).line.color.rgb = RGBColor(0xCC, 0xDD, 0xEE)
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(9.9), Inches(5.5), Inches(9.5), Inches(5.9)).line.color.rgb = RGBColor(0xCC, 0xDD, 0xEE)
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.3), Inches(7.3), Inches(7.7), Inches(6.9)).line.color.rgb = RGBColor(0xCC, 0xDD, 0xEE)
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(9.9), Inches(7.3), Inches(9.5), Inches(6.9)).line.color.rgb = RGBColor(0xCC, 0xDD, 0xEE)

    box_ws_inner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.8), Inches(5.9), Inches(1.6), Inches(1.0))
    box_ws_inner.fill.solid()
    box_ws_inner.fill.fore_color.rgb = WHITE
    box_ws_inner.line.fill.background()
    
    # Shadow effect simulation
    box_ws_inner_shadow = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.78), Inches(5.88), Inches(1.64), Inches(1.04))
    box_ws_inner_shadow.fill.background()
    box_ws_inner_shadow.line.color.rgb = RGBColor(0xEE, 0xEE, 0xEE)
    
    tb_core = slide.shapes.add_textbox(Inches(7.8), Inches(6.1), Inches(1.6), Inches(0.4))
    p_core = tb_core.text_frame.paragraphs[0]
    p_core.text = "核心内容"
    p_core.font.size = Pt(16)
    p_core.font.bold = True
    p_core.alignment = PP_ALIGN.CENTER
    
    core_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.2), Inches(6.55), Inches(0.8), Inches(0.05))
    core_line.fill.solid()
    core_line.fill.fore_color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    core_line.line.fill.background()

    # Whitespace Labels
    lbl_ws_t = slide.shapes.add_textbox(Inches(8.3), Inches(5.55), Inches(0.8), Inches(0.2))
    lbl_ws_t.text_frame.text = "呼吸空间"
    lbl_ws_t.text_frame.paragraphs[0].font.size = Pt(9)
    lbl_ws_t.text_frame.paragraphs[0].font.color.rgb = BLUE_TITLE
    
    lbl_ws_b = slide.shapes.add_textbox(Inches(8.3), Inches(7.0), Inches(0.8), Inches(0.2))
    lbl_ws_b.text_frame.text = "呼吸空间"
    lbl_ws_b.text_frame.paragraphs[0].font.size = Pt(9)
    lbl_ws_b.text_frame.paragraphs[0].font.color.rgb = BLUE_TITLE

    lbl_ws_l = slide.shapes.add_textbox(Inches(7.35), Inches(6.3), Inches(0.8), Inches(0.2))
    lbl_ws_l.text_frame.text = "呼吸空间"
    lbl_ws_l.text_frame.paragraphs[0].font.size = Pt(9)
    lbl_ws_l.text_frame.paragraphs[0].font.color.rgb = BLUE_TITLE

    lbl_ws_r = slide.shapes.add_textbox(Inches(9.45), Inches(6.3), Inches(0.8), Inches(0.2))
    lbl_ws_r.text_frame.text = "呼吸空间"
    lbl_ws_r.text_frame.paragraphs[0].font.size = Pt(9)
    lbl_ws_r.text_frame.paragraphs[0].font.color.rgb = BLUE_TITLE

    # Bad Layout Graphic (Cluttered)
    box_cl_outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.4), Inches(5.5), Inches(1.5), Inches(1.0))
    box_cl_outer.fill.solid()
    box_cl_outer.fill.fore_color.rgb = LIGHT_BLUE_FILL
    box_cl_outer.line.color.rgb = BLUE_LINE

    # Cluttered inner elements
    cl_title = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.45), Inches(5.55), Inches(0.6), Inches(0.15))
    cl_title.fill.solid()
    cl_title.fill.fore_color.rgb = BLUE_TITLE
    cl_title.line.fill.background()
    
    for i in range(4):
        cl_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.45), Inches(5.75 + i*0.1), Inches(0.75), Inches(0.05))
        cl_line.fill.solid()
        cl_line.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
        cl_line.line.fill.background()
        
    cl_img = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.3), Inches(5.55), Inches(0.55), Inches(0.4))
    cl_img.fill.solid()
    cl_img.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
    cl_img.line.color.rgb = BLUE_LINE
    
    cl_box2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.45), Inches(6.2), Inches(0.75), Inches(0.25))
    cl_box2.fill.solid()
    cl_box2.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
    cl_box2.line.color.rgb = BLUE_LINE

    cl_box3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.3), Inches(6.0), Inches(0.55), Inches(0.45))
    cl_box3.fill.solid()
    cl_box3.fill.fore_color.rgb = RGBColor(0xBD, 0xD0, 0xE6)
    cl_box3.line.color.rgb = BLUE_LINE

    # Cluttered Label
    err_circle2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.6), Inches(6.75), Inches(0.2), Inches(0.2))
    err_circle2.fill.solid()
    err_circle2.fill.fore_color.rgb = RED_ERR
    err_circle2.line.fill.background()
    
    lbl_cluttered = slide.shapes.add_textbox(Inches(10.8), Inches(6.65), Inches(1.5), Inches(0.3))
    lbl_cluttered.text_frame.text = "拥挤布局"
    lbl_cluttered.text_frame.paragraphs[0].font.size = Pt(14)
    lbl_cluttered.text_frame.paragraphs[0].font.bold = True

    # 5. Page Number
    tb_page = slide.shapes.add_textbox(Inches(11.8), Inches(6.8), Inches(1.2), Inches(0.5))
    p_page = tb_page.text_frame.paragraphs[0]
    p_page.text = "08 / 11"
    p_page.font.size = Pt(20)
    p_page.font.bold = True
    p_page.font.color.rgb = GRAY_TEXT
    p_page.font.name = "Microsoft YaHei"
    p_page.alignment = PP_ALIGN.RIGHT