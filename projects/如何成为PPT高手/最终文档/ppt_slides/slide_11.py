def build_slide(slide):
    # 1. Background Grid (Top part - Light Blue)
    for i in range(1, 14):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(i), 0, Inches(i), Inches(7.5))
        line.line.color.rgb = RGBColor(0xE8, 0xF0, 0xF8)
        line.line.width = Pt(0.5)
    for i in range(1, 8):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, Inches(i), Inches(13.333), Inches(i))
        line.line.color.rgb = RGBColor(0xE8, 0xF0, 0xF8)
        line.line.width = Pt(0.5)

    # 2. Bottom Dark Blue Background
    bottom_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.2), Inches(13.333), Inches(3.3))
    bottom_rect.fill.solid()
    bottom_rect.fill.fore_color.rgb = RGBColor(0x15, 0x43, 0x85)
    bottom_rect.line.fill.background()

    # Bottom Grid (Overlay on dark blue)
    for i in range(1, 14):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(i), Inches(4.2), Inches(i), Inches(7.5))
        line.line.color.rgb = RGBColor(0x25, 0x53, 0x95)
        line.line.width = Pt(0.5)
    for i in range(5, 8):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, Inches(i), Inches(13.333), Inches(i))
        line.line.color.rgb = RGBColor(0x25, 0x53, 0x95)
        line.line.width = Pt(0.5)

    # 3. Title Text
    title_box = slide.shapes.add_textbox(Inches(2.66), Inches(0.8), Inches(8), Inches(1.2))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "感谢观看"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(64)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0x52, 0xCC)
    p.font.name = "Microsoft YaHei"

    subtitle_box = slide.shapes.add_textbox(Inches(2.66), Inches(2.1), Inches(8), Inches(0.8))
    tf_sub = subtitle_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "立即开始你的专业演示之旅"
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.font.size = Pt(28)
    p_sub.font.color.rgb = RGBColor(0x00, 0x52, 0xCC)
    p_sub.font.name = "Microsoft YaHei"

    # 4. Middle Icons and Text
    icon_color = RGBColor(0x00, 0x52, 0xCC)
    text_color = RGBColor(0x00, 0x00, 0x00)

    # Item 1: Compass
    compass = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.5), Inches(3.2), Inches(0.6), Inches(0.6))
    compass.fill.background()
    compass.line.color.rgb = icon_color
    compass.line.width = Pt(2)
    needle = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1.65), Inches(3.65), Inches(1.95), Inches(3.35))
    needle.line.color.rgb = icon_color
    needle.line.width = Pt(2)
    
    tb1 = slide.shapes.add_textbox(Inches(2.2), Inches(3.1), Inches(2.5), Inches(0.8))
    tf1 = tb1.text_frame
    p1 = tf1.paragraphs[0]
    p1.text = "1. 实践是提升PPT能\n力的唯一捷径"
    p1.font.size = Pt(16)
    p1.font.color.rgb = text_color
    p1.font.name = "Microsoft YaHei"

    # Item 2: Lightbulb
    bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.2), Inches(3.2), Inches(0.5), Inches(0.5))
    bulb.fill.background()
    bulb.line.color.rgb = icon_color
    bulb.line.width = Pt(2)
    base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.35), Inches(3.7), Inches(0.2), Inches(0.15))
    base.fill.background()
    base.line.color.rgb = icon_color
    base.line.width = Pt(2)
    ray1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.45), Inches(3.1), Inches(5.45), Inches(2.95))
    ray1.line.color.rgb = icon_color
    ray1.line.width = Pt(1.5)
    ray2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.05), Inches(3.45), Inches(4.9), Inches(3.45))
    ray2.line.color.rgb = icon_color
    ray2.line.width = Pt(1.5)
    ray3 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.85), Inches(3.45), Inches(6.0), Inches(3.45))
    ray3.line.color.rgb = icon_color
    ray3.line.width = Pt(1.5)
    
    tb2 = slide.shapes.add_textbox(Inches(5.9), Inches(3.1), Inches(2.5), Inches(0.8))
    tf2 = tb2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = "2. 保持好奇，持续优\n化视觉表达"
    p2.font.size = Pt(16)
    p2.font.color.rgb = text_color
    p2.font.name = "Microsoft YaHei"

    # Item 3: Rocket
    rocket = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.9), Inches(3.2), Inches(0.4), Inches(0.6))
    rocket.rotation = 45
    rocket.fill.background()
    rocket.line.color.rgb = icon_color
    rocket.line.width = Pt(2)
    wing1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(8.9), Inches(3.6), Inches(8.7), Inches(3.8))
    wing1.line.color.rgb = icon_color
    wing1.line.width = Pt(2)
    wing2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(9.2), Inches(3.7), Inches(9.4), Inches(3.9))
    wing2.line.color.rgb = icon_color
    wing2.line.width = Pt(2)
    
    tb3 = slide.shapes.add_textbox(Inches(9.6), Inches(3.2), Inches(2.5), Inches(0.8))
    tf3 = tb3.text_frame
    p3 = tf3.paragraphs[0]
    p3.text = "3. 期待您的精彩呈现"
    p3.font.size = Pt(16)
    p3.font.color.rgb = text_color
    p3.font.name = "Microsoft YaHei"

    # 5. Bottom Section - Faint Chart
    chart_color = RGBColor(0x4A, 0x76, 0xB5)
    
    # Chart Grid
    for x in [4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0]:
        v_line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(4.8), Inches(x), Inches(6.2))
        v_line.line.color.rgb = chart_color
        v_line.line.width = Pt(0.5)
        
    # X and Y axis
    x_axis = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(4.0), Inches(6.2), Inches(9.3), Inches(6.2))
    x_axis.line.color.rgb = chart_color
    x_axis.line.width = Pt(1)
    y_axis = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(4.0), Inches(4.8), Inches(4.0), Inches(6.2))
    y_axis.line.color.rgb = chart_color
    y_axis.line.width = Pt(1)
    
    # Trend line
    t1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(4.0), Inches(6.0), Inches(5.5), Inches(5.7))
    t1.line.color.rgb = chart_color
    t1.line.width = Pt(1.5)
    t2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.5), Inches(5.7), Inches(7.0), Inches(5.9))
    t2.line.color.rgb = chart_color
    t2.line.width = Pt(1.5)
    t3 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.0), Inches(5.9), Inches(9.0), Inches(4.8))
    t3.line.color.rgb = chart_color
    t3.line.width = Pt(1.5)

    # 6. "Thank You" Text
    ty_box = slide.shapes.add_textbox(Inches(2.66), Inches(4.6), Inches(8), Inches(1.5))
    tf_ty = ty_box.text_frame
    p_ty = tf_ty.paragraphs[0]
    p_ty.text = "Thank You"
    p_ty.alignment = PP_ALIGN.CENTER
    p_ty.font.size = Pt(72)
    p_ty.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p_ty.font.name = "Arial"

    # 7. Footer
    # QR Code Scanner Icon
    # Brackets
    brackets = [
        (3.5, 6.5, 3.6, 6.5), (3.5, 6.5, 3.5, 6.6), # Top-left
        (4.0, 6.5, 4.1, 6.5), (4.1, 6.5, 4.1, 6.6), # Top-right
        (3.5, 7.0, 3.6, 7.0), (3.5, 6.9, 3.5, 7.0), # Bottom-left
        (4.0, 7.0, 4.1, 7.0), (4.1, 6.9, 4.1, 7.0)  # Bottom-right
    ]
    for x1, y1, x2, y2 in brackets:
        bl = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        bl.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        bl.line.width = Pt(1.5)
        
    # Inner square and line
    sq = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.65), Inches(6.65), Inches(0.3), Inches(0.3))
    sq.fill.background()
    sq.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    sq.line.width = Pt(1)
    scan_line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(3.4), Inches(6.75), Inches(4.2), Inches(6.75))
    scan_line.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    scan_line.line.width = Pt(1.5)

    # Footer Text
    footer_box = slide.shapes.add_textbox(Inches(4.2), Inches(6.55), Inches(7.0), Inches(0.5))
    tf_footer = footer_box.text_frame
    p_footer = tf_footer.paragraphs[0]
    p_footer.text = "扫描二维码联系 | 联系方式: support@example.com | www.example.com"
    p_footer.font.size = Pt(11)
    p_footer.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p_footer.font.name = "Microsoft YaHei"

    # Page Number
    page_box = slide.shapes.add_textbox(Inches(11.8), Inches(6.8), Inches(1.0), Inches(0.5))
    tf_page = page_box.text_frame
    p_page = tf_page.paragraphs[0]
    p_page.text = "11 / 11"
    p_page.alignment = PP_ALIGN.RIGHT
    p_page.font.size = Pt(12)
    p_page.font.color.rgb = RGBColor(0x8A, 0xB4, 0xF8)
    p_page.font.name = "Arial"