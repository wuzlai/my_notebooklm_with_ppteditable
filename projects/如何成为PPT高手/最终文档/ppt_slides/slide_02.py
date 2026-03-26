def build_slide(slide):
    # Colors
    TITLE_BLUE = RGBColor(0x1B, 0x5E, 0xB8)
    SUBTITLE_GRAY = RGBColor(0x40, 0x40, 0x40)
    LINE_BLUE = RGBColor(0x4A, 0x86, 0xC8)
    HIGHLIGHT_ORANGE = RGBColor(0xDF, 0x9A, 0x2A)
    BOX_BORDER = RGBColor(0xE0, 0xE0, 0xE0)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    BLACK = RGBColor(0x20, 0x20, 0x20)

    # Add Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(10.0), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "目录：构建专业PPT的蓝图"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = TITLE_BLUE

    # Add Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(10.0), Inches(0.5))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = "本次分享的核心框架"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = SUBTITLE_GRAY

    # Helper function to draw connecting lines
    def add_line(left, top, width, height):
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        line.fill.solid()
        line.fill.fore_color.rgb = LINE_BLUE
        line.line.fill.background()

    # Helper function to draw junction circles
    def add_circle(cx, cy):
        r = 0.06
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx-r), Inches(cy-r), Inches(r*2), Inches(r*2))
        circle.fill.solid()
        circle.fill.fore_color.rgb = WHITE
        circle.line.color.rgb = LINE_BLUE
        circle.line.width = Pt(1.5)

    # Helper function to draw simple icons
    def add_icon(shape_type, left, top, width=0.4, height=0.4):
        icon = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
        icon.fill.background()
        icon.line.color.rgb = LINE_BLUE
        icon.line.width = Pt(1.5)

    # Helper function to draw text boxes with highlighted text
    def add_node_box(left, top, text_parts):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(5.8), Inches(0.8))
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = BOX_BORDER
        box.line.width = Pt(1)
        
        tf = box.text_frame
        tf.vertical_anchor = 3  # MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.2)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        
        for text, color in text_parts:
            run = p.add_run()
            run.text = text
            run.font.name = "Microsoft YaHei"
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = color

    # Draw Tree Lines (Staggered structure)
    line_thickness = 0.025
    
    # Node 1 horizontal connection
    add_line(1.8, 2.8 - line_thickness/2, 0.4, line_thickness)
    
    # Vertical line 1 (Node 1 to Node 2 level)
    add_line(1.8 - line_thickness/2, 2.8, line_thickness, 1.2)
    
    # Horizontal line 2 (Node 2 level)
    add_line(1.8, 4.0 - line_thickness/2, 1.6, line_thickness)
    
    # Vertical line 2 (Node 2 to Node 3 level)
    add_line(3.0 - line_thickness/2, 4.0, line_thickness, 1.2)
    
    # Horizontal line 3 (Node 3 level)
    add_line(3.0, 5.2 - line_thickness/2, 1.6, line_thickness)
    
    # Vertical line 3 (Node 3 to Node 4 level)
    add_line(4.2 - line_thickness/2, 5.2, line_thickness, 1.2)
    
    # Horizontal line 4 (Node 4 level)
    add_line(4.2, 6.4 - line_thickness/2, 1.6, line_thickness)

    # Draw Icons
    add_icon(MSO_SHAPE.DOCUMENT, 1.0, 2.5, 0.5, 0.6)
    add_icon(MSO_SHAPE.CAN, 2.4, 3.75, 0.4, 0.5)
    add_icon(MSO_SHAPE.SUN, 3.6, 4.95, 0.5, 0.5)
    add_icon(MSO_SHAPE.ISOSCELES_TRIANGLE, 4.8, 6.15, 0.5, 0.5)

    # Draw Junction Circles
    add_circle(1.8, 2.8)
    add_circle(1.8, 4.0)
    add_circle(3.0, 5.2)
    add_circle(4.2, 6.4)

    # Draw Text Boxes
    add_node_box(2.2, 2.4, [
        ("1. 内容法则：一页一事，", BLACK), 
        ("结论先行", HIGHLIGHT_ORANGE)
    ])
    
    add_node_box(3.4, 3.6, [
        ("2. 减法艺术：拒绝文字堆砌，追求", BLACK), 
        ("秒懂", HIGHLIGHT_ORANGE)
    ])
    
    add_node_box(4.6, 4.8, [
        ("3. 设计规范：高度统一，建立", BLACK), 
        ("专业感", HIGHLIGHT_ORANGE)
    ])
    
    add_node_box(5.8, 6.0, [
        ("4. 高手境界：简洁有力的视觉哲学", BLACK)
    ])

    # Add Page Number
    page_num = slide.shapes.add_textbox(Inches(12.5), Inches(6.8), Inches(0.5), Inches(0.5))
    p = page_num.text_frame.paragraphs[0]
    p.text = "02"
    p.font.name = "Arial"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(0x80, 0x80, 0x80)