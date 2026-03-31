def build_slide(slide):
    # --- Colors ---
    BG_COLOR = RGBColor(245, 242, 235)
    TEXT_BLACK = RGBColor(20, 20, 20)
    ORANGE = RGBColor(244, 122, 32)
    GREEN = RGBColor(0, 168, 112)
    DARK_BLUE = RGBColor(20, 40, 80)
    LIGHT_TEAL = RGBColor(160, 220, 200)
    WHITE = RGBColor(255, 255, 255)

    # Set background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # --- Edge Decorations (Torn Paper Effect) ---
    # Top Left
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(-1.5), Inches(-1.5), Inches(3), Inches(3))
    shape.rotation = 45
    shape.fill.solid()
    shape.fill.fore_color.rgb = GREEN
    shape.line.fill.background()

    # Top Right
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.5), Inches(-1), Inches(3), Inches(2))
    shape.rotation = -20
    shape.fill.solid()
    shape.fill.fore_color.rgb = ORANGE
    shape.line.fill.background()
    
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12), Inches(0), Inches(2), Inches(3))
    shape.rotation = 10
    shape.fill.solid()
    shape.fill.fore_color.rgb = DARK_BLUE
    shape.line.fill.background()

    # Bottom Left
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(-1), Inches(6.5), Inches(3), Inches(2))
    shape.rotation = 15
    shape.fill.solid()
    shape.fill.fore_color.rgb = ORANGE
    shape.line.fill.background()

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(-0.5), Inches(7), Inches(3), Inches(2))
    shape.rotation = -10
    shape.fill.solid()
    shape.fill.fore_color.rgb = TEXT_BLACK
    shape.line.fill.background()

    # --- Main Title ---
    # Line 1
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(8), Inches(1))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "桂林山水“甲”天下，"
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.name = "Microsoft YaHei"
    p.font.color.rgb = TEXT_BLACK

    # Highlight under Line 2
    highlight = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(2.6), Inches(6.2), Inches(0.25))
    highlight.fill.solid()
    highlight.fill.fore_color.rgb = ORANGE
    highlight.line.fill.background()

    # Line 2
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(8), Inches(1))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "我的脑洞“假”不了"
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.name = "Microsoft YaHei"
    p.font.color.rgb = TEXT_BLACK

    # --- Subtitle ---
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(8), Inches(0.6))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "一个猎奇博主的桂林“真香”探险报告"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.name = "Microsoft YaHei"
    p.font.color.rgb = TEXT_BLACK

    # --- 20 RMB Graphic ---
    rmb_left = Inches(8.0)
    rmb_top = Inches(0.6)
    
    # Base
    rmb_base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, rmb_left, rmb_top, Inches(4.8), Inches(2.5))
    rmb_base.fill.solid()
    rmb_base.fill.fore_color.rgb = WHITE
    rmb_base.line.color.rgb = DARK_BLUE
    rmb_base.line.width = Pt(2)

    # Inner fill
    rmb_inner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, rmb_left + Inches(0.1), rmb_top + Inches(0.1), Inches(4.6), Inches(2.3))
    rmb_inner.fill.solid()
    rmb_inner.fill.fore_color.rgb = LIGHT_TEAL
    rmb_inner.line.fill.background()

    # Sun
    sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, rmb_left + Inches(2.0), rmb_top + Inches(0.3), Inches(0.6), Inches(0.6))
    sun.fill.solid()
    sun.fill.fore_color.rgb = WHITE
    sun.line.fill.background()

    # Mountains (Triangles)
    m1 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, rmb_left + Inches(0.5), rmb_top + Inches(0.8), Inches(1.0), Inches(1.2))
    m1.fill.solid()
    m1.fill.fore_color.rgb = DARK_BLUE
    m1.line.fill.background()

    m2 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, rmb_left + Inches(1.2), rmb_top + Inches(0.6), Inches(1.2), Inches(1.4))
    m2.fill.solid()
    m2.fill.fore_color.rgb = GREEN
    m2.line.fill.background()

    m3 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, rmb_left + Inches(2.5), rmb_top + Inches(0.9), Inches(1.0), Inches(1.1))
    m3.fill.solid()
    m3.fill.fore_color.rgb = GREEN
    m3.line.fill.background()

    # River (Rectangle at bottom)
    river = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, rmb_left + Inches(0.1), rmb_top + Inches(1.8), Inches(4.6), Inches(0.6))
    river.fill.solid()
    river.fill.fore_color.rgb = DARK_BLUE
    river.line.fill.background()

    # Text "20" Top Left
    txBox = slide.shapes.add_textbox(rmb_left + Inches(0.2), rmb_top + Inches(0.1), Inches(1), Inches(0.5))
    p = txBox.text_frame.paragraphs[0]
    p.text = "20"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE

    # Text "20 RMB" Bottom Left
    txBox = slide.shapes.add_textbox(rmb_left + Inches(0.2), rmb_top + Inches(1.8), Inches(1.5), Inches(0.5))
    p = txBox.text_frame.paragraphs[0]
    p.text = "20 RMB"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = WHITE

    # Text "中国人民银行" Top Right
    txBox = slide.shapes.add_textbox(rmb_left + Inches(2.5), rmb_top + Inches(0.1), Inches(2), Inches(0.4))
    p = txBox.text_frame.paragraphs[0]
    p.text = "中国人民银行"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE

    # --- Labels & Character (Bottom Left) ---
    # Character Placeholder
    char_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.0), Inches(4.5), Inches(1.5), Inches(2.5))
    char_bg.fill.solid()
    char_bg.fill.fore_color.rgb = ORANGE
    char_bg.line.color.rgb = TEXT_BLACK
    char_bg.line.width = Pt(2)
    p = char_bg.text_frame.paragraphs[0]
    p.text = "猎奇\n博主"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    def add_label(text, left, top, width, height, rotation, border_color):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.rotation = rotation
        shape.fill.solid()
        shape.fill.fore_color.rgb = WHITE
        shape.line.color.rgb = border_color
        shape.line.width = Pt(2.5)
        
        tf = shape.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.name = "Microsoft YaHei"
        p.font.color.rgb = TEXT_BLACK
        p.alignment = PP_ALIGN.CENTER
        tf.margin_top = Pt(3)
        tf.margin_bottom = Pt(3)
        tf.margin_left = Pt(5)
        tf.margin_right = Pt(5)

    add_label("奇葩角落", Inches(1.2), Inches(4.0), Inches(1.4), Inches(0.4), -15, GREEN)
    add_label("真香！！", Inches(1.2), Inches(5.0), Inches(1.4), Inches(0.4), -5, ORANGE)
    add_label("不错的青", Inches(1.5), Inches(6.0), Inches(1.4), Inches(0.4), 10, DARK_BLUE)
    
    add_label("拒绝中老年团", Inches(4.8), Inches(4.2), Inches(2.0), Inches(0.4), -15, ORANGE)
    add_label("冷知识", Inches(5.2), Inches(5.2), Inches(1.2), Inches(0.4), 5, GREEN)
    add_label("美景背后", Inches(5.0), Inches(6.0), Inches(1.4), Inches(0.4), 15, DARK_BLUE)

    # Add some arrows/lines pointing to center
    line1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(2.5), Inches(4.2), Inches(2.9), Inches(4.8))
    line1.line.color.rgb = TEXT_BLACK
    line1.line.width = Pt(2)
    
    line2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(4.8), Inches(4.6), Inches(4.5), Inches(5.0))
    line2.line.color.rgb = TEXT_BLACK
    line2.line.width = Pt(2)

    # --- Right List ---
    def add_list_item(text, left, top, bg_color, icon_border_color, icon_type):
        # Background Parallelogram
        bg = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, left + Inches(0.4), top, Inches(4.5), Inches(0.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = bg_color
        bg.line.fill.background()

        # Text
        txBox = slide.shapes.add_textbox(left + Inches(0.8), top - Inches(0.05), Inches(4.0), Inches(0.6))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.name = "Microsoft YaHei"
        p.font.color.rgb = TEXT_BLACK

        # Icon Circle
        icon_bg = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top - Inches(0.1), Inches(0.7), Inches(0.7))
        icon_bg.fill.solid()
        icon_bg.fill.fore_color.rgb = WHITE
        icon_bg.line.color.rgb = icon_border_color
        icon_bg.line.width = Pt(2.5)

        # Simple icon drawing inside the circle
        if icon_type == 'bus':
            bus = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.15), top + Inches(0.1), Inches(0.4), Inches(0.3))
            bus.fill.solid()
            bus.fill.fore_color.rgb = DARK_BLUE
            bus.line.fill.background()
            cross = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left + Inches(0.1), top, left + Inches(0.6), top + Inches(0.5))
            cross.line.color.rgb = ORANGE
            cross.line.width = Pt(3)
        elif icon_type == 'mag':
            mag_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.15), top + Inches(0.05), Inches(0.3), Inches(0.3))
            mag_circle.fill.background()
            mag_circle.line.color.rgb = GREEN
            mag_circle.line.width = Pt(2)
            handle = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left + Inches(0.4), top + Inches(0.3), left + Inches(0.55), top + Inches(0.45))
            handle.line.color.rgb = ORANGE
            handle.line.width = Pt(3)
        elif icon_type == 'person':
            head = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.25), top + Inches(0.05), Inches(0.2), Inches(0.2))
            head.fill.solid()
            head.fill.fore_color.rgb = ORANGE
            head.line.fill.background()
            body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.15), top + Inches(0.3), Inches(0.4), Inches(0.2))
            body.fill.solid()
            body.fill.fore_color.rgb = GREEN
            body.line.fill.background()

    add_list_item("拒绝传统中老年旅行团画风", Inches(7.8), Inches(4.0), GREEN, ORANGE, 'bus')
    add_list_item("深度挖掘桂林不为人知的“奇葩”角落", Inches(7.8), Inches(5.0), GREEN, ORANGE, 'mag')
    add_list_item("搞笑博主的生存视角：\n美景背后的冷知识", Inches(7.8), Inches(6.0), ORANGE, GREEN, 'person')

    # --- Page Number ---
    # Green background circle
    circle_bg = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(12.1), Inches(6.8), Inches(0.4), Inches(0.4))
    circle_bg.fill.solid()
    circle_bg.fill.fore_color.rgb = GREEN
    circle_bg.line.fill.background()

    # Orange Circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(12.2), Inches(6.8), Inches(0.4), Inches(0.4))
    circle.fill.solid()
    circle.fill.fore_color.rgb = ORANGE
    circle.line.fill.background()

    # Text
    txBox = slide.shapes.add_textbox(Inches(12.4), Inches(6.7), Inches(0.6), Inches(0.6))
    p = txBox.text_frame.paragraphs[0]
    p.text = "01"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.name = "Microsoft YaHei"
    p.font.color.rgb = TEXT_BLACK