def build_slide(slide):
    # 1. Background
    # Left background (Light beige)
    bg_left = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg_left.fill.solid()
    bg_left.fill.fore_color.rgb = RGBColor(0xF4, 0xF1, 0xEA)
    bg_left.line.fill.background()

    # Right background (Orange, angled)
    bg_right = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(-3), Inches(10), Inches(14))
    bg_right.rotation = 12
    bg_right.fill.solid()
    bg_right.fill.fore_color.rgb = RGBColor(0xFF, 0x7A, 0x00)
    bg_right.line.fill.background()

    # 2. Top Left Tag ("第11页")
    tag_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.4), Inches(1.5), Inches(0.6))
    tag_box.fill.solid()
    tag_box.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
    tag_box.line.fill.background()
    
    tf_tag = tag_box.text_frame
    tf_tag.clear()
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = "第11页"
    p_tag.font.name = "Microsoft YaHei"
    p_tag.font.size = Pt(20)
    p_tag.font.bold = True
    p_tag.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p_tag.alignment = PP_ALIGN.CENTER

    # 3. Titles
    # Main Title Background
    title1_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.8), Inches(0.5), Inches(9.5), Inches(1.0))
    title1_bg.rotation = -2
    title1_bg.fill.solid()
    title1_bg.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    title1_bg.line.fill.background()

    # Main Title Text
    title1_tx = slide.shapes.add_textbox(Inches(2.9), Inches(0.5), Inches(9.3), Inches(1.0))
    title1_tx.rotation = -2
    tf1 = title1_tx.text_frame
    tf1.clear()
    p1 = tf1.paragraphs[0]
    p1.text = "感谢观看，记得给个“五星好评”"
    p1.font.name = "Microsoft YaHei"
    p1.font.size = Pt(40)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(0xE6, 0x4A, 0x19) # Deep Orange

    # Subtitle Background
    title2_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(1.5), Inches(7.5), Inches(0.7))
    title2_bg.fill.solid()
    title2_bg.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    title2_bg.line.fill.background()

    # Subtitle Text
    title2_tx = slide.shapes.add_textbox(Inches(5.8), Inches(1.55), Inches(7.0), Inches(0.6))
    tf2 = title2_tx.text_frame
    tf2.clear()
    p2 = tf2.paragraphs[0]
    p2.text = "关注我，带你解锁更多奇葩目的地"
    p2.font.name = "Microsoft YaHei"
    p2.font.size = Pt(26)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # 4. Left Visuals (Mountains & River Placeholder)
    # Mountains (Overlapping green triangles)
    colors_mtn = [RGBColor(0x81, 0xC7, 0x84), RGBColor(0xA5, 0xD6, 0xA7), RGBColor(0x66, 0xBB, 0x6A)]
    mtn_coords = [(0.2, 3.5, 2.5, 4.0), (1.5, 2.5, 2.5, 5.0), (3.0, 3.5, 2.5, 4.0)]
    
    for i, (x, y, w, h) in enumerate(mtn_coords):
        mtn = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        mtn.fill.solid()
        mtn.fill.fore_color.rgb = colors_mtn[i]
        mtn.line.color.rgb = RGBColor(0x1B, 0x5E, 0x20)
        mtn.line.width = Pt(2)

    # River (Curved shape)
    river = slide.shapes.add_shape(MSO_SHAPE.MOON, Inches(1.0), Inches(5.0), Inches(4.5), Inches(2.5))
    river.rotation = 50
    river.fill.solid()
    river.fill.fore_color.rgb = RGBColor(0xE0, 0xF2, 0xF1)
    river.line.color.rgb = RGBColor(0x00, 0x4D, 0x40)
    river.line.width = Pt(2)

    # Decorations (Stars)
    star_coords = [(0.5, 2.0), (1.0, 2.8), (4.8, 4.2), (1.2, 6.5)]
    for x, y in star_coords:
        star = slide.shapes.add_shape(MSO_SHAPE.5_POINT_STAR, Inches(x), Inches(y), Inches(0.4), Inches(0.4))
        star.rotation = 15
        star.fill.background()
        star.line.color.rgb = RGBColor(0x9C, 0x27, 0xB0)
        star.line.width = Pt(2)

    # Text "BOOM!"
    boom_tx = slide.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(1.2), Inches(0.5))
    boom_tx.rotation = -15
    p_boom = boom_tx.text_frame.paragraphs[0]
    p_boom.text = "BOOM!"
    p_boom.font.name = "Arial"
    p_boom.font.size = Pt(18)
    p_boom.font.bold = True
    p_boom.font.color.rgb = RGBColor(0x9C, 0x27, 0xB0)

    # Text "OMG!"
    omg_coords = [(4.0, 2.2), (4.8, 3.2), (4.5, 5.5)]
    for x, y in omg_coords:
        omg_tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(1.0), Inches(0.5))
        omg_tx.rotation = -10
        p_omg = omg_tx.text_frame.paragraphs[0]
        p_omg.text = "OMG!"
        p_omg.font.name = "Arial"
        p_omg.font.size = Pt(16)
        p_omg.font.bold = True
        p_omg.font.color.rgb = RGBColor(0xF4, 0x43, 0x36)

    # 5. Center Visual (Silhouette Placeholder)
    # Head
    head = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.5), Inches(2.6), Inches(0.8), Inches(1.0))
    head.fill.solid()
    head.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
    head.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    head.line.width = Pt(4)

    # Body
    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(3.5), Inches(2.0), Inches(3.5))
    body.fill.solid()
    body.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
    body.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    body.line.width = Pt(4)

    # Arm (waving)
    arm = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(2.8), Inches(0.6), Inches(1.5))
    arm.rotation = -30
    arm.fill.solid()
    arm.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
    arm.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    arm.line.width = Pt(4)

    # 6. Right Visual (QR Code Area)
    # Sticky Note (Yellow)
    note = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.0), Inches(1.8), Inches(3.8), Inches(3.8))
    note.rotation = 3
    note.fill.solid()
    note.fill.fore_color.rgb = RGBColor(0xFF, 0xEB, 0x3B)
    note.line.fill.background()

    # QR Code Base (Orange)
    qr_base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.4), Inches(2.2), Inches(3.0), Inches(3.0))
    qr_base.rotation = 3
    qr_base.fill.solid()
    qr_base.fill.fore_color.rgb = RGBColor(0xFF, 0x98, 0x00)
    qr_base.line.fill.background()

    # QR Code Inner details (Yellow squares)
    qr_inners = [(9.6, 2.4), (11.6, 2.5), (9.7, 4.4)]
    for x, y in qr_inners:
        inner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.6), Inches(0.6))
        inner.rotation = 3
        inner.fill.solid()
        inner.fill.fore_color.rgb = RGBColor(0xFF, 0xEB, 0x3B)
        inner.line.fill.background()

    # Hand-drawn circles around QR code
    circ1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.1), Inches(2.1), Inches(3.6), Inches(1.2))
    circ1.rotation = 3
    circ1.fill.background()
    circ1.line.color.rgb = RGBColor(0x00, 0x00, 0x00)
    circ1.line.width = Pt(1.5)

    circ2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.0), Inches(4.3), Inches(1.5), Inches(0.8))
    circ2.rotation = -5
    circ2.fill.background()
    circ2.line.color.rgb = RGBColor(0x00, 0x00, 0x00)
    circ2.line.width = Pt(1.5)

    # 7. Bottom Right Text Block ("内容要点")
    # Title
    tx_points_title = slide.shapes.add_textbox(Inches(8.8), Inches(5.3), Inches(3.0), Inches(0.5))
    p_title = tx_points_title.text_frame.paragraphs[0]
    p_title.text = "内容要点"
    p_title.font.name = "Microsoft YaHei"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # Points Data
    points_data = [
        ("?", "互动区：你见过最奇葩的景点在哪里？"),
        ("@", "联系方式：微博/小红书/抖音同名"),
        ("!", "结束语：山水不改，脑洞常在！")
    ]

    start_y = 5.9
    for icon_char, text_content in points_data:
        # Icon circle
        icon_shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.8), Inches(start_y), Inches(0.3), Inches(0.3))
        icon_shape.fill.solid()
        icon_shape.fill.fore_color.rgb = RGBColor(0xFF, 0xEB, 0x3B)
        icon_shape.line.color.rgb = RGBColor(0x00, 0x00, 0x00)
        icon_shape.line.width = Pt(1)
        
        icon_p = icon_shape.text_frame.paragraphs[0]
        icon_p.text = icon_char
        icon_p.font.name = "Arial"
        icon_p.font.size = Pt(12)
        icon_p.font.bold = True
        icon_p.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        icon_p.alignment = PP_ALIGN.CENTER

        # Text
        tx_point = slide.shapes.add_textbox(Inches(9.2), Inches(start_y - 0.05), Inches(4.0), Inches(0.4))
        p_point = tx_point.text_frame.paragraphs[0]
        p_point.text = text_content
        p_point.font.name = "Microsoft YaHei"
        p_point.font.size = Pt(14)
        p_point.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        
        start_y += 0.45