def build_slide(slide):
    # 1. Background (Green irregular shape approximated by a rectangle)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.6), Inches(11.7), Inches(6.3))
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0x2E, 0x9E, 0x7B)
    bg.line.fill.background()

    # 2. Title
    title_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.333), Inches(1.0))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "总结：桂林，一个越怪越美的地方"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.font.name = "Microsoft YaHei"
    p.font.shadow = True

    # 3. Subtitle
    sub_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(11.333), Inches(0.8))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = "这里的山水有灵，这里的人们有戏"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.font.name = "Microsoft YaHei"
    p.font.shadow = True

    # 4. Item 1: Core Viewpoint
    # Camera Icon (Constructed with shapes)
    cam_base = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.6), Inches(3.4), Inches(0.8), Inches(0.6))
    cam_base.fill.solid()
    cam_base.fill.fore_color.rgb = RGBColor(0xE0, 0xF7, 0xFA)
    cam_base.line.color.rgb = RGBColor(0x00, 0x00, 0x00)
    cam_base.line.width = Pt(2)
    
    cam_lens = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.75), Inches(3.45), Inches(0.5), Inches(0.5))
    cam_lens.fill.solid()
    cam_lens.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    cam_lens.line.color.rgb = RGBColor(0x00, 0x00, 0x00)
    cam_lens.line.width = Pt(2)

    cam_heart = slide.shapes.add_shape(MSO_SHAPE.HEART, Inches(1.85), Inches(3.55), Inches(0.3), Inches(0.3))
    cam_heart.fill.solid()
    cam_heart.fill.fore_color.rgb = RGBColor(0xFF, 0x66, 0x00)
    cam_heart.line.fill.background()

    cam_flash = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.15), Inches(3.48), Inches(0.12), Inches(0.12))
    cam_flash.fill.solid()
    cam_flash.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    cam_flash.line.color.rgb = RGBColor(0x00, 0x00, 0x00)
    cam_flash.line.width = Pt(1)

    # Text
    tb1 = slide.shapes.add_textbox(Inches(2.6), Inches(3.4), Inches(9.0), Inches(0.8))
    p1 = tb1.text_frame.paragraphs[0]
    p1.text = "核心观点：风景是背景，有趣才是旅行的灵魂"
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p1.font.name = "Microsoft YaHei"
    p1.font.shadow = True

    # 5. Item 2: Novelty Index
    # Star Icon
    star_icon1 = slide.shapes.add_shape(MSO_SHAPE.5_POINT_STAR, Inches(1.7), Inches(4.5), Inches(0.6), Inches(0.6))
    star_icon1.fill.solid()
    star_icon1.fill.fore_color.rgb = RGBColor(0xFF, 0x66, 0x00)
    star_icon1.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    star_icon1.line.width = Pt(1.5)

    # Text
    tb2 = slide.shapes.add_textbox(Inches(2.6), Inches(4.4), Inches(2.5), Inches(0.8))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = "猎奇指数："
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p2.font.name = "Microsoft YaHei"
    p2.font.shadow = True

    # Rating Stars
    for i in range(5):
        star = slide.shapes.add_shape(MSO_SHAPE.5_POINT_STAR, Inches(4.8 + i*0.7), Inches(4.5), Inches(0.6), Inches(0.6))
        star.fill.solid()
        star.fill.fore_color.rgb = RGBColor(0xFF, 0x66, 0x00)
        star.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        star.line.width = Pt(1.5)

    # 6. Item 3: Recommendation Index
    # Star Icon
    star_icon2 = slide.shapes.add_shape(MSO_SHAPE.5_POINT_STAR, Inches(1.7), Inches(5.5), Inches(0.6), Inches(0.6))
    star_icon2.fill.solid()
    star_icon2.fill.fore_color.rgb = RGBColor(0xFF, 0x66, 0x00)
    star_icon2.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    star_icon2.line.width = Pt(1.5)

    # Text
    tb3 = slide.shapes.add_textbox(Inches(2.6), Inches(5.4), Inches(2.5), Inches(0.8))
    p3 = tb3.text_frame.paragraphs[0]
    p3.text = "推荐指数："
    p3.font.size = Pt(28)
    p3.font.bold = True
    p3.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p3.font.name = "Microsoft YaHei"
    p3.font.shadow = True

    # Rating Stars
    for i in range(5):
        star = slide.shapes.add_shape(MSO_SHAPE.5_POINT_STAR, Inches(4.8 + i*0.7), Inches(5.5), Inches(0.6), Inches(0.6))
        if i < 4:
            star.fill.solid()
            star.fill.fore_color.rgb = RGBColor(0xFF, 0x66, 0x00)
            star.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        else:
            star.fill.solid()
            star.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            star.line.color.rgb = RGBColor(0xFF, 0x66, 0x00)
        star.line.width = Pt(1.5)

    # Comment Text
    tb_comment = slide.shapes.add_textbox(Inches(8.2), Inches(5.45), Inches(4.0), Inches(0.8))
    p_comment = tb_comment.text_frame.paragraphs[0]
    p_comment.text = "(扣一星怕它太骄傲)"
    p_comment.font.size = Pt(24)
    p_comment.font.bold = True
    p_comment.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p_comment.font.name = "Microsoft YaHei"
    p_comment.font.shadow = True

    # 7. Decorations
    # OMG!
    omg_box = slide.shapes.add_textbox(Inches(10.2), Inches(3.2), Inches(2.0), Inches(1.0))
    omg_box.rotation = 15
    p_omg = omg_box.text_frame.paragraphs[0]
    p_omg.text = "OMG!"
    p_omg.font.size = Pt(32)
    p_omg.font.bold = True
    p_omg.font.color.rgb = RGBColor(0xFF, 0x66, 0x00)
    p_omg.font.name = "Arial"

    # Boom!
    boom_box = slide.shapes.add_textbox(Inches(10.5), Inches(4.8), Inches(2.0), Inches(1.0))
    boom_box.rotation = -15
    p_boom = boom_box.text_frame.paragraphs[0]
    p_boom.text = "Boom!"
    p_boom.font.size = Pt(32)
    p_boom.font.bold = True
    p_boom.font.color.rgb = RGBColor(0x80, 0x00, 0x80)
    p_boom.font.name = "Arial"

    # Arrow
    arrow = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(7.6), Inches(4.2), Inches(0.3), Inches(0.5))
    arrow.rotation = 225
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = RGBColor(0xFF, 0x33, 0x33)
    arrow.line.fill.background()

    # Decorative circles
    circle1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.5), Inches(3.1), Inches(0.15), Inches(0.15))
    circle1.fill.background()
    circle1.line.color.rgb = RGBColor(0xFF, 0x66, 0x00)
    circle1.line.width = Pt(2)

    circle2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.8), Inches(4.8), Inches(0.15), Inches(0.15))
    circle2.fill.background()
    circle2.line.color.rgb = RGBColor(0xFF, 0x33, 0x33)
    circle2.line.width = Pt(2)
    
    # Decorative arc
    arc = slide.shapes.add_shape(MSO_SHAPE.ARC, Inches(10.5), Inches(6.0), Inches(0.8), Inches(0.4))
    arc.rotation = 180
    arc.fill.background()
    arc.line.color.rgb = RGBColor(0x80, 0x00, 0x80)
    arc.line.width = Pt(2)

    # 8. Page Number
    page_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.8), Inches(6.5), Inches(1.0), Inches(0.4))
    page_box.fill.solid()
    page_box.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    page_box.line.color.rgb = RGBColor(0x2E, 0x9E, 0x7B)
    page_box.line.width = Pt(2)
    
    tf_page = page_box.text_frame
    tf_page.margin_left = 0
    tf_page.margin_right = 0
    tf_page.margin_top = 0
    tf_page.margin_bottom = 0
    p_page = tf_page.paragraphs[0]
    p_page.text = "第10页"
    p_page.alignment = PP_ALIGN.CENTER
    p_page.font.size = Pt(14)
    p_page.font.bold = True
    p_page.font.color.rgb = RGBColor(0x2E, 0x9E, 0x7B)
    p_page.font.name = "Microsoft YaHei"