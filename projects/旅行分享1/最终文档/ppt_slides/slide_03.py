def build_slide(slide):
    # Colors
    BLACK = RGBColor(0x00, 0x00, 0x00)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    GREEN_LBL = RGBColor(0x8A, 0xE0, 0xA1)
    ORANGE_LBL = RGBColor(0xFF, 0xB3, 0x47)
    GREEN_BORDER = RGBColor(0x4C, 0xAF, 0x50)
    ORANGE_BORDER = RGBColor(0xFF, 0x57, 0x22)
    PURPLE = RGBColor(0x9C, 0x27, 0xB0)
    GRAY_TEXT = RGBColor(0x75, 0x75, 0x75)

    # 1. Title
    title_box = slide.shapes.add_textbox(Inches(1.5), Inches(0.4), Inches(10.333), Inches(1))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    
    run1 = p.add_run()
    run1.text = "20元人民币打卡： "
    run1.font.size = Pt(36)
    run1.font.bold = True
    run1.font.name = "Microsoft YaHei"
    
    run2 = p.add_run()
    run2.text = "买家秀"
    run2.font.size = Pt(36)
    run2.font.bold = True
    run2.font.name = "Microsoft YaHei"
    
    run3 = p.add_run()
    run3.text = " vs "
    run3.font.size = Pt(36)
    run3.font.bold = True
    run3.font.name = "Microsoft YaHei"
    
    run4 = p.add_run()
    run4.text = "卖家秀"
    run4.font.size = Pt(36)
    run4.font.bold = True
    run4.font.name = "Microsoft YaHei"

    # Purple underline for "卖家秀" (approximate position)
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(9.3), Inches(1.15), Inches(10.8), Inches(1.15))
    line.line.color.rgb = PURPLE
    line.line.width = Pt(4)

    # 2. Subtitle
    sub_box = slide.shapes.add_textbox(Inches(3), Inches(1.1), Inches(7.333), Inches(0.6))
    tf_sub = sub_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.alignment = PP_ALIGN.CENTER
    run_sub = p_sub.add_run()
    run_sub.text = "理想很丰满，现实很骨感"
    run_sub.font.size = Pt(22)
    run_sub.font.bold = True
    run_sub.font.name = "Microsoft YaHei"

    # 3. Center Divider (Lightning Bolt)
    lightning = slide.shapes.add_shape(MSO_SHAPE.LIGHTNING_BOLT, Inches(6.3), Inches(1.8), Inches(0.8), Inches(5.5))
    lightning.fill.solid()
    lightning.fill.fore_color.rgb = ORANGE_BORDER
    lightning.line.fill.background()

    # --- LEFT SIDE (IDEAL) ---

    # Label: 理想 (Ideal)
    lbl_ideal = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.3), Inches(1.8), Inches(2.0), Inches(0.5))
    lbl_ideal.fill.solid()
    lbl_ideal.fill.fore_color.rgb = GREEN_LBL
    lbl_ideal.line.fill.background()
    tf_ideal = lbl_ideal.text_frame
    tf_ideal.text = "理想 (Ideal)"
    tf_ideal.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_ideal.paragraphs[0].font.size = Pt(18)
    tf_ideal.paragraphs[0].font.bold = True
    tf_ideal.paragraphs[0].font.color.rgb = BLACK
    tf_ideal.paragraphs[0].font.name = "Microsoft YaHei"

    # Image Placeholder: 20 RMB
    img_ideal = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.4), Inches(5.0), Inches(2.6))
    img_ideal.fill.solid()
    img_ideal.fill.fore_color.rgb = RGBColor(0xE8, 0xF5, 0xE9)
    img_ideal.line.fill.solid()
    img_ideal.line.fore_color.rgb = GREEN_BORDER
    img_ideal.line.width = Pt(3)
    tf_img1 = img_ideal.text_frame
    tf_img1.text = "[20元人民币风景图]"
    tf_img1.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_img1.paragraphs[0].font.color.rgb = RGBColor(0x9E, 0x9E, 0x9E)

    # Photographer Placeholder
    photo_guy = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.8), Inches(5.2), Inches(1.5), Inches(1.8))
    photo_guy.fill.solid()
    photo_guy.fill.fore_color.rgb = RGBColor(0xBB, 0xDE, 0xFB)
    photo_guy.line.fill.solid()
    photo_guy.line.fore_color.rgb = BLACK
    photo_guy.line.width = Pt(2)
    tf_guy = photo_guy.text_frame
    tf_guy.text = "摄影师\n(插画)"
    tf_guy.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_guy.paragraphs[0].font.color.rgb = BLACK

    # Speech Bubble: PERFECT SHOT!
    bubble1 = slide.shapes.add_shape(MSO_SHAPE.OVAL_CALLOUT, Inches(3.1), Inches(4.8), Inches(1.5), Inches(0.8))
    bubble1.fill.solid()
    bubble1.fill.fore_color.rgb = WHITE
    bubble1.line.fill.solid()
    bubble1.line.fore_color.rgb = BLACK
    bubble1.line.width = Pt(2)
    tf_b1 = bubble1.text_frame
    tf_b1.text = "PERFECT\nSHOT!"
    tf_b1.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_b1.paragraphs[0].font.size = Pt(11)
    tf_b1.paragraphs[0].font.bold = True
    tf_b1.paragraphs[0].font.color.rgb = BLACK

    # Conclusion Box
    box1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.5), Inches(6.4), Inches(2.8), Inches(0.8))
    box1.fill.solid()
    box1.fill.fore_color.rgb = WHITE
    box1.line.fill.solid()
    box1.line.fore_color.rgb = BLACK
    box1.line.width = Pt(2)
    tf_box1 = box1.text_frame
    tf_box1.text = "结论：找准角度，你就是\n人民币上的那个男人/女人"
    tf_box1.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_box1.paragraphs[0].font.size = Pt(11)
    tf_box1.paragraphs[0].font.bold = True
    tf_box1.paragraphs[0].font.color.rgb = BLACK
    tf_box1.paragraphs[0].font.name = "Microsoft YaHei"

    # Arrow to Conclusion Box
    arrow1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.1), Inches(6.7), Inches(0.4), Inches(0.2))
    arrow1.fill.solid()
    arrow1.fill.fore_color.rgb = ORANGE_BORDER
    arrow1.line.fill.background()

    # --- RIGHT SIDE (REALITY) ---

    # Label: 现实 (Reality)
    lbl_reality = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.9), Inches(1.8), Inches(2.2), Inches(0.5))
    lbl_reality.fill.solid()
    lbl_reality.fill.fore_color.rgb = ORANGE_LBL
    lbl_reality.line.fill.background()
    tf_reality = lbl_reality.text_frame
    tf_reality.text = "现实 (Reality)"
    tf_reality.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_reality.paragraphs[0].font.size = Pt(18)
    tf_reality.paragraphs[0].font.bold = True
    tf_reality.paragraphs[0].font.color.rgb = BLACK
    tf_reality.paragraphs[0].font.name = "Microsoft YaHei"

    # Image Placeholder: Crowd
    img_reality = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.5), Inches(2.4), Inches(5.0), Inches(3.0))
    img_reality.fill.solid()
    img_reality.fill.fore_color.rgb = RGBColor(0xFF, 0xE0, 0xB2)
    img_reality.line.fill.solid()
    img_reality.line.fore_color.rgb = ORANGE_BORDER
    img_reality.line.width = Pt(3)
    tf_img2 = img_reality.text_frame
    tf_img2.text = "[拥挤的人群举着20元拍照]"
    tf_img2.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_img2.paragraphs[0].font.color.rgb = RGBColor(0x9E, 0x9E, 0x9E)

    # Stressed Guy Placeholder
    stress_guy = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.5), Inches(3.5), Inches(1.5), Inches(1.8))
    stress_guy.fill.solid()
    stress_guy.fill.fore_color.rgb = RGBColor(0xFF, 0xCD, 0xD2)
    stress_guy.line.fill.solid()
    stress_guy.line.fore_color.rgb = BLACK
    stress_guy.line.width = Pt(2)
    tf_guy2 = stress_guy.text_frame
    tf_guy2.text = "崩溃游客\n(插画)"
    tf_guy2.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_guy2.paragraphs[0].font.color.rgb = BLACK

    # Speech Bubble: OMG! TOO MANY PEOPLE!
    bubble2 = slide.shapes.add_shape(MSO_SHAPE.CLOUD_CALLOUT, Inches(11.0), Inches(2.2), Inches(1.8), Inches(1.2))
    bubble2.fill.solid()
    bubble2.fill.fore_color.rgb = WHITE
    bubble2.line.fill.solid()
    bubble2.line.fore_color.rgb = BLACK
    bubble2.line.width = Pt(2)
    tf_b2 = bubble2.text_frame
    tf_b2.text = "OMG!\nTOO MANY\nPEOPLE!"
    tf_b2.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_b2.paragraphs[0].font.size = Pt(11)
    tf_b2.paragraphs[0].font.bold = True
    tf_b2.paragraphs[0].font.color.rgb = BLACK

    # Reality Text Box
    box2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.2), Inches(4.8), Inches(2.5), Inches(1.0))
    box2.fill.solid()
    box2.fill.fore_color.rgb = WHITE
    box2.line.fill.solid()
    box2.line.fore_color.rgb = BLACK
    box2.line.width = Pt(2)
    tf_box2 = box2.text_frame
    tf_box2.text = "现实：岸边挤满了100个\n同样拿着20块钱的人"
    tf_box2.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_box2.paragraphs[0].font.size = Pt(11)
    tf_box2.paragraphs[0].font.bold = True
    tf_box2.paragraphs[0].font.color.rgb = BLACK
    tf_box2.paragraphs[0].font.name = "Microsoft YaHei"

    # Cormorant Placeholder
    bird = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.2), Inches(5.5), Inches(1.2), Inches(1.8))
    bird.fill.solid()
    bird.fill.fore_color.rgb = RGBColor(0xCF, 0xD8, 0xDC)
    bird.line.fill.solid()
    bird.line.fore_color.rgb = BLACK
    bird.line.width = Pt(2)
    tf_bird = bird.text_frame
    tf_bird.text = "鸬鹚\n(插画)"
    tf_bird.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_bird.paragraphs[0].font.color.rgb = BLACK

    # RETIRED sign
    sign = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.3), Inches(6.5), Inches(1.0), Inches(0.4))
    sign.fill.solid()
    sign.fill.fore_color.rgb = RGBColor(0xFF, 0xF5, 0x9D)
    sign.line.fill.solid()
    sign.line.fore_color.rgb = BLACK
    sign.line.width = Pt(1)
    tf_sign = sign.text_frame
    tf_sign.text = "RETIRED"
    tf_sign.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_sign.paragraphs[0].font.size = Pt(10)
    tf_sign.paragraphs[0].font.bold = True
    tf_sign.paragraphs[0].font.color.rgb = BLACK

    # Fact Text Box
    box3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), Inches(6.4), Inches(2.8), Inches(0.8))
    box3.fill.solid()
    box3.fill.fore_color.rgb = WHITE
    box3.line.fill.solid()
    box3.line.fore_color.rgb = BLACK
    box3.line.width = Pt(2)
    tf_box3 = box3.text_frame
    tf_box3.text = "猎奇点：那只配合拍照的\n鸬鹚其实已经“退休”了"
    tf_box3.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf_box3.paragraphs[0].font.size = Pt(11)
    tf_box3.paragraphs[0].font.bold = True
    tf_box3.paragraphs[0].font.color.rgb = BLACK
    tf_box3.paragraphs[0].font.name = "Microsoft YaHei"

    # Arrow to Fact Box
    arrow2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.3), Inches(6.7), Inches(0.4), Inches(0.2))
    arrow2.fill.solid()
    arrow2.fill.fore_color.rgb = ORANGE_BORDER
    arrow2.line.fill.background()

    # 4. Page Number
    page_num = slide.shapes.add_textbox(Inches(12.5), Inches(7.0), Inches(0.8), Inches(0.4))
    tf_page = page_num.text_frame
    p_page = tf_page.paragraphs[0]
    p_page.text = "3 / 11"
    p_page.font.size = Pt(14)
    p_page.font.color.rgb = GRAY_TEXT
    p_page.alignment = PP_ALIGN.RIGHT