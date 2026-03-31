def build_slide(slide):
    # 1. 添加左上角页码
    tx_box_page = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(2.0), Inches(0.8))
    tf_page = tx_box_page.text_frame
    p_page = tf_page.paragraphs[0]
    p_page.text = "第4页"
    p_page.font.name = "Microsoft YaHei"
    p_page.font.size = Pt(32)
    p_page.font.bold = True
    p_page.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # 2. 添加主标题
    tx_box_title = slide.shapes.add_textbox(Inches(3.5), Inches(0.6), Inches(9.0), Inches(0.8))
    tf_title = tx_box_title.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.text = "象鼻山：这头大象到底在喝什么？"
    p_title.font.name = "Microsoft YaHei"
    p_title.font.size = Pt(36)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # 3. 添加副标题
    tx_box_subtitle = slide.shapes.add_textbox(Inches(6.5), Inches(1.3), Inches(6.0), Inches(0.6))
    tf_subtitle = tx_box_subtitle.text_frame
    p_subtitle = tf_subtitle.paragraphs[0]
    p_subtitle.text = "关于桂林城徽的终极猜想"
    p_subtitle.font.name = "Microsoft YaHei"
    p_subtitle.font.size = Pt(24)
    p_subtitle.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p_subtitle.alignment = PP_ALIGN.RIGHT

    # 4. 左侧大象喝奶茶插图占位符及“奶茶？”文字
    # 插图占位符
    pic_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.2), Inches(5.5), Inches(4.5))
    pic_shape.fill.solid()
    pic_shape.fill.fore_color.rgb = RGBColor(0xF5, 0xF5, 0xF5)
    pic_shape.line.fill.background()
    tf_pic = pic_shape.text_frame
    p_pic = tf_pic.paragraphs[0]
    p_pic.text = "[大象喝奶茶插图区域]"
    p_pic.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    p_pic.alignment = PP_ALIGN.CENTER

    # “奶茶？”文字
    tx_box_tea = slide.shapes.add_textbox(Inches(5.2), Inches(4.8), Inches(1.5), Inches(0.8))
    tx_box_tea.rotation = -15
    tf_tea = tx_box_tea.text_frame
    p_tea = tf_tea.paragraphs[0]
    p_tea.text = "奶茶？"
    p_tea.font.name = "Microsoft YaHei"
    p_tea.font.size = Pt(28)
    p_tea.font.bold = True
    p_tea.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # 5. 右侧图文列表项
    # 列表项 1
    icon1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.8), Inches(2.6), Inches(0.8), Inches(0.8))
    icon1.fill.solid()
    icon1.fill.fore_color.rgb = RGBColor(0xE1, 0xF5, 0xFE)
    icon1.line.fill.background()
    icon1.text_frame.text = "📷"
    
    tx1 = slide.shapes.add_textbox(Inches(7.8), Inches(2.5), Inches(5.0), Inches(1.0))
    tf1 = tx1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "视觉错位："
    p1.font.name = "Microsoft YaHei"
    p1.font.bold = True
    p1.font.size = Pt(18)
    p1.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    run1 = p1.add_run()
    run1.text = "从哪个角度看最像一只喝醉的大象"
    run1.font.bold = False
    run1.font.size = Pt(18)

    # 列表项 2
    icon2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.8), Inches(4.1), Inches(0.8), Inches(0.8))
    icon2.fill.solid()
    icon2.fill.fore_color.rgb = RGBColor(0xF3, 0xE5, 0xF5)
    icon2.line.fill.background()
    icon2.text_frame.text = "🍷"
    
    tx2 = slide.shapes.add_textbox(Inches(7.8), Inches(4.0), Inches(5.0), Inches(1.0))
    tf2 = tx2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "猎奇冷知识："
    p2.font.name = "Microsoft YaHei"
    p2.font.bold = True
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    run2 = p2.add_run()
    run2.text = "象鼻山内部其实是空的（藏酒洞）"
    run2.font.bold = False
    run2.font.size = Pt(18)

    # 列表项 3
    icon3 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.8), Inches(5.6), Inches(0.8), Inches(0.8))
    icon3.fill.solid()
    icon3.fill.fore_color.rgb = RGBColor(0xE0, 0xF2, 0xF1)
    icon3.line.fill.background()
    icon3.text_frame.text = "💦"
    
    tx3 = slide.shapes.add_textbox(Inches(7.8), Inches(5.5), Inches(5.0), Inches(1.0))
    tf3 = tx3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.text = "吐槽点："
    p3.font.name = "Microsoft YaHei"
    p3.font.bold = True
    p3.font.size = Pt(18)
    p3.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    run3 = p3.add_run()
    run3.text = "为了拍一张“喂象照”，我差点掉进江里"
    run3.font.bold = False
    run3.font.size = Pt(18)

    # 6. 底部右侧装饰元素 (BooM!, OMG!, 箭头)
    # BooM!
    boom_box = slide.shapes.add_textbox(Inches(8.5), Inches(6.8), Inches(1.2), Inches(0.6))
    boom_box.rotation = -10
    tf_boom = boom_box.text_frame
    p_boom = tf_boom.paragraphs[0]
    p_boom.text = "BooM!"
    p_boom.font.name = "Microsoft YaHei"
    p_boom.font.size = Pt(16)
    p_boom.font.bold = True
    p_boom.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # 青色箭头
    arrow1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.8), Inches(6.9), Inches(0.4), Inches(0.25))
    arrow1.fill.solid()
    arrow1.fill.fore_color.rgb = RGBColor(0x00, 0xBC, 0xD4)
    arrow1.line.fill.background()

    # OMG!
    omg_box = slide.shapes.add_textbox(Inches(10.3), Inches(6.8), Inches(1.0), Inches(0.6))
    tf_omg = omg_box.text_frame
    p_omg = tf_omg.paragraphs[0]
    p_omg.text = "OMG!"
    p_omg.font.name = "Microsoft YaHei"
    p_omg.font.size = Pt(16)
    p_omg.font.bold = True
    p_omg.font.color.rgb = RGBColor(0x00, 0xBC, 0xD4)

    # 紫色箭头
    arrow2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(11.5), Inches(6.9), Inches(0.4), Inches(0.25))
    arrow2.fill.solid()
    arrow2.fill.fore_color.rgb = RGBColor(0x9C, 0x27, 0xB0)
    arrow2.line.fill.background()

    # 绿色箭头
    arrow3 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(12.1), Inches(6.9), Inches(0.4), Inches(0.25))
    arrow3.fill.solid()
    arrow3.fill.fore_color.rgb = RGBColor(0x4C, 0xAF, 0x50)
    arrow3.line.fill.background()