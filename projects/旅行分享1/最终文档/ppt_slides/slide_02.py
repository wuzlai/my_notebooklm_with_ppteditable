def build_slide(slide):
    # 1. 设置背景颜色 (浅灰白)
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(248, 248, 248)

    # 右侧边缘装饰条 (渐变橙/绿效果的简化)
    edge = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(13.1), Inches(0), Inches(0.233), Inches(7.5))
    edge.fill.solid()
    edge.fill.fore_color.rgb = RGBColor(230, 120, 50)
    edge.line.fill.background()

    # 2. 标题与副标题
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(10), Inches(1))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "角色设定：谁在带你逛桂林？"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.color.rgb = RGBColor(10, 10, 10)

    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(0.8))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "颜值不够，脑洞来凑"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = RGBColor(20, 20, 20)

    # 页码
    pg_box = slide.shapes.add_textbox(Inches(11.5), Inches(0.4), Inches(1.5), Inches(0.6))
    tf = pg_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "第2页"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(22)
    run.font.bold = True

    # 3. 左侧视觉区域 (橙色泼墨背景 + 人物抠图占位)
    splash_color = RGBColor(255, 102, 0)
    
    # 使用多个椭圆组合模拟泼墨形状
    s1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.2), Inches(2.8), Inches(3.8), Inches(3.8))
    s1.fill.solid(); s1.fill.fore_color.rgb = splash_color; s1.line.fill.background()
    s2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), Inches(4.0), Inches(2.0), Inches(2.0))
    s2.fill.solid(); s2.fill.fore_color.rgb = splash_color; s2.line.fill.background()
    s3 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.5), Inches(2.5), Inches(2.0), Inches(2.0))
    s3.fill.solid(); s3.fill.fore_color.rgb = splash_color; s3.line.fill.background()
    s4 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.2), Inches(5.2), Inches(2.2), Inches(1.8))
    s4.fill.solid(); s4.fill.fore_color.rgb = splash_color; s4.line.fill.background()

    # 人物抠图占位符 (带白色粗边框的圆角矩形)
    cutout = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.8), Inches(2.8), Inches(2.4), Inches(3.8))
    cutout.fill.solid()
    cutout.fill.fore_color.rgb = RGBColor(30, 30, 30) # 深色衣服
    cutout.line.color.rgb = RGBColor(255, 255, 255)
    cutout.line.width = Pt(6)
    
    # 脸部占位
    face = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.4), Inches(3.0), Inches(1.2), Inches(1.5))
    face.fill.solid(); face.fill.fore_color.rgb = RGBColor(255, 218, 185); face.line.fill.background()
    
    # 相机占位
    cam = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.5), Inches(4.5), Inches(1.0), Inches(0.7))
    cam.fill.solid(); cam.fill.fore_color.rgb = RGBColor(50, 50, 50); cam.line.fill.background()
    lens = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.65), Inches(4.55), Inches(0.7), Inches(0.6))
    lens.fill.solid(); lens.fill.fore_color.rgb = RGBColor(10, 10, 10); lens.line.color.rgb = RGBColor(100, 100, 100)

    # "粗糙抠图" 文本
    cutout_txt = slide.shapes.add_textbox(Inches(2.2), Inches(6.7), Inches(2.0), Inches(0.5))
    tf = cutout_txt.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "粗糙抠图"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(14)
    run.font.bold = True

    # 装饰性英文文本
    def add_deco_text(text, left, top, rotation, color, size=22):
        box = slide.shapes.add_textbox(left, top, Inches(1.5), Inches(0.8))
        box.rotation = rotation
        tf = box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text
        run.font.name = "Arial Black"
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = True

    add_deco_text("OMG!", Inches(1.0), Inches(3.0), -15, RGBColor(0, 160, 140))
    add_deco_text("!?", Inches(3.5), Inches(2.8), 15, RGBColor(255, 255, 255), size=28)
    add_deco_text("BOOM!", Inches(3.8), Inches(3.8), 15, RGBColor(255, 255, 255))
    add_deco_text("LOOK!", Inches(4.2), Inches(5.5), 20, RGBColor(255, 255, 255))
    add_deco_text("NEW!", Inches(0.8), Inches(6.2), -10, RGBColor(0, 160, 140))
    add_deco_text("NEW!", Inches(11.5), Inches(2.0), 20, RGBColor(255, 112, 0))

    # 4. 右侧内容区域
    def add_content_block(left, top, width, title, desc):
        box = slide.shapes.add_textbox(left, top, width, Inches(1.5))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.line_spacing = 1.3
        
        run1 = p.add_run()
        run1.text = title
        run1.font.name = "Microsoft YaHei"
        run1.font.size = Pt(20)
        run1.font.bold = True
        run1.font.color.rgb = RGBColor(0, 0, 0)
        
        run2 = p.add_run()
        run2.text = desc
        run2.font.name = "Microsoft YaHei"
        run2.font.size = Pt(20)
        run2.font.bold = True
        run2.font.color.rgb = RGBColor(0, 0, 0)

    # --- 条目 1: 身份 ---
    i1_left, i1_top = Inches(5.5), Inches(2.5)
    # 图标占位 (人物 + 地图)
    p_body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, i1_left+Inches(0.2), i1_top+Inches(0.6), Inches(0.8), Inches(0.6))
    p_body.fill.solid(); p_body.fill.fore_color.rgb = RGBColor(220, 50, 50); p_body.line.color.rgb = RGBColor(0, 0, 0)
    p_head = slide.shapes.add_shape(MSO_SHAPE.OVAL, i1_left+Inches(0.35), i1_top+Inches(0.1), Inches(0.5), Inches(0.5))
    p_head.fill.solid(); p_head.fill.fore_color.rgb = RGBColor(255, 200, 180); p_head.line.color.rgb = RGBColor(0, 0, 0)
    
    map_s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, i1_left+Inches(1.1), i1_top+Inches(0.3), Inches(0.8), Inches(0.7))
    map_s.rotation = 10
    map_s.fill.solid(); map_s.fill.fore_color.rgb = RGBColor(150, 200, 150); map_s.line.color.rgb = RGBColor(0, 0, 0)
    
    # 地图上的红叉
    x1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, i1_left+Inches(1.2), i1_top+Inches(0.4), i1_left+Inches(1.8), i1_top+Inches(0.9))
    x1.line.color.rgb = RGBColor(255, 0, 0); x1.line.width = Pt(3)
    x2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, i1_left+Inches(1.8), i1_top+Inches(0.4), i1_left+Inches(1.2), i1_top+Inches(0.9))
    x2.line.color.rgb = RGBColor(255, 0, 0); x2.line.width = Pt(3)

    # allergy 文本
    alg_box = slide.shapes.add_textbox(i1_left+Inches(1.0), i1_top+Inches(1.0), Inches(1.0), Inches(0.3))
    alg_box.rotation = -5
    tf = alg_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "allergy"
    run.font.name = "Arial"
    run.font.size = Pt(12)
    run.font.bold = True
    
    add_content_block(Inches(7.8), Inches(2.5), Inches(5.0), "身份：", "一个对“正常景点”\n过敏的猎奇博主")

    # --- 条目 2: 装备 ---
    add_content_block(Inches(5.8), Inches(4.2), Inches(4.5), "装备：", "自拍杆、扩音器、\n以及随时准备跑路的运动鞋")
    
    i2_left, i2_top = Inches(10.2), Inches(4.0)
    # 图标占位 (紫色背景 + 喇叭 + 鞋子)
    bg2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, i2_left, i2_top, Inches(1.2), Inches(1.2))
    bg2.fill.solid(); bg2.fill.fore_color.rgb = RGBColor(160, 32, 240); bg2.line.fill.background()
    
    mega = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, i2_left+Inches(0.6), i2_top-Inches(0.2), Inches(0.6), Inches(0.8))
    mega.rotation = 90
    mega.fill.solid(); mega.fill.fore_color.rgb = RGBColor(255, 255, 255); mega.line.color.rgb = RGBColor(160, 32, 240); mega.line.width = Pt(2)
    
    shoe = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, i2_left+Inches(1.0), i2_top+Inches(0.8), Inches(1.0), Inches(0.4))
    shoe.rotation = -15
    shoe.fill.solid(); shoe.fill.fore_color.rgb = RGBColor(255, 150, 0); shoe.line.color.rgb = RGBColor(0, 0, 0)
    
    add_deco_text("AHA!", i2_left+Inches(1.2), i2_top-Inches(0.5), -10, RGBColor(160, 32, 240))

    # --- 条目 3: 目标 ---
    add_content_block(Inches(5.8), Inches(6.0), Inches(3.5), "目标：\n", "寻找桂林最“野”的打开\n方式")
    
    i3_left, i3_top = Inches(9.5), Inches(5.8)
    # 图标占位 (青色背景 + 山水 + 指南针)
    bg3 = slide.shapes.add_shape(MSO_SHAPE.OVAL, i3_left+Inches(0.5), i3_top, Inches(1.4), Inches(1.4))
    bg3.fill.solid(); bg3.fill.fore_color.rgb = RGBColor(0, 160, 140); bg3.line.fill.background()
    
    m1 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, i3_left+Inches(0.6), i3_top+Inches(0.4), Inches(0.6), Inches(0.8))
    m1.fill.solid(); m1.fill.fore_color.rgb = RGBColor(20, 20, 20); m1.line.fill.background()
    m2 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, i3_left+Inches(1.0), i3_top+Inches(0.2), Inches(0.8), Inches(1.0))
    m2.fill.solid(); m2.fill.fore_color.rgb = RGBColor(10, 10, 10); m2.line.fill.background()
    
    comp = slide.shapes.add_shape(MSO_SHAPE.OVAL, i3_left, i3_top+Inches(0.5), Inches(0.8), Inches(0.8))
    comp.fill.solid(); comp.fill.fore_color.rgb = RGBColor(255, 220, 150); comp.line.color.rgb = RGBColor(0, 0, 0); comp.line.width = Pt(2)
    
    add_deco_text("WILD?", i3_left+Inches(1.5), i3_top-Inches(0.4), 15, RGBColor(0, 0, 0))

    # 5. 添加手绘风格的连接线/箭头指示
    a1 = slide.shapes.add_connector(MSO_CONNECTOR.CURVE, Inches(7.3), Inches(3.5), Inches(7.7), Inches(3.4))
    a1.line.color.rgb = RGBColor(150, 150, 150)
    
    a2 = slide.shapes.add_connector(MSO_CONNECTOR.CURVE, Inches(9.5), Inches(5.0), Inches(10.0), Inches(4.8))
    a2.line.color.rgb = RGBColor(150, 150, 150)
    
    a3 = slide.shapes.add_connector(MSO_CONNECTOR.CURVE, Inches(8.8), Inches(6.8), Inches(9.3), Inches(6.5))
    a3.line.color.rgb = RGBColor(150, 150, 150)