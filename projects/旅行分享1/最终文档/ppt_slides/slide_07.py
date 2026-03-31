def build_slide(slide):
    # --- 颜色定义 ---
    ORANGE = RGBColor(0xF2, 0x71, 0x27)
    DARK_TEXT = RGBColor(0x20, 0x20, 0x20)
    GRAY_TEXT = RGBColor(0x60, 0x60, 0x60)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    BLUE_LIGHT = RGBColor(0x00, 0xBC, 0xD4)
    RED = RGBColor(0xE5, 0x39, 0x35)
    IMG_BG = RGBColor(0xE8, 0xF0, 0xF4)

    # --- 1. 左上角页码 ---
    page_num = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.2), Inches(0.2), Inches(1.0), Inches(1.0))
    page_num.fill.solid()
    page_num.fill.fore_color.rgb = ORANGE
    page_num.line.fill.background()
    tf = page_num.text_frame
    p = tf.paragraphs[0]
    p.text = "第7页"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_TEXT
    p.alignment = PP_ALIGN.CENTER

    # --- 2. 主标题与副标题 ---
    title_box = slide.shapes.add_textbox(Inches(1.5), Inches(0.3), Inches(10), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "漓江竹筏：塑料椅子上的“速度与激情”"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.name = "Microsoft YaHei"
    p.font.color.rgb = DARK_TEXT

    # 标题下划线
    underline = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.6), Inches(0.9), Inches(8.5), Inches(0.1))
    underline.fill.solid()
    underline.fill.fore_color.rgb = ORANGE
    underline.line.fill.background()

    # 副标题
    subtitle_box = slide.shapes.add_textbox(Inches(2.5), Inches(1.1), Inches(8), Inches(0.5))
    tf = subtitle_box.text_frame
    p = tf.paragraphs[0]
    p.text = "别被照片骗了，这其实是水上拖拉机"
    p.font.size = Pt(20)
    p.font.name = "Microsoft YaHei"
    p.font.color.rgb = GRAY_TEXT
    p.alignment = PP_ALIGN.CENTER

    # --- 辅助函数：添加带下划线的区块标题 ---
    def add_section_title(left, top, text):
        box = slide.shapes.add_textbox(left, top, Inches(2), Inches(0.5))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.name = "Microsoft YaHei"
        
        ul = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + Inches(0.4), Inches(1.5), Inches(0.08))
        ul.fill.solid()
        ul.fill.fore_color.rgb = ORANGE
        ul.line.fill.background()

    # ==========================================
    # --- Section 1: 现实反差 (左侧) ---
    # ==========================================
    add_section_title(Inches(0.8), Inches(1.8), "现实反差")
    
    t1 = slide.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(2), Inches(0.4))
    t1.text_frame.text = "理想中：优雅的竹筏"
    
    # 图片占位 1 (传统竹筏)
    img1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.8), Inches(2.2), Inches(1.5))
    img1.fill.solid()
    img1.fill.fore_color.rgb = IMG_BG
    img1.line.color.rgb = WHITE
    img1.line.width = Pt(3)
    
    # 红色大叉
    cross = slide.shapes.add_shape(MSO_SHAPE.MATH_MULTIPLY, Inches(1.5), Inches(3.1), Inches(0.8), Inches(0.8))
    cross.fill.solid()
    cross.fill.fore_color.rgb = RED
    cross.line.fill.background()

    # 气泡提示
    bubble = slide.shapes.add_shape(MSO_SHAPE.CLOUD_CALLOUT, Inches(2.5), Inches(2.5), Inches(1.5), Inches(1.0))
    bubble.fill.solid()
    bubble.fill.fore_color.rgb = BLUE_LIGHT
    bubble.line.fill.background()
    tf = bubble.text_frame
    p = tf.paragraphs[0]
    p.text = "其实是\nPVC管+马达"
    p.font.size = Pt(12)
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    # 图片占位 2 (PVC竹筏)
    img2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(3.8), Inches(2.8), Inches(2.0))
    img2.fill.solid()
    img2.fill.fore_color.rgb = IMG_BG
    img2.line.color.rgb = WHITE
    img2.line.width = Pt(3)

    # 底部说明文字
    cap1 = slide.shapes.add_textbox(Inches(0.8), Inches(5.9), Inches(3.5), Inches(0.5))
    tf = cap1.text_frame
    p = tf.paragraphs[0]
    p.text = "优雅的竹筏其实是PVC管+马达"
    p.font.size = Pt(14)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    # ==========================================
    # --- Section 2: 搞笑瞬间 (中间) ---
    # ==========================================
    add_section_title(Inches(5.2), Inches(1.8), "搞笑瞬间")

    # 图片占位
    img3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0), Inches(2.5), Inches(3.5), Inches(2.8))
    img3.fill.solid()
    img3.fill.fore_color.rgb = IMG_BG
    img3.line.color.rgb = WHITE
    img3.line.width = Pt(3)

    # 拟声词
    t_boom = slide.shapes.add_textbox(Inches(6.5), Inches(2.8), Inches(1.5), Inches(0.5))
    t_boom.text_frame.text = "轰隆隆!\nBOOM!"
    t_boom.text_frame.paragraphs[0].font.bold = True

    # 吐槽气泡
    t_omg = slide.shapes.add_shape(MSO_SHAPE.OVAL_CALLOUT, Inches(7.0), Inches(4.0), Inches(1.2), Inches(0.8))
    t_omg.fill.solid()
    t_omg.fill.fore_color.rgb = RGBColor(0xA0, 0xD0, 0xA0)
    t_omg.line.fill.background()
    tf = t_omg.text_frame
    p = tf.paragraphs[0]
    p.text = "OMG!\n诗意呢?"
    p.font.size = Pt(10)
    p.alignment = PP_ALIGN.CENTER

    # 底部说明文字
    cap2 = slide.shapes.add_textbox(Inches(4.8), Inches(5.4), Inches(4.0), Inches(0.5))
    tf = cap2.text_frame
    p = tf.paragraphs[0]
    p.text = "马达启动时的黑烟与诗意山水的完美融合"
    p.font.size = Pt(14)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    # ==========================================
    # --- Section 3: 猎奇体验 (右侧) ---
    # ==========================================
    add_section_title(Inches(9.5), Inches(1.8), "猎奇体验")

    # 图片占位
    img4 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.2), Inches(2.5), Inches(3.5), Inches(2.8))
    img4.fill.solid()
    img4.fill.fore_color.rgb = IMG_BG
    img4.line.color.rgb = WHITE
    img4.line.width = Pt(3)

    # 标签：移动超市
    t_shop = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.0), Inches(2.8), Inches(1.2), Inches(0.4))
    t_shop.fill.solid()
    t_shop.fill.fore_color.rgb = RGBColor(0xFF, 0xDD, 0xAA)
    t_shop.line.fill.background()
    tf = t_shop.text_frame
    p = tf.paragraphs[0]
    p.text = "移动超市"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT
    p.alignment = PP_ALIGN.CENTER

    # 标签：现烤鱼
    t_fish = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.5), Inches(4.5), Inches(1.0), Inches(0.4))
    t_fish.fill.solid()
    t_fish.fill.fore_color.rgb = RGBColor(0xFF, 0xDD, 0xAA)
    t_fish.line.fill.background()
    tf = t_fish.text_frame
    p = tf.paragraphs[0]
    p.text = "现烤鱼"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT
    p.alignment = PP_ALIGN.CENTER

    # 底部说明文字
    cap3 = slide.shapes.add_textbox(Inches(9.0), Inches(5.4), Inches(4.0), Inches(0.5))
    tf = cap3.text_frame
    p = tf.paragraphs[0]
    p.text = "江上的“移动超市”，划着竹筏卖烤鱼"
    p.font.size = Pt(14)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    # --- 区块之间的引导箭头 ---
    arrow1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.5), Inches(1.9), Inches(0.8), Inches(0.3))
    arrow1.fill.solid()
    arrow1.fill.fore_color.rgb = GRAY_TEXT
    arrow1.line.fill.background()

    arrow2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(7.8), Inches(1.9), Inches(0.8), Inches(0.3))
    arrow2.fill.solid()
    arrow2.fill.fore_color.rgb = GRAY_TEXT
    arrow2.line.fill.background()

    # ==========================================
    # --- 底部图表区域: 马达轰鸣声 vs 优雅度 ---
    # ==========================================
    bot_title = slide.shapes.add_textbox(Inches(6.5), Inches(5.8), Inches(3.5), Inches(0.5))
    tf = bot_title.text_frame
    p = tf.paragraphs[0]
    p.text = "马达轰鸣声 vs 优雅度"
    p.font.size = Pt(18)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    
    bot_ul = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(6.2), Inches(2.9), Inches(0.08))
    bot_ul.fill.solid()
    bot_ul.fill.fore_color.rgb = ORANGE
    bot_ul.line.fill.background()

    # 渐变进度条 (使用纯色箭头代替)
    bar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.0), Inches(6.4), Inches(5.0), Inches(0.3))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(0xFF, 0x98, 0x00) # 橙色
    bar.line.color.rgb = DARK_TEXT
    bar.line.width = Pt(1)

    # 左侧文本 (优雅度)
    left_txt = slide.shapes.add_textbox(Inches(5.0), Inches(6.1), Inches(1.0), Inches(0.8))
    tf = left_txt.text_frame
    tf.text = "🕊️\n优雅度\n(最高)"
    tf.paragraphs[0].font.size = Pt(14)
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    if len(tf.paragraphs) > 1:
        tf.paragraphs[1].font.size = Pt(12)
        tf.paragraphs[1].font.bold = True
        tf.paragraphs[1].alignment = PP_ALIGN.CENTER
    if len(tf.paragraphs) > 2:
        tf.paragraphs[2].font.size = Pt(10)
        tf.paragraphs[2].alignment = PP_ALIGN.CENTER

    # 右侧文本 (马达轰鸣声)
    right_txt = slide.shapes.add_textbox(Inches(11.2), Inches(6.1), Inches(1.5), Inches(0.8))
    tf = right_txt.text_frame
    tf.text = "🚜\n马达轰鸣声\n(最大)"
    tf.paragraphs[0].font.size = Pt(14)
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    if len(tf.paragraphs) > 1:
        tf.paragraphs[1].font.size = Pt(12)
        tf.paragraphs[1].font.bold = True
        tf.paragraphs[1].alignment = PP_ALIGN.CENTER
    if len(tf.paragraphs) > 2:
        tf.paragraphs[2].font.size = Pt(10)
        tf.paragraphs[2].alignment = PP_ALIGN.CENTER

    # 进度条上的标签
    labels = ["静音", "嗡嗡...", "轰隆隆!!!", "VROOM!!!"]
    x_positions = [6.3, 7.3, 8.5, 9.8]
    for i, text in enumerate(labels):
        lbl = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_positions[i]), Inches(6.8), Inches(1.0), Inches(0.3))
        lbl.fill.solid()
        lbl.fill.fore_color.rgb = WHITE
        lbl.line.color.rgb = DARK_TEXT
        lbl.line.width = Pt(1)
        tf = lbl.text_frame
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = DARK_TEXT
        p.alignment = PP_ALIGN.CENTER

    # 底部总结文字
    bot_cap = slide.shapes.add_textbox(Inches(6.0), Inches(7.2), Inches(5.0), Inches(0.4))
    tf = bot_cap.text_frame
    p = tf.paragraphs[0]
    p.text = "马达一响，优雅全无 (进度条图表)"
    p.font.size = Pt(14)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER