def build_slide(slide):
    # 颜色常量定义
    BLUE_DARK = RGBColor(0x00, 0x55, 0xA4)
    BLUE_VERY_DARK = RGBColor(0x1A, 0x2B, 0x3C)
    ORANGE_TEXT = RGBColor(0xFF, 0x98, 0x00)
    GRAY_DARK = RGBColor(0x33, 0x33, 0x33)
    GRAY_LIGHT = RGBColor(0x55, 0x55, 0x55)
    GRAY_LINE = RGBColor(0xCC, 0xCC, 0xCC)
    GRAY_HANDLE = RGBColor(0x88, 0x88, 0x88)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    # 1. 标题
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(10), Inches(0.8))
    tf_title = title_box.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.text = "秒懂原则：视觉化的力量"
    p_title.font.name = FONT_NAME
    p_title.font.size = Pt(36)
    p_title.font.bold = True
    p_title.font.color.rgb = BLUE_DARK

    # 2. 副标题
    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(10), Inches(0.5))
    tf_sub = sub_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "让观众在3秒内捕捉核心信息"
    p_sub.font.name = FONT_NAME
    p_sub.font.size = Pt(20)
    p_sub.font.color.rgb = BLUE_DARK

    # 3. 左侧主图：放大镜及内部元素
    # 放大镜手柄连接处 (灰色)
    conn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.8), Inches(5.2), Inches(0.6), Inches(0.8))
    conn.rotation = 45
    conn.fill.solid()
    conn.fill.fore_color.rgb = GRAY_HANDLE
    conn.line.color.rgb = BLUE_VERY_DARK
    conn.line.width = Pt(3)

    # 放大镜手柄主体 (蓝色)
    handle = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(5.8), Inches(0.9), Inches(2.2))
    handle.rotation = 45
    handle.fill.solid()
    handle.fill.fore_color.rgb = BLUE_DARK
    handle.line.color.rgb = BLUE_VERY_DARK
    handle.line.width = Pt(4)

    # 放大镜外圈 (深色粗边框，白色填充以遮挡手柄顶部)
    outer_ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.0), Inches(2.2), Inches(4.0), Inches(4.0))
    outer_ring.fill.solid()
    outer_ring.fill.fore_color.rgb = WHITE
    outer_ring.line.color.rgb = BLUE_VERY_DARK
    outer_ring.line.width = Pt(12)

    # 放大镜内圈 (蓝色细边框)
    inner_ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.2), Inches(2.4), Inches(3.6), Inches(3.6))
    inner_ring.fill.background()
    inner_ring.line.color.rgb = BLUE_DARK
    inner_ring.line.width = Pt(6)

    # 内部放射状虚线/浅色线
    lines_coords = [
        (4.0, 2.9, 4.0, 2.6), # 上
        (4.0, 5.5, 4.0, 5.8), # 下
        (2.7, 4.2, 2.4, 4.2), # 左
        (5.3, 4.2, 5.6, 4.2), # 右
        (3.1, 3.3, 2.9, 3.1), # 左上
        (4.9, 3.3, 5.1, 3.1), # 右上
        (3.1, 5.1, 2.9, 5.3), # 左下
        (4.9, 5.1, 5.1, 5.3), # 右下
    ]
    for x1, y1, x2, y2 in lines_coords:
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        line.line.color.rgb = GRAY_LINE
        line.line.width = Pt(2)

    # 内部文字 "秒懂"
    text_box_md = slide.shapes.add_textbox(Inches(2.0), Inches(3.4), Inches(4.0), Inches(1.5))
    tf_md = text_box_md.text_frame
    p_md = tf_md.paragraphs[0]
    p_md.text = "秒懂"
    p_md.alignment = PP_ALIGN.CENTER
    p_md.font.name = FONT_NAME
    p_md.font.size = Pt(65)
    p_md.font.bold = True
    p_md.font.color.rgb = ORANGE_TEXT

    # 内部小图标：眼睛 (上方)
    eye_outer = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.3), Inches(3.1), Inches(0.4), Inches(0.25))
    eye_outer.fill.background()
    eye_outer.line.color.rgb = BLUE_DARK
    eye_outer.line.width = Pt(2)
    eye_inner = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.45), Inches(3.17), Inches(0.1), Inches(0.1))
    eye_inner.fill.solid()
    eye_inner.fill.fore_color.rgb = BLUE_DARK
    eye_inner.line.fill.background()

    # 内部小图标：大脑/云朵 (右下方)
    brain = slide.shapes.add_shape(MSO_SHAPE.CLOUD, Inches(4.7), Inches(4.5), Inches(0.4), Inches(0.3))
    brain.fill.background()
    brain.line.color.rgb = BLUE_DARK
    brain.line.width = Pt(1.5)

    # 放大镜外部装饰弧线 (使用简单的曲线近似)
    arc1 = slide.shapes.add_shape(MSO_SHAPE.ARC, Inches(1.6), Inches(1.8), Inches(4.8), Inches(4.8))
    arc1.fill.background()
    arc1.line.color.rgb = BLUE_VERY_DARK
    arc1.line.width = Pt(3)
    arc1.adjustments[0] = 110
    arc1.adjustments[1] = 160

    arc2 = slide.shapes.add_shape(MSO_SHAPE.ARC, Inches(1.8), Inches(2.0), Inches(4.4), Inches(4.4))
    arc2.fill.background()
    arc2.line.color.rgb = BLUE_VERY_DARK
    arc2.line.width = Pt(3)
    arc2.adjustments[0] = 20
    arc2.adjustments[1] = 70

    # 4. 右侧要点列表
    # --- 要点 1：降低认知负荷 ---
    # 图标 1 (人脑齿轮)
    head = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.4), Inches(2.5), Inches(0.5), Inches(0.6))
    head.fill.background()
    head.line.color.rgb = BLUE_DARK
    head.line.width = Pt(2)
    neck = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.55), Inches(3.05), Inches(0.2), Inches(0.15))
    neck.fill.background()
    neck.line.color.rgb = BLUE_DARK
    neck.line.width = Pt(2)
    gear1 = slide.shapes.add_shape(MSO_SHAPE.GEAR_6, Inches(7.45), Inches(2.6), Inches(0.25), Inches(0.25))
    gear1.fill.background()
    gear1.line.color.rgb = BLUE_DARK
    gear1.line.width = Pt(1.5)
    gear2 = slide.shapes.add_shape(MSO_SHAPE.GEAR_6, Inches(7.6), Inches(2.8), Inches(0.2), Inches(0.2))
    gear2.fill.background()
    gear2.line.color.rgb = BLUE_DARK
    gear2.line.width = Pt(1.5)

    # 文本 1
    tb1 = slide.shapes.add_textbox(Inches(8.4), Inches(2.4), Inches(4.5), Inches(1.0))
    tf1 = tb1.text_frame
    p1_1 = tf1.paragraphs[0]
    p1_1.text = "降低认知负荷："
    p1_1.font.name = FONT_NAME
    p1_1.font.size = Pt(18)
    p1_1.font.bold = True
    p1_1.font.color.rgb = GRAY_DARK
    p1_2 = tf1.add_paragraph()
    p1_2.text = "一眼就能看懂逻辑"
    p1_2.font.name = FONT_NAME
    p1_2.font.size = Pt(14)
    p1_2.font.color.rgb = GRAY_LIGHT
    p1_2.space_before = Pt(6)

    # --- 要点 2：视觉层级明确 ---
    # 图标 2 (层级金字塔与箭头)
    base_y2 = 4.8
    rect1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.3), Inches(base_y2), Inches(0.8), Inches(0.15))
    rect1.fill.background()
    rect1.line.color.rgb = BLUE_DARK
    rect1.line.width = Pt(2)
    rect2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.4), Inches(base_y2 - 0.25), Inches(0.6), Inches(0.15))
    rect2.fill.background()
    rect2.line.color.rgb = BLUE_DARK
    rect2.line.width = Pt(2)
    rect3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.5), Inches(base_y2 - 0.5), Inches(0.4), Inches(0.15))
    rect3.fill.background()
    rect3.line.color.rgb = BLUE_DARK
    rect3.line.width = Pt(2)
    arrow = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(8.2), Inches(base_y2 - 0.5), Inches(0.15), Inches(0.65))
    arrow.fill.background()
    arrow.line.color.rgb = BLUE_DARK
    arrow.line.width = Pt(1.5)

    # 文本 2
    tb2 = slide.shapes.add_textbox(Inches(8.4), Inches(4.1), Inches(4.5), Inches(1.0))
    tf2 = tb2.text_frame
    p2_1 = tf2.paragraphs[0]
    p2_1.text = "视觉层级明确："
    p2_1.font.name = FONT_NAME
    p2_1.font.size = Pt(18)
    p2_1.font.bold = True
    p2_1.font.color.rgb = GRAY_DARK
    p2_2 = tf2.add_paragraph()
    p2_2.text = "通过大小、颜色区分主次"
    p2_2.font.name = FONT_NAME
    p2_2.font.size = Pt(14)
    p2_2.font.color.rgb = GRAY_LIGHT
    p2_2.space_before = Pt(6)

    # --- 要点 3：善用图表 ---
    # 图标 3 (柱状图与饼图)
    base_y3 = 6.5
    axis_x = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.3), Inches(base_y3), Inches(8.0), Inches(base_y3))
    axis_x.line.color.rgb = BLUE_DARK
    axis_x.line.width = Pt(2)
    axis_y = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.3), Inches(base_y3), Inches(7.3), Inches(base_y3 - 0.8))
    axis_y.line.color.rgb = BLUE_DARK
    axis_y.line.width = Pt(2)
    
    bar1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.4), Inches(base_y3 - 0.3), Inches(0.12), Inches(0.3))
    bar1.fill.background()
    bar1.line.color.rgb = BLUE_DARK
    bar1.line.width = Pt(1.5)
    bar2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.6), Inches(base_y3 - 0.5), Inches(0.12), Inches(0.5))
    bar2.fill.background()
    bar2.line.color.rgb = BLUE_DARK
    bar2.line.width = Pt(1.5)
    bar3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.8), Inches(base_y3 - 0.2), Inches(0.12), Inches(0.2))
    bar3.fill.background()
    bar3.line.color.rgb = BLUE_DARK
    bar3.line.width = Pt(1.5)
    
    pie = slide.shapes.add_shape(MSO_SHAPE.PIE, Inches(7.7), Inches(base_y3 - 0.9), Inches(0.4), Inches(0.4))
    pie.fill.background()
    pie.line.color.rgb = BLUE_DARK
    pie.line.width = Pt(1.5)

    # 文本 3
    tb3 = slide.shapes.add_textbox(Inches(8.4), Inches(5.8), Inches(4.5), Inches(1.0))
    tf3 = tb3.text_frame
    p3_1 = tf3.paragraphs[0]
    p3_1.text = "善用图表："
    p3_1.font.name = FONT_NAME
    p3_1.font.size = Pt(18)
    p3_1.font.bold = True
    p3_1.font.color.rgb = GRAY_DARK
    p3_2 = tf3.add_paragraph()
    p3_2.text = "数据关系一目了然"
    p3_2.font.name = FONT_NAME
    p3_2.font.size = Pt(14)
    p3_2.font.color.rgb = GRAY_LIGHT
    p3_2.space_before = Pt(6)

    # 5. 页码
    page_num = slide.shapes.add_textbox(Inches(12.0), Inches(6.8), Inches(1.0), Inches(0.5))
    tf_page = page_num.text_frame
    p_page = tf_page.paragraphs[0]
    p_page.text = "5 / 11"
    p_page.font.name = FONT_NAME
    p_page.font.size = Pt(12)
    p_page.font.color.rgb = GRAY_HANDLE
    p_page.alignment = PP_ALIGN.RIGHT