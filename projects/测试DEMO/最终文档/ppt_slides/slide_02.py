def build_slide(slide):
    # Colors
    DARK_BLUE = RGBColor(0x1A, 0x36, 0x5D)
    RED = RGBColor(0xE5, 0x39, 0x35)
    LIGHT_RED = RGBColor(0xFD, 0xED, 0xEC)
    LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)
    GRAY_TEXT = RGBColor(0x55, 0x55, 0x55)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    LIGHT_GRAY = RGBColor(0xE0, 0xE0, 0xE0)
    
    # 1. Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(10), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "中等场景瓶颈：AI “幻觉”引发崩溃"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE
    p.font.name = "Microsoft YaHei"

    # 2. Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(10), Inches(0.5))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = "采购配额维护（中等复杂度）开发验证"
    p.font.size = Pt(18)
    p.font.color.rgb = GRAY_TEXT
    p.font.name = "Microsoft YaHei"

    # 3. Left Bullets
    bullets_data = [
        ("🔗", "逻辑误区", "Copilot 完全混淆业务概念（配额误作货源），代码完全不可用。"),
        ("⚠️", "幻觉严重", "Claude Code 虚构字段比例高达 50%，导致 21 个连锁语法错误。"),
        ("🚫", "效率归零", "AI 修复成本大于重写成本，整体效率对比手写无任何提升。"),
        ("⚠️", "规则缺失", "AI 无法准确遵循 SAP 函数接口规范，直接忽略“优先 API”的指令。")
    ]
    
    start_y = 1.8
    for icon, label, desc in bullets_data:
        # Icon
        icon_box = slide.shapes.add_textbox(Inches(0.5), Inches(start_y), Inches(0.6), Inches(0.6))
        tf = icon_box.text_frame
        p = tf.paragraphs[0]
        p.text = icon
        p.font.size = Pt(28)
        p.font.color.rgb = RED
        p.alignment = PP_ALIGN.CENTER
        
        # Label
        label_box = slide.shapes.add_textbox(Inches(1.3), Inches(start_y), Inches(4.5), Inches(0.4))
        tf = label_box.text_frame
        p = tf.paragraphs[0]
        p.text = label
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.name = "Microsoft YaHei"
        
        # Desc
        desc_box = slide.shapes.add_textbox(Inches(1.3), Inches(start_y + 0.35), Inches(4.8), Inches(0.8))
        tf = desc_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = desc
        p.font.size = Pt(14)
        p.font.color.rgb = GRAY_TEXT
        p.font.name = "Microsoft YaHei"
        
        start_y += 1.3

    # 4. Right Data Card
    card_left = Inches(6.8)
    card_top = Inches(1.4)
    card_width = Inches(6.0)
    card_height = Inches(1.4)
    
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_left, card_top, card_width, card_height)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = LIGHT_GRAY
    card.line.width = Pt(1)

    # Card Text - 50%
    val_box = slide.shapes.add_textbox(card_left + Inches(0.2), card_top + Inches(0.1), Inches(3), Inches(0.8))
    tf = val_box.text_frame
    p = tf.paragraphs[0]
    p.text = "50% ↗"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RED
    p.font.name = "Arial"

    # Card Text - Label
    lbl_box = slide.shapes.add_textbox(card_left + Inches(0.2), card_top + Inches(0.9), Inches(4), Inches(0.4))
    tf = lbl_box.text_frame
    p = tf.paragraphs[0]
    p.text = "虚构字段比例 (Fictional Field Ratio)"
    p.font.size = Pt(12)
    p.font.color.rgb = GRAY_TEXT
    p.font.name = "Microsoft YaHei"

    # Card Icon - Warning
    warn_box = slide.shapes.add_textbox(card_left + Inches(4.8), card_top + Inches(0.2), Inches(1), Inches(1))
    tf = warn_box.text_frame
    p = tf.paragraphs[0]
    p.text = "⚠️"
    p.font.size = Pt(60)
    p.font.color.rgb = RED
    p.alignment = PP_ALIGN.CENTER

    # 5. Right Table
    table_top = Inches(3.0)
    
    # Table Title
    tt_box = slide.shapes.add_textbox(card_left, table_top, card_width, Inches(0.4))
    tf = tt_box.text_frame
    p = tf.paragraphs[0]
    p.text = "字段对比示例"
    p.font.size = Pt(14)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Microsoft YaHei"

    # Table Shape
    rows = 6
    cols = 2
    table_shape = slide.shapes.add_table(rows, cols, card_left, table_top + Inches(0.4), card_width, Inches(2.0)).table
    table_shape.columns[0].width = Inches(3.0)
    table_shape.columns[1].width = Inches(3.0)

    table_data = [
        ("❌ 虚构字段", "✅ 正确字段"),
        ("MSEG-QUOTA_ID (不存在)", "EKKO-EBELN (采购凭证)"),
        ("EKPO-ALLOC_QTY (错误逻辑)", "EKPO-MENGE (数量)"),
        ("EKPO-ALLOC_QTY (错误逻辑)", "EKPO-MENGE (数量)"),
        ("EKPO-BLG_ID (不存在)", "EKKO-QUOTA (数量)"),
        ("...", "...")
    ]

    for r in range(rows):
        for c in range(cols):
            cell = table_shape.cell(r, c)
            cell.text = table_data[r][c]
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.font.name = "Microsoft YaHei"
            
            if r == 0:
                p.font.bold = True
            
            # Set background colors
            if r == 0:
                if c == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT_RED
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT_GREEN
            else:
                if c == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(0xFE, 0xF5, 0xF5)
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(0xF1, 0xF8, 0xF1)

    # 6. Right Flowchart
    flow_top = Inches(5.7)
    
    # Box 1
    b1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_left, flow_top, Inches(1.2), Inches(1.2))
    b1.fill.solid()
    b1.fill.fore_color.rgb = RED
    b1.line.fill.background()
    tf = b1.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "50%"
    p1.font.size = Pt(20)
    p1.font.color.rgb = WHITE
    p1.font.bold = True
    p1.alignment = PP_ALIGN.CENTER
    p1.font.name = "Arial"
    p2 = tf.add_paragraph()
    p2.text = "虚构字段"
    p2.font.size = Pt(12)
    p2.font.color.rgb = WHITE
    p2.alignment = PP_ALIGN.CENTER
    p2.font.name = "Microsoft YaHei"

    # Arrow 1
    a1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, card_left + Inches(1.3), flow_top + Inches(0.5), Inches(0.3), Inches(0.2))
    a1.fill.solid()
    a1.fill.fore_color.rgb = GRAY_TEXT
    a1.line.fill.background()

    # Box 2
    b2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_left + Inches(1.7), flow_top, Inches(1.6), Inches(1.2))
    b2.fill.solid()
    b2.fill.fore_color.rgb = RED
    b2.line.fill.background()
    tf = b2.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "🔗"
    p1.font.size = Pt(20)
    p1.font.color.rgb = WHITE
    p1.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = "21 个连锁语法错误"
    p2.font.size = Pt(11)
    p2.font.color.rgb = WHITE
    p2.alignment = PP_ALIGN.CENTER
    p2.font.name = "Microsoft YaHei"

    # Arrow 2
    a2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, card_left + Inches(3.4), flow_top + Inches(0.5), Inches(0.3), Inches(0.2))
    a2.fill.solid()
    a2.fill.fore_color.rgb = GRAY_TEXT
    a2.line.fill.background()

    # Box 3
    b3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_left + Inches(3.8), flow_top, Inches(2.2), Inches(1.2))
    b3.fill.solid()
    b3.fill.fore_color.rgb = RED
    b3.line.fill.background()
    tf = b3.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "🚫"
    p1.font.size = Pt(20)
    p1.font.color.rgb = WHITE
    p1.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = "AI 修复成本 >> 重写成本，\n整体效率对比手写无提升"
    p2.font.size = Pt(10)
    p2.font.color.rgb = WHITE
    p2.alignment = PP_ALIGN.CENTER
    p2.font.name = "Microsoft YaHei"

    # 7. Footer
    footer = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.3))
    tf = footer.text_frame
    p = tf.paragraphs[0]
    p.text = "PAGE 2 OF 4"
    p.font.size = Pt(10)
    p.font.color.rgb = GRAY_TEXT
    p.alignment = PP_ALIGN.RIGHT