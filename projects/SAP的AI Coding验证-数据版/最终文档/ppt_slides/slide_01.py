def build_slide(slide):
    # Define Colors
    TITLE_COLOR = RGBColor(0x1A, 0x3B, 0x5C)  # Dark Blue/Teal
    GREEN_COLOR = RGBColor(0x00, 0xC8, 0x75)  # Bright Green
    ORANGE_COLOR = RGBColor(0xF2, 0x8C, 0x28) # Bright Orange
    LINE_COLOR = RGBColor(0xD0, 0xD0, 0xD0)   # Light Gray
    MUTED_TEXT = RGBColor(0x59, 0x59, 0x59)   # Gray for page number
    BLACK_TEXT = RGBColor(0x00, 0x00, 0x00)

    # --- Title Section ---
    # Main Title
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(0.6))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "简单报表：AI 提效的“甜点区”"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = TITLE_COLOR
    p.font.name = "Microsoft YaHei"

    # Subtitle
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(1.15), Inches(10), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "案例A - 销售订单报表查询（低复杂度）验证"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TITLE_COLOR
    p.font.name = "Microsoft YaHei"

    # Horizontal Separator Line
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.8), Inches(1.8), Inches(12.5), Inches(1.8))
    connector.line.color.rgb = LINE_COLOR
    connector.line.width = Pt(1)

    # --- Middle Section ---
    # Left Column (Claude Code)
    # Icon
    txBox = slide.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(0.8), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "⏱"
    p.font.size = Pt(48)
    p.font.color.rgb = GREEN_COLOR

    # Title
    txBox = slide.shapes.add_textbox(Inches(2.0), Inches(2.3), Inches(4.5), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Claude Code"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = BLACK_TEXT
    p.font.name = "Microsoft YaHei"
    run = p.add_run()
    run.text = " （AI 提效点）"
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = BLACK_TEXT
    run.font.name = "Microsoft YaHei"

    # Highlight Text
    txBox = slide.shapes.add_textbox(Inches(2.0), Inches(2.7), Inches(4.0), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "提效 "
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = GREEN_COLOR
    p.font.name = "Microsoft YaHei"
    run = p.add_run()
    run.text = "50%"
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = GREEN_COLOR
    run.font.name = "Microsoft YaHei"

    # Bullet Icon
    txBox = slide.shapes.add_textbox(Inches(1.0), Inches(3.8), Inches(0.4), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "✅"
    p.font.size = Pt(16)

    # Bullet Text
    txBox = slide.shapes.add_textbox(Inches(1.4), Inches(3.8), Inches(4.5), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "仅需 2 轮人工干预即可运行"
    p.font.size = Pt(16)
    p.font.color.rgb = BLACK_TEXT
    p.font.name = "Microsoft YaHei"

    # Right Column (GitHub Copilot)
    # Icon
    txBox = slide.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(0.8), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "⏱"
    p.font.size = Pt(48)
    p.font.color.rgb = ORANGE_COLOR

    # Title
    txBox = slide.shapes.add_textbox(Inches(8.2), Inches(2.3), Inches(4.5), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "GitHub Copilot"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = BLACK_TEXT
    p.font.name = "Microsoft YaHei"
    run = p.add_run()
    run.text = " （风险点）"
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = BLACK_TEXT
    run.font.name = "Microsoft YaHei"

    # Highlight Text
    txBox = slide.shapes.add_textbox(Inches(8.2), Inches(2.7), Inches(4.0), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "提效 "
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = ORANGE_COLOR
    p.font.name = "Microsoft YaHei"
    run = p.add_run()
    run.text = "0%"
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = ORANGE_COLOR
    run.font.name = "Microsoft YaHei"

    # Bullet Icon
    txBox = slide.shapes.add_textbox(Inches(7.2), Inches(3.8), Inches(0.4), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "⚠️"
    p.font.size = Pt(16)

    # Bullet Text
    txBox = slide.shapes.add_textbox(Inches(7.6), Inches(3.8), Inches(5.0), Inches(0.8))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "因 OpenSQL 字段不兼容导致 1 次运行崩溃（Short Dump）"
    p.font.size = Pt(16)
    p.font.color.rgb = BLACK_TEXT
    p.font.name = "Microsoft YaHei"

    # Vertical Separator Line
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(6.66), Inches(2.2), Inches(6.66), Inches(4.5))
    connector.line.color.rgb = LINE_COLOR
    connector.line.width = Pt(1)

    # --- Bottom Section ---
    # Section Title
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(5.0), Inches(10), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "共性问题与性能隐患"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = TITLE_COLOR
    p.font.name = "Microsoft YaHei"

    # Bullet 1 Icon
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(5.7), Inches(0.4), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "⚠️"
    p.font.size = Pt(18)

    # Bullet 1 Text
    txBox = slide.shapes.add_textbox(Inches(1.3), Inches(5.7), Inches(11.5), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "共性问题："
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = BLACK_TEXT
    p.font.name = "Microsoft YaHei"
    run = p.add_run()
    run.text = "初始生成均不可直接编译，AI 难以准确处理地址实例化及过账状态逻辑。"
    run.font.bold = False
    run.font.size = Pt(16)
    run.font.color.rgb = BLACK_TEXT
    run.font.name = "Microsoft YaHei"

    # Bullet 2 Icon
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(6.4), Inches(0.4), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "⚠️"
    p.font.size = Pt(18)

    # Bullet 2 Text
    txBox = slide.shapes.add_textbox(Inches(1.3), Inches(6.4), Inches(11.5), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "性能隐患："
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = BLACK_TEXT
    p.font.name = "Microsoft YaHei"
    run = p.add_run()
    run.text = "AI 生成的 SQL 逻辑存在冗余查询和未去重问题，大数据量下性能堪忧。"
    run.font.bold = False
    run.font.size = Pt(16)
    run.font.color.rgb = BLACK_TEXT
    run.font.name = "Microsoft YaHei"

    # --- Footer ---
    txBox = slide.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(1.0), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "1 / 4"
    p.font.size = Pt(12)
    p.font.color.rgb = MUTED_TEXT
    p.alignment = PP_ALIGN.RIGHT
    p.font.name = "Microsoft YaHei"