def build_slide(slide):
    # 自定义颜色
    LIGHT_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
    DENSE_TEXT_COLOR = RGBColor(0x99, 0x99, 0x99)
    
    # 1. 标题
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(10), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "减法艺术：拒绝文字堆砌"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = BLUE_DARK
    p.font.name = "Microsoft YaHei"

    # 2. 副标题
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(10), Inches(0.6))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "别让你的PPT变成Word搬家"
    p.font.size = Pt(22)
    p.font.color.rgb = GRAY_TEXT
    p.font.name = "Microsoft YaHei"

    # 3. 左侧栏 (错误示例)
    # 标题
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(5.5), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Word搬家（错误示例）"
    p.font.size = Pt(18)
    p.font.color.rgb = RED
    p.font.name = "Microsoft YaHei"
    
    # 红色下划线
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.5), Inches(2.7), Inches(6.0), Inches(2.7))
    line.line.color.rgb = RED
    line.line.width = Pt(1.5)

    # 密集文本块
    dense_text = "这里是一段非常长且密集的文字，用来模拟将Word文档直接复制粘贴到PPT中的错误做法。在实际的演示中，观众根本无法在短时间内阅读并理解这么多文字。这种做法不仅会让幻灯片显得杂乱无章，还会严重分散观众的注意力，导致他们无法专心听讲。优秀的PPT应该只保留核心观点和关键词，通过演讲者的口述来补充细节。如果把所有内容都写在屏幕上，那么演讲者就失去了存在的意义，PPT也就变成了一份阅读材料而不是辅助演示的工具。因此，我们必须学会做减法，拒绝文字堆砌，提炼出最精炼的信息，用视觉化的方式呈现出来，从而提高沟通的效率和效果。" * 4
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(2.9), Inches(5.5), Inches(3.4))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = dense_text
    p.font.size = Pt(8)
    p.font.color.rgb = DENSE_TEXT_COLOR
    p.font.name = "Microsoft YaHei"
    p.alignment = PP_ALIGN.JUSTIFY

    # 红色大叉
    cross_center_x = 3.25
    cross_center_y = 4.6
    cross_size = 1.2
    line1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cross_center_x - cross_size), Inches(cross_center_y - cross_size), Inches(cross_center_x + cross_size), Inches(cross_center_y + cross_size))
    line1.line.color.rgb = RED
    line1.line.width = Pt(25)
    line2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cross_center_x - cross_size), Inches(cross_center_y + cross_size), Inches(cross_center_x + cross_size), Inches(cross_center_y - cross_size))
    line2.line.color.rgb = RED
    line2.line.width = Pt(25)

    # 左侧底部结论
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(6.0), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "密密麻麻的文字，信息过载，观众无法聚焦。"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = BLACK
    p.font.name = "Microsoft YaHei"

    # 4. 中间垂直分割线
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(6.66), Inches(2.2), Inches(6.66), Inches(6.8))
    line.line.color.rgb = LIGHT_GRAY
    line.line.width = Pt(1)

    # 5. 右侧栏 (成功示例)
    # 标题
    txBox = slide.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.5), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "极简要点（成功示例）"
    p.font.size = Pt(18)
    p.font.color.rgb = GREEN
    p.font.name = "Microsoft YaHei"
    
    # 绿色下划线
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.2), Inches(2.7), Inches(12.8), Inches(2.7))
    line.line.color.rgb = GREEN
    line.line.width = Pt(1.5)

    # 绘制要点条目的内部函数
    def draw_bullet(slide, left, top, icon_text, label, desc):
        # 图标
        txBox = slide.shapes.add_textbox(left, top, Inches(0.6), Inches(0.6))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = icon_text
        p.font.size = Pt(28)
        p.font.name = "Segoe UI Emoji"
        
        # 标签
        txBox = slide.shapes.add_textbox(left + Inches(0.8), top, Inches(4.5), Inches(0.35))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = label
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = BLACK
        p.font.name = "Microsoft YaHei"
        
        # 描述
        txBox = slide.shapes.add_textbox(left + Inches(0.8), top + Inches(0.35), Inches(4.5), Inches(0.35))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = desc
        p.font.size = Pt(14)
        p.font.color.rgb = BLACK
        p.font.name = "Microsoft YaHei"

    # 添加三个要点
    draw_bullet(slide, Inches(7.2), Inches(3.1), "👂", "专注聆听", "观众阅读文字时，无法同时听取演讲。")
    draw_bullet(slide, Inches(7.2), Inches(4.3), "💎", "提炼金句", "删除冗余的修饰词，只保留核心观点。")
    draw_bullet(slide, Inches(7.2), Inches(5.5), "🖼️", "视觉替代", "用视觉元素（图标/图片）替代长篇大论。")

    # 绿色大勾
    check_x = 9.8
    check_y = 4.8
    line1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(check_x - 0.6), Inches(check_y - 0.1), Inches(check_x), Inches(check_y + 0.5))
    line1.line.color.rgb = GREEN
    line1.line.width = Pt(25)
    line2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(check_x), Inches(check_y + 0.5), Inches(check_x + 1.2), Inches(check_y - 1.0))
    line2.line.color.rgb = GREEN
    line2.line.width = Pt(25)

    # 右侧底部结论
    txBox = slide.shapes.add_textbox(Inches(7.2), Inches(6.5), Inches(6.0), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "精简内容，视觉引导，提升传递效率。"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = GREEN
    p.font.name = "Microsoft YaHei"

    # 6. 页码
    txBox = slide.shapes.add_textbox(Inches(11.8), Inches(6.8), Inches(1.0), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "4 / 11"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = BLUE_DARK
    p.font.name = "Microsoft YaHei"
    p.alignment = PP_ALIGN.RIGHT