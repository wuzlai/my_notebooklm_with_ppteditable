def build_slide(slide):
    # 1. 添加顶部标题横幅
    add_header_banner(
        slide, 
        "简单场景：AI 提效显著但仍需人工干预", 
        bg_color=RGBColor(0x15, 0x55, 0xC0)
    )

    # 2. 添加副标题
    add_subtitle(
        slide, 
        "案例A — 销售订单报表查询（低复杂度验证）", 
        left=Inches(0.5), 
        top=Inches(1.2), 
        font_size=Pt(18)
    )

    # 3. 添加左侧要点列表 (带图标、标签和描述)
    left_col_x = Inches(0.5)
    
    add_bullet_item(
        slide, 
        left=left_col_x, 
        top=Inches(2.0), 
        symbol="⚡", 
        label="效率对比：", 
        description="Claude Code 耗时仅 30 分钟，较手写开发提效 50%，而 Copilot 无明显提升。",
        width=Inches(5.5)
    )
    
    add_bullet_item(
        slide, 
        left=left_col_x, 
        top=Inches(3.2), 
        symbol="🐛", 
        label="生成质量：", 
        description="两款 AI 初次生成均无法直接运行，均存在地址取数逻辑或字段虚构问题。",
        width=Inches(5.5)
    )
    
    add_bullet_item(
        slide, 
        left=left_col_x, 
        top=Inches(4.4), 
        symbol="🛡️", 
        label="运行风险：", 
        description="Copilot 生成代码导致系统崩溃 (Short Dump)，Claude Code 运行相对平稳。",
        width=Inches(5.5)
    )
    
    add_bullet_item(
        slide, 
        left=left_col_x, 
        top=Inches(5.6), 
        symbol="⚙️", 
        label="性能瓶颈：", 
        description="AI 在 SQL 优化方面表现欠缺，存在大表关联不当及未去重等性能隐患。",
        width=Inches(5.5)
    )

    # 4. 添加中间的垂直分割线
    separator = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(6.4), 
        top=Inches(2.0), 
        width=Pt(1), 
        height=Inches(4.5)
    )
    separator.fill.solid()
    separator.fill.fore_color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    separator.line.fill.background()

    # 5. 添加右侧柱状图
    chart_left = Inches(6.8)
    gray_color = RGBColor(0xA6, 0xA6, 0xA6)
    cyan_color = RGBColor(0x00, 0xBC, 0xD4)
    
    add_bar_chart(
        slide,
        left=chart_left,
        top=Inches(1.8),
        width=Inches(6.0),
        height=Inches(3.5),
        categories=["Copilot", "Claude Code", "手写开发"],
        values=[60, 30, 60],
        title="开发耗时对比（分钟）",
        bar_colors=[gray_color, cyan_color, gray_color]
    )

    # 6. 添加图表上的标注标签
    add_callout_label(
        slide, 
        left=Inches(10.0), 
        top=Inches(3.0), 
        text="提效 50%", 
        bg_color=cyan_color
    )
    
    add_callout_label(
        slide, 
        left=Inches(11.6), 
        top=Inches(4.2), 
        text="无明显提升", 
        bg_color=gray_color
    )

    # 7. 添加右下角核心结论框
    conclusion_text = "核心结论：在简单场景下，Claude Code 展现出显著的速度优势，但两款工具生成的代码均需人工深度介入进行逻辑修正和性能优化才能交付使用。"
    add_conclusion_box(
        slide,
        left=chart_left,
        top=Inches(5.6),
        width=Inches(6.0),
        text=conclusion_text
    )