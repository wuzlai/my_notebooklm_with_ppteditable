def build_slide(slide):
    # 1. Background
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = RGBColor(0x05, 0x22, 0x55) # Dark blue
    bg_shape.line.fill.background()

    # Grid lines
    grid_color = RGBColor(0x10, 0x35, 0x70)
    for i in range(1, 14):
        line = slide.shapes.add_connector(1, Inches(i), Inches(0), Inches(i), Inches(7.5))
        line.line.color.rgb = grid_color
        line.line.width = Pt(0.5)
    for i in range(1, 8):
        line = slide.shapes.add_connector(1, Inches(0), Inches(i), Inches(13.333), Inches(i))
        line.line.color.rgb = grid_color
        line.line.width = Pt(0.5)

    # 2. Central Lightbulb Icon
    center_x = 13.333 / 2
    bulb_y = 1.0
    bulb_color = RGBColor(0xAA, 0xD4, 0xFF)
    
    # Glow effect (simulated with a larger faint circle)
    glow = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(center_x - 1.0), Inches(bulb_y - 0.4), Inches(2.0), Inches(2.0))
    glow.fill.solid()
    glow.fill.fore_color.rgb = RGBColor(0x15, 0x45, 0x85)
    glow.line.fill.background()

    # Main bulb (Oval)
    bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(center_x - 0.6), Inches(bulb_y), Inches(1.2), Inches(1.2))
    bulb.fill.background()
    bulb.line.color.rgb = bulb_color
    bulb.line.width = Pt(3)
    
    # Bulb base (Rectangle)
    base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(center_x - 0.3), Inches(bulb_y + 1.15), Inches(0.6), Inches(0.4))
    base.fill.background()
    base.line.color.rgb = bulb_color
    base.line.width = Pt(3)
    
    # Base screw lines
    for i in range(3):
        line_y = bulb_y + 1.6 + i * 0.12
        line = slide.shapes.add_connector(1, Inches(center_x - 0.25), Inches(line_y), Inches(center_x + 0.25), Inches(line_y))
        line.line.color.rgb = bulb_color
        line.line.width = Pt(2.5)
        
    # Filament (Inner lines)
    fil_left = slide.shapes.add_connector(1, Inches(center_x - 0.2), Inches(bulb_y + 1.15), Inches(center_x - 0.2), Inches(bulb_y + 0.6))
    fil_left.line.color.rgb = bulb_color
    fil_left.line.width = Pt(2)
    
    fil_right = slide.shapes.add_connector(1, Inches(center_x + 0.2), Inches(bulb_y + 1.15), Inches(center_x + 0.2), Inches(bulb_y + 0.6))
    fil_right.line.color.rgb = bulb_color
    fil_right.line.width = Pt(2)
    
    fil_top = slide.shapes.add_shape(MSO_SHAPE.ARC, Inches(center_x - 0.2), Inches(bulb_y + 0.4), Inches(0.4), Inches(0.4))
    fil_top.rotation = 180
    fil_top.fill.background()
    fil_top.line.color.rgb = bulb_color
    fil_top.line.width = Pt(2)

    # 3. Main Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(3.2), Inches(11.333), Inches(1))
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "简而不凡：高效演示文稿的制作之道"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 4. Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(2), Inches(4.2), Inches(9.333), Inches(0.8))
    tf_sub = subtitle_box.text_frame
    tf_sub.clear()
    p_sub = tf_sub.paragraphs[0]
    p_sub.alignment = PP_ALIGN.CENTER
    run_sub = p_sub.add_run()
    run_sub.text = "掌握专业PPT的核心逻辑与设计法则"
    run_sub.font.name = "Microsoft YaHei"
    run_sub.font.size = Pt(24)
    run_sub.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 5. Bullet Points
    bullet_texts = [
        "1. 演示文稿不仅仅是工具，更是思维的视觉化表达",
        "2. 核心目标：降低沟通成本，提升说服力",
        "3. 专家级PPT的两大底层支柱：内容清晰与设计统一"
    ]
    
    start_y = 5.2
    spacing = 0.6
    icon_x = 3.8
    text_x = 4.3
    icon_color = RGBColor(0xAA, 0xD4, 0xFF)
    
    for i, text in enumerate(bullet_texts):
        icon_y = start_y + i * spacing + 0.08
        
        # Draw Icons
        if i == 0: # Idea (Lightbulb)
            icon_bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(icon_x), Inches(icon_y), Inches(0.2), Inches(0.2))
            icon_bulb.fill.background()
            icon_bulb.line.color.rgb = icon_color
            icon_bulb.line.width = Pt(1.5)
            icon_base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(icon_x + 0.05), Inches(icon_y + 0.18), Inches(0.1), Inches(0.08))
            icon_base.fill.background()
            icon_base.line.color.rgb = icon_color
            icon_base.line.width = Pt(1.5)
            # Rays
            slide.shapes.add_connector(1, Inches(icon_x+0.1), Inches(icon_y-0.05), Inches(icon_x+0.1), Inches(icon_y-0.1)).line.color.rgb = icon_color
            slide.shapes.add_connector(1, Inches(icon_x-0.05), Inches(icon_y+0.1), Inches(icon_x-0.1), Inches(icon_y+0.1)).line.color.rgb = icon_color
            slide.shapes.add_connector(1, Inches(icon_x+0.25), Inches(icon_y+0.1), Inches(icon_x+0.3), Inches(icon_y+0.1)).line.color.rgb = icon_color

        elif i == 1: # Target
            icon1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(icon_x), Inches(icon_y), Inches(0.25), Inches(0.25))
            icon1.fill.background()
            icon1.line.color.rgb = icon_color
            icon1.line.width = Pt(1.5)
            icon2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(icon_x+0.075), Inches(icon_y+0.075), Inches(0.1), Inches(0.1))
            icon2.fill.background()
            icon2.line.color.rgb = icon_color
            icon2.line.width = Pt(1.5)
            # Arrow
            arrow = slide.shapes.add_connector(1, Inches(icon_x+0.15), Inches(icon_y+0.1), Inches(icon_x+0.35), Inches(icon_y-0.1))
            arrow.line.color.rgb = icon_color
            arrow.line.width = Pt(1.5)

        elif i == 2: # Balance
            # Base triangle
            tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(icon_x+0.05), Inches(icon_y+0.1), Inches(0.15), Inches(0.15))
            tri.fill.background()
            tri.line.color.rgb = icon_color
            tri.line.width = Pt(1.5)
            # Top bar
            bar = slide.shapes.add_connector(1, Inches(icon_x-0.05), Inches(icon_y+0.1), Inches(icon_x+0.3), Inches(icon_y+0.1))
            bar.line.color.rgb = icon_color
            bar.line.width = Pt(1.5)
            # Left pan
            slide.shapes.add_connector(1, Inches(icon_x-0.05), Inches(icon_y+0.1), Inches(icon_x-0.05), Inches(icon_y+0.25)).line.color.rgb = icon_color
            slide.shapes.add_connector(1, Inches(icon_x-0.1), Inches(icon_y+0.25), Inches(icon_x), Inches(icon_y+0.25)).line.color.rgb = icon_color
            # Right pan
            slide.shapes.add_connector(1, Inches(icon_x+0.3), Inches(icon_y+0.1), Inches(icon_x+0.3), Inches(icon_y+0.25)).line.color.rgb = icon_color
            slide.shapes.add_connector(1, Inches(icon_x+0.25), Inches(icon_y+0.25), Inches(icon_x+0.35), Inches(icon_y+0.25)).line.color.rgb = icon_color

        # Text
        text_box = slide.shapes.add_textbox(Inches(text_x), Inches(start_y + i * spacing), Inches(8), Inches(0.4))
        tf_bullet = text_box.text_frame
        tf_bullet.clear()
        p_bullet = tf_bullet.paragraphs[0]
        run_bullet = p_bullet.add_run()
        run_bullet.text = text
        run_bullet.font.name = "Microsoft YaHei"
        run_bullet.font.size = Pt(16)
        run_bullet.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)