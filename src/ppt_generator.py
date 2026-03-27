"""Generate PPTX from infographic images using Gemini to write python-pptx code."""

import os
import re
import traceback

from src.gemini_client import generate_text_with_images
from src.prompts import PPT_CODE_GEN_SYSTEM_PROMPT, PPT_CODE_GEN_USER_PROMPT

DEFAULT_MODEL = "gemini-3-pro-preview"

# ── Shared boilerplate embedded in every generated script ──
_SCRIPT_HEADER = '''\
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.chart.data import CategoryChartData

# ── Color Palette ──
BLUE_HEADER = RGBColor(0x5B, 0x9B, 0xD5)
BLUE_DARK   = RGBColor(0x4A, 0x86, 0xC8)
CYAN        = RGBColor(0x00, 0xBC, 0xD4)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
BLACK       = RGBColor(0x33, 0x33, 0x33)
GRAY_TEXT   = RGBColor(0x55, 0x55, 0x55)
GRAY_BAR    = RGBColor(0xB0, 0xBE, 0xC5)
RED         = RGBColor(0xE5, 0x39, 0x35)
GREEN       = RGBColor(0x43, 0xA0, 0x47)
ORANGE      = RGBColor(0xFF, 0x98, 0x00)

ICON_BG  = RGBColor(0xE3, 0xE8, 0xED)
ICON_FG  = RGBColor(0x54, 0x6E, 0x7A)
FONT_NAME = "Microsoft YaHei"
SLIDE_WIDTH = Inches(13.333)
HEADER_H    = Inches(0.75)
SUBTITLE_Y  = Inches(0.95)


def add_header_banner(slide, title_text, bg_color=None):
    if bg_color is None:
        bg_color = BLUE_HEADER
    banner = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, HEADER_H
    )
    banner.fill.solid()
    banner.fill.fore_color.rgb = bg_color
    banner.line.fill.background()
    tf = banner.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.6)
    tf.margin_top = Inches(0.08)
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(26)
    p.font.color.rgb = WHITE
    p.font.bold = True
    p.font.name = FONT_NAME


def add_subtitle(slide, text, left, top, width=Inches(12), font_size=Pt(18)):
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.4))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = font_size
    p.font.color.rgb = BLACK
    p.font.bold = True
    p.font.name = FONT_NAME
    return txBox


def add_icon_box(slide, left, top, symbol, size=Inches(0.48)):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, size, size
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = ICON_BG
    shape.line.fill.background()
    shape.adjustments[0] = 0.25
    tf = shape.text_frame
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = symbol
    p.font.size = Pt(18)
    p.font.color.rgb = ICON_FG
    p.font.bold = False
    return shape


def add_bullet_item(slide, left, top, symbol, label, description,
                    width=Inches(5.5), desc_size=Pt(13)):
    add_icon_box(slide, left, top, symbol)
    text_left = left + Inches(0.65)
    txBox = slide.shapes.add_textbox(text_left, top - Inches(0.02), width, Inches(0.65))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run_label = p.add_run()
    run_label.text = label + "\\uff1a"
    run_label.font.size = Pt(14)
    run_label.font.color.rgb = BLACK
    run_label.font.bold = True
    run_label.font.name = FONT_NAME
    run_desc = p.add_run()
    run_desc.text = description
    run_desc.font.size = desc_size
    run_desc.font.color.rgb = GRAY_TEXT
    run_desc.font.bold = False
    run_desc.font.name = FONT_NAME
    return txBox


def add_conclusion_box(slide, left, top, width, text, font_size=Pt(13)):
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.7))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.color.rgb = BLACK
    run.font.bold = True
    run.font.name = FONT_NAME
    return txBox


def add_table(slide, left, top, width, height, rows, cols, data,
              header_color=None, col_widths=None):
    if header_color is None:
        header_color = BLUE_HEADER
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = str(data[r][c]) if data[r][c] is not None else ""
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
                for run in paragraph.runs:
                    run.font.size = Pt(11)
                    run.font.name = FONT_NAME
                    if r == 0:
                        run.font.color.rgb = WHITE
                        run.font.bold = True
                    else:
                        run.font.color.rgb = BLACK
                        run.font.bold = False
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_color
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r % 2 == 1 else RGBColor(0xF5, 0xF5, 0xF5)
            cell.margin_left = Pt(5)
            cell.margin_right = Pt(5)
            cell.margin_top = Pt(3)
            cell.margin_bottom = Pt(3)
    return table_shape


def add_bar_chart(slide, left, top, width, height,
                  categories, values, title="", bar_colors=None):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series('', values)
    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, chart_data
    )
    chart = chart_frame.chart
    chart.has_legend = False
    chart.chart_style = 2
    plot = chart.plots[0]
    plot.gap_width = 100
    series = plot.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = CYAN
    series.has_data_labels = True
    dl = series.data_labels
    dl.font.size = Pt(13)
    dl.font.bold = True
    dl.font.color.rgb = BLACK
    dl.number_format = '0.#'
    dl.show_value = True
    dl.label_position = XL_LABEL_POSITION.OUTSIDE_END
    if bar_colors:
        for i, color in enumerate(bar_colors):
            pt = series.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = color
    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.size = Pt(12)
    cat_axis.tick_labels.font.name = FONT_NAME
    cat_axis.major_tick_mark = 2
    cat_axis.format.line.fill.background()
    val_axis = chart.value_axis
    val_axis.visible = False
    val_axis.major_tick_mark = 2
    val_axis.format.line.fill.background()
    val_axis.major_gridlines.format.line.fill.background()
    if title:
        chart.has_title = True
        ct = chart.chart_title.text_frame.paragraphs[0]
        ct.text = title
        ct.font.size = Pt(14)
        ct.font.bold = True
        ct.font.name = FONT_NAME
    else:
        chart.has_title = False
    return chart_frame


def add_callout_label(slide, left, top, text, bg_color=None, font_size=Pt(11)):
    if bg_color is None:
        bg_color = CYAN
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(1.3), Inches(0.3)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.fill.background()
    tf = shape.text_frame
    tf.margin_left = Pt(4)
    tf.margin_right = Pt(4)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = text
    p.font.size = font_size
    p.font.color.rgb = WHITE
    p.font.bold = True
    p.font.name = FONT_NAME
    return shape


def add_data_card(slide, left, top, width, height, value, label,
                  value_color=None, bg_color=None):
    if value_color is None:
        value_color = CYAN
    if bg_color is None:
        bg_color = WHITE
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    shape.line.width = Pt(1)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(8)
    tf.margin_right = Pt(8)
    tf.margin_top = Pt(6)
    tf.margin_bottom = Pt(3)
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    run1 = p1.add_run()
    run1.text = str(value)
    run1.font.size = Pt(24)
    run1.font.color.rgb = value_color
    run1.font.bold = True
    run1.font.name = FONT_NAME
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = label
    run2.font.size = Pt(10)
    run2.font.color.rgb = GRAY_TEXT
    run2.font.bold = False
    run2.font.name = FONT_NAME
    return shape
'''


def _extract_code(response_text: str) -> str:
    """Extract ONLY the build_slide function from the AI response."""
    # Step 1: extract from markdown code block if present (```python / ```py / plain ```)
    code_blocks = re.findall(
        r'```(?:python|py)?\s*\n(.*?)```',
        response_text,
        re.DOTALL | re.IGNORECASE,
    )
    if code_blocks:
        # Prefer a block containing build_slide, fallback to the first block
        preferred = next(
            (blk for blk in code_blocks if re.search(r'def\s+build_slide\s*\(', blk)),
            code_blocks[0],
        )
        code = preferred.strip()
    else:
        code = response_text.strip()

    # Step 2: find the build_slide function and extract only it
    match = re.search(r'(def\s+build_slide\s*\(\s*slide[^)]*\)\s*:.*)', code, re.DOTALL)
    if not match:
        return code

    func_text = match.group(1)
    lines = func_text.split('\n')
    result = [lines[0]]  # "def build_slide(slide):"
    for line in lines[1:]:
        # Stop at the next top-level definition or non-indented code
        # (but allow blank lines and comments inside the function)
        if line and not line[0].isspace() and not line.startswith('#') and line.strip():
            break
        result.append(line)

    # Remove trailing blank lines
    while result and not result[-1].strip():
        result.pop()

    return '\n'.join(result)


def generate_slide_code(
    image_path: str,
    page_num: int,
    total_pages: int,
    model: str = DEFAULT_MODEL,
) -> str:
    """Send an infographic image to Gemini and get back python-pptx code."""
    user_prompt = PPT_CODE_GEN_USER_PROMPT.format(
        page_num=page_num,
        total_pages=total_pages,
    )
    response = generate_text_with_images(
        model=model,
        system_prompt=PPT_CODE_GEN_SYSTEM_PROMPT,
        user_prompt=user_prompt,
        image_paths=[image_path],
    )
    return _extract_code(response)


def _make_pptx_script(build_func_codes: list[tuple[str, str]], output_path: str) -> str:
    """Assemble a full runnable script.

    build_func_codes: list of (func_name, func_body) pairs.
    """
    parts = [_SCRIPT_HEADER]
    parts.append(f'\nOUTPUT_PATH = r"{output_path}"\n')

    slide_calls = []
    for i, (func_name, code) in enumerate(build_func_codes):
        parts.append(f"\n# ── Slide {i + 1} ──\n")
        parts.append(code)
        parts.append("\n")
        slide_calls.append(
            f"s{i} = prs.slides.add_slide(prs.slide_layouts[6])\n"
            f"{func_name}(s{i})"
        )

    parts.append(f"""
# ── Main ──
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

{chr(10).join(slide_calls)}
prs.save(OUTPUT_PATH)
""")
    return "\n".join(parts)


def _rename_func(code: str, new_name: str) -> str:
    """Rename build_slide -> new_name in the code."""
    renamed, count = re.subn(
        r'def\s+build_slide\s*\(\s*slide[^)]*\)\s*:',
        f'def {new_name}(slide):',
        code,
        count=1,
    )
    if count == 0:
        raise ValueError("未找到函数定义: def build_slide(slide)")
    return renamed


def build_single_slide_pptx(slide_code: str, output_path: str) -> tuple[bool, str]:
    """Generate a single-slide PPTX from one slide's code. Returns (success, error)."""
    func_name = "build_slide_1"
    try:
        renamed = _rename_func(slide_code, func_name)
    except ValueError as e:
        return False, str(e)
    script = _make_pptx_script([(func_name, renamed)], output_path)

    # Save assembled script for debugging (use _full suffix to avoid overwriting code file)
    script_path = output_path.replace(".pptx", "_full.py")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)

    return _exec_script(script)


def build_full_pptx(slide_codes: dict[int, str], output_path: str) -> tuple[bool, str]:
    """Generate a full PPTX from multiple slide codes.

    slide_codes: {page_num: code_string} (1-based page numbers).
    Returns (success, error).
    """
    func_pairs = []
    for page_num in sorted(slide_codes.keys()):
        func_name = f"build_slide_{page_num}"
        try:
            renamed = _rename_func(slide_codes[page_num], func_name)
        except ValueError as e:
            return False, f"第 {page_num} 页代码无效: {e}"
        func_pairs.append((func_name, renamed))

    script = _make_pptx_script(func_pairs, output_path)

    script_path = output_path.replace(".pptx", "_full.py")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)

    return _exec_script(script)


def _patch_common_errors(code: str) -> str:
    """Auto-fix common mistakes the AI makes in generated python-pptx code."""
    # .line.background() -> .line.fill.background()
    code = re.sub(r'\.line\.background\(\)', '.line.fill.background()', code)
    # .line.no_fill() -> .line.fill.background()
    code = re.sub(r'\.line\.no_fill\(\)', '.line.fill.background()', code)
    # MSO_CONNECTOR not imported -> add import if used
    if 'MSO_CONNECTOR' in code and 'from pptx.enum.shapes import MSO_CONNECTOR' not in code:
        code = 'from pptx.enum.shapes import MSO_CONNECTOR\n' + code
    # Replace invalid MSO_SHAPE members with ROUNDED_RECTANGLE
    from pptx.enum.shapes import MSO_SHAPE
    _valid_shapes = set(MSO_SHAPE.__members__.keys())
    def _fix_shape(m):
        name = m.group(1)
        if name in _valid_shapes:
            return m.group(0)
        return f'MSO_SHAPE.ROUNDED_RECTANGLE'
    code = re.sub(r'MSO_SHAPE\.([A-Z_0-9]+)', _fix_shape, code)
    return code


def _exec_script(script: str) -> tuple[bool, str]:
    """Execute a generated pptx script in-process. Returns (success, error)."""
    script = _patch_common_errors(script)
    try:
        exec(compile(script, "<pptx_gen>", "exec"), {"__builtins__": __builtins__})
        return True, ""
    except Exception:
        return False, traceback.format_exc()


# ── Slide code persistence ──

def get_slides_dir(proj_dir: str) -> str:
    d = os.path.join(proj_dir, "最终文档", "ppt_slides")
    os.makedirs(d, exist_ok=True)
    return d


def save_slide_code(proj_dir: str, page_num: int, code: str):
    path = os.path.join(get_slides_dir(proj_dir), f"slide_{page_num:02d}.py")
    with open(path, "w", encoding="utf-8") as f:
        f.write(code)


def load_slide_code(proj_dir: str, page_num: int) -> str | None:
    path = os.path.join(get_slides_dir(proj_dir), f"slide_{page_num:02d}.py")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return None


def load_all_slide_codes(proj_dir: str) -> dict[int, str]:
    """Return {page_num: code} for all saved slide codes."""
    slides_dir = get_slides_dir(proj_dir)
    result = {}
    for fname in sorted(os.listdir(slides_dir)):
        m = re.match(r"slide_(\d+)\.py$", fname)
        if m:
            page_num = int(m.group(1))
            with open(os.path.join(slides_dir, fname), "r", encoding="utf-8") as f:
                result[page_num] = f.read()
    return result


def get_single_pptx_path(proj_dir: str, page_num: int) -> str:
    return os.path.join(get_slides_dir(proj_dir), f"slide_{page_num:02d}.pptx")
