from src.gemini_client import generate_image
from src.prompts import SLIDE_IMAGE_PROMPT_TEMPLATE

DEFAULT_MODEL = "gemini-3-pro-image-preview"


def generate_slide_image(
    slide_content: str,
    style_desc: str,
    page_num: int,
    total_pages: int,
    model: str = DEFAULT_MODEL,
) -> bytes | None:
    """Generate an infographic image for a single slide."""
    prompt = SLIDE_IMAGE_PROMPT_TEMPLATE.format(
        style_description=style_desc,
        page_num=page_num,
        slide_content=slide_content,
        total_pages=total_pages,
    )
    return generate_image(model=model, prompt=prompt)
