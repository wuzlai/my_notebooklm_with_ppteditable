import os
import img2pdf
from PIL import Image


def build_pdf(image_dir: str, output_path: str) -> str:
    """Merge numbered images in image_dir into a single PDF.

    Images are sorted by filename (expected format: 01.jpg, 02.jpg, ...).
    Returns the output path.
    """
    image_files = sorted(
        f for f in os.listdir(image_dir)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    )
    if not image_files:
        raise ValueError(f"No images found in {image_dir}")

    image_paths = [os.path.join(image_dir, f) for f in image_files]

    # Convert any non-JPEG images to JPEG for img2pdf compatibility
    jpeg_paths = []
    for path in image_paths:
        if path.lower().endswith(".png"):
            img = Image.open(path).convert("RGB")
            jpeg_path = path.rsplit(".", 1)[0] + ".jpg"
            img.save(jpeg_path, "JPEG", quality=95)
            jpeg_paths.append(jpeg_path)
        else:
            jpeg_paths.append(path)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(jpeg_paths))

    return output_path
