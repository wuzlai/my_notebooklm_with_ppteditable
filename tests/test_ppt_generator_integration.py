import os
import tempfile
import unittest

from src.ppt_generator import build_full_pptx, build_single_slide_pptx


MINIMAL_SLIDE_CODE = """\
def build_slide(slide):
    title = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    title.text_frame.text = "Hello PPT"
"""


class TestPptGeneratorIntegration(unittest.TestCase):
    def test_build_single_slide_pptx_creates_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "single.pptx")
            ok, err = build_single_slide_pptx(MINIMAL_SLIDE_CODE, output_path)
            self.assertTrue(ok, msg=err)
            self.assertTrue(os.path.exists(output_path))
            self.assertGreater(os.path.getsize(output_path), 0)

    def test_build_full_pptx_creates_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "full.pptx")
            slide_codes = {
                1: MINIMAL_SLIDE_CODE,
                2: MINIMAL_SLIDE_CODE.replace("Hello PPT", "Slide 2"),
            }
            ok, err = build_full_pptx(slide_codes, output_path)
            self.assertTrue(ok, msg=err)
            self.assertTrue(os.path.exists(output_path))
            self.assertGreater(os.path.getsize(output_path), 0)


if __name__ == "__main__":
    unittest.main()
