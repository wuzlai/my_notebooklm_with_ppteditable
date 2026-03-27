import unittest

from src.ppt_generator import _extract_code, _rename_func


class TestPptGeneratorHelpers(unittest.TestCase):
    def test_extract_code_prefers_block_with_build_slide(self):
        response = (
            "```python\nprint('helper')\n```\n"
            "```Python\n"
            "def build_slide(slide: object):\n"
            "    a = 1\n"
            "    return a\n"
            "\n"
            "def another():\n"
            "    pass\n"
            "```"
        )
        code = _extract_code(response)
        self.assertIn("def build_slide", code)
        self.assertNotIn("def another", code)

    def test_rename_func_supports_typed_signature(self):
        code = "def build_slide(slide: object):\n    return 1\n"
        renamed = _rename_func(code, "build_slide_9")
        self.assertIn("def build_slide_9(slide):", renamed)

    def test_rename_func_raises_on_invalid_input(self):
        with self.assertRaises(ValueError):
            _rename_func("def other(slide):\n    pass", "build_slide_1")


if __name__ == "__main__":
    unittest.main()
