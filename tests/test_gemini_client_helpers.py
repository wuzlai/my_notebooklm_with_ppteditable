import unittest
from types import SimpleNamespace

from src.gemini_client import _extract_text_response


class TestGeminiClientHelpers(unittest.TestCase):
    def test_extract_text_direct(self):
        response = SimpleNamespace(text="hello")
        self.assertEqual(_extract_text_response(response), "hello")

    def test_extract_text_empty_string_is_valid(self):
        response = SimpleNamespace(text="")
        self.assertEqual(_extract_text_response(response), "")

    def test_extract_text_from_candidates(self):
        parts = [SimpleNamespace(text="A"), SimpleNamespace(text="B")]
        candidate = SimpleNamespace(content=SimpleNamespace(parts=parts))
        response = SimpleNamespace(text=None, candidates=[candidate])
        self.assertEqual(_extract_text_response(response), "A\nB")


if __name__ == "__main__":
    unittest.main()
