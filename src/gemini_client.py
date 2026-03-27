import os
from google import genai
from google.genai import types
from dotenv import load_dotenv

load_dotenv()

_client = None


def get_client() -> genai.Client:
    global _client
    if _client is None:
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY not found in environment variables")
        _client = genai.Client(api_key=api_key)
    return _client


def _extract_text_response(response) -> str:
    """Best-effort extraction for text responses across SDK variants."""
    text = getattr(response, "text", None)
    if text is not None:
        return text

    candidates = getattr(response, "candidates", None) or []
    for candidate in candidates:
        content = getattr(candidate, "content", None)
        parts = getattr(content, "parts", None) or []
        text_chunks = [getattr(part, "text", None) for part in parts if getattr(part, "text", None)]
        if text_chunks:
            return "\n".join(text_chunks)
    raise ValueError("Gemini response does not contain text content")


def generate_text(model: str, system_prompt: str, user_prompt: str) -> str:
    client = get_client()
    response = client.models.generate_content(
        model=model,
        contents=user_prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.7,
        ),
    )
    return _extract_text_response(response)


def generate_text_with_images(
    model: str, system_prompt: str, user_prompt: str, image_paths: list[str]
) -> str:
    """Send text + images to Gemini and get text response."""
    client = get_client()
    parts = []
    for img_path in image_paths:
        with open(img_path, "rb") as f:
            img_bytes = f.read()
        mime = "image/png" if img_path.lower().endswith(".png") else "image/jpeg"
        parts.append(types.Part.from_bytes(data=img_bytes, mime_type=mime))
    parts.append(types.Part.from_text(text=user_prompt))
    response = client.models.generate_content(
        model=model,
        contents=parts,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.3,
        ),
    )
    return _extract_text_response(response)


def generate_image(model: str, prompt: str) -> bytes | None:
    client = get_client()
    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_modalities=["IMAGE", "TEXT"],
        ),
    )
    candidates = getattr(response, "candidates", None) or []
    if not candidates:
        return None
    parts = getattr(candidates[0].content, "parts", None) or []
    for part in parts:
        inline_data = getattr(part, "inline_data", None)
        if inline_data is not None:
            return inline_data.data
    return None
