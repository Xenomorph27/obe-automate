# backend/core/llm.py
"""
Multi-provider LLM fallback chain
----------------------------------
Call order: Gemini 2.0 Flash → Groq (llama-3.3-70b) → OpenAI (gpt-4o-mini)

Vision fallback chain (for image/page transcription)
-----------------------------------------------------
Call order: Gemini 2.0 Flash (vision) → OpenAI gpt-4o-mini (vision)
Groq's llama-3.3-70b-versatile is text-only and is intentionally skipped here.
"""

import asyncio
import base64
from backend.core.config import GEMINI_API_KEY, GROQ_API_KEY, OPENAI_API_KEY
from backend.core.exceptions import LLMError
from backend.core.logger import get_logger

logger = get_logger(__name__)


async def _call_gemini(prompt: str) -> str:
    from google import genai
    def _sync_call():
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt,
            )
            return response.text.strip()
        except Exception as e:
            # Re-raise as standard Exception so outer try/except catches it
            raise Exception(f"Gemini error: {e}") from e
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync_call)


async def _call_groq(prompt: str) -> str:
    from groq import Groq
    def _sync_call():
        client = Groq(api_key=GROQ_API_KEY)
        completion = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        return completion.choices[0].message.content.strip()
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync_call)


async def _call_openai(prompt: str) -> str:
    from openai import OpenAI
    def _sync_call():
        client = OpenAI(api_key=OPENAI_API_KEY)
        completion = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        return completion.choices[0].message.content.strip()
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync_call)


async def get_llm_response(prompt: str) -> str:
    providers = []
    if GEMINI_API_KEY:
        providers.append(("Gemini", _call_gemini))
    if GROQ_API_KEY:
        providers.append(("Groq", _call_groq))
    if OPENAI_API_KEY:
        providers.append(("OpenAI", _call_openai))
    if not providers:
        raise LLMError("No LLM API keys configured.")
    last_error = None
    for name, fn in providers:
        try:
            logger.info(f"Calling {name}...")
            result = await fn(prompt)
            logger.info(f"{name} responded successfully")
            return result
        except Exception as e:
            logger.warning(f"{name} failed: {e}. Trying next provider...")
            last_error = e
    raise LLMError(f"All LLM providers failed. Last error: {last_error}")


# ── Vision (image → text) chain ─────────────────────────────────────────
# Used for pages/screenshots embedded as raster images, where PyMuPDF's
# text layer is empty or incomplete. Only routed to providers whose
# models actually accept image input.

async def _call_gemini_vision(image_bytes: bytes, mime_type: str, prompt: str) -> str:
    from google import genai
    from google.genai import types

    def _sync_call():
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=[
                    types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
                    prompt,
                ],
            )
            return response.text.strip()
        except Exception as e:
            raise Exception(f"Gemini vision error: {e}") from e

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync_call)


async def _call_openai_vision(image_bytes: bytes, mime_type: str, prompt: str) -> str:
    from openai import OpenAI

    def _sync_call():
        client = OpenAI(api_key=OPENAI_API_KEY)
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        completion = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:{mime_type};base64,{b64}"},
                        },
                    ],
                }
            ],
            temperature=0.1,
        )
        return completion.choices[0].message.content.strip()

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync_call)


async def get_vision_response(image_bytes: bytes, prompt: str, mime_type: str = "image/png") -> str:
    """
    Sends a single image to a vision-capable LLM and returns the transcription/
    description text. Falls back across providers the same way get_llm_response
    does, but skips Groq's llama-3.3-70b-versatile since it is text-only.
    """
    providers = []
    if GEMINI_API_KEY:
        providers.append(("Gemini-Vision", _call_gemini_vision))
    if OPENAI_API_KEY:
        providers.append(("OpenAI-Vision", _call_openai_vision))
    if not providers:
        raise LLMError("No vision-capable LLM API keys configured (need Gemini or OpenAI).")
    last_error = None
    for name, fn in providers:
        try:
            logger.info(f"Calling {name}...")
            result = await fn(image_bytes, mime_type, prompt)
            logger.info(f"{name} responded successfully")
            return result
        except Exception as e:
            logger.warning(f"{name} failed: {e}. Trying next vision provider...")
            last_error = e
    raise LLMError(f"All vision LLM providers failed. Last error: {last_error}")
