# backend/core/llm.py
"""
Multi-provider LLM fallback chain
----------------------------------
Call order: Gemini 2.0 Flash → Groq (llama-3.3-70b) → OpenAI (gpt-4o-mini)

Usage (anywhere in services):
    from backend.core.llm import get_llm_response
    raw_text = await get_llm_response(prompt)

The function returns the model's raw text response (string).
The caller is responsible for stripping markdown fences and JSON-parsing,
exactly as before — this layer only handles transport + fallback.

Install requirements (add to requirements.txt / pip install):
    groq
    openai
"""

import asyncio
from backend.core.config import GEMINI_API_KEY, GROQ_API_KEY, OPENAI_API_KEY
from backend.core.exceptions import LLMError
from backend.core.logger import get_logger

logger = get_logger(__name__)


# ---------------------------------------------------------------------------
# Provider implementations
# ---------------------------------------------------------------------------

async def _call_gemini(prompt: str) -> str:
    """Call Gemini 2.0 Flash via google-genai (sync SDK wrapped in executor)."""
    from google import genai

    def _sync_call():
        client = genai.Client(api_key=GEMINI_API_KEY)
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt,
        )
        return response.text.strip()

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync_call)


async def _call_groq(prompt: str) -> str:
    """Call Groq (llama-3.3-70b-versatile) — fast, generous free tier."""
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
    """Call OpenAI gpt-4o-mini — last resort fallback."""
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


# ---------------------------------------------------------------------------
# Public interface
# ---------------------------------------------------------------------------

async def get_llm_response(prompt: str) -> str:
    """
    Try each provider in order. Returns the first successful raw text response.
    Raises LLMError only if ALL providers fail.
    """
    providers = []

    if GEMINI_API_KEY:
        providers.append(("Gemini", _call_gemini))
    if GROQ_API_KEY:
        providers.append(("Groq", _call_groq))
    if OPENAI_API_KEY:
        providers.append(("OpenAI", _call_openai))

    if not providers:
        raise LLMError("No LLM API keys configured. Set at least one of: GEMINI_API_KEY, GROQ_API_KEY, OPENAI_API_KEY")

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