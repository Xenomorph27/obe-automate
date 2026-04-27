# backend/services/syllabus_service.py
import json
import asyncio
import fitz  # PyMuPDF

from fastapi import UploadFile

from backend.core.config import MAX_FILE_SIZE_BYTES, ALLOWED_FILE_TYPES
from backend.core.exceptions import FileValidationError, ExtractionError, LLMError
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger

logger = get_logger(__name__)


class SyllabusService:

    async def process_syllabus(self, file: UploadFile) -> dict:
        self._validate_file(file)
        contents = await file.read()
        if len(contents) > MAX_FILE_SIZE_BYTES:
            raise FileValidationError(
                f"File size exceeds {MAX_FILE_SIZE_BYTES // (1024 * 1024)}MB limit."
            )
        raw_text = self._extract_pdf_text(contents)
        structured_data = await self._extract_with_llm(raw_text)
        return structured_data

    def _validate_file(self, file: UploadFile):
        if file.content_type not in ALLOWED_FILE_TYPES:
            raise FileValidationError(
                f"Invalid file type '{file.content_type}'. Only PDF files are accepted."
            )

    def _extract_pdf_text(self, file_bytes: bytes) -> str:
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            full_text = ""
            for page in doc:
                full_text += page.get_text()
            doc.close()
            if not full_text.strip():
                raise ExtractionError(
                    "PDF appears to be empty or scanned-only (no extractable text)."
                )
            logger.debug(f"Extracted {len(full_text)} characters from PDF")
            return full_text
        except ExtractionError:
            raise
        except Exception as e:
            logger.exception("Failed to parse PDF")
            raise ExtractionError(f"Could not read PDF file: {str(e)}")

    async def _extract_with_llm(self, raw_text: str) -> dict:
        prompt = f"""You are an expert in engineering education and Outcome-Based Education (OBE).

Analyze the following syllabus text and extract structured information.

Return ONLY a valid JSON object — no explanation, no markdown, no code blocks.

The JSON must have this exact structure:
{{
  "course_name": "string",
  "course_code": "string or null",
  "units": [
    {{
      "unit_number": 1,
      "unit_title": "string",
      "topics": ["topic1", "topic2"]
    }}
  ],
  "course_outcomes": [
    {{
      "co_id": "CO1",
      "statement": "string",
      "bloom_level": "Remember|Understand|Apply|Analyze|Evaluate|Create"
    }}
  ]
}}

Syllabus text:
---
{raw_text[:8000]}
---
"""
        try:
            logger.debug("Sending syllabus to LLM for extraction")
            raw_response = await get_llm_response(prompt)

            if raw_response.startswith("```"):
                lines = raw_response.split("\n")
                raw_response = "\n".join(lines[1:-1])

            parsed = json.loads(raw_response.strip())
            logger.debug(f"LLM returned {len(parsed.get('course_outcomes', []))} COs")
            return parsed

        except json.JSONDecodeError:
            logger.error(f"LLM returned invalid JSON")
            raise LLMError("AI returned an unparseable response. Please try again.")
        except LLMError:
            raise
        except Exception as e:
            logger.exception("LLM call failed")
            raise LLMError(f"AI service unavailable: {str(e)}")