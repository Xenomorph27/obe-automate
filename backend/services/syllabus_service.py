# backend/services/syllabus_service.py
import json
import fitz  # PyMuPDF

from fastapi import UploadFile
from google import genai
from google.genai import types

from backend.core.config import GEMINI_API_KEY, MAX_FILE_SIZE_BYTES, ALLOWED_FILE_TYPES
from backend.core.exceptions import FileValidationError, ExtractionError, LLMError
from backend.core.logger import get_logger

logger = get_logger(__name__)

# Initialise the new client
client = genai.Client(api_key=GEMINI_API_KEY)


class SyllabusService:

    async def process_syllabus(self, file: UploadFile) -> dict:
        """Full pipeline: validate → extract text → call Gemini → return structured data."""

        # Step 1: Validate file type
        self._validate_file(file)

        # Step 2: Read file bytes
        contents = await file.read()

        # Step 3: Validate file size
        if len(contents) > MAX_FILE_SIZE_BYTES:
            raise FileValidationError(
                f"File size exceeds {MAX_FILE_SIZE_BYTES // (1024 * 1024)}MB limit."
            )

        # Step 4: Extract raw text from PDF
        raw_text = self._extract_pdf_text(contents)

        # Step 5: Send to Gemini and parse response
        structured_data = self._extract_with_gemini(raw_text)

        return structured_data

    def _validate_file(self, file: UploadFile):
        """Check file type before reading any bytes."""
        if file.content_type not in ALLOWED_FILE_TYPES:
            raise FileValidationError(
                f"Invalid file type '{file.content_type}'. Only PDF files are accepted."
            )

    def _extract_pdf_text(self, file_bytes: bytes) -> str:
        """Use PyMuPDF to extract all text from PDF pages."""
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

    def _extract_with_gemini(self, raw_text: str) -> dict:
        """Send syllabus text to Gemini and get back structured JSON."""

        prompt = f"""
You are an expert in engineering education and Outcome-Based Education (OBE).

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
            logger.debug("Sending syllabus to Gemini for extraction")

            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt,
            )

            raw_response = response.text.strip()

            # Clean up — Gemini sometimes wraps output in markdown fences
            if raw_response.startswith("```"):
                lines = raw_response.split("\n")
                raw_response = "\n".join(lines[1:-1])

            parsed = json.loads(raw_response)
            logger.debug(
                f"Gemini returned {len(parsed.get('course_outcomes', []))} COs"
            )
            return parsed

        except json.JSONDecodeError:
            logger.error(f"Gemini returned invalid JSON: {raw_response[:200]}")
            raise LLMError("AI returned an unparseable response. Please try again.")

        except Exception as e:
            logger.exception("Gemini API call failed")
            raise LLMError(f"AI service unavailable: {str(e)}")