# backend/services/syllabus_service.py
import json
import re
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

    def _clean_json_response(self, raw: str) -> str:
        """Strip markdown fences and extract the first JSON object from the response."""
        # Remove ```json or ``` fences
        raw = raw.strip()
        if raw.startswith("```"):
            lines = raw.split("\n")
            # Drop first line (```json) and last line (```)
            raw = "\n".join(lines[1:])
            if raw.endswith("```"):
                raw = raw[: raw.rfind("```")]
        raw = raw.strip()

        # If response has extra text before/after JSON, extract just the JSON object
        start = raw.find("{")
        end = raw.rfind("}")
        if start != -1 and end != -1:
            raw = raw[start : end + 1]

        return raw

    async def _extract_with_llm(self, raw_text: str) -> dict:
        # Use first 10000 chars to capture more of the syllabus
        text_chunk = raw_text[:10000]

        prompt = f"""You are an academic data extraction assistant specializing in engineering college syllabi for Outcome-Based Education (OBE) systems.

TASK: Extract structured information from the syllabus text below.

CRITICAL RULES:
1. Return ONLY a raw JSON object. No explanation, no markdown, no code blocks, no backticks.
2. Your entire response must start with {{ and end with }}
3. Never return null for any field — use empty string "" or empty array [] if not found.
4. For credits: look for "Course Credit", "Credits", "Credit Hours" — extract the number as a string.
5. For course_outcomes: Look for sections labelled "Course Outcomes", "COs", or "Learning Objectives". 
   - If you find "Course Outcomes" or "COs" → use those directly as CO1, CO2, CO3...
   - If you only find "Learning Objectives" → treat each objective as a CO (CO1, CO2, CO3...).
   - NEVER leave course_outcomes empty if any objectives or outcomes exist in the text.
6. For bloom_level: infer from the action verb in the statement. 
   - Explain/Describe/List/Define → Understand
   - Apply/Use/Implement/Solve → Apply
   - Analyze/Compare/Distinguish → Analyze
   - Evaluate/Justify/Assess → Evaluate
   - Design/Create/Build/Develop → Create
   - Remember/Recall/Identify → Remember
7. Extract all units/modules/topics from the Course Outline section.

OUTPUT FORMAT (copy this structure exactly):
{{
  "course_name": "Full course name as written in syllabus",
  "course_code": "Course code e.g. TE7760",
  "credits": "3",
  "units": [
    {{
      "unit_number": 1,
      "unit_title": "Title of Unit 1",
      "topics": ["Topic A", "Topic B", "Topic C"]
    }},
    {{
      "unit_number": 2,
      "unit_title": "Title of Unit 2",
      "topics": ["Topic D", "Topic E"]
    }}
  ],
  "course_outcomes": [
    {{
      "co_id": "CO1",
      "statement": "Full statement as written in the syllabus",
      "bloom_level": "Understand"
    }},
    {{
      "co_id": "CO2",
      "statement": "Full statement as written in the syllabus",
      "bloom_level": "Apply"
    }}
  ]
}}

SYLLABUS TEXT:
---
{text_chunk}
---

Remember: Start your response with {{ immediately. No preamble. Extract ALL objectives/outcomes — do not skip any."""

        try:
            logger.info("Sending syllabus to LLM for extraction")
            raw_response = await get_llm_response(prompt)
            logger.debug(f"LLM raw response length: {len(raw_response)}")

            cleaned = self._clean_json_response(raw_response)
            parsed = json.loads(cleaned)

            # Validate minimum required fields
            if not parsed.get("course_name"):
                logger.warning("LLM returned empty course_name — attempting repair")
                parsed["course_name"] = self._fallback_course_name(raw_text)

            if not parsed.get("course_outcomes"):
                logger.warning("LLM returned no course outcomes")
                parsed["course_outcomes"] = []

            if not parsed.get("units"):
                parsed["units"] = []

            co_count = len(parsed.get("course_outcomes", []))
            unit_count = len(parsed.get("units", []))
            logger.info(f"Extraction complete: {co_count} COs, {unit_count} units, course='{parsed.get('course_name')}'")
            return parsed

        except json.JSONDecodeError as e:
            logger.error(f"LLM returned invalid JSON: {e}")
            logger.debug(f"Raw response was: {raw_response[:500]}")
            raise LLMError("AI returned an unparseable response. Please try again.")
        except LLMError:
            raise
        except Exception as e:
            logger.exception("LLM call failed")
            raise LLMError(f"AI service unavailable: {str(e)}")

    def _fallback_course_name(self, raw_text: str) -> str:
        """Best-effort course name extraction from raw text when LLM fails."""
        lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
        for line in lines[:20]:
            # Skip very short or very long lines
            if 5 < len(line) < 100:
                # Skip lines that look like headers/metadata
                lower = line.lower()
                if not any(kw in lower for kw in ["university", "department", "semester", "year", "page", "credit"]):
                    return line
        return "Unknown Course"
