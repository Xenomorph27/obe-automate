# backend/services/syllabus_service.py
import json
import re
import asyncio
import fitz  # PyMuPDF

from fastapi import UploadFile

from backend.core.config import MAX_FILE_SIZE_BYTES, ALLOWED_FILE_TYPES
from backend.core.exceptions import FileValidationError, ExtractionError, LLMError
from backend.core.llm import get_llm_response, get_vision_response
from backend.core.logger import get_logger

logger = get_logger(__name__)

# A page is treated as "image-heavy" (and gets a vision pass) when either:
#  - it has very little extractable text but does contain raster images
#    (classic case: syllabus table/unit list pasted in as a screenshot), or
#  - its embedded images cover a large fraction of the page area even if
#    some text is also present (e.g. a scanned page with a thin text header).
MIN_TEXT_CHARS_PER_PAGE = 120
IMAGE_AREA_COVERAGE_THRESHOLD = 0.35
MAX_VISION_PAGES = 8  # safety cap so a huge scanned PDF can't trigger 50+ API calls

VISION_PROMPT = """You are looking at one page of an engineering college course syllabus (PDF).
This page's content was embedded as an image/screenshot rather than as selectable text, so
transcribe everything relevant to the syllabus that is visible in the image.

Include, if present: course name, course code, credits, unit/module numbers and titles,
topics within each unit, course outcomes (COs) / learning objectives, and any tables of
content from the course outline.

Output plain text only — no markdown, no commentary, no preamble. Preserve unit/topic
structure using line breaks. If the image contains nothing relevant to a syllabus
(e.g. it's a logo, signature, or decorative banner), output exactly: NO_SYLLABUS_CONTENT"""


class SyllabusService:

    async def process_syllabus(self, file: UploadFile) -> dict:
        self._validate_file(file)
        contents = await file.read()
        if len(contents) > MAX_FILE_SIZE_BYTES:
            raise FileValidationError(
                f"File size exceeds {MAX_FILE_SIZE_BYTES // (1024 * 1024)}MB limit."
            )
        raw_text = await self._extract_pdf_text(contents)
        structured_data = await self._extract_with_llm(raw_text)
        return structured_data

    def _validate_file(self, file: UploadFile):
        if file.content_type not in ALLOWED_FILE_TYPES:
            raise FileValidationError(
                f"Invalid file type '{file.content_type}'. Only PDF files are accepted."
            )

    async def _extract_pdf_text(self, file_bytes: bytes) -> str:
        """
        Extracts text from the PDF's text layer AND, for any page whose content
        is mostly/entirely embedded as an image (scanned page, pasted screenshot,
        table-as-picture), runs a vision-model pass over that page so the
        syllabus content baked into the image isn't silently dropped.
        """
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
        except Exception as e:
            logger.exception("Failed to open PDF")
            raise ExtractionError(f"Could not read PDF file: {str(e)}")

        try:
            page_texts = []
            vision_candidates = []  # (page_index, has_some_text)

            for i, page in enumerate(doc):
                page_text = page.get_text()
                page_texts.append(page_text)

                images = page.get_images(full=True)
                if not images:
                    continue

                page_area = page.rect.width * page.rect.height
                image_area = 0.0
                for img in images:
                    try:
                        # img[0] is the xref; bbox lookup gives the placed rect on the page
                        for rect in page.get_image_rects(img[0]):
                            image_area += rect.width * rect.height
                    except Exception:
                        continue
                coverage = (image_area / page_area) if page_area else 0

                sparse_text = len(page_text.strip()) < MIN_TEXT_CHARS_PER_PAGE
                image_heavy = coverage >= IMAGE_AREA_COVERAGE_THRESHOLD

                if sparse_text or image_heavy:
                    vision_candidates.append(i)

            if len(vision_candidates) > MAX_VISION_PAGES:
                logger.warning(
                    f"{len(vision_candidates)} image-heavy pages found, capping vision "
                    f"pass at {MAX_VISION_PAGES} to limit API usage"
                )
                vision_candidates = vision_candidates[:MAX_VISION_PAGES]

            vision_text_by_page = {}
            if vision_candidates:
                logger.info(
                    f"Running vision extraction on {len(vision_candidates)} "
                    f"image-heavy page(s): {vision_candidates}"
                )
                vision_text_by_page = await self._extract_text_from_pages(doc, vision_candidates)

            full_text_parts = []
            for i, page_text in enumerate(page_texts):
                vision_text = vision_text_by_page.get(i, "")
                if vision_text and vision_text.strip() != "NO_SYLLABUS_CONTENT":
                    full_text_parts.append(page_text + "\n" + vision_text)
                else:
                    full_text_parts.append(page_text)
            full_text = "\n".join(full_text_parts)

            doc.close()

            if not full_text.strip():
                raise ExtractionError(
                    "PDF appears to be empty or scanned-only (no extractable text "
                    "and vision extraction found nothing usable)."
                )
            logger.debug(f"Extracted {len(full_text)} characters from PDF "
                         f"({len(vision_candidates)} page(s) via vision pass)")
            return full_text
        except ExtractionError:
            raise
        except Exception as e:
            logger.exception("Failed to parse PDF")
            raise ExtractionError(f"Could not read PDF file: {str(e)}")

    async def _extract_text_from_pages(self, doc, page_indices: list) -> dict:
        """Rasterizes each candidate page and runs it through the vision LLM chain
        concurrently. A failure on one page is logged and skipped, not fatal —
        the rest of the document (and that page's text layer, if any) still gets used."""

        async def _render_and_extract(i: int):
            try:
                page = doc[i]
                pix = page.get_pixmap(dpi=200)
                image_bytes = pix.tobytes("png")
                text = await get_vision_response(image_bytes, VISION_PROMPT, mime_type="image/png")
                return i, text
            except Exception as e:
                logger.warning(f"Vision extraction failed for page {i + 1}: {e}")
                return i, ""

        results = await asyncio.gather(*[_render_and_extract(i) for i in page_indices])
        return dict(results)

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
        text_chunk = raw_text[:16000]

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
