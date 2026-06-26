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
MAX_VISION_PAGES = 8
MAX_LLM_CHARS = 24000   # Gemini 2.0 Flash handles long contexts easily

# Keywords that strongly signal a page contains syllabus-relevant content
_RELEVANT_KEYWORDS = [
    "course outcome", "learning objective", "unit", "module", "syllabus",
    "co1", "co2", "co3", "co4", "co5",
    "program outcome", "peo", "pso", "bloom", "course credit",
    "course code", "course name", "dimensionality", "clustering",
    "deep learning", "autoencoder", "neural", "unsupervised",
    "contact hours", "course outline", "prerequisite",
]

# Keywords that strongly signal a page is bulk filler — skip it
_FILLER_KEYWORDS = [
    "prn", "roll no", "sign", "signature", "attendance", "absent",
    "cgpa", "sgpa", "result declaration", "marks declaration",
    "dear students", "submission deadline", "google meet",
    "inbox", "reply", "forward", "gmail", "moodle",
    "sr. no", "sr.no",
]

def _page_relevance_score(text: str) -> int:
    """
    Returns a score for how relevant a page is to syllabus extraction.
    Positive = relevant, negative = filler, 0 = neutral.
    """
    lower = text.lower()
    score = 0
    for kw in _RELEVANT_KEYWORDS:
        if kw in lower:
            score += 2
    for kw in _FILLER_KEYWORDS:
        if kw in lower:
            score -= 3
    # Penalise pages that are mostly numeric rows (student lists, marks sheets)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if lines:
        numeric_lines = sum(1 for l in lines if sum(c.isdigit() for c in l) / max(len(l), 1) > 0.4)
        if numeric_lines / len(lines) > 0.5:
            score -= 4
    return score

VISION_PROMPT = """You are an OCR engine reading one page of an engineering college course document.

YOUR ONLY JOB IS TO COPY TEXT EXACTLY AS IT APPEARS. Do not summarize, rephrase, shorten, or infer anything.

STRICT RULES:
- Copy every word VERBATIM — character by character, exactly as printed.
- Preserve original capitalization, punctuation, and spacing.
- For tables: copy each cell value exactly. Use " | " to separate columns and a new line for each row.
- For numbered lists: copy the number and text exactly as printed.
- Do NOT paraphrase. Do NOT add context. Do NOT omit words for brevity.
- If a unit title says "Introduction to Unsupervised Learning" copy that EXACTLY — never shorten to "Introduction to Machine Learning" or anything else.
- If the image contains no text at all (pure logo, photo, decorative graphic), output exactly: NO_SYLLABUS_CONTENT

Output: plain copied text only. No markdown, no commentary, no preamble."""


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

            # ── Build final text: vision pages first, then relevance-ranked ──
            # Vision pages MUST come first in the budget — they contain the
            # course outline table (unit titles) which is the most critical
            # structured content. Text-layer pages fill the remaining budget
            # sorted by relevance score, skipping bulk filler.

            budget = MAX_LLM_CHARS
            selected = []  # (original_page_index, text)

            # 1. Pin all vision-extracted pages at the top (they earned their place)
            for i in sorted(vision_text_by_page.keys()):
                if budget <= 0:
                    break
                vision_text = vision_text_by_page.get(i, "")
                if not vision_text or vision_text.strip() == "NO_SYLLABUS_CONTENT":
                    continue
                combined = page_texts[i] + "\n" + vision_text
                chunk = combined[:budget]
                selected.append((i, chunk))
                budget -= len(chunk)

            # 2. Fill remaining budget with text-layer pages, highest score first
            vision_indices = set(vision_text_by_page.keys())
            text_pages = []
            for i, page_text in enumerate(page_texts):
                if i in vision_indices:
                    continue  # already handled above
                score = _page_relevance_score(page_text)
                text_pages.append((score, i, page_text))

            text_pages.sort(key=lambda x: (-x[0], x[1]))
            for score, i, page_text in text_pages:
                if budget <= 0:
                    break
                chunk = page_text[:budget]
                selected.append((i, chunk))
                budget -= len(chunk)

            # 3. Re-sort into document order so context reads coherently
            selected.sort(key=lambda x: x[0])
            full_text = "\n".join(text for _, text in selected)

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
        text_chunk = raw_text  # budget already enforced by page relevance filter

        prompt = f"""You are an academic data extraction assistant specializing in engineering college syllabi for Outcome-Based Education (OBE) systems.

TASK: Extract structured information from the syllabus text below.

CRITICAL RULES:
1. Return ONLY a raw JSON object. No explanation, no markdown, no code blocks, no backticks.
2. Your entire response must start with {{ and end with }}
3. Never return null for any field — use empty string "" or empty array [] if not found.
4. For credits: look for "Course Credit", "Credits", "Credit Hours" — extract the number as a string.
5. VERBATIM COPY RULE — THIS IS THE MOST IMPORTANT RULE:
   For the following fields, copy the text EXACTLY as it appears in the source. Do NOT rephrase,
   shorten, paraphrase, or infer. Copy character by character:
   - unit_title: copy the exact unit/module heading as printed. If it says
     "Introduction to Unsupervised Learning" write that exactly — never write
     "Introduction to Machine Learning" or any other variation.
   - CO statement: copy the exact sentence as written. Do not shorten it.
   - course_name, course_code: copy exactly as printed.
   - PO, PEO, PSO statements: copy exactly as written.
   If you are not 100% sure of a word, copy what you see. Never substitute a different word.
6. There are FOUR distinct outcome types in this document. Do not mix them up, and do not let
   one bucket "borrow" content that belongs in another:
   - course_outcomes (COs): SPECIFIC to this one course only. Usually 4-8 items. Found under a
     heading like "Course Outcomes", "COs", or "Learning Objectives" that is scoped to this
     course's syllabus (not the department/program in general).
   - program_outcomes (POs): The standard generic 12-item engineering graduate list (PO1-PO12),
     found under "Program Outcomes" or "POs". These are IDENTICAL across every course in the
     department — if you see items like "Engineering Knowledge", "Problem analysis", "Modern
     tool usage", "Ethics", "Life-long learning" etc., these are POs, NOT course outcomes.
   - peos (PEOs): 3-5 broad career/professional statements under "Program Educational
     Objectives" — these describe what graduates will achieve years after graduating, not
     course-level skills. NOT course outcomes.
   - psos (PSOs): 2-4 department-specialization statements under "Program Specific Outcomes" —
     these describe the specific specialization area (e.g. AI/ML, VLSI, etc). NOT course outcomes.
   - If you only find "Learning Objectives" with no separate PO/PEO/PSO section, treat each
     objective as a CO (CO1, CO2, CO3...). NEVER put PO/PEO/PSO content into course_outcomes.
   - If a genuine course_outcomes section is not present in the text at all, return an empty
     array for course_outcomes — do NOT substitute PO/PEO/PSO content as a fallback.
   - learning_outcomes (LOs): statements listed under "Learning Outcomes" or "Course Learning
     Outcomes" (distinct from Course Outcomes / COs). Copy verbatim. If not present, return [].
7. For bloom_level: this is the ONLY field where you infer rather than copy — infer from the
   action verb in the CO statement:
   - Explain/Describe/List/Define/Understand → Understand
   - Apply/Use/Implement/Solve/Demonstrate → Apply
   - Analyze/Compare/Distinguish/Contrast → Analyze
   - Evaluate/Justify/Assess → Evaluate
   - Design/Create/Build/Develop/Model → Create
   - Remember/Recall/Identify → Remember
8. For units: look ONLY in the "Course Outline" table (columns: Sr.No. | Topic | Contact Hours).
   - unit_title: copy the FIRST LINE / main heading of each row exactly as printed.
     e.g. "Introduction to Unsupervised Learning -" → unit_title is "Introduction to Unsupervised Learning"
   - topics: copy ALL the remaining sub-topic text from that same row, split into individual
     topic strings. Do NOT truncate, do NOT summarize. Every sub-topic listed in that row
     must appear as a separate string in the topics array, copied verbatim.
     e.g. for Unit 1 the topics array should contain strings like:
       "Introduction to Machine Learning, applications",
       "Types of Learning: Supervised, Unsupervised and Semi-Supervised Learning",
       "Data Types and distance measures Numeric Data and Euclidean Distance, Categorical data, Graph data, Spatial data, Trajectory data, Time Series Data and Distance measures, Manhattan, Minkowski, Chessboard and others"
   - NEVER use session plan lecture topics (Lect. No. 1, 2, 3...) as unit titles or topics.
     The session plan is a different section entirely. The Course Outline table is the source.
   - If the Course Outline table is not present in the text (it may be image-only), return [].

OUTPUT FORMAT (copy this structure exactly):
{{
  "course_name": "Full course name exactly as written",
  "course_code": "Course code exactly as written",
  "credits": "3",
  "units": [
    {{
      "unit_number": 1,
      "unit_title": "Introduction to Unsupervised Learning",
      "topics": [
        "Introduction to Machine Learning, applications",
        "Types of Learning: Supervised, Unsupervised and Semi-Supervised Learning",
        "Data Types and distance measures Numeric Data and Euclidean Distance, Categorical data, Graph data, Spatial data, Trajectory data, Time Series Data and Distance measures, Manhattan, Minkowski, Chessboard and others"
      ]
    }},
    {{
      "unit_number": 2,
      "unit_title": "Dimensionality Reduction Techniques",
      "topics": [
        "Data size, Feature size and scalability issues in Machine Learning",
        "Linear Discriminate Analysis, Principal Component Analysis, Independent Component Analysis, Non-Negative Matrix Factorization and types, Singular Value Decomposition, Manifold Learning methods: MDS and T-SNE, normalization of input data, Density estimation"
      ]
    }}
  ],
  "course_outcomes": [
    {{
      "co_id": "CO1",
      "statement": "VERBATIM full statement as written",
      "bloom_level": "Analyze"
    }}
  ],
  "program_outcomes": [
    {{ "po_id": "PO1", "statement": "VERBATIM full statement" }}
  ],
  "peos": [
    {{ "peo_id": "PEO1", "statement": "VERBATIM full statement" }}
  ],
  "psos": [
    {{ "pso_id": "PSO1", "statement": "VERBATIM full statement" }}
  ],
  "learning_outcomes": [
    {{ "lo_id": "LO1", "statement": "VERBATIM full statement" }}
  ]
}}

SYLLABUS TEXT:
---
{text_chunk}
---

Remember: Start your response with {{ immediately. No preamble. COPY DON'T PARAPHRASE for all fields except bloom_level."""

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

            for key in ("course_outcomes", "program_outcomes", "peos", "psos", "units", "learning_outcomes"):
                if not parsed.get(key):
                    parsed[key] = []

            if not parsed.get("course_outcomes"):
                logger.warning("LLM returned no course outcomes")

            co_count   = len(parsed.get("course_outcomes", []))
            po_count   = len(parsed.get("program_outcomes", []))
            peo_count  = len(parsed.get("peos", []))
            pso_count  = len(parsed.get("psos", []))
            unit_count = len(parsed.get("units", []))
            lo_count  = len(parsed.get("learning_outcomes", []))
            logger.info(
                f"Extraction complete: {co_count} COs, {po_count} POs, {peo_count} PEOs, "
                f"{pso_count} PSOs, {lo_count} LOs, {unit_count} units, course='{parsed.get('course_name')}'"
            )
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
