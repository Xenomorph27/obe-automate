# backend/services/session_plan_service.py
"""
SessionPlanService
------------------
Generates a unit-wise session plan using the LLM fallback chain
(Gemini → Groq → OpenAI), then writes a formatted Word document.

Output folder : generated_docs/session_plans/
Filename      : session_plan_<course_id>.docx
"""

import json
import os
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.exceptions import LLMError, OBEException
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger
from backend.services.course_service import CourseService

logger = get_logger(__name__)

from backend.core.storage import get_storage
_CATEGORY = "session_plans"


class SessionPlanService:

    def __init__(self, db: AsyncSession):
        self.db = db

    # ------------------------------------------------------------------
    # Static helper — used by route for file-existence check
    # ------------------------------------------------------------------

    @staticmethod
    def get_filepath(course_id: int) -> str:
        storage = get_storage()
        p = storage.get_path(_CATEGORY, f"session_plan_{course_id}.docx")
        return str(p) if p else str(get_storage()._dir(_CATEGORY) / f"session_plan_{course_id}.docx")

    # ------------------------------------------------------------------
    # Main entry point
    # ------------------------------------------------------------------

    async def generate(self, course_id: int) -> dict:
        # 1. Load course (raises OBEException 404 if missing)
        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)

        course_name = course.course_name
        course_code = course.course_code
        cos = course.cos  # list of dicts: co_id, statement, bloom_level

        logger.info(
            f"Generating session plan for '{course_name}' ({course_code}) "
            f"with {len(cos)} COs"
        )

        # 2. Build prompt and call LLM fallback chain
        prompt = self._build_prompt(course_name, course_code, cos, course.total_hours)
        plan = await self._call_llm(prompt)

        # 3. Write Word doc via storage abstraction
        _storage = get_storage()
        _filename = f"session_plan_{course_id}.docx"
        filepath = self._build_docx(course_name, course_code, cos, plan, _storage, _filename)

        total_sessions = sum(len(u.get("sessions", [])) for u in plan.get("units", []))

        return {
            "course_id": course_id,
            "course_name": course_name,
            "filename": os.path.basename(filepath),
            "download_url": f"/session-plan/download/{course_id}",
            "total_sessions": total_sessions,
            "units": plan.get("units", []),
        }

    # ------------------------------------------------------------------
    # Prompt builder
    # ------------------------------------------------------------------

    def _build_prompt(
        self,
        course_name: str,
        course_code: str,
        cos: list,
        total_hours: int,
    ) -> str:
        cos_text = "\n".join(
            f"  {co['co_id']}: {co['statement']} [Bloom: {co['bloom_level']}]"
            for co in cos
        )
        co_ids = [co["co_id"] for co in cos]

        return f"""You are an expert curriculum designer for engineering colleges using OBE.

Course : {course_name} ({course_code})
Total contact hours : {total_hours}

Course Outcomes:
{cos_text}

Task: Generate a complete SESSION PLAN that covers the full syllabus.

Rules:
- Split the hours across 4-5 units logically based on the course outcomes.
- Each unit gets roughly {total_hours // 5} sessions of 50 minutes each.
- Every session must map to at least one of these COs: {co_ids}
- teaching_method must be one of: Lecture, Tutorial, Lab, Flipped Classroom, Case Study, Problem Solving, Group Discussion.
- Return ONLY valid JSON — no markdown, no explanation.

Schema:
{{
  "units": [
    {{
      "unit_number": 1,
      "unit_title": "string",
      "sessions": [
        {{
          "session_number": 1,
          "topic": "string",
          "teaching_method": "Lecture",
          "co_mapped": ["CO1"],
          "bloom_level": "Remember",
          "duration_minutes": 50
        }}
      ]
    }}
  ]
}}"""

    # ------------------------------------------------------------------
    # LLM call with fallback
    # ------------------------------------------------------------------

    async def _call_llm(self, prompt: str) -> dict:
        logger.info("Calling LLM for session plan")
        raw = await get_llm_response(prompt)

        # Strip markdown fences — Critical Pattern (never break)
        if raw.startswith("```"):
            lines = raw.split("\n")
            raw = "\n".join(lines[1:-1])
        raw = raw.strip()

        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            logger.error(f"LLM JSON parse error: {e} | raw[:300]={raw[:300]}")
            raise LLMError("LLM returned invalid JSON for session plan")

    # ------------------------------------------------------------------
    # Word document builder
    # ------------------------------------------------------------------

    def _build_docx(
        self,
        course_name: str,
        course_code: str,
        cos: list,
        data: dict,
        _storage,
        _filename: str,
    ) -> str:
        doc = Document()

        # Page margins — 1 inch all sides
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # ── Title block
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("SESSION PLAN")
        run.bold = True
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        srun = subtitle.add_run(f"{course_name}   |   {course_code}")
        srun.bold = True
        srun.font.size = Pt(12)

        doc.add_paragraph()

        # ── CO Reference Table
        self._section_heading(doc, "Course Outcomes (CO) Reference")
        co_tbl = doc.add_table(rows=1, cols=3)
        co_tbl.style = "Table Grid"
        hdrs = ["CO ID", "Statement", "Bloom's Level"]
        for cell, text in zip(co_tbl.rows[0].cells, hdrs):
            self._header_cell(cell, text)

        for co in cos:
            row = co_tbl.add_row().cells
            row[0].text = co.get("co_id", "")
            row[1].text = co.get("statement", "")
            row[2].text = co.get("bloom_level", "")
            for c in row:
                c.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

        # ── Per-unit session tables
        for unit in data.get("units", []):
            self._section_heading(
                doc,
                f"Unit {unit.get('unit_number', '')}: {unit.get('unit_title', '')}",
            )

            tbl = doc.add_table(rows=1, cols=6)
            tbl.style = "Table Grid"
            col_headers = ["#", "Topic", "Method", "CO Mapped", "Bloom's Level", "Duration"]
            for cell, text in zip(tbl.rows[0].cells, col_headers):
                self._header_cell(cell, text, color="2E75B6")

            for s in unit.get("sessions", []):
                row = tbl.add_row().cells
                row[0].text = str(s.get("session_number", ""))
                row[1].text = s.get("topic", "")
                row[2].text = s.get("teaching_method", "")
                row[3].text = ", ".join(s.get("co_mapped", []))
                row[4].text = s.get("bloom_level", "")
                row[5].text = f"{s.get('duration_minutes', 50)} min"
                for c in row:
                    c.paragraphs[0].runs[0].font.size = Pt(9)

            doc.add_paragraph()

        import tempfile as _tmp
        with _tmp.TemporaryDirectory() as _t:
            _p = Path(_t) / _filename
            doc.save(str(_p))
            _storage.save_from_path(_CATEGORY, _filename, _p)
        filepath = str(_storage.get_path(_CATEGORY, _filename))
        logger.info(f"Session plan saved -> {filepath}")
        return filepath

    # ------------------------------------------------------------------
    # Small helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _section_heading(doc: Document, text: str):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)

    @staticmethod
    def _header_cell(cell, text: str, color: str = "1F497D"):
        """Bold white text on coloured background."""
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)