# backend/services/evaluation_plan_service.py
"""
EvaluationPlanService
----------------------
Generates a full CIE + SEE evaluation plan using the LLM fallback chain
(Gemini → Groq → OpenAI), then writes a formatted Word document.

Output folder : generated_docs/evaluation_plans/
Filename      : evaluation_plan_<course_id>.docx
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

OUTPUT_DIR = Path("generated_docs/evaluation_plans")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


class EvaluationPlanService:

    def __init__(self, db: AsyncSession):
        self.db = db

    # ------------------------------------------------------------------
    # Static helper — used by route for file-existence check
    # ------------------------------------------------------------------

    @staticmethod
    def get_filepath(course_id: int) -> str:
        return str(OUTPUT_DIR / f"evaluation_plan_{course_id}.docx")

    # ------------------------------------------------------------------
    # Main entry point
    # ------------------------------------------------------------------

    async def generate(self, course_id: int) -> dict:
        # 1. Load course (raises OBEException 404 if missing)
        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)

        course_name = course.course_name
        course_code = course.course_code
        cos = course.cos            # list of {co_id, statement, bloom_level}
        eval_cfg = course.evaluation_config  # {continuous_assessment_total, components, end_sem_total}

        logger.info(
            f"Generating evaluation plan for '{course_name}' ({course_code})"
        )

        # 2. Build prompt and call LLM fallback chain
        prompt = self._build_prompt(course_name, course_code, cos, eval_cfg)
        plan = await self._call_llm(prompt)

        # 3. Write Word doc
        filepath = self.get_filepath(course_id)
        self._build_docx(course_name, course_code, cos, eval_cfg, plan, filepath)

        return {
            "course_id": course_id,
            "course_name": course_name,
            "filename": os.path.basename(filepath),
            "download_url": f"/evaluation-plan/download/{course_id}",
            "cie_total": eval_cfg.get("continuous_assessment_total", 30),
            "see_total": eval_cfg.get("end_sem_total", 60),
        }

    # ------------------------------------------------------------------
    # Prompt builder
    # ------------------------------------------------------------------

    def _build_prompt(
        self,
        course_name: str,
        course_code: str,
        cos: list,
        eval_cfg: dict,
    ) -> str:
        cos_text = "\n".join(
            f"  {co['co_id']}: {co['statement']} [Bloom: {co['bloom_level']}]"
            for co in cos
        )
        components = eval_cfg.get("components", {})
        components_text = "\n".join(
            f"  - {name}: {marks} marks" for name, marks in components.items()
        )
        cie_total = eval_cfg.get("continuous_assessment_total", 30)
        see_total = eval_cfg.get("end_sem_total", 60)
        co_ids = [co["co_id"] for co in cos]

        return f"""You are an expert in Outcome-Based Education (OBE) evaluation design for engineering colleges.

Course : {course_name} ({course_code})

Course Outcomes:
{cos_text}

Evaluation Structure:
  CIE (Continuous Internal Evaluation) — Total: {cie_total} marks
  Components:
{components_text}

  SEE (Semester End Exam) — Total: {see_total} marks

Task: Generate a detailed EVALUATION PLAN.

Rules:
- For each CIE component, specify: which COs it assesses, Bloom's level targeted, and suggested question types.
- For SEE, provide a mark distribution across COs and a recommended question paper pattern.
- CO attainment targets: suggest a threshold percentage (e.g. 60%) for each CO.
- Return ONLY valid JSON — no markdown, no extra text.

Schema:
{{
  "cie_plan": [
    {{
      "component": "string",
      "marks": 10,
      "cos_assessed": ["CO1", "CO2"],
      "bloom_levels": ["Apply", "Analyze"],
      "question_types": ["Short Answer", "Problem Solving"],
      "description": "string describing what to assess"
    }}
  ],
  "see_plan": {{
    "total_marks": {see_total},
    "duration_hours": 3,
    "pattern": "string (e.g. 5 units x 2 questions, attempt any 1)",
    "co_mark_distribution": {{
      "CO1": 20,
      "CO2": 20
    }},
    "bloom_distribution": {{
      "Remember": 10,
      "Understand": 20,
      "Apply": 30
    }}
  }},
  "co_attainment_targets": [
    {{
      "co_id": "CO1",
      "threshold_percentage": 60,
      "measurement_method": "string"
    }}
  ]
}}"""

    # ------------------------------------------------------------------
    # LLM call with fallback
    # ------------------------------------------------------------------

    async def _call_llm(self, prompt: str) -> dict:
        logger.info("Calling LLM for evaluation plan")
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
            raise LLMError("LLM returned invalid JSON for evaluation plan")

    # ------------------------------------------------------------------
    # Word document builder
    # ------------------------------------------------------------------

    def _build_docx(
        self,
        course_name: str,
        course_code: str,
        cos: list,
        eval_cfg: dict,
        data: dict,
        filepath: str,
    ):
        doc = Document()

        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # ── Title block
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("EVALUATION PLAN")
        run.bold = True
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        srun = subtitle.add_run(f"{course_name}   |   {course_code}")
        srun.bold = True
        srun.font.size = Pt(12)

        doc.add_paragraph()

        # ── CO Reference
        self._section_heading(doc, "Course Outcomes (CO) Reference")
        co_tbl = doc.add_table(rows=1, cols=3)
        co_tbl.style = "Table Grid"
        for cell, text in zip(co_tbl.rows[0].cells, ["CO ID", "Statement", "Bloom's Level"]):
            self._header_cell(cell, text)
        for co in cos:
            row = co_tbl.add_row().cells
            row[0].text = co.get("co_id", "")
            row[1].text = co.get("statement", "")
            row[2].text = co.get("bloom_level", "")
            for c in row:
                c.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

        # ── CIE Plan
        cie_total = eval_cfg.get("continuous_assessment_total", 30)
        self._section_heading(doc, f"Continuous Internal Evaluation (CIE) — {cie_total} Marks")

        cie_tbl = doc.add_table(rows=1, cols=5)
        cie_tbl.style = "Table Grid"
        cie_headers = ["Component", "Marks", "COs Assessed", "Bloom's Levels", "Question Types"]
        for cell, text in zip(cie_tbl.rows[0].cells, cie_headers):
            self._header_cell(cell, text, color="375623")

        for comp in data.get("cie_plan", []):
            row = cie_tbl.add_row().cells
            row[0].text = comp.get("component", "")
            row[1].text = str(comp.get("marks", ""))
            row[2].text = ", ".join(comp.get("cos_assessed", []))
            row[3].text = ", ".join(comp.get("bloom_levels", []))
            row[4].text = ", ".join(comp.get("question_types", []))
            for c in row:
                c.paragraphs[0].runs[0].font.size = Pt(9)

            desc = comp.get("description", "")
            if desc:
                note = doc.add_paragraph()
                note_run = note.add_run(f"  ↳ {comp.get('component','')}: {desc}")
                note_run.italic = True
                note_run.font.size = Pt(9)

        doc.add_paragraph()

        # ── SEE Plan
        see_data = data.get("see_plan", {})
        see_total = eval_cfg.get("end_sem_total", 60)
        self._section_heading(
            doc,
            f"Semester End Exam (SEE) — {see_total} Marks  |  "
            f"Duration: {see_data.get('duration_hours', 3)} hrs",
        )

        pattern_para = doc.add_paragraph()
        pattern_run = pattern_para.add_run(
            f"Pattern: {see_data.get('pattern', 'N/A')}"
        )
        pattern_run.font.size = Pt(10)

        doc.add_paragraph()

        # CO mark distribution
        self._section_heading(doc, "SEE — CO-wise Mark Distribution", sub=True)
        co_dist = see_data.get("co_mark_distribution", {})
        if co_dist:
            see_co_tbl = doc.add_table(rows=1, cols=2)
            see_co_tbl.style = "Table Grid"
            for cell, text in zip(see_co_tbl.rows[0].cells, ["CO", "Marks Allocated"]):
                self._header_cell(cell, text, color="7F3F98")
            for co_id, marks in co_dist.items():
                row = see_co_tbl.add_row().cells
                row[0].text = co_id
                row[1].text = str(marks)
                for c in row:
                    c.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

        bloom_dist = see_data.get("bloom_distribution", {})
        if bloom_dist:
            self._section_heading(doc, "SEE — Bloom's Level Distribution", sub=True)
            bloom_tbl = doc.add_table(rows=1, cols=2)
            bloom_tbl.style = "Table Grid"
            for cell, text in zip(bloom_tbl.rows[0].cells, ["Bloom's Level", "Marks"]):
                self._header_cell(cell, text, color="7F3F98")
            for level, marks in bloom_dist.items():
                row = bloom_tbl.add_row().cells
                row[0].text = level
                row[1].text = str(marks)
                for c in row:
                    c.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

        # ── CO Attainment Targets
        self._section_heading(doc, "CO Attainment Targets")
        att_tbl = doc.add_table(rows=1, cols=3)
        att_tbl.style = "Table Grid"
        for cell, text in zip(
            att_tbl.rows[0].cells, ["CO ID", "Threshold %", "Measurement Method"]
        ):
            self._header_cell(cell, text, color="C55A11")

        for target in data.get("co_attainment_targets", []):
            row = att_tbl.add_row().cells
            row[0].text = target.get("co_id", "")
            row[1].text = f"{target.get('threshold_percentage', 60)}%"
            row[2].text = target.get("measurement_method", "")
            for c in row:
                c.paragraphs[0].runs[0].font.size = Pt(9)

        doc.save(filepath)
        logger.info(f"Evaluation plan saved → {filepath}")

    # ------------------------------------------------------------------
    # Small helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _section_heading(doc: Document, text: str, sub: bool = False):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10 if sub else 11)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)

    @staticmethod
    def _header_cell(cell, text: str, color: str = "1F497D"):
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