# backend/services/evaluation_plan_service.py
"""
Generates evaluation plan matching the exact SIT format:
Table: Sr.No | Component | Unit Syllabus | CO | Marks | Weightage | Tentative Date
"""
import json, os
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from sqlalchemy.ext.asyncio import AsyncSession
from backend.core.exceptions import LLMError
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger
from backend.services.course_service import CourseService

logger = get_logger(__name__)
from backend.core.storage import get_storage
_CATEGORY = "evaluation_plans"

_NAVY  = "1F3864"
_LIGHT = "D6DCE4"
_WHITE = RGBColor(0xFF,0xFF,0xFF)

class EvaluationPlanService:
    def __init__(self, db: AsyncSession):
        self.db = db

    @staticmethod
    def get_filepath(course_id: int) -> str:
        storage = get_storage()
        p = storage.get_path(_CATEGORY, f"evaluation_plan_{course_id}.docx")
        return str(p) if p else str(get_storage()._dir(_CATEGORY) / f"evaluation_plan_{course_id}.docx")

    async def generate(self, course_id: int) -> dict:
        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)
        course_name   = course.course_name
        course_code   = course.course_code
        faculty_name  = course.faculty_name
        department    = course.department
        semester      = course.semester
        academic_year = course.academic_year
        cos           = course.cos
        eval_cfg      = course.evaluation_config
        logger.info(f"Generating evaluation plan for '{course_name}' ({course_code})")
        prompt = self._build_prompt(course_name, course_code, cos, eval_cfg)
        plan   = await self._call_llm(prompt)
        _storage  = get_storage()
        _filename = f"evaluation_plan_{course_id}.docx"
        filepath  = self._build_docx(
            course_name, course_code, faculty_name, department,
            semester, academic_year, cos, eval_cfg, plan, _storage, _filename
        )
        return {
            "course_id":       course_id,
            "course_name":     course_name,
            "filename":        os.path.basename(filepath),
            "download_url":    f"/evaluation-plan/download/{course_id}",
            "cie_total":       eval_cfg.get("continuous_assessment_total", 30),
            "see_total":       eval_cfg.get("end_sem_total", 60),
            "evaluation_plan": plan,
        }

    def _build_prompt(self, course_name, course_code, cos, eval_cfg):
        cos_text = "\n".join(f"  {c['co_id']}: {c['statement']}" for c in cos)
        co_ids   = [c["co_id"] for c in cos]
        components = eval_cfg.get("components", {})
        comp_text  = "\n".join(f"  - {k}: {v} marks" for k,v in components.items())
        cie_total  = eval_cfg.get("continuous_assessment_total", 30)
        num_units  = len(cos)

        return f"""You are an OBE evaluation expert for engineering colleges.

Course: {course_name} ({course_code})
Course Outcomes: {co_ids}
{cos_text}

CIE Total: {cie_total} marks
Components:
{comp_text}

Generate an EVALUATION PLAN with exactly these CA components (CA1, CA2, CA3...) mapped to units and COs.
Each component must specify: which units it covers, which COs, marks, weightage %, and a tentative date (month 2026).

Return ONLY valid JSON, no markdown:
{{"ca_components":[{{"sr_no":"CA1","component":"Experiential Learning","unit_syllabus":"Unit 1, Unit 2","co_mapped":"CO1, CO2","marks":10,"weightage":"33%","tentative_date":"February 2026"}}]}}"""

    async def _call_llm(self, prompt: str) -> dict:
        raw = await get_llm_response(prompt)
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:-1])
        raw = raw.strip()
        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            logger.error(f"LLM JSON parse error: {e}")
            raise LLMError("LLM returned invalid JSON for evaluation plan")

    def _build_docx(self, course_name, course_code, faculty_name, department,
                    semester, academic_year, cos, eval_cfg, data, _storage, _filename) -> str:
        doc = Document()
        for sec in doc.sections:
            sec.top_margin = sec.bottom_margin = Inches(0.6)
            sec.left_margin = sec.right_margin = Inches(0.7)

        # ── HEADER ────────────────────────────────────────────────────
        def _hdr_row(tbl, text, size=12, bold=True):
            row = tbl.add_row()
            c = row.cells[0]
            for other in row.cells[1:]: c = c.merge(other)
            self._shade(c, _NAVY)
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text); r.bold=bold; r.font.size=Pt(size); r.font.color.rgb=_WHITE

        hdr = doc.add_table(rows=0, cols=7)
        hdr.style = "Table Grid"
        _hdr_row(hdr, "Symbiosis Institute of Technology, Pune", size=13)
        _hdr_row(hdr, f"Department of {department}", size=11)
        _hdr_row(hdr, "Evaluation Plan", size=12)
        _hdr_row(hdr, f"Course: {course_name} ({course_code})  |  Faculty: {faculty_name}  |  Semester: {semester}  |  AY: {academic_year}", size=9, bold=False)

        doc.add_paragraph()

        # ── EVALUATION TABLE ──────────────────────────────────────────
        col_headers = ["Sr. No.","Component","Unit Syllabus","CO","Marks","Weightage","Tentative Date"]
        col_widths  = [Inches(0.5),Inches(1.5),Inches(1.8),Inches(1.0),Inches(0.55),Inches(0.75),Inches(1.4)]

        tbl = doc.add_table(rows=1, cols=7)
        tbl.style = "Table Grid"
        for i,(cell,text) in enumerate(zip(tbl.rows[0].cells, col_headers)):
            self._shade(cell, _NAVY)
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text); r.bold=True; r.font.size=Pt(9); r.font.color.rgb=_WHITE
            cell.width = col_widths[i]

        components = data.get("ca_components", [])
        for i, comp in enumerate(components):
            row = tbl.add_row().cells
            if i % 2 == 1:
                for c in row: self._shade(c, _LIGHT)
            values = [
                comp.get("sr_no",""),
                comp.get("component",""),
                comp.get("unit_syllabus",""),
                comp.get("co_mapped",""),
                str(comp.get("marks","")),
                comp.get("weightage",""),
                comp.get("tentative_date",""),
            ]
            for j,(c,v) in enumerate(zip(row, values)):
                p = c.paragraphs[0]; p.clear()
                r = p.add_run(v); r.font.size=Pt(9)
                if j in (0,4,5): p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                c.width = col_widths[j]

        # ── TOTAL ROW ─────────────────────────────────────────────────
        total_row = tbl.add_row().cells
        self._shade(total_row[0], _NAVY)
        tc = total_row[0].merge(total_row[3])
        self._shade(tc, _NAVY)
        p = tc.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("Total"); r.bold=True; r.font.size=Pt(9); r.font.color.rgb=_WHITE

        marks_c = total_row[4]
        self._shade(marks_c, _NAVY)
        p2 = marks_c.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(str(eval_cfg.get("continuous_assessment_total",30)))
        r2.bold=True; r2.font.size=Pt(9); r2.font.color.rgb=_WHITE

        doc.add_paragraph()

        # ── CO ATTAINMENT NOTE ────────────────────────────────────────
        note_p = doc.add_paragraph()
        nr = note_p.add_run("CO Attainment Target: 60% of students should score ≥60% marks in each assessment component.")
        nr.italic=True; nr.font.size=Pt(9)

        import tempfile as _tmp
        with _tmp.TemporaryDirectory() as _t:
            _p = Path(_t) / _filename
            doc.save(str(_p))
            _storage.save_from_path(_CATEGORY, _filename, _p)
        filepath = str(_storage.get_path(_CATEGORY, _filename))
        logger.info(f"Evaluation plan saved -> {filepath}")
        return filepath

    @staticmethod
    def _shade(cell, hex_color):
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hex_color)
        tcPr.append(shd)
