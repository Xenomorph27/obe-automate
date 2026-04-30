# backend/services/session_plan_service.py
"""
Generates session plan matching the exact SIT (Symbiosis Institute of Technology)
branded format:
  Header: Institute | Dept | Session Plan | Course info rows
  Table columns: Lect.No | Unit No. | Points to Cover | Methodology |
                 Faculty Conducting | Lecture/Exp.Learning/Evaluation | CO
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
_CATEGORY = "session_plans"

_NAVY  = "1F3864"
_LIGHT = "D6DCE4"
_WHITE = RGBColor(0xFF,0xFF,0xFF)
_NAVY_R = RGBColor(0x1F,0x38,0x64)

class SessionPlanService:
    def __init__(self, db: AsyncSession):
        self.db = db

    @staticmethod
    def get_filepath(course_id: int) -> str:
        storage = get_storage()
        p = storage.get_path(_CATEGORY, f"session_plan_{course_id}.docx")
        return str(p) if p else str(get_storage()._dir(_CATEGORY) / f"session_plan_{course_id}.docx")

    async def generate(self, course_id: int) -> dict:
        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)
        course_name   = course.course_name
        course_code   = course.course_code
        faculty_name  = course.faculty_name
        department    = course.department
        semester      = course.semester
        academic_year = course.academic_year
        credits       = course.credits
        cos           = course.cos
        logger.info(f"Generating session plan for '{course_name}' ({course_code})")
        prompt = self._build_prompt(course_name, course_code, cos, course.total_hours)
        plan   = await self._call_llm(prompt)
        _storage  = get_storage()
        _filename = f"session_plan_{course_id}.docx"
        filepath  = self._build_docx(
            course_name, course_code, faculty_name, department,
            semester, academic_year, credits, cos, plan, _storage, _filename
        )
        total_sessions = sum(len(u.get("sessions",[])) for u in plan.get("units",[]))
        return {
            "course_id":      course_id,
            "course_name":    course_name,
            "filename":       os.path.basename(filepath),
            "download_url":   f"/session-plan/download/{course_id}",
            "total_sessions": total_sessions,
            "units":          plan.get("units",[]),
        }

    def _build_prompt(self, course_name, course_code, cos, total_hours):
        cos_text = "\n".join(f"  {c['co_id']}: {c['statement']} [Bloom: {c['bloom_level']}]" for c in cos)
        co_ids   = [c["co_id"] for c in cos]
        return f"""You are an expert curriculum designer for engineering colleges using OBE.

Course: {course_name} ({course_code})
Total contact hours: {total_hours}
Course Outcomes:
{cos_text}

Generate a complete SESSION PLAN covering the full syllabus.
Rules:
- Split hours across 4-5 units. Each unit ~{total_hours//5} sessions of 50 min.
- Every session maps to exactly one CO from: {co_ids}
- teaching_method: one of: Classroom Teaching, Tutorial, Flipped Classroom, Case Study, Problem Solving, Group Discussion
- type: one of: Lecture, Exp. Learning, Evaluation
- Include at least 1 "Exp. Learning" row per unit (topic="Experiential Learning", type="Exp. Learning")
- Include at least 1 "Quiz" and 1 "Unit Test" row (type="Evaluation")
- Return ONLY valid JSON, no markdown.

Schema:
{{"units":[{{"unit_number":1,"unit_title":"string","sessions":[{{"session_number":1,"topic":"string","teaching_method":"Classroom Teaching","type":"Lecture","co_mapped":"CO1"}}]}}]}}"""

    async def _call_llm(self, prompt: str) -> dict:
        raw = await get_llm_response(prompt)
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:-1])
        raw = raw.strip()
        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            logger.error(f"LLM JSON parse error: {e}")
            raise LLMError("LLM returned invalid JSON for session plan")

    def _build_docx(self, course_name, course_code, faculty_name, department,
                    semester, academic_year, credits, cos, data, _storage, _filename) -> str:
        doc = Document()
        for sec in doc.sections:
            sec.top_margin = sec.bottom_margin = Inches(0.6)
            sec.left_margin = sec.right_margin = Inches(0.7)

        # ── HEADER TABLE (merged single-column rows) ──────────────────
        def _hdr_row(tbl, text, size=12, bold=True):
            if len(tbl.rows) == 0:
                row = tbl.add_row()
            else:
                row = tbl.add_row()
            c = row.cells[0]
            # merge all cols
            for other in row.cells[1:]:
                c = c.merge(other)
            self._shade(c, _NAVY)
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text); r.bold = bold
            r.font.size = Pt(size); r.font.color.rgb = _WHITE

        hdr = doc.add_table(rows=0, cols=9)
        hdr.style = "Table Grid"
        _hdr_row(hdr, "Symbiosis Institute of Technology, Pune", size=13)
        _hdr_row(hdr, f"Department of {department}", size=11)
        _hdr_row(hdr, "Session Plan", size=12)
        _hdr_row(hdr, f"Name of the Department – {department}", size=10, bold=False)
        _hdr_row(hdr, f"Name of the course – {course_name}", size=10, bold=False)

        # Split row: credit on right
        cr = hdr.add_row()
        lc = cr.cells[0].merge(cr.cells[5])
        rc = cr.cells[6].merge(cr.cells[8])
        self._shade(lc, "FFFFFF"); self._shade(rc, "FFFFFF")
        lc.paragraphs[0].add_run(f"Semester – {semester}").font.size = Pt(9)
        rc.paragraphs[0].add_run(f"Credit – {credits}").font.size = Pt(9)

        br = hdr.add_row()
        bc = br.cells[0].merge(br.cells[8])
        self._shade(bc, "FFFFFF")
        bc.paragraphs[0].add_run(f"Name of the faculty – {faculty_name}").font.size = Pt(9)

        yr = hdr.add_row()
        yc = yr.cells[0].merge(yr.cells[8])
        self._shade(yc, "FFFFFF")
        yc.paragraphs[0].add_run(f"Batch – {academic_year}").font.size = Pt(9)

        doc.add_paragraph()

        # ── SESSION TABLE ─────────────────────────────────────────────
        col_headers = ["Lect.\nNo","Unit\nNo.","Points to Cover","Methodology",
                       "Faculty Conducting","Lecture/Exp.\nLearning/Evaluation","CO"]
        col_widths  = [Inches(0.45),Inches(0.45),Inches(2.7),Inches(1.3),
                       Inches(1.5),Inches(1.05),Inches(0.45)]

        tbl = doc.add_table(rows=1, cols=7)
        tbl.style = "Table Grid"
        for i,(cell,text) in enumerate(zip(tbl.rows[0].cells, col_headers)):
            self._shade(cell, _NAVY)
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text); r.bold=True; r.font.size=Pt(8); r.font.color.rgb=_WHITE
            cell.width = col_widths[i]

        lect_num = 0
        for unit in data.get("units",[]):
            for s in unit.get("sessions",[]):
                lect_num += 1
                row = tbl.add_row().cells
                stype = s.get("type","Lecture")
                is_eval = stype in ("Evaluation","Exp. Learning")

                if lect_num % 2 == 0:
                    for c in row: self._shade(c, _LIGHT)

                values = [
                    str(s.get("session_number", lect_num)),
                    str(unit.get("unit_number","")) if not is_eval else "",
                    s.get("topic",""),
                    s.get("teaching_method","Classroom Teaching"),
                    faculty_name,
                    stype,
                    s.get("co_mapped",""),
                ]
                for i,(c,v) in enumerate(zip(row,values)):
                    p = c.paragraphs[0]
                    p.clear()
                    r = p.add_run(v); r.font.size = Pt(8)
                    if i in (0,1,6): p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if is_eval: r.bold=True; r.font.color.rgb=_NAVY_R
                    c.width = col_widths[i]

        import tempfile as _tmp
        with _tmp.TemporaryDirectory() as _t:
            _p = Path(_t) / _filename
            doc.save(str(_p))
            _storage.save_from_path(_CATEGORY, _filename, _p)
        filepath = str(_storage.get_path(_CATEGORY, _filename))
        logger.info(f"Session plan saved -> {filepath}")
        return filepath

    @staticmethod
    def _shade(cell, hex_color):
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hex_color)
        tcPr.append(shd)
