# backend/services/attainment_service.py
"""
AttainmentService
-----------------
1. save_marks()   — bulk-upsert student marks for a course
2. calculate()    — compute CO attainment % and PO attainment, return dict
3. generate_report() — write a formatted Word document (.docx)

Attainment formula (standard OBE practice):
  CO attainment % = (students scoring >= threshold%) / total_students * 100
  Default threshold = 60% of max marks for that CO's components.

Output folder: generated_docs/attainment_reports/
"""

import json
import os
from pathlib import Path
from collections import defaultdict

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, delete

from backend.core.config import GEMINI_API_KEY
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.database.models import Course, COAttainment
from backend.services.course_service import CourseService

logger = get_logger(__name__)

OUTPUT_DIR = Path("generated_docs/attainment_reports")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ATTAINMENT_THRESHOLD_PCT = 60   # students must score >= 60% of max CO marks


class AttainmentService:

    def __init__(self, db: AsyncSession):
        self.db = db

    # ------------------------------------------------------------------
    # Static helper
    # ------------------------------------------------------------------

    @staticmethod
    def get_filepath(course_id: int) -> str:
        return str(OUTPUT_DIR / f"attainment_report_{course_id}.docx")

    # ------------------------------------------------------------------
    # 1. Save / replace student marks
    # ------------------------------------------------------------------

    async def save_marks(self, course_id: int, students: list) -> dict:
        """
        Accepts a list of student mark dicts, replaces all existing records
        for this course, and commits.

        students = [
          {
            "student_id": "USN001",
            "student_name": "Alice",
            "marks": {
              "CO1": {"Quiz": 8, "Unit Test": 18},
              "CO2": {"Quiz": 7, "Unit Test": 15}
            }
          }, ...
        ]
        """
        # Verify course exists
        course_svc = CourseService(self.db)
        await course_svc.get_course(course_id)   # raises 404 if missing

        # Delete all existing records for this course
        await self.db.execute(
            delete(COAttainment).where(COAttainment.course_id == course_id)
        )

        records = []
        for s in students:
            rec = COAttainment(
                course_id=course_id,
                student_id=s["student_id"],
                student_name=s["student_name"],
            )
            rec.marks = s["marks"]
            self.db.add(rec)
            records.append(rec)

        await self.db.commit()
        logger.info(f"Saved marks for {len(records)} students, course_id={course_id}")

        return {
            "course_id": course_id,
            "students_saved": len(records),
            "message": "Marks saved successfully. Call /attainment/report/{course_id} to generate the report.",
        }

    # ------------------------------------------------------------------
    # 2. Calculate attainment
    # ------------------------------------------------------------------

    async def calculate(self, course_id: int) -> dict:
        """
        Returns full attainment calculation dict without writing any file.
        """
        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)

        result = await self.db.execute(
            select(COAttainment).where(COAttainment.course_id == course_id)
        )
        records = result.scalars().all()

        if not records:
            raise OBEException(
                f"No student marks found for course_id={course_id}. Upload marks first.",
                status_code=400,
            )

        cos = course.cos           # [{co_id, statement, bloom_level}]
        eval_cfg = course.evaluation_config  # {components: {name: max_marks}, ...}
        co_po_matrix = course.co_po_matrix
        pos = course.pos

        components = eval_cfg.get("components", {})   # {"Quiz": 10, "Unit Test": 10, ...}

        # Build CO → max marks map
        # We assume each component is mapped equally to all COs unless
        # marks dict says otherwise. Max per CO = sum of all component maxes.
        # (Faculty can structure marks dict per-CO with per-component sub-marks.)
        all_co_ids = [co["co_id"] for co in cos]

        # Per-CO max = sum of component maximums present in any student's marks for that CO
        co_max: dict[str, float] = {}
        for co_id in all_co_ids:
            # Find all component names used for this CO across all students
            component_names = set()
            for rec in records:
                co_marks = rec.marks.get(co_id, {})
                component_names.update(co_marks.keys())
            # Use the eval_cfg max for each component; fallback to 10 if unknown
            co_max[co_id] = sum(components.get(c, 10) for c in component_names) or sum(components.values())

        threshold = ATTAINMENT_THRESHOLD_PCT / 100

        # Per-student total per CO
        co_student_totals: dict[str, list[float]] = defaultdict(list)
        for rec in records:
            for co_id in all_co_ids:
                co_marks = rec.marks.get(co_id, {})
                total = sum(co_marks.values())
                co_student_totals[co_id].append(total)

        # CO attainment %
        co_attainment: dict[str, dict] = {}
        for co_id in all_co_ids:
            max_m = co_max.get(co_id, 1)
            totals = co_student_totals[co_id]
            passed = sum(1 for t in totals if t >= threshold * max_m)
            attainment_pct = round((passed / len(totals)) * 100, 2) if totals else 0.0
            avg_marks = round(sum(totals) / len(totals), 2) if totals else 0.0

            co_attainment[co_id] = {
                "co_id": co_id,
                "statement": next((c["statement"] for c in cos if c["co_id"] == co_id), ""),
                "bloom_level": next((c["bloom_level"] for c in cos if c["co_id"] == co_id), ""),
                "max_marks": max_m,
                "avg_marks_scored": avg_marks,
                "students_passed_threshold": passed,
                "total_students": len(totals),
                "attainment_percentage": attainment_pct,
                "attainment_level": self._attainment_level(attainment_pct),
                "target_met": attainment_pct >= 60,
            }

        # PO attainment (weighted average via CO-PO matrix)
        po_attainment: dict[str, dict] = {}
        for po in pos:
            po_id = po["po_id"]
            weights = []
            for co_id in all_co_ids:
                weight = co_po_matrix.get(co_id, {}).get(po_id, 0)
                if weight > 0:
                    weights.append((co_attainment[co_id]["attainment_percentage"], weight))
            if weights:
                total_w = sum(w for _, w in weights)
                po_pct = round(sum(p * w for p, w in weights) / total_w, 2)
            else:
                po_pct = 0.0
            po_attainment[po_id] = {
                "po_id": po_id,
                "statement": po["statement"],
                "attainment_percentage": po_pct,
                "attainment_level": self._attainment_level(po_pct),
            }

        return {
            "course_id": course_id,
            "course_name": course.course_name,
            "course_code": course.course_code,
            "total_students": len(records),
            "threshold_percentage": ATTAINMENT_THRESHOLD_PCT,
            "co_attainment": co_attainment,
            "po_attainment": po_attainment,
            "overall_co_attainment": round(
                sum(v["attainment_percentage"] for v in co_attainment.values()) / len(co_attainment), 2
            ) if co_attainment else 0.0,
        }

    # ------------------------------------------------------------------
    # 3. Generate Word report
    # ------------------------------------------------------------------

    async def generate_report(self, course_id: int) -> dict:
        data = await self.calculate(course_id)

        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)

        filepath = self.get_filepath(course_id)
        self._build_docx(course, data, filepath)

        return {
            "course_id": course_id,
            "course_name": data["course_name"],
            "filename": os.path.basename(filepath),
            "download_url": f"/attainment/download/{course_id}",
            "overall_co_attainment": data["overall_co_attainment"],
            "total_students": data["total_students"],
        }

    # ------------------------------------------------------------------
    # Word document builder
    # ------------------------------------------------------------------

    def _build_docx(self, course: Course, data: dict, filepath: str):
        doc = Document()

        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # ── Title
        t = doc.add_paragraph()
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = t.add_run("CO & PO ATTAINMENT REPORT")
        r.bold = True
        r.font.size = Pt(18)
        r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

        s = doc.add_paragraph()
        s.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sr = s.add_run(
            f"{data['course_name']}   |   {data['course_code']}   |   "
            f"Students: {data['total_students']}   |   "
            f"Threshold: {data['threshold_percentage']}%"
        )
        sr.font.size = Pt(11)
        sr.bold = True

        doc.add_paragraph()

        # ── CO Attainment Table
        self._heading(doc, "Course Outcome (CO) Attainment")

        co_tbl = doc.add_table(rows=1, cols=7)
        co_tbl.style = "Table Grid"
        for cell, text in zip(
            co_tbl.rows[0].cells,
            ["CO", "Statement", "Bloom's", "Max Marks", "Avg Scored", "Students ≥ Threshold", "Attainment %"],
        ):
            self._hdr_cell(cell, text)

        for co_id, co_data in data["co_attainment"].items():
            row = co_tbl.add_row().cells
            row[0].text = co_id
            row[1].text = co_data["statement"]
            row[2].text = co_data["bloom_level"]
            row[3].text = str(co_data["max_marks"])
            row[4].text = str(co_data["avg_marks_scored"])
            row[5].text = f"{co_data['students_passed_threshold']} / {co_data['total_students']}"
            pct = co_data["attainment_percentage"]
            row[6].text = f"{pct}%  ({co_data['attainment_level']})"
            # Colour attainment cell by level
            fill = "70AD47" if co_data["target_met"] else "FF0000"
            self._shade_cell(row[6], fill)
            row[6].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            for c in row:
                c.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

        # ── Overall CO Attainment callout
        overall_para = doc.add_paragraph()
        overall_run = overall_para.add_run(
            f"Overall CO Attainment: {data['overall_co_attainment']}%"
        )
        overall_run.bold = True
        overall_run.font.size = Pt(11)
        overall_run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

        doc.add_paragraph()

        # ── PO Attainment Table
        self._heading(doc, "Program Outcome (PO) Attainment (via CO-PO Matrix)")

        po_tbl = doc.add_table(rows=1, cols=4)
        po_tbl.style = "Table Grid"
        for cell, text in zip(
            po_tbl.rows[0].cells,
            ["PO", "Statement", "Attainment %", "Level"],
        ):
            self._hdr_cell(cell, text, color="7F3F98")

        for po_id, po_data in data["po_attainment"].items():
            row = po_tbl.add_row().cells
            row[0].text = po_id
            row[1].text = po_data["statement"]
            row[2].text = f"{po_data['attainment_percentage']}%"
            row[3].text = po_data["attainment_level"]
            for c in row:
                c.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

        # ── CO-PO Matrix section
        self._heading(doc, "CO-PO Correlation Matrix (Reference)")
        pos_list = course.pos
        po_ids = [p["po_id"] for p in pos_list]
        co_ids = list(data["co_attainment"].keys())

        matrix_tbl = doc.add_table(rows=1, cols=len(po_ids) + 1)
        matrix_tbl.style = "Table Grid"
        header_cells = matrix_tbl.rows[0].cells
        header_cells[0].text = "CO \\ PO"
        header_cells[0].paragraphs[0].runs[0].bold = True
        for i, po_id in enumerate(po_ids):
            self._hdr_cell(header_cells[i + 1], po_id, color="C55A11")

        for co_id in co_ids:
            row = matrix_tbl.add_row().cells
            r0 = row[0].paragraphs[0].add_run(co_id)
            r0.bold = True
            r0.font.size = Pt(9)
            for i, po_id in enumerate(po_ids):
                val = course.co_po_matrix.get(co_id, {}).get(po_id, 0)
                row[i + 1].text = str(val) if val else "-"
                row[i + 1].paragraphs[0].runs[0].font.size = Pt(9)

        doc.save(filepath)
        logger.info(f"Attainment report saved → {filepath}")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _attainment_level(pct: float) -> str:
        if pct >= 70:
            return "High"
        if pct >= 50:
            return "Medium"
        return "Low"

    @staticmethod
    def _heading(doc: Document, text: str):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)

    @staticmethod
    def _hdr_cell(cell, text: str, color: str = "1F497D"):
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        AttainmentService._shade_cell(cell, color)

    @staticmethod
    def _shade_cell(cell, color: str):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)