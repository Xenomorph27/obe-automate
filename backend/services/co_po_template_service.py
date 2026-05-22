# backend/services/co_po_template_service.py
"""
Generates a pre-filled CO-PO Attainment Excel workbook (.xlsx) for a course.

Structure (mirrors the SIT template exactly):
  Course_Info   — auto-filled from DB
  Roll_List     — students from DB
  CO_List       — COs + attainment rubric
  ESE_QP        — question paper template (empty, faculty fills)
  ESE_MKS       — marks sheet with student list + formulas
  CA{n}_QP      — one per CA in eval config (questions pre-filled from QB or blank)
  CA{n}_Marks   — one per CA, student list + formulas
  Final_CO_Attn — formula-driven CO attainment summary
  PO_Attainment — CO-PO matrix from course data + formula-driven PO attainment
"""

import os
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import (Alignment, Border, Font, PatternFill, Side,
                              numbers)
from openpyxl.utils import get_column_letter
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.services.course_service import CourseService

logger = get_logger(__name__)

_CATEGORY = "co_po_templates"

# ── Colours ───────────────────────────────────────────────────────────────────
_NAVY_HEX   = "1F3864"
_LIGHT_HEX  = "D6DCE4"
_YELLOW_HEX = "FFFF00"
_GREEN_HEX  = "E2EFDA"
_ORANGE_HEX = "FCE4D6"
_HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_SUBHDR_FONT = Font(name="Calibri", bold=True, size=10)
_BODY_FONT   = Font(name="Calibri", size=10)
_NAVY_FILL   = PatternFill("solid", fgColor=_NAVY_HEX)
_LIGHT_FILL  = PatternFill("solid", fgColor=_LIGHT_HEX)
_GREEN_FILL  = PatternFill("solid", fgColor=_GREEN_HEX)
_ORANGE_FILL = PatternFill("solid", fgColor=_ORANGE_HEX)

_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


def _hdr(ws, row, col, value, fill=None, font=None, align=None, number_format=None):
    """Write a styled cell."""
    c = ws.cell(row=row, column=col, value=value)
    c.font   = font  or _SUBHDR_FONT
    c.fill   = fill  or PatternFill()
    c.border = _BORDER
    c.alignment = align or _CENTER
    if number_format:
        c.number_format = number_format
    return c


def _navy(ws, row, col, value, align=None):
    return _hdr(ws, row, col, value, fill=_NAVY_FILL, font=_HEADER_FONT, align=align or _CENTER)


def _body(ws, row, col, value, align=None, fill=None, bold=False, number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font   = Font(name="Calibri", size=10, bold=bold)
    c.border = _BORDER
    c.alignment = align or _LEFT
    if fill:
        c.fill = fill
    if number_format:
        c.number_format = number_format
    return c


def _merge_hdr(ws, row, c1, c2, value, fill=None, font=None):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1, value=value)
    c.font   = font  or _SUBHDR_FONT
    c.fill   = fill  or PatternFill()
    c.border = _BORDER
    c.alignment = _CENTER
    return c


def _course_header_block(ws, course, start_row=1):
    """Write the standard 4-row course info block used on most sheets."""
    r = start_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1, value=f"Department of : {course.department}")
    c.font = _SUBHDR_FONT; c.alignment = _LEFT

    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=12)
    c2 = ws.cell(row=r, column=7, value="CO Attainment")
    c2.font = _SUBHDR_FONT; c2.alignment = _CENTER

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = ws.cell(row=r, column=1, value=course.course_name)
    c.font = Font(name="Calibri", bold=True, size=12); c.alignment = _CENTER

    r += 1
    pairs = [
        ("Academic Year", course.academic_year),
        ("Batch",         course.academic_year),
        ("Semester",      course.semester),
    ]
    for i, (label, val) in enumerate(pairs):
        col = i * 4 + 1
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        ws.cell(row=r, column=col, value=label).font = _SUBHDR_FONT
        ws.merge_cells(start_row=r, start_column=col+2, end_row=r, end_column=col+3)
        ws.cell(row=r, column=col+2, value=val).font = _BODY_FONT

    r += 1
    meta = [
        ("Course Name", course.course_name),
        ("Course Code", course.course_code),
        ("Faculty Name", course.faculty_name),
    ]
    for i, (label, val) in enumerate(meta):
        col = i * 4 + 1
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        ws.cell(row=r, column=col, value=label).font = _SUBHDR_FONT
        ws.merge_cells(start_row=r, start_column=col+2, end_row=r, end_column=col+3)
        ws.cell(row=r, column=col+2, value=val).font = _BODY_FONT

    return r + 1  # next free row


# ── Sheet builders ────────────────────────────────────────────────────────────

def _build_course_info(wb, course):
    ws = wb.create_sheet("Course_Info")
    info = [
        ("Department of :", f"Department of : {course.department}"),
        ("CO Attainment", "CO Attainment"),
        ("Academic Year", course.academic_year),
        ("Batch",         course.academic_year),
        ("Examination Season", ""),
        ("Course Name",   course.course_name),
        ("Course Code",   course.course_code),
        ("Semester",      course.semester),
        ("Credit",        course.credits),
        ("Faculty Name",  course.faculty_name),
    ]
    for i, (label, val) in enumerate(info, 1):
        ws.cell(row=i, column=1, value=label).font = _SUBHDR_FONT
        ws.cell(row=i, column=3, value=val).font = _BODY_FONT
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["C"].width = 45


def _build_roll_list(wb, course, students):
    ws = wb.create_sheet("Roll_List")
    # Header block
    r = _course_header_block(ws, course)
    # Column headers
    headers = ["Sr. No.", "Seat No", "PRN", "Name of the Student", "Section"]
    for ci, h in enumerate(headers, 1):
        _navy(ws, r, ci, h)
    r += 1
    for idx, s in enumerate(students, 1):
        fill = _LIGHT_FILL if idx % 2 == 0 else PatternFill()
        _body(ws, r, 1, idx,          align=_CENTER, fill=fill)
        _body(ws, r, 2, "",           align=_CENTER, fill=fill)
        _body(ws, r, 3, s["prn"],     align=_CENTER, fill=fill)
        _body(ws, r, 4, s["name"],    align=_LEFT,   fill=fill)
        _body(ws, r, 5, s["section"], align=_CENTER, fill=fill)
        r += 1
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 10


def _build_co_list(wb, course):
    ws = wb.create_sheet("CO_List")
    r = _course_header_block(ws, course)

    # Rubric block
    _merge_hdr(ws, r, 1, 8, "Rubric for deciding level of attainment", fill=_ORANGE_FILL)
    _navy(ws, r, 9, "Range")
    _navy(ws, r, 11, "Level")
    r += 1
    rubric = [
        ("If the percentage of students is less than equal to 40% secured >= 60% marks", "<= 40%",         1),
        ("If the percentage of students is > 40% and < 70% secured >= 60% marks",        "> 40% & < 70%",  2),
        ("If the percentage of students >= 70% secured >= 60% marks",                    ">=70%",          3),
    ]
    for text_val, rng, lvl in rubric:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.cell(row=r, column=1, value=text_val).font = _BODY_FONT
        _body(ws, r, 9, rng, align=_CENTER)
        _body(ws, r, 11, lvl, align=_CENTER)
        r += 1

    # CO table headers
    _navy(ws, r, 1, "CO No")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    _navy(ws, r, 2, "Statement")
    _navy(ws, r, 9, "Target (% of max marks)")
    r += 1

    for co in course.cos:
        _body(ws, r, 1, co["co_id"], align=_CENTER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        _body(ws, r, 2, co["statement"])
        _body(ws, r, 9, 60, align=_CENTER)
        r += 1

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 60


def _build_qp_sheet(wb, sheet_name, course, ca_label, questions=None):
    """Build a question paper sheet (ESE_QP or CA{n}_QP)."""
    ws = wb.create_sheet(sheet_name)
    r = _course_header_block(ws, course)

    # Column headers
    qp_cols = ["Q. No", "Question", "", "", "", "", "Marks",
               "CO Map to question", "BL", "",
               "CO", "Marks for the CO", "Percentage", "",
               "Bloom's Taxonomy Level (BL)", "", "Marks for BL", "Percentage"]
    for ci, h in enumerate(qp_cols, 1):
        _navy(ws, r, ci, h)
    r += 1

    cos = [c["co_id"] for c in course.cos]
    bloom_levels = ["L1", "L2", "L3", "L4", "L5", "L6"]
    bloom_labels = ["Remembering", "Understanding", "Applying", "Analyzing", "Evaluating", "Creating"]

    # Summary formula rows (CO summary columns K-N, BL summary columns O-R)
    for i, co in enumerate(cos):
        row_offset = r + i
        _body(ws, row_offset, 11, co, align=_CENTER)
        # Marks for this CO (SUMIF over col H)
        _body(ws, row_offset, 12,
              f"=IF(SUMIF($H${r}:$H${r+39},{get_column_letter(11)}{row_offset},$G${r}:$G${r+39})=0,"
              f"\"\",SUMIF($H${r}:$H${r+39},{get_column_letter(11)}{row_offset},$G${r}:$G${r+39}))",
              align=_CENTER)
        _body(ws, row_offset, 13,
              f"=IFERROR({get_column_letter(12)}{row_offset}/SUM($L${r}:$L${r+5})*100,\"\")",
              align=_CENTER, number_format="0.0%")

    for i, (bl, lbl) in enumerate(zip(bloom_levels, bloom_labels)):
        row_offset = r + i
        _body(ws, row_offset, 15, bl,  align=_CENTER)
        _body(ws, row_offset, 16, lbl, align=_LEFT)
        _body(ws, row_offset, 17,
              f"=IF(SUMIF($I${r}:$I${r+39},{get_column_letter(15)}{row_offset},$G${r}:$G${r+39})=0,"
              f"\"\",SUMIF($I${r}:$I${r+39},{get_column_letter(15)}{row_offset},$G${r}:$G${r+39}))",
              align=_CENTER)
        _body(ws, row_offset, 18,
              f"=IFERROR({get_column_letter(17)}{row_offset}/SUM($Q${r}:$Q${r+5})*100,\"\")",
              align=_CENTER, number_format="0.0%")

    # Pre-fill questions if provided, else leave blank rows
    if questions:
        for q in questions:
            _body(ws, r, 1, q.get("q_no", ""), align=_CENTER)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            _body(ws, r, 2, q.get("question_text", ""), align=_LEFT)
            _body(ws, r, 7, q.get("marks", ""), align=_CENTER)
            _body(ws, r, 8, q.get("co_id", ""), align=_CENTER)
            _body(ws, r, 9, f"L{q.get('bloom_level', '')}", align=_CENTER)
            r += 1
    else:
        # 15 blank rows for faculty to fill
        for i in range(1, 16):
            _body(ws, r, 1, i, align=_CENTER)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            _body(ws, r, 2, "", align=_LEFT)
            _body(ws, r, 7, "", align=_CENTER)
            _body(ws, r, 8, "", align=_CENTER)
            _body(ws, r, 9, "", align=_CENTER)
            r += 1

    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["H"].width = 14
    return ws


def _build_marks_sheet(wb, sheet_name, course, ca_label, qp_sheet_name, students, total_marks):
    """Build a marks entry sheet (CA{n}_Marks or ESE_MKS)."""
    ws = wb.create_sheet(sheet_name)
    r = _course_header_block(ws, course)

    # Fixed col headers
    fixed = ["Sr. No.", "Seat No", "Roll No.", "Name of the Student", ca_label]
    for ci, h in enumerate(fixed, 1):
        _navy(ws, r, ci, h)

    # Dynamic question columns (Q1–Q15 placeholders)
    for qi in range(1, 16):
        _navy(ws, r, 5 + qi, f"Q{qi}")
    r += 1

    # Marks row
    _body(ws, r, 5, "Marks", align=_CENTER, fill=_LIGHT_FILL)
    for qi in range(1, 16):
        _body(ws, r, 5 + qi, "", align=_CENTER, fill=_LIGHT_FILL)
    r += 1

    # Max marks row
    _body(ws, r, 5, total_marks, align=_CENTER, fill=_GREEN_FILL, bold=True)
    for qi in range(1, 16):
        _body(ws, r, 5 + qi, "", align=_CENTER, fill=_GREEN_FILL)
    data_start = r + 1
    r += 1

    # Student rows
    for idx, s in enumerate(students, 1):
        fill = _LIGHT_FILL if idx % 2 == 0 else PatternFill()
        _body(ws, r, 1, idx, align=_CENTER, fill=fill)
        _body(ws, r, 2, "",  align=_CENTER, fill=fill)
        _body(ws, r, 3, s["prn"],  align=_CENTER, fill=fill)
        _body(ws, r, 4, s["name"], align=_LEFT,   fill=fill)
        # Total formula
        _body(ws, r, 5,
              f"=SUM({get_column_letter(6)}{r}:{get_column_letter(20)}{r})",
              align=_CENTER, fill=fill)
        # Blank mark cells
        for qi in range(1, 16):
            _body(ws, r, 5 + qi, None, align=_CENTER, fill=fill)
        r += 1

    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["C"].width = 14
    return ws


def _build_final_co_attn(wb, course, ca_names, marks_sheet_names, ese_sheet_name):
    """Final CO attainment sheet — weighted average of CA + ESE."""
    ws = wb.create_sheet("Final_CO_Attn")
    r = _course_header_block(ws, course)

    cos = [c["co_id"] for c in course.cos]

    # Header
    _merge_hdr(ws, r, 1, 1, "CO No / Weightage", fill=_NAVY_FILL, font=_HEADER_FONT)
    _merge_hdr(ws, r, 2, len(ca_names)+1, "CO Attainment using CIE", fill=_NAVY_FILL, font=_HEADER_FONT)
    _merge_hdr(ws, r, len(ca_names)+2, len(ca_names)+4, "Final CO attainment", fill=_NAVY_FILL, font=_HEADER_FONT)
    _navy(ws, r, len(ca_names)+5, "Overall Att")
    r += 1

    # Sub headers
    _body(ws, r, 1, "", align=_CENTER)
    for i, ca in enumerate(ca_names, 2):
        _body(ws, r, i, ca, align=_CENTER, bold=True)
    offset = len(ca_names) + 2
    _body(ws, r, offset,   "Internal", align=_CENTER, bold=True)
    _body(ws, r, offset+1, "External", align=_CENTER, bold=True)
    _body(ws, r, offset+2, "Final",    align=_CENTER, bold=True)
    r += 1

    # Weights row
    _body(ws, r, 1, "", align=_CENTER)
    for i in range(len(ca_names)):
        _body(ws, r, i+2, "", align=_CENTER)
    _body(ws, r, offset,   40, align=_CENTER, fill=_GREEN_FILL)
    _body(ws, r, offset+1, 60, align=_CENTER, fill=_GREEN_FILL)
    _body(ws, r, offset+2, 100, align=_CENTER, fill=_GREEN_FILL)
    int_w_row = r
    r += 1

    data_start = r
    for co in cos:
        _body(ws, r, 1, co, align=_CENTER, bold=True)
        # Placeholder — faculty fills CA attainment or links from marks sheets
        for i in range(len(ca_names)):
            _body(ws, r, i+2, "", align=_CENTER)
        # Internal average
        ca_cols = [get_column_letter(i+2) for i in range(len(ca_names))]
        ca_range = ":".join([f"{c}{r}" for c in ca_cols])
        _body(ws, r, offset,
              f"=IFERROR(AVERAGE({','.join([f'{c}{r}' for c in ca_cols])}),''))",
              align=_CENTER)
        # External placeholder
        _body(ws, r, offset+1, "", align=_CENTER)
        # Final weighted
        _body(ws, r, offset+2,
              f"=IFERROR(({get_column_letter(offset)}{r}*${get_column_letter(offset)}${int_w_row}"
              f"+{get_column_letter(offset+1)}{r}*${get_column_letter(offset+1)}${int_w_row})/100,'')",
              align=_CENTER)
        r += 1

    # Overall attainment
    _body(ws, r, 1, "Overall CO Attainment", bold=True)
    final_col = get_column_letter(offset+2)
    _body(ws, r, offset+5,
          f"=IFERROR(AVERAGE({final_col}{data_start}:{final_col}{r-1}),'')",
          align=_CENTER, bold=True, fill=_GREEN_FILL)

    ws.column_dimensions["A"].width = 20
    for i in range(len(ca_names)):
        ws.column_dimensions[get_column_letter(i+2)].width = 12


def _build_po_attainment(wb, course):
    """PO attainment sheet — uses CO-PO matrix from course data."""
    ws = wb.create_sheet("PO_Attainment")
    r = _course_header_block(ws, course)

    co_po = course.co_po_matrix  # dict: {"CO1": {"PO1": 3, "PO2": 2, ...}, ...}
    cos   = [c["co_id"] for c in course.cos]
    pos   = course.pos           # list of {"po_id": "PO1", "description": "..."}
    po_ids = [p["po_id"] for p in pos] if pos else [f"PO{i}" for i in range(1, 13)]

    # Headers
    _navy(ws, r, 1, "CO")
    _navy(ws, r, 2, "Attainment")
    for ci, po in enumerate(po_ids, 3):
        _navy(ws, r, ci, po)
    r += 1

    data_start = r
    for co in cos:
        _body(ws, r, 1, co, align=_CENTER, bold=True)
        _body(ws, r, 2, "", align=_CENTER)  # filled from Final_CO_Attn
        mapping = co_po.get(co, {})
        for ci, po in enumerate(po_ids, 3):
            val = mapping.get(po, None)
            _body(ws, r, ci, val if val else "", align=_CENTER,
                  fill=_GREEN_FILL if val else PatternFill())
        r += 1

    # Articulation average
    _body(ws, r, 1, "Articulation Average", bold=True)
    _body(ws, r, 2, "")
    for ci, po in enumerate(po_ids, 3):
        col = get_column_letter(ci)
        _body(ws, r, ci,
              f"=IFERROR(AVERAGE({col}{data_start}:{col}{r-1}),\"-\")",
              align=_CENTER, fill=_ORANGE_FILL)
    r += 1

    # CO-PO Attainment (weighted)
    _body(ws, r, 1, "CO-PO Attainment", bold=True)
    _body(ws, r, 2, "")
    attn_col = get_column_letter(2)
    for ci, po in enumerate(po_ids, 3):
        col = get_column_letter(ci)
        # Weighted: sum(CO_attainment * mapping) / (3 * count of mapped COs)
        terms = "+".join(
            [f"{col}{data_start+i}*${attn_col}${data_start+i}"
             for i in range(len(cos))]
        )
        _body(ws, r, ci,
              f"=IFERROR(({terms})/(3*COUNT({col}{data_start}:{col}{data_start+len(cos)-1})),\"-\")",
              align=_CENTER, fill=_LIGHT_FILL)

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 12
    for ci in range(3, 3 + len(po_ids)):
        ws.column_dimensions[get_column_letter(ci)].width = 8


# ── Main service class ────────────────────────────────────────────────────────

class COPOTemplateService:
    def __init__(self, db: AsyncSession):
        self.db = db

    @staticmethod
    def get_filepath(course_id: int) -> str:
        storage = get_storage()
        p = storage.get_path(_CATEGORY, f"co_po_template_{course_id}.xlsx")
        return str(p) if p else str(
            get_storage()._dir(_CATEGORY) / f"co_po_template_{course_id}.xlsx"
        )

    async def _get_students(self, course_id: int):
        try:
            result = await self.db.execute(
                text("SELECT prn, name, section FROM students WHERE course_id=:cid ORDER BY section, name"),
                {"cid": course_id}
            )
            rows = result.fetchall()
            return [{"prn": r[0], "name": r[1], "section": r[2]} for r in rows]
        except Exception as e:
            logger.warning(f"Could not fetch students: {e}")
            return []

    async def _get_questions(self, course_id: int, co_id: Optional[str] = None):
        from sqlalchemy import select
        from backend.database.models import Question
        stmt = select(Question).where(Question.course_id == course_id)
        if co_id:
            stmt = stmt.where(Question.co_id == co_id)
        result = await self.db.execute(stmt)
        return result.scalars().all()

    async def generate(self, course_id: int, qp_source: str = "blank") -> dict:
        """
        Generate the CO-PO attainment Excel workbook.

        qp_source: "blank"     — leave QP sheets empty (faculty fills)
                   "question_bank" — pre-fill from question bank
        """
        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)
        students = await self._get_students(course_id)
        eval_cfg = course.evaluation_config
        components = eval_cfg.get("components", {})

        # Determine how many CAs from the evaluation config
        ca_names = sorted([k for k in components.keys()
                          if k.upper().startswith("CA") or "quiz" in k.lower()
                          or "unit test" in k.lower() or "test" in k.lower()])
        # Fallback: use numeric suffix detection
        if not ca_names:
            ca_names = [f"CA{i}" for i in range(1, 4)]
        # Cap at 5
        ca_names = ca_names[:5]

        wb = Workbook()
        # Remove default sheet
        del wb[wb.sheetnames[0]]

        # 1. Course_Info
        _build_course_info(wb, course)

        # 2. Roll_List
        _build_roll_list(wb, course, students)

        # 3. CO_List
        _build_co_list(wb, course)

        # 4. ESE_QP + ESE_MKS
        ese_questions = None
        if qp_source == "question_bank":
            qs = await self._get_questions(course_id)
            ese_questions = [
                {"q_no": i+1, "question_text": q.question_text,
                 "marks": q.marks, "co_id": q.co_id,
                 "bloom_level": q.bloom_level}
                for i, q in enumerate(qs[:20])
            ]
        _build_qp_sheet(wb, "ESE_QP", course, "ESE", questions=ese_questions)
        ese_total = eval_cfg.get("end_sem_total", 60)
        _build_marks_sheet(wb, "ESE_MKS", course, "ESE", "ESE_QP", students, ese_total)

        # 5. CA sheets
        marks_sheets = []
        for ca in ca_names:
            qp_name  = f"{ca}_QP"
            mks_name = f"{ca}_Marks"
            ca_questions = None
            if qp_source == "question_bank":
                qs = await self._get_questions(course_id)
                ca_questions = [
                    {"q_no": i+1, "question_text": q.question_text,
                     "marks": q.marks, "co_id": q.co_id,
                     "bloom_level": q.bloom_level}
                    for i, q in enumerate(qs[:10])
                ]
            _build_qp_sheet(wb, qp_name, course, ca, questions=ca_questions)
            ca_total = components.get(ca, components.get(ca.lower(), 10))
            if not isinstance(ca_total, (int, float)):
                ca_total = 10
            _build_marks_sheet(wb, mks_name, course, ca, qp_name, students, ca_total)
            marks_sheets.append(mks_name)

        # 6. Final_CO_Attn
        _build_final_co_attn(wb, course, ca_names, marks_sheets, "ESE_MKS")

        # 7. PO_Attainment
        _build_po_attainment(wb, course)

        # Save
        import tempfile
        _storage  = get_storage()
        _filename = f"co_po_template_{course_id}.xlsx"
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp) / _filename
            wb.save(str(tmp_path))
            _storage.save_from_path(_CATEGORY, _filename, tmp_path)

        filepath = str(_storage.get_path(_CATEGORY, _filename))
        logger.info(f"CO-PO template saved -> {filepath}")

        return {
            "course_id":    course_id,
            "course_name":  course.course_name,
            "filename":     _filename,
            "download_url": f"/co-po-template/download/{course_id}",
            "sheets":       wb.sheetnames,
            "students":     len(students),
            "ca_count":     len(ca_names),
        }
