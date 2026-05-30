# backend/services/co_po_template_service.py
"""
Generates a pre-filled CO-PO Attainment Excel workbook (.xlsx) for a course.
Matches the SIT template structure exactly.

Sheet layout (fixed rows):
  Course_Info   — labels in col A, values in col C (rows 1-11)
  Roll_List     — header block; students from row 8
  CO_List       — header block; rubric rows 7-10, CO table from row 13
  ESE_QP / CA{n}_QP  — header block; col headers row 7; questions from row 8
  ESE_MKS / CA{n}_Marks — header block; col headers row 7; "Marks" row 8;
                            max row 9; students from row 10; summary rows after last student
  Final_CO_Attn — references CA_Marks summary rows and CO_List
  PO_Attainment — CO-PO matrix + attainment formulas
"""

import os
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.services.course_service import CourseService

logger = get_logger(__name__)

_CATEGORY = "co_po_templates"

# ── Colours ──────────────────────────────────────────────────────────────────
_NAVY_HEX    = "1F3864"
_LIGHT_HEX   = "D6DCE4"
_GREEN_HEX   = "E2EFDA"
_ORANGE_HEX  = "FCE4D6"
_YELLOW_HEX  = "FFFF00"
_SKYBLUE_HEX = "BDD7EE"
_TEAL_HEX    = "00B0F0"
_LIME_HEX    = "92D050"
_PINK_HEX    = "FF99CC"

_HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_SUBHDR_FONT = Font(name="Calibri", bold=True, size=10)
_BODY_FONT   = Font(name="Calibri", size=10)
_NAVY_FILL   = PatternFill("solid", fgColor=_NAVY_HEX)
_LIGHT_FILL  = PatternFill("solid", fgColor=_LIGHT_HEX)
_GREEN_FILL  = PatternFill("solid", fgColor=_GREEN_HEX)
_ORANGE_FILL = PatternFill("solid", fgColor=_ORANGE_HEX)
_YELLOW_FILL = PatternFill("solid", fgColor=_YELLOW_HEX)
_SKYBLUE_FILL= PatternFill("solid", fgColor=_SKYBLUE_HEX)
_TEAL_FILL   = PatternFill("solid", fgColor=_TEAL_HEX)
_LIME_FILL   = PatternFill("solid", fgColor=_LIME_HEX)
_PINK_FILL   = PatternFill("solid", fgColor=_PINK_HEX)

_THICK  = Side(style="medium")
_THIN   = Side(style="thin")
_BORDER = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)
_THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# Maximum number of question columns in marks sheets (matches template)
_MAX_Q = 30


def _c(ws, row, col, value=None, font=None, fill=None, align=None, border=True, bold=False, number_format=None):
    """Write a styled cell with bold black border by default."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font or Font(name="Calibri", size=10, bold=bold)
    if fill:
        cell.fill = fill
    cell.border = _BORDER if border else _THIN_BORDER
    cell.alignment = align or _CENTER
    if number_format:
        cell.number_format = number_format
    return cell


def _navy(ws, row, col, value, align=None):
    return _c(ws, row, col, value, font=_HEADER_FONT, fill=_NAVY_FILL, align=align or _CENTER)


def _merge(ws, r1, r2, c1, c2, value=None, font=None, fill=None, align=None):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=value)
    cell.font = font or _SUBHDR_FONT
    if fill:
        cell.fill = fill
    cell.border = _BORDER
    cell.alignment = align or _CENTER
    return cell


def _qp_sheet_name(name):
    """Return sheet name quoted with single quotes if it contains spaces."""
    if " " in name:
        return f"'{name}'"
    return name


# ─────────────────────────────────────────────────────────────────────────────
# Standard 6-row course header block (used in every sheet)
# ─────────────────────────────────────────────────────────────────────────────
def _course_header_block(ws, course):
    dept_text = f"Department of : {course.department}"
    # Row 1: dept in A (merged A1:J1), "CO Attainment" in C1
    ws.merge_cells("A1:J1")
    c = ws.cell(row=1, column=1, value=dept_text)
    c.font = Font(name="Calibri", bold=True, size=12)
    c.fill = _SKYBLUE_FILL
    c.border = _BORDER
    c.alignment = _CENTER

    # Row 2: "CO Attainment" merged
    ws.merge_cells("A2:J2")
    c = ws.cell(row=2, column=1, value="CO Attainment")
    c.font = Font(name="Calibri", bold=True, size=11)
    c.fill = _TEAL_FILL
    c.border = _BORDER
    c.alignment = _CENTER

    academic_year = course.academic_year
    batch = getattr(course, "batch", course.academic_year)
    exam_season = getattr(course, "exam_season", "")

    # Row 4: Academic Year | Batch | Exam Season
    _c(ws, 4, 1, "Academic Year",      bold=True, fill=_LIGHT_FILL, align=_LEFT)
    _c(ws, 4, 3, academic_year,         fill=_YELLOW_FILL, align=_CENTER)
    _c(ws, 4, 5, "Batch",              bold=True, fill=_LIGHT_FILL, align=_LEFT)
    _c(ws, 4, 6, batch,                fill=_YELLOW_FILL, align=_CENTER)
    _c(ws, 4, 8, "Examination Season", bold=True, fill=_LIGHT_FILL, align=_LEFT)
    _c(ws, 4, 10, exam_season,          fill=_YELLOW_FILL, align=_CENTER)

    # Row 5: Course Name | Course Code
    _c(ws, 5, 1, "Course Name",        bold=True, fill=_LIGHT_FILL, align=_LEFT)
    _c(ws, 5, 3, course.course_name,   fill=_YELLOW_FILL, align=_LEFT)
    _c(ws, 5, 8, "Course Code",        bold=True, fill=_LIGHT_FILL, align=_LEFT)
    _c(ws, 5, 10, course.course_code,  fill=_YELLOW_FILL, align=_CENTER)

    return 7  # first usable content row


# ─────────────────────────────────────────────────────────────────────────────
# Course_Info sheet
# ─────────────────────────────────────────────────────────────────────────────
def _build_course_info(wb, course):
    ws = wb.create_sheet("Course_Info")

    ws.merge_cells("A1:K1")
    c = ws.cell(row=1, column=1, value=f"Department of : {course.department}")
    c.font = Font(name="Calibri", bold=True, size=12)
    c.fill = _SKYBLUE_FILL
    c.border = _BORDER
    c.alignment = _CENTER

    ws.merge_cells("A2:K2")
    c = ws.cell(row=2, column=1, value="CO Attainment")
    c.font = Font(name="Calibri", bold=True, size=11)
    c.fill = _TEAL_FILL
    c.border = _BORDER
    c.alignment = _CENTER

    labels = ["Academic Year", "Batch", "Examination Season", "Course Name",
              "Course Code", "Semester", "Credit", "Faculty Name"]
    vals = [
        course.academic_year,
        getattr(course, "batch", course.academic_year),
        getattr(course, "exam_season", ""),
        course.course_name,
        course.course_code,
        course.semester,
        course.credits,
        course.faculty_name,
    ]
    fills = [_LIGHT_FILL, _LIGHT_FILL, _LIGHT_FILL, _GREEN_FILL,
             _GREEN_FILL, _LIGHT_FILL, _LIGHT_FILL, _ORANGE_FILL]
    for i, (lbl, val, fill) in enumerate(zip(labels, vals, fills)):
        r = i + 4
        _c(ws, r, 1, lbl, bold=True, fill=_LIGHT_FILL, align=_LEFT)
        _c(ws, r, 3, val, fill=fill, align=_LEFT)

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["C"].width = 50


# ─────────────────────────────────────────────────────────────────────────────
# Roll_List
# ─────────────────────────────────────────────────────────────────────────────
def _build_roll_list(wb, course, students):
    ws = wb.create_sheet("Roll_List")
    r = _course_header_block(ws, course)

    headers = ["Sr. No.", "Seat No", "PRN", "Name of the Student", "Section"]
    hdr_fills = [_NAVY_FILL, _NAVY_FILL, _NAVY_FILL, _NAVY_FILL, _NAVY_FILL]
    for ci, (h, hf) in enumerate(zip(headers, hdr_fills), 1):
        _c(ws, r, ci, h, font=_HEADER_FONT, fill=hf)
    r += 1

    row_fills = [_GREEN_FILL, _LIGHT_FILL]
    for idx, s in enumerate(students, 1):
        fill = row_fills[idx % 2]
        prn_val = str(s["prn"]) if s["prn"] is not None else ""
        _c(ws, r, 1, idx,                       fill=fill, align=_CENTER)
        _c(ws, r, 2, "",                         fill=fill, align=_CENTER)
        _c(ws, r, 3, prn_val,                    fill=fill, align=_CENTER)
        _c(ws, r, 4, s["name"],                  fill=fill, align=_LEFT)
        _c(ws, r, 5, s.get("section", ""),       fill=fill, align=_CENTER)
        r += 1

    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 32
    ws.column_dimensions["A"].width = 8


# ─────────────────────────────────────────────────────────────────────────────
# CO_List
# ─────────────────────────────────────────────────────────────────────────────
def _build_co_list(wb, course):
    ws = wb.create_sheet("CO_List")
    _course_header_block(ws, course)

    r = 7
    # Rubric header — merged A7:H7
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = ws.cell(row=r, column=1, value="Rubric for deciding level of attainment")
    c.font = Font(name="Calibri", bold=True, size=10)
    c.fill = _ORANGE_FILL
    c.border = _BORDER
    c.alignment = _CENTER
    _c(ws, r, 9,  "Range", bold=True, fill=_NAVY_FILL, font=_HEADER_FONT)
    _c(ws, r, 10, "",      fill=_NAVY_FILL)
    _c(ws, r, 11, "Level", bold=True, fill=_NAVY_FILL, font=_HEADER_FONT)
    r += 1

    rubric_data = [
        ("If the percentage of students is less than equal to 40% secured >= 60%  marks ", "<= 40%",        1, _LIGHT_FILL),
        ("If the percentage of students is > 40% and  < 70% secured >= 60% marks ",        "> 40% & < 70%", 2, _GREEN_FILL),
        ("If the percentage of students is greater than or equal to  70% secured >= 60%  marks ", ">=70%",  3, _SKYBLUE_FILL),
    ]
    for text_val, rng, lvl, rfill in rubric_data:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(row=r, column=1, value=text_val)
        c.font = _BODY_FONT
        c.fill = rfill
        c.border = _BORDER
        c.alignment = _LEFT
        _c(ws, r, 9,  rng, fill=rfill,        align=_CENTER)
        _c(ws, r, 10, "",  fill=rfill)
        _c(ws, r, 11, lvl, fill=_YELLOW_FILL, align=_CENTER, bold=True)
        r += 1

    # Blank row 11, header row 12
    r = 13
    _c(ws, r, 1,  "CO No",    bold=True, fill=_NAVY_FILL, font=_HEADER_FONT, align=_CENTER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    c = ws.cell(row=r, column=2, value="Statement")
    c.font = _HEADER_FONT
    c.fill = _NAVY_FILL
    c.border = _BORDER
    c.alignment = _LEFT
    _c(ws, r, 10, "Target (% of maximum marks)", bold=True, fill=_NAVY_FILL, font=_HEADER_FONT, align=_CENTER)
    r += 1

    co_fills = [_GREEN_FILL, _LIGHT_FILL, _YELLOW_FILL, _ORANGE_FILL, _PINK_FILL]
    for i, co in enumerate(course.cos):
        fill = co_fills[i % len(co_fills)]
        _c(ws, r, 1, co["co_id"], fill=fill, bold=True, align=_CENTER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        c = ws.cell(row=r, column=2, value=co["statement"])
        c.font = _BODY_FONT
        c.fill = fill
        c.border = _BORDER
        c.alignment = _LEFT
        _c(ws, r, 10, 60, fill=_YELLOW_FILL, bold=True, align=_CENTER)
        r += 1

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 70


# ─────────────────────────────────────────────────────────────────────────────
# QP sheet (ESE_QP or CA{n}_QP)
# ─────────────────────────────────────────────────────────────────────────────
def _build_qp_sheet(wb, sheet_name, course, ca_label, questions=None):
    ws = wb.create_sheet(sheet_name)
    _course_header_block(ws, course)

    r = 7
    hdr_labels = ["Q. No", "Question", "", "", "", "", "Marks",
                  "CO Map to question", "BL", "",
                  "CO", "Marks for the CO", "Percentage", "",
                  "Bloom's Taxonomy Level (BL)", "", "Marks for BL", "Percentage"]
    hdr_fills  = [_NAVY_FILL, _NAVY_FILL, _NAVY_FILL, _NAVY_FILL, _NAVY_FILL,
                  _NAVY_FILL, _TEAL_FILL, _ORANGE_FILL, _LIME_FILL, _NAVY_FILL,
                  _NAVY_FILL, _SKYBLUE_FILL, _SKYBLUE_FILL, _NAVY_FILL,
                  _NAVY_FILL, _NAVY_FILL, _SKYBLUE_FILL, _SKYBLUE_FILL]
    for ci, (h, hf) in enumerate(zip(hdr_labels, hdr_fills), 1):
        _c(ws, r, ci, h, font=_HEADER_FONT, fill=hf)

    bloom = [("L1", "Remembering"), ("L2", "Understanding"), ("L3", "Applying"),
             ("L4", "Analyzing"),   ("L5", "Evaluating"),    ("L6", "Creating")]

    bl_row_fills = [_LIGHT_FILL, _GREEN_FILL, _YELLOW_FILL,
                    _ORANGE_FILL, _PINK_FILL, _LIME_FILL]

    for i in range(6):
        sr = r + 1 + i   # rows 8-13
        co_list_row = 14 + i
        bf = bl_row_fills[i]
        _c(ws, sr, 11, f'=IF(ISBLANK(CO_List!A{co_list_row}),"",CO_List!A{co_list_row})',
           fill=bf, align=_CENTER)
        _c(ws, sr, 12,
           f'=IF(SUMIF($H$8:$H$39,K{sr},$G$8:$G$39)=0,"",SUMIF($H$8:$H$39,K{sr},$G$8:$G$39))',
           fill=bf, align=_CENTER)
        _c(ws, sr, 13,
           f'=IFERROR(L{sr}/SUM($L$8:$L$13)*100,"")',
           fill=bf, align=_CENTER, number_format="0.00")
        bl_label, bl_name = bloom[i]
        _c(ws, sr, 15, bl_label, fill=bf, bold=True, align=_CENTER)
        _c(ws, sr, 16, bl_name,  fill=bf, align=_LEFT)
        _c(ws, sr, 17,
           f'=IF(SUMIF($I$8:$I$39,O{sr},$G$8:$G$39)=0,"",SUMIF($I$8:$I$39,O{sr},$G$8:$G$39))',
           fill=bf, align=_CENTER)
        _c(ws, sr, 18,
           f'=IFERROR(Q{sr}/SUM($Q$8:$Q$13)*100,"")',
           fill=bf, align=_CENTER, number_format="0.00")

    q_row = r + 1
    q_row_fills = [_GREEN_FILL, _LIGHT_FILL]
    if questions:
        for qi, q in enumerate(questions):
            qf = q_row_fills[qi % 2]
            _c(ws, q_row, 1, q.get("q_no", qi + 1), fill=qf, bold=True, align=_CENTER)
            ws.merge_cells(start_row=q_row, start_column=2, end_row=q_row, end_column=6)
            c = ws.cell(row=q_row, column=2, value=q.get("question_text", ""))
            c.font = _BODY_FONT
            c.fill = qf
            c.border = _BORDER
            c.alignment = _LEFT
            _c(ws, q_row, 7, q.get("marks", ""),  fill=_YELLOW_FILL, bold=True, align=_CENTER)
            _c(ws, q_row, 8, q.get("co_id", ""),  fill=_ORANGE_FILL, bold=True, align=_CENTER)
            bl_raw = str(q.get("bloom_level", ""))
            bl_str = bl_raw if bl_raw.startswith("L") else f"L{bl_raw}" if bl_raw else ""
            _c(ws, q_row, 9, bl_str, fill=_LIME_FILL, bold=True, align=_CENTER)
            q_row += 1
    else:
        for i in range(1, 16):
            qf = q_row_fills[i % 2]
            _c(ws, q_row, 1, i, fill=qf, bold=True, align=_CENTER)
            ws.merge_cells(start_row=q_row, start_column=2, end_row=q_row, end_column=6)
            c = ws.cell(row=q_row, column=2, value="")
            c.fill = qf
            c.border = _BORDER
            c.alignment = _LEFT
            _c(ws, q_row, 7, "", fill=_YELLOW_FILL, align=_CENTER)
            _c(ws, q_row, 8, "", fill=_ORANGE_FILL, align=_CENTER)
            _c(ws, q_row, 9, "", fill=_LIME_FILL,   align=_CENTER)
            q_row += 1

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 6
    ws.column_dimensions["K"].width = 6
    ws.column_dimensions["L"].width = 14
    ws.column_dimensions["M"].width = 12
    ws.column_dimensions["O"].width = 10
    ws.column_dimensions["P"].width = 14
    ws.column_dimensions["Q"].width = 12
    ws.column_dimensions["R"].width = 12
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# Marks sheet (ESE_MKS or CA{n}_Marks)
# ─────────────────────────────────────────────────────────────────────────────
def _build_marks_sheet(wb, sheet_name, course, ca_label, qp_sheet_name,
                       students, total_marks, saved_marks=None, saved_qp=None):
    ws = wb.create_sheet(sheet_name)
    _course_header_block(ws, course)

    n_students = len(students)
    data_start = 10
    data_end   = data_start + n_students - 1
    count_end  = max(data_end, 224)

    # ── Row 7: column headers ─────────────────────────────────────────────
    r = 7
    hdr_info  = ["Sr. No.", "Seat No", "Roll No.", "Name of the Student", ca_label]
    hdr_fills = [_NAVY_FILL, _NAVY_FILL, _NAVY_FILL, _NAVY_FILL, _TEAL_FILL]
    for ci, (h, hf) in enumerate(zip(hdr_info, hdr_fills), 1):
        _c(ws, r, ci, h, font=_HEADER_FONT, fill=hf)

    qp_ref = _qp_sheet_name(qp_sheet_name)
    n_qs = len(saved_qp) if saved_qp else _MAX_Q
    for qi in range(n_qs):
        qp_row = 8 + qi
        col = 6 + qi
        _c(ws, r, col,
           f'=IF({qp_ref}!$A{qp_row}=0,"",{qp_ref}!$A{qp_row})',
           font=_HEADER_FONT, fill=_NAVY_FILL)

    # ── Row 8: "Marks" label + CO mapping from QP col H ──────────────────
    r = 8
    _c(ws, r, 5, "Marks", fill=_LIGHT_FILL, bold=True, align=_CENTER)
    for qi in range(n_qs):
        qp_row = 8 + qi
        col = 6 + qi
        _c(ws, r, col,
           f'=IF({qp_ref}!$H{qp_row}=0,"",{qp_ref}!$H{qp_row})',
           fill=_ORANGE_FILL, bold=True, align=_CENTER)

    # ── Row 9: total marks + per-question max (from QP col G) ────────────
    r = 9
    _c(ws, r, 5, total_marks, fill=_GREEN_FILL, bold=True, align=_CENTER)
    for qi in range(n_qs):
        qp_row = 8 + qi
        col = 6 + qi
        _c(ws, r, col,
           f'=IF({qp_ref}!$G{qp_row}=0,"",{qp_ref}!$G{qp_row})',
           fill=_LIME_FILL, bold=True, align=_CENTER)

    # ── Build normalised saved marks lookup ───────────────────────────────
    _saved = {}
    if saved_marks:
        for prn_key, qmarks in saved_marks.items():
            try:
                norm = str(int(float(str(prn_key))))
            except (ValueError, TypeError):
                norm = str(prn_key).strip()
            _saved[norm] = qmarks

    # ── Rows 10+: students ────────────────────────────────────────────────
    last_q_col = get_column_letter(6 + n_qs - 1)
    row_fills = [_GREEN_FILL, _LIGHT_FILL]
    for idx, s in enumerate(students, 1):
        r = data_start + idx - 1
        fill = row_fills[idx % 2]
        prn_val = str(s["prn"]) if s["prn"] is not None else ""
        try:
            prn_norm = str(int(float(prn_val))) if prn_val else ""
        except (ValueError, TypeError):
            prn_norm = prn_val

        _c(ws, r, 1, idx,        fill=fill, bold=True,  align=_CENTER)
        _c(ws, r, 2, "",         fill=fill, align=_CENTER)
        _c(ws, r, 3, prn_val,    fill=fill, align=_CENTER)
        _c(ws, r, 4, s["name"],  fill=fill, align=_LEFT)
        _c(ws, r, 5,
           f"=SUM(F{r}:{last_q_col}{r})",
           fill=_YELLOW_FILL, bold=True, align=_CENTER)

        student_marks = _saved.get(prn_norm, {})
        for qi in range(n_qs):
            col = 6 + qi
            mark_val = None
            if student_marks:
                mark_val = student_marks.get(str(qi + 1))
                if mark_val is None and saved_qp and qi < len(saved_qp):
                    actual_qno = str(saved_qp[qi].get("q_no", qi + 1))
                    mark_val = student_marks.get(actual_qno)
                if mark_val is None:
                    mark_val = student_marks.get(qi + 1)
            _c(ws, r, col, mark_val, fill=fill, align=_CENTER)

    # ── Summary rows ──────────────────────────────────────────────────────
    s0 = count_end + 3

    cos    = [c["co_id"] for c in course.cos]
    n_cos  = len(cos)

    # s0: "No of students who attempted"
    _c(ws, s0, 1, "CO No",  bold=True, fill=_NAVY_FILL, font=_HEADER_FONT, align=_CENTER)
    _c(ws, s0, 2, "Level",  bold=True, fill=_NAVY_FILL, font=_HEADER_FONT, align=_CENTER)
    _c(ws, s0, 4, "No of students who attempted",
       bold=True, fill=_TEAL_FILL, align=_LEFT)
    for qi in range(n_qs):
        col = 6 + qi
        col_ltr = get_column_letter(col)
        _c(ws, s0, col,
           f'=IF(OR({col_ltr}$8="",{col_ltr}$9=""),"",COUNT({col_ltr}{data_start}:{col_ltr}{count_end}))',
           fill=_LIGHT_FILL, align=_CENTER)

    row_labels      = ["CO No", "Max", "Target", "No. of students scored >= target", "Percentage"]
    summary_fills   = [_LIGHT_FILL, _GREEN_FILL, _YELLOW_FILL, _ORANGE_FILL, _SKYBLUE_FILL]

    for i, co_id in enumerate(cos):
        r_co   = s0 + 1 + i
        sfill  = summary_fills[i % len(summary_fills)]
        _c(ws, r_co, 1, co_id,
           bold=True, fill=_NAVY_FILL, font=_HEADER_FONT, align=_CENTER)
        _c(ws, r_co, 4,
           row_labels[i] if i < len(row_labels) else "",
           bold=True, fill=sfill, align=_LEFT)

        for qi in range(n_qs):
            col     = 6 + qi
            col_ltr = get_column_letter(col)
            if i == 0:
                _c(ws, r_co, col,
                   f'=IF(OR({col_ltr}$8="",{col_ltr}$9=""),"",{col_ltr}8)',
                   fill=sfill, align=_CENTER)
            elif i == 1:
                _c(ws, r_co, col,
                   f'=IF(OR({col_ltr}$8="",{col_ltr}$9=""),"",{col_ltr}9)',
                   fill=sfill, align=_CENTER)
            elif i == 2:
                _c(ws, r_co, col,
                   f'=IF(OR({col_ltr}$8="",{col_ltr}$9=""),'
                   f'"",{col_ltr}{r_co-1}*VLOOKUP({col_ltr}{r_co-2},CO_List!$A$14:$J$19,10,0)/100)',
                   fill=sfill, align=_CENTER)
            elif i == 3:
                _c(ws, r_co, col,
                   f'=IF(OR({col_ltr}$8="",{col_ltr}$9=""),'
                   f'"",COUNTIFS({col_ltr}{data_start}:{col_ltr}{count_end},">="&{col_ltr}{r_co-1}))',
                   fill=sfill, align=_CENTER)
            elif i == 4:
                _c(ws, r_co, col,
                   f'=IF(OR({col_ltr}$8="",{col_ltr}$9=""),'
                   f'"",IFERROR({col_ltr}{r_co-1}/{col_ltr}{s0}*100,0))',
                   fill=sfill, align=_CENTER, number_format="0.00")

    # CO attainment level in col B (AVERAGEIF across percentage row)
    pct_row       = s0 + 5
    co_row_start  = s0 + 1
    q_start_col   = get_column_letter(6)
    q_end_col     = get_column_letter(6 + n_qs - 1)

    for i, co_id in enumerate(cos):
        r_co   = s0 + 1 + i
        sfill  = summary_fills[i % len(summary_fills)]
        _c(ws, r_co, 2,
           f'=IF(A{r_co}="","",IFERROR('
           f'IF(AVERAGEIF({q_start_col}{co_row_start}:{q_end_col}{co_row_start},A{r_co},'
           f'{q_start_col}{pct_row}:{q_end_col}{pct_row})>=70,3,'
           f'IF(AVERAGEIF({q_start_col}{co_row_start}:{q_end_col}{co_row_start},A{r_co},'
           f'{q_start_col}{pct_row}:{q_end_col}{pct_row})>40,2,1)),"")'
           f')',
           fill=_YELLOW_FILL, bold=True, align=_CENTER)

    # Level row
    r_level = s0 + 1 + n_cos
    _c(ws, r_level, 4, "Level", bold=True, fill=_TEAL_FILL, align=_LEFT)
    for qi in range(n_qs):
        col     = 6 + qi
        col_ltr = get_column_letter(col)
        _c(ws, r_level, col,
           f'=IF(OR({col_ltr}$8="",{col_ltr}$9=""),'
           f'"",IF({col_ltr}{pct_row}>=70,3,IF({col_ltr}{pct_row}>40,2,IF({col_ltr}{pct_row}<=40,1,0))))',
           fill=_SKYBLUE_FILL, bold=True, align=_CENTER)

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 12

    return {"s0": s0, "co_rows": [s0 + 1 + i for i in range(n_cos)]}


# ─────────────────────────────────────────────────────────────────────────────
# Final_CO_Attn
# ─────────────────────────────────────────────────────────────────────────────
def _build_final_co_attn(wb, course, ca_names, marks_meta, ese_meta):
    ws = wb.create_sheet("Final_CO_Attn")
    _course_header_block(ws, course)

    cos  = [c["co_id"] for c in course.cos]
    n_ca = len(ca_names)

    r = 7
    # Header row 7
    _c(ws, r, 1, "CO No / Weightage",
       fill=_NAVY_FILL, font=_HEADER_FONT, bold=True, align=_CENTER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=n_ca + 1)
    c = ws.cell(row=r, column=2, value="CO Attainment using CIE")
    c.font = _HEADER_FONT; c.fill = _TEAL_FILL; c.border = _BORDER; c.alignment = _CENTER
    ws.merge_cells(start_row=r, start_column=n_ca + 2, end_row=r, end_column=n_ca + 4)
    c = ws.cell(row=r, column=n_ca + 2, value="Final CO attainment")
    c.font = _HEADER_FONT; c.fill = _ORANGE_FILL; c.border = _BORDER; c.alignment = _CENTER
    _c(ws, r, n_ca + 5, "Overall Att",
       fill=_LIME_FILL, font=Font(name="Calibri", bold=True, size=10), bold=True)

    # Row 8: CA component names
    r = 8
    _c(ws, r, 1, "", fill=_LIGHT_FILL, align=_CENTER)
    for i, ca in enumerate(ca_names):
        _c(ws, r, i + 2, ca, bold=True, fill=_SKYBLUE_FILL, align=_CENTER)
    g = n_ca + 2
    _c(ws, r, g,     "Internal", bold=True, fill=_GREEN_FILL,  align=_CENTER)
    _c(ws, r, g + 1, "External", bold=True, fill=_ORANGE_FILL, align=_CENTER)
    _c(ws, r, g + 2, "Final",    bold=True, fill=_TEAL_FILL,   align=_CENTER)

    # Row 9: weightages
    r = 9
    _c(ws, r, 1, "", fill=_LIGHT_FILL, align=_CENTER)
    for i in range(n_ca):
        _c(ws, r, i + 2, "", fill=_LIGHT_FILL, align=_CENTER)
    _c(ws, r, g,     40,  fill=_GREEN_FILL,  bold=True, align=_CENTER)
    _c(ws, r, g + 1, 60,  fill=_ORANGE_FILL, bold=True, align=_CENTER)
    _c(ws, r, g + 2, 100, fill=_TEAL_FILL,   bold=True, align=_CENTER)
    wt_row = r

    # Rows 10+: one row per CO
    r = 10
    co_fills = [_GREEN_FILL, _LIGHT_FILL, _YELLOW_FILL, _ORANGE_FILL, _PINK_FILL]
    for ci, co_id in enumerate(cos):
        cfill = co_fills[ci % len(co_fills)]
        co_list_row = 14 + ci
        _c(ws, r, 1,
           f'=IF(CO_List!A{co_list_row}="","",CO_List!A{co_list_row})',
           fill=_NAVY_FILL, font=_HEADER_FONT, bold=True, align=_CENTER)

        for cai, ca in enumerate(ca_names):
            mks_sheet = _qp_sheet_name(f"{ca}_Marks")
            co_row    = marks_meta[cai]["co_rows"][ci]
            _c(ws, r, cai + 2, f"={mks_sheet}!B{co_row}",
               fill=cfill, bold=True, align=_CENTER)

        ca_cols = ",".join([f"{get_column_letter(cai + 2)}{r}" for cai in range(n_ca)])
        _c(ws, r, g,
           f'=IF(A{r}="","",IFERROR(AVERAGE({ca_cols}),""))',
           fill=_GREEN_FILL, bold=True, align=_CENTER)

        ese_co_row = ese_meta["co_rows"][ci]
        _c(ws, r, g + 1, f"=ESE_MKS!B{ese_co_row}",
           fill=_ORANGE_FILL, bold=True, align=_CENTER)

        g_ltr   = get_column_letter(g)
        gp1_ltr = get_column_letter(g + 1)
        _c(ws, r, g + 2,
           f'=IF(A{r}="","",IFERROR({g_ltr}{r}*${g_ltr}${wt_row}/100,0)'
           f'+IFERROR({gp1_ltr}{r}*${gp1_ltr}${wt_row}/100,0))',
           fill=_TEAL_FILL, bold=True, align=_CENTER, number_format="0.00")

        r += 1

    final_col = get_column_letter(g + 2)
    # Overall CO Attainment row
    _c(ws, r, 1, "Overall CO Attainment",
       bold=True, fill=_SKYBLUE_FILL, align=_LEFT)
    _c(ws, r, g + 4,
       f'=IFERROR(AVERAGE({final_col}10:{final_col}{r-1}),"")',
       fill=_LIME_FILL, bold=True, align=_CENTER, number_format="0.00")
    r += 2

    # Legend rows
    _c(ws, r, 1, "Final CO attainment", bold=True, fill=_LIGHT_FILL, align=_LEFT)
    r += 1
    _c(ws, r, 1, "External"); _c(ws, r, 2, 0.6, fill=_YELLOW_FILL, bold=True)
    _c(ws, r, 3, "OR"); _c(ws, r, 4, 1.0, fill=_GREEN_FILL, bold=True)
    _c(ws, r, 5, "OR"); _c(ws, r, 6, "Nil", fill=_ORANGE_FILL)
    r += 1
    _c(ws, r, 1, "Internal"); _c(ws, r, 2, 0.4, fill=_YELLOW_FILL, bold=True)
    _c(ws, r, 4, "NIL", fill=_LIGHT_FILL); _c(ws, r, 6, 1.0, fill=_GREEN_FILL, bold=True)
    r += 1
    _c(ws, r, 2, "Both", fill=_GREEN_FILL)
    _c(ws, r, 4, "Only ESE", fill=_ORANGE_FILL)
    _c(ws, r, 6, "Only CIE", fill=_SKYBLUE_FILL)

    ws.column_dimensions["A"].width = 20
    for i in range(n_ca + 6):
        ws.column_dimensions[get_column_letter(i + 2)].width = 12


# ─────────────────────────────────────────────────────────────────────────────
# PO_Attainment
# ─────────────────────────────────────────────────────────────────────────────
def _build_po_attainment(wb, course, n_ca):
    ws = wb.create_sheet("PO_Attainment")
    _course_header_block(ws, course)

    co_po  = course.co_po_matrix
    cos    = [c["co_id"] for c in course.cos]
    pos    = course.pos
    po_ids = [p["po_id"] for p in pos] if pos else [f"PO{i}" for i in range(1, 13)]

    r = 7
    _c(ws, r, 1, "CO",         bold=True, fill=_NAVY_FILL, font=_HEADER_FONT)
    _c(ws, r, 2, "Attainment", bold=True, fill=_NAVY_FILL, font=_HEADER_FONT)
    po_hdr_fills = [_TEAL_FILL, _GREEN_FILL, _ORANGE_FILL, _SKYBLUE_FILL,
                    _LIME_FILL, _PINK_FILL, _YELLOW_FILL, _LIGHT_FILL]
    for ci, po in enumerate(po_ids, 3):
        hf = po_hdr_fills[ci % len(po_hdr_fills)]
        _c(ws, r, ci, po, bold=True, fill=_NAVY_FILL, font=_HEADER_FONT)

    r = 8
    data_start = r
    final_col  = get_column_letter(n_ca + 4)

    co_fills = [_GREEN_FILL, _LIGHT_FILL, _YELLOW_FILL, _ORANGE_FILL, _PINK_FILL]
    for ci_idx, co_id in enumerate(cos):
        cfill = co_fills[ci_idx % len(co_fills)]
        _c(ws, r, 1, co_id, bold=True, fill=cfill, align=_CENTER)
        final_co_attn_row = 10 + ci_idx
        _c(ws, r, 2,
           f"=Final_CO_Attn!{final_col}{final_co_attn_row}",
           fill=_YELLOW_FILL, bold=True, align=_CENTER, number_format="0.00")
        mapping = co_po.get(co_id, {})
        for pi, po in enumerate(po_ids, 3):
            val = mapping.get(po, None)
            cell_fill = _GREEN_FILL if val else PatternFill()
            _c(ws, r, pi, val if val else "", fill=cell_fill,
               align=_CENTER, bold=bool(val))
        r += 1

    # Articulation Average row
    _c(ws, r, 1, "Articulation Average", bold=True, fill=_TEAL_FILL, align=_LEFT)
    _c(ws, r, 2, "", fill=_TEAL_FILL)
    for pi, po in enumerate(po_ids, 3):
        col = get_column_letter(pi)
        _c(ws, r, pi,
           f'=IFERROR(AVERAGEIF({col}{data_start}:{col}{r-1},"<>",{col}{data_start}:{col}{r-1}),"-")',
           fill=_ORANGE_FILL, bold=True, align=_CENTER, number_format="0.00")
    r += 1

    # CO-PO Attainment row
    _c(ws, r, 1, "CO-PO_PSO Attainment", bold=True, fill=_LIME_FILL, align=_LEFT)
    _c(ws, r, 2, "", fill=_LIME_FILL)
    attn_col_ltr = "B"
    for pi, po in enumerate(po_ids, 3):
        col = get_column_letter(pi)
        terms = "+".join(
            [f"{col}{data_start + i}*${attn_col_ltr}${data_start + i}"
             for i in range(len(cos))]
        )
        _c(ws, r, pi,
           f'=IFERROR(({terms})/(3*COUNT({col}{data_start}:{col}{data_start+len(cos)-1})),"-")',
           fill=_SKYBLUE_FILL, bold=True, align=_CENTER, number_format="0.00")

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 12
    for pi in range(3, 3 + len(po_ids)):
        ws.column_dimensions[get_column_letter(pi)].width = 8


# ─────────────────────────────────────────────────────────────────────────────
# Main service
# ─────────────────────────────────────────────────────────────────────────────
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

    async def _get_questions(self, course_id: int):
        from sqlalchemy import select
        from backend.database.models import Question
        result = await self.db.execute(
            select(Question).where(Question.course_id == course_id)
        )
        return result.scalars().all()

    async def _get_saved_sheets(self, course_id: int) -> dict:
        """Load all saved CASheet records. Returns {ca_label: {qp:[...], marks:{...}}}."""
        from sqlalchemy import select
        from backend.database.models import CASheet
        try:
            result = await self.db.execute(
                select(CASheet).where(CASheet.course_id == course_id)
            )
            return {s.ca_label: {"qp": s.qp, "marks": s.marks}
                    for s in result.scalars().all()}
        except Exception as e:
            logger.warning(f"Could not load saved CA sheets: {e}")
            return {}

    async def generate(self, course_id: int, qp_source: str = "blank") -> dict:
        course_svc = CourseService(self.db)
        course     = await course_svc.get_course(course_id)
        students   = await self._get_students(course_id)
        eval_cfg   = course.evaluation_config
        components = eval_cfg.get("components", {})

        # Load all saved CA sheets (QP + marks entered by the user in the frontend)
        saved_sheets = await self._get_saved_sheets(course_id)

        # Build CA name list: start from evaluation_config components (excluding ESE),
        # then add any saved sheet labels not already present
        ESE_KEYWORDS = {"end semester", "ese", "end-semester", "final exam"}
        def _is_ese(name):
            return any(k in name.lower() for k in ESE_KEYWORDS)

        comp_ca_names = sorted([k for k in components.keys() if not _is_ese(k)])
        saved_ca_names = [k for k in saved_sheets.keys() if not _is_ese(k)]

        ca_names = list(comp_ca_names)
        for label in saved_ca_names:
            if label not in ca_names:
                ca_names.append(label)
        if not ca_names:
            ca_names = [f"CA{i}" for i in range(1, 4)]
        ca_names = ca_names[:5]

        wb = Workbook()
        del wb[wb.sheetnames[0]]

        # 1. Course_Info
        _build_course_info(wb, course)

        # 2. Roll_List
        _build_roll_list(wb, course, students)

        # 3. CO_List
        _build_co_list(wb, course)

        # 4. ESE sheets — prefer saved ESE data, fall back to question bank
        ese_saved       = saved_sheets.get("ESE", {})
        ese_saved_qp    = ese_saved.get("qp") or []
        ese_saved_marks = ese_saved.get("marks") or {}
        ese_questions   = ese_saved_qp or None
        if not ese_questions and qp_source == "question_bank":
            qs = await self._get_questions(course_id)
            ese_questions = [
                {"q_no": i + 1, "question_text": q.question_text,
                 "marks": q.marks, "co_id": q.co_id, "bloom_level": q.bloom_level}
                for i, q in enumerate(qs[:20])
            ]
        _build_qp_sheet(wb, "ESE_QP", course, "ESE", questions=ese_questions)
        ese_total = eval_cfg.get("end_sem_total", 60)
        ese_meta  = _build_marks_sheet(
            wb, "ESE_MKS", course, "ESE", "ESE_QP", students, ese_total,
            saved_marks=ese_saved_marks,
            saved_qp=ese_saved_qp or ese_questions,
        )

        # 5. CA sheets
        marks_meta_list = []
        for ca in ca_names:
            qp_name  = f"{ca}_QP"
            mks_name = f"{ca}_Marks"

            ca_saved       = saved_sheets.get(ca, {})
            ca_saved_qp    = ca_saved.get("qp") or []
            ca_saved_marks = ca_saved.get("marks") or {}
            ca_questions   = ca_saved_qp or None
            if not ca_questions and qp_source == "question_bank":
                qs = await self._get_questions(course_id)
                ca_questions = [
                    {"q_no": i + 1, "question_text": q.question_text,
                     "marks": q.marks, "co_id": q.co_id, "bloom_level": q.bloom_level}
                    for i, q in enumerate(qs[:15])
                ]
            _build_qp_sheet(wb, qp_name, course, ca, questions=ca_questions)
            ca_total = components.get(ca, components.get(ca.lower(), 10))
            if not isinstance(ca_total, (int, float)):
                ca_total = 10
            meta = _build_marks_sheet(
                wb, mks_name, course, ca, qp_name, students, ca_total,
                saved_marks=ca_saved_marks,
                saved_qp=ca_saved_qp or ca_questions,
            )
            marks_meta_list.append(meta)

        # 6. Final_CO_Attn
        _build_final_co_attn(wb, course, ca_names, marks_meta_list, ese_meta)

        # 7. PO_Attainment
        _build_po_attainment(wb, course, len(ca_names))

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
