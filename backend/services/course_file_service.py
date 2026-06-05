"""
backend/services/course_file_service.py

Generates the complete OBE Course File (.docx) with all 13 sections
using python-docx only — no Node.js dependency.
"""

import io
import json
from pathlib import Path
from typing import Optional

from sqlalchemy import text, select
from sqlalchemy.ext.asyncio import AsyncSession

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.services.course_service import CourseService

logger = get_logger(__name__)

_CATEGORY = "course_files"

# ── Colours ───────────────────────────────────────────────────────────────────
_NAVY   = (31,  56, 100)
_LIGHT  = (214, 220, 228)
_GREEN  = (226, 239, 218)
_ORANGE = (252, 228, 214)
_WHITE  = (255, 255, 255)
_LGRAY  = (245, 245, 245)


# ── Low-level XML helpers ─────────────────────────────────────────────────────

def _set_cell_bg(cell, rgb: tuple):
    r, g, b = rgb
    hex_color = f"{r:02X}{g:02X}{b:02X}"
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _rgb(tup):
    return RGBColor(*tup)


# ── Style helpers ─────────────────────────────────────────────────────────────

def _run(para, text, bold=False, size=10, color=None, underline=False):
    run = para.add_run(str(text or ""))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = _rgb(color or (0, 0, 0))
    run.underline = underline
    return run


def _add_para(doc_or_cell, text="", bold=False, size=10, color=None,
              align=WD_ALIGN_PARAGRAPH.LEFT, space_before=3, space_after=3):
    if hasattr(doc_or_cell, "add_paragraph"):
        p = doc_or_cell.add_paragraph()
    else:
        p = doc_or_cell.paragraphs[0] if doc_or_cell.paragraphs else doc_or_cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if text:
        _run(p, text, bold=bold, size=size, color=color)
    return p


def _section_title(doc, num, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(6)
    if num > 1:
        run = p.add_run()
        run.add_break(__import__("docx.enum.text", fromlist=["WD_BREAK"]).WD_BREAK.PAGE)
    _run(p, f"{num}. {title}", bold=True, size=14, color=_NAVY)
    return p


def _heading2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(4)
    _run(p, text, bold=True, size=12)
    return p


# ── Table builder ─────────────────────────────────────────────────────────────

def _make_table(doc, headers, rows, col_widths_cm):
    num_cols = len(headers)
    tbl = doc.add_table(rows=1 + len(rows), cols=num_cols)
    tbl.style = "Table Grid"

    hdr_row = tbl.rows[0]
    for i, hdr in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.width = Cm(col_widths_cm[i])
        _set_cell_bg(cell, _NAVY)
        _set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, hdr, bold=True, size=9, color=_WHITE)

    for ri, row in enumerate(rows):
        bg = _WHITE if ri % 2 == 0 else _LIGHT
        tr = tbl.rows[ri + 1]
        for ci, val in enumerate(row):
            cell = tr.cells[ci]
            cell.width = Cm(col_widths_cm[ci])
            _set_cell_bg(cell, bg)
            _set_cell_margins(cell)
            p = cell.paragraphs[0]
            _run(p, str(val or ""), size=9)

    return tbl


def _make_header_table(doc, institution_name, institution_address, title, subtitle=""):
    """Creates the header block used in the reference docx (SIT-style merged header)."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    _set_cell_bg(cell, _NAVY)
    _set_cell_margins(cell, top=100, bottom=100, left=160, right=160)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, institution_name or "Symbiosis Institute of Technology", bold=True, size=11, color=_WHITE)
    if institution_address:
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p2, institution_address, size=9, color=_WHITE)
    if title:
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p3, title, bold=True, size=10, color=_WHITE)
    if subtitle:
        p4 = cell.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p4, subtitle, size=9, color=_WHITE)
    return tbl


# ── Key-safe getter (handles camelCase / snake_case) ─────────────────────────

def _g(row, *keys, default=""):
    """Try multiple key variants; return first non-empty value."""
    for k in keys:
        v = row.get(k)
        if v is not None and str(v).strip():
            return v
    return default


# ── Timetable grid renderer ───────────────────────────────────────────────────

def _render_timetable(doc, timetable: dict):
    """Render the timetable as a proper day × time-slot grid."""
    if not timetable:
        _add_para(doc, "[Timetable not yet uploaded. Upload via the dashboard timetable upload.]",
                  color=(136, 136, 136))
        return

    faculty = timetable.get("faculty_name", "")
    dept    = timetable.get("department", "")
    ay      = timetable.get("academic_year", "")
    slots   = timetable.get("time_slots", [])
    schedule= timetable.get("schedule", {})  # {day: {slot: entry}}

    if not slots and not schedule:
        _add_para(doc, "[Timetable data is incomplete. Re-upload the timetable docx.]",
                  color=(136, 136, 136))
        return

    DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

    num_cols = 1 + len(slots)
    tbl = doc.add_table(rows=1, cols=num_cols)
    tbl.style = "Table Grid"

    # Merged header rows
    for label in [dept or "Department of AIML",
                  f"Individual Timetable {ay}",
                  faculty]:
        row = tbl.add_row()
        # Merge all cells in row
        merged = row.cells[0]
        for ci in range(1, num_cols):
            merged = merged.merge(row.cells[ci])
        _set_cell_bg(merged, _NAVY)
        _set_cell_margins(merged, top=60, bottom=60)
        p = merged.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, label, bold=True, size=9, color=_WHITE)

    # Delete the auto-created first row (it was a placeholder)
    tbl._tbl.remove(tbl.rows[0]._tr)

    # Header row: Day/Time | slot1 | slot2 ...
    hdr_row = tbl.add_row()
    hdr_row.cells[0].width = Cm(2.0)
    _set_cell_bg(hdr_row.cells[0], _NAVY)
    _set_cell_margins(hdr_row.cells[0])
    p = hdr_row.cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, "Day / Time", bold=True, size=8, color=_WHITE)

    slot_w = round(14.0 / max(len(slots), 1), 2)
    for si, slot in enumerate(slots):
        c = hdr_row.cells[si + 1]
        c.width = Cm(slot_w)
        _set_cell_bg(c, _NAVY)
        _set_cell_margins(c, left=60, right=60)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, slot, bold=True, size=8, color=_WHITE)

    # Build a normalised slot → entry dict for each day
    # The parser stores schedule[day] as either:
    #   - list of {time, course, section, room}
    #   - dict of {time_slot_str: entry_str}
    def _normalise_day(day_data, slots):
        if isinstance(day_data, dict):
            return day_data  # already {slot: entry}
        if isinstance(day_data, list):
            mapping = {}
            for item in day_data:
                t = item.get("time", "")
                course  = item.get("course","")
                section = item.get("section","")
                room    = item.get("room","")
                parts = [p for p in [course, section, room] if p]
                mapping[t] = "\n".join(parts)
            return mapping
        return {}

    # Data rows
    for di, day in enumerate(DAYS):
        raw_day = schedule.get(day, {})
        day_map = _normalise_day(raw_day, slots)
        row = tbl.add_row()
        bg = _WHITE if di % 2 == 0 else _LIGHT
        row.cells[0].width = Cm(2.0)
        _set_cell_bg(row.cells[0], (220, 225, 235))
        _set_cell_margins(row.cells[0])
        p = row.cells[0].paragraphs[0]
        _run(p, day, bold=True, size=9)

        for si, slot in enumerate(slots):
            entry = day_map.get(slot, "")
            c = row.cells[si + 1]
            c.width = Cm(slot_w)
            _set_cell_bg(c, bg)
            _set_cell_margins(c, left=60, right=60)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p, str(entry or ""), size=8)


# ── Main document builder ─────────────────────────────────────────────────────

def _build_docx(data: dict) -> bytes:
    doc = Document()

    for section in doc.sections:
        section.page_width  = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin   = section.right_margin  = Cm(2.5)
        section.top_margin    = section.bottom_margin = Cm(2)

    inst_name = data.get("institution_name") or "Symbiosis Institute of Technology"
    inst_addr = data.get("institution_address") or "SIU Pune 412115, Maharashtra, India"

    # ── Cover ──────────────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, inst_name, bold=True, size=14)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, inst_addr, size=10, color=(80, 80, 80))

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, f"Department of {data.get('department','')}", bold=True, size=13)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, f"{data.get('course_name','')} ({data.get('course_code','')}) — Course File",
         bold=True, size=16, color=_NAVY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, f"A.Y {data.get('academic_year','')}  |  {data.get('semester','')} Semester", size=12)

    if data.get("batch"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, f"Batch {data['batch']}", size=12)

    if data.get("faculty_name"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, f"Faculty: {data['faculty_name']}", size=12)

    doc.add_paragraph()

    # Table of Contents
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, "Course File Contents", bold=True, size=13, color=_NAVY)

    toc_entries = [
        "Vision & Mission of the Department",
        "Program Outcomes (POs), Program Educational Objectives (PEOs) and Program Specific Outcomes (PSOs)",
        "Syllabus, Personal Timetable",
        "CO Statements, CO-PO-PSO Mapping with justification",
        "CO Attainment from previous academic year and the action plan",
        "Session Plan",
        "Evaluation plan with CO Mapping",
        "List of Slow and Advanced learners and the action plans",
        "CO Attainment of internal evaluation",
        "Reports of activities planned and conducted",
        "Learning Material",
        "Question Bank",
        "Compiled Attendance",
    ]
    _make_table(doc,
                ["Sr. No", "Section"],
                [[str(i+1), t] for i, t in enumerate(toc_entries)],
                [1.5, 14.5])

    # ── 1. Vision & Mission ────────────────────────────────────────────────────
    _section_title(doc, 1, "Vision & Mission of the Department")
    if data.get("vision_text"):
        _heading2(doc, "VISION OF THE DEPARTMENT")
        _add_para(doc, data["vision_text"])
    if data.get("mission_text"):
        _heading2(doc, "MISSION OF THE DEPARTMENT")
        for line in (data["mission_text"] or "").split("\n"):
            if line.strip():
                _add_para(doc, line)
    if not data.get("vision_text") and not data.get("mission_text"):
        _add_para(doc, "[Vision & Mission not yet filled. Edit in Course File section.]",
                  color=(136, 136, 136))

    # ── 2. POs, PEOs, PSOs ────────────────────────────────────────────────────
    _section_title(doc, 2, "Program Outcomes (POs), Program Educational Objectives (PEOs) and Program Specific Outcomes (PSOs)")
    _heading2(doc, "Program Outcomes (POs)")
    pos = data.get("pos") or []
    _STANDARD_POS = [
        ("PO 1",  "Engineering Knowledge: Apply the knowledge of mathematics, science, engineering fundamentals, and an engineering specialization to the solution of complex engineering problems."),
        ("PO 2",  "Problem analysis: Identify, formulate, review research literature, and analyze complex engineering problems reaching substantiated conclusions using first principles of mathematics, natural sciences, and engineering sciences."),
        ("PO 3",  "Design/development of solutions: Design solutions for complex engineering problems and design system components or processes that meet the specified needs with appropriate consideration for the public health and safety, and the cultural, societal, and environmental considerations."),
        ("PO 4",  "Conduct investigations of complex problems: Use research-based knowledge and research methods including design of experiments, analysis and interpretation of data, and synthesis of the information to provide valid conclusions."),
        ("PO 5",  "Modern tool usage: Create, select, and apply appropriate techniques, resources, and modern engineering and IT tools including prediction and modeling to complex engineering activities with an understanding of the limitations."),
        ("PO 6",  "The engineer and society: Apply reasoning informed by the contextual knowledge to assess societal, health, safety, legal and cultural issues and the consequent responsibilities relevant to the professional engineering practice."),
        ("PO 7",  "Environment and sustainability: Understand the impact of the professional engineering solutions in societal and environmental contexts, and demonstrate the knowledge of, and need for sustainable development."),
        ("PO 8",  "Ethics: Apply ethical principles and commit to professional ethics and responsibilities and norms of the engineering practice."),
        ("PO 9",  "Individual and team work: Function effectively as an individual, and as a member or leader in diverse teams, and in multidisciplinary settings."),
        ("PO 10", "Communication: Communicate effectively on complex engineering activities with the engineering community and with society at large, such as, being able to comprehend and write effective reports and design documentation, make effective presentations, and give and receive clear instructions."),
        ("PO 11", "Project management and finance: Demonstrate knowledge and understanding of the engineering and management principles and apply these to one's own work, as a member and leader in a team, to manage projects and in multidisciplinary environments."),
        ("PO 12", "Life-long learning: Recognize the need for, and have the preparation and ability to engage in independent and life-long learning in the broadest context of technological change."),
    ]
    db_pos_have_text = any(p.get("statement", p.get("description", "")).strip() for p in pos)
    po_rows = (
        [[p.get("po_id",""), p.get("statement", p.get("description",""))] for p in pos]
        if db_pos_have_text
        else [[pid, stmt] for pid, stmt in _STANDARD_POS]
    )
    _make_table(doc, ["", "Program Outcomes"], po_rows, [1.5, 14.5])

    _heading2(doc, "Program Educational Objectives (PEOs)")
    peos = [
        ("PEO1", "Apply the knowledge of the latest trends of AIML and will be engaged in technology development and deployment for engineering systems in their profession."),
        ("PEO2", "To be competent AIML engineers with innovative thinking and research attitude to solve the real-world problems."),
        ("PEO3", "To have enhanced interpersonal and managerial skills to function effectively in their profession with social awareness and responsibility."),
    ]
    for pid, ptext in peos:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after  = Pt(3)
        r = p.add_run(f"{pid}: ")
        r.bold = True
        r.font.size = Pt(10)
        r2 = p.add_run(ptext)
        r2.font.size = Pt(10)

    _heading2(doc, "Program Specific Outcomes (PSOs)")
    psos = [
        ("PSO1", "To apply the concepts of Artificial Intelligence and Machine Learning with practical knowledge in analysis, design and development of intelligent systems and applications to multi-disciplinary problems."),
        ("PSO2", "To provide a concrete foundation to the students in the cutting-edge areas Artificial Intelligence and Machine Learning and excelling in the specialized areas like Natural Language Processing, Computer Vision, Reinforcement Learning, Internet of Things, Cloud computing, Data Security and privacy etc."),
    ]
    _make_table(doc, ["", "Program Specific Outcomes"], [[pid, ptext] for pid, ptext in psos], [1.5, 14.5])

    # ── 3. Syllabus & Timetable ───────────────────────────────────────────────
    _section_title(doc, 3, "Syllabus, Personal Timetable")
    _heading2(doc, "Syllabus")
    syllabus_units = data.get("syllabus_units") or []
    if syllabus_units:
        for u in syllabus_units:
            _add_para(doc, f"Unit {u.get('unit_number','')}: {u.get('unit_title','')}",
                      bold=True, size=11, space_before=6, space_after=2)
            for t in u.get("topics", []):
                _add_para(doc, f"  • {t}", space_before=1, space_after=1)
    else:
        _add_para(doc, "[Syllabus will be extracted from session plan. Generate session plan first.]",
                  color=(136, 136, 136))

    _heading2(doc, "Individual Timetable:")
    timetable = data.get("timetable") or {}
    _render_timetable(doc, timetable)

    # Additional timetables from attachments (other faculty)
    extra_tt_atts = [a for a in (data.get("attachments") or []) if a.get("section_no") == 3]
    if extra_tt_atts:
        _heading2(doc, "Additional Timetable Uploads")
        for a in extra_tt_atts:
            _add_para(doc, f"📎 {a['label']}  ({a['filename']})", size=10)

    # ── List of Students ──────────────────────────────────────────────────────
    _heading2(doc, "List of Students")
    students_all = data.get("students") or []
    if students_all:
        from collections import defaultdict as _dd
        by_section = _dd(list)
        for s in students_all:
            sec = (s.get("section") or "").strip().upper() or "All"
            by_section[sec].append(s)
        for sec_label in sorted(by_section.keys()):
            sec_students = by_section[sec_label]
            if len(by_section) > 1:
                _add_para(doc, f"Section {sec_label}", bold=True, size=11, space_before=6, space_after=2)
            _make_table(
                doc,
                ["Sr. No", "PRN", "Name"],
                [[str(i+1), s.get("prn",""), s.get("name","")] for i, s in enumerate(sec_students)],
                [1.2, 3.5, 11.3],
            )
    else:
        _add_para(doc, "[Student list not available. Add students via the Students page.]",
                  color=(136, 136, 136))

    # ── 4. CO Statements + CO-PO Mapping ─────────────────────────────────────
    _section_title(doc, 4, "CO Statements, CO-PO-PSO Mapping with justification")
    cos = data.get("cos") or []
    if cos:
        _make_table(doc,
                    ["CO", "Statement", "Bloom's Level"],
                    [[c.get("co_id",""), c.get("statement",""), c.get("bloom_level","")] for c in cos],
                    [1.5, 12.5, 2.0])

    _heading2(doc, "CO-PO Mapping")
    co_po_matrix = data.get("co_po_matrix") or {}
    # Use standard 12 POs + 2 PSOs for the matrix header
    po_ids_matrix = [f"PO{i}" for i in range(1, 13)] + ["PSO1", "PSO2"]
    if co_po_matrix and cos:
        matrix_rows = []
        for co in cos:
            mapping = co_po_matrix.get(co.get("co_id", "")) or {}
            # Try both "PO1" and "PO 1" styles
            def _get_mapping(pid):
                return mapping.get(pid) or mapping.get(pid.replace("PO", "PO ")) or mapping.get(pid.replace("PO ", "PO")) or "-"
            matrix_rows.append([co.get("co_id","")] + [_get_mapping(pid) for pid in po_ids_matrix])
        n = len(po_ids_matrix)
        col_w = [1.5] + [round(14.5/max(n,1), 2)] * n

        # Build table manually so we can colour cells by strength value
        num_cols = 1 + n
        tbl = doc.add_table(rows=1 + len(matrix_rows), cols=num_cols)
        tbl.style = "Table Grid"
        hdr_row = tbl.rows[0]
        for ci, hdr in enumerate(["CO \\ PO/PSO"] + po_ids_matrix):
            cell = hdr_row.cells[ci]
            cell.width = Cm(col_w[ci])
            _set_cell_bg(cell, _NAVY)
            _set_cell_margins(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p, hdr, bold=True, size=8, color=_WHITE)

        # Strength colour map: 1=light blue, 2=medium blue, 3=dark blue — all with BLACK text
        strength_bg = {
            "1": (209, 231, 246),
            "2": (130, 188, 235),
            "3": ( 56, 136, 195),
            1:   (209, 231, 246),
            2:   (130, 188, 235),
            3:   ( 56, 136, 195),
        }
        for ri, row in enumerate(matrix_rows):
            tr = tbl.rows[ri + 1]
            for ci, val in enumerate(row):
                cell = tr.cells[ci]
                cell.width = Cm(col_w[ci])
                _set_cell_margins(cell)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT
                str_val = str(val) if val is not None else "-"
                if ci == 0:
                    _set_cell_bg(cell, _LGRAY)
                    _run(p, str_val, bold=True, size=8, color=(0, 0, 0))
                elif str_val in strength_bg or val in strength_bg:
                    _set_cell_bg(cell, strength_bg.get(str_val) or strength_bg.get(val) or _WHITE)
                    _run(p, str_val, bold=True, size=8, color=(0, 0, 0))
                else:
                    _set_cell_bg(cell, _WHITE)
                    _run(p, "-" if str_val in ("-", "0", "") else str_val, size=8, color=(150, 150, 150))
    elif cos:
        _add_para(doc, "[CO-PO mapping not yet configured. Set it up in Course Setup.]", color=(136, 136, 136))

    # CO-PO Justification
    co_po_justification = data.get("co_po_justification") or ""
    if co_po_justification:
        _heading2(doc, "CO-PO Mapping Justification")
        for line in co_po_justification.split("\n"):
            if line.strip():
                _add_para(doc, line)

    # ── 5. Previous CO Attainment ─────────────────────────────────────────────
    _section_title(doc, 5, "CO Attainment from previous academic year and the action plan")
    if data.get("prev_co_attainment"):
        _add_para(doc, data["prev_co_attainment"])
    else:
        _add_para(doc, "[Previous year CO attainment data not yet entered.]", color=(136, 136, 136))
    _heading2(doc, "Action Plan")
    if data.get("action_plan"):
        _add_para(doc, data["action_plan"])
    else:
        _add_para(doc, "[Action plan not yet entered.]", color=(136, 136, 136))

    # ── 6. Session Plan ───────────────────────────────────────────────────────
    _section_title(doc, 6, "Session Plan with CO mapping to each lecture")

    # Header block like real doc
    _make_header_table(doc, inst_name, None,
                       f"Session Plan — {data.get('department','')}",
                       f"Course: {data.get('course_name','')} ({data.get('course_code','')})  |  Faculty: {data.get('faculty_name','')}")

    doc.add_paragraph()

    session_rows = data.get("session_rows") or []
    if session_rows:
        table_data = []
        for i, r in enumerate(session_rows):
            lect   = _g(r, "lect", "lect_no", "lectNo", "lecture_no")
            unit   = _g(r, "unit", "unit_no", "unitNo", "unit_number")
            topic  = _g(r, "topic", "points_to_cover", "pointsToCover", "content", "description")
            method = _g(r, "method", "methodology", "lecture_method")
            ltype  = _g(r, "type", "lecture_exp_eval", "lectureType", default="Lecture")
            co     = _g(r, "co", "co_mapped", "co_id")
            if isinstance(co, list):
                co = ", ".join(str(x) for x in co)
            table_data.append([str(lect or i+1), str(unit or ""), topic, method, ltype, co])
        _make_table(doc,
                    ["Lect. No", "Unit No", "Points to Cover", "Methodology", "Type", "CO Mapped"],
                    table_data,
                    [1.5, 1.5, 8.0, 2.5, 2.0, 2.0])
    else:
        _add_para(doc, "[Session plan not yet generated. Use the Session Plan page first.]",
                  color=(136, 136, 136))

    # Textbooks & References
    materials = data.get("study_materials") or {}
    textbooks = materials.get("textbooks") or []
    ref_books  = materials.get("reference_books") or materials.get("references") or []
    web_links  = materials.get("web_links") or materials.get("web") or []
    journals   = materials.get("journals") or []
    moocs      = materials.get("moocs") or []

    if textbooks:
        _heading2(doc, "Textbooks")
        _make_table(doc, ["Book", "Author", "Publisher"],
                    [[b.get("title",b.get("book","") if isinstance(b,dict) else str(b)),
                      b.get("author","") if isinstance(b,dict) else "",
                      b.get("publisher","") if isinstance(b,dict) else ""] for b in textbooks],
                    [7.0, 4.0, 5.0])

    if ref_books:
        _heading2(doc, "Reference Books")
        _make_table(doc, ["Book", "Author", "Publisher"],
                    [[b.get("title",b.get("book","") if isinstance(b,dict) else str(b)),
                      b.get("author","") if isinstance(b,dict) else "",
                      b.get("publisher","") if isinstance(b,dict) else ""] for b in ref_books],
                    [7.0, 4.0, 5.0])

    if web_links:
        _heading2(doc, "Web Links / NPTEL")
        _make_table(doc, ["Sr. No.", "Web Link", "Module"],
                    [[str(i+1),
                      w.get("title",w.get("url","") if isinstance(w,dict) else str(w)),
                      w.get("unit",w.get("module","")) if isinstance(w,dict) else ""] for i,w in enumerate(web_links)],
                    [1.0, 10.0, 5.0])

    if journals:
        _heading2(doc, "Journals / Research Articles")
        _make_table(doc, ["Sr. No.", "Journal"],
                    [[str(i+1),
                      j.get("title","") if isinstance(j,dict) else str(j)] for i,j in enumerate(journals)],
                    [1.0, 15.0])

    if moocs:
        _heading2(doc, "MOOC Courses")
        _make_table(doc, ["S.No.", "MOOC Course Link", "Course conducted by", "Course Duration", "Certificate (Y/N)"],
                    [[str(i+1),
                      m.get("title",m.get("url","") if isinstance(m,dict) else str(m)),
                      m.get("platform",m.get("conducted_by","")) if isinstance(m,dict) else "",
                      m.get("duration","") if isinstance(m,dict) else "",
                      m.get("certificate","Y") if isinstance(m,dict) else "Y"] for i,m in enumerate(moocs)],
                    [1.0, 6.0, 3.0, 2.5, 2.0])

    # Tutorial questions
    if data.get("tutorial_questions"):
        _heading2(doc, "Tutorial Questions with CO Mapping")
        _make_table(doc,
                    ["Q No", "Question", "CO"],
                    [[str(i+1), q.get("question_text",""), q.get("co_id","")] for i, q in enumerate(data["tutorial_questions"])],
                    [1.0, 13.5, 2.0])

    # ── 7. Evaluation Plan & Marksheets ──────────────────────────────────────
    _section_title(doc, 7, "Evaluation plan with CO Mapping")
    eval_rows = data.get("eval_rows") or []
    if eval_rows:
        table_data = []
        for i, r in enumerate(eval_rows):
            sr   = str(i + 1)   # always auto-number
            comp = _g(r, "comp", "component", "name")
            units= _g(r, "unit_syllabus", "units", "syllabus", "unit")
            co   = _g(r, "co", "co_mapped")
            if isinstance(co, list):
                co = ", ".join(str(x) for x in co)
            marks= _g(r, "marks", "total_marks", "max_marks")
            wt   = _g(r, "weightage", "weight")
            date = _g(r, "date", "tentative_date")
            table_data.append([sr, comp, units, co, marks, wt, date])
        _make_table(doc,
                    ["Sr.No", "Component", "Units/Syllabus", "CO Mapped", "Marks", "Weightage", "Tentative Date"],
                    table_data,
                    [1.0, 3.0, 4.5, 2.0, 1.2, 1.8, 3.0])
    else:
        _add_para(doc, "[Evaluation plan not yet generated. Use the Evaluation Plan page first.]",
                  color=(136, 136, 136))

    # ── 7b. Question Papers ───────────────────────────────────────────────────
    _heading2(doc, "Question Papers")
    bloom_map = {1: "Remember", 2: "Understand", 3: "Apply",
                 4: "Analyse",  5: "Evaluate",   6: "Create"}

    ca_sheets_all = data.get("ca_sheets") or []
    any_qp = any((ca.get("qp") or []) for ca in ca_sheets_all)
    if any_qp:
        for ca in ca_sheets_all:
            qp = ca.get("qp") or []
            if not qp:
                continue
            _heading2(doc, f"{ca.get('ca_label','')} — Question Paper")
            bl_map_for_ca = data.get("bloom_ai_map", {}).get(ca.get("ca_label",""), {})
            qp_rows = []
            for q in qp:
                q_no   = q.get("q_no","")
                q_text = q.get("question_text","")
                marks  = q.get("marks","")
                co     = q.get("co_id","")
                # AI-mapped bloom level: use stored value, AI override, or raw int→name
                bl_raw = q.get("bloom_level","")
                if bl_raw and isinstance(bl_raw, int):
                    bl_str = bloom_map.get(bl_raw, str(bl_raw))
                elif bl_raw:
                    bl_str = str(bl_raw)
                else:
                    bl_str = ""
                # Apply AI override if present
                bl_str = bl_map_for_ca.get(str(q_no), bl_str) or bl_str
                qp_rows.append([q_no, q_text, marks, co, bl_str])
            _make_table(doc,
                        ["Q.No", "Question", "Marks", "CO", "Bloom's Level"],
                        qp_rows,
                        [1.0, 9.5, 1.2, 1.8, 2.0])
    else:
        _add_para(doc, "[No question papers uploaded yet. Upload via the Evaluation Plan page.]",
                  color=(136, 136, 136))

    # ── 8. Student Marks ─────────────────────────────────────────────────────
    _section_title(doc, 8, "Student Marks")
    _add_para(doc, "Exam-wise and question-wise marks auto-populated from uploaded marks and "
                   "master attainment file.", size=9, color=(80, 80, 80))

    student_map = data.get("student_map") or {}
    students_list = data.get("students") or []
    has_any_marks_section = False

    for ca in ca_sheets_all:
        qp         = ca.get("qp") or []
        marks_data = ca.get("marks") or {}
        ca_label   = ca.get("ca_label", "")
        if not qp:
            continue

        has_real_marks = any(
            any(float(v or 0) > 0 for v in mks.values())
            for mks in marks_data.values()
            if isinstance(mks, dict)
        )

        q_nos       = [q.get("q_no","") for q in qp]
        q_max_marks = [float(q.get("marks", 0) or 0) for q in qp]
        total_max   = sum(q_max_marks)
        num_students = len(students_list)

        # Compute per-question averages and overall avg
        if has_real_marks and num_students:
            q_avgs = []
            for qi, q in enumerate(qp):
                vals = [
                    float((marks_data.get(s["prn"]) or {}).get(q.get("q_no"), 0) or 0)
                    for s in students_list
                ]
                q_avgs.append(f"{sum(vals)/len(vals):.1f}" if vals else "—")
            all_totals = []
            for s in students_list:
                mks = marks_data.get(s["prn"]) or {}
                all_totals.append(sum(float(mks.get(q.get("q_no"),0) or 0) for q in qp))
            overall_avg = f"{sum(all_totals)/len(all_totals):.1f}" if all_totals else "—"
        else:
            q_avgs = ["—"] * len(qp)
            overall_avg = "—"

        # Header summary info for this exam
        _heading2(doc, f"{ca_label}")
        info_tbl = doc.add_table(rows=1, cols=5)
        info_tbl.style = "Table Grid"
        info_labels = ["Exam", "Students", "Avg (Assignment)" if "assign" in ca_label.lower()
                       else f"Avg ({ca_label})", "Max Marks", "Questions"]
        info_vals   = [ca_label, str(num_students), overall_avg, str(int(total_max)) if total_max else "—", str(len(qp))]
        for ci, (lbl, val) in enumerate(zip(info_labels, info_vals)):
            cell = info_tbl.rows[0].cells[ci]
            _set_cell_bg(cell, _NAVY)
            _set_cell_margins(cell)
            p1 = cell.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p1, lbl, bold=True, size=8, color=_WHITE)
            p2 = cell.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p2, val, bold=True, size=10, color=_WHITE)

        doc.add_paragraph()

        if has_real_marks:
            has_any_marks_section = True
            mk_rows = []

            # Avg row first
            avg_row = ["—", "Class Average"] + q_avgs + [overall_avg]
            mk_rows.append(("avg", avg_row))

            for s in students_list:
                prn  = s["prn"]
                name = s.get("name","")
                mks  = marks_data.get(prn) or {}
                row  = [prn, name]
                tot  = 0.0
                for q in qp:
                    v = float(mks.get(q.get("q_no"), 0) or 0)
                    row.append(str(v) if v else "")
                    tot += v
                row.append(f"{tot:.1f}" if tot else "")
                mk_rows.append(("student", row))

            # Also include students from marks_data not in students_list
            known_prns = {s["prn"] for s in students_list}
            for prn, mks in marks_data.items():
                if prn not in known_prns and isinstance(mks, dict):
                    name = student_map.get(prn) or "—"
                    row  = [prn, name]
                    tot  = 0.0
                    for q in qp:
                        v = float(mks.get(q.get("q_no"), 0) or 0)
                        row.append(str(v) if v else "")
                        tot += v
                    row.append(f"{tot:.1f}" if tot else "")
                    mk_rows.append(("student", row))

            n = len(q_nos)
            col_w = [2.5, 4.5] + [round(7.5/max(n,1), 2)]*n + [1.5]
            # Build table manually to colour avg row differently
            num_cols = 2 + n + 1
            tbl = doc.add_table(rows=1 + len(mk_rows), cols=num_cols)
            tbl.style = "Table Grid"
            # Header
            hdr_row = tbl.rows[0]
            headers = ["PRN", "Name"] + [f"{q_nos[i]}\n(/{int(q_max_marks[i]) if q_max_marks[i] else '?'})" for i in range(n)] + ["Total"]
            for ci2, hdr_txt in enumerate(headers):
                cell = hdr_row.cells[ci2]
                cell.width = Cm(col_w[ci2])
                _set_cell_bg(cell, _NAVY)
                _set_cell_margins(cell)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _run(p, hdr_txt, bold=True, size=8, color=_WHITE)
            # Data rows
            for ri, (row_type, row_data) in enumerate(mk_rows):
                bg = _GREEN if row_type == "avg" else (_WHITE if ri % 2 == 1 else _LIGHT)
                tr = tbl.rows[ri + 1]
                for ci2, val in enumerate(row_data):
                    cell = tr.cells[ci2]
                    cell.width = Cm(col_w[ci2])
                    _set_cell_bg(cell, bg)
                    _set_cell_margins(cell)
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci2 >= 2 else WD_ALIGN_PARAGRAPH.LEFT
                    is_avg_row = (row_type == "avg")
                    _run(p, str(val or ""), bold=is_avg_row, size=8,
                         color=(0,100,0) if is_avg_row else (0,0,0))
        else:
            _add_para(doc, f"[Marks not yet entered for {ca_label}. Upload via the marks upload.]",
                      color=(136, 136, 136))
        doc.add_paragraph()

    if not ca_sheets_all:
        _add_para(doc, "[No evaluation components found. Generate evaluation plan first.]",
                  color=(136, 136, 136))

    # ── 9. Slow & Advanced Learners ───────────────────────────────────────────
    _section_title(doc, 9, "List of Slow and Advanced learners and the action plans")

    _heading2(doc, "Advance Learners")
    advanced_list = data.get("advanced_learners_parsed") or []
    if advanced_list:
        _make_table(doc,
                    ["Sr.No", "PRN", "Name", "Marks Obtained"],
                    [[str(i+1), s.get("prn",""), s.get("name",""), s.get("marks","")] for i, s in enumerate(advanced_list)],
                    [1.0, 2.5, 9.0, 4.0])
    elif data.get("advanced_learners"):
        _add_para(doc, data["advanced_learners"])
    else:
        _add_para(doc, "[Advanced learner list will appear once CA marks are entered.]", color=(136, 136, 136))

    _heading2(doc, "Slow Learners")
    slow_list = data.get("slow_learners_parsed") or []
    if slow_list:
        _make_table(doc,
                    ["Sr.No", "PRN", "Name", "Marks Obtained"],
                    [[str(i+1), s.get("prn",""), s.get("name",""), s.get("marks","")] for i, s in enumerate(slow_list)],
                    [1.0, 2.5, 9.0, 4.0])
    elif data.get("slow_learners"):
        _add_para(doc, data["slow_learners"])
    else:
        _add_para(doc, "[Slow learner list will appear once CA marks are entered.]", color=(136, 136, 136))

    # ── 9. CO Attainment (internal) ───────────────────────────────────────────
    _section_title(doc, 10, "CO Attainment of Internal Evaluation")
    co_attainment = data.get("co_attainment") or {}
    if co_attainment:
        ca_rows = []
        for co, val in co_attainment.items():
            pct = float(val) if isinstance(val, (int, float)) else 0.0
            level = 3 if pct >= 70 else (2 if pct >= 40 else 1)
            ca_rows.append([co, f"{pct:.1f}%", str(level)])
        _make_table(doc, ["CO", "Attainment (%)", "Level"], ca_rows, [2.0, 5.0, 9.5])
    else:
        _add_para(doc, "[CO attainment will appear here once marks are entered in the Master Attainment File.]",
                  color=(136, 136, 136))

    # ── 10. Activity Reports ──────────────────────────────────────────────────
    _section_title(doc, 11, "Reports of activities planned and conducted")
    activity_reports = data.get("activity_reports") or ""
    if activity_reports.strip():
        # Try to parse as JSON structured reports
        try:
            reports = json.loads(activity_reports)
            if isinstance(reports, list):
                _heading2(doc, "Best Practice and Innovative Activities-")
                for i, rpt in enumerate(reports):
                    if isinstance(rpt, dict):
                        _add_para(doc, f"{i+1}. {rpt.get('title','')}", bold=True, size=11)
                        for k, v in rpt.items():
                            if k != "title" and v:
                                p = doc.add_paragraph()
                                p.paragraph_format.space_before = Pt(2)
                                p.paragraph_format.space_after  = Pt(2)
                                r1 = p.add_run(f"{k.replace('_',' ').title()}: ")
                                r1.bold = True
                                r1.font.size = Pt(10)
                                p.add_run(str(v)).font.size = Pt(10)
                    else:
                        _add_para(doc, f"{i+1}. {rpt}")
            else:
                raise ValueError("not a list")
        except (json.JSONDecodeError, ValueError):
            # Fall back to free text — numbered lines
            lines = [l.strip() for l in activity_reports.split("\n") if l.strip()]
            _heading2(doc, "Best Practice and Innovative Activities-")
            for i, line in enumerate(lines):
                _add_para(doc, f"{i+1}.\t{line}", space_before=2, space_after=2)
    else:
        _add_para(doc, "[Activity reports not yet entered. Add them in the Course File section.]",
                  color=(136, 136, 136))

    # ── 11. Learning Material ─────────────────────────────────────────────────
    _section_title(doc, 12, "Learning Material")
    if data.get("learning_material_links"):
        _heading2(doc, "LMS / Online Resources")
        for link in data["learning_material_links"].split("\n"):
            link = link.strip()
            if link:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after  = Pt(2)
                run = p.add_run(link)
                run.font.color.rgb = _rgb((5, 99, 193))
                run.underline = True
                run.font.size = Pt(10)

    # Also render study materials tables if available
    materials = data.get("study_materials") or {}
    tb2 = materials.get("textbooks") or []
    rb2 = materials.get("reference_books") or materials.get("references") or []
    wl2 = materials.get("web_links") or materials.get("web") or []
    j2  = materials.get("journals") or []
    m2  = materials.get("moocs") or []

    if tb2:
        _heading2(doc, "Textbooks")
        _make_table(doc, ["Book", "Author", "Publisher"],
                    [[b.get("title","") if isinstance(b,dict) else str(b),
                      b.get("author","") if isinstance(b,dict) else "",
                      b.get("publisher","") if isinstance(b,dict) else ""] for b in tb2],
                    [7.0, 4.0, 5.0])
    if rb2:
        _heading2(doc, "Reference Books")
        _make_table(doc, ["Book", "Author", "Publisher"],
                    [[b.get("title","") if isinstance(b,dict) else str(b),
                      b.get("author","") if isinstance(b,dict) else "",
                      b.get("publisher","") if isinstance(b,dict) else ""] for b in rb2],
                    [7.0, 4.0, 5.0])
    if wl2:
        _heading2(doc, "Web Links / NPTEL / SWAYAM")
        _make_table(doc, ["Sr. No.", "Web Link", "Module"],
                    [[str(i+1),
                      w.get("title",w.get("url","") if isinstance(w,dict) else str(w)),
                      w.get("unit",w.get("module","")) if isinstance(w,dict) else ""] for i,w in enumerate(wl2)],
                    [1.0, 10.0, 5.0])
    if j2:
        _heading2(doc, "Journals")
        _make_table(doc, ["Sr. No.", "Journal"],
                    [[str(i+1), j.get("title","") if isinstance(j,dict) else str(j)] for i,j in enumerate(j2)],
                    [1.0, 15.0])
    if m2:
        _heading2(doc, "MOOC Courses")
        _make_table(doc, ["S.No.", "MOOC Course Link", "Course conducted by", "Course Duration", "Certificate (Y/N)"],
                    [[str(i+1),
                      m.get("title",m.get("url","") if isinstance(m,dict) else str(m)),
                      m.get("platform",m.get("conducted_by","")) if isinstance(m,dict) else "",
                      m.get("duration","") if isinstance(m,dict) else "",
                      m.get("certificate","Y") if isinstance(m,dict) else "Y"] for i,m in enumerate(m2)],
                    [1.0, 6.0, 3.0, 2.5, 2.0])

    if not data.get("learning_material_links") and not any([tb2, rb2, wl2, j2, m2]):
        _add_para(doc, "[Learning material links not yet entered. Add them in the Course File section.]",
                  color=(136, 136, 136))

    # ── 12. Question Bank ─────────────────────────────────────────────────────
    _section_title(doc, 13, "Question Bank")
    questions = data.get("questions") or []
    if questions:
        from collections import defaultdict as _dd2
        by_unit = _dd2(list)
        for q in questions:
            unit_key = q.get("unit_no", q.get("unit", "General"))
            by_unit[str(unit_key or "General")].append(q)

        for unit_label in sorted(by_unit.keys()):
            unit_qs = by_unit[unit_label]
            _heading2(doc, f"Unit - {unit_label}")
            for i, q in enumerate(unit_qs):
                co_tag = f"  [{q.get('co_id','')}]" if q.get("co_id") else ""
                _add_para(doc, f"{i+1}. {q.get('question_text','')}{co_tag}", space_before=2, space_after=1)
    else:
        _add_para(doc, "[Question bank is empty. Use the Question Bank page to generate questions.]",
                  color=(136, 136, 136))

    # ── 13. Attendance ────────────────────────────────────────────────────────
    _section_title(doc, 14, "Compiled Attendance")
    if data.get("attendance_links"):
        for link in data["attendance_links"].split("\n"):
            link = link.strip()
            if link:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after  = Pt(2)
                run = p.add_run(link)
                run.font.color.rgb = _rgb((5, 99, 193))
                run.underline = True
                run.font.size = Pt(10)
    else:
        _add_para(doc, "[Attendance links not yet entered. Add them in the Course File section.]",
                  color=(136, 136, 136))

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Service class
# ─────────────────────────────────────────────────────────────────────────────

class CourseFileService:
    def __init__(self, db: AsyncSession):
        self.db = db

    @staticmethod
    def get_filepath(course_id: int) -> str:
        storage = get_storage()
        p = storage.get_path(_CATEGORY, f"course_file_{course_id}.docx")
        return str(p) if p else str(
            get_storage()._dir(_CATEGORY) / f"course_file_{course_id}.docx"
        )

    async def _get_students(self, course_id: int):
        try:
            result = await self.db.execute(
                text("SELECT prn, name, section FROM students WHERE course_id=:cid ORDER BY section, name"),
                {"cid": course_id}
            )
            return [{"prn": r[0], "name": r[1], "section": r[2]} for r in result.fetchall()]
        except Exception:
            return []

    async def _get_session_rows(self, course_id: int):
        from backend.database.models import SessionPlanRow
        result = await self.db.execute(
            select(SessionPlanRow).where(SessionPlanRow.course_id == course_id)
        )
        row = result.scalar_one_or_none()
        return row.rows if row else []

    async def _get_eval_rows(self, course_id: int):
        from backend.database.models import EvalPlanRow
        result = await self.db.execute(
            select(EvalPlanRow).where(EvalPlanRow.course_id == course_id)
        )
        row = result.scalar_one_or_none()
        return row.rows if row else []

    async def _get_ca_sheets(self, course_id: int):
        from backend.database.models import CASheet
        result = await self.db.execute(
            select(CASheet).where(CASheet.course_id == course_id)
        )
        sheets = result.scalars().all()
        return [{"ca_label": s.ca_label, "qp": s.qp, "marks": s.marks} for s in sheets]

    async def _get_questions(self, course_id: int):
        from backend.database.models import Question
        result = await self.db.execute(
            select(Question).where(Question.course_id == course_id)
        )
        qs = result.scalars().all()
        return [{"question_text": q.question_text, "co_id": q.co_id,
                 "bloom_level": q.bloom_level, "marks": q.marks,
                 "unit_no": getattr(q, "unit_no", None)} for q in qs]

    async def _get_timetable(self) -> dict:
        """Pull the uploaded timetable from storage."""
        try:
            storage = get_storage()
            path = storage.get_path("timetables", "current_timetable.json")
            if path and Path(path).exists():
                return json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as e:
            logger.warning(f"Could not read timetable: {e}")
        return {}

    async def _get_study_materials(self, course_id: int) -> dict:
        """Try to pull study materials from the session plan materials endpoint data."""
        try:
            from backend.database.models import SessionPlanRow
            result = await self.db.execute(
                select(SessionPlanRow).where(SessionPlanRow.course_id == course_id)
            )
            sp_row = result.scalar_one_or_none()
            if not sp_row or not sp_row.rows:
                return {}

            rows = sp_row.rows
            cols = sp_row.cols or []

            textbooks, ref_books, web_links, journals, moocs = [], [], [], [], []

            for col in cols:
                label = col.get("label", "").lower()
                key   = col.get("key", "")
                if not key:
                    continue
                if any(k in label for k in ["textbook", "text book"]):
                    for r in rows:
                        v = r.get(key, "")
                        if v and v not in [t.get("title","") for t in textbooks]:
                            textbooks.append({"title": v, "author": "", "publisher": ""})
                elif any(k in label for k in ["reference", "ref book"]):
                    for r in rows:
                        v = r.get(key, "")
                        if v and v not in [t.get("title","") for t in ref_books]:
                            ref_books.append({"title": v, "author": "", "publisher": ""})
                elif any(k in label for k in ["web", "link", "nptel", "url", "online", "swayam"]):
                    for r in rows:
                        v = r.get(key, "")
                        if v and v not in [w.get("title","") for w in web_links]:
                            web_links.append({"title": v, "unit": r.get("unit", ""), "url": ""})
                elif any(k in label for k in ["journal", "paper", "research", "article"]):
                    for r in rows:
                        v = r.get(key, "")
                        if v and v not in [j.get("title","") for j in journals]:
                            journals.append({"title": v, "url": ""})
                elif any(k in label for k in ["mooc", "course", "coursera", "swayam", "edx"]):
                    for r in rows:
                        v = r.get(key, "")
                        if v and v not in [m.get("title","") for m in moocs]:
                            moocs.append({"title": v, "platform": "", "duration": "", "certificate": "Y"})

            return {
                "textbooks": textbooks,
                "reference_books": ref_books,
                "web_links": web_links,
                "journals": journals,
                "moocs": moocs,
            }
        except Exception as e:
            logger.warning(f"Could not read study materials: {e}")
            return {}

    async def _get_co_attainment(self, course_id, students, ca_sheets, cos):
        attainment = {}
        for co in cos:
            cid = co["co_id"]
            total_pct, count = 0, 0
            for sheet in ca_sheets:
                qp = [q for q in (sheet.get("qp") or []) if q.get("co_id") == cid]
                max_marks = sum(float(q.get("marks", 0)) for q in qp)
                if not max_marks or not students:
                    continue
                marks_data = sheet.get("marks") or {}
                # Skip if no actual marks entered
                has_real = any(
                    any(float(v or 0) > 0 for v in (mks or {}).values())
                    for mks in marks_data.values() if isinstance(mks, dict)
                )
                if not has_real:
                    continue
                passed = sum(
                    1 for s in students
                    if sum(float((marks_data.get(s["prn"]) or {}).get(q.get("q_no"), 0))
                           for q in qp) / max_marks * 100 >= 60
                )
                total_pct += (passed / len(students)) * 100
                count += 1
            attainment[cid] = round(total_pct / count, 1) if count else None  # None = no data
        # Return only COs with real data
        return {k: v for k, v in attainment.items() if v is not None}

    async def _get_slow_advanced(self, course_id, students, ca_sheets, cos):
        if not students or not ca_sheets:
            return [], []
        totals = {s["prn"]: 0.0 for s in students}
        maxes  = {s["prn"]: 0.0 for s in students}
        has_any_marks = False
        for sheet in ca_sheets:
            qp = sheet.get("qp") or []
            marks_data = sheet.get("marks") or {}
            total_marks = sum(float(q.get("marks", 0)) for q in qp)
            if not total_marks:
                continue
            for s in students:
                obtained = sum(float((marks_data.get(s["prn"]) or {}).get(q.get("q_no"), 0))
                               for q in qp)
                if obtained > 0:
                    has_any_marks = True
                totals[s["prn"]] += obtained
                maxes[s["prn"]]  += total_marks

        if not has_any_marks:
            return [], []  # No marks entered — don't classify anyone

        scored = []
        for s in students:
            mx  = maxes.get(s["prn"], 0)
            tot = totals.get(s["prn"], 0)
            pct = (tot / mx * 100) if mx else 0
            scored.append({"prn": s["prn"], "name": s["name"],
                           "marks": f"{tot:.1f}/{mx:.0f}", "pct": pct})
        scored.sort(key=lambda x: x["pct"])
        return [s for s in scored if s["pct"] < 40], [s for s in scored if s["pct"] >= 75]

    async def _get_extra(self, course_id: int):
        from backend.database.models import CourseFileExtra
        result = await self.db.execute(
            select(CourseFileExtra).where(CourseFileExtra.course_id == course_id)
        )
        extra = result.scalar_one_or_none()
        return extra.to_dict() if extra else {}

    async def _get_attachments(self, course_id: int):
        from backend.database.models import CourseFileAttachment
        result = await self.db.execute(
            select(CourseFileAttachment)
            .where(CourseFileAttachment.course_id == course_id)
            .order_by(CourseFileAttachment.section_no, CourseFileAttachment.uploaded_at)
        )
        return [a.to_dict() for a in result.scalars().all()]

    def _extract_syllabus_from_session(self, session_rows):
        units = {}
        for row in session_rows:
            unit_no = _g(row, "unit", "unit_no", "unitNo", "unit_number")
            topic   = _g(row, "topic", "points_to_cover", "pointsToCover", "content")
            if not unit_no:
                continue
            key = str(unit_no)
            if key not in units:
                units[key] = {"unit_number": unit_no,
                              "unit_title": row.get("unit_title", f"Unit {unit_no}"),
                              "topics": []}
            if topic and topic not in units[key]["topics"]:
                units[key]["topics"].append(topic)
        return list(units.values())

    def _extract_tutorial_questions(self, questions, max_per_co=5):
        if not questions:
            return []
        by_co = {}
        for q in questions:
            co = q.get("co_id") or "General"
            by_co.setdefault(co, []).append(q)
        result = []
        for co, qs in by_co.items():
            result.extend(qs[:max_per_co])
        return result



    async def _ai_bloom_map(self, ca_sheets: list) -> dict:
        """
        For each CA sheet, AI-classify any questions missing a bloom level.
        Returns {ca_label: {q_no: bloom_level_str}}.
        Falls back gracefully if LLM unavailable.
        """
        import json as _json
        from backend.core.llm import get_llm_response

        BLOOM_NAMES = ["Remember", "Understand", "Apply", "Analyse", "Evaluate", "Create"]
        result = {}

        for ca in (ca_sheets or []):
            qp = ca.get("qp") or []
            ca_label = ca.get("ca_label", "")
            if not qp:
                continue

            # Collect questions with missing bloom level
            needs_ai = [
                q for q in qp
                if not q.get("bloom_level") or q.get("bloom_level") == 0
            ]
            if not needs_ai:
                continue

            lines = []
            for q in needs_ai:
                q_no = str(q.get("q_no", "?"))
                text = str(q.get("question_text", ""))[:120]
                lines.append("  Q" + q_no + ". " + text)
            q_list = "\n".join(lines)

            prompt = (
                "You are an expert in Bloom's Taxonomy for engineering education. "
                "Classify each question below into exactly ONE Bloom level from: "
                "Remember, Understand, Apply, Analyse, Evaluate, Create. "
                "Return a JSON object mapping Q-number (e.g. \"Q1\") to Bloom level string. "
                "Respond ONLY with the JSON, no markdown fences, no explanation.\n\n"
                "Questions:\n" + q_list
            )

            try:
                raw = await get_llm_response(prompt)
                raw = raw.strip()
                if raw.startswith("```"):
                    raw = raw.split("\n", 1)[-1].rsplit("```", 1)[0].strip()
                mapping = _json.loads(raw)
                clean = {}
                for k, v in mapping.items():
                    v_str = str(v).strip()
                    matched = next(
                        (b for b in BLOOM_NAMES if b.lower() in v_str.lower()), None
                    )
                    if matched:
                        key = str(k).replace("Q", "").strip()
                        clean[key] = matched
                result[ca_label] = clean
            except Exception as e:
                logger.warning("AI bloom mapping failed for %s: %s", ca_label, e)

        return result

    async def generate(self, course_id: int) -> dict:
        course_svc = CourseService(self.db)
        course     = await course_svc.get_course(course_id)
        students   = await self._get_students(course_id)
        session_rows = await self._get_session_rows(course_id)
        eval_rows    = await self._get_eval_rows(course_id)
        ca_sheets    = await self._get_ca_sheets(course_id)
        questions    = await self._get_questions(course_id)
        extra        = await self._get_extra(course_id)
        timetable    = await self._get_timetable()
        study_mats   = await self._get_study_materials(course_id)
        cos          = course.cos

        co_attainment              = await self._get_co_attainment(course_id, students, ca_sheets, cos)
        slow_list, advanced_list   = await self._get_slow_advanced(course_id, students, ca_sheets, cos)
        syllabus_units             = self._extract_syllabus_from_session(session_rows)
        tutorial_qs                = self._extract_tutorial_questions(questions)

        data = {
            "institution_name":    extra.get("institution_name", "Symbiosis Institute of Technology"),
            "institution_address": extra.get("institution_address", "SIU Pune 412115, Maharashtra, India"),
            "course_name":    course.course_name,
            "course_code":    course.course_code,
            "department":     course.department,
            "faculty_name":   course.faculty_name,
            "semester":       course.semester,
            "academic_year":  course.academic_year,
            "credits":        course.credits,
            "batch":          extra.get("batch", ""),
            "cos":            cos,
            "pos":            course.pos,
            "co_po_matrix":   course.co_po_matrix,
            "co_po_justification": extra.get("co_po_justification", ""),
            "vision_text":    extra.get("vision_text", ""),
            "mission_text":   extra.get("mission_text", ""),
            "syllabus_units": syllabus_units,
            "prev_co_attainment": extra.get("prev_co_attainment", ""),
            "action_plan":    extra.get("action_plan", ""),
            "session_rows":   session_rows,
            "tutorial_questions": tutorial_qs,
            "eval_rows":      eval_rows,
            "ca_sheets":      ca_sheets,
            "student_map":    {s["prn"]: s["name"] for s in students},
            "slow_learners":  extra.get("slow_learners", ""),
            "advanced_learners": extra.get("advanced_learners", ""),
            "slow_learners_parsed":     slow_list,
            "advanced_learners_parsed": advanced_list,
            "students":       students,
            "timetable":      timetable,
            "study_materials": study_mats,
            "attachments":    await self._get_attachments(course_id),
            "co_attainment":  co_attainment,
            "activity_reports": extra.get("activity_reports", ""),
            "learning_material_links": extra.get("learning_material_links", ""),
            "questions":      questions,
            "attendance_links": extra.get("attendance_links", ""),
        }

        # ── AI Bloom-level mapping for question papers ─────────────────────
        # Auto-classify any questions missing a bloom level (no button needed)
        data["bloom_ai_map"] = await self._ai_bloom_map(ca_sheets)

        docx_bytes = _build_docx(data)

        _storage  = get_storage()
        _filename = f"course_file_{course_id}.docx"
        _storage.save(_CATEGORY, _filename, docx_bytes)

        filepath = str(_storage.get_path(_CATEGORY, _filename))
        logger.info(f"Course file saved -> {filepath}")

        return {
            "course_id":   course_id,
            "course_name": course.course_name,
            "filename":    _filename,
            "download_url": f"/course-file/download/{course_id}",
            "sections_with_data": {
                "session_plan":     len(session_rows) > 0,
                "evaluation_plan":  len(eval_rows) > 0,
                "ca_marks":         len(ca_sheets) > 0,
                "question_bank":    len(questions) > 0,
                "vision_mission":   bool(extra.get("vision_text")),
                "timetable":        bool(timetable),
                "study_materials":  bool(any(study_mats.values())),
            },
        }
