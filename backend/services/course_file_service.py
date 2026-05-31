"""
backend/services/course_file_service.py

Generates the complete OBE Course File (.docx) with all 13 sections
using python-docx only — no Node.js dependency.
"""

import io
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
        p.runs  # ensure para exists before page break
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
    """
    col_widths_cm: list of column widths in cm.
    headers: list of header strings.
    rows: list of lists (str values).
    """
    num_cols = len(headers)
    tbl = doc.add_table(rows=1 + len(rows), cols=num_cols)
    tbl.style = "Table Grid"

    # Header row
    hdr_row = tbl.rows[0]
    for i, hdr in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.width = Cm(col_widths_cm[i])
        _set_cell_bg(cell, _NAVY)
        _set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, hdr, bold=True, size=9, color=_WHITE)

    # Data rows
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


# ── Main document builder ─────────────────────────────────────────────────────

def _build_docx(data: dict) -> bytes:
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.page_width  = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin   = section.right_margin  = Cm(2.5)
        section.top_margin    = section.bottom_margin = Cm(2)

    # ── Cover ──────────────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, f"Department of {data.get('department','')}", bold=True, size=13)

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
    # Use DB POs if they have real statement text; otherwise fall back to the
    # standard NBA 12 POs for engineering programmes.
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
    # Check if DB POs have actual statement text (not just IDs)
    db_pos_have_text = any(
        p.get("statement", p.get("description", "")).strip()
        for p in pos
    )
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
    _make_table(doc,
                ["", "Program Specific Outcomes"],
                [[pid, ptext] for pid, ptext in psos],
                [1.5, 14.5])

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
                  color=(136,136,136))
    _heading2(doc, "Personal Timetable")
    timetable_attachments = [
        a for a in (data.get("attachments") or [])
        if a.get("section_no") == 3
    ]
    if timetable_attachments:
        for a in timetable_attachments:
            _add_para(doc, f"Attached: {a['label']}  ({a['filename']})", size=10)
    else:
        _add_para(doc, "[Faculty timetable — upload via Dept. Uploads tab, tagged to Section 3.]",
                  color=(136,136,136))

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
                _add_para(doc, f"Section {sec_label}", bold=True, size=11,
                          space_before=6, space_after=2)
            _make_table(
                doc,
                ["Sr. No", "PRN", "Name"],
                [[str(i+1), s.get("prn",""), s.get("name","")] for i, s in enumerate(sec_students)],
                [1.2, 3.5, 11.3],
            )
    else:
        _add_para(doc, "[Student list not available. Add students via the Students page.]",
                  color=(136,136,136))

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
    if co_po_matrix and pos:
        po_ids = [p.get("po_id","") for p in pos]
        matrix_rows = []
        for co in cos:
            mapping = co_po_matrix.get(co.get("co_id","")) or {}
            matrix_rows.append([co.get("co_id","")] + [mapping.get(pid, "-") for pid in po_ids])
        n = len(po_ids)
        col_w = [1.5] + [round(14.5/max(n,1), 2)] * n
        _make_table(doc, ["CO"] + po_ids, matrix_rows, col_w)
    else:
        _add_para(doc, "[CO-PO mapping not yet configured.]", color=(136,136,136))

    # ── 5. Previous CO Attainment ─────────────────────────────────────────────
    _section_title(doc, 5, "CO Attainment from previous academic year and the action plan")
    if data.get("prev_co_attainment"):
        _add_para(doc, data["prev_co_attainment"])
    else:
        _add_para(doc, "[Previous year CO attainment data not yet entered.]", color=(136,136,136))
    _heading2(doc, "Action Plan")
    if data.get("action_plan"):
        _add_para(doc, data["action_plan"])
    else:
        _add_para(doc, "[Action plan not yet entered.]", color=(136,136,136))

    # ── 6. Session Plan ───────────────────────────────────────────────────────
    _section_title(doc, 6, "Session Plan with CO mapping to each lecture")
    session_rows = data.get("session_rows") or []
    if session_rows:
        _make_table(doc,
                    ["Lect. No", "Unit No", "Points to Cover", "Methodology", "Type", "CO Mapped"],
                    [[r.get("lect_no", r.get("lectNo","")),
                      r.get("unit_no", r.get("unitNo","")),
                      r.get("points_to_cover", r.get("pointsToCover", r.get("topic",""))),
                      r.get("methodology",""),
                      r.get("lecture_exp_eval", r.get("type","Lecture")),
                      r.get("co", r.get("co_mapped",""))] for r in session_rows],
                    [1.5, 1.5, 8.0, 2.5, 2.0, 2.0])
    else:
        _add_para(doc, "[Session plan not yet generated. Use the Session Plan page first.]",
                  color=(136,136,136))

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
        _make_table(doc,
                    ["Sr.No", "Component", "Units/Syllabus", "CO Mapped", "Marks", "Weightage", "Tentative Date"],
                    [[r.get("sr_no", r.get("srNo","")),
                      r.get("component", r.get("comp", r.get("name",""))),
                      r.get("unit_syllabus", r.get("units","")),
                      r.get("co", r.get("co_mapped","")),
                      r.get("marks", r.get("total_marks","")),
                      r.get("weightage",""),
                      r.get("date", r.get("tentative_date",""))] for r in eval_rows],
                    [1.0, 3.0, 4.5, 2.0, 1.2, 1.8, 3.0])
    else:
        _add_para(doc, "[Evaluation plan not yet generated. Use the Evaluation Plan page first.]",
                  color=(136,136,136))

    for ca in (data.get("ca_sheets") or []):
        qp = ca.get("qp") or []
        if not qp:
            continue
        _heading2(doc, f"{ca.get('ca_label','')} — Question Paper")
        _make_table(doc,
                    ["Q.No", "Question", "Marks", "CO", "BL"],
                    [[q.get("q_no",""), q.get("question_text",""), q.get("marks",""),
                      q.get("co_id",""), q.get("bloom_level","")] for q in qp],
                    [1.0, 10.0, 1.2, 1.8, 1.5])

        marks_data = ca.get("marks") or {}
        if marks_data:
            _heading2(doc, f"{ca.get('ca_label','')} — Marks")
            q_nos = [q.get("q_no","") for q in qp]
            student_map = data.get("student_map") or {}
            mk_rows = []
            for prn, mks in marks_data.items():
                row = [prn, student_map.get(prn,"")]
                tot = 0.0
                for q in qp:
                    v = float((mks or {}).get(q.get("q_no"), 0) or 0)
                    row.append(v or "")
                    tot += v
                row.append(tot or "")
                mk_rows.append(row)
            n = len(q_nos)
            col_w = [2.0, 3.5] + [round(8.0/max(n,1),2)]*n + [1.5]
            _make_table(doc, ["PRN", "Name"] + q_nos + ["Total"], mk_rows, col_w)

    # ── 8. Slow & Advanced Learners ───────────────────────────────────────────
    _section_title(doc, 8, "List of Slow and Advanced learners and the action plans")
    _heading2(doc, "Slow Learners")
    slow = data.get("slow_learners_parsed") or []
    if slow:
        _make_table(doc,
                    ["Sr.No", "PRN", "Name", "Marks Obtained"],
                    [[str(i+1), s.get("prn",""), s.get("name",""), s.get("marks","")] for i, s in enumerate(slow)],
                    [1.0, 2.5, 9.0, 4.0])
    elif data.get("slow_learners"):
        _add_para(doc, data["slow_learners"])
    else:
        _add_para(doc, "[Slow learner list not yet entered. Complete CA marks to auto-generate.]",
                  color=(136,136,136))

    _heading2(doc, "Advanced Learners")
    if data.get("advanced_learners"):
        _add_para(doc, data["advanced_learners"])
    else:
        _add_para(doc, "[Advanced learner list not yet entered.]", color=(136,136,136))

    # ── 9. CO Attainment (internal) ───────────────────────────────────────────
    _section_title(doc, 9, "CO Attainment of internal evaluation")
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
                  color=(136,136,136))

    # ── 10. Activity Reports ──────────────────────────────────────────────────
    _section_title(doc, 10, "Reports of activities planned and conducted")
    if data.get("activity_reports"):
        for line in data["activity_reports"].split("\n"):
            if line.strip():
                _add_para(doc, line)
    else:
        _add_para(doc, "[Activity reports not yet entered. Add them in the Course File section.]",
                  color=(136,136,136))

    # ── 11. Learning Material ─────────────────────────────────────────────────
    _section_title(doc, 11, "Learning Material")
    if data.get("learning_material_links"):
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
    else:
        _add_para(doc, "[Learning material links not yet entered. Add them in the Course File section.]",
                  color=(136,136,136))

    # ── 12. Question Bank ─────────────────────────────────────────────────────
    _section_title(doc, 12, "Question Bank")
    questions = data.get("questions") or []
    if questions:
        for co in cos:
            co_qs = [q for q in questions if q.get("co_id") == co.get("co_id")]
            if not co_qs:
                continue
            _heading2(doc, co.get("co_id",""))
            for i, q in enumerate(co_qs):
                _add_para(doc, f"{i+1}. {q.get('question_text','')}", space_before=2, space_after=1)
        unmapped = [q for q in questions if not q.get("co_id")]
        if unmapped:
            _heading2(doc, "General")
            for i, q in enumerate(unmapped):
                _add_para(doc, f"{i+1}. {q.get('question_text','')}", space_before=2, space_after=1)
    else:
        _add_para(doc, "[Question bank is empty. Use the Question Bank page to generate questions.]",
                  color=(136,136,136))

    # ── 13. Attendance ────────────────────────────────────────────────────────
    _section_title(doc, 13, "Compiled Attendance")
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
                  color=(136,136,136))

    # ── Serialise ─────────────────────────────────────────────────────────────
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
                 "bloom_level": q.bloom_level, "marks": q.marks} for q in qs]

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
                passed = sum(
                    1 for s in students
                    if sum(float((marks_data.get(s["prn"]) or {}).get(q.get("q_no"), 0))
                           for q in qp) / max_marks * 100 >= 60
                )
                total_pct += (passed / len(students)) * 100
                count += 1
            attainment[cid] = round(total_pct / count, 1) if count else 0.0
        return attainment

    async def _get_slow_advanced(self, course_id, students, ca_sheets, cos):
        if not students or not ca_sheets:
            return [], []
        totals = {s["prn"]: 0.0 for s in students}
        maxes  = {s["prn"]: 0.0 for s in students}
        for sheet in ca_sheets:
            qp = sheet.get("qp") or []
            marks_data = sheet.get("marks") or {}
            total_marks = sum(float(q.get("marks", 0)) for q in qp)
            if not total_marks:
                continue
            for s in students:
                obtained = sum(float((marks_data.get(s["prn"]) or {}).get(q.get("q_no"), 0))
                               for q in qp)
                totals[s["prn"]] += obtained
                maxes[s["prn"]]  += total_marks
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
            unit_no = row.get("unit_no", row.get("unitNo", ""))
            topic   = row.get("points_to_cover", row.get("pointsToCover", row.get("topic", "")))
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

    async def generate(self, course_id: int) -> dict:
        course_svc = CourseService(self.db)
        course     = await course_svc.get_course(course_id)
        students   = await self._get_students(course_id)
        session_rows = await self._get_session_rows(course_id)
        eval_rows    = await self._get_eval_rows(course_id)
        ca_sheets    = await self._get_ca_sheets(course_id)
        questions    = await self._get_questions(course_id)
        extra        = await self._get_extra(course_id)
        cos          = course.cos

        co_attainment          = await self._get_co_attainment(course_id, students, ca_sheets, cos)
        slow_list, advanced_list = await self._get_slow_advanced(course_id, students, ca_sheets, cos)
        syllabus_units         = self._extract_syllabus_from_session(session_rows)
        tutorial_qs            = self._extract_tutorial_questions(questions)

        data = {
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
            "slow_learners":  extra.get("slow_learners", "") or
                              "\n".join(f"{s['prn']} — {s['name']} ({s['marks']})" for s in slow_list),
            "advanced_learners": extra.get("advanced_learners", "") or
                                 "\n".join(f"{s['prn']} — {s['name']} ({s['marks']})" for s in advanced_list),
            "slow_learners_parsed": slow_list,
            "students":       students,
            "attachments":    await self._get_attachments(course_id),
            "co_attainment":  co_attainment,
            "activity_reports": extra.get("activity_reports", ""),
            "learning_material_links": extra.get("learning_material_links", ""),
            "questions":      questions,
            "attendance_links": extra.get("attendance_links", ""),
        }

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
            },
        }
