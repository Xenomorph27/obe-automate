# backend/services/nba_report_service.py
"""
NBAReportService
----------------
Day 6 — CO-PO Gap Analysis + NBA/NAAC-format PDF report using ReportLab.

What it does:
1. gap_analysis()     — computes CO/PO attainment gaps (actual vs target)
2. generate_pdf()     — builds a print-ready PDF with:
     • Cover page (course metadata)
     • Section A: NBA format — CO attainment table + gap analysis + bar chart
     • Section B: NAAC format — PO attainment table + level descriptors
     • Section C: AI-generated recommendations (via LLM fallback chain)
     • Section D: CO-PO correlation matrix

Output folder : generated_docs/nba_reports/
Filename      : nba_report_<course_id>.pdf
"""

import os
from pathlib import Path
from collections import defaultdict

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    HRFlowable,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics import renderPDF
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.exceptions import OBEException
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger
from backend.services.attainment_service import AttainmentService
from backend.services.course_service import CourseService

logger = get_logger(__name__)

from backend.core.storage import get_storage
_CATEGORY = "nba_reports"

# ── Colour palette ──────────────────────────────────────────────────────
C_NAVY    = colors.HexColor("#1F497D")
C_BLUE    = colors.HexColor("#2E75B6")
C_GREEN   = colors.HexColor("#70AD47")
C_ORANGE  = colors.HexColor("#C55A11")
C_PURPLE  = colors.HexColor("#7F3F98")
C_RED     = colors.HexColor("#FF0000")
C_YELLOW  = colors.HexColor("#FFD700")
C_LGRAY   = colors.HexColor("#F2F2F2")
C_DGRAY   = colors.HexColor("#595959")
C_WHITE   = colors.white
C_BLACK   = colors.black

TARGET_PCT = 60   # default attainment target


class NBAReportService:

    def __init__(self, db: AsyncSession):
        self.db = db

    # ------------------------------------------------------------------
    # Static helper
    # ------------------------------------------------------------------

    @staticmethod
    def get_filepath(course_id: int) -> str:
        storage = get_storage()
        p = storage.get_path(_CATEGORY, f"nba_report_{course_id}.pdf")
        return str(p) if p else str(get_storage()._dir(_CATEGORY) / f"nba_report_{course_id}.pdf")

    # ------------------------------------------------------------------
    # 1. Gap Analysis
    # ------------------------------------------------------------------

    async def gap_analysis(self, course_id: int) -> dict:
        """
        Returns full attainment data enriched with gap analysis fields.
        gap = target_pct - actual_pct  (positive = underperforming)
        """
        att_svc = AttainmentService(self.db)
        data = await att_svc.calculate(course_id)

        co_gaps = {}
        for co_id, co in data["co_attainment"].items():
            actual = co["attainment_percentage"]
            gap = round(TARGET_PCT - actual, 2)
            co_gaps[co_id] = {
                **co,
                "target_percentage": TARGET_PCT,
                "gap": gap,
                "status": "AT RISK" if gap > 0 else "ACHIEVED",
            }

        po_gaps = {}
        for po_id, po in data["po_attainment"].items():
            actual = po["attainment_percentage"]
            gap = round(TARGET_PCT - actual, 2)
            po_gaps[po_id] = {
                **po,
                "target_percentage": TARGET_PCT,
                "gap": gap,
                "status": "AT RISK" if gap > 0 else "ACHIEVED",
            }

        at_risk_cos = [co_id for co_id, v in co_gaps.items() if v["status"] == "AT RISK"]
        at_risk_pos = [po_id for po_id, v in po_gaps.items() if v["status"] == "AT RISK"]

        return {
            **data,
            "co_gaps": co_gaps,
            "po_gaps": po_gaps,
            "at_risk_cos": at_risk_cos,
            "at_risk_pos": at_risk_pos,
            "target_percentage": TARGET_PCT,
        }

    # ------------------------------------------------------------------
    # 2. Get AI recommendations
    # ------------------------------------------------------------------

    async def _get_recommendations(self, course: object, gap_data: dict) -> str:
        at_risk_cos = gap_data["at_risk_cos"]
        at_risk_pos = gap_data["at_risk_pos"]

        co_summary = "\n".join(
            f"  {co_id}: actual={v['attainment_percentage']}%, gap={v['gap']}%, status={v['status']}"
            for co_id, v in gap_data["co_gaps"].items()
        )
        po_summary = "\n".join(
            f"  {po_id}: actual={v['attainment_percentage']}%, gap={v['gap']}%, status={v['status']}"
            for po_id, v in gap_data["po_gaps"].items()
        )

        prompt = f"""You are an expert OBE academic consultant helping an engineering college improve NBA/NAAC accreditation scores.

Course: {course.course_name} ({course.course_code})
Department: {course.department}
Total Students: {gap_data['total_students']}
Overall CO Attainment: {gap_data['overall_co_attainment']}%
Target Attainment: {TARGET_PCT}%

CO Attainment Summary:
{co_summary}

PO Attainment Summary:
{po_summary}

At-Risk COs (below target): {at_risk_cos if at_risk_cos else 'None — all COs achieved target'}
At-Risk POs (below target): {at_risk_pos if at_risk_pos else 'None — all POs achieved target'}

Task: Write a concise, actionable RECOMMENDATIONS section for the NBA/NAAC report.

Structure your response EXACTLY as follows (plain text, no markdown, no bullet symbols):

OVERALL ASSESSMENT
[2-3 sentences on the overall attainment performance]

STRENGTHS
[2-3 specific strengths observed from the data]

AREAS FOR IMPROVEMENT
[For each at-risk CO/PO, give one specific, actionable recommendation]

ACTION PLAN
[3-5 concrete steps faculty should take next semester to improve attainment]

CONCLUSION
[1-2 sentences suitable for an accreditation report conclusion]

Keep the tone formal and suitable for submission to NBA/NAAC auditors."""

        logger.info("Calling LLM for NBA report recommendations")
        try:
            return await get_llm_response(prompt)
        except Exception as e:
            logger.warning(f"LLM recommendations failed: {e}. Using fallback text.")
            return (
                "OVERALL ASSESSMENT\n"
                f"The course achieved an overall CO attainment of {gap_data['overall_co_attainment']}% "
                f"against a target of {TARGET_PCT}%.\n\n"
                "ACTION PLAN\n"
                "1. Review teaching methods for at-risk COs.\n"
                "2. Increase formative assessments to identify struggling students early.\n"
                "3. Conduct remedial sessions for students below threshold.\n"
                "4. Review and update CO-PO mapping for next academic year.\n"
                "5. Collect student feedback to identify pedagogical gaps."
            )

    # ------------------------------------------------------------------
    # 3. Generate PDF
    # ------------------------------------------------------------------

    async def generate_pdf(self, course_id: int) -> dict:
        gap_data = await self.gap_analysis(course_id)

        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)

        recommendations = await self._get_recommendations(course, gap_data)

        _storage = get_storage()
        _filename = f"nba_report_{course_id}.pdf"
        import tempfile as _tmp
        from pathlib import Path as _Path
        with _tmp.TemporaryDirectory() as _t:
            filepath = str(_Path(_t) / _filename)
            self._build_pdf(course, gap_data, recommendations, filepath)
            _storage.save_from_path(_CATEGORY, _filename, _Path(filepath))
        filepath = str(_storage.get_path(_CATEGORY, _filename))

        return {
            "course_id": course_id,
            "course_name": gap_data["course_name"],
            "filename": os.path.basename(filepath),
            "download_url": f"/attainment/nba-report/download/{course_id}",
            "overall_co_attainment": gap_data["overall_co_attainment"],
            "at_risk_cos": gap_data["at_risk_cos"],
            "at_risk_pos": gap_data["at_risk_pos"],
            "total_students": gap_data["total_students"],
        }

    # ------------------------------------------------------------------
    # 4. PDF builder
    # ------------------------------------------------------------------

    def _build_pdf(self, course, gap_data: dict, recommendations: str, filepath: str):
        styles = self._make_styles()

        # Page template with header/footer
        def on_page(canvas, doc):
            self._draw_header_footer(canvas, doc, course, gap_data)

        frame = Frame(1.5*cm, 2.5*cm, A4[0] - 3*cm, A4[1] - 4.5*cm, id="main")
        template = PageTemplate(id="main", frames=[frame], onPage=on_page)
        doc = BaseDocTemplate(filepath, pagesize=A4, pageTemplates=[template])

        story = []

        # ── Cover Page
        story += self._cover_page(course, gap_data, styles)
        story.append(PageBreak())

        # ── Section A: NBA Format
        story += self._section_nba(course, gap_data, styles)
        story.append(PageBreak())

        # ── Section B: NAAC Format
        story += self._section_naac(course, gap_data, styles)
        story.append(PageBreak())

        # ── Section C: Recommendations
        story += self._section_recommendations(recommendations, styles)
        story.append(PageBreak())

        # ── Section D: CO-PO Matrix
        story += self._section_matrix(course, gap_data, styles)

        doc.build(story)
        logger.info(f"NBA/NAAC PDF report saved → {filepath}")

    # ------------------------------------------------------------------
    # Header / Footer
    # ------------------------------------------------------------------

    def _draw_header_footer(self, canvas, doc, course, gap_data):
        canvas.saveState()
        w, h = A4

        # Header bar
        canvas.setFillColor(C_NAVY)
        canvas.rect(0, h - 1.8*cm, w, 1.8*cm, fill=1, stroke=0)
        canvas.setFillColor(C_WHITE)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(1.5*cm, h - 1.1*cm, "OBE AUTOMATE — NBA/NAAC ACCREDITATION REPORT")
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(w - 1.5*cm, h - 1.1*cm,
            f"{course.course_name} ({course.course_code})")

        # Footer bar
        canvas.setFillColor(C_LGRAY)
        canvas.rect(0, 0, w, 1.8*cm, fill=1, stroke=0)
        canvas.setFillColor(C_DGRAY)
        canvas.setFont("Helvetica", 7.5)
        canvas.drawString(1.5*cm, 0.8*cm,
            f"{course.department}  |  {course.academic_year}  |  "
            f"Overall CO Attainment: {gap_data['overall_co_attainment']}%  |  "
            f"Target: {TARGET_PCT}%")
        canvas.drawRightString(w - 1.5*cm, 0.8*cm, f"Page {doc.page}")

        canvas.restoreState()

    # ------------------------------------------------------------------
    # Cover Page
    # ------------------------------------------------------------------

    def _cover_page(self, course, gap_data, styles):
        items = []
        items.append(Spacer(1, 3*cm))

        # Top accent bar
        items.append(HRFlowable(width="100%", thickness=6, color=C_NAVY))
        items.append(Spacer(1, 0.5*cm))

        items.append(Paragraph("NBA / NAAC", styles["cover_sub"]))
        items.append(Paragraph("CO-PO ATTAINMENT REPORT", styles["cover_title"]))
        items.append(Spacer(1, 0.3*cm))
        items.append(HRFlowable(width="100%", thickness=2, color=C_BLUE))
        items.append(Spacer(1, 1.5*cm))

        # Course info table
        info = [
            ["Course Name",   course.course_name],
            ["Course Code",   course.course_code],
            ["Department",    course.department],
            ["Faculty",       course.faculty_name],
            ["Semester",      course.semester],
            ["Academic Year", course.academic_year],
            ["Credits",       str(course.credits)],
            ["Total Students", str(gap_data["total_students"])],
        ]
        t = Table(info, colWidths=[5*cm, 10*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (0, -1), C_NAVY),
            ("TEXTCOLOR",   (0, 0), (0, -1), C_WHITE),
            ("TEXTCOLOR",   (1, 0), (1, -1), C_BLACK),
            ("FONTNAME",    (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTNAME",    (1, 0), (1, -1), "Helvetica"),
            ("FONTSIZE",    (0, 0), (-1, -1), 10),
            ("ROWBACKGROUNDS", (1, 0), (1, -1), [C_WHITE, C_LGRAY]),
            ("GRID",        (0, 0), (-1, -1), 0.5, C_DGRAY),
            ("PADDING",     (0, 0), (-1, -1), 8),
        ]))
        items.append(t)
        items.append(Spacer(1, 1.5*cm))

        # Attainment summary pills
        overall = gap_data["overall_co_attainment"]
        at_risk_co = len(gap_data["at_risk_cos"])
        at_risk_po = len(gap_data["at_risk_pos"])
        achieved_co = len(gap_data["co_gaps"]) - at_risk_co

        summary = [
            ["Overall CO\nAttainment", "Target\nThreshold", "COs\nAchieved", "COs\nAt Risk"],
            [f"{overall}%", f"{TARGET_PCT}%", str(achieved_co), str(at_risk_co)],
        ]
        st = Table(summary, colWidths=[3.75*cm]*4)
        pill_bg = C_GREEN if overall >= TARGET_PCT else C_RED
        st.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, 0), C_NAVY),
            ("TEXTCOLOR",   (0, 0), (-1, 0), C_WHITE),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND",  (0, 1), (0, 1), pill_bg),
            ("BACKGROUND",  (1, 1), (1, 1), C_BLUE),
            ("BACKGROUND",  (2, 1), (2, 1), C_GREEN),
            ("BACKGROUND",  (3, 1), (3, 1), C_RED if at_risk_co else C_GREEN),
            ("TEXTCOLOR",   (0, 1), (-1, 1), C_WHITE),
            ("FONTNAME",    (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE",    (0, 1), (-1, 1), 16),
            ("FONTSIZE",    (0, 0), (-1, 0), 9),
            ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("ROWHEIGHTS",  (0, 0), (-1, -1), 1*cm),
            ("GRID",        (0, 0), (-1, -1), 1, C_WHITE),
            ("PADDING",     (0, 0), (-1, -1), 10),
        ]))
        items.append(st)
        items.append(Spacer(1, 2*cm))
        items.append(HRFlowable(width="100%", thickness=6, color=C_NAVY))

        return items

    # ------------------------------------------------------------------
    # Section A — NBA Format
    # ------------------------------------------------------------------

    def _section_nba(self, course, gap_data, styles):
        items = []
        items.append(Paragraph("SECTION A — NBA FORMAT", styles["section_label"]))
        items.append(Paragraph("Course Outcome (CO) Attainment & Gap Analysis",
                               styles["section_title"]))
        items.append(HRFlowable(width="100%", thickness=1.5, color=C_BLUE))
        items.append(Spacer(1, 0.4*cm))

        items.append(Paragraph(
            f"The following table presents the attainment of each Course Outcome (CO) "
            f"as per NBA criterion. The threshold for attainment is set at <b>{TARGET_PCT}%</b> "
            f"of the maximum marks per CO. Students scoring at or above this threshold "
            f"are considered to have attained the CO.",
            styles["body"]
        ))
        items.append(Spacer(1, 0.4*cm))

        # CO Attainment + Gap Table
        header = ["CO", "Statement", "Bloom's\nLevel", "Max\nMarks",
                  "Avg\nScored", "Students\n≥ Threshold", "Attainment\n%",
                  "Target\n%", "Gap", "Status"]
        rows = [header]
        for co_id, co in gap_data["co_gaps"].items():
            gap_val = co["gap"]
            rows.append([
                co_id,
                co["statement"][:45] + ("..." if len(co["statement"]) > 45 else ""),
                co["bloom_level"],
                str(co["max_marks"]),
                str(co["avg_marks_scored"]),
                f"{co['students_passed_threshold']}/{co['total_students']}",
                f"{co['attainment_percentage']}%",
                f"{TARGET_PCT}%",
                f"{gap_val:+.1f}%" if gap_val != 0 else "0%",
                co["status"],
            ])

        col_w = [1.2*cm, 4.5*cm, 1.8*cm, 1.3*cm, 1.3*cm, 2*cm, 1.8*cm, 1.5*cm, 1.3*cm, 1.8*cm]
        t = Table(rows, colWidths=col_w, repeatRows=1)

        ts = [
            ("BACKGROUND",  (0, 0), (-1, 0), C_NAVY),
            ("TEXTCOLOR",   (0, 0), (-1, 0), C_WHITE),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0), (-1, -1), 7.5),
            ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
            ("ALIGN",       (1, 1), (1, -1), "LEFT"),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",        (0, 0), (-1, -1), 0.4, C_DGRAY),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_WHITE, C_LGRAY]),
            ("PADDING",     (0, 0), (-1, -1), 5),
            ("FONTNAME",    (0, 1), (0, -1), "Helvetica-Bold"),
        ]
        # Colour status column
        for i, (co_id, co) in enumerate(gap_data["co_gaps"].items(), start=1):
            bg = C_GREEN if co["status"] == "ACHIEVED" else C_RED
            ts.append(("BACKGROUND", (9, i), (9, i), bg))
            ts.append(("TEXTCOLOR",  (9, i), (9, i), C_WHITE))
            ts.append(("FONTNAME",   (9, i), (9, i), "Helvetica-Bold"))
            # Gap column — red if positive
            if co["gap"] > 0:
                ts.append(("TEXTCOLOR", (8, i), (8, i), C_RED))
                ts.append(("FONTNAME",  (8, i), (8, i), "Helvetica-Bold"))
        t.setStyle(TableStyle(ts))
        items.append(t)
        items.append(Spacer(1, 0.5*cm))

        # Bar chart
        items.append(Paragraph("CO Attainment vs Target — Visual Overview", styles["subsection"]))
        items.append(Spacer(1, 0.2*cm))
        items.append(self._bar_chart(gap_data["co_gaps"], "CO"))
        items.append(Spacer(1, 0.4*cm))

        # NBA criterion note
        items.append(Paragraph(
            "<b>NBA Criterion Note:</b> As per NBA Self-Assessment Report (SAR) guidelines, "
            "CO attainment is computed using direct assessment methods (CIE + SEE). "
            f"COs with attainment ≥ {TARGET_PCT}% are considered attained. "
            f"Overall CO attainment for this course: <b>{gap_data['overall_co_attainment']}%</b>.",
            styles["note"]
        ))

        return items

    # ------------------------------------------------------------------
    # Section B — NAAC Format
    # ------------------------------------------------------------------

    def _section_naac(self, course, gap_data, styles):
        items = []
        items.append(Paragraph("SECTION B — NAAC FORMAT", styles["section_label"]))
        items.append(Paragraph("Program Outcome (PO) Attainment Analysis",
                               styles["section_title"]))
        items.append(HRFlowable(width="100%", thickness=1.5, color=C_PURPLE))
        items.append(Spacer(1, 0.4*cm))

        items.append(Paragraph(
            "Program Outcome attainment is computed by mapping CO attainment to POs "
            "via the CO-PO correlation matrix. The weighted average method is used "
            "where correlation strength (1=Low, 2=Medium, 3=High) serves as the weight.",
            styles["body"]
        ))
        items.append(Spacer(1, 0.4*cm))

        # PO Attainment Table
        header = ["PO", "Statement", "Attainment %", "Target %", "Gap", "Level", "Status"]
        rows = [header]
        for po_id, po in gap_data["po_gaps"].items():
            gap_val = po["gap"]
            rows.append([
                po_id,
                po["statement"][:55] + ("..." if len(po["statement"]) > 55 else ""),
                f"{po['attainment_percentage']}%",
                f"{TARGET_PCT}%",
                f"{gap_val:+.1f}%" if gap_val != 0 else "0%",
                po["attainment_level"],
                po["status"],
            ])

        col_w = [1.2*cm, 6.5*cm, 2.2*cm, 1.8*cm, 1.5*cm, 1.8*cm, 2*cm]
        t = Table(rows, colWidths=col_w, repeatRows=1)
        ts = [
            ("BACKGROUND",  (0, 0), (-1, 0), C_PURPLE),
            ("TEXTCOLOR",   (0, 0), (-1, 0), C_WHITE),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0), (-1, -1), 8),
            ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
            ("ALIGN",       (1, 1), (1, -1), "LEFT"),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",        (0, 0), (-1, -1), 0.4, C_DGRAY),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_WHITE, C_LGRAY]),
            ("PADDING",     (0, 0), (-1, -1), 6),
            ("FONTNAME",    (0, 1), (0, -1), "Helvetica-Bold"),
        ]
        for i, (po_id, po) in enumerate(gap_data["po_gaps"].items(), start=1):
            bg = C_GREEN if po["status"] == "ACHIEVED" else C_RED
            ts.append(("BACKGROUND", (6, i), (6, i), bg))
            ts.append(("TEXTCOLOR",  (6, i), (6, i), C_WHITE))
            ts.append(("FONTNAME",   (6, i), (6, i), "Helvetica-Bold"))
            if po["gap"] > 0:
                ts.append(("TEXTCOLOR", (4, i), (4, i), C_RED))
        t.setStyle(TableStyle(ts))
        items.append(t)
        items.append(Spacer(1, 0.5*cm))

        # PO bar chart
        items.append(Paragraph("PO Attainment vs Target — Visual Overview", styles["subsection"]))
        items.append(Spacer(1, 0.2*cm))
        items.append(self._bar_chart(gap_data["po_gaps"], "PO"))
        items.append(Spacer(1, 0.4*cm))

        # NAAC level descriptors
        items.append(Paragraph("NAAC Attainment Level Descriptors", styles["subsection"]))
        level_data = [
            ["Level", "Attainment Range", "Descriptor"],
            ["High",   "≥ 70%", "CO/PO is well attained. Continue current practices."],
            ["Medium", "50% – 69%", "Moderate attainment. Targeted improvement needed."],
            ["Low",    "< 50%", "CO/PO not attained. Immediate corrective action required."],
        ]
        lt = Table(level_data, colWidths=[2.5*cm, 4*cm, 10.5*cm])
        lt.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, 0), C_PURPLE),
            ("TEXTCOLOR",   (0, 0), (-1, 0), C_WHITE),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND",  (0, 1), (-1, 1), colors.HexColor("#EBF3E8")),
            ("BACKGROUND",  (0, 2), (-1, 2), colors.HexColor("#FFF8E1")),
            ("BACKGROUND",  (0, 3), (-1, 3), colors.HexColor("#FFEBEE")),
            ("FONTSIZE",    (0, 0), (-1, -1), 8.5),
            ("ALIGN",       (0, 0), (-1, -1), "LEFT"),
            ("GRID",        (0, 0), (-1, -1), 0.4, C_DGRAY),
            ("PADDING",     (0, 0), (-1, -1), 6),
            ("FONTNAME",    (0, 1), (0, -1), "Helvetica-Bold"),
        ]))
        items.append(lt)

        return items

    # ------------------------------------------------------------------
    # Section C — Recommendations
    # ------------------------------------------------------------------

    def _section_recommendations(self, recommendations: str, styles):
        items = []
        items.append(Paragraph("SECTION C — RECOMMENDATIONS", styles["section_label"]))
        items.append(Paragraph("AI-Generated Academic Improvement Recommendations",
                               styles["section_title"]))
        items.append(HRFlowable(width="100%", thickness=1.5, color=C_ORANGE))
        items.append(Spacer(1, 0.4*cm))

        items.append(Paragraph(
            "The following recommendations were generated by AI analysis of the attainment "
            "data above. They are intended as a starting point for faculty and curriculum "
            "committees and should be reviewed in the context of institutional policies.",
            styles["note"]
        ))
        items.append(Spacer(1, 0.5*cm))

        # Parse and render recommendation sections
        current_heading = None
        for line in recommendations.split("\n"):
            line = line.strip()
            if not line:
                items.append(Spacer(1, 0.2*cm))
                continue
            # Detect headings (ALL CAPS lines)
            if line.isupper() or (len(line) < 60 and line == line.upper() and len(line) > 5):
                items.append(Paragraph(line, styles["rec_heading"]))
            else:
                items.append(Paragraph(line, styles["rec_body"]))

        return items

    # ------------------------------------------------------------------
    # Section D — CO-PO Matrix
    # ------------------------------------------------------------------

    def _section_matrix(self, course, gap_data, styles):
        items = []
        items.append(Paragraph("SECTION D — CO-PO CORRELATION MATRIX", styles["section_label"]))
        items.append(Paragraph("Mapping of Course Outcomes to Program Outcomes",
                               styles["section_title"]))
        items.append(HRFlowable(width="100%", thickness=1.5, color=C_ORANGE))
        items.append(Spacer(1, 0.4*cm))

        items.append(Paragraph(
            "The CO-PO matrix below shows the correlation between each Course Outcome "
            "and Program Outcome. Correlation levels: <b>3 = High</b>, <b>2 = Medium</b>, "
            "<b>1 = Low</b>, <b>— = No correlation</b>. "
            "This matrix is the basis for PO attainment computation in Section B.",
            styles["body"]
        ))
        items.append(Spacer(1, 0.4*cm))

        pos = course.pos
        co_po_matrix = course.co_po_matrix
        po_ids = [p["po_id"] for p in pos]
        co_ids = list(gap_data["co_gaps"].keys())

        header = ["CO \\ PO"] + po_ids
        rows = [header]
        for co_id in co_ids:
            row = [co_id]
            for po_id in po_ids:
                val = co_po_matrix.get(co_id, {}).get(po_id, 0)
                row.append(str(val) if val else "—")
            rows.append(row)

        n_cols = len(header)
        col_w = [2.5*cm] + [max(1.2*cm, 15*cm / max(len(po_ids), 1))] * len(po_ids)
        t = Table(rows, colWidths=col_w, repeatRows=1)
        ts = [
            ("BACKGROUND",  (0, 0), (-1, 0), C_ORANGE),
            ("TEXTCOLOR",   (0, 0), (-1, 0), C_WHITE),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND",  (0, 0), (0, -1), C_NAVY),
            ("TEXTCOLOR",   (0, 0), (0, -1), C_WHITE),
            ("FONTNAME",    (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0), (-1, -1), 9),
            ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",        (0, 0), (-1, -1), 0.5, C_DGRAY),
            ("ROWBACKGROUNDS", (1, 1), (-1, -1), [C_WHITE, C_LGRAY]),
            ("PADDING",     (0, 0), (-1, -1), 7),
        ]
        # Colour cells by correlation strength
        for r_i, co_id in enumerate(co_ids, start=1):
            for c_i, po_id in enumerate(po_ids, start=1):
                val = co_po_matrix.get(co_id, {}).get(po_id, 0)
                if val == 3:
                    ts.append(("BACKGROUND", (c_i, r_i), (c_i, r_i), colors.HexColor("#C6EFCE")))
                elif val == 2:
                    ts.append(("BACKGROUND", (c_i, r_i), (c_i, r_i), colors.HexColor("#FFEB9C")))
                elif val == 1:
                    ts.append(("BACKGROUND", (c_i, r_i), (c_i, r_i), colors.HexColor("#FFC7CE")))
        t.setStyle(TableStyle(ts))
        items.append(t)
        items.append(Spacer(1, 0.5*cm))

        # Legend
        legend_data = [["3 = High Correlation", "2 = Medium Correlation",
                         "1 = Low Correlation", "— = No Correlation"]]
        lt = Table(legend_data, colWidths=[4*cm]*4)
        lt.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (0, 0), colors.HexColor("#C6EFCE")),
            ("BACKGROUND",  (1, 0), (1, 0), colors.HexColor("#FFEB9C")),
            ("BACKGROUND",  (2, 0), (2, 0), colors.HexColor("#FFC7CE")),
            ("BACKGROUND",  (3, 0), (3, 0), C_LGRAY),
            ("FONTSIZE",    (0, 0), (-1, -1), 8),
            ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
            ("GRID",        (0, 0), (-1, -1), 0.4, C_DGRAY),
            ("PADDING",     (0, 0), (-1, -1), 5),
            ("FONTNAME",    (0, 0), (-1, -1), "Helvetica-Bold"),
        ]))
        items.append(lt)

        return items

    # ------------------------------------------------------------------
    # Bar Chart (pure ReportLab Drawing)
    # ------------------------------------------------------------------

    def _bar_chart(self, gaps: dict, label_prefix: str) -> Drawing:
        ids = list(gaps.keys())
        actuals = [gaps[k]["attainment_percentage"] for k in ids]
        n = len(ids)

        chart_w = 450
        bar_h = 120
        drawing = Drawing(chart_w, bar_h + 40)

        bar_width = min(30, (chart_w - 60) / max(n * 2, 1))
        spacing = (chart_w - 60) / max(n, 1)
        x_start = 40

        # Y-axis line
        drawing.add(Line(x_start, 10, x_start, bar_h + 10,
                        strokeColor=C_DGRAY, strokeWidth=0.5))
        # X-axis line
        drawing.add(Line(x_start, 10, chart_w - 10, 10,
                        strokeColor=C_DGRAY, strokeWidth=0.5))

        # Target line
        target_y = 10 + (TARGET_PCT / 100) * bar_h
        drawing.add(Line(x_start, target_y, chart_w - 10, target_y,
                        strokeColor=C_RED, strokeWidth=1.2, strokeDashArray=[4, 3]))
        drawing.add(String(chart_w - 8, target_y - 3, f"Target {TARGET_PCT}%",
                          fontSize=6, fillColor=C_RED))

        # Y-axis labels
        for pct in [0, 25, 50, 75, 100]:
            y = 10 + (pct / 100) * bar_h
            drawing.add(Line(x_start - 4, y, x_start, y,
                            strokeColor=C_DGRAY, strokeWidth=0.5))
            drawing.add(String(2, y - 3, f"{pct}%", fontSize=6, fillColor=C_DGRAY))

        for i, (co_id, actual) in enumerate(zip(ids, actuals)):
            x = x_start + i * spacing + spacing / 2 - bar_width / 2
            bar_h_val = (actual / 100) * bar_h
            fill = C_GREEN if actual >= TARGET_PCT else C_RED

            # Actual bar
            drawing.add(Rect(x, 10, bar_width, bar_h_val,
                            fillColor=fill, strokeColor=C_WHITE, strokeWidth=0.5))

            # Value label on bar
            drawing.add(String(x + bar_width / 2, 10 + bar_h_val + 2,
                              f"{actual}%", fontSize=6.5,
                              fillColor=C_BLACK, textAnchor="middle"))

            # X-axis label
            drawing.add(String(x + bar_width / 2, 1,
                              co_id, fontSize=7,
                              fillColor=C_NAVY, textAnchor="middle"))

        return drawing

    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------

    def _make_styles(self):
        base = getSampleStyleSheet()
        s = {}

        s["cover_title"] = ParagraphStyle("cover_title",
            fontSize=28, textColor=C_NAVY, alignment=TA_CENTER,
            fontName="Helvetica-Bold", spaceAfter=8)

        s["cover_sub"] = ParagraphStyle("cover_sub",
            fontSize=13, textColor=C_BLUE, alignment=TA_CENTER,
            fontName="Helvetica", spaceAfter=4, spaceBefore=12)

        s["section_label"] = ParagraphStyle("section_label",
            fontSize=9, textColor=C_WHITE, alignment=TA_LEFT,
            fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4,
            backColor=C_NAVY, leftIndent=-10, rightIndent=-10,
            borderPad=5)

        s["section_title"] = ParagraphStyle("section_title",
            fontSize=14, textColor=C_NAVY, alignment=TA_LEFT,
            fontName="Helvetica-Bold", spaceBefore=8, spaceAfter=4)

        s["subsection"] = ParagraphStyle("subsection",
            fontSize=10, textColor=C_NAVY, alignment=TA_LEFT,
            fontName="Helvetica-Bold", spaceBefore=8, spaceAfter=4)

        s["body"] = ParagraphStyle("body",
            fontSize=9, textColor=C_BLACK, alignment=TA_LEFT,
            fontName="Helvetica", spaceAfter=4, leading=14)

        s["note"] = ParagraphStyle("note",
            fontSize=8.5, textColor=C_DGRAY, alignment=TA_LEFT,
            fontName="Helvetica-Oblique", spaceAfter=4,
            backColor=C_LGRAY, borderPad=6, leading=13)

        s["rec_heading"] = ParagraphStyle("rec_heading",
            fontSize=10, textColor=C_NAVY, alignment=TA_LEFT,
            fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=3)

        s["rec_body"] = ParagraphStyle("rec_body",
            fontSize=9, textColor=C_BLACK, alignment=TA_LEFT,
            fontName="Helvetica", spaceAfter=3, leading=14,
            leftIndent=10)

        return s