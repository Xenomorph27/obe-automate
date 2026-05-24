# backend/routes/attainment.py
import io
import os
from typing import Any, Dict, List, Optional
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from backend.core.auth import require_auth
from backend.database.user_models import User
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.database.connection import get_db
from backend.services.attainment_service import AttainmentService
from backend.services.nba_report_service import NBAReportService

logger = get_logger(__name__)
router = APIRouter(prefix="/attainment", tags=["CO Attainment"])


# ── Request Models ──────────────────────────────────────────────────────

class StudentMarks(BaseModel):
    student_id: str = Field(example="USN001")
    student_name: str = Field(example="Alice Kumar")
    marks: Dict[str, Any] = Field(
        description="CO → component → marks (nested), OR exam → mark (flat exam-wise)",
        example={
            "CO1": {"Quiz": 8.0, "Unit Test": 18.0},
        }
    )


class MarksUploadRequest(BaseModel):
    students: List[StudentMarks] = Field(min_length=1)


# ── XLSX Parser ─────────────────────────────────────────────────────────

def parse_marks_xlsx(file_bytes: bytes) -> List[dict]:
    """
    Parse a marks xlsx file with this structure:
      - Header rows at top (institute name, batch info, branch) — skipped
      - Row with 'SR. No.' / 'PRN' / name = column header row
      - Section rows (e.g. 'Section A') — skipped
      - Student rows: col B = PRN (student_id), col C = name
      - CO mark columns start from col D onward, named like CO1, CO2...
      - Rows with formula strings in col A (=IF...) or numeric serial → student rows
    """
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    all_rows = list(ws.iter_rows(values_only=True))

    # Find the header row — contains 'PRN' or 'prn'
    header_row_idx = None
    co_columns = {}  # col_index → CO name
    for i, row in enumerate(all_rows):
        row_str = [str(v).strip().lower() if v else '' for v in row]
        if 'prn' in row_str:
            header_row_idx = i
            # Find CO columns (D onward)
            for j, val in enumerate(row):
                if val and str(val).strip().upper().startswith('CO'):
                    # Normalize: "CO1\n(Quiz /10)" -> "CO1"
                    import re as _re
                    raw = str(val).strip().upper().split('\n')[0].strip()
                    co_name = _re.match(r'(CO\d+)', raw)
                    if co_name:
                        co_columns[j] = co_name.group(1)
            break

    if header_row_idx is None:
        raise ValueError("Could not find header row with 'PRN' column in the xlsx file.")

    # Also look for exam-wise columns (non-CO columns like "Quiz", "Unit Test", "Case Study", etc.)
    header_row = all_rows[header_row_idx]
    exam_columns = {}   # col_index → exam name (cleaned)
    SKIP_COLS = {'sr no', 'sr. no', 'sr.no', 'prn', 'roll no', 'student name', 'name',
                 'sec', 'section', 'ca total', 'grand total', 'grade', 'scaled'}
    for j, val in enumerate(header_row):
        if val is None:
            continue
        # Normalize header: strip newlines and extra whitespace
        raw = str(val).strip().replace('\n', ' ').strip()
        raw_lower = raw.lower()
        # Skip known non-exam columns
        if any(s in raw_lower for s in SKIP_COLS):
            continue
        # Skip CO columns (already handled above)
        if raw.upper().startswith('CO') and raw[2:3].isdigit():
            continue
        # Include columns that look like exam components: contain letters and are reasonable headers
        if raw and not raw.replace('.','').replace('/','').isdigit():
            # Strip marks suffix like "/10", "/5", "/60"
            import re as _re2
            clean = _re2.sub(r'\s*/\s*\d+', '', raw).strip()
            if clean and clean.lower() not in SKIP_COLS:
                exam_columns[j] = clean

    students = []
    for row in all_rows[header_row_idx + 1:]:
        if not row or len(row) < 3:
            continue
        prn = row[1]   # Column B
        name = row[2]  # Column C

        # Skip section headers, empty rows, formula-only rows, max-marks rows
        if not prn or not name:
            continue
        prn_str = str(prn).strip()
        # Skip non-PRN rows (headers, "Max →", section markers, class average)
        if isinstance(prn, str) and not prn_str.replace('.','').isdigit():
            continue
        if isinstance(name, str) and any(x in name.lower() for x in ['section','max','average','class avg','max marks']):
            continue
        # PRN must be a reasonable number (at least 6 digits)
        try:
            prn_int = int(float(prn_str))
            if prn_int < 100000:  # not a valid PRN
                continue
        except (ValueError, TypeError):
            continue

        student_id = str(int(float(str(prn)))).strip()
        student_name = str(name).strip().replace('\xa0', '').strip()

        # Build marks dict
        marks = {}
        if exam_columns:
            # Exam-wise format: store each exam as its own component under a "Marks" key
            # Use component name as the CO key so the frontend can display columns per exam
            for col_idx, exam_name in exam_columns.items():
                val = row[col_idx] if col_idx < len(row) else None
                # Skip "—", "N/A", empty
                if val is None or str(val).strip() in ('', '—', 'N/A'):
                    continue
                try:
                    marks[exam_name] = float(val)
                except (TypeError, ValueError):
                    pass
            # Store as flat exam→mark mapping under a special key
            if marks:
                students.append({
                    "student_id": student_id,
                    "student_name": student_name,
                    "marks": marks,   # flat: {"Quiz": 8.5, "Unit Test": 7.0, ...}
                    "_format": "exam_wise",
                })
                continue
        if co_columns:
            for col_idx, co_name in co_columns.items():
                val = row[col_idx] if col_idx < len(row) else None
                try:
                    marks[co_name] = {"Total": float(val) if val is not None else 0.0}
                except (TypeError, ValueError):
                    marks[co_name] = {"Total": 0.0}
        else:
            # No CO columns found — put all numeric values from col D onward as CO1, CO2...
            co_idx = 1
            for j in range(3, len(row)):
                val = row[j]
                if val is not None:
                    try:
                        marks[f"CO{co_idx}"] = {"Total": float(val)}
                        co_idx += 1
                    except (TypeError, ValueError):
                        pass

        if student_id and student_name:
            students.append({
                "student_id": student_id,
                "student_name": student_name,
                "marks": marks if marks else {"CO1": {"Total": 0.0}},
            })

    return students


# ── Routes ─────────────────────────────────────────────────────────────

@router.post("/marks/{course_id}/xlsx", status_code=201)
async def upload_marks_xlsx(
    course_id: int,
    file: UploadFile = File(..., description="Student marks Excel file (.xlsx)"),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Upload student marks from an Excel (.xlsx) file.
    Expected columns: PRN (col B), Student Name (col C), CO marks (col D+).
    Handles institute header rows, section separators, and formula cells automatically.
    Replaces all previous marks for this course.
    """
    logger.info(f"XLSX marks upload for course_id={course_id}, file={file.filename}")
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")
    try:
        file_bytes = await file.read()
        students = parse_marks_xlsx(file_bytes)
        if not students:
            raise HTTPException(status_code=400, detail="No student records found in the file. Check the format.")
        logger.info(f"Parsed {len(students)} students from xlsx")
        # Detect if this is an exam-wise file (flat marks dict)
        is_exam_wise = any(s.get("_format") == "exam_wise" for s in students)
        # Clean up the _format key and normalize marks for saving
        cleaned_students = []
        for s in students:
            fmt = s.pop("_format", None)
            if fmt == "exam_wise":
                # Store flat exam marks directly — frontend handles display
                cleaned_students.append({
                    "student_id": s["student_id"],
                    "student_name": s["student_name"],
                    "marks": s["marks"],  # {"Quiz": 8.5, "Unit Test": 7.0, ...}
                })
            else:
                cleaned_students.append(s)
        service = AttainmentService(db)
        result = await service.save_marks(course_id, cleaned_students)
        return {
            "status": "success",
            "data": result,
            "parsed_students": len(cleaned_students),
            "format": "exam_wise" if is_exam_wise else "co_wise",
        }
    except HTTPException:
        raise
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except OBEException as e:
        logger.error(f"Marks upload error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error uploading xlsx marks for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.post("/marks/{course_id}", status_code=201)
async def upload_marks(
    course_id: int,
    request: MarksUploadRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Upload student marks per CO per evaluation component (JSON body).
    Replaces all previous marks for this course.
    """
    logger.info(f"Marks upload for course_id={course_id}, students={len(request.students)}")
    try:
        service = AttainmentService(db)
        result = await service.save_marks(
            course_id,
            [s.dict() for s in request.students],
        )
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"Marks upload error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error uploading marks for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/marks/{course_id}")
async def get_marks(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Return all student marks records for a course as a list of dicts.
    Used by the frontend to display and edit uploaded marks.
    """
    from sqlalchemy import select as _select
    from backend.database.models import COAttainment
    logger.info(f"Fetching marks for course_id={course_id}")
    try:
        result = await db.execute(
            _select(COAttainment).where(COAttainment.course_id == course_id)
        )
        records = result.scalars().all()
        return {"status": "success", "data": [r.to_dict() for r in records], "total": len(records)}
    except Exception:
        logger.exception(f"Error fetching marks for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.put("/marks/{course_id}/student/{student_id}", status_code=200)
async def update_student_marks(
    course_id: int,
    student_id: str,
    request: MarksUploadRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Update marks for a single student in a course.
    Finds the existing COAttainment record by course_id + student_id and patches marks.
    """
    from sqlalchemy import select as _select
    from backend.database.models import COAttainment
    logger.info(f"Updating marks for course_id={course_id} student_id={student_id}")
    try:
        result = await db.execute(
            _select(COAttainment).where(
                COAttainment.course_id == course_id,
                COAttainment.student_id == student_id,
            )
        )
        rec = result.scalar_one_or_none()
        if not rec:
            raise HTTPException(status_code=404, detail=f"No marks found for student_id={student_id}")
        student_data = request.students[0]
        rec.student_name = student_data.student_name
        rec.marks = student_data.marks
        await db.commit()
        await db.refresh(rec)
        return {"status": "success", "data": rec.to_dict()}
    except HTTPException:
        raise
    except Exception:
        logger.exception(f"Error updating marks for course {course_id} student {student_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/documents/{course_id}")
async def list_documents(course_id: int, current_user: User = Depends(require_auth)):
    """
    List all generated documents for a course (session plan, evaluation plan,
    attainment report, NBA report, question papers).
    Returns filename, category, size, download URL.
    """
    from backend.core.storage import get_storage
    storage = get_storage()
    docs = []
    checks = [
        ("session_plans",     f"session_plan_{course_id}.docx",    "Session Plan",        f"/session-plan/download/{course_id}",            "📅"),
        ("evaluation_plans",  f"eval_plan_{course_id}.docx",       "Evaluation Plan",     f"/evaluation-plan/download/{course_id}",          "📋"),
        ("attainment_reports",f"attainment_report_{course_id}.docx","Attainment Report",  f"/attainment/download/{course_id}",               "📊"),
        ("nba_reports",       f"nba_report_{course_id}.pdf",        "NBA/NAAC PDF",        f"/attainment/nba-report/download/{course_id}",    "🏆"),
        ("question_papers",   f"qpaper_{course_id}_latest.docx",    "Question Paper (.docx)",f"/questions/paper/download/{course_id}/docx",  "❓"),
        ("question_papers",   f"qpaper_{course_id}_latest.pdf",     "Question Paper (.pdf)", f"/questions/paper/download/{course_id}/pdf",   "❓"),
    ]
    for category, filename, label, url, icon in checks:
        p = storage.get_path(category, filename)
        if p and p.exists():
            stat = p.stat()
            docs.append({
                "label": label,
                "filename": filename,
                "category": category,
                "icon": icon,
                "size_kb": round(stat.st_size / 1024, 1),
                "download_url": url,
                "generated_at": stat.st_mtime,
            })
        else:
            # Also try listing by prefix for question papers which may have timestamps
            if category == "question_papers":
                files = storage.list_files(category, prefix=f"qpaper_{course_id}_")
                for f in files:
                    stat = f.stat()
                    ext = f.suffix
                    lbl = f"Question Paper ({ext})"
                    dl_url = f"/questions/paper/download/{course_id}/docx" if ext == ".docx" else f"/questions/paper/download/{course_id}/pdf"
                    docs.append({
                        "label": lbl, "filename": f.name, "category": category,
                        "icon": "❓", "size_kb": round(stat.st_size / 1024, 1),
                        "download_url": dl_url, "generated_at": stat.st_mtime,
                    })
    return {"status": "success", "data": docs, "total": len(docs)}


@router.get("/calculate/{course_id}")
async def calculate_attainment(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Calculate and return CO + PO attainment as JSON (no file generated).
    Use this to preview attainment before generating the report.
    """
    try:
        service = AttainmentService(db)
        result = await service.calculate(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Error calculating attainment for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.post("/report/{course_id}", status_code=201)
async def generate_attainment_report(
    course_id: int,
    template: Optional[UploadFile] = File(None, description="Optional .docx template to guide report structure"),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Generate a formatted Word document attainment report.
    Requires marks to be uploaded first via POST /attainment/marks/{course_id}.
    Optionally accepts a .docx template — if provided, the report follows that template's structure.
    """
    logger.info(f"Attainment report requested for course_id={course_id}, template={'yes' if template else 'no'}")
    try:
        service = AttainmentService(db)
        if template and template.filename:
            template_bytes = await template.read()
            result = await service.generate_report_from_template(course_id, template_bytes)
        else:
            result = await service.generate_report(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"Attainment report error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error generating attainment report for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/download/{course_id}")
async def download_attainment_report(course_id: int, current_user: User = Depends(require_auth)):
    """Download the generated attainment report Word document."""
    filepath = AttainmentService.get_filepath(course_id)
    if not os.path.exists(filepath):
        raise HTTPException(
            status_code=404,
            detail=f"Report not found for course_id={course_id}. Run POST /attainment/report/{course_id} first.",
        )
    return FileResponse(
        path=filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"attainment_report_{course_id}.docx",
    )


# ── Day 6: NBA/NAAC PDF Routes ──────────────────────────────────────────

@router.post("/nba-report/{course_id}", status_code=201)
async def generate_nba_report(
    course_id: int,
    template: Optional[UploadFile] = File(None, description="Optional .docx template to guide report structure"),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Generate a full NBA/NAAC-format PDF report with:
    - Section A: NBA CO attainment + gap analysis + bar charts
    - Section B: NAAC PO attainment + level descriptors
    - Section C: AI-generated recommendations
    - Section D: CO-PO correlation matrix

    Requires marks to be uploaded first via POST /attainment/marks/{course_id}.
    Optionally accepts a .docx template — if provided, the PDF follows that template's section structure.
    """
    logger.info(f"NBA/NAAC PDF report requested for course_id={course_id}, template={'yes' if template else 'no'}")
    try:
        service = NBAReportService(db)
        if template and template.filename:
            template_bytes = await template.read()
            result = await service.generate_pdf_from_template(course_id, template_bytes)
        else:
            result = await service.generate_pdf(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"NBA report error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error generating NBA report for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/nba-report/download/{course_id}")
async def download_nba_report(course_id: int, current_user: User = Depends(require_auth)):
    """Download the generated NBA/NAAC PDF report."""
    filepath = NBAReportService.get_filepath(course_id)
    if not os.path.exists(filepath):
        raise HTTPException(
            status_code=404,
            detail=f"NBA report not found for course_id={course_id}. Run POST /attainment/nba-report/{course_id} first.",
        )
    return FileResponse(
        path=filepath,
        media_type="application/pdf",
        filename=f"nba_report_{course_id}.pdf",
    )


@router.get("/nba-report/gap-analysis/{course_id}")
async def get_gap_analysis(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Preview the CO-PO gap analysis as JSON without generating the PDF.
    Shows which COs and POs are at risk (below target threshold).
    """
    try:
        service = NBAReportService(db)
        result = await service.gap_analysis(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Error running gap analysis for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")