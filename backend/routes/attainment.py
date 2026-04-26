# backend/routes/attainment.py
import io
import os
from typing import Dict, List, Optional
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
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
    marks: Dict[str, Dict[str, float]] = Field(
        description="CO → component → marks scored",
        example={
            "CO1": {"Quiz": 8.0, "Unit Test": 18.0},
            "CO2": {"Quiz": 7.0, "Unit Test": 15.0},
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
                    co_columns[j] = str(val).strip().upper()
            break

    if header_row_idx is None:
        raise ValueError("Could not find header row with 'PRN' column in the xlsx file.")

    students = []
    for row in all_rows[header_row_idx + 1:]:
        if not row or len(row) < 3:
            continue
        prn = row[1]   # Column B
        name = row[2]  # Column C

        # Skip section headers, empty rows, formula-only rows
        if not prn or not name:
            continue
        if isinstance(prn, str) and not prn.strip().isdigit():
            continue
        if isinstance(name, str) and ('section' in name.lower() or not name.strip()):
            continue

        student_id = str(int(float(str(prn)))).strip()
        student_name = str(name).strip().replace('\xa0', '').strip()

        # Build marks dict — if CO columns exist use them, else put all numeric cols under CO1
        marks = {}
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
        service = AttainmentService(db)
        result = await service.save_marks(course_id, students)
        return {"status": "success", "data": result, "parsed_students": len(students)}
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


@router.get("/calculate/{course_id}")
async def calculate_attainment(
    course_id: int,
    db: AsyncSession = Depends(get_db),
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
    db: AsyncSession = Depends(get_db),
):
    """
    Generate a formatted Word document attainment report.
    Requires marks to be uploaded first via POST /attainment/marks/{course_id}.
    """
    logger.info(f"Attainment report requested for course_id={course_id}")
    try:
        service = AttainmentService(db)
        result = await service.generate_report(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"Attainment report error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error generating attainment report for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/download/{course_id}")
async def download_attainment_report(course_id: int):
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
    db: AsyncSession = Depends(get_db),
):
    """
    Generate a full NBA/NAAC-format PDF report with:
    - Section A: NBA CO attainment + gap analysis + bar charts
    - Section B: NAAC PO attainment + level descriptors
    - Section C: AI-generated recommendations
    - Section D: CO-PO correlation matrix

    Requires marks to be uploaded first via POST /attainment/marks/{course_id}.
    """
    logger.info(f"NBA/NAAC PDF report requested for course_id={course_id}")
    try:
        service = NBAReportService(db)
        result = await service.generate_pdf(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"NBA report error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error generating NBA report for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/nba-report/download/{course_id}")
async def download_nba_report(course_id: int):
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