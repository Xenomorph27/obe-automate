# backend/routes/attainment.py
import os
from typing import Dict, List
from fastapi import APIRouter, Depends, HTTPException
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


# ── Existing Routes ─────────────────────────────────────────────────────

@router.post("/marks/{course_id}", status_code=201)
async def upload_marks(
    course_id: int,
    request: MarksUploadRequest,
    db: AsyncSession = Depends(get_db),
):
    """
    Upload student marks per CO per evaluation component.
    Replaces all previous marks for this course.
    Send marks for ALL students at once.
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