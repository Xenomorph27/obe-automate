# backend/routes/session_plan.py
import os
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.database.connection import get_db
from backend.services.session_plan_service import SessionPlanService

logger = get_logger(__name__)
router = APIRouter(prefix="/session-plan", tags=["Session Plan"])


@router.post("/generate/{course_id}", status_code=201)
async def generate_session_plan(
    course_id: int,
    db: AsyncSession = Depends(get_db),
):
    """
    Generates a unit-wise session plan for the course using Gemini AI.
    Returns metadata and a download URL for the Word document.
    """
    logger.info(f"Session plan generation requested for course_id={course_id}")
    try:
        service = SessionPlanService(db)
        result = await service.generate(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"Session plan error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error generating session plan for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/download/{course_id}")
async def download_session_plan(course_id: int):
    """Download the generated session plan Word document."""
    filepath = SessionPlanService.get_filepath(course_id)
    if not os.path.exists(filepath):
        raise HTTPException(
            status_code=404,
            detail=f"Session plan not found for course_id={course_id}. Run POST /session-plan/generate/{course_id} first.",
        )
    return FileResponse(
        path=filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"session_plan_{course_id}.docx",
    )