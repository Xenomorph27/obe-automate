# backend/routes/course_file.py
"""
Routes for generating and managing the full OBE Course File.

POST /course-file/generate/{course_id}    — generate the complete .docx
GET  /course-file/download/{course_id}    — download the generated .docx
GET  /course-file/extra/{course_id}       — get saved extra fields (vision, mission, etc.)
POST /course-file/extra/{course_id}       — save extra fields
"""
import os
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from backend.core.auth import require_auth
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.database.connection import get_db
from backend.database.user_models import User
from backend.database.models import CourseFileExtra
from backend.services.course_file_service import CourseFileService

logger = get_logger(__name__)
router = APIRouter(prefix="/course-file", tags=["Course File"])


class CourseFileExtraPayload(BaseModel):
    vision_text: Optional[str] = ""
    mission_text: Optional[str] = ""
    batch: Optional[str] = ""
    prev_co_attainment: Optional[str] = ""
    action_plan: Optional[str] = ""
    slow_learners: Optional[str] = ""
    advanced_learners: Optional[str] = ""
    activity_reports: Optional[str] = ""
    learning_material_links: Optional[str] = ""
    attendance_links: Optional[str] = ""


@router.post("/generate/{course_id}", status_code=201)
async def generate_course_file(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Generate the complete 13-section OBE course file as a Word document."""
    logger.info(f"Course file generation requested for course_id={course_id}")
    try:
        svc = CourseFileService(db)
        result = await svc.generate(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"Course file error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception as e:
        logger.exception(f"Unexpected error generating course file for course {course_id}")
        raise HTTPException(status_code=500, detail=f"Internal server error: {str(e)}")


@router.get("/download/{course_id}")
async def download_course_file(
    course_id: int,
    current_user: User = Depends(require_auth),
):
    """Download the generated course file Word document."""
    filepath = CourseFileService.get_filepath(course_id)
    if not os.path.exists(filepath):
        raise HTTPException(
            status_code=404,
            detail=f"Course file not found for course_id={course_id}. "
                   f"Run POST /course-file/generate/{course_id} first.",
        )
    return FileResponse(
        path=filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"course_file_{course_id}.docx",
    )


@router.get("/extra/{course_id}")
async def get_course_file_extra(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Get the saved extra fields for course file generation (vision, mission, etc.)."""
    result = await db.execute(
        select(CourseFileExtra).where(CourseFileExtra.course_id == course_id)
    )
    extra = result.scalar_one_or_none()
    return {"status": "success", "data": extra.to_dict() if extra else {}}


@router.post("/extra/{course_id}", status_code=200)
async def save_course_file_extra(
    course_id: int,
    payload: CourseFileExtraPayload,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Save extra fields needed for course file generation."""
    result = await db.execute(
        select(CourseFileExtra).where(CourseFileExtra.course_id == course_id)
    )
    extra = result.scalar_one_or_none()

    if not extra:
        extra = CourseFileExtra(course_id=course_id)
        db.add(extra)

    extra.vision_text = payload.vision_text or ""
    extra.mission_text = payload.mission_text or ""
    extra.batch = payload.batch or ""
    extra.prev_co_attainment = payload.prev_co_attainment or ""
    extra.action_plan = payload.action_plan or ""
    extra.slow_learners = payload.slow_learners or ""
    extra.advanced_learners = payload.advanced_learners or ""
    extra.activity_reports = payload.activity_reports or ""
    extra.learning_material_links = payload.learning_material_links or ""
    extra.attendance_links = payload.attendance_links or ""

    await db.commit()
    logger.info(f"Course file extra saved for course_id={course_id}")
    return {"status": "success", "message": "Saved successfully"}
