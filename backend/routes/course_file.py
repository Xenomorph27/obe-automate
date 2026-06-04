# backend/routes/course_file.py
"""
Routes for generating and managing the full OBE Course File.

POST /course-file/generate/{course_id}          — generate the complete .docx
GET  /course-file/download/{course_id}          — download the generated .docx
GET  /course-file/extra/{course_id}             — get saved extra fields
POST /course-file/extra/{course_id}             — save extra fields
POST /course-file/upload/{course_id}            — upload a dept-specific attachment
GET  /course-file/attachments/{course_id}       — list all attachments for a course
DELETE /course-file/attachment/{attachment_id}  — delete an attachment
GET  /course-file/attachment/{attachment_id}/download — download an attachment
"""
import os
import uuid
import mimetypes
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from backend.core.auth import require_auth
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.database.connection import get_db
from backend.database.user_models import User
from backend.database.models import CourseFileExtra, CourseFileAttachment
from backend.services.course_file_service import CourseFileService

logger = get_logger(__name__)
router = APIRouter(prefix="/course-file", tags=["Course File"])

_ATTACH_CATEGORY = "course_file_attachments"
_MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB


# ── Pydantic schemas ──────────────────────────────────────────────────────────

class CourseFileExtraPayload(BaseModel):
    vision_text: Optional[str] = ""
    mission_text: Optional[str] = ""
    batch: Optional[str] = ""
    po_peo_pso_text: Optional[str] = ""
    peo_text: Optional[str] = ""
    pso_text: Optional[str] = ""
    co_po_justification: Optional[str] = ""
    prev_co_attainment: Optional[str] = ""
    action_plan: Optional[str] = ""
    slow_learners: Optional[str] = ""
    advanced_learners: Optional[str] = ""
    activity_reports: Optional[str] = ""
    learning_material_links: Optional[str] = ""
    attendance_links: Optional[str] = ""
    student_list: Optional[str] = ""
    custom_tabs: Optional[str] = "[]"


# ── Existing endpoints ────────────────────────────────────────────────────────

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
    extra.po_peo_pso_text = payload.po_peo_pso_text or ""
    extra.peo_text = payload.peo_text or ""
    extra.pso_text = payload.pso_text or ""
    extra.co_po_justification = payload.co_po_justification or ""
    extra.prev_co_attainment = payload.prev_co_attainment or ""
    extra.action_plan = payload.action_plan or ""
    extra.slow_learners = payload.slow_learners or ""
    extra.advanced_learners = payload.advanced_learners or ""
    extra.activity_reports = payload.activity_reports or ""
    extra.learning_material_links = payload.learning_material_links or ""
    extra.attendance_links = payload.attendance_links or ""
    extra.student_list = payload.student_list or ""
    extra.custom_tabs = payload.custom_tabs or "[]"

    await db.commit()
    logger.info(f"Course file extra saved for course_id={course_id}")
    return {"status": "success", "message": "Saved successfully"}


# ── Attachment endpoints ──────────────────────────────────────────────────────

@router.post("/upload/{course_id}", status_code=201)
async def upload_attachment(
    course_id: int,
    file: UploadFile = File(...),
    label: str = Form(...),
    section_no: Optional[int] = Form(None),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Upload a department-specific file (timetable, event photo, email screenshot, etc.)
    and attach it to the given course's file. Max 50 MB.
    """
    data = await file.read()
    if len(data) > _MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="File too large — maximum size is 50 MB.")

    # Build a unique stored filename so collisions are impossible
    ext = os.path.splitext(file.filename or "")[-1].lower()
    stored_name = f"course{course_id}_{uuid.uuid4().hex}{ext}"

    storage = get_storage()
    stored_path = storage.save(_ATTACH_CATEGORY, stored_name, data)

    mime = file.content_type or mimetypes.guess_type(file.filename or "")[0] or "application/octet-stream"

    attachment = CourseFileAttachment(
        course_id=course_id,
        section_no=section_no,
        label=label.strip(),
        filename=file.filename or stored_name,
        stored_path=str(stored_path),
        mime_type=mime,
        file_size=len(data),
    )
    db.add(attachment)
    await db.commit()
    await db.refresh(attachment)

    logger.info(f"Attachment uploaded: course={course_id} label='{label}' file={stored_name}")
    return {"status": "success", "data": attachment.to_dict()}


@router.get("/attachments/{course_id}")
async def list_attachments(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """List all department-specific attachments for a course."""
    result = await db.execute(
        select(CourseFileAttachment)
        .where(CourseFileAttachment.course_id == course_id)
        .order_by(CourseFileAttachment.section_no, CourseFileAttachment.uploaded_at)
    )
    attachments = result.scalars().all()
    return {"status": "success", "data": [a.to_dict() for a in attachments]}


@router.delete("/attachment/{attachment_id}", status_code=200)
async def delete_attachment(
    attachment_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Delete a department attachment (removes DB record and stored file)."""
    result = await db.execute(
        select(CourseFileAttachment).where(CourseFileAttachment.id == attachment_id)
    )
    attachment = result.scalar_one_or_none()
    if not attachment:
        raise HTTPException(status_code=404, detail="Attachment not found.")

    # Remove from disk
    storage = get_storage()
    stored_name = os.path.basename(attachment.stored_path)
    storage.delete(_ATTACH_CATEGORY, stored_name)

    await db.delete(attachment)
    await db.commit()
    logger.info(f"Attachment deleted: id={attachment_id}")
    return {"status": "success", "message": "Attachment deleted."}


@router.get("/attachment/{attachment_id}/download")
async def download_attachment(
    attachment_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Download a specific department attachment."""
    result = await db.execute(
        select(CourseFileAttachment).where(CourseFileAttachment.id == attachment_id)
    )
    attachment = result.scalar_one_or_none()
    if not attachment:
        raise HTTPException(status_code=404, detail="Attachment not found.")

    if not os.path.exists(attachment.stored_path):
        raise HTTPException(status_code=404, detail="File not found on disk.")

    return FileResponse(
        path=attachment.stored_path,
        media_type=attachment.mime_type or "application/octet-stream",
        filename=attachment.filename,
    )
