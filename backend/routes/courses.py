# backend/routes/courses.py
from fastapi import APIRouter, Depends, HTTPException
from backend.core.auth import require_auth
from backend.database.user_models import User
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from backend.database.connection import get_db
from backend.services.course_service import CourseService
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.database.models import CourseFileAttachment
from backend.services.course_file_service import CourseFileService
from pydantic import BaseModel, Field
from typing import Dict, List
import os

logger = get_logger(__name__)
router = APIRouter(prefix="/courses", tags=["Courses"])

# --- Request Models (what faculty sends) ---

class CourseOutcome(BaseModel):
    co_id: str = Field(example="CO1")
    statement: str = Field(example="Understand the fundamentals of OS")
    bloom_level: str = Field(example="Understand")


class ProgramOutcome(BaseModel):
    po_id: str = Field(example="PO1")
    statement: str = Field(example="Engineering Knowledge")


class ProgramSpecificOutcome(BaseModel):
    pso_id: str = Field(example="PSO1")
    statement: str = Field(example="Apply AI/ML concepts to real-world problems")


class Unit(BaseModel):
    unit_number: int = Field(example=1)
    unit_title: str = Field(example="Introduction to Unsupervised Learning")
    topics: List[str] = Field(default_factory=list, example=[
        "Introduction to Machine Learning, applications",
        "Types of Learning: Supervised, Unsupervised and Semi-Supervised Learning"
    ])


class EvaluationConfig(BaseModel):
    continuous_assessment_total: int = Field(example=30)
    components: Dict[str, int] = Field(
        example={"Mind Map": 5, "Quiz": 10, "Unit Test": 10, "Case Study": 5}
    )
    end_sem_total: int = Field(example=60)


class CourseSetupRequest(BaseModel):
    course_name: str = Field(example="Operating Systems")
    course_code: str = Field(example="CS301")
    credits: int = Field(example=4)
    total_hours: int = Field(example=60)
    faculty_name: str = Field(example="Dr. Sharma")
    department: str = Field(example="Computer Science")
    semester: str = Field(example="5th")
    academic_year: str = Field(example="2024-25")
    cos: List[CourseOutcome]
    pos: List[ProgramOutcome]
    psos: List[ProgramSpecificOutcome] = Field(default_factory=list)
    units: List[Unit] = Field(default_factory=list)  # ★ ADDED — was missing, caused syllabus topics to be dropped
    co_po_matrix: Dict[str, Dict[str, int]] = Field(
        description="CO-PO mapping. Keys are CO IDs, values are dicts of PO ID → correlation (0,1,2,3)",
        example={"CO1": {"PO1": 3, "PO2": 2}, "CO2": {"PO1": 1, "PO2": 3}}
    )
    co_pso_matrix: Dict[str, Dict[str, int]] = Field(
        default_factory=dict,
        description="CO-PSO mapping. Same correlation scale as CO-PO.",
        example={"CO1": {"PSO1": 3}, "CO2": {"PSO1": 2, "PSO2": 1}}
    )
    evaluation_config: EvaluationConfig


class CourseUpdateRequest(BaseModel):
    course_name: str | None = None
    course_code: str | None = None
    credits: int | None = None
    total_hours: int | None = None
    faculty_name: str | None = None
    department: str | None = None
    semester: str | None = None
    academic_year: str | None = None
    cos: List[CourseOutcome] | None = None
    pos: List[ProgramOutcome] | None = None
    psos: List[ProgramSpecificOutcome] | None = None
    units: List[Unit] | None = None  # ★ ADDED
    co_po_matrix: Dict[str, Dict[str, int]] | None = None
    co_pso_matrix: Dict[str, Dict[str, int]] | None = None
    evaluation_config: EvaluationConfig | None = None


# --- Routes ---

@router.post("/setup", status_code=201)
async def setup_course(
    request: CourseSetupRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Faculty submits all start-of-semester data.
    Creates the course record that all other features depend on.
    """
    logger.info(f"Course setup request: {request.course_code} — {request.course_name}")
    try:
        service = CourseService(db)
        course = await service.create_course(request)
        return {
            "status": "success",
            "message": "Course created successfully",
            "course_id": course.id,
            "data": course.to_dict()
        }
    except OBEException as e:
        logger.error(f"Course setup failed: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception as e:
        logger.exception("Unexpected error during course setup")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/{course_id}")
async def get_course(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Retrieve full course data by ID."""
    try:
        service = CourseService(db)
        course = await service.get_course(course_id)
        return {"status": "success", "data": course.to_dict()}
    except OBEException as e:
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Error fetching course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/")
async def list_courses(db: AsyncSession = Depends(get_db), current_user: User = Depends(require_auth)):
    """List all courses."""
    try:
        service = CourseService(db)
        courses = await service.list_courses()
        return {"status": "success", "count": len(courses), "data": courses}
    except Exception:
        logger.exception("Error listing courses")
        raise HTTPException(status_code=500, detail="Internal server error")

@router.put("/{course_id}")
async def update_course(
    course_id: int,
    request: CourseUpdateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Update an existing course's fields (partial update — only non-None fields applied)."""
    logger.info(f"Course update request for id={course_id}")
    try:
        service = CourseService(db)
        course = await service.update_course(course_id, request)
        return {
            "status": "success",
            "message": "Course updated successfully",
            "data": course.to_dict()
        }
    except OBEException as e:
        logger.error(f"Course update failed: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception as e:
        logger.exception("Unexpected error during course update")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.delete("/{course_id}")
async def delete_course(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Permanently delete a course and ALL associated data:
    - DB rows (session plan, eval plan, marks, questions, attachments, course file extra) via CASCADE
    - Generated .docx file from storage
    - All uploaded attachment files from storage
    """
    logger.info(f"Delete request for course_id={course_id} by user={current_user.id}")

    # 1. Verify course exists
    service = CourseService(db)
    try:
        course = await service.get_course(course_id)
    except OBEException as e:
        raise HTTPException(status_code=e.status_code, detail=e.message)

    storage = get_storage()

    # 2. Delete generated .docx from storage
    try:
        docx_path = CourseFileService.get_filepath(course_id)
        if os.path.exists(docx_path):
            os.remove(docx_path)
            logger.info(f"Deleted docx for course {course_id}: {docx_path}")
    except Exception as e:
        logger.warning(f"Could not delete docx for course {course_id}: {e}")

    # 3. Delete all attachment files from storage
    try:
        result = await db.execute(
            select(CourseFileAttachment).where(CourseFileAttachment.course_id == course_id)
        )
        attachments = result.scalars().all()
        for att in attachments:
            try:
                if att.stored_path and os.path.exists(att.stored_path):
                    os.remove(att.stored_path)
                    logger.info(f"Deleted attachment file: {att.stored_path}")
            except Exception as e:
                logger.warning(f"Could not delete attachment {att.id}: {e}")
    except Exception as e:
        logger.warning(f"Error fetching attachments for course {course_id}: {e}")

    # 4. Delete course DB row (cascades handle all child rows automatically)
    try:
        await db.delete(course)
        await db.commit()
        logger.info(f"Course {course_id} and all related DB data deleted.")
    except Exception as e:
        await db.rollback()
        logger.exception(f"Failed to delete course {course_id} from DB")
        raise HTTPException(status_code=500, detail="Failed to delete course from database")

    return {"status": "success", "message": f"Course {course_id} and all associated data deleted."}
