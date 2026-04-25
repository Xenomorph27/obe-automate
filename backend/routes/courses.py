# backend/routes/courses.py
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from backend.database.connection import get_db
from backend.services.course_service import CourseService
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from pydantic import BaseModel, Field
from typing import Dict, List

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
    co_po_matrix: Dict[str, Dict[str, int]] = Field(
        description="CO-PO mapping. Keys are CO IDs, values are dicts of PO ID → correlation (0,1,2,3)",
        example={"CO1": {"PO1": 3, "PO2": 2}, "CO2": {"PO1": 1, "PO2": 3}}
    )
    evaluation_config: EvaluationConfig


# --- Routes ---

@router.post("/setup", status_code=201)
async def setup_course(
    request: CourseSetupRequest,
    db: AsyncSession = Depends(get_db)
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
    db: AsyncSession = Depends(get_db)
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
async def list_courses(db: AsyncSession = Depends(get_db)):
    """List all courses."""
    try:
        service = CourseService(db)
        courses = await service.list_courses()
        return {"status": "success", "count": len(courses), "data": courses}
    except Exception:
        logger.exception("Error listing courses")
        raise HTTPException(status_code=500, detail="Internal server error")