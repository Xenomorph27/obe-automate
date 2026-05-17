# backend/services/course_service.py
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from backend.database.models import Course
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger

logger = get_logger(__name__)


class CourseService:

    def __init__(self, db: AsyncSession):
        self.db = db

    async def create_course(self, request) -> Course:
        """Saves a new course to the database."""

        # Check if course code already exists for this academic year
        existing = await self.db.execute(
            select(Course).where(
                Course.course_code == request.course_code,
                Course.academic_year == request.academic_year
            )
        )
        if existing.scalar_one_or_none():
            raise OBEException(
                f"Course {request.course_code} already exists for {request.academic_year}",
                status_code=409
            )

        course = Course(
            course_name=request.course_name,
            course_code=request.course_code,
            credits=request.credits,
            total_hours=request.total_hours,
            faculty_name=request.faculty_name,
            department=request.department,
            semester=request.semester,
            academic_year=request.academic_year,
        )

        # Use the property setters — they handle JSON serialisation
        course.cos = [co.dict() for co in request.cos]
        course.pos = [po.dict() for po in request.pos]
        course.co_po_matrix = request.co_po_matrix
        course.evaluation_config = request.evaluation_config.dict()

        self.db.add(course)
        await self.db.commit()
        await self.db.refresh(course)

        logger.info(f"Course created: ID={course.id}, code={course.course_code}")
        return course

    async def get_course(self, course_id: int) -> Course:
        """Fetch a single course by ID."""
        result = await self.db.execute(
            select(Course).where(Course.id == course_id)
        )
        course = result.scalar_one_or_none()
        if not course:
            raise OBEException(f"Course ID {course_id} not found", status_code=404)
        return course

    async def list_courses(self) -> list:
        """Return all courses as dicts."""
        result = await self.db.execute(select(Course))
        courses = result.scalars().all()
        return [c.to_dict() for c in courses]
    async def update_course(self, course_id: int, request) -> Course:
        """Partially update an existing course — only non-None fields are changed."""
        course = await self.get_course(course_id)

        if request.course_name is not None:
            course.course_name = request.course_name
        if request.course_code is not None:
            course.course_code = request.course_code
        if request.credits is not None:
            course.credits = request.credits
        if request.total_hours is not None:
            course.total_hours = request.total_hours
        if request.faculty_name is not None:
            course.faculty_name = request.faculty_name
        if request.department is not None:
            course.department = request.department
        if request.semester is not None:
            course.semester = request.semester
        if request.academic_year is not None:
            course.academic_year = request.academic_year
        if request.cos is not None:
            course.cos = [co.dict() for co in request.cos]
        if request.pos is not None:
            course.pos = [po.dict() for po in request.pos]
        if request.co_po_matrix is not None:
            course.co_po_matrix = request.co_po_matrix
        if request.evaluation_config is not None:
            course.evaluation_config = request.evaluation_config.dict()

        await self.db.commit()
        await self.db.refresh(course)

        logger.info(f"Course updated: ID={course.id}, code={course.course_code}")
        return course
