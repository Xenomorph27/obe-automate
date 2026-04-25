# backend/services/dashboard_service.py
"""
DashboardService
----------------
Aggregates CO/PO attainment across ALL courses in the department.
Used by the HOD Dashboard page.

GET /dashboard/department  → returns aggregated data for heatmap + summary
"""

from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from backend.core.logger import get_logger
from backend.database.models import Course, COAttainment
from backend.services.attainment_service import AttainmentService

logger = get_logger(__name__)


class DashboardService:

    def __init__(self, db: AsyncSession):
        self.db = db

    async def get_department_summary(self) -> dict:
        """
        Returns aggregated attainment data for all courses that have marks.
        """
        # Get all courses
        result = await self.db.execute(select(Course))
        courses = result.scalars().all()

        if not courses:
            return {
                "total_courses": 0,
                "courses_with_data": 0,
                "courses": [],
                "po_heatmap": {},
                "department_po_average": {},
                "summary": {
                    "avg_co_attainment": 0.0,
                    "courses_above_target": 0,
                    "total_students": 0,
                }
            }

        attainment_svc = AttainmentService(self.db)
        courses_data = []
        po_accumulator: dict[str, list[float]] = {}
        total_students = 0
        co_attainments = []
        courses_above_target = 0

        for course in courses:
            # Check if this course has marks
            marks_result = await self.db.execute(
                select(COAttainment).where(COAttainment.course_id == course.id).limit(1)
            )
            has_marks = marks_result.scalar() is not None

            if not has_marks:
                continue

            try:
                data = await attainment_svc.calculate(course.id)
            except Exception as e:
                logger.warning(f"Could not calculate attainment for course {course.id}: {e}")
                continue

            overall_co = data["overall_co_attainment"]
            co_attainments.append(overall_co)
            total_students += data["total_students"]
            if overall_co >= 60:
                courses_above_target += 1

            # Build PO row for heatmap
            po_row = {}
            for po_id, po_data in data["po_attainment"].items():
                pct = po_data["attainment_percentage"]
                po_row[po_id] = pct
                if po_id not in po_accumulator:
                    po_accumulator[po_id] = []
                po_accumulator[po_id].append(pct)

            courses_data.append({
                "course_id": course.id,
                "course_name": course.course_name,
                "course_code": course.course_code,
                "faculty_name": course.faculty_name,
                "semester": course.semester,
                "academic_year": course.academic_year,
                "total_students": data["total_students"],
                "overall_co_attainment": overall_co,
                "co_attainment": {
                    co_id: {
                        "attainment_percentage": v["attainment_percentage"],
                        "attainment_level": v["attainment_level"],
                        "target_met": v["target_met"],
                    }
                    for co_id, v in data["co_attainment"].items()
                },
                "po_attainment": po_row,
            })

        # Department-level PO averages
        dept_po_avg = {
            po_id: round(sum(vals) / len(vals), 2)
            for po_id, vals in po_accumulator.items()
        }

        avg_co = round(sum(co_attainments) / len(co_attainments), 2) if co_attainments else 0.0

        return {
            "total_courses": len(courses),
            "courses_with_data": len(courses_data),
            "courses": courses_data,
            "po_heatmap": {c["course_code"]: c["po_attainment"] for c in courses_data},
            "department_po_average": dept_po_avg,
            "summary": {
                "avg_co_attainment": avg_co,
                "courses_above_target": courses_above_target,
                "total_students": total_students,
            }
        }