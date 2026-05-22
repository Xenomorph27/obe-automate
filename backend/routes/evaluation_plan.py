# backend/routes/evaluation_plan.py
import json
import os
import tempfile
from pathlib import Path
from typing import Any, Dict, List

from fastapi import APIRouter, Depends, HTTPException
from backend.core.auth import require_auth
from backend.database.user_models import User
from fastapi.responses import FileResponse
from pydantic import BaseModel
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.database.connection import get_db
from backend.services.evaluation_plan_service import EvaluationPlanService

logger = get_logger(__name__)
router = APIRouter(prefix="/evaluation-plan", tags=["Evaluation Plan"])

_CATEGORY = "evaluation_plans"
_META_SUFFIX = "_edited.json"


class SavePlanRequest(BaseModel):
    cols: List[Dict[str, Any]]
    rows: List[Dict[str, Any]]


@router.post("/generate/{course_id}", status_code=201)
async def generate_evaluation_plan(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Generates a full CIE + SEE evaluation plan for the course using Gemini AI.
    Returns metadata and a download URL for the Word document.
    """
    logger.info(f"Evaluation plan generation requested for course_id={course_id}")
    try:
        service = EvaluationPlanService(db)
        result = await service.generate(course_id)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"Evaluation plan error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception:
        logger.exception(f"Unexpected error generating evaluation plan for course {course_id}")
        raise HTTPException(status_code=500, detail="Internal server error")


@router.get("/download/{course_id}")
async def download_evaluation_plan(course_id: int, current_user: User = Depends(require_auth)):
    """Download the generated evaluation plan Word document."""
    filepath = EvaluationPlanService.get_filepath(course_id)
    if not os.path.exists(filepath):
        raise HTTPException(
            status_code=404,
            detail=f"Evaluation plan not found for course_id={course_id}. Run POST /evaluation-plan/generate/{course_id} first.",
        )
    return FileResponse(
        path=filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"evaluation_plan_{course_id}.docx",
    )


@router.post("/save/{course_id}", status_code=200)
async def save_evaluation_plan(
    course_id: int,
    payload: SavePlanRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Saves the user-edited evaluation plan table (cols + rows).
    Persists as JSON and rebuilds a fresh .docx from the edited data.
    """
    storage = get_storage()

    # 1. Persist edited table as JSON
    meta_filename = f"evaluation_plan_{course_id}{_META_SUFFIX}"
    meta_bytes = json.dumps({"cols": payload.cols, "rows": payload.rows}, ensure_ascii=False).encode()
    storage.save(_CATEGORY, meta_filename, meta_bytes)

    # 2. Rebuild docx from edited rows using the proper SIT-formatted service
    try:
        from backend.services.course_service import CourseService
        from backend.services.evaluation_plan_service import EvaluationPlanService

        # Load course info
        course_svc = CourseService(db)
        course = await course_svc.get_course(course_id)

        # Convert edited rows back into ca_components structure
        ca_components = []
        for row_data in payload.rows:
            ca_components.append({
                "sr_no":          row_data.get("sr_no", row_data.get("srNo", "")),
                "component":      row_data.get("component", ""),
                "unit_syllabus":  row_data.get("unit_syllabus", row_data.get("unitSyllabus", "")),
                "co_mapped":      row_data.get("co", row_data.get("co_mapped", row_data.get("coMapped", ""))),
                "marks":          row_data.get("marks", ""),
                "weightage":      row_data.get("weightage", ""),
                "tentative_date": row_data.get("tentative_date", row_data.get("tentativeDate", "")),
            })

        plan_data = {"ca_components": ca_components}

        eval_cfg = {**course.evaluation_config, "credits": str(course.credits)}

        svc = EvaluationPlanService(db)
        _filename = f"evaluation_plan_{course_id}.docx"
        svc._build_docx(
            course_name=course.course_name,
            course_code=course.course_code,
            faculty_name=course.faculty_name,
            department=course.department,
            semester=course.semester,
            academic_year=course.academic_year,
            cos=course.cos,
            eval_cfg=eval_cfg,
            data=plan_data,
            _storage=storage,
            _filename=_filename,
        )

        logger.info(f"Evaluation plan saved (edited) for course_id={course_id}: {len(payload.rows)} rows")
        return {
            "status": "success",
            "message": f"Saved {len(payload.rows)} rows",
            "download_url": f"/evaluation-plan/download/{course_id}",
        }

    except Exception as exc:
        logger.exception(f"Failed to rebuild eval plan docx for course_id={course_id}: {exc}")
        return {
            "status": "partial",
            "message": f"Data saved but docx rebuild failed: {exc}",
        }