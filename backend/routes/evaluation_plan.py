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
from sqlalchemy import select

from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.database.connection import get_db
from backend.database.models import EvalPlanRow
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


@router.get("/view/{course_id}")
async def view_evaluation_plan(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Return saved evaluation plan rows as JSON. Reads from DB first, falls back to filesystem."""
    # 1. Try DB first (persistent)
    result = await db.execute(
        select(EvalPlanRow).where(EvalPlanRow.course_id == course_id)
    )
    row = result.scalar_one_or_none()
    if row and row.rows:
        return {"data": row.rows, "cols": row.cols}

    # 2. Fallback: old filesystem JSON (for data created before this migration)
    storage = get_storage()
    meta_filename = f"evaluation_plan_{course_id}{_META_SUFFIX}"
    fs_path = storage._dir(_CATEGORY) / meta_filename
    if fs_path.exists():
        data = json.loads(fs_path.read_text())
        rows = data.get("rows", data) if isinstance(data, dict) else data
        return {"data": rows}

    # 3. Legacy path
    legacy_path = os.path.join("generated_docs", "evaluation_plans", f"evaluation_plan_{course_id}_edited.json")
    if os.path.exists(legacy_path):
        with open(legacy_path) as jf:
            data = json.load(jf)
        return {"data": data.get("rows", data) if isinstance(data, dict) else data}

    raise HTTPException(status_code=404, detail="No evaluation plan found. Generate first.")


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
    """Saves the user-edited evaluation plan table to DB (and rebuilds docx)."""

    # 1. Save to DB (persistent — survives restarts)
    result = await db.execute(
        select(EvalPlanRow).where(EvalPlanRow.course_id == course_id)
    )
    ep_row = result.scalar_one_or_none()
    if ep_row:
        ep_row.rows = payload.rows
        ep_row.cols = payload.cols
    else:
        ep_row = EvalPlanRow(course_id=course_id)
        ep_row.rows = payload.rows
        ep_row.cols = payload.cols
        db.add(ep_row)
    await db.commit()

    # 2. Also write to filesystem for docx generation (best-effort)
    try:
        storage = get_storage()
        meta_filename = f"evaluation_plan_{course_id}{_META_SUFFIX}"
        meta_bytes = json.dumps({"cols": payload.cols, "rows": payload.rows}, ensure_ascii=False).encode()
        storage.save(_CATEGORY, meta_filename, meta_bytes)
    except Exception as e:
        logger.warning(f"Could not write eval plan JSON to filesystem: {e}")

    # 3. Rebuild docx from edited rows
    try:
        from backend.services.course_service import CourseService

        course_svc = CourseService(db)
        course = await course_svc.get_course(course_id)

        ca_components = []
        for row_data in payload.rows:
            ca_components.append({
                "sr_no":          row_data.get("sr_no", row_data.get("srNo", row_data.get("sr", ""))),
                "component":      row_data.get("component", row_data.get("comp", row_data.get("name", ""))),
                "unit_syllabus":  row_data.get("unit_syllabus", row_data.get("unitSyllabus", row_data.get("units", ""))),
                "co_mapped":      row_data.get("co", row_data.get("co_mapped", row_data.get("coMapped", ""))),
                "marks":          row_data.get("marks", row_data.get("total_marks", "")),
                "weightage":      row_data.get("weightage", ""),
                "tentative_date": row_data.get("date", row_data.get("tentative_date", row_data.get("tentativeDate", ""))),
            })

        plan_data = {"ca_components": ca_components}
        eval_cfg = {**course.evaluation_config, "credits": str(course.credits)}
        svc = EvaluationPlanService(db)
        _storage = get_storage()
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
            _storage=_storage,
            _filename=_filename,
        )
        logger.info(f"Evaluation plan saved to DB for course_id={course_id}: {len(payload.rows)} rows")
        return {
            "status": "success",
            "message": f"Saved {len(payload.rows)} rows",
            "download_url": f"/evaluation-plan/download/{course_id}",
        }
    except Exception as exc:
        logger.exception(f"Failed to rebuild eval plan docx for course_id={course_id}: {exc}")
        return {
            "status": "partial",
            "message": f"Data saved to DB but docx rebuild failed: {exc}",
        }
