# backend/routes/session_plan.py
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
from backend.database.models import SessionPlanRow
from backend.services.session_plan_service import SessionPlanService

logger = get_logger(__name__)
router = APIRouter(prefix="/session-plan", tags=["Session Plan"])

_CATEGORY = "session_plans"
_META_SUFFIX = "_edited.json"


class SavePlanRequest(BaseModel):
    cols: List[Dict[str, Any]]
    rows: List[Dict[str, Any]]


@router.post("/generate/{course_id}", status_code=201)
async def generate_session_plan(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
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
async def download_session_plan(course_id: int, current_user: User = Depends(require_auth)):
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


@router.get("/view/{course_id}")
async def view_session_plan(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Return saved session plan rows as JSON. Reads from DB first, then filesystem fallback."""
    # 1. Try DB first (persistent — survives restarts)
    result = await db.execute(
        select(SessionPlanRow).where(SessionPlanRow.course_id == course_id)
    )
    row = result.scalar_one_or_none()
    if row and row.rows:
        return {"data": row.rows, "cols": row.cols}

    # 2. Fallback: filesystem JSON (old data before this migration)
    storage = get_storage()
    meta_filename = f"session_plan_{course_id}{_META_SUFFIX}"
    try:
        fs_path = storage._dir(_CATEGORY) / meta_filename
        if fs_path.exists():
            data = json.loads(fs_path.read_text())
            rows = data.get("rows", data) if isinstance(data, dict) else data
            return {"data": rows, "cols": data.get("cols", []) if isinstance(data, dict) else []}
    except Exception:
        pass

    # 3. Legacy path
    legacy_path = os.path.join("generated_docs", "session_plans", f"session_plan_{course_id}_edited.json")
    if os.path.exists(legacy_path):
        with open(legacy_path) as jf:
            data = json.load(jf)
        rows = data.get("rows", data) if isinstance(data, dict) else data
        return {"data": rows, "cols": data.get("cols", []) if isinstance(data, dict) else []}

    raise HTTPException(status_code=404, detail="No session plan found. Generate first.")


@router.post("/save/{course_id}", status_code=200)
async def save_session_plan(
    course_id: int,
    payload: SavePlanRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Save the user-edited session plan to DB (and filesystem for docx rebuild)."""
    storage = get_storage()

    # 1. Save to DB (persistent — survives restarts and Render sleep/wake)
    result = await db.execute(
        select(SessionPlanRow).where(SessionPlanRow.course_id == course_id)
    )
    sp_row = result.scalar_one_or_none()
    if sp_row:
        sp_row.rows = payload.rows
        sp_row.cols = payload.cols
    else:
        sp_row = SessionPlanRow(course_id=course_id)
        sp_row.rows = payload.rows
        sp_row.cols = payload.cols
        db.add(sp_row)
    await db.commit()

    # 2. Also write filesystem JSON (best-effort, for docx rebuild)
    try:
        meta_filename = f"session_plan_{course_id}{_META_SUFFIX}"
        meta_bytes = json.dumps({"cols": payload.cols, "rows": payload.rows}, ensure_ascii=False).encode()
        storage.save(_CATEGORY, meta_filename, meta_bytes)
    except Exception as e:
        logger.warning(f"Could not write session plan JSON to filesystem: {e}")

    # 3. Rebuild docx from the edited rows
    try:
        from backend.services.course_service import CourseService
        from backend.services.session_plan_service import SessionPlanService

        course_svc = CourseService(db)
        course = await course_svc.get_course(course_id)

        units_dict: dict = {}
        for row_data in payload.rows:
            unit_no = row_data.get("unit_no", row_data.get("unitNo", ""))
            topic   = row_data.get("points_to_cover", row_data.get("pointsToCover", row_data.get("topic", "")))
            method  = row_data.get("methodology", "Classroom Teaching")
            stype   = row_data.get("lecture_exp_eval", row_data.get("lectureExpEval", row_data.get("type", "Lecture")))
            co      = row_data.get("co", row_data.get("co_mapped", ""))
            lect_no = row_data.get("lect_no", row_data.get("lectNo", ""))

            key = str(unit_no) if unit_no else "_misc"
            if key not in units_dict:
                units_dict[key] = {"unit_number": unit_no, "unit_title": f"Unit {unit_no}", "sessions": []}
            units_dict[key]["sessions"].append({
                "session_number": lect_no,
                "topic": topic,
                "teaching_method": method,
                "type": stype,
                "co_mapped": co,
            })

        plan_data = {"units": list(units_dict.values())}
        svc = SessionPlanService(db)
        _filename = f"session_plan_{course_id}.docx"
        svc._build_docx(
            course_name=course.course_name,
            course_code=course.course_code,
            faculty_name=course.faculty_name,
            department=course.department,
            semester=course.semester,
            academic_year=course.academic_year,
            credits=course.credits,
            cos=course.cos,
            data=plan_data,
            _storage=storage,
            _filename=_filename,
        )

        logger.info(f"Session plan saved to DB for course_id={course_id}: {len(payload.rows)} rows")
        return {
            "status": "success",
            "message": f"Saved {len(payload.rows)} rows",
            "download_url": f"/session-plan/download/{course_id}",
        }

    except Exception as exc:
        logger.exception(f"Failed to rebuild docx for course_id={course_id}: {exc}")
        return {
            "status": "partial",
            "message": f"Data saved to DB but docx rebuild failed: {exc}",
        }


@router.get("/materials/{course_id}")
async def get_session_materials(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Returns study materials extracted from the session plan."""
    textbooks: list = []
    web_links: list = []
    journals: list = []
    moocs: list = []

    # Try DB first
    result = await db.execute(
        select(SessionPlanRow).where(SessionPlanRow.course_id == course_id)
    )
    sp_row = result.scalar_one_or_none()
    rows = []
    cols = []
    if sp_row and sp_row.rows:
        rows = sp_row.rows
        cols = sp_row.cols
    else:
        # Fallback to filesystem
        storage = get_storage()
        meta_filename = f"session_plan_{course_id}{_META_SUFFIX}"
        try:
            fs_path = storage._dir(_CATEGORY) / meta_filename
            if fs_path.exists():
                data = json.loads(fs_path.read_text(encoding="utf-8"))
                rows = data.get("rows", [])
                cols = data.get("cols", [])
        except Exception as e:
            logger.warning(f"Could not parse session plan JSON: {e}")

    if rows and cols:
        for col in cols:
            label = col.get("label", "").lower()
            key = col.get("key", "")
            if any(k in label for k in ["textbook", "reference", "book"]):
                textbooks += [{"title": r.get(key, ""), "author": "", "publisher": ""} for r in rows if r.get(key)]
            elif any(k in label for k in ["web", "link", "nptel", "url", "online"]):
                web_links += [{"title": r.get(key, ""), "unit": r.get("unit", ""), "url": ""} for r in rows if r.get(key)]
            elif any(k in label for k in ["journal", "paper", "research"]):
                journals += [{"title": r.get(key, ""), "url": ""} for r in rows if r.get(key)]
            elif any(k in label for k in ["mooc", "course", "swayam", "coursera"]):
                moocs += [{"title": r.get(key, ""), "platform": "", "url": ""} for r in rows if r.get(key)]

    if not any([textbooks, web_links, journals, moocs]):
        raise HTTPException(
            status_code=404,
            detail="No study materials found. Generate the session plan first, then add material columns via AI chat.",
        )

    return {
        "status": "success",
        "data": {
            "textbooks": textbooks,
            "web_links": web_links,
            "journals": journals,
            "mooc_courses": moocs,
        },
    }
