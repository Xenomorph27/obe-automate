# backend/routes/co_po_template.py
"""
Routes for generating and downloading the CO-PO Attainment Excel template.

POST /co-po-template/generate/{course_id}   — generate the workbook
GET  /co-po-template/download/{course_id}   — download the xlsx
"""
import os
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.auth import require_auth
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.database.connection import get_db
from backend.database.user_models import User
from backend.services.co_po_template_service import COPOTemplateService

logger = get_logger(__name__)
router = APIRouter(prefix="/co-po-template", tags=["CO-PO Template"])

_CATEGORY = "co_po_templates"


class GenerateRequest(BaseModel):
    qp_source: str = "blank"   # "blank" | "question_bank"


@router.post("/generate/{course_id}", status_code=201)
async def generate_co_po_template(
    course_id: int,
    body: GenerateRequest = GenerateRequest(),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Generate the CO-PO attainment Excel workbook for a course.

    qp_source options:
    - "blank"         : Leave question paper sheets empty (faculty fills manually)
    - "question_bank" : Pre-fill question paper sheets from the course question bank
    """
    logger.info(f"CO-PO template generation requested for course_id={course_id}, source={body.qp_source}")
    try:
        svc = COPOTemplateService(db)
        result = await svc.generate(course_id, qp_source=body.qp_source)
        return {"status": "success", "data": result}
    except OBEException as e:
        logger.error(f"CO-PO template error: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)
    except Exception as e:
        logger.exception(f"Unexpected error generating CO-PO template for course {course_id}")
        raise HTTPException(status_code=500, detail=f"Internal server error: {str(e)}")


@router.get("/download/{course_id}")
async def download_co_po_template(
    course_id: int,
    current_user: User = Depends(require_auth),
):
    """Download the generated CO-PO attainment Excel workbook."""
    filepath = COPOTemplateService.get_filepath(course_id)
    if not os.path.exists(filepath):
        raise HTTPException(
            status_code=404,
            detail=f"CO-PO template not found for course_id={course_id}. "
                   f"Run POST /co-po-template/generate/{course_id} first.",
        )
    return FileResponse(
        path=filepath,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"CO_PO_Attainment_{course_id}.xlsx",
    )


@router.post("/upload-qp/{course_id}/{ca_label}", status_code=200)
async def upload_question_paper(
    course_id: int,
    ca_label: str,
    file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """
    Upload a question paper PDF/XLSX for a specific CA and parse questions from it.
    Extracted questions are stored in the question bank with source='uploaded'.
    Returns the list of parsed questions.
    """
    if not file.filename.endswith((".xlsx", ".xls", ".pdf")):
        raise HTTPException(400, "Only .xlsx, .xls, or .pdf files supported")

    file_bytes = await file.read()
    questions  = []

    if file.filename.endswith((".xlsx", ".xls")):
        try:
            import io, openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            # Find header row
            hdr_idx = None
            for i, row in enumerate(rows):
                vals = [str(v).strip().lower() if v else "" for v in row]
                if "question" in vals or "q. no" in " ".join(vals):
                    hdr_idx = i
                    break
            if hdr_idx is not None:
                for row in rows[hdr_idx+1:]:
                    if not any(row):
                        continue
                    q_text = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                    marks  = row[6] if len(row) > 6 and row[6] else 0
                    co     = str(row[7]).strip() if len(row) > 7 and row[7] else ""
                    bl_raw = str(row[8]).strip() if len(row) > 8 and row[8] else "L1"
                    bl     = int(bl_raw.replace("L", "")) if bl_raw.startswith("L") else 1
                    if q_text and q_text not in ("Question", ""):
                        questions.append({
                            "question_text": q_text,
                            "marks": int(marks) if marks else 5,
                            "co_id": co,
                            "bloom_level": bl,
                            "source": "uploaded",
                            "ca_label": ca_label,
                        })
        except Exception as e:
            raise HTTPException(400, f"Failed to parse xlsx: {str(e)}")

    # Save parsed questions to question bank
    if questions:
        from sqlalchemy import text as _text
        from backend.database.models import Question, BLOOM_LEVELS
        for q in questions:
            bl = q["bloom_level"]
            new_q = Question(
                course_id=course_id,
                question_text=q["question_text"],
                marks=q["marks"],
                co_id=q["co_id"],
                bloom_level=bl,
                bloom_label=BLOOM_LEVELS.get(bl, "Remember"),
                source="uploaded",
                question_type="Short Answer",
            )
            db.add(new_q)
        await db.commit()

    return {
        "status":    "success",
        "ca_label":  ca_label,
        "extracted": len(questions),
        "questions": questions,
        "message":   f"Parsed {len(questions)} questions from {file.filename}. "
                     f"Re-generate the template with qp_source='question_bank' to include them.",
    }
