# backend/routes/questions.py
from fastapi import APIRouter, Depends, Query
from backend.core.auth import require_auth
from backend.database.user_models import User
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession
from typing import Optional
from pydantic import BaseModel

from backend.core.logger import get_logger
from backend.database.connection import get_db
from backend.database.models import BLOOM_LEVELS
from backend.services.question_service import QuestionService
from backend.services.question_bank_service import QuestionBankService
from backend.services.question_classifier_service import QuestionClassifierService
from backend.services.question_manipulator_service import QuestionManipulatorService
from backend.services.question_paper_service import QuestionPaperService

logger = get_logger(__name__)
router = APIRouter(prefix="/questions", tags=["Questions"])

# ── Pydantic models ──────────────────────────────────────────────────────────

class GenerateSingleRequest(BaseModel):
    course_id: int
    course_name: str
    topic: str
    bloom_level: int
    question_type: str = "Short Answer"
    co_id: Optional[str] = None
    marks: int = 5
    extra_instructions: str = ""
    save_to_bank: bool = True

class GeneratePaperRequest(BaseModel):
    course_id: int
    course_name: str
    course_code: str
    duration_hours: int = 3
    total_marks: int = 100
    cos: list
    sections: list

class ClassifyRequest(BaseModel):
    course_id: int
    course_name: str
    question_text: str
    cos: list
    pos: list
    save_to_bank: bool = True

class ManipulateRequest(BaseModel):
    course_id: int
    course_name: str
    original_question: str
    original_bloom_level: int
    target_bloom_level: int
    topic: str = ""
    teacher_hint: str = ""

class SaveManipulatedRequest(BaseModel):
    course_id: int
    question_text: str
    topic: str = ""
    bloom_level: int = 1
    co_id: Optional[str] = None
    po_id: Optional[str] = None
    marks: int = 5
    question_type: str = "Short Answer"
    parent_question_id: Optional[int] = None

class DeleteRequest(BaseModel):
    question_id: int

# ── Routes ───────────────────────────────────────────────────────────────────

@router.post("/generate-single")
async def generate_single(req: GenerateSingleRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(require_auth)):
    svc = QuestionService()
    q = await svc.generate_single(
        topic=req.topic, bloom_level=req.bloom_level,
        question_type=req.question_type, course_name=req.course_name,
        co_id=req.co_id, marks=req.marks,
        extra_instructions=req.extra_instructions,
    )
    q["course_id"] = req.course_id
    saved_id = None
    if req.save_to_bank:
        bank = QuestionBankService(db)
        saved = await bank.save_question({**q, "source": "generated"})
        saved_id = saved.id
    return {"question": q, "saved_id": saved_id, "saved_to_bank": req.save_to_bank}


@router.post("/generate-paper/{course_id}")
async def generate_paper(course_id: int, req: GeneratePaperRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(require_auth)):
    svc = QuestionService()
    questions = await svc.generate_paper(
        course_name=req.course_name, course_code=req.course_code,
        cos=req.cos, total_marks=req.total_marks,
        duration_hours=req.duration_hours, sections=req.sections,
    )
    # Save all to bank
    bank = QuestionBankService(db)
    for q in questions:
        q["course_id"] = course_id
        q["source"] = "generated"
        await bank.save_question(q)
    # Generate documents
    paper_svc = QuestionPaperService()
    paths = paper_svc.generate_documents(
        course_id=course_id, course_name=req.course_name,
        course_code=req.course_code, duration=req.duration_hours,
        total_marks=req.total_marks, questions=questions,
    )
    return {"questions": questions, "total": len(questions), "documents": paths}


@router.get("/bank/{course_id}")
async def get_bank(
    course_id: int,
    bloom_level: Optional[int] = Query(None),
    co_id: Optional[str] = Query(None),
    question_type: Optional[str] = Query(None),
    source: Optional[str] = Query(None),
    db: AsyncSession = Depends(get_db),
):
    bank = QuestionBankService(db)
    questions = await bank.get_bank(course_id, bloom_level, co_id, question_type, source)
    stats = await bank.get_stats(course_id)
    return {"questions": questions, "stats": stats, "total": len(questions)}


@router.delete("/bank/{question_id}")
async def delete_question(question_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(require_auth)):
    bank = QuestionBankService(db)
    return await bank.delete_question(question_id)


@router.post("/classify")
async def classify_question(req: ClassifyRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(require_auth)):
    svc = QuestionClassifierService()
    result = await svc.classify(
        question_text=req.question_text, course_name=req.course_name,
        cos=req.cos, pos=req.pos,
    )
    saved_id = None
    if req.save_to_bank:
        bank = QuestionBankService(db)
        saved = await bank.save_question({
            "course_id": req.course_id,
            "question_text": req.question_text,
            "bloom_level": result["bloom_level"],
            "co_id": result["co_id"],
            "po_id": result["po_id"],
            "question_type": result["question_type"],
            "source": "classified",
        })
        saved_id = saved.id
    return {"classification": result, "saved_id": saved_id}


@router.post("/manipulate")
async def manipulate_question(req: ManipulateRequest, current_user: User = Depends(require_auth)):
    svc = QuestionManipulatorService()
    result = await svc.manipulate(
        original_question=req.original_question,
        original_bloom_level=req.original_bloom_level,
        target_bloom_level=req.target_bloom_level,
        course_name=req.course_name,
        topic=req.topic, teacher_hint=req.teacher_hint,
    )
    return result


@router.post("/manipulate/save")
async def save_manipulated(req: SaveManipulatedRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(require_auth)):
    bank = QuestionBankService(db)
    saved = await bank.save_question({
        "course_id": req.course_id,
        "question_text": req.question_text,
        "topic": req.topic,
        "bloom_level": req.bloom_level,
        "co_id": req.co_id,
        "po_id": req.po_id,
        "marks": req.marks,
        "question_type": req.question_type,
        "source": "manipulated",
        "parent_question_id": req.parent_question_id,
    })
    return saved.to_dict()


@router.get("/paper/download/{course_id}/docx")
async def download_paper_docx(course_id: int, current_user: User = Depends(require_auth)):
    svc = QuestionPaperService()
    paths = svc.get_latest_paths(course_id)
    if not paths["docx_path"]:
        from backend.core.exceptions import OBEException
        raise OBEException("No question paper found. Generate one first.", 404)
    return FileResponse(paths["docx_path"], filename=f"question_paper_{course_id}.docx",
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@router.get("/paper/download/{course_id}/pdf")
async def download_paper_pdf(course_id: int, current_user: User = Depends(require_auth)):
    svc = QuestionPaperService()
    paths = svc.get_latest_paths(course_id)
    if not paths["pdf_path"]:
        from backend.core.exceptions import OBEException
        raise OBEException("No question paper found. Generate one first.", 404)
    return FileResponse(paths["pdf_path"], filename=f"question_paper_{course_id}.pdf",
                        media_type="application/pdf")


@router.get("/bloom-levels")
async def get_bloom_levels(current_user: User = Depends(require_auth)):
    return [{"level": k, "label": v} for k, v in BLOOM_LEVELS.items()]
