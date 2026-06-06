# backend/routes/co_po_template.py
"""
Routes for generating and downloading the CO-PO Attainment Excel template.

POST /co-po-template/generate/{course_id}         — generate the workbook
GET  /co-po-template/download/{course_id}         — download the xlsx
POST /co-po-template/save-sheet/{course_id}/{ca}  — save QP + marks for one CA to DB
GET  /co-po-template/load-all-sheets/{course_id}  — load all CA sheets from DB
"""
import os
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Any, Dict, List
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from backend.core.auth import require_auth
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.database.connection import get_db
from backend.database.user_models import User
from backend.database.models import CASheet
from backend.services.co_po_template_service import COPOTemplateService

logger = get_logger(__name__)
router = APIRouter(prefix="/co-po-template", tags=["CO-PO Template"])

_CATEGORY = "co_po_templates"


class GenerateRequest(BaseModel):
    qp_source: str = "blank"   # "blank" | "question_bank"


class SaveSheetRequest(BaseModel):
    qp: List[Dict[str, Any]] = []
    marks: Dict[str, Any] = {}


@router.post("/generate/{course_id}", status_code=201)
async def generate_co_po_template(
    course_id: int,
    body: GenerateRequest = GenerateRequest(),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
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


@router.post("/save-sheet/{course_id}/{ca_label}", status_code=200)
async def save_sheet(
    course_id: int,
    ca_label: str,
    payload: SaveSheetRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Save QP + marks for one CA component to the database."""
    from urllib.parse import unquote
    ca_label = unquote(ca_label)

    result = await db.execute(
        select(CASheet).where(CASheet.course_id == course_id, CASheet.ca_label == ca_label)
    )
    sheet = result.scalar_one_or_none()

    if sheet:
        sheet.qp = payload.qp
        sheet.marks = payload.marks
    else:
        sheet = CASheet(course_id=course_id, ca_label=ca_label)
        sheet.qp = payload.qp
        sheet.marks = payload.marks
        db.add(sheet)

    await db.commit()
    logger.info(f"Saved CA sheet for course={course_id} ca={ca_label}: {len(payload.qp)} questions, {len(payload.marks)} students")
    return {"status": "success", "ca_label": ca_label, "questions": len(payload.qp), "students": len(payload.marks)}


@router.get("/load-all-sheets/{course_id}")
async def load_all_sheets(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_auth),
):
    """Load all saved CA sheets for a course from the database."""
    result = await db.execute(
        select(CASheet).where(CASheet.course_id == course_id)
    )
    sheets = result.scalars().all()
    data = {s.ca_label: {"qp": s.qp, "marks": s.marks} for s in sheets}
    return {"sheets": data}


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
            hdr_idx = None
            for i, row in enumerate(rows):
                vals = [str(v).strip().lower() if v else "" for v in row]
                if "question" in vals or "q. no" in " ".join(vals):
                    hdr_idx = i
                    break
            if hdr_idx is not None:
                hdr_row = rows[hdr_idx]
                # Dynamically find column indices from header row
                import re as _re
                def _ci(keywords):
                    for ki, kw in enumerate(keywords):
                        for ci, h in enumerate(hdr_row):
                            if h and kw.lower() in str(h).lower():
                                return ci
                    return -1
                q_no_ci  = _ci(["q. no", "q no", "q#", "qno"])
                q_txt_ci = _ci(["question text", "question"])
                marks_ci = _ci(["marks"])
                co_ci    = _ci(["co mapped", "co"])
                bl_ci    = _ci(["bloom"])
                q_counter = 0
                for row in rows[hdr_idx+1:]:
                    if not any(row):
                        continue
                    q_text = str(row[q_txt_ci]).strip() if q_txt_ci >= 0 and q_txt_ci < len(row) and row[q_txt_ci] else ""
                    if not q_text or q_text.lower() in ("question", "question text", "total", ""):
                        continue
                    if str(row[0] if row else "").strip().lower() in ("total", "max", "max marks", ""):
                        if not q_text:
                            continue
                    raw_qno = row[q_no_ci] if q_no_ci >= 0 and q_no_ci < len(row) else None
                    try:
                        q_no = int(float(str(raw_qno))) if raw_qno is not None and str(raw_qno).strip() != "" else None
                    except (ValueError, TypeError):
                        q_no = None
                    if q_no is None:
                        q_counter += 1
                        q_no = q_counter
                    marks_val = row[marks_ci] if marks_ci >= 0 and marks_ci < len(row) else None
                    try:
                        marks = float(marks_val) if marks_val is not None else 5
                    except (TypeError, ValueError):
                        marks = 5
                    co_val = row[co_ci] if co_ci >= 0 and co_ci < len(row) else ""
                    co = str(co_val).strip() if co_val else ""
                    bl_raw = str(row[bl_ci]).strip() if bl_ci >= 0 and bl_ci < len(row) and row[bl_ci] else "L1"
                    # Extract L1-L6
                    bl_match = _re.search(r'L([1-6])', bl_raw, _re.IGNORECASE)
                    bl = int(bl_match.group(1)) if bl_match else 1
                    questions.append({
                        "question_text": q_text,
                        "marks": marks,
                        "co_id": co,
                        "bloom_level": bl,
                        "q_no": q_no,
                        "source": "uploaded",
                        "ca_label": ca_label,
                    })
        except Exception as e:
            raise HTTPException(400, f"Failed to parse xlsx: {str(e)}")

    if questions:
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


class AICOPORequest(BaseModel):
    cos: list  # list of {co_id, co_statement}
    pos: list  # list of {po_id, po_statement}
    psos: list = []  # list of {pso_id, pso_statement}


@router.post("/ai-mapping/{course_id}", status_code=200)
async def ai_co_po_mapping(
    course_id: int,
    body: AICOPORequest,
    current_user: User = Depends(require_auth),
):
    """
    Generate CO-PO mapping values (0/1/2/3) using the existing LLM fallback chain
    (Gemini → Groq → OpenAI — whichever key is configured).
    Returns {co_id: {po_id: strength}} for each CO.
    """
    import json as _json
    from backend.core.llm import get_llm_response
    from backend.core.exceptions import LLMError

    co_text = "\n".join(
        f"- {c['co_id']}: {c.get('co_statement', c.get('description', c['co_id']))}"
        for c in body.cos
    )
    # Use full NBA standard PO descriptions regardless of what frontend sends
    # This gives the AI real content to reason against instead of empty labels
    _NBA_PO = {
        "PO1":  "Engineering Knowledge: Apply knowledge of mathematics, science, engineering fundamentals to solve complex engineering problems.",
        "PO2":  "Problem Analysis: Identify, formulate, review research literature, and analyze complex engineering problems to reach substantiated conclusions.",
        "PO3":  "Design/Development of Solutions: Design solutions for complex engineering problems and design systems/components/processes that meet specifications with societal and environmental considerations.",
        "PO4":  "Conduct Investigations of Complex Problems: Use research-based knowledge and methods including design of experiments, analysis and interpretation of data to provide valid conclusions.",
        "PO5":  "Modern Tool Usage: Create, select, and apply appropriate techniques, resources, and modern engineering and IT tools including prediction and modelling for complex activities.",
        "PO6":  "The Engineer and Society: Apply reasoning informed by contextual knowledge to assess societal, health, safety, legal and cultural issues relevant to professional engineering practice.",
        "PO7":  "Environment and Sustainability: Understand the impact of professional engineering solutions in societal and environmental contexts and demonstrate knowledge of sustainable development.",
        "PO8":  "Ethics: Apply ethical principles and commit to professional ethics, responsibilities, and norms of engineering practice.",
        "PO9":  "Individual and Team Work: Function effectively as an individual, and as a member or leader in diverse teams and multidisciplinary settings.",
        "PO10": "Communication: Communicate effectively on complex engineering activities with engineering community and society at large — reports, documentation, presentations.",
        "PO11": "Project Management and Finance: Demonstrate knowledge and understanding of engineering and management principles and apply them to manage projects in multidisciplinary environments.",
        "PO12": "Life-long Learning: Recognize the need for and engage in independent and life-long learning in the broadest context of technological change.",
    }
    po_text = "\n".join(
        f"- {p['po_id']}: {_NBA_PO.get(p['po_id'], p.get('po_statement', p.get('description', p.get('po_id', ''))))}"
        for p in body.pos
    )
    pso_text = "\n".join(
        f"- {p['pso_id']}: {p.get('pso_statement', p.get('description', ''))}"
        for p in body.psos
    ) if body.psos else ""

    po_ids = [p["po_id"] for p in body.pos]
    pso_ids = [p["pso_id"] for p in body.psos] if body.psos else []
    all_po_ids = po_ids + pso_ids
    co_ids = [c["co_id"] for c in body.cos]

    co_ids_str = ", ".join(co_ids)
    po_ids_str = ", ".join(all_po_ids)
    pso_block  = ("Program Specific Outcomes (PSOs):\n" + pso_text) if pso_text else ""

    # Build a CO-by-CO analysis section so the LLM reasons per CO, not in a pattern
    co_analysis_lines = []
    for co in body.cos:
        cid  = co["co_id"]
        stmt = co.get("co_statement", co.get("description", ""))
        co_analysis_lines.append(f"{cid}: \"{stmt}\"")
        co_analysis_lines.append(f"  -> For each PO ask: Does this CO explicitly require or develop the skill/knowledge described by that PO?")
        co_analysis_lines.append(f"  -> Assign 3 only if strongly and directly, 2 if moderately, 1 if peripherally, 0 if unrelated.")
        co_analysis_lines.append("")
    co_analysis = "\n".join(co_analysis_lines)

    prompt = f"""You are a senior NBA/NAAC accreditation expert. Your job is to produce a justified, academically rigorous CO-PO mapping for an engineering course.

=== COURSE OUTCOMES ===
{co_text}

=== PROGRAM OUTCOMES ===
{po_text}
{pso_block}

=== YOUR TASK ===
For every CO, assign a mapping strength (0/1/2/3) to every PO by following this exact reasoning process:

STEP 1 — Read each CO statement carefully. Identify the ACTION VERB and SUBJECT MATTER.
  - "Explain" = knowledge recall and comprehension → strong PO1, moderate PO2
  - "Describe" = conceptual understanding → strong PO1, moderate PO2
  - "Discuss" = analysis and evaluation → strong PO1, strong PO2
  - "Design/Implement/Develop" = solution building → strong PO1, PO2, PO3
  - "Investigate/Analyze/Compare" = research skills → PO1, PO2, PO4
  - "Apply/Use tools/software" = modern tools → PO5
  - "Understand deep learning/neural networks" = tools + design → PO1, PO3, PO5

STEP 2 — For each PO, ask: "Does THIS CO's content require THIS skill to achieve it?"
  PO1 (Engineering Knowledge): Is domain knowledge required? → Yes for almost all technical COs = 3
  PO2 (Problem Analysis): Does the CO involve analyzing, comparing, or reasoning about systems? → 2-3 if yes
  PO3 (Design Solutions): Does the CO involve building, designing, or implementing something? → 2-3 if yes
  PO4 (Investigation): Does the CO involve experiments, data analysis, or research methods? → 1-2 if yes
  PO5 (Modern Tool Usage): Does achieving this CO require using software, frameworks, or computational tools? → 1-3 if yes
  PO6 (Engineer & Society): Does the CO explicitly address societal impact? → 0 unless stated
  PO7 (Environment): Does the CO address sustainability or environmental concerns? → 0 unless stated
  PO8 (Ethics): Does the CO involve ethical decision-making? → 0 unless stated
  PO9 (Individual/Team Work): Does the CO involve collaborative work explicitly? → 0 unless stated
  PO10 (Communication): Does the CO require producing reports, presentations, or documentation? → 0 unless stated
  PO11 (Project Management): Does the CO involve managing resources or a project? → 0 unless stated
  PO12 (Life-long Learning): All COs support ongoing learning = 1 universally

STEP 3 — MANDATORY COVERAGE CHECK before finalizing:
  - Every CO must map to AT LEAST 2-3 POs with value > 0
  - PO1 should be 3 for almost every technical CO
  - PO12 must be 1 for every CO
  - PO2 should be ≥1 for any CO with analytical verbs (explain, describe, discuss, analyze)
  - PO5 should be ≥1 for any CO involving algorithms, deep learning, or computational methods
  - PO6, PO7, PO8, PO9, PO10, PO11 = 0 unless the CO statement explicitly mentions those domains

STEP 4 — ANTI-PATTERN CHECK:
  WRONG: CO1→PO1 only, CO2→PO2 only, CO3→PO3 only (diagonal — forbidden)
  WRONG: All COs identical mapping (copy-paste — forbidden)
  RIGHT: Each CO has a unique profile based on its verb and subject matter

=== ANALYSIS FOR THIS COURSE ===
{co_analysis}

=== OUTPUT ===
Return ONLY a valid JSON object. No markdown, no explanation, no code fences.
Every CO must list every PO. Values must be 0, 1, 2, or 3.
{{
  "CO1": {{"PO1": 3, "PO2": 2, "PO3": 0, "PO4": 0, "PO5": 1, "PO6": 0, "PO7": 0, "PO8": 0, "PO9": 0, "PO10": 0, "PO11": 0, "PO12": 1}},
  ...
}}
COs required: {co_ids_str}
POs required: {po_ids_str}"""

    # ── Step 1: Rule engine (deterministic, always runs) ────────────────────
    from backend.services.co_po_engine import rule_engine, merge_rule_ai
    rule_map = rule_engine(body.cos, all_po_ids)
    logger.info(f"Rule engine mapping for course_id={course_id}: {list(rule_map.keys())}")

    # ── Step 2: AI layer (validates + fills gaps) ────────────────────────────
    ai_map = {}
    try:
        text = await get_llm_response(prompt)

        text = text.strip()
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.lstrip().startswith("json"):
                text = text.lstrip()[4:]
        text = text.strip()

        start = text.find("{")
        end   = text.rfind("}") + 1
        if start == -1 or end == 0:
            raise ValueError("No JSON object in LLM response")
        raw_ai = _json.loads(text[start:end])

        # Normalise AI output — integers, all CO/PO keys present
        for co_id in co_ids:
            co_map = raw_ai.get(co_id, {})
            ai_map[co_id] = {po: int(co_map.get(po, 0)) for po in all_po_ids}

        logger.info(f"AI mapping received for course_id={course_id}")

    except LLMError as e:
        logger.warning(f"LLM unavailable for course_id={course_id}: {e} — using rule engine only")
    except Exception as e:
        logger.warning(f"AI mapping failed for course_id={course_id}: {e} — using rule engine only")

    # ── Step 3: Merge — rules anchor, AI adjusts within ±1 ──────────────────
    if ai_map:
        final = merge_rule_ai(rule_map, ai_map, co_ids, all_po_ids)
        source = "rule_engine + AI"
    else:
        # AI unavailable — rule engine alone (still good)
        final = rule_map
        source = "rule_engine_only"

    logger.info(f"CO-PO mapping final ({source}) for course_id={course_id}: {list(final.keys())}")
    return {"status": "success", "data": final, "source": source}
