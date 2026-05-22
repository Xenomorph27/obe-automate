# backend/routes/syllabus.py
from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from backend.services.syllabus_service import SyllabusService
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.core.auth import require_auth
from backend.database.user_models import User

from pydantic import BaseModel
from typing import List
from backend.core.llm import get_llm_response
import json, re

logger = get_logger(__name__)
router = APIRouter(prefix="/syllabus", tags=["Syllabus"])

syllabus_service = SyllabusService()


class UnitContext(BaseModel):
    unit_number: int
    unit_title: str
    topics: List[str] = []

class BloomRequest(BaseModel):
    statements: List[str]
    units: List[UnitContext] = []  # optional — used for content-aware classification


@router.post("/classify-bloom")
async def classify_bloom(req: BloomRequest, current_user: User = Depends(require_auth)):
    """
    Classify CO statements into Bloom's Taxonomy levels using LLM.
    When units are provided, uses unit content depth for accurate level inference
    instead of relying solely on the CO verb (which is often poorly written).
    """
    if not req.statements:
        return {"levels": []}

    # Build unit context block if units were provided
    unit_context_block = ""
    if req.units:
        lines = ["\nCOURSE UNITS (use these to judge cognitive depth required):"]
        for u in req.units:
            topics_str = ", ".join(u.topics) if u.topics else "N/A"
            lines.append(f"  Unit {u.unit_number} — {u.unit_title}: {topics_str}")
        unit_context_block = "\n".join(lines)

    prompt = f"""You are a Bloom's Taxonomy expert for university-level engineering courses in an OBE (Outcome-Based Education) system.

Your task: Classify each Course Outcome (CO) into the most appropriate Bloom's level.

IMPORTANT: CO statements in Indian university syllabi are often written with weak verbs (explain, describe, discuss) regardless of the actual cognitive depth required. You MUST look at the UNIT CONTENT to determine the true cognitive level — not just the verb.

Bloom's Taxonomy levels:
- Remember (L1): Pure rote recall — list facts, define terms, name items. No comprehension needed.
- Understand (L2): Comprehend concepts — describe, explain, summarize, discuss ideas. Foundation level.
- Apply (L3): Use knowledge in real situations — implement algorithms, compute values, perform procedures, simulate, operate tools.
- Analyze (L4): Break down and find relationships — compare algorithms, select appropriate method for a problem, examine tradeoffs, differentiate approaches.
- Evaluate (L5): Make judgments — assess performance, justify choices, critique methods, recommend solutions with reasoning.
- Create (L6): Produce new work — design systems, formulate new approaches, build novel solutions.

CLASSIFICATION RULES:
1. VERB IS A WEAK SIGNAL in Indian syllabi — weight unit content heavily
2. If the unit covers: multiple algorithms + selection criteria + performance evaluation → L4 Analyze minimum
3. If the unit covers: implementation, computation, simulation → L3 Apply minimum  
4. If the unit covers: only definitions, types, introductions → L2 Understand
5. "How to select the appropriate algorithm" = L4 Analyze (requires comparing and deciding)
6. "Performance evaluation" or "comparison of algorithms" in unit = L4 Analyze
7. Advanced/deep learning topics with practical application = L3 Apply minimum
8. A course should ideally have a spread of levels — avoid mapping everything to L2
9. Reply ONLY with a JSON array of strings, same order as input. No explanation, no markdown.
{unit_context_block}

CO statements to classify:
{chr(10).join(f"{i+1}. {s}" for i, s in enumerate(req.statements))}

Reply format (example for 5 COs): ["Understand","Analyze","Apply","Analyze","Apply"]"""

    try:
        raw = await get_llm_response(prompt)
        clean = re.sub(r'```json|```', '', raw).strip()
        # Extract JSON array from response
        start = clean.find('[')
        end = clean.rfind(']')
        if start != -1 and end != -1:
            clean = clean[start:end+1]
        levels = json.loads(clean)
        valid = {'Remember', 'Understand', 'Apply', 'Analyze', 'Evaluate', 'Create'}
        safe = [l if l in valid else 'Understand' for l in levels]
        while len(safe) < len(req.statements):
            safe.append('Understand')
        return {"levels": safe[:len(req.statements)]}
    except Exception as e:
        logger.warning(f"Bloom classification failed: {e}")
        return {"levels": ["Understand"] * len(req.statements)}


@router.post("/upload")
async def upload_syllabus(file: UploadFile = File(...), current_user: User = Depends(require_auth)):
    """
    Upload a syllabus PDF.
    Extracts topics and Course Outcomes using Gemini AI.
    """
    logger.info(f"Received file upload: {file.filename}")

    try:
        result = await syllabus_service.process_syllabus(file)
        logger.info(f"Successfully processed: {file.filename}")
        return {
            "status": "success",
            "filename": file.filename,
            "data": result
        }

    except OBEException as e:
        logger.error(f"OBE error processing {file.filename}: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)

    except Exception as e:
        logger.exception(f"Unexpected error processing {file.filename}")
        raise HTTPException(status_code=500, detail="Internal server error")