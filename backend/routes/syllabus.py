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


class BloomRequest(BaseModel):
    statements: List[str]


@router.post("/classify-bloom")
async def classify_bloom(req: BloomRequest, current_user: User = Depends(require_auth)):
    """
    Classify a list of CO statements into Bloom's Taxonomy levels using LLM.
    Returns a list of level names in the same order as input.
    """
    if not req.statements:
        return {"levels": []}

    prompt = f"""You are a Bloom's Taxonomy expert for university-level engineering courses.

Classify each Course Outcome (CO) statement into exactly one Bloom's level:
- Remember (L1): Recalling facts/terms — verbs: list, define, state, recall, name, describe, memorize, enumerate
- Understand (L2): Explaining ideas in own words — verbs: explain, summarize, discuss, identify, recognize, interpret, paraphrase
- Apply (L3): Using in real-world situations — verbs: apply, solve, implement, demonstrate, calculate, use, execute, perform
- Analyze (L4): Breaking into parts, drawing connections — verbs: analyze, compare, contrast, differentiate, examine, dissect, distinguish
- Evaluate (L5): Justifying decisions based on criteria — verbs: evaluate, assess, justify, critique, judge, defend, recommend
- Create (L6): Producing new original work — verbs: design, formulate, construct, create, develop, synthesize, build, generate, plan

RULES:
1. The ACTION VERB is the strongest signal
2. Context matters: "describe the architecture" = Remember; "describe how you would design" = Create
3. "Discuss" alone = Understand; "Discuss and compare tradeoffs" = Analyze
4. Multiple verbs → pick the HIGHEST level verb
5. Reply ONLY with a JSON array of strings, same order as input, no explanation, no markdown.

CO statements:
{chr(10).join(f"{i+1}. {s}" for i, s in enumerate(req.statements))}

Reply format (example for 3 COs): ["Understand","Remember","Apply"]"""

    try:
        raw = await get_llm_response(prompt)
        # Strip markdown fences if present
        clean = re.sub(r'```json|```', '', raw).strip()
        levels = json.loads(clean)
        valid = {'Remember', 'Understand', 'Apply', 'Analyze', 'Evaluate', 'Create'}
        # Validate each — fallback to Understand if invalid
        safe = [l if l in valid else 'Understand' for l in levels]
        # Pad if LLM returned fewer items
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