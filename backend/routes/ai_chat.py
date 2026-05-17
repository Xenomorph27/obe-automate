# backend/routes/ai_chat.py
"""
AI Table Chat endpoint
Receives table context + user message, returns a structured action via Gemini.
Used by Session Plan and Evaluation Plan pages for the AI assistant panel.
"""
from fastapi import APIRouter, Depends
from pydantic import BaseModel
from typing import List
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger
from backend.core.auth import get_current_user
from backend.database.user_models import User

logger = get_logger(__name__)
router = APIRouter(prefix="/ai", tags=["AI"])


class TableCol(BaseModel):
    key: str
    label: str


class TableChatRequest(BaseModel):
    user_message: str
    cols: List[TableCol]
    rows: List[dict]
    plan_type: str = "session"  # "session" or "evaluation"


@router.post("/table-chat")
async def table_chat(
    req: TableChatRequest,
    current_user: User = Depends(get_current_user),
):
    col_labels = [c.label for c in req.cols]
    context = "session plan (OBE)" if req.plan_type == "session" else "evaluation/assessment plan (OBE)"

    # Send ALL rows with their 0-based index so Gemini can find exact rows
    # Cap at 200 rows to stay within token limits — sufficient for any real session plan
    rows_to_send = req.rows[:200]
    all_rows_text = "\n".join(
        f"[{i}] " + " | ".join(str(row.get(c.key, "")) for c in req.cols)
        for i, row in enumerate(rows_to_send)
    )

    prompt = f"""You are an AI assistant editing a {context} table for an engineering college NBA/NAAC accreditation platform.

Current table:
Columns: {", ".join(col_labels)}
Total rows: {len(req.rows)}

ALL rows (format: [index] col1 | col2 | ...):
{all_rows_text}

User instruction: {req.user_message}

IMPORTANT for remove_rows: scan ALL rows above to find matching ones by topic/content. Use the [index] number shown. Return ALL matching indices.

Respond ONLY with a single JSON object — no prose, no markdown fences, no explanation.

Available actions:
- {{"action":"add_column","label":"Column Name","values":["val for row 0","val for row 1",...]}}  — values array must have exactly {len(req.rows)} entries
- {{"action":"remove_column","label":"exact column label"}}
- {{"action":"rename_column","old_label":"Old","new_label":"New"}}
- {{"action":"fill_column","label":"Column Name","values":["val0","val1",...]}}  — exactly {len(req.rows)} entries
- {{"action":"add_row","data":{{"Column Label":"value",...}}}}
- {{"action":"remove_rows","indices":[0,2,...]}}  — use the [index] numbers from the rows listed above
- {{"action":"update_cell","row_index":0,"col_label":"Column Name","value":"new value"}}
- {{"action":"reorder_rows","from_index":3,"to_index":0}}
- {{"action":"message","text":"explanation"}}  — only when no table change is needed

NBA/OBE context:
- CO column: CO1–CO6 (Course Outcomes)
- Program Indicators (PI) map to POs (PO1–PO12)
- Type column values: Lecture, Exp. Learning, Evaluation
- Bloom's levels: Remember, Understand, Apply, Analyse, Evaluate, Create

Respond with ONLY the JSON object."""

    try:
        raw = await get_llm_response(prompt)
        clean = raw.strip()
        if clean.startswith("```"):
            clean = clean.split("```")[1]
            if clean.startswith("json"):
                clean = clean[4:]
        clean = clean.strip()
        return {"ok": True, "raw": clean}
    except Exception as e:
        logger.error(f"AI table chat error: {e}")
        return {"ok": False, "error": str(e)}
