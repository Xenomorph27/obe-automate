# backend/routes/ai_chat.py
"""
AI Table Chat endpoint
Receives table context + user message, returns a structured action via Gemini.
Used by Session Plan and Evaluation Plan pages for the AI assistant panel.
"""
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from typing import List, Optional
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger
from backend.core.auth import get_current_user
from backend.database.user_models import User
from backend.database.connection import get_db
from sqlalchemy.ext.asyncio import AsyncSession

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


class StudyMaterialRequest(BaseModel):
    course_name: str
    course_code: str
    department: Optional[str] = ""
    semester: Optional[str] = ""
    cos: Optional[List[dict]] = []
    syllabus_units: Optional[List[dict]] = []


@router.post("/study-materials/{course_id}")
async def get_study_material_recommendations(
    course_id: int,
    req: StudyMaterialRequest,
    current_user: User = Depends(get_current_user),
):
    co_list = "\n".join(
        f"{c.get('co_id','')}: {c.get('statement','')}"
        for c in (req.cos or [])
        if c.get("co_id")
    ) or "Not specified"

    syllabus = "\n".join(
        f"Unit {u.get('unit_no','')}: {u.get('title') or u.get('unit_title','')} — {u.get('topics','')}"
        for u in (req.syllabus_units or [])
    ) or "Not specified"

    prompt = f"""You are an academic resource advisor for an Indian engineering college.
Given the course below, recommend specific study materials used in Indian universities.

Course: {req.course_name} ({req.course_code})
Department: {req.department or 'Engineering'}
Semester: {req.semester or ''}

Course Outcomes:
{co_list}

Syllabus:
{syllabus}

Respond ONLY with a valid JSON object (no markdown, no explanation, no code fences) in exactly this format:
{{
  "textbooks": [
    {{"title":"","author":"","publisher":"","reason":"why this book fits the COs"}}
  ],
  "web": [
    {{"title":"","url":"","unit":"which unit/CO it covers","reason":""}}
  ],
  "journals": [
    {{"title":"","url":"","reason":""}}
  ],
  "moocs": [
    {{"title":"","platform":"","url":"","duration":"","reason":""}}
  ]
}}

Rules:
- Give 3-4 items per category
- Use real book titles commonly prescribed in Indian engineering syllabi
- For web: prefer NPTEL (nptel.ac.in), GeeksforGeeks, Coursera free courses, YouTube playlists
- For journals: use IEEE Xplore, Springer, Elsevier — relevant to the topics
- For MOOCs: prefer NPTEL SWAYAM, Coursera, NPTEL YouTube
- Tailor everything to the exact course topics and COs above
- reason field must explain which specific CO or topic this covers"""

    import json, re
    try:
        raw = await get_llm_response(prompt)
        clean = raw.strip()
        if clean.startswith("```"):
            clean = re.sub(r"^```(?:json)?\s*", "", clean)
            clean = re.sub(r"\s*```$", "", clean)
        clean = clean.strip()
        parsed = json.loads(clean)
        return {"ok": True, "data": parsed}
    except json.JSONDecodeError as e:
        logger.error(f"Study materials JSON parse error: {e} | raw: {raw[:300] if 'raw' in dir() else 'N/A'}")
        raise HTTPException(status_code=500, detail="AI returned invalid response. Please try again.")
    except Exception as e:
        logger.error(f"Study materials recommendation error: {e}")
        raise HTTPException(status_code=500, detail=str(e))




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

    n_rows = len(req.rows)
    prompt = f"""You are an AI table editor for an OBE (Outcome-Based Education) {context} used in NBA/NAAC accreditation at an Indian engineering college.

=== CURRENT TABLE ===
Columns (in order): {", ".join(col_labels)}
Total rows: {n_rows}

ALL rows (format: [index] col1 | col2 | col3 | ...):
{all_rows_text}

=== USER INSTRUCTION ===
{req.user_message}

=== YOUR TASK ===
Step 1 — THINK: Identify exactly what the user wants. Map their intent to ONE of the actions below.
Step 2 — PICK the single best action. Do not combine multiple actions into one response.
Step 3 — OUTPUT: Respond with ONLY a valid JSON object. No prose, no markdown, no code fences, no explanation.

=== ACTION SCHEMAS (pick exactly one) ===

ACTION: add_column
Use when: user wants a new column added to the table.
Schema: {{"action":"add_column","label":"Column Name","values":["value for row 0","value for row 1",...]}}
Rules: "values" array MUST have exactly {n_rows} entries (one per row). Generate meaningful values based on context.
Example — user says "add a Program Indicator column":
→ {{"action":"add_column","label":"Program Indicator","values":["PO1,PO2","PO1,PO3","PO2","PO1","PO2,PO4","PO3"]}}

ACTION: remove_column
Use when: user wants to delete/remove an existing column.
Schema: {{"action":"remove_column","label":"Exact Column Label"}}
Rules: label must exactly match one of: {", ".join(col_labels)}
Example — user says "remove the Methodology column":
→ {{"action":"remove_column","label":"Methodology"}}

ACTION: rename_column
Use when: user wants to rename/relabel a column.
Schema: {{"action":"rename_column","old_label":"Current Label","new_label":"New Label"}}
Rules: old_label must match an existing column exactly.
Example — user says "rename Topic to Points to Cover":
→ {{"action":"rename_column","old_label":"Topic / Points to Cover","new_label":"Points to Cover"}}

ACTION: fill_column
Use when: user wants to populate/fill values in an existing column.
Schema: {{"action":"fill_column","label":"Column Name","values":["val0","val1",...]}}
Rules: "values" array MUST have exactly {n_rows} entries. "label" must match an existing column exactly.
Example — user says "fill the CO column based on topics":
→ {{"action":"fill_column","label":"CO","values":["CO1","CO1","CO2","CO2","CO3","CO3"]}}

ACTION: add_row
Use when: user wants to add/insert a new row to the table.
Schema: {{"action":"add_row","data":{{"Column Label":"value",...}}}}
Rules: include ALL column labels as keys. "data" keys must be column labels (not key names).
Example — user says "add a row for Unit 3 quiz":
→ {{"action":"add_row","data":{{"Lect No":"19","Unit No.":"3","Topic / Points to Cover":"Unit 3 Quiz","Methodology":"Classroom Teaching","Faculty":"","Lecture/Exp.Learning/Eval":"Evaluation","CO":"CO3"}}}}

ACTION: remove_rows
Use when: user wants to delete/remove one or more rows by content or index.
Schema: {{"action":"remove_rows","indices":[0,2,5,...]}}
Rules: scan ALL rows shown above. Match rows by topic/content/index. Return ALL matching [index] numbers.
Example — user says "remove all evaluation rows" (rows [4] and [9] contain "Evaluation"):
→ {{"action":"remove_rows","indices":[4,9]}}

ACTION: update_cell
Use when: user wants to change one specific cell.
Schema: {{"action":"update_cell","row_index":0,"col_label":"Column Label","value":"new value"}}
Rules: row_index is the [index] from the table above. col_label must match an existing column exactly.
Example — user says "set row 3 CO to CO2":
→ {{"action":"update_cell","row_index":3,"col_label":"CO","value":"CO2"}}

ACTION: reorder_rows
Use when: user wants to move a row to a different position.
Schema: {{"action":"reorder_rows","from_index":3,"to_index":0}}
Example — user says "move row 5 to the top":
→ {{"action":"reorder_rows","from_index":5,"to_index":0}}

ACTION: message
Use ONLY when no table change is needed — e.g. user asks a question, or the request is ambiguous/impossible.
Schema: {{"action":"message","text":"your explanation here"}}
Example — user says "what is CO attainment?":
→ {{"action":"message","text":"CO Attainment measures how well students achieved each Course Outcome based on marks in CA and End-Sem exams."}}

=== OBE CONTEXT ===
- CO column values: CO1 through CO6 (Course Outcomes)
- Type column values: Lecture | Exp. Learning | Evaluation
- Methodology values: Classroom Teaching | Tutorial | Flipped Classroom | Case Study | Problem Solving | Group Discussion | Lab
- Bloom's levels: Remember | Understand | Apply | Analyse | Evaluate | Create
- PI = Program Indicator, maps to PO1–PO12

=== CRITICAL RULES ===
1. Output ONLY the raw JSON object. No backticks. No "```json". No text before or after.
2. "values" arrays for add_column and fill_column MUST contain exactly {n_rows} items.
3. "label" in remove_column and fill_column must be an EXACT match to an existing column label.
4. "data" keys in add_row must be column labels (the human-readable labels, NOT the key names).
5. If the instruction is unclear, return a "message" action asking for clarification."""

    import json, re
    try:
        raw = await get_llm_response(prompt)
        clean = raw.strip()

        # Strip markdown fences if present (model sometimes forgets rule 1)
        if clean.startswith("```"):
            clean = re.sub(r"^```(?:json)?\s*", "", clean)
            clean = re.sub(r"\s*```$", "", clean)
        clean = clean.strip()

        # Validate it's parseable JSON before sending to frontend
        parsed = json.loads(clean)

        # Sanity-check required "action" field
        if "action" not in parsed:
            logger.warning(f"AI response missing 'action' field: {clean[:200]}")
            return {"ok": True, "raw": json.dumps({"action": "message", "text": "Sorry, I couldn't understand that. Could you rephrase?"})}

        # Coerce values arrays to correct length if off-by-one (add_column / fill_column)
        if parsed["action"] in ("add_column", "fill_column"):
            vals = parsed.get("values", [])
            if len(vals) != n_rows:
                logger.warning(f"AI values array length {len(vals)} != {n_rows}, padding/trimming")
                if len(vals) < n_rows:
                    vals = vals + [""] * (n_rows - len(vals))
                else:
                    vals = vals[:n_rows]
                parsed["values"] = vals
                clean = json.dumps(parsed)

        logger.info(f"AI table chat action={parsed.get('action')} for plan_type={req.plan_type}")
        return {"ok": True, "raw": clean}

    except json.JSONDecodeError as e:
        logger.error(f"AI table chat JSON parse error: {e} | raw: {raw[:300] if 'raw' in dir() else 'N/A'}")
        fallback = json.dumps({"action": "message", "text": "I had trouble processing that request. Please try rephrasing it."})
        return {"ok": True, "raw": fallback}
    except Exception as e:
        logger.error(f"AI table chat error: {e}")
        return {"ok": False, "error": str(e)}


# ── CO-PO Justification ────────────────────────────────────────────────────────

class CoPoJustificationRequest(BaseModel):
    course_name: str
    course_code: str
    department: str
    cos: list
    pos: list
    co_po_matrix: dict

@router.post("/co-po-justification/{course_id}")
async def generate_co_po_justification(
    course_id: int,
    req: CoPoJustificationRequest,
    current_user: User = Depends(get_current_user),
):
    """Generate CO-PO mapping justifications using backend LLM."""
    mapping_lines = []
    for co in req.cos:
        co_map = req.co_po_matrix.get(co.get("co_id", ""), {})
        for po_id, val in co_map.items():
            if val and str(val) not in ("0", ""):
                po_stmt = next((p.get("statement", po_id) for p in req.pos if p.get("po_id") == po_id), po_id)
                mapping_lines.append(
                    f'{co["co_id"]} → {po_id} ({val}): CO="{co.get("statement","")}" | PO="{po_stmt}"'
                )

    prompt = f"""You are an OBE (Outcome Based Education) expert for engineering colleges. Generate concise CO-PO mapping justifications.

Course: {req.course_name} ({req.course_code})
Department: {req.department}

Course Outcomes:
{chr(10).join(f'{c["co_id"]}: {c.get("statement","")} [Bloom\'s: {c.get("bloom_level","")}]' for c in req.cos)}

CO-PO Mapping:
{chr(10).join(mapping_lines)}

Write one justification line per CO-PO mapping in this exact format:
CO1 → PO1 (3): <short reason why this CO directly maps to this PO at this strength>.

Rules:
- Level 3 = direct, strong alignment; Level 2 = moderate; Level 1 = slight/indirect
- Keep each justification to one line, under 120 characters
- Be technically specific to the course content
- Only list mappings that have a non-zero value
- Do NOT include any preamble or closing remarks, just the justification lines"""

    try:
        text = await get_llm_response(prompt)
        return {"status": "success", "data": text.strip()}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ── Action Plan Generator ──────────────────────────────────────────────────────

class ActionPlanRequest(BaseModel):
    course_name: str
    course_code: str
    prev_co_attainment: str

@router.post("/action-plan/{course_id}")
async def generate_action_plan(
    course_id: int,
    req: ActionPlanRequest,
    current_user: User = Depends(get_current_user),
):
    """Generate a remedial action plan based on previous CO attainment."""
    prompt = f"""You are an OBE expert for engineering colleges. Based on the previous year CO attainment data below, generate a concise remedial action plan.

Course: {req.course_name} ({req.course_code})

Previous CO Attainment:
{req.prev_co_attainment}

Write a bullet-point action plan (4-6 bullets) addressing the low-attainment COs (below 70%). Each bullet should be a specific, actionable intervention. Format:
• <action>

Keep it practical: remedial sessions, tutorials, peer learning, reassessment, extra assignments, etc.
Do NOT include preamble or closing remarks, just the bullet points."""

    try:
        text = await get_llm_response(prompt)
        return {"status": "success", "data": text.strip()}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ── Learning Materials Generator ──────────────────────────────────────────────

class LearningMaterialsRequest(BaseModel):
    course_name: str
    course_code: str
    department: str
    semester: Optional[str] = ""
    cos: list = []
    syllabus_units: list = []

@router.post("/learning-materials/{course_id}")
async def generate_learning_materials(
    course_id: int,
    req: LearningMaterialsRequest,
    current_user: User = Depends(get_current_user),
):
    """Generate recommended learning materials for the course."""
    co_text = "\n".join(f'- {c.get("co_id","")}: {c.get("statement","")}' for c in req.cos) if req.cos else "Not specified"
    units_text = "\n".join(f'- {u.get("unit_no","")}: {u.get("title","")}' for u in req.syllabus_units) if req.syllabus_units else "Not specified"

    prompt = f"""You are an expert academic resource curator for Indian engineering colleges (NBA/NAAC context). Recommend learning materials for the following course.

Course: {req.course_name} ({req.course_code})
Department: {req.department}
Semester: {req.semester}

Course Outcomes:
{co_text}

Syllabus Units:
{units_text}

Generate 5-8 learning material recommendations in this exact format (one per line):
Textbook: <Title> — <Author(s)>, <Publisher>, <Edition/Year>
or
<URL>  (<brief description>)

Include: 1-2 standard textbooks, 1-2 NPTEL/Swayam links, 1 YouTube playlist, 1-2 relevant reference books.
Use real, widely-known resources. Do NOT include preamble or closing remarks, just the material lines."""

    try:
        text = await get_llm_response(prompt)
        return {"status": "success", "data": text.strip()}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
