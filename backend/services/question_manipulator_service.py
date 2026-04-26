# backend/services/question_manipulator_service.py
import json
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger
from backend.core.exceptions import OBEException
from backend.database.models import BLOOM_LEVELS

logger = get_logger(__name__)

def _strip(raw):
    raw = raw.strip()
    if raw.startswith("```"):
        lines = raw.split("\n")
        raw = "\n".join(lines[1:-1])
    return raw.strip()

BLOOM_GUIDANCE = {
    1:"Ask students to recall, list, define, or identify facts.",
    2:"Ask students to explain, summarise, describe, or interpret.",
    3:"Ask students to solve, use, demonstrate, or implement.",
    4:"Ask students to compare, differentiate, examine, or break down.",
    5:"Ask students to justify, critique, assess, or argue.",
    6:"Ask students to design, create, construct, or formulate.",
}

class QuestionManipulatorService:
    async def manipulate(self, original_question, original_bloom_level,
                          target_bloom_level, course_name, topic="", teacher_hint="") -> dict:
        orig_label = BLOOM_LEVELS.get(original_bloom_level, "Remember")
        target_label = BLOOM_LEVELS.get(target_bloom_level, "Remember")
        prompt = f"""You are an expert in Bloom's Taxonomy for engineering education.

Original question (Bloom Level {original_bloom_level} — {orig_label}):
"{original_question}"

Course: {course_name}
{f'Topic: {topic}' if topic else ''}
{f'Teacher hint: {teacher_hint}' if teacher_hint else ''}

Rewrite to Bloom Level {target_bloom_level} — {target_label}.
Guidance: {BLOOM_GUIDANCE.get(target_bloom_level,'')}

Keep same topic. Provide 2 alternatives. Respond ONLY with valid JSON (no markdown fences):
{{"original_question":"{original_question}","original_bloom_level":{original_bloom_level},"original_bloom_label":"{orig_label}","target_bloom_level":{target_bloom_level},"target_bloom_label":"{target_label}","suggested_rewrite":"...","alternative_rewrite":"...","explanation":"..."}}"""
        raw = _strip(await get_llm_response(prompt))
        try: return json.loads(raw)
        except: raise OBEException("AI returned invalid JSON for manipulation.", 502)
