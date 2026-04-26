# backend/services/question_service.py
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

class QuestionService:
    async def generate_single(self, topic, bloom_level, question_type, course_name,
                               co_id=None, marks=5, extra_instructions="") -> dict:
        bloom_label = BLOOM_LEVELS.get(bloom_level, "Remember")
        mcq_note = 'Also provide exactly 4 options (A,B,C,D) and mark correct in "correct_option".' if question_type=="MCQ" else ""
        prompt = f"""You are an expert question paper setter for engineering colleges.
Generate ONE {question_type} question:
- Course: {course_name}
- Topic: {topic}
- Bloom Level: {bloom_level} — {bloom_label}
- Marks: {marks}
{f'- CO: {co_id}' if co_id else ''}
{f'- Instructions: {extra_instructions}' if extra_instructions else ''}
{mcq_note}
Bloom guide: 1=Remember,2=Understand,3=Apply,4=Analyse,5=Evaluate,6=Create

Respond ONLY with valid JSON (no markdown fences):
{{"question_text":"...","bloom_level":{bloom_level},"bloom_label":"{bloom_label}","question_type":"{question_type}","topic":"{topic}","marks":{marks},"co_id":"{co_id or ''}","options":null,"correct_option":null}}"""
        raw = _strip(await get_llm_response(prompt))
        try: data = json.loads(raw)
        except: raise OBEException("AI returned invalid JSON for question generation.", 502)
        data["bloom_level"] = bloom_level
        data["bloom_label"] = bloom_label
        return data

    async def generate_paper(self, course_name, course_code, cos, total_marks, duration_hours, sections) -> list:
        co_text = "\n".join([f"  - {c['co_id']}: {c['statement']} (Bloom: {c.get('bloom_level','')})" for c in cos])
        prompt = f"""You are an expert question paper setter for Indian engineering colleges (NBA/NAAC).
Generate a COMPLETE question paper:
- Course: {course_name} ({course_code})
- Total Marks: {total_marks}, Duration: {duration_hours}h
Course Outcomes:\n{co_text}
Sections: {json.dumps(sections)}

Rules: map each question to a CO and Bloom level. MCQ must have 4 options.

Respond ONLY with a valid JSON array (no markdown fences). Each element:
{{"section":"...","question_number":1,"question_text":"...","question_type":"...","bloom_level":1,"bloom_label":"Remember","co_id":"CO1","marks":1,"options":null,"correct_option":null}}"""
        raw = _strip(await get_llm_response(prompt))
        try:
            qs = json.loads(raw)
            if not isinstance(qs, list): raise ValueError
        except: raise OBEException("AI returned invalid JSON for paper generation.", 502)
        return qs
