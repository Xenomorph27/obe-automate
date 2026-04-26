# backend/services/question_classifier_service.py
import json
from backend.core.llm import get_llm_response
from backend.core.logger import get_logger
from backend.core.exceptions import OBEException

logger = get_logger(__name__)

def _strip(raw):
    raw = raw.strip()
    if raw.startswith("```"):
        lines = raw.split("\n")
        raw = "\n".join(lines[1:-1])
    return raw.strip()

class QuestionClassifierService:
    async def classify(self, question_text, course_name, cos, pos) -> dict:
        co_text = "\n".join([f"  - {c['co_id']}: {c['statement']}" for c in cos])
        po_text = "\n".join([f"  - {p['po_id']}: {p['statement']}" for p in pos])
        prompt = f"""You are an expert in Bloom's Taxonomy and NBA/NAAC OBE.
Classify this question for course: {course_name}

Question: "{question_text}"

Course Outcomes:\n{co_text}
Program Outcomes:\n{po_text}
Bloom levels: 1=Remember,2=Understand,3=Apply,4=Analyse,5=Evaluate,6=Create

Respond ONLY with valid JSON (no markdown fences):
{{"bloom_level":3,"bloom_label":"Apply","co_id":"CO1","po_id":"PO1","question_type":"Short Answer","reasoning":"Brief explanation"}}"""
        raw = _strip(await get_llm_response(prompt))
        try: data = json.loads(raw)
        except: raise OBEException("AI returned invalid JSON for classification.", 502)
        return {"question_text":question_text,"bloom_level":data.get("bloom_level",1),
                "bloom_label":data.get("bloom_label","Remember"),"co_id":data.get("co_id"),
                "po_id":data.get("po_id"),"question_type":data.get("question_type","Short Answer"),
                "reasoning":data.get("reasoning","")}
