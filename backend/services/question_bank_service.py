# backend/services/question_bank_service.py
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, delete
from typing import Optional
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger
from backend.database.models import Question, BLOOM_LEVELS

logger = get_logger(__name__)

class QuestionBankService:
    def __init__(self, db: AsyncSession):
        self.db = db

    async def save_question(self, data: dict) -> Question:
        bloom_level = data.get("bloom_level", 1)
        q = Question(
            course_id=data["course_id"], question_text=data["question_text"],
            topic=data.get("topic"), question_type=data.get("question_type","Short Answer"),
            bloom_level=bloom_level, bloom_label=BLOOM_LEVELS.get(bloom_level,"Remember"),
            co_id=data.get("co_id"), po_id=data.get("po_id"),
            marks=data.get("marks",5), source=data.get("source","generated"),
            parent_question_id=data.get("parent_question_id"),
        )
        if data.get("options"):
            q.options = data["options"]
        self.db.add(q)
        await self.db.commit()
        await self.db.refresh(q)
        return q

    async def save_many(self, questions: list) -> list:
        return [await self.save_question(q) for q in questions]

    async def get_bank(self, course_id: int, bloom_level: Optional[int]=None,
                       co_id: Optional[str]=None, question_type: Optional[str]=None,
                       source: Optional[str]=None) -> list:
        stmt = select(Question).where(Question.course_id == course_id)
        if bloom_level: stmt = stmt.where(Question.bloom_level == bloom_level)
        if co_id: stmt = stmt.where(Question.co_id == co_id)
        if question_type: stmt = stmt.where(Question.question_type == question_type)
        if source: stmt = stmt.where(Question.source == source)
        stmt = stmt.order_by(Question.bloom_level, Question.created_at.desc())
        result = await self.db.execute(stmt)
        return [q.to_dict() for q in result.scalars().all()]

    async def get_question(self, question_id: int) -> Question:
        result = await self.db.execute(select(Question).where(Question.id == question_id))
        q = result.scalar_one_or_none()
        if not q: raise OBEException(f"Question {question_id} not found", 404)
        return q

    async def delete_question(self, question_id: int) -> dict:
        await self.db.execute(delete(Question).where(Question.id == question_id))
        await self.db.commit()
        return {"deleted": question_id}

    async def get_stats(self, course_id: int) -> dict:
        questions = await self.get_bank(course_id)
        stats = {"total": len(questions), "by_bloom":{}, "by_type":{}, "by_co":{}, "by_source":{}}
        for q in questions:
            for key, field in [("by_bloom","bloom_label"),("by_type","question_type"),("by_co","co_id"),("by_source","source")]:
                val = q[field] or "Unclassified"
                stats[key][val] = stats[key].get(val, 0) + 1
        return stats
