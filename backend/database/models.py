# backend/database/models.py
import json
from datetime import datetime
from sqlalchemy import Column, Integer, String, Float, Text, DateTime, ForeignKey
from sqlalchemy.orm import relationship
from backend.database.connection import Base


class Course(Base):
    __tablename__ = "courses"

    id = Column(Integer, primary_key=True, index=True)
    course_name = Column(String(200), nullable=False)
    course_code = Column(String(50), nullable=False)
    credits = Column(Integer, nullable=False)
    total_hours = Column(Integer, nullable=False)
    faculty_name = Column(String(200), nullable=False)
    department = Column(String(200), nullable=False)
    semester = Column(String(50), nullable=False)
    academic_year = Column(String(20), nullable=False)

    _cos = Column("cos", Text, nullable=False)
    _pos = Column("pos", Text, nullable=False)
    _co_po_matrix = Column("co_po_matrix", Text, nullable=False)
    _evaluation_config = Column("evaluation_config", Text, nullable=False)

    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    attainment_records = relationship("COAttainment", back_populates="course", cascade="all, delete-orphan")

    @property
    def cos(self):
        return json.loads(self._cos)

    @cos.setter
    def cos(self, value):
        self._cos = json.dumps(value)

    @property
    def pos(self):
        return json.loads(self._pos)

    @pos.setter
    def pos(self, value):
        self._pos = json.dumps(value)

    @property
    def co_po_matrix(self):
        return json.loads(self._co_po_matrix)

    @co_po_matrix.setter
    def co_po_matrix(self, value):
        self._co_po_matrix = json.dumps(value)

    @property
    def evaluation_config(self):
        return json.loads(self._evaluation_config)

    @evaluation_config.setter
    def evaluation_config(self, value):
        self._evaluation_config = json.dumps(value)

    def to_dict(self):
        return {
            "id": self.id,
            "course_name": self.course_name,
            "course_code": self.course_code,
            "credits": self.credits,
            "total_hours": self.total_hours,
            "faculty_name": self.faculty_name,
            "department": self.department,
            "semester": self.semester,
            "academic_year": self.academic_year,
            "cos": self.cos,
            "pos": self.pos,
            "co_po_matrix": self.co_po_matrix,
            "evaluation_config": self.evaluation_config,
            "created_at": self.created_at.isoformat() if self.created_at else None,
        }


class COAttainment(Base):
    """
    Stores per-student, per-CO marks for a course.
    One row = one student's complete marks entry for the course.
    """
    __tablename__ = "co_attainment"

    id = Column(Integer, primary_key=True, index=True)
    course_id = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, index=True)
    student_id = Column(String(50), nullable=False)
    student_name = Column(String(200), nullable=False)

    # {"CO1": {"Quiz": 8, "Unit Test": 18}, "CO2": {"Quiz": 6, "Unit Test": 15}}
    _marks = Column("marks", Text, nullable=False)

    created_at = Column(DateTime, default=datetime.utcnow)

    course = relationship("Course", back_populates="attainment_records")

    @property
    def marks(self):
        return json.loads(self._marks)

    @marks.setter
    def marks(self, value):
        self._marks = json.dumps(value)

    def to_dict(self):
        return {
            "id": self.id,
            "course_id": self.course_id,
            "student_id": self.student_id,
            "student_name": self.student_name,
            "marks": self.marks,
            "created_at": self.created_at.isoformat() if self.created_at else None,
        }