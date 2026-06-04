# backend/database/models.py
import json
from datetime import datetime
from sqlalchemy import Column, Integer, String, Float, Text, DateTime, ForeignKey, UniqueConstraint
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
    questions = relationship("Question", back_populates="course", cascade="all, delete-orphan")
    ca_sheets = relationship("CASheet", back_populates="course", cascade="all, delete-orphan")
    eval_plan_rows = relationship("EvalPlanRow", back_populates="course", cascade="all, delete-orphan")
    session_plan_rows = relationship("SessionPlanRow", back_populates="course", cascade="all, delete-orphan")
    course_file_extra = relationship("CourseFileExtra", back_populates="course", uselist=False, cascade="all, delete-orphan")

    @property
    def cos(self): return json.loads(self._cos)
    @cos.setter
    def cos(self, value): self._cos = json.dumps(value)
    @property
    def pos(self): return json.loads(self._pos)
    @pos.setter
    def pos(self, value): self._pos = json.dumps(value)
    @property
    def co_po_matrix(self): return json.loads(self._co_po_matrix)
    @co_po_matrix.setter
    def co_po_matrix(self, value): self._co_po_matrix = json.dumps(value)
    @property
    def evaluation_config(self): return json.loads(self._evaluation_config)
    @evaluation_config.setter
    def evaluation_config(self, value): self._evaluation_config = json.dumps(value)

    def to_dict(self):
        return {"id":self.id,"course_name":self.course_name,"course_code":self.course_code,
                "credits":self.credits,"total_hours":self.total_hours,"faculty_name":self.faculty_name,
                "department":self.department,"semester":self.semester,"academic_year":self.academic_year,
                "cos":self.cos,"pos":self.pos,"co_po_matrix":self.co_po_matrix,
                "evaluation_config":self.evaluation_config,
                "created_at":self.created_at.isoformat() if self.created_at else None}


class COAttainment(Base):
    __tablename__ = "co_attainment"
    id = Column(Integer, primary_key=True, index=True)
    course_id = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, index=True)
    student_id = Column(String(50), nullable=False)
    student_name = Column(String(200), nullable=False)
    _marks = Column("marks", Text, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    course = relationship("Course", back_populates="attainment_records")

    @property
    def marks(self): return json.loads(self._marks)
    @marks.setter
    def marks(self, value): self._marks = json.dumps(value)

    def to_dict(self):
        return {"id":self.id,"course_id":self.course_id,"student_id":self.student_id,
                "student_name":self.student_name,"marks":self.marks,
                "created_at":self.created_at.isoformat() if self.created_at else None}


BLOOM_LEVELS = {1:"Remember",2:"Understand",3:"Apply",4:"Analyse",5:"Evaluate",6:"Create"}


class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, index=True)
    course_id = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, index=True)
    question_text = Column(Text, nullable=False)
    topic = Column(String(300), nullable=True)
    question_type = Column(String(50), nullable=False, default="Short Answer")
    bloom_level = Column(Integer, nullable=False, default=1)
    bloom_label = Column(String(50), nullable=False, default="Remember")
    co_id = Column(String(20), nullable=True)
    po_id = Column(String(20), nullable=True)
    marks = Column(Integer, nullable=False, default=5)
    source = Column(String(50), nullable=False, default="generated")
    _options = Column("options", Text, nullable=True)
    parent_question_id = Column(Integer, ForeignKey("questions.id", ondelete="SET NULL"), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    course = relationship("Course", back_populates="questions")

    @property
    def options(self): return json.loads(self._options) if self._options else None
    @options.setter
    def options(self, value): self._options = json.dumps(value) if value else None

    def to_dict(self):
        return {"id":self.id,"course_id":self.course_id,"question_text":self.question_text,
                "topic":self.topic,"question_type":self.question_type,"bloom_level":self.bloom_level,
                "bloom_label":self.bloom_label,"co_id":self.co_id,"po_id":self.po_id,
                "marks":self.marks,"source":self.source,"options":self.options,
                "parent_question_id":self.parent_question_id,
                "created_at":self.created_at.isoformat() if self.created_at else None}


class CASheet(Base):
    """Stores master attainment question paper + marks per CA component, persisted in DB."""
    __tablename__ = "ca_sheets"
    id = Column(Integer, primary_key=True, index=True)
    course_id = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, index=True)
    ca_label = Column(String(100), nullable=False)   # e.g. "Quiz 1", "Unit Test 1", "ESE"
    _qp = Column("qp", Text, nullable=False, default="[]")      # JSON list of question objects
    _marks = Column("marks", Text, nullable=False, default="{}")  # JSON dict {prn: {q_no: mark}}
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    course = relationship("Course", back_populates="ca_sheets")
    __table_args__ = (UniqueConstraint("course_id", "ca_label", name="uq_ca_sheet"),)

    @property
    def qp(self): return json.loads(self._qp)
    @qp.setter
    def qp(self, value): self._qp = json.dumps(value)
    @property
    def marks(self): return json.loads(self._marks)
    @marks.setter
    def marks(self, value): self._marks = json.dumps(value)


class EvalPlanRow(Base):
    """Stores saved evaluation plan rows in DB so they survive server restarts."""
    __tablename__ = "eval_plan_rows"
    id = Column(Integer, primary_key=True, index=True)
    course_id = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, index=True)
    _rows = Column("rows", Text, nullable=False, default="[]")   # full rows JSON
    _cols = Column("cols", Text, nullable=False, default="[]")   # cols JSON
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    course = relationship("Course", back_populates="eval_plan_rows")
    __table_args__ = (UniqueConstraint("course_id", name="uq_eval_plan"),)

    @property
    def rows(self): return json.loads(self._rows)
    @rows.setter
    def rows(self, value): self._rows = json.dumps(value)
    @property
    def cols(self): return json.loads(self._cols)
    @cols.setter
    def cols(self, value): self._cols = json.dumps(value)


class SessionPlanRow(Base):
    """Stores saved session plan rows in DB so they survive server restarts."""
    __tablename__ = "session_plan_rows"
    id = Column(Integer, primary_key=True, index=True)
    course_id = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, index=True)
    _rows = Column("rows", Text, nullable=False, default="[]")
    _cols = Column("cols", Text, nullable=False, default="[]")
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    course = relationship("Course", back_populates="session_plan_rows")
    __table_args__ = (UniqueConstraint("course_id", name="uq_session_plan"),)

    @property
    def rows(self): return json.loads(self._rows)
    @rows.setter
    def rows(self, value): self._rows = json.dumps(value)
    @property
    def cols(self): return json.loads(self._cols)
    @cols.setter
    def cols(self, value): self._cols = json.dumps(value)


class CourseFileExtra(Base):
    """
    Stores additional fields needed for course file generation that aren't
    in the core Course model. All fields are optional — faculty fills them
    in the UI before generating the course file.
    """
    __tablename__ = "course_file_extra"
    id = Column(Integer, primary_key=True, index=True)
    course_id = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, unique=True, index=True)

    # Cover — institution details
    institution_name    = Column(String(200), nullable=True, default="Symbiosis Institute of Technology")
    institution_address = Column(String(300), nullable=True, default="SIU Pune 412115, Maharashtra, India")

    # Section 1 — Vision & Mission (changes per department)
    vision_text = Column(Text, nullable=True, default="")
    mission_text = Column(Text, nullable=True, default="")   # M1, M2, M3 as newline-separated

    # Section 2 — Batch info (e.g. "2024-28")
    batch = Column(String(20), nullable=True, default="")

    # Section 2 — POs / PEOs / PSOs (free-text, one per line)
    po_peo_pso_text = Column(Text, nullable=True, default="")   # PO1–PO12
    peo_text        = Column(Text, nullable=True, default="")   # PEO1–PEO3
    pso_text        = Column(Text, nullable=True, default="")   # PSO1–PSO2

    # Section 4 — CO-PO justification text
    co_po_justification = Column(Text, nullable=True, default="")

    # Section 5 — Previous year CO attainment + action plan
    prev_co_attainment = Column(Text, nullable=True, default="")   # free-text / JSON summary
    action_plan = Column(Text, nullable=True, default="")

    # Section 8 — Slow/Advanced learners
    slow_learners = Column(Text, nullable=True, default="")        # JSON list or free-text
    advanced_learners = Column(Text, nullable=True, default="")

    # Section 10 — Activity reports
    activity_reports = Column(Text, nullable=True, default="")

    # Section 11 — Learning material links
    learning_material_links = Column(Text, nullable=True, default="")  # newline-separated URLs

    # Section 13 — Attendance links
    attendance_links = Column(Text, nullable=True, default="")

    # Section 14 — Student list
    student_list = Column(Text, nullable=True, default="")

    # Custom tabs (JSON list)
    custom_tabs = Column(Text, nullable=True, default="[]")

    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    course = relationship("Course", back_populates="course_file_extra")

    def to_dict(self):
        return {
            "institution_name": getattr(self, "institution_name", None) or "Symbiosis Institute of Technology",
            "institution_address": getattr(self, "institution_address", None) or "SIU Pune 412115, Maharashtra, India",
            "vision_text": self.vision_text or "",
            "mission_text": self.mission_text or "",
            "batch": self.batch or "",
            "po_peo_pso_text": getattr(self, "po_peo_pso_text", None) or "",
            "peo_text": getattr(self, "peo_text", None) or "",
            "pso_text": getattr(self, "pso_text", None) or "",
            "co_po_justification": getattr(self, "co_po_justification", None) or "",
            "prev_co_attainment": self.prev_co_attainment or "",
            "action_plan": self.action_plan or "",
            "slow_learners": self.slow_learners or "",
            "advanced_learners": self.advanced_learners or "",
            "activity_reports": self.activity_reports or "",
            "learning_material_links": self.learning_material_links or "",
            "attendance_links": self.attendance_links or "",
            "student_list": getattr(self, "student_list", None) or "",
            "custom_tabs": getattr(self, "custom_tabs", None) or "[]",
        }


class CourseFileAttachment(Base):
    """
    Stores arbitrary file uploads attached to a course file section.
    Faculty can upload timetables, event photos, email screenshots, etc.
    Each attachment belongs to a course and is tagged with a section number + label.
    """
    __tablename__ = "course_file_attachments"
    id          = Column(Integer, primary_key=True, index=True)
    course_id   = Column(Integer, ForeignKey("courses.id", ondelete="CASCADE"), nullable=False, index=True)
    section_no  = Column(Integer, nullable=True)          # 1-13, which section it belongs to
    label       = Column(String(200), nullable=False)     # human-readable label e.g. "Faculty Timetable"
    filename    = Column(String(300), nullable=False)     # original filename
    stored_path = Column(String(500), nullable=False)     # path on disk
    mime_type   = Column(String(100), nullable=True)
    file_size   = Column(Integer, nullable=True)          # bytes
    uploaded_at = Column(DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            "id":          self.id,
            "course_id":   self.course_id,
            "section_no":  self.section_no,
            "label":       self.label,
            "filename":    self.filename,
            "mime_type":   self.mime_type,
            "file_size":   self.file_size,
            "uploaded_at": self.uploaded_at.isoformat() if self.uploaded_at else None,
        }


# Import user models so they're registered in Base metadata
from backend.database.user_models import User  # noqa
