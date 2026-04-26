# backend/database/user_models.py
"""
User model for JWT auth — Day 13
-----------------------------------
Added separately to avoid breaking existing models.py
Imported in init_db() via models.py.
"""
from datetime import datetime
from sqlalchemy import Column, Integer, String, Boolean, DateTime
from backend.database.connection import Base


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(100), unique=True, nullable=False, index=True)
    hashed_password = Column(String(255), nullable=False)
    full_name = Column(String(200), nullable=False, default="")
    email = Column(String(200), nullable=True, unique=True, index=True)
    role = Column(String(50), nullable=False, default="faculty")  # faculty | hod | admin
    department = Column(String(200), nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    last_login = Column(DateTime, nullable=True)

    def to_dict(self):
        return {
            "id": self.id,
            "username": self.username,
            "full_name": self.full_name,
            "email": self.email,
            "role": self.role,
            "department": self.department,
            "is_active": self.is_active,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "last_login": self.last_login.isoformat() if self.last_login else None,
        }
