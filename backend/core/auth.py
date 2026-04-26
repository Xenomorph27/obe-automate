# backend/core/auth.py
"""
JWT Authentication — Day 13
------------------------------
Replaces hardcoded faculty/faculty123 + hod/hod123 with real DB-backed users.

Features:
  - bcrypt password hashing (passlib)
  - HS256 JWT tokens with role + user_id claims
  - 8-hour access tokens (configurable via JWT_EXPIRE_HOURS)
  - FastAPI dependency: require_auth, require_hod
  - Seed users on startup if DB is empty

Config (env vars):
  JWT_SECRET_KEY  — set in Railway Variables (auto-generated if absent)
  JWT_EXPIRE_HOURS — default 8
"""

import os
import secrets
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from jose import JWTError, jwt
from passlib.context import CryptContext
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.logger import get_logger
from backend.database.connection import get_db
from backend.database.user_models import User

logger = get_logger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────

JWT_SECRET_KEY = os.getenv("JWT_SECRET_KEY") or secrets.token_hex(32)
JWT_ALGORITHM = "HS256"
JWT_EXPIRE_HOURS = int(os.getenv("JWT_EXPIRE_HOURS", "8"))

if not os.getenv("JWT_SECRET_KEY"):
    logger.warning(
        "JWT_SECRET_KEY not set in env — using ephemeral key. "
        "All sessions will be invalidated on restart. "
        "Set JWT_SECRET_KEY in Railway Variables for persistence."
    )

# ── Passlib ───────────────────────────────────────────────────────────────────

pwd_ctx = CryptContext(schemes=["bcrypt"], deprecated="auto")


def hash_password(plain: str) -> str:
    return pwd_ctx.hash(plain)


def verify_password(plain: str, hashed: str) -> bool:
    return pwd_ctx.verify(plain, hashed)


# ── JWT ───────────────────────────────────────────────────────────────────────

def create_access_token(user_id: int, username: str, role: str) -> str:
    expire = datetime.now(timezone.utc) + timedelta(hours=JWT_EXPIRE_HOURS)
    payload = {
        "sub": str(user_id),
        "username": username,
        "role": role,
        "exp": expire,
        "iat": datetime.now(timezone.utc),
    }
    return jwt.encode(payload, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)


def decode_token(token: str) -> dict:
    try:
        return jwt.decode(token, JWT_SECRET_KEY, algorithms=[JWT_ALGORITHM])
    except JWTError as e:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Invalid or expired token: {e}",
            headers={"WWW-Authenticate": "Bearer"},
        )


# ── OAuth2 scheme ─────────────────────────────────────────────────────────────

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/auth/login")


# ── FastAPI dependencies ──────────────────────────────────────────────────────

async def get_current_user(
    token: str = Depends(oauth2_scheme),
    db: AsyncSession = Depends(get_db),
) -> User:
    """Dependency: validate JWT and return the User row."""
    payload = decode_token(token)
    user_id = int(payload.get("sub", 0))

    result = await db.execute(select(User).where(User.id == user_id, User.is_active == True))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(status_code=401, detail="User not found or inactive")
    return user


async def require_auth(user: User = Depends(get_current_user)) -> User:
    """Any authenticated user (faculty or hod)."""
    return user


async def require_hod(user: User = Depends(get_current_user)) -> User:
    """Only HOD or admin."""
    if user.role not in ("hod", "admin"):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="HOD or admin role required",
        )
    return user


async def require_admin(user: User = Depends(get_current_user)) -> User:
    """Only admin."""
    if user.role != "admin":
        raise HTTPException(status_code=403, detail="Admin role required")
    return user


# ── User helpers ──────────────────────────────────────────────────────────────

async def get_user_by_username(db: AsyncSession, username: str) -> Optional[User]:
    result = await db.execute(select(User).where(User.username == username))
    return result.scalar_one_or_none()


async def authenticate_user(db: AsyncSession, username: str, password: str) -> Optional[User]:
    user = await get_user_by_username(db, username)
    if not user or not user.is_active:
        return None
    if not verify_password(password, user.hashed_password):
        return None
    return user


# ── Seed default users on startup ─────────────────────────────────────────────

DEFAULT_USERS = [
    {
        "username": "faculty",
        "password": os.getenv("FACULTY_PASSWORD", "faculty123"),
        "full_name": "Faculty User",
        "role": "faculty",
        "email": "faculty@obe.edu",
        "department": "Computer Science & Engineering",
    },
    {
        "username": "hod",
        "password": os.getenv("HOD_PASSWORD", "hod123"),
        "full_name": "Head of Department",
        "role": "hod",
        "email": "hod@obe.edu",
        "department": "Computer Science & Engineering",
    },
    {
        "username": "admin",
        "password": os.getenv("ADMIN_PASSWORD", "admin@OBE2024"),
        "full_name": "System Administrator",
        "role": "admin",
        "email": "admin@obe.edu",
        "department": None,
    },
]


async def seed_default_users(db: AsyncSession) -> None:
    """
    Called on startup. Creates default users only if the users table is empty.
    Override passwords via env: FACULTY_PASSWORD, HOD_PASSWORD, ADMIN_PASSWORD.
    """
    result = await db.execute(select(User).limit(1))
    if result.scalar_one_or_none() is not None:
        logger.info("Users already seeded — skipping")
        return

    logger.info("Seeding default users...")
    for u in DEFAULT_USERS:
        user = User(
            username=u["username"],
            hashed_password=hash_password(u["password"]),
            full_name=u["full_name"],
            role=u["role"],
            email=u["email"],
            department=u["department"],
            is_active=True,
        )
        db.add(user)
    await db.commit()
    logger.info(f"Seeded {len(DEFAULT_USERS)} default users")
