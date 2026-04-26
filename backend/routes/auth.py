# backend/routes/auth.py
"""
Auth routes — Day 13
----------------------
POST /auth/login           — get JWT token (form or JSON)
GET  /auth/me              — current user profile
PUT  /auth/me/password     — change own password
GET  /auth/users           — list all users (admin only)
POST /auth/users           — create user (admin only)
PUT  /auth/users/{id}      — update user (admin only)
DELETE /auth/users/{id}    — deactivate user (admin only)
POST /auth/users/{id}/reset-password — admin resets a user's password
"""

from datetime import datetime, timezone
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordRequestForm
from pydantic import BaseModel, EmailStr
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.auth import (
    authenticate_user, create_access_token, get_current_user,
    get_user_by_username, hash_password, require_admin, require_auth, verify_password
)
from backend.core.logger import get_logger
from backend.database.connection import get_db
from backend.database.user_models import User

logger = get_logger(__name__)
router = APIRouter(prefix="/auth", tags=["Auth"])


# ── Pydantic schemas ──────────────────────────────────────────────────────────

class LoginRequest(BaseModel):
    username: str
    password: str


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user: dict


class ChangePasswordRequest(BaseModel):
    current_password: str
    new_password: str


class CreateUserRequest(BaseModel):
    username: str
    password: str
    full_name: str
    role: str = "faculty"
    email: Optional[str] = None
    department: Optional[str] = None


class UpdateUserRequest(BaseModel):
    full_name: Optional[str] = None
    role: Optional[str] = None
    email: Optional[str] = None
    department: Optional[str] = None
    is_active: Optional[bool] = None


class ResetPasswordRequest(BaseModel):
    new_password: str


# ── Login ─────────────────────────────────────────────────────────────────────

@router.post("/login", response_model=TokenResponse)
async def login(req: LoginRequest, db: AsyncSession = Depends(get_db)):
    """
    JSON login endpoint.
    Returns JWT access_token + user profile.
    """
    user = await authenticate_user(db, req.username, req.password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password",
        )
    # Update last_login
    user.last_login = datetime.now(timezone.utc)
    await db.commit()

    token = create_access_token(user.id, user.username, user.role)
    logger.info(f"Login: {user.username} ({user.role})")
    return {"access_token": token, "token_type": "bearer", "user": user.to_dict()}


@router.post("/token")  # OAuth2 form-compatible (for /docs "Authorize" button)
async def login_form(form: OAuth2PasswordRequestForm = Depends(), db: AsyncSession = Depends(get_db)):
    """Form-based login — used by FastAPI /docs Authorize."""
    user = await authenticate_user(db, form.username, form.password)
    if not user:
        raise HTTPException(status_code=401, detail="Incorrect username or password")
    user.last_login = datetime.now(timezone.utc)
    await db.commit()
    token = create_access_token(user.id, user.username, user.role)
    return {"access_token": token, "token_type": "bearer"}


# ── Current user ──────────────────────────────────────────────────────────────

@router.get("/me")
async def get_me(current_user: User = Depends(require_auth)):
    """Return the authenticated user's profile."""
    return current_user.to_dict()


@router.put("/me/password")
async def change_password(
    req: ChangePasswordRequest,
    current_user: User = Depends(require_auth),
    db: AsyncSession = Depends(get_db),
):
    """Change own password. Requires current password for verification."""
    if not verify_password(req.current_password, current_user.hashed_password):
        raise HTTPException(status_code=400, detail="Current password is incorrect")
    if len(req.new_password) < 6:
        raise HTTPException(status_code=400, detail="New password must be at least 6 characters")
    current_user.hashed_password = hash_password(req.new_password)
    await db.commit()
    logger.info(f"Password changed for user: {current_user.username}")
    return {"message": "Password changed successfully"}


# ── Admin: user management ────────────────────────────────────────────────────

@router.get("/users")
async def list_users(
    _: User = Depends(require_admin),
    db: AsyncSession = Depends(get_db),
):
    """List all users (admin only)."""
    result = await db.execute(select(User).order_by(User.created_at))
    return [u.to_dict() for u in result.scalars().all()]


@router.post("/users", status_code=201)
async def create_user(
    req: CreateUserRequest,
    _: User = Depends(require_admin),
    db: AsyncSession = Depends(get_db),
):
    """Create a new user (admin only)."""
    existing = await get_user_by_username(db, req.username)
    if existing:
        raise HTTPException(status_code=409, detail=f"Username '{req.username}' already exists")
    if req.role not in ("faculty", "hod", "admin"):
        raise HTTPException(status_code=400, detail="Role must be faculty, hod, or admin")
    if len(req.password) < 6:
        raise HTTPException(status_code=400, detail="Password must be at least 6 characters")

    user = User(
        username=req.username,
        hashed_password=hash_password(req.password),
        full_name=req.full_name,
        role=req.role,
        email=req.email,
        department=req.department,
        is_active=True,
    )
    db.add(user)
    await db.commit()
    await db.refresh(user)
    logger.info(f"Created user: {user.username} ({user.role})")
    return user.to_dict()


@router.put("/users/{user_id}")
async def update_user(
    user_id: int,
    req: UpdateUserRequest,
    _: User = Depends(require_admin),
    db: AsyncSession = Depends(get_db),
):
    """Update user details (admin only)."""
    result = await db.execute(select(User).where(User.id == user_id))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(status_code=404, detail=f"User {user_id} not found")

    if req.full_name is not None:
        user.full_name = req.full_name
    if req.role is not None:
        if req.role not in ("faculty", "hod", "admin"):
            raise HTTPException(status_code=400, detail="Role must be faculty, hod, or admin")
        user.role = req.role
    if req.email is not None:
        user.email = req.email
    if req.department is not None:
        user.department = req.department
    if req.is_active is not None:
        user.is_active = req.is_active

    await db.commit()
    await db.refresh(user)
    logger.info(f"Updated user: {user.username}")
    return user.to_dict()


@router.post("/users/{user_id}/reset-password")
async def reset_password(
    user_id: int,
    req: ResetPasswordRequest,
    _: User = Depends(require_admin),
    db: AsyncSession = Depends(get_db),
):
    """Admin resets any user's password."""
    result = await db.execute(select(User).where(User.id == user_id))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(status_code=404, detail=f"User {user_id} not found")
    if len(req.new_password) < 6:
        raise HTTPException(status_code=400, detail="Password must be at least 6 characters")

    user.hashed_password = hash_password(req.new_password)
    await db.commit()
    logger.info(f"Admin reset password for: {user.username}")
    return {"message": f"Password reset for {user.username}"}


@router.delete("/users/{user_id}")
async def deactivate_user(
    user_id: int,
    current_admin: User = Depends(require_admin),
    db: AsyncSession = Depends(get_db),
):
    """Deactivate (soft-delete) a user. Cannot deactivate yourself."""
    if user_id == current_admin.id:
        raise HTTPException(status_code=400, detail="Cannot deactivate your own account")
    result = await db.execute(select(User).where(User.id == user_id))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(status_code=404, detail=f"User {user_id} not found")

    user.is_active = False
    await db.commit()
    logger.info(f"Deactivated user: {user.username}")
    return {"message": f"User '{user.username}' deactivated"}
