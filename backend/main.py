# backend/main.py
from contextlib import asynccontextmanager
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import os
import time
from collections import defaultdict

from backend.routes.syllabus import router as syllabus_router
from backend.routes.courses import router as courses_router
from backend.routes.session_plan import router as session_plan_router
from backend.routes.evaluation_plan import router as evaluation_plan_router
from backend.routes.attainment import router as attainment_router
from backend.routes.dashboard import router as dashboard_router
from backend.routes.questions import router as questions_router
from backend.routes.auth import router as auth_router
from backend.routes.students import router as students_router
from backend.database.connection import init_db
from backend.core.config import APP_NAME, APP_VERSION, STORAGE_PATH
from backend.core.logger import get_logger
from backend.routes.ai_chat import router as ai_chat_router
from backend.routes.co_po_template import router as co_po_template_router

logger = get_logger(__name__)
FRONTEND_DIR = Path("frontend")
FRONTEND_DIR.mkdir(exist_ok=True)

@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("Starting up — initialising database...")
    await init_db()
    # Seed default users (Day 13)
    from backend.core.auth import seed_default_users
    from backend.database.connection import AsyncSessionLocal
    async with AsyncSessionLocal() as _db:
        await seed_default_users(_db)
    yield
    logger.info("Shutting down")

# Hide API docs in production — only expose in dev
_IS_DEV = os.getenv("ENVIRONMENT", "production").lower() in ("dev", "development", "local")

app = FastAPI(
    title=APP_NAME,
    version=APP_VERSION,
    description="OBE Automate — Outcome-Based Education platform",
    lifespan=lifespan,
    docs_url="/docs" if _IS_DEV else None,
    redoc_url="/redoc" if _IS_DEV else None,
    openapi_url="/openapi.json" if _IS_DEV else None,
)

# ── CORS ─────────────────────────────────────────────────────────────────────
# Only allow your Railway domain + localhost in dev
_ALLOWED_ORIGINS = [
    "https://obe-automate.vercel.app",
    "http://localhost:8000",
    "http://localhost:3000",
    "http://127.0.0.1:8000",
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type"],
    max_age=3600,
)

# ── Login Rate Limiter ─────────────────────────────────────────────────────────
# Simple in-memory limiter: max 10 login attempts per IP per 15 minutes
_LOGIN_ATTEMPTS: dict = defaultdict(list)
_LOGIN_MAX = 10
_LOGIN_WINDOW = 900  # 15 minutes in seconds

@app.middleware("http")
async def rate_limit_login(request: Request, call_next):
    if request.url.path in ("/auth/login", "/auth/token") and request.method == "POST":
        ip = request.client.host if request.client else "unknown"
        now = time.time()
        window_start = now - _LOGIN_WINDOW
        # Purge old attempts
        _LOGIN_ATTEMPTS[ip] = [t for t in _LOGIN_ATTEMPTS[ip] if t > window_start]
        if len(_LOGIN_ATTEMPTS[ip]) >= _LOGIN_MAX:
            return JSONResponse(
                status_code=429,
                content={"detail": "Too many login attempts. Please wait 15 minutes and try again."},
                headers={"Retry-After": "900"},
            )
        _LOGIN_ATTEMPTS[ip].append(now)
    return await call_next(request)

app.include_router(syllabus_router)
app.include_router(courses_router)
app.include_router(session_plan_router)
app.include_router(evaluation_plan_router)
app.include_router(attainment_router)
app.include_router(dashboard_router)
app.include_router(questions_router)
app.include_router(auth_router)
app.include_router(students_router)
app.include_router(ai_chat_router)
app.include_router(co_po_template_router)

@app.get("/health")
def health_check():
    # Keep health check public (needed by Railway uptime checks) but return minimal info
    return {"status": "ok", "app": APP_NAME, "version": APP_VERSION}


@app.get("/storage/info", tags=["System"])
def storage_info():
    """Returns storage configuration — admin use only (protected by ENVIRONMENT check)."""
    if not _IS_DEV:
        from fastapi import HTTPException
        raise HTTPException(status_code=403, detail="Not available in production")
    from backend.core.storage import get_storage
    s = get_storage()
    cats = ["session_plans", "evaluation_plans", "attainment_reports", "nba_reports", "question_papers"]
    counts = {}
    for c in cats:
        try:
            counts[c] = len(list((s.base / c).glob("*"))) if (s.base / c).exists() else 0
        except Exception:
            counts[c] = -1
    return {"base_path": str(s.base.resolve()), "storage_backend": STORAGE_PATH or "local (generated_docs/)",
            "file_counts": counts, "is_volume_mounted": bool(STORAGE_PATH)}

# Mount generated files for direct URL access
# CRITICAL: mkdir BEFORE StaticFiles — it crashes if the directory doesn't exist on fresh containers
from backend.core.storage import get_storage as _gs
_storage_base = _gs().base
_storage_base.mkdir(parents=True, exist_ok=True)
# Pre-create all category subdirs so StaticFiles doesn't fail on first request
for _cat in ["session_plans", "evaluation_plans", "attainment_reports", "nba_reports", "question_papers", "co_po_templates"]:
    (_storage_base / _cat).mkdir(parents=True, exist_ok=True)
app.mount("/files", StaticFiles(directory=str(_storage_base)), name="files")

assets_dir = FRONTEND_DIR / "assets"
assets_dir.mkdir(parents=True, exist_ok=True)
app.mount("/assets", StaticFiles(directory=str(assets_dir)), name="assets")

@app.get("/", include_in_schema=False)
@app.get("/app", include_in_schema=False)
async def serve_frontend():
    index = FRONTEND_DIR / "index.html"
    if index.exists():
        return FileResponse(str(index))
    return {"message": "Frontend not found."}