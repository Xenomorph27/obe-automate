# backend/main.py
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pathlib import Path

from backend.routes.syllabus import router as syllabus_router
from backend.routes.courses import router as courses_router
from backend.routes.session_plan import router as session_plan_router
from backend.routes.evaluation_plan import router as evaluation_plan_router
from backend.routes.attainment import router as attainment_router
from backend.routes.dashboard import router as dashboard_router
from backend.routes.questions import router as questions_router
from backend.routes.auth import router as auth_router
from backend.database.connection import init_db
from backend.core.config import APP_NAME, APP_VERSION, STORAGE_PATH
from backend.core.logger import get_logger

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

app = FastAPI(title=APP_NAME, version=APP_VERSION,
              description="OBE Automate — Outcome-Based Education platform",
              lifespan=lifespan, docs_url="/docs", redoc_url="/redoc")

app.include_router(syllabus_router)
app.include_router(courses_router)
app.include_router(session_plan_router)
app.include_router(evaluation_plan_router)
app.include_router(attainment_router)
app.include_router(dashboard_router)
app.include_router(questions_router)
app.include_router(auth_router)

@app.get("/health")
def health_check():
    from backend.core.storage import get_storage
    storage = get_storage()
    return {"status": "ok", "app": APP_NAME, "version": APP_VERSION,
            "storage_path": str(storage.base), "storage_configured": bool(STORAGE_PATH)}


@app.get("/storage/info", tags=["System"])
def storage_info():
    """Returns current storage configuration — useful for verifying Volume is mounted."""
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
for _cat in ["session_plans", "evaluation_plans", "attainment_reports", "nba_reports", "question_papers"]:
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