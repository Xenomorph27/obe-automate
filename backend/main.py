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
from backend.database.connection import init_db
from backend.core.config import APP_NAME, APP_VERSION
from backend.core.logger import get_logger

logger = get_logger(__name__)

FRONTEND_DIR = Path("frontend")
FRONTEND_DIR.mkdir(exist_ok=True)


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("Starting up — initialising database...")
    await init_db()
    yield
    logger.info("Shutting down")


app = FastAPI(
    title=APP_NAME,
    version=APP_VERSION,
    description="OBE Automate — Outcome-Based Education platform",
    lifespan=lifespan,
    docs_url="/docs",
    redoc_url="/redoc",
)

# ── API Routers ─────────────────────────────────────────────────────────
app.include_router(syllabus_router)
app.include_router(courses_router)
app.include_router(session_plan_router)
app.include_router(evaluation_plan_router)
app.include_router(attainment_router)
app.include_router(dashboard_router)


# ── Health check ────────────────────────────────────────────────────────
@app.get("/health")
def health_check():
    return {"status": "ok", "app": APP_NAME, "version": APP_VERSION}


# ── Frontend — serve React SPA ──────────────────────────────────────────
assets_dir = FRONTEND_DIR / "assets"
assets_dir.mkdir(exist_ok=True)
app.mount("/assets", StaticFiles(directory=str(assets_dir)), name="assets")


@app.get("/", include_in_schema=False)
@app.get("/app", include_in_schema=False)
async def serve_frontend():
    """Serve the React SPA."""
    index = FRONTEND_DIR / "index.html"
    if index.exists():
        return FileResponse(str(index))
    return {"message": "Frontend not found. Place index.html in the frontend/ folder."}