# backend/routes/dashboard.py
from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.logger import get_logger
from backend.database.connection import get_db
from backend.services.dashboard_service import DashboardService

logger = get_logger(__name__)
router = APIRouter(prefix="/dashboard", tags=["Dashboard"])


@router.get("/department")
async def get_department_dashboard(db: AsyncSession = Depends(get_db)):
    """
    HOD Dashboard — aggregated CO/PO attainment across all courses.
    Returns heatmap data, per-course summaries, and department averages.
    """
    svc = DashboardService(db)
    return await svc.get_department_summary()