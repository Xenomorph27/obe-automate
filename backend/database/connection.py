# backend/database/connection.py
"""
Database connection — supports both SQLite (dev) and PostgreSQL (production).

SQLite   : DATABASE_URL=sqlite:///./obe_automate.db  (default)
PostgreSQL: DATABASE_URL=postgresql://user:pass@host/db  (Railway injects this)

The Railway PostgreSQL plugin sets DATABASE_URL automatically.
We convert it to the async driver variant here.
"""

from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession
from sqlalchemy.orm import sessionmaker, DeclarativeBase
from backend.core.config import DATABASE_URL
from backend.core.logger import get_logger

logger = get_logger(__name__)


def _make_async_url(url: str) -> str:
    """
    Convert a sync DB URL to its async driver equivalent.
      sqlite:///       → sqlite+aiosqlite:///
      postgresql://    → postgresql+asyncpg://
      postgres://      → postgresql+asyncpg://   (Railway uses this older form)
    """
    if url.startswith("sqlite:///"):
        return url.replace("sqlite:///", "sqlite+aiosqlite:///")
    if url.startswith("postgres://"):
        # Railway legacy format
        return url.replace("postgres://", "postgresql+asyncpg://", 1)
    if url.startswith("postgresql://"):
        return url.replace("postgresql://", "postgresql+asyncpg://", 1)
    # Already has a driver prefix — return as-is
    return url


async_db_url = _make_async_url(DATABASE_URL)
logger.info(f"Database driver: {async_db_url.split('://')[0]}")

# PostgreSQL needs pool settings; SQLite doesn't support them
is_postgres = "postgresql" in async_db_url or "asyncpg" in async_db_url

engine_kwargs = {
    "echo": False,
}
if is_postgres:
    engine_kwargs.update({
        "pool_size": 5,
        "max_overflow": 10,
        "pool_pre_ping": True,   # detect stale connections
        "pool_recycle": 300,     # recycle connections every 5 min
    })

engine = create_async_engine(async_db_url, **engine_kwargs)

AsyncSessionLocal = sessionmaker(
    bind=engine,
    class_=AsyncSession,
    expire_on_commit=False,
)


class Base(DeclarativeBase):
    pass


async def get_db():
    """Dependency — gives each request its own DB session, auto-closes after."""
    async with AsyncSessionLocal() as session:
        try:
            yield session
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()


async def init_db():
    """Creates all tables on startup if they don't exist."""
    async with engine.begin() as conn:
        from backend.database import models  # noqa — ensures models are registered
        await conn.run_sync(Base.metadata.create_all)
    logger.info("Database initialised successfully")