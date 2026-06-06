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
    """Creates all tables on startup if they don't exist.
    Also runs safe ADD COLUMN migrations for any new columns added to existing tables.
    Uses IF NOT EXISTS / DO NOTHING patterns so it's safe to run on every startup.
    """
    async with engine.begin() as conn:
        from backend.database import models  # noqa — ensures models are registered
        await conn.run_sync(Base.metadata.create_all)

    # ── Safe column migrations ─────────────────────────────────────────────
    # SQLAlchemy create_all won't add new columns to existing tables.
    # We use raw SQL with IF NOT EXISTS (SQLite) / DO NOTHING (Postgres) guards.
    _new_columns = [
        # (table, column, type_sql)
        ("course_file_extra", "institution_name",    "VARCHAR(200) DEFAULT 'Symbiosis Institute of Technology'"),
        ("course_file_extra", "institution_address", "VARCHAR(300) DEFAULT 'SIU Pune 412115, Maharashtra, India'"),
        ("course_file_extra", "co_po_justification", "TEXT DEFAULT ''"),
        # New columns added in v2
        ("course_file_extra", "po_peo_pso_text",     "TEXT DEFAULT ''"),
        ("course_file_extra", "peo_text",             "TEXT DEFAULT ''"),
        ("course_file_extra", "pso_text",             "TEXT DEFAULT ''"),
        ("course_file_extra", "student_list",         "TEXT DEFAULT ''"),
        ("course_file_extra", "custom_tabs",          "TEXT DEFAULT '[]'"),
        # Timetable persisted per-user in DB (survives redeploys)
        ("users",             "timetable_json",          "TEXT DEFAULT NULL"),
    ]
    async with engine.begin() as conn:
        for table, column, col_type in _new_columns:
            try:
                if is_postgres:
                    # PostgreSQL: ADD COLUMN IF NOT EXISTS
                    await conn.execute(
                        __import__("sqlalchemy", fromlist=["text"]).text(
                            f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {column} {col_type}"
                        )
                    )
                else:
                    # SQLite: check if column exists first
                    result = await conn.execute(
                        __import__("sqlalchemy", fromlist=["text"]).text(
                            f"PRAGMA table_info({table})"
                        )
                    )
                    cols = [r[1] for r in result.fetchall()]
                    if column not in cols:
                        await conn.execute(
                            __import__("sqlalchemy", fromlist=["text"]).text(
                                f"ALTER TABLE {table} ADD COLUMN {column} {col_type}"
                            )
                        )
            except Exception as e:
                logger.warning(f"Migration {table}.{column}: {e}")

    logger.info("Database initialised and migrations applied successfully")