# backend/core/storage.py
"""
Storage Abstraction Layer — Day 12
-----------------------------------
Provides a unified interface for saving and retrieving generated files.

Backends:
  - LocalStorage  : default, writes to generated_docs/ (ephemeral on Railway unless Volume mounted)
  - VolumeStorage : same as LocalStorage but path points to Railway Volume mount

Configuration via env:
  STORAGE_BACKEND=local        # default — generated_docs/ in project root
  STORAGE_PATH=/mnt/storage    # override base path (set this to Railway Volume mount path)

Railway Volume setup:
  1. Railway project → + Add → Volume
  2. Mount path: /mnt/storage
  3. Set env var: STORAGE_PATH=/mnt/storage
  4. Redeploy — files now survive redeploys

Usage (in any service):
    from backend.core.storage import get_storage
    storage = get_storage()

    # Save bytes
    path = storage.save("session_plans", "plan_5.docx", file_bytes)

    # Get path for FileResponse
    path = storage.get_path("session_plans", "plan_5.docx")

    # List files
    files = storage.list_files("question_papers", prefix="qpaper_5_")

    # Get latest file matching a pattern
    latest = storage.get_latest("question_papers", prefix="qpaper_5_", ext=".docx")
"""

import os
import shutil
from pathlib import Path
from typing import Optional
from backend.core.config import STORAGE_PATH, STORAGE_BACKEND
from backend.core.logger import get_logger

logger = get_logger(__name__)


class FileStorage:
    """
    Local/Volume file storage — same implementation, different base paths.
    """

    def __init__(self, base_path: Path):
        self.base = base_path
        self.base.mkdir(parents=True, exist_ok=True)
        logger.info(f"Storage initialised at: {self.base.resolve()}")

    def _dir(self, category: str) -> Path:
        d = self.base / category
        d.mkdir(parents=True, exist_ok=True)
        return d

    def save(self, category: str, filename: str, data: bytes) -> Path:
        """Save raw bytes. Returns the full Path."""
        out = self._dir(category) / filename
        out.write_bytes(data)
        logger.info(f"Stored: {out}")
        return out

    def save_from_path(self, category: str, filename: str, src: Path) -> Path:
        """Move/copy an existing file into storage. Returns destination Path."""
        out = self._dir(category) / filename
        shutil.copy2(str(src), str(out))
        logger.info(f"Copied to storage: {out}")
        return out

    def get_path(self, category: str, filename: str) -> Optional[Path]:
        """Return Path if file exists, else None."""
        p = self._dir(category) / filename
        return p if p.exists() else None

    def list_files(self, category: str, prefix: str = "", ext: str = "") -> list[Path]:
        """List all files in category matching optional prefix + ext."""
        d = self._dir(category)
        files = [f for f in d.iterdir() if f.is_file()]
        if prefix:
            files = [f for f in files if f.name.startswith(prefix)]
        if ext:
            files = [f for f in files if f.suffix == ext]
        return sorted(files)

    def get_latest(self, category: str, prefix: str = "", ext: str = "") -> Optional[Path]:
        """Return the most recently modified file matching criteria, or None."""
        files = self.list_files(category, prefix=prefix, ext=ext)
        return files[-1] if files else None

    def delete(self, category: str, filename: str) -> bool:
        """Delete a file. Returns True if deleted, False if not found."""
        p = self._dir(category) / filename
        if p.exists():
            p.unlink()
            return True
        return False

    def exists(self, category: str, filename: str) -> bool:
        return (self._dir(category) / filename).exists()

    def get_url_path(self, category: str, filename: str) -> str:
        """Return the URL path for serving this file via FastAPI /files endpoint."""
        return f"/files/{category}/{filename}"


# ── Singleton ─────────────────────────────────────────────────────────────────

_storage_instance: Optional[FileStorage] = None


def get_storage() -> FileStorage:
    """
    Returns the configured storage singleton.
    Call once per request or reuse — it's thread/async safe for reads.
    """
    global _storage_instance
    if _storage_instance is None:
        if STORAGE_PATH:
            base = Path(STORAGE_PATH)
            logger.info(f"Using configured storage path: {base}")
        else:
            # Default: generated_docs/ in project root (same as before Day 12)
            base = Path("generated_docs")
            logger.info("Using default local storage: generated_docs/")
        _storage_instance = FileStorage(base)
    return _storage_instance


def reset_storage():
    """Force re-initialise storage (used in tests)."""
    global _storage_instance
    _storage_instance = None
