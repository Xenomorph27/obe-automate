# ── Dockerfile ──────────────────────────────────────────────────────────
# Multi-stage build: keeps final image lean (~200MB vs ~800MB)
# Stage 1: dependency install
# Stage 2: runtime image

FROM python:3.11-slim AS base

# Prevent .pyc files and enable unbuffered stdout (important for Railway logs)
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

# ── Stage 1: install dependencies ───────────────────────────────────────
FROM base AS builder

# System libs needed by PyMuPDF and ReportLab
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libffi-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --prefix=/install -r requirements.txt

# ── Stage 2: runtime ────────────────────────────────────────────────────
FROM base AS runtime

# Copy installed packages from builder
COPY --from=builder /install /usr/local

# System libs needed at runtime by PyMuPDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 \
    libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy source code
COPY backend/ ./backend/
COPY frontend/ ./frontend/

# Create output directories (Railway has ephemeral filesystem but these
# are needed for the process to start; use Railway volumes for persistence)
RUN mkdir -p generated_docs/session_plans \
             generated_docs/evaluation_plans \
             generated_docs/attainment_reports \
             generated_docs/nba_reports

# Non-root user for security
RUN adduser --disabled-password --gecos '' appuser && chown -R appuser /app
USER appuser

# Railway injects PORT env var; default to 8000
ENV PORT=8000

EXPOSE $PORT

# Use shell form so $PORT is expanded
CMD uvicorn backend.main:app --host 0.0.0.0 --port $PORT --workers 1
