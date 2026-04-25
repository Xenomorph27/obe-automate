# Dockerfile — single stage, Railway-compatible
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Runtime libs needed by PyMuPDF and ReportLab
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 \
    libgl1 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies first (Docker cache layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source
COPY backend/ ./backend/
COPY frontend/ ./frontend/

# Create output directories
RUN mkdir -p generated_docs/session_plans \
             generated_docs/evaluation_plans \
             generated_docs/attainment_reports \
             generated_docs/nba_reports

ENV PORT=8000
EXPOSE $PORT

CMD uvicorn backend.main:app --host 0.0.0.0 --port $PORT --workers 1
