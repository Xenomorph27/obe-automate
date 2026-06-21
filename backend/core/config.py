# backend/core/config.py
import os
from dotenv import load_dotenv

load_dotenv()

# App settings
APP_NAME = "OBE Automate"
APP_VERSION = "0.1.0"
DEBUG = os.getenv("DEBUG", "false").lower() == "true"

# API Keys — LLM fallback chain: Gemini → Groq → OpenAI
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")

# File upload settings
MAX_FILE_SIZE_MB = 25
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
ALLOWED_FILE_TYPES = ["application/pdf"]

# Database
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./obe_automate.db")

# File Storage — Day 12
# STORAGE_PATH: set to Railway Volume mount path (e.g. /mnt/storage) for persistence
# Leave empty to use default generated_docs/ (ephemeral on Railway)
STORAGE_PATH = os.getenv("STORAGE_PATH", "")
STORAGE_BACKEND = os.getenv("STORAGE_BACKEND", "local")  # reserved for future S3 backend