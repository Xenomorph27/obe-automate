# backend/routes/syllabus.py
from fastapi import APIRouter, UploadFile, File, HTTPException
from backend.services.syllabus_service import SyllabusService
from backend.core.exceptions import OBEException
from backend.core.logger import get_logger

logger = get_logger(__name__)
router = APIRouter(prefix="/syllabus", tags=["Syllabus"])

syllabus_service = SyllabusService()


@router.post("/upload")
async def upload_syllabus(file: UploadFile = File(...)):
    """
    Upload a syllabus PDF.
    Extracts topics and Course Outcomes using Gemini AI.
    """
    logger.info(f"Received file upload: {file.filename}")

    try:
        result = await syllabus_service.process_syllabus(file)
        logger.info(f"Successfully processed: {file.filename}")
        return {
            "status": "success",
            "filename": file.filename,
            "data": result
        }

    except OBEException as e:
        logger.error(f"OBE error processing {file.filename}: {e.message}")
        raise HTTPException(status_code=e.status_code, detail=e.message)

    except Exception as e:
        logger.exception(f"Unexpected error processing {file.filename}")
        raise HTTPException(status_code=500, detail="Internal server error")