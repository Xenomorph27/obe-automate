# backend/core/exceptions.py


class OBEException(Exception):
    """Base exception for all OBE Automate errors."""
    def __init__(self, message: str, status_code: int = 500):
        self.message = message
        self.status_code = status_code
        super().__init__(self.message)


class FileValidationError(OBEException):
    """Raised when an uploaded file fails validation (wrong type, too large)."""
    def __init__(self, message: str):
        super().__init__(message, status_code=400)


class ExtractionError(OBEException):
    """Raised when PDF text extraction fails."""
    def __init__(self, message: str):
        super().__init__(message, status_code=422)


class LLMError(OBEException):
    """Raised when the Gemini API call fails or returns bad data."""
    def __init__(self, message: str):
        super().__init__(message, status_code=503)