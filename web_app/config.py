"""Web app configuration for excel_datamerger."""
import os
from pathlib import Path


class WebConfig:
    """Configuration values, overridable via environment variables."""

    # Authentication
    USERNAME: str = os.getenv("MERGER_USERNAME", "admin")
    PASSWORD: str = os.getenv("MERGER_PASSWORD", "admin123")

    # Flask secret key for session cookies
    SECRET_KEY: str = os.getenv("MERGER_SECRET_KEY", "replace-this-secret")

    # Upload and output settings
    UPLOAD_ROOT: Path = Path(
        os.getenv("MERGER_UPLOAD_ROOT", "/tmp/excel_datamerger")
    )
    ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv", ".txt"}
    MAX_CONTENT_LENGTH: int = int(
        float(os.getenv("MERGER_MAX_CONTENT_MB", "50")) * 1024 * 1024
    )

    # Cleanup policy (in minutes) for temporary results; currently informational
    CLEANUP_MINUTES: int = int(os.getenv("MERGER_CLEANUP_MINUTES", "120"))
