import os
import tempfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# On Vercel or cloud serverless, root directory is read-only.
# We automatically route storage to /tmp in serverless environments.
IS_VERCEL = bool(os.environ.get("VERCEL") or os.environ.get("AWS_LAMBDA_FUNCTION_NAME"))
if IS_VERCEL:
    STORAGE_DIR = Path(tempfile.gettempdir()) / "puredoc"
else:
    STORAGE_DIR = BASE_DIR / "storage"

UPLOAD_DIR = STORAGE_DIR / "uploads"
OUTPUT_DIR = STORAGE_DIR / "outputs"
TEMP_DIR = STORAGE_DIR / "temp"

# Ensure directories exist
for directory in [STORAGE_DIR, UPLOAD_DIR, OUTPUT_DIR, TEMP_DIR]:
    try:
        directory.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

MAX_FILE_SIZE_MB = 100
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

# Auto cleanup after 1 hour (3600 seconds)
FILE_TTL_SECONDS = 3600

# Server settings
HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8000"))

# Cloud Conversion Service (Gotenberg or external LibreOffice service)
GOTENBERG_URL = os.environ.get("GOTENBERG_URL", "").rstrip("/")
