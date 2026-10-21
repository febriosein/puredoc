import os
import time
import shutil
import logging
from pathlib import Path
from app.config import STORAGE_DIR, UPLOAD_DIR, OUTPUT_DIR, TEMP_DIR, FILE_TTL_SECONDS

logger = logging.getLogger("puredoc.cleanup")

def cleanup_old_files(max_age_seconds: int = FILE_TTL_SECONDS) -> int:
    """
    Deletes files older than max_age_seconds in uploads, outputs, and temp folders.
    Returns number of deleted items.
    """
    now = time.time()
    deleted_count = 0

    target_dirs = [UPLOAD_DIR, OUTPUT_DIR, TEMP_DIR]
    for d in target_dirs:
        if not d.exists():
            continue
        for item in d.iterdir():
            try:
                mtime = item.stat().st_mtime
                if (now - mtime) > max_age_seconds:
                    if item.is_file():
                        item.unlink(missing_ok=True)
                        deleted_count += 1
                    elif item.is_dir():
                        shutil.rmtree(item, ignore_errors=True)
                        deleted_count += 1
            except Exception as e:
                logger.warning(f"Error cleaning up {item}: {e}")

    return deleted_count

def delete_job_files(job_id: str) -> bool:
    """
    Immediately deletes all files associated with a specific job_id across all storage directories.
    (FR-20: Tombol 'Hapus sekarang').
    """
    target_dirs = [UPLOAD_DIR, OUTPUT_DIR, TEMP_DIR]
    deleted_any = False

    for d in target_dirs:
        if not d.exists():
            continue
        for item in d.iterdir():
            if item.name.startswith(job_id):
                try:
                    if item.is_file():
                        item.unlink(missing_ok=True)
                        deleted_any = True
                    elif item.is_dir():
                        shutil.rmtree(item, ignore_errors=True)
                        deleted_any = True
                except Exception as e:
                    logger.warning(f"Failed to delete job item {item}: {e}")

    return deleted_any
