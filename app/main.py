import os
import uuid
import asyncio
import logging
from pathlib import Path
from typing import Dict, Any

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Query
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from app.config import (
    BASE_DIR,
    UPLOAD_DIR,
    OUTPUT_DIR,
    TEMP_DIR,
    MAX_FILE_SIZE_BYTES,
    MAX_FILE_SIZE_MB,
    FILE_TTL_SECONDS
)
from app.docx_analyzer import analyze_docx, validate_docx_file
from app.word_converter import convert_docx_to_pdf_word
from app.pdf_restorer import restore_pdf_original_images
from app.verifier import verify_conversion
from app.cleanup import cleanup_old_files, delete_job_files
from app.font_detector import get_installed_fonts

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("puredoc")

app = FastAPI(
    title="PureDoc PDF",
    description="Konverter Word ke PDF Tanpa Kompresi Gambar & Tata Letak Identik",
    version="1.0"
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job state cache
JOBS: Dict[str, Dict[str, Any]] = {}

@app.on_event("startup")
async def startup_event():
    logger.info("PureDoc PDF backend initialized.")
    # Pre-warm font cache in background
    asyncio.create_task(asyncio.to_thread(get_installed_fonts))
    # Start periodic background cleanup
    async def periodic_cleanup():
        while True:
            await asyncio.sleep(600)  # Check every 10 minutes
            cleaned = cleanup_old_files(FILE_TTL_SECONDS)
            if cleaned > 0:
                logger.info(f"Auto-cleanup removed {cleaned} expired items.")
    asyncio.create_task(periodic_cleanup())

@app.get("/api/health")
def health_check():
    fonts = get_installed_fonts()
    return {
        "status": "healthy",
        "service": "PureDoc PDF",
        "word_engine": "Microsoft Word Automation (Native)",
        "installed_system_fonts_count": len(fonts),
        "max_upload_size_mb": MAX_FILE_SIZE_MB,
        "auto_cleanup_ttl_hours": round(FILE_TTL_SECONDS / 3600, 1)
    }

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    # 1. Validate file extension
    filename = file.filename or "document.docx"
    ext = os.path.splitext(filename)[1].lower()
    if ext != ".docx":
        raise HTTPException(
            status_code=400,
            detail="Hanya file berekstensi .docx yang didukung pada versi ini."
        )

    # 2. Read and validate file content & size
    job_id = str(uuid.uuid4())
    safe_name = f"{job_id}_{filename}"
    upload_path = UPLOAD_DIR / safe_name

    content = await file.read()
    if len(content) > MAX_FILE_SIZE_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Ukuran file melebihi batas maksimal ({MAX_FILE_SIZE_MB} MB)."
        )

    if len(content) < 4 or content[:4] != b"PK\x03\x04":
        raise HTTPException(
            status_code=400,
            detail="File rusak atau bukan format .docx yang valid."
        )

    with open(upload_path, "wb") as f:
        f.write(content)

    # 3. Analyze document (FR-4, FR-5, FR-6)
    try:
        analysis = analyze_docx(str(upload_path))
    except Exception as e:
        upload_path.unlink(missing_ok=True)
        raise HTTPException(
            status_code=400,
            detail=f"Gagal menganalisis isi file Word: {str(e)}"
        )

    JOBS[job_id] = {
        "job_id": job_id,
        "original_filename": filename,
        "docx_path": str(upload_path),
        "pdf_path": None,
        "analysis": analysis,
        "status": "ANALYZED",
        "created_at": asyncio.get_event_loop().time()
    }

    return {
        "success": True,
        "job_id": job_id,
        "filename": filename,
        "analysis": analysis
    }

@app.post("/api/jobs/{job_id}/convert")
async def convert_document(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Sesi pekerjaan tidak ditemukan.")

    docx_path = job["docx_path"]
    if not os.path.exists(docx_path):
        raise HTTPException(status_code=404, detail="File sumber Word telah kadaluarsa atau dihapus.")

    base_name = os.path.splitext(job["original_filename"])[0]
    pdf_filename = f"{job_id}_{base_name}.pdf"
    pdf_path = OUTPUT_DIR / pdf_filename

    job["status"] = "CONVERTING"

    try:
        # Step 1: Native Microsoft Word High-Fidelity Conversion
        conv_res = await asyncio.to_thread(convert_docx_to_pdf_word, docx_path, str(pdf_path))
        word_pages = conv_res.get("word_page_count", 1)

        # Step 2: Lossless Image Restoration (Pass-through original raw bytes)
        restore_res = await asyncio.to_thread(restore_pdf_original_images, docx_path, str(pdf_path))

        # Step 3: Automated Post-conversion Verification
        verification = await asyncio.to_thread(
            verify_conversion,
            docx_path,
            str(pdf_path),
            word_pages,
            job["analysis"]
        )

        job["pdf_path"] = str(pdf_path)
        job["status"] = "COMPLETED"
        job["conversion_result"] = conv_res
        job["restoration_result"] = restore_res
        job["verification"] = verification

        return {
            "success": True,
            "job_id": job_id,
            "filename": f"{base_name}.pdf",
            "verification": verification,
            "restoration": restore_res,
            "download_url": f"/api/jobs/{job_id}/pdf?download=true",
            "preview_url": f"/api/jobs/{job_id}/pdf?download=false"
        }

    except Exception as e:
        job["status"] = "ERROR"
        job["error"] = str(e)
        logger.exception(f"Conversion error for job {job_id}: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"Terjadi kesalahan saat memproses konversi: {str(e)}"
        )

@app.get("/api/jobs/{job_id}/pdf")
def get_pdf(job_id: str, download: bool = Query(False)):
    job = JOBS.get(job_id)
    if not job or not job.get("pdf_path") or not os.path.exists(job["pdf_path"]):
        # Check direct file match in outputs
        candidates = list(OUTPUT_DIR.glob(f"{job_id}_*.pdf"))
        if not candidates:
            raise HTTPException(status_code=404, detail="File PDF tidak ditemukan atau sudah dihapus.")
        pdf_path = str(candidates[0])
        filename = candidates[0].name.replace(f"{job_id}_", "")
    else:
        pdf_path = job["pdf_path"]
        base_name = os.path.splitext(job["original_filename"])[0]
        filename = f"{base_name}.pdf"

    disposition = "attachment" if download else "inline"
    return FileResponse(
        path=pdf_path,
        media_type="application/pdf",
        filename=filename,
        headers={"Content-Disposition": f'{disposition}; filename="{filename}"'}
    )

@app.get("/api/jobs/{job_id}/report")
def get_quality_report(job_id: str):
    job = JOBS.get(job_id)
    if not job or not job.get("verification"):
        raise HTTPException(status_code=404, detail="Laporan kualitas belum tersedia untuk sesi ini.")

    return {
        "job_id": job_id,
        "filename": job["original_filename"],
        "verification": job["verification"],
        "restoration": job.get("restoration_result"),
        "analysis": job.get("analysis")
    }

@app.delete("/api/jobs/{job_id}")
def delete_job(job_id: str):
    success = delete_job_files(job_id)
    if job_id in JOBS:
        del JOBS[job_id]
    return {
        "success": True,
        "message": "Seluruh file dokumen dan PDF terkait telah dihapus permanen dari server."
    }

# Mount static files
static_dir = BASE_DIR / "static"
if static_dir.exists():
    app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")

@app.get("/")
def serve_index():
    index_file = BASE_DIR / "static" / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file))
    return JSONResponse({"message": "PureDoc PDF API running. Static UI not yet created."})
