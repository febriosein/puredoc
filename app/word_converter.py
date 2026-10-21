import os
import sys
import time
import shutil
import subprocess
from typing import Dict, Any
import httpx
from app.config import GOTENBERG_URL

# Word WdStatistic enumeration
WD_STATISTIC_PAGES = 2
WD_STATISTIC_WORDS = 0
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1

def _convert_with_word_com(abs_docx: str, abs_pdf: str) -> int:
    """Windows-only: Microsoft Word COM Automation."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError("Pustaka pywin32 tidak terpasang di sistem ini.")

    pythoncom.CoInitialize()
    word = None
    doc = None
    word_page_count = 1

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        try:
            word.Options.DoNotCompressImages = True
        except Exception:
            pass

        doc = word.Documents.Open(
            FileName=abs_docx,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False
        )

        try:
            doc.DoNotCompressImages = True
        except Exception:
            pass

        try:
            word_page_count = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        except Exception:
            try:
                word_page_count = doc.BuiltInDocumentProperties("Number of Pages").Value
            except Exception:
                word_page_count = 1

        doc.ExportAsFixedFormat(
            OutputFileName=abs_pdf,
            ExportFormat=WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
            Range=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return word_page_count

def _convert_with_gotenberg(abs_docx: str, abs_pdf: str, gotenberg_url: str) -> int:
    """Converts DOCX to PDF using remote Gotenberg LibreOffice service."""
    endpoint = f"{gotenberg_url.rstrip('/')}/forms/libreoffice/convert"
    filename = os.path.basename(abs_docx)

    with open(abs_docx, "rb") as f:
        file_bytes = f.read()

    files = {
        "files": (filename, file_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    }
    data = {
        "losslessImageCompression": "true",
        "reduceImageResolution": "false"
    }

    with httpx.Client(timeout=120.0) as client:
        response = client.post(endpoint, files=files, data=data)
        if response.status_code != 200:
            raise RuntimeError(f"Gotenberg error {response.status_code}: {response.text}")
        with open(abs_pdf, "wb") as out_f:
            out_f.write(response.content)

    return 1

def _convert_with_libreoffice_cli(abs_docx: str, abs_pdf: str) -> int:
    """Converts DOCX to PDF using local LibreOffice headless CLI."""
    soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_cmd and sys.platform == "win32":
        for p in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]:
            if os.path.exists(p):
                soffice_cmd = p
                break

    if not soffice_cmd:
        raise RuntimeError("LibreOffice tidak ditemukan pada sistem.")

    out_dir = os.path.dirname(abs_pdf)
    cmd = [soffice_cmd, "--headless", "--convert-to", "pdf", "--outdir", out_dir, abs_docx]
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
    if res.returncode != 0:
        raise RuntimeError(f"LibreOffice gagal: {res.stderr.decode('utf-8', errors='ignore')}")

    # LibreOffice outputs with same base name in out_dir
    expected_out = os.path.join(out_dir, os.path.splitext(os.path.basename(abs_docx))[0] + ".pdf")
    if os.path.exists(expected_out) and expected_out != abs_pdf:
        shutil.move(expected_out, abs_pdf)

    return 1

def convert_docx_to_pdf_word(docx_path: str, pdf_path: str) -> Dict[str, Any]:
    """
    Multi-Tier Conversion Engine:
    1. If on Windows and Word is installed -> Native Word COM Automation (100% Identical Layout).
    2. If GOTENBERG_URL is configured (e.g. on Vercel/Cloud) -> Gotenberg LibreOffice Cloud Service.
    3. If local LibreOffice CLI is available -> Local LibreOffice Headless.
    """
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)

    if not os.path.exists(abs_docx):
        raise FileNotFoundError(f"File dokumen tidak ditemukan: {abs_docx}")

    os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)
    if os.path.exists(abs_pdf):
        os.remove(abs_pdf)

    start_time = time.time()
    word_page_count = 1
    engine_used = "none"

    # Tier 1: Windows Native Microsoft Word
    if sys.platform == "win32":
        try:
            word_page_count = _convert_with_word_com(abs_docx, abs_pdf)
            engine_used = "Microsoft Word Native COM"
        except Exception as e:
            # If Word COM failed, fall through to Gotenberg or LibreOffice
            pass

    # Tier 2: Gotenberg Cloud Service (Ideal for Vercel / Cloud deployments)
    if not os.path.exists(abs_pdf) and GOTENBERG_URL:
        try:
            word_page_count = _convert_with_gotenberg(abs_docx, abs_pdf, GOTENBERG_URL)
            engine_used = f"Gotenberg Service ({GOTENBERG_URL})"
        except Exception as e:
            pass

    # Tier 3: Local LibreOffice CLI
    if not os.path.exists(abs_pdf):
        try:
            word_page_count = _convert_with_libreoffice_cli(abs_docx, abs_pdf)
            engine_used = "LibreOffice Headless CLI"
        except Exception:
            pass

    if not os.path.exists(abs_pdf) or os.path.getsize(abs_pdf) == 0:
        if sys.platform != "win32" and not GOTENBERG_URL:
            raise RuntimeError(
                "Layanan konversi dokumen belum tersedia di server Linux/Vercel. "
                "Silakan pasang environment variable 'GOTENBERG_URL' di Vercel Dashboard "
                "(misal menggunakan instance Gotenberg di Railway/Render), "
                "atau jalankan aplikasi di komputer Windows dengan Microsoft Word."
            )
        raise RuntimeError("Gagal mengekspor PDF dari seluruh mesin konversi yang tersedia.")

    elapsed = round(time.time() - start_time, 2)
    return {
        "success": True,
        "engine_used": engine_used,
        "pdf_path": abs_pdf,
        "word_page_count": word_page_count,
        "elapsed_seconds": elapsed,
        "pdf_size_bytes": os.path.getsize(abs_pdf)
    }
