import os
import time
import pythoncom
import win32com.client
from typing import Dict, Any

# Word WdStatistic enumeration
WD_STATISTIC_PAGES = 2
WD_STATISTIC_WORDS = 0

# Word WdExportFormat enumeration
WD_EXPORT_FORMAT_PDF = 17

# Word WdExportOptimizeFor
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0  # High quality / standard print

# Word WdExportCreateBookmarks
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1

def convert_docx_to_pdf_word(docx_path: str, pdf_path: str) -> Dict[str, Any]:
    """
    Converts a .docx document to PDF using Microsoft Word COM automation.
    Configured for maximum fidelity and print quality.
    Returns metadata including Word's reported page count.
    """
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)

    if not os.path.exists(abs_docx):
        raise FileNotFoundError(f"File dokumen tidak ditemukan: {abs_docx}")

    # Ensure output directory exists
    os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)
    if os.path.exists(abs_pdf):
        os.remove(abs_pdf)

    pythoncom.CoInitialize()
    word = None
    doc = None
    start_time = time.time()
    word_page_count = 1

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        # Try to enforce application level DoNotCompressImages if supported
        try:
            word.Options.DoNotCompressImages = True
        except Exception:
            pass

        # Open docx in read-only mode without conversion prompts
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

        # Retrieve page count as calculated by Word's layout engine
        try:
            word_page_count = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        except Exception:
            try:
                word_page_count = doc.BuiltInDocumentProperties("Number of Pages").Value
            except Exception:
                word_page_count = 1

        # Export with maximum quality settings
        doc.ExportAsFixedFormat(
            OutputFileName=abs_pdf,
            ExportFormat=WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
            Range=0,  # wdExportAllDocument
            Item=0,   # wdExportDocumentContent
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

    elapsed = round(time.time() - start_time, 2)

    if not os.path.exists(abs_pdf) or os.path.getsize(abs_pdf) == 0:
        raise RuntimeError("Gagal mengekspor PDF dari Microsoft Word.")

    return {
        "success": True,
        "pdf_path": abs_pdf,
        "word_page_count": word_page_count,
        "elapsed_seconds": elapsed,
        "pdf_size_bytes": os.path.getsize(abs_pdf)
    }
