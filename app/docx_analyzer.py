import os
import zipfile
import hashlib
import io
import xml.etree.ElementTree as ET
from typing import Dict, Any, List
from PIL import Image
from app.font_detector import check_fonts

XML_NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "v": "urn:schemas-microsoft-com:vml"
}

def validate_docx_file(file_path: str) -> bool:
    """Validates that the file exists, has zip magic bytes, and contains word/document.xml."""
    if not os.path.exists(file_path):
        return False
    if os.path.getsize(file_path) < 4:
        return False
    with open(file_path, "rb") as f:
        magic = f.read(4)
        if magic != b"PK\x03\x04":
            return False
    try:
        with zipfile.ZipFile(file_path, "r") as z:
            names = z.namelist()
            return "word/document.xml" in names
    except Exception:
        return False

def analyze_docx(file_path: str) -> Dict[str, Any]:
    """
    Analyzes docx package:
    - Extracts media image metadata (format, dimensions, hash, size)
    - Extracts fonts used
    - Checks complex elements (SmartArt, equation, tables, textboxes, charts)
    - Runs font availability check
    """
    if not validate_docx_file(file_path):
        raise ValueError("File bukan dokumen Word (.docx) yang valid atau file rusak.")

    images_meta = []
    total_image_bytes = 0
    min_res = None
    max_res = None

    fonts_found = set()
    complex_elements = {
        "tables_count": 0,
        "equations_count": 0,
        "textboxes_count": 0,
        "charts_count": 0,
        "smartart_count": 0,
        "shapes_count": 0
    }

    with zipfile.ZipFile(file_path, "r") as z:
        names = z.namelist()

        # 1. Analyze Media Files
        media_names = [n for n in names if n.startswith("word/media/")]
        for m_name in sorted(media_names):
            raw_bytes = z.read(m_name)
            byte_len = len(raw_bytes)
            total_image_bytes += byte_len
            sha256 = hashlib.sha256(raw_bytes).hexdigest()
            ext = os.path.splitext(m_name)[1].lower().replace(".", "")

            width, height = 0, 0
            is_bitmap = False
            format_name = ext.upper()

            try:
                # Open with Pillow to get dimensions and format
                with Image.open(io.BytesIO(raw_bytes)) as pil_img:
                    width, height = pil_img.size
                    format_name = pil_img.format or format_name
                    is_bitmap = True

                    total_pixels = width * height
                    if min_res is None or total_pixels < (min_res["width"] * min_res["height"]):
                        min_res = {"name": m_name, "width": width, "height": height, "pixels": total_pixels}
                    if max_res is None or total_pixels > (max_res["width"] * max_res["height"]):
                        max_res = {"name": m_name, "width": width, "height": height, "pixels": total_pixels}
            except Exception:
                # Non-standard or vector formats (EMF, WMF, SVG, etc.)
                is_bitmap = False

            images_meta.append({
                "filename": os.path.basename(m_name),
                "full_path": m_name,
                "format": format_name,
                "is_bitmap": is_bitmap,
                "width": width,
                "height": height,
                "resolution_str": f"{width}×{height} px" if width > 0 else "Vektor / Asli",
                "size_bytes": byte_len,
                "size_kb": round(byte_len / 1024, 1),
                "sha256": sha256
            })

        # 2. Extract Fonts from fontTable.xml and document.xml
        if "word/fontTable.xml" in names:
            try:
                tree = ET.fromstring(z.read("word/fontTable.xml"))
                for font_elem in tree.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}font"):
                    name = font_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")
                    if name:
                        fonts_found.add(name)
            except Exception:
                pass

        if "word/document.xml" in names:
            try:
                doc_tree = ET.fromstring(z.read("word/document.xml"))
                # Fonts in run properties
                for rFonts in doc_tree.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts"):
                    for attr in ["ascii", "hAnsi", "cs", "eastAsia"]:
                        fname = rFonts.get(f"{{{XML_NAMESPACES['w']}}}{attr}")
                        if fname:
                            fonts_found.add(fname)

                # Complex elements detection
                tables = doc_tree.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl")
                complex_elements["tables_count"] = len(tables)

                # Equations
                eqs = doc_tree.findall(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath")
                complex_elements["equations_count"] = len(eqs)

                # Textboxes
                txbx = doc_tree.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent")
                complex_elements["textboxes_count"] = len(txbx)

                # Drawings/shapes
                drawings = doc_tree.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing")
                complex_elements["shapes_count"] = len(drawings)
            except Exception:
                pass

        # SmartArt and Charts detection from zip package structure
        complex_elements["smartart_count"] = len([n for n in names if "diagrams/" in n])
        complex_elements["charts_count"] = len([n for n in names if "charts/" in n])

    # Filter out empty or standard generic font tags
    clean_fonts = [f for f in sorted(list(fonts_found)) if f and not f.startswith("@")]

    # Check font availability on the Windows host
    font_status = check_fonts(clean_fonts)

    # Human-readable summary for FR-6
    img_count = len(images_meta)
    if img_count > 0:
        summary_msg = f"{img_count} gambar ditemukan. Seluruh resolusi asli akan dipertahankan 100% tanpa kompresi."
    else:
        summary_msg = "Dokumen teks murni (tidak ada gambar yang terdeteksi)."

    return {
        "is_valid": True,
        "file_name": os.path.basename(file_path),
        "file_size_bytes": os.path.getsize(file_path),
        "file_size_mb": round(os.path.getsize(file_path) / (1024 * 1024), 2),
        "image_count": img_count,
        "total_image_bytes": total_image_bytes,
        "min_resolution": f"{min_res['width']}×{min_res['height']} px" if min_res else "-",
        "max_resolution": f"{max_res['width']}×{max_res['height']} px" if max_res else "-",
        "images": images_meta,
        "fonts": clean_fonts,
        "font_status": font_status,
        "complex_elements": complex_elements,
        "summary_message": summary_msg
    }
