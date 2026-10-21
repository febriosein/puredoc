import io
import os
import zipfile
import hashlib
from typing import Dict, Any, List, Tuple
from PIL import Image
import pymupdf

def _get_thumbnail_data(image_bytes: bytes) -> List[int]:
    """Generates a 16x16 grayscale thumbnail pixel list for perceptual matching."""
    try:
        with Image.open(io.BytesIO(image_bytes)) as im:
            im_gray = im.convert("L").resize((16, 16), Image.Resampling.BILINEAR)
            return list(im_gray.getdata())
    except Exception:
        return [0] * 256

def _calculate_mse(t1: List[int], t2: List[int]) -> float:
    """Calculates Mean Squared Error between two 16x16 grayscale thumbnails."""
    if len(t1) != len(t2) or len(t1) == 0:
        return float("inf")
    return sum((a - b) ** 2 for a, b in zip(t1, t2)) / len(t1)

def extract_docx_media(docx_path: str) -> List[Dict[str, Any]]:
    """Extracts all media from docx with dimensions, byte hashes, and thumbnails."""
    docx_media = []
    with zipfile.ZipFile(docx_path, "r") as z:
        for name in sorted(z.namelist()):
            if not name.startswith("word/media/"):
                continue
            raw_bytes = z.read(name)
            sha256 = hashlib.sha256(raw_bytes).hexdigest()
            ext = os.path.splitext(name)[1].lower().replace(".", "")
            
            width, height = 0, 0
            is_bitmap = False
            format_name = ext.upper()
            thumb = []

            try:
                with Image.open(io.BytesIO(raw_bytes)) as pil_img:
                    width, height = pil_img.size
                    format_name = pil_img.format or format_name
                    is_bitmap = True
                thumb = _get_thumbnail_data(raw_bytes)
            except Exception:
                is_bitmap = False

            docx_media.append({
                "zip_path": name,
                "filename": os.path.basename(name),
                "bytes": raw_bytes,
                "format": format_name,
                "is_bitmap": is_bitmap,
                "width": width,
                "height": height,
                "aspect_ratio": (width / height) if height > 0 else 1.0,
                "sha256": sha256,
                "thumb": thumb,
                "size_bytes": len(raw_bytes)
            })
    return docx_media

def restore_pdf_original_images(docx_path: str, pdf_path: str) -> Dict[str, Any]:
    """
    Scans the exported PDF for images. If any image was downsampled or re-encoded by Word,
    it maps it to the original docx image and restores it to full lossless / pass-through quality.
    """
    docx_media = extract_docx_media(docx_path)
    if not docx_media:
        return {
            "restored_count": 0,
            "total_images": 0,
            "details": []
        }

    doc = pymupdf.open(pdf_path)
    replaced_count = 0
    details = []
    used_docx_indices = set()

    # Iterate through all pages and their image XObjects
    for page_num in range(len(doc)):
        page = doc[page_num]
        image_list = page.get_images(full=True)

        for img_info in image_list:
            xref = img_info[0]
            try:
                extracted = doc.extract_image(xref)
            except Exception:
                continue

            pdf_img_bytes = extracted["image"]
            pdf_w = extracted["width"]
            pdf_h = extracted["height"]
            pdf_ext = extracted["ext"]
            pdf_hash = hashlib.sha256(pdf_img_bytes).hexdigest()
            pdf_ar = (pdf_w / pdf_h) if pdf_h > 0 else 1.0
            pdf_thumb = _get_thumbnail_data(pdf_img_bytes)

            # Find matching docx image
            best_match = None
            best_match_idx = -1
            min_mse = float("inf")

            for idx, d_img in enumerate(docx_media):
                if not d_img["is_bitmap"]:
                    continue
                # Aspect ratio tolerance check (~10%)
                if abs(d_img["aspect_ratio"] - pdf_ar) > 0.12:
                    continue

                mse = _calculate_mse(d_img["thumb"], pdf_thumb)
                if mse < min_mse:
                    min_mse = mse
                    best_match = d_img
                    best_match_idx = idx

            # If match found with acceptable visual similarity threshold
            if best_match and min_mse < 3000:
                is_exact_hash = (pdf_hash == best_match["sha256"])
                is_exact_dim = (pdf_w == best_match["width"] and pdf_h == best_match["height"])

                # If Word downsampled or modified the image, restore the original
                needs_restoration = not is_exact_dim or not is_exact_hash

                if needs_restoration:
                    try:
                        page.replace_image(xref, stream=best_match["bytes"])
                        replaced_count += 1
                        used_docx_indices.add(best_match_idx)
                        details.append({
                            "page": page_num + 1,
                            "xref": xref,
                            "filename": best_match["filename"],
                            "original_resolution": f"{best_match['width']}×{best_match['height']}",
                            "downsampled_resolution": f"{pdf_w}×{pdf_h}",
                            "action": "Restored to 100% Original Resolution",
                            "status": "PASS_THROUGH_RESTORED"
                        })
                    except Exception as e:
                        details.append({
                            "page": page_num + 1,
                            "xref": xref,
                            "filename": best_match["filename"],
                            "action": f"Restore failed: {str(e)}",
                            "status": "RESTORE_ERROR"
                        })
                else:
                    details.append({
                        "page": page_num + 1,
                        "xref": xref,
                        "filename": best_match["filename"],
                        "original_resolution": f"{best_match['width']}×{best_match['height']}",
                        "action": "Already 100% Original (Bit-identical)",
                        "status": "BIT_IDENTICAL"
                    })

    # Save the updated PDF with restored images
    if replaced_count > 0:
        temp_out = pdf_path + ".restored.tmp"
        doc.save(temp_out, garbage=3, deflate=True)
        doc.close()
        # Replace original pdf
        os.replace(temp_out, pdf_path)
    else:
        doc.close()

    return {
        "restored_count": replaced_count,
        "total_images": len(docx_media),
        "details": details
    }
