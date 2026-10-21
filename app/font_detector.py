import os
import sys
from typing import Set, Dict, Any, List

if sys.platform == "win32":
    try:
        import winreg
    except ImportError:
        winreg = None
else:
    winreg = None

_SYSTEM_FONTS_CACHE: Set[str] = set()

# Standard metric-compatible and popular fallback mappings
FONT_EQUIVALENTS = {
    "calibri": ["carlito", "arial", "segoe ui"],
    "cambria": ["caladea", "georgia", "times new roman"],
    "times new roman": ["times", "liberation serif", "tinos"],
    "arial": ["liberation sans", "arimo", "helvetica"],
    "courier new": ["courier", "liberation mono", "consolas"],
}

def get_installed_fonts() -> Set[str]:
    """
    Scans system font registry (Windows) or font directories (Linux/macOS)
    and returns a set of lowercase font family names.
    Caches results for subsequent calls.
    """
    global _SYSTEM_FONTS_CACHE
    if _SYSTEM_FONTS_CACHE:
        return _SYSTEM_FONTS_CACHE

    fonts = set()

    if winreg is not None and sys.platform == "win32":
        roots = [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]
        subkeys = [
            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Fonts",
        ]

        for root in roots:
            for subkey in subkeys:
                try:
                    with winreg.OpenKey(root, subkey) as key:
                        i = 0
                        while True:
                            try:
                                name, val, _ = winreg.EnumValue(key, i)
                                clean_name = name.split(" (")[0].strip()
                                for part in clean_name.split("&"):
                                    p = part.strip().lower()
                                    if p:
                                        fonts.add(p)
                                i += 1
                            except OSError:
                                break
                except Exception:
                    continue
    else:
        # Linux / container fallback: scan font directories if present
        linux_dirs = ["/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts")]
        for d in linux_dirs:
            if os.path.exists(d):
                for root_dir, _, files in os.walk(d):
                    for f in files:
                        base = os.path.splitext(f)[0].lower()
                        fonts.add(base)

    # Also add standard universally available Windows fonts
    standard_windows_fonts = {
        "calibri", "calibri light", "cambria", "cambria math", "arial", "arial black",
        "times new roman", "segoe ui", "segoe ui symbol", "tahoma", "verdana",
        "georgia", "comic sans ms", "trebuchet ms", "impact", "courier new",
        "consolas", "garamond", "century gothic", "franklin gothic", "palatino linotype",
        "book antiqua", "aptos", "aptos display" # Modern Office 365 default fonts
    }
    fonts.update(standard_windows_fonts)

    _SYSTEM_FONTS_CACHE = fonts
    return _SYSTEM_FONTS_CACHE

def check_fonts(document_fonts: List[str]) -> Dict[str, Any]:
    """
    Verifies a list of fonts found in a .docx against the installed Windows system fonts.
    Returns:
    {
        "all_available": bool,
        "available_fonts": [...],
        "missing_fonts": [...],
        "warnings": [...]
    }
    """
    installed = get_installed_fonts()
    available = []
    missing = []
    warnings = []

    for font in document_fonts:
        f_norm = font.strip().lower()
        if not f_norm:
            continue
        
        # Check direct match or partial family match
        is_installed = False
        if f_norm in installed:
            is_installed = True
        else:
            # Check if any installed font starts with this name (e.g. 'Calibri Light' for 'Calibri')
            for inst in installed:
                if f_norm == inst or f_norm in inst or inst in f_norm:
                    is_installed = True
                    break

        if is_installed:
            available.append(font)
        else:
            fallback = FONT_EQUIVALENTS.get(f_norm, ["Arial", "Times New Roman"])
            missing.append({
                "font": font,
                "suggested_fallback": fallback[0] if fallback else "Standard Font"
            })
            warnings.append(
                f"Font '{font}' tidak ditemukan pada server/sistem. Word mungkin menggunakan font pengganti."
            )

    return {
        "all_available": len(missing) == 0,
        "available_fonts": available,
        "missing_fonts": missing,
        "warnings": warnings,
        "total_doc_fonts": len(document_fonts)
    }
