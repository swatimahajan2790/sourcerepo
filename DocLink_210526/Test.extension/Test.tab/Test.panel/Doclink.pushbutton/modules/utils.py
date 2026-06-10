# -*- coding: utf-8 -*-
"""
utils.py
--------
Pure helper functions with no Revit or UI dependencies.
Imported by every other module.
"""

import os
import re
import uuid


# ── Numeric / type helpers ────────────────────────────────────────────────────

def _safe_int(value, default):
    try:
        return int(value)
    except Exception:
        return default


def _as_bool(value):
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    try:
        return str(value).strip().lower() in ("1", "true", "yes", "y", "on")
    except Exception:
        return False


def _combo_selected_text(combo, default=""):
    """Return the Content string of the selected ComboBoxItem, or default."""
    try:
        item = combo.SelectedItem
        if item and hasattr(item, "Content"):
            return str(item.Content)
    except Exception:
        pass
    return default


# ── File-type detection ───────────────────────────────────────────────────────

def detect_file_type(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        return "excel"
    if ext in (".doc", ".docx"):
        return "word"
    if ext == ".pdf":
        return "pdf"
    return "image"


SUPPORTED_FILES_FILTER = (
    "Supported Files|*.xlsx;*.xls;*.xlsm;*.doc;*.docx;*.pdf;"
    "*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff|All Files|*.*"
)

DPI_CHOICES = ["72", "96", "150", "200", "300", "600"]


# ── Record description helpers ─────────────────────────────────────────────────

def describe_record_options(record):
    ftype       = record.get("file_type", "")
    dpi         = record.get("dpi", "300")
    transparent = _as_bool(record.get("transparent", False))

    if ftype == "excel":
        sheet = record.get("sheet_name", "")
        rng   = record.get("range_addr", "")
        bits  = []
        if sheet:
            bits.append(sheet)
        if rng:
            bits.append(rng)
        bits.append("{} DPI".format(dpi))
        if transparent:
            bits.append("Transparent")
        return " | ".join(bits) if bits else "—"

    if ftype in ("pdf", "word"):
        page_number = record.get("page_number", "1")
        bits = ["Page {}".format(page_number), "{} DPI".format(dpi)]
        if transparent:
            bits.append("Transparent")
        return " | ".join(bits)

    bits = ["{} DPI".format(dpi)]
    if transparent:
        bits.append("Transparent")
    return " | ".join(bits)


# ── Import name / user helpers ────────────────────────────────────────────────

def _sanitize_import_name(raw):
    """Convert a file basename to a safe, human-readable import name.

    Replaces characters that are illegal in Revit element names with spaces,
    collapses whitespace, and trims to 128 characters.
    """
    if not raw:
        return ""
    name = re.sub(r'[^\w\s\-]', ' ', raw, flags=re.UNICODE)
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:128]


def _get_auto_user():
    """Return a best-effort username string."""
    for env in ("USERNAME", "USER", "LOGNAME"):
        val = os.environ.get(env, "")
        if val:
            return val
    return "Unknown"


# ── Page number parsing ───────────────────────────────────────────────────────

def _parse_page_numbers(page_str):
    """Parse a page number string such as '1', '2-4', '1,3,5' etc.

    Returns a sorted list of 1-based integers.
    Falls back to [1] on any parse error.
    """
    result = set()
    try:
        for part in str(page_str).split(","):
            part = part.strip()
            if "-" in part:
                lo, hi = part.split("-", 1)
                for n in range(int(lo.strip()), int(hi.strip()) + 1):
                    result.add(n)
            else:
                result.add(int(part))
    except Exception:
        return [1]
    return sorted(result) or [1]


# ── Path resolution helpers ──────────────────────────────────────────────────

def resolve_path(raw_path, path_type, doc_path):
    """
    Resolve a potentially relative path to an absolute path.
    
    raw_path  : The path stored in the record.
    path_type : 'Absolute' or 'Relative'.
    doc_path  : Full path to the Revit model (doc.PathName).
    """
    if not raw_path:
        return ""
    if path_type == "Relative" and doc_path:
        try:
            return os.path.normpath(os.path.join(os.path.dirname(doc_path), raw_path))
        except Exception:
            pass
    return raw_path


def to_relative_path(abs_path, doc_path):
    """
    Try to convert an absolute path to one relative to the Revit model.
    Falls back to abs_path on error (e.g. different drives).
    """
    if not abs_path or not doc_path:
        return abs_path
    try:
        return os.path.relpath(abs_path, os.path.dirname(doc_path))
    except (ValueError, Exception):
        return abs_path


# ── UUID helper ───────────────────────────────────────────────────────────────

def _uid():
    import uuid
    return str(uuid.uuid4())
