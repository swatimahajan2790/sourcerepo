# -*- coding: utf-8 -*-
"""
pdf_ops.py
----------
PDF + image processing operations for DocLink.

Relies on the _pdf_helper.py CPython subprocess bridge for PyMuPDF
(fitz) since PyMuPDF is a CPython C-extension only.

Public API
----------
convert_pdf_page_to_png(pdf_path, page_number, dpi)  -> str|None
get_pdf_page_count(pdf_path)                         -> int
make_white_background_transparent(src_path, tolerance) -> str
crop_white_margins(src_path, padding_px, whiteness_threshold) -> str
remove_background(src_path, page_number, dpi, tolerance)     -> str|None
"""

import os

import System
from _imports import (
    Bitmap, Graphics, PixelFormat, ImageFormat, ImageLockMode, Marshal,
    _CPYTHON, _FITZ_AVAILABLE, _PDF_HELPER, _LIB_DIR,
    _subprocess,
)
from word_ops import _render_xps_page_to_png
from logger import LogManager


# ─────────────────────────────────────────────────────────────────────────────
# Safe TEMP directory helper — works in C# and PyRevit contexts
# ─────────────────────────────────────────────────────────────────────────────

def _get_safe_temp_dir():
    """
    Get a writable temp directory with fallbacks.
    
    C# compiled tools may not have TEMP set, or it might be inaccessible.
    This function tries multiple sources with fallbacks.
    """
    # Try 1: TEMP environment variable
    temp_dir = os.environ.get("TEMP")
    if temp_dir and os.path.isdir(temp_dir):
        try:
            # Test write access
            test_file = os.path.join(temp_dir, ".doclink_write_test")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
            return temp_dir
        except Exception:
            pass
    
    # Try 2: TMP environment variable
    tmp_dir = os.environ.get("TMP")
    if tmp_dir and os.path.isdir(tmp_dir):
        try:
            test_file = os.path.join(tmp_dir, ".doclink_write_test")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
            return tmp_dir
        except Exception:
            pass
    
    # Try 3: USERPROFILE\AppData\Local\Temp (Windows default)
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        win_temp = os.path.join(userprofile, "AppData", "Local", "Temp")
        if os.path.isdir(win_temp):
            try:
                test_file = os.path.join(win_temp, ".doclink_write_test")
                with open(test_file, "w") as f:
                    f.write("test")
                os.remove(test_file)
                return win_temp
            except Exception:
                pass
    
    # Try 4: User home directory
    home = os.path.expanduser("~")
    if home and home != "~":
        home_temp = os.path.join(home, "AppData", "Local", "Temp")
        if os.path.isdir(home_temp):
            try:
                test_file = os.path.join(home_temp, ".doclink_write_test")
                with open(test_file, "w") as f:
                    f.write("test")
                os.remove(test_file)
                return home_temp
            except Exception:
                pass
    
    # Try 5: Create a DocLink-specific temp in current directory (desperate fallback)
    doclink_temp = os.path.abspath(os.path.join(".", "DocLink_Temp"))
    try:
        if not os.path.exists(doclink_temp):
            os.makedirs(doclink_temp)
        test_file = os.path.join(doclink_temp, ".doclink_write_test")
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        return doclink_temp
    except Exception:
        pass
    
    # Absolute last resort — just return TEMP and let it fail with a clear error
    return os.environ.get("TEMP", "C:\\Windows\\Temp")
# ─────────────────────────────────────────────────────────────────────────────
# PDF via PyMuPDF CPython bridge
# ─────────────────────────────────────────────────────────────────────────────

def convert_pdf_page_to_png(pdf_path, page_number=1, dpi=300):
    """Render a single PDF page to PNG via PyMuPDF (CPython subprocess).

    page_number is 1-based.
    Returns temp PNG path, or None on failure.
    """
    LogManager.info("▶ convert_pdf_page_to_png START: file='{}', page={}, dpi={}".format(
        pdf_path, page_number, dpi))

    if not _FITZ_AVAILABLE:
        LogManager.error("✗ PyMuPDF (fitz) not available - CPython environment missing")
        LogManager.warning("  Expected at: DocLink.pushbutton/runtime/python.exe")
        return None

    # Check if source file exists
    if not os.path.isfile(pdf_path):
        LogManager.error("✗ PDF file not found: {}".format(pdf_path))
        return None

    LogManager.debug("  File size: {} bytes".format(os.path.getsize(pdf_path)))

    base = os.path.splitext(os.path.basename(pdf_path))[0]
    temp_dir = _get_safe_temp_dir()
    output_png = os.path.join(
        temp_dir,
        "DocLinkManager_PDFPage_{}_{}.png".format(base, page_number)
    )
    LogManager.debug("  Output PNG path: {}".format(output_png))

    try:
        LogManager.debug("  Launching PDF helper subprocess via Python subprocess...")
        LogManager.debug("    CPython: {}".format(_CPYTHON))
        LogManager.debug("    Helper: {}".format(_PDF_HELPER))

        creationflags = 0x08000000 if os.name == 'nt' else 0
        proc = _subprocess.Popen(
            [_CPYTHON, _PDF_HELPER, 'convert', _LIB_DIR, pdf_path, output_png,
             str(page_number), str(dpi)],
            stdout=_subprocess.PIPE,
            stderr=_subprocess.PIPE,
            creationflags=creationflags,
            shell=False,
            cwd=temp_dir,
        )
        out, err = proc.communicate()
        out_text = out.decode('utf-8', 'replace').strip() if out else ''
        err_text = err.decode('utf-8', 'replace').strip() if err else ''

        LogManager.debug("  Subprocess completed with exit code: {}".format(proc.returncode))

        if proc.returncode != 0:
            LogManager.error("  ✗ PDF helper process error (code {}):".format(proc.returncode))
            if err_text:
                LogManager.error("    STDERR: {}".format(err_text))
            if out_text:
                LogManager.debug("    STDOUT: {}".format(out_text))
            return None

        if os.path.isfile(output_png):
            sz = os.path.getsize(output_png)
            LogManager.info("✓ PDF→PNG conversion successful: {} bytes".format(sz))
            LogManager.debug("  Output PNG: {}".format(output_png))
            return output_png
        if out_text and os.path.isfile(out_text):
            sz = os.path.getsize(out_text)
            LogManager.info("✓ PDF→PNG conversion successful: {} bytes".format(sz))
            LogManager.debug("  Output PNG from helper stdout: {}".format(out_text))
            return out_text

        LogManager.error("✗ Output PNG file was not created at: {}".format(output_png))
        return None

    except Exception as ex:
        LogManager.error("✗ PDF helper exception: {}".format(ex))
        LogManager.exception("  Traceback details:")
        return None


def get_pdf_page_count(pdf_path):
    """Return total number of pages in a PDF, or 0 on failure."""
    if not _FITZ_AVAILABLE:
        return 0
    try:
        creationflags = 0x08000000 if os.name == 'nt' else 0
        proc = _subprocess.Popen(
            [_CPYTHON, _PDF_HELPER, 'pagecount', _LIB_DIR, pdf_path],
            stdout=_subprocess.PIPE,
            stderr=_subprocess.PIPE,
            creationflags=creationflags,
            shell=False,
            cwd=_get_safe_temp_dir(),
        )
        out, err = proc.communicate()
        if proc.returncode == 0:
            out_text = out.decode('utf-8', 'replace').strip() if out else ''
            return int(out_text)
        return 0
    except Exception:
        return 0


# ─────────────────────────────────────────────────────────────────────────────
# Image processing — LockBits-based (performance-critical)
# ─────────────────────────────────────────────────────────────────────────────

def make_white_background_transparent(src_path, tolerance=252):
    """
    Convert only near-WHITE pixels to transparent and save as a temp PNG.

    Uses LockBits + Marshal.Copy for ~10× speedup vs per-pixel GetPixel/SetPixel.
    BGRA byte order: [B, G, R, A] per pixel.

    tolerance   : pixels where min(R,G,B) >= this are fully erased (default 252)
    LOWER_BOUND : pixels below this are always fully opaque (246)
    """
    LogManager.info("▶ make_white_background_transparent START: file='{}'".format(src_path))
    LogManager.debug("  Tolerance: {}".format(tolerance))

    ext = os.path.splitext(src_path)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"):
        LogManager.warning("  Unsupported file type '{}'; returning original".format(ext))
        return src_path

    try:
        if not os.path.exists(src_path):
            LogManager.error("  ✗ Source file does not exist: {}".format(src_path))
            return src_path

        file_size = os.path.getsize(src_path)
        if file_size < 64:
            LogManager.warning("  ✗ File too small ({} bytes); returning original".format(file_size))
            return src_path

        LogManager.debug("  Source file: {} bytes".format(file_size))
    except Exception as ex:
        LogManager.warning("  Could not check file: {}".format(ex))
        pass

    LOWER_BOUND = 246
    src_bmp = bmp = bmp_data = g = None

    try:
        LogManager.debug("  Loading image with System.Drawing.Bitmap...")
        src_bmp = Bitmap(src_path)
        w = src_bmp.Width
        h = src_bmp.Height
        LogManager.info("  Image dimensions: {}×{} pixels".format(w, h))

        LogManager.debug("  Creating 32-bit ARGB working bitmap...")
        bmp = Bitmap(w, h, PixelFormat.Format32bppArgb)
        g = Graphics.FromImage(bmp)
        g.DrawImage(src_bmp, 0, 0, w, h)
        g.Dispose()
        g = None

        LogManager.debug("  Locking bitmap bits for pixel processing...")
        rect     = System.Drawing.Rectangle(0, 0, w, h)
        bmp_data = bmp.LockBits(rect, ImageLockMode.ReadWrite,
                                 PixelFormat.Format32bppArgb)
        stride    = bmp_data.Stride
        num_bytes = stride * h
        pixels    = System.Array.CreateInstance(System.Byte, num_bytes)
        LogManager.debug("  Copying {} bytes from bitmap memory...".format(num_bytes))
        Marshal.Copy(bmp_data.Scan0, pixels, 0, num_bytes)

        LogManager.debug("  Processing pixels for transparency...")
        transparent_count = 0
        total_pixels      = w * h

        for py in range(h):
            base = py * stride
            for px in range(w):
                idx  = base + px * 4
                b_ch = int(pixels[idx])
                g_ch = int(pixels[idx + 1])
                r_ch = int(pixels[idx + 2])

                whiteness = min(r_ch, g_ch, b_ch)
                if whiteness >= tolerance:
                    pixels[idx + 3] = 0
                    transparent_count += 1
                elif whiteness > LOWER_BOUND:
                    alpha = int(255 * (1.0 - float(whiteness - LOWER_BOUND)
                                           / float(tolerance - LOWER_BOUND)))
                    pixels[idx + 3] = max(0, min(255, alpha))

        LogManager.debug("  Copying pixel data back to bitmap...")
        Marshal.Copy(pixels, 0, bmp_data.Scan0, num_bytes)
        bmp.UnlockBits(bmp_data)
        bmp_data = None

        pct = 100.0 * transparent_count / max(1, total_pixels)
        if transparent_count == 0:
            LogManager.warning("  ⚠ WARNING: 0/{} pixels made transparent (tolerance={})".format(
                total_pixels, tolerance))
            LogManager.warning("    Consider using a lower tolerance value")
        else:
            LogManager.info("  ✓ {}/{} pixels ({:.1f}%) made transparent".format(
                transparent_count, total_pixels, pct))

        output_png = os.path.join(
            os.environ["TEMP"],
            "DocLinkManager_Transparent_{}.png".format(
                os.path.splitext(os.path.basename(src_path))[0])
        )
        LogManager.debug("  Saving to: {}".format(output_png))
        bmp.Save(output_png, ImageFormat.Png)

        try:
            out_size = os.path.getsize(output_png)
            LogManager.info("✓ Transparency conversion complete: {} bytes".format(out_size))
            if out_size < 64:
                LogManager.warning("  ⚠ Output PNG too small ({} bytes) — fallback to original".format(out_size))
                return src_path
        except Exception:
            pass

        return output_png

    except Exception as ex:
        LogManager.error("✗ make_white_background_transparent EXCEPTION: {}".format(ex))
        LogManager.exception("  Traceback:")
        return src_path

    finally:
        try:
            if bmp_data is not None and bmp is not None:
                bmp.UnlockBits(bmp_data)
        except Exception:
            pass
        try:
            if g is not None:
                g.Dispose()
        except Exception:
            pass
        try:
            if bmp is not None:
                bmp.Dispose()
        except Exception:
            pass
        try:
            if src_bmp is not None:
                src_bmp.Dispose()
        except Exception:
            pass
            pass
        try:
            if g is not None:
                g.Dispose()
        except Exception:
            pass
        try:
            if bmp is not None:
                bmp.Dispose()
        except Exception:
            pass
        try:
            if src_bmp is not None:
                src_bmp.Dispose()
        except Exception:
            pass


def crop_white_margins(src_path, padding_px=2, whiteness_threshold=250):
    """
    Trim outer blank/white rows and columns from a raster image using LockBits.

    Returns path to cropped PNG, or src_path if the operation is skipped.
    """
    LogManager.info("▶ crop_white_margins START: file='{}'".format(src_path))
    LogManager.debug("  Padding: {} px, Whiteness threshold: {}".format(padding_px, whiteness_threshold))

    ext = os.path.splitext(src_path)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"):
        LogManager.debug("  Unsupported file type '{}'; skipping".format(ext))
        return src_path

    try:
        if not os.path.exists(src_path) or os.path.getsize(src_path) < 64:
            LogManager.debug("  File missing or too small; skipping")
            return src_path
    except Exception:
        return src_path

    src_bmp = bmp = bmp_data = g = None

    try:
        LogManager.debug("  Loading image...")
        src_bmp = Bitmap(src_path)
        w = src_bmp.Width
        h = src_bmp.Height
        LogManager.debug("  Dimensions: {}×{}".format(w, h))

        if w < 4 or h < 4:
            LogManager.debug("  Image too small; skipping")
            return src_path

        bmp = Bitmap(w, h, PixelFormat.Format32bppArgb)
        g = Graphics.FromImage(bmp)
        g.DrawImage(src_bmp, 0, 0, w, h)
        g.Dispose()
        g = None

        LogManager.debug("  Analyzing pixel data...")
        rect     = System.Drawing.Rectangle(0, 0, w, h)
        bmp_data = bmp.LockBits(rect, ImageLockMode.ReadOnly,
                                 PixelFormat.Format32bppArgb)
        stride    = bmp_data.Stride
        num_bytes = stride * h
        pixels    = System.Array.CreateInstance(System.Byte, num_bytes)
        Marshal.Copy(bmp_data.Scan0, pixels, 0, num_bytes)
        bmp.UnlockBits(bmp_data)
        bmp_data = None

        def row_is_blank(py):
            base = py * stride
            for px in range(w):
                idx = base + px * 4
                b_ch = int(pixels[idx])
                g_ch = int(pixels[idx + 1])
                r_ch = int(pixels[idx + 2])
                a_ch = int(pixels[idx + 3])
                if a_ch > 16 and min(r_ch, g_ch, b_ch) < whiteness_threshold:
                    return False
            return True

        def col_is_blank(px):
            for py in range(h):
                idx = py * stride + px * 4
                b_ch = int(pixels[idx])
                g_ch = int(pixels[idx + 1])
                r_ch = int(pixels[idx + 2])
                a_ch = int(pixels[idx + 3])
                if a_ch > 16 and min(r_ch, g_ch, b_ch) < whiteness_threshold:
                    return False
            return True

        LogManager.debug("  Detecting margins...")
        top = 0
        while top < h and row_is_blank(top):
            top += 1
        bottom = h - 1
        while bottom > top and row_is_blank(bottom):
            bottom -= 1
        left = 0
        while left < w and col_is_blank(left):
            left += 1
        right = w - 1
        while right > left and col_is_blank(right):
            right -= 1

        LogManager.debug("  Margins detected: top={}, bottom={}, left={}, right={}".format(
            top, bottom, left, right))

        top    = max(0,     top    - padding_px)
        bottom = min(h - 1, bottom + padding_px)
        left   = max(0,     left   - padding_px)
        right  = min(w - 1, right  + padding_px)

        crop_w = right  - left  + 1
        crop_h = bottom - top   + 1

        if crop_w < 4 or crop_h < 4:
            LogManager.debug("  Result too small ({}×{}); skipping".format(crop_w, crop_h))
            return src_path

        trimmed_pct = 100.0 * (1.0 - float(crop_w * crop_h) / float(w * h))
        if trimmed_pct < 0.5:
            LogManager.debug("  Trim amount negligible ({:.2f}%); skipping".format(trimmed_pct))
            return src_path

        LogManager.info("✓ Margins detected: {}×{} → {}×{} ({:.1f}% trimmed)".format(
            w, h, crop_w, crop_h, trimmed_pct))

        LogManager.debug("  Creating cropped bitmap...")
        cropped = Bitmap(crop_w, crop_h, PixelFormat.Format32bppArgb)
        g = Graphics.FromImage(cropped)
        src_rect  = System.Drawing.Rectangle(left, top, crop_w, crop_h)
        dest_rect = System.Drawing.Rectangle(0, 0, crop_w, crop_h)
        g.DrawImage(bmp, dest_rect, src_rect, System.Drawing.GraphicsUnit.Pixel)
        g.Dispose()
        g = None

        output_png = os.path.join(
            os.environ["TEMP"],
            "DocLinkManager_Cropped_{}.png".format(
                os.path.splitext(os.path.basename(src_path))[0])
        )
        LogManager.debug("  Saving cropped image to: {}".format(output_png))
        cropped.Save(output_png, ImageFormat.Png)
        cropped.Dispose()
        LogManager.info("✓ Crop complete: {}".format(output_png))
        return output_png

    except Exception as ex:
        LogManager.error("✗ crop_white_margins EXCEPTION: {}".format(ex))
        LogManager.exception("  Traceback:")
        return src_path

    finally:
        try:
            if bmp_data is not None and bmp is not None:
                bmp.UnlockBits(bmp_data)
        except Exception:
            pass
        try:
            if g is not None:
                g.Dispose()
        except Exception:
            pass
        try:
            if bmp is not None:
                bmp.Dispose()
        except Exception:
            pass
        try:
            if src_bmp is not None:
                src_bmp.Dispose()
        except Exception:
            pass


def remove_background(src_path, page_number=1, dpi=300, tolerance=252):
    """
    Unified background-removal entry point.

    • Image files (.png/.jpg/.bmp/.tif) → make_white_background_transparent
    • PDF files (.pdf) → convert_pdf_page_to_png → make_white_background_transparent
    • XPS files → _render_xps_page_to_png → make_white_background_transparent
    • Unknown formats → returned as-is

    Returns the path to the transparent PNG, or None if a renderer is unavailable.
    """
    LogManager.section("BACKGROUND REMOVAL PIPELINE")
    LogManager.info("Source file: {}".format(src_path))
    LogManager.info("Page number: {}, DPI: {}, Tolerance: {}".format(page_number, dpi, tolerance))

    ext = os.path.splitext(src_path)[1].lower()
    LogManager.debug("File extension: '{}'".format(ext))

    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"):
        LogManager.info("→ Path: Image file → make_white_background_transparent")
        return make_white_background_transparent(src_path, tolerance=tolerance)

    if ext == ".xps":
        LogManager.info("→ Path: XPS file → _render_xps_page_to_png → make_white_background_transparent")
        png_path = _render_xps_page_to_png(src_path, dpi=dpi)
        if png_path:
            LogManager.debug("XPS render succeeded: {}".format(png_path))
            return make_white_background_transparent(png_path, tolerance=tolerance)
        LogManager.error("XPS render failed, returning original")
        return src_path

    if ext == ".pdf":
        LogManager.info("→ Path: PDF file → convert_pdf_page_to_png → make_white_background_transparent")
        png_path = convert_pdf_page_to_png(src_path, page_number=page_number, dpi=dpi)
        if png_path:
            LogManager.debug("PDF→PNG conversion succeeded: {}".format(png_path))
            return make_white_background_transparent(png_path, tolerance=tolerance)
        LogManager.error("PyMuPDF PDF render failed, cannot remove background")
        return None

    LogManager.warning("Unknown file type '{}', returning original file".format(ext))
    return src_path
