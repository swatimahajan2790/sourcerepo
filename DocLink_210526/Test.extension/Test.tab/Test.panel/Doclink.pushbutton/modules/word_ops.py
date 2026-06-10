# -*- coding: utf-8 -*-
"""
word_ops.py
-----------
Microsoft Word COM automation for DocLink.

Public API
----------
export_word_to_pdf(word_path)             -> (pdf_path, total_pages)
get_word_page_count(word_path)            -> int
capture_word_page_as_image(word_path, ..) -> str|None  (via XPS → WPF → PNG)
_render_xps_page_to_png(xps_path, dpi)   -> str|None
"""

import os
import _imports

# Late binding for Word - get from _imports at runtime, not import time
_WORD_AVAILABLE = _imports._WORD_AVAILABLE
Marshal = _imports.Marshal
clr = _imports.clr

def _get_word():
    """Get Word object from _imports, handling None case."""
    return getattr(_imports, 'Word', None)


# ─────────────────────────────────────────────────────────────────────────────
# Safe TEMP directory helper — works in C# and PyRevit contexts
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# Revit 2025 Workaround: PowerShell/VBScript conversion (when COM unavailable)
# ─────────────────────────────────────────────────────────────────────────────

def _convert_word_to_pdf_powershell(word_path, output_pdf):
    """
    Convert Word to PDF using PowerShell COM (works if Word installed).
    Bypasses IronPython COM limitations in Revit 2025.
    """
    try:
        # Escape paths properly for PowerShell
        word_escaped = word_path.replace('"', '""')
        output_escaped = output_pdf.replace('"', '""')
        
        ps_script = '''
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open("{0}", $false, $true)
$doc.ExportAsFixedFormat("{1}", 17)
$doc.Close($false)
$word.Quit()
'''.format(word_escaped, output_escaped)
        
        import subprocess
        proc = subprocess.Popen(
            ["powershell", "-NoProfile", "-Command", ps_script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        try:
            stdout, stderr = proc.communicate(timeout=120)
        except TypeError:
            # IronPython or older Python: timeout not supported
            stdout, stderr = proc.communicate()
        
        if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
            return True
    except Exception as ex:
        print("[DocLinkManager] PowerShell Word conversion failed: {}".format(ex))
    return False

def _get_word_page_count_powershell(word_path):
    """Get Word page count via PowerShell (fallback for Revit 2025)."""
    try:
        word_escaped = word_path.replace('"', '""')
        ps_script = '''
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
    $doc = $word.Documents.Open("{0}", $false, $true)
    $count = $doc.ComputeStatistics(2)
    $doc.Close($false)
    Write-Host $count
}} finally {{
    $word.Quit()
}}
'''.format(word_escaped)
        
        import subprocess
        proc = subprocess.Popen(
            ["powershell", "-NoProfile", "-Command", ps_script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        try:
            stdout, stderr = proc.communicate(timeout=30)
        except TypeError:
            # IronPython or older Python: timeout not supported
            stdout, stderr = proc.communicate()
        
        if stdout and stdout.strip().isdigit():
            return int(stdout.strip())
    except Exception:
        pass
    return 0

def _convert_word_to_pdf_vbscript(word_path, output_pdf):
    """
    Convert Word to PDF using VBScript (works if Word installed).
    Fallback method for PowerShell.
    """
    try:
        vbs_script = '''
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objDoc = objWord.Documents.Open("{0}")
objDoc.ExportAsFixedFormat "{1}", 17
objDoc.Close
objWord.Quit
Set objDoc = Nothing
Set objWord = Nothing
'''.format(word_path, output_pdf)
        
        vbs_path = os.path.join(_get_safe_temp_dir(), "doclink_convert_word.vbs")
        with open(vbs_path, "w") as f:
            f.write(vbs_script)
        
        import subprocess
        proc = subprocess.Popen(
            ["cscript.exe", vbs_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        try:
            stdout, stderr = proc.communicate(timeout=120)
        except TypeError:
            # IronPython or older Python: timeout not supported
            stdout, stderr = proc.communicate()
        
        try:
            os.remove(vbs_path)
        except Exception:
            pass
        
        if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
            return True
    except Exception as ex:
        print("[DocLinkManager] VBScript Word conversion failed: {}".format(ex))
    return False


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


def export_word_to_pdf(word_path):
    """
    Export a Word document to PDF.
    Returns (output_pdf_path, total_pages).
    Revit 2025: Falls back to PowerShell/VBScript if COM unavailable.
    """
    base = os.path.splitext(os.path.basename(word_path))[0]
    temp_dir = _get_safe_temp_dir()
    output_pdf = os.path.join(temp_dir, "DocLinkManager_Word_{}.pdf".format(base))
    if os.path.exists(output_pdf):
        try:
            os.remove(output_pdf)
        except Exception:
            pass

    # Try 1: Use COM if available (Revit 2024 and earlier)
    if _WORD_AVAILABLE:
        word_app = None
        doc = None
        try:
            word_app = _get_word().ApplicationClass()
            word_app.Visible = False
            doc = word_app.Documents.Open(word_path, False, True)

            try:
                total_pages = int(doc.ComputeStatistics(2))  # wdStatisticPages = 2
            except Exception:
                total_pages = 1

            doc.ExportAsFixedFormat(output_pdf, 17)  # wdExportFormatPDF = 17
            doc.Close(False)
            doc = None
            
            if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
                return output_pdf, max(1, total_pages)

        except Exception as ex:
            pass  # Fall through to subprocess methods
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if doc is not None:
                    Marshal.ReleaseComObject(doc)
            except Exception:
                pass
            try:
                if word_app is not None:
                    try:
                        word_app.Quit()
                    except Exception:
                        pass
                    Marshal.ReleaseComObject(word_app)
            except Exception:
                pass

    # Try 2: PowerShell (Word must be installed)
    if _convert_word_to_pdf_powershell(word_path, output_pdf):
        # If we got here, we might need the real page count
        pc = _get_word_page_count_powershell(word_path)
        return output_pdf, max(1, pc)

    # Try 3: VBScript (Word must be installed)
    if _convert_word_to_pdf_vbscript(word_path, output_pdf):
        return output_pdf, 1

    # All methods failed
    raise RuntimeError(
        "Word → PDF conversion failed.\n"
        "Requirements:\n"
        "  - Revit 2024 & earlier: Works with bundled runtime\n"
        "  - Revit 2025+: Requires Microsoft Word installed on system\n"
        "Error: Could not export to PDF via COM, PowerShell, or VBScript."
    )


def get_word_page_count(word_path):
    """Return total number of pages in a Word document (without exporting)."""
    if not _WORD_AVAILABLE:
        return 0
    word_app = doc = None
    try:
        word_app = _get_word().ApplicationClass()
        word_app.Visible = False
        doc = word_app.Documents.Open(word_path, False, True)
        count = int(doc.ComputeStatistics(2))
        doc.Close(False)
        doc = None
        return max(1, count)
    except Exception:
        # Try PowerShell fallback for Revit 2025
        return _get_word_page_count_powershell(word_path)
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if doc is not None:
                Marshal.ReleaseComObject(doc)
        except Exception:
            pass
        try:
            if word_app is not None:
                try:
                    word_app.Quit()
                except Exception:
                    pass
                Marshal.ReleaseComObject(word_app)
        except Exception:
            pass


def _render_xps_page_to_png(xps_path, dpi=300):
    """
    Render the first page of an XPS document to PNG using WPF.
    Returns temp PNG path, or None on failure.
    """
    xps_doc = None
    try:
        clr.AddReference("ReachFramework")
        from System.Windows.Xps.Packaging import XpsDocument as XpsDoc
        from System.IO import FileAccess as FA
        from System.Windows.Media.Imaging import (
            RenderTargetBitmap, PngBitmapEncoder, BitmapFrame
        )
        from System.Windows.Media import PixelFormats, VisualBrush, DrawingVisual
        from System.Windows import Rect

        xps_doc = XpsDoc(xps_path, FA.Read)
        seq = xps_doc.GetFixedDocumentSequence()
        if seq is None:
            print("[DocLinkManager] XPS: no FixedDocumentSequence")
            return None

        paginator = seq.DocumentPaginator
        if paginator.PageCount < 1:
            print("[DocLinkManager] XPS: page count is 0")
            return None

        page = paginator.GetPage(0)
        visual = page.Visual
        scale = float(dpi) / 96.0
        pw = int(page.Size.Width  * scale)
        ph = int(page.Size.Height * scale)

        dv = DrawingVisual()
        dc = dv.RenderOpen()
        vb = VisualBrush(visual)
        dc.DrawRectangle(vb, None, Rect(0, 0, pw, ph))
        dc.Close()

        rtb = RenderTargetBitmap(pw, ph, 96, 96, PixelFormats.Pbgra32)
        rtb.Render(dv)

        encoder = PngBitmapEncoder()
        encoder.Frames.Add(BitmapFrame.Create(rtb))

        temp_dir = _get_safe_temp_dir()
        output_png = os.path.join(
            temp_dir,
            "DocLinkManager_XPSRender_{}.png".format(
                os.path.splitext(os.path.basename(xps_path))[0])
        )
        from System.IO import FileStream, FileMode
        fs = FileStream(output_png, FileMode.Create)
        try:
            encoder.Save(fs)
        finally:
            fs.Close()

        print("[DocLinkManager] XPS rendered to PNG: {}".format(output_png))
        return output_png

    except Exception as ex:
        print("[DocLinkManager] XPS rendering failed: {}".format(ex))
        return None
    finally:
        try:
            if xps_doc is not None:
                xps_doc.Close()
        except Exception:
            pass


def capture_word_page_as_image(word_path, page_number=1, dpi=300):
    """
    Render a specific Word page to PNG via: Word COM → XPS → WPF → PNG.
    Returns temp PNG path, or None on failure.
    """
    if not _WORD_AVAILABLE:
        print("[DocLinkManager] capture_word_page_as_image: Word COM not available")
        return None

    base = os.path.splitext(os.path.basename(word_path))[0]
    temp_dir = _get_safe_temp_dir()
    xps_path = os.path.join(
        temp_dir,
        "DocLinkManager_WordXPS_{}_p{}.xps".format(base, page_number)
    )
    if os.path.exists(xps_path):
        try:
            os.remove(xps_path)
        except Exception:
            pass

    word_app = doc = None
    try:
        word_app = _get_word().ApplicationClass()
        word_app.Visible = False
        doc = word_app.Documents.Open(word_path, False, True)

        try:
            total_pages = int(doc.ComputeStatistics(2))
        except Exception:
            total_pages = 1

        pn = max(1, min(page_number, total_pages))
        doc.ExportAsFixedFormat(xps_path, 18)  # wdExportFormatXPS = 18
        doc.Close(False)
        doc = None

        if not os.path.exists(xps_path):
            print("[DocLinkManager] capture_word_page_as_image: XPS not created")
            return None

        return _render_xps_page_to_png(xps_path, dpi=dpi)

    except Exception as ex:
        print("[DocLinkManager] capture_word_page_as_image failed: {}".format(ex))
        return None
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if doc is not None:
                Marshal.ReleaseComObject(doc)
        except Exception:
            pass
        try:
            if word_app is not None:
                try:
                    word_app.Quit()
                except Exception:
                    pass
                Marshal.ReleaseComObject(word_app)
        except Exception:
            pass
