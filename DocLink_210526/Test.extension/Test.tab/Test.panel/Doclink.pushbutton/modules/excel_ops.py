# -*- coding: utf-8 -*-
"""
excel_ops.py
------------
All Microsoft Excel operations for DocLink — COM when available,
PowerShell/VBScript fallback for Revit 2025+ (no in-process Office),
and direct OpenXML parsing for read-only metadata (defined names,
sheet list, print area) so basic discovery works without Office at all.

Auto-format conversion: Any Excel file format (.xls, .xlsb, .xlsm, etc.)
is automatically converted to .xlsx and cached in temp before processing.
This ensures consistent behavior across all file formats.

Public API
----------
export_excel_range_to_pdf(excel_path, range_address, sheet_name, named_range=None) -> str
capture_excel_range_as_image(excel_path, range_address, sheet_name, dpi) -> str|None
get_excel_print_area(excel_path, sheet_name) -> str|None
list_excel_defined_names(excel_path, sheet_name=None) -> list[dict]
resolve_defined_name(excel_path, name, sheet_name=None) -> dict|None
_read_excel_display_values(excel_path, sheet_name, bounds) -> dict
"""

import os
import hashlib
import _imports

# Late binding for Excel - get from _imports at runtime, not import time
_EXCEL_AVAILABLE = _imports._EXCEL_AVAILABLE
Marshal = _imports.Marshal
Bitmap = _imports.Bitmap
Graphics = _imports.Graphics
PixelFormat = _imports.PixelFormat
ImageFormat = _imports.ImageFormat
WinForms = _imports.WinForms
Drawing2D = getattr(_imports, 'Drawing2D', None)

def _get_excel():
    """Get Excel object from _imports, handling None case."""
    return getattr(_imports, 'Excel', None)


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
# Excel Format Normalization — Convert any format (.xls, .xlsb, etc.) to .xlsx
# ─────────────────────────────────────────────────────────────────────────────

def _ensure_xlsx_format(excel_path):
    """
    Convert any Excel format to .xlsx if needed.
    
    If the file is already .xlsx, returns the original path.
    If the file is .xls, .xlsb, .xlsm, or other format, converts to .xlsx,
    saves in temp directory, and returns the temp path.
    
    Returns the path to an .xlsx file (original or converted).
    """
    if not excel_path or not os.path.exists(excel_path):
        return excel_path
    
    # Check if already .xlsx
    if excel_path.lower().endswith('.xlsx'):
        return excel_path
    
    # Need conversion
    temp_dir = _get_safe_temp_dir()
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    source_key = os.path.abspath(excel_path).lower().encode('utf-8')
    source_hash = hashlib.md5(source_key).hexdigest()[:10]
    output_xlsx = os.path.join(
        temp_dir,
        "DocLink_Normalized_{}_{}.xlsx".format(base_name, source_hash)
    )

    # Reuse an up-to-date normalized copy to avoid repeated SaveAs prompts.
    if os.path.exists(output_xlsx):
        try:
            if os.path.getmtime(output_xlsx) >= os.path.getmtime(excel_path):
                return output_xlsx
            os.remove(output_xlsx)
        except Exception:
            pass
    
    # Try COM first (Revit 2024 and earlier)
    if _EXCEL_AVAILABLE:
        xl_app = None
        wb = None
        try:
            xl_app = _get_excel().ApplicationClass()
            xl_app.Visible = False
            xl_app.DisplayAlerts = False
            try:
                xl_app.AlertBeforeOverwriting = False
            except Exception:
                pass
            try:
                xl_app.AskToUpdateLinks = False
            except Exception:
                pass
            try:
                xl_app.AutomationSecurity = 3
            except Exception:
                pass
            
            wb = xl_app.Workbooks.Open(excel_path, 0, True)
            try:
                wb.CheckCompatibility = False
            except Exception:
                pass
            # SaveAs with format 51 = xlOpenXMLWorkbook (.xlsx)
            wb.SaveAs(output_xlsx, 51)
            try:
                wb.Saved = True
            except Exception:
                pass
            wb.Close(False)
            wb = None
            
            if os.path.exists(output_xlsx) and os.path.getsize(output_xlsx) > 0:
                return output_xlsx
        except Exception:
            pass
        finally:
            try:
                if wb:
                    wb.Close(False)
            except Exception:
                pass
            try:
                if xl_app:
                    xl_app.Quit()
                    Marshal.ReleaseComObject(xl_app)
            except Exception:
                pass
    
    # Fallback: PowerShell conversion
    try:
        excel_escaped = excel_path.replace('"', '""')
        output_escaped = output_xlsx.replace('"', '""')
        
        ps_script = '''
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
    try {{ $excel.AlertBeforeOverwriting = $false }} catch {{ }}
    try {{ $excel.AskToUpdateLinks = $false }} catch {{ }}
    try {{ $excel.AutomationSecurity = 3 }} catch {{ }}
    $wb = $excel.Workbooks.Open("{0}", 0, $true)
    try {{ $wb.CheckCompatibility = $false }} catch {{ }}
$wb.SaveAs("{1}", 51)
    try {{ $wb.Saved = $true }} catch {{ }}
$wb.Close($false)
$excel.Quit()
'''.format(excel_escaped, output_escaped)
        
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
        
        if os.path.exists(output_xlsx) and os.path.getsize(output_xlsx) > 0:
            return output_xlsx
    except Exception as ex:
        print("[DocLinkManager] PowerShell format conversion failed: {}".format(ex))
    
    # Last resort: return original and let caller handle the error
    print("[DocLinkManager] Format conversion to .xlsx failed, using original file")
    return excel_path


# ─────────────────────────────────────────────────────────────────────────────
# Defined Names (Name Manager) — OpenXML primary, COM fallback for .xls
# ─────────────────────────────────────────────────────────────────────────────
# Mirrors C# DocLink.Core/Ingestion/Excel/NpoiExcelReader.GetNamedRanges():
#   - Workbook-scoped names returned regardless of active sheet
#   - Sheet-scoped names returned only when their sheet matches sheet_name
#   - Built-ins (_xlnm.*) excluded
#   - External-workbook refs (RefersTo contains '[') excluded
#   - Non-range constants (RefersTo has no '!') excluded

def _strip_quoted_sheet(s):
    """Strip surrounding single quotes that Excel adds when a sheet name has spaces."""
    if len(s) >= 2 and s[0] == "'" and s[-1] == "'":
        return s[1:-1].replace("''", "'")
    return s


def _parse_refers_to(refers_to):
    """
    Parse a `RefersTo` formula like `Sheet1!$A$1:$G$25` or `'My Sheet'!$A$1`.
    Returns (sheet_name, a1_range) or (None, None) if it doesn't match.
    Strips the leading '=' Excel sometimes includes.
    """
    if not refers_to:
        return None, None
    s = refers_to.strip()
    if s.startswith("="):
        s = s[1:].strip()
    if "!" not in s:
        return None, None
    if "[" in s:  # external workbook reference
        return None, None
    sheet_part, _, range_part = s.rpartition("!")
    sheet_part = _strip_quoted_sheet(sheet_part.strip())
    range_part = range_part.strip().replace("$", "")
    if not sheet_part or not range_part:
        return None, None
    return sheet_part, range_part


def _list_defined_names_openxml(excel_path):
    """Parse defined names directly from .xlsx/.xlsm OpenXML. No Office needed."""
    import zipfile
    import xml.etree.ElementTree as ET
    if not zipfile.is_zipfile(excel_path):
        return None  # not openxml — caller should try COM
    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    def tag(local): return "{%s}%s" % (NS, local)
    results = []
    try:
        with zipfile.ZipFile(excel_path, "r") as zf:
            wb_path = None
            for n in zf.namelist():
                if n.lower().endswith("workbook.xml"):
                    wb_path = n
                    break
            if wb_path is None:
                return []
            root = ET.fromstring(zf.read(wb_path))
            sheet_names = [sh.get("name", "") for sh in root.iter(tag("sheet"))]
            for dn in root.iter(tag("definedName")):
                name = dn.get("name", "") or ""
                if not name or name.startswith("_xlnm."):
                    continue
                refers_to = (dn.text or "").strip()
                sheet, rng = _parse_refers_to(refers_to)
                if not sheet or not rng:
                    continue
                local_id_attr = dn.get("localSheetId")
                if local_id_attr is None:
                    scope_sheet = None
                    is_workbook = True
                else:
                    try:
                        idx = int(local_id_attr)
                        scope_sheet = sheet_names[idx] if 0 <= idx < len(sheet_names) else None
                    except (ValueError, TypeError):
                        scope_sheet = None
                    is_workbook = False
                results.append({
                    "name":              name,
                    "refers_to":         refers_to,
                    "ref_sheet":         sheet,
                    "range_address":     rng,
                    "scope_sheet":       scope_sheet,
                    "is_workbook_scope": is_workbook,
                })
    except Exception as ex:
        print("[DocLinkManager] _list_defined_names_openxml failed: {}".format(ex))
        return []
    return results


def _list_defined_names_com(excel_path):
    """COM fallback for legacy .xls files (no OpenXML parser path)."""
    if not _EXCEL_AVAILABLE:
        return []
    xl_app = wb = None
    results = []
    try:
        xl_app = _get_excel().ApplicationClass()
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        wb = xl_app.Workbooks.Open(excel_path)
        names_col = wb.Names
        count = 0
        try:
            count = int(names_col.Count)
        except Exception:
            pass
        for i in range(1, count + 1):
            try:
                nm = names_col.Item(i)
                name = str(nm.Name) if nm.Name else ""
                if not name or name.startswith("_xlnm."):
                    continue
                # Sheet-scoped names come back as "Sheet!Name"; normalise.
                bare = name.rsplit("!", 1)[-1]
                refers_to = str(nm.RefersTo) if nm.RefersTo else ""
                sheet, rng = _parse_refers_to(refers_to)
                if not sheet or not rng:
                    continue
                scope_sheet = None
                is_workbook = True
                try:
                    # Worksheet-scoped names expose a Parent that is the worksheet.
                    parent = nm.Parent
                    parent_type = getattr(parent, "Name", None)
                    if parent_type and parent_type != wb.Name:
                        scope_sheet = str(parent_type)
                        is_workbook = False
                except Exception:
                    pass
                results.append({
                    "name":              bare,
                    "refers_to":         refers_to,
                    "ref_sheet":         sheet,
                    "range_address":     rng,
                    "scope_sheet":       scope_sheet,
                    "is_workbook_scope": is_workbook,
                })
            except Exception:
                continue
    except Exception as ex:
        print("[DocLinkManager] _list_defined_names_com failed: {}".format(ex))
    finally:
        try:
            if wb:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl_app:
                xl_app.Quit()
                Marshal.ReleaseComObject(xl_app)
        except Exception:
            pass
    return results


def list_excel_defined_names(excel_path, sheet_name=None):
    """
    Return user-defined names visible from `sheet_name`.

    When `sheet_name` is provided, the result includes workbook-scoped names
    plus any names scoped to that sheet (matching C# NpoiExcelReader scope
    logic). When omitted, every workbook- and sheet-scoped name is returned.

    Each entry is a dict with: name, refers_to, ref_sheet, range_address,
    scope_sheet (None = workbook), is_workbook_scope.
    """
    if not excel_path or not os.path.exists(excel_path):
        return []
    
    # Normalize to .xlsx format
    excel_path = _ensure_xlsx_format(excel_path)
    
    items = _list_defined_names_openxml(excel_path)
    if items is None:
        items = _list_defined_names_com(excel_path)
    if not items:
        return []
    if sheet_name:
        items = [n for n in items if n["is_workbook_scope"] or n["scope_sheet"] == sheet_name]
    items.sort(key=lambda n: (not n["is_workbook_scope"], n["name"].lower()))
    return items


def resolve_defined_name(excel_path, name, sheet_name=None):
    """Return the single defined-name dict matching `name`, or None.

    Prefers a sheet-scoped match when sheet_name is supplied; otherwise
    falls back to workbook scope.
    """
    if not name:
        return None
    candidates = list_excel_defined_names(excel_path, sheet_name)
    if sheet_name:
        for n in candidates:
            if n["name"] == name and n["scope_sheet"] == sheet_name:
                return n
    for n in candidates:
        if n["name"] == name and n["is_workbook_scope"]:
            return n
    for n in candidates:
        if n["name"] == name:
            return n
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Revit 2025 Workaround: PowerShell/VBScript conversion (when COM unavailable)
# ─────────────────────────────────────────────────────────────────────────────

def _convert_excel_to_pdf_powershell(excel_path, output_pdf, sheet_name=None,
                                     range_address=None, named_range=None):
    """
    Convert Excel to PDF using PowerShell COM (works if Excel installed).
    Bypasses IronPython COM limitations in Revit 2025.

    When `named_range` is supplied, the PS script resolves the name in-process
    via `$wb.Names.Item(...).RefersToRange` and uses its absolute address as
    the Print_Area, ignoring `range_address`.
    """
    try:
        # Escape paths properly for PowerShell
        excel_escaped = excel_path.replace('"', '""')
        output_escaped = output_pdf.replace('"', '""')

        sheet_logic = ""
        if sheet_name:
            sheet_logic += '$ws = $wb.Sheets.Item("{0}")\n$ws.Activate()\n'.format(sheet_name.replace('"', '""'))
        else:
            sheet_logic += '$ws = $wb.ActiveSheet\n'

        if named_range:
            # Resolve the named range inside Excel so workbook-scoped and
            # sheet-scoped names both work. Switch the active worksheet to the
            # one the name resolves to before setting PrintArea.
            nm_escaped = named_range.replace('"', '""')
            sheet_logic += (
                '$nm = $wb.Names.Item("{0}")\n'
                '$rng = $nm.RefersToRange\n'
                '$ws = $rng.Worksheet\n'
                '$ws.Activate()\n'
                '$ws.PageSetup.PrintArea = $rng.Address($true, $true)\n'
            ).format(nm_escaped)
        elif range_address:
            sheet_logic += '$ws.PageSetup.PrintArea = "{0}"\n'.format(range_address)

        sheet_logic += '''
$ps = $ws.PageSetup
$ps.LeftHeader = ""
$ps.CenterHeader = ""
$ps.RightHeader = ""
$ps.LeftFooter = ""
$ps.CenterFooter = ""
$ps.RightFooter = ""
try { $ps.LeftMargin = 0 } catch { }
try { $ps.RightMargin = 0 } catch { }
try { $ps.TopMargin = 0 } catch { }
try { $ps.BottomMargin = 0 } catch { }
try { $ps.HeaderMargin = 0 } catch { }
try { $ps.FooterMargin = 0 } catch { }
try { $ps.PrintGridlines = $false } catch { }
try {
    $ps.Zoom = $false
    $ps.FitToPagesWide = 1
    $ps.FitToPagesTall = 1
} catch { }
'''

        # Export from the WORKSHEET, not the workbook — workbook export
        # ignores per-sheet PrintArea and emits every visible sheet (and on
        # .xlsm macro workbooks effectively returns sheet 1 first). Sheet
        # export honours both PrintArea and the active sheet selection.
        # AutomationSecurity=3 suppresses macro-execution prompts when
        # opening .xlsm files headlessly.
        ps_script = '''
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{ $excel.AutomationSecurity = 3 }} catch {{ }}
$wb = $excel.Workbooks.Open("{0}")
{1}
$ws.ExportAsFixedFormat(0, "{2}")
$wb.Close($false)
$excel.Quit()
'''.format(excel_escaped, sheet_logic, output_escaped)
        
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
        print("[DocLinkManager] PowerShell conversion failed: {}".format(ex))
    return False

def _convert_excel_to_pdf_vbscript(excel_path, output_pdf, sheet_name=None,
                                   range_address=None, named_range=None):
    """
    Convert Excel to PDF using VBScript COM (alternative fallback).

    `named_range` (if supplied) is resolved inside Excel via Names.Item(...)
    .RefersToRange and overrides `range_address`.
    """
    try:
        sheet_logic = ""
        if sheet_name:
            sheet_logic += 'Set ws = wb.Sheets("{0}")\nws.Activate\n'.format(sheet_name.replace('"', '""'))
        else:
            sheet_logic += 'Set ws = wb.ActiveSheet\n'

        if named_range:
            nm_escaped = named_range.replace('"', '""')
            sheet_logic += (
                'Set nm = wb.Names.Item("{0}")\n'
                'Set rng = nm.RefersToRange\n'
                'Set ws = rng.Worksheet\n'
                'ws.Activate\n'
                'ws.PageSetup.PrintArea = rng.Address(True, True)\n'
            ).format(nm_escaped)
        elif range_address:
            sheet_logic += 'ws.PageSetup.PrintArea = "{0}"\n'.format(range_address)

        sheet_logic += (
            'Set ps = ws.PageSetup\n'
            'ps.LeftHeader = ""\n'
            'ps.CenterHeader = ""\n'
            'ps.RightHeader = ""\n'
            'ps.LeftFooter = ""\n'
            'ps.CenterFooter = ""\n'
            'ps.RightFooter = ""\n'
            'On Error Resume Next\n'
            'ps.LeftMargin = 0\n'
            'ps.RightMargin = 0\n'
            'ps.TopMargin = 0\n'
            'ps.BottomMargin = 0\n'
            'ps.HeaderMargin = 0\n'
            'ps.FooterMargin = 0\n'
            'ps.PrintGridlines = False\n'
            'ps.Zoom = False\n'
            'ps.FitToPagesWide = 1\n'
            'ps.FitToPagesTall = 1\n'
            'On Error Goto 0\n'
        )

        # Export from the WORKSHEET (`ws`), not the workbook — see PS variant
        # for the same rationale. AutomationSecurity=3 (msoAutomationSecurityForceDisable)
        # blocks macro prompts on .xlsm files.
        vbs_content = '''
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
On Error Resume Next
excel.AutomationSecurity = 3
On Error Goto 0
Set wb = excel.Workbooks.Open("{0}")
{1}
ws.ExportAsFixedFormat 0, "{2}"
wb.Close False
excel.Quit
Set ws = Nothing
Set wb = Nothing
Set excel = Nothing
'''.format(excel_path, sheet_logic, output_pdf)

        vbs_path = os.path.join(_get_safe_temp_dir(), "doclink_convert_excel.vbs")
        with open(vbs_path, "w") as f:
            f.write(vbs_content)
        
        import subprocess
        proc = subprocess.Popen(
            ["cscript.exe", vbs_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        stdout, stderr = proc.communicate(timeout=120)
        
        try:
            os.remove(vbs_path)
        except Exception:
            pass
        
        if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
            return True
    except Exception:
        pass
    return False


def export_excel_range_to_pdf(excel_path, range_address, sheet_name, named_range=None):
    """
    Set the print area to exactly the chosen range, export to PDF, then restore.
    Uses Names("Print_Area") to avoid IronPython COM type mismatch on PageSetup.PrintArea.
    Returns the path to the temp PDF.

    When `named_range` is supplied, the workbook's existing defined name is
    resolved (workbook- or sheet-scoped) and its range overrides
    `range_address` and `sheet_name`. This mirrors the C# Name-Manager flow.

    Revit 2025: Falls back to PowerShell/VBScript if COM unavailable.
    """
    # Normalize to .xlsx format
    excel_path = _ensure_xlsx_format(excel_path)
    
    temp_dir = _get_safe_temp_dir()
    output_pdf = os.path.join(temp_dir, "DocLinkManager_Temp.pdf")
    if os.path.exists(output_pdf):
        try:
            os.remove(output_pdf)
        except Exception:
            pass

    # Try 1: Use COM if available (Revit 2024 and earlier)
    xl_app = None
    try:
        if _EXCEL_AVAILABLE:
            xl_app = _get_excel().ApplicationClass()
            xl_app.Visible = False
            xl_app.DisplayAlerts = False

            wb = xl_app.Workbooks.Open(excel_path)

            if named_range:
                # Resolve via the workbook's Names collection so both
                # workbook-scoped and sheet-scoped names work. The Range's
                # parent worksheet becomes the export target.
                nm = wb.Names.Item(named_range)
                rng = nm.RefersToRange
                ws = rng.Worksheet
            else:
                ws = wb.Sheets[sheet_name] if sheet_name else wb.ActiveSheet
                rng = ws.Range[range_address]
            abs_address = str(rng.Address[True, True])

            original_ref = None
            try:
                pa_name = ws.Names.Item("Print_Area")
                original_ref = str(pa_name.RefersTo)
            except Exception:
                pass

            ws.Names.Add("Print_Area", "={}!{}".format(ws.Name, abs_address))

            ps = ws.PageSetup
            saved_hf = (
                ps.LeftHeader,   ps.CenterHeader,   ps.RightHeader,
                ps.LeftFooter,   ps.CenterFooter,   ps.RightFooter,
            )
            saved_margins         = None
            saved_zoom            = None
            saved_print_gridlines = None
            
            try:
                saved_margins = (
                    ps.LeftMargin, ps.RightMargin,
                    ps.TopMargin, ps.BottomMargin,
                    ps.HeaderMargin, ps.FooterMargin,
                )
            except Exception:
                pass
            
            try:
                saved_zoom = (ps.Zoom, ps.FitToPagesWide, ps.FitToPagesTall)
            except Exception:
                pass
            
            try:
                saved_print_gridlines = ps.PrintGridlines
                ps.PrintGridlines = False
            except Exception:
                pass

            ps.LeftHeader = ps.CenterHeader = ps.RightHeader = ""
            ps.LeftFooter = ps.CenterFooter = ps.RightFooter = ""
            
            try:
                ps.LeftMargin   = 0
                ps.RightMargin  = 0
                ps.TopMargin    = 0
                ps.BottomMargin = 0
                ps.HeaderMargin = 0
                ps.FooterMargin = 0
            except Exception:
                pass
            
            try:
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = 1
            except Exception:
                pass

            try:
                ws.ExportAsFixedFormat(0, output_pdf, 0)
            finally:
                (ps.LeftHeader, ps.CenterHeader, ps.RightHeader,
                 ps.LeftFooter, ps.CenterFooter, ps.RightFooter) = saved_hf
                if saved_margins:
                    try:
                        (ps.LeftMargin, ps.RightMargin,
                         ps.TopMargin, ps.BottomMargin,
                         ps.HeaderMargin, ps.FooterMargin) = saved_margins
                    except Exception:
                        pass
                if saved_zoom:
                    try:
                        ps.FitToPagesWide = saved_zoom[1]
                        ps.FitToPagesTall = saved_zoom[2]
                        ps.Zoom = saved_zoom[0]
                    except Exception:
                        pass
                if saved_print_gridlines is not None:
                    try:
                        ps.PrintGridlines = saved_print_gridlines
                    except Exception:
                        pass

            try:
                if original_ref:
                    ws.Names.Add("Print_Area", original_ref)
                else:
                    ws.Names.Item("Print_Area").Delete()
            except Exception:
                pass

            wb.Close(False)
            
            if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
                return output_pdf
    except Exception:
        pass  # Fall through to subprocess methods
    finally:
        try:
            if xl_app:
                xl_app.Quit()
                Marshal.ReleaseComObject(xl_app)
        except Exception:
            pass

    # Try 2: PowerShell (Excel must be installed)
    if _convert_excel_to_pdf_powershell(excel_path, output_pdf, sheet_name,
                                        range_address, named_range=named_range):
        return output_pdf

    # Try 3: VBScript (Excel must be installed)
    if _convert_excel_to_pdf_vbscript(excel_path, output_pdf, sheet_name,
                                      range_address, named_range=named_range):
        return output_pdf

    # All three Office-dependent paths failed. The cause is almost always
    # one of: (a) Microsoft Excel not installed; (b) PowerShell running in
    # ConstrainedLanguage mode (typically enforced by WDAC); (c) AppLocker
    # blocking COM activation. There is no in-process pyRevit alternative
    # that produces Office-fidelity Excel→PDF output — surface the cause
    # plainly so the user knows what to ask their IT for.
    raise RuntimeError(
        "Excel → PDF conversion failed on this machine.\n\n"
        "DocLink tried every available Office-COM path:\n"
        "  1. In-process Excel COM (Revit 2024 & earlier only)\n"
        "  2. PowerShell + Excel COM (Revit 2025+ fallback)\n"
        "  3. VBScript + Excel COM (final fallback)\n\n"
        "All three failed. The cause is one of:\n"
        "  • Microsoft Excel is not installed on this PC.\n"
        "  • PowerShell is in ConstrainedLanguage mode (WDAC policy).\n"
        "  • AppLocker / SmartScreen is blocking COM activation.\n\n"
        "Run  DocLink.pushbutton\\diagnose.py  to identify the specific cause."
    )


def get_excel_sheets_com(excel_path):
    """
    List sheets in an Excel workbook using COM (fallback for macro-enabled files).
    
    Returns list of dict:
        {
            'name': sheet_name,
            'has_print_area': bool,
            'print_area': str or None
        }
    
    Returns empty list if COM is unavailable or file cannot be opened.
    """
    if not _EXCEL_AVAILABLE:
        return []
    
    xl_app = wb = None
    try:
        xl_app = _get_excel().ApplicationClass()
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        try:
            xl_app.AlertBeforeOverwriting = False
        except Exception:
            pass
        
        # Normalize to .xlsx first if it's a different format
        try:
            normalized_path = _ensure_xlsx_format(excel_path)
        except Exception:
            # If normalization fails, try the original path
            normalized_path = excel_path
        
        wb = xl_app.Workbooks.Open(normalized_path)
        results = []
        
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            sheet_name = str(ws.Name)
            
            # Try to detect print area
            has_pa = False
            pa_addr = None
            try:
                pa_name = ws.Names.Item("Print_Area")
                refers_to = str(pa_name.RefersTo)
                if "!" in refers_to:
                    pa_addr = refers_to.split("!")[-1].replace("$", "")
                has_pa = True
            except Exception:
                pass
            
            results.append({
                'name': sheet_name,
                'has_print_area': has_pa,
                'print_area': pa_addr
            })
        
        return results
    except Exception as ex:
        # Silently fail and return empty list
        return []
    finally:
        try:
            if wb:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl_app:
                xl_app.Quit()
                Marshal.ReleaseComObject(xl_app)
        except Exception:
            pass


def capture_excel_range_as_image(excel_path, range_address, sheet_name, dpi=96):
    """
    Capture a range of Excel cells as an image (PNG), then return the file path.
    Disables gridlines during capture, then restores them.
    """
    if not _EXCEL_AVAILABLE:
        return None
    
    # Normalize to .xlsx format
    excel_path = _ensure_xlsx_format(excel_path)

    XL_SCREEN = 1
    XL_BITMAP  = 2
    xl_app = None
    try:
        xl_app = _get_excel().ApplicationClass()
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        wb = xl_app.Workbooks.Open(excel_path)
        ws = wb.Sheets[sheet_name] if sheet_name else wb.ActiveSheet
        try:
            ws.Activate()
        except Exception:
            pass
        rng = ws.Range[range_address]
        # Ensure we don't capture millions of empty cells if user provides A:Z
        rng = _excel_range_content_extents(ws, rng)

        orig_display_gridlines = True
        orig_window_gridlines  = True
        try:
            # Ensure the sheet is visible and active for CopyPicture
            ws.Visible = -1 # xlSheetVisible
            ws.Activate()
            # Set zoom to 100% to ensure CopyPicture captures at standard resolution
            xl_app.ActiveWindow.Zoom = 100
            orig_display_gridlines = ws.DisplayGridlines
        except Exception:
            pass
        try:
            orig_window_gridlines = xl_app.ActiveWindow.DisplayGridlines
        except Exception:
            pass
        try:
            ws.DisplayGridlines = False
        except Exception:
            pass
        try:
            xl_app.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass

        # xlScreen=1, xlBitmap=2
        # Use xlScreen to capture exactly what is shown (macros, buttons, etc.)
        rng.CopyPicture(1, 2) 

        # Restore gridlines
        try:
            ws.DisplayGridlines = orig_display_gridlines
        except Exception:
            pass
        try:
            xl_app.ActiveWindow.DisplayGridlines = orig_window_gridlines
        except Exception:
            pass

        # Give Excel a moment to process the clipboard
        import time
        time.sleep(0.2)
        
        img = WinForms.Clipboard.GetImage()
        if img is None:
            # Fallback: some Excel versions fail on first CopyPicture if clipboard is busy
            import time
            time.sleep(0.1)
            rng.CopyPicture(1, 2)
            img = WinForms.Clipboard.GetImage()
            
        if img is None:
            print("[DocLinkManager] capture_excel_range_as_image: clipboard returned None")
            return None

        scale = max(1.0, float(dpi) / 96.0)
        new_w = int(img.Width  * scale)
        new_h = int(img.Height * scale)
        scaled = Bitmap(new_w, new_h, PixelFormat.Format32bppArgb)
        g = Graphics.FromImage(scaled)
        try:
            if Drawing2D is not None:
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            g.DrawImage(img, 0, 0, new_w, new_h)
        finally:
            g.Dispose()

        temp_dir = _get_safe_temp_dir()
        output_png = os.path.join(
            temp_dir,
            "DocLinkManager_ExcelCapture_{}.png".format(
                os.path.splitext(os.path.basename(excel_path))[0])
        )
        scaled.Save(output_png, ImageFormat.Png)
        wb.Close(False)
        print("[DocLinkManager] capture_excel_range_as_image: saved {}".format(output_png))
        return output_png

    except Exception as ex:
        print("[DocLinkManager] capture_excel_range_as_image failed: {}".format(ex))
        return None
    finally:
        try:
            if xl_app:
                xl_app.Quit()
                Marshal.ReleaseComObject(xl_app)
        except Exception:
            pass


def _excel_range_content_extents(ws, rng):
    """
    Limit a range ONLY if it is exceptionally large (e.g. entire columns A:Z)
    to avoid capturing millions of empty cells, while respecting smaller
    user-defined ranges even if they contain empty cells.
    """
    try:
        # If the range is "reasonable" (e.g. less than 1000 rows/columns),
        # respect the user's exact selection.
        if rng.Rows.Count < 2000 and rng.Columns.Count < 200:
            return rng

        used = ws.UsedRange
        if not used:
            return rng
            
        xl_app = ws.Application
        intersection = xl_app.Intersect(rng, used)
        if intersection:
            return intersection
    except Exception:
        pass
    return rng


def _excel_col_index_to_letters(idx):
    idx = int(idx or 0)
    if idx < 1:
        return "A"
    letters = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters.append(chr(65 + rem))
    return ''.join(reversed(letters))


def _bounds_to_excel_range(bounds):
    if not bounds:
        return None
    min_col, min_row, max_col, max_row = bounds
    return "{}{}:{}{}".format(
        _excel_col_index_to_letters(min_col), min_row,
        _excel_col_index_to_letters(max_col), max_row,
    )


def _read_excel_display_values(excel_path, sheet_name, bounds):
    """Read exact displayed Excel text (Cell.Text) for cells in bounds range."""
    if (not _EXCEL_AVAILABLE) or (not bounds):
        return {}
    min_col, min_row, max_col, max_row = bounds
    xl_app = wb = None
    values = {}
    try:
        xl_app = _get_excel().ApplicationClass()
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        wb = xl_app.Workbooks.Open(excel_path)
        ws = wb.Sheets[sheet_name] if sheet_name else wb.ActiveSheet
        for abs_row in range(min_row, max_row + 1):
            for abs_col in range(min_col, max_col + 1):
                try:
                    txt = ws.Cells(abs_row, abs_col).Text
                    if txt is None:
                        txt = ''
                    txt = str(txt)
                    values[(abs_row, abs_col)] = txt.replace('\r\n', '\n').replace('\r', '\n')
                except Exception:
                    pass
        return values
    finally:
        try:
            if wb:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl_app:
                xl_app.Quit()
                Marshal.ReleaseComObject(xl_app)
        except Exception:
            pass


def get_excel_print_area(excel_path, sheet_name=None):
    """Return the Print_Area range address (e.g. 'A1:G25') or None."""
    if not _EXCEL_AVAILABLE:
        return None
    
    # Normalize to .xlsx format
    excel_path = _ensure_xlsx_format(excel_path)
    
    xl_app = wb = None
    try:
        xl_app = _get_excel().ApplicationClass()
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        wb = xl_app.Workbooks.Open(excel_path)
        ws = wb.Sheets[sheet_name] if sheet_name else wb.ActiveSheet
        try:
            pa_name = ws.Names.Item("Print_Area")
            refers_to = str(pa_name.RefersTo)
            if "!" in refers_to:
                range_part = refers_to.split("!")[-1]
                return range_part.replace("$", "")
        except Exception:
            pass
        try:
            used_range = ws.UsedRange
            if used_range:
                return str(used_range.Address).replace("$", "")
        except Exception:
            pass
        return None
    except Exception:
        return None
    finally:
        try:
            if wb:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl_app:
                xl_app.Quit()
                Marshal.ReleaseComObject(xl_app)
        except Exception:
            pass
