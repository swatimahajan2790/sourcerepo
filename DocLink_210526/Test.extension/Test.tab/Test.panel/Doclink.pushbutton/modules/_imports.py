# -*- coding: utf-8 -*-
"""
_imports.py
-----------
All CLR / WPF / WinForms / Revit API imports and global constants
shared across every DocLink module.

Every other module starts with:
    from _imports import *
"""

import os
import sys
import clr
import re
import json
import datetime
import tempfile
import zipfile
import subprocess as _subprocess
import xml.etree.ElementTree as ET


# ── Shared extension-dir helper ───────────────────────────────────────────────
def _get_ext_dir():
    """Walk upward to find the .extension root or lib/shared folders."""
    curr = os.path.dirname(os.path.abspath(__file__))
    # __file__ is inside modules/, so go up one level first
    curr = os.path.dirname(curr)
    for _ in range(5):
        if (os.path.basename(curr).endswith('.extension')
                or os.path.exists(os.path.join(curr, "lib"))
                or os.path.exists(os.path.join(curr, "shared"))):
            return curr
        parent = os.path.dirname(curr)
        if parent == curr:
            break
        curr = parent
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")


_EXT_DIR = _get_ext_dir()

# Add shared + lib to sys.path so revit_ui_loader etc. can be found
for _p in (os.path.join(_EXT_DIR, "shared"), os.path.join(_EXT_DIR, "lib")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

try:
    from revit_ui_loader import apply_template_to_window, TEMPLATE_PATH as _TMPL_PATH, _load_xaml
    _RES = _load_xaml(_TMPL_PATH)
except Exception:
    def apply_template_to_window(*a, **kw):
        pass
    _TMPL_PATH = None
    _RES = None

# ── pyRevit ───────────────────────────────────────────────────────────────────
from pyrevit import forms, revit

# ── System / WPF ─────────────────────────────────────────────────────────────
import System
from System import Array, Guid, String
from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    Visibility, GridLength, GridUnitType, FontWeights
)
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, StackPanel, WrapPanel, ScrollViewer,
    Button, Label, TextBox, TextBlock, ComboBox, ComboBoxItem, CheckBox,
    RadioButton,
    Separator, DataGrid, DataGridTextColumn, DataGridSelectionMode,
    Orientation as WPFOrientation, SelectionChangedEventArgs,
    TabControl, TabItem, ProgressBar
)
from System.Windows.Data import Binding
from System.Windows.Media import SolidColorBrush, Color, Brushes
from System.Collections.ObjectModel import ObservableCollection
clr.AddReference("System")
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows import SystemParameters
from System.Diagnostics import Process, ProcessStartInfo
clr.AddReference("WindowsBase")
from System.Windows.Threading import Dispatcher, DispatcherPriority, DispatcherFrame
from System.Runtime.InteropServices import Marshal

# ── System.Drawing (WinForms and image processing) ────────────────────────────
clr.AddReference("System.Drawing")
from System.Drawing import (
    Bitmap, Graphics, Color as DrawingColor,
    Size as DSize, Point as DPoint, Font as DFont,
)
DColor = DrawingColor   # alias used in ExcelHeaderMapper dialogs
from System.Drawing.Imaging import ImageFormat, PixelFormat, ImageLockMode
import System.Windows.Forms as WinForms

# ── WinForms aliases (avoid name clash with WPF controls) ─────────────────────
WFForm              = WinForms.Form
WFLabel             = WinForms.Label
WFTextBox           = WinForms.TextBox
WFButton            = WinForms.Button
WFCheckBox          = WinForms.CheckBox
WFListBox           = WinForms.ListBox
WFOpenFileDialog    = WinForms.OpenFileDialog
WFDialogResult      = WinForms.DialogResult
WFFormBorderStyle   = WinForms.FormBorderStyle
WFFormStartPosition = WinForms.FormStartPosition
WFSelectionMode     = WinForms.SelectionMode
WFBorderStyle       = WinForms.BorderStyle
WFNumericUpDown     = WinForms.NumericUpDown
WFGroupBox          = WinForms.GroupBox
WFComboBox          = WinForms.ComboBox
WFComboBoxStyle     = WinForms.ComboBoxStyle
WFDialogResultEnum  = WinForms.DialogResult
WFRadioButton       = WinForms.RadioButton

# ── Excel COM ─────────────────────────────────────────────────────────────────
Excel = None
_EXCEL_AVAILABLE = False
try:
    clr.AddReference("Microsoft.Office.Interop.Excel")
    from Microsoft.Office.Interop import Excel
    _EXCEL_AVAILABLE = True
except Exception as _e:
    # Revit 2025+ (.NET 8) fallback via Late Binding
    try:
        import System
        _excel_type = System.Type.GetTypeFromProgID("Excel.Application")
        if _excel_type is not None:
            class _DummyExcel(object):
                @staticmethod
                def ApplicationClass():
                    return System.Activator.CreateInstance(_excel_type)
            Excel = _DummyExcel()
            _EXCEL_AVAILABLE = True
    except Exception:
        Excel = None
        _EXCEL_AVAILABLE = False

# ── Word COM ──────────────────────────────────────────────────────────────────
Word = None
_WORD_AVAILABLE = False
try:
    clr.AddReference("Microsoft.Office.Interop.Word")
    from Microsoft.Office.Interop import Word
    _WORD_AVAILABLE = True
except Exception as _e:
    # Revit 2025+ (.NET 8) fallback via Late Binding
    try:
        import System
        _word_type = System.Type.GetTypeFromProgID("Word.Application")
        if _word_type is not None:
            class _DummyWord(object):
                @staticmethod
                def ApplicationClass():
                    return System.Activator.CreateInstance(_word_type)
            Word = _DummyWord()
            _WORD_AVAILABLE = True
    except Exception:
        Word = None
        _WORD_AVAILABLE = False

# ── Revit API ─────────────────────────────────────────────────────────────────
from Autodesk.Revit.DB import (
    Transaction, TransactionGroup,
    ImageType, ImageTypeOptions, ImageTypeSource,
    ImagePlacementOptions, ImageInstance, XYZ, BoxPlacement,
    BuiltInParameter, ElementId, FilteredElementCollector,
    # ExcelHeaderMapper additions:
    SectionType, TableMergedCell, TableCellStyle,
    TableCellStyleOverrideOptions,
    Color as RevitColor,
    ViewSchedule, GraphicsStyle, ScheduleFieldType,
    ScheduleFilter, ScheduleFilterType,
    BuiltInCategory, CellType, ViewSheet,
    ScheduleSheetInstance,
)
from Autodesk.Revit.DB.ExtensibleStorage import (
    Schema, SchemaBuilder, Entity, AccessLevel
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons

try:
    from Autodesk.Revit.DB import HorizontalAlignmentStyle
    _HAS_H = True
except ImportError:
    _HAS_H = False

try:
    from Autodesk.Revit.DB import VerticalAlignmentStyle
    _HAS_V = True
except ImportError:
    _HAS_V = False

try:
    from Autodesk.Revit.DB import TextNoteType
    _HAS_TNT = True
except Exception:
    TextNoteType = None
    _HAS_TNT = False

# ── DataStorage type (lazy, reflection-based) ─────────────────────────────────
def _get_data_storage_type():
    for _asm in System.AppDomain.CurrentDomain.GetAssemblies():
        if _asm.GetName().Name == "RevitAPI":
            try:
                ds_type = _asm.GetType("Autodesk.Revit.DB.DataStorage")
                if ds_type:
                    return ds_type
                for t in _asm.GetTypes():
                    if t.Name == "DataStorage":
                        return t
            except Exception:
                pass
    return None

_DataStorageType = None

def DataStorageType():
    global _DataStorageType
    if _DataStorageType is None:
        _DataStorageType = _get_data_storage_type()
        if _DataStorageType is None:
            raise ImportError(
                "Could not resolve Autodesk.Revit.DB.DataStorage from RevitAPI assembly.")
    return _DataStorageType


# ── PyMuPDF / CPython bridge ──────────────────────────────────────────────────
_MODULES_DIR = os.path.dirname(os.path.abspath(__file__))
_PDF_HELPER_SRC = os.path.join(_MODULES_DIR, "_pdf_helper.py")
_LIB_DIR_SRC    = os.path.join(_EXT_DIR, "lib")
_PDF_HELPER     = _PDF_HELPER_SRC
_LIB_DIR        = _LIB_DIR_SRC


def _runtime_site_packages_dir(runtime_root):
    """Return bundled runtime site-packages path when present."""
    if not runtime_root:
        return ""
    candidates = [
        os.path.join(runtime_root, "Lib", "site-packages"),
        os.path.join(runtime_root, "Lib", "site-packages.zip"),
    ]
    for candidate in candidates:
        if os.path.isdir(candidate) or os.path.isfile(candidate):
            return candidate
    return ""


def _default_pdf_lib_dir():
    """Prefer external lib/, else bundled runtime site-packages, else empty."""
    if os.path.isdir(_LIB_DIR_SRC):
        return _LIB_DIR_SRC
    _pushbutton_root = os.path.dirname(_MODULES_DIR)
    runtime_root = os.path.join(_pushbutton_root, "runtime")
    return _runtime_site_packages_dir(runtime_root)


_LIB_DIR = _default_pdf_lib_dir()


def _doclink_local_dir():
    """Per-user local cache root for DocLink (created if missing)."""
    base = (os.environ.get("LOCALAPPDATA")
            or os.path.expandvars("%LOCALAPPDATA%")
            or os.path.join(os.path.expanduser("~"), "AppData", "Local"))
    out = os.path.join(base, "DocLink")
    try:
        if not os.path.isdir(out):
            os.makedirs(out)
    except Exception:
        pass
    return out


def _bridge_cache_dir():
    """Per-user local cache for the full PDF bridge bundle."""
    out = os.path.join(_doclink_local_dir(), "pdf_bridge")
    try:
        if not os.path.isdir(out):
            os.makedirs(out)
    except Exception:
        pass
    return out


def _is_network_path(path):
    """Best-effort UNC/network path detection."""
    if not path:
        return False
    try:
        norm = os.path.abspath(path)
    except Exception:
        norm = path
    return norm.startswith("\\\\") or norm.startswith("//")


def _path_signature(path):
    """Small file/dir signature used for cache invalidation."""
    if not path or not os.path.exists(path):
        return ""
    try:
        st = os.stat(path)
        return "{0}|{1}|{2}".format(path, int(st.st_mtime), st.st_size)
    except Exception:
        return path


# Cache the resolved CPython path so we don't re-probe the bundled runtime
# on every script load (= every button click). Invalidated when the source
# runtime's mtime changes (= you deployed a new version).
_CPYTHON_CACHE_FILE = os.path.join(_doclink_local_dir(), "runtime_cache.json")


def _runtime_source_token():
    """Stable token for the bundled PDF bridge so the cache invalidates on deploy."""
    _pushbutton_root = os.path.dirname(_MODULES_DIR)
    _runtime_src     = os.path.join(_pushbutton_root, "runtime")
    _bundled_python  = os.path.join(_runtime_src, "python.exe")
    if not os.path.isfile(_bundled_python):
        return ""
    parts = [
        _path_signature(_bundled_python),
        _path_signature(_PDF_HELPER_SRC),
        _path_signature(_LIB_DIR_SRC),
    ]
    return "||".join([p for p in parts if p])


def _load_cpython_cache():
    """Return cached cpython path if still valid, else None."""
    try:
        if not os.path.isfile(_CPYTHON_CACHE_FILE):
            return None
        import json as _json
        with open(_CPYTHON_CACHE_FILE, "r") as f:
            data = _json.load(f)
        exe   = data.get("cpython")
        token = data.get("source_token")
        if not exe or not os.path.isfile(exe):
            return None
        # Invalidate when the bundled runtime has been updated upstream.
        if token != _runtime_source_token():
            return None
        return exe
    except Exception:
        return None


def _save_cpython_cache(exe):
    """Persist the resolved cpython path keyed by the bundled runtime's token."""
    try:
        import json as _json
        with open(_CPYTHON_CACHE_FILE, "w") as f:
            _json.dump({
                "cpython":      exe,
                "source_token": _runtime_source_token(),
            }, f)
    except Exception:
        pass


def _strip_motw(root_dir):
    """
    Remove the Zone.Identifier alternate data stream from every file under
    `root_dir`. Files copied off a network share or unzipped from a download
    carry MOTW, which Windows/SmartScreen uses to block `python.exe` on some
    locked-down configurations. Stripping it makes the local cache execute
    cleanly. Safe no-op on non-NTFS / files without the ADS.
    """
    try:
        for dirpath, _dirnames, filenames in os.walk(root_dir):
            for fn in filenames:
                ads = os.path.join(dirpath, fn) + ":Zone.Identifier"
                try:
                    if os.path.exists(ads):
                        os.remove(ads)
                except Exception:
                    # Some files may be locked or non-NTFS — keep going.
                    pass
    except Exception:
        pass


def _stage_pdf_bridge_locally():
    """Copy runtime + PDF helper + lib to local AppData for reliable execution."""
    import shutil

    _pushbutton_root = os.path.dirname(_MODULES_DIR)
    runtime_src      = os.path.join(_pushbutton_root, "runtime")
    helper_src       = _PDF_HELPER_SRC
    lib_src          = _LIB_DIR_SRC

    if not os.path.isdir(runtime_src) or not os.path.isfile(os.path.join(runtime_src, "python.exe")):
        raise IOError("Bundled runtime not found: {}".format(runtime_src))
    if not os.path.isfile(helper_src):
        raise IOError("PDF helper not found: {}".format(helper_src))

    bridge_root      = _bridge_cache_dir()
    bridge_token     = os.path.join(bridge_root, "bridge_token.txt")
    local_runtime    = os.path.join(bridge_root, "runtime")
    local_modules    = os.path.join(bridge_root, "modules")
    local_lib        = os.path.join(bridge_root, "lib")
    local_python     = os.path.join(local_runtime, "python.exe")
    local_helper     = os.path.join(local_modules, "_pdf_helper.py")
    source_token     = _runtime_source_token()

    current_token = ""
    try:
        if os.path.isfile(bridge_token):
            with open(bridge_token, "r") as f:
                current_token = (f.read() or "").strip()
    except Exception:
        current_token = ""

    need_refresh = (current_token != source_token
                    or not os.path.isfile(local_python)
                    or not os.path.isfile(local_helper))
    if os.path.isdir(lib_src):
        need_refresh = need_refresh or (not os.path.isdir(local_lib))

    if need_refresh:
        print("[DocLink] Refreshing local PDF bridge cache...")
        try:
            if os.path.isdir(local_runtime):
                shutil.rmtree(local_runtime)
        except Exception:
            pass
        try:
            if os.path.isdir(local_modules):
                shutil.rmtree(local_modules)
        except Exception:
            pass
        try:
            if os.path.isdir(local_lib):
                shutil.rmtree(local_lib)
        except Exception:
            pass

        shutil.copytree(runtime_src, local_runtime)
        os.makedirs(local_modules)
        shutil.copy2(helper_src, local_helper)
        if os.path.isdir(lib_src):
            shutil.copytree(lib_src, local_lib)
        _strip_motw(bridge_root)
        try:
            with open(bridge_token, "w") as f:
                f.write(source_token)
        except Exception:
            pass

    resolved_local_lib = local_lib if os.path.isdir(local_lib) else _runtime_site_packages_dir(local_runtime)
    return local_python, local_helper, resolved_local_lib


def _resolve_pdf_bridge_paths(cpython_exe):
    """Resolve helper and lib paths for the selected CPython runtime."""
    global _PDF_HELPER, _LIB_DIR

    _PDF_HELPER = _PDF_HELPER_SRC
    _LIB_DIR    = _default_pdf_lib_dir()

    if not cpython_exe:
        return _PDF_HELPER, _LIB_DIR

    try:
        exe_norm = os.path.normcase(os.path.abspath(cpython_exe))
    except Exception:
        exe_norm = cpython_exe

    bridge_root = _bridge_cache_dir()
    try:
        bridge_norm = os.path.normcase(os.path.abspath(bridge_root))
    except Exception:
        bridge_norm = bridge_root

    if exe_norm.startswith(bridge_norm) or (_is_network_path(_MODULES_DIR) and exe_norm.startswith(os.path.normcase(os.path.abspath(_doclink_local_dir())))):
        try:
            _local_python, local_helper, local_lib = _stage_pdf_bridge_locally()
            _PDF_HELPER = local_helper
            _LIB_DIR    = local_lib
        except Exception as ex:
            print("[DocLink] Failed to resolve local PDF bridge paths: {}".format(ex))

    return _PDF_HELPER, _LIB_DIR


def _find_cpython():
    """
    Locate a CPython 3 executable that has PyMuPDF (fitz) installed.

    Search order:
      0. Bundled runtime/ next to this add-in  ← works in isolated C# add-ins
      1. PATH-style short names  (python / python3 / py)
      2. Well-known absolute paths on Windows
         – per-user AppData installs
         – system-wide Program Files installs
         – Windows Store alias directory
         – conda / Anaconda / Miniconda
         – pyenv-win shims
      3. Windows registry  (HKCU and HKLM Python core keys)

    Each candidate is first checked for Python 3, then confirmed that
    ``import fitz`` succeeds  (so we don't pick a Python that lacks PyMuPDF).
    Returns the full path string, or None if nothing suitable is found.
    """
    import platform

    from System.Diagnostics import Process, ProcessStartInfo

    # Fast path: a previous successful resolution wins until the source
    # runtime changes (mtime+size mismatch invalidates the cache).
    _cached = _load_cpython_cache()
    if _cached:
        print("[DocLink] Using cached runtime: {}".format(_cached))
        return _cached

    def _probe(exe):
        """Return exe if it is a real file, is Python 3, and has fitz.

        Uses .NET System.Diagnostics.Process to avoid IronPython subprocess
        socket/firewall issues.
        """
        if not exe or not os.path.isfile(exe):
            return None

        # ── Step 1: confirm it is Python 3 ─────────────────────────────────
        try:
            psi = ProcessStartInfo()
            psi.FileName = exe
            psi.Arguments = "--version"
            psi.UseShellExecute = False
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True
            psi.CreateNoWindow = True
            # FIX: Set local working directory to avoid UNC path issues
            psi.WorkingDirectory = tempfile.gettempdir()
            
            proc = Process.Start(psi)
            out = proc.StandardOutput.ReadToEnd()
            err = proc.StandardError.ReadToEnd()
            proc.WaitForExit(5000)
            
            version_text = (out or "") + (err or "")
            if proc.ExitCode != 0 or "Python 3" not in version_text:
                return None
        except Exception:
            return None

        # ── Step 2: confirm PyMuPDF (fitz) is importable ───────────────────
        try:
            psi2 = ProcessStartInfo()
            psi2.FileName = exe
            psi2.Arguments = '-c "import fitz; print(\'ok\')"'
            psi2.UseShellExecute = False
            psi2.RedirectStandardOutput = True
            psi2.RedirectStandardError = True
            psi2.CreateNoWindow = True
            psi2.WorkingDirectory = tempfile.gettempdir()
            
            proc2 = Process.Start(psi2)
            out2 = proc2.StandardOutput.ReadToEnd()
            proc2.WaitForExit(5000)
            
            if proc2.ExitCode == 0 and "ok" in (out2 or ""):
                return exe
        except Exception:
            pass
        return None

    # ── 0. Bundled portable runtime shipped WITH the add-in ───────────────────
    _pushbutton_root = os.path.dirname(_MODULES_DIR)
    _runtime_src     = os.path.join(_pushbutton_root, "runtime")
    _bundled_python  = os.path.join(_runtime_src, "python.exe")

    # Prefer a fully local staged bridge when the add-in is running from UNC.
    if _is_network_path(_MODULES_DIR) or _is_network_path(_EXT_DIR) or _is_network_path(_bundled_python):
        try:
            local_python, _local_helper, _local_lib = _stage_pdf_bridge_locally()
            result = _probe(local_python)
            if result:
                print("[DocLink] Using local staged PDF bridge runtime: {}".format(result))
                _save_cpython_cache(result)
                return result
        except Exception as ex:
            print("[DocLink] Local PDF bridge staging failed: {}".format(ex))
    
    result = _probe(_bundled_python)
    if result:
        print("[DocLink] Using bundled runtime: {}".format(result))
        _save_cpython_cache(result)
        return result

    # ── 0b. Network Fallback: Copy to Local AppData if network execution blocked ──
    if os.path.isdir(_runtime_src) and (_bundled_python.startswith("\\") or ":" not in _bundled_python):
        try:
            local_python, _local_helper, _local_lib = _stage_pdf_bridge_locally()

            result = _probe(local_python)
            if result:
                print("[DocLink] Using local cached PDF bridge runtime: {}".format(result))
                _save_cpython_cache(result)
                return result
        except Exception as e:
            print("[DocLink] Failed to copy PDF bridge locally: {}".format(e))

    # ── 1. PATH-based short names ─────────────────────────────────────────────
    for name in ("python", "python3", "py"):
        result = _probe(name)
        if result:
            _save_cpython_cache(result)
            return result

    # ── 2. Well-known absolute paths (Windows only) ───────────────────────────
    if platform.system() == "Windows":
        candidates = []

        # Safe env lookups – use hardcoded fallbacks if env is stripped
        user_profile = (os.environ.get("USERPROFILE")
                        or os.environ.get("HOMEPATH")
                        or r"C:\Users\Default")
        local_app    = (os.environ.get("LOCALAPPDATA")
                        or os.path.join(user_profile, "AppData", "Local"))
        app_data     = (os.environ.get("APPDATA")
                        or os.path.join(user_profile, "AppData", "Roaming"))
        prog_files   = (os.environ.get("ProgramFiles")   or r"C:\Program Files")
        prog_files86 = (os.environ.get("ProgramFiles(x86)") or r"C:\Program Files (x86)")
        prog_data    = (os.environ.get("ProgramData")    or r"C:\ProgramData")

        # Per-user AppData installs
        for ver in ("313", "312", "311", "310", "39", "38"):
            candidates.append(
                os.path.join(local_app, "Programs", "Python",
                             "Python{}".format(ver), "python.exe")
            )

        # System-wide installs
        for pf in (prog_files, prog_files86):
            for ver in ("313", "312", "311", "310", "39", "38"):
                candidates.append(
                    os.path.join(pf, "Python{}".format(ver), "python.exe")
                )
            candidates.append(os.path.join(pf, "Python", "python.exe"))

        # Windows Store Python alias directory
        for alias in ("python3.exe", "python.exe"):
            candidates.append(
                os.path.join(local_app, "Microsoft", "WindowsApps", alias)
            )

        # Anaconda / Miniconda – per-user and system
        for conda_name in ("Anaconda3", "Miniconda3", "anaconda3", "miniconda3"):
            candidates.append(os.path.join(user_profile, conda_name, "python.exe"))
            candidates.append(os.path.join(local_app,   conda_name, "python.exe"))
            candidates.append(os.path.join(prog_data,   conda_name, "python.exe"))

        # pyenv-win shims
        candidates.append(
            os.path.join(user_profile, ".pyenv", "pyenv-win", "shims", "python.exe")
        )

        for exe in candidates:
            result = _probe(exe)
            if result:
                _save_cpython_cache(result)
                return result

        # ── 3. Windows registry ───────────────────────────────────────────────
        try:
            import winreg
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                for base_key in (
                    r"SOFTWARE\Python\PythonCore",
                    r"SOFTWARE\WOW6432Node\Python\PythonCore",
                ):
                    try:
                        with winreg.OpenKey(hive, base_key) as bk:
                            idx = 0
                            while True:
                                try:
                                    ver_name = winreg.EnumKey(bk, idx)
                                    idx += 1
                                    try:
                                        with winreg.OpenKey(
                                            bk,
                                            "{}\\InstallPath".format(ver_name)
                                        ) as ik:
                                            install_dir, _ = winreg.QueryValueEx(ik, "")
                                            exe = os.path.join(
                                                install_dir.strip(), "python.exe"
                                            )
                                            result = _probe(exe)
                                            if result:
                                                _save_cpython_cache(result)
                                                return result
                                    except Exception:
                                        pass
                                except OSError:
                                    break
                    except Exception:
                        pass
        except Exception:
            pass

    return None


_CPYTHON        = _find_cpython()
_PDF_HELPER, _LIB_DIR = _resolve_pdf_bridge_paths(_CPYTHON)
_FITZ_AVAILABLE = (_CPYTHON is not None
                   and os.path.isfile(_PDF_HELPER))


# ── Screen helper ─────────────────────────────────────────────────────────────
def SystemParameters_WorkArea_Height():
    try:
        return SystemParameters.WorkArea.Height - 40
    except Exception:
        return 700
