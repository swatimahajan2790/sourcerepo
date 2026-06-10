# -*- coding: utf-8 -*-
"""
revit_ui_loader.py
==================
Shared helper for loading RevitUITemplate.xaml (and optionally the dark overlay)
into a WPF Window's resource scope before ShowDialog().

Compatible with IronPython 2.7 (pyRevit) running inside Revit 2021-2026.

USAGE — per tool (add these 3 lines before win.ShowDialog()):
-----------------------------------------------------------------
    import os, sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "shared"))
    from revit_ui_loader import apply_template_to_window, TEMPLATE_PATH

    win = MyToolWindow()
    apply_template_to_window(win, TEMPLATE_PATH)           # light theme
    # apply_template_to_window(win, TEMPLATE_PATH, dark=True)  # dark theme (GeometryControl)
    win.ShowDialog()

THEN INSIDE MyToolWindow.__init__:
    btn = Button()
    btn.Style = self.FindResource("Style.Button.Primary")

STYLE KEY REFERENCE (see RevitUITemplate.xaml for full list):
-----------------------------------------------------------------
  Buttons:
    Style.Button.Primary          Blue accent  — main action
    Style.Button.Secondary        Gray         — default/cancel
    Style.Button.Danger           Red          — destructive
    Style.Button.Toolbar          Transparent  — DocLink toolbar buttons
    Style.Button.Accent.Green     Green        — SmartPick select
    Style.Button.Icon             Square gray  — icon-only
    Style.Button.Compact          24px height  — tight UI rows
    Style.Button.Dark.Primary     Blue dark    — GeometryControl
    Style.Button.Dark.Secondary   Gray dark    — GeometryControl

  Inputs:
    Style.TextBox.Standard
    Style.TextBox.Search
    Style.TextBox.Dark
    Style.ComboBox.Standard
    Style.CheckBox.Standard
    Style.RadioButton.Standard

  ListBox:
    Style.ListBox.Standard        Single-select, light
    Style.ListBox.Extended        Multi-select,  light
    Style.ListBox.Dark            Single-select, dark
    Style.ListBox.Dark.Extended   Multi-select,  dark

  TabControl/TabItem:
    Style.TabControl.Standard
    Style.TabItem.Standard
    Style.TabItem.Tinted.Blue     D-Sync Tab 1
    Style.TabItem.Tinted.Green    D-Sync Tab 2
    Style.TabControl.Dark
    Style.TabItem.Dark            GeometryControl

  DataGrid:
    Style.DataGrid.Standard

  Text:
    Style.TextBlock.Standard
    Style.TextBlock.Heading
    Style.TextBlock.Title
    Style.TextBlock.Secondary
    Style.TextBlock.SectionLabel
    Style.TextBlock.StatusBar
    Style.TextBlock.OnDark
    Style.TextBlock.Status.OK / .Error / .Warn / .Info
    Style.TextBlock.Dark.Standard / .Heading / .Secondary
    Style.TextBlock.Dark.Status.Success / .Warn / .Error
    Style.Label.Standard
    Style.Label.Dark

  Notifications / Borders:
    Style.Border.Notification.Info / .Warning / .Error / .Success
    Style.Border.Notification.Dark.Info / .Warning / .Error
    Style.Border.AccentBar.Green    SmartPick top bar
    Style.Border.AccentBar.Blue
    Style.Border.StatusBar
    Style.Border.Toolbar.NavyBlue   DocLink Tab1 toolbar
    Style.Border.Toolbar.Purple     DocLink Tab2 toolbar

  Misc:
    Style.ProgressBar.Standard
    Style.ToolTip.Standard
    Style.ContextMenu.Standard
    Style.MenuItem.Standard
    Style.GroupBox.Standard
    Style.Separator.Standard

  Brushes (direct use via FindResource or StaticResource in XAML):
    Brush.Accent.Primary / .Green
    Brush.Text.Primary / .Secondary / .OnAccent / .OnDark
    Brush.Window.Background
    Brush.Panel.Background / .Secondary
    Brush.Status.OK / .Error / .Warn / .Info
    Brush.Toolbar.NavyBlue / .Purple
    Dark.Brush.Accent.Primary / .Green
    Dark.Brush.Text.Primary / .Secondary
    (see §2 and §3 in RevitUITemplate.xaml for complete list)

  Sizing constants:
    Height.Button.Standard (28) / .Compact (24) / .Large (34)
    Height.Input (24) / .Large (28)
    Height.DataGridRow (36)
    Font.Size.Standard (11) / .Body (12) / .DataGrid (13) / .Heading (14)
    Padding.Button.Standard / .Toolbar / .Compact
    Radius.Standard (3) / .Small (2) / .Medium (4) / .Pill (12)
"""

import os

# ── Lazy CLR imports (only resolved when this module is imported inside Revit) ──
try:
    import clr
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")

    from System.IO import FileStream, FileMode
    from System.Windows.Markup import XamlReader
    from System.Windows import ResourceDictionary

    _CLR_AVAILABLE = True
except Exception:
    # Allow the module to be imported outside Revit for documentation / linting
    _CLR_AVAILABLE = False

# ── Absolute path to this file's directory ──────────────────────────────────
_SHARED_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Convenience constants — use these in tool scripts ───────────────────────
TEMPLATE_PATH      = os.path.join(_SHARED_DIR, "RevitUITemplate.xaml")
TEMPLATE_DARK_PATH = os.path.join(_SHARED_DIR, "RevitUITemplate_Dark.xaml")

# ── Load-once cache (keyed by normalised absolute path) ─────────────────────
# Prevents adding duplicate MergedDictionaries when two tools open in the
# same Revit session, which would cause ambiguous StaticResource resolution.
_LOADED_DICTS = {}   # {norm_path: ResourceDictionary}


def _load_xaml(path):
    """
    Load a XAML ResourceDictionary from *path*, using the session cache.

    Parameters
    ----------
    path : str
        Absolute path to a .xaml file whose root element is ResourceDictionary.

    Returns
    -------
    System.Windows.ResourceDictionary

    Raises
    ------
    IOError
        If the file does not exist.
    TypeError
        If the XAML root is not a ResourceDictionary.
    """
    if not _CLR_AVAILABLE:
        raise RuntimeError("revit_ui_loader: CLR / WPF not available. "
                           "Must run inside Revit / IronPython.")

    norm = os.path.normcase(os.path.normpath(path))

    if norm in _LOADED_DICTS:
        return _LOADED_DICTS[norm]

    if not os.path.isfile(path):
        raise IOError("XAML template not found: {}".format(path))

    with FileStream(path, FileMode.Open) as fs:
        obj = XamlReader.Load(fs)

    if not isinstance(obj, ResourceDictionary):
        raise TypeError("Expected ResourceDictionary root in {}, got {}".format(
            path, type(obj).__name__))

    _LOADED_DICTS[norm] = obj
    return obj


def apply_template_to_window(window, xaml_path=None, dark=False):
    """
    Merge the Revit UI template into *window*'s resource scope.

    Call this before ``window.ShowDialog()`` or ``window.Show()``.
    After the call, the window and all child controls can reference style
    keys via ``FindResource("Style.Button.Primary")`` or directly in XAML
    markup as ``Style="{StaticResource Style.Button.Primary}"``.

    Parameters
    ----------
    window : System.Windows.Window
        The WPF Window to receive the styles.
    xaml_path : str, optional
        Absolute path to RevitUITemplate.xaml.
        Defaults to the ``RevitUITemplate.xaml`` in this shared folder.
    dark : bool, optional
        If True, also merge RevitUITemplate_Dark.xaml so all Brush.*
        keys take dark values.  Use for GeometryControl.

    Example
    -------
        win = LauncherWindow()
        apply_template_to_window(win)                    # light theme
        apply_template_to_window(win, dark=True)         # dark theme
        win.ShowDialog()
    """
    if xaml_path is None:
        xaml_path = TEMPLATE_PATH

    base_rd = _load_xaml(xaml_path)
    window.Resources.MergedDictionaries.Add(base_rd)

    if dark:
        dark_path = os.path.join(os.path.dirname(os.path.abspath(xaml_path)),
                                 "RevitUITemplate_Dark.xaml")
        dark_rd = _load_xaml(dark_path)
        window.Resources.MergedDictionaries.Add(dark_rd)


def get_resource(window, key):
    """
    Convenience wrapper around ``window.FindResource(key)``.

    Parameters
    ----------
    window : System.Windows.Window
    key    : str — resource key, e.g. "Style.Button.Primary"

    Returns
    -------
    object (Style, Brush, Double, Thickness, etc.)

    Raises
    ------
    System.Windows.ResourceReferenceKeyNotFoundException
        If the key is not found.  Ensure apply_template_to_window() was called.
    """
    return window.FindResource(key)


def get_brush(window, key):
    """Return a SolidColorBrush resource by key, e.g. "Brush.Accent.Primary"."""
    return window.FindResource(key)


def get_style(window, key):
    """Return a Style resource by key, e.g. "Style.Button.Primary"."""
    return window.FindResource(key)


# ── Quick smoke-test (run standalone, not inside Revit) ─────────────────────
if __name__ == "__main__":
    print("Template path  :", TEMPLATE_PATH)
    print("Dark path      :", TEMPLATE_DARK_PATH)
    print("Template exists:", os.path.isfile(TEMPLATE_PATH))
    print("Dark exists    :", os.path.isfile(TEMPLATE_DARK_PATH))
    print("CLR available  :", _CLR_AVAILABLE)
    print()
    print("Key style names exported by this template:")
    keys = [
        "Style.Button.Primary", "Style.Button.Secondary", "Style.Button.Danger",
        "Style.Button.Toolbar", "Style.Button.Accent.Green", "Style.Button.Icon",
        "Style.Button.Compact", "Style.Button.Dark.Primary", "Style.Button.Dark.Secondary",
        "Style.TextBox.Standard", "Style.TextBox.Search", "Style.TextBox.Dark",
        "Style.ComboBox.Standard",
        "Style.CheckBox.Standard", "Style.RadioButton.Standard",
        "Style.ListBox.Standard", "Style.ListBox.Extended",
        "Style.ListBox.Dark", "Style.ListBox.Dark.Extended",
        "Style.TabControl.Standard", "Style.TabItem.Standard",
        "Style.TabItem.Tinted.Blue", "Style.TabItem.Tinted.Green",
        "Style.TabControl.Dark", "Style.TabItem.Dark",
        "Style.DataGrid.Standard",
        "Style.ProgressBar.Standard",
        "Style.ToolTip.Standard",
        "Style.ContextMenu.Standard", "Style.MenuItem.Standard",
        "Style.GroupBox.Standard", "Style.Separator.Standard",
        "Style.Border.Notification.Info", "Style.Border.Notification.Warning",
        "Style.Border.Notification.Error", "Style.Border.Notification.Success",
        "Style.Border.Notification.Dark.Info",
        "Style.Border.Notification.Dark.Warning",
        "Style.Border.Notification.Dark.Error",
        "Style.Border.AccentBar.Green", "Style.Border.AccentBar.Blue",
        "Style.Border.StatusBar",
        "Style.Border.Toolbar.NavyBlue", "Style.Border.Toolbar.Purple",
        "Style.TextBlock.Standard", "Style.TextBlock.Heading", "Style.TextBlock.Title",
        "Style.TextBlock.Secondary", "Style.TextBlock.SectionLabel",
        "Style.TextBlock.StatusBar", "Style.TextBlock.OnDark",
        "Style.TextBlock.Status.OK", "Style.TextBlock.Status.Error",
        "Style.TextBlock.Status.Warn", "Style.TextBlock.Status.Info",
        "Style.TextBlock.Dark.Standard", "Style.TextBlock.Dark.Heading",
        "Style.TextBlock.Dark.Secondary",
        "Style.TextBlock.Dark.Status.Success",
        "Style.TextBlock.Dark.Status.Warn", "Style.TextBlock.Dark.Status.Error",
        "Style.Label.Standard", "Style.Label.Dark",
        "Style.Window.Base",
    ]
    for k in sorted(keys):
        print("  ", k)
