# -*- coding: utf-8 -*-
"""
dialogs.py
----------
Secondary WPF dialog windows for DocLink Tab 1.

Classes
-------
AddImportDialog       – collects settings for a new import
EditImportDialog      – edits settings for an existing import
InstanceHistoryWindow – shows the full lifecycle log for one record
"""

import os, sys, re, json, datetime

from _imports import (
    System, Window, Thickness, HorizontalAlignment, VerticalAlignment,
    Visibility, GridLength, GridUnitType, FontWeights,
    Grid, RowDefinition, ColumnDefinition, StackPanel, WrapPanel, ScrollViewer,
    Button, Label, TextBox, TextBlock, ComboBox, ComboBoxItem, CheckBox,
    RadioButton,
    Separator, DataGrid, DataGridTextColumn, DataGridSelectionMode,
    WPFOrientation, TabControl, TabItem,
    Binding, SolidColorBrush, Color, Brushes, ObservableCollection,
    SystemParameters,
    WinForms, WFOpenFileDialog, WFDialogResult,
    apply_template_to_window, _TMPL_PATH, _RES,
    SystemParameters_WorkArea_Height,
    forms, revit,
    _FITZ_AVAILABLE, _WORD_AVAILABLE, _EXCEL_AVAILABLE,
)
from utils import (
    _safe_int, _as_bool, _combo_selected_text,
    detect_file_type, SUPPORTED_FILES_FILTER, DPI_CHOICES,
    _sanitize_import_name, _get_auto_user,
)
from models import DocLinkRow
from pdf_ops  import get_pdf_page_count
from word_ops import get_word_page_count
from excel_ops import get_excel_print_area, list_excel_defined_names, get_excel_sheets_com
from schedule_tab import list_sheets


# ─────────────────────────────────────────────────────────────────────────────
# Shared helpers for the Excel "Name Manager" picker used by both Add/Edit.
# ─────────────────────────────────────────────────────────────────────────────

def _populate_named_range_combo(combo, excel_path, sheet_name, preferred_name=None):
    """
    Fill the Named-Range ComboBox with workbook-scoped + sheet-scoped names
    visible from `sheet_name`. Each item stores the full name-info dict on
    its Tag for later resolution. Returns True iff at least one name exists.
    """
    combo.Items.Clear()
    if not excel_path:
        return False
    try:
        names = list_excel_defined_names(excel_path, sheet_name)
    except Exception as ex:
        print("[DocLinkManager] list_excel_defined_names failed: {}".format(ex))
        names = []
    if not names:
        return False
    selected_idx = 0
    for i, n in enumerate(names):
        item = ComboBoxItem()
        scope_tag = "Workbook" if n["is_workbook_scope"] else "Sheet: {}".format(n["scope_sheet"] or "?")
        item.Content = "{}    ({}  —  {})".format(n["name"], n["range_address"], scope_tag)
        item.Tag = n
        combo.Items.Add(item)
        if preferred_name and n["name"] == preferred_name:
            selected_idx = i
    combo.SelectedIndex = selected_idx
    return True

class AddImportDialog(Window):
    """
    Gathers: name, source file, path type, excel options,
    page/DPI/transparent import settings.
    "Created By" is no longer a UI field — user metadata is derived
    automatically from the Revit / OS session (Phase 3).
    After ShowDialog() → True, read .result dict.
    """

    def __init__(self):
        self.result           = None
        self._src_path        = None
        self._file_type       = None
        # Phase 2: track whether the Import Name was auto-generated so
        # browsing again only overwrites the name if the user hasn't
        # manually customised it.
        self._auto_named      = True   # True = current name came from autofill
        self._prev_auto_name  = ""     # last value we wrote via autofill
        self._updating_name   = False  # re-entrancy guard for TextChanged
        self._sheet_data      = {}     # {name: {has_print_area: bool, print_area: str}}
        self._setup()

    def _setup(self):
        self.Title  = "Add New DocLink Import"
        self.Width  = 560
        self.MinHeight = 240
        self.SizeToContent = System.Windows.SizeToContent.Height
        self.MaxHeight = SystemParameters_WorkArea_Height()
        self.ResizeMode = System.Windows.ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        apply_template_to_window(self, _TMPL_PATH)
        if _RES: self.Background = _RES["Brush.Window.Background"]

        outer = Grid()
        r0 = RowDefinition(); r0.Height = GridLength(1, GridUnitType.Star)
        r1 = RowDefinition(); r1.Height = GridLength.Auto
        outer.RowDefinitions.Add(r0)
        outer.RowDefinitions.Add(r1)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled

        form = StackPanel()
        form.Margin = Thickness(18, 14, 18, 8)

        def section(text):
            lbl = Label()
            lbl.Content = text
            lbl.FontWeight = FontWeights.Bold
            lbl.Padding = Thickness(0)
            lbl.Margin = Thickness(0, 8, 0, 2)
            return lbl

        def textbox(default=""):
            tb = TextBox()
            tb.Text = default
            tb.Padding = Thickness(6, 5, 6, 5)
            tb.Margin = Thickness(0, 0, 0, 2)
            return tb

        # Metadata
        form.Children.Add(section("Import Name *"))
        self._name_box = textbox("")
        # Phase 2: detect manual edits so autofill won't clobber them
        self._name_box.TextChanged += self._on_name_changed
        form.Children.Add(self._name_box)

        # (Phase 3: "Created By" field removed — user is derived automatically)

        form.Children.Add(section("Source File *"))

        file_row_grid = Grid()
        c0 = ColumnDefinition(); c0.Width = GridLength(1, GridUnitType.Star)
        c1 = ColumnDefinition(); c1.Width = GridLength.Auto
        file_row_grid.ColumnDefinitions.Add(c0)
        file_row_grid.ColumnDefinitions.Add(c1)
        file_row_grid.Margin = Thickness(0, 0, 0, 2)

        self._src_box = TextBox()
        self._src_box.IsReadOnly = True
        self._src_box.Padding = Thickness(6, 5, 6, 5)
        self._src_box.Background = SolidColorBrush(Color.FromRgb(245, 247, 250))
        self._src_box.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(self._src_box, 0)
        file_row_grid.Children.Add(self._src_box)

        browse = Button()
        browse.Content = " Browse… "
        browse.Margin = Thickness(6, 0, 0, 0)
        browse.Padding = Thickness(12, 5, 12, 5)
        browse.VerticalAlignment = VerticalAlignment.Center
        browse.Click += self._browse
        Grid.SetColumn(browse, 1)
        file_row_grid.Children.Add(browse)
        form.Children.Add(file_row_grid)

        self._detect_lbl = TextBlock()
        self._detect_lbl.Text = ""
        self._detect_lbl.Foreground = SolidColorBrush(Color.FromRgb(0, 110, 180))
        self._detect_lbl.Margin = Thickness(2, 2, 0, 4)
        self._detect_lbl.FontSize = 12
        form.Children.Add(self._detect_lbl)

        form.Children.Add(section("Path Type *"))
        self._path_combo = ComboBox()
        self._path_combo.Margin = Thickness(0, 0, 0, 4)
        self._path_combo.Padding = Thickness(4, 4, 4, 4)
        for opt in ["Absolute", "Relative"]:
            item = ComboBoxItem(); item.Content = opt
            self._path_combo.Items.Add(item)
        self._path_combo.SelectedIndex = 0
        form.Children.Add(self._path_combo)

        # Excel-only panel
        self._excel_panel = StackPanel()
        self._excel_panel.Visibility = Visibility.Collapsed
        self._excel_panel.Margin = Thickness(0, 4, 0, 0)

        xl_border = System.Windows.Controls.Border()
        xl_border.BorderBrush = SolidColorBrush(Color.FromRgb(180, 200, 230))
        xl_border.BorderThickness = Thickness(1)
        xl_border.CornerRadius = System.Windows.CornerRadius(4)
        xl_border.Padding = Thickness(10, 8, 10, 10)
        xl_border.Background = SolidColorBrush(Color.FromRgb(245, 250, 255))
        xl_border.Margin = Thickness(0, 4, 0, 4)

        xl_inner = StackPanel()

        xl_title = TextBlock()
        xl_title.Text = "Excel Options"
        xl_title.FontWeight = FontWeights.Bold
        xl_title.Foreground = SolidColorBrush(Color.FromRgb(0, 80, 160))
        xl_title.Margin = Thickness(0, 0, 0, 6)
        xl_inner.Children.Add(xl_title)

        sheet_lbl = Label(); sheet_lbl.Content = "Sheet Name  (blank = active sheet)"
        sheet_lbl.Padding = Thickness(0); sheet_lbl.Margin = Thickness(0, 0, 0, 2)
        xl_inner.Children.Add(sheet_lbl)
        self._sheet_combo = ComboBox()
        self._sheet_combo.IsEditable = True
        self._sheet_combo.Padding = Thickness(4, 3, 4, 3)
        self._sheet_combo.Margin = Thickness(0, 0, 0, 2)
        self._sheet_combo.SelectionChanged += self._on_sheet_selection_changed
        xl_inner.Children.Add(self._sheet_combo)

        # ── Range mode: Named Range (from Excel Name Manager) vs Manual ──
        mode_label = Label()
        mode_label.Content = "Range Source"
        mode_label.Padding = Thickness(0)
        mode_label.Margin = Thickness(0, 8, 0, 2)
        mode_label.FontWeight = FontWeights.Bold
        xl_inner.Children.Add(mode_label)

        mode_row = StackPanel()
        mode_row.Orientation = WPFOrientation.Horizontal
        mode_row.Margin = Thickness(0, 0, 0, 2)

        self._rb_named = RadioButton()
        self._rb_named.Content = "Named Range"
        self._rb_named.GroupName = "AddRangeMode"
        self._rb_named.Margin = Thickness(0, 0, 16, 0)
        self._rb_named.Checked += self._on_range_mode_changed
        mode_row.Children.Add(self._rb_named)

        self._rb_manual = RadioButton()
        self._rb_manual.Content = "Manual Range"
        self._rb_manual.GroupName = "AddRangeMode"
        self._rb_manual.IsChecked = True  # default
        self._rb_manual.Checked += self._on_range_mode_changed
        mode_row.Children.Add(self._rb_manual)

        xl_inner.Children.Add(mode_row)

        self._named_combo = ComboBox()
        self._named_combo.Margin = Thickness(0, 4, 0, 2)
        self._named_combo.Padding = Thickness(4, 3, 4, 3)
        self._named_combo.IsEnabled = False
        self._named_combo.SelectionChanged += self._on_named_range_changed
        xl_inner.Children.Add(self._named_combo)

        self._named_hint = TextBlock()
        self._named_hint.Text = ""
        self._named_hint.FontSize = 11
        self._named_hint.Foreground = SolidColorBrush(Color.FromRgb(110, 110, 110))
        self._named_hint.Margin = Thickness(2, 0, 0, 4)
        xl_inner.Children.Add(self._named_hint)

        range_lbl = Label(); range_lbl.Content = "Cell Range  e.g. A1:G25  *"
        range_lbl.Padding = Thickness(0); range_lbl.Margin = Thickness(0, 6, 0, 2)
        xl_inner.Children.Add(range_lbl)
        self._range_box = textbox("A1:G25")
        xl_inner.Children.Add(self._range_box)

        xl_border.Child = xl_inner
        self._excel_panel.Children.Add(xl_border)
        form.Children.Add(self._excel_panel)

        # Common import options
        self._common_panel = StackPanel()
        self._common_panel.Visibility = Visibility.Collapsed
        self._common_panel.Margin = Thickness(0, 4, 0, 0)

        imp_border = System.Windows.Controls.Border()
        imp_border.BorderBrush = SolidColorBrush(Color.FromRgb(190, 210, 200))
        imp_border.BorderThickness = Thickness(1)
        imp_border.CornerRadius = System.Windows.CornerRadius(4)
        imp_border.Padding = Thickness(10, 8, 10, 10)
        imp_border.Background = SolidColorBrush(Color.FromRgb(247, 252, 248))
        imp_border.Margin = Thickness(0, 4, 0, 4)

        imp_inner = StackPanel()

        imp_title = TextBlock()
        imp_title.Text = "Import Options"
        imp_title.FontWeight = FontWeights.Bold
        imp_title.Foreground = SolidColorBrush(Color.FromRgb(0, 110, 70))
        imp_title.Margin = Thickness(0, 0, 0, 6)
        imp_inner.Children.Add(imp_title)

        self._page_panel = StackPanel()
        self._page_panel.Visibility = Visibility.Collapsed
        page_lbl = Label(); page_lbl.Content = "Page Number  (PDF / Word)"
        page_lbl.Padding = Thickness(0); page_lbl.Margin = Thickness(0, 0, 0, 2)
        self._page_panel.Children.Add(page_lbl)
        self._page_combo = ComboBox()
        self._page_combo.Margin = Thickness(0, 0, 0, 2)
        self._page_combo.Padding = Thickness(4, 4, 4, 4)
        # Default: single page until a file is browsed
        ci = ComboBoxItem(); ci.Content = "1"
        self._page_combo.Items.Add(ci)
        self._page_combo.SelectedIndex = 0
        self._page_panel.Children.Add(self._page_combo)
        imp_inner.Children.Add(self._page_panel)

        dpi_lbl = Label(); dpi_lbl.Content = "Import DPI"
        dpi_lbl.Padding = Thickness(0); dpi_lbl.Margin = Thickness(0, 6, 0, 2)
        imp_inner.Children.Add(dpi_lbl)
        self._dpi_combo = ComboBox()
        self._dpi_combo.Margin = Thickness(0, 0, 0, 2)
        self._dpi_combo.Padding = Thickness(4, 4, 4, 4)
        for val in DPI_CHOICES:
            item = ComboBoxItem(); item.Content = val
            self._dpi_combo.Items.Add(item)
        self._dpi_combo.SelectedIndex = 4  # 300 DPI
        imp_inner.Children.Add(self._dpi_combo)

        self._transparent_chk = CheckBox()
        self._transparent_chk.Content = "Remove background"
        self._transparent_chk.Margin = Thickness(0, 8, 0, 0)
        self._transparent_chk.IsChecked = False
        imp_inner.Children.Add(self._transparent_chk)

        imp_border.Child = imp_inner
        self._common_panel.Children.Add(imp_border)
        form.Children.Add(self._common_panel)

        scroll.Content = form
        Grid.SetRow(scroll, 0)
        outer.Children.Add(scroll)

        btn_bar = Grid()
        btn_bar.Background = SolidColorBrush(Color.FromRgb(240, 244, 250))

        btn_sep = Separator()
        btn_bar.Children.Add(btn_sep)

        btn_inner = StackPanel()
        btn_inner.Orientation = WPFOrientation.Horizontal
        btn_inner.HorizontalAlignment = HorizontalAlignment.Right
        btn_inner.Margin = Thickness(16, 10, 16, 10)

        ok = Button()
        ok.Content = "✔  Add Import"
        ok.Width = 130; ok.Height = 36
        ok.Margin = Thickness(0, 0, 8, 0)
        ok.Background = SolidColorBrush(Color.FromRgb(0, 130, 80))
        ok.Foreground = Brushes.White
        ok.FontWeight = FontWeights.Bold
        ok.BorderThickness = Thickness(0)
        ok.FontSize = 13
        ok.Click += self._ok
        btn_inner.Children.Add(ok)

        cancel = Button()
        cancel.Content = "Cancel"
        cancel.Width = 86; cancel.Height = 36
        cancel.FontSize = 13
        cancel.Click += self._cancel
        btn_inner.Children.Add(cancel)

        btn_bar.Children.Add(btn_inner)
        Grid.SetRow(btn_bar, 1)
        outer.Children.Add(btn_bar)

        self.Content = outer

    def _on_name_changed(self, s, e):
        """Track whether the name box still contains an auto-generated value."""
        if self._updating_name:
            return
        # If the user manually changed the text away from our last autofill
        # value, remember that so we won't overwrite it on next browse.
        self._auto_named = (self._name_box.Text == self._prev_auto_name)

    def _on_sheet_selection_changed(self, s, e):
        """When the user picks a sheet, try to auto-detect its print area
        and refresh the Name-Manager dropdown for that sheet's scope."""
        if self._file_type != "excel" or not self._src_path:
            return

        # Get sheet name from combo (either selected or typed)
        item = self._sheet_combo.SelectedItem
        sheet_name = None
        if item and hasattr(item, "Tag"):
            sheet_name = item.Tag
        else:
            sheet_name = self._sheet_combo.Text.strip()

        # Refresh defined-name list for the new sheet scope (kept regardless
        # of which range mode is active — switching modes shouldn't drop it).
        try:
            had_any = _populate_named_range_combo(
                self._named_combo, self._src_path, sheet_name)
            if not had_any:
                self._named_combo.IsEnabled = False
                self._rb_named.IsEnabled = False
                if self._rb_named.IsChecked:
                    self._rb_manual.IsChecked = True
            else:
                self._rb_named.IsEnabled = True
                # only enable the combo if the user is in named-range mode
                self._named_combo.IsEnabled = bool(self._rb_named.IsChecked)
        except Exception:
            pass

        if sheet_name:
            # 1. Try cached OpenXML data first (exact print area defined in file)
            sh_info = self._sheet_data.get(sheet_name)
            if sh_info and sh_info.get("print_area"):
                pa = sh_info["print_area"]
                self._range_box.Text = pa
                self._detect_lbl.Text = "✔  Print area detected (from file): {}".format(pa)
                return

            # 2. Fallback to COM for used range or typed sheet names
            try:
                print_area = get_excel_print_area(self._src_path, sheet_name)
                if print_area:
                    self._range_box.Text = print_area
                    self._detect_lbl.Text = "✔  Range detected: {}".format(print_area)
            except Exception:
                pass

    def _on_range_mode_changed(self, s, e):
        """Toggle which controls are active when the radio selection changes."""
        if not hasattr(self, "_rb_named"):
            return
        use_named = bool(self._rb_named.IsChecked)
        # Enable the named-range combo only if there are entries to pick.
        has_items = self._named_combo.Items.Count > 0
        self._named_combo.IsEnabled = use_named and has_items
        # In Named mode the chosen name drives both the sheet and the range,
        # so grey those inputs to make it obvious they aren't user-editable.
        self._range_box.IsEnabled = not use_named
        self._sheet_combo.IsEnabled = not use_named
        # When switching INTO named mode, reflect the selected name's range
        # into the textbox so the user can see what will be exported.
        if use_named and has_items:
            self._on_named_range_changed(None, None)
        else:
            self._named_hint.Text = ""

    def _on_named_range_changed(self, s, e):
        """Mirror the selected name's resolved A1 reference into the textbox,
        and force the sheet combo to the name's target sheet so a
        workbook-scoped name that points off-sheet still resolves correctly
        at export time."""
        item = self._named_combo.SelectedItem
        if item is None or not hasattr(item, "Tag") or item.Tag is None:
            self._named_hint.Text = ""
            return
        info = item.Tag
        rng  = info.get("range_address") or ""
        sheet = info.get("ref_sheet") or ""
        if rng:
            self._range_box.Text = rng
            scope = "workbook-scoped" if info.get("is_workbook_scope") else "sheet-scoped"
            self._named_hint.Text = "→ {}!{}  ({})".format(sheet, rng, scope)
        # Sync the sheet combo to the name's resolved sheet — guarded so we
        # don't re-trigger the named-range refresh recursively.
        if sheet:
            self._sync_sheet_combo_to(sheet)

    def _sync_sheet_combo_to(self, sheet_name):
        """Select the sheet whose Tag matches `sheet_name`. Suppresses the
        SelectionChanged side-effect that would otherwise reload the
        named-range list and clobber the user's pick."""
        if not sheet_name or self._sheet_combo.Items.Count == 0:
            return
        for i in range(self._sheet_combo.Items.Count):
            it = self._sheet_combo.Items[i]
            tag = getattr(it, "Tag", None)
            if tag == sheet_name:
                if self._sheet_combo.SelectedIndex != i:
                    # Detach the handler while we move the selection so
                    # _on_sheet_selection_changed doesn't repopulate the
                    # named-range combo (and drop the user's current pick).
                    try:
                        self._sheet_combo.SelectionChanged -= self._on_sheet_selection_changed
                    except Exception:
                        pass
                    try:
                        self._sheet_combo.SelectedIndex = i
                    finally:
                        try:
                            self._sheet_combo.SelectionChanged += self._on_sheet_selection_changed
                        except Exception:
                            pass
                return

    def _set_type_ui(self, ftype):
        self._file_type = ftype
        self._excel_panel.Visibility = Visibility.Visible if ftype == "excel" else Visibility.Collapsed
        self._common_panel.Visibility = Visibility.Visible if ftype in ("excel", "pdf", "word", "image") else Visibility.Collapsed
        self._page_panel.Visibility = Visibility.Visible if ftype in ("pdf", "word") else Visibility.Collapsed
        self._transparent_chk.Visibility = Visibility.Visible if ftype in ("excel", "word", "image", "pdf") else Visibility.Collapsed

    def _browse(self, s, e):
        dlg = System.Windows.Forms.OpenFileDialog()
        dlg.Title  = "Select Source File"
        dlg.Filter = SUPPORTED_FILES_FILTER
        if dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK:
            return

        path = dlg.FileName
        self._src_path = path
        self._src_box.Text = path
        ftype = detect_file_type(path)
        self._set_type_ui(ftype)

        if ftype == "excel":
            self._detect_lbl.Text = "✔  Excel workbook detected"

            # Populate sheets
            sheets = None
            try:
                # Try OpenXML parsing first (works for all Excel formats)
                sheets = list_sheets(path)
            except Exception as ex:
                # Fallback to COM for macro-enabled files or if OpenXML fails
                sheets = get_excel_sheets_com(path)
                if sheets:
                    print("[DocLinkManager] OpenXML parsing failed, using COM fallback: {}".format(ex))
            
            if sheets:
                self._sheet_data = {sh['name']: sh for sh in sheets}
                self._sheet_combo.Items.Clear()
                auto_sel_idx = -1
                for i, sh in enumerate(sheets):
                    item = ComboBoxItem()
                    marker = "  [Print Area]" if sh.get("has_print_area") else ""
                    item.Content = "{}{}".format(sh['name'], marker)
                    item.Tag = sh['name'] # store clean name
                    self._sheet_combo.Items.Add(item)
                    if sh.get('has_print_area') and auto_sel_idx == -1:
                        auto_sel_idx = i

                if auto_sel_idx >= 0:
                    self._sheet_combo.SelectedIndex = auto_sel_idx
                elif self._sheet_combo.Items.Count > 0:
                    self._sheet_combo.SelectedIndex = 0
                # _on_sheet_selection_changed will refresh the named-range
                # combo for the active sheet's scope; if no sheet was
                # auto-selected, populate workbook-scoped names now.
                if auto_sel_idx < 0:
                    _populate_named_range_combo(
                        self._named_combo, self._src_path, None)
                    if self._named_combo.Items.Count == 0:
                        self._rb_named.IsEnabled = False
            else:
                self._detect_lbl.Text = "⚠  Could not list sheets from this workbook"
        elif ftype == "word":
            msg = "✔  Word document detected"
            if not _WORD_AVAILABLE:
                msg += "  (Word Interop not available on this machine)"
            else:
                count = get_word_page_count(path)
                if count > 0:
                    self._populate_page_combo(count)
                    msg += "  ({} page{})".format(count, "s" if count > 1 else "")
            self._detect_lbl.Text = msg
        elif ftype == "pdf":
            count = get_pdf_page_count(path)
            if count > 0:
                self._populate_page_combo(count)
                self._detect_lbl.Text = "✔  PDF detected  ({} page{})".format(
                    count, "s" if count > 1 else "")
            else:
                self._detect_lbl.Text = "✔  PDF detected  (via PyMuPDF)"
        else:
            ext = os.path.splitext(path)[1].lower()
            self._detect_lbl.Text = "✔  Image detected ({})".format(ext)

        # ── Phase 2: autofill Import Name from file basename ─────────────
        suggested = _sanitize_import_name(
            os.path.splitext(os.path.basename(path))[0])
        if suggested and (self._auto_named or not self._name_box.Text.strip()):
            self._updating_name = True
            try:
                self._name_box.Text = suggested
                self._prev_auto_name = suggested
                self._auto_named = True
            finally:
                self._updating_name = False

    def _populate_page_combo(self, total_pages, selected_page=1):
        """Fill the page ComboBox with page numbers 1..total_pages."""
        self._page_combo.Items.Clear()
        sel_idx = 0
        for p in range(1, total_pages + 1):
            ci = ComboBoxItem(); ci.Content = str(p)
            self._page_combo.Items.Add(ci)
            if p == selected_page:
                sel_idx = p - 1
        self._page_combo.SelectedIndex = min(sel_idx, max(0, total_pages - 1))

    def _ok(self, s, e):
        name = self._name_box.Text.strip()
        if not name:
            forms.alert("Enter an import name."); return
        # Phase 3: user no longer entered manually
        if not self._src_path:
            forms.alert("Browse to a source file."); return
        if self._file_type == "excel" and not self._range_box.Text.strip():
            forms.alert("Enter a cell range."); return
        sel = self._path_combo.SelectedItem
        path_type = sel.Content if sel else "Absolute"

        page_sel = self._page_combo.SelectedItem
        page_num = page_sel.Content if page_sel else "1"

        # Extract clean sheet name from Tag if available
        sheet_name = ""
        named_range = ""
        if self._file_type == "excel":
            sel_item = self._sheet_combo.SelectedItem
            if sel_item and hasattr(sel_item, "Tag"):
                sheet_name = sel_item.Tag
            else:
                sheet_name = self._sheet_combo.Text.strip()
            if self._rb_named.IsChecked:
                nm_item = self._named_combo.SelectedItem
                if nm_item is None or not hasattr(nm_item, "Tag") or nm_item.Tag is None:
                    forms.alert("Pick a Named Range or switch to Manual Range.")
                    return
                info = nm_item.Tag
                named_range = info.get("name", "")
                # When the name is sheet-scoped, force the export to use that
                # sheet — otherwise Excel could resolve to a different sheet.
                if not info.get("is_workbook_scope"):
                    sheet_name = info.get("scope_sheet") or sheet_name

        self.result = {
            "name":        name,
            "user":        _get_auto_user(),          # Phase 3: auto-derived
            "auto_named":  self._auto_named,           # Phase 2: schema field
            "source_path": self._src_path,
            "path_type":   path_type,
            "file_type":   self._file_type,
            "sheet_name":  sheet_name,
            "range_addr":  self._range_box.Text.strip() if self._file_type == "excel" else "",
            "named_range": named_range,
            "page_number": page_num if self._file_type in ("pdf", "word") else "1",
            "dpi":         _combo_selected_text(self._dpi_combo, "300"),
            "transparent": _as_bool(self._transparent_chk.IsChecked),
        }
        self.DialogResult = True
        self.Close()

    def _cancel(self, s, e):
        self.DialogResult = False
        self.Close()


class EditImportDialog(Window):
    """
    Pre-populated dialog that lets the user change:
      - Source file (browse)
      - Path type
      - Sheet name & cell range (Excel only)
      - Page number (PDF/Word)
      - DPI / transparency
    After ShowDialog() → True, read .result dict.
    """

    def __init__(self, record):
        self.result      = None
        self._record     = record
        self._src_path   = record.get("source_path", "")
        self._file_type  = record.get("file_type", "image")
        self._sheet_data = {}
        self._setup()

    def _setup(self):
        self.Title  = "Edit DocLink Import"
        self.Width  = 560
        self.MinHeight = 240
        self.SizeToContent = System.Windows.SizeToContent.Height
        self.MaxHeight = SystemParameters_WorkArea_Height()
        self.ResizeMode = System.Windows.ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        apply_template_to_window(self, _TMPL_PATH)
        if _RES: self.Background = _RES["Brush.Window.Background"]

        outer = Grid()
        r0 = RowDefinition(); r0.Height = GridLength(1, GridUnitType.Star)
        r1 = RowDefinition(); r1.Height = GridLength.Auto
        outer.RowDefinitions.Add(r0)
        outer.RowDefinitions.Add(r1)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled

        form = StackPanel()
        form.Margin = Thickness(18, 14, 18, 8)

        def section(text):
            lbl = Label()
            lbl.Content = text
            lbl.FontWeight = FontWeights.Bold
            lbl.Padding = Thickness(0)
            lbl.Margin = Thickness(0, 8, 0, 2)
            return lbl

        def textbox(default=""):
            tb = TextBox()
            tb.Text = default
            tb.Padding = Thickness(6, 5, 6, 5)
            tb.Margin = Thickness(0, 0, 0, 2)
            return tb

        form.Children.Add(section("Source File *"))

        file_row_grid = Grid()
        c0 = ColumnDefinition(); c0.Width = GridLength(1, GridUnitType.Star)
        c1 = ColumnDefinition(); c1.Width = GridLength.Auto
        file_row_grid.ColumnDefinitions.Add(c0)
        file_row_grid.ColumnDefinitions.Add(c1)
        file_row_grid.Margin = Thickness(0, 0, 0, 2)

        self._src_box = TextBox()
        self._src_box.IsReadOnly = True
        self._src_box.Text = self._src_path
        self._src_box.Padding = Thickness(6, 5, 6, 5)
        self._src_box.Background = SolidColorBrush(Color.FromRgb(245, 247, 250))
        self._src_box.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(self._src_box, 0)
        file_row_grid.Children.Add(self._src_box)

        browse = Button()
        browse.Content = " Browse… "
        browse.Margin = Thickness(6, 0, 0, 0)
        browse.Padding = Thickness(12, 5, 12, 5)
        browse.VerticalAlignment = VerticalAlignment.Center
        browse.Click += self._browse
        Grid.SetColumn(browse, 1)
        file_row_grid.Children.Add(browse)
        form.Children.Add(file_row_grid)

        self._detect_lbl = TextBlock()
        ext = os.path.splitext(self._src_path)[1].lower() if self._src_path else ""
        self._detect_lbl.Text = "Current file: {}".format(ext.upper() if ext else "unknown")
        self._detect_lbl.Foreground = SolidColorBrush(Color.FromRgb(0, 110, 180))
        self._detect_lbl.Margin = Thickness(2, 2, 0, 4)
        self._detect_lbl.FontSize = 12
        form.Children.Add(self._detect_lbl)

        form.Children.Add(section("Import Name"))
        self._name_box = textbox(self._record.get("name", ""))
        self._name_box.ToolTip = "Optional label for this import (leave blank to keep current name)"
        form.Children.Add(self._name_box)

        form.Children.Add(section("Path Type *"))
        self._path_combo = ComboBox()
        self._path_combo.Margin = Thickness(0, 0, 0, 4)
        self._path_combo.Padding = Thickness(4, 4, 4, 4)
        for opt in ["Absolute", "Relative"]:
            item = ComboBoxItem(); item.Content = opt
            self._path_combo.Items.Add(item)
        current_path_type = self._record.get("path_type", "Absolute")
        self._path_combo.SelectedIndex = 1 if current_path_type == "Relative" else 0
        form.Children.Add(self._path_combo)

        self._excel_panel = StackPanel()
        self._excel_panel.Margin = Thickness(0, 4, 0, 0)

        xl_border = System.Windows.Controls.Border()
        xl_border.BorderBrush = SolidColorBrush(Color.FromRgb(180, 200, 230))
        xl_border.BorderThickness = Thickness(1)
        xl_border.CornerRadius = System.Windows.CornerRadius(4)
        xl_border.Padding = Thickness(10, 8, 10, 10)
        xl_border.Background = SolidColorBrush(Color.FromRgb(245, 250, 255))
        xl_border.Margin = Thickness(0, 4, 0, 4)

        xl_inner = StackPanel()

        xl_title = TextBlock()
        xl_title.Text = "Excel Options"
        xl_title.FontWeight = FontWeights.Bold
        xl_title.Foreground = SolidColorBrush(Color.FromRgb(0, 80, 160))
        xl_title.Margin = Thickness(0, 0, 0, 6)
        xl_inner.Children.Add(xl_title)

        sheet_lbl = Label(); sheet_lbl.Content = "Sheet Name  (blank = active sheet)"
        sheet_lbl.Padding = Thickness(0); sheet_lbl.Margin = Thickness(0, 0, 0, 2)
        xl_inner.Children.Add(sheet_lbl)
        self._sheet_combo = ComboBox()
        self._sheet_combo.IsEditable = True
        self._sheet_combo.Padding = Thickness(4, 3, 4, 3)
        self._sheet_combo.Margin = Thickness(0, 0, 0, 2)
        self._sheet_combo.Text = self._record.get("sheet_name", "")
        self._sheet_combo.SelectionChanged += self._on_sheet_selection_changed
        xl_inner.Children.Add(self._sheet_combo)

        # ── Range mode (mirrors AddImportDialog) ──
        saved_named = self._record.get("named_range", "") or ""
        use_named_default = bool(saved_named)

        mode_label = Label()
        mode_label.Content = "Range Source"
        mode_label.Padding = Thickness(0)
        mode_label.Margin = Thickness(0, 8, 0, 2)
        mode_label.FontWeight = FontWeights.Bold
        xl_inner.Children.Add(mode_label)

        mode_row = StackPanel()
        mode_row.Orientation = WPFOrientation.Horizontal
        mode_row.Margin = Thickness(0, 0, 0, 2)

        self._rb_named = RadioButton()
        self._rb_named.Content = "Named Range"
        self._rb_named.GroupName = "EditRangeMode"
        self._rb_named.Margin = Thickness(0, 0, 16, 0)
        self._rb_named.IsChecked = use_named_default
        self._rb_named.Checked += self._on_range_mode_changed
        mode_row.Children.Add(self._rb_named)

        self._rb_manual = RadioButton()
        self._rb_manual.Content = "Manual Range"
        self._rb_manual.GroupName = "EditRangeMode"
        self._rb_manual.IsChecked = not use_named_default
        self._rb_manual.Checked += self._on_range_mode_changed
        mode_row.Children.Add(self._rb_manual)

        xl_inner.Children.Add(mode_row)

        self._named_combo = ComboBox()
        self._named_combo.Margin = Thickness(0, 4, 0, 2)
        self._named_combo.Padding = Thickness(4, 3, 4, 3)
        self._named_combo.IsEnabled = use_named_default
        self._named_combo.SelectionChanged += self._on_named_range_changed
        # When opening on a record that used a named range, the sheet is
        # driven by that name — disable the sheet picker so it's obvious.
        self._sheet_combo.IsEnabled = not use_named_default
        xl_inner.Children.Add(self._named_combo)

        self._named_hint = TextBlock()
        self._named_hint.Text = ""
        self._named_hint.FontSize = 11
        self._named_hint.Foreground = SolidColorBrush(Color.FromRgb(110, 110, 110))
        self._named_hint.Margin = Thickness(2, 0, 0, 4)
        xl_inner.Children.Add(self._named_hint)

        # Pre-populate the named-range combo from the saved record.
        if self._src_path and self._file_type == "excel":
            try:
                had_any = _populate_named_range_combo(
                    self._named_combo, self._src_path,
                    self._record.get("sheet_name") or None,
                    preferred_name=saved_named or None)
                if not had_any:
                    self._rb_named.IsEnabled = False
                    if self._rb_named.IsChecked:
                        self._rb_manual.IsChecked = True
            except Exception:
                pass

        range_lbl = Label(); range_lbl.Content = "Cell Range  e.g. A1:G25  *"
        range_lbl.Padding = Thickness(0); range_lbl.Margin = Thickness(0, 6, 0, 2)
        xl_inner.Children.Add(range_lbl)
        self._range_box = textbox(self._record.get("range_addr", "A1:G25"))
        self._range_box.IsEnabled = not use_named_default
        xl_inner.Children.Add(self._range_box)

        xl_border.Child = xl_inner
        self._excel_panel.Children.Add(xl_border)
        form.Children.Add(self._excel_panel)

        self._common_panel = StackPanel()
        self._common_panel.Margin = Thickness(0, 4, 0, 0)

        imp_border = System.Windows.Controls.Border()
        imp_border.BorderBrush = SolidColorBrush(Color.FromRgb(190, 210, 200))
        imp_border.BorderThickness = Thickness(1)
        imp_border.CornerRadius = System.Windows.CornerRadius(4)
        imp_border.Padding = Thickness(10, 8, 10, 10)
        imp_border.Background = SolidColorBrush(Color.FromRgb(247, 252, 248))
        imp_border.Margin = Thickness(0, 4, 0, 4)

        imp_inner = StackPanel()

        imp_title = TextBlock()
        imp_title.Text = "Import Options"
        imp_title.FontWeight = FontWeights.Bold
        imp_title.Foreground = SolidColorBrush(Color.FromRgb(0, 110, 70))
        imp_title.Margin = Thickness(0, 0, 0, 6)
        imp_inner.Children.Add(imp_title)

        self._page_panel = StackPanel()
        page_lbl = Label(); page_lbl.Content = "Page Number  (PDF / Word)"
        page_lbl.Padding = Thickness(0); page_lbl.Margin = Thickness(0, 0, 0, 2)
        self._page_panel.Children.Add(page_lbl)
        self._page_combo = ComboBox()
        self._page_combo.Margin = Thickness(0, 0, 0, 2)
        self._page_combo.Padding = Thickness(4, 4, 4, 4)
        saved_page = _safe_int(self._record.get("page_number", "1"), 1)
        # Try to populate with actual page count from source
        self._init_page_combo(saved_page)
        self._page_panel.Children.Add(self._page_combo)
        imp_inner.Children.Add(self._page_panel)

        dpi_lbl = Label(); dpi_lbl.Content = "Import DPI"
        dpi_lbl.Padding = Thickness(0); dpi_lbl.Margin = Thickness(0, 6, 0, 2)
        imp_inner.Children.Add(dpi_lbl)
        self._dpi_combo = ComboBox()
        self._dpi_combo.Margin = Thickness(0, 0, 0, 2)
        self._dpi_combo.Padding = Thickness(4, 4, 4, 4)
        dpi_value = str(self._record.get("dpi", "300"))
        dpi_index = 0
        for i, val in enumerate(DPI_CHOICES):
            item = ComboBoxItem(); item.Content = val
            self._dpi_combo.Items.Add(item)
            if val == dpi_value:
                dpi_index = i
        self._dpi_combo.SelectedIndex = dpi_index
        imp_inner.Children.Add(self._dpi_combo)

        self._transparent_chk = CheckBox()
        self._transparent_chk.Content = "Remove background"
        self._transparent_chk.Margin = Thickness(0, 8, 0, 0)
        self._transparent_chk.IsChecked = _as_bool(self._record.get("transparent", False))
        imp_inner.Children.Add(self._transparent_chk)

        imp_border.Child = imp_inner
        self._common_panel.Children.Add(imp_border)
        form.Children.Add(self._common_panel)

        scroll.Content = form
        Grid.SetRow(scroll, 0)
        outer.Children.Add(scroll)

        btn_bar = Grid()
        btn_bar.Background = SolidColorBrush(Color.FromRgb(240, 244, 250))
        btn_sep = Separator()
        btn_bar.Children.Add(btn_sep)

        btn_inner = StackPanel()
        btn_inner.Orientation = WPFOrientation.Horizontal
        btn_inner.HorizontalAlignment = HorizontalAlignment.Right
        btn_inner.Margin = Thickness(16, 10, 16, 10)

        ok = Button()
        ok.Content = "✔  Apply Changes"
        ok.Width = 150; ok.Height = 36
        ok.Margin = Thickness(0, 0, 8, 0)
        ok.Background = SolidColorBrush(Color.FromRgb(0, 110, 180))
        ok.Foreground = Brushes.White
        ok.FontWeight = FontWeights.Bold
        ok.BorderThickness = Thickness(0)
        ok.FontSize = 13
        ok.Click += self._ok
        btn_inner.Children.Add(ok)

        cancel = Button()
        cancel.Content = "Cancel"
        cancel.Width = 86; cancel.Height = 36
        cancel.FontSize = 13
        cancel.Click += self._cancel
        btn_inner.Children.Add(cancel)

        btn_bar.Children.Add(btn_inner)
        Grid.SetRow(btn_bar, 1)
        outer.Children.Add(btn_bar)

        self.Content = outer
        self._set_type_ui(self._file_type)

    def _on_sheet_selection_changed(self, s, e):
        """When the user picks a sheet, refresh the print-area textbox and
        the Name-Manager combo for that sheet's visible scope."""
        if self._file_type != "excel" or not self._src_path:
            return

        item = self._sheet_combo.SelectedItem
        sheet_name = None
        if item and hasattr(item, "Tag"):
            sheet_name = item.Tag
        else:
            sheet_name = self._sheet_combo.Text.strip()

        # Refresh defined-name list for the new sheet scope.
        try:
            had_any = _populate_named_range_combo(
                self._named_combo, self._src_path, sheet_name)
            if not had_any:
                self._named_combo.IsEnabled = False
                self._rb_named.IsEnabled = False
                if self._rb_named.IsChecked:
                    self._rb_manual.IsChecked = True
            else:
                self._rb_named.IsEnabled = True
                self._named_combo.IsEnabled = bool(self._rb_named.IsChecked)
        except Exception:
            pass

        if sheet_name:
            # 1. Try cached OpenXML data first
            sh_info = self._sheet_data.get(sheet_name)
            if sh_info and sh_info.get("print_area"):
                pa = sh_info["print_area"]
                if not self._rb_named.IsChecked:
                    self._range_box.Text = pa
                self._detect_lbl.Text = "✔  Print area detected (from file): {}".format(pa)
                return

            # 2. Fallback to COM
            try:
                print_area = get_excel_print_area(self._src_path, sheet_name)
                if print_area and not self._rb_named.IsChecked:
                    self._range_box.Text = print_area
                    self._detect_lbl.Text = "✔  Range detected: {}".format(print_area)
            except Exception:
                pass

    def _on_range_mode_changed(self, s, e):
        if not hasattr(self, "_rb_named"):
            return
        use_named = bool(self._rb_named.IsChecked)
        has_items = self._named_combo.Items.Count > 0
        self._named_combo.IsEnabled = use_named and has_items
        # Same coupling as AddImportDialog: name drives both sheet and range.
        self._range_box.IsEnabled = not use_named
        self._sheet_combo.IsEnabled = not use_named
        if use_named and has_items:
            self._on_named_range_changed(None, None)
        else:
            self._named_hint.Text = ""

    def _on_named_range_changed(self, s, e):
        item = self._named_combo.SelectedItem
        if item is None or not hasattr(item, "Tag") or item.Tag is None:
            self._named_hint.Text = ""
            return
        info = item.Tag
        rng  = info.get("range_address") or ""
        sheet = info.get("ref_sheet") or ""
        if rng:
            self._range_box.Text = rng
            scope = "workbook-scoped" if info.get("is_workbook_scope") else "sheet-scoped"
            self._named_hint.Text = "→ {}!{}  ({})".format(sheet, rng, scope)

    def _set_type_ui(self, ftype):
        self._file_type = ftype
        self._excel_panel.Visibility = Visibility.Visible if ftype == "excel" else Visibility.Collapsed
        self._common_panel.Visibility = Visibility.Visible if ftype in ("excel", "pdf", "word", "image") else Visibility.Collapsed
        self._page_panel.Visibility = Visibility.Visible if ftype in ("pdf", "word") else Visibility.Collapsed
        self._transparent_chk.Visibility = Visibility.Visible if ftype in ("excel", "word", "image", "pdf") else Visibility.Collapsed

    def _browse(self, s, e):
        dlg = System.Windows.Forms.OpenFileDialog()
        dlg.Title  = "Select Source File"
        dlg.Filter = SUPPORTED_FILES_FILTER
        if dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK:
            return
        path = dlg.FileName
        self._src_path = path
        self._src_box.Text = path
        ftype = detect_file_type(path)
        self._set_type_ui(ftype)

        if ftype == "excel":
            self._detect_lbl.Text = "✔  Excel workbook detected"
            # Populate sheets
            sheets = None
            try:
                # Try OpenXML parsing first (works for all Excel formats)
                sheets = list_sheets(path)
            except Exception as ex:
                # Fallback to COM for macro-enabled files or if OpenXML fails
                sheets = get_excel_sheets_com(path)
                if sheets:
                    print("[DocLinkManager] OpenXML parsing failed, using COM fallback: {}".format(ex))
            
            if sheets:
                self._sheet_data = {sh['name']: sh for sh in sheets}
                self._sheet_combo.Items.Clear()
                auto_sel_idx = -1
                for i, sh in enumerate(sheets):
                    item = ComboBoxItem()
                    marker = "  [Print Area]" if sh.get("has_print_area") else ""
                    item.Content = "{}{}".format(sh['name'], marker)
                    item.Tag = sh['name']
                    self._sheet_combo.Items.Add(item)
                    if sh.get('name') == self._record.get("sheet_name"):
                        auto_sel_idx = i

                if auto_sel_idx >= 0:
                    self._sheet_combo.SelectedIndex = auto_sel_idx
                elif self._sheet_combo.Items.Count > 0:
                    self._sheet_combo.SelectedIndex = 0
                # Browsing to a new file: re-evaluate which range modes apply.
                # If no defined names exist, force Manual mode.
                if self._named_combo.Items.Count == 0:
                    _populate_named_range_combo(
                        self._named_combo, self._src_path, None)
                if self._named_combo.Items.Count == 0:
                    self._rb_named.IsEnabled = False
                    if self._rb_named.IsChecked:
                        self._rb_manual.IsChecked = True
                else:
                    self._rb_named.IsEnabled = True
            else:
                self._detect_lbl.Text = "⚠  Could not list sheets from this workbook"
        elif ftype == "word":
            msg = "✔  Word document detected"
            if not _WORD_AVAILABLE:
                msg += "  (Word Interop not available on this machine)"
            else:
                count = get_word_page_count(path)
                if count > 0:
                    self._populate_page_combo(count)
                    msg += "  ({} page{})".format(count, "s" if count > 1 else "")
            self._detect_lbl.Text = msg
        elif ftype == "pdf":
            count = get_pdf_page_count(path)
            if count > 0:
                self._populate_page_combo(count)
                self._detect_lbl.Text = "✔  PDF detected  ({} page{})".format(
                    count, "s" if count > 1 else "")
            else:
                self._detect_lbl.Text = "✔  PDF detected  (via PyMuPDF)"
        else:
            ext = os.path.splitext(path)[1].lower()
            self._detect_lbl.Text = "✔  Image detected ({})".format(ext)

    def _init_page_combo(self, saved_page):
        """Populate the page ComboBox on dialog open from the existing source file."""
        total = 0
        src = self._src_path
        if src and os.path.exists(src):
            ftype = self._file_type
            if ftype == "pdf":
                total = get_pdf_page_count(src)
            elif ftype == "word" and _WORD_AVAILABLE:
                total = get_word_page_count(src)
        if total > 0:
            self._populate_page_combo(total, saved_page)
        else:
            # Fallback: just show the saved page number
            self._page_combo.Items.Clear()
            ci = ComboBoxItem(); ci.Content = str(saved_page)
            self._page_combo.Items.Add(ci)
            self._page_combo.SelectedIndex = 0

    def _populate_page_combo(self, total_pages, selected_page=1):
        """Fill the page ComboBox with page numbers 1..total_pages."""
        self._page_combo.Items.Clear()
        sel_idx = 0
        for p in range(1, total_pages + 1):
            ci = ComboBoxItem(); ci.Content = str(p)
            self._page_combo.Items.Add(ci)
            if p == selected_page:
                sel_idx = p - 1
        self._page_combo.SelectedIndex = min(sel_idx, max(0, total_pages - 1))

    def _ok(self, s, e):
        if not self._src_path:
            forms.alert("Browse to a source file."); return
        if self._file_type == "excel" and not self._range_box.Text.strip():
            forms.alert("Enter a cell range."); return

        sel = self._path_combo.SelectedItem
        path_type = sel.Content if sel else "Absolute"

        page_sel = self._page_combo.SelectedItem
        page_num = page_sel.Content if page_sel else "1"

        # Extract clean sheet name from Tag if available
        sheet_name = ""
        named_range = ""
        if self._file_type == "excel":
            sel_item = self._sheet_combo.SelectedItem
            if sel_item and hasattr(sel_item, "Tag"):
                sheet_name = sel_item.Tag
            else:
                sheet_name = self._sheet_combo.Text.strip()
            if self._rb_named.IsChecked:
                nm_item = self._named_combo.SelectedItem
                if nm_item is None or not hasattr(nm_item, "Tag") or nm_item.Tag is None:
                    forms.alert("Pick a Named Range or switch to Manual Range.")
                    return
                info = nm_item.Tag
                named_range = info.get("name", "")
                if not info.get("is_workbook_scope"):
                    sheet_name = info.get("scope_sheet") or sheet_name

        self.result = {
            "name":         self._name_box.Text.strip(),
            "source_path":  self._src_path,
            "path_type":    path_type,
            "file_type":    self._file_type,
            "sheet_name":   sheet_name,
            "range_addr":   self._range_box.Text.strip() if self._file_type == "excel" else "",
            "named_range":  named_range,
            "page_number":  page_num if self._file_type in ("pdf", "word") else "1",
            "dpi":          _combo_selected_text(self._dpi_combo, "300"),
            "transparent":  _as_bool(self._transparent_chk.IsChecked),
        }
        self.DialogResult = True
        self.Close()

    def _cancel(self, s, e):
        self.DialogResult = False
        self.Close()

class InstanceHistoryWindow(Window):
    """
    Shows the full instance lifecycle log for one doclink record.
    Open by double-clicking a row in the main grid.
    """

    # Colour-coding per action type
    _ACTION_COLORS = {
        "Created":            Color.FromRgb(200, 240, 210),   # soft green
        "Placed":             Color.FromRgb(200, 235, 255),   # soft blue
        "Replaced":           Color.FromRgb(255, 245, 200),   # soft amber
        "Edited":             Color.FromRgb(230, 215, 255),   # soft purple
        "Deleted":            Color.FromRgb(255, 210, 210),   # soft red
        "Externally Deleted": Color.FromRgb(255, 180, 180),   # stronger red
    }

    def __init__(self, record):
        self._record = record
        self._setup()

    def _setup(self):
        self.Title  = "Instance History  –  {}".format(self._record.get("name", "?"))
        self.Width  = 720
        self.Height = 480
        self.MinWidth  = 500
        self.MinHeight = 300
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
        apply_template_to_window(self, _TMPL_PATH)
        if _RES: self.Background = _RES["Brush.Window.Background"]

        root = Grid()
        r0 = RowDefinition(); r0.Height = GridLength.Auto
        r1 = RowDefinition(); r1.Height = GridLength(1, GridUnitType.Star)
        r2 = RowDefinition(); r2.Height = GridLength.Auto
        root.RowDefinitions.Add(r0)
        root.RowDefinitions.Add(r1)
        root.RowDefinitions.Add(r2)

        # ── Summary header ────────────────────────────────────────────────────
        header = StackPanel()
        header.Background = SolidColorBrush(Color.FromRgb(28, 48, 80))
        header.Margin = Thickness(0)
        header.Orientation = WPFOrientation.Horizontal

        def hdr_lbl(text, bold=False, width=None):
            lb = Label()
            lb.Content = text
            lb.Foreground = Brushes.White
            lb.VerticalAlignment = VerticalAlignment.Center
            lb.Padding = Thickness(12, 8, 12, 8)
            lb.FontSize = 13
            if bold:
                lb.FontWeight = FontWeights.Bold
            if width:
                lb.Width = width
            return lb

        ic = self._record.get("instance_count", 0)
        dc = self._record.get("deleted_count",  0)
        header.Children.Add(hdr_lbl(self._record.get("name", ""), bold=True))
        header.Children.Add(hdr_lbl("│"))
        header.Children.Add(hdr_lbl("Instances placed: {}".format(ic)))
        header.Children.Add(hdr_lbl("│"))
        header.Children.Add(hdr_lbl("Deleted: {}".format(dc)))
        header.Children.Add(hdr_lbl("│"))
        header.Children.Add(hdr_lbl("Type: {}".format(self._record.get("file_type","").upper())))

        Grid.SetRow(header, 0)
        root.Children.Add(header)

        # ── History list (scrollable) ─────────────────────────────────────────
        scroll = ScrollViewer()
        scroll.Margin = Thickness(8, 6, 8, 4)
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto

        history_panel = StackPanel()

        history = self._record.get("instance_history", [])
        if not history:
            no_data = TextBlock()
            no_data.Text = "No history recorded yet."
            no_data.Margin = Thickness(16, 20, 0, 0)
            no_data.Foreground = SolidColorBrush(Color.FromRgb(120, 120, 120))
            no_data.FontSize = 13
            history_panel.Children.Add(no_data)
        else:
            for idx_h, entry in enumerate(reversed(history)):
                action    = entry.get("action", "?")
                timestamp = entry.get("timestamp", "—")
                elem_id   = entry.get("element_unique_id") or entry.get("element_id")
                user      = entry.get("user", "")
                detail    = entry.get("detail", "")

                bg_color = self._ACTION_COLORS.get(action, Color.FromRgb(240, 240, 240))

                row_border = System.Windows.Controls.Border()
                row_border.Background = SolidColorBrush(bg_color)
                row_border.BorderBrush = SolidColorBrush(Color.FromRgb(200, 210, 225))
                row_border.BorderThickness = Thickness(0, 0, 0, 1)
                row_border.Padding = Thickness(12, 8, 12, 8)

                row_grid = Grid()
                c0 = ColumnDefinition(); c0.Width = GridLength(150)
                c1 = ColumnDefinition(); c1.Width = GridLength(110)
                c2 = ColumnDefinition(); c2.Width = GridLength(120)
                c3 = ColumnDefinition(); c3.Width = GridLength(1, GridUnitType.Star)
                row_grid.ColumnDefinitions.Add(c0)
                row_grid.ColumnDefinitions.Add(c1)
                row_grid.ColumnDefinitions.Add(c2)
                row_grid.ColumnDefinitions.Add(c3)

                def cell(text, col_idx, bold=False, color=None):
                    tb = TextBlock()
                    tb.Text = str(text) if text is not None else "—"
                    tb.VerticalAlignment = VerticalAlignment.Center
                    tb.FontSize = 12
                    if bold:
                        tb.FontWeight = FontWeights.Bold
                    if color:
                        tb.Foreground = SolidColorBrush(color)
                    Grid.SetColumn(tb, col_idx)
                    return tb

                is_delete = action in ("Deleted", "Externally Deleted")
                action_color = (Color.FromRgb(180, 0, 0) if is_delete
                                else Color.FromRgb(0, 100, 0) if action in ("Placed", "Created")
                                else Color.FromRgb(40, 40, 120))

                row_grid.Children.Add(cell(timestamp,   0))
                row_grid.Children.Add(cell(action,      1, color=action_color))
                short_id = (elem_id[:8] + "…" + elem_id[-8:]) if elem_id and len(elem_id) > 18 else elem_id
                row_grid.Children.Add(cell("UID: {}".format(short_id) if short_id else "—", 2,
                                           color=Color.FromRgb(100, 100, 100)))
                row_grid.Children.Add(cell(detail,      3))

                row_border.Child = row_grid
                history_panel.Children.Add(row_border)

        scroll.Content = history_panel
        Grid.SetRow(scroll, 1)
        root.Children.Add(scroll)

        # ── Column headers strip ──────────────────────────────────────────────
        # Insert column headers just above the list by rearranging grid rows
        col_hdr = Grid()
        col_hdr.Background = SolidColorBrush(Color.FromRgb(210, 220, 240))
        col_hdr.Margin = Thickness(8, 0, 8, 0)

        ch0 = ColumnDefinition(); ch0.Width = GridLength(150)
        ch1 = ColumnDefinition(); ch1.Width = GridLength(110)
        ch2 = ColumnDefinition(); ch2.Width = GridLength(120)
        ch3 = ColumnDefinition(); ch3.Width = GridLength(1, GridUnitType.Star)
        col_hdr.ColumnDefinitions.Add(ch0)
        col_hdr.ColumnDefinitions.Add(ch1)
        col_hdr.ColumnDefinitions.Add(ch2)
        col_hdr.ColumnDefinitions.Add(ch3)

        def ch_lbl(text, col_idx):
            lb = Label()
            lb.Content = text
            lb.FontWeight = FontWeights.Bold
            lb.FontSize = 11
            lb.Foreground = SolidColorBrush(Color.FromRgb(40, 60, 110))
            lb.Padding = Thickness(12, 4, 4, 4)
            Grid.SetColumn(lb, col_idx)
            return lb

        col_hdr.Children.Add(ch_lbl("Timestamp",   0))
        col_hdr.Children.Add(ch_lbl("Action",       1))
        col_hdr.Children.Add(ch_lbl("Element UID",  2))
        col_hdr.Children.Add(ch_lbl("Detail",       3))

        # We need a 4-row root grid to fit the column headers between header and list
        # Rebuild root grid with 4 rows
        root4 = Grid()
        rr0 = RowDefinition(); rr0.Height = GridLength.Auto
        rr1 = RowDefinition(); rr1.Height = GridLength.Auto
        rr2 = RowDefinition(); rr2.Height = GridLength(1, GridUnitType.Star)
        rr3 = RowDefinition(); rr3.Height = GridLength.Auto
        root4.RowDefinitions.Add(rr0)
        root4.RowDefinitions.Add(rr1)
        root4.RowDefinitions.Add(rr2)
        root4.RowDefinitions.Add(rr3)

        Grid.SetRow(header, 0)
        root4.Children.Add(header)

        Grid.SetRow(col_hdr, 1)
        root4.Children.Add(col_hdr)

        Grid.SetRow(scroll, 2)
        root4.Children.Add(scroll)

        # ── Close button ──────────────────────────────────────────────────────
        close_bar = StackPanel()
        close_bar.Orientation = WPFOrientation.Horizontal
        close_bar.HorizontalAlignment = HorizontalAlignment.Right
        close_bar.Margin = Thickness(0, 6, 12, 10)

        close_btn = Button()
        close_btn.Content = "Close"
        close_btn.Width = 90
        close_btn.Height = 32
        close_btn.FontSize = 13
        close_btn.Click += lambda s, e: self.Close()
        close_bar.Children.Add(close_btn)

        Grid.SetRow(close_bar, 3)
        root4.Children.Add(close_bar)

        self.Content = root4


# ─────────────────────────────────────────────────────────────────────────────
# MAIN WINDOW
# ─────────────────────────────────────────────────────────────────────────────
