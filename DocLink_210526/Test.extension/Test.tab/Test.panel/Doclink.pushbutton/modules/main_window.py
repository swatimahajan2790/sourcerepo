# -*- coding: utf-8 -*-
"""
main_window.py
--------------
DocLink main two-tab window plus the tool entry point.

Classes
-------
DocLinkManagerWindow  – Tab 1 (DocLink Images) + Tab 2 (Schedule Import) manager

Functions
---------
main()  – called by pyRevit engine via script.py
"""

import os, sys, re, json, datetime, traceback

from _imports import (
    System, Window, Thickness, HorizontalAlignment, VerticalAlignment,
    Visibility, GridLength, GridUnitType, FontWeights,
    Grid, RowDefinition, ColumnDefinition, StackPanel, WrapPanel, ScrollViewer,
    Button, Label, TextBox, TextBlock, ComboBox, ComboBoxItem, CheckBox,
    Separator, DataGrid, DataGridTextColumn, DataGridSelectionMode,
    WPFOrientation, SelectionChangedEventArgs, TabControl, TabItem, ProgressBar,
    Binding, SolidColorBrush, Color, Brushes, ObservableCollection,
    INotifyPropertyChanged, PropertyChangedEventArgs, SystemParameters,
    Dispatcher, DispatcherPriority, DispatcherFrame,
    WinForms, WFOpenFileDialog, WFDialogResult,
    apply_template_to_window, _TMPL_PATH, _RES,
    SystemParameters_WorkArea_Height,
    Transaction, TransactionGroup,
    ImageInstance, ElementId, FilteredElementCollector,
    BuiltInParameter, BuiltInCategory,
    ViewSheet, ScheduleSheetInstance, XYZ,
    TaskDialog, TaskDialogCommonButtons,
    forms, revit,
    _FITZ_AVAILABLE, _CPYTHON,
    _EXCEL_AVAILABLE, _WORD_AVAILABLE,
)
from utils import (
    _safe_int, _as_bool, _combo_selected_text,
    detect_file_type, SUPPORTED_FILES_FILTER, DPI_CHOICES,
    describe_record_options, _sanitize_import_name, _get_auto_user,
    _parse_page_numbers, _uid
)
from models import DocLinkRow, ScheduleRow
from persistence import (
    load_records, save_records,
    load_schedule_records, save_schedule_records,
    _normalize_loaded_records,
)
from excel_ops import (
    export_excel_range_to_pdf, capture_excel_range_as_image,
    get_excel_print_area,
)
from word_ops import (
    export_word_to_pdf, get_word_page_count, capture_word_page_as_image,
)
from pdf_ops import (
    convert_pdf_page_to_png, get_pdf_page_count,
    make_white_background_transparent, crop_white_margins, remove_background,
)
from revit_ops import (
    place_or_replace, _element_by_unique_id, _get_view_center,
    _get_image_center_point, _all_instances_by_type,
    _normalize_unique_ids, _record_element_unique_ids,
    _set_record_element_unique_ids, _find_instances_doc_wide,
    _get_element_name, _image_type_unique_id_from_instance,
    _resolve_existing_instance_unique_ids
)
from dialogs import AddImportDialog, EditImportDialog, InstanceHistoryWindow
from schedule_tab import (
    ScheduleSetupDialog, run_document_importer,
    _build_schedule_from_record, _extract_column_width_snapshot,
    _get_active_sheet, _resize_schedule_in_place
)
from logger import LogManager

class DocLinkManagerWindow(Window):

    def __init__(self, doc, view, uidoc=None):
        self._doc     = doc
        self._uidoc   = uidoc
        self._view    = view
        self._records = load_records(doc)
        self._rows    = ObservableCollection[DocLinkRow]()
        # Tab 2: Schedule Import
        self._sched_records = load_schedule_records(doc)
        self._sched_rows    = ObservableCollection[ScheduleRow]()
        self._setup_ui()
        self._refresh()
        self._sched_refresh()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        view_name = self._view.Name[:50] if self._view else "No Active View"
        self.Title  = "DocLink Manager  –  {}".format(view_name)
        self.Width  = 1060
        self.Height = 660
        self.MinWidth  = 800
        self.MinHeight = 460
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        apply_template_to_window(self, _TMPL_PATH)
        if _RES: self.Background = _RES["Brush.Window.Background"]

        root = Grid()
        r0 = RowDefinition(); r0.Height = GridLength(1, GridUnitType.Star)
        r1 = RowDefinition(); r1.Height = GridLength(38)
        root.RowDefinitions.Add(r0)
        root.RowDefinitions.Add(r1)

        # ══ TabControl ════════════════════════════════════════════════════════
        tabs = TabControl()
        tabs.Margin = Thickness(0)

        # ── TAB 1: DocLink Images ──────────────────────────────────────────────
        tab1 = TabItem()
        tab1.Header = "  DocLink Images  "
        tab1.FontSize = 13
        tab1.FontWeight = FontWeights.Bold
        tab1_grid = self._build_tab1_content(view_name)
        tab1.Content = tab1_grid
        tabs.Items.Add(tab1)

        # ── TAB 2: Schedule Import ────────────────────────────────────────────
        tab2 = TabItem()
        tab2.Header = "  Schedule Import  "
        tab2.FontSize = 13
        tab2.FontWeight = FontWeights.Bold
        tab2_grid = self._build_tab2_content()
        tab2.Content = tab2_grid
        tabs.Items.Add(tab2)

        Grid.SetRow(tabs, 0)
        root.Children.Add(tabs)

        # ── Shared status bar ─────────────────────────────────────────────────
        sb = StackPanel()
        sb.Orientation = WPFOrientation.Horizontal
        if _RES:
            sb.Background = _RES["Brush.Panel.Secondary"]
        else:
            sb.Background = SolidColorBrush(Color.FromRgb(230, 236, 248))

        self._status = Label()
        self._status.Content = "Ready"
        self._status.VerticalAlignment = VerticalAlignment.Center
        if _RES:
            self._status.Style = _RES["Style.Label.Standard"]
        else:
            self._status.FontSize  = 12
            self._status.Foreground = SolidColorBrush(Color.FromRgb(40, 70, 130))
        sb.Children.Add(self._status)

        self._progress = ProgressBar()
        self._progress.Minimum = 0
        self._progress.Maximum = 100
        self._progress.Value   = 0
        self._progress.Width   = 180
        self._progress.Height  = 14
        self._progress.Margin  = Thickness(6, 0, 6, 0)
        self._progress.VerticalAlignment = VerticalAlignment.Center
        self._progress.Visibility = Visibility.Collapsed
        sb.Children.Add(self._progress)

        Grid.SetRow(sb, 1)
        root.Children.Add(sb)

        self.Content = root

    # ── Tab 1 builder ─────────────────────────────────────────────────────────

    def _build_tab1_content(self, view_name):
        tab1 = Grid()
        for h in [GridLength(1, GridUnitType.Auto), GridLength(1, GridUnitType.Star)]:
            rd = RowDefinition(); rd.Height = h
            tab1.RowDefinitions.Add(rd)

        # toolbar
        bar = WrapPanel()
        bar.Orientation = WPFOrientation.Horizontal
        bar.Background  = SolidColorBrush(Color.FromRgb(234, 238, 245))
        bar.Margin = Thickness(8, 6, 8, 0)

        def tbtn(text, handler, tip=""):
            b = Button()
            b.Content = text
            b.Margin  = Thickness(0, 0, 8, 8)
            b.Padding = Thickness(16, 7, 16, 7)
            b.FontWeight = FontWeights.Bold
            b.Background = SolidColorBrush(Color.FromRgb(76, 96, 122))
            b.Foreground = Brushes.White
            b.BorderThickness = Thickness(0)
            b.Click += handler
            if tip:
                b.ToolTip = tip
            return b

        bar.Children.Add(tbtn("+ Add Import", self._on_add,
                               "Browse Excel / Word / PDF / Image and place into active view."))
        bar.Children.Add(tbtn("✎  Edit", self._on_edit,
                               "Edit source path or cell range of the selected import."))
        bar.Children.Add(tbtn("⟳  Update Selected", self._on_update_sel,
                               "Re-export and re-place the selected row(s)."))
        bar.Children.Add(tbtn("⟳  Update All", self._on_update_all,
                               "Re-export and re-place every import in this view."))
        bar.Children.Add(tbtn("⟳  Update Entire Project", self._on_update_project_all,
                               "Refresh tracked stickies from all views and all tracked schedules in one run."))
        bar.Children.Add(tbtn("✕  Remove", self._on_remove,
                               "Delete selected records and their Revit elements."))

        lbl = Label()
        lbl.Content   = "  View: {}".format(view_name)
        lbl.Foreground = SolidColorBrush(Color.FromRgb(160, 190, 230))
        lbl.VerticalAlignment = VerticalAlignment.Center
        bar.Children.Add(lbl)

        Grid.SetRow(bar, 0)
        tab1.Children.Add(bar)

        # DataGrid
        scroll = ScrollViewer()
        scroll.Margin = Thickness(8, 6, 8, 0)

        self._dg = DataGrid()
        self._dg.AutoGenerateColumns = False
        self._dg.CanUserAddRows    = False
        self._dg.CanUserDeleteRows = False
        self._dg.IsReadOnly        = True
        self._dg.SelectionMode     = DataGridSelectionMode.Extended
        self._dg.RowHeight = 36
        self._dg.FontSize  = 13
        self._dg.AlternatingRowBackground = SolidColorBrush(Color.FromRgb(245, 248, 255))
        self._dg.GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.Horizontal
        self._dg.HorizontalGridLinesBrush = SolidColorBrush(Color.FromRgb(210, 218, 232))

        def col(header, prop, width=None, star=False):
            c = DataGridTextColumn()
            c.Header  = header
            c.Binding = Binding(prop)
            DGL = System.Windows.Controls.DataGridLength
            DGLU = System.Windows.Controls.DataGridLengthUnitType
            if star:
                c.Width = DGL(1, DGLU.Star)
            elif width:
                c.Width = DGL(width)
            else:
                c.Width = DGL(1, DGLU.Star)
            return c

        self._dg.Columns.Add(col("#",            "idx",            40))
        self._dg.Columns.Add(col("Import Name",  "name",           star=True))
        self._dg.Columns.Add(col("Type",         "file_type",      70))
        self._dg.Columns.Add(col("Created By",   "user",           120))
        self._dg.Columns.Add(col("View",         "view_name",      140))
        self._dg.Columns.Add(col("Path Type",    "path_type",      90))
        self._dg.Columns.Add(col("Options",      "range_info",     180))
        self._dg.Columns.Add(col("Instances",    "instance_count", 80))
        self._dg.Columns.Add(col("Last Updated", "last_updated",   140))
        self._dg.Columns.Add(col("Status",       "status",         80))

        self._dg.MouseDoubleClick += self._on_row_double_click
        self._dg.ItemsSource = self._rows

        scroll.Content = self._dg
        Grid.SetRow(scroll, 1)
        tab1.Children.Add(scroll)

        return tab1

    # ── Tab 2 builder ─────────────────────────────────────────────────────────

    def _build_tab2_content(self):
        tab2 = Grid()
        for h in [GridLength(1, GridUnitType.Auto), GridLength(1, GridUnitType.Star)]:
            rd = RowDefinition(); rd.Height = h
            tab2.RowDefinitions.Add(rd)

        # toolbar
        bar = WrapPanel()
        bar.Orientation = WPFOrientation.Horizontal
        bar.Background  = SolidColorBrush(Color.FromRgb(234, 238, 245))
        bar.Margin = Thickness(8, 6, 8, 0)

        def tbtn(text, handler, tip=""):
            b = Button()
            b.Content = text
            b.Margin  = Thickness(0, 0, 8, 8)
            b.Padding = Thickness(16, 7, 16, 7)
            b.FontWeight = FontWeights.Bold
            b.Background = SolidColorBrush(Color.FromRgb(76, 96, 122))
            b.Foreground = Brushes.White
            b.BorderThickness = Thickness(0)
            b.Click += handler
            if tip:
                b.ToolTip = tip
            return b

        bar.Children.Add(tbtn("+ Import Excel → Schedule", self._on_sched_add,
                               "Pick an Excel file, choose sheet and options, create a Revit schedule."))
        bar.Children.Add(tbtn("✏  Edit Setup", self._on_sched_edit,
                               "Edit the source file, worksheet, and import options for the selected schedule."))
        bar.Children.Add(tbtn("⟳  Update Selected", self._on_sched_update,
                               "Delete and recreate the selected schedule(s) from their source Excel."))
        bar.Children.Add(tbtn("⟳  Update All", self._on_sched_update_all,
                               "Delete and recreate every tracked schedule from its saved Excel setup."))
        bar.Children.Add(tbtn("⟳  Update Entire Project", self._on_update_project_all,
                               "Refresh tracked stickies from all views and all tracked schedules in one run."))
        bar.Children.Add(tbtn("✕  Remove", self._on_sched_remove,
                               "Delete the selected schedule(s) from Revit and remove tracking."))

        Grid.SetRow(bar, 0)
        tab2.Children.Add(bar)

        # DataGrid
        scroll = ScrollViewer()
        scroll.Margin = Thickness(8, 6, 8, 0)

        self._sched_dg = DataGrid()
        self._sched_dg.AutoGenerateColumns = False
        self._sched_dg.CanUserAddRows    = False
        self._sched_dg.CanUserDeleteRows = False
        self._sched_dg.IsReadOnly        = False
        self._sched_dg.SelectionMode     = DataGridSelectionMode.Extended
        self._sched_dg.RowHeight = 36
        self._sched_dg.FontSize  = 13
        self._sched_dg.AlternatingRowBackground = SolidColorBrush(Color.FromRgb(250, 245, 255))
        self._sched_dg.GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.Horizontal
        self._sched_dg.HorizontalGridLinesBrush = SolidColorBrush(Color.FromRgb(220, 210, 232))

        def col(header, prop, width=None, star=False, readonly=True):
            c = DataGridTextColumn()
            c.Header  = header
            c.Binding = Binding(prop)
            c.IsReadOnly = readonly
            DGL = System.Windows.Controls.DataGridLength
            DGLU = System.Windows.Controls.DataGridLengthUnitType
            if star:
                c.Width = DGL(1, DGLU.Star)
            elif width:
                c.Width = DGL(width)
            else:
                c.Width = DGL(1, DGLU.Star)
            return c

        def chk_col(header, prop, width=60):
            from System.Windows.Controls import DataGridCheckBoxColumn
            c = DataGridCheckBoxColumn()
            c.Header = header
            c.Binding = Binding(prop)
            c.Width = System.Windows.Controls.DataGridLength(width)
            return c

        self._sched_dg.Columns.Add(col("#",            "idx",          40))
        self._sched_dg.Columns.Add(chk_col("Retain",   "retain_settings"))
        self._sched_dg.Columns.Add(col("Schedule Name","name",         star=True))
        self._sched_dg.Columns.Add(col("Source File",  "source_file",  220))
        self._sched_dg.Columns.Add(col("Sheet",        "sheet_name",   120))
        self._sched_dg.Columns.Add(col("Options",      "options_info", 180))
        self._sched_dg.Columns.Add(col("Last Updated", "last_updated", 155))
        self._sched_dg.Columns.Add(col("Status",       "status",       90))

        self._sched_dg.ItemsSource = self._sched_rows
        self._sched_selected_ids_cache = set()
        self._sched_dg.SelectionChanged += self._on_sched_selection_changed

        scroll.Content = self._sched_dg
        Grid.SetRow(scroll, 1)
        tab2.Children.Add(scroll)

        return tab2

    # ── helpers ───────────────────────────────────────────────────────────────

    def _bring_to_focus(self):
        """Bring this window to focus after any process or error dialog."""
        try:
            # Schedule on UI dispatcher to ensure focus happens after alert closes
            def focus_on_dispatcher():
                try:
                    # Ensure window is visible and not minimized
                    if self.WindowState == System.Windows.WindowState.Minimized:
                        self.WindowState = System.Windows.WindowState.Normal
                    
                    self.Show()
                    # Force window to top temporarily, then restore normal state
                    self.Topmost = True
                    self.Activate()
                    self.Focus()
                    self.Topmost = False
                except Exception:
                    pass
            
            # Use dispatcher to defer focus call  
            self.Dispatcher.BeginInvoke(
                System.Action(focus_on_dispatcher),
                DispatcherPriority.Normal
            )
        except Exception:
            pass

    def _view_id(self):
        return self._view.Id.IntegerValue if self._view else -1

    def _view_records(self):
        vid = self._view_id()
        return [r for r in self._records
                if _safe_int(r.get("view_id"), -999) == vid]

    def _resolve_view_name(self, view_id_int):
        """Return the view name for a given ElementId integer, or '—'."""
        try:
            el = self._doc.GetElement(ElementId(view_id_int))
            if el is not None:
                return el.Name
        except Exception:
            pass
        return "—"

    def _resolve_record_view(self, record):
        """Return the Revit view element for a record, falling back to active view."""
        vid = _safe_int(record.get("view_id"), -1)
        if vid > 0:
            try:
                el = self._doc.GetElement(ElementId(vid))
                if el is not None:
                    return el
            except Exception:
                pass
        return self._view

    # ── progress bar helpers ────────────────────────────────────────────────

    def _progress_start(self, total):
        """Show the progress bar and set its range."""
        self._progress.Value = 0
        self._progress.Maximum = max(1, total)
        self._progress.Visibility = Visibility.Visible
        self._flush_ui()

    def _progress_update(self, current, message=""):
        """Advance the progress bar and optionally update the status text."""
        self._progress.Value = current
        if message:
            self._status.Content = message
        self._flush_ui()

    def _progress_end(self):
        """Hide the progress bar."""
        self._progress.Visibility = Visibility.Collapsed
        self._flush_ui()

    @staticmethod
    def _flush_ui():
        """Pump the WPF dispatcher so the UI repaints during a synchronous loop."""
        try:
            frame = DispatcherFrame()
            Dispatcher.CurrentDispatcher.BeginInvoke(
                DispatcherPriority.Background,
                System.Action(lambda: setattr(frame, 'Continue', False))
            )
            Dispatcher.PushFrame(frame)
        except Exception:
            pass

    def _refresh(self):
        self._rows.Clear()
        # Scan for any externally deleted elements across ALL records
        changed = False
        for r in self._records:
            if self._check_external_deletion(r):
                changed = True
        if changed:
            save_records(self._doc, self._records)

        # Build a view-name cache to avoid repeated lookups
        vid_cache = {}
        current_vid = self._view_id()
        for i, r in enumerate(self._records):
            vid = _safe_int(r.get("view_id"), -999)
            if vid not in vid_cache:
                vid_cache[vid] = self._resolve_view_name(vid)
            vname = vid_cache[vid]
            if vid == current_vid:
                vname = vname + " *"   # mark active view
            ri = describe_record_options(r)
            row = DocLinkRow(
                record_id      = r.get("id", ""),
                idx            = i + 1,
                name           = r.get("name", ""),
                file_type      = r.get("file_type", ""),
                user           = r.get("user", ""),
                path_type      = r.get("path_type", "Absolute"),
                range_info     = ri,
                last_updated   = r.get("last_updated", "—"),
                status         = r.get("status", "—"),
                instance_count = r.get("instance_count", 0),
                deleted_count  = r.get("deleted_count", 0),
                view_name      = vname,
            )
            self._rows.Add(row)
        view_count = len([r for r in self._records
                          if _safe_int(r.get("view_id"), -999) == current_vid])
        self._status.Content = "{} import(s) total  |  {} in active view.".format(
            len(self._rows), view_count)

    def _resolve_path(self, record):
        src = record.get("source_path", "")
        if record.get("path_type") == "Relative":
            rvt = self._doc.PathName
            if rvt:
                src = os.path.join(os.path.dirname(rvt), src)
        return src

    # ── instance tracking ─────────────────────────────────────────────────────

    @staticmethod
    def _now():
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _log_event(self, record, action, element_unique_id=None, detail=""):
        """Append one entry to record['instance_history']."""
        if "instance_history" not in record:
            record["instance_history"] = []
        entry = {
            "timestamp":         self._now(),
            "action":            action,
            "element_unique_id": element_unique_id,
            "user":              record.get("user", ""),
            "detail":            detail,
        }
        record["instance_history"].append(entry)

    def _sync_counters(self, record):
        """Recompute current/live instance counters for one doclink record.

        Uses the ImageType to discover ALL instances project-wide, including
        user-copied duplicates that are not yet in the tracked list.  Newly
        found copies are added to the tracked IDs automatically.
        """
        tracked_ids = _record_element_unique_ids(record)
        tracked_set = set(tracked_ids)

        # Discover copies by ImageType (project-wide)
        type_uid = record.get("image_type_unique_id")
        if type_uid:
            all_proj = _all_instances_by_type(self._doc, type_uid)
            for info in all_proj:
                uid = info["unique_id"]
                if uid not in tracked_set:
                    tracked_ids.append(uid)
                    tracked_set.add(uid)
            _set_record_element_unique_ids(record, tracked_ids)

        record["instance_count"] = len(_record_element_unique_ids(record))
        history = record.get("instance_history", [])
        deleted = sum(1 for e in history if e["action"] in ("Deleted", "Externally Deleted"))
        record["deleted_count"] = deleted

    def _check_external_deletion(self, record):
        """
        Remove any tracked element IDs that no longer exist in the model and log them.
        Returns True if any missing instances were found.
        """
        tracked_ids = _record_element_unique_ids(record)
        if not tracked_ids:
            _set_record_element_unique_ids(record, [])
            self._sync_counters(record)
            return False

        live_ids = []
        missing_found = False
        for eid in tracked_ids:
            try:
                el = _element_by_unique_id(self._doc, eid)
            except Exception:
                el = None
            if el is None:
                self._log_event(record, "Externally Deleted", eid,
                                "Element not found in Revit model (deleted outside tool)")
                missing_found = True
            else:
                live_ids.append(eid)

        _set_record_element_unique_ids(record, live_ids)
        if missing_found:
            record["status"] = "Missing" if not live_ids else "Partially Missing"
        self._sync_counters(record)
        return missing_found

    def _process(self, record, target_view=None):
        """Export (if needed) and place/replace in Revit. Mutates record. Returns (ok, msg)."""
        src = self._resolve_path(record)
        if not os.path.exists(src):
            return False, "Source not found:\n{}".format(src)

        # Clean tracked instances that may have been deleted outside the tool
        self._check_external_deletion(record)

        target_view = target_view or self._view
        if target_view is None:
            return False, "Target view could not be resolved for this doclink record."

        ftype = record.get("file_type", "image")
        page_number = _parse_page_numbers(record.get("page_number", "1"))[0]  # single page only
        dpi = _safe_int(record.get("dpi", 300), 300)
        transparent = _as_bool(record.get("transparent", False))

        tracked_ids = _record_element_unique_ids(record)
        existing_type_id = record.get("image_type_unique_id")
        if not existing_type_id and tracked_ids:
            existing_type_id = _image_type_unique_id_from_instance(self._doc, tracked_ids[0])
            if existing_type_id:
                record["image_type_unique_id"] = existing_type_id

        existing_ids = _resolve_existing_instance_unique_ids(
            self._doc, target_view, tracked_ids, existing_type_id
        )
        is_replace = len(existing_ids) > 0

        try:
            import_path = None   # single file to place

            LogManager.section("IMPORT PROCESSING PIPELINE")
            LogManager.info("File type: {}".format(ftype))
            LogManager.info("Source: {}".format(src))
            LogManager.info("Transparent: {}, DPI: {}".format(transparent, dpi))
            LogManager.info("Page number: {}".format(page_number))

            if ftype == "excel":
                LogManager.info("→ Step 1: Excel → PDF export")
                sheet    = record.get("sheet_name") or None
                rng      = record.get("range_addr", "A1:G20")
                # Name-Manager mode: when set, the workbook's defined name is
                # resolved at export time and overrides sheet+range.
                named    = record.get("named_range") or None
                LogManager.debug("  Sheet: '{}', Range: '{}', Named: '{}'".format(sheet, rng, named))
                try:
                    pdf_path = export_excel_range_to_pdf(src, rng, sheet, named_range=named)
                    LogManager.info("  PDF export: {}".format(pdf_path))
                except Exception as e:
                    LogManager.error("✗ Excel to PDF conversion failed: {}".format(e))
                    return False, (
                        "Excel to PDF conversion failed.\n"
                        "Requirements:\n"
                        "  - Revit 2024 & earlier: Works with bundled runtime\n"
                        "  - Revit 2025+: Requires Microsoft Excel installed on system\n"
                        "Error: {}".format(str(e))
                    )

                if transparent:
                    LogManager.info("→ Step 2: Transparent Excel mode - PDF → PNG")
                    if not _FITZ_AVAILABLE:
                        LogManager.error("✗ Transparent Excel requires CPython 3 + PyMuPDF")
                        return False, (
                            "Transparent Excel import requires CPython 3 + PyMuPDF.\n"
                            "Bundled runtime not found.\n"
                            "Expected: DocLink.pushbutton/runtime/python.exe"
                        )
                    # Excel → PDF → PNG (via PyMuPDF) → background removal below
                    LogManager.debug("  Converting PDF to PNG (page 1)...")
                    png = convert_pdf_page_to_png(pdf_path, page_number=1, dpi=dpi)
                    if not png:
                        LogManager.error("✗ Failed to convert Excel PDF to PNG")
                        return False, "Failed to render Excel export to PNG (page 1)."
                    LogManager.info("  Success: {}".format(png))
                    import_path = png
                else:
                    # transparent=OFF → import PDF directly (fast path)
                    LogManager.info("  Direct PDF import (transparen = False)")
                    import_path = pdf_path

            elif ftype == "word":
                LogManager.info("→ Step 1: Word → PDF export")
                try:
                    pdf_path, total_pages = export_word_to_pdf(src)
                    LogManager.info("  PDF export: {} ({} pages)".format(pdf_path, total_pages))
                except Exception as e:
                    LogManager.error("✗ Word to PDF conversion failed: {}".format(e))
                    return False, (
                        "Word to PDF conversion failed.\n"
                        "Requirements:\n"
                        "  - Revit 2024 & earlier: Works with bundled runtime\n"
                        "  - Revit 2025+: Requires Microsoft Word installed on system\n"
                        "Error: {}".format(str(e))
                    )

                if transparent:
                    LogManager.info("→ Step 2: Transparent Word mode - PDF → PNG")
                    if not _FITZ_AVAILABLE:
                        LogManager.error("✗ Transparent Word requires CPython 3 + PyMuPDF")
                        return False, (
                            "Transparent Word import requires CPython 3 + PyMuPDF.\n"
                            "Bundled runtime not found.\n"
                            "Expected: DocLink.pushbutton/runtime/python.exe"
                        )
                    # Word → PDF → PNG (via PyMuPDF) → background removal below
                    pn  = min(page_number, total_pages)
                    LogManager.debug("  Converting page {} of {} to PNG...".format(pn, total_pages))
                    png = convert_pdf_page_to_png(pdf_path, page_number=pn, dpi=dpi)
                    if not png:
                        LogManager.error("✗ Failed to convert Word PDF page {} to PNG".format(pn))
                        return False, "Failed to render Word page {} to PNG.".format(pn)
                    LogManager.info("  Success: {}".format(png))
                    import_path = png
                else:
                    # transparent=OFF → import PDF directly (fast path)
                    LogManager.info("  Direct PDF import (transparent = False)")
                    import_path = pdf_path

            elif ftype == "pdf":
                if transparent:
                    LogManager.info("→ Step 1: Transparent PDF mode - PDF → PNG")
                    if not _FITZ_AVAILABLE:
                        LogManager.error("✗ Transparent PDF requires CPython 3 + PyMuPDF")
                        return False, (
                            "Transparent PDF import requires CPython 3 + PyMuPDF.\n"
                            "Bundled runtime not found.\n"
                            "Expected: DocLink.pushbutton/runtime/python.exe"
                        )
                    LogManager.debug("  Converting page {} to PNG...".format(page_number))
                    png = convert_pdf_page_to_png(src, page_number=page_number, dpi=dpi)
                    if not png:
                        LogManager.error("✗ Failed to convert PDF page {} to PNG".format(page_number))
                        return False, "Failed to render PDF page {} to image.".format(page_number)
                    LogManager.info("  Success: {}".format(png))
                    import_path = png
                else:
                    # transparent=OFF → import PDF directly into Revit (fast path)
                    LogManager.info("  Direct PDF import (transparent = False)")
                    import_path = src

            else:
                LogManager.info("→ Image file (direct import)")
                import_path = src

            # ── Background removal (only when transparent=ON) ─────────────────
            # At this point import_path is always a raster PNG (never a PDF)
            # when transparent=True — the ftype branches above guarantee this.
            if transparent:
                LogManager.info("→ Step 3: Background removal")
                LogManager.debug("  Input: {}".format(import_path))
                processed = remove_background(import_path, dpi=dpi)
                if processed is None:
                    LogManager.error("✗ Background removal failed")
                    return False, "Background removal failed for:\n{}".format(import_path)
                LogManager.info("✓ Background removal output: {}".format(processed))
                import_path = processed

            # ── Trim outer blank margins (only when transparent=ON) ───────────
            if transparent:
                LogManager.info("→ Step 4: Margin trimming")
                _png_ext = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")
                if os.path.splitext(import_path)[1].lower() in _png_ext:
                    LogManager.debug("  Input: {}".format(import_path))
                    cropped = crop_white_margins(import_path)
                    if cropped != import_path:
                        LogManager.info("✓ Margins trimmed: {}".format(cropped))
                    else:
                        LogManager.debug("  No trimming needed")
                    import_path = cropped

            # ── Placement ─────────────────────────────────────────────────────
            LogManager.info("→ Step 5: Place into Revit view")
            LogManager.debug("  Final import file: {}".format(import_path))
            LogManager.debug("  View: {}".format(target_view.Name))
            LogManager.debug("  Is replace: {}".format(is_replace))

            result = place_or_replace(
                self._doc, target_view, import_path,
                existing_element_unique_ids=existing_ids,
                existing_type_unique_id=existing_type_id,
                dpi=dpi,
                transparent=transparent,
                import_name=record.get("name", "")
            )
            new_ids = result.get("element_unique_ids", [])
            record["image_type_unique_id"] = result.get("image_type_unique_id")

            action = "Replaced" if is_replace else "Placed"
            detail = describe_record_options(record)
            LogManager.info("✓ Image {} successfully: {} elements".format(action.lower(), len(new_ids)))

            if is_replace:
                for old_id in existing_ids:
                    self._log_event(record, "Deleted", old_id,
                                    "Auto-deleted on {}".format(action.lower()))
            for new_id in new_ids:
                self._log_event(record, action, new_id, detail)

            _set_record_element_unique_ids(record, new_ids)
            record["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            record["status"] = "OK"
            self._sync_counters(record)
            LogManager.info("✓ Import completed successfully")
            LogManager.info("Log file: {}".format(LogManager.get_log_path()))
            return True, "OK"

        except Exception as ex:
            record["status"] = "Error"
            LogManager.error("✗ IMPORT FAILED: An exception occurred")
            LogManager.exception("Exception details:")
            LogManager.error("Error message: {}".format(str(ex)))
            LogManager.error("Log file: {}".format(LogManager.get_log_path()))
            LogManager.section("END OF FAILED IMPORT SESSION")
            return False, str(ex)


    # ── event handlers ────────────────────────────────────────────────────────

    def _on_row_double_click(self, s, e):
        """Double-click a row → show full instance history for that record."""
        item = self._dg.SelectedItem
        if not item:
            return
        rid = item._record_id
        record = next((r for r in self._records if r.get("id") == rid), None)
        if not record:
            return
        dlg = InstanceHistoryWindow(record)
        dlg.Owner = self
        dlg.ShowDialog()

    def _on_add(self, s, e):
        dlg = AddImportDialog()
        dlg.Owner = self
        if not dlg.ShowDialog():
            return
        info = dlg.result
        if not info:
            return

        record = {
            "id":                    _uid(),
            "view_id":               self._view_id(),
            "name":                  info["name"],
            "user":                  info["user"],            # auto-derived (Phase 3)
            "auto_named":            info.get("auto_named", False),  # Phase 2
            "path_type":             info["path_type"],
            "file_type":             info["file_type"],
            "sheet_name":            info.get("sheet_name", ""),
            "range_addr":            info.get("range_addr", ""),
            "page_number":           info.get("page_number", "1"),
            "dpi":                   info.get("dpi", "300"),
            "transparent":           info.get("transparent", False),
            "element_unique_id":     None,
            "element_unique_ids":    [],
            "image_type_unique_id":  None,
            "last_updated":          "—",
            "status":                "Pending",
            "instance_count":        0,
            "deleted_count":         0,
            "instance_history":      [],
        }

        # Store relative path as-is (relative to project)
        raw_path = info["source_path"]
        if info["path_type"] == "Relative":
            rvt = self._doc.PathName
            if rvt:
                try:
                    raw_path = os.path.relpath(raw_path, os.path.dirname(rvt))
                except ValueError:
                    pass   # different drives – keep absolute
        record["source_path"] = raw_path

        # Log initial creation event
        self._log_event(record, "Created", detail="Registered in DocLink Manager")

        self._status.Content = "Importing '{}' …".format(record["name"])
        self._progress_start(1)
        self._progress_update(0, "Importing '{}' …".format(record["name"]))
        try:
            ok, msg = self._process(record)
            self._progress_end()
            if ok:
                self._records.append(record)
                save_records(self._doc, self._records)
                self._refresh()
                self._status.Content = "✔  '{}' placed successfully.".format(record["name"])
            else:
                forms.alert("Import failed:\n" + msg)
                self._status.Content = "✖  Import failed."
            self._bring_to_focus()
        except Exception as ex:
            self._progress_end()
            LogManager.error("✗ UNHANDLED EXCEPTION in _on_add for '{}'".format(record.get("name", "?")))
            LogManager.error("Exception: {}".format(str(ex)))
            LogManager.error("Full traceback:")
            for line in traceback.format_exc().split("\n"):
                if line.strip():
                    LogManager.error(line)
            forms.alert("Import failed with error:\n" + str(ex))
            self._status.Content = "✖  Import failed."
            self._bring_to_focus()
            LogManager.debug("[DocLinkManager] _on_add exception traceback:\n{}".format(traceback.format_exc()))

    def _selected_record_ids(self):
        ids = set()
        for item in self._dg.SelectedItems:
            try:
                ids.add(item._record_id)
            except Exception:
                pass
        return ids

    def _on_update_sel(self, s, e):
        ids = self._selected_record_ids()
        if not ids:
            forms.alert("Select one or more rows first.")
            self._bring_to_focus()
            return

        targets = [r for r in self._records if r.get("id") in ids]
        total = len(targets)
        self._progress_start(total)

        updated, errors = 0, []
        for i, r in enumerate(targets):
            self._progress_update(i, "Updating '{0}' … ({1}/{2})".format(
                r.get("name", "?"), i + 1, total))
            try:
                ok, msg = self._process(r, target_view=self._resolve_record_view(r))
                if ok: 
                    updated += 1
                else:  
                    errors.append("{}: {}".format(r.get("name","?"), msg))
            except Exception as ex:
                LogManager.error("✗ UNHANDLED EXCEPTION in _on_update_sel for '{}'".format(r.get("name", "?")))
                LogManager.error("Exception: {}".format(str(ex)))
                LogManager.error("Full traceback:")
                for line in traceback.format_exc().split("\n"):
                    if line.strip():
                        LogManager.error(line)
                errors.append("{}: Unhandled exception - {}".format(r.get("name","?"), str(ex)))
                LogManager.debug("[DocLinkManager] _on_update_sel exception traceback:\n{}".format(traceback.format_exc()))

        self._progress_end()
        save_records(self._doc, self._records)
        self._refresh()
        if errors:
            forms.alert("Errors:\n" + "\n".join(errors))
        self._status.Content = "✔  {} record(s) updated.".format(updated)
        self._bring_to_focus()

    def _on_update_all(self, s, e):
        all_recs = list(self._records)
        if not all_recs:
            forms.alert("No imports tracked in the project.")
            self._bring_to_focus()
            return

        total = len(all_recs)
        self._progress_start(total)

        updated, errors = 0, []
        for i, r in enumerate(all_recs):
            self._progress_update(i, "Updating '{0}' … ({1}/{2})".format(
                r.get("name", "?"), i + 1, total))
            try:
                ok, msg = self._process(r, target_view=self._resolve_record_view(r))
                if ok: 
                    updated += 1
                else:  
                    errors.append("{}: {}".format(r.get("name","?"), msg))
            except Exception as ex:
                LogManager.error("✗ UNHANDLED EXCEPTION in _on_update_all for '{}'".format(r.get("name", "?")))
                LogManager.error("Exception: {}".format(str(ex)))
                LogManager.error("Full traceback:")
                for line in traceback.format_exc().split("\n"):
                    if line.strip():
                        LogManager.error(line)
                errors.append("{}: Unhandled exception - {}".format(r.get("name","?"), str(ex)))
                LogManager.debug("[DocLinkManager] _on_update_all exception traceback:\n{}".format(traceback.format_exc()))

        self._progress_end()
        save_records(self._doc, self._records)
        self._refresh()
        if errors:
            forms.alert("Errors during update:\n" + "\n".join(errors))
        self._status.Content = "✔  {}/{} record(s) updated.".format(updated, total)
        self._bring_to_focus()

    def _on_update_project_all(self, s, e):
        self._sync_sched_ui_to_records()
        all_stickies = list(self._records)
        all_scheds   = list(self._sched_records)
        total = len(all_stickies) + len(all_scheds)
        self._progress_start(total)
        step = 0

        doclink_updated, doclink_errors = 0, []
        for r in all_stickies:
            step += 1
            self._progress_update(step, "Updating doclink '{0}' … ({1}/{2})".format(
                r.get("name", "?"), step, total))
            target_view = self._record_view_element(r)
            if target_view is None:
                doclink_errors.append("{}: target view not found".format(r.get("name", "?")))
                r["status"] = "Missing View"
                continue
            try:
                ok, msg = self._process(r, target_view=target_view)
                if ok:
                    doclink_updated += 1
                else:
                    doclink_errors.append("{}: {}".format(r.get("name", "?"), msg))
            except Exception as ex:
                r["status"] = "Error"
                LogManager.error("✗ UNHANDLED EXCEPTION in _on_update_project_all [doclink '{0}']".format(
                    r.get("name", "?")))
                LogManager.error("Exception: {}".format(str(ex)))
                LogManager.error("Full traceback:")
                for line in traceback.format_exc().split("\n"):
                    if line.strip():
                        LogManager.error(line)
                doclink_errors.append("{}: {}".format(r.get("name", "?"), str(ex)))
                print("[DocLinkManager] [_on_update_project_all doclink] exception traceback:")
                print(traceback.format_exc())

        schedule_updated, schedule_errors = 0, []
        for r in all_scheds:
            step += 1
            self._progress_update(step, "Updating schedule '{0}' … ({1}/{2})".format(
                r.get("schedule_name", "?"), step, total))
            try:
                self._rebuild_schedule_record(r)
                schedule_updated += 1
            except Exception as ex:
                r["status"] = "Error"
                LogManager.error("✗ UNHANDLED EXCEPTION in _on_update_project_all [schedule '{0}']".format(
                    r.get("schedule_name", "?")))
                LogManager.error("Exception: {}".format(str(ex)))
                LogManager.error("Full traceback:")
                for line in traceback.format_exc().split("\n"):
                    if line.strip():
                        LogManager.error(line)
                schedule_errors.append("{}: {}".format(r.get("schedule_name", "?"), str(ex)))
                LogManager.debug("[DocLinkManager] [_on_update_project_all schedule] exception traceback:\n{}".format(traceback.format_exc()))

        self._progress_end()
        save_records(self._doc, self._records)
        save_schedule_records(self._doc, self._sched_records)
        self._refresh()
        self._sched_refresh()

        all_errors = doclink_errors + schedule_errors
        if all_errors:
            forms.alert("Update completed with errors:\n" + "\n".join(all_errors))
        self._status.Content = "✔  Project update complete — {} doclink import(s), {} schedule(s).".format(
            doclink_updated, schedule_updated)
        self._bring_to_focus()

    def _on_edit(self, s, e):
        ids = self._selected_record_ids()
        if len(ids) != 1:
            forms.alert("Select exactly one row to edit.")
            self._bring_to_focus()
            return

        rid = next(iter(ids))
        record = next((r for r in self._records if r.get("id") == rid), None)
        if not record:
            forms.alert("Record not found.")
            self._bring_to_focus()
            return

        dlg = EditImportDialog(record)
        dlg.Owner = self
        if not dlg.ShowDialog():
            self._bring_to_focus()
            return
        info = dlg.result
        if not info:
            self._bring_to_focus()
            return

        # Apply edits to record
        record["source_path"] = to_relative_path(info["source_path"], self._doc.PathName) if info["path_type"] == "Relative" else info["source_path"]
        record["path_type"]   = info["path_type"]
        record["file_type"]    = info["file_type"]
        record["sheet_name"]   = info.get("sheet_name", "")
        record["range_addr"]   = info.get("range_addr", "")
        record["page_number"]  = info.get("page_number", "1")
        record["dpi"]          = info.get("dpi", "300")
        record["transparent"]  = info.get("transparent", False)
        if info.get("name"):
            record["name"] = info["name"]

        # Log edit intent before processing
        self._log_event(record, "Edited",
                        detail="Source/import settings changed via Edit dialog")

        self._progress_start(1)
        self._progress_update(0, "Applying edits to '{}' …".format(record["name"]))
        try:
            ok, msg = self._process(record, target_view=self._resolve_record_view(record))
        except Exception as ex:
            LogManager.error("✗ UNHANDLED EXCEPTION in _on_edit for '{}'".format(record.get("name", "?")))
            LogManager.error("Exception: {}".format(str(ex)))
            LogManager.error("Full traceback:")
            for line in traceback.format_exc().split("\n"):
                if line.strip():
                    LogManager.error(line)
            LogManager.debug("[DocLinkManager] _on_edit exception traceback:\n{}".format(traceback.format_exc()))
            ok = False
            msg = "Unhandled exception: {}".format(str(ex))
        
        self._progress_end()
        if ok:
            save_records(self._doc, self._records)
            self._refresh()
            self._status.Content = "✔  '{}' updated and replaced at the same location(s).".format(record["name"])
        else:
            forms.alert("Re-import failed:\n" + msg)
            self._status.Content = "✖  Edit failed."
        self._bring_to_focus()

    def _on_remove(self, s, e):
        ids = self._selected_record_ids()
        if not ids:
            forms.alert("Select rows to remove.")
            self._bring_to_focus()
            return

        if not forms.alert(
            "Remove {} selected record(s)?\n\nRevit elements will also be deleted.\n\n"
            "Click OK to confirm, or Cancel to abort.".format(len(ids)),
            title="Confirm Remove", cancel=True
        ):
            self._bring_to_focus()
            return

        records_to_remove = [r for r in self._records if r.get("id") in ids]

        LogManager.debug("[DocLinkManager] _on_remove: selected ids={}, matched records={}".format(ids, len(records_to_remove)))

        removed_count = 0

        for r in records_to_remove:
            tracked_ids = _record_element_unique_ids(r)
            image_type_uid = r.get("image_type_unique_id")

            LogManager.debug("[DocLinkManager] Removing '{}': tracked_ids={}, image_type_uid={}".format(
                r.get("name", "?"), tracked_ids, image_type_uid))

            # ── Resolve live elements ────────────────────────────────────────
            live_ids = _resolve_existing_instance_unique_ids(
                self._doc, self._view, tracked_ids,
                image_type_uid, r.get("name")
            )

            LogManager.debug("[DocLinkManager] Resolved live_ids={}".format(live_ids))

            deleted_count = 0

            with Transaction(self._doc, "DocLink Manager – Remove") as t:
                t.Start()

                if live_ids:
                    # Primary path: delete all tracked live elements
                    for eid in live_ids:
                        try:
                            el = _element_by_unique_id(self._doc, eid)
                            if el:
                                self._log_event(r, "Deleted", eid,
                                                "Removed by user via DocLink Manager")
                                self._doc.Delete(el.Id)
                                deleted_count += 1
                                LogManager.debug("[DocLinkManager] Deleted element: {}".format(eid))
                            else:
                                self._log_event(r, "Externally Deleted", eid,
                                                "Element missing when Remove triggered")
                                LogManager.debug("[DocLinkManager] Element not found for deletion: {}".format(eid))
                        except Exception as ex:
                            LogManager.debug("[DocLinkManager] Delete failed for {}: {}".format(eid, ex))
                else:
                    # Patch C fallback: view-scoped search found nothing —
                    # try a doc-wide search before giving up
                    doc_wide = _find_instances_doc_wide(self._doc, r.get("name", ""))
                    LogManager.debug("[DocLinkManager] Doc-wide fallback found {} elements".format(len(doc_wide)))
                    if doc_wide:
                        for uid, inst in doc_wide:
                            try:
                                self._log_event(r, "Deleted", uid,
                                                "Found via doc-wide fallback, removed")
                                self._doc.Delete(inst.Id)
                                deleted_count += 1
                                LogManager.debug("[DocLinkManager] Deleted via doc-wide: {}".format(uid))
                            except Exception as ex:
                                LogManager.debug("[DocLinkManager] Doc-wide delete failed for {}: {}".format(
                                    uid, ex))
                    else:
                        # Truly no elements found — log each tracked id
                        for eid in tracked_ids:
                            self._log_event(r, "Externally Deleted", eid,
                                            "Element not found on Remove (deleted externally?)")

                # Also clean up the ImageType if no other instances reference it
                if image_type_uid and deleted_count > 0:
                    try:
                        type_el = _element_by_unique_id(self._doc, image_type_uid)
                        if type_el is not None:
                            # Check if any other ImageInstance still uses this type
                            remaining = FilteredElementCollector(self._doc).OfClass(ImageInstance)
                            still_used = False
                            for inst in remaining:
                                try:
                                    if inst.GetTypeId() == type_el.Id:
                                        still_used = True
                                        break
                                except Exception:
                                    pass
                            if not still_used:
                                self._doc.Delete(type_el.Id)
                                LogManager.debug("[DocLinkManager] Cleaned up unused ImageType: {}".format(
                                    image_type_uid))
                    except Exception as ex:
                        LogManager.debug("[DocLinkManager] ImageType cleanup failed: {}".format(ex))

                t.Commit()

            LogManager.debug("[DocLinkManager] deleted_count={}, tracked_ids={}".format(
                deleted_count, tracked_ids))

            # Patch D: if we expected elements but found none, ask the user
            # whether to clean up the orphaned record or keep it
            if deleted_count == 0 and tracked_ids:
                keep = not forms.alert(
                    "No Revit elements were found for '{}'.\n\n"
                    "The image may have been deleted manually outside this tool.\n\n"
                    "Click OK to remove the orphaned record from the tracker,\n"
                    "or Cancel to keep the record.".format(r.get("name", "?")),
                    title="Element not found", cancel=True
                )
                if keep:
                    # User chose to keep the record — do not remove it
                    continue

            # Remove the record from tracking
            _set_record_element_unique_ids(r, [])
            self._sync_counters(r)
            self._records.remove(r)
            removed_count += 1

        save_records(self._doc, self._records)
        self._refresh()
        self._status.Content = "✔  {} record(s) removed.".format(removed_count)
        self._bring_to_focus()

    # ── Tab 2: Schedule Import handlers ───────────────────────────────────────

    def _sched_refresh(self):
        self._sched_rows.Clear()
        for i, r in enumerate(self._sched_records):
            opts = r.get("options", {})
            opts_str = "Row {:.1f}x | Col {:.1f}x | Txt {:.1f}x | Sc 1:{}".format(
                opts.get("row_scale", 1.0),
                opts.get("col_scale", 1.0),
                opts.get("text_scale", 1.0),
                opts.get("view_scale", 1),
            )
            row = ScheduleRow(
                record_id   = r.get("id", ""),
                idx         = i + 1,
                name        = r.get("schedule_name", ""),
                source_file = os.path.basename(r.get("source_path", "")),
                sheet_name  = r.get("sheet_name", ""),
                options_info = opts_str,
                last_updated = r.get("last_updated", "—"),
                status       = r.get("status", "—"),
                retain_settings = r.get("retain_settings", True),
                path_type    = r.get("path_type", "Absolute"),
            )
            self._sched_rows.Add(row)
        self._status.Content = "{} schedule import(s) tracked.".format(
            len(self._sched_rows))

    def _on_sched_selection_changed(self, s, e):
        """Cache selected record IDs immediately on selection change.
        WPF can clear SelectedItems when a button click commits DataGrid
        cell edits via PropertyChanged (INotifyPropertyChanged race).
        Snapshotting here guarantees the IDs are available in button handlers."""
        ids = set()
        for item in self._sched_dg.SelectedItems:
            try:
                ids.add(item._record_id)
            except Exception:
                pass
        self._sched_selected_ids_cache = ids

    def _sched_selected_record_ids(self):
        # Primary: read directly from SelectedItems
        ids = set()
        for item in self._sched_dg.SelectedItems:
            try:
                ids.add(item._record_id)
            except Exception:
                pass
        # Fallback: use snapshot from SelectionChanged event if SelectedItems lost
        if not ids:
            ids = getattr(self, "_sched_selected_ids_cache", set())
        return ids

    def _record_view_element(self, record):
        vid = _safe_int(record.get("view_id"), None)
        if vid is None:
            return self._view
        try:
            return self._doc.GetElement(ElementId(int(vid)))
        except Exception:
            return None

    def _rebuild_schedule_record(self, r):
        if self._uidoc is None:
            raise RuntimeError("uidoc not available.")

        # Use individual Transactions (no TransactionGroup) for safety in WPF handlers
        # tg = TransactionGroup(self._doc, "Schedule Import")
        # tg.Start()

        old_ssi_uid    = r.get("ssi_unique_id")
        old_sched_uid  = r.get("schedule_unique_id")
        saved_pt       = None
        target_sheet   = None
        saved_sheet_uid = r.get("target_sheet_unique_id")
        
        output_type = r.get("output_type", "Schedule")
        is_graphic = output_type != "Schedule"

        if old_ssi_uid and not is_graphic:
            try:
                old_ssi = self._doc.GetElement(str(old_ssi_uid))
                if old_ssi is not None:
                    raw_pt = old_ssi.Point
                    saved_pt = XYZ(raw_pt.X, raw_pt.Y, raw_pt.Z)
                    owner_s = self._doc.GetElement(old_ssi.OwnerViewId)
                    if owner_s is not None and isinstance(owner_s, ViewSheet):
                        saved_sheet_uid = owner_s.UniqueId
                        target_sheet = owner_s

                    with Transaction(self._doc, "Schedule Import – Delete Old SSI") as t:
                        t.Start()
                        self._doc.Delete(old_ssi.Id)
                        t.Commit()
            except Exception as ex:
                LogManager.debug("[DocLinkManager] SSI delete/read error for {}: {}".format(old_ssi_uid, ex))

        if old_sched_uid and not is_graphic:
            try:
                old_sched_el = self._doc.GetElement(str(old_sched_uid))
                if old_sched_el is not None:
                    # Capture current widths before deletion if retain logic is on
                    retain_logic = _as_bool(r.get("retain_settings", True))
                    if retain_logic:
                        fresh_widths = _extract_column_width_snapshot(self._doc, old_sched_el)
                        if fresh_widths:
                            r["width_snapshot"] = fresh_widths

                    with Transaction(self._doc, "Schedule Import – Delete Old View") as t:
                        t.Start()
                        self._doc.Delete(old_sched_el.Id)
                        t.Commit()
            except Exception as ex:
                LogManager.debug("[DocLinkManager] Sched delete error for {}: {}".format(old_sched_uid, ex))

        if target_sheet is None and saved_sheet_uid:
            try:
                ts = self._doc.GetElement(str(saved_sheet_uid))
                if ts is not None and isinstance(ts, ViewSheet):
                    target_sheet = ts
            except Exception:
                pass

        try:
            result = _build_schedule_from_record(
                self._doc, self._uidoc, r,
                placement_pt=saved_pt,
                target_sheet=target_sheet,
            )
        except Exception as _ex:
            LogManager.error("✗ UNHANDLED EXCEPTION in _rebuild_schedule_record")
            LogManager.error("Exception: {}".format(str(_ex)))
            r["status"] = "Error"
            # tg.RollBack()
            return

        # tg.Commit()

        r["schedule_unique_id"]     = result.get("schedule_unique_id")
        r["ssi_unique_id"]          = result.get("ssi_unique_id")
        r["merge_snapshot"]         = result.get("merge_snapshot")
        r["border_snapshot"]        = result.get("border_snapshot")
        r["width_snapshot"]         = result.get("width_snapshot")
        r["target_sheet_unique_id"] = saved_sheet_uid
        new_ssi_uid = result.get("ssi_unique_id")
        if new_ssi_uid:
            try:
                new_ssi = self._doc.GetElement(str(new_ssi_uid))
                if new_ssi is not None:
                    raw_pt = new_ssi.Point
                    r["last_placement_point"] = [raw_pt.X, raw_pt.Y, raw_pt.Z]
                    LogManager.debug("[DocLink] New SSI placed at ({:.4f}, {:.4f})".format(
                        raw_pt.X, raw_pt.Y))
            except Exception:
                pass
        new_ssi_placed = bool(result.get("ssi_unique_id"))
        r["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        r["status"] = "OK" if new_ssi_placed else "Not on Sheet"

    def _reimport_schedule_records(self, records, done_message):
        total = len(records)
        self._progress_start(total)

        updated, errors = 0, []
        for i, r in enumerate(records):
            self._progress_update(i, "Re-importing '{0}' … ({1}/{2})".format(
                r.get("schedule_name", "?"), i + 1, total))
            try:
                self._rebuild_schedule_record(r)
                updated += 1
            except Exception as ex:
                r["status"] = "Error"
                LogManager.error("✗ UNHANDLED EXCEPTION in _reimport_schedule_records ['{0}']".format(
                    r.get("schedule_name", "?")))
                LogManager.error("Exception: {}".format(str(ex)))
                LogManager.error("Full traceback:")
                for line in traceback.format_exc().split("\n"):
                    if line.strip():
                        LogManager.error(line)
                errors.append("{}: {}".format(r.get("schedule_name", "?"), str(ex)))
                LogManager.debug("[DocLinkManager] [_reimport_schedule_records] exception traceback:\n{}".format(traceback.format_exc()))

        self._progress_end()
        save_schedule_records(self._doc, self._sched_records)
        self._sched_refresh()
        if errors:
            forms.alert("Errors:\n" + "\n".join(errors))
        self._status.Content = done_message.format(updated)
        return updated, errors

    def _on_sched_add(self, s, e):
        if self._uidoc is None:
            forms.alert("uidoc not available – cannot place schedule on sheet.")
            return

        self._status.Content = "Running Document Importer …"
        self._progress_start(1)
        self._progress_update(0, "Running Document Importer …")
        try:
            result = run_document_importer(self._doc, self._uidoc)
        except Exception as ex:
            self._progress_end()
            LogManager.debug("[DocLinkManager] _on_sched_add exception traceback:\n{}".format(traceback.format_exc()))
            forms.alert("Schedule import failed:\n{}: {}".format(type(ex).__name__, str(ex)))
            self._status.Content = "✖  Import failed."
            return

        self._progress_end()
        if result is None:
            self._status.Content = "Import cancelled."
            return

        # Phase 5: record the sheet and placement point for update preservation
        target_sheet_uid = None
        last_pt = None
        ssi_uid = result.get("ssi_unique_id")
        if ssi_uid:
            try:
                ssi_el = self._doc.GetElement(str(ssi_uid))
                if ssi_el is not None:
                    owner_sheet = self._doc.GetElement(ssi_el.OwnerViewId)
                    if owner_sheet is not None and isinstance(owner_sheet, ViewSheet):
                        target_sheet_uid = owner_sheet.UniqueId
                    raw_pt = ssi_el.Point
                    last_pt = [raw_pt.X, raw_pt.Y, raw_pt.Z]
            except Exception as ex:
                LogManager.debug("[DocLinkManager] _on_sched_add: could not read SSI point: {}".format(ex))

        record = {
            "id":                       _uid(),
            "schedule_name":            result.get("schedule_name", ""),
            "source_path":              result.get("source_path", ""),
            "path_type":                result.get("path_type", "Absolute"),
            "sheet_name":               result.get("sheet_name", ""),
            "options":                  result.get("options", {}),
            "schedule_unique_id":       result.get("schedule_unique_id"),
            "ssi_unique_id":            ssi_uid,
            "target_sheet_unique_id":   target_sheet_uid,       # Phase 5
            "last_placement_point":     last_pt,                # Phase 5
            "merge_snapshot":           result.get("merge_snapshot"),  # Phase 7
            "border_snapshot":          result.get("border_snapshot"),  # border retain
            "width_snapshot":           result.get("width_snapshot"),   # V25 width retain
            "retain_settings":          True,  # Default to True for new imports
            "options_version":          1,
            "last_updated":             datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "status":                   "OK" if ssi_uid else "Not on Sheet",
        }
        self._sched_records.append(record)
        save_schedule_records(self._doc, self._sched_records)
        self._sched_refresh()
        self._status.Content = "✔  Schedule '{}' created and tracked.".format(
            record["schedule_name"])

    def _sync_sched_ui_to_records(self):
        """Sync DataGrid state back to persistence list."""
        lookup = {r.get("id"): r for r in self._sched_records}
        for row in self._sched_rows:
            rid = row._record_id
            if rid in lookup:
                rec = lookup[rid]
                rec["retain_settings"] = row.retain_settings
                # path_type is usually handled via dialog, but we sync it here too for correctness
                rec["path_type"] = row.path_type

    def _on_sched_update(self, s, e):
        self._sync_sched_ui_to_records()
        ids = self._sched_selected_record_ids()
        if not ids:
            forms.alert("Select one or more rows to re-import.")
            self._bring_to_focus()
            return
        selected = [r for r in self._sched_records if r.get("id") in ids]
        self._reimport_schedule_records(selected, "✔  {} schedule(s) re-imported.")

    def _on_sched_update_all(self, s, e):
        self._sync_sched_ui_to_records()
        if not self._sched_records:
            forms.alert("No tracked schedules found.")
            self._bring_to_focus()
            return
        self._reimport_schedule_records(list(self._sched_records), "✔  {} schedule(s) re-imported.")

    def _on_sched_remove(self, s, e):
        ids = self._sched_selected_record_ids()
        if not ids:
            forms.alert("Select rows to remove.")
            self._bring_to_focus()
            return

        if not forms.alert(
            "Remove {} selected schedule(s)?\n\n"
            "The Revit schedule elements will also be deleted.\n\n"
            "Click OK to confirm, or Cancel to abort.".format(len(ids)),
            title="Confirm Remove", cancel=True
        ):
            return

        removed = 0
        for r in list(self._sched_records):
            if r.get("id") not in ids:
                continue

            # Delete the schedule from Revit
            sched_uid = r.get("schedule_unique_id")
            ssi_uid   = r.get("ssi_unique_id")

            with Transaction(self._doc, "Schedule Import – Remove") as t:
                t.Start()
                for uid in (ssi_uid, sched_uid):
                    if not uid:
                        continue
                    try:
                        el = self._doc.GetElement(str(uid))
                        if el is not None:
                            self._doc.Delete(el.Id)
                    except Exception as ex:
                        LogManager.debug("[DocLinkManager] Sched delete error for {}: {}".format(uid, ex))
                t.Commit()

            self._sched_records.remove(r)
            removed += 1

        save_schedule_records(self._doc, self._sched_records)
        self._sched_refresh()
        self._status.Content = "✔  {} schedule(s) removed.".format(removed)

    def _on_sched_edit(self, s, e):
        """Phase 6: open ScheduleSetupDialog for exactly one selected record,
        persist edited options, and trigger a non-interactive re-import."""
        ids = self._sched_selected_record_ids()
        if len(ids) != 1:
            forms.alert("Select exactly one schedule row to edit.")
            self._bring_to_focus()
            return
        if self._uidoc is None:
            forms.alert("uidoc not available.")
            self._bring_to_focus()
            return

        rid = next(iter(ids))
        record = next((r for r in self._sched_records if r.get("id") == rid), None)
        if not record:
            forms.alert("Record not found.")
            self._bring_to_focus()
            return

        dlg = ScheduleSetupDialog(record, doc=self._doc)
        _dlg_result = dlg.ShowDialog()
        # Bring WPF window back to front after WinForms dialog steals focus.
        # Without this the DocLink window goes behind Revit and appears frozen.
        try:
            self.Activate()
        except Exception:
            pass
        if _dlg_result != WFDialogResult.OK:
            return
        edited = dlg.result
        if not edited:
            return

        # Apply edited values to the record.
        # `source_path` is read-only in the schedule edit dialog and older
        # dialog builds (before the round-trip fix) didn't put it in `res`;
        # fall back to whatever's already on the record so an incomplete
        # dialog result can't break the rebuild.
        raw_path = edited.get("source_path") or record.get("source_path") or ""
        path_type = edited.get("path_type", "Absolute")
        record["source_path"] = to_relative_path(raw_path, self._doc.PathName) if path_type == "Relative" else raw_path
        record["path_type"]   = path_type
        record["sheet_name"]  = edited.get("sheet_name") or record.get("sheet_name") or None
        opts = record.get("options") or {}

        # Detect scale-factor changes BEFORE overwriting opts so we can
        # decide whether to invalidate the cached snapshots.
        _prev_col_scale  = float(opts.get("col_scale",  1.0))
        _prev_row_scale  = float(opts.get("row_scale",  1.0))
        _prev_txt_scale  = float(opts.get("text_scale", 1.0))
        _prev_v_scale    = int(opts.get("view_scale", 1))
        _prev_dyn_merge  = bool(opts.get("enable_dynamic_merge", True))

        opts["row_scale"]                    = edited["row_scale"]
        opts["col_scale"]                    = edited["col_scale"]
        opts["text_scale"]                   = edited["text_scale"]
        opts["view_scale"]                   = edited["view_scale"]
        opts["output_type"]                  = edited.get("output_type", opts.get("output_type", "Schedule"))
        opts["scale_by_view"]                = edited.get("scale_by_view", True)
        opts["enable_dynamic_merge"]         = edited["enable_dynamic_merge"]
        opts["show_empty_gridlines"]         = edited["show_empty_gridlines"]
        opts["retain_previous_merge_layout"] = edited["retain_settings"]
        opts["default_text_note_type"]       = edited.get("default_text_note_type")
        opts["base_text_size_mm"]            = edited.get("base_text_size_mm", 0.0)
        opts["excel_points_per_mm"]          = edited.get("excel_points_per_mm", 2.834)
        opts["line_style_map"]               = edited.get("line_style_map", {})
        opts["text_note_map"]                = edited.get("text_note_map", {})
        opts["single_text_note_type"]        = edited.get("single_text_note_type")
        opts["fill_region_map"]              = edited.get("fill_region_map", {})
        # Range-source selection from the schedule dialog. Empty strings mean
        # the user picked Default (Print Area) mode — explicitly persist them
        # so a previous named/manual choice is cleared when the user opts back
        # out. parse_excel treats the priority named > manual > print area.
        opts["named_range"]                  = edited.get("named_range",  "") or ""
        opts["manual_range"]                 = edited.get("manual_range", "") or ""

        record["retain_settings"]            = edited["retain_settings"]
        record["output_type"]                = opts["output_type"]
        record["options"] = opts

        # If col_scale changed, the stored width_snapshot reflect the OLD
        # scale and would override the new one -- drop it so _build_schedule
        # computes fresh widths from the new col_scale.
        _col_changed = abs(edited["col_scale"]  - _prev_col_scale) > 0.001
        _row_changed = abs(edited["row_scale"]  - _prev_row_scale) > 0.001
        _txt_changed = abs(edited["text_scale"] - _prev_txt_scale) > 0.001
        _mrg_changed = (edited["enable_dynamic_merge"] != _prev_dyn_merge)
        _resize_only = bool(edited.get("resize_only", False))

        if _col_changed:
            record["width_snapshot"] = None   # recompute from new col_scale
            LogManager.debug("[DocLink] col_scale changed {:.2f}->{:.2f}: width_snapshot cleared".format(
                _prev_col_scale, edited["col_scale"]))
        if _mrg_changed or _col_changed:
            record["merge_snapshot"] = None  # recompute merges with new settings
            LogManager.debug("[DocLink] scale/merge change: merge_snapshot cleared")

        # Persist the edited setup first so it survives even if rebuild fails
        save_schedule_records(self._doc, self._sched_records)
        self._sched_refresh()

        if _resize_only and record.get("output_type") == "Schedule":
            schedule_el = None
            try:
                sched_uid = record.get("schedule_unique_id")
                if sched_uid:
                    schedule_el = self._doc.GetElement(str(sched_uid))
            except Exception:
                schedule_el = None

            if schedule_el is not None:
                try:
                    _resize_schedule_in_place(
                        self._doc,
                        schedule_el,
                        row_ratio=(edited["row_scale"] / _prev_row_scale) if _prev_row_scale else 1.0,
                        col_ratio=(edited["col_scale"] / _prev_col_scale) if _prev_col_scale else 1.0,
                        text_ratio=(edited["text_scale"] / _prev_txt_scale) if _prev_txt_scale else 1.0,
                        resize_columns=(not _as_bool(record.get("retain_settings", True))),
                    )
                    record["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    record["status"] = "OK"
                    save_schedule_records(self._doc, self._sched_records)
                    self._sched_refresh()
                    self._status.Content = "Setup saved. Resized '{}' without re-import.".format(
                        record.get("schedule_name", "?"))
                    self._bring_to_focus()
                    return
                except Exception as ex:
                    forms.alert("Resize failed:\n{}".format(str(ex)))
                    self._status.Content = "Resize failed for '{}'".format(
                        record.get("schedule_name", "?"))
                    self._bring_to_focus()
                    return

            self._status.Content = "Setup saved. Existing schedule not found; no re-import was run."
            self._bring_to_focus()
            return

        self._status.Content = "Setup saved. Re-importing '{}' …".format(
            record.get("schedule_name", "?"))

        # Trigger non-interactive rebuild using the new settings
        # Re-use _on_sched_update logic by temporarily limiting selection
        saved_pt       = None
        target_sheet   = None
        saved_sheet_uid = record.get("target_sheet_unique_id")
        old_ssi_uid    = record.get("ssi_unique_id")
        if old_ssi_uid:
            try:
                old_ssi = self._doc.GetElement(str(old_ssi_uid))
                if old_ssi is not None:
                    raw_pt = old_ssi.Point
                    saved_pt = XYZ(raw_pt.X, raw_pt.Y, raw_pt.Z)
                    owner_s = self._doc.GetElement(old_ssi.OwnerViewId)
                    if owner_s is not None and isinstance(owner_s, ViewSheet):
                        saved_sheet_uid = owner_s.UniqueId
                        target_sheet = owner_s
            except Exception as ex:
                LogManager.debug("[DocLinkManager] _on_sched_edit: SSI read failed: {}".format(ex))

        if target_sheet is None and saved_sheet_uid:
            try:
                ts = self._doc.GetElement(str(saved_sheet_uid))
                if ts is not None and isinstance(ts, ViewSheet):
                    target_sheet = ts
            except Exception:
                pass
        if target_sheet is None:
            target_sheet = _get_active_sheet(self._uidoc)

        old_sched_uid = record.get("schedule_unique_id")

        # Phase 25/Capture: Extract current widths BEFORE deletion so
        # width_snapshot survives the rebuild -- BUT only when col_scale has
        # NOT changed.  If the user explicitly changed the column-width scale
        # we must NOT restore the old widths; we let _build_schedule_from_record
        # compute fresh widths from the new col_scale instead.
        retain_logic = _as_bool(record.get("retain_settings", True))
        if retain_logic and old_sched_uid and not _col_changed:
            try:
                old_sched_el = self._doc.GetElement(str(old_sched_uid))
                if old_sched_el is not None:
                    fresh_widths = _extract_column_width_snapshot(self._doc, old_sched_el)
                    if fresh_widths:
                        record["width_snapshot"] = fresh_widths
                        LogManager.info("[DocLink] _on_sched_edit: captured {} col widths (scale unchanged)".format(
                            len(fresh_widths)))
            except Exception as _w_ex:
                LogManager.debug("[DocLink] _on_sched_edit: width capture failed: {}".format(_w_ex))
        elif _col_changed:
            LogManager.info("[DocLink] _on_sched_edit: col_scale changed -- skipping width recapture, fresh widths from new scale")

        # Delete BOTH old SSI AND old schedule view safely.
        # Use individual Transactions (no TransactionGroup) to avoid
        # Revit fatal errors from nested transaction contexts in WPF handlers.
        for _uid_edit, _lbl_edit in [
            (old_ssi_uid,   "SSI (Edit)"),
            (old_sched_uid, "Schedule (Edit)"),
        ]:
            if not _uid_edit:
                continue
            try:
                _old_el = self._doc.GetElement(str(_uid_edit))
                if _old_el is not None:
                    with Transaction(
                        self._doc,
                        "Schedule Import – Delete Old {}".format(_lbl_edit)
                    ) as _t:
                        _t.Start()
                        self._doc.Delete(_old_el.Id)
                        _t.Commit()
                    LogManager.debug("[DocLink] _on_sched_edit: deleted old {} ({})".format(
                        _lbl_edit, _uid_edit))
            except Exception as _del_ex:
                LogManager.debug("[DocLink] _on_sched_edit: delete {} failed: {}".format(
                    _lbl_edit, _del_ex))

        try:
            result = _build_schedule_from_record(
                self._doc, self._uidoc, record,
                placement_pt=saved_pt,
                target_sheet=target_sheet,
            )
        except Exception as _ex:
            LogManager.error("✗ UNHANDLED EXCEPTION in _on_sched_edit/_build_schedule_from_record")
            LogManager.error("Exception: {}".format(str(_ex)))
            LogManager.error("Full traceback:")
            for line in traceback.format_exc().split("\n"):
                if line.strip():
                    LogManager.error(line)
            LogManager.debug("[DocLink] _on_sched_edit exception traceback:\n{}".format(traceback.format_exc()))
            forms.alert("Re-import failed:\n{}".format(str(_ex)))
            record["status"] = "Error"
            save_schedule_records(self._doc, self._sched_records)
            self._sched_refresh()
            return

        record["schedule_unique_id"]     = result.get("schedule_unique_id")
        record["ssi_unique_id"]          = result.get("ssi_unique_id")
        record["merge_snapshot"]         = result.get("merge_snapshot")
        record["border_snapshot"]        = result.get("border_snapshot")
        record["width_snapshot"]         = result.get("width_snapshot")
        record["target_sheet_unique_id"] = saved_sheet_uid
        _new_ssi_uid = result.get("ssi_unique_id")
        if _new_ssi_uid:
            try:
                _new_ssi = self._doc.GetElement(str(_new_ssi_uid))
                if _new_ssi is not None:
                    _raw_pt = _new_ssi.Point
                    record["last_placement_point"] = [_raw_pt.X, _raw_pt.Y, _raw_pt.Z]
                    LogManager.debug("[DocLink] _on_sched_edit: new SSI at ({:.4f}, {:.4f})".format(
                        _raw_pt.X, _raw_pt.Y))
            except Exception:
                pass
        record["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        record["status"] = "OK" if _new_ssi_uid else "Not on Sheet"

        save_schedule_records(self._doc, self._sched_records)
        self._sched_refresh()
        self._status.Content = "✔  '{}' updated with new setup.".format(
            record.get("schedule_name", "?"))


# ─────────────────────────────────────────────────────────────────────────────
# SCHEDULE SETUP DIALOG  (Phase 6 – Tab 2 editable setup)
# ─────────────────────────────────────────────────────────────────────────────



def main():
    LogManager.section("DOCLINK TOOL LAUNCHED")
    LogManager.info("Python version: {}".format(sys.version))
    LogManager.info("Log file: {}".format(LogManager.get_log_path()))
    LogManager.info("PyMuPDF available: {}".format(_FITZ_AVAILABLE))

    # Show pyrevit output window for real-time logging
    _output = None
    try:
        from pyrevit import script as _script
        _output = _script.get_output()
        # Connect pyRevit output to logger for real-time display
        LogManager.set_pyrevit_output(_output)
        # Display header in pyRevit window
        _output.print_md("# **DocLink Transparent Import Tool**")
        _output.print_md("Logging real-time to disk: `{}`\n".format(LogManager.get_log_path()))
        _output.print_md("**Note:** All operations are logged to disk for debugging.\n")
        _output.print_md("---\n")
    except Exception as e:
        LogManager.debug("Could not open pyRevit output: {}".format(e))

    doc  = revit.doc
    uidoc = revit.uidoc
    view = revit.active_view
    LogManager.debug("Active view: {}".format(view.Name if view else "None"))

    if view is None:
        LogManager.error("✗ No active view found")
        forms.alert("No active view. Please open a view before launching this tool.")
        return

    LogManager.info("Opening DocLink Manager window...")
    win = DocLinkManagerWindow(doc, view, uidoc=uidoc)
    win.ShowDialog()
    LogManager.section("DOCLINK TOOL CLOSED")


if __name__ == "__main__":
    main()
