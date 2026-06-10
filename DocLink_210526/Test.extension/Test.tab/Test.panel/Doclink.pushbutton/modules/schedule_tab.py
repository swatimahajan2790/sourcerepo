# -*- coding: utf-8 -*-
"""
schedule_tab.py
---------------
Tab 2 – Schedule Import engine.

Contains
--------
ScheduleSetupDialog             – WinForms settings dialog for schedule records
Excel XML parser                – parse_excel, list_sheets, build_grid
Revit schedule builder          – create_generic_model_schedule, write_header
Cell/border styling             – _apply_style, _restamp_borders_only, BorderGraphicsStyles
Snapshot utilities              – _extract_merge_snapshot, _apply_column_width_snapshot, …
Schedule persistence helpers    – _build_schedule_from_record, run_excel_header_mapper
"""

import os, sys, re, json, datetime, tempfile, zipfile
from datetime import date as _date, timedelta as _timedelta
import xml.etree.ElementTree as ET

from _imports import (
    System, clr,
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    Visibility, GridLength, GridUnitType, FontWeights,
    Grid, RowDefinition, ColumnDefinition, StackPanel, WrapPanel, ScrollViewer,
    Button, Label, TextBox, TextBlock, ComboBox, ComboBoxItem, CheckBox,
    Separator, DataGrid, DataGridTextColumn, DataGridSelectionMode,
    WPFOrientation,
    Binding, SolidColorBrush, Color, Brushes, ObservableCollection,
    INotifyPropertyChanged, PropertyChangedEventArgs,
    WinForms, WFForm, WFLabel, WFTextBox, WFButton, WFCheckBox, WFListBox,
    WFOpenFileDialog, WFDialogResult, WFFormBorderStyle, WFFormStartPosition,
    WFSelectionMode, WFBorderStyle, WFNumericUpDown, WFGroupBox,
    WFComboBox, WFComboBoxStyle,
    WFDialogResultEnum, WFRadioButton,
    Bitmap, Graphics, DrawingColor, DColor, ImageFormat, PixelFormat,
    DSize, DPoint, DFont,
    Array, Guid, String, Marshal,
    Transaction, TransactionGroup,
    ImageType, ImageTypeOptions, ImageTypeSource,
    ImagePlacementOptions, ImageInstance, XYZ, BoxPlacement,
    BuiltInParameter, ElementId, FilteredElementCollector,
    SectionType, TableMergedCell, TableCellStyle,
    TableCellStyleOverrideOptions, RevitColor,
    ViewSchedule, GraphicsStyle, ScheduleFieldType,
    ScheduleFilter, ScheduleFilterType,
    BuiltInCategory, CellType, ViewSheet, ScheduleSheetInstance,
    Schema, SchemaBuilder, Entity, AccessLevel,
    TaskDialog, TaskDialogCommonButtons,
    _HAS_H, _HAS_V, _HAS_TNT, TextNoteType,
    DataStorageType,
    apply_template_to_window, _TMPL_PATH, _RES,
    forms, revit,
    _EXCEL_AVAILABLE,
)

try:
    from Autodesk.Revit.DB import HorizontalAlignmentStyle
except ImportError:
    pass
try:
    from Autodesk.Revit.DB import VerticalAlignmentStyle
except ImportError:
    pass

from logger import LogManager
from utils import (
    _safe_int, _as_bool, _combo_selected_text,
    SUPPORTED_FILES_FILTER,
    _sanitize_import_name, _get_auto_user,
)
from models import ScheduleRow
from persistence import (
    load_schedule_records, save_schedule_records,
    _normalize_sched_records,
)
from excel_ops import (
    _read_excel_display_values,
    _excel_col_index_to_letters, _bounds_to_excel_range,
    list_excel_defined_names, resolve_defined_name,
    get_excel_print_area,
    _ensure_xlsx_format,
)
from revit_ops import (
    _element_by_unique_id, _normalize_unique_ids,
)

def _show_resize_only_dialog(parent, row_scale, col_scale, text_scale, col_disabled):
    """
    Compact WinForms popup with just the three scale spinners (row / column
    width / text). Column spinner is greyed when `col_disabled` is True (=
    Retain Layout is on in the parent — widths come from snapshot anyway).

    Returns a dict {'row_scale','col_scale','text_scale'} on Apply, else None.
    """
    import System.Windows.Forms as _WF
    from System import Decimal as System_dec

    form = WFForm()
    form.Text                = "Resize Schedule"
    form.FormBorderStyle     = WFFormBorderStyle.FixedDialog
    form.MaximizeBox         = False
    form.MinimizeBox         = False
    form.StartPosition       = WFFormStartPosition.CenterParent
    form.ClientSize          = DSize(340, 200)
    form.AcceptButton        = None
    form.CancelButton        = None

    def _spin(val, mn, mx, dec, gx, gy):
        n = WFNumericUpDown()
        n.Value         = System_dec(val)
        n.Minimum       = System_dec(mn)
        n.Maximum       = System_dec(mx)
        n.DecimalPlaces = dec
        n.Increment     = System_dec(0.05) if dec > 0 else System_dec(5)
        n.Location      = DPoint(gx, gy)
        n.Size          = DSize(80, 24)
        n.Font          = DFont("Segoe UI", 9)
        form.Controls.Add(n)
        return n

    def _lbl(text, gy, color=None):
        l = WFLabel()
        l.Text     = text
        l.Location = DPoint(14, gy + 4)
        l.Size     = DSize(180, 20)
        l.Font     = DFont("Segoe UI", 9)
        if color is not None:
            l.ForeColor = color
        form.Controls.Add(l)
        return l

    y = 16
    _lbl("Row Height Scale (0.1–2.0):", y)
    sp_row  = _spin(row_scale, 0.10, 2.00, 2, 220, y); y += 32

    _lbl("Column Width Scale (0.1–3.0):", y,
         color=(DColor.FromArgb(140, 140, 140) if col_disabled else None))
    sp_col  = _spin(col_scale, 0.10, 3.00, 2, 220, y)
    sp_col.Enabled = not col_disabled
    if col_disabled:
        hint = WFLabel()
        hint.Text     = "(disabled — Retain Layout is on)"
        hint.Location = DPoint(14, y + 26)
        hint.Size     = DSize(310, 16)
        hint.Font     = DFont("Segoe UI Italic", 8)
        hint.ForeColor = DColor.FromArgb(110, 110, 110)
        form.Controls.Add(hint)
        y += 18
    y += 32

    _lbl("Text Size Scale (0.1–3.0):", y)
    sp_text = _spin(text_scale, 0.10, 3.00, 2, 220, y); y += 40

    # Resize body to fit content.
    form.ClientSize = DSize(340, y + 50)

    result_holder = {'val': None}

    btn_apply = WFButton()
    btn_apply.Text     = "Apply Resize"
    btn_apply.Size     = DSize(120, 28)
    btn_apply.Location = DPoint(form.ClientSize.Width - 14 - 120 - 90, y)
    def _on_apply(s, e):
        result_holder['val'] = {
            'row_scale':  float(sp_row.Value),
            'col_scale':  float(sp_col.Value),
            'text_scale': float(sp_text.Value),
        }
        form.DialogResult = WFDialogResultEnum.OK
        form.Close()
    btn_apply.Click += _on_apply
    form.Controls.Add(btn_apply)
    form.AcceptButton = btn_apply

    btn_cancel = WFButton()
    btn_cancel.Text     = "Cancel"
    btn_cancel.Size     = DSize(80, 28)
    btn_cancel.Location = DPoint(form.ClientSize.Width - 14 - 80, y)
    btn_cancel.DialogResult = WFDialogResultEnum.Cancel
    form.Controls.Add(btn_cancel)
    form.CancelButton = btn_cancel

    if parent is not None:
        try:
            form.Owner = parent
        except Exception:
            pass
    form.ShowDialog()
    return result_holder['val']


class ScheduleSetupDialog(object):
    def __init__(self, record=None, default_name=None, existing_names=None, 
                 filepath=None, sheet_name=None, doc=None, scan_results=None,
                 initial_output_type=None):
        self.result = None
        self._resize_only_requested = False
        self._record = record or {}
        self._opts = self._record.get("options") or {}
        
        self._default_name = default_name or self._record.get("schedule_name") or ""
        self._existing_names = existing_names or set()
        
        self._filepath = filepath or self._record.get("source_path") or ""
        self._sheet_name = sheet_name or self._record.get("sheet_name") or ""
        self._doc = doc
        
        _scan = scan_results or {}
        self._scan_fonts   = _scan.get('fonts', None)
        self._scan_fills   = _scan.get('fills', None)
        self._scan_borders = _scan.get('borders', None)
        
        self._initial_output_type = initial_output_type or self._record.get("output_type") or self._opts.get("output_type") or "Schedule"

        self.line_map_holder = [self._opts.get("line_style_map", {})]
        self.text_map_holder = [self._opts.get("text_note_map", {})]
        self.fill_map_holder = [self._opts.get("fill_region_map", {})]
        
        self._form = self._build()

    def _build(self):
        # Width matches the reference module (600 px).
        W   = 600
        PAD = 14
        BW, BH = 100, 30

        try:
            from System import Decimal as System_dec
        except ImportError:
            System_dec = float

        import System.Windows.Forms as _WF
        from System import Array

        form = WFForm()
        if self._record and self._record.get("schedule_name"):
            form.Text = "DocLink  –  Edit Import Setup"
        else:
            form.Text = "DocLink  –  New Import Setup"
        # Use FixedDialog like the reference — simple, reliable button layout.
        form.FormBorderStyle = WFFormBorderStyle.FixedDialog
        form.MaximizeBox     = False
        form.MinimizeBox     = False
        form.StartPosition   = WFFormStartPosition.CenterScreen

        y = PAD

        # ── 1. Name & Calibration ───────────────────────────────────────── #
        lbl_name = WFLabel()
        lbl_name.Text     = "Import Name:"
        lbl_name.Location = DPoint(PAD, y)
        lbl_name.Size     = DSize(180, 16)
        lbl_name.Font     = DFont("Segoe UI", 9)
        form.Controls.Add(lbl_name)
        
        txt_name = WFTextBox()
        txt_name.Text     = self._default_name
        txt_name.Location = DPoint(PAD + 184, y - 2)
        txt_name.Size     = DSize(W - PAD * 2 - 184, 24)
        txt_name.Font     = DFont("Segoe UI", 9)
        form.Controls.Add(txt_name)
        y += 28
        
        self.err_lbl = WFLabel()
        self.err_lbl.Text      = ""
        self.err_lbl.Location  = DPoint(PAD, y)
        self.err_lbl.Size      = DSize(W - PAD * 2, 16)
        self.err_lbl.Font      = DFont("Segoe UI", 8)
        self.err_lbl.ForeColor = DColor.FromArgb(200, 0, 0)
        form.Controls.Add(self.err_lbl);  y += 20
        
        lbl_src = WFLabel()
        lbl_src.Text     = "Source File (Excel/Word):"
        lbl_src.Location = DPoint(PAD, y)
        lbl_src.Size     = DSize(W - PAD * 2, 16)
        lbl_src.Font     = DFont("Segoe UI", 9)
        form.Controls.Add(lbl_src); y += 20
        
        txt_src = WFTextBox()
        txt_src.Text     = self._filepath
        txt_src.Location = DPoint(PAD, y)
        txt_src.Size     = DSize(W - PAD * 2 - 235, 24)
        txt_src.Font     = DFont("Segoe UI", 9)
        txt_src.ReadOnly = True
        txt_src.BackColor = DColor.FromArgb(245, 247, 250)
        form.Controls.Add(txt_src)

        btn_browse = WFButton()
        btn_browse.Text = "Browse…"
        btn_browse.Location = DPoint(W - PAD - 225, y - 1)
        btn_browse.Size = DSize(75, 26)
        btn_browse.Font = DFont("Segoe UI", 9)

        def _on_browse(s, e):
            dlg = WFOpenFileDialog()
            dlg.Filter = SUPPORTED_FILES_FILTER
            if dlg.ShowDialog() == WFDialogResult.OK:
                self._filepath = dlg.FileName
                txt_src.Text = self._filepath
                cmb_sheet.Items.Clear()
                if self._filepath and not self._filepath.lower().endswith(('.doc', '.docx')):
                    try:
                        sheets = list_sheets(self._filepath)
                        if sheets:
                            for sh in sheets:
                                cmb_sheet.Items.Add(sh['name'])
                            if cmb_sheet.Items.Count > 0:
                                cmb_sheet.SelectedIndex = 0
                    except Exception: pass
        btn_browse.Click += _on_browse
        form.Controls.Add(btn_browse)
        
        cmb_path_type = _WF.ComboBox()
        cmb_path_type.DropDownStyle = _WF.ComboBoxStyle.DropDownList
        cmb_path_type.Items.AddRange(Array[str](["Absolute", "Relative"]))
        cmb_path_type.Location = DPoint(W - PAD - 140, y)
        cmb_path_type.Size     = DSize(140, 24)
        cmb_path_type.Font     = DFont("Segoe UI", 9)
        current_pt = self._record.get("path_type", "Absolute")
        cmb_path_type.SelectedIndex = 1 if current_pt == "Relative" else 0
        form.Controls.Add(cmb_path_type)
        y += 30
        
        lbl_sh = WFLabel()
        lbl_sh.Text     = "Worksheet / Word Page:"
        lbl_sh.Location = DPoint(PAD, y)
        lbl_sh.Size     = DSize(W - PAD * 2, 16)
        lbl_sh.Font     = DFont("Segoe UI", 9)
        form.Controls.Add(lbl_sh); y += 20
        
        cmb_sheet = _WF.ComboBox()
        cmb_sheet.DropDownStyle = _WF.ComboBoxStyle.DropDown
        cmb_sheet.Location = DPoint(PAD, y)
        cmb_sheet.Size     = DSize(W - PAD * 2, 24)
        cmb_sheet.Font     = DFont("Segoe UI", 9)
        form.Controls.Add(cmb_sheet); y += 34
        
        current_sh = self._sheet_name
        cmb_sheet.Text = current_sh
        # Cache sheet metadata (incl. print_area) for fast Manual-mode auto-fill.
        self._sched_sheet_data = {}
        if self._filepath and not self._filepath.lower().endswith(('.doc', '.docx')):
            try:
                sheets = list_sheets(self._filepath)
                if sheets:
                    self._sched_sheet_data = {sh['name']: sh for sh in sheets}
                    for sh in sheets:
                        cmb_sheet.Items.Add(sh['name'])
            except Exception: pass

        saved_named  = (self._opts.get('named_range')  if self._opts else '') or ''
        saved_manual = (self._opts.get('manual_range') if self._opts else '') or ''
        is_excel_src = bool(self._filepath) and not self._filepath.lower().endswith(('.doc', '.docx'))

        # Range Source controls are added to page_sched BELOW, after it is created.


        grp = WFGroupBox()
        grp.Text     = "Scale & Calibration"
        grp.Location = DPoint(PAD, y)
        grp.Size     = DSize(W - PAD * 2, 130)
        grp.Font     = DFont("Segoe UI", 8)
        form.Controls.Add(grp)
        
        LW = 170
        NW = 72
        GY = 22
        
        def _lbl(text, gx, gy, ctl=grp):
            l = WFLabel()
            l.Text     = text
            l.Location = DPoint(gx, gy)
            l.Size     = DSize(LW, 18)
            l.Font     = DFont("Segoe UI", 8)
            ctl.Controls.Add(l)

        def _spin(val, mn, mx, dec, gx, gy, ctl=grp):
            n = WFNumericUpDown()
            n.Value         = System_dec(val)
            n.Minimum       = System_dec(mn)
            n.Maximum       = System_dec(mx)
            n.DecimalPlaces = dec
            n.Increment     = System_dec(0.05) if dec > 0 else System_dec(5)
            n.Location      = DPoint(gx, gy)
            n.Size          = DSize(NW, 22)
            n.Font          = DFont("Segoe UI", 8)
            ctl.Controls.Add(n)
            return n

        _lbl("Row Height Scale (0.1–2.0):",   8, GY)
        spin_row = _spin(self._opts.get("row_scale",  1.00), 0.10, 2.00, 2, LW + 14, GY); GY += 32

        # Label for Column Width Scale is kept as a local so the retain
        # checkbox below can grey it out alongside the spinner.
        lbl_col_scale = WFLabel()
        lbl_col_scale.Text     = "Column Width Scale (0.1–3.0):"
        lbl_col_scale.Location = DPoint(8, GY)
        lbl_col_scale.Size     = DSize(LW, 18)
        lbl_col_scale.Font     = DFont("Segoe UI", 8)
        grp.Controls.Add(lbl_col_scale)
        spin_col = _spin(self._opts.get("col_scale",  1.50), 0.10, 3.00, 2, LW + 14, GY); GY += 32

        # Hint shown when retain is on — explains why col_scale is greyed.
        lbl_col_retain_hint = WFLabel()
        lbl_col_retain_hint.Text      = "(retained from snapshot)"
        lbl_col_retain_hint.Location  = DPoint(LW + 14 + 78, GY - 30)
        lbl_col_retain_hint.Size      = DSize(150, 18)
        lbl_col_retain_hint.Font      = DFont("Segoe UI Italic", 8)
        lbl_col_retain_hint.ForeColor = DColor.FromArgb(110, 110, 110)
        lbl_col_retain_hint.Visible   = False
        grp.Controls.Add(lbl_col_retain_hint)

        _lbl("Text Size Scale (0.1–3.0):",     8, GY)
        spin_text = _spin(self._opts.get("text_scale", 1.00), 0.10, 3.00, 2, LW + 14, GY); GY += 32
        
        y += 140

        
        # ── 2. Tab Control ─────────────────────────────────────────────────── #
        tab_ctrl = _WF.TabControl()
        tab_ctrl.Location = DPoint(PAD, y)
        # 260 (was 230): leaves room for the Schedule sub-tab's Manual/Named
        # rows that now stack instead of overlap, plus the Configure button.
        tab_ctrl.Size     = DSize(W - PAD * 2, 260)
        tab_ctrl.Font     = DFont("Segoe UI", 9)
        # Width-anchor only — the form is FixedDialog, so a Bottom anchor
        # combined with the later ClientSize assignment would stretch the
        # TabControl downward, covering the OK/Cancel buttons.
        try:
            tab_ctrl.Anchor = (_WF.AnchorStyles.Top | _WF.AnchorStyles.Left
                               | _WF.AnchorStyles.Right)
        except Exception:
            pass
        form.Controls.Add(tab_ctrl)
        
        page_sched = _WF.TabPage("Schedule")
        page_draft = _WF.TabPage("Drafting View")
        page_legend = _WF.TabPage("Legend")

        tab_ctrl.TabPages.Add(page_sched)
        tab_ctrl.TabPages.Add(page_draft)
        tab_ctrl.TabPages.Add(page_legend)

        # ── Range Source controls: now page_sched exists ────────────────────
        sy = 6

        lbl_range_mode = WFLabel()
        lbl_range_mode.Text     = "Range Source:"
        lbl_range_mode.Location = DPoint(14, sy)
        lbl_range_mode.Size     = DSize(W - 40, 16)
        lbl_range_mode.Font     = DFont("Segoe UI", 9)
        page_sched.Controls.Add(lbl_range_mode); sy += 18

        # Two-mode toggle, matching the Import Image dialog:
        #   Manual (default) — pick sheet; textbox auto-fills with the
        #     sheet's Print Area (if defined) and is user-editable.
        #   Named Range      — pick a defined name; sheet and range are
        #     driven by the name and are greyed out as read-only display.
        is_named_default = bool(saved_named) and is_excel_src

        rb_manual = WFRadioButton()
        rb_manual.Text     = "Manual Range"
        rb_manual.Location = DPoint(14, sy)
        rb_manual.Size     = DSize(140, 22)
        rb_manual.Font     = DFont("Segoe UI", 9)
        rb_manual.Checked  = (not is_named_default)
        rb_manual.Enabled  = is_excel_src
        page_sched.Controls.Add(rb_manual)

        rb_named = WFRadioButton()
        rb_named.Text     = "Named Range"
        rb_named.Location = DPoint(14 + 160, sy)
        rb_named.Size     = DSize(140, 22)
        rb_named.Font     = DFont("Segoe UI", 9)
        rb_named.Checked  = is_named_default
        rb_named.Enabled  = is_excel_src
        page_sched.Controls.Add(rb_named); sy += 26

        cmb_named = _WF.ComboBox()
        cmb_named.DropDownStyle = _WF.ComboBoxStyle.DropDownList
        cmb_named.Location = DPoint(14, sy)
        cmb_named.Size     = DSize(W - 40, 24)
        cmb_named.Font     = DFont("Segoe UI", 9)
        cmb_named.Enabled  = rb_named.Checked
        page_sched.Controls.Add(cmb_named); sy += 28

        txt_manual = WFTextBox()
        txt_manual.Text     = saved_manual or ""
        txt_manual.Location = DPoint(14, sy)
        txt_manual.Size     = DSize(W - 40, 24)
        txt_manual.Font     = DFont("Segoe UI", 9)
        txt_manual.Enabled  = rb_manual.Checked
        page_sched.Controls.Add(txt_manual); sy += 26

        lbl_named_hint = WFLabel()
        lbl_named_hint.Text      = ""
        lbl_named_hint.Location  = DPoint(16, sy)
        lbl_named_hint.Size      = DSize(W - 44, 16)
        lbl_named_hint.Font      = DFont("Segoe UI", 8)
        lbl_named_hint.ForeColor = DColor.FromArgb(110, 110, 110)
        page_sched.Controls.Add(lbl_named_hint); sy += 20

        self._sched_named_state = {'items': [], 'suppress_sheet_sync': False, 'initializing': True}

        def _populate_named_for_sheet(sheet_for_scope):
            cmb_named.BeginUpdate()
            try:
                cmb_named.Items.Clear()
                self._sched_named_state['items'] = []
                if not is_excel_src:
                    return False
                try:
                    items = list_excel_defined_names(self._filepath, sheet_for_scope)
                except Exception:
                    items = []
                if not items:
                    return False
                preferred_idx = 0
                for i, n in enumerate(items):
                    scope_tag = ("Workbook" if n['is_workbook_scope']
                                 else "Sheet: {0}".format(n.get('scope_sheet') or '?'))
                    label = "{0}    ({1}  \u2014  {2})".format(n['name'], n['range_address'], scope_tag)
                    cmb_named.Items.Add(label)
                    self._sched_named_state['items'].append(n)
                    if saved_named and n['name'] == saved_named:
                        preferred_idx = i
                cmb_named.SelectedIndex = preferred_idx
                return True
            finally:
                cmb_named.EndUpdate()

        def _update_hint_from_selection():
            idx = cmb_named.SelectedIndex
            items = self._sched_named_state['items']
            if 0 <= idx < len(items):
                info = items[idx]
                scope = "workbook-scoped" if info['is_workbook_scope'] else "sheet-scoped"
                lbl_named_hint.Text = "\u2192 {0}!{1}  ({2})".format(
                    info.get('ref_sheet') or '', info.get('range_address') or '', scope)
                # Mirror the resolved range into the (disabled-in-Named-mode)
                # manual textbox so the user always sees what'll be exported.
                rng = info.get('range_address') or ''
                if rng:
                    txt_manual.Text = rng
            else:
                lbl_named_hint.Text = ""

        def _sync_sheet_combo_to(name):
            if not name: return
            self._sched_named_state['suppress_sheet_sync'] = True
            try:
                found = False
                for i in range(cmb_sheet.Items.Count):
                    if str(cmb_sheet.Items[i]) == name:
                        cmb_sheet.SelectedIndex = i; found = True; break
                if not found:
                    cmb_sheet.Text = name
            finally:
                self._sched_named_state['suppress_sheet_sync'] = False

        def _autofill_manual_from_print_area():
            """Manual mode: pre-fill the manual textbox with the active
            sheet's Print Area. Cheap OpenXML cache lookup, COM fallback."""
            name = (cmb_sheet.Text or '').strip()
            if not name:
                return
            sh = (self._sched_sheet_data.get(name)
                  if self._sched_sheet_data else None)
            pa = sh.get('print_area') if sh else None
            if not pa:
                try:
                    pa = get_excel_print_area(self._filepath, name)
                except Exception:
                    pa = None
            if pa:
                txt_manual.Text = pa

        def _on_sheet_changed(s, e):
            if self._sched_named_state.get('initializing'): return
            if self._sched_named_state['suppress_sheet_sync']: return
            had_any = _populate_named_for_sheet(cmb_sheet.Text or None)
            if not had_any:
                rb_named.Enabled = False; cmb_named.Enabled = False
                if rb_named.Checked:
                    rb_manual.Checked = True
                lbl_named_hint.Text = ""
            else:
                rb_named.Enabled = True
                cmb_named.Enabled = bool(rb_named.Checked)
                if rb_named.Checked: _update_hint_from_selection()
            # Manual mode: refresh textbox with the new sheet's Print Area.
            if rb_manual.Checked:
                _autofill_manual_from_print_area()

        def _on_named_picked(s, e):
            if self._sched_named_state.get('initializing'): return
            if not rb_named.Checked: return
            _update_hint_from_selection()
            idx = cmb_named.SelectedIndex
            items = self._sched_named_state['items']
            if 0 <= idx < len(items):
                tgt = items[idx].get('ref_sheet')
                if tgt and tgt != cmb_sheet.Text: _sync_sheet_combo_to(tgt)

        def _on_mode_changed(s, e):
            if self._sched_named_state.get('initializing'): return
            use_named  = rb_named.Checked  and rb_named.Enabled
            use_manual = rb_manual.Checked and rb_manual.Enabled
            cmb_named.Enabled = use_named and cmb_named.Items.Count > 0
            # In Named mode the chosen name drives the sheet AND the range —
            # grey both inputs to make it visually obvious.
            cmb_sheet.Enabled = not use_named
            txt_manual.Enabled = use_manual
            if use_named:
                _update_hint_from_selection(); _on_named_picked(None, None)
            elif use_manual:
                lbl_named_hint.Text = (
                    "Manual range — auto-filled from Print Area; edit freely.")
                _autofill_manual_from_print_area()
            else:
                lbl_named_hint.Text = ""

        cmb_sheet.TextChanged          += _on_sheet_changed
        cmb_named.SelectedIndexChanged += _on_named_picked
        rb_named.CheckedChanged        += _on_mode_changed
        rb_manual.CheckedChanged       += _on_mode_changed

        if is_excel_src:
            had_any = _populate_named_for_sheet(current_sh or None)
            if not had_any:
                rb_named.Enabled = False; rb_named.Checked = False
                rb_manual.Checked = True; cmb_named.Enabled = False
            elif rb_named.Checked:
                _update_hint_from_selection()
            # Manual default: pre-fill textbox from the sheet's Print Area
            # (only when the saved record didn't already supply one).
            if rb_manual.Checked and not saved_manual:
                _autofill_manual_from_print_area()
            # Named-mode initial state: grey sheet picker and manual textbox.
            if rb_named.Checked:
                cmb_sheet.Enabled = False
                txt_manual.Enabled = False
        self._sched_named_state['initializing'] = False
        self._sched_rb_named = rb_named; self._sched_cmb_named = cmb_named
        self._sched_rb_manual = rb_manual; self._sched_txt_manual = txt_manual

        # ── Checkboxes below range source ───────────────────────────────────
        def _chk(text, checked, yy, page):
            c = WFCheckBox()
            c.Text = text; c.Checked = checked
            c.Location = DPoint(14, yy)
            c.Size = DSize(tab_ctrl.Width - 40, 20)
            c.Font = DFont("Segoe UI", 9)
            page.Controls.Add(c); return c

        chk_merge     = _chk("Enable Dynamic Merging", _as_bool(self._opts.get("enable_dynamic_merge", True)), sy, page_sched); sy += 22
        chk_gridlines = _chk("Show Grid Lines (empty borders)", _as_bool(self._opts.get("show_empty_gridlines", False)), sy, page_sched); sy += 22
        chk_retain    = _chk("Retain Layout on Re-import (Merges, Borders, Widths)", _as_bool(self._record.get("retain_settings", True) if self._record else True), sy, page_sched); sy += 22

        def _apply_retain_to_col_scale():
            # New imports should always allow the user to set the initial
            # column scale. Lock it only in edit mode when Retain is on.
            on = bool(self._record) and bool(chk_retain.Checked)
            spin_col.Enabled = not on
            lbl_col_scale.Enabled = not on
            lbl_col_retain_hint.Visible = on

        chk_retain.CheckedChanged += lambda s, e: _apply_retain_to_col_scale()
        _apply_retain_to_col_scale()


        btn_line_map_sched = WFButton()
        btn_line_map_sched.Text     = u'\u2699 Configure Line Styles\u2026'
        btn_line_map_sched.Location = DPoint(14, sy)
        btn_line_map_sched.Size     = DSize(200, 26)
        btn_line_map_sched.Font     = DFont('Segoe UI', 9)
        page_sched.Controls.Add(btn_line_map_sched)
        
        lbl_line_status_sched = WFLabel()
        lbl_line_status_sched.Text      = '{} style(s) mapped'.format(len([v for v in self.line_map_holder[0].values() if v])) if self.line_map_holder[0] else 'Not configured'
        lbl_line_status_sched.Location  = DPoint(220, sy + 4)
        lbl_line_status_sched.Size      = DSize(W - PAD * 2 - 230, 18)
        lbl_line_status_sched.Font      = DFont('Segoe UI Italic', 8)
        lbl_line_status_sched.ForeColor = DColor.FromArgb(100, 100, 100) if not self.line_map_holder[0] else DColor.FromArgb(0, 120, 0)
        page_sched.Controls.Add(lbl_line_status_sched)
        
        def _on_line_map_sched(s, e):
            try:
                xl_styles  = self._scan_borders if self._scan_borders is not None else _scan_excel_border_styles(self._filepath, cmb_sheet.Text)
                rv_styles  = _get_project_line_style_names(self._doc)
                dlg = LineStyleMappingDialog(xl_styles, rv_styles, self.line_map_holder[0])
                if dlg.ShowDialog() == WFDialogResultEnum.OK:
                    self.line_map_holder[0] = dlg.result or {}
                    n = len([v for v in self.line_map_holder[0].values() if v])
                    lbl_line_status_sched.Text      = '{} style(s) mapped'.format(n)
                    lbl_line_status_sched.ForeColor = DColor.FromArgb(0, 120, 0)
            except Exception as _ex:
                pass
                
        btn_line_map_sched.Click += _on_line_map_sched
        
        def _build_graphic_page(page, output_type):
            gy = 14
            chk_scaling = _chk("Compensate Geometry for View Scale (Maintain Sheet Size)", _as_bool(self._opts.get("scale_by_view", True)), gy, page); gy += 28
            
            _lbl("Revit View Scale (1:X):", 14, gy+2, page)
            spin_vscale = _spin(self._opts.get("view_scale", 1), 1, 5000, 0, LW + 20, gy, page); gy += 28
            
            _lbl("Excel Points per 1 mm:", 14, gy+2, page)
            spin_pt_mm = _spin(self._opts.get("excel_points_per_mm", 2.834), 1.0, 10.0, 3, LW + 20, gy, page); gy += 34
            
            btn_line_map = WFButton()
            btn_line_map.Text     = u'⚙ Configure Line Styles…'
            btn_line_map.Location = DPoint(14, gy)
            btn_line_map.Size     = DSize(200, 26)
            btn_line_map.Font     = DFont('Segoe UI', 9)
            page.Controls.Add(btn_line_map)
            
            lbl_line_status = WFLabel()
            lbl_line_status.Text      = '{} style(s) mapped'.format(len([v for v in self.line_map_holder[0].values() if v])) if self.line_map_holder[0] else 'Not configured'
            lbl_line_status.Location  = DPoint(220, gy + 4)
            lbl_line_status.Size      = DSize(W - PAD * 2 - 230, 18)
            lbl_line_status.Font      = DFont('Segoe UI Italic', 8)
            lbl_line_status.ForeColor = DColor.FromArgb(100, 100, 100) if not self.line_map_holder[0] else DColor.FromArgb(0, 120, 0)
            page.Controls.Add(lbl_line_status); gy += 34
            
            btn_text_map = WFButton()
            btn_text_map.Text     = u'⚙ Configure Text Note Types…'
            btn_text_map.Location = DPoint(14, gy)
            btn_text_map.Size     = DSize(200, 26)
            btn_text_map.Font     = DFont('Segoe UI', 9)
            page.Controls.Add(btn_text_map)
            
            lbl_text_status = WFLabel()
            lbl_text_status.Text      = '{} type(s) mapped'.format(len([v for v in self.text_map_holder[0].values() if v])) if self.text_map_holder[0] else 'Not configured'
            lbl_text_status.Location  = DPoint(220, gy + 4)
            lbl_text_status.Size      = DSize(W - PAD * 2 - 230, 18)
            lbl_text_status.Font      = DFont('Segoe UI Italic', 8)
            lbl_text_status.ForeColor = DColor.FromArgb(100, 100, 100) if not self.text_map_holder[0] else DColor.FromArgb(0, 120, 0)
            page.Controls.Add(lbl_text_status); gy += 34
            
            btn_fill_map = WFButton()
            btn_fill_map.Text     = u'⚙ Configure Filled Regions…'
            btn_fill_map.Location = DPoint(14, gy)
            btn_fill_map.Size     = DSize(200, 26)
            btn_fill_map.Font     = DFont('Segoe UI', 9)
            page.Controls.Add(btn_fill_map)
            
            lbl_fill_status = WFLabel()
            lbl_fill_status.Text      = '{} colour(s) mapped'.format(len([v for v in self.fill_map_holder[0].values() if v])) if self.fill_map_holder[0] else 'Not configured'
            lbl_fill_status.Location  = DPoint(220, gy + 4)
            lbl_fill_status.Size      = DSize(W - PAD * 2 - 230, 18)
            lbl_fill_status.Font      = DFont('Segoe UI Italic', 8)
            lbl_fill_status.ForeColor = DColor.FromArgb(100, 100, 100) if not self.fill_map_holder[0] else DColor.FromArgb(0, 120, 0)
            page.Controls.Add(lbl_fill_status)
            
            def _on_line_map(s, e):
                try:
                    xl_styles  = self._scan_borders if self._scan_borders is not None else _scan_excel_border_styles(self._filepath, cmb_sheet.Text)
                    rv_styles  = _get_project_line_style_names(self._doc)
                    dlg = LineStyleMappingDialog(xl_styles, rv_styles, self.line_map_holder[0])
                    if dlg.ShowDialog() == WFDialogResultEnum.OK:
                        self.line_map_holder[0] = dlg.result or {}
                        n = len([v for v in self.line_map_holder[0].values() if v])
                        lbl_line_status.Text      = '{} style(s) mapped'.format(n)
                        lbl_line_status.ForeColor = DColor.FromArgb(0, 120, 0)
                except Exception as _ex:
                    pass
            
            def _on_text_map(s, e):
                try:
                    xl_fonts   = self._scan_fonts if self._scan_fonts is not None else _scan_excel_fonts(self._filepath, cmb_sheet.Text)
                    rv_types   = _get_project_text_note_type_names(self._doc)
                    dlg = TextNoteTypeMappingDialog(xl_fonts, rv_types, self.text_map_holder[0])
                    if dlg.ShowDialog() == WFDialogResultEnum.OK:
                        self.text_map_holder[0] = dlg.result or {}
                        n = len([v for v in self.text_map_holder[0].values() if v])
                        lbl_text_status.Text      = '{} type(s) mapped'.format(n)
                        lbl_text_status.ForeColor = DColor.FromArgb(0, 120, 0)
                except Exception as _ex:
                    pass
                    
            def _on_fill_map(s, e):
                try:
                    xl_fills   = self._scan_fills if self._scan_fills is not None else _scan_excel_fill_colors(self._filepath, cmb_sheet.Text)
                    rv_types   = _get_project_filled_region_type_names(self._doc)
                    dlg = FilledRegionMappingDialog(xl_fills, rv_types, self.fill_map_holder[0])
                    if dlg.ShowDialog() == WFDialogResultEnum.OK:
                        self.fill_map_holder[0] = dlg.result or {}
                        n = len([v for v in self.fill_map_holder[0].values() if v])
                        lbl_fill_status.Text      = '{} colour(s) mapped'.format(n)
                        lbl_fill_status.ForeColor = DColor.FromArgb(0, 120, 0)
                except Exception as _ex:
                    pass
            
            btn_line_map.Click += _on_line_map
            btn_text_map.Click += _on_text_map
            btn_fill_map.Click += _on_fill_map
            
            return {
                "chk_scaling": chk_scaling,
                "spin_vscale": spin_vscale,
                "spin_pt_mm": spin_pt_mm
            }
            
        draft_ctls = _build_graphic_page(page_draft, "Drafting View")
        legend_ctls = _build_graphic_page(page_legend, "Legend")

        # Advance past the (now 260 px) TabControl plus a 10 px gap before
        # the OK/Resize/Cancel row at the bottom of the form.
        y += 270

        if self._initial_output_type == "Drafting View":
            tab_ctrl.SelectedIndex = 1
        elif self._initial_output_type == "Legend":
            tab_ctrl.SelectedIndex = 2
        else:
            tab_ctrl.SelectedIndex = 0

        y += 10
        # Fix height to fit all content + button row (matches reference approach).
        H = y + BH + PAD
        form.ClientSize = DSize(W, H)

        btn_cancel = WFButton()
        btn_cancel.Text         = "Cancel"
        btn_cancel.Size         = DSize(BW, BH)
        btn_cancel.Location     = DPoint(W - PAD - BW, y)
        btn_cancel.DialogResult = WFDialogResultEnum.Cancel
        form.Controls.Add(btn_cancel)
        form.CancelButton = btn_cancel

        btn_ok = WFButton()
        btn_ok.Text     = "Apply & Re-import" if self._record else "Create"
        btn_ok.Size     = DSize(BW + 40, BH)
        btn_ok.Location = DPoint(W - PAD - BW * 2 - 50, y)
        form.Controls.Add(btn_ok)
        form.AcceptButton = btn_ok

        # ── Edit-mode shortcut: Resize Only ────────────────────────────────
        if self._record:
            btn_resize = WFButton()
            btn_resize.Text     = "Resize"
            btn_resize.Size     = DSize(BW + 20, BH)
            btn_resize.Location = DPoint(PAD, y)

            def _open_resize_only(_s, _e):
                vals = _show_resize_only_dialog(
                    parent=form,
                    row_scale=float(spin_row.Value),
                    col_scale=float(spin_col.Value),
                    text_scale=float(spin_text.Value),
                    col_disabled=bool(chk_retain.Checked),
                )
                if vals is None:
                    return
                try:
                    spin_row.Value  = System_dec(vals['row_scale'])
                    spin_text.Value = System_dec(vals['text_scale'])
                    if not chk_retain.Checked:
                        spin_col.Value = System_dec(vals['col_scale'])
                except Exception:
                    pass
                self._resize_only_requested = True
                on_ok(_s, _e)

            btn_resize.Click += _open_resize_only
            form.Controls.Add(btn_resize)

        def on_ok(sender, e):
            name = txt_name.Text.strip()
            if not name:
                self.err_lbl.Text = "⚠  Name cannot be empty."
                return
            if name in self._existing_names:
                self.err_lbl.Text = "⚠  A view named '{}' already exists.".format(name)
                return
                
            idx = tab_ctrl.SelectedIndex
            if idx == 0:
                output_type = "Schedule"
                g_ctls = None
            elif idx == 1:
                output_type = "Drafting View"
                g_ctls = draft_ctls
            else:
                output_type = "Legend"
                g_ctls = legend_ctls
            
            # Resolve range-source selection for persistence. Only one of
            # named_range / manual_range is non-empty at a time; default
            # (Print Area) leaves both empty.
            picked_named  = ''
            picked_manual = ''
            if self._sched_rb_named.Checked and self._sched_rb_named.Enabled:
                idx = self._sched_cmb_named.SelectedIndex
                items = self._sched_named_state['items']
                if 0 <= idx < len(items):
                    picked_named = items[idx].get('name', '') or ''
            elif self._sched_rb_manual.Checked and self._sched_rb_manual.Enabled:
                picked_manual = (self._sched_txt_manual.Text or '').strip()
                if not picked_manual:
                    self.err_lbl.Text = (
                        "⚠  Enter a cell range (e.g. A1:G25) for Manual Range mode.")
                    return

            res = {
                'name':                         name,
                # source_path is read-only in this dialog (txt_src is
                # ReadOnly), but main_window.py expects the result to
                # round-trip it so the edit flow can rebuild path_type-aware.
                'source_path':                  self._filepath,
                'row_scale':                    float(spin_row.Value),
                'col_scale':                    float(spin_col.Value),
                'text_scale':                   float(spin_text.Value),
                'max_rows':                     0,
                'enable_dynamic_merge':         chk_merge.Checked,
                'show_empty_gridlines':         chk_gridlines.Checked,
                'retain_previous_merge_layout': chk_retain.Checked,
                'retain_settings':              chk_retain.Checked,
                'base_text_size_mm':            0.0,
                'path_type':                    "Relative" if cmb_path_type.SelectedIndex == 1 else "Absolute",
                'output_type':                  output_type,
                'resize_only':                  bool(self._resize_only_requested),
                'sheet_name':                   cmb_sheet.Text,
                'named_range':                  picked_named,
                'manual_range':                 picked_manual,
                'line_style_map':               dict(self.line_map_holder[0]),
                'fill_region_map':              dict(self.fill_map_holder[0])
            }
            
            if g_ctls:
                res['view_scale']          = int(g_ctls["spin_vscale"].Value)
                res['scale_by_view']       = g_ctls["chk_scaling"].Checked
                res['excel_points_per_mm'] = float(g_ctls["spin_pt_mm"].Value)
                res['text_note_map']       = dict(self.text_map_holder[0])
                res['single_text_note_type'] = None
            else:
                res['view_scale'] = 1
                res['scale_by_view'] = True
                res['excel_points_per_mm'] = 2.834
                res['text_note_map'] = {}
                res['single_text_note_type'] = None

            self.result = res
            form.DialogResult = WFDialogResultEnum.OK
            form.Close()

        def _on_apply(sender, e):
            self._resize_only_requested = False
            on_ok(sender, e)

        btn_ok.Click += _on_apply
        return form

    def ShowDialog(self):
        dr = self._form.ShowDialog()
        return dr


# ─────────────────────────────────────────────────────────────────────────────
# UTILS
# ─────────────────────────────────────────────────────────────────────────────

def _uid():
    import uuid
    return str(uuid.uuid4())


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

# ═══════════════════════════════════════════════════════════════════════════════
# ══  EXCEL HEADER MAPPER  (business logic from ExcelHeaderMapper V24)  ════════
# ═══════════════════════════════════════════════════════════════════════════════

_NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DEFAULT_MAX_ROWS = 0

_T = {}
def _tag(name, ns=None):
    ns = ns or _NS
    key = (name, ns)
    t = _T.get(key)
    if t is None:
        t = '{%s}%s' % (ns, name)
        _T[key] = t
    return t

for _n in (
    'fill', 'patternFill', 'gradientFill', 'fgColor', 'bgColor', 'stop', 'color',
    'font', 'name', 'b', 'i', 'u', 'sz',
    'border', 'top', 'bottom', 'left', 'right',
    'numFmt', 'numFmts', 'cellXfs', 'xf', 'alignment',
    'si', 't', 'sheet', 'definedName',
    'sheetFormatPr', 'cols', 'col',
    'sheetData', 'row', 'c', 'v', 'is', 'mergeCell', 'mergeCells', 'dimension',
):
    _tag(_n)


def _col_letter_to_index(letters):
    result = 0
    for ch in letters.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def _cell_ref_to_rowcol(ref):
    ref = ref.strip().replace('$', '')
    m = re.match(r'^([A-Za-z]+)(\d+)$', ref)
    if not m:
        return None, None
    return int(m.group(2)), _col_letter_to_index(m.group(1))


def _parse_range(s):
    if not s:
        return None
    s = s.strip().strip("'")
    if '!' in s:
        s = s.split('!')[-1]
    if ',' in s:
        s = s.split(',')[0].strip()
    s = s.replace('$', '').upper().strip()
    m = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', s)
    if not m:
        return None
    return (
        _col_letter_to_index(m.group(1)),
        int(m.group(2)),
        _col_letter_to_index(m.group(3)),
        int(m.group(4)),
    )


# Date handling
_DATE_FORMAT_IDS = frozenset([
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
    45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
])
_DATE_FMT_RE = re.compile(r'[yYmMdDhHsS]')


def _is_date_format(num_fmt_id, custom_fmt_str):
    if num_fmt_id in _DATE_FORMAT_IDS:
        return True
    if custom_fmt_str:
        stripped = re.sub(r'"[^"]*"', '', custom_fmt_str)
        stripped = re.sub(r'\[[^\]]*\]', '', stripped)
        if _DATE_FMT_RE.search(stripped):
            return True
    return False


def _serial_to_date_str(serial):
    try:
        n = int(float(serial))
        if n <= 0:
            return str(serial)
        if n >= 60:
            n -= 1
        epoch = _date(1899, 12, 31)
        d = epoch + _timedelta(days=n)
        return d.strftime('%d-%b-%Y')
    except Exception:
        return str(serial)


# ─────────────────────────────────────────────────────────────────────────────
# Number-format applier — mirrors NPOI DataFormatter for the common cases.
# Used on Revit 2025+ where Cell.Text via COM isn't available and we have to
# reconstruct the displayed string from the raw <v> and the numFmt in styles.
# ─────────────────────────────────────────────────────────────────────────────

# OOXML built-in numFmtIds (spec, ECMA-376 §18.8.30). Custom formats with
# id >= 164 are looked up in <numFmts>; these are the implicit defaults.
_BUILTIN_NUM_FMTS = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yyyy h:mm',
    37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
}

# Quoted literal `"..."` or escaped `\x` — extracted before parsing digits.
_FMT_LITERAL_RE = re.compile(r'"([^"]*)"|\\(.)')
# Locale / colour brackets `[Red]`, `[$-409]`, `[h]` (the time-elapsed `[h]`
# only matters for date formats which we skip — safe to strip here).
_FMT_BRACKET_RE = re.compile(r'\[[^\]]*\]')


def _fmt_section(fmt_code, value):
    """Pick the correct sub-format from a `pos;neg;zero;text` chain."""
    if not fmt_code:
        return fmt_code, value
    parts = fmt_code.split(';')
    try:
        v = float(value)
    except (ValueError, TypeError):
        # Text → use the 4th section if present (Excel convention).
        if len(parts) >= 4:
            return parts[3], value
        return parts[0], value
    if v > 0 or len(parts) == 1:
        return parts[0], v
    if v < 0:
        # Negative section consumes the minus sign — give it the absolute
        # value so the format code's own parentheses/literals supply it.
        return (parts[1] if len(parts) >= 2 else parts[0]), abs(v)
    return (parts[2] if len(parts) >= 3 else parts[0]), v


def _split_format_pieces(section):
    """
    Strip color/locale brackets and `"..."` literals from `section`, returning
    (cleaned_code, prefix, suffix). The literals/brackets that sat before the
    first digit token go into `prefix`; those after the last digit go into
    `suffix`. Underscore-pad (`_X`) becomes a single space.
    """
    section = _FMT_BRACKET_RE.sub('', section)
    pieces  = []  # list of ('lit', text) | ('code', char)
    i = 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            end = section.find('"', i + 1)
            if end == -1:
                pieces.append(('lit', section[i+1:]))
                break
            pieces.append(('lit', section[i+1:end]))
            i = end + 1
        elif ch == '\\' and i + 1 < len(section):
            pieces.append(('lit', section[i+1]))
            i += 2
        elif ch == '_' and i + 1 < len(section):
            # `_X` reserves the width of X — render as a single space.
            pieces.append(('lit', ' '))
            i += 2
        elif ch == '*' and i + 1 < len(section):
            # `*X` repeats X to fill column width — we collapse to nothing.
            i += 2
        else:
            pieces.append(('code', ch))
            i += 1

    # A section qualifies as a "number format" only if it has at least one
    # real digit placeholder (`0`, `#`, or `?`). Otherwise it's pure literals
    # — e.g. Excel's `"-"` zero-section, where the hyphen is meant to render
    # as-is rather than be parsed as a sign.
    has_digit_placeholder = any(
        p[0] == 'code' and p[1] in '0#?' for p in pieces)
    if not has_digit_placeholder:
        return '', ''.join(p[1] for p in pieces), ''

    # Walk to find the digit-token span.
    is_digit_tok = lambda p: p[0] == 'code' and p[1] in '0#?.,%Ee+- '
    first = next((idx for idx, p in enumerate(pieces) if is_digit_tok(p)), None)
    last  = None
    for idx in range(len(pieces) - 1, -1, -1):
        if is_digit_tok(pieces[idx]):
            last = idx
            break
    if first is None or last is None:
        # No digit codes — treat the whole thing as a literal pass-through.
        return '', ''.join(p[1] for p in pieces), ''
    prefix  = ''.join(p[1] for p in pieces[:first])
    suffix  = ''.join(p[1] for p in pieces[last+1:])
    cleaned = ''.join(p[1] for p in pieces[first:last+1])
    return cleaned, prefix, suffix


def _apply_number_format(raw, num_fmt_id, fmt_code, is_date=False):
    """
    Reproduce what Excel would display for a numeric cell given its raw
    stored value and its number-format. Returns the original raw string on
    anything we can't confidently format.
    """
    if raw is None or raw == '':
        return ''
    if is_date:
        return _serial_to_date_str(raw)
    code = fmt_code or _BUILTIN_NUM_FMTS.get(num_fmt_id, '')
    if not code or code == 'General':
        return raw
    if code == '@':
        return raw
    try:
        v = float(raw)
    except (ValueError, TypeError):
        return raw

    section, v = _fmt_section(code, v)
    cleaned, prefix, suffix = _split_format_pieces(section)
    if not cleaned:
        # All-literal section (e.g. zero-section ""): keep prefix/suffix only.
        return (prefix + suffix).strip()

    # Percentage → multiply by 100, drop the % from the digit token but keep
    # it in the suffix so it appears in the output.
    percent_count = cleaned.count('%')
    if percent_count:
        v = v * (100 ** percent_count)
        cleaned = cleaned.replace('%', '')
        suffix = suffix + ('%' * percent_count)

    # Scientific → fall back to Python's `%.<n>E` using the digits before E.
    if 'E' in cleaned.upper():
        # Count the decimals in the mantissa portion (`0.00E+00` → 2).
        mantissa = cleaned.split('E')[0].split('e')[0]
        decimals = len(mantissa.split('.')[1]) if '.' in mantissa else 0
        try:
            return prefix + ('%.*E' % (decimals, v)) + suffix
        except Exception:
            return raw

    # Decide decimal count and thousands flag from the cleaned digit token.
    int_part, _, dec_part = cleaned.partition('.')
    decimals = len(re.sub(r'[^0#?]', '', dec_part))
    thousands = ',' in int_part

    sign = ''
    if v < 0 and ('(' not in section):
        # Default negative sign — when the section has explicit `(...)` the
        # parentheses already came through in prefix/suffix, so don't add `-`.
        sign = '-'
        v = abs(v)

    try:
        if thousands:
            formatted = '{0:,.{1}f}'.format(v, decimals)
        else:
            formatted = '{0:.{1}f}'.format(v, decimals)
    except Exception:
        return raw

    return prefix + sign + formatted + suffix


# Colour helpers
_INDEXED_COLORS = {
    0:(0,0,0),1:(255,255,255),2:(255,0,0),3:(0,255,0),4:(0,0,255),
    5:(255,255,0),6:(255,0,255),7:(0,255,255),8:(0,0,0),9:(255,255,255),
    10:(255,0,0),11:(0,255,0),12:(0,0,255),13:(255,255,0),14:(255,0,255),
    15:(0,255,255),16:(128,0,0),17:(0,128,0),18:(0,0,128),19:(128,128,0),
    20:(128,0,128),21:(0,128,128),22:(192,192,192),23:(128,128,128),
    24:(153,153,255),25:(153,51,102),26:(255,255,204),27:(204,255,255),
    28:(102,0,102),29:(255,128,128),30:(0,102,204),31:(204,204,255),
    32:(0,0,128),33:(255,0,255),34:(255,255,0),35:(0,255,255),
    36:(128,0,128),37:(128,0,0),38:(0,128,128),39:(0,0,255),
    40:(0,204,255),41:(204,255,255),42:(204,255,204),43:(255,255,153),
    44:(153,204,255),45:(255,153,204),46:(204,153,255),47:(255,204,153),
    48:(51,102,255),49:(51,204,204),50:(153,204,0),51:(255,204,0),
    52:(255,153,0),53:(255,102,0),54:(102,102,153),55:(150,150,150),
    56:(0,51,102),57:(51,153,102),58:(0,51,0),59:(51,51,0),
    60:(153,51,0),61:(153,51,102),62:(51,51,153),63:(51,51,51),
    64:(0,0,0),65:(255,255,255),
}


def _argb_to_rgb(argb):
    if not argb:
        return None
    argb = argb.strip().lstrip('#')
    if len(argb) == 8:
        try:
            return (int(argb[2:4], 16), int(argb[4:6], 16), int(argb[6:8], 16))
        except ValueError:
            return None
    elif len(argb) == 6:
        try:
            return (int(argb[0:2], 16), int(argb[2:4], 16), int(argb[4:6], 16))
        except ValueError:
            return None
    return None


def _is_white_or_none(rgb):
    return rgb is None or rgb == (255, 255, 255)


def _indexed_to_rgb(idx_str):
    try:
        idx = int(idx_str)
        # 64 = system foreground (automatic), 65 = system background (automatic)
        # These are NOT literal black/white — treat as no color (transparent)
        if idx in (64, 65):
            return None
        return _INDEXED_COLORS.get(idx)
    except (ValueError, TypeError):
        return None


def _read_theme_colors(zf, _names=None):
    _nl = _names if _names is not None else set(zf.namelist())
    theme_path = None
    for n in _nl:
        low = n.lower()
        if 'theme' in low and low.endswith('.xml') and 'theme1' in low:
            theme_path = n
            break
    if not theme_path:
        for n in _nl:
            if re.search(r'xl/theme/.*\.xml$', n, re.IGNORECASE):
                theme_path = n
                break
    if not theme_path:
        return []
    try:
        root = ET.fromstring(zf.read(theme_path))
    except Exception:
        return []
    colors = []
    for ns in ('http://schemas.openxmlformats.org/drawingml/2006/main',
               'http://schemas.openxmlformats.org/drawingml/2008/main'):
        t_ns = '{%s}' % ns
        clr_scheme = root.find('.//' + t_ns + 'clrScheme')
        if clr_scheme is None:
            continue
        # OpenXML theme color indexing:
        # 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4=accent1, etc.
        slots = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3',
                 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
        for slot in slots:
            el = clr_scheme.find(t_ns + slot)
            if el is None:
                colors.append(None)
                continue
            srgb = el.find(t_ns + 'srgbClr')
            if srgb is not None:
                colors.append(_argb_to_rgb(srgb.get('val', '')))
                continue
            sys_clr = el.find(t_ns + 'sysClr')
            if sys_clr is not None:
                last = sys_clr.get('lastClr', '')
                rgb = _argb_to_rgb(last) if last else None
                if rgb is None:
                    sval = sys_clr.get('val', '')
                    if sval in ('windowText', 'btnText'):
                        rgb = (0, 0, 0)
                    elif sval in ('window', 'btnFace'):
                        rgb = (255, 255, 255)
                colors.append(rgb)
                continue
            colors.append(None)
        if colors:
            break
    return colors


def _apply_tint(rgb, tint):
    r, g, b = rgb[0]/255.0, rgb[1]/255.0, rgb[2]/255.0
    if tint >= 0:
        r = r + (1-r)*tint; g = g + (1-g)*tint; b = b + (1-b)*tint
    else:
        r = r*(1+tint);     g = g*(1+tint);     b = b*(1+tint)
    return (int(round(r*255)), int(round(g*255)), int(round(b*255)))


def _resolve_color(el, theme_colors):
    if el is None:
        return None
    if el.get('auto', '0') == '1':
        return None
    rgb_attr = el.get('rgb', '')
    if rgb_attr:
        return _argb_to_rgb(rgb_attr)
    idx_attr = el.get('indexed', '')
    if idx_attr:
        return _indexed_to_rgb(idx_attr)
    theme_attr = el.get('theme', '')
    if theme_attr and theme_colors:
        try:
            tidx = int(theme_attr)
            if 0 <= tidx < len(theme_colors):
                base_rgb = theme_colors[tidx]
                if base_rgb is None:
                    return None
                tint_attr = el.get('tint', '')
                if tint_attr:
                    try:
                        base_rgb = _apply_tint(base_rgb, float(tint_attr))
                    except (ValueError, TypeError):
                        pass
                return base_rgb
        except (ValueError, TypeError):
            pass
    return None


def _read_styles(zf, theme_colors, _names=None):
    _nl = _names if _names is not None else set(zf.namelist())
    styles_path = None
    for n in _nl:
        if n.lower() == 'xl/styles.xml':
            styles_path = n
            break
    if styles_path is None:
        return [], 'Calibri', 11.0, {}, [], []

    root = ET.fromstring(zf.read(styles_path))

    num_fmt_map = {}
    nf_node = root.find(_tag('numFmts'))
    if nf_node is not None:
        for nf in nf_node.findall(_tag('numFmt')):
            try:
                nf_id   = int(nf.get('numFmtId', -1))
                nf_code = nf.get('formatCode', '')
                if nf_id >= 0:
                    num_fmt_map[nf_id] = nf_code
            except (ValueError, TypeError):
                pass

    # Grey pattern types: fgColor is the dot/pattern color (usually black),
    # bgColor is the actual cell background. Map directly to avoid black bleed-through.
    _GREY_PATTERN_RGB = {
        'darkGray':   (64,  64,  64),
        'mediumGray': (128, 128, 128),
        'lightGray':  (192, 192, 192),
        'gray0625':   (242, 242, 242),
        'gray125':    (217, 217, 217),
    }

    fills = []
    for fill in root.findall('.//' + _tag('fill')):
        pf = fill.find(_tag('patternFill'))
        if pf is not None:
            pat_type = pf.get('patternType', 'none')
            if pat_type == 'none':
                fills.append(None); continue
            if pat_type == 'solid':
                fg  = pf.find(_tag('fgColor'))
                rgb = _resolve_color(fg, theme_colors)
                # FIX: fgColor indexed=64 means "auto/system". For Excel grey fills,
                # the actual colour is often stored in bgColor instead.
                if _is_white_or_none(rgb):
                    bg     = pf.find(_tag('bgColor'))
                    bg_rgb = _resolve_color(bg, theme_colors) if bg is not None else None
                    if bg_rgb and not _is_white_or_none(bg_rgb):
                        rgb = bg_rgb
                fills.append(None if _is_white_or_none(rgb) else rgb); continue
            # Named grey patterns — use hardcoded RGB, ignore fgColor (dot color)
            if pat_type in _GREY_PATTERN_RGB:
                fills.append(_GREY_PATTERN_RGB[pat_type]); continue
            # Other non-solid patterns — prefer bgColor (cell bg) over fgColor (dots)
            bg  = pf.find(_tag('bgColor'))
            rgb = _resolve_color(bg, theme_colors)
            if not _is_white_or_none(rgb):
                fills.append(rgb); continue
            fg  = pf.find(_tag('fgColor'))
            rgb = _resolve_color(fg, theme_colors)
            fills.append(None if _is_white_or_none(rgb) else rgb); continue
        gf = fill.find(_tag('gradientFill'))
        if gf is not None:
            stops = gf.findall(_tag('stop'))
            if stops:
                rgb = _resolve_color(stops[0].find(_tag('color')), theme_colors)
                fills.append(None if _is_white_or_none(rgb) else rgb); continue
        fills.append(None)

    fonts = []
    for font in root.findall('.//' + _tag('font')):
        name_el   = font.find(_tag('name'))
        font_name = name_el.get('val', 'Calibri') if name_el is not None else 'Calibri'
        bold      = font.find(_tag('b'))  is not None
        italic    = font.find(_tag('i'))  is not None
        u_el      = font.find(_tag('u'))
        underline = (u_el is not None) and (u_el.get('val', 'single') != 'none')
        sz_el     = font.find(_tag('sz'))
        size      = float(sz_el.get('val', 11)) if sz_el is not None else None
        col_el    = font.find(_tag('color'))
        color     = _resolve_color(col_el, theme_colors)
        fonts.append({'name': font_name, 'bold': bold, 'italic': italic,
                      'underline': underline, 'size': size, 'color': color})

    default_font_name = fonts[0].get('name', 'Calibri') if fonts else 'Calibri'
    default_font_size = fonts[0].get('size', 11.0)      if fonts else 11.0
    if not default_font_name: default_font_name = 'Calibri'
    if not default_font_size: default_font_size = 11.0

    borders = []
    for border in root.findall('.//' + _tag('border')):
        def _bs(side):
            el = border.find(_tag(side))
            if el is None: return None
            s  = el.get('style', '')
            # Excel writes <top style="none"/> when a border is explicitly
            # removed.  Treat that the same as a missing border so the mapping
            # logic doesn't fall back to 'Thin Lines' for cells the user
            # specifically wanted borderless.
            if not s or s == 'none':
                return None
            return s
        borders.append({'top':    _bs('top'),    'bottom': _bs('bottom'),
                        'left':   _bs('left'),   'right':  _bs('right')})

    styles          = []
    style_is_date   = []
    style_fmt_codes = []   # parallel to styles; format code per cellXf
    cell_xfs = root.find(_tag('cellXfs'))
    if cell_xfs is None:
        return [], default_font_name, default_font_size, num_fmt_map, [], []

    for xf in cell_xfs.findall(_tag('xf')):
        fill_id    = int(xf.get('fillId',   0))
        font_id    = int(xf.get('fontId',   0))
        border_id  = int(xf.get('borderId', 0))
        num_fmt_id = int(xf.get('numFmtId', 0))

        fill_rgb = fills[fill_id]     if fill_id   < len(fills)   else None
        font     = fonts[font_id]     if font_id   < len(fonts)   else {}
        bord     = borders[border_id] if border_id < len(borders) else {}

        h_align   = None
        v_align   = None
        wrap_text = False
        align_el  = xf.find(_tag('alignment'))
        if align_el is not None:
            h_align   = align_el.get('horizontal') or None
            v_align   = align_el.get('vertical')   or None
            wrap_text = align_el.get('wrapText', '0') in ('1', 'true', 'True')
            try:
                text_rotation = int(align_el.get('textRotation', 0) or 0)
            except (ValueError, TypeError):
                text_rotation = 0
        else:
            text_rotation = 0

        _raw_fc     = font.get('color', None)
        _font_color = (None if (fill_rgb is None and _raw_fc == (255, 255, 255))
                       else _raw_fc)

        styles.append({
            'fill_rgb':       fill_rgb,
            'font_name':      font.get('name',      None),
            'font_bold':      font.get('bold',       False),
            'font_italic':    font.get('italic',     False),
            'font_underline': font.get('underline',  False),
            'font_size':      font.get('size',       None),
            'font_color':     _font_color,
            'border_top':     bord.get('top',        None),
            'border_bottom':  bord.get('bottom',     None),
            'border_left':    bord.get('left',       None),
            'border_right':   bord.get('right',      None),
            'h_align':        h_align,
            'v_align':        v_align,
            'wrap_text':      wrap_text,
            'text_rotation':  text_rotation,
        })
        custom_code = num_fmt_map.get(num_fmt_id, '')
        style_is_date.append(_is_date_format(num_fmt_id, custom_code))
        # Effective format string for this cell style: prefer the workbook's
        # custom code; otherwise fall back to the OOXML built-in for this id.
        style_fmt_codes.append(custom_code or _BUILTIN_NUM_FMTS.get(num_fmt_id, ''))

    return (styles, default_font_name, default_font_size, num_fmt_map,
            style_is_date, style_fmt_codes)


_EMPTY_STYLE = {
    'fill_rgb': None, 'font_name': None, 'font_bold': False,
    'font_italic': False, 'font_underline': False, 'font_size': None,
    'font_color': None, 'border_top': None, 'border_bottom': None,
    'border_left': None, 'border_right': None,
    'h_align': None, 'v_align': None, 'wrap_text': False,
    'text_rotation': 0,
}


def _read_shared_strings(zf, _names=None):
    _nl = _names if _names is not None else set(zf.namelist())
    target = None
    for n in _nl:
        if n.lower() == 'xl/sharedstrings.xml':
            target = n
            break
    if target is None:
        return []
    root = ET.fromstring(zf.read(target))
    strings = []
    for si in root.findall('.//' + _tag('si')):
        parts = []
        for t in si.iter(_tag('t')):
            text = t.text
            if text is not None:
                parts.append(text)
        strings.append(''.join(parts))
    return strings


def _read_workbook(zf, _names=None):
    _nl = _names if _names is not None else set(zf.namelist())
    wb_path = None
    for n in _nl:
        if n.lower().endswith('workbook.xml'):
            wb_path = n
            break
    if wb_path is None:
        return [], {}, {}

    root = ET.fromstring(zf.read(wb_path))
    sheets = []
    for sh in root.iter(_tag('sheet')):
        sheets.append({
            'name':     sh.get('name', 'Sheet'),
            'sheet_id': sh.get('sheetId', ''),
            'rel_id':   sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', ''),
        })

    print_areas = {}
    for dn in root.iter(_tag('definedName')):
        if dn.get('name', '') == '_xlnm.Print_Area':
            local_id = int(dn.get('localSheetId', 0))
            if local_id < len(sheets):
                print_areas[sheets[local_id]['rel_id']] = dn.text or None

    rels = {}
    wb_dir    = wb_path.rsplit('/', 1)[0] if '/' in wb_path else 'xl'
    rels_path = wb_dir + '/_rels/' + wb_path.split('/')[-1] + '.rels'
    if rels_path in _nl:
        rels_root = ET.fromstring(zf.read(rels_path))
        NS_REL  = 'http://schemas.openxmlformats.org/package/2006/relationships'
        WS_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
        for rel in rels_root.findall('{%s}Relationship' % NS_REL):
            if rel.get('Type', '') == WS_TYPE:
                tgt = rel.get('Target', '')
                if not tgt.startswith('xl/'):
                    tgt = wb_dir + '/' + tgt
                rels[rel.get('Id', '')] = tgt

    return sheets, print_areas, rels


def _get_cell_value(cell_elem, shared_strings, is_date=False,
                    fmt_code='', num_fmt_id=0):
    """
    Return the display string for a cell, applying its number format when
    available. `fmt_code` / `num_fmt_id` are looked up from the cell's
    style (cellXfs[s]) at the call site; when both are empty the raw value
    is returned unchanged (back-compat for any caller that doesn't pass
    them).
    """
    t = cell_elem.get('t', 'n')
    if t == 's':
        v = cell_elem.find(_tag('v'))
        if v is not None and v.text is not None:
            try:
                idx = int(v.text)
                if 0 <= idx < len(shared_strings):
                    return shared_strings[idx]
            except (ValueError, TypeError):
                pass
        return ''
    elif t == 'inlineStr':
        is_node = cell_elem.find(_tag('is'))
        if is_node is not None:
            parts = []
            for t_node in is_node.iter(_tag('t')):
                text = t_node.text
                if text is not None:
                    parts.append(text)
            return ''.join(parts)
        return ''
    elif t in ('str', 'e'):
        v = cell_elem.find(_tag('v'))
        return v.text if (v is not None and v.text) else ''
    elif t == 'b':
        v = cell_elem.find(_tag('v'))
        if v is not None and v.text is not None:
            return 'TRUE' if v.text.strip() == '1' else 'FALSE'
        return ''
    else:
        v = cell_elem.find(_tag('v'))
        if v is not None and v.text is not None:
            raw = v.text.strip()
            if is_date:
                return _serial_to_date_str(raw)
            # Apply the cell's number format when one is known. Dates have
            # already returned above; this covers `0.000`, `#,##0.00`, `0%`,
            # currency, scientific, etc. Returns the raw value if the format
            # is "General" / empty or if anything goes sideways.
            if fmt_code or num_fmt_id:
                return _apply_number_format(raw, num_fmt_id, fmt_code)
            try:
                f = float(raw)
                return str(int(f)) if f == int(f) else raw
            except (ValueError, TypeError):
                return raw
        return ''


def list_sheets(filepath):
    # Validate file exists first
    if not filepath or not os.path.exists(filepath):
        raise IOError('File not found or path is incomplete: {0}'.format(filepath))
    
    # Auto-convert any Excel format to .xlsx
    try:
        filepath = _ensure_xlsx_format(filepath)
    except Exception as e:
        raise IOError('Failed to normalize Excel format: {0}'.format(str(e)))
    
    # Verify it's now a valid xlsx/xlsm file (zipfile)
    if not zipfile.is_zipfile(filepath):
        raise IOError('Not a valid Excel file (.xlsx/.xlsm): {0}'.format(filepath))
    try:
        zf = zipfile.ZipFile(filepath, 'r')
    except (IOError, OSError) as e:
        raise IOError('Cannot open file: {0}'.format(str(e)))
    try:
        sheets, print_areas, _ = _read_workbook(zf)
        results = []
        for sh in sheets:
            pa = print_areas.get(sh['rel_id'])
            # pa might be like "'Sheet Name'!$A$1:$G$25" or just "$A$1:$G$25"
            pa_clean = None
            if pa:
                pa_clean = pa.split('!')[-1].replace('$', '')
            
            results.append({
                'name': sh['name'],
                'has_print_area': pa is not None,
                'print_area': pa_clean
            })
        return results
    finally:
        zf.close()


# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 1.5: IMAGE EXTRACTOR  (OpenXML image parsing)
# ══════════════════════════════════════════════════════════════════════════════

# Namespaces for drawing/relationship parsing
_NS_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
_NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
_NS_OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'



def _extract_images_from_excel(filepath, sheet_name, temp_dir):
    """Extract embedded images from an Excel .xlsx sheet.

    V23:
      - Supports xdr:twoCellAnchor (existing)
      - Adds xdr:oneCellAnchor
      - Adds xdr:absoluteAnchor (no cell ref; mapped as fallback)

    Returns list of dict:
        image_path, image_format,
        from_row/from_col/to_row/to_col (1-based),
        image_id, anchor_type
    """
    images = []

    # Validate file exists
    if not filepath or not os.path.exists(filepath):
        return images
    
    # Auto-convert any Excel format to .xlsx
    try:
        filepath = _ensure_xlsx_format(filepath)
    except Exception:
        return images

    if not zipfile.is_zipfile(filepath):
        return images

    try:
        zf = zipfile.ZipFile(filepath, 'r')
    except (IOError, OSError):
        return images

    def _resolve_target(base_dir, tgt):
        if not tgt:
            return None
        if tgt.startswith('/'):
            return tgt[1:]
        if tgt.startswith('xl/'):
            return tgt
        if tgt.startswith('../'):
            parts = [p for p in tgt.split('/') if p != '..']
            return 'xl/' + '/'.join(parts)
        return base_dir + '/' + tgt

    def _read_marker(marker):
        row = marker.find('{%s}row' % _NS_DRAWING)
        col = marker.find('{%s}col' % _NS_DRAWING)
        row_off = marker.find('{%s}rowOff' % _NS_DRAWING)
        col_off = marker.find('{%s}colOff' % _NS_DRAWING)
        return {
            'row': int(row.text) if row is not None and row.text else 0,
            'col': int(col.text) if col is not None and col.text else 0,
            'row_off': int(row_off.text) if row_off is not None and row_off.text else 0,
            'col_off': int(col_off.text) if col_off is not None and col_off.text else 0,
        }

    def _extract_pic(anchor, image_id_to_path, anchor_type, from_coords, to_coords, names_set):
        pic = anchor.find('{%s}pic' % _NS_DRAWING)
        if pic is None:
            return
        blip = pic.find('.//{%s}blip' % 'http://schemas.openxmlformats.org/drawingml/2006/main')
        if blip is None:
            return
        embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if not embed:
            return

        img_path = image_id_to_path.get(embed)
        if not img_path:
            return
        if img_path not in names_set:
            return

        img_data = zf.read(img_path)
        img_format = img_path.split('.')[-1].lower() if '.' in img_path else 'png'
        if img_format not in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff', 'webp'):
            img_format = 'png'

        temp_path = os.path.join(temp_dir, 'excel_image_{0}.{1}'.format(len(images), img_format))
        with open(temp_path, 'wb') as f:
            f.write(img_data)

        images.append({
            'image_path': temp_path,
            'image_format': img_format,
            'from_row': from_coords.get('row', 0) + 1,
            'from_col': from_coords.get('col', 0) + 1,
            'to_row': to_coords.get('row', 0) + 1,
            'to_col': to_coords.get('col', 0) + 1,
            'image_id': embed,
            'anchor_type': anchor_type,
        })

    try:
        names_set = set(zf.namelist())
        sheets, print_areas, wb_rels = _read_workbook(zf, names_set)

        target_sheet = None
        for sh in sheets:
            if sh['name'] == sheet_name:
                target_sheet = sh
                break
        if target_sheet is None:
            return images

        rel_id = target_sheet['rel_id']
        sheet_path = wb_rels.get(rel_id)
        if not sheet_path:
            return images

        sheet_dir = sheet_path.rsplit('/', 1)[0] if '/' in sheet_path else 'xl'
        sheet_rels_path = sheet_dir + '/_rels/' + sheet_path.split('/')[-1] + '.rels'

        drawing_target = None
        if sheet_rels_path in names_set:
            try:
                rels_root = ET.fromstring(zf.read(sheet_rels_path))
                for rel in rels_root.findall('{%s}Relationship' % _NS_REL):
                    rel_type = rel.get('Type', '')
                    if 'drawing' in (rel_type or '').lower():
                        drawing_target = rel.get('Target', '')
                        break
            except Exception:
                drawing_target = None

        if not drawing_target:
            return images

        drawing_path = _resolve_target(sheet_dir, drawing_target)
        if not drawing_path or drawing_path not in names_set:
            return images

        try:
            drawing_root = ET.fromstring(zf.read(drawing_path))
        except Exception:
            return images

        drawing_dir = drawing_path.rsplit('/', 1)[0] if '/' in drawing_path else 'xl/drawings'
        drawing_rels_path = drawing_dir + '/_rels/' + drawing_path.split('/')[-1] + '.rels'

        image_id_to_path = {}
        if drawing_rels_path in names_set:
            try:
                rels_root = ET.fromstring(zf.read(drawing_rels_path))
                for rel in rels_root.findall('{%s}Relationship' % _NS_REL):
                    rel_type = rel.get('Type', '')
                    if 'image' in (rel_type or '').lower():
                        img_id = rel.get('Id', '')
                        img_target = rel.get('Target', '')
                        img_path = _resolve_target(drawing_dir, img_target)
                        if img_id and img_path:
                            image_id_to_path[img_id] = img_path
            except Exception:
                pass

        for anchor in drawing_root.findall('.//{%s}twoCellAnchor' % _NS_DRAWING):
            try:
                fm = anchor.find('{%s}from' % _NS_DRAWING)
                to = anchor.find('{%s}to' % _NS_DRAWING)
                if fm is None or to is None:
                    continue
                _extract_pic(anchor, image_id_to_path, 'twoCellAnchor', _read_marker(fm), _read_marker(to), names_set)
            except Exception:
                continue

        for anchor in drawing_root.findall('.//{%s}oneCellAnchor' % _NS_DRAWING):
            try:
                fm = anchor.find('{%s}from' % _NS_DRAWING)
                if fm is None:
                    continue
                c = _read_marker(fm)
                _extract_pic(anchor, image_id_to_path, 'oneCellAnchor', c, c, names_set)
            except Exception:
                continue

        for anchor in drawing_root.findall('.//{%s}absoluteAnchor' % _NS_DRAWING):
            try:
                c = {'row': 0, 'col': 0, 'row_off': 0, 'col_off': 0}
                _extract_pic(anchor, image_id_to_path, 'absoluteAnchor', c, c, names_set)
            except Exception:
                continue

    finally:
        try:
            zf.close()
        except Exception:
            pass

    return images


def _map_images_to_cells(images, print_area_bounds):
    """Map extracted images to grid cells based on print area.

    V23: absoluteAnchor has no cell reference; map it to the print-area top-left (0,0)
    as best-effort fallback.
    """
    if not images or not print_area_bounds:
        return {}

    min_col, min_row, max_col, max_row = print_area_bounds
    cell_images = {}

    for img in images:
        if img.get('anchor_type') == 'absoluteAnchor':
            if img.get('image_path'):
                cell_images[(0, 0)] = img.get('image_path')
            continue

        from_row = img.get('from_row', 0)
        from_col = img.get('from_col', 0)

        if from_row < min_row or from_row > max_row:
            continue
        if from_col < min_col or from_col > max_col:
            continue

        rel_row = from_row - min_row
        rel_col = from_col - min_col
        cell_images[(rel_row, rel_col)] = img.get('image_path')

    return cell_images

def parse_excel(filepath, sheet_name=None, max_rows=DEFAULT_MAX_ROWS,
                strict_print_area=False, named_range=None, manual_range=None):
    """
    Parse an .xlsx/.xlsm worksheet into a tabular `parsed_data` payload.

    Range precedence (first hit wins):
      1. `named_range`   — workbook defined name; resolved sheet wins.
      2. `manual_range`  — A1 reference like "A1:G25", applied to sheet_name.
      3. `_xlnm.Print_Area` for the sheet.
      4. Worksheet dimension / used range.

    `named_range` resolution failures fall back to (3) with a warning.
    """
    warnings = []

    # Validate file exists first
    if not filepath or not os.path.exists(filepath):
        raise IOError('File not found or path is incomplete: {0}'.format(filepath))
    
    # Auto-convert any Excel format to .xlsx
    try:
        filepath = _ensure_xlsx_format(filepath)
    except Exception as e:
        raise IOError('Failed to normalize Excel format: {0}'.format(str(e)))

    if not zipfile.is_zipfile(filepath):
        raise IOError('Not a valid .xlsx file: {0}'.format(filepath))
    try:
        zf = zipfile.ZipFile(filepath, 'r')
    except (IOError, OSError) as e:
        raise IOError('Cannot open file: {0}'.format(str(e)))

    try:
        _names = set(zf.namelist())
        shared_strings = _read_shared_strings(zf, _names)
        theme_colors   = _read_theme_colors(zf, _names)
        (style_table, default_font_name, default_font_size, num_fmt_map,
         style_is_date, style_fmt_codes) = _read_styles(zf, theme_colors, _names)
        sheets, print_areas, rels = _read_workbook(zf, _names)

        if not sheets:
            raise ValueError('No sheets found in workbook.')

        # Named-Range mode: resolve the workbook's defined name FIRST. The
        # name's `ref_sheet` becomes the target so a workbook-scoped name
        # pointing off the user's selected sheet still resolves correctly.
        # If resolution fails (name was renamed/deleted since the record
        # was saved, file replaced, etc.) we fall back to print-area / used
        # range instead of crashing the import — surface the issue via the
        # parsed result's `warnings` list.
        nm_info = None
        nm_bounds_str = None
        if named_range:
            try:
                nm_info = resolve_defined_name(filepath, named_range, sheet_name)
            except Exception as _nm_ex:
                nm_info = None
                warnings.append(
                    "Could not look up defined name '{0}' ({1}). "
                    "Falling back to print area / used range.".format(named_range, _nm_ex))
            if nm_info is None:
                warnings.append(
                    "Defined name '{0}' not found in workbook. "
                    "Falling back to print area / used range.".format(named_range))
            else:
                sheet_name    = nm_info.get('ref_sheet') or sheet_name
                nm_bounds_str = "{0}!{1}".format(
                    nm_info.get('ref_sheet') or '', nm_info.get('range_address') or '')

        target_sheet = None
        if sheet_name:
            for sh in sheets:
                if sh['name'] == sheet_name:
                    target_sheet = sh
                    break
            if target_sheet is None:
                raise ValueError(
                    "Sheet '{0}' not found. Available: {1}".format(
                        sheet_name, ', '.join(s['name'] for s in sheets)))
        else:
            for sh in sheets:
                if sh['rel_id'] in print_areas:
                    target_sheet = sh
                    break
            if target_sheet is None:
                target_sheet = sheets[0]

        rel_id     = target_sheet['rel_id']
        sheet_path = rels.get(rel_id)
        if not sheet_path or sheet_path not in _names:
            raise ValueError(
                "Cannot locate worksheet file for sheet '{0}'.".format(
                    target_sheet['name']))

        sheet_root = ET.fromstring(zf.read(sheet_path))

        print_area_str = print_areas.get(rel_id)
        # Named-Range mode wins over Print_Area: when the caller supplied a
        # defined name, its resolved range is the parse bounds (and `strict`
        # is irrelevant because we already have a range).
        if nm_bounds_str:
            print_area_str = nm_bounds_str
        # Manual-Range mode is second priority — a user-typed A1 ref applied
        # to the selected sheet. Same effect as defining an ad-hoc name.
        elif manual_range:
            print_area_str = "{0}!{1}".format(target_sheet['name'], manual_range.strip())
        # V23: optional strict Print Area requirement
        elif strict_print_area and not print_area_str:
            raise ValueError(
                'PRINT AREA NOT DEFINED\n\n'
                'Sheet "{0}" has no Print Area set.\n\n'
                'To fix: In Excel → Page Layout → Print Area → Set Print Area, then re-run.'
                .format(target_sheet['name']))

        bounds = _parse_range(print_area_str) if print_area_str else None

        if bounds is None:
            dim = sheet_root.find(_tag('dimension'))
            if dim is not None:
                bounds = _parse_range(dim.get('ref', ''))
            if bounds is None:
                raise ValueError(
                    "No print area or dimension found on sheet '{0}'. "
                    "Please set a Print Area in Excel.".format(target_sheet['name']))
            warnings.append(
                "No print area on sheet '{0}'. Using worksheet used range.".format(
                    target_sheet['name']))

        min_col, min_row, max_col, max_row = bounds
        num_rows = max_row - min_row + 1
        num_cols = max_col - min_col + 1

        if max_rows and num_rows > max_rows:
            raise ValueError(
                "Sheet '{0}' print area has {1} rows which exceeds the "
                "configured limit of {2}. "
                "Either reduce the print area or increase max_rows.".format(
                    target_sheet['name'], num_rows, max_rows))

        fmt_node = sheet_root.find(_tag('sheetFormatPr'))
        default_col_width_chars = 8.43
        default_row_height_pt   = 15.0
        if fmt_node is not None:
            try:
                default_col_width_chars = float(fmt_node.get('defaultColWidth', 8.43))
            except (ValueError, TypeError):
                pass
            try:
                default_row_height_pt = float(fmt_node.get('defaultRowHeight', 15.0))
            except (ValueError, TypeError):
                pass
        if default_row_height_pt <= 0:
            default_row_height_pt = 15.0

        col_width_map = {}
        cols_node = sheet_root.find(_tag('cols'))
        if cols_node is not None:
            for col_def in cols_node.findall(_tag('col')):
                try:
                    c_min  = int(col_def.get('min', 1))
                    c_max  = int(col_def.get('max', 1))
                    w      = float(col_def.get('width', default_col_width_chars))
                    custom = col_def.get('customWidth', '0')
                    if custom in ('1', 'true', 'True') or w != default_col_width_chars:
                        for ci in range(c_min, c_max + 1):
                            col_width_map[ci] = w
                except (ValueError, TypeError):
                    pass

        col_widths = [
            col_width_map.get(c, default_col_width_chars)
            for c in range(min_col, max_col + 1)
        ]

        row_height_map = {}
        sd_for_heights = sheet_root.find(_tag('sheetData'))
        if sd_for_heights is not None:
            for row_elem in sd_for_heights.findall(_tag('row')):
                try:
                    r_num = int(row_elem.get('r', 0))
                except (ValueError, TypeError):
                    continue
                ht = row_elem.get('ht')
                if ht is not None:
                    try:
                        h = float(ht)
                        if h > 0:
                            row_height_map[r_num] = h
                    except (ValueError, TypeError):
                        pass

        row_heights = [
            row_height_map.get(abs_row, default_row_height_pt)
            for abs_row in range(min_row, max_row + 1)
        ]

        merge_map = {}
        mc_node = sheet_root.find(_tag('mergeCells'))
        if mc_node is not None:
            for mc in mc_node.findall(_tag('mergeCell')):
                b = _parse_range(mc.get('ref', ''))
                if b is None:
                    warnings.append('Could not parse merge: {0}'.format(mc.get('ref')))
                    continue
                mc_min_col, mc_min_row, mc_max_col, mc_max_row = b
                r_span = mc_max_row - mc_min_row + 1
                c_span = mc_max_col - mc_min_col + 1
                for r in range(mc_min_row, mc_max_row + 1):
                    for c in range(mc_min_col, mc_max_col + 1):
                        if r == mc_min_row and c == mc_min_col:
                            merge_map[(r, c)] = {'type': 'master', 'span': (r_span, c_span)}
                        else:
                            merge_map[(r, c)] = {'type': 'spanned', 'span': None}

        display_value_map = {}
        try:
            display_value_map = _read_excel_display_values(
                filepath, target_sheet.get('name'), bounds)
        except Exception as ex:
            warnings.append('Could not read exact display text from Excel; using raw workbook values. ({})'.format(ex))

        cell_map = {}
        sheet_data = sheet_root.find(_tag('sheetData'))
        if sheet_data is not None:
            for row_elem in sheet_data.findall(_tag('row')):
                for c_elem in row_elem.findall(_tag('c')):
                    abs_r, abs_c = _cell_ref_to_rowcol(c_elem.get('r', ''))
                    if abs_r is not None:
                        s_idx   = int(c_elem.get('s', 0))
                        is_date = (style_is_date[s_idx]
                                   if s_idx < len(style_is_date) else False)
                        fmt_code = (style_fmt_codes[s_idx]
                                    if s_idx < len(style_fmt_codes) else '')
                        raw_value = _get_cell_value(
                            c_elem, shared_strings, is_date,
                            fmt_code=fmt_code)
                        display_value = display_value_map.get((abs_r, abs_c))
                        if display_value is None or display_value == '':
                            display_value = raw_value
                        cell_map[(abs_r, abs_c)] = {
                            'value':     display_value,
                            'style_idx': s_idx,
                        }

        cells = []
        for abs_row in range(min_row, max_row + 1):
            for abs_col in range(min_col, max_col + 1):
                rel_row = abs_row - min_row
                rel_col = abs_col - min_col
                key     = (abs_row, abs_col)
                mi      = merge_map.get(key)

                if mi is None:
                    is_master  = True
                    is_spanned = False
                    merge_span = None
                elif mi['type'] == 'master':
                    is_master  = True
                    is_spanned = False
                    merge_span = mi['span']
                else:
                    is_master  = False
                    is_spanned = True
                    merge_span = None

                raw   = cell_map.get(key, {})
                value = raw.get('value', '') if is_master else ''
                s_idx = raw.get('style_idx', 0)
                style = style_table[s_idx] if s_idx < len(style_table) else _EMPTY_STYLE

                cell_dict = {
                    'row': rel_row, 'col': rel_col, 'value': value,
                    'is_master': is_master, 'is_spanned': is_spanned,
                    'merge_span': merge_span,
                }
                cell_dict.update(style)
                cells.append(cell_dict)

    finally:
        zf.close()

    return {
        'rows':              num_rows,
        'cols':              num_cols,
        'col_widths':        col_widths,
        'row_heights':       row_heights,
        'default_font_name': default_font_name,
        'default_font_size': default_font_size,
        'cells':             cells,
        'warnings':          warnings,
        'bounds':            bounds,
        'print_area_str':    print_area_str,
    }


# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 2: GRID MODEL  (unit conversion + CellData normalisation)
# ══════════════════════════════════════════════════════════════════════════════

def _mm_to_feet(mm):
    """Millimeters to Revit internal feet."""
    return mm / 304.8


def _pt_to_feet_exact(size_pt):
    """Font size points -> Revit internal feet.
    1 ft = 12 in * 72 pt/in = 864 pt.
    Note: Revit's SCHEDULE_TEXT_SIZE and TableCellStyle.TextSize interpret
    the raw point value directly, so we pass size_pt unchanged.
    """
    return size_pt / 1


def _compute_schedule_baseline_text_size_ft(grid_model, options=None):
    """Return the baseline text size in feet for SCHEDULE_TEXT_SIZE.
    
    If 'base_text_size_mm' is provided in options, it overrides the scale-based 
    logic to ensure a consistent physical size.
    """
    options = options or {}
    base_mm = options.get("base_text_size_mm")
    
    if base_mm and float(base_mm) > 0:
        return _mm_to_feet(float(base_mm))
        
    try:
        scale = options.get("text_scale", 1.0)
        return _pt_to_feet_exact(grid_model.default_font_size_pt * scale)
    except Exception:
        return _pt_to_feet_exact(11.0)


_MDW_TABLE = {
    'calibri':7,'arial':7,'arial narrow':6,'tahoma':6,'times new roman':8,
    'segoe ui':6,'verdana':8,'courier new':7,'trebuchet ms':7,'georgia':8,
    'comic sans ms':8,'impact':8,'lucida console':7,'palatino linotype':8,
    'book antiqua':8,'garamond':7,'century gothic':7,'franklin gothic':7,
    'helvetica':7,'myriad pro':7,
}
_MDW_DEFAULT = 7

# Tunable scale factors — exposed so the options dialog can override them
COL_WIDTH_SCALE  = 1.0
ROW_HEIGHT_SCALE = 1.0
TEXT_SIZE_SCALE  = 1.0
ENABLE_DYNAMIC_MERGE = True
SHOW_EMPTY_GRIDLINES = False

COL_WIDTH_MIN_FT  = 0.04
COL_WIDTH_MAX_FT  = 5.0
ROW_HEIGHT_MIN_FT = 0.015
ROW_HEIGHT_MAX_FT = 1.0


def _get_mdw(font_name):
    if not font_name:
        return _MDW_DEFAULT
    return _MDW_TABLE.get(font_name.lower().strip(), _MDW_DEFAULT)


def _chars_to_feet(width_chars, mdw_px=_MDW_DEFAULT, scale=1.0):
    inches = (width_chars * mdw_px) / 96.0
    return (inches / 12.0) * scale


def _pt_to_feet(height_pt, scale=1.0):
    """Excel row height points -> Revit internal feet."""
    return (height_pt / (72.0 * 12.0)) * scale





def _find_best_text_note_type(doc, font_name, size_ft):
    """Pick a TextNoteType close to desired font + size.

    Used to drive schedule-level "Text Appearance" / "Text Type" style when
    the sheet renderer ignores per-cell overrides.
    """
    if doc is None or (not _HAS_TNT):
        return None

    desired_font = (font_name or '').strip()
    desired_size = float(size_ft) if size_ft else None

    best_id = None
    best_score = None

    try:
        tnts = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
    except Exception:
        tnts = []

    for tnt in tnts or []:
        try:
            p_font = tnt.get_Parameter(BuiltInParameter.TEXT_FONT)
            p_size = tnt.get_Parameter(BuiltInParameter.TEXT_SIZE)
            cur_font = p_font.AsString() if p_font else None
            cur_size = p_size.AsDouble() if p_size else None

            if not cur_size or cur_size <= 0:
                continue

            # Score: size delta (primary) + font mismatch penalty
            if desired_size:
                score = abs(cur_size - desired_size)
            else:
                score = 0.0

            if desired_font and cur_font and cur_font.strip().lower() != desired_font.lower():
                score += 1.0  # big penalty in feet-units context; just needs to dominate

            if best_score is None or score < best_score:
                best_score = score
                best_id = tnt.Id
        except Exception:
            pass

    return best_id


def _set_schedule_text_appearance(doc, schedule, text_type_id):
    """Attempt to set schedule 'Text Appearance' / 'Text Type' element id."""
    if schedule is None or text_type_id is None:
        return

    # Known BIP seen in the BuiltInParameter list (Revit 2026):
    # VIEW_GRAPH_SCHED_TEXT_APPEARANCE  -> "Text Appearance"
    for bip_name in ('VIEW_GRAPH_SCHED_TEXT_APPEARANCE',):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = schedule.get_Parameter(bip)
            if p is not None and (not p.IsReadOnly) and p.StorageType.ToString() == 'ElementId':
                p.Set(text_type_id)
        except Exception:
            pass

    # Locale-safe fallback: scan parameters by name
    try:
        for p in schedule.Parameters:
            try:
                if p is None or p.IsReadOnly:
                    continue
                if p.StorageType.ToString() != 'ElementId':
                    continue
                nm = (p.Definition.Name or '').lower()
                if any(kw in nm for kw in ('text appearance', 'text type', 'textstyle', 'text style')):
                    p.Set(text_type_id)
            except Exception:
                pass
    except Exception:
        pass


# Font fallback mapping for Excel to Revit font compatibility
_FONT_FALLBACK_MAP = {
    'calibri': 'Calibri',
    'arial': 'Arial',
    'arial narrow': 'Arial Narrow',
    'times new roman': 'Times New Roman',
    'tahoma': 'Tahoma',
    'segoe ui': 'Segoe UI',
    'verdana': 'Verdana',
    'courier new': 'Courier New',
    'trebuchet ms': 'Trebuchet MS',
    'georgia': 'Georgia',
    'comic sans ms': 'Comic Sans MS',
    'impact': 'Impact',
    'lucida console': 'Lucida Console',
    'palatino linotype': 'Palatino Linotype',
    'book antiqua': 'Book Antiqua',
    'garamond': 'Garamond',
    'century gothic': 'Century Gothic',
    'franklin gothic': 'Franklin Gothic',
    'helvetica': 'Helvetica',
    'myriad pro': 'Myriad Pro',
    'wingdings': 'Wingdings',
    'webdings': 'Webdings',
    'symbol': 'Symbol',
}


def _resolve_font_name(excel_font_name):
    """
    Map Excel font names to Revit-available fonts with fallback.
    Returns normalized font name or original if no mapping found.
    """
    if not excel_font_name:
        return 'Calibri'
    
    font_lower = excel_font_name.lower().strip()
    
    # Direct lookup
    if font_lower in _FONT_FALLBACK_MAP:
        return _FONT_FALLBACK_MAP[font_lower]
    
    # Try partial match
    for key, value in _FONT_FALLBACK_MAP.items():
        if key in font_lower or font_lower in key:
            return value
    
    # Return original (Revit will fallback if not available)
    return excel_font_name


class CellData(object):
    __slots__ = (
        'row', 'col', 'value',
        'is_master', 'is_spanned', 'merge_span',
        'fill_rgb', 'font_name',
        'font_bold', 'font_italic', 'font_underline',
        'font_size', 'font_color',
        'border_top', 'border_bottom', 'border_left', 'border_right',
        'h_align', 'v_align', 'wrap_text', 'text_rotation',
        'image_path', 'is_image',
    )

    def __init__(self, row, col, value='',
                 is_master=True, is_spanned=False, merge_span=None,
                 fill_rgb=None, font_name=None,
                 font_bold=False, font_italic=False, font_underline=False,
                 font_size=None, font_color=None,
                 border_top=None, border_bottom=None,
                 border_left=None, border_right=None,
                 h_align=None, v_align=None, wrap_text=False,
                 text_rotation=0,
                 image_path=None, is_image=False):
        self.row=row; self.col=col; self.value=value
        self.is_master=is_master; self.is_spanned=is_spanned; self.merge_span=merge_span
        self.fill_rgb=fill_rgb; self.font_name=font_name
        self.font_bold=font_bold; self.font_italic=font_italic
        self.font_underline=font_underline; self.font_size=font_size
        self.font_color=font_color
        self.border_top=border_top; self.border_bottom=border_bottom
        self.border_left=border_left; self.border_right=border_right
        self.h_align=h_align; self.v_align=v_align; self.wrap_text=wrap_text
        self.text_rotation=text_rotation
        self.image_path=image_path; self.is_image=is_image

    def as_dict(self):
        return {s: getattr(self, s) for s in self.__slots__}


class GridModel(object):
    def __init__(self, rows, cols, cells, col_widths_ft, row_heights_ft,
                 default_font_name, default_font_size_pt, warnings):
        self.rows=rows; self.cols=cols; self.cells=cells
        self.col_widths_ft=col_widths_ft; self.row_heights_ft=row_heights_ft
        self.default_font_name=default_font_name
        self.default_font_size_pt=default_font_size_pt
        self.warnings=warnings
        self._grid = {(c.row, c.col): c for c in cells}

    def get(self, row, col):
        return self._grid.get((row, col))

    def master_cells(self):
        return [c for c in self.cells if c.is_master]

    def merged_masters(self):
        return [c for c in self.cells
                if c.is_master and c.merge_span is not None
                and (c.merge_span[0] > 1 or c.merge_span[1] > 1)]


def _validate_merges(cells, warnings):
    covered = {}
    for cell in cells:
        if not cell.is_master or cell.merge_span is None:
            continue
        r_span, c_span = cell.merge_span
        if r_span == 1 and c_span == 1:
            continue
        master_key = (cell.row, cell.col)
        for dr in range(r_span):
            for dc in range(c_span):
                key = (cell.row + dr, cell.col + dc)
                if key in covered and key != master_key:
                    warnings.append(
                        'Overlapping merge at [{0},{1}]: {2} vs {3}. Skipped.'.format(
                            key[0], key[1], covered[key], master_key))
                else:
                    covered[key] = master_key

_BORDER_PRIORITY = {
    'thick': 10, 'double': 10,
    'medium': 7, 'mediumDashed': 6, 'mediumDashDot': 5, 'mediumDashDotDot': 5,
    'thin': 3, 'dashed': 2, 'dashDot': 2, 'slantDashDot': 2,
    'hair': 1, 'dotted': 1, 'dashDotDot': 1,
}

def _bp(style):
    """Border priority — higher means thicker/more prominent."""
    return _BORDER_PRIORITY.get(style or '', 0)

def _unified_border_pass(cells, num_rows, num_cols):
    hz_borders = [[None]*num_cols for _ in range(num_rows + 1)]
    vt_borders = [[None]*(num_cols + 1) for _ in range(num_rows)]

    # 1. Gather specified borders — thicker style wins on shared edges.
    # Excel uses thicker-wins semantics when adjacent cells specify the same
    # shared edge with conflicting styles (e.g. cell A border_bottom='thick',
    # cell B border_top='thin' -> the shared line is 'thick').
    # Without priority, the last cell processed silently downgrades the weight.
    for cell in cells:
        if cell.border_top:
            r, c = cell.row, cell.col
            if _bp(cell.border_top) > _bp(hz_borders[r][c]):
                hz_borders[r][c] = cell.border_top
        if cell.border_bottom:
            r, c = cell.row + 1, cell.col
            if _bp(cell.border_bottom) > _bp(hz_borders[r][c]):
                hz_borders[r][c] = cell.border_bottom
        if cell.border_left:
            r, c = cell.row, cell.col
            if _bp(cell.border_left) > _bp(vt_borders[r][c]):
                vt_borders[r][c] = cell.border_left
        if cell.border_right:
            r, c = cell.row, cell.col + 1
            if _bp(cell.border_right) > _bp(vt_borders[r][c]):
                vt_borders[r][c] = cell.border_right

    # 2. Write unified borders back; None means no Excel border on that edge.
    for cell in cells:
        cell.border_top    = hz_borders[cell.row][cell.col]
        cell.border_bottom = hz_borders[cell.row + 1][cell.col]
        cell.border_left   = vt_borders[cell.row][cell.col]
        cell.border_right  = vt_borders[cell.row][cell.col + 1]


def _gobble_signature(cell):
    """Signature for gobble pass: fill + top/bottom borders only.

    Left/right borders are checked separately as hard stops in the gobble
    loop (they indicate column separators).  The signature captures the
    horizontal band-continuation pattern: same fill colour and same
    top/bottom border style means the cells belong to the same visual row
    band and can be merged.
    """
    if cell is None:
        return (None, None, None)
    return (cell.fill_rgb, cell.border_top, cell.border_bottom)


def _format_signature(cell):
    """Hashable tuple of visual format for row pattern inheritance.

    Two source cells in different rows must share the same font size, font
    style (bold + italic + name), fill colour, and top/bottom borders to be
    considered matching.  This ensures row pattern inheritance only applies
    when the rows are visually identical.
    """
    if cell is None:
        return (None, None, None, None, None, None, None)
    return (cell.fill_rgb, cell.border_top, cell.border_bottom,
            cell.font_size, bool(cell.font_bold), bool(cell.font_italic),
            cell.font_name)


def _apply_row_pattern_inheritance(cells, grid, overflow_spans):
    """Cap overflow spans to match earlier rows with same format at same column."""
    if not overflow_spans:
        return

    col_history = {}
    for (row, col), span in sorted(overflow_spans.items()):
        sig = _format_signature(grid.get((row, col)))
        col_history.setdefault(col, []).append((row, span, sig))

    for col, entries in col_history.items():
        for idx, (row, span, sig) in enumerate(entries):
            cap_span = None
            for prev_idx in range(idx - 1, -1, -1):
                prev_row, prev_span, prev_sig = entries[prev_idx]
                if prev_sig == sig:
                    cap_span = prev_span
                    break

            if cap_span is not None and span > cap_span:
                cell = grid.get((row, col))
                if cell is None:
                    continue
                r_span = cell.merge_span[0] if cell.merge_span else 1
                old_c_span = cell.merge_span[1] if cell.merge_span else 1
                for release_col in range(col + cap_span, col + old_c_span):
                    released = grid.get((row, release_col))
                    if released is not None and released.is_spanned:
                        released.is_master = True
                        released.is_spanned = False
                cell.merge_span = (r_span, cap_span)
                overflow_spans[(row, col)] = cap_span


def _neighbor_allows_extension(neighbor, source_cell, source_sig):
    """
    Decide whether *neighbor* can be absorbed into source_cell's merge.

    Returns a tuple (allowed: bool, absorbs_right_border: bool).

    Each condition is checked one by one.  The first failure is a hard stop
    that immediately returns (False, False).  Only when every condition passes
    is the merge allowed.

    Conditions in order:
      1. neighbor is None → only allowed when source itself has no fill and no
         top/bottom borders (i.e. source_sig matches the empty signature).
      2. Neighbor must be a master cell (not already spanned).
      3. Neighbor must have no value (empty cells only).
      4. Neighbor must have NO left border — left border = column separator → hard stop.
      5. Neighbor fill must exactly match source fill.
      6. Neighbor border_top must exactly match source border_top → hard stop if different.
      7. Neighbor border_bottom must exactly match source border_bottom → hard stop if different.
      8. At this point all conditions pass.
         - If neighbor has no right border → normal extension.
         - If neighbor has a right border → extension allowed; caller must
           transfer that right border to the far-right edge of the merged region.
    """
    # ── Condition 1: None neighbor ───────────────────────────────────────────
    if neighbor is None:
        # A missing cell has signature (None, None, None).
        # Only allow if the source also has no fill and no top/bottom borders.
        allowed = (source_sig == _gobble_signature(None))
        return allowed, False

    # ── Condition 2: must be a live master cell ──────────────────────────────
    if neighbor.is_spanned or not neighbor.is_master:
        return False, False

    # ── Condition 3: must have no value ─────────────────────────────────────
    if str(neighbor.value).strip():
        return False, False

    # ── Condition 4: no left border (column separator = hard stop) ───────────
    if neighbor.border_left:
        return False, False

    # ── Condition 5: fill must match source ──────────────────────────────────
    if neighbor.fill_rgb != source_cell.fill_rgb:
        return False, False

    # ── Condition 6: top border must match source (hard stop) ────────────────
    if neighbor.border_top != source_cell.border_top:
        return False, False

    # ── Condition 7: bottom border must match source (hard stop) ─────────────
    if neighbor.border_bottom != source_cell.border_bottom:
        return False, False

    # ── All conditions passed ────────────────────────────────────────────────
    # If there is a right border on the neighbor, we absorb it and the caller
    # must apply it to the final merged edge instead.
    absorbs_right = bool(neighbor.border_right)
    return True, absorbs_right


def _dynamic_overflow_merge(cells, num_rows, num_cols, col_widths_ft, default_font_size_pt):
    """
    Phase 5: revised overflow merge logic.

    Changes from prior version:
    A. Wrapped / already-merged cells are no longer skipped.  They may extend
       further if the estimated sheet-view text width still exceeds available width.
    B. The hard stop on any right border is replaced by a softer rule: if the
       candidate neighbor has matching fill + no value + matching top/bottom +
       right border only (no left border), the merge is allowed and the right
       border is transferred to the final merged edge.
    C. If text does NOT require more width, no merge happens regardless of cell state.
    """
    grid = {(c.row, c.col): c for c in cells}

    baseline_ft = _pt_to_feet(default_font_size_pt) if default_font_size_pt else _pt_to_feet(11.0)
    if baseline_ft <= 0:
        baseline_ft = 0.015

    overflow_spans = {}  # (row, col) -> final_col_span for Pass 2

    for cell in cells:
        if not cell.is_master:
            continue

        val_str = str(cell.value).strip()
        if not val_str:
            continue

        # Estimate physical text width for sheet view
        char_count = len(val_str)
        mdw_px = _get_mdw(cell.font_name)
        baseline_char_width_ft = _chars_to_feet(1.0, mdw_px) * 1.10  # 10% safety margin

        font_scale = 1.0
        if cell.font_size and baseline_ft > 0:
            font_scale = cell.font_size / baseline_ft

        text_width_ft = char_count * baseline_char_width_ft * font_scale

        # Current available width (accounts for existing merge span)
        current_c_span = cell.merge_span[1] if cell.merge_span else 1
        available_width_ft = sum(
            col_widths_ft[cell.col + i]
            for i in range(current_c_span)
            if (cell.col + i) < num_cols
        )

        # Gate: only merge when text width actually requires it (Phase 5 rule C)
        if text_width_ft <= available_width_ft:
            continue

        r_span      = cell.merge_span[0] if cell.merge_span else 1
        source_sig  = _gobble_signature(cell)
        added_span  = 0
        # Track whether we absorbed a right border so we can move it to the edge
        pending_right_border = None

        c_check = cell.col + current_c_span
        while c_check < num_cols and available_width_ft < text_width_ft:
            neighbor = grid.get((cell.row, c_check))

            allowed, absorbs_right = _neighbor_allows_extension(
                neighbor, cell, source_sig)

            if not allowed:
                break

            available_width_ft += col_widths_ft[c_check]
            added_span += 1

            if neighbor is not None:
                # If this neighbor had a right border, remember it for the edge
                if absorbs_right:
                    pending_right_border = neighbor.border_right
                    neighbor.border_right = None  # clear interior right border
                neighbor.is_master  = False
                neighbor.is_spanned = True
            else:
                dummy = CellData(row=cell.row, col=c_check, is_master=False, is_spanned=True)
                cells.append(dummy)
                grid[(cell.row, c_check)] = dummy

            c_check += 1

        if added_span > 0:
            new_span = current_c_span + added_span
            cell.merge_span = (r_span, new_span)
            overflow_spans[(cell.row, cell.col)] = new_span

            # Apply the absorbed right border to the far-right edge cell of the merge
            if pending_right_border is not None:
                cell.border_right = pending_right_border

    # Pass 2: cap spans based on earlier rows with same format at same column
    _apply_row_pattern_inheritance(cells, grid, overflow_spans)


def build_grid(parsed_data, cell_images=None, options=None):
    options = options or {}
    warnings  = list(parsed_data.get('warnings', []))
    num_rows  = parsed_data['rows']
    num_cols  = parsed_data['cols']

    default_font_name = parsed_data.get('default_font_name', 'Calibri') or 'Calibri'
    default_font_size = parsed_data.get('default_font_size', 11.0)      or 11.0
    mdw_px            = _get_mdw(default_font_name)

    # MM-based scaling logic (User Request)
    base_mm = options.get('base_text_size_mm')
    final_text_scale = TEXT_SIZE_SCALE
    final_row_scale  = ROW_HEIGHT_SCALE
    final_col_scale  = COL_WIDTH_SCALE
    
    if base_mm and float(base_mm) > 0:
        # Excel default point size to mm: (pt / 72.0) * 25.4
        excel_default_mm = default_font_size * (25.4 / 72.0)
        # Calculate scale factor to reach target mm from Excel default
        auto_scale = float(base_mm) / excel_default_mm
        # Apply this scale factor proportionally to all dimensions
        final_text_scale *= auto_scale
        final_row_scale  *= auto_scale
        final_col_scale  *= auto_scale
        
        # Update options in-place so view_generator sees the final scales
        options['text_scale'] = final_text_scale
        options['row_scale']  = final_row_scale
        options['col_scale']  = final_col_scale
        
        warnings.append('Applied mm-based scaling: {0}mm (Scale: {1:.3f})'.format(base_mm, auto_scale))
    mdw_px            = _get_mdw(default_font_name)

    # Column Widths
    cw_ft_pre = parsed_data.get('col_widths_ft')
    raw_widths = parsed_data.get('col_widths', [])
    col_widths_ft = []
    
    if cw_ft_pre:
        for w_ft in cw_ft_pre:
            col_widths_ft.append(w_ft * final_col_scale)
    else:
        for i in range(num_cols):
            w_chars = raw_widths[i] if (i < len(raw_widths) and raw_widths[i]) else 8.43
            w_ft    = _chars_to_feet(w_chars, mdw_px, scale=final_col_scale)
            col_widths_ft.append(w_ft)

    # Row Heights
    rh_ft_pre = parsed_data.get('row_heights_ft')
    raw_heights = parsed_data.get('row_heights', [])
    row_heights_ft = []
    
    if rh_ft_pre:
        for h_ft in rh_ft_pre:
            row_heights_ft.append(h_ft * final_row_scale)
    else:
        for i in range(num_rows):
            h_pt = (raw_heights[i] if (i < len(raw_heights) and raw_heights[i]
                                       and raw_heights[i] > 0) else 15.0)
            h_ft = _pt_to_feet(h_pt, scale=final_row_scale)
            row_heights_ft.append(h_ft)

    # Image mapping
    cell_images = cell_images or {}
    image_count = 0

    cells = []
    for raw in parsed_data.get('cells', []):
        r = raw['row']
        c = raw['col']
        if r >= num_rows or c >= num_cols:
            continue
        merge_span = raw.get('merge_span')
        if merge_span is not None:
            rs = max(1, min(merge_span[0], num_rows - r))
            cs = max(1, min(merge_span[1], num_cols - c))
            merge_span = (rs, cs)
        raw_fs       = raw.get('font_size')
        font_size_ft = _pt_to_feet_exact(raw_fs * final_text_scale) if (raw_fs is not None and raw_fs > 0) else None
        
        # Resolve font name with fallback
        raw_font_name = raw.get('font_name')
        resolved_font = _resolve_font_name(raw_font_name) if raw_font_name else None
        
        # Check for image at this cell
        img_path = cell_images.get((r, c))
        is_img = img_path is not None
        if is_img:
            image_count += 1
        
        cells.append(CellData(
            row=r, col=c, value=raw.get('value',''),
            is_master=raw.get('is_master', True),
            is_spanned=raw.get('is_spanned', False),
            merge_span=merge_span,
            fill_rgb=raw.get('fill_rgb'), font_name=resolved_font,
            font_bold=raw.get('font_bold', False),
            font_italic=raw.get('font_italic', False),
            font_underline=raw.get('font_underline', False),
            font_size=font_size_ft, font_color=raw.get('font_color'),
            border_top=raw.get('border_top'), border_bottom=raw.get('border_bottom'),
            border_left=raw.get('border_left'), border_right=raw.get('border_right'),
            h_align=raw.get('h_align'), v_align=raw.get('v_align'),
            wrap_text=raw.get('wrap_text', False),
            text_rotation=raw.get('text_rotation', 0) or 0,
            image_path=img_path, is_image=is_img,
        ))

    _validate_merges(cells, warnings)

    # Simulate Excel text overflow into adjacent empty cells
    if ENABLE_DYNAMIC_MERGE:
        _dynamic_overflow_merge(cells, num_rows, num_cols, col_widths_ft, default_font_size)

    # Resolve border conflicts geometrically
    _unified_border_pass(cells, num_rows, num_cols)
    
    if image_count > 0:
        warnings.append('Found {0} image(s) in Excel sheet.'.format(image_count))

    return GridModel(
        rows=num_rows, cols=num_cols, cells=cells,
        col_widths_ft=col_widths_ft, row_heights_ft=row_heights_ft,
        default_font_name=default_font_name,
        default_font_size_pt=default_font_size,
        warnings=warnings,
    )


# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 3: SCHEDULE CREATOR
# ══════════════════════════════════════════════════════════════════════════════

def _find_mark_schedulable_field(defn):
    mark_id = ElementId(BuiltInParameter.ALL_MODEL_MARK)
    for sf in defn.GetSchedulableFields():
        if sf.ParameterId == mark_id:
            return sf
    return None


def _get_header_section(schedule):
    return schedule.GetTableData().GetSectionData(SectionType.Header)


def _get_hidden_graphics_style_id(doc):
    """Get the ElementId of the '<Hidden>' GraphicsStyle for suppressing grid lines.

    Tries multiple strategies because the name varies by Revit version/locale
    and system styles can have negative ElementIds.
    """
    _hidden_names = {'<Hidden>', 'Hidden', '<Invisible>', 'Invisible',
                     '<No Lines>', 'No Lines', '<None>'}
    try:
        collector = FilteredElementCollector(doc).OfClass(GraphicsStyle)
        for gs in collector:
            if gs.Name in _hidden_names:
                return gs.Id
    except Exception:
        pass
    # Fallback: Revit's system hidden-line style is ElementId(-3) in most builds
    try:
        candidate = doc.GetElement(ElementId(-3))
        if candidate is not None:
            return ElementId(-3)
    except Exception:
        pass
    return None


def create_generic_model_schedule(doc, schedule_name, num_cols, col_widths_ft):
    warnings = []

    def _w(i):
        if 0 <= i < len(col_widths_ft) and col_widths_ft[i]:
            return col_widths_ft[i]
        return 0.0833

    total_width_ft = sum(w for w in col_widths_ft if w) if col_widths_ft else _w(0)

    # Get the hidden graphics style ID before transaction
    hidden_gs_id = _get_hidden_graphics_style_id(doc)

    t = Transaction(doc, 'Create Schedule')
    t.Start()
    try:
        cat_id   = ElementId(BuiltInCategory.OST_GenericModel)
        schedule = ViewSchedule.CreateSchedule(doc, cat_id)
        schedule.Name = schedule_name
        defn = schedule.Definition

        # FIX Issue 3: Hide the body column header row ("Mark" field name row)
        try:
            defn.ShowHeaders = False
        except Exception:
            pass

        # Show schedule grid lines — conditionally controlled by UI checkbox
        try:
            defn.ShowGridLines = options.get('show_empty_gridlines', False)
        except Exception:
            pass

        mark_sf    = _find_mark_schedulable_field(defn)
        mark_field = None
        if mark_sf is not None:
            mark_field = defn.AddField(mark_sf)
        else:
            try:
                mark_field = defn.AddField(ScheduleFieldType.Count)
            except Exception:
                pass

        if mark_field is not None:
            fid = mark_field.FieldId
            try:
                defn.AddFilter(ScheduleFilter(fid, ScheduleFilterType.HasValue))
            except Exception:
                try:
                    defn.AddFilter(ScheduleFilter(fid, ScheduleFilterType.NotEqual, ''))
                except Exception:
                    pass
            try:
                defn.AddFilter(ScheduleFilter(fid, ScheduleFilterType.HasNoValue))
            except Exception:
                try:
                    defn.AddFilter(ScheduleFilter(fid, ScheduleFilterType.Equal, ''))
                except Exception:
                    pass

        # CRITICAL FIX: Clear the default title cell content
        # The schedule name should only appear in Project Browser, not in cells
        try:
            header_data = schedule.GetTableData().GetSectionData(SectionType.Header)
            if header_data and header_data.NumberOfRows > 0:
                first_row = header_data.FirstRowNumber
                first_col = header_data.FirstColumnNumber
                # Set to single space to maintain cell structure without visible text
                header_data.SetCellText(first_row, first_col, " ")
        except Exception:
            pass

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise Exception('Create schedule failed: {0}'.format(str(ex)))

    cols_inserted = 1
    for i in range(1, num_cols):
        t = Transaction(doc, 'Insert Header Col {0}'.format(i))
        t.Start()
        try:
            hd       = _get_header_section(schedule)
            last_abs = hd.FirstColumnNumber + hd.NumberOfColumns - 1
            hd.InsertColumn(last_abs)
            t.Commit()
            cols_inserted += 1
        except Exception as ex:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            warnings.append('Insert col {0} failed: {1}'.format(i, str(ex)))

    t = Transaction(doc, 'All Column Widths')
    t.Start()
    try:
        hd        = _get_header_section(schedule)
        first_col = hd.FirstColumnNumber
        for i in range(min(num_cols, hd.NumberOfColumns)):
            try:
                hd.SetColumnWidth(first_col + i, _w(i))
            except Exception:
                pass
        try:
            defn = schedule.Definition
            if defn.GetFieldCount() > 0:
                defn.GetField(0).ColumnWidth = total_width_ft
        except Exception:
            pass
        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        warnings.append('Column widths tx failed: {0}'.format(str(ex)))

    if warnings:
        try:
            schedule.__dict__['_creator_warnings'] = warnings
        except Exception:
            pass

    return schedule


# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 4: REVIT WRITER  (V19 — GraphicsStyle borders + alignment fix)
# ══════════════════════════════════════════════════════════════════════════════

def _collect_line_styles_dict(doc):
    """Return {name: ElementId} for every line style usable on table cell
    borders / detail lines.

    Single source of truth used by BOTH BorderGraphicsStyles (the cache) AND
    _get_project_line_style_names (the dialog dropdown).  Sharing the source
    guarantees that names the user picks in the dialog will resolve to a real
    cache entry — eliminates the silent "lookup-by-name returned None" bug.

    Strategy
    --------
    1. Canonical: enumerate subcategories of OST_Lines via
       doc.Settings.Categories.get_Item(OST_Lines).SubCategories and call
       sub.GetGraphicsStyle(Projection).  This yields actual line styles —
       NOT the same-named Projection styles that Walls / Doors / Floors /
       etc. expose.  Without this filter Revit silently rejects non-line
       ids when assigned to BorderTopLineStyle/etc., which is the root
       cause of the "configured mapping appears to do nothing" bug.

    2. Fallback: if step 1 yielded nothing (older Revit, restricted
       template), iterate FilteredElementCollector(GraphicsStyle) filtered
       to Projection type.  May include non-line styles but at least
       produces a populated dictionary.

    Both bracket and plain forms ('<Thin Lines>' and 'Thin Lines') are
    registered for every entry so callers don't have to guess.
    """
    result = {}
    if doc is None:
        return result

    try:
        from Autodesk.Revit.DB import GraphicsStyleType as _GST
        _gst_proj = _GST.Projection
    except Exception:
        _gst_proj = None

    def _register(name, eid):
        if not name or eid is None:
            return
        result[name] = eid
        if name.startswith('<') and name.endswith('>'):
            plain = name[1:-1]
            if plain and plain not in result:
                result[plain] = eid
        else:
            bracketed = '<' + name + '>'
            if bracketed not in result:
                result[bracketed] = eid

    # ── Step 1: canonical SubCategories of OST_Lines ─────────────────────────
    try:
        lines_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        if lines_cat is not None:
            try:
                if _gst_proj is not None:
                    top_gs = lines_cat.GetGraphicsStyle(_gst_proj)
                    if top_gs is not None:
                        _register(lines_cat.Name, top_gs.Id)
            except Exception:
                pass
            try:
                for sub in lines_cat.SubCategories:
                    try:
                        if _gst_proj is None:
                            continue
                        gs = sub.GetGraphicsStyle(_gst_proj)
                        if gs is None:
                            continue
                        _register(sub.Name, gs.Id)
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass

    # ── Step 2: fallback collector (only runs if Step 1 was empty) ──────────
    if not result:
        try:
            for gs in FilteredElementCollector(doc).OfClass(GraphicsStyle):
                try:
                    if not gs.Name:
                        continue
                    if _gst_proj is not None:
                        try:
                            if gs.GraphicsStyleType != _gst_proj:
                                continue
                        except Exception:
                            pass
                    _register(gs.Name, gs.Id)
                except Exception:
                    pass
        except Exception:
            pass

    return result


class BorderGraphicsStyles(object):
    """
    Cache of GraphicsStyle ElementIds for schedule cell border control.
    """
    _EXCEL_TO_REVIT = {
        'thin':             'Thin Lines',
        'hair':             'Thin Lines',
        'medium':           'Medium Lines',
        'thick':            'Wide Lines',
        'double':           'Wide Lines',
        'dashed':           'Thin Lines',
        'mediumDashed':     'Medium Lines',
        'dotted':           'Thin Lines',
        'dashDot':          'Thin Lines',
        'mediumDashDot':    'Medium Lines',
        'dashDotDot':       'Thin Lines',
        'mediumDashDotDot': 'Medium Lines',
        'slantDashDot':     'Thin Lines',
    }

    # Names used as the canonical hidden-line style (no border)
    _HIDDEN_NAMES = {'<Hidden>', 'Hidden', '<Invisible>', 'Invisible',
                     '<No Lines>', 'No Lines', '<None>'}

    # Keyword hints used to auto-pick a visually-matching project line style
    # for dashed/dotted Excel borders when the user hasn't manually configured
    # a mapping.  These are scanned against actual project GraphicsStyle names.
    _AUTO_KEYWORD_HINTS = {
        'dashed':           ['dashed', 'dash'],
        'mediumDashed':     ['dashed', 'dash'],
        'dotted':           ['dotted', 'dot'],
        'dashDot':          ['dash-dot', 'dashdot', 'dash dot'],
        'mediumDashDot':    ['dash-dot', 'dashdot', 'dash dot'],
        'dashDotDot':       ['dash-dot-dot', 'dashdotdot'],
        'mediumDashDotDot': ['dash-dot-dot', 'dashdotdot'],
        'slantDashDot':     ['dash-dot', 'dashdot', 'slant'],
    }

    def __init__(self, doc, override_map=None):
        """
        override_map: optional {xl_border_style: revit_line_style_name} dict
        from the user's Configure Line Styles dialog. Takes precedence over
        the built-in _EXCEL_TO_REVIT mapping.
        """
        # Single source of truth for {name: ElementId} — same dict the dialog
        # dropdown is sourced from, so any name the user picks is guaranteed
        # to resolve back to a valid ElementId in this cache.
        self._cache = _collect_line_styles_dict(doc)

        # Start with the class-level defaults, then layer user overrides on top
        self._excel_to_revit = dict(self._EXCEL_TO_REVIT)
        if override_map:
            for xl_s, rv_name in override_map.items():
                if rv_name:
                    self._excel_to_revit[xl_s] = rv_name

        # Auto-upgrade dashed/dotted defaults: if the project has a line style
        # whose name contains a matching keyword and the user hasn't explicitly
        # mapped this Excel style, promote that project style as the default.
        try:
            override_keys = set((override_map or {}).keys())
            project_names = list(self._cache.keys())
            for xl_style, hints in self._AUTO_KEYWORD_HINTS.items():
                if xl_style in override_keys:
                    continue  # User already configured this style
                for hint in hints:
                    matched = None
                    for nm in project_names:
                        nm_low = nm.lower().strip('<>')
                        if hint in nm_low:
                            matched = nm
                            break
                    if matched:
                        self._excel_to_revit[xl_style] = matched
                        break
        except Exception:
            pass

        # ── Step 1: Normalize bracket forms ───────────────────────────────────
        # Revit stores line styles as either '<Thin Lines>' (angle-bracket) or
        # 'Thin Lines' (plain) depending on the project template. Add both forms
        # to the cache so all downstream lookups work regardless of which variant
        # the current project uses.
        for _nm in list(self._cache.keys()):
            if _nm.startswith('<') and _nm.endswith('>'):
                _plain = _nm[1:-1]
                if _plain not in self._cache:
                    self._cache[_plain] = self._cache[_nm]
            else:
                _bracketed = '<' + _nm + '>'
                if _bracketed not in self._cache:
                    self._cache[_bracketed] = self._cache[_nm]

        # ── Step 2: Ensure '<Hidden>' / 'Hidden' are accessible ──────────────
        if '<Hidden>' not in self._cache:
            for _nm in self._HIDDEN_NAMES:
                if _nm in self._cache:
                    self._cache['<Hidden>'] = self._cache[_nm]
                    break
        # Last-resort: Revit's system hidden-line style is ElementId(-3)
        if '<Hidden>' not in self._cache:
            try:
                if doc.GetElement(ElementId(-3)) is not None:
                    self._cache['<Hidden>'] = ElementId(-3)
            except Exception:
                pass

        # ── Step 3: Ensure 'Thin Lines' / 'Medium Lines' / 'Wide Lines' exist ─
        # After bracket normalization above, '<Thin Lines>' → 'Thin Lines' has
        # already been added if the project uses angle-bracket names.  This block
        # only runs when NEITHER form was found (unusual/non-standard template).
        _need_thin   = 'Thin Lines'   not in self._cache
        _need_medium = 'Medium Lines' not in self._cache
        _need_wide   = 'Wide Lines'   not in self._cache

        if _need_thin or _need_medium or _need_wide:
            _skip = self._HIDDEN_NAMES | {'<Hidden>', 'Hidden'}
            # Prefer styles that sound like border lines; fall back to any non-hidden
            _ranked = sorted(
                [(name, gid) for name, gid in self._cache.items()
                 if name not in _skip and gid is not None],
                key=lambda pair: (
                    0 if any(kw in pair[0].lower()
                             for kw in ('thin',)) else
                    1 if any(kw in pair[0].lower()
                             for kw in ('medium', 'lw-2', 'lw-3')) else
                    2 if any(kw in pair[0].lower()
                             for kw in ('wide', 'lw-4', 'lw-5', 'lw-6')) else
                    3
                )
            )
            _ids = [gid for _, gid in _ranked]
            if _need_thin   and _ids:
                self._cache['Thin Lines']   = _ids[0]
                self._cache['<Thin Lines>'] = _ids[0]
            if _need_medium and _ids:
                self._cache['Medium Lines']   = _ids[min(1, len(_ids)-1)]
                self._cache['<Medium Lines>'] = _ids[min(1, len(_ids)-1)]
            if _need_wide   and _ids:
                self._cache['Wide Lines']   = _ids[min(2, len(_ids)-1)]
                self._cache['<Wide Lines>'] = _ids[min(2, len(_ids)-1)]

    def get(self, style_name, default=None):
        """Robust lookup: tries direct, bracket-variant, then case-insensitive match."""
        if style_name is None:
            return default
        if style_name in self._cache:
            return self._cache[style_name]
        # Try the opposite bracket variant
        if style_name.startswith('<') and style_name.endswith('>'):
            plain = style_name[1:-1]
            if plain in self._cache:
                return self._cache[plain]
        else:
            bracketed = '<' + style_name + '>'
            if bracketed in self._cache:
                return self._cache[bracketed]
        # Case-insensitive last resort (handles user typing variations from dialog)
        try:
            sl = style_name.lower().strip('<>')
            for k, v in self._cache.items():
                if k.lower().strip('<>') == sl:
                    return v
        except Exception:
            pass
        return default

    def get_border_id(self, excel_style, use_hidden_for_empty=False):
        """Map an Excel border style string to a Revit GraphicsStyle ElementId.

        When use_hidden_for_empty=True and excel_style is empty/None/'none',
        returns ElementId.InvalidElementId as a sentinel meaning "suppress
        border override" — preserves the previously-shipped invisible-lines
        behaviour for cells that have no Excel border on a given edge.
        For non-empty Excel styles, looks up the configured Revit line style;
        if not found in cache, falls back to 'Thin Lines' so the border at
        least renders rather than disappearing silently.
        """
        # Excel's explicit-no-border ('none') is normalised to None upstream
        # in _bs(), but treat it defensively here too in case raw cell data
        # arrives via another code path.
        if not excel_style or excel_style == 'none':
            if use_hidden_for_empty:
                return ElementId.InvalidElementId  # sentinel: suppress border override
            return None
        if excel_style == '<Hidden>':
            return self.get('<Hidden>')
        revit_name = self._excel_to_revit.get(excel_style, 'Thin Lines')
        gid = self.get(revit_name)
        if gid is None:
            # Fallback: keep the border visible with the closest available style
            for fallback in ('Thin Lines', '<Thin Lines>', 'Medium Lines', 'Wide Lines'):
                gid = self.get(fallback)
                if gid is not None:
                    break
        return gid


def _flag(o, *names):
    # V23 FIX (Bug 4): Removed early `return True` so ALL candidate flag names
    # are attempted.  In IronPython, setattr() on a .NET object never raises
    # for unknown attribute names — it silently creates a Python-side attribute
    # that Revit ignores.  The old `return True` fired on the very first name
    # regardless of whether Revit actually registered the flag.  Now every name
    # is tried; valid .NET properties take effect, invalid ones are harmless.
    _any = False
    for name in names:
        try:
            setattr(o, name, True)
            _any = True
        except Exception:
            pass
    return _any


def _color(rgb, fix_white_for_sheet=True):
    if not rgb:
        return None
    try:
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
        # FIX: Pure white (255,255,255) reverts to black on sheet view
        if fix_white_for_sheet and r == 255 and g == 255 and b == 255:
            r, g, b = 254, 254, 254
        return RevitColor(r, g, b)
    except Exception:
        return None


def _halign(s):
    if not _HAS_H or not s:
        return None
    return {
        'left':             HorizontalAlignmentStyle.Left,
        'center':           HorizontalAlignmentStyle.Center,
        'right':            HorizontalAlignmentStyle.Right,
        'fill':             HorizontalAlignmentStyle.Left,
        'justify':          HorizontalAlignmentStyle.Left,
        'distributed':      HorizontalAlignmentStyle.Center,
        'general':          None,
        'centerContinuous': HorizontalAlignmentStyle.Center,
    }.get(s)


def _valign(s):
    if not _HAS_V or not s:
        return None
    return {
        'top':         VerticalAlignmentStyle.Top,
        'center':      VerticalAlignmentStyle.Middle,
        'bottom':      VerticalAlignmentStyle.Bottom,
        'justify':     VerticalAlignmentStyle.Middle,
        'distributed': VerticalAlignmentStyle.Middle,
    }.get(s)


_CHAR_MAP = {
    # Wingdings/Symbol-encoded characters that arrive as Latin-1 codepoints
    # because Excel stored them under a symbol font.  Revit cannot render
    # these raw codepoints in a non-symbol font, so remap to real Unicode.
    u'ü': u'✓',  # Wingdings tick (chr 252)    -> U+2713 check
    u'û': u'✓',  # alt Wingdings tick (chr 251) -> U+2713 check
    u'ý': u'✗',  # Wingdings ballot X (chr 253) -> U+2717 cross
    u'þ': u'✘',  # Wingdings heavy X  (chr 254) -> U+2718 heavy cross
    # Normalise check-mark variants to one canonical tick for consistency
    u'✔': u'✓',  # heavy check (U+2714) -> U+2713
    u'√': u'✓',  # sqrt (U+221A)        -> U+2713
    # All other Unicode (multiplication signs, bullets, dashes, smart quotes,
    # degree, plus-minus, legal marks, etc.) is now PRESERVED verbatim --
    # modern Revit fonts render BMP characters correctly, so the previous
    # ASCII collapse (mapping bullets to -, degree to deg, plus-minus to +/-,
    # trademark to TM, multiplication signs to x, en/em dash to -/--, smart
    # quotes to straight quotes) has been removed.
}


def _sanitize_cell_text(value):
    """
    IronPython 2.7 mangles non-BMP and Wingdings/Symbol-encoded Unicode characters
    when passed to Revit's SetCellText.  Map the most common offenders to safe
    ASCII/Latin equivalents that Revit schedule cells can display cleanly.
    """
    if not value:
        return value
    try:
        if isinstance(value, bytes):
            value = value.decode('utf-8', errors='replace')
        for src, dst in _CHAR_MAP.items():
            if src in value:
                value = value.replace(src, dst)
        # Strip remaining non-BMP chars that Revit can't render (codepoint > U+FFFF)
        value = u''.join(
            ch if ord(ch) <= 0xFFFF else '?' for ch in value
        )
    except Exception:
        pass
    return value


def _get_header_view(view):
    return view.GetTableData().GetSectionData(SectionType.Header)


def _is_numeric_value(v):
    """Return True when v looks like a pure number."""
    if not v:
        return False
    try:
        float(str(v).replace(',', ''))
        return True
    except (ValueError, TypeError):
        return False


def _is_rotated(text_rotation):
    """
    Convert an Excel textRotation integer to a Revit IsRotated bool.

    Excel textRotation values:
        0         = normal horizontal
        1 – 90    = CCW rotation in degrees (90 = fully vertical CCW)
        91 – 180  = CW rotation (91 = 1° CW, 180 = 90° CW)
        255       = stacked vertical text (each character upright, top-to-bottom)

    Revit TableCellStyle.IsRotated = True rotates text 90° CCW (vertical).
    We map to True whenever the angle is meaningfully non-horizontal:
        - CCW ≥ 45°  (value 45–90)
        - CW  ≥ 45°  (value 136–180, i.e. 90 - (val-90) = val-90 ≥ 45)
        - Stacked    (255)
    """
    try:
        r = int(text_rotation)
    except (TypeError, ValueError):
        return False
    if r == 255:          # stacked — treat as vertical
        return True
    if 45 <= r <= 90:     # meaningful CCW rotation
        return True
    if 136 <= r <= 180:   # meaningful CW rotation (≥ 45°)
        return True
    return False


def _apply_style(hd, abs_row, abs_col, cd, border_styles):
    fill   = cd.get('fill_rgb')
    fname  = cd.get('font_name')
    bold   = cd.get('font_bold',       False)
    italic = cd.get('font_italic',     False)
    uline  = cd.get('font_underline',  False)
    fsize  = cd.get('font_size')
    fcolor = cd.get('font_color')
    bt     = cd.get('border_top')
    bb     = cd.get('border_bottom')
    bl     = cd.get('border_left')
    br_    = cd.get('border_right')
    ha     = cd.get('h_align')
    va     = cd.get('v_align')
    wrap   = cd.get('wrap_text', False)
    trot   = cd.get('text_rotation', 0) or 0
    value  = cd.get('value', '')

    # Resolve effective horizontal alignment
    if not ha or ha == 'general':
        ha = 'right' if _is_numeric_value(value) else 'left'

    has_explicit = any([fill, fname, bold, italic, uline, fsize, fcolor,
                        bt, bb, bl, br_, ha, va, wrap, trot])
    has_baseline = _HAS_H or _HAS_V or (border_styles is not None)

    if not has_explicit and not has_baseline:
        return

    # V24 FIX: Read-Modify-Write pattern for TableCellStyle + OverrideOptions.
    # Read the EXISTING style and override options so we MERGE flags rather than
    # replacing them.  This prevents Phase F restamp from wiping Phase E flags.
    try:
        s = hd.GetTableCellStyle(abs_row, abs_col)
    except Exception:
        s = TableCellStyle()

    try:
        o = s.GetCellStyleOverrideOptions()
    except Exception:
        o = TableCellStyleOverrideOptions()

    # 1. Background fill
    if fill:
        try:
            c = _color(fill)
            if c:
                s.BackgroundColor = c
                _flag(o, 'BackgroundColor')
        except Exception:
            pass

    # 2. Font name  (canonical flag: FontName)
    if fname:
        try:
            s.FontName = fname
        except Exception:
            pass
        _flag(o, 'Font')

    # 3. Bold  (canonical flag: Bold)
    try:
        s.IsFontBold = bool(bold)
        _flag(o, 'Bold')
    except Exception:
        pass

    # 4. Italic  (canonical flag: Italics)
    try:
        s.IsFontItalic = bool(italic)
        _flag(o, 'Italics')
    except Exception:
        pass

    # 5. Underline  (canonical flag: Underline)
    try:
        s.IsFontUnderline = bool(uline)
        _flag(o, 'Underline')
    except Exception:
        pass

    # 6. Font size  (canonical flag: FontSize)
    if fsize and fsize > 0:
        try:
            s.TextSize = fsize
        except Exception:
            pass
        _flag(o, 'FontSize')

    # 7. Font colour  (canonical flag: FontColor)
    if fcolor:
        try:
            c = _color(fcolor)
            if c:
                color_set = False
                for prop in ('FontColor', 'TextColor', 'Color'):
                    try:
                        setattr(s, prop, c)
                        color_set = True
                        break
                    except Exception:
                        pass
                if color_set:
                    _flag(o, 'FontColor')
        except Exception:
            pass

    # 8. Borders — set per-side line-style override + override flag.
    # Use only the canonical 'BorderTopLineStyle'/'BorderBottomLineStyle'/etc.
    # Adding extra flag-name variants like 'BorderTop'/'TopBorder' silently
    # creates Python attributes in IronPython but, where they do happen to be
    # real .NET booleans, they force ALL borders visible (regression seen).
    if border_styles is not None:
        try:
            _hide = not SHOW_EMPTY_GRIDLINES
            t_id = border_styles.get_border_id(bt,  use_hidden_for_empty=_hide)
            b_id = border_styles.get_border_id(bb,  use_hidden_for_empty=_hide)
            l_id = border_styles.get_border_id(bl,  use_hidden_for_empty=_hide)
            r_id = border_styles.get_border_id(br_, use_hidden_for_empty=_hide)

            if t_id is not None:
                s.BorderTopLineStyle    = t_id
                _flag(o, 'BorderTopLineStyle')

            if b_id is not None:
                s.BorderBottomLineStyle = b_id
                _flag(o, 'BorderBottomLineStyle')

            if l_id is not None:
                s.BorderLeftLineStyle   = l_id
                _flag(o, 'BorderLeftLineStyle')

            if r_id is not None:
                s.BorderRightLineStyle  = r_id
                _flag(o, 'BorderRightLineStyle')

        except Exception:
            pass

    # 9. Horizontal alignment  (canonical flag: HorizontalAlignment)
    if _HAS_H:
        try:
            rv = _halign(ha)
            if rv is not None:
                s.FontHorizontalAlignment = rv
                _flag(o, 'HorizontalAlignment')
        except Exception:
            pass

    # 10. Vertical alignment  (canonical flag: VerticalAlignment)
    if _HAS_V:
        try:
            rv = _valign(va)
            if rv is not None:
                s.FontVerticalAlignment = rv
                _flag(o, 'VerticalAlignment')
        except Exception:
            pass

    # 11. Word wrap — ALWAYS set the override flag so Revit doesn't use its default.
    # Only set True when Excel explicitly has wrap; otherwise force False.
    try:
        s.FontWordWrap = bool(wrap)
        _flag(o, 'WordWrap', 'FontWordWrap')
    except Exception:
        pass

    # 12. Text rotation — map Excel textRotation to Revit IsRotated.
    # Revit schedule cells support a 90° CCW rotation (IsRotated = True).
    # We activate it whenever the Excel angle is meaningfully non-horizontal
    # (≥ 45° CCW, ≥ 45° CW, or stacked vertical text = 255).
    # NOTE: do NOT break after the first setattr — IronPython silently creates
    # a Python attribute for unknown .NET names, so the first attempt always
    # appears to succeed even when the property doesn't exist on this version.
    try:
        rotated = _is_rotated(trot)
        set_any = False
        for prop in ('IsRotated', 'TextRotation', 'Rotation'):
            try:
                setattr(s, prop, rotated)
                set_any = True
            except Exception:
                pass
        if set_any:
            _flag(o, 'Rotation', 'IsRotated', 'TextRotation')
    except Exception:
        pass

    try:
        s.SetCellStyleOverrideOptions(o)
        hd.SetCellStyle(abs_row, abs_col, s)
    except Exception:
        pass


def _is_symbol_font(font_name):
    """Return True when font_name is a dingbat/symbol font (Wingdings, Webdings, Symbol).
    Setting FontName/FontSize/WordWrap override flags on these cells causes Revit
    to render Unicode glyphs (e.g. ✓ U+2713) as □ (U+25A1).
    """
    if not font_name:
        return False
    fn = font_name.lower().strip()
    return 'wingding' in fn or 'webding' in fn or 'symbol' in fn


def _apply_symbol_cell_style(hd, abs_row, abs_col, cd):
    """Apply ONLY Bold to a symbol cell (Wingdings/Webdings/Symbol).
    
    Setting ANY other override flag (Font, FontSize, Fill, WordWrap, etc.)
    on a cell containing Unicode symbols like \u2713 causes Revit to render
    them as \u25a1 (empty square).  The user confirmed that manually resetting
    ALL overrides and then re-applying only Bold produces correct symbols.
    This function replicates that exact behaviour.
    """
    try:
        s = hd.GetTableCellStyle(abs_row, abs_col)
    except Exception:
        s = TableCellStyle()
    o = TableCellStyleOverrideOptions()
    
    bold = cd.get('font_bold', False)
    try:
        s.IsFontBold = bool(bold)
        _flag(o, 'Bold')
    except Exception:
        pass
    
    try:
        s.SetCellStyleOverrideOptions(o)
        hd.SetCellStyle(abs_row, abs_col, s)
    except Exception:
        pass

def _restamp_borders_only(hd, abs_row, abs_col, cell, border_styles):
    """Border-only restamp for a master cell (matches V1.4 _restamp_borders).
    
    Re-reads the existing TableCellStyle, sets ONLY border properties,
    and always enables all four border override flags.
    This is a lightweight targeted pass used in Phase F.
    Does NOT touch font, fill, text size, or word wrap — preserving
    overrides from Phase E and preventing symbol glyph corruption.
    """
    if border_styles is None:
        return
    try:
        bt  = cell.border_top    if cell is not None else None
        bb  = cell.border_bottom if cell is not None else None
        bl  = cell.border_left   if cell is not None else None
        br_ = cell.border_right  if cell is not None else None

        _hide = not SHOW_EMPTY_GRIDLINES
        t_id = border_styles.get_border_id(bt,  use_hidden_for_empty=_hide)
        b_id = border_styles.get_border_id(bb,  use_hidden_for_empty=_hide)
        l_id = border_styles.get_border_id(bl,  use_hidden_for_empty=_hide)
        r_id = border_styles.get_border_id(br_, use_hidden_for_empty=_hide)

        try:
            s = hd.GetTableCellStyle(abs_row, abs_col)
        except Exception:
            s = TableCellStyle()
        o = s.GetCellStyleOverrideOptions()

        if t_id is not None:
            s.BorderTopLineStyle    = t_id
            _flag(o, 'BorderTopLineStyle')

        if b_id is not None:
            s.BorderBottomLineStyle = b_id
            _flag(o, 'BorderBottomLineStyle')

        if l_id is not None:
            s.BorderLeftLineStyle   = l_id
            _flag(o, 'BorderLeftLineStyle')

        if r_id is not None:
            s.BorderRightLineStyle  = r_id
            _flag(o, 'BorderRightLineStyle')

        s.SetCellStyleOverrideOptions(o)
        hd.SetCellStyle(abs_row, abs_col, s)
    except Exception:
        pass

def _apply_hidden_borders(hd, abs_row, abs_col, border_styles):
    """Suppress all four borders on a spanned or empty cell using '<Hidden>'."""
    if border_styles is None:
        return
    hidden = border_styles.get('<Hidden>')
    if hidden is None:
        return
    try:
        # Read-Modify-Write: preserve existing style flags
        try:
            s = hd.GetTableCellStyle(abs_row, abs_col)
        except Exception:
            s = TableCellStyle()
        try:
            o = s.GetCellStyleOverrideOptions()
        except Exception:
            o = TableCellStyleOverrideOptions()

        s.BorderTopLineStyle    = hidden
        s.BorderBottomLineStyle = hidden
        s.BorderLeftLineStyle   = hidden
        s.BorderRightLineStyle  = hidden
        _flag(o, 'BorderTopLineStyle')
        _flag(o, 'BorderBottomLineStyle')
        _flag(o, 'BorderLeftLineStyle')
        _flag(o, 'BorderRightLineStyle')
        s.SetCellStyleOverrideOptions(o)
        hd.SetCellStyle(abs_row, abs_col, s)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 4.5: IN-CELL IMAGE IMPORTER
# ══════════════════════════════════════════════════════════════════════════════


def _import_image_to_schedule_cell(doc, schedule, abs_row, abs_col, image_path):
    """
    Import an image file into a schedule header cell.
    
    Args:
        doc: Revit Document
        schedule: ViewSchedule
        abs_row: Absolute row number in header section
        abs_col: Absolute column number in header section
        image_path: Path to image file
    
    Returns:
        ElementId of created ImageType, or None on failure
    """
    LogManager.debug('[INCELL] row={0} col={1} path={2}'.format(abs_row, abs_col, image_path))
    if not image_path or not os.path.exists(image_path):
        LogManager.debug('[INCELL]   SKIP: file not found')
        return None

    try:
        # Create ImageType from file
        img_opts = ImageTypeOptions(image_path, False, ImageTypeSource.Import)
        img_opts.Resolution = 96

        image_type = ImageType.Create(doc, img_opts)
        if image_type is None:
            LogManager.debug('[INCELL]   ImageType.Create returned None')
            return None
        LogManager.debug('[INCELL]   ImageType.Create OK id={0}'.format(image_type.Id))

        # Insert into schedule header cell
        header = schedule.GetTableData().GetSectionData(SectionType.Header)
        max_r = header.NumberOfRows - 1
        max_c = header.NumberOfColumns - 1
        LogManager.debug('[INCELL]   header rows={0} cols={1} target=({2},{3})'.format(
            max_r + 1, max_c + 1, abs_row, abs_col))

        if abs_row > max_r or abs_col > max_c:
            LogManager.debug('[INCELL]   SKIP: target out of header bounds')
            return None

        # Set cell type to Graphic before inserting image
        try:
            header.SetCellType(abs_row, abs_col, CellType.Graphic)
            LogManager.debug('[INCELL]   SetCellType(Graphic) OK')
        except Exception as _e:
            LogManager.debug('[INCELL]   SetCellType(Graphic) FAILED: {0}'.format(str(_e)))

        # Insert the image
        header.InsertImage(abs_row, abs_col, image_type.Id)
        print('[INCELL]   InsertImage OK')

        return image_type.Id

    except Exception as _ex:
        print('[INCELL]   EXCEPTION: {0}'.format(str(_ex)))
        return None


# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 4.6: SHEET PLACEMENT + SHEET-VIEW APPEARANCE  (V20 NEW)
# ══════════════════════════════════════════════════════════════════════════════

def _get_active_sheet(uidoc):
    """
    Returns the ViewSheet that is currently active (if any), otherwise None.
    Handles the case where the active view IS a sheet, or is a schedule/drafting
    view that happens to be opened from a sheet context.
    """
    try:
        active_view = uidoc.ActiveView
        if active_view is None:
            return None
        if isinstance(active_view, ViewSheet):
            return active_view
        # Check if the active view is owned by a sheet
        # (some Revit versions allow opening schedule directly on a sheet)
        sheet_num = active_view.get_Parameter(
            BuiltInParameter.VIEWER_SHEET_NUMBER)
        if sheet_num is not None:
            sheet_val = sheet_num.AsString()
            if sheet_val:
                collector = FilteredElementCollector(uidoc.Document).OfClass(ViewSheet)
                for s in collector:
                    if s.SheetNumber == sheet_val:
                        return s
    except Exception:
        pass
    return None


def _get_schedule_total_size(schedule, grid_model):
    """
    Calculate the total bounding-box width and height of the schedule
    in Revit internal units (feet), based on the grid model dimensions.
    Returns (width_ft, height_ft).
    """
    width_ft  = sum(w for w in grid_model.col_widths_ft if w) if grid_model.col_widths_ft else 0.1
    height_ft = sum(h for h in grid_model.row_heights_ft if h) if grid_model.row_heights_ft else 0.05
    # Clamp to sane values
    width_ft  = max(0.01, width_ft)
    height_ft = max(0.005, height_ft)
    return width_ft, height_ft


def _place_schedule_on_sheet(doc, uidoc, sheet, schedule, grid_model,
                              explicit_pt=None):
    """
    Place the given ViewSchedule on the given ViewSheet.

    explicit_pt supplied  →  update path: place at that exact saved point.
    explicit_pt is None   →  first-time path: place at sheet origin (0, 0, 0).

    Returns the ScheduleSheetInstance, or None on failure.
    """
    if sheet is None or schedule is None:
        return None

    try:
        # Create at origin first — atomic position set inside same transaction
        target_pt = explicit_pt if explicit_pt is not None else XYZ(0.0, 0.0, 0.0)

        with Transaction(doc, 'Place Schedule on Sheet') as t:
            t.Start()
            ssi = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, XYZ(0.0, 0.0, 0.0))
            # Set exact position inside same transaction (atomic, no inter-tx drift)
            if explicit_pt is not None:
                try:
                    ssi.Point = explicit_pt
                    print("[DocLink] SSI.Point set to ({:.4f}, {:.4f})".format(explicit_pt.X, explicit_pt.Y))
                except Exception:
                    try:
                        from Autodesk.Revit.DB import ElementTransformUtils as _ETU
                        _ETU.MoveElement(doc, ssi.Id, explicit_pt)
                        print("[DocLink] SSI moved via ElementTransformUtils")
                    except Exception as _mv_ex:
                        print("[DocLink] SSI move failed: {}".format(_mv_ex))
            t.Commit()

        actual = ssi.Point
        print("[DocLink] SSI at ({:.4f},{:.4f}) target ({:.4f},{:.4f})".format(
            actual.X, actual.Y, target_pt.X, target_pt.Y))
        return ssi

    except Exception as ex:
        LogManager.debug("[DocLinkManager] _place_schedule_on_sheet EXCEPTION: {}".format(ex))
        return None




# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 4.6: SCHEDULE GLOBAL TEXT STYLE  (V22 – Sheet-View Fix)
# ══════════════════════════════════════════════════════════════════════════════

def _set_schedule_global_text_style(doc, schedule, default_size_ft, default_font_name):
    """
    V22 ROOT-CAUSE FIX for "text correct in schedule view but tiny on sheet view".

    Problem:
        TableCellStyle.TextSize per-cell override works in the schedule view
        (which is a standalone Revit view), but the SHEET VIEW renderer for
        ScheduleSheetInstance uses the schedule's own global text-size setting
        (stored as a BuiltInParameter on the ViewSchedule element) as the
        rendering baseline.  If that parameter is still at Revit's factory
        default (3/32" ≈ 2.4 mm ≈ ~6.5 pt), all text on the sheet appears tiny
        regardless of per-cell TextSize overrides.

    Fix (three complementary approaches):
        1. Set BuiltInParameter.SCHEDULE_TEXT_SIZE on the schedule element.
           This is the parameter that Revit's sheet renderer actually reads.
        2. Call TableSectionData.SetGlobalCellStyle / equivalent on the header
           section so the TableCellStyle baseline also carries the right size.
        3. Try any schedule-level text parameters by enumerating the element's
           parameter set (locale-safe fallback for non-English Revit installs).

    Call this function:
        a) Right after create_generic_model_schedule()
        b) Right after write_header() completes
        c) Inside _apply_schedule_sheet_appearance() after sheet placement
    """
    if schedule is None or default_size_ft is None or default_size_ft <= 0:
        return

    t = Transaction(doc, 'Set Schedule Global Text Style')
    t.Start()
    try:
        # ── Approach 1: BuiltInParameter.SCHEDULE_TEXT_SIZE ──────────────────
        # This is the primary parameter that governs sheet-view text rendering.
        _bip_names = [
            'SCHEDULE_TEXT_SIZE',
            'SCHEDULE_FILTER_PARAM_TEXT_SIZE',
            'SCHEDULE_HEADER_TEXT_SIZE',
        ]
        for bip_name in _bip_names:
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
                if bip is None:
                    continue
                p = schedule.get_Parameter(bip)
                if p is not None:
                    # V23 FIX (Bug 1): Removed `not p.IsReadOnly` guard.
                    # SCHEDULE_TEXT_SIZE often reports IsReadOnly=True via
                    # parameter introspection, but the Revit API still accepts
                    # .Set() on it in most 2022-2026 builds.  The old guard
                    # silently skipped the primary fix.  We now always attempt
                    # .Set() and let Revit raise if it is truly immutable.
                    try:
                        p.Set(default_size_ft)
                    except Exception:
                        pass
            except Exception:
                pass

        # ── Approach 2: Enumerate ALL parameters for text/size keywords ───────
        # Covers non-English Revit and future BIP renames.
        try:
            for p in schedule.Parameters:
                try:
                    nm = (p.Definition.Name or '').lower()
                    if any(kw in nm for kw in ('text size', 'textsize', 'font size',
                                               'text height', 'schrift')):
                        # V23 FIX (Bug 1 follow-up): same IsReadOnly guard
                        # removed; attempt .Set() regardless and let Revit raise.
                        if p.StorageType.ToString() == 'Double':
                            try:
                                p.Set(default_size_ft)
                            except Exception:
                                pass
                except Exception:
                    pass
        except Exception:
            pass

        # ── Approach 3: REMOVED ───────────────────────────────────────────────
        # The per-cell sweep previously here called SetCellStyleOverrideOptions
        # with a font-only TableCellStyleOverrideOptions object, which REPLACED
        # (not merged) the existing override options on every cell.  This silently
        # wiped the border and fill override flags that write_header Phase E had
        # just set, causing all borders and fills to disappear on the sheet view.
        # Approach 1 (SCHEDULE_TEXT_SIZE parameter) is sufficient for driving the
        # schedule's sheet-view baseline text size; the per-cell sweep is not needed.

        # ── Approach 4 (V23): Set schedule "Text Appearance" / "Text Type" ─────
        # Some Revit builds drive sheet text from a TextNoteType-based appearance
        # parameter rather than SCHEDULE_TEXT_SIZE.
        try:
            tnt_id = _find_best_text_note_type(doc, default_font_name, default_size_ft)
            if tnt_id is not None:
                _set_schedule_text_appearance(doc, schedule, tnt_id)
        except Exception:
            pass

        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()

def _apply_schedule_sheet_appearance(doc, schedule, ssi, grid_model, border_styles):
    """
    After placing the schedule on a sheet (ScheduleSheetInstance),
    re-apply all cell-level overrides so they survive the sheet placement.

    Revit internally re-evaluates TableCellStyle when a ScheduleSheetInstance
    is created. This function re-stamps every cell's style a second time
    (within a new transaction) so the sheet view reflects:
      - Correct borders (per-cell GraphicsStyle line-weight)
      - Correct background fills
      - Correct font name/size/bold/italic
      - Suppressed borders on merged spanned cells

    This is the KEY fix for the 'schedule looks fine but sheet view strips styling' bug.
    """
    if ssi is None or schedule is None:
        return

    _default_size_ft = _pt_to_feet_exact(grid_model.default_font_size_pt)
    _default_font    = _resolve_font_name(grid_model.default_font_name)

    # V23: Re-apply schedule-level global text style FIRST.
    # Use the largest font size found in the grid as the schedule baseline.
    _baseline_size_ft = _compute_schedule_baseline_text_size_ft(grid_model)
    _set_schedule_global_text_style(doc, schedule, _baseline_size_ft, _default_font)

    t = Transaction(doc, 'Sheet-View Style Override')
    t.Start()
    try:
        hd        = _get_header_view(schedule)
        first_row = hd.FirstRowNumber
        first_col = hd.FirstColumnNumber
        max_r     = first_row + hd.NumberOfRows    - 1
        max_c     = first_col + hd.NumberOfColumns - 1

        for row_i in range(grid_model.rows):
            for col_i in range(grid_model.cols):
                abs_r = first_row + row_i
                abs_c = first_col + col_i
                if abs_r > max_r or abs_c > max_c:
                    continue

                cell = grid_model.get(row_i, col_i)

                if cell is not None and cell.is_master:
                    cd = cell.as_dict()
                    if _is_symbol_font(cell.font_name):
                        _apply_symbol_cell_style(hd, abs_r, abs_c, cd)
                        _restamp_borders_only(hd, abs_r, abs_c, cell, border_styles)
                    else:
                        if not cd.get('font_size'):
                            cd['font_size'] = _default_size_ft
                        if not cd.get('font_name'):
                            cd['font_name'] = _default_font
                        _apply_style(hd, abs_r, abs_c, cd, border_styles)
                else:
                    _apply_hidden_borders(hd, abs_r, abs_c, border_styles)

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        # Surface as a Revit dialog so the user knows sheet-view styling was partial.
        try:
            TaskDialog.Show('Sheet Style Warning',
                            'Sheet-view style pass incomplete: {0}'.format(str(ex)))
        except Exception:
            pass

def write_header(doc, view, grid_model, import_images=True, image_mode='incell',
                 line_style_map=None):
    if not isinstance(view, ViewSchedule):
        raise TypeError('View is not a ViewSchedule.')

    warnings = list(grid_model.warnings)
    imported_images = 0

    border_styles = BorderGraphicsStyles(doc, override_map=line_style_map)
    # Diagnostic: surface the effective Excel -> Revit mapping after user overrides merged
    try:
        _map_diag = ', '.join(
            '{} -> {} (id={})'.format(
                xl_s, rv_nm,
                getattr(border_styles.get(rv_nm), 'IntegerValue', '?'))
            for xl_s, rv_nm in sorted(border_styles._excel_to_revit.items())
        )
        warnings.append('[BORDER MAP] Effective excel->revit after override merge: ' + _map_diag)
    except Exception:
        pass
    # Check both bracket and plain forms since projects vary
    _diag_keys = ['<Hidden>', 'Hidden', '<Thin Lines>', 'Thin Lines',
                  '<Medium Lines>', 'Medium Lines', '<Wide Lines>', 'Wide Lines']
    available = [n for n in _diag_keys if border_styles.get(n) is not None]
    warnings.append('[BORDER DIAG] GraphicsStyles resolved: {0}'.format(
        ', '.join(available) if available else 'none -- borders will not apply'))

    # Count how many master cells have at least one explicit Excel border side.
    # If this is 0 the Excel file uses Table Styles or conditional formatting for
    # borders, which this tool does not parse — borders will always be invisible.
    _cells_with_borders = sum(
        1 for c in grid_model.cells
        if c.is_master and any([c.border_top, c.border_bottom,
                                c.border_left, c.border_right])
    )
    _total_masters = sum(1 for c in grid_model.cells if c.is_master)
    warnings.append(
        '[BORDER DIAG] Cells with explicit Excel border data: {0}/{1}{2}'.format(
            _cells_with_borders, _total_masters,
            ' -- WARNING: 0 means file uses Table Styles (not parsed)'
            if _cells_with_borders == 0 else ''))

    # List writable properties of TableCellStyleOverrideOptions via .NET reflection
    # so we can verify the border flag names are correct for this Revit build.
    try:
        import clr
        _tco_inst = TableCellStyleOverrideOptions()
        _tco_type = _tco_inst.GetType()  # .NET GetType(), not Python type()
        _tco_props = sorted(p.Name for p in _tco_type.GetProperties() if p.CanWrite)
        warnings.append('[TCO PROPS] ' + ', '.join(_tco_props))
    except Exception as _rex:
        warnings.append('[TCO PROPS] reflection failed: ' + str(_rex))

    # Diagnostic: list Revit project line style names (plain, no brackets) so user
    # can see exactly which styles are available and whether Thin/Medium/Wide Lines
    # resolved to distinct ElementIds or all fell back to the same style.
    try:
        _gs_plain = sorted(
            '{0}(id={1})'.format(k, border_styles._cache[k].IntegerValue)
            for k in border_styles._cache
            if not k.startswith('<') and border_styles._cache[k] is not None
        )
        warnings.append('[BORDER DIAG] Project line styles: {0}'.format(
            ', '.join(_gs_plain) if _gs_plain else 'none found'))
    except Exception:
        pass

    # Diagnostic: Excel border style distribution → Revit mapping
    try:
        _bsd = {}
        for _c in grid_model.cells:
            if _c.is_master:
                for _side in (_c.border_top, _c.border_bottom,
                              _c.border_left, _c.border_right):
                    if _side:
                        _bsd[_side] = _bsd.get(_side, 0) + 1
        if _bsd:
            _bsd_str = ', '.join(
                '{0}(x{1})->{2}(id={3})'.format(
                    k, v,
                    border_styles._excel_to_revit.get(k, 'Thin Lines'),
                    getattr(border_styles.get(
                        border_styles._excel_to_revit.get(k, 'Thin Lines')),
                        'IntegerValue', '?'))
                for k, v in sorted(_bsd.items())
            )
            warnings.append('[BORDER DIAG] Excel border style->Revit id: ' + _bsd_str)
        else:
            warnings.append('[BORDER DIAG] No per-cell Excel border styles found')
    except Exception:
        pass

    def _row_h(i):
        h = grid_model.row_heights_ft
        return h[i] if (h and i < len(h)) else None

    # Phase A0: clear to one row
    t = Transaction(doc, 'Clear Header Rows')
    t.Start()
    try:
        hd = _get_header_view(view)
        for i in range(hd.NumberOfRows - 1, 0, -1):
            hd.RemoveRow(i)
        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise Exception('Clear rows failed: {0}'.format(str(ex)))

    # Phase A1: insert rows
    for i in range(1, grid_model.rows):
        t = Transaction(doc, 'Insert Header Row {0}'.format(i))
        t.Start()
        try:
            hd = _get_header_view(view)
            hd.InsertRow(i - 1)
            t.Commit()
        except Exception as ex:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            warnings.append('Insert row {0}: {1}'.format(i, str(ex)))

    # Phase A2: row heights
    t = Transaction(doc, 'Header Row Heights')
    t.Start()
    try:
        hd        = _get_header_view(view)
        first_row = hd.FirstRowNumber
        for i in range(grid_model.rows):
            hi = _row_h(i)
            if hi is not None:
                try:
                    hd.SetRowHeight(first_row + i, hi)
                except Exception as ex:
                    warnings.append('Row {0} height: {1}'.format(i, str(ex)))
        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        warnings.append('Row heights tx: {0}'.format(str(ex)))

    rows_written = grid_model.rows

    # Phase B: column widths
    t = Transaction(doc, 'Header Column Widths')
    t.Start()
    try:
        total_width_ft = (sum(w for w in grid_model.col_widths_ft if w)
                          if grid_model.col_widths_ft else 0.0833)
        try:
            defn = view.Definition
            if defn.GetFieldCount() > 0:
                defn.GetField(0).ColumnWidth = total_width_ft
        except Exception:
            pass
        hd        = _get_header_view(view)
        first_col = hd.FirstColumnNumber
        for rel_c in range(min(hd.NumberOfColumns, len(grid_model.col_widths_ft))):
            try:
                hd.SetColumnWidth(first_col + rel_c, grid_model.col_widths_ft[rel_c])
            except Exception:
                pass
        try:
            bd = view.GetTableData().GetSectionData(SectionType.Body)
            bd.SetColumnWidth(bd.FirstColumnNumber, total_width_ft)
        except Exception:
            pass
        t.Commit()
        cols_written = grid_model.cols
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        warnings.append('Column widths: {0}'.format(str(ex)))
        cols_written = grid_model.cols

    # Phase C: cell text + images ONLY  (styles deliberately deferred to Phase E)
    # Styles must come AFTER merges — MergeCells() resets TableCellStyle on every
    # cell it touches, wiping any borders/fills applied before the merge.
    t = Transaction(doc, 'Header Content')
    t.Start()
    try:
        hd        = _get_header_view(view)
        first_row = hd.FirstRowNumber
        first_col = hd.FirstColumnNumber
        max_r     = first_row + hd.NumberOfRows    - 1
        max_c     = first_col + hd.NumberOfColumns - 1

        for row_i in range(grid_model.rows):
            for col_i in range(grid_model.cols):
                abs_r = first_row + row_i
                abs_c = first_col + col_i
                if abs_r > max_r or abs_c > max_c:
                    continue

                cell = grid_model.get(row_i, col_i)

                if cell is not None and cell.is_master:
                    # Handle image cells
                    if import_images and (image_mode == 'incell') and cell.is_image and cell.image_path:
                        img_id = _import_image_to_schedule_cell(
                            doc, view, abs_r, abs_c, cell.image_path)
                        if img_id is not None:
                            imported_images += 1
                        else:
                            warnings.append('Failed to import image at [{0},{1}]'.format(
                                row_i, col_i))

                    # Write text for non-image cells
                    if cell.value and not cell.is_image:
                        try:
                            hd.SetCellText(abs_r, abs_c, _sanitize_cell_text(str(cell.value)))
                        except Exception:
                            pass

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise Exception('Content write failed: {0}'.format(str(ex)))

    # Phase D: merges
    merges_applied = 0
    t = Transaction(doc, 'Header Merges')
    t.Start()
    try:
        hd        = _get_header_view(view)
        first_row = hd.FirstRowNumber
        first_col = hd.FirstColumnNumber
        max_r     = first_row + hd.NumberOfRows    - 1
        max_c     = first_col + hd.NumberOfColumns - 1

        for cell in grid_model.merged_masters():
            r_span, c_span = cell.merge_span
            if r_span == 1 and c_span == 1:
                continue
            abs_top    = first_row + cell.row
            abs_left   = first_col + cell.col
            abs_bottom = abs_top  + r_span - 1
            abs_right  = abs_left + c_span - 1
            if abs_bottom > max_r or abs_right > max_c:
                warnings.append('Merge [{0},{1}] out of bounds -- skipped.'.format(
                    cell.row, cell.col))
                continue
            try:
                mc        = TableMergedCell()
                mc.Top    = abs_top
                mc.Left   = abs_left
                mc.Bottom = abs_bottom
                mc.Right  = abs_right
                hd.MergeCells(mc)
                merges_applied += 1
            except Exception as e:
                warnings.append('Merge [{0},{1}]: {2}'.format(
                    cell.row, cell.col, str(e)))

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        warnings.append('Merges tx: {0}'.format(str(ex)))

    # Phase E: ALL cell styles — borders, fills, fonts — applied AFTER merges.
    # This is the correct order: MergeCells() resets TableCellStyle, so styles
    # must be written after all merge operations are committed.
    #
    # SHEET-VIEW FIX: cells without explicit Excel font_size / font_name get
    # the workbook defaults injected so that the FontSize & FontName override
    # flags are always set.  Without these flags Revit falls back to its own
    # default schedule text appearance (often 3/32" Calibri) when the schedule
    # is placed on a sheet, making text appear wrong even though the schedule
    # view itself looks correct.
    _default_size_ft = _pt_to_feet_exact(grid_model.default_font_size_pt)
    _default_font    = _resolve_font_name(grid_model.default_font_name)

    t = Transaction(doc, 'Header Styles')
    t.Start()
    try:
        hd        = _get_header_view(view)
        first_row = hd.FirstRowNumber
        first_col = hd.FirstColumnNumber
        max_r     = first_row + hd.NumberOfRows    - 1
        max_c     = first_col + hd.NumberOfColumns - 1

        for row_i in range(grid_model.rows):
            for col_i in range(grid_model.cols):
                abs_r = first_row + row_i
                abs_c = first_col + col_i
                if abs_r > max_r or abs_c > max_c:
                    continue

                cell = grid_model.get(row_i, col_i)

                if cell is not None and cell.is_master:
                    cd = cell.as_dict()
                    if _is_symbol_font(cell.font_name):
                        # Symbol cells (Wingdings/Webdings/Symbol): never set
                        # FontName/FontSize/WordWrap flags or Revit renders ✓ as □.
                        _apply_symbol_cell_style(hd, abs_r, abs_c, cd)
                        _restamp_borders_only(hd, abs_r, abs_c, cell, border_styles)
                    else:
                        # Inject workbook defaults so override flags are always set
                        if not cd.get('font_size'):
                            cd['font_size'] = _default_size_ft
                        if not cd.get('font_name'):
                            cd['font_name'] = _default_font
                        _apply_style(hd, abs_r, abs_c, cd, border_styles)
                else:
                    # Spanned cell inside a merge region — suppress its borders
                    _apply_hidden_borders(hd, abs_r, abs_c, border_styles)

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        warnings.append('Styles tx: {0}'.format(str(ex)))

    if imported_images > 0:
        warnings.append('Successfully imported {0} image(s).'.format(imported_images))

    # Phase F: Border-only second pass (matches V1.4)
    # Revit sometimes resets border flags on the first write when ShowGridLines
    # changes state.  A second isolated transaction that ONLY stamps border styles
    # ensures they persist in the schedule view and survive placement onto a sheet.
    # IMPORTANT: This is a border-only pass, NOT a full _apply_style pass.
    # Full _apply_style here would set SetCellStyleOverrideOptions again which
    # replaces the entire override object from Phase E, corrupting symbol glyphs.
    t = Transaction(doc, 'Header Borders Restamp')
    t.Start()
    try:
        hd        = _get_header_view(view)
        first_row = hd.FirstRowNumber
        first_col = hd.FirstColumnNumber
        max_r     = first_row + hd.NumberOfRows    - 1
        max_c     = first_col + hd.NumberOfColumns - 1

        for row_i in range(grid_model.rows):
            for col_i in range(grid_model.cols):
                abs_r = first_row + row_i
                abs_c = first_col + col_i
                if abs_r > max_r or abs_c > max_c:
                    continue
                cell = grid_model.get(row_i, col_i)
                if cell is not None and cell.is_master:
                    # Re-stamp border overrides only (matches V1.4 _restamp_borders)
                    _restamp_borders_only(hd, abs_r, abs_c, cell, border_styles)
                else:
                    _apply_hidden_borders(hd, abs_r, abs_c, border_styles)
        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        warnings.append('Border restamp tx: {0}'.format(str(ex)))

    # Phase G: Symbol cell cleanup — reset + re-apply only Bold + borders.
    # After all previous phases, Wingdings/Webdings cells may have accumulated
    # override flags (Font, FontSize, Fill etc.) that corrupt Unicode glyph
    # rendering (✓ becomes □).  The user confirmed that manually resetting
    # ALL overrides and re-applying ONLY Bold produces correct symbols.
    # This phase runs LAST to ensure it has the final word on these cells.
    t = Transaction(doc, 'Symbol Cell Cleanup')
    t.Start()
    try:
        hd        = _get_header_view(view)
        first_row = hd.FirstRowNumber
        first_col = hd.FirstColumnNumber
        max_r     = first_row + hd.NumberOfRows    - 1
        max_c     = first_col + hd.NumberOfColumns - 1

        for row_i in range(grid_model.rows):
            for col_i in range(grid_model.cols):
                abs_r = first_row + row_i
                abs_c = first_col + col_i
                if abs_r > max_r or abs_c > max_c:
                    continue
                cell = grid_model.get(row_i, col_i)
                if cell is None or not cell.is_master:
                    continue
                # Only target symbol font cells
                fn = cell.font_name or ''
                if not fn:
                    continue
                fn_low = fn.lower()
                if 'wingding' not in fn_low and 'webding' not in fn_low and 'symbol' not in fn_low:
                    continue

                # RESET: create completely fresh style + override options
                s = TableCellStyle()
                o = TableCellStyleOverrideOptions()

                # Re-apply Bold only if Excel had it
                if cell.font_bold:
                    try:
                        s.IsFontBold = True
                        _flag(o, 'Bold', 'FontBold')
                    except Exception:
                        pass

                # Re-apply borders from cell data
                if border_styles is not None:
                    bt  = cell.border_top
                    bb  = cell.border_bottom
                    bl  = cell.border_left
                    br_ = cell.border_right

                    _hide = not SHOW_EMPTY_GRIDLINES
                    t_id = border_styles.get_border_id(bt,  use_hidden_for_empty=_hide)
                    b_id = border_styles.get_border_id(bb,  use_hidden_for_empty=_hide)
                    l_id = border_styles.get_border_id(bl,  use_hidden_for_empty=_hide)
                    r_id = border_styles.get_border_id(br_, use_hidden_for_empty=_hide)

                    if t_id is not None:
                        s.BorderTopLineStyle = t_id
                        _flag(o, 'BorderTopLineStyle')

                    if b_id is not None:
                        s.BorderBottomLineStyle = b_id
                        _flag(o, 'BorderBottomLineStyle')

                    if l_id is not None:
                        s.BorderLeftLineStyle = l_id
                        _flag(o, 'BorderLeftLineStyle')

                    if r_id is not None:
                        s.BorderRightLineStyle = r_id
                        _flag(o, 'BorderRightLineStyle')

                try:
                    s.SetCellStyleOverrideOptions(o)
                    hd.SetCellStyle(abs_r, abs_c, s)
                except Exception:
                    pass

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        warnings.append('Symbol cell cleanup tx: {0}'.format(str(ex)))

    return {
        'rows_written':   rows_written,
        'cols_written':   cols_written,
        'merges_applied': merges_applied,
        'images_imported': imported_images,
        'warnings':       warnings,
    }


# ══════════════════════════════════════════════════════════════════════════════
# ── SECTION 5: UI + MAIN
# ══════════════════════════════════════════════════════════════════════════════

def _center_form(form):
    form.StartPosition = WFFormStartPosition.CenterScreen


def _pick_excel_file():
    dlg = WFOpenFileDialog()
    dlg.Title       = 'Step 1 of 3  -  Select Excel Workbook'
    dlg.Filter      = 'Excel Workbooks (*.xlsx)|*.xlsx'
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == WFDialogResultEnum.OK else None


def _make_list_form(title, instruction, items, legend=None, preselect=0, btn_label='Select'):
    W, H   = 440, 340
    PAD    = 12
    LIST_H = 190
    BW, BH = 90, 28

    form = WFForm()
    form.Text            = title
    form.ClientSize      = DSize(W, H)
    form.FormBorderStyle = WFFormBorderStyle.FixedDialog
    form.MaximizeBox     = False
    form.MinimizeBox     = False
    _center_form(form)

    lbl = WFLabel()
    lbl.Text     = instruction
    lbl.Location = DPoint(PAD, PAD)
    lbl.Size     = DSize(W - PAD * 2, 20)
    lbl.Font     = DFont('Segoe UI', 9)
    form.Controls.Add(lbl)

    lb = WFListBox()
    lb.Location      = DPoint(PAD, PAD + 26)
    lb.Size          = DSize(W - PAD * 2, LIST_H)
    lb.SelectionMode = WFSelectionMode.One
    lb.Font          = DFont('Segoe UI', 9)
    lb.BorderStyle   = WFBorderStyle.FixedSingle
    for item in items:
        lb.Items.Add(item)
    if items:
        lb.SelectedIndex = max(0, min(preselect, len(items) - 1))
    form.Controls.Add(lb)

    y_after = PAD + 26 + LIST_H + 6
    if legend:
        leg = WFLabel()
        leg.Text      = legend
        leg.Location  = DPoint(PAD, y_after)
        leg.Size      = DSize(W - PAD * 2, 16)
        leg.Font      = DFont('Segoe UI', 8)
        leg.ForeColor = DColor.FromArgb(100, 100, 100)
        form.Controls.Add(leg)
        btn_y = y_after + 22
    else:
        btn_y = y_after + 8

    btn_cancel = WFButton()
    btn_cancel.Text         = 'Cancel'
    btn_cancel.Size         = DSize(BW, BH)
    btn_cancel.Location     = DPoint(W - PAD - BW, btn_y)
    btn_cancel.DialogResult = WFDialogResultEnum.Cancel
    form.Controls.Add(btn_cancel)
    form.CancelButton = btn_cancel

    btn_ok = WFButton()
    btn_ok.Text         = btn_label
    btn_ok.Size         = DSize(BW, BH)
    btn_ok.Location     = DPoint(W - PAD - BW * 2 - 8, btn_y)
    btn_ok.DialogResult = WFDialogResultEnum.OK
    form.Controls.Add(btn_ok)
    form.AcceptButton = btn_ok

    return form, lb


def _clamp(val, lo, hi):
    return max(lo, min(hi, val))


def _suggest_schedule_scales(filepath, sheet_name=None):
    """Return gentle auto-suggested row/column/text scale factors.

    The goal is not to hard-force layout, but to start from workbook-aware values
    that reduce clipping for long formatted strings while still letting the user
    adjust them manually.
    """
    suggestion = {
        'row_scale': 1.00,
        'col_scale': 1.50,
        'text_scale': 1.00,
    }
    try:
        parsed = parse_excel(filepath, sheet_name=sheet_name, max_rows=0, strict_print_area=False)
        col_widths = parsed.get('col_widths', []) or []
        row_heights = parsed.get('row_heights', []) or []
        cells = parsed.get('cells', []) or []
        default_font = float(parsed.get('default_font_size', 11.0) or 11.0)

        max_overflow = 1.0
        row_need = 1.0
        long_count = 0
        text_count = 0

        for cell in cells:
            if not cell.get('is_master', True):
                continue
            value = cell.get('value', '')
            if value is None:
                continue
            try:
                txt = unicode(value)
            except Exception:
                txt = str(value)
            txt = txt.replace('\r\n', '\n').replace('\r', '\n').strip()
            if not txt:
                continue

            text_count += 1
            lines = txt.split('\n')
            line_count = max(1, len(lines))
            max_line_len = max(len(line) for line in lines) if lines else len(txt)

            c0 = int(cell.get('col', 0) or 0)
            r0 = int(cell.get('row', 0) or 0)
            merge_span = cell.get('merge_span') or (1, 1)
            c_span = max(1, int(merge_span[1] or 1))
            r_span = max(1, int(merge_span[0] or 1))

            width_chars = 0.0
            for c_idx in range(c0, min(c0 + c_span, len(col_widths))):
                width_chars += float(col_widths[c_idx] or 8.43)
            width_chars = max(6.0, width_chars)

            available_chars = width_chars * 0.92
            overflow = float(max_line_len) / available_chars
            if overflow > 1.05:
                long_count += 1
            max_overflow = max(max_overflow, overflow)

            row_height_pt = 0.0
            for r_idx in range(r0, min(r0 + r_span, len(row_heights))):
                row_height_pt += float(row_heights[r_idx] or 15.0)
            row_height_pt = max(12.0, row_height_pt)

            font_pt = float(cell.get('font_size') or default_font or 11.0)
            est_lines = line_count
            if cell.get('wrap_text') and overflow > 1.0:
                est_lines = max(est_lines, int(round(overflow + 0.35)))
            required_row_pt = max(15.0, (font_pt * 1.25 * est_lines) + 3.0)
            row_need = max(row_need, required_row_pt / row_height_pt)

        if text_count > 0:
            density = float(long_count) / float(text_count)
        else:
            density = 0.0

        suggestion['col_scale'] = round(_clamp(max(1.05, max_overflow * (1.02 if density < 0.25 else 1.08)), 0.90, 1.80), 2)
        suggestion['row_scale'] = round(_clamp(max(0.90, row_need * (1.02 if density < 0.30 else 1.08)), 0.75, 1.80), 2)

        text_scale = 1.00
        if default_font >= 12.5 and max_overflow > 1.15:
            text_scale = 0.95
        elif default_font <= 9.0:
            text_scale = 1.05
        suggestion['text_scale'] = round(_clamp(text_scale, 0.85, 1.20), 2)
    except Exception:
        pass
    return suggestion


def _pick_sheet(sheets):
    items     = []
    preselect = 0
    for i, sh in enumerate(sheets):
        marker = '  [Print Area defined]' if sh['has_print_area'] else ''
        items.append(sh['name'] + marker)
        if sh['has_print_area'] and preselect == 0:
            preselect = i

    form, lb = _make_list_form(
        title       = 'Step 2 of 3  -  Select Worksheet',
        instruction = 'Select the worksheet whose print area will be used as the schedule header:',
        items       = items,
        legend      = 'Sheets with a defined Print Area are recommended.',
        preselect   = preselect,
        btn_label   = 'Next',
    )
    if form.ShowDialog() == WFDialogResultEnum.OK and lb.SelectedIndex >= 0:
        return sheets[lb.SelectedIndex]['name']
    return None


# ── Document scanners (run on button click, not on dialog open) ───────────────

def _scan_excel_border_styles(filepath, sheet_name=None):
    """Unique Excel border style strings used in the sheet."""
    try:
        parsed = parse_excel(filepath, sheet_name=sheet_name, max_rows=0,
                             strict_print_area=False)
        styles = set()
        for cell in parsed.get('cells', []):
            for side in ('border_top', 'border_bottom', 'border_left', 'border_right'):
                v = cell.get(side)
                if v:
                    styles.add(v)
        return sorted(styles)
    except Exception:
        return []


def _scan_excel_fonts(filepath, sheet_name=None):
    """Unique (font_name, size_pt) tuples used in the sheet.

    Includes any cell with an explicit font_name OR font_size so that the
    Mapping Dialog presents every (name, size) combination the user can
    actually encounter — including cells that override only the font name
    while inheriting the default size.
    """
    try:
        parsed = parse_excel(filepath, sheet_name=sheet_name, max_rows=0,
                             strict_print_area=False)
        fonts = set()
        default_font = _resolve_font_name(
            parsed.get('default_font_name', 'Calibri') or 'Calibri')
        default_size = round(float(parsed.get('default_font_size', 11.0) or 11.0), 1)
        fonts.add((default_font, default_size))
        for cell in parsed.get('cells', []):
            raw_sz = cell.get('font_size')
            raw_nm = cell.get('font_name')
            has_sz = raw_sz and float(raw_sz) > 0
            has_nm = bool(raw_nm)
            if has_sz or has_nm:
                nm = _resolve_font_name(raw_nm) if has_nm else default_font
                sz = round(float(raw_sz), 1) if has_sz else default_size
                fonts.add((nm, sz))
        return sorted(fonts, key=lambda x: (x[0].lower(), x[1]))
    except Exception:
        return []


def _scan_excel_fill_colors(filepath, sheet_name=None):
    """Unique non-white background fill (r,g,b) tuples used in the sheet."""
    try:
        parsed = parse_excel(filepath, sheet_name=sheet_name, max_rows=0,
                             strict_print_area=False)
        colors = set()
        for cell in parsed.get('cells', []):
            rgb = cell.get('fill_rgb')
            if rgb:
                try:
                    r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
                    # Skip white / near-white
                    if r >= 250 and g >= 250 and b >= 250:
                        continue
                    colors.add((r, g, b))
                except Exception:
                    pass
        return sorted(colors)
    except Exception:
        return []


def scan_document(filepath, sheet_name=None):
    """Eagerly scan an imported document and return a summary dict:
        { 'fonts'  : [(font_name, size_pt), ...],
          'fills'  : [(r, g, b), ...],
          'borders': ['thin', 'medium', ...] }

    Used by the Import Options dialog so the user can map every style to
    Revit types or auto-create new ones.
    """
    if not filepath:
        return {'fonts': [], 'fills': [], 'borders': []}
    is_word = filepath.lower().endswith(('.doc', '.docx'))
    if is_word:
        # Word scanning not yet supported — return empty sets so dialogs
        # gracefully show "no styles detected".
        return {'fonts': [], 'fills': [], 'borders': []}
    return {
        'fonts'  : _scan_excel_fonts(filepath, sheet_name),
        'fills'  : _scan_excel_fill_colors(filepath, sheet_name),
        'borders': _scan_excel_border_styles(filepath, sheet_name),
    }


def _get_element_name(elem):
    try:
        from pyrevit import revit
        p = elem.get_Parameter(revit.DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p and p.AsString():
            return p.AsString()
    except Exception: pass
    try:
        from pyrevit import revit
        p = elem.get_Parameter(revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.AsString():
            return p.AsString()
    except Exception: pass
    try:
        nm = getattr(elem, "Name", None)
        if nm: return nm
    except Exception: pass
    try:
        from pyrevit import revit
        return revit.DB.Element.Name.GetValue(elem)
    except Exception: pass
    return None

def _get_project_filled_region_type_names(doc):
    """Sorted list of FilledRegionType names in the project."""
    names = []
    if doc is None:
        return names
    try:
        from pyrevit import revit
        DB = revit.DB
        seen = set()
        
        print("\n--- DEBUG: COLLECTING FILLED REGION TYPES ---")
        
        # 1. Standard collection using pyrevit.revit.DB
        frt_count = 0
        for frt in FilteredElementCollector(doc).OfClass(DB.FilledRegionType):
            try:
                frt_count += 1
                frt_name = _get_element_name(frt)
                if frt_name and frt.Id.IntegerValue not in seen:
                    names.append(frt_name)
                    seen.add(frt.Id.IntegerValue)
            except Exception as e:
                print("Error in #1: " + str(e))
        print("Method 1 (OfClass) found {} types. Unique so far: {}".format(frt_count, len(names)))

        # 2. Category-based fallback
        cat_count = 0
        for et in FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_FilledRegion).WhereElementIsElementType():
            try:
                cat_count += 1
                et_name = _get_element_name(et)
                if et_name and et.Id.IntegerValue not in seen:
                    names.append(et_name)
                    seen.add(et.Id.IntegerValue)
            except Exception as e:
                print("Error in #2: " + str(e))
        print("Method 2 (OfCategory) checked {} types. Unique so far: {}".format(cat_count, len(names)))

        # 3. Instance-based fallback (find types of existing instances in the project)
        inst_count = 0
        for fr in FilteredElementCollector(doc).OfClass(DB.FilledRegion):
            try:
                inst_count += 1
                frt = doc.GetElement(fr.GetTypeId())
                if frt:
                    frt_name = _get_element_name(frt)
                    if frt_name and frt.Id.IntegerValue not in seen:
                        names.append(frt_name)
                        seen.add(frt.Id.IntegerValue)
            except Exception as e:
                print("Error in #3: " + str(e))
        print("Method 3 (Instances) checked {} instances. Unique so far: {}".format(inst_count, len(names)))

        # 4. Fallback to string name check if STILL empty
        if not names:
            print("Trying fallback Method 4 (String Check on ALL Element Types)...")
            all_types = FilteredElementCollector(doc).WhereElementIsElementType()
            print("Total Element Types in project: {}".format(all_types.GetElementCount()))
            for et in all_types:
                try:
                    t_name = et.GetType().Name
                    if "FilledRegionType" in t_name:
                        et_name = _get_element_name(et)
                        print("Found class {} with Name: {}".format(t_name, et_name))
                        if et_name and et.Id.IntegerValue not in seen:
                            names.append(et_name)
                            seen.add(et.Id.IntegerValue)
                except Exception as e:
                    pass
        
        print("--- DEBUG: FINISHED COLLECTING. Total found: {} ---\n".format(len(names)))
    except Exception as e:
        import traceback
        print("CRITICAL ERROR collecting FilledRegionTypes: " + str(e))
        traceback.print_exc()

    return sorted(list(set(names)))


def _get_project_line_style_names(doc):
    """Sorted list of Revit project line style names usable for detail lines.

    Sourced from _collect_line_styles_dict — the same helper that populates
    BorderGraphicsStyles._cache.  This guarantees that any name shown in the
    dialog dropdown will resolve to a valid ElementId at apply-time.
    """
    if doc is None:
        return []
    styles = _collect_line_styles_dict(doc)
    if not styles:
        return []
    # _collect_line_styles_dict registers both '<Foo>' and 'Foo' for every
    # entry — keep the bracketed form when both exist (matches Revit's UI
    # convention for system styles like '<Sketch>' / '<Centerline>') and
    # collapse plain forms that have a bracketed sibling pointing to the
    # same ElementId so the dialog doesn't show duplicates.
    bracketed_ids = {
        v.IntegerValue: True
        for k, v in styles.items()
        if k.startswith('<') and k.endswith('>')
    }
    out = set()
    for name, eid in styles.items():
        if name.startswith('<') and name.endswith('>'):
            out.add(name)
        else:
            try:
                if bracketed_ids.get(eid.IntegerValue):
                    continue  # bracketed sibling already in the set
            except Exception:
                pass
            out.add(name)
    return sorted(out)


def _get_project_text_note_type_names(doc):
    """Sorted list of TextNoteType names in the project."""
    names = []
    if doc is None or not _HAS_TNT:
        return names
    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
        try:
            tnt_name = _get_element_name(tnt)
            if tnt_name:
                names.append(tnt_name)
        except Exception:
            pass
    return sorted(set(names))


# ── Mapping dialogs ───────────────────────────────────────────────────────────

class LineStyleMappingDialog(object):
    """
    Maps each Excel border style found in the loaded document to a Revit
    project line style. Only shown for Drafting View / Legend output types.

    After ShowDialog() == WFDialogResultEnum.OK, read .result:
        { 'thin': 'Thin Lines', 'medium': 'Medium Lines', ... }
    """
    _DEFAULTS = {
        'thin': 'Thin Lines',           'hair': 'Thin Lines',
        'medium': 'Medium Lines',       'mediumDashed': 'Medium Lines',
        'thick': 'Wide Lines',          'double': 'Wide Lines',
        'dashed': 'Thin Lines',         'dotted': 'Thin Lines',
        'dashDot': 'Thin Lines',        'mediumDashDot': 'Medium Lines',
        'dashDotDot': 'Thin Lines',     'mediumDashDotDot': 'Medium Lines',
        'slantDashDot': 'Thin Lines',
    }

    def __init__(self, xl_styles, revit_names, existing_map=None):
        self.result  = None
        self._xl     = xl_styles or []
        self._rv     = revit_names or []
        self._prior  = existing_map or {}
        self._h      = [None]
        self._form   = self._build()

    # Keyword hints used to find a project line style that matches the Excel
    # border style's visual character (dashed/dotted) before falling back to
    # the generic Thin/Medium/Wide defaults.
    _KEYWORD_HINTS = {
        'thin':             ['thin'],
        'hair':             ['hair', 'thin'],
        'medium':           ['medium'],
        'thick':            ['thick', 'wide'],
        'double':           ['double', 'wide'],
        'dashed':           ['dashed', 'dash'],
        'mediumDashed':     ['dashed', 'dash', 'medium'],
        'dotted':           ['dotted', 'dot'],
        'dashDot':          ['dash-dot', 'dashdot', 'dash dot', 'dash'],
        'mediumDashDot':    ['dash-dot', 'dashdot', 'dash dot', 'dash'],
        'dashDotDot':       ['dash-dot-dot', 'dashdotdot', 'dash dot dot', 'dash'],
        'mediumDashDotDot': ['dash-dot-dot', 'dashdotdot', 'dash dot dot', 'dash'],
        'slantDashDot':     ['dash-dot', 'dashdot', 'slant', 'dash'],
    }

    def _best_match(self, xl_style):
        # 1. Prefer a project line style whose name contains a keyword
        #    matching the Excel style's visual character (e.g. dashed → "Dash")
        for hint in self._KEYWORD_HINTS.get(xl_style, []):
            for n in self._rv:
                n_low = n.lower().strip('<>')
                if hint in n_low:
                    return n
        # 2. Fall back to the generic Thin/Medium/Wide default
        default = self._DEFAULTS.get(xl_style, 'Thin Lines')
        for candidate in (default, '<' + default + '>'):
            if candidate in self._rv:
                return candidate
        kw = default.lower().split()[0]
        for n in self._rv:
            if kw in n.lower():
                return n
        return self._rv[0] if self._rv else None

    def _build(self):
        import System.Windows.Forms as _WF
        W, PAD = 520, 12
        ROW_H  = 30
        BODY_H = min(max(len(self._xl), 1) * ROW_H + 8, 300)
        H      = 72 + BODY_H + 50

        form = WFForm()
        form.Text            = 'Line Style Mapping  –  Excel → Revit'
        form.ClientSize      = DSize(W, H)
        form.FormBorderStyle = WFFormBorderStyle.FixedDialog
        form.MaximizeBox     = False
        form.MinimizeBox     = False
        form.StartPosition   = WFFormStartPosition.CenterScreen

        hdr = WFLabel()
        hdr.Text     = 'Map each Excel border style to a Revit project line style:'
        hdr.Location = DPoint(PAD, PAD); hdr.Size = DSize(W - PAD*2, 18)
        hdr.Font     = DFont('Segoe UI', 9)
        form.Controls.Add(hdr)

        c1 = WFLabel(); c1.Text = 'Excel Border Style'
        c1.Location = DPoint(PAD, PAD+24); c1.Size = DSize(200, 16)
        c1.Font = DFont('Segoe UI', 8); form.Controls.Add(c1)

        c2 = WFLabel(); c2.Text = 'Revit Line Style'
        c2.Location = DPoint(PAD+210, PAD+24); c2.Size = DSize(W-PAD-220, 16)
        c2.Font = DFont('Segoe UI', 8); form.Controls.Add(c2)

        panel = _WF.Panel()
        panel.AutoScroll  = True
        panel.Location    = DPoint(PAD, 60)
        panel.Size        = DSize(W-PAD*2, BODY_H)
        panel.BorderStyle = _WF.BorderStyle.FixedSingle
        form.Controls.Add(panel)

        combos = []
        if not self._xl:
            msg = WFLabel()
            msg.Text = 'No explicit border styles detected in this sheet.'
            msg.Location = DPoint(8, 8); msg.Size = DSize(W-PAD*2-16, 20)
            msg.Font = DFont('Segoe UI Italic', 9)
            panel.Controls.Add(msg)
        else:
            y2 = 4
            for xl_s in self._xl:
                lx = WFLabel(); lx.Text = xl_s
                lx.Location = DPoint(6, y2+5); lx.Size = DSize(200, 20)
                lx.Font = DFont('Segoe UI', 9)
                panel.Controls.Add(lx)

                cmb = WFComboBox()
                cmb.DropDownStyle = WFComboBoxStyle.DropDownList
                cmb.Location = DPoint(210, y2+2)
                cmb.Size     = DSize(W-PAD*2-218, 24)
                cmb.Font     = DFont('Segoe UI', 9)
                for n in self._rv:
                    cmb.Items.Add(n)

                sel = self._prior.get(xl_s) or self._best_match(xl_s)
                if sel and sel in self._rv:
                    cmb.SelectedItem = sel
                elif self._rv:
                    cmb.SelectedIndex = 0

                panel.Controls.Add(cmb)
                combos.append((xl_s, cmb))
                y2 += ROW_H

        btn_y = H - 42
        btn_c = WFButton(); btn_c.Text = 'Cancel'; btn_c.Size = DSize(90, 28)
        btn_c.Location = DPoint(W-PAD-90, btn_y)
        btn_c.DialogResult = WFDialogResultEnum.Cancel
        form.Controls.Add(btn_c); form.CancelButton = btn_c

        btn_ok = WFButton(); btn_ok.Text = 'Apply'; btn_ok.Size = DSize(90, 28)
        btn_ok.Location = DPoint(W-PAD-190, btn_y)
        form.Controls.Add(btn_ok); form.AcceptButton = btn_ok

        _h = self._h
        def on_ok(s, e):
            _h[0] = {xs: (str(cb.SelectedItem) if cb.SelectedItem else None)
                     for xs, cb in combos}
            form.DialogResult = WFDialogResultEnum.OK
            form.Close()
        btn_ok.Click += on_ok
        return form

    def ShowDialog(self):
        dr = self._form.ShowDialog()
        if dr == WFDialogResultEnum.OK:
            self.result = self._h[0] or {}
        return dr


class TextNoteTypeMappingDialog(object):
    """
    Maps each (font_name, size_pt) combination found in the loaded document to a
    Revit TextNoteType. Only shown for Drafting View / Legend output types.

    After ShowDialog() == WFDialogResultEnum.OK, read .result:
        { 'Calibri_11.0': '2.5mm Arial', ... }
    """

    @staticmethod
    def make_key(font_name, size_pt):
        return '{}_{}'.format(font_name, round(float(size_pt), 1))

    def __init__(self, xl_fonts, revit_tnt_names, existing_map=None):
        self.result  = None
        self._fonts  = xl_fonts or []
        self._types  = revit_tnt_names or []
        self._prior  = existing_map or {}
        self._h      = [None]
        self._form   = self._build()

    def _build(self):
        import System.Windows.Forms as _WF
        W, PAD = 560, 12
        ROW_H  = 30
        BODY_H = min(max(len(self._fonts), 1) * ROW_H + 8, 300)
        H      = 72 + BODY_H + 50

        form = WFForm()
        form.Text            = 'Text Note Type Mapping  –  Excel → Revit'
        form.ClientSize      = DSize(W, H)
        form.FormBorderStyle = WFFormBorderStyle.FixedDialog
        form.MaximizeBox     = False
        form.MinimizeBox     = False
        form.StartPosition   = WFFormStartPosition.CenterScreen

        hdr = WFLabel()
        hdr.Text     = 'Map each Excel font to a Revit Text Note Type:'
        hdr.Location = DPoint(PAD, PAD); hdr.Size = DSize(W-PAD*2, 18)
        hdr.Font     = DFont('Segoe UI', 9)
        form.Controls.Add(hdr)

        c1 = WFLabel(); c1.Text = 'Excel Font  (name, pt)'
        c1.Location = DPoint(PAD, PAD+24); c1.Size = DSize(220, 16)
        c1.Font = DFont('Segoe UI', 8); form.Controls.Add(c1)

        c2 = WFLabel(); c2.Text = 'Revit Text Note Type'
        c2.Location = DPoint(PAD+230, PAD+24); c2.Size = DSize(W-PAD-240, 16)
        c2.Font = DFont('Segoe UI', 8); form.Controls.Add(c2)

        panel = _WF.Panel()
        panel.AutoScroll  = True
        panel.Location    = DPoint(PAD, 60)
        panel.Size        = DSize(W-PAD*2, BODY_H)
        panel.BorderStyle = _WF.BorderStyle.FixedSingle
        form.Controls.Add(panel)

        combos = []
        if not self._fonts:
            msg = WFLabel()
            msg.Text = 'No font data found in this sheet.'
            msg.Location = DPoint(8, 8); msg.Size = DSize(W-PAD*2-16, 20)
            msg.Font = DFont('Segoe UI Italic', 9)
            panel.Controls.Add(msg)
        else:
            y2 = 4
            for (fname, fsize) in self._fonts:
                key  = self.make_key(fname, fsize)
                disp = '{},  {:.1f} pt'.format(fname, fsize)

                lx = WFLabel(); lx.Text = disp
                lx.Location = DPoint(6, y2+5); lx.Size = DSize(220, 20)
                lx.Font = DFont('Segoe UI', 9)
                panel.Controls.Add(lx)

                cmb = WFComboBox()
                cmb.DropDownStyle = WFComboBoxStyle.DropDownList
                cmb.Location = DPoint(228, y2+2)
                cmb.Size     = DSize(W-PAD*2-236, 24)
                cmb.Font     = DFont('Segoe UI', 9)
                # Special "Create New" option -- generates DocLink_Text_<Font>_<Size>pt
                _create_label = u'« Create New: {}_{:.1f}pt »'.format(fname, fsize)
                cmb.Items.Add(_create_label)
                for n in self._types:
                    cmb.Items.Add(n)

                sel = self._prior.get(key)
                if sel == '__CREATE__':
                    cmb.SelectedIndex = 0
                elif sel and sel in self._types:
                    cmb.SelectedItem = sel
                else:
                    cmb.SelectedIndex = 0  # default to Create New

                panel.Controls.Add(cmb)
                combos.append((key, cmb, _create_label))
                y2 += ROW_H

        btn_y = H - 42
        btn_c = WFButton(); btn_c.Text = 'Cancel'; btn_c.Size = DSize(90, 28)
        btn_c.Location = DPoint(W-PAD-90, btn_y)
        btn_c.DialogResult = WFDialogResultEnum.Cancel
        form.Controls.Add(btn_c); form.CancelButton = btn_c

        btn_ok = WFButton(); btn_ok.Text = 'Apply'; btn_ok.Size = DSize(90, 28)
        btn_ok.Location = DPoint(W-PAD-190, btn_y)
        form.Controls.Add(btn_ok); form.AcceptButton = btn_ok

        _h = self._h
        def on_ok(s, e):
            out = {}
            for k, cb, create_lbl in combos:
                sel = str(cb.SelectedItem) if cb.SelectedItem else None
                if sel == create_lbl:
                    out[k] = '__CREATE__'
                else:
                    out[k] = sel
            _h[0] = out
            form.DialogResult = WFDialogResultEnum.OK
            form.Close()
        btn_ok.Click += on_ok
        return form

    def ShowDialog(self):
        dr = self._form.ShowDialog()
        if dr == WFDialogResultEnum.OK:
            self.result = self._h[0] or {}
        return dr


class FilledRegionMappingDialog(object):
    """
    Maps each unique RGB fill colour found in the loaded document to a
    Revit FilledRegionType, OR auto-creates a new one named
    DocLink_Fill_RRGGBB.

    After ShowDialog() == WFDialogResultEnum.OK, read .result:
        { (r, g, b): 'Existing Type' or '__CREATE__' or None }
    """

    @staticmethod
    def make_key(rgb):
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
        return '{:02X}{:02X}{:02X}'.format(r, g, b)

    def __init__(self, xl_fills, revit_frt_names, existing_map=None):
        self.result = None
        self._fills = xl_fills or []
        self._types = revit_frt_names or []
        self._prior = existing_map or {}
        self._h     = [None]
        self._form  = self._build()

    def _build(self):
        import System.Windows.Forms as _WF
        W, PAD = 540, 12
        ROW_H  = 30
        BODY_H = min(max(len(self._fills), 1) * ROW_H + 8, 300)
        H      = 72 + BODY_H + 50

        form = WFForm()
        form.Text            = 'Filled Region Mapping ({} types found)'.format(len(self._types))
        form.ClientSize      = DSize(W, H)
        form.FormBorderStyle = WFFormBorderStyle.FixedDialog
        form.MaximizeBox     = False
        form.MinimizeBox     = False
        form.StartPosition   = WFFormStartPosition.CenterScreen

        hdr = WFLabel()
        hdr.Text     = 'Map each Excel fill colour to a Revit Filled Region Type:'
        hdr.Location = DPoint(PAD, PAD); hdr.Size = DSize(W-PAD*2, 18)
        hdr.Font     = DFont('Segoe UI', 9)
        form.Controls.Add(hdr)

        c1 = WFLabel(); c1.Text = 'Excel Colour'
        c1.Location = DPoint(PAD, PAD+24); c1.Size = DSize(220, 16)
        c1.Font = DFont('Segoe UI', 8); form.Controls.Add(c1)

        c2 = WFLabel(); c2.Text = 'Revit Filled Region Type'
        c2.Location = DPoint(PAD+230, PAD+24); c2.Size = DSize(W-PAD-240, 16)
        c2.Font = DFont('Segoe UI', 8); form.Controls.Add(c2)

        panel = _WF.Panel()
        panel.AutoScroll  = True
        panel.Location    = DPoint(PAD, 60)
        panel.Size        = DSize(W-PAD*2, BODY_H)
        panel.BorderStyle = _WF.BorderStyle.FixedSingle
        form.Controls.Add(panel)

        combos = []
        if not self._fills:
            msg = WFLabel()
            msg.Text = 'No background fill colours found in this sheet.'
            msg.Location = DPoint(8, 8); msg.Size = DSize(W-PAD*2-16, 20)
            msg.Font = DFont('Segoe UI Italic', 9)
            panel.Controls.Add(msg)
        else:
            y2 = 4
            for rgb in self._fills:
                key = self.make_key(rgb)
                r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])

                # Colour swatch panel
                sw = _WF.Panel()
                sw.Location = DPoint(6, y2+4)
                sw.Size     = DSize(22, 22)
                sw.BackColor = DColor.FromArgb(r, g, b)
                sw.BorderStyle = _WF.BorderStyle.FixedSingle
                panel.Controls.Add(sw)

                lx = WFLabel(); lx.Text = '#{:02X}{:02X}{:02X}  ({},{},{})'.format(r, g, b, r, g, b)
                lx.Location = DPoint(34, y2+5); lx.Size = DSize(190, 20)
                lx.Font = DFont('Consolas', 9)
                panel.Controls.Add(lx)

                cmb = WFComboBox()
                cmb.DropDownStyle = WFComboBoxStyle.DropDownList
                cmb.Location = DPoint(228, y2+2)
                cmb.Size     = DSize(W-PAD*2-236, 24)
                cmb.Font     = DFont('Segoe UI', 9)
                _create_label = u'« Create New: DocLink_Fill_{:02X}{:02X}{:02X} »'.format(r, g, b)
                cmb.Items.Add(_create_label)
                for n in self._types:
                    cmb.Items.Add(n)

                sel = self._prior.get(key)
                if sel == '__CREATE__':
                    cmb.SelectedIndex = 0
                elif sel and sel in self._types:
                    cmb.SelectedItem = sel
                else:
                    cmb.SelectedIndex = 0  # default Create New

                panel.Controls.Add(cmb)
                combos.append((key, cmb, _create_label))
                y2 += ROW_H

        btn_y = H - 42
        btn_c = WFButton(); btn_c.Text = 'Cancel'; btn_c.Size = DSize(90, 28)
        btn_c.Location = DPoint(W-PAD-90, btn_y)
        btn_c.DialogResult = WFDialogResultEnum.Cancel
        form.Controls.Add(btn_c); form.CancelButton = btn_c

        btn_ok = WFButton(); btn_ok.Text = 'Apply'; btn_ok.Size = DSize(90, 28)
        btn_ok.Location = DPoint(W-PAD-190, btn_y)
        form.Controls.Add(btn_ok); form.AcceptButton = btn_ok

        _h = self._h
        def on_ok(s, e):
            out = {}
            for k, cb, create_lbl in combos:
                sel = str(cb.SelectedItem) if cb.SelectedItem else None
                if sel == create_lbl:
                    out[k] = '__CREATE__'
                else:
                    out[k] = sel
            _h[0] = out
            form.DialogResult = WFDialogResultEnum.OK
            form.Close()
        btn_ok.Click += on_ok
        return form

    def ShowDialog(self):
        dr = self._form.ShowDialog()
        if dr == WFDialogResultEnum.OK:
            self.result = self._h[0] or {}
        return dr




def _ehm_error(title, msg):
    td = TaskDialog(title)
    td.MainContent   = msg
    td.CommonButtons = TaskDialogCommonButtons.Ok
    td.Show()


def _ehm_delete_schedule(doc, schedule):
    t = None
    try:
        t = Transaction(doc, 'Delete Orphaned Schedule')
        t.Start()
        doc.Delete(schedule.Id)
        t.Commit()
    except Exception:
        try:
            if t is not None and t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════



# ─────────────────────────────────────────────────────────────────────────────
# ExcelHeaderMapper – entry point (called from Tab 2)
# ─────────────────────────────────────────────────────────────────────────────

def _extract_column_width_snapshot(doc, schedule):
    """
    Capture the width (in feet) of every column in the Revit schedule header.
    Returns a list of floats.
    """
    try:
        header = schedule.GetTableData().GetSectionData(SectionType.Header)
        count = header.NumberOfColumns
        widths = []
        for c in range(count):
            widths.append(header.GetColumnWidth(c))
        return widths
    except Exception:
        return None


def _apply_column_width_snapshot(doc, schedule, widths):
    """
    Re-apply saved column widths to a Revit schedule header and sync body.
    """
    if not widths:
        return
    with Transaction(doc, "Apply Retained Widths") as t:
        t.Start()
        try:
            hd = schedule.GetTableData().GetSectionData(SectionType.Header)
            count = min(len(widths), hd.NumberOfColumns)
            total_ft = 0.0
            for c in range(count):
                w = float(widths[c])
                try:
                    hd.SetColumnWidth(c, w)
                except Exception:
                    pass
                total_ft += w
            
            # Sync the Body and overall Field width to match the new total
            try:
                bd = schedule.GetTableData().GetSectionData(SectionType.Body)
                bd.SetColumnWidth(bd.FirstColumnNumber, total_ft)
                defn = schedule.Definition
                if defn.GetFieldCount() > 0:
                    defn.GetField(0).ColumnWidth = total_ft
            except Exception:
                pass
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()


def _get_schedule_text_size_ft(schedule):
    """Best-effort read of the schedule baseline text size in feet."""
    if schedule is None:
        return None

    for bip_name in (
        'SCHEDULE_TEXT_SIZE',
        'SCHEDULE_FILTER_PARAM_TEXT_SIZE',
        'SCHEDULE_HEADER_TEXT_SIZE',
    ):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = schedule.get_Parameter(bip)
            if p is not None:
                val = p.AsDouble()
                if val and val > 0:
                    return float(val)
        except Exception:
            pass

    try:
        for p in schedule.Parameters:
            try:
                if p is None or p.StorageType.ToString() != 'Double':
                    continue
                nm = (p.Definition.Name or '').lower()
                if any(kw in nm for kw in ('text size', 'textsize', 'font size',
                                           'text height', 'schrift')):
                    val = p.AsDouble()
                    if val and val > 0:
                        return float(val)
            except Exception:
                pass
    except Exception:
        pass

    return None


def _resize_schedule_in_place(doc, schedule, row_ratio=1.0, col_ratio=1.0,
                              text_ratio=1.0, resize_columns=True):
    """Resize an existing schedule view without rebuilding from Excel."""
    if doc is None or schedule is None:
        return False

    changed = False
    row_changed = abs(float(row_ratio or 1.0) - 1.0) > 0.001
    col_changed = resize_columns and abs(float(col_ratio or 1.0) - 1.0) > 0.001
    text_changed = abs(float(text_ratio or 1.0) - 1.0) > 0.001

    if row_changed or col_changed:
        with Transaction(doc, 'Resize Schedule In Place') as t:
            t.Start()
            try:
                table = schedule.GetTableData()
                hd = table.GetSectionData(SectionType.Header)

                if row_changed:
                    first_row = hd.FirstRowNumber
                    for i in range(hd.NumberOfRows):
                        row_idx = first_row + i
                        try:
                            cur_h = float(hd.GetRowHeight(row_idx))
                            if cur_h > 0:
                                new_h = cur_h * float(row_ratio)
                                hd.SetRowHeight(row_idx, new_h)
                                changed = True
                        except Exception:
                            pass

                if col_changed:
                    first_col = hd.FirstColumnNumber
                    total_ft = 0.0
                    for i in range(hd.NumberOfColumns):
                        col_idx = first_col + i
                        try:
                            cur_w = float(hd.GetColumnWidth(col_idx))
                            if cur_w > 0:
                                new_w = cur_w * float(col_ratio)
                                hd.SetColumnWidth(col_idx, new_w)
                                total_ft += new_w
                                changed = True
                        except Exception:
                            pass

                    if total_ft > 0:
                        try:
                            bd = table.GetSectionData(SectionType.Body)
                            bd.SetColumnWidth(bd.FirstColumnNumber, total_ft)
                        except Exception:
                            pass
                        try:
                            defn = schedule.Definition
                            if defn.GetFieldCount() > 0:
                                defn.GetField(0).ColumnWidth = total_ft
                        except Exception:
                            pass

                t.Commit()
            except Exception:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                raise

    if text_changed:
        cur_text_ft = _get_schedule_text_size_ft(schedule)
        if cur_text_ft and cur_text_ft > 0:
            _set_schedule_global_text_style(
                doc,
                schedule,
                cur_text_ft * float(text_ratio),
                None,
            )
            changed = True

    return changed


# ── Phase 7: capture merge layout from a completed GridModel ─────────────────

def _extract_merge_snapshot(grid_model):
    """
    Capture the current merge layout from a completed GridModel so it can
    be replayed on a future rebuild when retain_previous_merge_layout is ON.

    Returns a dict:
        { "rows": int, "cols": int,
          "merges": [ {"row":r, "col":c, "r_span":rs, "c_span":cs}, ... ] }
    """
    merges = []
    for cell in grid_model.merged_masters():
        rs, cs = cell.merge_span
        if rs > 1 or cs > 1:
            merges.append({
                "row": cell.row, "col": cell.col,
                "r_span": rs, "c_span": cs,
            })
    return {"rows": grid_model.rows, "cols": grid_model.cols, "merges": merges}


def _apply_merge_snapshot(grid_model, snapshot, warnings):
    """
    Replay a previously saved merge layout onto grid_model.

    Structure compatibility check:
        - same row count
        - same col count
    If the structure doesn't match, logs a warning and returns False so the
    caller can fall back to normal dynamic-merge behaviour.

    Returns True if the snapshot was applied, False on structure mismatch.
    """
    if snapshot is None:
        return False
    if snapshot.get("rows") != grid_model.rows or snapshot.get("cols") != grid_model.cols:
        warnings.append(
            "[MergeRetain] Snapshot structure mismatch "
            "(snapshot {}x{} vs current {}x{}) — using dynamic merge instead.".format(
                snapshot.get("rows"), snapshot.get("cols"),
                grid_model.rows, grid_model.cols))
        return False

    grid = {(c.row, c.col): c for c in grid_model.cells}

    # Reset all existing merge state so we start clean
    for cell in grid_model.cells:
        cell.is_master  = True
        cell.is_spanned = False
        cell.merge_span = (1, 1)

    applied = 0
    for entry in snapshot.get("merges", []):
        r, c = entry.get("row", -1), entry.get("col", -1)
        rs, cs = entry.get("r_span", 1), entry.get("c_span", 1)
        if r < 0 or c < 0 or rs < 1 or cs < 1:
            continue
        # Validate bounds
        if r + rs > grid_model.rows or c + cs > grid_model.cols:
            warnings.append(
                "[MergeRetain] Merge [{},{}] span {}x{} out of bounds — skipped.".format(
                    r, c, rs, cs))
            continue
        master = grid.get((r, c))
        if master is None:
            continue
        master.merge_span = (rs, cs)
        # Mark spanned cells
        for dr in range(rs):
            for dc in range(cs):
                if dr == 0 and dc == 0:
                    continue
                spanned = grid.get((r + dr, c + dc))
                if spanned is not None:
                    spanned.is_master  = False
                    spanned.is_spanned = True
        applied += 1

    warnings.append("[MergeRetain] Replayed {} merge(s) from snapshot.".format(applied))
    return True


def _extract_border_snapshot(grid_model):
    """
    Capture the per-cell border style strings from a completed GridModel so
    they can be replayed on a future rebuild when retain layout is ON.

    Only master cells are stored (spanned cells inherit borders from their master
    or have them suppressed).

    Returns a dict:
        { "rows": int, "cols": int,
          "borders": [ {"row":r, "col":c,
                        "top":str|None, "bottom":str|None,
                        "left":str|None, "right":str|None}, ... ] }
    """
    borders = []
    for cell in grid_model.cells:
        if not cell.is_master:
            continue
        # Only store cells that have at least one border set
        if not any([cell.border_top, cell.border_bottom,
                    cell.border_left, cell.border_right]):
            continue
        borders.append({
            "row":    cell.row,
            "col":    cell.col,
            "top":    cell.border_top,
            "bottom": cell.border_bottom,
            "left":   cell.border_left,
            "right":  cell.border_right,
        })
    return {
        "rows":    grid_model.rows,
        "cols":    grid_model.cols,
        "borders": borders,
    }


def _apply_border_snapshot(grid_model, snapshot, warnings):
    """
    Replay a previously saved border layout onto grid_model.

    Structure compatibility check:
        - same row count AND same col count as when snapshot was taken
    If the structure doesn't match, logs a warning and returns False so the
    caller can fall back to the freshly parsed border data.

    Returns True if the snapshot was applied, False on structure mismatch.
    """
    if snapshot is None:
        return False
    if snapshot.get("rows") != grid_model.rows or snapshot.get("cols") != grid_model.cols:
        warnings.append(
            "[BorderRetain] Snapshot structure mismatch "
            "(snapshot {}×{} vs current {}×{}) — using parsed borders instead.".format(
                snapshot.get("rows"), snapshot.get("cols"),
                grid_model.rows, grid_model.cols))
        return False

    grid = {(c.row, c.col): c for c in grid_model.cells}

    # First clear all border data so we start from a clean slate
    for cell in grid_model.cells:
        cell.border_top    = None
        cell.border_bottom = None
        cell.border_left   = None
        cell.border_right  = None

    applied = 0
    for entry in snapshot.get("borders", []):
        r, c = entry.get("row", -1), entry.get("col", -1)
        if r < 0 or c < 0:
            continue
        cell = grid.get((r, c))
        if cell is None:
            continue
        cell.border_top    = entry.get("top")
        cell.border_bottom = entry.get("bottom")
        cell.border_left   = entry.get("left")
        cell.border_right  = entry.get("right")
        applied += 1

    warnings.append("[BorderRetain] Replayed borders for {} cell(s) from snapshot.".format(
        applied))
    return True


# ── Phase 4+5: sched record normalisation + non-interactive rebuild ──────────

def _normalize_sched_records(records):
    """
    Ensure all schedule record dicts carry the fields added in Phases 4–7.
    Safe to call on old records — missing keys receive neutral defaults.
    """
    if not isinstance(records, list):
        return []
    for r in records:
        if not isinstance(r, dict):
            continue
        # Phase 5 — placement fields
        if "target_sheet_unique_id" not in r:
            r["target_sheet_unique_id"] = None
        if "last_placement_point" not in r:
            r["last_placement_point"] = None
        # Phase 7 — merge retention
        opts = r.get("options") if isinstance(r.get("options"), dict) else {}
        if "retain_previous_merge_layout" not in opts:
            opts["retain_previous_merge_layout"] = True  # V25: Default to True
        r["options"] = opts
        
        if "retain_settings" not in r:
            r["retain_settings"] = True  # Simplified top-level flag
            
        if "merge_snapshot" not in r:
            r["merge_snapshot"] = None
        # Border retention — stored alongside merge snapshot under same flag
        if "border_snapshot" not in r:
            r["border_snapshot"] = None
        # V25 Column Width retention
        if "width_snapshot" not in r:
            r["width_snapshot"] = None
        # Phase 4 — options_version marker
        if "options_version" not in r:
            r["options_version"] = 1
        # Path type — default to Absolute for old records
        if "path_type" not in r:
            r["path_type"] = "Absolute"
    return records


def _parse_document_to_grid(doc, filepath, selected_sheet, options, temp_dir):
    """
    Common parser for Excel and Word into a GridModel.
    """
    is_word = filepath.lower().endswith(('.doc', '.docx'))
    
    if is_word:
        from word_text_parser import parse_word
        parsed = parse_word(filepath)
        cell_images = {}
    else:
        raw_images = []
        if temp_dir:
            try:
                raw_images = _extract_images_from_excel(filepath, selected_sheet, temp_dir)
            except Exception: pass
            
        # Range source: named_range > manual_range > Print_Area > used range.
        # parse_excel handles the priority; this side just hands all three in.
        nm = options.get("named_range") or None
        mr = options.get("manual_range") or None
        parsed = parse_excel(
            filepath,
            sheet_name = selected_sheet or "Sheet1",
            max_rows = options.get("max_rows", 0),
            strict_print_area = (not (nm or mr)),
            named_range = nm,
            manual_range = mr,
        )
        
        cell_images = {}
        if raw_images:
            try:
                bounds = parsed.get('bounds')
                if bounds: cell_images = _map_images_to_cells(raw_images, bounds)
            except Exception: pass
            
    grid = build_grid(parsed, cell_images=cell_images, options=options)
    return grid

def _build_schedule_from_record(doc, uidoc, record,
                                 placement_pt=None,
                                 target_sheet=None):
    """
    Unified rebuild path for existing records.
    """
    filepath      = record.get("source_path", "")
    if record.get("path_type") == "Relative" and doc.PathName:
        filepath = os.path.join(os.path.dirname(doc.PathName), filepath)
    selected_sheet = record.get("sheet_name") or None
    options        = record.get("options", {}) or {}
    schedule_name  = record.get("schedule_name", "Rebuilt Document")
    output_type    = record.get("output_type", "Schedule")

    if not filepath or not os.path.exists(filepath):
        raise RuntimeError("Source file not found: {}".format(filepath))

    globals()['ROW_HEIGHT_SCALE']      = options.get("row_scale",    1.0)
    globals()['COL_WIDTH_SCALE']       = options.get("col_scale",    1.0)
    globals()['TEXT_SIZE_SCALE']       = options.get("text_scale",   1.0)
    globals()['ENABLE_DYNAMIC_MERGE']  = options.get("enable_dynamic_merge", True)
    globals()['SHOW_EMPTY_GRIDLINES']  = options.get("show_empty_gridlines", False)

    retain_logic   = _as_bool(record.get("retain_settings", True))
    
    merge_snapshot  = record.get("merge_snapshot")
    border_snapshot = record.get("border_snapshot")
    width_snapshot  = record.get("width_snapshot")

    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix='doclink_rebuild_')
        grid = _parse_document_to_grid(doc, filepath, selected_sheet, options, temp_dir)
        
        if output_type != "Schedule":
            from view_generator import create_graphic_view
            ext_id = record.get("schedule_unique_id")
            v_id = ElementId(int(ext_id)) if (ext_id and str(ext_id).isdigit()) else None
            # If not digit, might be UniqueId (string)
            if not v_id and ext_id:
                try:
                    v_el = doc.GetElement(str(ext_id))
                    if v_el: v_id = v_el.Id
                except Exception: pass

            view = create_graphic_view(doc, grid, options, existing_view_id=v_id)
            if not view: raise RuntimeError("Graphic view build failed.")
            
            return {
                "schedule_name":      schedule_name,
                "source_path":        record.get("source_path"),
                "sheet_name":         selected_sheet,
                "options":            options,
                "schedule_unique_id": view.UniqueId,
                "ssi_unique_id":      None,
                "merge_snapshot":     _extract_merge_snapshot(grid),
                "border_snapshot":    _extract_border_snapshot(grid),
                "width_snapshot":     None,
                "output_type":        output_type
            }

        # Schedule Path
        # ... logic for schedule build ...

        # Phase 6: replay merge snapshot if enabled and compatible
        if retain_logic and merge_snapshot:
            applied = _apply_merge_snapshot(grid, merge_snapshot, grid.warnings)
            if not applied:
                LogManager.debug("[DocLinkManager] MergeRetain: structure mismatch, "
                      "re-running dynamic merge as fallback.")
                # Structure changed — run dynamic merge now as fallback
                from copy import deepcopy as _dc
                try:
                    _dynamic_overflow_merge(
                        grid.cells, grid.rows, grid.cols,
                        grid.col_widths_ft, grid.default_font_size_pt)
                    _unified_border_pass(grid.cells, grid.rows, grid.cols)
                except Exception as _ex:
                    LogManager.debug("[DocLinkManager] MergeRetain fallback merge error: "
                          "{}".format(_ex))

        # Replay border snapshot under the same retain flag
        if retain_logic and border_snapshot:
            b_applied = _apply_border_snapshot(grid, border_snapshot, grid.warnings)
            if not b_applied:
                LogManager.debug("[DocLinkManager] BorderRetain: structure mismatch — "
                      "using freshly parsed borders.")

        try:
            new_schedule = create_generic_model_schedule(
                doc           = doc,
                schedule_name = schedule_name,
                num_cols      = grid.cols,
                col_widths_ft = grid.col_widths_ft,
            )
        except Exception as e:
            raise RuntimeError("Schedule creation failed: {}".format(str(e)))

        _baseline_size_ft = _compute_schedule_baseline_text_size_ft(grid, options)
        _baseline_font    = _resolve_font_name(grid.default_font_name)
        _set_schedule_global_text_style(
            doc, new_schedule, _baseline_size_ft, _baseline_font)

        try:
            write_header(doc, new_schedule, grid,
                         import_images=True, image_mode='incell',
                         line_style_map=options.get('line_style_map'))

            # Phase 25: Column Width retention
            if retain_logic and width_snapshot:
                _apply_column_width_snapshot(doc, new_schedule, width_snapshot)
                
        except Exception as e:
            _ehm_delete_schedule(doc, new_schedule)
            raise RuntimeError("Header write failed: {}".format(str(e)))

        _set_schedule_global_text_style(
            doc, new_schedule, _baseline_size_ft, _baseline_font)

        ssi = None
        border_styles_main = BorderGraphicsStyles(
            doc, override_map=options.get('line_style_map'))

        # Determine placement sheet
        sheet = target_sheet
        if sheet is None and uidoc is not None:
            sheet = _get_active_sheet(uidoc)

        if sheet is not None:
            try:
                # Phase 5: use the caller-supplied explicit point if present,
                # otherwise place at sheet origin (first-placement behaviour).
                ssi = _place_schedule_on_sheet(
                    doc, uidoc, sheet, new_schedule, grid,
                    explicit_pt=placement_pt)
                if ssi is not None:
                    _apply_schedule_sheet_appearance(
                        doc, new_schedule, ssi, grid, border_styles_main)


            except Exception as ex:
                LogManager.debug("[DocLinkManager] _build_schedule_from_record: "
                      "sheet placement failed: {}".format(ex))

        try:
            if ssi is not None and sheet is not None:
                if uidoc is not None:
                    uidoc.ActiveView = sheet
            else:
                if uidoc is not None:
                    uidoc.ActiveView = new_schedule
        except Exception:
            pass

        # Capture new merge + border + width snapshots after build
        new_merge_snapshot  = _extract_merge_snapshot(grid)
        new_border_snapshot = _extract_border_snapshot(grid)
        new_width_snapshot  = _extract_column_width_snapshot(doc, new_schedule)

        return {
            "schedule_name":      schedule_name,
            "source_path":        filepath,
            "sheet_name":         selected_sheet,
            "options":            options,
            "schedule_unique_id": new_schedule.UniqueId if new_schedule else None,
            "ssi_unique_id":      ssi.UniqueId if ssi else None,
            "merge_snapshot":     new_merge_snapshot,
            "border_snapshot":    new_border_snapshot,
            "width_snapshot":     new_width_snapshot,
        }

    finally:
        if temp_dir and os.path.exists(temp_dir):
            try:
                import shutil
                shutil.rmtree(temp_dir)
            except Exception:
                pass


def run_document_importer(doc, uidoc, existing_record=None):
    """
    Unified entry point for doc link imports.
    Supports Excel/Word source -> Schedule/Drafting/Legend output.
    If existing_record is provided, it re-imports using saved settings.
    """
    from pyrevit import forms
    
    if existing_record:
        filepath       = existing_record.get('source_path')
        selected_sheet = existing_record.get('sheet_name')
        is_word        = bool(filepath and filepath.lower().endswith(('.doc', '.docx')))
        options        = existing_record.get('options')
        if not options: options = {}
        # Back-fill keys that old persisted records may not have
        options.setdefault('name',                   existing_record.get('schedule_name'))
        options.setdefault('output_type',            existing_record.get('output_type', 'Schedule'))
        options.setdefault('row_scale',              1.0)
        options.setdefault('col_scale',              1.5)
        options.setdefault('text_scale',             1.0)
        options.setdefault('view_scale',             1)
        options.setdefault('max_rows',               0)
        options.setdefault('enable_dynamic_merge',   True)
        options.setdefault('show_empty_gridlines',   False)
        options.setdefault('scale_by_view',          True)
        options.setdefault('line_style_map',         {})
        options.setdefault('text_note_map',          {})
        options.setdefault('single_text_note_type',  None)
        options.setdefault('fill_region_map',        {})
        options.setdefault('default_text_note_type', None)
        options.setdefault('excel_points_per_mm',    2.834)
    else:
        filepath = forms.pick_file(
            files_filter="Support Files (*.xlsx, *.xls, *.xlsm, *.docx, *.doc)|*.xlsx;*.xls;*.xlsm;*.docx;*.doc",
            title="Select Document for Link"
        )
        if filepath is None:
            return None

        is_word = filepath.lower().endswith(('.doc', '.docx'))
        selected_sheet = None
        
        if is_word:
            from word_text_parser import get_word_info
            try:
                info = get_word_info(filepath)
                pages = info.get('pages', 1)
                selected_sheet = None # Placeholder for Word "page"
                
                if pages > 1:
                    page_items = ["All Pages"] + ["Page {}".format(p) for p in range(1, pages + 1)]
                    sel = forms.SelectFromList.show(page_items, title="Select Word Page", multiselect=False)
                    if sel:
                        if sel == "All Pages":
                            selected_sheet = None
                        else:
                            selected_sheet = int(sel.split()[-1])
                    else:
                        return None
            except Exception as e:
                _ehm_error('Could Not Get Word Info', str(e))
                return None
            base_name = os.path.splitext(os.path.basename(filepath))[0]
        else:
            try:
                sheets = list_sheets(filepath)
            except Exception as e:
                _ehm_error('Could Not Read Workbook', str(e))
                return None

            if not sheets:
                _ehm_error('No Sheets', 'The workbook contains no sheets.')
                return None

            selected_sheet = sheets[0]['name'] if len(sheets) == 1 else _pick_sheet(sheets)
            if selected_sheet is None:
                return None
            base_name = os.path.splitext(os.path.basename(filepath))[0]

        try:
            existing_names = set(v.Name for v in FilteredElementCollector(doc).OfClass(View))
        except Exception:
            existing_names = set()

        # Eager document scan — collects fonts, fills, borders before showing options
        try:
            scan_results = scan_document(filepath, sheet_name=selected_sheet)
        except Exception:
            scan_results = {'fonts': [], 'fills': [], 'borders': []}

        try:
            dlg = ScheduleSetupDialog(default_name=base_name, existing_names=existing_names,
                                      filepath=filepath, sheet_name=selected_sheet,
                                      doc=doc, scan_results=scan_results)
            if dlg.ShowDialog() == WFDialogResultEnum.OK:
                options = dlg.result
            else:
                options = None
        except Exception as _e:
            _ehm_error('Options Dialog Error', '{}: {}'.format(type(_e).__name__, _e))
            return None
        if options is None:
            return None

    schedule_name = options.get('name') or ''
    output_type   = options.get('output_type', 'Schedule')

    globals()['ROW_HEIGHT_SCALE']     = options.get('row_scale',            1.0)
    globals()['COL_WIDTH_SCALE']      = options.get('col_scale',            1.5)
    globals()['TEXT_SIZE_SCALE']      = options.get('text_scale',           1.0)
    globals()['ENABLE_DYNAMIC_MERGE'] = options.get('enable_dynamic_merge', True)
    globals()['SHOW_EMPTY_GRIDLINES'] = options.get('show_empty_gridlines', False)

    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix='doclink_importer_')
    except Exception:
        temp_dir = None

    try:
        if is_word:
            from word_text_parser import parse_word
            parsed = parse_word(filepath, selected_page=selected_sheet)
            cell_images = {} # Image extraction for Word tables not yet implemented
        else:
            raw_images = []
            if temp_dir:
                try:
                    raw_images = _extract_images_from_excel(filepath, selected_sheet, temp_dir)
                except Exception: pass

            try:
                # Named-Range / Manual-Range overrides strict_print_area.
                nm = options.get('named_range')  or None
                mr = options.get('manual_range') or None
                # The dialog may have snapped the sheet combo to the name's
                # resolved sheet — prefer the post-dialog sheet over the
                # pre-dialog `selected_sheet`.
                effective_sheet = options.get('sheet_name') or selected_sheet
                parsed = parse_excel(
                    filepath,
                    sheet_name = effective_sheet,
                    max_rows = options.get('max_rows', 0),
                    strict_print_area = (not (nm or mr)),
                    named_range = nm,
                    manual_range = mr,
                )
            except Exception as e:
                _ehm_error('Parse Error', str(e))
                return None

            cell_images = {}
            if raw_images:
                try:
                    bounds = parsed.get('bounds')
                    if bounds: cell_images = _map_images_to_cells(raw_images, bounds)
                except Exception: pass

        try:
            grid = build_grid(parsed, cell_images=cell_images)
        except Exception as e:
            _ehm_error('Grid Build Failed', str(e))
            return None

        # ── Route to Output Generator ─────────────────────────────────────────
        if output_type == "Schedule":
            LogManager.info("Generating Revit Schedule View...")
            return _execute_schedule_path(doc, uidoc, grid, options, filepath, selected_sheet)
        else:
            LogManager.info("Generating Graphic View: {} ({})".format(schedule_name, output_type))
            from view_generator import create_graphic_view, LegendTemplateMissingError

            # Check for existing view if re-import
            v_id = None
            if existing_record:
                ext_id = existing_record.get("schedule_unique_id")
                if ext_id:
                    try:
                        v_el = doc.GetElement(str(ext_id))
                        if v_el: v_id = v_el.Id
                    except Exception: pass

            try:
                view = create_graphic_view(doc, grid, options, existing_view_id=v_id)
            except LegendTemplateMissingError as _lme:
                _ehm_error('Cannot Create Legend', str(_lme))
                return None
            except Exception as _vge:
                _ehm_error('Graphic View Error', '{}: {}'.format(type(_vge).__name__, _vge))
                return None
            if not view: return None
            
            # Switch to view
            try:
                uidoc.ActiveView = view
            except Exception: pass
            
            # Return persistence record
            return {
                "schedule_name":      schedule_name,
                "source_path":        filepath,
                "sheet_name":         selected_sheet,
                "options":            options,
                "schedule_unique_id": view.UniqueId,
                "ssi_unique_id":      None,
                "merge_snapshot":     _extract_merge_snapshot(grid),
                "border_snapshot":    _extract_border_snapshot(grid),
                "width_snapshot":     None,
                "output_type":        output_type
            }

    finally:
        if temp_dir and os.path.exists(temp_dir):
            try:
                import shutil
                shutil.rmtree(temp_dir)
            except Exception:
                pass


def _execute_schedule_path(doc, uidoc, grid, options, filepath, selected_sheet):
    schedule_name = options['name']
    try:
        # Wrap all Revit transactions in a TransactionGroup so the
        # entire schedule import appears as one undo step.
        tg = TransactionGroup(doc, "Schedule Import")
        tg.Start()
        try:
            try:
                new_schedule = create_generic_model_schedule(
                    doc           = doc,
                    schedule_name = schedule_name,
                    num_cols      = grid.cols,
                    col_widths_ft = grid.col_widths_ft,
                )
            except Exception as e:
                tg.RollBack()
                _ehm_error('Schedule Creation Failed', str(e))
                return None

            _baseline_size_ft = _compute_schedule_baseline_text_size_ft(grid, options)
            _baseline_font    = _resolve_font_name(grid.default_font_name)
            _set_schedule_global_text_style(
                doc, new_schedule, _baseline_size_ft, _baseline_font)

            try:
                result = write_header(doc, new_schedule, grid,
                                      import_images=True, image_mode='incell',
                                      line_style_map=options.get('line_style_map'))
            except Exception as e:
                _ehm_delete_schedule(doc, new_schedule)
                tg.RollBack()
                _ehm_error('Header Write Failed', str(e))
                return None

            _set_schedule_global_text_style(
                doc, new_schedule, _baseline_size_ft, _baseline_font)

            ssi = None
            border_styles_main = BorderGraphicsStyles(
                doc, override_map=options.get('line_style_map'))

            active_sheet = _get_active_sheet(uidoc)
            if active_sheet is not None:
                try:
                    # explicit_pt=None → centre on sheet (first-time creation)
                    ssi = _place_schedule_on_sheet(
                        doc, uidoc, active_sheet, new_schedule, grid,
                        explicit_pt=None)
                    if ssi is not None:
                        _apply_schedule_sheet_appearance(
                            doc, new_schedule, ssi, grid, border_styles_main)
                except Exception:
                    pass

            try:
                if ssi is not None and active_sheet is not None:
                    uidoc.ActiveView = active_sheet
                else:
                    uidoc.ActiveView = new_schedule
            except Exception:
                pass

            tg.Assimilate()   # merge all inner transactions into one undo entry
        except Exception:
            tg.RollBack()
            raise

        # Capture merge + border + width snapshots for retention on future updates
        new_merge_snapshot  = _extract_merge_snapshot(grid)
        new_border_snapshot = _extract_border_snapshot(grid)
        new_width_snapshot  = _extract_column_width_snapshot(doc, new_schedule)

        return {
            "schedule_name":      schedule_name,
            "source_path":        filepath,
            "sheet_name":         selected_sheet,
            "options":            options,
            "schedule_unique_id": new_schedule.UniqueId if new_schedule else None,
            "ssi_unique_id":      ssi.UniqueId if ssi else None,
            "merge_snapshot":     new_merge_snapshot,
            "border_snapshot":    new_border_snapshot,
            "width_snapshot":     new_width_snapshot,
            "output_type":        "Schedule"
        }
    except Exception as ex:
        _ehm_error('Schedule Build Failed', str(ex))
        return None



# ─────────────────────────────────────────────────────────────────────────────


