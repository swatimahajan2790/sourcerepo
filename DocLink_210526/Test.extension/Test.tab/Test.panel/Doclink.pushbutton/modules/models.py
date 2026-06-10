# -*- coding: utf-8 -*-
"""
models.py
---------
Observable data-row models for WPF DataGrid binding.

Classes
-------
DocLinkRow   – one row in the Tab 1 (DocLink Images) DataGrid.
ScheduleRow  – one row in the Tab 2 (Schedule Import) DataGrid.

Both implement INotifyPropertyChanged so WPF two-way binding works
under IronPython.
"""

from _imports import (
    INotifyPropertyChanged, PropertyChangedEventArgs
)


# ─────────────────────────────────────────────────────────────────────────────
# DocLinkRow
# ─────────────────────────────────────────────────────────────────────────────

class DocLinkRow(INotifyPropertyChanged):
    """One row in the Tab-1 manager grid."""

    _fields = ["idx", "name", "file_type", "user", "path_type",
               "range_info", "last_updated", "status", "instance_count",
               "deleted_count", "view_name", "_record_id"]

    def __init__(self, record_id, idx, name, file_type, user,
                 path_type, range_info, last_updated, status,
                 instance_count=0, deleted_count=0, view_name=""):
        self._idx            = str(idx)
        self._name           = name
        self._file_type      = file_type.upper() if file_type else ""
        self._user           = user
        self._path_type      = path_type
        self._range_info     = range_info or "—"
        self._last_updated   = last_updated or "—"
        self._status         = status or "—"
        self._instance_count = str(instance_count)
        self._deleted_count  = str(deleted_count)
        self._view_name      = view_name or "—"
        self.__record_id     = record_id
        self._handlers       = []

    # INotifyPropertyChanged
    def add_PropertyChanged(self, handler):
        self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def _notify(self, prop):
        args = PropertyChangedEventArgs(prop)
        for h in self._handlers:
            h(self, args)

    def _prop(name):
        private = "_" + name
        def getter(self):
            return getattr(self, private)
        def setter(self, v):
            setattr(self, private, v)
            self._notify(name)
        return property(getter, setter)

    idx            = _prop("idx")
    name           = _prop("name")
    file_type      = _prop("file_type")
    user           = _prop("user")
    path_type      = _prop("path_type")
    range_info     = _prop("range_info")
    last_updated   = _prop("last_updated")
    status         = _prop("status")
    instance_count = _prop("instance_count")
    deleted_count  = _prop("deleted_count")
    view_name      = _prop("view_name")

    @property
    def _record_id(self):
        return self.__record_id


# ─────────────────────────────────────────────────────────────────────────────
# ScheduleRow
# ─────────────────────────────────────────────────────────────────────────────

class ScheduleRow(INotifyPropertyChanged):
    """One row in the Tab-2 Schedule Import DataGrid."""

    def __init__(self, record_id, idx, name, source_file, sheet_name,
                 options_info, last_updated, status,
                 retain_settings=True, path_type="Absolute"):
        self._idx            = str(idx)
        self._name           = name
        self._source_file    = source_file
        self._sheet_name     = sheet_name
        self._options_info   = options_info or "—"
        self._last_updated   = last_updated or "—"
        self._status         = status or "—"
        self._retain_settings = retain_settings
        self._path_type      = path_type
        self.__record_id     = record_id
        self._handlers       = []

    def add_PropertyChanged(self, handler):
        self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def _notify(self, prop):
        args = PropertyChangedEventArgs(prop)
        for h in self._handlers:
            h(self, args)

    def _prop(name):
        private = "_" + name
        def getter(self):
            return getattr(self, private)
        def setter(self, v):
            setattr(self, private, v)
            self._notify(name)
        return property(getter, setter)

    idx             = _prop("idx")
    name            = _prop("name")
    source_file     = _prop("source_file")
    sheet_name      = _prop("sheet_name")
    options_info    = _prop("options_info")
    last_updated    = _prop("last_updated")
    status          = _prop("status")
    retain_settings = _prop("retain_settings")
    path_type       = _prop("path_type")

    @property
    def _record_id(self):
        return self.__record_id
