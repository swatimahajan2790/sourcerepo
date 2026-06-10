# -*- coding: utf-8 -*-
"""
persistence.py
--------------
DocLink + Schedule Import persistence layer.

Uses Revit ExtensibleStorage (DataStorage + Schema) to save/load
JSON-serialised record lists inside the Revit model file.

Public API
----------
load_records(doc)             -> list[dict]
save_records(doc, records)    -> None
load_schedule_records(doc)    -> list[dict]
save_schedule_records(doc, records) -> None
"""

import json
import os

from _imports import (
    Array, Guid, String,
    Schema, SchemaBuilder, Entity, AccessLevel,
    Transaction, FilteredElementCollector, ElementId,
    forms,
    DataStorageType,
)
from utils import _safe_int

# ── DocLink schema constants ──────────────────────────────────────────────────
_DOCLINK_SCHEMA_GUID = Guid("9D1D7A8D-5D8E-4F3A-8A6B-0CFE9C6F7A11")
_DOCLINK_SCHEMA_NAME = "DocLinkManagerRecords"
_DOCLINK_FIELD_NAME  = "RecordsJson"

# ── Schedule schema constants ─────────────────────────────────────────────────
# IMPORTANT: these GUIDs must match the original script.py values exactly to
# preserve data saved by earlier versions of the tool in existing Revit models.
_SCHED_SCHEMA_GUID = Guid("A2B3C4D5-E6F7-8901-2345-6789ABCDEF02")
_SCHED_SCHEMA_NAME = "ScheduleImportRecords"
_SCHED_FIELD_NAME  = "ScheduleRecordsJson"


# ─────────────────────────────────────────────────────────────────────────────
# Internal schema helpers
# ─────────────────────────────────────────────────────────────────────────────

def _doclink_schema():
    schema = Schema.Lookup(_DOCLINK_SCHEMA_GUID)
    if schema is not None:
        return schema
    builder = SchemaBuilder(_DOCLINK_SCHEMA_GUID)
    builder.SetSchemaName(_DOCLINK_SCHEMA_NAME)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.AddSimpleField(_DOCLINK_FIELD_NAME, String)
    return builder.Finish()


def _sched_schema():
    schema = Schema.Lookup(_SCHED_SCHEMA_GUID)
    if schema is not None:
        return schema
    builder = SchemaBuilder(_SCHED_SCHEMA_GUID)
    builder.SetSchemaName(_SCHED_SCHEMA_NAME)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.AddSimpleField(_SCHED_FIELD_NAME, String)
    return builder.Finish()


def _find_storage(doc, schema):
    """Return the first DataStorage element that holds this schema, or None."""
    try:
        collector = FilteredElementCollector(doc).OfClass(DataStorageType())
        for ds in collector:
            try:
                ent = ds.GetEntity(schema)
                if ent is not None and ent.IsValid():
                    return ds
            except Exception:
                pass
    except Exception:
        pass
    return None


def _find_doclink_storage(doc, schema=None):
    return _find_storage(doc, schema or _doclink_schema())


def _find_sched_storage(doc, schema=None):
    return _find_storage(doc, schema or _sched_schema())


# ─────────────────────────────────────────────────────────────────────────────
# Legacy helpers (migration path from old JSON sidecar files)
# ─────────────────────────────────────────────────────────────────────────────

def _legacy_data_file_path(doc):
    rvt_path = doc.PathName
    if not rvt_path:
        return os.path.join(os.environ["TEMP"], "DocLinkManager_unsaved.json")
    base = os.path.splitext(rvt_path)[0]
    return base + "_DocLinkManager.json"


def _load_legacy_records_from_json(doc):
    path = _legacy_data_file_path(doc)
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            pass
    return []


def _safe_element_int(val):
    try:
        return int(val)
    except Exception:
        return None


def _normalize_element_ids(ids):
    result = []
    for v in (ids or []):
        n = _safe_element_int(v)
        if n is not None:
            result.append(n)
    return result


def _normalize_loaded_records(records):
    if not isinstance(records, list):
        return []
    normalized = []
    for r in records:
        if not isinstance(r, dict):
            continue
        if "element_unique_ids" not in r:
            r["element_unique_ids"] = []
        if "element_unique_id" not in r:
            r["element_unique_id"] = None
        if "image_type_unique_id" not in r:
            r["image_type_unique_id"] = None
        if "instance_history" not in r or not isinstance(r.get("instance_history"), list):
            r["instance_history"] = []
        if "auto_named" not in r:
            r["auto_named"] = False
        if not r.get("user"):
            r["user"] = "Unknown User"
        for entry in r.get("instance_history", []):
            if isinstance(entry, dict) and "element_unique_id" not in entry:
                entry["element_unique_id"] = None
        normalized.append(r)
    return normalized


def _migrate_legacy_record_ids_to_unique_ids(doc, records):
    migrated = []
    for r in records:
        if not isinstance(r, dict):
            continue
        tracked_uids = []
        for eid in _normalize_element_ids(r.get("element_ids", [])):
            try:
                el = doc.GetElement(ElementId(eid))
                if el is not None and el.UniqueId not in tracked_uids:
                    tracked_uids.append(el.UniqueId)
            except Exception:
                pass
        primary_eid = _safe_element_int(r.get("element_id"))
        if primary_eid:
            try:
                el = doc.GetElement(ElementId(primary_eid))
                if el is not None and el.UniqueId not in tracked_uids:
                    tracked_uids.insert(0, el.UniqueId)
            except Exception:
                pass
        image_type_uid = None
        image_type_eid = _safe_element_int(r.get("image_type_id"))
        if image_type_eid:
            try:
                type_el = doc.GetElement(ElementId(image_type_eid))
                if type_el is not None:
                    image_type_uid = type_el.UniqueId
            except Exception:
                pass
        r["element_unique_ids"] = tracked_uids
        r["element_unique_id"] = tracked_uids[0] if tracked_uids else None
        r["image_type_unique_id"] = image_type_uid
        for entry in r.get("instance_history", []):
            if not isinstance(entry, dict):
                continue
            hist_uid = None
            hist_eid = _safe_element_int(entry.get("element_id"))
            if hist_eid:
                try:
                    hist_el = doc.GetElement(ElementId(hist_eid))
                    if hist_el is not None:
                        hist_uid = hist_el.UniqueId
                except Exception:
                    pass
            entry["element_unique_id"] = hist_uid
        r.pop("element_id", None)
        r.pop("element_ids", None)
        r.pop("image_type_id", None)
        migrated.append(r)
    return _normalize_loaded_records(migrated)


def _normalize_sched_records(records):
    if not isinstance(records, list):
        return []
    return [r for r in records if isinstance(r, dict)]


# ─────────────────────────────────────────────────────────────────────────────
# Public API — DocLink Image records
# ─────────────────────────────────────────────────────────────────────────────

def load_records(doc):
    schema = _doclink_schema()
    ds = _find_doclink_storage(doc, schema)
    if ds is not None:
        try:
            ent = ds.GetEntity(schema)
            if ent is not None and ent.IsValid():
                payload = ent.Get[String](_DOCLINK_FIELD_NAME)
                if payload:
                    return _normalize_loaded_records(json.loads(payload))
        except Exception:
            pass
    # One-time legacy migration from old JSON sidecar file
    legacy = _load_legacy_records_from_json(doc)
    if legacy:
        migrated = _migrate_legacy_record_ids_to_unique_ids(doc, legacy)
        if migrated:
            save_records(doc, migrated)
            return migrated
    return []


def save_records(doc, records):
    schema = _doclink_schema()
    payload = json.dumps(records, indent=2)
    try:
        with Transaction(doc, "DocLink Manager – Save Data") as t:
            t.Start()
            ds = _find_doclink_storage(doc, schema)
            if ds is None:
                ds = DataStorageType().GetMethod("Create").Invoke(
                    None, Array[object]([doc]))
            ent = Entity(schema)
            ent.Set[String](_DOCLINK_FIELD_NAME, payload)
            ds.SetEntity(ent)
            t.Commit()
    except Exception as ex:
        forms.alert("Could not save DocLink Manager data in the model:\n{}".format(str(ex)))


# ─────────────────────────────────────────────────────────────────────────────
# Public API — Schedule Import records
# ─────────────────────────────────────────────────────────────────────────────

def load_schedule_records(doc):
    schema = _sched_schema()
    ds = _find_sched_storage(doc, schema)
    if ds is not None:
        try:
            ent = ds.GetEntity(schema)
            if ent is not None and ent.IsValid():
                payload = ent.Get[String](_SCHED_FIELD_NAME)
                if payload:
                    records = json.loads(payload)
                    if isinstance(records, list):
                        return _normalize_sched_records(records)
        except Exception:
            pass
    return []


def save_schedule_records(doc, records):
    schema = _sched_schema()
    payload = json.dumps(records, indent=2)
    try:
        with Transaction(doc, "Schedule Import – Save Data") as t:
            t.Start()
            ds = _find_sched_storage(doc, schema)
            if ds is None:
                ds = DataStorageType().GetMethod("Create").Invoke(
                    None, Array[object]([doc]))
            ent = Entity(schema)
            ent.Set[String](_SCHED_FIELD_NAME, payload)
            ds.SetEntity(ent)
            t.Commit()
    except Exception as ex:
        forms.alert("Could not save Schedule Import data:\n{}".format(str(ex)))
