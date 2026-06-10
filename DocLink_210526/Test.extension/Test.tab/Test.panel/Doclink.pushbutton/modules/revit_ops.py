# -*- coding: utf-8 -*-
"""
revit_ops.py
------------
Revit API operations for placing, updating, and removing DocLink images.

Public API
----------
place_or_replace(doc, view, file_path, ...)  -> dict
_make_image_type(doc, file_path, dpi)        -> ImageType|None
_element_by_unique_id(doc, unique_id)        -> Element|None
_get_view_center(view)                       -> XYZ
_get_image_center_point(element, view)       -> XYZ|None
_resolve_existing_instance_unique_ids(...)   -> list[str]
_all_instances_by_type(doc, type_uid)        -> list[dict]
_find_instances_doc_wide(doc, import_name)   -> list[(str, inst)]
_normalize_unique_ids(values)               -> list[str]
_record_element_unique_ids(record)           -> list[str]
_set_record_element_unique_ids(record, ids) -> None
"""

import re
import os

import System

from _imports import (
    Transaction, ImageType, ImageTypeOptions, ImageTypeSource,
    ImagePlacementOptions, ImageInstance, XYZ, BoxPlacement,
    BuiltInParameter, ElementId, FilteredElementCollector,
    Array, revit,
)
from utils import _safe_int, _as_bool


# ─────────────────────────────────────────────────────────────────────────────
# Name / identity helpers
# ─────────────────────────────────────────────────────────────────────────────

def _sanitize_import_name(raw):
    """Convert a file basename to a safe, human-readable import name."""
    if not raw:
        return ""
    name = re.sub(r'[^\w\s\-]', ' ', raw, flags=re.UNICODE)
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:128]


def _get_auto_user():
    """Derive best available user identity for record metadata."""
    try:
        app = revit.doc.Application
        u = str(app.Username).strip()
        if u:
            return u
    except Exception:
        pass
    try:
        import getpass
        u = getpass.getuser()
        if u:
            return u
    except Exception:
        pass
    try:
        u = (os.environ.get("USERNAME") or os.environ.get("USER") or "").strip()
        if u:
            return u
    except Exception:
        pass
    return "Unknown User"


# ─────────────────────────────────────────────────────────────────────────────
# Unique-ID helpers
# ─────────────────────────────────────────────────────────────────────────────

def _safe_element_int(value):
    try:
        value = int(value)
        return value if value > 0 else None
    except Exception:
        return None


def _normalize_element_ids(values):
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        values = [values]
    result, seen = [], set()
    for value in values:
        eid = _safe_element_int(value)
        if eid and eid not in seen:
            result.append(eid)
            seen.add(eid)
    return result


def _normalize_unique_ids(values):
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        values = [values]
    result, seen = [], set()
    for value in values:
        if value is None:
            continue
        try:
            uid = str(value).strip()
        except Exception:
            continue
        if uid and uid not in seen:
            result.append(uid)
            seen.add(uid)
    return result


def _record_element_unique_ids(record):
    ids = _normalize_unique_ids(record.get("element_unique_ids", []))
    primary = record.get("element_unique_id")
    if primary:
        primary = str(primary).strip()
        if primary and primary not in ids:
            ids.insert(0, primary)
    return ids


def _set_record_element_unique_ids(record, unique_ids):
    ids = _normalize_unique_ids(unique_ids)
    record["element_unique_ids"] = ids
    record["element_unique_id"]  = ids[0] if ids else None


# ─────────────────────────────────────────────────────────────────────────────
# Element lookups
# ─────────────────────────────────────────────────────────────────────────────

def _element_by_unique_id(doc, unique_id):
    """Resolve a Revit element by its UniqueId string (IronPython safe)."""
    if not unique_id:
        return None
    uid = str(unique_id).strip()
    if not uid:
        return None

    # Tier 1 — standard string overload
    try:
        el = doc.GetElement(uid)
        if el is not None:
            return el
    except Exception:
        pass

    # Tier 2 — force String overload via reflection
    try:
        mi = doc.GetType().GetMethod(
            "GetElement", Array[System.Type]([System.String])
        )
        if mi is not None:
            el = mi.Invoke(doc, Array[object]([uid]))
            if el is not None:
                return el
    except Exception:
        pass

    # Tier 3 — brute-force scan all ImageInstances
    try:
        collector = FilteredElementCollector(doc).OfClass(ImageInstance)
        for inst in collector:
            try:
                if inst.UniqueId == uid:
                    return inst
            except Exception:
                pass
    except Exception:
        pass

    return None


def _get_element_name(element):
    """Safely retrieve the Name of a Revit element."""
    if element is None:
        return ""
    try:
        val = element.Name
        if val is not None:
            return str(val)
    except Exception:
        pass
    try:
        p = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p is not None:
            s = p.AsString()
            if s:
                return s
    except Exception:
        pass
    try:
        p = element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p is not None:
            s = p.AsString()
            if s:
                return s
    except Exception:
        pass
    return ""


def _image_type_unique_id_from_instance(doc, element_unique_id):
    el = _element_by_unique_id(doc, element_unique_id)
    if el is None:
        return None
    try:
        tid = el.GetTypeId()
        if tid is not None and tid != ElementId.InvalidElementId:
            type_el = doc.GetElement(tid)
            if type_el is not None:
                return type_el.UniqueId
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Location / bounding-box helpers
# ─────────────────────────────────────────────────────────────────────────────

def _element_point(element):
    try:
        loc = element.Location
        if loc and hasattr(loc, "Point"):
            return loc.Point
    except Exception:
        pass
    return None


def _get_image_center_point(element, view=None):
    """Get the center XYZ of an ImageInstance."""
    try:
        if view is not None:
            bbox = element.get_BoundingBox(view)
            if bbox:
                return XYZ(
                    (bbox.Min.X + bbox.Max.X) / 2.0,
                    (bbox.Min.Y + bbox.Max.Y) / 2.0,
                    (bbox.Min.Z + bbox.Max.Z) / 2.0,
                )
    except Exception:
        pass
    try:
        bbox = element.get_BoundingBox(None)
        if bbox:
            return XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
    except Exception:
        pass
    try:
        loc = element.Location
        if loc and hasattr(loc, "Point"):
            return loc.Point
    except Exception:
        pass
    return None


def _get_view_center(view):
    """Return the center XYZ of a view's visible area."""
    try:
        bb = view.CropBox
        if bb is not None:
            return XYZ((bb.Min.X + bb.Max.X) / 2.0,
                       (bb.Min.Y + bb.Max.Y) / 2.0, 0.0)
    except Exception:
        pass
    try:
        bb = view.get_BoundingBox(None)
        if bb is not None:
            return XYZ((bb.Min.X + bb.Max.X) / 2.0,
                       (bb.Min.Y + bb.Max.Y) / 2.0, 0.0)
    except Exception:
        pass
    try:
        outline = view.Outline
        if outline is not None:
            return XYZ((outline.Min.U + outline.Max.U) / 2.0,
                       (outline.Min.V + outline.Max.V) / 2.0, 0.0)
    except Exception:
        pass
    print("[DocLinkManager] _get_view_center: using origin")
    return XYZ.Zero


# ─────────────────────────────────────────────────────────────────────────────
# Instance search helpers
# ─────────────────────────────────────────────────────────────────────────────

def _instances_in_view_by_type(doc, view, image_type_unique_id):
    type_el = _element_by_unique_id(doc, image_type_unique_id)
    if type_el is None:
        return []
    ids = []
    try:
        collector = FilteredElementCollector(doc, view.Id).OfClass(ImageInstance)
        for inst in collector:
            try:
                tid = inst.GetTypeId()
                if tid and tid == type_el.Id:
                    ids.append(inst.UniqueId)
            except Exception:
                pass
    except Exception:
        pass
    return _normalize_unique_ids(ids)


def _all_instances_by_type(doc, image_type_unique_id):
    """Find ALL ImageInstances of a given ImageType across the entire project."""
    type_el = _element_by_unique_id(doc, image_type_unique_id)
    if type_el is None:
        return []
    results = []
    try:
        collector = FilteredElementCollector(doc).OfClass(ImageInstance)
        for inst in collector:
            try:
                tid = inst.GetTypeId()
                if tid and tid == type_el.Id:
                    owner_view = None
                    try:
                        owner_view = doc.GetElement(inst.OwnerViewId)
                    except Exception:
                        pass
                    loc = _get_image_center_point(inst, owner_view)
                    results.append({
                        "unique_id": inst.UniqueId,
                        "view_id":   inst.OwnerViewId,
                        "location":  loc,
                    })
            except Exception:
                pass
    except Exception:
        pass
    return results


def _get_fallback_instances(doc, view, import_name=None):
    """Fallback: find ImageInstances by count or name match."""
    result = []
    try:
        collector = FilteredElementCollector(doc, view.Id).OfClass(ImageInstance)
        instances = list(collector)
        print("[DocLinkManager] Fallback: found {} ImageInstance(s) in view".format(len(instances)))
        if len(instances) == 1:
            pt = _get_image_center_point(instances[0], view)
            if pt is not None:
                return [(instances[0].UniqueId, pt)]
        if import_name and instances:
            for inst in instances:
                try:
                    type_el = doc.GetElement(inst.GetTypeId())
                    if type_el is None:
                        continue
                    type_name = _get_element_name(type_el)
                    if import_name.lower() in type_name.lower():
                        pt = _get_image_center_point(inst, view)
                        if pt is not None:
                            result.append((inst.UniqueId, pt))
                except Exception:
                    pass
    except Exception as ex:
        print("[DocLinkManager] Error in fallback: {}".format(ex))
    return result


def _find_instances_doc_wide(doc, import_name):
    """Last-resort: search ALL ImageInstances document-wide by type name."""
    results = []
    if not import_name:
        return results
    try:
        collector = FilteredElementCollector(doc).OfClass(ImageInstance)
        for inst in collector:
            try:
                type_el = doc.GetElement(inst.GetTypeId())
                if type_el is None:
                    continue
                type_name = _get_element_name(type_el)
                if import_name.lower() in type_name.lower():
                    results.append((inst.UniqueId, inst))
            except Exception:
                pass
    except Exception as ex:
        print("[DocLinkManager] _find_instances_doc_wide error: {}".format(ex))
    return results


def _resolve_existing_instance_unique_ids(doc, view, element_unique_ids=None,
                                          image_type_unique_id=None,
                                          import_name=None):
    """Resolve existing ImageInstance unique IDs from multiple fallback paths."""
    result, seen = [], set()
    normalized_ids = _normalize_unique_ids(element_unique_ids)
    for uid in normalized_ids:
        try:
            el = _element_by_unique_id(doc, uid)
            if el is not None and uid not in seen:
                result.append(uid)
                seen.add(uid)
        except Exception:
            pass
    if image_type_unique_id:
        type_matches = _instances_in_view_by_type(doc, view, image_type_unique_id)
        for uid in type_matches:
            if uid not in seen:
                result.append(uid)
                seen.add(uid)
    if not result and import_name:
        for uid, _ in _get_fallback_instances(doc, view, import_name):
            if uid not in seen:
                result.append(uid)
                seen.add(uid)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# ImageType factory
# ─────────────────────────────────────────────────────────────────────────────

def _make_image_type(doc, file_path, dpi=300):
    """Create a Revit ImageType from a raster file path."""
    dpi = _safe_int(dpi, 300)
    for attempt in range(3):
        try:
            if attempt == 0:
                opts = ImageTypeOptions(file_path, False, ImageTypeSource.Import)
                try:
                    opts.Resolution = dpi
                except Exception:
                    pass
                return ImageType.Create(doc, opts)
            elif attempt == 1:
                opts = ImageTypeOptions(file_path)
                try:
                    opts.Resolution = dpi
                except Exception:
                    pass
                return ImageType.Create(doc, opts)
            else:
                return ImageType.Create(doc, file_path)
        except Exception:
            continue
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Main placement operation
# ─────────────────────────────────────────────────────────────────────────────

def place_or_replace(doc, view, file_path,
                     existing_element_unique_ids=None,
                     existing_type_unique_id=None,
                     dpi=300, transparent=False, import_name=None):
    """
    Place a DocLink image, or replace ALL matching instances project-wide
    while preserving each instance's current location and owner view.

    Returns dict with 'element_unique_ids' and 'image_type_unique_id'.
    """
    with Transaction(doc, "DocLink Manager – Place") as t:
        t.Start()
        try:
            is_fresh = (not existing_element_unique_ids and not existing_type_unique_id)
            effective_name = None if is_fresh else import_name

            placements  = []
            deleted_uids = set()

            if not is_fresh and existing_type_unique_id:
                all_proj = _all_instances_by_type(doc, existing_type_unique_id)
                for info in all_proj:
                    placements.append((info["unique_id"], info["view_id"], info["location"]))

            old_unique_ids = _resolve_existing_instance_unique_ids(
                doc, view, existing_element_unique_ids,
                existing_type_unique_id, effective_name
            )
            for uid in old_unique_ids:
                if uid not in {p[0] for p in placements}:
                    try:
                        el = _element_by_unique_id(doc, uid)
                        if el is not None:
                            inst_view = None
                            try:
                                inst_view = doc.GetElement(el.OwnerViewId)
                            except Exception:
                                inst_view = view
                            pt = _get_image_center_point(el, inst_view or view)
                            owner_vid = el.OwnerViewId if inst_view else view.Id
                            placements.append((uid, owner_vid, pt))
                    except Exception:
                        pass

            fallback_instances = []
            if not placements and not is_fresh:
                fallback_instances = _get_fallback_instances(doc, view, import_name)
                for uid, pt in fallback_instances:
                    placements.append((uid, view.Id, pt))

            view_locs = []
            if placements:
                for uid, owner_vid, loc in placements:
                    try:
                        v = doc.GetElement(owner_vid)
                    except Exception:
                        v = None
                    if v is None:
                        v = view
                    if loc is None:
                        loc = _get_view_center(v)
                    view_locs.append((v, loc))

            if not view_locs:
                center = _get_view_center(view)
                view_locs = [(view, center)]

            # Create new ImageType and place in each view
            img_type = _make_image_type(doc, file_path, dpi=dpi)
            if not img_type:
                raise RuntimeError(
                    "Could not create ImageType.\n"
                    "The image file may be corrupt or unsupported by this Revit version."
                )

            new_unique_ids = []
            for target_v, pt in view_locs:
                opts = ImagePlacementOptions()
                opts.Location = pt if pt is not None else XYZ.Zero
                opts.PlacementPoint = BoxPlacement.Center
                inst = ImageInstance.Create(doc, target_v, img_type.Id, opts)
                bg = inst.get_Parameter(BuiltInParameter.IMPORT_BACKGROUND)
                if bg and not bg.IsReadOnly:
                    bg.Set(0 if _as_bool(transparent) else 1)
                new_unique_ids.append(inst.UniqueId)

            # Delete all old instances
            for uid, _, _ in placements:
                try:
                    old = _element_by_unique_id(doc, uid)
                    if old is not None:
                        doc.Delete(old.Id)
                        deleted_uids.add(uid)
                except Exception:
                    pass
            for uid, _ in fallback_instances:
                if uid not in deleted_uids:
                    try:
                        old = _element_by_unique_id(doc, uid)
                        if old is not None:
                            doc.Delete(old.Id)
                    except Exception:
                        pass

            t.Commit()
            return {
                "element_unique_ids":  new_unique_ids,
                "image_type_unique_id": img_type.UniqueId,
            }

        except Exception:
            t.RollBack()
            raise
