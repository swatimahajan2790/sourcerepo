# -*- coding: utf-8 -*-

import clr

import os

import math



from Autodesk.Revit.DB import (

    Transaction, View, ViewDrafting, ViewType, ElementId,

    FilteredElementCollector, ViewFamilyType, ViewFamily,

    Line, XYZ, DetailLine, DetailCurve, TextNote, TextNoteOptions,

    HorizontalTextAlignment, VerticalTextAlignment,

    FormattedText, TextRange, ElementTypeGroup, TextNoteType,

    ViewDuplicateOption, ImageType, ImageTypeOptions, ImagePlacementOptions,

    ImageInstance, BoxPlacement, ImageTypeSource, BuiltInParameter,

    FilledRegion, FilledRegionType, CurveLoop, Color as RevitColor,

    FillPatternElement, GraphicsStyle, GraphicsStyleType,

    ElementTransformUtils, BuiltInCategory,

)

import math as _math

from System.Collections.Generic import List



from logger import LogManager





#   Filled-region colour cache: (r,g,b) -> ElementId of FilledRegionType  

_FILL_TYPE_CACHE = {}



# Cache of ImageType ElementIds keyed by absolute file path so we don't

# re-import the same image twice in a run.

_IMAGE_TYPE_CACHE = {}





def _get_element_name(elem):
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        p = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p and p.AsString():
            return p.AsString()
    except Exception: pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.AsString():
            return p.AsString()
    except Exception: pass
    try:
        nm = getattr(elem, "Name", None)
        if nm: return nm
    except Exception: pass
    try:
        from Autodesk.Revit.DB import Element
        return Element.Name.GetValue(elem)
    except Exception: pass
    return None


def _ensure_image_type(doc, image_path):

    """Find or create an ImageType for the given file path. Caller must be

    inside an active transaction.



    Mirrors the schedule path's `_import_image_to_schedule_cell` ImageType

    construction, which is the proven-working code path.

    """

    if not image_path:

        return None

    cached = _IMAGE_TYPE_CACHE.get(image_path)

    if cached is not None:

        try:

            if doc.GetElement(cached) is not None:

                return cached

        except Exception:

            pass

    if not os.path.exists(image_path):

        LogManager.warning("[IMG] file missing: {}".format(image_path))

        return None

    try:

        opts = ImageTypeOptions(image_path, False, ImageTypeSource.Import)

        try:

            opts.Resolution = 96

        except Exception:

            pass

        itype = ImageType.Create(doc, opts)

        if itype is None:

            LogManager.warning("[IMG] ImageType.Create returned None for {}".format(image_path))

            return None

        _IMAGE_TYPE_CACHE[image_path] = itype.Id

        LogManager.debug("[IMG] ImageType created id={} for {}".format(itype.Id, image_path))

        return itype.Id

    except Exception as ex:

        LogManager.warning(

            "[IMG] ImageType.Create EXCEPTION for {}: {}".format(image_path, ex))

        return None





def _place_image_in_cell(doc, view, x_left, x_right, y_top, y_bottom, image_path):

    """Place an ImageInstance fitted to the cell rectangle. Caller must be

    in an active transaction.



    Mirrors the schedule path's image-import flow:

        1. ImageType.Create(doc, ImageTypeOptions(path, False, Import))

        2. ImageInstance.Create(doc, viewId, typeId, XYZ location,

                                ImagePlacementOptions(BoxPlacement.Center))



    All failures are logged with [IMG] prefix so the user can see exactly

    where the placement is breaking.

    """

    if not image_path:

        return None



    #   Step 1: ImageType  

    type_id = _ensure_image_type(doc, image_path)

    if type_id is None:

        return None



    #   Step 2: compute placement geometry  

    cell_w = abs(x_right - x_left)

    cell_h = abs(y_top - y_bottom)

    cx = (x_left + x_right) / 2.0

    cy = (y_top + y_bottom) / 2.0

    loc = XYZ(cx, cy, 0)



    #   Step 3: ImagePlacementOptions  

    # Construct options and set the placement geometry on the options object.

    # The Location property holds the XYZ where the image is anchored;

    # BoxPlacementPoint says which corner/center is at that point.

    place_opts = None

    try:

        place_opts = ImagePlacementOptions(BoxPlacement.Center)

    except Exception:

        try:

            place_opts = ImagePlacementOptions()

        except Exception as ex:

            LogManager.warning(

                "[IMG] ImagePlacementOptions ctor failed: {}".format(ex))

            return None



    # Set the XYZ location on the options   required by the Revit 2022 4-arg

    # form of ImageInstance.Create.  setattr is harmless on builds where the

    # property doesn't exist (the 5-arg fallback below passes XYZ separately).

    try:

        setattr(place_opts, 'Location', loc)

    except Exception:

        pass

    # Defensive: also set BoxPlacementPoint in case the ctor overload was

    # the no-arg one (default BoxPlacementPoint may be TopLeft on some builds).

    for _pp in ('BoxPlacementPoint', 'PlacementPoint'):

        try:

            setattr(place_opts, _pp, BoxPlacement.Center)

        except Exception:

            pass



    #   Step 4: ImageInstance.Create  

    # Revit 2022 (initial): Create(doc, viewId, typeId, ImagePlacementOptions)

    #                         4 args, Location lives on the options.

    # Revit 2023+ / 2024+:  Create(doc, viewId, typeId, XYZ, ImagePlacementOptions)

    #                         5 args, XYZ passed separately.

    # Try 4-arg first since that's what your Revit reported, then fall back

    # to 5-arg for newer builds.

    inst = None

    _err4 = _err5 = None

    try:

        inst = ImageInstance.Create(doc, view, type_id, place_opts)

    except Exception as ex:

        _err4 = ex

        try:

            inst = ImageInstance.Create(doc, view, type_id, loc, place_opts)

        except Exception as ex2:

            _err5 = ex2

    if inst is None:

        LogManager.warning(

            "[IMG] ImageInstance.Create FAILED for {} in view {}   "

            "4-arg: {} / 5-arg: {}".format(

                image_path, getattr(view, 'Name', '?'), _err4, _err5))

        return None



    LogManager.debug(

        "[IMG] ImageInstance placed id={} at ({:.3f}, {:.3f}) in {}".format(

            inst.Id, cx, cy, getattr(view, 'Name', '?')))



    #   Step 5: size to fit cell, preserving aspect ratio  

    # Failures here are non-fatal: image stays at its native ImageInstance

    # size, which is still better than no image at all.

    try:

        target_w = cell_w

        target_h = cell_h

        try:

            itype = doc.GetElement(type_id)

            orig_w_px = float(getattr(itype, 'Width', 0) or 0)

            orig_h_px = float(getattr(itype, 'Height', 0) or 0)

            if orig_w_px > 0 and orig_h_px > 0 and cell_w > 0 and cell_h > 0:

                aspect = orig_w_px / orig_h_px

                cell_aspect = cell_w / cell_h

                if aspect > cell_aspect:

                    target_w = cell_w

                    target_h = cell_w / aspect

                else:

                    target_h = cell_h

                    target_w = cell_h * aspect

        except Exception:

            pass



        # Unlock proportions before sizing so width/height don't drag together.

        try:

            inst.LockProportions = False

        except Exception:

            pass



        # Try direct property assignment first (Revit 2024+).

        sized = False

        try:

            inst.Width = target_w

            inst.Height = target_h

            sized = True

        except Exception:

            pass

        # Fall back to BuiltInParameter   try both MODEL (drafting/legend/plan)

        # and SHEET (sheet placement) so this works regardless of view type.

        if not sized:

            for _wp in ('RASTER_MODELWIDTH', 'RASTER_SHEETWIDTH',

                        'IMAGE_SHEETWIDTH'):

                _bip = getattr(BuiltInParameter, _wp, None)

                if _bip is None:

                    continue

                try:

                    p = inst.get_Parameter(_bip)

                    if p and not p.IsReadOnly and target_w > 0:

                        p.Set(target_w)

                        break

                except Exception:

                    pass

            for _hp in ('RASTER_MODELHEIGHT', 'RASTER_SHEETHEIGHT',

                        'IMAGE_SHEETHEIGHT'):

                _bip = getattr(BuiltInParameter, _hp, None)

                if _bip is None:

                    continue

                try:

                    p = inst.get_Parameter(_bip)

                    if p and not p.IsReadOnly and target_h > 0:

                        p.Set(target_h)

                        break

                except Exception:

                    pass



        try:

            inst.LockProportions = True

        except Exception:

            pass

    except Exception as ex:

        LogManager.debug("[IMG] sizing skipped: {}".format(ex))



    return inst





# Wingdings-symbol   real Unicode mapping for common cases (matches schedule path).

# Only Wingdings-encoded Latin-1 codepoints are remapped   all other Unicode

# (bullets, dashes, smart quotes, degree, plus-minus, trademark, etc.) is

# preserved verbatim because Revit TextNotes render BMP characters correctly.

_SYMBOL_CHAR_MAP = {
    # Wingdings-encoded Latin-1 codepoints -> proper Unicode
    u'\xfc': u'\u2713',  # Wingdings tick (chr 252)       -> U+2713 CHECK MARK
    u'\xfb': u'\u2713',  # alt Wingdings tick (chr 251)   -> U+2713 CHECK MARK
    u'\xfd': u'\u2717',  # Wingdings ballot X (chr 253)   -> U+2717 BALLOT X
    u'\xfe': u'\u2718',  # Wingdings heavy X  (chr 254)   -> U+2718 HEAVY BALLOT X
    # Normalise check-mark variants to one canonical tick
    u'\u2714': u'\u2713',  # heavy check mark (U+2714)    -> U+2713
    u'\u221a': u'\u2713',  # square root / sqrt (U+221A)  -> U+2713
    u'\u2714': u'\u2713',  # heavy check (alternate)      -> U+2713
}





# Resolve the unicode-string constructor for both IronPython 2.7 and Python 3.

# Use getattr on the builtins module so the bare name `unicode` never appears

# (which keeps Py3 linters happy) while still picking it up at runtime in IPy.

try:

    import __builtin__ as _builtins  # IronPython 2.7

except ImportError:

    import builtins as _builtins      # Python 3

_UnicodeStr = getattr(_builtins, 'unicode', str)





def _excel_rotation_to_radians(text_rotation):

    """Convert Excel textRotation (0 180, or 255 = stacked) to a CCW angle in radians.



    Excel encoding:

        0         = horizontal

        1   90    = CCW rotation in degrees (90 = fully vertical CCW)

        91   180  = CW rotation, where degrees CW = value - 90 (180 = 90  CW)

        255       = stacked vertical text   treated as 90  CCW for layout

    """

    try:

        r = int(text_rotation or 0)

    except (TypeError, ValueError):

        return 0.0

    if r == 0:

        return 0.0

    if r == 255:

        return _math.pi / 2.0  # treat stacked as 90  CCW

    if 1 <= r <= 90:

        return _math.radians(r)

    if 91 <= r <= 180:

        return -_math.radians(r - 90)  # negative = CW

    return 0.0





def _normalize_text(s):

    """Coerce cell value to a Unicode string and remap legacy Wingdings glyphs."""

    if s is None:

        return u''

    if not isinstance(s, _UnicodeStr):

        try:

            s = _UnicodeStr(s)

        except Exception:

            try:

                s = str(s).decode('utf-8', 'replace')

            except Exception:

                s = u''

    out_chars = []

    for ch in s:

        out_chars.append(_SYMBOL_CHAR_MAP.get(ch, ch))

    return u''.join(out_chars)



# These are approximately white   don't waste a filled region on them

_WHITE_THRESHOLD = 250





def _rgb_is_white(rgb):

    if rgb is None:

        return True

    r, g, b = rgb

    return r >= _WHITE_THRESHOLD and g >= _WHITE_THRESHOLD and b >= _WHITE_THRESHOLD





def _get_solid_fill_pattern_id(doc):

    """Return the ElementId of the first solid-fill FillPatternElement found."""

    for fpe in FilteredElementCollector(doc).OfClass(FillPatternElement):

        try:

            if fpe.GetFillPattern().IsSolidFill:

                return fpe.Id

        except Exception:

            pass

    return ElementId.InvalidElementId





def _resolve_element_id(val):

    """Safely coerce a value returned by Duplicate() to an ElementId.



    Revit API inconsistency: some builds return the ElementId directly,

    others return the newly created Element."""

    if val is None:

        return ElementId.InvalidElementId

    if hasattr(val, 'Id'):

        return val.Id

    return val



def _get_fill_type_id(doc, rgb):

    """Find or create a solid FilledRegionType for the given RGB tuple.



    Colour is applied via the ForegroundPatternColor / BackgroundPatternColor

    properties (Revit 2019+ direct API) with a BuiltInParameter fallback for

    older builds.  The fill pattern is forced to solid so the region renders

    as a flat colour block.



    NOTE: FilteredElementCollector(doc).OfClass(FilledRegionType) returns 0

    elements in many Revit project templates even though types exist.

    WhereElementIsElementType() + isinstance() is the reliable alternative.

    """

    key = (int(rgb[0]), int(rgb[1]), int(rgb[2]))

    # Module-level cache check

    cached = _FILL_TYPE_CACHE.get(key)

    if cached is not None:

        try:

            if doc.GetElement(cached) is not None:

                return cached

        except Exception:

            pass

        _FILL_TYPE_CACHE.pop(key, None)  # stale   retry



    r, g, b = key

    name = 'DocLink_Fill_{:02X}{:02X}{:02X}'.format(r, g, b)



    #  

    # Shared helper: collect ALL FilledRegionType elements using the approach

    # that actually works when OfClass() returns nothing.

    #  

    def _all_frt():

        """Yield all FilledRegionType elements in the document."""

        # Primary: OfClass (fast, but sometimes misses)

        for _e in FilteredElementCollector(doc).OfClass(FilledRegionType):

            yield _e

        # 2. Secondary: OfCategory (OST_FilledRegion)
        try:
            for _e in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FilledRegion).WhereElementIsElementType():
                if _e.Id.IntegerValue not in seen:
                    seen.add(_e.Id.IntegerValue)
                    yield _e
        except: pass

        # Fallback: walk all element types (always works)

        try:

            seen = set()

            for _e in FilteredElementCollector(doc).WhereElementIsElementType():

                if ("FilledRegionType" in _e.GetType().FullName) and _e.Id.IntegerValue not in seen:

                    seen.add(_e.Id.IntegerValue)

                    yield _e

        except Exception:

            pass



    #   Phase 1: Search for an existing DocLink type with this exact name  
    for et in FilteredElementCollector(doc).WhereElementIsElementType():
        try:
            et_name = _get_element_name(et)
            if et_name and et_name.lower() == name.lower():
                LogManager.debug("_get_fill_type_id: found existing '{}' by direct name search id={}".format(name, et.Id))
                _FILL_TYPE_CACHE[key] = et.Id
                return et.Id
        except Exception:
            pass

    for frt in _all_frt():
        try:
            frt_name = _get_element_name(frt)
            if frt_name and frt_name.lower() == name.lower():
                LogManager.debug("_get_fill_type_id: found existing '{}' id={}".format(name, frt.Id))
                _FILL_TYPE_CACHE[key] = frt.Id
                return frt.Id
        except Exception:
            pass



    #   Phase 2: Find a base type to duplicate  

    solid_pat_id = _get_solid_fill_pattern_id(doc)

    base_type = None



    # Approach A: GetDefaultElementTypeId (most reliable)

    try:

        _def_id = doc.GetDefaultElementTypeId(ElementTypeGroup.FilledRegionType)

        if _def_id and _def_id != ElementId.InvalidElementId:

            base_type = doc.GetElement(_def_id)

            LogManager.debug("_get_fill_type_id: base via GetDefaultElementTypeId id={}".format(_def_id))

    except Exception as _ea:

        LogManager.debug("_get_fill_type_id: GetDefaultElementTypeId failed: {}".format(_ea))



    # Approach B: first result from _all_frt()

    if base_type is None:

        for frt in _all_frt():

            try:

                base_type = frt

                LogManager.debug("_get_fill_type_id: base via _all_frt() name={}".format(frt.Name))

                break

            except Exception:

                pass



    if base_type is None:

        try:

            for et in FilteredElementCollector(doc).WhereElementIsElementType():

                if isinstance(et, FilledRegionType):

                    base_type = et

                    LogManager.debug("_get_fill_type_id: base via WhereElementIsElementType")

                    break

        except Exception as _e3:

            LogManager.debug("_get_fill_type_id: WhereElementIsElementType failed: {}".format(_e3))



    if base_type is None:

        LogManager.warning("_get_fill_type_id: no FilledRegionType found to duplicate")

        return ElementId.InvalidElementId



    try:

        raw_dup = base_type.Duplicate(name)

        new_id  = _resolve_element_id(raw_dup)

        if new_id is None or new_id == ElementId.InvalidElementId:

            LogManager.warning("_get_fill_type_id: Duplicate returned invalid id for '{}'".format(name))

            return ElementId.InvalidElementId



        new_type = doc.GetElement(new_id)

        if new_type is None:

            LogManager.warning("_get_fill_type_id: GetElement returned None after Duplicate")

            return ElementId.InvalidElementId



        col = RevitColor(r, g, b)



        #   Method 1: direct property API (Revit 2019+)  

        _col_set = False

        try:

            new_type.ForegroundPatternColor = col

            _col_set = True

        except Exception:

            pass

        try:

            new_type.BackgroundPatternColor = col

        except Exception:

            pass



        #   Method 2: BuiltInParameter fallback  

        if not _col_set:

            for _bip in (BuiltInParameter.FILL_PATTERN_COLOR,

                         BuiltInParameter.FILL_PATTERN_COLOR_2):

                try:

                    p = new_type.get_Parameter(_bip)

                    if p and not p.IsReadOnly:

                        p.Set(col)

                        _col_set = True

                except Exception:

                    pass



        if not _col_set:

            LogManager.warning("_get_fill_type_id: could not set colour for '{}'".format(name))



        #   Force solid fill pattern  

        if solid_pat_id != ElementId.InvalidElementId:

            try:

                new_type.ForegroundPatternId = solid_pat_id

            except Exception:

                pass

            try:

                new_type.BackgroundPatternId = solid_pat_id

            except Exception:

                pass

            for _bip in (BuiltInParameter.FILL_PATTERN_ID,):

                try:

                    p = new_type.get_Parameter(_bip)

                    if p and not p.IsReadOnly:

                        p.Set(solid_pat_id)

                except Exception:

                    pass



        LogManager.debug("_get_fill_type_id: created '{}' id={} colour=({},{},{})".format(

            name, new_id, r, g, b))

        _FILL_TYPE_CACHE[key] = new_id

        return new_id



    except Exception as ex:

        LogManager.warning("_get_fill_type_id failed for ({},{},{}): {}".format(r, g, b, ex))

        return ElementId.InvalidElementId





def _draw_fill(doc, view, x1, x2, y1, y2, fill_type_id):

    """Draw a solid filled region covering the cell rectangle.



    y1 is top_y (larger value in Revit coords), y2 is bot_y (smaller).

    Revit requires a counter-clockwise CurveLoop when viewed from +Z.

    With top_y > bot_y the CCW order is:

        bottom-left -> bottom-right -> top-right -> top-left.

    """

    if fill_type_id is None or fill_type_id == ElementId.InvalidElementId:

        return

    # Ensure correct CCW winding: lo = min(y), hi = max(y)

    lo_y = min(y1, y2)

    hi_y = max(y1, y2)

    lo_x = min(x1, x2)

    hi_x = max(x1, x2)

    # Guard: zero-dimension cells crash Line.CreateBound

    _TOL = 1e-6

    if (hi_x - lo_x) < _TOL or (hi_y - lo_y) < _TOL:

        LogManager.debug("_draw_fill: cell too small ({:.6f} x {:.6f}), skipped".format(

            hi_x - lo_x, hi_y - lo_y))

        return

    try:

        loop = CurveLoop()

        loop.Append(Line.CreateBound(XYZ(lo_x, lo_y, 0), XYZ(hi_x, lo_y, 0)))  # bottom L->R

        loop.Append(Line.CreateBound(XYZ(hi_x, lo_y, 0), XYZ(hi_x, hi_y, 0)))  # right  up

        loop.Append(Line.CreateBound(XYZ(hi_x, hi_y, 0), XYZ(lo_x, hi_y, 0)))  # top    R->L

        loop.Append(Line.CreateBound(XYZ(lo_x, hi_y, 0), XYZ(lo_x, lo_y, 0)))  # left   down

        profiles = List[CurveLoop]([loop])

        fr = FilledRegion.Create(doc, fill_type_id, view.Id, profiles)

        if fr is None:

            LogManager.warning("FilledRegion.Create returned None for type_id={}".format(fill_type_id))

    except Exception as ex:

        LogManager.warning("FilledRegion creation failed: {}".format(ex))





#   Line style cache  



# Built-in Excel border-style   Revit line-style-name defaults.

# Mirrors BorderGraphicsStyles._EXCEL_TO_REVIT in schedule_tab.py so both the

# Schedule path and the Drafting View / Legend path produce the same visual

# weights when the user has not configured an explicit mapping.

_EXCEL_TO_REVIT_DEFAULTS = {

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



# Keyword hints for auto-matching dashed/dotted project line styles  

# matches the logic in BorderGraphicsStyles._AUTO_KEYWORD_HINTS.

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





def _collect_detail_line_style_ids(doc):

    """

    Return {name: ElementId} for GraphicsStyles that are valid for DetailCurve.



    Uses the canonical OST_Lines.SubCategories approach (same as

    schedule_tab._collect_line_styles_dict).  Using the broad

    FilteredElementCollector(GraphicsStyle) returns IDs for Walls/Doors/etc.

    which silently fail when assigned to DetailCurve.LineStyle.



    Both bracketed and plain forms are registered so lookup is

    format-agnostic.

    """

    result = {}



    def _reg(nm, eid):

        if not nm or eid is None:

            return

        result[nm] = eid

        if nm.startswith('<') and nm.endswith('>'):

            plain = nm[1:-1]

            if plain and plain not in result:

                result[plain] = eid

        else:

            br = '<' + nm + '>'

            if br not in result:

                result[br] = eid



    #   Step 1: OST_Lines subcategories (canonical, valid for DetailCurve)  

    _populated = False

    try:

        lines_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)

        if lines_cat is not None:

            # Register the parent style itself

            try:

                top_gs = lines_cat.GetGraphicsStyle(GraphicsStyleType.Projection)

                if top_gs is not None:

                    _reg(lines_cat.Name, top_gs.Id)

            except Exception:

                pass

            # Register every subcategory (Thin Lines, Medium Lines, etc.)

            try:

                for sub in lines_cat.SubCategories:

                    try:

                        gs = sub.GetGraphicsStyle(GraphicsStyleType.Projection)

                        if gs is not None:

                            _reg(sub.Name, gs.Id)

                            _populated = True

                    except Exception:

                        pass

            except Exception:

                pass

    except Exception as ex:

        LogManager.debug("_collect_detail_line_style_ids OST_Lines failed: {}".format(ex))



    #   Step 2: broad fallback (older Revit / restricted template)  

    if not _populated:

        LogManager.debug("_collect_detail_line_style_ids: OST_Lines empty, using GraphicsStyle fallback")

        for gs in FilteredElementCollector(doc).OfClass(GraphicsStyle):

            try:

                if gs.GraphicsStyleType == GraphicsStyleType.Projection and gs.Name:

                    _reg(gs.Name, gs.Id)

            except Exception:

                pass



    LogManager.debug("_collect_detail_line_style_ids: {} style(s) found".format(len(result)))

    return result





def _build_line_style_cache(doc, line_style_map):

    """

    Convert {xl_border_style: revit_line_style_name} to {xl_border_style: ElementId}.



    Applies the same built-in Excel Revit defaults as BorderGraphicsStyles so

    borders render with correct visual weight even when the user has not opened

    the Configure Line Styles dialog.  User-configured overrides always win.

    Bracket/plain form variants (<Thin Lines>   Thin Lines) are normalised so

    a saved mapping survives project templates that differ on this convention.

    """

    # Use OST_Lines-based collector so only valid DetailCurve styles are returned

    name_to_id = _collect_detail_line_style_ids(doc)



    # Merge built-in defaults with user overrides (user wins).

    effective_map = dict(_EXCEL_TO_REVIT_DEFAULTS)

    if line_style_map:

        for xl_s, rv_name in line_style_map.items():

            if rv_name:

                effective_map[xl_s] = rv_name



    # Auto-upgrade dashed/dotted defaults when the project has a line style

    # whose name contains a matching keyword and the user hasn't explicitly

    # mapped that Excel style   mirrors BorderGraphicsStyles auto-upgrade logic.

    user_keys = set(line_style_map.keys()) if line_style_map else set()

    project_names = list(name_to_id.keys())

    for xl_style, hints in _AUTO_KEYWORD_HINTS.items():

        if xl_style in user_keys:

            continue

        for hint in hints:

            for nm in project_names:

                if hint in nm.lower().strip('<>'):

                    effective_map[xl_style] = nm

                    break

            else:

                continue

            break



    cache = {}

    for xl_style, rv_name in effective_map.items():

        if rv_name and rv_name in name_to_id:

            cache[xl_style] = name_to_id[rv_name]



    LogManager.debug("_build_line_style_cache: {} style(s) mapped from {} xl styles".format(

        len(cache), len(effective_map)))

    return cache





#   TextNoteType helpers  



def _ensure_text_note_type(doc, font_name, size_pt, pt_mm=2.834):

    """

    Find or create a TextNoteType matching font_name+size_pt.

    Naming: DocLink_Text_<Font>_<Size>pt. Caller must be inside a transaction.

    """

    target_name = 'DocLink_Text_{}_{}pt'.format(

        (font_name or 'Default').replace(' ', ''), round(float(size_pt or 0), 1))

    size_ft = (float(size_pt or 0) / float(pt_mm) / 25.4 / 12.0)



    # First check if our managed type already exists

    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):

        try:
            tnt_name = _get_element_name(tnt)
            if tnt_name == target_name:
                return tnt.Id
        except Exception:

            pass



    # Find a base type to duplicate

    base_type = None

    default_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

    if default_id and default_id != ElementId.InvalidElementId:

        try:

            base_type = doc.GetElement(default_id)

        except Exception:

            base_type = None

    if base_type is None:

        for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):

            base_type = tnt

            break

    if base_type is None:

        return ElementId.InvalidElementId



    try:

        new_id   = base_type.Duplicate(target_name)

        new_type = doc.GetElement(new_id)



        # Set font name

        try:

            font_param = new_type.get_Parameter(BuiltInParameter.TEXT_FONT)

            if font_param and not font_param.IsReadOnly and font_name:

                font_param.Set(font_name)

        except Exception:

            pass



        # Set text size (in feet)

        try:

            size_param = new_type.get_Parameter(BuiltInParameter.TEXT_SIZE)

            if size_param and not size_param.IsReadOnly and size_ft > 0:

                size_param.Set(size_ft)

        except Exception:

            pass



        # Set transparent background so filled regions show through text areas

        try:

            bg_param = new_type.get_Parameter(BuiltInParameter.TEXT_BACKGROUND)

            if bg_param and not bg_param.IsReadOnly:

                bg_param.Set(1)  # 1 = Transparent

        except Exception:

            pass



        # Enable word wrap on the type

        try:

            wrap_param = new_type.get_Parameter(BuiltInParameter.TEXT_WRAPPING)

            if wrap_param and not wrap_param.IsReadOnly:

                wrap_param.Set(1)  # 1 = Wrap

        except Exception:

            pass



        return new_id

    except Exception as ex:

        LogManager.debug("_ensure_text_note_type failed for {} {}pt: {}".format(

            font_name, size_pt, ex))

        return ElementId.InvalidElementId





def _build_tnt_map_cache(doc, text_note_map, text_scale=1.0, pt_mm=2.834):

    """

    Convert {'FontName_size.0': tnt_name | '__CREATE__'} to {'FontName_size.0': ElementId}.

    For '__CREATE__' values, calls _ensure_text_note_type.  Caller must be in a transaction

    if any '__CREATE__' is present.

    """

    if not text_note_map:

        return {}

    name_to_id = {}

    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):

        try:
            tnt_name = _get_element_name(tnt)
            if tnt_name:
                name_to_id[tnt_name] = tnt.Id
        except Exception:

            pass

    cache = {}

    for key, tnt_name in text_note_map.items():

        if tnt_name == '__CREATE__':

            # Parse key "FontName_size.X" into font + size and ensure the type

            try:

                parts = key.rsplit('_', 1)

                fnm = parts[0] if parts else ''

                fpt = float(parts[1]) if len(parts) > 1 else 0

            except Exception:

                fnm, fpt = '', 0

            fpt_scaled = fpt * float(text_scale)
            fid = _ensure_text_note_type(doc, fnm, fpt_scaled, pt_mm)

            if fid and fid != ElementId.InvalidElementId:

                cache[key] = fid

        elif tnt_name and tnt_name in name_to_id:

            cache[key] = name_to_id[tnt_name]

    return cache





def _build_fill_map_cache(doc, fill_region_map):
    """
    Convert {'RRGGBB': frt_name | '__CREATE__'} to {(r,g,b): ElementId}.
    Caller must be in a transaction if any '__CREATE__' is present.
    """
    if not fill_region_map:
        return {}

    name_to_id = {}
    # 1. Class-based collection
    for frt in FilteredElementCollector(doc).OfClass(FilledRegionType):
        try:
            if frt.Name:
                name_to_id[frt.Name] = frt.Id
        except Exception:
            pass

    # 2. String-based fallback (ensures we catch everything)
    try:
        seen_ids = set(name_to_id.values())
        for et in FilteredElementCollector(doc).WhereElementIsElementType():
            if ("FilledRegionType" in et.GetType().Name) and et.Id not in seen_ids:
                try:
                    if et.Name:
                        name_to_id[et.Name] = et.Id
                        seen_ids.add(et.Id)
                except Exception:
                    pass
    except Exception:
        pass

    cache = {}

    """

    Convert {'RRGGBB': frt_name | '__CREATE__'} to {(r,g,b): ElementId}.

    Caller must be in a transaction if any '__CREATE__' is present.

    """

    if not fill_region_map:

        return {}

    name_to_id = {}

    # Collect existing FilledRegionType names using the same multi-approach

    # as _get_fill_type_id to handle templates where OfClass returns nothing.

    for frt in FilteredElementCollector(doc).OfClass(FilledRegionType):
        try:
            frt_name = _get_element_name(frt)
            if frt_name:
                name_to_id[frt_name] = frt.Id
        except Exception:
            pass

    # Fallback: walk all element types (always works)
    try:
        seen_ids = set(name_to_id.values())
        for et in FilteredElementCollector(doc).WhereElementIsElementType():
            if ("FilledRegionType" in et.GetType().Name) and et.Id not in seen_ids:
                try:
                    et_name = _get_element_name(et)
                    if et_name:
                        name_to_id[et_name] = et.Id
                        seen_ids.add(et.Id)
                except Exception:
                    pass
    except Exception:
        pass

    cache = {}

    for hex_key, frt_name in fill_region_map.items():

        try:

            r = int(hex_key[0:2], 16)

            g = int(hex_key[2:4], 16)

            b = int(hex_key[4:6], 16)

        except Exception:

            continue

        rgb_tuple = (r, g, b)

        if frt_name == '__CREATE__':

            fid = _get_fill_type_id(doc, rgb_tuple)

            if fid and fid != ElementId.InvalidElementId:

                cache[rgb_tuple] = fid

        elif frt_name and frt_name in name_to_id:

            cache[rgb_tuple] = name_to_id[frt_name]

    return cache





def _make_tnt_key(font_name, size_pt):

    return '{}_{}'.format(font_name or '', round(float(size_pt or 0), 1))





def _make_fill_key(rgb):

    r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])

    return (r, g, b)





#   Main entry point  



def create_graphic_view(doc, grid, options, existing_view_id=None):

    """

    Create a Drafting View or Legend and populate it with filled regions,

    TextNotes, and detail lines based on the GridModel 'grid'.

    """

    # Clear per-run caches so duplicate types/images aren't created across runs

    _FILL_TYPE_CACHE.clear()

    _IMAGE_TYPE_CACHE.clear()



    output_type = options.get('output_type', 'Drafting View')

    view_name   = options.get('name', 'DocLink_Table')



    #   Step 1: Find or Create View  

    view = None

    if existing_view_id:

        try:

            view = doc.GetElement(existing_view_id)

        except Exception:

            pass



    if not view or not view.IsValidObject:

        view = _find_existing_view(doc, view_name, output_type)



    if not view or not view.IsValidObject:

        view = _create_view(doc, view_name, output_type)



    if not view or not view.IsValidObject:

        LogManager.error("Could not create/find view: {}".format(view_name))

        return None



    view_id = view.Id



    # Line styles never auto-create   always safe to build outside transaction

    line_style_cache = _build_line_style_cache(doc, options.get('line_style_map') or {})

    default_tnt_name = options.get('default_text_note_type') or None



    # Grid-level font fallbacks   cells without explicit font use these

    grid_default_font_name = getattr(grid, 'default_font_name', 'Calibri') or 'Calibri'

    grid_default_font_pt   = getattr(grid, 'default_font_size_pt', 11.0) or 11.0



    #   Step 3: Build Table Geometry  

    with Transaction(doc, "Build Graphic Table") as t:

        t.Start()



        view = doc.GetElement(view_id)

        if not view or not view.IsValidObject:

            LogManager.error("View object invalid at start of transaction.")

            t.RollBack()

            return None



        # Build TNT and Fill caches inside the transaction so '__CREATE__'

        # markers can spawn new types as part of the same undo entry.

        text_scale = float(options.get('text_scale', 1.0))
        pt_mm      = float(options.get('excel_points_per_mm', 2.834))
        tnt_map_cache  = _build_tnt_map_cache(doc, options.get('text_note_map') or {}, text_scale, pt_mm)

        fill_map_cache = _build_fill_map_cache(doc, options.get('fill_region_map') or {})



        # Resolve "Apply one type to all fonts" mode   single_text_note_type

        # overrides per-font lookup for every cell in this view.

        # Format: None | '__CREATE__' (legacy) | '__CREATE__:FontName|size_pt' | existing_type_name

        _single_tnt_raw = options.get('single_text_note_type')  # None | str

        _single_tnt_id  = None

        if _single_tnt_raw and _single_tnt_raw.startswith('__CREATE__'):

            # Decode font/size if encoded in new format '__CREATE__:FontName|11.0'

            _cr_font = grid_default_font_name

            _cr_size = grid_default_font_pt

            if ':' in _single_tnt_raw:

                try:

                    _payload = _single_tnt_raw.split(':', 1)[1]

                    _parts   = _payload.split('|', 1)

                    _cr_font = _parts[0].strip() or _cr_font

                    _cr_size = float(_parts[1]) if len(_parts) > 1 else _cr_size

                except Exception:

                    pass

            _cr_size_scaled = _cr_size * text_scale
            _single_tnt_id = _ensure_text_note_type(doc, _cr_font, _cr_size_scaled, pt_mm)

            LogManager.debug("[TNT] Single mode create: {} {:.1f}pt -> id={}".format(

                _cr_font, _cr_size, _single_tnt_id))

        elif _single_tnt_raw:

            for _tnt in FilteredElementCollector(doc).OfClass(TextNoteType):

                try:
                    tnt_name = _get_element_name(_tnt)
                    if tnt_name == _single_tnt_raw:
                        _single_tnt_id = _tnt.Id
                        break
                except Exception:

                    pass

            if _single_tnt_id is None:

                LogManager.debug("[TNT] Single type '{}' not found, using project default".format(

                    _single_tnt_raw))

                _single_tnt_id = doc.GetDefaultElementTypeId(

                    ElementTypeGroup.TextNoteType)



        v_scale_val  = options.get('view_scale', 1)

        scale_by_view = options.get('scale_by_view', True)



        try:

            view.Scale = int(v_scale_val)

        except Exception:

            pass



        _clear_view(doc, view)



        row_scale = options.get('row_scale', 1.0)

        col_scale = options.get('col_scale', 1.0)

        # v_scale_geom : applied to lines/cells (constant model size unless user opts in)

        # v_scale_text : ALWAYS view_scale because TextNoteType is annotative  

        #                its actual model footprint is tnt_size_ft   view_scale_denominator

        v_scale_geom = float(v_scale_val) if scale_by_view else 1.0

        v_scale_text = float(v_scale_val)



        # Per-TNT size cache so we don't query each cell

        _tnt_size_cache = {}

        def _tnt_size_ft(tnt_id):

            if tnt_id in _tnt_size_cache:

                return _tnt_size_cache[tnt_id]

            sz = 0.0082  # ~2.5mm fallback

            try:

                elem = doc.GetElement(tnt_id)

                if elem is not None:

                    sp = elem.get_Parameter(BuiltInParameter.TEXT_SIZE)

                    if sp:

                        sz = sp.AsDouble()

            except Exception:

                pass

            _tnt_size_cache[tnt_id] = sz

            return sz



        current_y = 0.0



        x_positions = [0.0]

        curr_x = 0.0

        for w_ft in grid.col_widths_ft:

            curr_x += w_ft * v_scale_geom

            x_positions.append(curr_x)



        processed_cells = set()



        for r_idx in range(grid.rows):

            row_h_scaled = grid.row_heights_ft[r_idx] * v_scale_geom



            for c_idx in range(grid.cols):

                cell = grid.get(r_idx, c_idx)

                if not cell or cell in processed_cells:

                    continue



                rs, cs = cell.merge_span if cell.merge_span else (1, 1)

                row_h_total_scaled = sum(grid.row_heights_ft[r_idx:r_idx + rs]) * v_scale_geom



                top_y   = current_y

                bot_y   = current_y - row_h_total_scaled

                left_x  = x_positions[c_idx]

                right_x = x_positions[c_idx + cs]



                #   Filled region (background colour)  

                if cell.fill_rgb and not _rgb_is_white(cell.fill_rgb):

                    rgb_key = _make_fill_key(cell.fill_rgb)

                    # Prefer user-mapped FilledRegionType; fall back to auto-create

                    ftype_id = fill_map_cache.get(rgb_key)

                    if ftype_id is None:

                        ftype_id = _get_fill_type_id(doc, cell.fill_rgb)

                    _draw_fill(doc, view, left_x, right_x, top_y, bot_y, ftype_id)



                #   Borders  

                _draw_borders(doc, view, left_x, right_x, top_y, bot_y,

                              cell, line_style_cache)



                #   Embedded image   placed instead of text  

                _has_image = bool(getattr(cell, 'image_path', None)) or bool(getattr(cell, 'is_image', False))

                if _has_image and getattr(cell, 'image_path', None):

                    _place_image_in_cell(doc, view,

                                         left_x, right_x, top_y, bot_y,

                                         cell.image_path)



                #   Text  

                if cell.value and not _has_image:

                    txt = _normalize_text(cell.value)
                    txt = txt.strip() if txt else txt
                    if txt:

                        h_align = _map_halign(cell.h_align)



                        # Fall back to grid defaults if cell has no explicit font.

                        # The Mapping Dialog used the resolved default font/size,

                        # so cells without overrides MUST look up under those keys

                        # for the user's mapping to take effect.

                        eff_font = cell.font_name or grid_default_font_name



                        # cell.font_size is stored in FEET by build_grid;

                        # grid_default_font_pt is in POINTS.

                        # Normalise to points for the TNT key and size calculations.

                        _FT_TO_PT = 12.0 * 72.0  # 1 ft = 864 pt

                        if cell.font_size and cell.font_size > 0:

                            eff_size = cell.font_size * _FT_TO_PT  # ft -> pt

                        else:

                            eff_size = grid_default_font_pt          # already in pt



                        # Resolve TextNoteType:

                        # 1. Single-type mode (user picked one type for all)

                        # 2. Per-font map (user configured per font)

                        # 3. Auto-match by font & size

                        if _single_tnt_id is not None:

                            tnt_id = _single_tnt_id

                        else:

                            tnt_key = _make_tnt_key(eff_font, eff_size)

                            tnt_id  = tnt_map_cache.get(tnt_key)

                            # Fallbacks for when the (font, size) key doesn't

                            # exactly match the configured map: try the same

                            # font with the grid default size, then any entry

                            # for the same font name regardless of size.

                            if tnt_id is None and tnt_map_cache:

                                _alt_key = _make_tnt_key(eff_font, grid_default_font_pt)

                                tnt_id = tnt_map_cache.get(_alt_key)

                            if tnt_id is None and tnt_map_cache:

                                _font_prefix = '{}_'.format(eff_font or '')

                                for _k, _v in tnt_map_cache.items():

                                    if _k.startswith(_font_prefix):

                                        tnt_id = _v

                                        break

                            if tnt_id is None:

                                tnt_id = _find_best_text_note_type(

                                    doc, eff_font, eff_size,

                                    preferred_name=default_tnt_name)



                        if tnt_id == ElementId.InvalidElementId:

                            LogManager.debug("Invalid TextNoteType at ({},{})".format(r_idx, c_idx))

                            continue



                        # ── Text geometry ─────────────────────────────────────────────────
                        # TextNote is an *annotative* element:
                        #   model footprint  = annotation size * v_scale_text
                        #   TextNote.Width   = paper/annotation-space width (NOT model space)
                        #
                        # All cell boundary coords (left_x, right_x, top_y, bot_y) are in
                        # model space, already multiplied by v_scale_geom.
                        # To get annotation-space equivalents we divide by v_scale_text.
                        #
                        # cell_width_annotation = (right_x - left_x) / v_scale_text
                        # This ensures the wrap box exactly matches the cell footprint on
                        # paper regardless of the view scale the user entered.

                        tnt_h_model = _tnt_size_ft(tnt_id) * v_scale_text  # model-space text height (ft)
                        tnt_h_annot = _tnt_size_ft(tnt_id)                  # annotation-space text height (ft)

                        # Margin expressed in annotation space (paper units)
                        cell_margin_annot = tnt_h_annot * 0.30   # ~30% of text height in annotation space
                        # Same margin converted to model space for anchor offset calculations
                        cell_margin_model = cell_margin_annot * v_scale_text

                        if cell.h_align and 'center' in cell.h_align.lower():
                            anchor_x = left_x + ((right_x - left_x) / 2.0)
                        elif cell.h_align and 'right' in cell.h_align.lower():
                            anchor_x = right_x - cell_margin_model
                        else:
                            anchor_x = left_x + cell_margin_model

                        # TextNote anchor.Y is the TOP edge of the text bounding box
                        if cell.v_align == 'top':
                            anchor_y = top_y - cell_margin_model
                        elif cell.v_align == 'bottom':
                            anchor_y = bot_y + cell_margin_model + tnt_h_model
                        else:  # center
                            anchor_y = top_y - (row_h_total_scaled / 2.0) + (tnt_h_model / 2.0)

                        tno = TextNoteOptions(tnt_id)
                        tno.HorizontalAlignment = h_align

                        try:
                            if not view.IsValidObject:
                                LogManager.error("View became invalid during generation.")
                                return view

                            _note = TextNote.Create(doc, view.Id, XYZ(anchor_x, anchor_y, 0), txt, tno)

                            # Set TextNote.Width in annotation space so the wrap boundary
                            # precisely matches the cell's paper-space width.
                            if _note is not None:
                                try:
                                    # cell width in model space -> divide by v_scale_text -> annotation space
                                    _cell_w_annot = (right_x - left_x) / max(v_scale_text, 1.0)
                                    _wrap_w = max(_cell_w_annot - 2.0 * cell_margin_annot, cell_margin_annot)
                                    _note.Width = _wrap_w
                                    LogManager.debug(
                                        "[TNT] ({},{}) cell_w_model={:.4f} v_scale={} wrap_annot={:.4f}".format(
                                            r_idx, c_idx, right_x - left_x, v_scale_text, _wrap_w))
                                except Exception:
                                    pass

                                # Apply Excel cell text rotation, if any

                                _trot = getattr(cell, 'text_rotation', 0) or 0

                                if _trot:

                                    try:

                                        _angle = _excel_rotation_to_radians(_trot)

                                        if _angle != 0.0:

                                            _axis = Line.CreateBound(

                                                XYZ(anchor_x, anchor_y, 0),

                                                XYZ(anchor_x, anchor_y, 1))

                                            ElementTransformUtils.RotateElement(

                                                doc, _note.Id, _axis, _angle)

                                    except Exception as _rex:

                                        LogManager.debug("Rotate TextNote failed: {}".format(_rex))

                        except Exception as ex:

                            LogManager.debug("TextNote failed at row {}: {}".format(r_idx, ex))



                if rs > 1 or cs > 1:

                    for dr in range(rs):

                        for dc in range(cs):

                            m_cell = grid.get(r_idx + dr, c_idx + dc)

                            if m_cell:

                                processed_cells.add(m_cell)



            current_y -= row_h_scaled



        t.Commit()



    return view





#   Helpers  



def _find_existing_view(doc, name, type_str):

    vt = ViewType.Legend if "Legend" in type_str else ViewType.DraftingView

    for v in FilteredElementCollector(doc).OfClass(View):

        try:

            if not v.IsTemplate and v.ViewType == vt and v.Name == name:

                return v

        except Exception:

            pass

    return None





def _create_view(doc, name, type_str):

    if "Legend" in type_str:

        return _create_legend_view(doc, name)

    return _create_drafting_view(doc, name)





def _create_drafting_view(doc, name):

    dvt = None

    for f in FilteredElementCollector(doc).OfClass(ViewFamilyType):

        if f.ViewFamily == ViewFamily.Drafting:

            dvt = f

            break

    if not dvt:

        return None

    with Transaction(doc, "Create Drafting View") as t:

        t.Start()

        v = ViewDrafting.Create(doc, dvt.Id)

        v.Name = name

        t.Commit()

        return v





class LegendTemplateMissingError(Exception):

    """Raised when project has no legend view to duplicate from."""

    pass





def _create_legend_view(doc, name):

    legend_template = None

    for v in FilteredElementCollector(doc).OfClass(View):

        try:

            if v.ViewType == ViewType.Legend and not v.IsTemplate:

                legend_template = v

                break

        except Exception:

            pass

    if not legend_template:

        raise LegendTemplateMissingError(

            "Revit cannot create a Legend view from scratch. "

            "Please create at least one Legend view in the project first, "

            "then re-run this import."

        )

    t = Transaction(doc, "Create Legend")

    t.Start()

    try:

        new_id = legend_template.Duplicate(ViewDuplicateOption.Duplicate)

        new_v  = doc.GetElement(new_id)

        new_v.Name = name

        t.Commit()

        return new_v

    except Exception:

        if t.HasStarted() and not t.HasEnded():

            t.RollBack()

        raise





def _clear_view(doc, view):

    """Delete only DocLink-managed elements: TextNote, DetailCurve, FilledRegion,

    ImageInstance."""

    cl = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()

    to_delete = []

    for e in cl:

        if isinstance(e, (TextNote, DetailCurve, DetailLine, FilledRegion,

                          ImageInstance)):

            to_delete.append(e.Id)

    if to_delete:

        try:

            doc.Delete(List[ElementId](to_delete))

        except Exception as ex:

            LogManager.debug("Clear view partial failure: {}".format(ex))





def _draw_borders(doc, view, x1, x2, y1, y2, cell, line_style_cache=None):

    if cell.border_top:

        _draw_line(doc, view, XYZ(x1, y1, 0), XYZ(x2, y1, 0),

                   _resolve_style(cell.border_top, line_style_cache))

    if cell.border_bottom:

        _draw_line(doc, view, XYZ(x1, y2, 0), XYZ(x2, y2, 0),

                   _resolve_style(cell.border_bottom, line_style_cache))

    if cell.border_left:

        _draw_line(doc, view, XYZ(x1, y1, 0), XYZ(x1, y2, 0),

                   _resolve_style(cell.border_left, line_style_cache))

    if cell.border_right:

        _draw_line(doc, view, XYZ(x2, y1, 0), XYZ(x2, y2, 0),

                   _resolve_style(cell.border_right, line_style_cache))





def _resolve_style(border_style_str, cache):

    """Return the mapped ElementId or None (use Revit default)."""

    if not cache or not border_style_str:

        return None

    return cache.get(border_style_str)





def _draw_line(doc, view, p1, p2, style_id=None):

    try:

        line   = Line.CreateBound(p1, p2)

        detail = doc.Create.NewDetailCurve(view, line)

        if style_id and style_id != ElementId.InvalidElementId:

            # For DetailCurve the correct API is the LineStyle property.

            # BuiltInParameter.BUILDING_CURVE_GSTYLE targets ModelCurve and

            # will silently fail on DetailCurve, so skip it entirely.

            try:

                gs_elem = doc.GetElement(style_id)

                if gs_elem is not None:

                    detail.LineStyle = gs_elem

            except Exception as _ls_ex:

                LogManager.debug("DetailCurve LineStyle assign failed: {}".format(_ls_ex))

    except Exception as ex:

        LogManager.debug("DetailCurve creation failed: {}".format(ex))





def _map_halign(s):

    if not s:

        return HorizontalTextAlignment.Left

    s = s.lower()

    if 'center' in s:

        return HorizontalTextAlignment.Center

    if 'right' in s:

        return HorizontalTextAlignment.Right

    return HorizontalTextAlignment.Left





def _find_best_text_note_type(doc, font_name, size_pt, preferred_name=None):

    """

    Resolve TextNoteType:

      1. preferred_name match (default TNT chosen by user)

      2. Exact font + size match

      3. Project default

    """

    cl = FilteredElementCollector(doc).OfClass(TextNoteType)



    if preferred_name:

        for tnt in cl:

            try:

                n = tnt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)

                if n and n.AsString() == preferred_name:

                    return tnt.Id

            except Exception:

                pass



    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):

        try:

            f = tnt.get_Parameter(BuiltInParameter.TEXT_FONT).AsString()

            s = tnt.get_Parameter(BuiltInParameter.TEXT_SIZE).AsDouble()

            size_ft = (size_pt / 72.0 / 12.0) if size_pt else 0

            if f == font_name and abs(s - size_ft) < 0.0001:

                return tnt.Id

        except Exception:

            pass



    return doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)


