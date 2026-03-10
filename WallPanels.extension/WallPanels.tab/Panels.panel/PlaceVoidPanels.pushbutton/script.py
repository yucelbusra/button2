# coding: ascii
"""
REVIT PANEL PLACEMENT - STRUCTURAL CORE CENTER ALIGNMENT WITH VOID CONTROL
Reads the optimized panel placement CSV and places panels in Revit.

VOID FAMILY PARAMETER MAPPING (verified from Family Types dialog):

  Panel size (instance, writable):
    Overall Width (default)       -> panel overall width
    Overall Height (default)      -> panel overall height

  Void 1 - Opening 1 (instance, writable):
    UNIT 1 WIDTH (default)        -> rough opening width
    UNIT 1 HEIGHT (default)       -> rough opening height
    VOID 1 JAMB CLR (default)     -> horiz clearance each side
    VOID 1 HEAD CLR (default)     -> clearance above unit
    VOID 1 SILL CLR (default)     -> clearance below unit
    Void 1 X Offset (default)     -> horiz position from panel left
    Void 1 Y Offset (default)     -> vert position from panel bottom
    VOID 1 (Yes/No)               -> visibility toggle for void 1

  Void 1 - Formula-driven (READ-ONLY, greyed out):
    Void 1 Height (default)  = VOID 1 SILL CLR + VOID 1 HEAD CLR + UNIT 1 HEIGHT
    Void 1 Width (default)   = 2 * VOID 1 JAMB CLR + UNIT 1 WIDTH

  Void 2 - Opening 2 (instance, writable):
    UNIT 2 WIDTH (default)        -> rough opening width
    UNIT 2 HEIGHT (default)       -> rough opening height
    VOID 2 JAMB CLR (default)     -> horiz clearance each side
    VOID 2 HEAD CLR (default)     -> clearance above unit
    VOID 2 SILL CLR               -> clearance below unit  <- NO "(default)" suffix
    Void 2 X Offset (default)     -> horiz position from panel left
    Void 2 Y Offset (default)     -> vert position from panel bottom
    VOID 2 (Yes/No)               -> visibility toggle for void 2

  Void 2 - Formula-driven (READ-ONLY, greyed out):
    Void 2 Height (default)  = VOID 2 SILL CLR + VOID 2 HEAD CLR + UNIT 2 HEIGHT
    Void 2 Width (default)   = 2 * VOID 2 JAMB CLR + UNIT 2 WIDTH
    Void 2 X Offset (default) = formula-driven READ-ONLY (cannot be set by script)
    Void 2 Y Offset (default) = formula-driven READ-ONLY (cannot be set by script)
    --> To control void 2 position: remove formulas from Void 2 X/Y Offset in family editor

  Other writable:
    Stud Depth
    Finish + Air Barrier Depth
    Sheathing Wrap_Left (default)
    Sheathing Wrap_Right (default)
    Fully Finished

FIXES INCLUDED:
- get_wall_base_elevation(): Level elevation + Base Offset for Z (fixes upper floors).
- get_core_center_offset_from_ext_face(): IsCoreLayer() walk for true structural core.
- get_wall_geometry_normalized(): Physical face measurement + analytical fallback.
- compute_panel_base_point(): Level-relative Z, not raw curve Z.
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector, Wall, Transaction, XYZ, Line,
    FamilySymbol, BuiltInCategory, BuiltInParameter, Transform, ElementId,
    DirectShape, ElementTransformUtils, FamilyPlacementType,
    HostObjectUtils, ShellLayerType, PlanarFace
)

try:
    from Autodesk.Revit.DB.Structure import StructuralType
except:
    from Autodesk.Revit.DB import Structure
    StructuralType = Structure.StructuralType

from pyrevit import revit
import csv
import os
import json
import math

import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import FolderBrowserDialog, DialogResult

doc = revit.doc

# ========== SETTINGS ==========
DEFAULT_INPUT_DIR = None
PANELS_FILE = "optimized_panel_placement.csv"
USE_FOLDER_PICKER = True
PANEL_FAMILY_NAME = None

SHOW_CUTOUTS = False
ALLOW_TYPE_PARAM_CHANGE = True

# --- VOID CONTROL ---
ENABLE_VOID_CONTROL = True  # Push cutout data into void params

# --- DEPTH SETTINGS ---
PANEL_THICKNESS_IN = 4.0
FAMILY_ORIGIN_LOCATION = "Center"  # "Center", "Front", or "Back"
MANUAL_DEPTH_OFFSET_IN = 0.0

# --- COORDINATE SETTINGS ---
PANEL_COORD_DEFAULT_REF = "start"
USE_CSV_ROTATION = True

# Runtime overrides
X_REF_OVERRIDE = None
ROTATION_OVERRIDE_DEG = None

# -----------------------------------------------------------------------
# PANEL SIZE -- exact parameter names from family (with "(default)" suffix)
# -----------------------------------------------------------------------
WIDTH_PARAM_CANDIDATES  = ["Overall Width (default)", "Overall Width"]
HEIGHT_PARAM_CANDIDATES = ["Overall Height (default)", "Overall Height"]

# -----------------------------------------------------------------------
# VOID 1 PARAMETERS -- exact names from Family Types dialog
# -----------------------------------------------------------------------
V1_UNIT_WIDTH   = "UNIT 1 WIDTH (default)"
V1_UNIT_HEIGHT  = "UNIT 1 HEIGHT (default)"
V1_JAMB_CLR     = "VOID 1 JAMB CLR (default)"
V1_HEAD_CLR     = "VOID 1 HEAD CLR (default)"
V1_SILL_CLR     = "VOID 1 SILL CLR (default)"
V1_X_OFFSET     = "Void 1 X Offset (default)"
V1_Y_OFFSET     = "Void 1 Y Offset (default)"
V1_VISIBLE      = "VOID 1 (default)"      # Yes/No instance visibility toggle

# Read-only (formula-driven) -- listed for documentation, NOT set by code:
# V1_WIDTH_RO  = "Void 1 Width (default)"   = 2*JAMB_CLR + UNIT_WIDTH
# V1_HEIGHT_RO = "Void 1 Height (default)"  = SILL_CLR + HEAD_CLR + UNIT_HEIGHT

# -----------------------------------------------------------------------
# VOID 2 PARAMETERS -- exact names from Family Types dialog
# NOTE: VOID 2 SILL CLR has no "(default)" suffix -- intentional
# -----------------------------------------------------------------------
V2_UNIT_WIDTH   = "UNIT 2 WIDTH (default)"
V2_UNIT_HEIGHT  = "UNIT 2 HEIGHT (default)"
V2_JAMB_CLR     = "VOID 2 JAMB CLR (default)"
V2_HEAD_CLR     = "VOID 2 HEAD CLR (default)"
V2_SILL_CLR     = "VOID 2 SILL CLR (default)"      # No "(default)" suffix -- as shown in family
V2_X_OFFSET     = "VOID 2 X OFFSET (default)"  # all-caps as shown in family
V2_Y_OFFSET     = None  # Void 2 Y Offset is formula-driven READ-ONLY -- cannot be set
V2_VISIBLE      = "VOID 2 (default)"      # Yes/No instance visibility toggle

# Read-only (formula-driven) -- NOT set by code:
# V2_WIDTH_RO  = "Void 2 Width (default)"   = 2*JAMB_CLR + UNIT_WIDTH
# V2_HEIGHT_RO = "Void 2 Height (default)"  = SILL_CLR + HEAD_CLR + UNIT_HEIGHT

# -----------------------------------------------------------------------
# CLEARANCE DEFAULTS BY OPENING TYPE (inches)
# Used when the CSV cutout does not include jamb_clr_in/head_clr_in/sill_clr_in.
# Keys match the "type" field in cutouts_json (compared case-insensitively).
# "default" is the fallback for unrecognised types.
# Edit these values to match your project clearance requirements.
# -----------------------------------------------------------------------
CLEARANCE_BY_TYPE = {
    "door":               {"jamb": 0.0, "head": 0.0, "sill": 0.0},
    "storefront/curtain": {"jamb": 0.0, "head": 0.0, "sill": 0.0},
    "window":             {"jamb": 0.0, "head": 0.0, "sill": 0.0},
    "default":            {"jamb": 0.0, "head": 0.0, "sill": 0.0},
}

# Minimum rough opening size to avoid "extrusion too thin" Revit errors
MIN_VOID_DIMENSION_IN = 0.5

USE_WALL_ENDCAP_EXTENSION = False
PANEL_SIDE_SIGN = 1


# ========== UTILITIES ==========

def _pick_input_folder(default_dir=None):
    try:
        fbd = FolderBrowserDialog()
        fbd.Description = "Select the folder containing '{0}'".format(PANELS_FILE)
        if default_dir and os.path.isdir(default_dir):
            fbd.SelectedPath = default_dir
        result = fbd.ShowDialog()
        if result == DialogResult.OK and fbd.SelectedPath:
            return str(fbd.SelectedPath)
    except:
        pass
    return None

def norm_id(val):
    try: return str(int(float(val)))
    except: return str(val).strip()

def get_wall_by_id(wall_id):
    try:
        elem_id = int(float(wall_id))
        element = doc.GetElement(ElementId(elem_id))
        if isinstance(element, Wall):
            return element
    except:
        pass
    return None

def _feet(val_inch):
    return float(val_inch) / 12.0

def _safe_float(val, default=0.0):
    try: return float(val)
    except: return default


# ========== GEOMETRY CORE ==========

def get_wall_base_elevation(wall):
    """
    Returns the true base elevation of the wall in feet.
    Combines: Base Level elevation + Base Offset parameter.
    More reliable than using location curve Z directly, especially on upper floors.
    """
    base_z = 0.0
    try:
        lvl_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        if lvl_id and lvl_id.IntegerValue > 0:
            level = doc.GetElement(lvl_id)
            if level:
                base_z = level.Elevation
    except:
        pass
    try:
        base_offset_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        if base_offset_param:
            base_z += base_offset_param.AsDouble()
    except:
        pass
    return base_z


def get_core_center_offset_from_ext_face(wall):
    """
    Returns the distance in feet, measured INWARD from the exterior face,
    to the CENTER of the structural/core layer(s).
    Works for simple all-core walls and compound walls with finish layers.
    """
    total_wall_width = wall.Width
    target_depth_inward = total_wall_width / 2.0  # Safe fallback

    try:
        cs = wall.WallType.GetCompoundStructure()
        if cs:
            layers = list(cs.GetLayers())
            cumulative = 0.0
            core_start = None
            core_end = None

            for i, layer in enumerate(layers):
                if cs.IsCoreLayer(i):
                    if core_start is None:
                        core_start = cumulative
                    core_end = cumulative + layer.Width
                cumulative += layer.Width

            if core_start is not None and core_end is not None:
                target_depth_inward = (core_start + core_end) / 2.0
                print("  [CORE] Core span: {0:.4f} to {1:.4f} ft, center at {2:.4f} ft from ext face".format(
                    core_start, core_end, target_depth_inward))
            else:
                print("  [CORE] No IsCoreLayer layers found, using wall center.")
        else:
            print("  [CORE] No CompoundStructure, using wall center.")
    except Exception as e:
        print("  [CORE] CompoundStructure read failed: {0}".format(e))

    return target_depth_inward


def get_wall_geometry_normalized(wall):
    """
    Returns:
        visual_left        - XYZ: left endpoint when facing the exterior
        visual_right       - XYZ: right endpoint when facing the exterior
        normalized_dir     - XYZ unit vector left -> right
        normal             - XYZ unit vector toward exterior (wall.Orientation)
        core_center_offset - float (ft): shift along normal from location line to core center.
                             Positive = toward exterior, negative = toward interior.
    """
    lc = wall.Location.Curve
    p0 = lc.GetEndPoint(0)
    p1 = lc.GetEndPoint(1)

    normal = wall.Orientation  # Points OUTWARD
    up = XYZ(0, 0, 1)
    visual_right_dir = normal.CrossProduct(up)

    dot0 = p0.DotProduct(visual_right_dir)
    dot1 = p1.DotProduct(visual_right_dir)
    if dot0 < dot1:
        visual_left, visual_right = p0, p1
    else:
        visual_left, visual_right = p1, p0

    normalized_dir = (visual_right - visual_left).Normalize()

    # Step 1: core center depth from exterior face
    core_depth_from_ext = get_core_center_offset_from_ext_face(wall)

    # Step 2: location line depth from exterior face via physical face geometry
    loc_line_depth_from_ext = None
    try:
        refs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior)
        if refs:
            face = wall.GetGeometryObjectFromReference(refs[0])
            if isinstance(face, PlanarFace):
                vec = p0 - face.Origin
                signed = vec.DotProduct(face.FaceNormal)
                loc_line_depth_from_ext = -signed  # positive = inward from ext face

                w = wall.Width
                if not (-(w * 0.05) <= loc_line_depth_from_ext <= w * 1.05):
                    print("  [WARN] loc_line_depth={0:.4f} outside wall width={1:.4f}, clamping.".format(
                        loc_line_depth_from_ext, w))
                    loc_line_depth_from_ext = max(0.0, min(w, loc_line_depth_from_ext))
                print("  [CORE] Location line depth from ext face: {0:.4f} ft".format(loc_line_depth_from_ext))
    except Exception as e:
        print("  [WARN] GetSideFaces failed: {0}".format(e))

    # Step 3: fallback -- read location line parameter and compute analytically
    if loc_line_depth_from_ext is None:
        try:
            w = wall.Width
            cs = wall.WallType.GetCompoundStructure()
            layers = list(cs.GetLayers()) if cs else []

            loc_line_param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
            loc_line = loc_line_param.AsInteger() if loc_line_param else 0

            if loc_line == 0:    # Wall Centerline
                loc_line_depth_from_ext = w / 2.0
            elif loc_line == 1:  # Core Centerline
                loc_line_depth_from_ext = core_depth_from_ext
            elif loc_line == 2:  # Finish Face Exterior
                loc_line_depth_from_ext = 0.0
            elif loc_line == 3:  # Finish Face Interior
                loc_line_depth_from_ext = w
            elif loc_line == 4:  # Core Face Exterior
                ext_finish = sum(
                    layers[i].Width for i in range(len(layers))
                    if not cs.IsCoreLayer(i) and i < next(
                        (j for j in range(len(layers)) if cs.IsCoreLayer(j)), 0)
                ) if cs else 0.0
                loc_line_depth_from_ext = ext_finish
            elif loc_line == 5:  # Core Face Interior
                last_core = next(
                    (j for j in reversed(range(len(layers))) if cs.IsCoreLayer(j)), len(layers) - 1
                ) if cs else len(layers) - 1
                int_finish = sum(
                    layers[i].Width for i in range(last_core + 1, len(layers))
                    if not cs.IsCoreLayer(i)
                ) if cs else 0.0
                loc_line_depth_from_ext = w - int_finish
            else:
                loc_line_depth_from_ext = w / 2.0

            print("  [CORE] Fallback: loc_line param={0}, depth from ext={1:.4f} ft".format(
                loc_line, loc_line_depth_from_ext))
        except Exception as e:
            print("  [WARN] Location line param fallback failed: {0}. Using wall center.".format(e))
            loc_line_depth_from_ext = wall.Width / 2.0

    # Step 4: offset = loc_line_depth - core_depth
    # Both measured inward from ext face. wall_normal points outward.
    # positive result = shift toward exterior = correct when loc_line is deeper than core center.
    core_center_offset = loc_line_depth_from_ext - core_depth_from_ext
    print("  [CORE] Final offset loc_line -> core center: {0:.4f} ft".format(core_center_offset))

    return visual_left, visual_right, normalized_dir, normal, core_center_offset


def compute_panel_base_point(wall, panel, rotation_deg=0.0, extra_z_offset_in=0.0):
    """
    Returns the insertion XYZ point aligning the panel origin to the wall structural core center.
    Z is from level elevation + base offset (not raw curve Z).
    """
    vis_left, vis_right, wall_dir, wall_normal, core_center_off = get_wall_geometry_normalized(wall)

    x_in  = _safe_float(panel.get("x_in",  0.0))
    y_in  = _safe_float(panel.get("y_in",  0.0))
    x_ref = (panel.get("x_ref", PANEL_COORD_DEFAULT_REF) or PANEL_COORD_DEFAULT_REF).lower().strip()

    if X_REF_OVERRIDE == "start": x_ref = "start"
    if X_REF_OVERRIDE == "end":   x_ref = "end"

    x_ft = _feet(x_in)
    y_ft = _feet(y_in)

    pt_xy = (vis_left + (wall_dir * x_ft)) if x_ref == "start" else (vis_right - (wall_dir * x_ft))

    # Level-relative Z -- fixes upper-floor placement
    wall_base_z = get_wall_base_elevation(wall)
    curve_z = vis_left.Z
    if abs(wall_base_z - curve_z) > 0.5:
        print("  [WARN] Wall {0}: Level Z={1:.3f} ft, Curve Z={2:.3f} ft -- using Level Z.".format(
            wall.Id.IntegerValue, wall_base_z, curve_z))

    base_z = wall_base_z + y_ft
    base_point_loc = XYZ(pt_xy.X, pt_xy.Y, base_z)

    # Depth: shift from location line to core center + family origin compensation
    calculated_offset = core_center_off
    p_thickness_ft = _feet(PANEL_THICKNESS_IN)

    rot = rotation_deg % 360
    if rot > 180: rot -= 360
    is_flipped = abs(rot) > 90.1

    if FAMILY_ORIGIN_LOCATION.lower() == "front":
        calculated_offset += (-p_thickness_ft / 2.0) if not is_flipped else (p_thickness_ft / 2.0)
    elif FAMILY_ORIGIN_LOCATION.lower() == "back":
        calculated_offset += (p_thickness_ft / 2.0) if not is_flipped else (-p_thickness_ft / 2.0)
    # "center" needs no adjustment

    calculated_offset += _feet(MANUAL_DEPTH_OFFSET_IN)
    calculated_offset += _feet(extra_z_offset_in)

    final_point = base_point_loc + (wall_normal * calculated_offset)
    return final_point, wall_dir, wall_normal


# ========== VOID CONTROL ==========

# Set to True on first panel placed, used to trigger one-time param name dump
_PARAM_DUMP_DONE = [False]


def _dump_param_names(inst):
    """
    Prints all parameter names available on the instance and its type.
    Runs once on the first placed panel to help debug NOT FOUND errors.
    """
    if _PARAM_DUMP_DONE[0]:
        return
    _PARAM_DUMP_DONE[0] = True

    print("  [DIAG] ---- INSTANCE PARAMETERS ----")
    for p in inst.Parameters:
        try:
            print("  [DIAG]   inst | '{0}'".format(p.Definition.Name))
        except:
            pass

    print("  [DIAG] ---- TYPE PARAMETERS ----")
    try:
        sym = inst.Symbol
        for p in sym.Parameters:
            try:
                print("  [DIAG]   type | '{0}'".format(p.Definition.Name))
            except:
                pass
    except:
        pass
    print("  [DIAG] ---- END PARAM DUMP ----")


def _resolve_param(inst, param_name):
    """
    Find a parameter by name, trying 4 variants in order:
      1. Exact name on instance
      2. Exact name on type
      3. Name without " (default)" suffix on instance
      4. Name without " (default)" suffix on type
    Returns the Parameter object or None.
    """
    variants = [param_name]
    if param_name.endswith(" (default)"):
        variants.append(param_name[:-len(" (default)")])
    elif not param_name.endswith(" (default)"):
        variants.append(param_name + " (default)")

    for name in variants:
        p = inst.LookupParameter(name)
        if p is not None:
            return p
        try:
            p = inst.Symbol.LookupParameter(name)
            if p is not None:
                return p
        except:
            pass
    return None


def _set_param(inst, param_name, value, label=""):
    """
    Set a parameter by name, trying with and without the " (default)" suffix
    on both instance and type. This handles Revit's inconsistent naming where
    the Family Types dialog shows "(default)" but LookupParameter does not.
    value: float for length (feet), int 1/0 for Yes/No.
    Returns True if successfully set.
    """
    try:
        p = _resolve_param(inst, param_name)

        if p is None:
            print("    [VOID] NOT FOUND: '{0}'".format(param_name))
            return False

        if p.IsReadOnly:
            print("    [VOID] READ-ONLY (formula): '{0}'".format(param_name))
            return False

        p.Set(value)

        if isinstance(value, float):
            print("    [VOID] SET {0} = {1:.3f} in".format(
                label or param_name, value * 12.0))
        elif value == 1:
            print("    [VOID] SET {0} = Yes".format(label or param_name))
        elif value == 0:
            print("    [VOID] SET {0} = No".format(label or param_name))
        else:
            print("    [VOID] SET {0} = {1}".format(label or param_name, value))
        return True

    except Exception as e:
        print("    [VOID] Error setting '{0}': {1}".format(param_name, e))
        return False


def _get_clearances(cutout):
    """
    Returns (jamb_clr_in, head_clr_in, sill_clr_in) for a cutout dict.

    Priority order:
      1. Explicit fields in the cutout dict (jamb_clr_in / head_clr_in / sill_clr_in)
      2. CLEARANCE_BY_TYPE lookup using the cutout "type" field (case-insensitive)
      3. CLEARANCE_BY_TYPE["default"] as final fallback
    """
    opening_type = str(cutout.get("type", "")).strip().lower()
    type_clr = CLEARANCE_BY_TYPE.get(opening_type, CLEARANCE_BY_TYPE.get("default", {}))

    jamb = _safe_float(cutout.get("jamb_clr_in", type_clr.get("jamb", 0.0)))
    head = _safe_float(cutout.get("head_clr_in", type_clr.get("head", 0.0)))
    sill = _safe_float(cutout.get("sill_clr_in", type_clr.get("sill", 0.0)))
    return jamb, head, sill


def set_void_parameters_for_cutouts(inst, panel_data):
    """
    Maps CSV cutout data onto the void family parameters.

    Cutout fields read from cutouts_json (all dimensions in inches):
        width_in    -> UNIT N WIDTH  (rough opening width)
        height_in   -> UNIT N HEIGHT (rough opening height)
        x_in        -> Void N X Offset (from panel left edge)
        y_in        -> Void N Y Offset (from panel bottom)
        type        -> used to look up clearances in CLEARANCE_BY_TYPE
        jamb_clr_in -> VOID N JAMB CLR (overrides type lookup if present)
        head_clr_in -> VOID N HEAD CLR (overrides type lookup if present)
        sill_clr_in -> VOID N SILL CLR (overrides type lookup if present)

    VOID 1 / VOID 2 Yes/No toggles:
        1 cutout  -> VOID 1 = 1 (on),  VOID 2 = 0 (off)
        2 cutouts -> VOID 1 = 1 (on),  VOID 2 = 1 (on)
        0 cutouts -> VOID 1 = 0 (off), VOID 2 = 0 (off)

    Void 1 Width/Height and Void 2 Width/Height are formula-driven (read-only).
    They update automatically once UNIT sizes and clearances are set.
    """
    if not ENABLE_VOID_CONTROL:
        return

    # Dump all parameter names on first panel so we can verify exact names
    _dump_param_names(inst)

    cutouts = panel_data.get("cutouts", [])
    num = len(cutouts)

    print("  [VOID] Panel '{0}': {1} cutout(s)".format(
        panel_data.get("panel_name", "?"), num))

    # ----------------------------------------------------------------
    # VOID 1
    # ----------------------------------------------------------------
    if num >= 1:
        c = cutouts[0]
        opening_type = c.get("type", "unknown")
        unit_w_in   = max(_safe_float(c.get("width_in",  0)), MIN_VOID_DIMENSION_IN)
        unit_h_in   = max(_safe_float(c.get("height_in", 0)), MIN_VOID_DIMENSION_IN)
        x_offset_in = _safe_float(c.get("x_in", 0))
        y_offset_in = _safe_float(c.get("y_in", 0))
        jamb_clr_in, head_clr_in, sill_clr_in = _get_clearances(c)

        print("  [VOID1] ON  | type={0} | {1:.2f}x{2:.2f} in | pos=({3:.2f},{4:.2f}) in | "
              "clr: J={5} H={6} S={7}".format(
              opening_type, unit_w_in, unit_h_in, x_offset_in, y_offset_in,
              jamb_clr_in, head_clr_in, sill_clr_in))

        _set_param(inst, V1_VISIBLE,     1,                  "VOID 1")
        _set_param(inst, V1_UNIT_WIDTH,  _feet(unit_w_in),   "UNIT 1 WIDTH")
        _set_param(inst, V1_UNIT_HEIGHT, _feet(unit_h_in),   "UNIT 1 HEIGHT")
        _set_param(inst, V1_JAMB_CLR,    _feet(jamb_clr_in), "VOID 1 JAMB CLR")
        _set_param(inst, V1_HEAD_CLR,    _feet(head_clr_in), "VOID 1 HEAD CLR")
        _set_param(inst, V1_SILL_CLR,    _feet(sill_clr_in), "VOID 1 SILL CLR")
        _set_param(inst, V1_X_OFFSET,    _feet(x_offset_in), "Void 1 X Offset")
        _set_param(inst, V1_Y_OFFSET,    _feet(y_offset_in), "Void 1 Y Offset")

    else:
        print("  [VOID1] OFF | no cutout")
        min_ft = _feet(MIN_VOID_DIMENSION_IN)
        panel_w_ft = _feet(_safe_float(panel_data.get("width_in", 120)))
        _set_param(inst, V1_VISIBLE,     0,          "VOID 1")
        _set_param(inst, V1_UNIT_WIDTH,  min_ft,     "UNIT 1 WIDTH")
        _set_param(inst, V1_UNIT_HEIGHT, min_ft,     "UNIT 1 HEIGHT")
        _set_param(inst, V1_JAMB_CLR,    0.0,        "VOID 1 JAMB CLR")
        _set_param(inst, V1_HEAD_CLR,    0.0,        "VOID 1 HEAD CLR")
        _set_param(inst, V1_SILL_CLR,    0.0,        "VOID 1 SILL CLR")
        _set_param(inst, V1_X_OFFSET,    panel_w_ft, "Void 1 X Offset")
        _set_param(inst, V1_Y_OFFSET,    0.0,        "Void 1 Y Offset")

    # ----------------------------------------------------------------
    # VOID 2
    # ----------------------------------------------------------------
    if num >= 2:
        c2 = cutouts[1]
        opening_type2 = c2.get("type", "unknown")
        unit_w2_in   = max(_safe_float(c2.get("width_in",  0)), MIN_VOID_DIMENSION_IN)
        unit_h2_in   = max(_safe_float(c2.get("height_in", 0)), MIN_VOID_DIMENSION_IN)
        x2_offset_in = _safe_float(c2.get("x_in", 0))
        y2_offset_in = _safe_float(c2.get("y_in", 0))
        jamb2_clr_in, head2_clr_in, sill2_clr_in = _get_clearances(c2)

        print("  [VOID2] ON  | type={0} | {1:.2f}x{2:.2f} in | pos=({3:.2f},{4:.2f}) in | "
              "clr: J={5} H={6} S={7}".format(
              opening_type2, unit_w2_in, unit_h2_in, x2_offset_in, y2_offset_in,
              jamb2_clr_in, head2_clr_in, sill2_clr_in))

        _set_param(inst, V2_VISIBLE,     1,                   "VOID 2")
        _set_param(inst, V2_UNIT_WIDTH,  _feet(unit_w2_in),   "UNIT 2 WIDTH")
        _set_param(inst, V2_UNIT_HEIGHT, _feet(unit_h2_in),   "UNIT 2 HEIGHT")
        _set_param(inst, V2_JAMB_CLR,    _feet(jamb2_clr_in), "VOID 2 JAMB CLR")
        _set_param(inst, V2_HEAD_CLR,    _feet(head2_clr_in), "VOID 2 HEAD CLR")
        _set_param(inst, V2_SILL_CLR,    _feet(sill2_clr_in), "VOID 2 SILL CLR")
        if V2_X_OFFSET:
            _set_param(inst, V2_X_OFFSET, _feet(x2_offset_in), "Void 2 X Offset")
        if V2_Y_OFFSET:
            _set_param(inst, V2_Y_OFFSET, _feet(y2_offset_in), "Void 2 Y Offset")
        else:
            print("    [VOID] Void 2 Y Offset is formula-driven, position set by family")

    else:
        print("  [VOID2] OFF | no second cutout")
        min_ft = _feet(MIN_VOID_DIMENSION_IN)
        panel_w_ft = _feet(_safe_float(panel_data.get("width_in", 120)))
        _set_param(inst, V2_VISIBLE,     0,          "VOID 2")
        _set_param(inst, V2_UNIT_WIDTH,  min_ft,     "UNIT 2 WIDTH")
        _set_param(inst, V2_UNIT_HEIGHT, min_ft,     "UNIT 2 HEIGHT")
        _set_param(inst, V2_JAMB_CLR,    0.0,        "VOID 2 JAMB CLR")
        _set_param(inst, V2_HEAD_CLR,    0.0,        "VOID 2 HEAD CLR")
        _set_param(inst, V2_SILL_CLR,    0.0,        "VOID 2 SILL CLR")
        if V2_X_OFFSET:
            _set_param(inst, V2_X_OFFSET, panel_w_ft, "Void 2 X Offset")

    if num > 2:
        print("  [VOID] WARNING: {0} cutouts in CSV but family supports only 2. "
              "Cutouts beyond index 1 are ignored.".format(num))


# ========== PLACEMENT ==========

def place_panel_family(wall, panel, symbol, extra_z_offset_in=0.0, is_cutout=False):
    if not ensure_symbol_active(symbol):
        return None

    # 1. Rotation first (affects depth offset calculation)
    rot_deg = 0.0
    if ROTATION_OVERRIDE_DEG is not None:
        rot_deg = ROTATION_OVERRIDE_DEG
    elif USE_CSV_ROTATION:
        try: rot_deg = _safe_float(panel.get("rotation_deg", 0.0))
        except: pass

    # 2. Compute insertion point
    try:
        pt, w_dir, w_norm = compute_panel_base_point(wall, panel, rot_deg, extra_z_offset_in)
    except Exception as e:
        print("[ERROR] Geometry calc failed: {0}".format(e))
        return None

    # 3. Place instance (wall-hosted preferred, fall back to level/unhosted)
    inst = None
    try:
        inst = doc.Create.NewFamilyInstance(pt, symbol, wall, StructuralType.NonStructural)
        if extra_z_offset_in == 0:
            print("  [PLACE] Hosted: {0}".format(panel.get("panel_name", "")))
    except:
        pass

    if not inst:
        try:
            lvl = get_wall_base_level(wall)
            if lvl:
                inst = doc.Create.NewFamilyInstance(pt, symbol, lvl, StructuralType.NonStructural)
            else:
                inst = doc.Create.NewFamilyInstance(pt, symbol, StructuralType.NonStructural)
            print("  [PLACE] Non-hosted: {0}".format(panel.get("panel_name", "")))
        except Exception as e:
            print("[ERROR] Placement failed: {0}".format(e))
            return None

    doc.Regenerate()

    # 4. Apply rotation
    if abs(rot_deg) > 0.001:
        try:
            axis = Line.CreateBound(pt, pt + XYZ(0, 0, 10))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.radians(rot_deg))
        except:
            pass

    # 5. Overall panel size
    set_size_parameters(inst, panel["width_in"], panel["height_in"], symbol)

    # 6. Void/opening parameters (skip for cutout visualizations)
    if not is_cutout:
        set_void_parameters_for_cutouts(inst, panel)
        doc.Regenerate()  # force formula recalc (Void Width/Height are formula-driven)

    # 7. Name / mark
    try:
        p = _find_param_by_candidates(inst, ["Name", "Panel Name", "Mark"])
        if p and not p.IsReadOnly:
            p.Set(panel.get("panel_name", ""))
    except:
        pass

    return inst


# ========== STANDARD HELPERS ==========

def get_element_name(element):
    try:
        p = element.get_Parameter(BuiltInParameter.SYMBOL_NAME)
        if p: return p.AsString()
    except: pass
    try: return element.Name
    except: return "Unknown"

def get_family_name(symbol):
    try: return symbol.Family.Name
    except: return "Unknown"

def get_all_family_symbols():
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    families_dict = {}
    for symbol in collector:
        families_dict.setdefault(get_family_name(symbol), []).append(symbol)
    return families_dict

def get_panel_family_symbol(family_name):
    from pyrevit import forms
    if family_name:
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
        for s in collector:
            if get_family_name(s) == family_name:
                return s, False
    families_dict = get_all_family_symbols()
    family_names  = sorted(families_dict.keys())
    family_names.insert(0, "< Use DirectShape (3D Solid Panels) >")
    selected_family = forms.SelectFromList.show(
        family_names, title="Select Panel Placement Method", button_name="Select", multiselect=False)
    if not selected_family: return None, False
    if selected_family == "< Use DirectShape (3D Solid Panels) >": return None, True
    symbols = families_dict[selected_family]
    if len(symbols) == 1: return symbols[0], False
    symbol_names  = [get_element_name(s) for s in symbols]
    selected_type = forms.SelectFromList.show(
        symbol_names, title="Select Family Type", button_name="Select", multiselect=False)
    if not selected_type: return symbols[0], False
    for symbol in symbols:
        if get_element_name(symbol) == selected_type:
            return symbol, False
    return symbols[0], False

def ensure_symbol_active(symbol):
    try:
        if not symbol.IsActive: symbol.Activate()
        return True
    except: return False

def _find_param_by_candidates(element, candidates):
    """Three-pass search: exact -> case-insensitive exact -> partial contains."""
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(nm == c for c in candidates): return p
        except: continue
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(nm.lower() == c.lower() for c in candidates): return p
        except: continue
    lower_cands = [c.lower() for c in candidates]
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(c in nm.lower() for c in lower_cands): return p
        except: continue
    return None

def set_size_parameters(inst, width_in, height_in, symbol=None):
    width_ft  = _feet(width_in)
    height_ft = _feet(height_in)
    changed = False

    w_param = _find_param_by_candidates(inst, WIDTH_PARAM_CANDIDATES)
    h_param = _find_param_by_candidates(inst, HEIGHT_PARAM_CANDIDATES)

    if w_param: print("    [SIZE] Width param:  {0}".format(w_param.Definition.Name))
    if h_param: print("    [SIZE] Height param: {0}".format(h_param.Definition.Name))

    try:
        if w_param and not w_param.IsReadOnly:
            w_param.Set(width_ft);  changed = True
        if h_param and not h_param.IsReadOnly:
            h_param.Set(height_ft); changed = True
    except: pass

    if not changed and ALLOW_TYPE_PARAM_CHANGE and symbol:
        try:
            wtp = _find_param_by_candidates(symbol, WIDTH_PARAM_CANDIDATES)
            htp = _find_param_by_candidates(symbol, HEIGHT_PARAM_CANDIDATES)
            if wtp and not wtp.IsReadOnly: wtp.Set(width_ft)
            if htp and not htp.IsReadOnly: htp.Set(height_ft)
        except: pass

def get_wall_base_level(wall):
    try:
        lvl_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        if lvl_id and lvl_id.IntegerValue > 0:
            return doc.GetElement(lvl_id)
    except: pass
    try:
        return FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).FirstElement()
    except: return None

def create_panel_as_direct_shape(wall, panel):
    try:
        pt, w_dir, w_norm = compute_panel_base_point(wall, panel)
        w_ft = _feet(_safe_float(panel.get("width_in",  0)))
        h_ft = _feet(_safe_float(panel.get("height_in", 0)))
        thk  = 1.0 / 12.0
        v1 = pt + (w_norm * 0.01)
        v2 = v1 + (w_dir * w_ft)
        v3 = v2 + XYZ(0, 0, h_ft)
        v4 = v1 + XYZ(0, 0, h_ft)
        v5 = pt + (w_norm * (0.01 + thk))
        v6 = v5 + (w_dir * w_ft)
        v7 = v6 + XYZ(0, 0, h_ft)
        v8 = v5 + XYZ(0, 0, h_ft)
        lines = [
            Line.CreateBound(v1, v2), Line.CreateBound(v2, v3),
            Line.CreateBound(v3, v4), Line.CreateBound(v4, v1),
            Line.CreateBound(v5, v6), Line.CreateBound(v6, v7),
            Line.CreateBound(v7, v8), Line.CreateBound(v8, v5),
            Line.CreateBound(v1, v5), Line.CreateBound(v2, v6),
            Line.CreateBound(v3, v7), Line.CreateBound(v4, v8),
        ]
        ds = DirectShape.CreateElement(doc, ElementId(int(BuiltInCategory.OST_GenericModel)))
        ds.SetShape(lines)
        ds.Name = panel.get("panel_name", "PanelSolid")
        print("  [DS] Created: {0}".format(ds.Name))
        return ds
    except Exception as e:
        print("  [DS] Failed: {0}".format(e))
        return None

def create_cutout_visualization(wall, panel, cutout_data, symbol, use_ds):
    if not use_ds and symbol:
        try:
            g_x = _safe_float(panel.get("x_in", 0)) + _safe_float(cutout_data.get("x_in", 0))
            g_y = _safe_float(panel.get("y_in", 0)) + _safe_float(cutout_data.get("y_in", 0))
            fake_panel = panel.copy()
            fake_panel.update({
                "panel_name": "CUT_" + str(cutout_data.get("id", "")),
                "x_in":       g_x,
                "y_in":       g_y,
                "width_in":   cutout_data.get("width_in",  0),
                "height_in":  cutout_data.get("height_in", 0),
                "cutouts":    [],
            })
            place_panel_family(wall, fake_panel, symbol, extra_z_offset_in=2.0, is_cutout=True)
            return True
        except: pass
    return False


# ========== MAIN ==========

def main():
    print("--- PANEL PLACEMENT: CORE CENTER ALIGNMENT WITH VOID CONTROL ---")

    if USE_FOLDER_PICKER:
        path = _pick_input_folder(DEFAULT_INPUT_DIR)
        if not path: return
        panels_path = os.path.join(path, PANELS_FILE)
    else:
        path = DEFAULT_INPUT_DIR or os.getcwd()
        panels_path = os.path.join(path, PANELS_FILE)

    if not os.path.exists(panels_path):
        print("CSV not found: " + panels_path)
        return

    panels = []
    with open(panels_path, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:    cutouts = json.loads(row.get("cutouts_json", "[]"))
            except: cutouts = []
            panels.append({
                "wall_id":      norm_id(row.get("wall_id")),
                "x_in":         row.get("x_in"),
                "y_in":         row.get("y_in"),
                "width_in":     row.get("width_in"),
                "height_in":    row.get("height_in"),
                "x_ref":        row.get("x_ref"),
                "panel_name":   row.get("panel_name"),
                "rotation_deg": row.get("rotation_deg"),
                "cutouts":      cutouts,
            })

    print("Loaded {0} panels.".format(len(panels)))

    sym, use_ds = get_panel_family_symbol(PANEL_FAMILY_NAME)
    if not sym and not use_ds: return

    from pyrevit import forms
    global X_REF_OVERRIDE, ROTATION_OVERRIDE_DEG

    if not use_ds:
        print("Using Family: " + get_family_name(sym))

        xref_ops = ["Use CSV Default", "Force Start (Left)", "Force End (Right)"]
        res = forms.SelectFromList.show(xref_ops, button_name="Set X Ref", multiselect=False)
        if res == xref_ops[1]: X_REF_OVERRIDE = "start"
        elif res == xref_ops[2]: X_REF_OVERRIDE = "end"

        rot_ops = ["Use CSV Rotation", "Force 0", "Force 90", "Force -90", "Force 180"]
        res = forms.SelectFromList.show(rot_ops, button_name="Set Rotation", multiselect=False)
        if res == rot_ops[1]: ROTATION_OVERRIDE_DEG = 0.0
        elif res == rot_ops[2]: ROTATION_OVERRIDE_DEG = 90.0
        elif res == rot_ops[3]: ROTATION_OVERRIDE_DEG = -90.0
        elif res == rot_ops[4]: ROTATION_OVERRIDE_DEG = 180.0

    panels_map = {}
    for p in panels:
        panels_map.setdefault(p["wall_id"], []).append(p)

    t = Transaction(doc, "Place Panels with Void Control")
    t.Start()

    count = 0
    for wid, wall_panels in panels_map.items():
        wall = get_wall_by_id(wid)
        if not wall:
            print("Wall {0} not found.".format(wid))
            continue

        print("\n--- Wall {0} ---".format(wid))
        for p in wall_panels:
            if use_ds:
                res = create_panel_as_direct_shape(wall, p)
            else:
                res = place_panel_family(wall, p, sym)
            if res: count += 1

            if SHOW_CUTOUTS:
                for c in p["cutouts"]:
                    create_cutout_visualization(wall, p, c, sym, use_ds)

    t.Commit()
    print("\nDone. Placed {0} panels.".format(count))


if __name__ == "__main__":
    main()