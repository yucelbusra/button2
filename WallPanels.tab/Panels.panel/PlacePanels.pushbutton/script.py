# -*- coding: utf-8 -*-
"""
REVIT PANEL PLACEMENT - STRUCTURAL CORE CENTER ALIGNMENT
Reads the optimized panel placement CSV and places panels in Revit.

FIXES:
- Calculates the CENTER of the Structural Core layer.
- Aligns the Panel's Center to the Core's Center.
- Ensures correct placement regardless of Wall Location Line (Finish Face, Centerline, etc).
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

SHOW_CUTOUTS = True
CUTOUT_THICKNESS_IN = 2.0
CUTOUT_DEPTH_IN = 3.0
ALLOW_TYPE_PARAM_CHANGE = True

# --- DEPTH SETTINGS ---
PANEL_THICKNESS_IN = 4.0
FAMILY_ORIGIN_LOCATION = "Center" 
MANUAL_DEPTH_OFFSET_IN = 0.0

# --- COORDINATE SETTINGS ---
PANEL_COORD_DEFAULT_REF = "start"
USE_CSV_ROTATION = True

# Runtime overrides
X_REF_OVERRIDE = None
ROTATION_OVERRIDE_DEG = None

WIDTH_PARAM_CANDIDATES = ["Width", "Panel Width", "W", "Overall Width", "Length", "L"]
HEIGHT_PARAM_CANDIDATES = ["Height", "Panel Height", "H", "Overall Height", "Thickness", "Depth"]

# Disable endcap extension to match exact drawing points
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
    except: pass
    return None

def norm_id(val):
    try: return str(int(float(val)))
    except: return str(val).strip()

def get_wall_by_id(wall_id):
    try:
        elem_id = int(float(wall_id))
        element = doc.GetElement(ElementId(elem_id))
        if isinstance(element, Wall): return element
    except: pass
    return None

def _feet(val_inch):
    return float(val_inch) / 12.0

# ========== GEOMETRY CORE ==========

# ========== GEOMETRY CORE (ROBUST VERSION) ==========

def get_wall_base_elevation(wall):
    """
    Returns the true base elevation of the wall in feet.
    Accounts for: Base Level elevation + Base Offset parameter.
    This is more reliable than using location curve Z directly.
    """
    base_z = 0.0
    try:
        lvl_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        if lvl_id and lvl_id.IntegerValue > 0:
            level = doc.GetElement(lvl_id)
            if level:
                base_z = level.Elevation  # This is in feet
    except:
        pass

    # Add base offset (can be positive or negative)
    try:
        base_offset_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        if base_offset_param:
            base_z += base_offset_param.AsDouble()  # Already in feet
    except:
        pass

    return base_z


def get_core_center_offset_from_ext_face(wall):
    """
    Returns the distance (in feet, measured INWARD from exterior face) 
    to the CENTER of the structural/core layer(s).
    
    Works correctly for:
    - Simple walls (all core, no finish layers) like your Concrete 8"
    - Compound walls with finish layers on either side
    - Walls with multiple core layers
    """
    total_wall_width = wall.Width
    target_depth_inward = total_wall_width / 2.0  # Safe fallback = wall center
    
    try:
        cs = wall.WallType.GetCompoundStructure()
        if cs:
            layers = list(cs.GetLayers())
            
            # Walk from exterior side (index 0), accumulating thickness
            # until we find core layers, then find their center
            cumulative = 0.0
            core_start = None
            core_end = None

            for i, layer in enumerate(layers):
                if cs.IsCoreLayer(i):
                    if core_start is None:
                        core_start = cumulative  # Where core begins (from ext face)
                    core_end = cumulative + layer.Width  # Where core ends so far
                cumulative += layer.Width
            
            if core_start is not None and core_end is not None:
                target_depth_inward = (core_start + core_end) / 2.0
                print("  [CORE] Core from {0:.4f} to {1:.4f} ft, center at {2:.4f} ft from ext face".format(
                    core_start, core_end, target_depth_inward
                ))
            else:
                print("  [CORE] No core layers found, using wall center.")
    except Exception as e:
        print("  [CORE] CompoundStructure failed: {0}".format(e))

    return target_depth_inward  # Always positive, measured inward from ext face


def get_wall_geometry_normalized(wall):
    """
    Returns wall geometry + core_center_offset needed to shift from 
    the wall's location line to the structural core center.
    Positive offset = toward exterior (outward along wall.Orientation).
    """
    lc = wall.Location.Curve
    p0 = lc.GetEndPoint(0)
    p1 = lc.GetEndPoint(1)

    normal = wall.Orientation  # Points OUTWARD (exterior direction)
    up = XYZ(0, 0, 1)
    visual_right_dir = normal.CrossProduct(up)

    dot0 = p0.DotProduct(visual_right_dir)
    dot1 = p1.DotProduct(visual_right_dir)
    if dot0 < dot1:
        visual_left, visual_right = p0, p1
    else:
        visual_left, visual_right = p1, p0

    normalized_dir = (visual_right - visual_left).Normalize()

    # Step 1: How far inward is the core center from the exterior face?
    core_depth_from_ext = get_core_center_offset_from_ext_face(wall)

    # Step 2: How far is the location line from the exterior face?
    # We measure this geometrically using the actual exterior face.
    loc_line_depth_from_ext = None
    try:
        refs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior)
        if refs:
            face = wall.GetGeometryObjectFromReference(refs[0])
            if isinstance(face, PlanarFace):
                # face.FaceNormal points outward (same direction as wall.Orientation)
                # Project location line point onto inward direction = negate FaceNormal dot
                vec = p0 - face.Origin
                # Positive result = p0 is on the OUTWARD side of face (wrong, shouldn't happen)
                # Negative result = p0 is inward from face (normal case)
                signed = vec.DotProduct(face.FaceNormal)
                loc_line_depth_from_ext = -signed  # Convert to "inward depth", positive = inward
                
                # Sanity check: should be between 0 and wall width
                w = wall.Width
                if not (-(w * 0.05) <= loc_line_depth_from_ext <= w * 1.05):
                    print("  [WARN] loc_line_depth={0:.4f} outside wall width={1:.4f}, clamping.".format(
                        loc_line_depth_from_ext, w))
                    loc_line_depth_from_ext = max(0.0, min(w, loc_line_depth_from_ext))
                    
                print("  [CORE] Location line is {0:.4f} ft inward from ext face".format(loc_line_depth_from_ext))
    except Exception as e:
        print("  [WARN] GetSideFaces failed: {0}".format(e))

    # Step 3: If face measurement failed, fall back to reading the location line parameter
    if loc_line_depth_from_ext is None:
        try:
            w = wall.Width
            cs = wall.WallType.GetCompoundStructure()
            layers = list(cs.GetLayers()) if cs else []
            
            loc_line_param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
            loc_line = loc_line_param.AsInteger() if loc_line_param else 0
            
            # Revit location line integer values:
            # 0 = Wall Centerline
            # 1 = Core Centerline  
            # 2 = Finish Face: Exterior
            # 3 = Finish Face: Interior
            # 4 = Core Face: Exterior
            # 5 = Core Face: Interior
            if loc_line == 0:    # Wall Centerline
                loc_line_depth_from_ext = w / 2.0
            elif loc_line == 1:  # Core Centerline — already at core center
                loc_line_depth_from_ext = core_depth_from_ext
            elif loc_line == 2:  # Finish Face Exterior — at ext face
                loc_line_depth_from_ext = 0.0
            elif loc_line == 3:  # Finish Face Interior — at int face
                loc_line_depth_from_ext = w
            elif loc_line == 4:  # Core Face Exterior
                ext_finish = sum(
                    layers[i].Width for i in range(len(layers))
                    if not cs.IsCoreLayer(i) and i < next(
                        (j for j in range(len(layers)) if cs.IsCoreLayer(j)), 0)
                )
                loc_line_depth_from_ext = ext_finish
            elif loc_line == 5:  # Core Face Interior
                int_finish = sum(
                    layers[i].Width for i in range(len(layers))
                    if not cs.IsCoreLayer(i) and i > next(
                        (j for j in reversed(range(len(layers))) if cs.IsCoreLayer(j)), len(layers))
                )
                loc_line_depth_from_ext = w - int_finish
            else:
                loc_line_depth_from_ext = w / 2.0  # Unknown, use center
                
            print("  [CORE] Fallback loc_line param={0}, depth={1:.4f} ft".format(
                loc_line, loc_line_depth_from_ext))
        except Exception as e:
            print("  [WARN] Location line param fallback failed: {0}".format(e))
            loc_line_depth_from_ext = wall.Width / 2.0

    # Step 4: Offset = how far to move from location line to reach core center
    # Both are "inward from ext face", so:
    # positive core_offset = move toward exterior (outward along wall.Orientation)
    # negative core_offset = move toward interior
    core_center_offset = loc_line_depth_from_ext - core_depth_from_ext

    print("  [CORE] Final offset from loc line to core center: {0:.4f} ft".format(core_center_offset))
    return visual_left, visual_right, normalized_dir, normal, core_center_offset


def compute_panel_base_point(wall, panel, rotation_deg=0.0, extra_z_offset_in=0.0):
    """
    Calculates insertion point aligning Panel CENTER to Wall CORE CENTER.
    
    ROBUST VERSION:
    - Uses level-relative Z (base elevation + base offset) rather than raw curve Z
    - y_in is treated as height above wall base, not above location curve Z
    """
    vis_left, vis_right, wall_dir, wall_normal, core_center_off = get_wall_geometry_normalized(wall)

    # --- XY Location ---
    x_in = float(panel.get("x_in", 0.0) or 0.0)
    y_in = float(panel.get("y_in", 0.0) or 0.0)
    x_ref = (panel.get("x_ref", PANEL_COORD_DEFAULT_REF) or PANEL_COORD_DEFAULT_REF).lower().strip()

    if X_REF_OVERRIDE == "start":
        x_ref = "start"
    if X_REF_OVERRIDE == "end":
        x_ref = "end"

    x_ft = _feet(x_in)
    y_ft = _feet(y_in)

    if x_ref == "start":
        pt_xy = vis_left + (wall_dir * x_ft)
    else:
        pt_xy = vis_right - (wall_dir * x_ft)

    # --- ROBUST Z: Use level elevation + base offset, NOT raw location curve Z ---
    # This correctly handles walls on upper levels and walls with base offsets
    wall_base_z = get_wall_base_elevation(wall)
    
    # Sanity check: if level-based Z differs wildly from location curve Z, log a warning
    curve_z = vis_left.Z
    if abs(wall_base_z - curve_z) > 0.5:  # More than 6 inches difference
        print("  [WARN] Wall {0}: Level Z={1:.3f} ft vs Curve Z={2:.3f} ft. Using Level Z.".format(
            wall.Id.IntegerValue, wall_base_z, curve_z
        ))

    base_z = wall_base_z + y_ft
    base_point_loc = XYZ(pt_xy.X, pt_xy.Y, base_z)

    # --- DEPTH ALIGNMENT LOGIC ---
    calculated_offset = core_center_off
    p_thickness_ft = _feet(PANEL_THICKNESS_IN)

    rot = rotation_deg % 360
    if rot > 180:
        rot -= 360
    is_flipped = abs(rot) > 90.1

    if FAMILY_ORIGIN_LOCATION.lower() == "center":
        pass
    elif FAMILY_ORIGIN_LOCATION.lower() == "front":
        if not is_flipped:
            calculated_offset -= (p_thickness_ft / 2.0)
        else:
            calculated_offset += (p_thickness_ft / 2.0)
    elif FAMILY_ORIGIN_LOCATION.lower() == "back":
        if not is_flipped:
            calculated_offset += (p_thickness_ft / 2.0)
        else:
            calculated_offset -= (p_thickness_ft / 2.0)

    calculated_offset += _feet(MANUAL_DEPTH_OFFSET_IN)
    calculated_offset += _feet(extra_z_offset_in)

    final_point = base_point_loc + (wall_normal * calculated_offset)

    return final_point, wall_dir, wall_normal

# ========== PLACEMENT ==========
def place_panel_family(wall, panel, symbol, extra_z_offset_in=0.0):
    if not ensure_symbol_active(symbol): return None
    
    # 1. Determine Rotation FIRST
    rot_deg = 0.0
    if ROTATION_OVERRIDE_DEG is not None:
        rot_deg = ROTATION_OVERRIDE_DEG
    elif USE_CSV_ROTATION:
        try: rot_deg = float(panel.get("rotation_deg", 0.0) or 0.0)
        except: pass

    # 2. Pass rotation to computation so it can fix the offset
    try:
        pt, w_dir, w_norm = compute_panel_base_point(wall, panel, rot_deg, extra_z_offset_in)
    except Exception as e:
        print("[ERROR] Geometry calc failed: {0}".format(e))
        return None

    inst = None
    try:
        inst = doc.Create.NewFamilyInstance(pt, symbol, wall, StructuralType.NonStructural)
        if extra_z_offset_in == 0:
            print("  [PLACE] Hosted: {0}".format(panel.get("panel_name", "")))
    except: pass
        
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
    
    # 3. Apply Rotation
    if abs(rot_deg) > 0.001:
        try:
            axis = Line.CreateBound(pt, pt + XYZ(0,0,10))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.radians(rot_deg))
        except: pass

    # ... (rest of your params logic) ...
    set_size_parameters(inst, panel["width_in"], panel["height_in"], symbol)
    try:
        p = _find_param_by_candidates(inst, ["Name", "Panel Name", "Mark"])
        if p and not p.IsReadOnly: p.Set(panel.get("panel_name",""))
    except: pass
    
    return inst

# [Standard Helpers]
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
        family_name = get_family_name(symbol)
        families_dict.setdefault(family_name, []).append(symbol)
    return families_dict

def get_panel_family_symbol(family_name):
    from pyrevit import forms
    if family_name:
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
        for s in collector:
            if get_family_name(s) == family_name: return s, False
    families_dict = get_all_family_symbols()
    family_names = sorted(families_dict.keys())
    family_names.insert(0, "< Use DirectShape (3D Solid Panels) >")
    selected_family = forms.SelectFromList.show(family_names, title="Select Panel Placement Method", button_name="Select", multiselect=False)
    if not selected_family: return None, False
    if selected_family == "< Use DirectShape (3D Solid Panels) >": return None, True
    symbols = families_dict[selected_family]
    if len(symbols) == 1: return symbols[0], False
    symbol_names = [get_element_name(s) for s in symbols]
    selected_type = forms.SelectFromList.show(symbol_names, title="Select Family Type", button_name="Select", multiselect=False)
    if not selected_type: return symbols[0], False
    for symbol in symbols:
        if get_element_name(symbol) == selected_type: return symbol, False
    return symbols[0], False

def ensure_symbol_active(symbol):
    try:
        if not symbol.IsActive: symbol.Activate()
        return True
    except: return False

def _find_param_by_candidates(element, candidates):
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(nm.lower() == cand.lower() for cand in candidates): return p
        except: continue
    lower_cands = [c.lower() for c in candidates]
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(c in nm.lower() for c in lower_cands): return p
        except: continue
    return None

def set_size_parameters(inst, width_in, height_in, symbol=None):
    width_ft = _feet(width_in)
    height_ft = _feet(height_in)
    changed = False
    w_param = _find_param_by_candidates(inst, WIDTH_PARAM_CANDIDATES)
    h_param = _find_param_by_candidates(inst, HEIGHT_PARAM_CANDIDATES)
    try:
        if w_param and not w_param.IsReadOnly:
            w_param.Set(width_ft)
            changed = True
        if h_param and not h_param.IsReadOnly:
            h_param.Set(height_ft)
            changed = True
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
        if lvl_id and lvl_id.IntegerValue > 0: return doc.GetElement(lvl_id)
    except: pass
    try: return FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).FirstElement()
    except: return None

def create_panel_as_direct_shape(wall, panel):
    try:
        pt, w_dir, w_norm = compute_panel_base_point(wall, panel)
        w_ft = _feet(panel.get("width_in",0))
        h_ft = _feet(panel.get("height_in",0))
        thk = 1.0/12.0
        v1 = pt + (w_norm * 0.01)
        v2 = v1 + (w_dir * w_ft)
        v3 = v2 + XYZ(0,0,h_ft)
        v4 = v1 + XYZ(0,0,h_ft)
        v5 = pt + (w_norm * (0.01+thk))
        v6 = v5 + (w_dir * w_ft)
        v7 = v6 + XYZ(0,0,h_ft)
        v8 = v5 + XYZ(0,0,h_ft)
        lines = [Line.CreateBound(v1,v2), Line.CreateBound(v2,v3), Line.CreateBound(v3,v4), Line.CreateBound(v4,v1),
                 Line.CreateBound(v5,v6), Line.CreateBound(v6,v7), Line.CreateBound(v7,v8), Line.CreateBound(v8,v5),
                 Line.CreateBound(v1,v5), Line.CreateBound(v2,v6), Line.CreateBound(v3,v7), Line.CreateBound(v4,v8)]
        ds = DirectShape.CreateElement(doc, ElementId(int(BuiltInCategory.OST_GenericModel)))
        ds.SetShape(lines)
        ds.Name = panel.get("panel_name", "PanelSolid")
        print("  [DS] Created: {0}".format(ds.Name))
        return ds
    except Exception as e:
        print("DS Fail: {0}".format(e))
        return None

def create_cutout_visualization(wall, panel, cutout_data, symbol, use_ds):
    if not use_ds and symbol:
        try:
            c_x = float(cutout_data.get("x_in",0))
            c_y = float(cutout_data.get("y_in",0))
            g_x = float(panel.get("x_in",0)) + c_x
            g_y = float(panel.get("y_in",0)) + c_y
            fake_panel = panel.copy()
            fake_panel.update({
                "panel_name": "CUT_" + str(cutout_data.get("id","")),
                "x_in": g_x, "y_in": g_y,
                "width_in": cutout_data.get("width_in",0),
                "height_in": cutout_data.get("height_in",0),
                "cutouts": []
            })
            
            # [FIX] Visual pop-out for cutouts
            place_panel_family(wall, fake_panel, symbol, extra_z_offset_in=2.0)
            return True
        except: pass
    return False

# ========== MAIN ==========
def main():
    print("--- PANEL PLACEMENT: CORE CENTER ALIGNMENT ---")
    
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
            try: cutouts = json.loads(row.get("cutouts_json","[]"))
            except: cutouts = []
            p = {
                "wall_id": norm_id(row.get("wall_id")),
                "x_in": row.get("x_in"),
                "y_in": row.get("y_in"),
                "width_in": row.get("width_in"),
                "height_in": row.get("height_in"),
                "x_ref": row.get("x_ref"),
                "panel_name": row.get("panel_name"),
                "rotation_deg": row.get("rotation_deg"),
                "cutouts": cutouts
            }
            panels.append(p)

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

    # Group by wall
    panels_map = {}
    for p in panels:
        panels_map.setdefault(p["wall_id"], []).append(p)

    t = Transaction(doc, "Place Panels")
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