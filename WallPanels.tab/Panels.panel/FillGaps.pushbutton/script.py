# coding: ascii
"""
FILL GAPS - Panel Gap Closer (v5.3 - Fixed viewId error)
========================================================
Detects gaps and overlaps between adjacent panels on the same wall
by projecting their physical Bounding Box footprint.
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector, Wall, Transaction, XYZ, Line,
    FamilySymbol, FamilyInstance, BuiltInCategory, BuiltInParameter,
    ElementId, ElementTransformUtils, HostObjectUtils, ShellLayerType,
    PlanarFace
)

try:
    from Autodesk.Revit.DB.Structure import StructuralType
except:
    from Autodesk.Revit.DB import Structure
    StructuralType = Structure.StructuralType

from pyrevit import revit, forms
import os
import json
import math

doc   = revit.doc
uidoc = revit.uidoc

# ========== CONFIG ==========
DEFAULT_MIN_WIDTH     = 24.0
DEFAULT_MAX_WIDTH     = 138.0
DEFAULT_DIM_INCREMENT = 1.0
DEFAULT_PANEL_SPACING = 0.125
GAP_TOLERANCE_IN      = 0.25

PANEL_FAMILY_NAMES = [
    "RNGD_Optimizer Ext Wall Panel_Opening",
    "RNGD_Optimizer Ext Wall Panel",
]

# ========== UTILITIES ==========

def _feet(v): return float(v) / 12.0
def _inches(v): return float(v) * 12.0
def _safe_float(v, d=0.0):
    try: return float(v)
    except: return d

def _snap_up(value, inc):
    if inc <= 0:
        return value
    return math.ceil(value / inc) * inc

def _is_panel_family(inst):
    try:
        if not isinstance(inst, FamilyInstance):
            return False
        fam_name = inst.Symbol.Family.Name if inst.Symbol and inst.Symbol.Family else ""
        sym_name = inst.Symbol.Name if inst.Symbol else ""
        return fam_name in PANEL_FAMILY_NAMES or sym_name in PANEL_FAMILY_NAMES
    except:
        return False

# ========== GEOMETRY ENGINE ==========

def _get_panel_geom_info(inst, vis_left, wall_dir):
    """Calculates true X and Width by projecting the physical Bounding Box."""
    bbox = inst.get_BoundingBox(None)
    if not bbox:
        return None, None

    pts = [
        bbox.Min,
        bbox.Max,
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z)
    ]

    projections = [(pt - vis_left).DotProduct(wall_dir) for pt in pts]
    true_left_ft = min(projections)
    true_right_ft = max(projections)

    return _inches(true_left_ft), _inches(true_right_ft - true_left_ft)

def get_panels_on_wall(wall):
    vis_left, wall_dir = _get_wall_direction(wall)
    wall_base_z = _get_wall_base_elevation(wall)
    panels = []

    # FIX:
    # Do NOT use FilteredElementCollector(doc, wall.Id)
    # The second argument must be a VIEW id, not a wall id.
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)

    for inst in collector:
        if not _is_panel_family(inst):
            continue

        try:
            if not inst.Host or inst.Host.Id.IntegerValue != wall.Id.IntegerValue:
                continue
        except:
            continue

        if not _panel_is_near_wall(inst, wall, vis_left, wall_dir, wall_base_z):
            continue

        x_in, w_in = _get_panel_geom_info(inst, vis_left, wall_dir)
        if x_in is None or w_in is None:
            continue

        mark_p = inst.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        name = mark_p.AsString() if mark_p and mark_p.HasValue else str(inst.Id.IntegerValue)

        panels.append({
            "inst": inst,
            "x_in": x_in,
            "w_in": w_in,
            "id": inst.Id.IntegerValue,
            "name": name
        })

    panels.sort(key=lambda p: p["x_in"])
    return panels, vis_left, wall_dir

# ========== CONFIG LOADING ==========

def load_config():
    constraints = {
        "min_width": DEFAULT_MIN_WIDTH,
        "max_width": DEFAULT_MAX_WIDTH,
        "dimension_increment": DEFAULT_DIM_INCREMENT,
        "panel_spacing": DEFAULT_PANEL_SPACING,
    }
    try:
        from System.Windows.Forms import OpenFileDialog, DialogResult
        ofd = OpenFileDialog()
        ofd.Title = "Locate config_used.json (or Cancel for defaults)"
        ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        ofd.FileName = "config_used.json"
        result = ofd.ShowDialog()
        if result == DialogResult.OK:
            config_path = str(ofd.FileName)
            with open(config_path, "r") as f:
                data = json.load(f)
            pc = data.get("panel_constraints", {})
            constraints["min_width"]           = float(pc.get("min_width", DEFAULT_MIN_WIDTH))
            constraints["max_width"]           = float(pc.get("max_width", DEFAULT_MAX_WIDTH))
            constraints["dimension_increment"] = float(pc.get("dimension_increment", DEFAULT_DIM_INCREMENT))
            constraints["panel_spacing"]       = float(pc.get("panel_spacing", DEFAULT_PANEL_SPACING))
            print("[CONFIG] Loaded: {0}".format(config_path))
            print("[CONFIG] min_width={min_width}\" max_width={max_width}\" "
                  "increment={dimension_increment}\" spacing={panel_spacing}\"".format(**constraints))
    except Exception as e:
        print("[CONFIG] Using defaults ({0})".format(e))
    return constraints

# ========== CSV BASELINE ==========

CSV_FILENAME = "optimized_panel_placement.csv"

def load_csv_baseline():
    try:
        from System.Windows.Forms import OpenFileDialog, DialogResult
        ofd = OpenFileDialog()
        ofd.Title = "Locate optimized_panel_placement.csv"
        ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        ofd.FileName = CSV_FILENAME
        result = ofd.ShowDialog()
        if result != DialogResult.OK:
            return None
        csv_path = str(ofd.FileName)
    except:
        return None

    import csv
    baseline = {}
    try:
        with open(csv_path, 'r') as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row.get("panel_name", "").strip()
                if not name:
                    continue
                baseline[name] = {
                    "width_in": _safe_float(row.get("width_in", 0)),
                    "x_in":     _safe_float(row.get("x_in", 0)),
                    "wall_id":  str(row.get("wall_id", "")).strip(),
                }
        print("[CSV] Loaded baseline for {0} panels".format(len(baseline)))
    except Exception as e:
        print("[CSV] Failed to read CSV: {0}".format(e))
        return None

    return baseline

def detect_changed_panel(panels, baseline, tolerance_in=0.5):
    changed = []
    for p in panels:
        name = p["name"]
        if name not in baseline:
            continue
        orig_w = baseline[name]["width_in"]
        curr_w = p["w_in"]
        delta = abs(curr_w - orig_w)
        if delta > tolerance_in:
            changed.append((delta, p, orig_w))
    if not changed:
        return None, None
    changed.sort(key=lambda x: x[0], reverse=True)
    _, panel, orig_w = changed[0]
    return panel, orig_w

# ========== PARAMETER HELPERS ==========

def _resolve_param(inst, param_name):
    variants = [param_name, param_name + " (default)"]
    if " (default)" in param_name:
        variants.append(param_name.replace(" (default)", ""))
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

def _get_width_in(inst):
    for cand in ["Overall Width", "Overall Width (default)"]:
        p = _resolve_param(inst, cand)
        if p is not None:
            return _inches(p.AsDouble())
    return None

def _set_width_in(inst, new_width_in):
    for cand in ["Overall Width", "Overall Width (default)"]:
        p = _resolve_param(inst, cand)
        if p is not None and not p.IsReadOnly:
            try:
                p.Set(_feet(new_width_in))
                return True
            except Exception as e:
                print("  [WARN] Could not set width: {0}".format(e))
    return False

def _get_void_right_edge_in(inst):
    max_right = 0.0
    for x_name, w_name, vis_name in [
        ("Void 1 X Offset", "Void 1 Width", "VOID 1"),
        ("Void 2 X Offset", "Void 2 Width", "VOID 2")
    ]:
        vis_p = _resolve_param(inst, vis_name)
        if vis_p is None:
            continue
        try:
            if vis_p.AsInteger() == 0:
                continue
        except:
            pass
        x_p = _resolve_param(inst, x_name)
        w_p = _resolve_param(inst, w_name)
        if x_p and w_p:
            right = _inches(x_p.AsDouble()) + _inches(w_p.AsDouble())
            if right > max_right:
                max_right = right
    return max_right

# ========== GEOMETRY HELPERS ==========

def _get_wall_base_elevation(wall):
    base_z = 0.0
    try:
        lvl_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        level = doc.GetElement(lvl_id)
        if level:
            base_z = level.Elevation
    except:
        pass
    try:
        base_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        if base_offset:
            base_z += base_offset.AsDouble()
    except:
        pass
    return base_z

def _get_wall_direction(wall):
    lc = wall.Location.Curve
    p0, p1 = lc.GetEndPoint(0), lc.GetEndPoint(1)
    vrd = wall.Orientation.CrossProduct(XYZ(0, 0, 1))
    if p0.DotProduct(vrd) < p1.DotProduct(vrd):
        vis_left, vis_right = p0, p1
    else:
        vis_left, vis_right = p1, p0
    return vis_left, (vis_right - vis_left).Normalize()

def _move_panel_x(inst, delta_in, wall_dir):
    try:
        ElementTransformUtils.MoveElement(doc, inst.Id, wall_dir * _feet(delta_in))
        return True
    except:
        return False

def _panel_is_near_wall(inst, wall, vis_left, wall_dir, wall_base_z):
    try:
        loc_pt = inst.Location.Point
        vec = loc_pt - vis_left
        x_in = _inches(vec.DotProduct(wall_dir))
        if x_in < -12.0 or x_in > _inches(wall.Location.Curve.Length) + 12.0:
            return False
        if _inches(abs(vec.DotProduct(wall.Orientation))) > _inches(wall.Width) + 12.0:
            return False
        if loc_pt.Z < (wall_base_z - 1.0):
            return False
        return True
    except:
        return False

# ========== LOGIC & MAIN ==========

def find_issues(panels, spacing_in, tolerance_in):
    gaps, overlaps = [], []
    for i in range(len(panels) - 1):
        left, right = panels[i], panels[i + 1]
        delta = right["x_in"] - (left["x_in"] + left["w_in"])
        if delta > spacing_in + tolerance_in:
            gaps.append({"type": "gap", "left_idx": i, "right_idx": i + 1, "gap_size": delta - spacing_in})
        elif delta < -(tolerance_in):
            overlaps.append({"type": "overlap", "left_idx": i, "right_idx": i + 1, "overlap_size": abs(delta)})
    return gaps, overlaps

def apply_move_group(panels, gap_idx, gap_size, direction, wall_dir):
    move_by = _snap_up(gap_size, 1.0)
    group = panels[gap_idx + 1:] if direction == "right_group_left" else panels[:gap_idx + 1]
    delta = -move_by if direction == "right_group_left" else move_by
    for p in group:
        if _move_panel_x(p["inst"], delta, wall_dir):
            p["x_in"] += delta
    return len(group)

def close_gap(gap, panels, strategy, constraints, wall_dir, selected_panel_ids=None):
    inc, max_w, gap_sz = constraints["dimension_increment"], constraints["max_width"], gap["gap_size"]
    lp, rp = panels[gap["left_idx"]], panels[gap["right_idx"]]

    selected_ids = selected_panel_ids or set()
    is_left_sel = lp["id"] in selected_ids

    if strategy == "move_group":
        direction = "right_group_left" if is_left_sel else "left_group_right"
        return apply_move_group(panels, gap["left_idx"], gap_sz, direction, wall_dir) > 0

    changes = []
    absorb = _snap_up(gap_sz, inc)
    target = lp if is_left_sel or strategy == "stretch_neighbor" else rp
    new_w = target["w_in"] + absorb
    if new_w <= max_w:
        move = 0.0 if target == lp else -absorb
        changes.append({"panel": target, "old_w": target["w_in"], "new_w": new_w, "move_in": move})
    return changes

def resolve_overlap(overlap, panels, strategy, constraints, wall_dir, selected_panel_ids=None):
    min_w, inc, ovlp = constraints["min_width"], constraints["dimension_increment"], overlap["overlap_size"]
    lp, rp = panels[overlap["left_idx"]], panels[overlap["right_idx"]]

    selected_ids = selected_panel_ids or set()
    is_left_sel = lp["id"] in selected_ids

    if strategy == "move_group":
        direction = "left_group_right" if is_left_sel else "right_group_left"
        return apply_move_group(panels, overlap["left_idx"], ovlp, direction, wall_dir) > 0

    changes = []
    shrink = _snap_up(ovlp, inc)
    target = lp if is_left_sel or strategy == "stretch_neighbor" else rp
    new_w = target["w_in"] - shrink
    if new_w >= min_w:
        move = shrink if target == rp else 0.0
        changes.append({"panel": target, "old_w": target["w_in"], "new_w": new_w, "move_in": move})
    return changes

def apply_changes(changes, wall_dir):
    applied = 0
    for ch in changes:
        p = ch["panel"]
        if abs(ch["move_in"]) > 0.001:
            _move_panel_x(p["inst"], ch["move_in"], wall_dir)
        if abs(ch["new_w"] - ch["old_w"]) > 0.001:
            _set_width_in(p["inst"], ch["new_w"])
        p["w_in"], p["x_in"] = ch["new_w"], p["x_in"] + ch["move_in"]
        applied += 1
    return applied

def get_selection_info():
    sel_ids = uidoc.Selection.GetElementIds()
    walls, selected_panel_ids = {}, set()
    for eid in sel_ids:
        elem = doc.GetElement(eid)
        if isinstance(elem, Wall):
            walls[eid.IntegerValue] = elem
        elif isinstance(elem, FamilyInstance) and _is_panel_family(elem):
            selected_panel_ids.add(eid.IntegerValue)
            if elem.Host:
                walls[elem.Host.Id.IntegerValue] = elem.Host
    mode = "panels" if selected_panel_ids else "walls" if walls else "all"
    if mode == "all":
        res = forms.alert("Process ALL walls?", yes=True, no=True)
        if not res:
            return {}, set(), "cancelled"
        for w in FilteredElementCollector(doc).OfClass(Wall).ToElements():
            walls[w.Id.IntegerValue] = w
    return walls, selected_panel_ids, mode

def _run_gap_fix(walls, selected_panel_ids, strategy, constraints):
    total_gaps = total_closed = 0
    for wid, wall in walls.items():
        panels, vis_left, wall_dir = get_panels_on_wall(wall)
        if len(panels) < 2:
            continue

        gaps, overlaps = find_issues(panels, constraints["panel_spacing"], GAP_TOLERANCE_IN)

        if selected_panel_ids:
            gaps = [g for g in gaps if panels[g["left_idx"]]["id"] in selected_panel_ids or panels[g["right_idx"]]["id"] in selected_panel_ids]
            overlaps = [o for o in overlaps if panels[o["left_idx"]]["id"] in selected_panel_ids or panels[o["right_idx"]]["id"] in selected_panel_ids]

        for g in gaps:
            res = close_gap(g, panels, strategy, constraints, wall_dir, selected_panel_ids)
            if isinstance(res, list) and res:
                total_closed += 1 if apply_changes(res, wall_dir) else 0
            elif res is True:
                total_closed += 1

        for o in overlaps:
            res = resolve_overlap(o, panels, strategy, constraints, wall_dir, selected_panel_ids)
            if isinstance(res, list) and res:
                total_closed += 1 if apply_changes(res, wall_dir) else 0
            elif res is True:
                total_closed += 1

        total_gaps += len(gaps) + len(overlaps)

    return total_gaps, total_closed

def main():
    print("--- FILL GAPS ---")
    constraints = load_config()
    walls, selected_panel_ids, mode = get_selection_info()

    if mode == "cancelled":
        return

    if mode == "panels":
        res = forms.SelectFromList.show(
            ["MOVE GROUP", "STRETCH selected", "STRETCH neighbor", "split"],
            title="Strategy",
            multiselect=False
        )
        if not res:
            return
        strategy = res.lower().replace(" ", "_")
    else:
        baseline = load_csv_baseline()
        if not baseline:
            return
        strategy = "move_group"

    t = Transaction(doc, "Fill Panel Gaps")
    try:
        t.Start()
        g, c = _run_gap_fix(walls, selected_panel_ids, strategy, constraints)
        t.Commit()
        print("Done. Issues: {0}, Closed: {1}".format(g, c))
    except Exception as e:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except:
            pass
        print("Error occurred, transaction rolled back: {0}".format(e))
        raise

if __name__ == "__main__":
    main()