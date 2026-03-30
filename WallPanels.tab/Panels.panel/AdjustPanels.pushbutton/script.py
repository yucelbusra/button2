# coding: ascii
"""
FILL GAPS - Panel Gap Closer v3
Detects gaps between adjacent panels on the same wall and closes them
by adjusting the width and/or insertion point of neighboring panels.

WORKFLOW:
  1. Select one or more panels (or none to process all walls)
  2. Choose gap-fill strategy: move / stretch selected / stretch neighbor / split
  3. Script detects all gaps between adjacent panels on each wall
  4. Adjusts neighbor panels respecting config_used.json constraints

CONSTRAINTS RESPECTED:
  - Panel min_width (24")        -- never shrink below this
  - Panel max_width (138")       -- never grow beyond this
  - dimension_increment (1")     -- snap adjusted widths to increment
  - Void jamb clearance          -- never shrink so void falls outside panel
  - Panel flip/mirror state      -- correctly handles mirrored panels
  - Oversized panels (>max_width)-- treated as fixed columns, never resized

KEY FIXES IN v3:
  1. x_in is now the TRUE VISUAL LEFT EDGE of the panel, not insertion point.
       Non-flipped: left_edge = insertion_x
       Flipped:     left_edge = insertion_x - width
     This makes gap/overlap detection accurate regardless of flip state.
  2. Duplicate panels at identical positions are deduplicated (keeps one).
  3. Panels wider than max_width are flagged fixed=True and skipped for resizing
     but still used as boundary markers for gap detection.
  4. Overlap fallback strategy correctly maps to "left"/"right" instead of
     re-using the gap strategy name verbatim.
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector, Wall, Transaction, XYZ,
    FamilyInstance, BuiltInParameter, ElementId,
    ElementTransformUtils
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

# ========== CONFIG DEFAULTS ==========
DEFAULT_MIN_WIDTH     = 24.0
DEFAULT_MAX_WIDTH     = 138.0
DEFAULT_DIM_INCREMENT = 1.0
DEFAULT_PANEL_SPACING = 0.125   # nominal gap between panels in inches
GAP_TOLERANCE_IN      = 0.25    # ignore gaps/overlaps smaller than this

# Panels wider than max_width * OVERSIZE_FACTOR are treated as fixed columns.
# Set to 1.0 to use max_width exactly.
OVERSIZE_FACTOR = 1.0


# ========== UTILITIES ==========

def _feet(val_inch):
    return float(val_inch) / 12.0

def _inches(val_feet):
    return float(val_feet) * 12.0

def _snap_up(value, increment):
    if increment <= 0:
        return value
    return math.ceil(value / increment) * increment


# ========== CONFIG LOADING ==========

def load_config():
    config_path = None
    try:
        from System.Windows.Forms import OpenFileDialog, DialogResult
        ofd = OpenFileDialog()
        ofd.Title = "Locate config_used.json (or Cancel for defaults)"
        ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        ofd.FileName = "config_used.json"
        if ofd.ShowDialog() == DialogResult.OK:
            config_path = str(ofd.FileName)
    except:
        pass

    constraints = {
        "min_width":           DEFAULT_MIN_WIDTH,
        "max_width":           DEFAULT_MAX_WIDTH,
        "dimension_increment": DEFAULT_DIM_INCREMENT,
        "panel_spacing":       DEFAULT_PANEL_SPACING,
    }

    if config_path and os.path.exists(config_path):
        try:
            with open(config_path, "r") as f:
                data = json.load(f)
            pc = data.get("panel_constraints", {})
            constraints["min_width"]           = float(pc.get("min_width",           DEFAULT_MIN_WIDTH))
            constraints["max_width"]           = float(pc.get("max_width",           DEFAULT_MAX_WIDTH))
            constraints["dimension_increment"] = float(pc.get("dimension_increment", DEFAULT_DIM_INCREMENT))
            constraints["panel_spacing"]       = float(pc.get("panel_spacing",       DEFAULT_PANEL_SPACING))
            print("[CONFIG] Loaded: {0}".format(config_path))
            print("[CONFIG] min={min_width}\" max={max_width}\" "
                  "inc={dimension_increment}\" spacing={panel_spacing}\"".format(**constraints))
        except Exception as e:
            print("[CONFIG] Load failed: {0} -- using defaults".format(e))
    else:
        print("[CONFIG] Using defaults")

    return constraints


# ========== PARAMETER HELPERS ==========

WIDTH_PARAM_CANDIDATES = ["Overall Width (default)", "Overall Width"]

def _resolve_param(inst, param_name):
    variants = [param_name]
    if param_name.endswith(" (default)"):
        variants.append(param_name[:-len(" (default)")])
    else:
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

def _get_width_in(inst):
    for cand in WIDTH_PARAM_CANDIDATES:
        p = _resolve_param(inst, cand)
        if p is not None:
            return _inches(p.AsDouble())
    return None

def _set_width_in(inst, new_width_in):
    for cand in WIDTH_PARAM_CANDIDATES:
        p = _resolve_param(inst, cand)
        if p is not None and not p.IsReadOnly:
            try:
                p.Set(_feet(new_width_in))
                return True
            except Exception as e:
                print("  [WARN] Could not set width: {0}".format(e))
    return False

def _get_void_jamb_clearance_in(inst):
    """Return max right-edge of any active void (relative to panel left edge)."""
    max_right_edge = 0.0
    for x_name, w_name, vis_name in [
        ("Void 1 X Offset", "Void 1 Width", "VOID 1"),
        ("Void 2 X Offset", "Void 2 Width", "VOID 2"),
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
        if x_p is None or w_p is None:
            continue
        right_edge = _inches(x_p.AsDouble()) + _inches(w_p.AsDouble())
        if right_edge > max_right_edge:
            max_right_edge = right_edge
    return max_right_edge


# ========== FLIP DETECTION ==========

def _is_panel_flipped(inst):
    """
    Net mirror state: HandFlipped XOR FacingFlipped.
    One flip = mirrored; two flips = back to normal.
    Falls back to inst.Mirrored if individual props unavailable.
    """
    try:
        return inst.HandFlipped != inst.FacingFlipped
    except:
        pass
    try:
        return inst.Mirrored
    except:
        return False


# ========== GEOMETRY HELPERS ==========

def _get_wall_base_elevation(wall):
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
        base_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        if base_offset:
            base_z += base_offset.AsDouble()
    except:
        pass
    return base_z

def _get_wall_direction(wall):
    """Return (vis_left XYZ, wall_dir unit vector pointing left->right visually)."""
    lc  = wall.Location.Curve
    p0  = lc.GetEndPoint(0)
    p1  = lc.GetEndPoint(1)
    up  = XYZ(0, 0, 1)
    vr  = wall.Orientation.CrossProduct(up)   # visual right direction
    if p0.DotProduct(vr) < p1.DotProduct(vr):
        vis_left, vis_right = p0, p1
    else:
        vis_left, vis_right = p1, p0
    return vis_left, (vis_right - vis_left).Normalize()

def _get_insertion_x_in(inst, vis_left, wall_dir):
    """Insertion point projected onto wall axis, in inches from vis_left."""
    try:
        return _inches((inst.Location.Point - vis_left).DotProduct(wall_dir))
    except:
        return None

def _get_left_edge_x_in(inst, vis_left, wall_dir, w_in, flipped):
    """
    TRUE VISUAL LEFT EDGE of the panel in inches from vis_left.

    The insertion point is always at the panel's internal origin:
      Non-flipped: internal origin = visual LEFT  edge -> left_edge = insertion_x
      Flipped:     internal origin = visual RIGHT edge -> left_edge = insertion_x - width

    All gap/overlap math uses this value, not raw insertion_x.
    """
    ins_x = _get_insertion_x_in(inst, vis_left, wall_dir)
    if ins_x is None:
        return None
    return (ins_x - w_in) if flipped else ins_x

def _move_panel_x(inst, delta_in, wall_dir):
    """Move panel along wall_dir by delta_in inches (positive = rightward)."""
    try:
        ElementTransformUtils.MoveElement(doc, inst.Id, wall_dir * _feet(delta_in))
        return True
    except Exception as e:
        print("  [WARN] Could not move panel: {0}".format(e))
        return False

def _panel_is_near_wall(inst, wall, vis_left, wall_dir, wall_base_z):
    try:
        loc_pt   = inst.Location.Point
        vec      = loc_pt - vis_left
        x_in     = _inches(vec.DotProduct(wall_dir))
        wall_len = _inches(wall.Location.Curve.Length)
        if x_in < -12.0 or x_in > wall_len + 12.0:
            return False
        depth_in = _inches(abs(vec.DotProduct(wall.Orientation)))
        if depth_in > _inches(wall.Width) + 12.0:
            return False
        if loc_pt.Z < (wall_base_z - 1.0):
            return False
        return True
    except:
        return False


# ========== PANEL DISCOVERY ==========

PANEL_FAMILY_NAMES = [
    "RNGD_Optimizer Ext Wall Panel_Opening",
    "RNGD_Optimizer Ext Wall Panel",
]

def _is_panel_family(inst):
    try:
        fname = inst.Symbol.Family.Name
        if any(fname == n for n in PANEL_FAMILY_NAMES):
            return True
        for cand in ["Overall Width", "Overall Width (default)"]:
            if inst.LookupParameter(cand) is not None:
                return True
        return False
    except:
        return False

def get_panels_on_wall(wall, max_width):
    """
    Return panels sorted by TRUE VISUAL LEFT EDGE (left->right).

    - Deduplicates panels at the same position + same mark (stacked duplicates).
    - Panels wider than max_width are marked fixed=True: used as boundaries
      but never resized.
    """
    vis_left, wall_dir = _get_wall_direction(wall)
    wall_base_z        = _get_wall_base_elevation(wall)

    seen_ids = set()
    raw      = []

    for inst in FilteredElementCollector(doc).OfClass(FamilyInstance):
        try:
            if inst.Id.IntegerValue in seen_ids:
                continue
            if not _is_panel_family(inst):
                continue

            on_wall = False
            try:
                host    = inst.Host
                on_wall = (host is not None and host.Id == wall.Id)
            except:
                pass
            if not on_wall:
                if not _panel_is_near_wall(inst, wall, vis_left, wall_dir, wall_base_z):
                    continue

            w_in = _get_width_in(inst)
            if w_in is None:
                continue

            flipped = _is_panel_flipped(inst)
            x_in    = _get_left_edge_x_in(inst, vis_left, wall_dir, w_in, flipped)
            if x_in is None:
                continue

            mark_p = inst.LookupParameter("Mark")
            name   = mark_p.AsString() if mark_p and mark_p.AsString() \
                     else str(inst.Id.IntegerValue)

            seen_ids.add(inst.Id.IntegerValue)
            raw.append({
                "inst":    inst,
                "x_in":   x_in,
                "w_in":   w_in,
                "id":     inst.Id.IntegerValue,
                "name":   name,
                "flipped": flipped,
                "fixed":  w_in > max_width * OVERSIZE_FACTOR,
            })
        except:
            continue

    # Sort by left edge then id for stable deduplication
    raw.sort(key=lambda p: (round(p["x_in"], 0), p["id"]))

    # Remove stacked duplicates (same mark, left edge within 0.5")
    deduped = []
    for p in raw:
        if (deduped
                and abs(p["x_in"] - deduped[-1]["x_in"]) < 0.5
                and p["name"] == deduped[-1]["name"]):
            print("  [DEDUP] '{0}' @ {1:.1f}\" -- keeping id={2}, dropping id={3}".format(
                p["name"], p["x_in"], deduped[-1]["id"], p["id"]))
            continue
        deduped.append(p)

    deduped.sort(key=lambda p: p["x_in"])

    fixed_names = [p["name"] for p in deduped if p["fixed"]]
    if fixed_names:
        print("  [FIXED] Oversized panels (boundary-only): {0}".format(fixed_names))

    return deduped, vis_left, wall_dir


# ========== GAP & OVERLAP DETECTION ==========

def find_issues(panels, spacing_in, tolerance_in):
    """
    Detect gaps and overlaps between consecutive panels using visual left edges.
    Pairs that include a fixed (oversized) panel are skipped for modification
    but the boundaries are still respected.
    """
    gaps     = []
    overlaps = []

    for i in range(len(panels) - 1):
        left  = panels[i]
        right = panels[i + 1]

        left_right_edge = left["x_in"]  + left["w_in"]
        right_left_edge = right["x_in"]
        delta           = right_left_edge - left_right_edge

        # Even if fixed, report the issue for information -- but mark it
        if delta > spacing_in + tolerance_in:
            gaps.append({
                "type":      "gap",
                "left_idx":  i,
                "right_idx": i + 1,
                "gap_start": left_right_edge,
                "gap_end":   right_left_edge,
                "gap_size":  delta - spacing_in,
                "has_fixed": left["fixed"] or right["fixed"],
            })
        elif delta < -(tolerance_in):
            overlaps.append({
                "type":         "overlap",
                "left_idx":     i,
                "right_idx":    i + 1,
                "overlap_size": abs(delta),
                "has_fixed":    left["fixed"] or right["fixed"],
            })

    return gaps, overlaps


# ========== FLIP-AWARE MOVE CALCULATION ==========

def _calc_move_for_growth(panel, grow_direction, absorb):
    """
    Calculate the insertion-point translation needed when a panel grows
    in a given visual direction.

    The insertion point is at the panel's internal origin:
      Non-flipped: insertion = visual LEFT  edge
      Flipped:     insertion = visual RIGHT edge

    grow RIGHT (visual right edge extends rightward):
      Non-flipped: right = ins + w, origin fixed          -> move =  0
      Flipped:     right = ins,     origin moves right    -> move = +absorb

    grow LEFT (visual left edge extends leftward):
      Non-flipped: left = ins, origin moves left          -> move = -absorb
      Flipped:     left = ins - w, origin fixed           -> move =  0
    """
    flipped = panel.get("flipped", False)
    if grow_direction == "right":
        return absorb if flipped else 0.0
    else:   # "left"
        return 0.0 if flipped else -absorb


# ========== GAP CLOSING ==========

def close_gap(gap, panels, strategy, constraints, vis_left, wall_dir,
              selected_panel_ids=None):
    """
    Close a single gap.  Returns list of change dicts, or [] if constrained.

    strategy: move_selected | stretch_selected | stretch_neighbor | split
    """
    max_w  = constraints["max_width"]
    inc    = constraints["dimension_increment"]
    gap_sz = gap["gap_size"]

    if gap.get("has_fixed"):
        print("  [SKIP GAP] Involves a fixed/oversized panel -- cannot resize.")
        return []

    left_p  = panels[gap["left_idx"]]
    right_p = panels[gap["right_idx"]]

    # Determine selected vs neighbor
    if selected_panel_ids:
        left_sel  = left_p["id"]  in selected_panel_ids
        right_sel = right_p["id"] in selected_panel_ids
        if left_sel and not right_sel:
            sel_p, nbr_p        = left_p,  right_p
            gap_to_right_of_sel = True    # gap is on RIGHT of selected
        else:
            sel_p, nbr_p        = right_p, left_p
            gap_to_right_of_sel = False   # gap is on LEFT  of selected
    else:
        sel_p, nbr_p            = right_p, left_p
        gap_to_right_of_sel     = False

    changes = []

    # ---- MOVE SELECTED ----
    if strategy == "move_selected":
        move_by   = _snap_up(gap_sz, inc)
        direction = move_by if gap_to_right_of_sel else -move_by
        print("  [MOVE] '{0}': {1:+.2f}\" (flip={2})".format(
            sel_p["name"], direction, sel_p["flipped"]))
        changes.append({"panel": sel_p, "old_w": sel_p["w_in"],
                         "new_w": sel_p["w_in"], "move_in": direction})
        return changes

    # ---- STRETCH SELECTED ----
    if strategy == "stretch_selected":
        absorb = _snap_up(gap_sz, inc)
        new_w  = sel_p["w_in"] + absorb
        if new_w > max_w:
            print("  [SKIP] '{0}': {1:.1f}\" > max {2}\"".format(
                sel_p["name"], new_w, max_w))
            return []
        grow_dir = "right" if gap_to_right_of_sel else "left"
        move     = _calc_move_for_growth(sel_p, grow_dir, absorb)
        print("  [STRETCH SEL] '{0}': {1:.1f}\" -> {2:.1f}\" "
              "grow={3} flip={4} move={5:+.2f}\"".format(
              sel_p["name"], sel_p["w_in"], new_w, grow_dir, sel_p["flipped"], move))
        changes.append({"panel": sel_p, "old_w": sel_p["w_in"],
                         "new_w": new_w, "move_in": move})
        return changes

    # ---- STRETCH NEIGHBOR ----
    if strategy == "stretch_neighbor":
        absorb = _snap_up(gap_sz, inc)
        new_w  = nbr_p["w_in"] + absorb
        if new_w > max_w:
            print("  [SKIP] '{0}': {1:.1f}\" > max {2}\"".format(
                nbr_p["name"], new_w, max_w))
            return []
        # Neighbor is on the opposite side of the gap from selected
        grow_dir = "left" if gap_to_right_of_sel else "right"
        move     = _calc_move_for_growth(nbr_p, grow_dir, absorb)
        print("  [STRETCH NBR] '{0}': {1:.1f}\" -> {2:.1f}\" "
              "grow={3} flip={4} move={5:+.2f}\"".format(
              nbr_p["name"], nbr_p["w_in"], new_w, grow_dir, nbr_p["flipped"], move))
        changes.append({"panel": nbr_p, "old_w": nbr_p["w_in"],
                         "new_w": new_w, "move_in": move})
        return changes

    # ---- SPLIT ----
    if strategy == "split":
        absorb_r = _snap_up(gap_sz / 2.0, inc)
        absorb_l = _snap_up(gap_sz / 2.0, inc)
        new_w_r  = right_p["w_in"] + absorb_r
        new_w_l  = left_p["w_in"]  + absorb_l
        if new_w_r > max_w:
            print("  [SKIP SPLIT] Right '{0}': {1:.1f}\" > max".format(
                right_p["name"], new_w_r))
            return []
        if new_w_l > max_w:
            print("  [SKIP SPLIT] Left '{0}': {1:.1f}\" > max".format(
                left_p["name"], new_w_l))
            return []
        move_r = _calc_move_for_growth(right_p, "left",  absorb_r)
        move_l = _calc_move_for_growth(left_p,  "right", absorb_l)
        changes.append({"panel": right_p, "old_w": right_p["w_in"],
                         "new_w": new_w_r, "move_in": move_r})
        changes.append({"panel": left_p,  "old_w": left_p["w_in"],
                         "new_w": new_w_l, "move_in": move_l})
        return changes

    return changes


# ========== OVERLAP RESOLUTION ==========

def resolve_overlap(overlap, panels, strategy, constraints, vis_left, wall_dir):
    """
    Resolve an overlap by shrinking the relevant panel(s).
    strategy: right | left | split
    Shrinking is the inverse of growing -- uses same flip-aware logic inverted.
    """
    if overlap.get("has_fixed"):
        print("  [SKIP OVLP] Involves a fixed/oversized panel -- cannot resize.")
        return []

    min_w = constraints["min_width"]
    inc   = constraints["dimension_increment"]
    ovlp  = overlap["overlap_size"]

    left_p  = panels[overlap["left_idx"]]
    right_p = panels[overlap["right_idx"]]

    changes = []

    if strategy in ("right", "split"):
        shrink = _snap_up(ovlp if strategy == "right" else ovlp / 2.0, inc)
        new_w  = right_p["w_in"] - shrink
        if new_w < min_w:
            print("  [SKIP] '{0}': shrink to {1:.1f}\" < min {2}\"".format(
                right_p["name"], new_w, min_w))
            return []
        # Right panel shrinks from its LEFT visual edge (inverse of grow-left):
        #   Non-flipped grow-left: move=-absorb -> shrink: move=+shrink
        #   Flipped     grow-left: move=0       -> shrink: move=0
        move = 0.0 if right_p["flipped"] else shrink
        changes.append({"panel": right_p, "old_w": right_p["w_in"],
                         "new_w": new_w, "move_in": move})

    if strategy in ("left", "split"):
        shrink = _snap_up(ovlp if strategy == "left" else ovlp / 2.0, inc)
        new_w  = left_p["w_in"] - shrink
        if new_w < min_w:
            print("  [SKIP] '{0}': shrink to {1:.1f}\" < min {2}\"".format(
                left_p["name"], new_w, min_w))
            return []
        void_edge = _get_void_jamb_clearance_in(left_p["inst"])
        if void_edge > new_w:
            print("  [SKIP] '{0}': void right={1:.1f}\" > new w={2:.1f}\"".format(
                left_p["name"], void_edge, new_w))
            return []
        # Left panel shrinks from its RIGHT visual edge (inverse of grow-right):
        #   Non-flipped grow-right: move=0       -> shrink: move=0
        #   Flipped     grow-right: move=+absorb -> shrink: move=-shrink
        move = -shrink if left_p["flipped"] else 0.0
        changes.append({"panel": left_p, "old_w": left_p["w_in"],
                         "new_w": new_w, "move_in": move})

    return changes


# ========== APPLY CHANGES ==========

def apply_changes(changes, vis_left, wall_dir):
    applied = 0
    for ch in changes:
        panel = ch["panel"]
        inst  = panel["inst"]

        if panel.get("fixed", False):
            print("  [SKIP FIXED] '{0}' is oversized -- not resizing".format(panel["name"]))
            continue

        if abs(ch["move_in"]) > 0.001:
            if not _move_panel_x(inst, ch["move_in"], wall_dir):
                print("  [ERROR] Could not move '{0}'".format(panel["name"]))
                continue

        if _set_width_in(inst, ch["new_w"]):
            print("  [OK] '{0}': {1:.2f}\" -> {2:.2f}\" "
                  "(move={3:+.2f}\" flip={4})".format(
                  panel["name"], ch["old_w"], ch["new_w"],
                  ch["move_in"], panel["flipped"]))
            applied += 1
        else:
            print("  [ERROR] Could not set width on '{0}'".format(panel["name"]))

    return applied


# ========== SELECTION ==========

def get_selection_info(max_width):
    sel_ids = uidoc.Selection.GetElementIds()
    walls   = {}
    selected_panel_ids = set()
    mode    = "all"

    for eid in sel_ids:
        elem = doc.GetElement(eid)
        if isinstance(elem, Wall):
            walls[elem.Id.IntegerValue] = elem
            mode = "walls"
        elif isinstance(elem, FamilyInstance):
            if _is_panel_family(elem):
                selected_panel_ids.add(eid.IntegerValue)
                mode = "panels"
            try:
                host = elem.Host
                if isinstance(host, Wall):
                    walls[host.Id.IntegerValue] = host
            except:
                pass

    if not walls and not selected_panel_ids:
        res = forms.alert(
            "No panels or walls selected.\nProcess ALL walls in the model?",
            title="Fill Gaps", yes=True, no=True)
        if not res:
            return {}, set(), "cancelled"
        for w in FilteredElementCollector(doc).OfClass(Wall).ToElements():
            walls[w.Id.IntegerValue] = w
        mode = "all"

    # Non-hosted panels: find wall by proximity
    if selected_panel_ids and not walls:
        for w in FilteredElementCollector(doc).OfClass(Wall).ToElements():
            vis_left, wall_dir = _get_wall_direction(w)
            base_z = _get_wall_base_elevation(w)
            for eid_int in selected_panel_ids:
                inst = doc.GetElement(ElementId(eid_int))
                if inst and _panel_is_near_wall(inst, w, vis_left, wall_dir, base_z):
                    walls[w.Id.IntegerValue] = w
                    break

    return walls, selected_panel_ids, mode


# ========== DIAGNOSTICS ==========

def run_diagnostics(max_width):
    print("--- DIAGNOSTICS v3 ---")
    all_inst = list(FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements())
    print("Total FamilyInstances: {0}".format(len(all_inst)))

    wall_hosted = 0
    detected    = 0

    for inst in all_inst:
        try:
            fname = "?"
            try: fname = inst.Symbol.Family.Name
            except: pass

            in_list   = any(fname == n for n in PANEL_FAMILY_NAMES)
            has_width = (inst.LookupParameter("Overall Width") is not None or
                         inst.LookupParameter("Overall Width (default)") is not None)
            if not in_list and not has_width:
                continue

            detected += 1

            try:
                host      = inst.Host
                host_info = "wall {0}".format(host.Id.IntegerValue) if host else "level-hosted"
                if host: wall_hosted += 1
            except:
                host_info = "unknown"

            w_in    = _get_width_in(inst)
            flipped = _is_panel_flipped(inst)
            mark_p  = inst.LookupParameter("Mark")
            mark    = mark_p.AsString() if mark_p and mark_p.AsString() \
                      else str(inst.Id.IntegerValue)
            fixed   = w_in is not None and w_in > max_width * OVERSIZE_FACTOR

            try:
                pt       = inst.Location.Point
                ins_info = "ins=({0:.1f}\",{1:.1f}\")".format(pt.X * 12, pt.Y * 12)
            except:
                ins_info = "no-loc"

            print("  [{0}] '{1}' | {2} | {3} | w={4} | flip={5} | fixed={6}".format(
                mark, fname, host_info, ins_info,
                "{0:.1f}\"".format(w_in) if w_in else "?",
                flipped, fixed))
        except Exception as e:
            print("  [ERR] {0}".format(e))

    print("Wall-hosted: {0}  Detected: {1}".format(wall_hosted, detected))
    print("--- END DIAGNOSTICS ---")


# ========== MAIN ==========

def main():
    print("--- FILL GAPS v3 ---")

    from pyrevit import forms as _forms
    if _forms.alert(
            "Run diagnostics first?\n"
            "(Shows left-edge positions, flip state, fixed panels)",
            title="Fill Gaps", yes=True, no=True, warn_icon=False):
        c = load_config()
        run_diagnostics(c["max_width"])
        return

    constraints = load_config()
    max_w = constraints["max_width"]

    strategy_ops = [
        "MOVE selected panel -- slides to close gap, width unchanged",
        "STRETCH selected panel -- selected panel grows to close gap",
        "STRETCH neighbor -- the other panel grows to close gap",
        "SPLIT -- both panels grow toward each other",
    ]
    res = forms.SelectFromList.show(
        strategy_ops,
        title="Fill Gaps -- Strategy",
        button_name="Apply",
        multiselect=False)
    if not res:
        print("Cancelled.")
        return

    if   res.startswith("MOVE"):             strategy = "move_selected"
    elif res.startswith("STRETCH selected"): strategy = "stretch_selected"
    elif res.startswith("STRETCH neighbor"): strategy = "stretch_neighbor"
    else:                                    strategy = "split"

    print("[STRATEGY] {0}".format(strategy))

    walls, selected_panel_ids, mode = get_selection_info(max_w)
    if mode == "cancelled" or not walls:
        print("No walls to process.")
        return

    print("[WALLS] {0} wall(s) | mode={1}".format(len(walls), mode))

    t = Transaction(doc, "Fill Panel Gaps")
    t.Start()

    total_gaps = total_overlaps = total_closed = total_skipped = 0

    for wid, wall in walls.items():
        print("\n--- Wall {0} ---".format(wid))

        panels, vis_left, wall_dir = get_panels_on_wall(wall, max_w)
        if len(panels) < 2:
            print("  < 2 panels -- skipping.")
            continue

        print("  {0} panels (after dedup):".format(len(panels)))
        for p in panels:
            print("    '{0}' left={1:.1f}\" w={2:.1f}\" flip={3} fixed={4}".format(
                p["name"], p["x_in"], p["w_in"], p["flipped"], p["fixed"]))

        gaps, overlaps = find_issues(panels, constraints["panel_spacing"], GAP_TOLERANCE_IN)

        if mode == "panels" and selected_panel_ids:
            sel_idx  = set(i for i, p in enumerate(panels) if p["id"] in selected_panel_ids)
            gaps     = [g for g in gaps
                        if g["left_idx"] in sel_idx or g["right_idx"] in sel_idx]
            overlaps = [o for o in overlaps
                        if o["left_idx"] in sel_idx or o["right_idx"] in sel_idx]
            print("  Scope: pairs involving {0}".format(
                [panels[i]["name"] for i in sorted(sel_idx)]))

        if not gaps and not overlaps:
            print("  No issues found.")
            continue

        for g in gaps:
            flag = " [FIXED-skip]" if g.get("has_fixed") else ""
            print("  GAP  '{0}'..'{1}': {2:.3f}\" excess{3}".format(
                panels[g["left_idx"]]["name"],
                panels[g["right_idx"]]["name"],
                g["gap_size"], flag))
        for o in overlaps:
            flag = " [FIXED-skip]" if o.get("has_fixed") else ""
            print("  OVLP '{0}'..'{1}': {2:.3f}\"{3}".format(
                panels[o["left_idx"]]["name"],
                panels[o["right_idx"]]["name"],
                o["overlap_size"], flag))

        total_gaps     += len(gaps)
        total_overlaps += len(overlaps)

        # ---- Process gaps ----
        for g in gaps:
            changes = close_gap(g, panels, strategy, constraints,
                                vis_left, wall_dir, selected_panel_ids)
            if not changes and strategy not in ("split", "stretch_neighbor"):
                print("  [FALLBACK] trying stretch_neighbor")
                changes = close_gap(g, panels, "stretch_neighbor", constraints,
                                    vis_left, wall_dir, selected_panel_ids)
            if not changes:
                total_skipped += 1
                continue
            applied = apply_changes(changes, vis_left, wall_dir)
            if applied:
                total_closed += 1
                for ch in changes:
                    ch["panel"]["w_in"]  = ch["new_w"]
                    ch["panel"]["x_in"] += ch["move_in"]
            else:
                total_skipped += 1

        # ---- Process overlaps ----
        for o in overlaps:
            # Map gap strategies to overlap strategies
            if strategy in ("stretch_selected", "move_selected"):
                ovlp_strategy = "right"   # shrink the panel to the right
            elif strategy == "stretch_neighbor":
                ovlp_strategy = "left"
            elif strategy == "split":
                ovlp_strategy = "split"
            else:
                ovlp_strategy = "right"

            changes = resolve_overlap(o, panels, ovlp_strategy, constraints,
                                      vis_left, wall_dir)
            if not changes:
                fallback = "left" if ovlp_strategy == "right" else "right"
                print("  [FALLBACK] trying {0}".format(fallback))
                changes = resolve_overlap(o, panels, fallback, constraints,
                                          vis_left, wall_dir)
            if not changes:
                total_skipped += 1
                continue
            applied = apply_changes(changes, vis_left, wall_dir)
            if applied:
                total_closed += 1
                for ch in changes:
                    ch["panel"]["w_in"]  = ch["new_w"]
                    ch["panel"]["x_in"] += ch["move_in"]
            else:
                total_skipped += 1

    t.Commit()

    print("\n--- DONE ---")
    print("Gaps: {0}  Overlaps: {1}".format(total_gaps, total_overlaps))
    print("Resolved: {0}  Skipped: {1}".format(total_closed, total_skipped))


if __name__ == "__main__":
    main()