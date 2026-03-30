# -*- coding: utf-8 -*-
"""
Places largest possible wall panels on walls defined in walls.csv, 
considering openings from wall_openings.csv.
"""


from __future__ import print_function
import os
from datetime import datetime

# --- .NET UI imports ---
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (
    Application, FolderBrowserDialog, DialogResult, Form,
    Label, RadioButton, Button, FormBorderStyle, FormStartPosition,
    MessageBox, MessageBoxButtons, MessageBoxIcon, TextBox, IWin32Window,
    CheckBox, GroupBox, Panel, FlowLayoutPanel,
    BorderStyle, DockStyle, AnchorStyles, Padding
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# --- Additional refs to bind dialogs to Revit main window ---
clr.AddReference('System')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from System import IntPtr
from Autodesk.Revit.UI import UIApplication
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.DB import (
    XYZ, Line, Transaction, DirectShape, ElementId,
    BoundingBoxXYZ, SolidOptions
)
from Autodesk.Revit.DB import BuiltInCategory
try:
    from Autodesk.Revit.DB import TessellatedShapeBuilder, TessellatedFace
    from Autodesk.Revit.DB import TessellatedShapeBuilderResult
    from Autodesk.Revit.DB import TessellatedShapeBuilderTarget
    from Autodesk.Revit.DB import TessellatedShapeBuilderFallback
    _TESS_AVAILABLE = True
except Exception:
    _TESS_AVAILABLE = False
try:
    from pyrevit import revit as _pyrevit_revit
    _doc = _pyrevit_revit.doc
except Exception:
    _doc = None
try:
    _uidoc = _pyrevit_revit.uidoc
except Exception:
    _uidoc = None

# --- Import optimizer module sitting next to this script ---
try:
    import panel_calculator as opt
except Exception as e:
    raise Exception("Failed to import panel_calculator.py: {0}".format(e))

# =========================
# Helpers: Revit window owner
# =========================
class WindowWrapper(IWin32Window):
    def __init__(self, handle):
        self._hwnd = handle
    @property
    def Handle(self):
        return self._hwnd

def get_revit_owner():
    """Return an IWin32Window wrapper of Revit's main window; None on failure."""
    try:
        # In pyRevit, __revit__ is already a UIApplication
        uiapp = __revit__
        hwnd = uiapp.MainWindowHandle
        return WindowWrapper(IntPtr(hwnd))
    except Exception as e:
        print("Warning: could not retrieve Revit main window handle: {0}".format(e))
        return None

# =========================
# UI: Folder Picker (click)
# =========================

def pick_data_folder():
    """Show a FolderBrowserDialog; default to Desktop, fallback to Home; return selected path or script dir."""
    owner = get_revit_owner()
    dialog = FolderBrowserDialog()
    dialog.Description = "Select folder containing walls.csv and wall_openings.csv"
    # Default: Desktop; fallback: Home
    initial_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(initial_dir):
        initial_dir = os.path.expanduser("~")
    dialog.SelectedPath = initial_dir
    result = dialog.ShowDialog(owner) if owner else dialog.ShowDialog()
    if result == DialogResult.OK and dialog.SelectedPath and os.path.isdir(dialog.SelectedPath):
        return dialog.SelectedPath
    else:
        # If user cancels, fall back to the folder where this script lives
        return os.path.dirname(os.path.abspath(__file__))

# =====================================
# UI: Orientation Picker (radio buttons)
# =====================================

class OrientationDialog(Form):
    def __init__(self):
        # Basic form setup
        self.Text = "Select Panel Orientation"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(600, 320)
        self.TopMost = True

        # Label
        lbl = Label()
        lbl.Text = "Select panel orientation:"
        lbl.Location = Point(12, 12)
        lbl.AutoSize = True
        self.Controls.Add(lbl)

        # Radio buttons (default: Vertical)
        self.rbVertical = RadioButton()
        self.rbVertical.Text = "Vertical"
        self.rbVertical.Checked = True
        self.rbVertical.Location = Point(24, 40)
        self.Controls.Add(self.rbVertical)

        self.rbHorizontal = RadioButton()
        self.rbHorizontal.Text = "Horizontal"
        self.rbHorizontal.Location = Point(24, 66)
        self.Controls.Add(self.rbHorizontal)

        # OK / Cancel buttons
        btnOK = Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = Point(180, 130)

        btnCancel = Button()
        btnCancel.Text = "Cancel"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = Point(260, 130)

        self.AcceptButton = btnOK
        self.CancelButton = btnCancel
        self.Controls.Add(btnOK)
        self.Controls.Add(btnCancel)


def pick_orientation():
    """Show OrientationDialog and return 'vertical' or 'horizontal'; None if canceled."""
    owner = get_revit_owner()
    dlg = OrientationDialog()
    result = dlg.ShowDialog(owner) if owner else dlg.ShowDialog()
    if result == DialogResult.OK:
        if dlg.rbVertical.Checked:
            return "vertical"
        else:
            return "horizontal"
    else:
        return None

# =====================================
# NEW: UI: Panel Type Picker
# =====================================

class PanelTypeDialog(Form):
    def __init__(self):
        # Basic form setup
        self.Text = "Select Panel Type"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(600, 320)
        self.TopMost = True

        # Label
        lbl = Label()
        lbl.Text = "Select panel type:"
        lbl.Location = Point(12, 12)
        lbl.AutoSize = True
        self.Controls.Add(lbl)

        # Radio buttons (default: Backer)
        self.rbBacker = RadioButton()
        self.rbBacker.Text = "Backer Panels (1/8\" spacing)"
        self.rbBacker.Checked = True
        self.rbBacker.Location = Point(24, 40)
        self.rbBacker.AutoSize = True
        self.Controls.Add(self.rbBacker)

        self.rbFullyFinished = RadioButton()
        self.rbFullyFinished.Text = "Fully Finished Panels (3/4\" spacing)"
        self.rbFullyFinished.Location = Point(24, 66)
        self.rbFullyFinished.AutoSize = True
        self.Controls.Add(self.rbFullyFinished)

        # OK / Cancel buttons
        btnOK = Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = Point(180, 130)

        btnCancel = Button()
        btnCancel.Text = "Cancel"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = Point(260, 130)

        self.AcceptButton = btnOK
        self.CancelButton = btnCancel
        self.Controls.Add(btnOK)
        self.Controls.Add(btnCancel)


def pick_panel_type():
    """Show PanelTypeDialog and return spacing value (0.125 or 0.75); None if canceled."""
    owner = get_revit_owner()
    dlg = PanelTypeDialog()
    result = dlg.ShowDialog(owner) if owner else dlg.ShowDialog()
    if result == DialogResult.OK:
        if dlg.rbBacker.Checked:
            return 0.125  # 1/8" for backer panels
        else:
            return 0.75   # 3/4" for fully finished panels
    else:
        return None

# =====================================
# UI: Project Name Input Dialog
# =====================================

class ProjectNameDialog(Form):
    def __init__(self):
        # Basic form setup
        self.Text = "Project Name"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(600, 320)
        self.TopMost = True

        # Label
        lbl = Label()
        lbl.Text = "Enter project name for output folder:"
        lbl.Location = Point(12, 12)
        lbl.AutoSize = True
        self.Controls.Add(lbl)

        # TextBox for project name
        self.txtProjectName = TextBox()
        self.txtProjectName.Text = "PanelOptimization"
        self.txtProjectName.Location = Point(12, 40)
        self.txtProjectName.Size = Size(365, 20)
        self.Controls.Add(self.txtProjectName)

        # OK / Cancel buttons
        btnOK = Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = Point(210, 80)

        btnCancel = Button()
        btnCancel.Text = "Cancel"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = Point(290, 80)

        self.AcceptButton = btnOK
        self.CancelButton = btnCancel
        self.Controls.Add(btnOK)
        self.Controls.Add(btnCancel)


def get_project_name():
    """Show ProjectNameDialog and return project name; None if canceled."""
    owner = get_revit_owner()
    while True:
        dlg = ProjectNameDialog()
        result = dlg.ShowDialog(owner) if owner else dlg.ShowDialog()
        if result == DialogResult.OK:
            project_name = dlg.txtProjectName.Text.strip()
            if project_name:
                return project_name
            else:
                # Show error and loop to ask again
                MessageBox.Show(
                    "Project name cannot be empty.",
                    "Invalid Input",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
        else:
            return None

# =================
# Utility functions
# =================

def _ensure_dir(path):
    if not os.path.exists(path):
        try:
            os.makedirs(path)
        except Exception as e:
            print("Failed to create directory {0}: {1}".format(path, e))

def _backup_file(path):
    if os.path.exists(path):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = path.replace(".json", "_backup_{0}.json".format(ts))
        try:
            import shutil
            shutil.copy(path, backup)
        except Exception as e:
            print("Failed to backup file {0}: {1}".format(path, e))

def _sanitize_folder_name(name):
    """Remove invalid characters from folder name."""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name


# =============================================
# UI: Priority Zone -- draw on model (PickBox)
# =============================================
#
# Coordinate mapping from world -> wall-local inches
# -------------------------------------------------
# The export script stores the combined facade as a bounding box whose
# p0/p1 are written to walls.csv as Start(X,Y,Z) and End(X,Y,Z) in feet.
#
#   If facade runs along X  (|p1.X - p0.X| > |p1.Y - p0.Y|):
#       x_local_in = (world.X - p0.X) * 12
#   If facade runs along Y:
#       x_local_in = (world.Y - p0.Y) * 12
#   y_local_in = (world.Z - p0.Z) * 12   (height from base)


def _parse_xyz_str(s):
    """Parse '(X,Y,Z)' string from CSV into (float, float, float) in feet."""
    try:
        s = s.strip().strip("()")
        parts = s.split(",")
        return (float(parts[0]), float(parts[1]), float(parts[2]))
    except Exception:
        return (0.0, 0.0, 0.0)


def _world_to_wall_local(picked_min, picked_max, p0, p1):
    """
    Convert a PickedBox (world XYZ in feet) to wall-local inches.
    Returns (x_start, x_end, y_start, y_end) in inches, clamped >= 0.
    """
    delta_x = abs(p1[0] - p0[0])
    delta_y = abs(p1[1] - p0[1])
    facade_along_x = (delta_x >= delta_y)
    base_z = p0[2]

    if facade_along_x:
        origin = p0[0]
        lo = (min(picked_min.X, picked_max.X) - origin) * 12.0
        hi = (max(picked_min.X, picked_max.X) - origin) * 12.0
    else:
        origin = p0[1]
        lo = (min(picked_min.Y, picked_max.Y) - origin) * 12.0
        hi = (max(picked_min.Y, picked_max.Y) - origin) * 12.0

    y_lo = (min(picked_min.Z, picked_max.Z) - base_z) * 12.0
    y_hi = (max(picked_min.Z, picked_max.Z) - base_z) * 12.0

    return (max(0.0, lo), max(0.0, hi),
            max(0.0, y_lo), max(0.0, y_hi))


def _wall_local_to_world(x_start_in, x_end_in, y_start_in, y_end_in, p0, p1):
    """
    Convert wall-local inches back to world XYZ corners (in feet).
    Returns (min_xyz, max_xyz) as XYZ objects -- used to rebuild highlight.
    """
    delta_x = abs(p1[0] - p0[0])
    delta_y = abs(p1[1] - p0[1])
    facade_along_x = (delta_x >= delta_y)
    base_z = p0[2]

    x_lo_ft = x_start_in / 12.0
    x_hi_ft = x_end_in   / 12.0
    y_lo_ft = y_start_in / 12.0
    y_hi_ft = y_end_in   / 12.0

    if facade_along_x:
        world_min = XYZ(p0[0] + x_lo_ft, p0[1], base_z + y_lo_ft)
        world_max = XYZ(p0[0] + x_hi_ft, p0[1], base_z + y_hi_ft)
    else:
        world_min = XYZ(p0[0], p0[1] + x_lo_ft, base_z + y_lo_ft)
        world_max = XYZ(p0[0], p0[1] + x_hi_ft, base_z + y_hi_ft)

    return world_min, world_max


def _create_zone_highlight(doc, world_min, world_max, p0, p1, thickness_ft=0.1):
    """
    Create a bright orange DirectShape box to highlight the priority zone.
    The box is placed on the wall face, slightly proud of the surface.
    Returns the ElementId of the created DirectShape, or None on failure.
    """
    try:
        delta_x = abs(p1[0] - p0[0])
        delta_y = abs(p1[1] - p0[1])
        facade_along_x = (delta_x >= delta_y)

        # Wall face normal direction (outward) -- we push the box off the face
        # For X-facade: normal is along Y; for Y-facade: normal is along X
        # We use a small offset so the box is visible on the face
        offset = 0.05  # ft -- just enough to show on top of wall

        if facade_along_x:
            # Facade along X: normal is +Y or -Y. Use +Y.
            n = XYZ(0, 1, 0)
        else:
            # Facade along Y: normal is +X or -X. Use +X.
            n = XYZ(1, 0, 0)

        # 8 corners of a thin box sitting on the wall face
        x0 = world_min.X
        x1 = world_max.X
        z0 = world_min.Z
        z1 = world_max.Z

        if facade_along_x:
            y_base = (p0[1] + p1[1]) / 2.0  # mid Y of wall
            corners = [
                XYZ(x0, y_base + offset,             z0),
                XYZ(x1, y_base + offset,             z0),
                XYZ(x1, y_base + offset + thickness_ft, z0),
                XYZ(x0, y_base + offset + thickness_ft, z0),
                XYZ(x0, y_base + offset,             z1),
                XYZ(x1, y_base + offset,             z1),
                XYZ(x1, y_base + offset + thickness_ft, z1),
                XYZ(x0, y_base + offset + thickness_ft, z1),
            ]
        else:
            y0 = world_min.Y
            y1 = world_max.Y
            x_base = (p0[0] + p1[0]) / 2.0
            corners = [
                XYZ(x_base + offset,             y0, z0),
                XYZ(x_base + offset + thickness_ft, y0, z0),
                XYZ(x_base + offset + thickness_ft, y1, z0),
                XYZ(x_base + offset,             y1, z0),
                XYZ(x_base + offset,             y0, z1),
                XYZ(x_base + offset + thickness_ft, y0, z1),
                XYZ(x_base + offset + thickness_ft, y1, z1),
                XYZ(x_base + offset,             y1, z1),
            ]

        if not _TESS_AVAILABLE:
            return None

        builder = TessellatedShapeBuilder()
        builder.OpenConnectedFaceSet(True)

        # 6 faces of the box: bottom, top, front, back, left, right
        face_indices = [
            [0, 1, 2, 3],   # bottom (z0)
            [4, 7, 6, 5],   # top    (z1)
            [0, 4, 5, 1],   # front
            [2, 6, 7, 3],   # back
            [0, 3, 7, 4],   # left
            [1, 5, 6, 2],   # right
        ]
        for fi in face_indices:
            verts = [corners[i] for i in fi]
            builder.AddFace(TessellatedFace(verts, ElementId.InvalidElementId))

        builder.CloseConnectedFaceSet()
        builder.Target = TessellatedShapeBuilderTarget.Solid
        builder.Fallback = TessellatedShapeBuilderFallback.Abort
        builder.Build()
        result = builder.GetBuildResult()

        ds = DirectShape.CreateElement(
            doc, ElementId(int(BuiltInCategory.OST_GenericModel))
        )
        ds.SetShape(result.GetGeometricalObjects())
        ds.Name = "_PriorityZonePreview"
        return ds.Id

    except Exception as e:
        print("[PRIORITY] Highlight creation failed: {}".format(e))
        return None


def _delete_highlight(doc, elem_id):
    """Delete the temporary highlight DirectShape."""
    if elem_id is None:
        return
    try:
        t = Transaction(doc, "_DeletePriorityHighlight")
        t.Start()
        doc.Delete(elem_id)
        t.Commit()
    except Exception as e:
        print("[PRIORITY] Highlight deletion failed: {}".format(e))


# -----------------------------------------------------------------------
# Zone confirm/edit dialog
# -----------------------------------------------------------------------
class ZoneConfirmDialog(Form):
    """
    Shows the picked zone coordinates as editable fields.
    User can Confirm, Redraw, or Cancel this zone.
    Also previews the wall-local inch values numerically.
    """
    CONFIRM = "confirm"
    REDRAW  = "redraw"
    CANCEL  = "cancel"

    def __init__(self, x_start, x_end, y_start, y_end, order, wall_length_in, wall_height_in):
        self.result_action = self.CANCEL
        self._x_start = x_start
        self._x_end   = x_end
        self._y_start = y_start
        self._y_end   = y_end

        self.Text = "Priority Zone {} -- Review & Adjust".format(order)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(420, 290)
        self.TopMost = True

        y = 12

        # Info label
        info = Label()
        info.Text = ("Zone {} drawn. Adjust if needed, then click Confirm.  Wall: {:.0f} in wide x {:.0f} in tall").format(order, wall_length_in, wall_height_in)
        info.Location = Point(12, y)
        info.Size = Size(396, 36)
        self.Controls.Add(info)
        y += 44

        # Coordinate fields
        labels = ["X Start (in):", "X End (in):", "Y Start (in):", "Y End (in):"]
        vals   = [x_start, x_end, y_start, y_end]
        self._fields = []
        for i, (lbl_text, val) in enumerate(zip(labels, vals)):
            lbl = Label()
            lbl.Text = lbl_text
            lbl.Location = Point(12, y + 3)
            lbl.Size = Size(100, 20)
            self.Controls.Add(lbl)

            txt = TextBox()
            txt.Text = "{:.2f}".format(val)
            txt.Location = Point(120, y)
            txt.Size = Size(100, 22)
            self.Controls.Add(txt)
            self._fields.append(txt)

            # Show which axis/direction each coord represents
            hint = Label()
            hint.ForeColor = Color.Gray
            if i == 0:   hint.Text = "distance from wall start"
            elif i == 1: hint.Text = "distance from wall start"
            elif i == 2: hint.Text = "height from wall base"
            else:        hint.Text = "height from wall base"
            hint.Location = Point(228, y + 3)
            hint.Size = Size(180, 20)
            self.Controls.Add(hint)

            y += 30

        y += 10

        # Validation label
        self._val_lbl = Label()
        self._val_lbl.Text = ""
        self._val_lbl.ForeColor = Color.Red
        self._val_lbl.Location = Point(12, y)
        self._val_lbl.Size = Size(396, 20)
        self.Controls.Add(self._val_lbl)
        y += 28

        # Buttons
        btn_confirm = Button()
        btn_confirm.Text = "Confirm Zone"
        btn_confirm.Size = Size(110, 28)
        btn_confirm.Location = Point(12, y)
        btn_confirm.Click += self._on_confirm
        self.Controls.Add(btn_confirm)
        self.AcceptButton = btn_confirm

        btn_redraw = Button()
        btn_redraw.Text = "Redraw"
        btn_redraw.Size = Size(80, 28)
        btn_redraw.Location = Point(132, y)
        btn_redraw.Click += self._on_redraw
        self.Controls.Add(btn_redraw)

        btn_cancel = Button()
        btn_cancel.Text = "Cancel Zone"
        btn_cancel.Size = Size(90, 28)
        btn_cancel.Location = Point(222, y)
        btn_cancel.Click += self._on_cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

    def _parse_fields(self):
        try:
            xs = float(self._fields[0].Text.strip())
            xe = float(self._fields[1].Text.strip())
            ys = float(self._fields[2].Text.strip())
            ye = float(self._fields[3].Text.strip())
            return xs, xe, ys, ye
        except ValueError:
            return None

    def _on_confirm(self, sender, e):
        vals = self._parse_fields()
        if vals is None:
            self._val_lbl.Text = "Please enter valid numbers in all fields."
            return
        xs, xe, ys, ye = vals
        if xe <= xs:
            self._val_lbl.Text = "X End must be greater than X Start."
            return
        if ye <= ys:
            self._val_lbl.Text = "Y End must be greater than Y Start."
            return
        self._x_start, self._x_end = xs, xe
        self._y_start, self._y_end = ys, ye
        self.result_action = self.CONFIRM
        self.Close()

    def _on_redraw(self, sender, e):
        self.result_action = self.REDRAW
        self.Close()

    def _on_cancel(self, sender, e):
        self.result_action = self.CANCEL
        self.Close()

    @property
    def zone_coords(self):
        return (self._x_start, self._x_end, self._y_start, self._y_end)


def _ask_allow_cuts(owner):
    """Small Yes/No dialog: allow panel cuts at zone boundaries?"""
    result = MessageBox.Show(
        "Allow panels to be CUT at priority zone boundaries? Yes=trim, No=whole panels only.",
        "Panel Cut Setting",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question
    )
    from System.Windows.Forms import DialogResult as DR
    return (result == DR.Yes)


def pick_priority_zones(walls_rows):
    """
    Interactive priority zone picker with highlight-and-confirm workflow.

    Per zone:
      1. User draws a PickBox rectangle on the wall
      2. A bright DirectShape highlights the selected area in the model
      3. A dialog shows the coordinates as editable fields
      4. User can Confirm (store zone), Redraw (try again), or Cancel zone

    Returns (zones_enabled, allow_panel_cuts, priority_zones_dict)
    """
    owner = get_revit_owner()
    from System.Windows.Forms import DialogResult as DR

    # Step 0: ask whether to use priority zones at all
    use = MessageBox.Show(
        "Do you want to define priority zones? Priority zones are filled with panels first.",
        "Priority Zones",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question
    )
    if use != DR.Yes:
        return (False, False, {})

    # Step 1: allow_panel_cuts preference
    allow_cuts = _ask_allow_cuts(owner)

    # Step 2: resolve uidoc and doc
    uidoc = _uidoc
    doc   = _doc
    if uidoc is None or doc is None:
        MessageBox.Show(
            "Could not access the Revit document. Priority zones cannot be drawn. Continuing without them.",
            "Priority Zones",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )
        return (False, allow_cuts, {})

    # Step 3: parse wall geometry from walls_rows
    wall_geom = {}
    for wr in walls_rows:
        wid = None
        for col in ["WallId", "ElementId", "Id"]:
            v = wr.get(col, "")
            if v and str(v).strip():
                try:    wid = str(int(float(v)))
                except: wid = str(v).strip()
                break
        if not wid:
            name = wr.get("Name", "")
            wid = str(name).strip() if name else "unknown"

        start_str = wr.get("Start(X,Y,Z)", "") or wr.get("StartXYZ", "")
        end_str   = wr.get("End(X,Y,Z)",   "") or wr.get("EndXYZ",   "")
        if start_str and end_str:
            p0 = _parse_xyz_str(start_str)
            p1 = _parse_xyz_str(end_str)
            # Wall dimensions in inches for the dialog hints
            delta_x = abs(p1[0] - p0[0])
            delta_y = abs(p1[1] - p0[1])
            wall_length_in = max(delta_x, delta_y) * 12.0
            try:
                wall_height_in = float(wr.get("UnconnectedHeight(ft)", 0) or 0) * 12.0
            except (ValueError, TypeError):
                wall_height_in = 0.0
            wall_geom[wid] = {
                'p0': p0, 'p1': p1,
                'length_in': wall_length_in,
                'height_in': wall_height_in
            }

    if not wall_geom:
        MessageBox.Show(
            "Could not read wall geometry from walls.csv. Make sure Start(X,Y,Z) and End(X,Y,Z) columns are present.",
            "Priority Zones",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )
        return (False, allow_cuts, {})

    # Step 4: per-zone draw -> highlight -> confirm loop
    priority_zones = {}
    highlight_id = None  # track any active highlight so we always clean up

    for wid, geom in wall_geom.items():
        order = 1
        while True:

            # Clean up any leftover highlight from previous iteration
            if highlight_id is not None:
                _delete_highlight(doc, highlight_id)
                highlight_id = None

            # Prompt: draw next zone or stop
            cont = MessageBox.Show(
                "Wall {wid} - Zone {order}: Click OK then drag a rectangle on the wall. Click Cancel when done.".format(
                    wid=wid, order=order),
                "Draw Priority Zone",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information
            )
            if cont != DR.OK:
                break

            # Draw the PickBox
            try:
                picked = uidoc.Selection.PickBox(
                    PickBoxStyle.Directional,
                    "Drag to define priority zone {}".format(order)
                )
            except Exception as e:
                print("[PRIORITY] PickBox cancelled: {}".format(e))
                break

            # Convert to wall-local inches
            x_start, x_end, y_start, y_end = _world_to_wall_local(
                picked.Min, picked.Max, geom['p0'], geom['p1']
            )

            if x_end <= x_start or y_end <= y_start:
                MessageBox.Show(
                    "The rectangle was too small or outside wall bounds. Please try again.",
                    "Invalid Zone",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                continue

            # Create highlight in the model
            world_min, world_max = _wall_local_to_world(
                x_start, x_end, y_start, y_end, geom['p0'], geom['p1']
            )
            try:
                t = Transaction(doc, "_PriorityZoneHighlight")
                t.Start()
                highlight_id = _create_zone_highlight(
                    doc, world_min, world_max, geom['p0'], geom['p1']
                )
                t.Commit()
            except Exception as e:
                print("[PRIORITY] Could not create highlight: {}".format(e))
                highlight_id = None

            # Show confirm/edit dialog (loop if user edits and needs new highlight)
            while True:
                dlg = ZoneConfirmDialog(
                    x_start, x_end, y_start, y_end, order,
                    geom['length_in'], geom['height_in']
                )
                dlg.ShowDialog(owner) if owner else dlg.ShowDialog()
                action = dlg.result_action

                if action == ZoneConfirmDialog.CONFIRM:
                    # Read back potentially edited coords
                    x_start, x_end, y_start, y_end = dlg.zone_coords

                    # Delete highlight -- zone is confirmed
                    if highlight_id is not None:
                        _delete_highlight(doc, highlight_id)
                        highlight_id = None

                    zone = {
                        'x_start': round(x_start, 2),
                        'x_end':   round(x_end,   2),
                        'y_start': round(y_start, 2),
                        'y_end':   round(y_end,   2),
                        'order':   order
                    }
                    priority_zones.setdefault(wid, []).append(zone)
                    print("[PRIORITY] Wall {}: Zone {} confirmed  x={:.1f}-{:.1f}in  y={:.1f}-{:.1f}in".format(
                        wid, order, x_start, x_end, y_start, y_end))
                    order += 1
                    break  # inner while -- go to outer loop for next zone

                elif action == ZoneConfirmDialog.REDRAW:
                    # Delete highlight and go back to PickBox
                    if highlight_id is not None:
                        _delete_highlight(doc, highlight_id)
                        highlight_id = None
                    break  # inner while -- outer while will re-prompt

                else:  # CANCEL
                    # Delete highlight and stop adding zones for this wall
                    if highlight_id is not None:
                        _delete_highlight(doc, highlight_id)
                        highlight_id = None
                    break

            # If user cancelled the zone, stop the outer loop too
            if action == ZoneConfirmDialog.CANCEL:
                break

            # If user chose Redraw, the outer while continues naturally
            # (the inner break already happened; action check below skips the outer break)

    # Final cleanup -- belt-and-suspenders in case anything was missed
    if highlight_id is not None:
        _delete_highlight(doc, highlight_id)

    if not any(priority_zones.values()):
        return (False, allow_cuts, {})

    return (True, allow_cuts, priority_zones)


# =====
# Main
# =====
def main():
    try:
        Application.EnableVisualStyles()
    except Exception as e:
        print("EnableVisualStyles failed: {0}".format(e))

    # 1) Pick input folder (click)
    input_dir = pick_data_folder()
    _ensure_dir(input_dir)

    # 2) Get project name (type)
    project_name = get_project_name()
    if project_name is None:
        MessageBox.Show(
            "Operation canceled.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        return

    # 3) Pick orientation (click)
    orientation = pick_orientation()
    if orientation is None:
        MessageBox.Show(
            "Operation canceled.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        return

    # 4) NEW: Pick panel type to determine spacing
    panel_spacing = pick_panel_type()
    if panel_spacing is None:
        MessageBox.Show(
            "Operation canceled.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        return

    # 5) Priority zones (optional)
    # walls_rows not loaded yet -- load CSV now just for wall IDs, reuse later
    walls_csv_early = os.path.join(input_dir, "walls.csv")
    _early_walls_rows = []
    if os.path.exists(walls_csv_early):
        _early_walls_rows = opt.load_walls_from_csv(walls_csv_early)

    zones_enabled, allow_cuts_from_ui, priority_zones_dict = pick_priority_zones(_early_walls_rows)

    # 6) Create timestamped output directory
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_project_name = _sanitize_folder_name(project_name)
    output_folder_name = "{0}_{1}".format(safe_project_name, timestamp)
    output_dir = os.path.join(input_dir, output_folder_name)
    _ensure_dir(output_dir)

    # 7) Build config based on orientation choice
    presets = opt.get_preset_configs()
    if orientation == "custom":
        MessageBox.Show(
            "Custom configuration selected.\n\n" +
            "You'll be asked to configure parameters in the console window.",
            "Custom Configuration",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        config = opt.create_custom_config()
        config.project_name = project_name
    else:
        config = presets[orientation]
        config.project_name = project_name
    
    # Set panel spacing based on user selection
    config.panel_constraints.panel_spacing = panel_spacing
    # Apply the allow_panel_cuts setting chosen in the priority zone dialog
    config.optimization_strategy.allow_panel_cuts = allow_cuts_from_ui
    panel_type_name = "Backer" if panel_spacing == 0.125 else "Fully Finished"

    # Show parameters in console and ask for confirmation
    print("\n" + "=" * 70)
    print("SELECTED CONFIGURATION: {} ({})".format(project_name, orientation.upper()))
    print("Panel Type: {} ({}\")".format(panel_type_name, panel_spacing))
    print("=" * 70)
    opt.print_config_summary(config)
    try:
        user_input = raw_input  # IronPython
    except NameError:
        user_input = input  # CPython
    print("\nOptions:")
    print("1. Continue with these settings")
    print("2. Edit parameters")
    print("3. Cancel")
    choice = user_input("\nChoice (1-3) [1]: ").strip() or "1"
    if choice == "2":
        config = opt.edit_panel_constraints(config)
        config.project_name = project_name  # Restore project name
        config.panel_constraints.panel_spacing = panel_spacing  # Restore spacing
    elif choice == "3":
        print("Cancelled by user.")
        return

    # 8) Resolve CSV paths (in INPUT directory)
    walls_csv = os.path.join(input_dir, "walls.csv")
    openings_csv = os.path.join(input_dir, "wall_openings.csv")

    # Validate inputs with message boxes
    if not os.path.exists(walls_csv):
        MessageBox.Show(
            "Could not find walls.csv in:\n{0}".format(input_dir),
            "Missing Input",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )
        return

    if not os.path.exists(openings_csv):
        MessageBox.Show(
            "Could not find wall_openings.csv in:\n{0}\n\nProceeding without openings.".format(input_dir),
            "Missing Input",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )

    # 9) Run optimizer (output to OUTPUT directory)
    walls_rows = _early_walls_rows if _early_walls_rows else opt.load_walls_from_csv(walls_csv)
    openings_rows = opt.load_openings_from_csv(openings_csv)
    
    # UPDATED: process_all_walls now returns (panels_path, config_path)
    panels_path, config_path = opt.process_all_walls(
        walls_rows, openings_rows, output_dir,  # OUTPUT to timestamped folder
        config.door_clearances,
        config.window_clearances,
        config.storefront_clearances,
        config,       # Pass full config object
        orientation,  # Pass orientation
        priority_zones=priority_zones_dict if zones_enabled else None
    )

    # Config is automatically saved by process_all_walls
    if config_path:
        print("Configuration saved to: {}".format(config_path))
    else:
        print("WARNING: Configuration was not saved")

    # 10) Done message
    if panels_path and os.path.exists(panels_path):
        MessageBox.Show(
            "Optimization complete.\n\nPanel Type: {0}\nSpacing: {1}\"\n\nExported panels to:\n{2}".format(
                panel_type_name, panel_spacing, output_dir),
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    else:
        MessageBox.Show(
            "No panels generated.\nPlease check inputs and configuration.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )

# Entrypoint for pyRevit button
if __name__ == "__main__":
    main()