
"""
Place Family Along Path - V2D Tools
Cho phep user chon duong dan (line/polyline/CAD link),
nhap ten family, buoc rai, diem bat dau, xem truoc va rai hang loat.
WinForms duoc dung (tuong thich PickObject tren moi phien ban Revit).
"""
# region --- Imports ---
import clr, System, math
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System.Windows.Forms as WinForms
import System.Drawing as Drawing
from System.Windows.Forms import (
    Form, Label, Button, TextBox, Panel, Application,
    FormBorderStyle, FormStartPosition, FlatStyle,
    BorderStyle, AutoScaleMode, ComboBox, MessageBox,
    MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Point, Size, Font, FontStyle, GraphicsUnit, Color, ContentAlignment

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
from Autodesk.Revit.DB import (
    XYZ, Transaction, Line, Arc,
    FilteredElementCollector, BuiltInCategory,
    FamilyInstance, FamilySymbol,
    ElementTransformUtils, Transform,
    CurveElement
)
from Autodesk.Revit.UI.Selection import ObjectType
# endregion

# region --- Revit context ---
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument
# endregion

# region --- Global state ---
data = []
# endregion

# region --- Helpers ---
def mm_to_ft(mm): return mm / 304.8
def ft_to_mm(ft): return ft * 304.8

def get_curve_from_element(elem):
    """Lay curve tu CurveElement (ModelLine, DetailLine, CAD line...)."""
    try:
        if hasattr(elem, 'GeometryCurve'):
            return elem.GeometryCurve
        if hasattr(elem, 'Location'):
            loc = elem.Location
            if hasattr(loc, 'Curve'):
                return loc.Curve
    except Exception:
        pass
    return None

def _unnorm_to_norm(curve, raw_param):
    """
    Convert a raw (un-normalized) curve parameter to a normalized 0-1 parameter.
    """
    try:
        sp   = curve.GetEndParameter(0)
        ep   = curve.GetEndParameter(1)
        span = ep - sp
        if abs(span) < 1e-12:
            return 0.0
        # Normalize and clamp to [0, 1] range
        return max(0.0, min(1.0, (raw_param - sp) / span))
    except Exception:
        # Fallback for curves that don't support GetEndParameter or other issues
        return 0.0

def place_instance_with_tangent_rotation(symbol, pt, tangent, rotation_extra=0.0):
    """
    Places an instance, then rotates it on the XY plane to align with the tangent.
    """
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

    level = doc.ActiveView.GenLevel if hasattr(doc.ActiveView, 'GenLevel') else None
    if level is None:
        levels = FilteredElementCollector(doc).OfCategory(
            BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        if levels:
            level = sorted(levels, key=lambda lv: lv.Elevation)[0]

    # Place directly at the target point
    inst = doc.Create.NewFamilyInstance(
        pt, symbol, level,
        Autodesk.Revit.DB.Structure.StructuralType.NonStructural
    )
    if inst is None: return None
    doc.Regenerate()

    # Rotate the element to align with the tangent on the XY plane
    try:
        angle = math.atan2(tangent.Y, tangent.X) + rotation_extra
        axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
        ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle)
    except Exception:
        # Ignore rotation errors, but the placement will still succeed
        pass

    return inst
# endregion

# region --- UI constants ---
C_GREEN     = Color.FromArgb(255,  30, 180,  80)
C_GREEN_DK  = Color.FromArgb(255,  15,  90,  40)
C_PURPLE_DK = Color.FromArgb(255,  35,  10,  75)
C_PURPLE    = Color.FromArgb(255,  85,  40, 170)
C_PURPLE_LT = Color.FromArgb(255, 175, 140, 240)
C_BG        = Color.FromArgb(255,  14,  10,  35)
C_CARD      = Color.FromArgb(255,  26,  18,  55)
C_TEXT      = Color.FromArgb(255, 220, 210, 255)
C_TEXT_DIM  = Color.FromArgb(255, 145, 130, 190)
C_ORANGE    = Color.FromArgb(255, 255, 160,  30)

F_HDR   = Font("Segoe UI", 14, FontStyle.Bold,    GraphicsUnit.Point)
F_STEP  = Font("Segoe UI", 10, FontStyle.Bold,    GraphicsUnit.Point)
F_MAIN  = Font("Segoe UI", 10, FontStyle.Regular, GraphicsUnit.Point)
F_BTN   = Font("Segoe UI", 10, FontStyle.Bold,    GraphicsUnit.Point)
F_INFO  = Font("Segoe UI",  9, FontStyle.Regular, GraphicsUnit.Point)
F_SMALL = Font("Segoe UI",  8, FontStyle.Regular, GraphicsUnit.Point)
F_LOGO  = Font("Segoe UI",  9, FontStyle.Bold,    GraphicsUnit.Point)

def _btn(text, bg, fg, x, y, w, h):
    b = Button()
    b.Text      = text
    b.Font      = F_BTN
    b.ForeColor = fg
    b.BackColor = bg
    b.FlatStyle = FlatStyle.Flat
    b.FlatAppearance.BorderSize = 0
    b.Location  = Point(x, y)
    b.Size      = Size(w, h)
    b.Cursor    = WinForms.Cursors.Hand
    return b

def _lbl(text, font, fg, x, y, w, h, align=ContentAlignment.MiddleLeft):
    lb = Label()
    lb.Text      = text
    lb.Font      = font
    lb.ForeColor = fg
    lb.BackColor = Color.Transparent
    lb.Location  = Point(x, y)
    lb.Size      = Size(w, h)
    lb.TextAlign = align
    return lb

def _txt(default, x, y, w, h):
    t = TextBox()
    t.Text        = default
    t.Font        = F_MAIN
    t.ForeColor   = C_TEXT
    t.BackColor   = Color.FromArgb(255, 36, 24, 66)
    t.BorderStyle = BorderStyle.FixedSingle
    t.Location    = Point(x, y)
    t.Size        = Size(w, h)
    return t

def _card(x, y, w, h):
    p = Panel()
    p.BackColor = C_CARD
    p.Location  = Point(x, y)
    p.Size      = Size(w, h)
    return p
# endregion

# region --- Form ---
class PlaceFamilyForm(Form):
    def __init__(self):
        self._curves         = []   # list of Curve on path
        self._path_elem      = None # element holding the path
        self._first_inst     = None # first family instance
        self._symbol         = None # family symbol
        self._start_pt       = None # projected start point on curve
        self._offset_dy      = 0.0
        self._offset_dz      = 0.0
        self._angle_offset   = 0.0  # Relative rotation on XY plane
        Form.__init__(self)
        self._build()

    def _get_local_frame(self, T):
        if abs(T.Z) > 0.9999:
            temp_Z = XYZ.BasisX
        else:
            temp_Z = XYZ.BasisZ
        Y = temp_Z.CrossProduct(T).Normalize()
        Z = T.CrossProduct(Y).Normalize()
        return Y, Z

    def _build(self):
        FW   = 470
        PAD  = 14
        INN  = 10
        HDR  = 60
        FTRH = 36
        GAP  = 10

        H_S1 = 104   # Select path
        H_S2 = 104   # Select First Instance
        H_S3 = 74    # Spacing
        H_S4 = 74    # Run
        H_ST = 28    # Status

        Y_S1 = HDR + GAP
        Y_S2 = Y_S1 + H_S1 + GAP
        Y_S3 = Y_S2 + H_S2 + GAP
        Y_S4 = Y_S3 + H_S3 + GAP
        Y_ST = Y_S4 + H_S4 + GAP
        FH   = Y_ST + H_ST + GAP + FTRH

        CW = FW - PAD * 2 - INN * 2
        self.SuspendLayout()

        # === HEADER ===
        hdr = Panel()
        hdr.BackColor = C_PURPLE_DK
        hdr.Location  = Point(0, 0)
        hdr.Size      = Size(FW, HDR)

        ic = Label()
        ic.Text      = "PF"
        ic.Font      = Font("Segoe UI", 13, FontStyle.Bold, GraphicsUnit.Point)
        ic.ForeColor = Color.White
        ic.BackColor = C_GREEN
        ic.Location  = Point(PAD, (HDR - 36) // 2)
        ic.Size      = Size(40, 36)
        ic.TextAlign = ContentAlignment.MiddleCenter
        hdr.Controls.Add(ic)

        t1 = _lbl("Place Family Along Path", F_HDR, Color.White,
                  PAD + 50, 8, 350, 26)
        hdr.Controls.Add(t1)
        t2 = _lbl("Place families along a selected path - V2D Tools",
                  F_SMALL, C_PURPLE_LT, PAD + 50, 35, 360, 16)
        hdr.Controls.Add(t2)

        # === S1: Select path ===
        s1 = _card(PAD, Y_S1, FW - PAD * 2, H_S1)
        s1.Controls.Add(_lbl("STEP 1 - Select Path (Line / Polyline / CAD)",
                              F_STEP, C_GREEN, INN, 8, CW + INN * 2, 22))

        self._btn_pick_path = _btn("Pick Path Element", C_PURPLE, Color.White,
                                   INN, 34, CW + INN * 2, 34)
        self._btn_pick_path.Click += self._on_pick_path
        s1.Controls.Add(self._btn_pick_path)

        self._lbl_path = _lbl("No path selected.", F_INFO, C_TEXT_DIM,
                               INN, 74, CW + INN * 2, 22)
        s1.Controls.Add(self._lbl_path)

        # === S2: Select First Instance ===
        s2 = _card(PAD, Y_S2, FW - PAD * 2, H_S2)
        s2.Controls.Add(_lbl("STEP 2 - Select Placed Family Instance", F_STEP, C_GREEN,
                              INN, 8, CW + INN * 2, 22))

        self._btn_pick_inst = _btn("Pick First Instance", C_PURPLE, Color.White,
                                   INN, 34, CW + INN * 2, 34)
        self._btn_pick_inst.Click += self._on_pick_inst
        s2.Controls.Add(self._btn_pick_inst)

        self._lbl_inst = _lbl("No instance selected.", F_INFO, C_TEXT_DIM,
                               INN, 74, CW + INN * 2, 22)
        s2.Controls.Add(self._lbl_inst)

        # === S3: Spacing ===
        s3 = _card(PAD, Y_S3, FW - PAD * 2, H_S3)
        s3.Controls.Add(_lbl("STEP 3 - Spacing (distance between each placed item)",
                              F_STEP, C_GREEN, INN, 8, CW + INN * 2, 22))
        s3.Controls.Add(_lbl("Spacing (mm):", F_MAIN, C_TEXT, INN, 38, 110, 28))
        self._txt_step = _txt("1000", INN + 115, 38, CW + INN * 2 - 115, 28)
        s3.Controls.Add(self._txt_step)

        # === S4: Run ===
        s4 = _card(PAD, Y_S4, FW - PAD * 2, H_S4)
        s4.Controls.Add(_lbl("STEP 4 - Run", F_STEP, C_GREEN,
                              INN, 8, CW + INN * 2, 22))

        self._btn_run = _btn("Place All", C_GREEN, C_BG,
                             INN, 34, CW + INN * 2, 34)
        self._btn_run.Click += self._on_run
        s4.Controls.Add(self._btn_run)

        # === STATUS ===
        self._lbl_status = _lbl("Ready.", F_INFO, C_TEXT_DIM,
                                 PAD, Y_ST, FW - PAD * 2, H_ST)
        self._lbl_status.AutoSize = False

        # === FOOTER ===
        ftr = Panel()
        ftr.BackColor = Color.FromArgb(255, 8, 5, 20)
        ftr.Location  = Point(0, FH - FTRH)
        ftr.Size      = Size(FW, FTRH)

        lbl_v2d = Label()
        lbl_v2d.Text      = "V2D"
        lbl_v2d.Font      = F_LOGO
        lbl_v2d.ForeColor = C_GREEN
        lbl_v2d.BackColor = C_PURPLE_DK
        lbl_v2d.Location  = Point(PAD, (FTRH - 20) // 2)
        lbl_v2d.Size      = Size(36, 20)
        lbl_v2d.TextAlign = ContentAlignment.MiddleCenter
        ftr.Controls.Add(lbl_v2d)

        lbl_copy = _lbl("V2D Tools  |  Place Family Along Path v1.0",
                         F_SMALL, C_TEXT_DIM, PAD + 44, (FTRH - 16) // 2, 320, 16)
        ftr.Controls.Add(lbl_copy)

        # === Form settings ===
        self.AutoScaleDimensions = Drawing.SizeF(96, 96)
        self.AutoScaleMode       = AutoScaleMode.Dpi
        self.ClientSize          = Size(FW, FH)
        self.FormBorderStyle     = FormBorderStyle.FixedSingle
        self.StartPosition       = FormStartPosition.CenterScreen
        self.BackColor           = C_BG
        self.Font                = F_MAIN
        self.Text                = "Place Family Along Path  |  V2D Tools"
        self.MaximizeBox         = False
        self.TopMost             = True

        for ctrl in [hdr, s1, s2, s3, s4, self._lbl_status, ftr]:
            self.Controls.Add(ctrl)

        self.ResumeLayout(False)
        self.PerformLayout()

    # ---- Event: Pick Path ----
    def _on_pick_path(self, sender, e):
        self.Hide()
        try:
            ref  = uidoc.Selection.PickObject(ObjectType.Element,
                                              "Select path (Line / CAD line / Polyline)")
            elem = doc.GetElement(ref)
            curve = get_curve_from_element(elem)
            if curve is None:
                self._set_status("Element has no Curve geometry. Try another element.", C_ORANGE)
                return
            self._curves    = [curve]
            self._path_elem = elem
            length_mm = ft_to_mm(curve.Length)
            self._lbl_path.Text      = "Selected: ID {0}  |  Length: {1:.1f} mm".format(
                elem.Id.IntegerValue, length_mm)
            self._lbl_path.ForeColor = C_GREEN
            self._btn_pick_path.BackColor = C_GREEN
            self._btn_pick_path.Text      = "Path Selected"
            self._set_status("Path selected. Select the first placed family instance.", C_TEXT_DIM)
        except Exception as ex:
            msg = str(ex)
            if "OperationCanceledException" not in msg:
                self._set_status("Pick error: " + msg[:80], Color.Red)
        finally:
            self.Show()
            self.TopMost = True

    # ---- Event: Pick First Instance ----
    def _on_pick_inst(self, sender, e):
        if not self._curves:
            self._set_status("Please select a path first.", C_ORANGE)
            return
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, "Select the first placed family instance")
            inst = doc.GetElement(ref)
            if not isinstance(inst, FamilyInstance):
                self._set_status("Selected element is not a FamilyInstance.", C_ORANGE)
                self.Show(); self.TopMost = True
                return
            
            loc = inst.Location
            if not hasattr(loc, "Point"):
                self._set_status("FamilyInstance has no point location.", C_ORANGE)
                self.Show(); self.TopMost = True
                return
            
            self._first_inst = inst
            self._symbol = inst.Symbol
            inst_pt = loc.Point
            
            curve   = self._curves[0]
            prj = curve.Project(inst_pt)
            if prj is None:
                self._set_status("Instance is too far from path.", C_ORANGE)
                self.Show(); self.TopMost = True
                return
                
            prj_pt = prj.XYZPoint
            norm_param = _unnorm_to_norm(curve, prj.Parameter)
            self._start_pt = prj_pt 
            
            deriv = curve.ComputeDerivatives(norm_param, True)
            T = deriv.BasisX.Normalize()
            Y, Z = self._get_local_frame(T)
            
            offset_vec = inst_pt - prj_pt
            self._offset_dy = offset_vec.DotProduct(Y)
            self._offset_dz = offset_vec.DotProduct(Z)
            
            # --- Calculate the relative 2D rotation offset ---
            # This finds the angle difference between the instance's direction
            # and the path's tangent, on the XY plane.
            inst_transform = inst.GetTransform()
            inst_dir = inst_transform.BasisX
            angle_curve = math.atan2(T.Y, T.X)
            angle_inst = math.atan2(inst_dir.Y, inst_dir.X)
            self._angle_offset = angle_inst - angle_curve

            fam_name = self._symbol.Family.Name
            sym_name = self._symbol.Name
            name_disp = "{} - {}".format(fam_name, sym_name)
            if len(name_disp) > 30:
                name_disp = name_disp[:27] + "..."
            self._lbl_inst.Text = "Selected: {} (Offset: {:.1f} mm)".format(name_disp, ft_to_mm(self._offset_dy))
            self._lbl_inst.ForeColor = C_GREEN
            self._btn_pick_inst.BackColor = C_GREEN
            self._btn_pick_inst.Text = "Instance Selected"
            self._set_status("First instance selected. Enter spacing and click Place All.", C_TEXT_DIM)
            
        except Exception as ex:
            msg = str(ex)
            if "OperationCanceledException" not in msg:
                self._set_status("Pick error: " + msg[:80], Color.Red)
        finally:
            self.Show()
            self.TopMost = True

    # ---- Event: Run ----
    def _on_run(self, sender, e):
        if not self._curves:
            self._set_status("No path selected.", C_ORANGE); return
        if self._first_inst is None:
            self._set_status("No first instance selected.", C_ORANGE); return

        try:
            step_mm = float(self._txt_step.Text.strip())
            if step_mm <= 0:
                raise ValueError("Spacing must be > 0")
        except Exception:
            self._set_status("Invalid spacing value.", C_ORANGE); return

        step_ft = mm_to_ft(step_mm)
        curve   = self._curves[0]
        total   = curve.Length  # Revit arc-length in feet

        # Build arc-length LUT at ~1 mm resolution
        N_SAMPLES = max(500, int(total / mm_to_ft(1.0)))
        lut_arc  = [0.0]  # arc-length (ft) at each sample
        lut_norm = [0.0]  # normalized param [0..1] at each sample
        lut_pts  = []     # 3D XYZ at each sample (for start_dist lookup)

        prev_pt = curve.Evaluate(0.0, True)
        if prev_pt is None:
            self._set_status("Error: Could not evaluate the start of the path curve.", Color.Red)
            return

        lut_pts.append(prev_pt)
        for i in range(1, N_SAMPLES + 1):
            t_norm = float(i) / N_SAMPLES
            cur_pt = curve.Evaluate(t_norm, True)

            # If evaluation fails, skip this sample point to avoid errors.
            if cur_pt is None:
                continue

            lut_arc.append(lut_arc[-1] + prev_pt.DistanceTo(cur_pt))
            lut_norm.append(t_norm)
            lut_pts.append(cur_pt)
            prev_pt = cur_pt # Update prev_pt to the new valid point

        lut_total = lut_arc[-1]

        def arc_to_norm(arc_dist_val):
            """Invert LUT: arc-length -> normalized param (binary search)."""
            if arc_dist_val <= 0.0:       return 0.0
            if arc_dist_val >= lut_total: return 1.0
            lo, hi = 0, len(lut_arc) - 1
            while lo < hi - 1:
                mid = (lo + hi) // 2
                if lut_arc[mid] <= arc_dist_val: lo = mid
                else: hi = mid
            span = lut_arc[hi] - lut_arc[lo]
            frac = (arc_dist_val - lut_arc[lo]) / span if span > 1e-12 else 0.0
            return lut_norm[lo] + frac * (lut_norm[hi] - lut_norm[lo])

        # Find start_dist by nearest-point search in lut_pts.
        # This is robust for all curve types (Line, Arc, Spline) because
        # it works purely in 3D space with no parameter-unit assumptions.
        try:
            prj_pt  = curve.Project(self._start_pt).XYZPoint
        except Exception:
            prj_pt  = self._start_pt
        min_d    = 1e30
        best_idx = 0
        for idx in range(len(lut_pts)):
            d = prj_pt.DistanceTo(lut_pts[idx])
            if d < min_d:
                min_d    = d
                best_idx = idx
        start_dist = lut_arc[best_idx]

        placed   = 0
        errors   = 0
        # Start placing at the NEXT point after the first selected instance
        arc_dist = start_dist + step_ft

        try:
            with Transaction(doc, "Place Family Along Path") as t:
                t.Start()
                sym = self._symbol
                if not sym.IsActive:
                    sym.Activate()
                    doc.Regenerate()

                while arc_dist <= lut_total + 1e-9:
                    param_norm = arc_to_norm(arc_dist)
                    try:
                        pt   = curve.Evaluate(param_norm, True)
                        deriv = curve.ComputeDerivatives(param_norm, True)
                        T = deriv.BasisX.Normalize()
                        Y, Z = self._get_local_frame(T)
                        
                        # Apply fixed offset
                        new_pt = pt + Y * self._offset_dy + Z * self._offset_dz
                        
                        # Place instance with tangent rotation
                        place_instance_with_tangent_rotation(sym, new_pt, T, self._angle_offset)

                        placed += 1
                    except Exception as ex2:
                        errors += 1
                        data.append(str(ex2))
                    arc_dist += step_ft   # advance exactly step_ft along arc

                t.Commit()

            msg = "Done! Placed {0} family instance(s).".format(placed)
            if errors:
                msg += "  ({0} error(s) - see OUT)".format(errors)
            self._set_status(msg, C_GREEN)
        except Exception as ex:
            self._set_status("Error: " + str(ex)[:90], Color.Red)
            data.append(str(ex))

    # ---- Helper ----
    def _set_status(self, msg, color):
        self._lbl_status.Text      = msg
        self._lbl_status.ForeColor = color

# endregion

# region --- Run ---
form = PlaceFamilyForm()
Application.Run(form)
OUT = data
# endregion
