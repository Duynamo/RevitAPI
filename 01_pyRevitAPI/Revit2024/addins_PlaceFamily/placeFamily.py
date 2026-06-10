
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

def sample_points_along_curves(curves, step_ft):
    """
    Tra ve danh sach (point, tangent) cach deu step_ft doc theo danh sach curves.
    """
    results = []
    accumulated = 0.0
    total_len = sum(c.Length for c in curves)
    if total_len < 1e-9:
        return results

    # Xay dung danh sach doan theo tich luy chieu dai
    segments = []
    cum = 0.0
    for c in curves:
        segments.append((cum, cum + c.Length, c))
        cum += c.Length

    dist = 0.0
    while dist <= total_len + 1e-9:
        # Tim segment chua khoang cach dist
        for seg_start, seg_end, curve in segments:
            if dist >= seg_start - 1e-9 and dist <= seg_end + 1e-9:
                param = max(0.0, min(1.0, (dist - seg_start) / (seg_end - seg_start)))
                try:
                    pt = curve.Evaluate(param, True)
                    deriv = curve.ComputeDerivatives(param, True)
                    tangent = deriv.BasisX.Normalize()
                    results.append((pt, tangent))
                except Exception:
                    pass
                break
        dist += step_ft
    return results

def find_family_symbol(name_query):
    """Tim FamilySymbol khop voi ten (family hoac symbol)."""
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    name_lower = name_query.strip().lower()
    matches = []
    for sym in collector:
        fam_name = ""
        try: fam_name = sym.Family.Name.lower()
        except: pass
        sym_name = ""
        try: sym_name = sym.Name.lower()
        except: pass
        if name_lower in fam_name or name_lower in sym_name:
            matches.append(sym)
    if not matches:
        return None
    # Uu tien khop chinh xac
    for s in matches:
        try:
            if s.Family.Name.lower() == name_lower or s.Name.lower() == name_lower:
                return s
        except: pass
    return matches[0]

def place_instance(symbol, pt, tangent, rotation_extra=0.0):
    """
    Dat FamilyInstance tai pt, xoay theo tangent + rotation_extra (rad).
    Tra ve FamilyInstance da tao.
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

    inst = doc.Create.NewFamilyInstance(
        pt, symbol, level,
        Autodesk.Revit.DB.Structure.StructuralType.NonStructural
    )

    # Xoay theo huong tangent
    try:
        angle = math.atan2(tangent.Y, tangent.X) + rotation_extra
        axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
        ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle)
    except Exception:
        pass
    return inst

# endregion

# region --- Arc-length helpers ---
def _unnorm_to_norm(curve, raw_param):
    """
    Convert a raw (un-normalized) curve parameter returned by curve.Project()
    to a normalized 0-1 parameter for curve.Evaluate(t, True).
    Works correctly for Lines (raw_param = arc-length in ft) and Arcs/NurbSplines
    (raw_param is angle in radians or knot value) by using the curve's own
    GetEndParameter() range to normalize.
    """
    try:
        sp   = curve.GetEndParameter(0)
        ep   = curve.GetEndParameter(1)
        span = ep - sp
        if abs(span) < 1e-12:
            return 0.0
        return max(0.0, min(1.0, (raw_param - sp) / span))
    except Exception:
        return 0.0
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
        self._start_pt       = None # start point for placement
        self._start_tangent  = None
        self._start_arc_dist = 0.0  # arc-length from curve start to _start_pt (ft)
        self._preview_inst   = None # preview instance
        self._rotation       = 0.0  # extra rotation (rad)
        Form.__init__(self)
        self._build()

    def _build(self):
        FW   = 470
        PAD  = 14
        INN  = 10
        HDR  = 60
        FTRH = 36
        GAP  = 10

        H_S1 = 104   # Select path
        H_S2 = 74    # Family name
        H_S3 = 74    # Spacing
        H_S4 = 126   # Start point + Place First (taller so button is not clipped)
        H_S5 = 74    # Rotate + Run
        H_ST = 28    # Status

        Y_S1 = HDR + GAP
        Y_S2 = Y_S1 + H_S1 + GAP
        Y_S3 = Y_S2 + H_S2 + GAP
        Y_S4 = Y_S3 + H_S3 + GAP
        Y_S5 = Y_S4 + H_S4 + GAP
        Y_ST = Y_S5 + H_S5 + GAP
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

        # === S2: Family name ===
        s2 = _card(PAD, Y_S2, FW - PAD * 2, H_S2)
        s2.Controls.Add(_lbl("STEP 2 - Family Name to Place", F_STEP, C_GREEN,
                              INN, 8, CW + INN * 2, 22))
        s2.Controls.Add(_lbl("Family Name:", F_MAIN, C_TEXT, INN, 38, 100, 28))
        self._txt_family = _txt("", INN + 105, 38, CW + INN * 2 - 105, 28)
        s2.Controls.Add(self._txt_family)

        # === S3: Spacing ===
        s3 = _card(PAD, Y_S3, FW - PAD * 2, H_S3)
        s3.Controls.Add(_lbl("STEP 3 - Spacing (distance between each placed item)",
                              F_STEP, C_GREEN, INN, 8, CW + INN * 2, 22))
        s3.Controls.Add(_lbl("Spacing (mm):", F_MAIN, C_TEXT, INN, 38, 110, 28))
        self._txt_step = _txt("1000", INN + 115, 38, CW + INN * 2 - 115, 28)
        s3.Controls.Add(self._txt_step)

        # === S4: Start point + Place First ===
        s4 = _card(PAD, Y_S4, FW - PAD * 2, H_S4)
        s4.Controls.Add(_lbl("STEP 4 - Pick Start Point & Preview First Item",
                              F_STEP, C_GREEN, INN, 8, CW + INN * 2, 22))

        self._btn_pick_start = _btn("Pick Start Point", C_PURPLE, Color.White,
                                    INN, 34, CW + INN * 2, 34)
        self._btn_pick_start.Click += self._on_pick_start
        s4.Controls.Add(self._btn_pick_start)

        self._lbl_start = _lbl("No start point selected.", F_INFO, C_TEXT_DIM,
                                INN, 74, CW + INN * 2, 20)
        s4.Controls.Add(self._lbl_start)

        self._btn_place_first = _btn("Place First Item  (Preview)", C_GREEN_DK, C_GREEN,
                                     INN, 88, CW + INN * 2, 32)
        self._btn_place_first.FlatAppearance.BorderColor = C_GREEN
        self._btn_place_first.FlatAppearance.BorderSize  = 1
        self._btn_place_first.Click += self._on_place_first
        s4.Controls.Add(self._btn_place_first)

        # === S5: Rotate + Run ===
        s5 = _card(PAD, Y_S5, FW - PAD * 2, H_S5)
        s5.Controls.Add(_lbl("STEP 5 - Verify Direction & Run", F_STEP, C_GREEN,
                              INN, 8, CW + INN * 2, 22))

        half = (CW + INN * 2 - INN) // 2 - 4
        self._btn_rotate = _btn("Rotate 90 deg", C_PURPLE, Color.White,
                                INN, 34, half, 34)
        self._btn_rotate.Click += self._on_rotate
        s5.Controls.Add(self._btn_rotate)

        self._btn_run = _btn("Place All", C_GREEN, C_BG,
                             INN + half + 8, 34, half, 34)
        self._btn_run.Click += self._on_run
        s5.Controls.Add(self._btn_run)

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

        for ctrl in [hdr, s1, s2, s3, s4, s5, self._lbl_status, ftr]:
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
            self._set_status("Path selected. Enter family name and spacing.", C_TEXT_DIM)
        except Exception as ex:
            msg = str(ex)
            if "OperationCanceledException" not in msg:
                self._set_status("Pick error: " + msg[:80], Color.Red)
        finally:
            self.Show()
            self.TopMost = True

    # ---- Event: Pick Start Point ----
    def _on_pick_start(self, sender, e):
        if not self._curves:
            self._set_status("Please select a path first.", C_ORANGE)
            return
        self.Hide()
        try:
            pt_ref = uidoc.Selection.PickPoint("Pick start point on or near the path")
            pt     = pt_ref

            curve   = self._curves[0]
            closest = curve.Project(pt)
            self._start_pt = closest.XYZPoint

            # Store start distance as arc-length from curve start (feet).
            # We find it by sampling the curve and measuring up to the closest param.
            norm_param = _unnorm_to_norm(curve, closest.Parameter)
            self._start_arc_dist = norm_param * curve.Length  # arc-length in feet

            deriv = curve.ComputeDerivatives(norm_param, True)
            self._start_tangent = deriv.BasisX.Normalize()
            self._rotation = 0.0

            x = ft_to_mm(self._start_pt.X)
            y = ft_to_mm(self._start_pt.Y)
            z = ft_to_mm(self._start_pt.Z)
            self._lbl_start.Text      = "Start: X={0:.1f}  Y={1:.1f}  Z={2:.1f} mm".format(x, y, z)
            self._lbl_start.ForeColor = C_GREEN
            self._btn_pick_start.BackColor = C_GREEN
            self._btn_pick_start.Text      = "Start Point Selected"
            self._set_status("Start point set. Click 'Place First Item' to preview.", C_TEXT_DIM)
        except Exception as ex:
            msg = str(ex)
            if "OperationCanceledException" not in msg:
                self._set_status("Pick error: " + msg[:80], Color.Red)
        finally:
            self.Show()
            self.TopMost = True

    # ---- Helper: delete preview instance ----
    def _delete_preview(self):
        if self._preview_inst is not None:
            try:
                with Transaction(doc, "Delete Preview") as t:
                    t.Start()
                    doc.Delete(self._preview_inst.Id)
                    t.Commit()
            except Exception:
                pass
            self._preview_inst = None

    # ---- Event: Place First Item ----
    def _on_place_first(self, sender, e):
        if not self._curves:
            self._set_status("No path selected.", C_ORANGE); return
        if self._start_pt is None:
            self._set_status("No start point selected.", C_ORANGE); return

        fam_name = self._txt_family.Text.strip()
        if not fam_name:
            self._set_status("Please enter a family name.", C_ORANGE); return

        sym = find_family_symbol(fam_name)
        if sym is None:
            self._set_status("Family not found: " + fam_name, Color.Red); return

        self._delete_preview()
        try:
            with Transaction(doc, "Place First Preview") as t:
                t.Start()
                inst = place_instance(sym, self._start_pt, self._start_tangent, self._rotation)
                t.Commit()
            self._preview_inst = inst
            self._set_status("Preview placed. Check direction — click Rotate if needed.", C_GREEN)
        except Exception as ex:
            self._set_status("Place error: " + str(ex)[:80], Color.Red)

    # ---- Event: Rotate ----
    def _on_rotate(self, sender, e):
        if self._preview_inst is None:
            self._set_status("Click 'Place First Item' before rotating.", C_ORANGE); return
        self._rotation += math.pi / 2.0
        fam_name = self._txt_family.Text.strip()
        sym = find_family_symbol(fam_name)
        if sym is None:
            self._set_status("Family not found.", Color.Red); return
        self._delete_preview()
        try:
            with Transaction(doc, "Rotate Preview") as t:
                t.Start()
                inst = place_instance(sym, self._start_pt, self._start_tangent, self._rotation)
                t.Commit()
            self._preview_inst = inst
            deg = int(round(math.degrees(self._rotation))) % 360
            self._set_status("Rotated {0} deg total. Click RUN when direction is correct.".format(deg), C_PURPLE_LT)
        except Exception as ex:
            self._set_status("Rotate error: " + str(ex)[:80], Color.Red)

    # ---- Event: Run ----
    def _on_run(self, sender, e):
        if not self._curves:
            self._set_status("No path selected.", C_ORANGE); return
        if self._start_pt is None:
            self._set_status("No start point selected.", C_ORANGE); return

        fam_name = self._txt_family.Text.strip()
        if not fam_name:
            self._set_status("Please enter a family name.", C_ORANGE); return

        try:
            step_mm = float(self._txt_step.Text.strip())
            if step_mm <= 0:
                raise ValueError("Spacing must be > 0")
        except Exception:
            self._set_status("Invalid spacing value.", C_ORANGE); return

        sym = find_family_symbol(fam_name)
        if sym is None:
            self._set_status("Family not found: " + fam_name, Color.Red); return

        step_ft = mm_to_ft(step_mm)
        curve   = self._curves[0]
        total   = curve.Length  # Revit arc-length in feet

        # Build arc-length LUT at ~1 mm resolution
        N_SAMPLES = max(500, int(total / mm_to_ft(1.0)))
        lut_arc  = [0.0]  # arc-length (ft) at each sample
        lut_norm = [0.0]  # normalized param [0..1] at each sample
        lut_pts  = []     # 3D XYZ at each sample (for start_dist lookup)
        prev_pt  = curve.Evaluate(0.0, True)
        lut_pts.append(prev_pt)
        for i in range(1, N_SAMPLES + 1):
            t_norm = float(i) / N_SAMPLES
            cur_pt = curve.Evaluate(t_norm, True)
            lut_arc.append(lut_arc[-1] + prev_pt.DistanceTo(cur_pt))
            lut_norm.append(t_norm)
            lut_pts.append(cur_pt)
            prev_pt = cur_pt
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

        # Phase-based placement: the picked start point defines WHERE one family lands,
        # but families cover the ENTIRE curve (0 -> lut_total).
        # phase = start_dist % step_ft  => first family at 'phase', then every step_ft.
        # This guarantees start_dist is always one of the placement positions AND
        # full path coverage regardless of where the user picks (A, B, or middle).
        phase    = start_dist % step_ft if step_ft > 1e-9 else 0.0

        self._delete_preview()

        placed   = 0
        errors   = 0
        arc_dist = phase   # start from beginning of curve, phase-aligned to start_dist

        try:
            with Transaction(doc, "Place Family Along Path") as t:
                t.Start()
                if not sym.IsActive:
                    sym.Activate()
                    doc.Regenerate()

                while arc_dist <= lut_total + 1e-9:
                    param_norm = arc_to_norm(arc_dist)
                    try:
                        pt   = curve.Evaluate(param_norm, True)
                        deriv = curve.ComputeDerivatives(param_norm, True)
                        tang  = deriv.BasisX.Normalize()
                        place_instance(sym, pt, tang, self._rotation)
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
