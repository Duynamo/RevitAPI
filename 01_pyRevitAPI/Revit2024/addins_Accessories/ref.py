"""
Create Elbow Tool - V2D Tools
Ho tro: Pipe, Pipe Fitting (elbow/tee/reducer), Pipe Accessory (valve,...)
Quy trinh: pick element -> pick point -> nhap goc -> tao elbow (qua ong ao an).
WinForms duoc dung thay WPF vi hop tac voi Revit PickObject trong Dynamo.
"""
# region --- Imports ---
import clr, System, math
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System.Windows.Forms as WinForms
import System.Drawing as Drawing
from System.Windows.Forms import (
    Form, Label, Button, TextBox, Panel, Application,
    FormBorderStyle, FormStartPosition, FlatStyle, BorderStyle, AutoScaleMode
)
from System.Drawing import Point, Size, Font, FontStyle, GraphicsUnit, Color, ContentAlignment
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
from Autodesk.Revit.DB import (
    XYZ, Transaction, TransactionGroup,
    BuiltInCategory, BuiltInParameter, FilteredElementCollector
)
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes
# endregion

# region --- Revit context ---
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument
# endregion

# region --- Global state ---
data = []
# endregion

# region --- Backend helpers ---

def mm_to_ft(mm): return mm / 304.8
def ft_to_mm(ft): return ft * 304.8

def get_connector_manager(elem):
    """
    Lay ConnectorManager tu bat ky loai MEP element nao:
    Pipe, PipeFitting (elbow/tee/reducer), PipeAccessory (valve,...).
    """
    # Pipe co thuoc tinh ConnectorManager truc tiep
    if hasattr(elem, 'ConnectorManager') and elem.ConnectorManager is not None:
        return elem.ConnectorManager
    # Fitting va Accessory co ConnectorManager qua MEPModel
    try:
        mep = elem.MEPModel
        if mep is not None and mep.ConnectorManager is not None:
            return mep.ConnectorManager
    except Exception:
        pass
    return None

def get_nearest_connector(elem, point_xyz):
    """Tim connector cua element gan voi point_xyz nhat."""
    cm = get_connector_manager(elem)
    if cm is None:
        return None
    nearest, min_dist = None, 1e12
    for con in cm.Connectors:
        d = con.Origin.DistanceTo(point_xyz)
        if d < min_dist:
            min_dist, nearest = d, con
    return nearest

def get_other_connector(elem, chosen_con):
    """Tra ve connector doi dien cua pipe (dung cho ong ao)."""
    cm = get_connector_manager(elem)
    if cm is None:
        return None
    for con in cm.Connectors:
        if not con.Origin.IsAlmostEqualTo(chosen_con.Origin):
            return con
    return None

def get_pipe_params(elem):
    """
    Lay (system_id, type_id, level_id, diameter) tu element.
    Neu element la fitting/accessory, tim pipe da ket noi de lay thong so.
    """
    # Thu lay truc tiep tu Pipe
    try:
        s = elem.MEPSystem.GetTypeId()
        t = elem.GetTypeId()
        l = elem.LevelId
        d = elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
        return s, t, l, d
    except Exception:
        pass

    # Voi fitting/accessory: tim pipe ket noi qua connectors
    cm = get_connector_manager(elem)
    if cm is not None:
        for con in cm.Connectors:
            try:
                for ref_con in con.AllRefs:
                    linked = doc.GetElement(ref_con.Owner.Id)
                    if linked is None or linked.Id == elem.Id:
                        continue
                    try:
                        s = linked.MEPSystem.GetTypeId()
                        t = linked.GetTypeId()
                        l = linked.LevelId
                        d = linked.get_Parameter(
                            BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
                        return s, t, l, d
                    except Exception:
                        pass
            except Exception:
                pass

    # Fallback: lay pipe type dau tien trong du an
    pipes = FilteredElementCollector(doc).OfClass(
        Autodesk.Revit.DB.Plumbing.PipeType).ToElements()
    pipe_instances = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
    if len(pipe_instances) > 0:
        p0 = pipe_instances[0]
        try:
            s = p0.MEPSystem.GetTypeId()
            t = p0.GetTypeId()
            l = p0.LevelId
            d = p0.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
            return s, t, l, d
        except Exception:
            pass
    raise Exception("Cannot determine pipe parameters from the selected element.")

def rotate_vector(vec, axis, angle_rad):
    """Xoay vec quanh axis mot goc angle_rad (Rodrigues rotation)."""
    ax  = axis.Normalize()
    cos = math.cos(angle_rad)
    sin = math.sin(angle_rad)
    dot = vec.DotProduct(ax)
    crs = ax.CrossProduct(vec)
    return XYZ(
        cos*vec.X + sin*crs.X + (1-cos)*dot*ax.X,
        cos*vec.Y + sin*crs.Y + (1-cos)*dot*ax.Y,
        cos*vec.Z + sin*crs.Z + (1-cos)*dot*ax.Z
    )

def create_elbow_from_connector(elem, sel_con, angle_degrees):
    """
    Tao elbow tu connector duoc chon:
    1. Tao ong ao 500mm theo huong xoay.
    2. Tao elbow noi elem va ong ao.
    3. Di chuyen elbow ve vi tri goc de khong lam thay doi vi tri ong chinh.
    4. Xoa ong ao.
    Dung TransactionGroup.Assimilate() -> Revit chi render ket qua cuoi.
    """
    system_id, type_id, level_id, diameter = get_pipe_params(elem)

    # Lay duong kinh tu ban than connector la chinh xac nhat
    # Tranh truong hop get_pipe_params lay sai duong kinh -> Revit tu chen reducer
    try:
        diameter = sel_con.Radius * 2.0
    except Exception:
        pass  # giu diameter tu get_pipe_params

    origin = sel_con.Origin

    # Dung BasisZ cua connector lam huong ra ngoai CHINH XAC
    # Voi pipe: BasisZ = huong doc truc ong tai dau connector (outward)
    # Voi fitting (elbow/tee/reducer): BasisZ = huong thuc cua cong ket noi,
    # KHONG phai vector noi 2 connector (vector noi 2 connector cua 90-deg elbow
    # bi lech 45 do so voi huong thuc -> gay ra loi goc sai khi dung A-B)
    try:
        outward_dir = sel_con.CoordinateSystem.BasisZ
        # Kiem tra: neu vector qua ngan (degenerate) thi fallback
        if outward_dir.GetLength() < 0.001:
            raise Exception("degenerate BasisZ")
    except Exception:
        # Fallback: tinh tu vi tri 2 connector (chi chinh xac voi pipe thang)
        con_b = get_other_connector(elem, sel_con)
        if con_b is not None:
            outward_dir = origin.Subtract(con_b.Origin).Normalize()
        else:
            outward_dir = XYZ(0, 0, 1)

    # Chon truc xoay vuong goc voi outward_dir de xoay trong mat phang hop ly
    ref_vec  = XYZ(0, 0, 1) if abs(outward_dir.Z) < 0.9999 else XYZ(1, 0, 0)
    rot_axis = outward_dir.CrossProduct(ref_vec).Normalize()
    new_dir  = rotate_vector(outward_dir, rot_axis, math.radians(angle_degrees))
    end_pt   = origin.Add(new_dir.Multiply(mm_to_ft(500.0)))

    elbow_created = None
    tg = TransactionGroup(doc, "Create Elbow")
    tg.Start()
    try:
        # Buoc 1: tao ong ao cung system, type, level, diameter voi elem goc
        t1 = Transaction(doc, "Ghost Pipe")
        t1.Start()
        ghost = Pipe.Create(doc, system_id, type_id, level_id, origin, end_pt)
        ghost.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter)
        t1.Commit()

        # Buoc 2: tim dung connector tai diem origin cua ca hai ong
        # IsAlmostEqualTo dung tolerance mac dinh cua Revit (~1e-9 ft)
        elem_con  = None
        ghost_con = None
        cm_elem  = get_connector_manager(elem)
        cm_ghost = get_connector_manager(ghost)
        for c in cm_elem.Connectors:
            if c.Origin.IsAlmostEqualTo(origin):
                elem_con = c; break
        for c in cm_ghost.Connectors:
            if c.Origin.IsAlmostEqualTo(origin):
                ghost_con = c; break
        if elem_con is None or ghost_con is None:
            raise Exception("Cannot locate connectors at origin for elbow.")

        # Buoc 3: tao elbow noi hai connector
        # SAU BUOC NAY Revit co the da dich chuyen pipe goc (elem) de fit geometry.
        t2 = Transaction(doc, "Elbow Fitting")
        t2.Start()
        elbow_created = doc.Create.NewElbowFitting(elem_con, ghost_con)
        t2.Commit()

        # Buoc 4: Di chuyen elbow ve dung vi tri origin
        # Khac phuc loi ong chinh bi thay doi vi tri khi tao elbow.
        t_move = Transaction(doc, "Move Elbow")
        t_move.Start()
        try:
            elbow_cm = get_connector_manager(elbow_created)
            elbow_con_to_elem = None
            if elbow_cm is not None:
                for c in elbow_cm.Connectors:
                    if c.IsConnected:
                        for ref in c.AllRefs:
                            if ref.Owner.Id == elem.Id:
                                elbow_con_to_elem = c
                                break
                    if elbow_con_to_elem is not None:
                        break
            
            if elbow_con_to_elem is not None:
                from Autodesk.Revit.DB import ElementTransformUtils
                move_vec = origin.Subtract(elbow_con_to_elem.Origin)
                if move_vec.GetLength() > 0.0001:
                    ElementTransformUtils.MoveElement(doc, elbow_created.Id, move_vec)
        except Exception:
            pass
        t_move.Commit()

        # Buoc 5: xoa ong ao -> user chi thay elbow
        t3 = Transaction(doc, "Delete Ghost Pipe")
        t3.Start()
        doc.Delete(ghost.Id)
        t3.Commit()

        tg.Assimilate()  # gop tat ca trong 1 TransactionGroup -> chi render ket qua cuoi
    except Exception as ex:
        tg.RollBack()
        data.append(str(ex))
        raise
    return elbow_created

# endregion

# region --- UI constants ---
C_PURPLE_DK = Color.FromArgb(255, 42,  10,  90)
C_PURPLE    = Color.FromArgb(255, 90,  45, 175)
C_PURPLE_LT = Color.FromArgb(255,180, 145, 240)
C_GREEN     = Color.FromArgb(255, 25, 170,  75)
C_BG        = Color.FromArgb(255, 18,  12,  40)
C_CARD      = Color.FromArgb(255, 30,  20,  60)
C_TEXT      = Color.FromArgb(255,225, 215, 255)
C_TEXT_DIM  = Color.FromArgb(255,150, 135, 195)
C_SEP       = Color.FromArgb(255, 65,  45, 110)

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
# endregion

# region --- Form ---
class CreateElbowForm(Form):
    """
    Form chinh. WinForms duoc dung vi hop tac voi Revit PickObject.
    Ho tro: Pipe, PipeFitting (elbow/tee/reducer), PipeAccessory (valve,...).
    """
    def __init__(self):
        self._elem      = None   # element duoc chon (pipe / fitting / accessory)
        self._connector = None   # connector gan diem click nhat
        Form.__init__(self)
        self._build()

    def _build(self):
        # --- Layout constants ---
        FW   = 450    # chieu rong form
        PAD  = 14     # padding ngoai
        INN  = 10     # padding ben trong card
        HDR  = 58     # chieu cao header
        FTRH = 36     # chieu cao footer

        # Chieu cao cac card
        H_S1 = 112
        H_S2 = 110
        H_S3 = 118

        GAP  = 10     # khoang cach giua cac card

        # Vi tri Y cac section (tinh chinh xac tranh overlap)
        Y_S1     = HDR + GAP
        Y_S2     = Y_S1 + H_S1 + GAP
        Y_S3     = Y_S2 + H_S2 + GAP
        Y_STATUS = Y_S3 + H_S3 + GAP
        H_STATUS = 26
        # Chieu cao form = status + gap + footer (khong de overlap)
        FH = Y_STATUS + H_STATUS + GAP + FTRH

        self.SuspendLayout()
        CW = FW - PAD * 2 - INN * 2  # chieu rong noi dung trong card

        # === HEADER ===
        hdr = Panel()
        hdr.BackColor = C_PURPLE_DK
        hdr.Location  = Point(0, 0)
        hdr.Size      = Size(FW, HDR)

        icon = Label()
        icon.Text      = "E"
        icon.Font      = Font("Segoe UI", 16, FontStyle.Bold, GraphicsUnit.Point)
        icon.ForeColor = Color.White
        icon.BackColor = C_GREEN
        icon.Location  = Point(PAD, (HDR-36)//2)
        icon.Size      = Size(36, 36)
        icon.TextAlign = ContentAlignment.MiddleCenter
        hdr.Controls.Add(icon)

        t1 = Label()
        t1.Text      = "Create Elbow"
        t1.Font      = F_HDR
        t1.ForeColor = Color.White
        t1.BackColor = Color.Transparent
        t1.Location  = Point(PAD+44, 8)
        t1.Size      = Size(280, 26)
        hdr.Controls.Add(t1)

        t2 = Label()
        t2.Text      = "Pipe / Fitting / Accessory  ->  Pick Point  ->  Create"
        t2.Font      = F_SMALL
        t2.ForeColor = C_PURPLE_LT
        t2.BackColor = Color.Transparent
        t2.Location  = Point(PAD+44, 34)
        t2.Size      = Size(360, 16)
        hdr.Controls.Add(t2)

        # === STEP 1 card ===
        s1 = Panel()
        s1.BackColor = C_CARD
        s1.Location  = Point(PAD, Y_S1)
        s1.Size      = Size(FW - PAD*2, H_S1)

        s1.Controls.Add(_lbl("STEP 1 - Select Element", F_STEP, C_GREEN, INN, 8, 300, 22))

        self._btn_pick_elem = _btn("Pick Element  (Pipe / Fitting / Accessory)",
                                   C_PURPLE, Color.White,
                                   INN, 34, FW-PAD*2-INN*2, 36)
        self._btn_pick_elem.Click += self._on_pick_elem
        s1.Controls.Add(self._btn_pick_elem)

        self._lbl_elem_info = _lbl("No element selected.", F_INFO, C_TEXT_DIM,
                                   INN, 76, FW-PAD*2-INN*2, 28)
        s1.Controls.Add(self._lbl_elem_info)

        # === STEP 2 card ===
        s2 = Panel()
        s2.BackColor = C_CARD
        s2.Location  = Point(PAD, Y_S2)
        s2.Size      = Size(FW - PAD*2, H_S2)

        s2.Controls.Add(_lbl("STEP 2 - Pick Connector Center Point", F_STEP, C_GREEN, INN, 8, 280, 22))

        self._btn_pick_pt = _btn("Pick Point", C_PURPLE, Color.White,
                                 INN, 34, FW-PAD*2-INN*2, 36)
        self._btn_pick_pt.Click += self._on_pick_point
        s2.Controls.Add(self._btn_pick_pt)

        self._lbl_con_info = _lbl("No point selected.", F_INFO, C_TEXT_DIM,
                                  INN, 76, FW-PAD*2-INN*2, 28)
        s2.Controls.Add(self._lbl_con_info)

        # === STEP 3 card ===
        s3 = Panel()
        s3.BackColor = C_CARD
        s3.Location  = Point(PAD, Y_S3)
        s3.Size      = Size(FW - PAD*2, H_S3)

        s3.Controls.Add(_lbl("STEP 3 - Set Angle and Create Elbow",
                              F_STEP, C_GREEN, INN, 8, 320, 22))

        s3.Controls.Add(_lbl("Angle (deg):", F_MAIN, C_TEXT, INN, 38, 100, 28))

        self._txt_angle = TextBox()
        self._txt_angle.Text        = "90"
        self._txt_angle.Font        = F_MAIN
        self._txt_angle.ForeColor   = C_TEXT
        self._txt_angle.BackColor   = Color.FromArgb(255, 38, 26, 68)
        self._txt_angle.BorderStyle = BorderStyle.FixedSingle
        self._txt_angle.Location    = Point(INN + 108, 36)
        self._txt_angle.Size        = Size(FW-PAD*2-INN*2-108, 28)
        s3.Controls.Add(self._txt_angle)

        self._btn_create = _btn("Create Elbow", C_GREEN, C_BG,
                                INN, 70, FW-PAD*2-INN*2, 38)
        self._btn_create.Click += self._on_create_elbow
        s3.Controls.Add(self._btn_create)

        # === STATUS label ===
        self._lbl_status = _lbl("Ready.", F_INFO, C_TEXT_DIM,
                                 PAD, Y_STATUS, FW-PAD*2, H_STATUS)
        self._lbl_status.AutoSize = False

        # === FOOTER ===
        ftr = Panel()
        ftr.BackColor = Color.FromArgb(255, 10, 6, 22)
        ftr.Location  = Point(0, FH - FTRH)   # anker o cuoi chinh xac
        ftr.Size      = Size(FW, FTRH)

        lbl_v2d = Label()
        lbl_v2d.Text      = "V2D"
        lbl_v2d.Font      = F_LOGO
        lbl_v2d.ForeColor = C_GREEN
        lbl_v2d.BackColor = C_PURPLE_DK
        lbl_v2d.Location  = Point(PAD, (FTRH-20)//2)
        lbl_v2d.Size      = Size(36, 20)
        lbl_v2d.TextAlign = ContentAlignment.MiddleCenter
        ftr.Controls.Add(lbl_v2d)

        lbl_copy = _lbl("V2D Tools  |  Create Elbow v1.1",
                         F_SMALL, C_TEXT_DIM,
                         PAD+44, (FTRH-16)//2, 260, 16)
        ftr.Controls.Add(lbl_copy)

        # === Form settings ===
        self.AutoScaleDimensions = Drawing.SizeF(96, 96)
        self.AutoScaleMode       = AutoScaleMode.Dpi
        self.ClientSize          = Size(FW, FH)
        self.FormBorderStyle     = FormBorderStyle.FixedSingle
        self.StartPosition       = FormStartPosition.CenterScreen
        self.BackColor           = C_BG
        self.Font                = F_MAIN
        self.Text                = "Create Elbow  |  V2D Tools"
        self.MaximizeBox         = False
        self.TopMost             = True

        self.Controls.Add(hdr)
        self.Controls.Add(s1)
        self.Controls.Add(s2)
        self.Controls.Add(s3)
        self.Controls.Add(self._lbl_status)
        self.Controls.Add(ftr)

        self.ResumeLayout(False)
        self.PerformLayout()

    # --- Event: Pick Element ---
    def _on_pick_elem(self, sender, e):
        """
        User pick element: Pipe, PipeFitting hoac PipeAccessory.
        An form truoc khi goi PickObject, hien lai sau khi chon xong.
        """
        self.Hide()
        try:
            ref  = uidoc.Selection.PickObject(
                ObjectType.Element,
                "Select Pipe / Pipe Fitting / Pipe Accessory"
            )
            elem = doc.GetElement(ref)

            # Kiem tra element co ConnectorManager khong
            cm = get_connector_manager(elem)
            if cm is None:
                self._set_status("Selected element has no connectors.", Color.Orange)
                return

            self._elem      = elem
            self._connector = None

            # Lay ten hien thi
            try:
                elem_name = elem.Name if elem.Name else "Element"
            except Exception:
                elem_name = "Element"
            elem_id = str(elem.Id.IntegerValue)

            self._lbl_elem_info.Text      = "{0}  |  ID: {1}".format(elem_name, elem_id)
            self._lbl_elem_info.ForeColor = C_GREEN
            self._btn_pick_elem.BackColor = C_GREEN
            self._btn_pick_elem.Text      = "Element Selected"

            # Reset buoc 2
            self._lbl_con_info.Text       = "No point selected."
            self._lbl_con_info.ForeColor  = C_TEXT_DIM
            self._btn_pick_pt.BackColor   = C_PURPLE
            self._btn_pick_pt.Text        = "Pick Point"

            self._set_status("Element selected. Now pick a point near a connector.", C_TEXT_DIM)
        except Exception:
            pass  # user nhan ESC
        finally:
            self.Show()
            self.TopMost = True

    # --- Event: Pick Point ---
    def _on_pick_point(self, sender, e):
        """
        User click diem gan connector: tim connector gan nhat va hien toa do (mm).
        """
        if self._elem is None:
            self._set_status("Please select an element first.", Color.Orange)
            return
        self.Hide()
        try:
            # Thay vi PickPoint (hay bi loi No Work Plane o 3D), dung PickObject PointOnElement
            ref  = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "Click on any element near the desired connector"
            )
            pt = ref.GlobalPoint
            
            con = get_nearest_connector(self._elem, pt)
            if con is None:
                self._set_status("No connector found. Try again.", Color.Orange)
                return

            self._connector = con
            ox = ft_to_mm(con.Origin.X)
            oy = ft_to_mm(con.Origin.Y)
            oz = ft_to_mm(con.Origin.Z)
            self._lbl_con_info.Text = (
                "Connector (mm):  X={0:.1f}  Y={1:.1f}  Z={2:.1f}"
            ).format(ox, oy, oz)
            self._lbl_con_info.ForeColor = C_GREEN
            self._btn_pick_pt.BackColor  = C_GREEN
            self._btn_pick_pt.Text       = "Point Selected"
            self._set_status("Connector found. Enter angle then click Create Elbow.", C_TEXT_DIM)
        except Exception as ex:
            msg = str(ex)
            if "OperationCanceledException" not in msg:
                self._set_status("Pick error: " + msg[:60], Color.Orange)
        finally:
            self.Show()
            self.TopMost = True

    # --- Event: Create Elbow ---
    def _on_create_elbow(self, sender, e):
        """Tao elbow qua ong ao (TransactionGroup.Assimilate)."""
        if self._elem is None:
            self._set_status("Please select an element first.", Color.Orange)
            return
        if self._connector is None:
            self._set_status("Please pick an origin point first.", Color.Orange)
            return
        try:
            angle_val = float(self._txt_angle.Text.strip())
        except Exception:
            self._set_status("Invalid angle. Enter a number.", Color.Orange)
            return

        self._btn_create.Enabled = False
        self._set_status("Creating elbow...", C_PURPLE_LT)
        try:
            elbow = create_elbow_from_connector(self._elem, self._connector, angle_val)
            if elbow is not None:
                self._set_status(
                    "Elbow created!  ID: {0}".format(elbow.Id.IntegerValue), C_GREEN)
            else:
                self._set_status("Elbow returned None. Check log.", Color.Orange)
        except Exception as ex:
            err = str(ex)
            self._set_status("Error: " + err[:90], Color.Red)
            data.append(err)
        finally:
            self._btn_create.Enabled = True

    # --- Helper ---
    def _set_status(self, msg, color):
        self._lbl_status.Text      = msg
        self._lbl_status.ForeColor = color

# endregion

# region --- Run ---
form = CreateElbowForm()
Application.Run(form)
OUT = data
# endregion