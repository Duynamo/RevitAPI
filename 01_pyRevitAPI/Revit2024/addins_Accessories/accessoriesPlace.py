"""
same method as demo1 to create accessory
cut pipe instead of delete to keep properties
"""
#region library
import clr
import System
import math
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import *
from System.Drawing import Point, Size, Font, FontStyle, GraphicsUnit, Color
from System.Windows.Forms import KeyPressEventHandler
import os 
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
clr.AddReference("RevitAPI") 
clr.AddReference("RevitAPIUI") 
import Autodesk
from Autodesk.Revit import * 
from Autodesk.Revit.DB import XYZ,Transaction,FilteredElementCollector,BuiltInCategory,BuiltInParameter,Plane,SketchPlane,FamilySymbol,TransactionGroup
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ISelectionFilter
from abc import *   
#endregion
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
pipes_list=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
types=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()
level_ids=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
piping_system_types=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
data=[]
error_list=[]
import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *
#region back end
data=[]
pipe=None
picked_p=None
connector=None
dict={}
system_id,type_id,level_id,diameter=None,None,None,None
A,B=None,None
access_con=[]
accessory_selected=None
accessories=[]
con1,con2=None,None
pipe1,pipe2,accessory_created=None,None,None
accessories_name=[]
all_accessories_name=[]  # keep full list for filtering
offset=0
name=None
def select():
    global pipe,picked_p,system_id,type_id,level_id,diameter,connector
    
    pipe=doc.GetElement(pipe)
    system_id=pipe.MEPSystem.GetTypeId()
    type_id=pipe.GetTypeId()
    level_id=pipe.LevelId
    diameter=pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
    
    # picked_p=connector
    # connector=doc.GetElement(connector)
    return 0
def get_connector():
    min=10e6
    global connector,A,B
    for con in pipe.ConnectorManager.Connectors:
        if con.Origin.DistanceTo(picked_p)<min:
            connector=con
            min=con.Origin.DistanceTo(picked_p)
    A=connector.Origin
    for con in pipe.ConnectorManager.Connectors:
        if con.Origin.IsAlmostEqualTo(A):
            continue
        else:
            B=con.Origin
    return 0
def delete_pipe():
    t=Transaction(doc,"delete")
    t.Start()
    try:
        doc.Delete(pipe.Id)
    except Exception as e:
        data.append(e)
    finally:
        t.Commit()
    return 0
def create_pipe(A,B):
    t=Transaction(doc,"create pipe")
    t.Start()
    try:
        pipe3=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
													system_id,
													type_id,
                                                    # types[2].Id,
													level_id,
													A,
													B)
        param=pipe3.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        param.Set(diameter)

    except Exception as e:
        data.append(e)
    finally:
        t.Commit()
    return pipe3
def create_new_pipe_1(offset):
    global pipe1
    vector_AB=B.Subtract(A).Normalize()
    C=A.Add(vector_AB.Multiply(offset))
    pipe1=create_pipe(A,C)
    return C
def create_new_pipe_2(offset):
    global pipe2
    D=access_con[0].Origin.DistanceTo(access_con[1].Origin)
    vector_AB=B.Subtract(A).Normalize()
    C=A.Add(vector_AB.Multiply(offset+D))
    pipe2=create_pipe(B,C)
    return C
def get_accessories():
    global accessories
    accessories=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).ToElements()
    for accessory in accessories:
        for param in accessory.Parameters:
            if param.Definition.Name=="Family Name" and param.AsValueString() is not None:
                accessories_name.append(str(param.AsValueString()))
                all_accessories_name.append(str(param.AsValueString()))
    return accessories

def extract_diameter(name):
    """Extract nominal diameter number from accessory name. e.g. 'V_仕切弁_D150_D2' -> 150"""
    import re
    # Match pattern _D<digits> or D<digits>
    matches = re.findall(r'[_\-]?D(\d+)', name)
    if matches:
        # Return the first (largest) diameter found
        diameters = [int(m) for m in matches]
        return max(diameters)
    return None

def filter_accessories_by_diameter(diameter_text):
    """Return filtered accessories_name list based on diameter text."""
    if not diameter_text or diameter_text.strip() == '':
        return list(all_accessories_name)
    try:
        d = int(diameter_text.strip())
        result = [n for n in all_accessories_name if extract_diameter(n) == d]
        return result
    except ValueError:
        return list(all_accessories_name)
def get_accessory_data(name):
    """
    Params: 
        + name: name of accessory (paramter.family&type.AsValueString())
        + A: 1st point(XYZ)
        + B: 2nd point(XYZ)
    output:
        + accessory at A and perpendicular to AB 

    """
    global accessory_selected
    
    accessory_selected=None
    for accessory in accessories:
        for param in accessory.Parameters:
            if param.Definition.Name=="Family Name":
                if param.AsValueString()==name:
                    accessory_params=accessory.Parameters
                    accessory_selected=accessory
    # data.append(accessory_params)
def create_accessory(A):
    global accessory_created,access_con
    t=Transaction(doc,"test")
    t.Start()
    try:
        cr_view=doc.ActiveView
        # accessory_created=doc.Create.NewFamilyInstance(curve,symbol,level_ids[0],StructuralType.Column)
        accessory_selected.Activate()
        accessory_created=doc.Create.NewFamilyInstance(A,accessory_selected,StructuralType.Column)
        doc.Delete(accessory_created.Id)
        accessory_created=doc.Create.NewFamilyInstance(A,accessory_selected,StructuralType.Column)
        for con in accessory_created.MEPModel.ConnectorManager.Connectors:
            access_con.append(con)
    except Exception as e:
        data.append(e)
    finally:
        t.Commit()
    return 0
def connect_connectors(con1,con2):
    t=Transaction(doc,"connect")
    t.Start()
    try:
        con1.ConnectTo(con2)
    except Exception as e:
        data.append(e)
    finally:
        t.Commit()
    return 0
def connect_connectors_of_system():
    for con1 in pipe1.ConnectorManager.Connectors:
        for con2 in accessory_created.MEPModel.ConnectorManager.Connectors:
            if con1.Origin.IsAlmostEqualTo(con2.Origin):
                connect_connectors(con1,con2)
    for con1 in pipe2.ConnectorManager.Connectors:
        for con2 in accessory_created.MEPModel.ConnectorManager.Connectors:
            if con1.Origin.IsAlmostEqualTo(con2.Origin):
                connect_connectors(con1,con2)
    return 0
A=XYZ(216.3,40.72,8.5)
B=XYZ(223.39,40.72,8.53)
# name="Cast Iron Gate Flange Internal Threaded Valve JIS 7.5K 150A"
def move(element,des):
    """
    move element to destination
    """
    root=element.Location.Point
    vec=des.Subtract(root).Negate()
    t=Transaction(doc,"move")
    t.Start()
    error=""
    try:
        element.Location.Move(vec.Negate())
    except Exception as e:
        data.append(e)
    finally:
        t.Commit()
    return data
def move_location():
    global con1,con2
    min_distance=10e6
    for connector in pipe1.ConnectorManager.Connectors:
        for connector2 in pipe2.ConnectorManager.Connectors:
            if connector.Origin.DistanceTo(connector2.Origin)<min_distance:
                min_distance=connector.Origin.DistanceTo(connector2.Origin)
                con1=connector
                con2=connector2
    loc=con1.Origin.Add(con2.Origin).Multiply(0.5)
    move(accessory_created,loc)
    return 0
def rotate_accessory():
    pipe_vec=con1.Origin.Subtract(con2.Origin).Normalize()
    access_vec=access_con[0].Origin.Subtract(access_con[1].Origin).Normalize()
    angle=abs(pipe_vec.AngleTo(access_vec))
    root=access_con[0].Origin.Add(access_con[1].Origin).Multiply(0.5)
    axis=access_vec.CrossProduct(pipe_vec).Normalize()
    endpoint=root.Add(axis.Multiply(2))
    t=Transaction(doc,"rotate")
    t.Start()
    try:
        bound=Autodesk.Revit.DB.Line.CreateBound(root,endpoint)
        accessory_created.Location.Rotate(bound,angle)
        cs=[]
        for connector in accessory_created.MEPModel.ConnectorManager.Connectors:
            cs.append(connector.Origin)

    except Exception as e:
        data.append(e)
    finally:
        t.Commit()
    return 0
def init():
    t=TransactionGroup(doc,"ga")
    t.Start()
    
    get_accessory_data(name)
    select()
    get_connector()
    delete_pipe()
    p1=create_new_pipe_1(offset)
    create_accessory(p1)
    p2=create_new_pipe_2(offset)
    move_location()
    rotate_accessory()
    connect_connectors_of_system()
    t.Assimilate()
    return 0
# init()

#endregion

#region winform
class MainForm(Form):
    def __init__(self):
        # store colours for event handlers
        self._COLOR_PURPLE    = System.Drawing.Color.FromArgb(255, 102,  51, 153)
        self._COLOR_PURPLE_DK = System.Drawing.Color.FromArgb(255,  72,  21, 120)
        self._COLOR_GREEN     = System.Drawing.Color.FromArgb(255,  34, 139,  34)
        self._COLOR_WHITE     = System.Drawing.Color.White
        self.InitializeComponent()

    def InitializeComponent(self):
        # ── colour palette ──────────────────────────────────────────────
        PURPLE    = System.Drawing.Color.FromArgb(255, 102,  51, 153)
        PURPLE_DK = System.Drawing.Color.FromArgb(255,  72,  21, 120)
        PURPLE_LT = System.Drawing.Color.FromArgb(255, 237, 224, 255)
        WHITE     = System.Drawing.Color.White
        GRAY_BG   = System.Drawing.Color.FromArgb(255, 245, 244, 250)
        GRAY_BDR  = System.Drawing.Color.FromArgb(255, 200, 185, 220)
        TEXT_DK   = System.Drawing.Color.FromArgb(255,  40,  20,  70)

        # ── fonts ────────────────────────────────────────────────────────
        F_MAIN  = System.Drawing.Font(u"メイリオ", 10,   System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        F_LBL   = System.Drawing.Font(u"メイリオ",  9.5, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        F_BTN   = System.Drawing.Font(u"メイリオ", 10,   System.Drawing.FontStyle.Bold,    System.Drawing.GraphicsUnit.Point, 128)
        F_SMALL = System.Drawing.Font(u"メイリオ",  8,   System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        F_HDR   = System.Drawing.Font(u"メイリオ", 12,   System.Drawing.FontStyle.Bold,    System.Drawing.GraphicsUnit.Point, 128)

        # ── layout constants ─────────────────────────────────────────────
        FW    = 520   # form width
        PAD   = 14
        LW    = 110   # label column width
        CX    = PAD + LW + 6   # control x start
        CW    = FW - CX - PAD  # control width
        RH    = 36              # row height
        RG    = 10              # row gap
        HDR   = 44              # header height

        Y1 = HDR + RG            # accessories row
        Y2 = Y1 + RH + RG        # DN filter row
        Y3 = Y2 + RH + RG + 4   # pick buttons row
        Y4 = Y3 + 50  + RG + 4  # offset row
        Y5 = Y4 + RH + RG        # separator
        Y6 = Y5 + 12             # OK / Cancel
        FH = Y6 + 44 + 10        # form height

        self.SuspendLayout()
        get_accessories()

        # ── auto-fit DropDownWidth ────────────────────────────────────────
        g = self.CreateGraphics()
        ddw = CW
        for item in accessories_name:
            w = int(g.MeasureString(str(item), F_MAIN).Width) + 24
            if w > ddw:
                ddw = w
        g.Dispose()

        # ── header panel ─────────────────────────────────────────────────
        self._pnl_header = System.Windows.Forms.Panel()
        self._pnl_header.Location  = System.Drawing.Point(0, 0)
        self._pnl_header.Size      = System.Drawing.Size(FW, HDR)
        self._pnl_header.BackColor = PURPLE_DK

        lbl_hdr = System.Windows.Forms.Label()
        lbl_hdr.Text      = u"  Accessory Placer"
        lbl_hdr.Font      = F_HDR
        lbl_hdr.ForeColor = WHITE
        lbl_hdr.Location  = System.Drawing.Point(0, 0)
        lbl_hdr.Size      = System.Drawing.Size(FW, HDR)
        lbl_hdr.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        self._pnl_header.Controls.Add(lbl_hdr)

        # ── ROW 1 : Accessories ComboBox ──────────────────────────────────
        self._label1 = System.Windows.Forms.Label()
        self._label1.Text      = u"Accessories"
        self._label1.Font      = F_LBL
        self._label1.ForeColor = TEXT_DK
        self._label1.Location  = System.Drawing.Point(PAD, Y1 + 7)
        self._label1.Size      = System.Drawing.Size(LW, RH)

        self._cmb_accessories = System.Windows.Forms.ComboBox()
        self._cmb_accessories.FormattingEnabled = True
        self._cmb_accessories.DropDownStyle     = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._cmb_accessories.Font              = F_MAIN
        self._cmb_accessories.ForeColor         = TEXT_DK
        self._cmb_accessories.BackColor         = WHITE
        self._cmb_accessories.Location          = System.Drawing.Point(CX, Y1)
        self._cmb_accessories.Size              = System.Drawing.Size(CW, RH)
        self._cmb_accessories.DropDownWidth     = ddw
        self._cmb_accessories.TabIndex          = 0
        for item in accessories_name:
            self._cmb_accessories.Items.Add(item)
        if self._cmb_accessories.Items.Count > 0:
            self._cmb_accessories.SelectedIndex = 0

        # ── ROW 2 : DN Filter ─────────────────────────────────────────────
        self._label_diameter = System.Windows.Forms.Label()
        self._label_diameter.Text      = u"DN Filter"
        self._label_diameter.Font      = F_LBL
        self._label_diameter.ForeColor = PURPLE
        self._label_diameter.Location  = System.Drawing.Point(PAD, Y2 + 7)
        self._label_diameter.Size      = System.Drawing.Size(LW, RH)

        self._txt_diameter = System.Windows.Forms.TextBox()
        self._txt_diameter.Font        = F_MAIN
        self._txt_diameter.ForeColor   = TEXT_DK
        self._txt_diameter.BackColor   = PURPLE_LT
        self._txt_diameter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        self._txt_diameter.Location    = System.Drawing.Point(CX, Y2)
        self._txt_diameter.Size        = System.Drawing.Size(CW, RH)
        self._txt_diameter.TabIndex    = 1
        self._txt_diameter.TextChanged += self.Txt_diameterTextChanged

        # ── ROW 3 : Pick Buttons ──────────────────────────────────────────
        BPW = int((FW - PAD * 3) / 2)

        self._btn_pickPipe = System.Windows.Forms.Button()
        self._btn_pickPipe.Text      = u"Pick Pipe"
        self._btn_pickPipe.Font      = F_BTN
        self._btn_pickPipe.ForeColor = WHITE
        self._btn_pickPipe.BackColor = PURPLE
        self._btn_pickPipe.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        self._btn_pickPipe.FlatAppearance.BorderSize = 0
        self._btn_pickPipe.Location  = System.Drawing.Point(PAD, Y3)
        self._btn_pickPipe.Size      = System.Drawing.Size(BPW, 50)
        self._btn_pickPipe.TabIndex  = 2
        self._btn_pickPipe.Cursor    = System.Windows.Forms.Cursors.Hand
        self._btn_pickPipe.Click    += self.Btn_pickPipeClick

        self._btn_pickPoint = System.Windows.Forms.Button()
        self._btn_pickPoint.Text      = u"Pick Start Point"
        self._btn_pickPoint.Font      = F_BTN
        self._btn_pickPoint.ForeColor = WHITE
        self._btn_pickPoint.BackColor = PURPLE
        self._btn_pickPoint.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        self._btn_pickPoint.FlatAppearance.BorderSize = 0
        self._btn_pickPoint.Location  = System.Drawing.Point(PAD * 2 + BPW, Y3)
        self._btn_pickPoint.Size      = System.Drawing.Size(BPW, 50)
        self._btn_pickPoint.TabIndex  = 3
        self._btn_pickPoint.Cursor    = System.Windows.Forms.Cursors.Hand
        self._btn_pickPoint.Click    += self.Btn_pickPointClick

        # ── ROW 4 : Offset TextBox ────────────────────────────────────────
        self._label3 = System.Windows.Forms.Label()
        self._label3.Text      = u"Offset (mm)"
        self._label3.Font      = F_LBL
        self._label3.ForeColor = TEXT_DK
        self._label3.Location  = System.Drawing.Point(PAD, Y4 + 7)
        self._label3.Size      = System.Drawing.Size(LW, RH)

        self._txt_offset = System.Windows.Forms.TextBox()
        self._txt_offset.CausesValidation = False
        self._txt_offset.Font        = F_MAIN
        self._txt_offset.ForeColor   = TEXT_DK
        self._txt_offset.BackColor   = WHITE
        self._txt_offset.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        self._txt_offset.Location    = System.Drawing.Point(CX, Y4)
        self._txt_offset.Size        = System.Drawing.Size(CW, RH)
        self._txt_offset.TabIndex    = 4

        # ── separator ─────────────────────────────────────────────────────
        self._sep = System.Windows.Forms.Panel()
        self._sep.BackColor = GRAY_BDR
        self._sep.Location  = System.Drawing.Point(PAD, Y5)
        self._sep.Size      = System.Drawing.Size(FW - PAD * 2, 1)

        # ── OK / Cancel buttons ───────────────────────────────────────────
        BW = 90
        BH = 36
        BGP = 10
        BCX = FW - PAD - BW
        BOX = BCX - BGP - BW

        self._btn_ok = System.Windows.Forms.Button()
        self._btn_ok.Text      = u"OK"
        self._btn_ok.Font      = F_BTN
        self._btn_ok.ForeColor = WHITE
        self._btn_ok.BackColor = PURPLE_DK
        self._btn_ok.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        self._btn_ok.FlatAppearance.BorderSize = 0
        self._btn_ok.Location  = System.Drawing.Point(BOX, Y6)
        self._btn_ok.Size      = System.Drawing.Size(BW, BH)
        self._btn_ok.TabIndex  = 5
        self._btn_ok.Cursor    = System.Windows.Forms.Cursors.Hand
        self._btn_ok.Click    += self.Btn_okClick

        self._btn_cancel = System.Windows.Forms.Button()
        self._btn_cancel.Text      = u"Cancel"
        self._btn_cancel.Font      = F_BTN
        self._btn_cancel.ForeColor = PURPLE_DK
        self._btn_cancel.BackColor = PURPLE_LT
        self._btn_cancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        self._btn_cancel.FlatAppearance.BorderColor = GRAY_BDR
        self._btn_cancel.FlatAppearance.BorderSize  = 1
        self._btn_cancel.Location  = System.Drawing.Point(BCX, Y6)
        self._btn_cancel.Size      = System.Drawing.Size(BW, BH)
        self._btn_cancel.TabIndex  = 6
        self._btn_cancel.Cursor    = System.Windows.Forms.Cursors.Hand
        self._btn_cancel.Click    += self.Btn_cancelClick

        # ── @FVC label ────────────────────────────────────────────────────
        self._label2 = System.Windows.Forms.Label()
        self._label2.Text      = u"@FVC"
        self._label2.Font      = F_SMALL
        self._label2.ForeColor = PURPLE
        self._label2.Location  = System.Drawing.Point(PAD, Y6 + (BH - 14) // 2)
        self._label2.Size      = System.Drawing.Size(60, 18)
        self._label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

        # ── form properties ───────────────────────────────────────────────
        self.AutoScaleDimensions = System.Drawing.SizeF(96, 96)
        self.AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Dpi
        self.ClientSize          = System.Drawing.Size(FW, FH)
        self.FormBorderStyle     = System.Windows.Forms.FormBorderStyle.FixedSingle
        self.StartPosition       = System.Windows.Forms.FormStartPosition.CenterScreen
        self.BackColor           = GRAY_BG
        self.Font                = F_MAIN
        self.Name                = "MainForm"
        self.Text                = u"Accessory Placer"
        self.MaximizeBox         = False

        self.Controls.Add(self._pnl_header)
        self.Controls.Add(self._label1)
        self.Controls.Add(self._cmb_accessories)
        self.Controls.Add(self._label_diameter)
        self.Controls.Add(self._txt_diameter)
        self.Controls.Add(self._btn_pickPipe)
        self.Controls.Add(self._btn_pickPoint)
        self.Controls.Add(self._label3)
        self.Controls.Add(self._txt_offset)
        self.Controls.Add(self._sep)
        self.Controls.Add(self._btn_ok)
        self.Controls.Add(self._btn_cancel)
        self.Controls.Add(self._label2)

        self.ResumeLayout(False)
        self.PerformLayout()

    # ── event handlers ────────────────────────────────────────────────────
    def Txt_diameterTextChanged(self, sender, e):
        filtered = filter_accessories_by_diameter(self._txt_diameter.Text)
        current  = self._cmb_accessories.Text
        self._cmb_accessories.Items.Clear()
        for item in filtered:
            self._cmb_accessories.Items.Add(item)
        if current in filtered:
            self._cmb_accessories.SelectedItem = current
        elif self._cmb_accessories.Items.Count > 0:
            self._cmb_accessories.SelectedIndex = 0

    def Btn_okClick(self, sender, e):
        global name, offset
        name   = self._cmb_accessories.Text
        try:
            offset = float(self._txt_offset.Text) / 304.8
        except Exception:
            offset = 0.0
        try:
            init()
        except Exception as ex:
            data.append(ex)
        # Reset UI for next operation — do NOT close
        self._reset_ui()

    def Btn_cancelClick(self, sender, e):
        self.Close()

    def _reset_ui(self):
        """Reset pick state: button back to purple, clear globals."""
        global pipe, picked_p, connector, access_con, accessory_created, pipe1, pipe2
        pipe = None
        picked_p = None
        connector = None
        access_con = []
        accessory_created = None
        pipe1 = None
        pipe2 = None
        # Reset Pick Pipe button colour
        self._btn_pickPipe.BackColor = self._COLOR_PURPLE
        self._btn_pickPipe.Text = u"Pick Pipe"
        # Reset Pick Point button colour
        self._btn_pickPoint.BackColor = self._COLOR_PURPLE
        self._btn_pickPoint.Text = u"Pick Start Point"

    def Btn_pickPipeClick(self, sender, e):
        global pipe
        try:
            pipe = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "s")
            # Change button colour to green to indicate pipe selected
            self._btn_pickPipe.BackColor = self._COLOR_GREEN
            self._btn_pickPipe.Text = u"Pipe Selected "
            # Auto-fill DN filter from pipe diameter
            pipe_elem = doc.GetElement(pipe)
            dn_mm = pipe_elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
            dn_mm_int = int(round(dn_mm * 304.8))  # feet -> mm
            # temporarily disconnect TextChanged to avoid re-filtering during set
            self._txt_diameter.TextChanged -= self.Txt_diameterTextChanged
            self._txt_diameter.Text = str(dn_mm_int)
            self._txt_diameter.TextChanged += self.Txt_diameterTextChanged
            # trigger filter manually
            self.Txt_diameterTextChanged(None, None)
        except Exception:
            pass  # user cancelled pick

    def Btn_pickPointClick(self, sender, e):
        global picked_p, connector
        try:
            connector = uidoc.Selection.PickPoint(Autodesk.Revit.UI.Selection.ObjectSnapTypes.Nearest, "connector")
            picked_p  = connector
            self._btn_pickPoint.BackColor = self._COLOR_GREEN
            self._btn_pickPoint.Text = u"Point Selected"
        except Exception:
            pass  # user cancelled pick

#endregion
form = MainForm()
form.TopMost = True
Application.Run(form)
OUT = data

