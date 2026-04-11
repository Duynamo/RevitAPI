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
    return accessories
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
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self._cmb_accessories = System.Windows.Forms.ComboBox()
        self._label1 = System.Windows.Forms.Label()
        self._label2 = System.Windows.Forms.Label()
        self._btn_pickPipe = System.Windows.Forms.Button()
        self._btn_pickPoint = System.Windows.Forms.Button()
        self._label3 = System.Windows.Forms.Label()
        self._txt_offset = System.Windows.Forms.TextBox()
        self._btn_ok = System.Windows.Forms.Button()
        self._btn_cancel = System.Windows.Forms.Button()
        self.SuspendLayout()
        get_accessories()
        # 
        # cmb_accessories
        # 
        self._cmb_accessories.FormattingEnabled = True
        self._cmb_accessories.Location = System.Drawing.Point(125, 13)
        self._cmb_accessories.Name = "cmb_accessories"
        self._cmb_accessories.DataSource=accessories_name
        self._cmb_accessories.Size = System.Drawing.Size(447, 31)
        self._cmb_accessories.TabIndex = 0
        # 
        # label1
        # 
        self._label1.Location = System.Drawing.Point(13, 13)
        self._label1.Name = "label1"
        self._label1.Size = System.Drawing.Size(106, 28)
        self._label1.TabIndex = 1
        self._label1.Text = "Accessories"
        self._label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        # 
        # label2
        # 
        self._label2.Font = System.Drawing.Font("メイリオ", 8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._label2.ForeColor = System.Drawing.Color.Blue
        self._label2.Location = System.Drawing.Point(13, 534)
        self._label2.Name = "label2"
        self._label2.Size = System.Drawing.Size(47, 18)
        self._label2.TabIndex = 2
        self._label2.Text = "@FVC"
        # 
        # btn_pickPipe
        # 
        self._btn_pickPipe.Location = System.Drawing.Point(13, 100)
        self._btn_pickPipe.Name = "btn_pickPipe"
        self._btn_pickPipe.Size = System.Drawing.Size(190, 63)
        self._btn_pickPipe.TabIndex = 3
        self._btn_pickPipe.Text = "Pick Pipe"
        self._btn_pickPipe.UseVisualStyleBackColor = True
        self._btn_pickPipe.Click += self.Btn_pickPipeClick
        # 
        # btn_pickPoint
        # 
        self._btn_pickPoint.Location = System.Drawing.Point(382, 100)
        self._btn_pickPoint.Name = "btn_pickPoint"
        self._btn_pickPoint.Size = System.Drawing.Size(190, 63)
        self._btn_pickPoint.TabIndex = 4
        self._btn_pickPoint.Text = "Pick Start Point"
        self._btn_pickPoint.UseVisualStyleBackColor = True
        self._btn_pickPoint.Click += self.Btn_pickPointClick
        # 
        # label3
        # 
        self._label3.Location = System.Drawing.Point(13, 233)
        self._label3.Name = "label3"
        self._label3.Size = System.Drawing.Size(190, 24)
        self._label3.TabIndex = 5
        self._label3.Text = "Offset from Start Point"
        # 
        # txt_offset
        # 
        self._txt_offset.CausesValidation = False
        self._txt_offset.Location = System.Drawing.Point(382, 233)
        self._txt_offset.Name = "txt_offset"
        self._txt_offset.Size = System.Drawing.Size(190, 30)
        self._txt_offset.TabIndex = 6
        # 
        # btn_ok
        # 
        self._btn_ok.Location = System.Drawing.Point(360, 485)
        self._btn_ok.Name = "btn_ok"
        self._btn_ok.Size = System.Drawing.Size(88, 39)
        self._btn_ok.TabIndex = 7
        self._btn_ok.Text = "OK"
        self._btn_ok.UseVisualStyleBackColor = True
        self._btn_ok.Click += self.Btn_okClick
        # 
        # btn_cancel
        # 
        self._btn_cancel.Location = System.Drawing.Point(484, 485)
        self._btn_cancel.Name = "btn_cancel"
        self._btn_cancel.Size = System.Drawing.Size(88, 39)
        self._btn_cancel.TabIndex = 8
        self._btn_cancel.Text = "Cancel"
        self._btn_cancel.UseVisualStyleBackColor = True
        self._btn_cancel.Click += self.Btn_cancelClick
        # 
        # MainForm
        # 
        self.AutoScaleDimensions = System.Drawing.SizeF(96, 96)
        self.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        self.ClientSize = System.Drawing.Size(584, 561)
        self.Controls.Add(self._btn_cancel)
        self.Controls.Add(self._btn_ok)
        self.Controls.Add(self._txt_offset)
        self.Controls.Add(self._label3)
        self.Controls.Add(self._btn_pickPoint)
        self.Controls.Add(self._btn_pickPipe)
        self.Controls.Add(self._label2)
        self.Controls.Add(self._label1)
        self.Controls.Add(self._cmb_accessories)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.Font = System.Drawing.Font("メイリオ", 11.25, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self.Name = "MainForm"
        self.Text = "Accessory"
        self.ResumeLayout(False)
        self.PerformLayout()


    def Btn_okClick(self, sender, e):
        global name,offset
        name=self._cmb_accessories.Text
        offset=float(self._txt_offset.Text)/304.8
        
        init()

        self.Close()

    def Btn_cancelClick(self, sender, e):
        self.Close()

    def Btn_pickPipeClick(self, sender, e):
        global pipe
        pipe=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
        # self.Close()
    def Btn_pickPointClick(self, sender, e):
        global picked_p,connector

        connector=uidoc.Selection.PickPoint(Autodesk.Revit.UI.Selection.ObjectSnapTypes.Nearest , "connector")

        picked_p=connector
        # self.Close()

#endregion
form=MainForm()
form.TopMost=True
Application.Run(form)
OUT=data