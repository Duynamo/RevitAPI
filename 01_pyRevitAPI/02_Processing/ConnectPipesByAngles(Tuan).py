#region Library 
# import time
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
from Autodesk.Revit.DB import XYZ,Transaction,FilteredElementCollector,BuiltInCategory,BuiltInParameter,Plane,SketchPlane
from Autodesk.Revit.UI.Selection import ISelectionFilter

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

pipes_list=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
types=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()
level_ids=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
piping_system_types=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
#endregion
#region Global variables
numberOfFitting=0
angles_dict=[]
angle_list=[]
pipe1=None
pipe2=None
delete_list=[]
data=[]
connector_created=[]

#pipe's properties variables
system_id=None
type_id=None
level_id=None
diameter=None
A,B,C,D=None,None,None,None

#endregion
#region Mathematical method
def distance(a,b):
    """
    distance between 2 XYZ points
    """
    return math.sqrt((a.X-b.X)**2+(a.Y-b.Y)**2+(a.Z-b.Z)**2)
def PlaneFromThreePoints(A,B,C):
    """
    Get plane equation: Ax+By+Cz+D=0
    from 3 points A,B,C
    
    """
    vector_AB=B.Subtract(A)
    Vector_AC=C.Subtract(A)
    normal=vector_AB.CrossProduct(Vector_AC)
    normal=normal.Normalize()
    D=-normal.DotProduct(A)  
    A=normal.X
    B=normal.Y
    C=normal.Z
    return A,B,C,D
def solve_cosine_law(A,B,C,AC):

    
    """
    A,B,C is angle by degree
    length is opposite to B angle (AC)

    let D be projection of A on CB
    then AD othorgonal to DC and DB
    """
    sin_ACD=math.sin(C)
    AD=AC*sin_ACD

    sin_ABC=math.sin(math.pi-B)
    AB=AD/sin_ABC
    return AB
def Angle(A,B,C,D):
    """
    angle between AB and CD in degree
    """
    angle=0
    vectorAB=B.Subtract(A)
    vectorCD=D.Subtract(C)
    angle=vectorAB.AngleTo(vectorCD)
    angle=math.degrees(angle)
    return angle


#endregion
#region main Python Code

def verify():
    global A,B,C,D
    if distance(A.Origin,C.Origin) < distance(B.Origin,C.Origin):
        temp=A
        A=B
        B=temp
    if distance(B.Origin,D.Origin) < distance(B.Origin,C.Origin):
        temp=C
        C=D
        D=C
    return 0
# def select_pipes():
#     global pipe1, pipe2
#     pipe1=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
#     pipe1=doc.GetElement(pipe1)
#     pipe2=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
#     pipe2=doc.GetElement(pipe2)
#     return 0
def get_pipes_param(pipe1,pipe2):
    global system_id, type_id, level_id, diameter, A, B, C, D
    system_id=pipe1.MEPSystem.GetTypeId()
    type_id=pipe1.GetTypeId()
    level_id=pipe1.LevelId
    diameter=pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
    try:
        connectors1=pipe1.MEPModel.ConnectorManager.Connectors
        connectors2=pipe2.MEPModel.ConnectorManager.Connectors
    except:
        connectors1=pipe1.ConnectorManager.Connectors
        connectors2=pipe2.ConnectorManager.Connectors
    min_distance=10e6
    connectors=[]
    temp=[None,None]
    for con in connectors1:
        for con1 in connectors2:
            if distance(con.Origin,con1.Origin) <= min_distance :
                temp[0]=con
                temp[1]=con1
                min_distance=distance(con.Origin,con1.Origin)
    for con in connectors1:
        if con not in temp:
            connectors.append(con)
    connectors.append(temp[0])
    connectors.append(temp[1])
    for con in connectors2:
        if con not in temp:
            connectors.append(con)
    A=connectors[0] #nearconn 1 of pipe1
    B=connectors[1] #nearconn 1 of pipe2
    C=connectors[2] #nearconn 2 of pipe1
    D=connectors[3] #nearconn 2 of pipe2
    verify()
    return 0
def create_connectors(con,con1):
    """
    create connector between 2 input connectors
    """
    t=Transaction(doc,"create elbow")
    t.Start()
    connector=None
    error=""
    try:
       connector= DocumentManager.Instance.CurrentDBDocument.Create.NewElbowFitting(con,con1)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error if error is not  "" else connector
def disconnect_connectors(A,B):
    """
    disconnect 2 connectors
    """
    error=""
    t=Transaction(doc,"disconnect")
    t.Start()
    try:
        A.DisconnectFrom(B)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error 
def delete():
    """
    delete unuse element in delete_list
    """
    t=Transaction(doc,"delete")
    t.Start()
    error=""
    try:
        for i in delete_list:
            doc.Delete(i)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error
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
        element.Location.Move(vec)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error if error is "" else "moved"
def create_temp_pipe(A,B):
    error=""
    connect=None
    
    t=Transaction(doc,"rotate")
    t.Start()
    try:
        pipe3=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
													system_id,
													type_id,
                                                    # types[2].Id,
													level_id,
													B,
													A)
        param=pipe3.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        param.Set(diameter)
        connector3=pipe3.ConnectorManager.Connectors
        connect=None
        min_distance=10e6
        for con in connector3:
            if con.Origin.DistanceTo(B)<min_distance:
                connect=con
                min_distance=con.Origin.DistanceTo(B)
        delete_list.append(pipe3.Id)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    if error is not "":
        data.append(error)
    return error if error is not "" else connect
def get_vector_by_angle(A,B,C,angle):
    if distance(A,C) < distance(B,C):
        temp=A
        A=B
        B=temp
    """
    return a point lie on plane (A,B,C) and form with AB by angle and closest to C 
    angle in deg
    """
    angle=math.radians(angle)
    # A,B,C=A.Origin,B.Origin,C.Origin
    planeEquation=PlaneFromThreePoints(A,B,C)
    normal=XYZ(planeEquation[0],planeEquation[1],planeEquation[2]).Normalize()
    vector_AB=B.Subtract(A).Normalize() if distance(B,C) < distance(A,C) else A.Subtract(B).Normalize()
    vector_plane=normal.CrossProduct(vector_AB)
    vector1=vector_AB.Multiply(math.cos(angle)).Add(vector_plane.Multiply(math.sin(angle))).Normalize().Multiply(3.0)
    
    vector2=vector1.Negate()
    b1=B.Add(vector1)
    b2=B.Add(vector2)
    return b1 if distance(b1,C) < distance(b2,C) else b2
def create_independent_connector(A,B,C,D,angle):
    """
    create an independent connector between pipe AB and a pipe align with AB an angle closest to C
    """
    error=""
    v=get_vector_by_angle(A.Origin,B.Origin,C.Origin,angle)
    connect=create_temp_pipe(v,B.Origin)
    connector=create_connectors(B,connect)
    move(connector,B.Origin)
    for con in connector.MEPModel.ConnectorManager.Connectors:
        if con.IsConnectedTo(B):
            error=disconnect_connectors(con,B)
    return error if error is not "" else connector

def create_false_pipe(A,B):
    error=""
    connect=None
    pipe3=None
    t=Transaction(doc,"rotate")
    t.Start()
    try:
        pipe3=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
													system_id,
													type_id,
                                                    # types[2].Id,
													level_id,
													B,
													A)
        param=pipe3.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        param.Set(diameter)
        delete_list.append(pipe3.Id)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    if error is not "":
        data.append(error)
    return error if error is not "" else pipe3
def create_new_pipe(A,B):
    error=""
    t=Transaction(doc,"new_pipe")
    t.Start()
    new_p=None
    try:
        new_p=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
													system_id,
													type_id,
                                                    # types[2].Id,
													level_id,
													B,
													A)
        # diameter=pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
        param=new_p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        param.Set(diameter)
        # delete_list.append(new_p.Id)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return new_p
def create_pipe_angle(C,connector,angle):
    """
    create pipe from 2 points from the connector with angle
    """
    error=""
    connector2=None
    try:
        connect=[]
        for con in connector.MEPModel.ConnectorManager.Connectors:
            connect.append(con)
        if connect[0].Id==1:
            temp=connect[0]
            connect[0]=connect[1]
            connect[1]=temp
        v=get_vector_by_angle(connect[0].Origin,connector.Location.Point,C.Origin,angle)
        vector1=v.Subtract(connect[0].Origin).Normalize().Multiply(0.5)
        point=connect[0] if distance(C.Origin,connect[0].Origin) < distance(C.Origin,connect[1].Origin) else connect[1]
        vector2=vector1.Negate().Normalize().Multiply(0.5)
        vector=vector1 if distance(point.Origin.Add(vector1),C.Origin) < distance(point.Origin.Add(vector2),C.Origin) else vector2
        pipe=create_false_pipe(v.Add(vector),point.Origin.Add(vector))
        connector2=pipe.ConnectorManager.Connectors

        dis=10e6
        first,second=None,None
        for con1 in connect:
            for con2 in connector2:
                if distance(con1.Origin,con2.Origin)<dis:
                    first=con1
                    second=con2
                    dis=distance(con1.Origin,con2.Origin)
        connector3=create_connectors(second,first)
    except Exception as e:
        error=e
    return error if error != "" else connector3
def get_new_angle(connector,pipe,angle):
    """
    get angle that new pipe create with 2 connectors of previous fitting
    """
    # connector=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
    # pipe=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
    # connector=doc.GetElement(connector)
    # pipe=doc.GetElement(pipe)
    # angle=90
    pipe=doc.GetElement(pipe)
    new_angle=angle
    cons=[]
    pipes=[]
    for con in connector.MEPModel.ConnectorManager.Connectors:
        cons.append(con.Origin)
    vector_con=cons[0].Subtract(cons[1]).Normalize()
    for con in pipe.ConnectorManager.Connectors:
        pipes.append(con.Origin)
    """
    verify pipes and cons values

    that pipes[0] = cons[1]
    """
    for pipe in pipes:
        for con in cons:
            if pipe.X==con.X and pipe.Y==con.Y and pipe.Z==con.Z:
                if pipes.index(pipe)==0:
                    temp=pipes[0]
                    pipes[0]=pipes[1]
                    pipes[1]=temp
                if cons.index(con)==1:
                    temp=cons[0]
                    cons[0]=cons[1]
                    cons[1]=temp
                break
    vector_pipe=pipes[0].Subtract(pipes[1]).Normalize()
    # new_angle=vector_con.AngleTo(vector_pipe)
    new_angle=math.acos(vector_con.DotProduct(vector_pipe))
    new_angle=math.degrees(new_angle)
    angle=180-angle
    # new_angle=180-angle-new_angle
    result=180-angle-new_angle 
    if result <0:
        new_angle=180-new_angle
        result=180-angle-new_angle 
    return result
def replace_pipes():
    connect1,connect2=None,None
    for con in connector_created[0].MEPModel.ConnectorManager.UnusedConnectors:
        connect1=con
    for con in connector_created[1].MEPModel.ConnectorManager.UnusedConnectors:
        connect2=con
    t=Transaction(doc,"rotate")
    t.Start()
    try:
        pipe3=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
                                                    system_id,
                                                    type_id,
                                                    level_id,
													connect1.Origin,
													A.Origin)
        param=pipe3.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        param.Set(diameter)
        for con in pipe3.ConnectorManager.Connectors:
            if con.Origin.X == connect1.Origin.X and con.Origin.Y == connect1.Origin.Y and con.Origin.Z == connect1.Origin.Z:
                try:
                    con.ConnectTo(connect1)
                except:
                    pass
        pipe4=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
                                                    system_id,
                                                    type_id,
                                                    level_id,
													connect2.Origin,
													D.Origin)
        param=pipe4.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        for con in pipe4.ConnectorManager.Connectors:
            if con.Origin.X == connect2.Origin.X and con.Origin.Y == connect2.Origin.Y and con.Origin.Z == connect2.Origin.Z:
                try:
                    con.ConnectTo(connect2)
                except:
                    pass
        param.Set(diameter)
    except Exception as e:
        error=e
        data.append(error)
    finally:
        t.Commit()
    return 0
def check_connected(root,connector):
    """
    if 2 connector is connected, disconnect them
    """
    error=""
    connect1,connect2=None,None
    try:
        
        for con1 in root.MEPModel.ConnectorManager.Connectors:
            for con2 in connector.MEPModel.ConnectorManager.Connectors:
                if con1.Origin.X==con2.Origin.X and con1.Origin.Y==con2.Origin.Y and con1.Origin.Z==con2.Origin.Z:
                    connect1=con1
                    connect2=con2
                    t=Transaction(doc,"disconnect")
                    t.Start()
                    try:
                        con1.DisconnectFrom(con2)
                    except:
                        pass
                    t.Commit()
    except Exception as e:
        error=e
    return root,connector,connect1,connect2
def rotate__by_axis(connector,id):
    """
    Id==2: Big head
    Id==1: Small head
    """
    location=connector.Location.Point
    cons=[]
    error=""
    for con in connector.MEPModel.ConnectorManager.Connectors:
        cons.append(con)
    if cons[0].Id==id:
        temp=cons[0]
        cons[0]=cons[1]
        cons[1]=temp
    t=Transaction(doc,"rotate")
    t.Start()
    try:
        rotate_axis=cons[0].Origin.Subtract(location)
        rotate_line=Autodesk.Revit.DB.Line.CreateBound(cons[0].Origin,location)
        connector.Location.Rotate(rotate_line,math.radians(180))
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error
def rotate_last_con(connector):
    connect=None
    for con in connector.MEPModel.ConnectorManager.Connectors:
        if con.IsConnected==False:
            connect=con
    vector_pipe=D.Origin.Subtract(C.Origin).Normalize()
    vector_connector=connect.Origin.Subtract(connector.Location.Point).Normalize()
    if round(vector_pipe.AngleTo(vector_connector),2) >0.1:
        data.append("have rotated")
        rotate__by_axis(connector,connect.Id)
    return 0
def rotate_connector(root,connector):
    cons=[]
    connectors_params=[]
    FL="Flange Thickness"
    HL="HL"
    AN="Angle"
    pipe_angle=0
    params=connector.Parameters
    for param in params:
        if param.Definition.Name==FL:
            connectors_params.append(param.AsDouble())
        if param.Definition.Name==HL:
            connectors_params.append(param.AsDouble())
        if param.Definition.Name==AN:
            pipe_angle=param.AsDouble()
            pipe_angle=math.degrees(pipe_angle)
    c=check_connected(root,connector)
    con1,con2=c[2],c[3]
    cons=[]
    try:
        connectors=connector.MEPModel.ConnectorManager.Connectors
    except Exception as e:
        connectors=connector.ConnectorManager.Connectors
    for con in connectors:
        cons.append(con)
    if cons[0].Id == 2:
        temp=cons[0]
        cons[0]=cons[1]
        cons[1]=temp
    origin0=cons[0].Origin
    origin1=cons[1].Origin
    first_point=cons[0].Origin
    second_point=cons[1].Origin
    location_point=connector.Location.Point
    
    t=Transaction(doc,"rotate")
    t.Start()
    error=""
    try:
        plane_equation=PlaneFromThreePoints(first_point,second_point,location_point)
        normal_vector=XYZ(*plane_equation[:3])
     
        y_axis=first_point.Subtract(location_point).Normalize()
        z_axis=normal_vector
        x_axis=y_axis.CrossProduct(normal_vector).Normalize()
        x_axis=second_point.Subtract(location_point).Normalize()
        vector_endpoint=first_point.Subtract(location_point).Normalize().Multiply(sum(connectors_params))
        #rotate by x_axis by degree
        x_line=Autodesk.Revit.DB.Line.CreateBound(location_point,location_point.Add(x_axis))
        # rotated=connector.Location.Rotate(x_line,math.radians(180))
        geoplane=Plane.CreateByThreePoints(first_point,second_point,location_point)
        sketchPlane=SketchPlane.Create(doc,geoplane)
        #rotate by y_axis by degree
        y_line=Autodesk.Revit.DB.Line.CreateBound(location_point,location_point.Add(y_axis))
        
        # #big fit
        if con2.Id == con1.Id:
            rotated=connector.Location.Rotate(y_line,math.radians(180))
            #rotate by z_axis by 
            z_line=Autodesk.Revit.DB.Line.CreateBound(location_point,location_point.Add(z_axis))
            
            rotated=connector.Location.Rotate(z_line,math.radians(180-pipe_angle))
            if con2.Id==2:
                root_con=location_point.Subtract(origin1)
                move=connector.Location.Move(root_con.Subtract(location_point.Subtract(cons[0].Origin)).Negate())
            else:
            # #small fit
                root_con=location_point.Subtract(origin0)
                move=connector.Location.Move(root_con.Subtract(location_point.Subtract(cons[1].Origin)).Negate())   
        for con1 in root.MEPModel.ConnectorManager.Connectors:
            for con2 in connector.MEPModel.ConnectorManager.Connectors:
                if con1.Origin.X==con2.Origin.X and con1.Origin.Y==con2.Origin.Y and con1.Origin.Z==con2.Origin.Z:
                    try:
                        con1.ConnectTo(con2)
                    except Exception as e:
                        pass
        else:
            pass  
        
    except Exception as e:
        error=e
        data.append(error)
    finally :
        t.Commit()  
    return connector
def create_connector_set(A,B,C,D,angle_list):
    error=""
    global connector_created
    try:
        first_con=create_independent_connector(A,B,C,D,angle_list[0])
        temp_con=first_con
        connector=first_con
        connector_created.append(connector)
        temp_point=[]
        for i in range(1,len(angle_list)):
            for con in temp_con.MEPModel.ConnectorManager.Connectors:
                temp_point.append(con)
            # angle=get_new_angle(temp_con,delete_list[i-1],angle_list[i])
            # angle=angle % (math.pi)
            connector=create_pipe_angle(C,temp_con,angle_list[i])
            connector=rotate_connector(temp_con,connector)
            connector_created.append(connector)
            temp_con=connector
            
        move_param=find_move_vector(connector,A.Origin,B.Origin,C.Origin,D.Origin)
        move_vector=A.Origin.Subtract(B.Origin).Normalize()
        destination=connector.Location.Point.Add(move_vector.Multiply(move_param))
        delete()
        if len(angle_list) >1:
            rotate_last_con(connector)
        e=move(connector,destination)
        # create_connectors(connect,C)
        
    except Exception as e:
        error=e
    t=Transaction(doc,"delete")
    t.Start()
    doc.Delete(pipe1.Id)
    doc.Delete(pipe2.Id)
    t.Commit()
    replace_pipes()
    return error
def find_move_vector(connector,A,B,C,D):
    # closest=None
    # min_distance=10e6
    # for con in connector.MEPModel.ConnectorManager.Connectors:
    #     if distance(con.Origin,C) < min_distance:
    #         min_distance=distance(con.Origin,C)
    #         closest=con
    closest=connector.Location.Point
    vector_AB=B.Subtract(A)
    vector_CD=D.Subtract(C)
    vector_AC=C.Subtract(closest)
    AC=distance(C,closest)

    BAC=vector_AB.AngleTo(vector_AC)
    ABC=vector_CD.AngleTo(vector_AB)
    BCA=vector_AC.AngleTo(vector_CD)
    temp=solve_cosine_law(BAC,ABC,BCA,AC)
    return temp

def init():
    error=""
    global pipe2,pipe1
    # angle_list=[90,45]
    # select_pipes()
    # get_pipes_param(pipe1,pipe2)
    t=Transaction(doc,"erase p2")
    t.Start()
    doc.Delete(pipe2.Id)
    doc.Delete(pipe1.Id)
    t.Commit()
    C_location=D.Origin.Add(C.Origin).Multiply(0.5)
    B_location=A.Origin.Add(B.Origin).Multiply(0.5)
    pipe2=create_new_pipe(C_location,D.Origin)
    pipe1=create_new_pipe(B_location,A.Origin)
    get_pipes_param(pipe1,pipe2)
    
    connectors=None
    connectors=create_connector_set(A,B,C,D,angle_list)
    return error,connectors
#endregion
#region Winform
"""
Winform implementation
"""
class MainForm(Form):
    def __init__(self):
        self.InitializeComponent()
    def InitializeComponent(self):
        self._ElbowBox = System.Windows.Forms.GroupBox()
        self._PipesBox = System.Windows.Forms.GroupBox()
        self._AngleBox = System.Windows.Forms.GroupBox()
        self._NumberBox = System.Windows.Forms.GroupBox()
        self._RUN_BTN = System.Windows.Forms.Button()
        self._CANCEL_BTN = System.Windows.Forms.Button()
        self._label1 = System.Windows.Forms.Label()
        self._BasePipe_BTN = System.Windows.Forms.Button()
        self._SecondPipe_BTN = System.Windows.Forms.Button()
        self._label2 = System.Windows.Forms.Label()
        self._Angle_TXT = System.Windows.Forms.TextBox()
        self._label3 = System.Windows.Forms.Label()
        self._Number_TXT = System.Windows.Forms.TextBox()
        self._INPUT_TXT = System.Windows.Forms.Button()

        self._ElbowBox.SuspendLayout()
        self._PipesBox.SuspendLayout()
        self._AngleBox.SuspendLayout()
        self._NumberBox.SuspendLayout()
        self.SuspendLayout()
        # 
        # ElbowBox
        # 
        
        self._ElbowBox.Location = System.Drawing.Point(12, 192)
        self._ElbowBox.Name = "ElbowBox"
        self._ElbowBox.Size = System.Drawing.Size(508, 246)
        self._ElbowBox.TabIndex = 0
        self._ElbowBox.TabStop = False
        self._ElbowBox.Text = "Elbows"

        # 
        # PipesBox
        # 
        self._PipesBox.Controls.Add(self._SecondPipe_BTN)
        self._PipesBox.Controls.Add(self._BasePipe_BTN)
        self._PipesBox.Location = System.Drawing.Point(13, 70)
        self._PipesBox.Name = "PipesBox"
        self._PipesBox.Size = System.Drawing.Size(160, 116)
        self._PipesBox.TabIndex = 1
        self._PipesBox.TabStop = False
        self._PipesBox.Text = "Pipes"
        # 
        # AngleBox
        # 
        self._AngleBox.Controls.Add(self._Angle_TXT)
        self._AngleBox.Controls.Add(self._label2)
        self._AngleBox.Location = System.Drawing.Point(190, 70)
        self._AngleBox.Name = "AngleBox"
        self._AngleBox.Size = System.Drawing.Size(160, 116)
        self._AngleBox.TabIndex = 2
        self._AngleBox.TabStop = False
        self._AngleBox.Text = "Angle"

        # 
        # NumberBox
        # 
        self._NumberBox.Controls.Add(self._INPUT_TXT)
        self._NumberBox.Controls.Add(self._Number_TXT)
        self._NumberBox.Controls.Add(self._label3)
        self._NumberBox.Location = System.Drawing.Point(373, 70)
        self._NumberBox.Name = "NumberBox"
        self._NumberBox.Size = System.Drawing.Size(147, 116)
        self._NumberBox.TabIndex = 3
        self._NumberBox.TabStop = False
        self._NumberBox.Text = "Number"
        # 
        # RUN_BTN
        # 
        self._RUN_BTN.Location = System.Drawing.Point(309, 499)
        self._RUN_BTN.Name = "RUN_BTN"
        self._RUN_BTN.Size = System.Drawing.Size(92, 39)
        self._RUN_BTN.TabIndex = 4
        self._RUN_BTN.Text = "RUN"
        self._RUN_BTN.UseVisualStyleBackColor = True
        self._RUN_BTN.Click += self.RUN_BTNClick

        # 
        # CANCEL_BTN
        # 
        self._CANCEL_BTN.Location = System.Drawing.Point(428, 499)
        self._CANCEL_BTN.Name = "CANCEL_BTN"
        self._CANCEL_BTN.Size = System.Drawing.Size(92, 39)
        self._CANCEL_BTN.TabIndex = 5
        self._CANCEL_BTN.Text = "CANCEL"
        self._CANCEL_BTN.UseVisualStyleBackColor = True
        self._CANCEL_BTN.Click += self.CANCEL_BTNClick
        # 
        # label1
        # 
        self._label1.Font = System.Drawing.Font("メイリオ", 6.75, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._label1.ForeColor = System.Drawing.Color.CornflowerBlue
        self._label1.Location = System.Drawing.Point(13, 499)
        self._label1.Name = "label1"
        self._label1.Size = System.Drawing.Size(100, 23)
        self._label1.TabIndex = 6
        self._label1.Text = "@FVC"
        # 
        # BasePipe_BTN
        # 
        self._BasePipe_BTN.Location = System.Drawing.Point(25, 19)
        
        self._BasePipe_BTN.Name = "BasePipe_BTN"
        self._BasePipe_BTN.Size = System.Drawing.Size(110, 38)
        self._BasePipe_BTN.TabIndex = 0
        if pipe1 != None:
            self._BasePipe_BTN.Text=str(pipe1.Id)
        else:
            self._BasePipe_BTN.Text = "Base Pipe"
        self._BasePipe_BTN.UseVisualStyleBackColor = True
        self._BasePipe_BTN.Click += self.BasePipe_BTNClick
        # 
        # SecondPipe_BTN
        # 
        self._SecondPipe_BTN.Location = System.Drawing.Point(25, 63)
        self._SecondPipe_BTN.Name = "SecondPipe_BTN"
        self._SecondPipe_BTN.Size = System.Drawing.Size(110, 38)
        self._SecondPipe_BTN.TabIndex = 1
        if pipe2 != None:
            self._SecondPipe_BTN.Text=str(pipe2.Id)
        else:
            self._SecondPipe_BTN.Text = "Second Pipe"
        self._SecondPipe_BTN.UseVisualStyleBackColor = True
        self._SecondPipe_BTN.Click += self.SecondPipe_BTNClick
        # 
        # label2
        # 
        self._label2.Location = System.Drawing.Point(20, 22)
        self._label2.Name = "label2"
        self._label2.Size = System.Drawing.Size(121, 35)
        self._label2.TabIndex = 0
        self._label2.Text = "Angle between 2 pipes"
        self._label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        # 
        # Angle_TXT
        # 
        self._Angle_TXT.Location = System.Drawing.Point(0, 63)
        self._Angle_TXT.Name = "Angle_TXT"
        self._Angle_TXT.ReadOnly = True
        self._Angle_TXT.Size = System.Drawing.Size(160, 20)
        self._Angle_TXT.TabIndex = 1
        # 
        # label3
        # 
        self._label3.Location = System.Drawing.Point(6, 20)
        self._label3.Name = "label3"
        self._label3.Size = System.Drawing.Size(74, 49)
        self._label3.TabIndex = 0
        self._label3.Text = "Number Of Elbow"
        # 
        # Number_TXT
        # 
        self._Number_TXT.Location = System.Drawing.Point(89, 19)
        self._Number_TXT.Name = "Number_TXT"
        self._Number_TXT.Size = System.Drawing.Size(49, 20)
        self._Number_TXT.TabIndex = 1
        # 
        # INPUT_TXT
        # 
        self._INPUT_TXT.Location = System.Drawing.Point(28, 63)
        self._INPUT_TXT.Name = "INPUT_TXT"
        self._INPUT_TXT.Size = System.Drawing.Size(110, 38)
        self._INPUT_TXT.TabIndex = 2
        self._INPUT_TXT.Text = "Create Elbows"
        self._INPUT_TXT.UseVisualStyleBackColor = True
        self._INPUT_TXT.Click += self.INPUT_TXTClick
        # 
        
        # 
        # MainForm
        # 
        self.ClientSize = System.Drawing.Size(532, 550)
        self.Controls.Add(self._label1)
        self.Controls.Add(self._CANCEL_BTN)
        self.Controls.Add(self._RUN_BTN)
        self.Controls.Add(self._NumberBox)
        self.Controls.Add(self._AngleBox)
        self.Controls.Add(self._PipesBox)
        self.Controls.Add(self._ElbowBox)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.Name = "MainForm"
        self.Text = "Create Elbows"
        self.Load += self.MainFormLoad
        self._ElbowBox.ResumeLayout(False)
        self._ElbowBox.PerformLayout()
        self._PipesBox.ResumeLayout(False)
        self._AngleBox.ResumeLayout(False)
        self._AngleBox.PerformLayout()
        self._NumberBox.ResumeLayout(False)
        self._NumberBox.PerformLayout()
        self.ResumeLayout(False)


    def MainFormLoad(self, sender, e):
        pass

    def TextBox1TextChanged(self, sender, e):
        pass

    

    def CANCEL_BTNClick(self, sender, e):
        self.DialogResult=DialogResult.Cancel
        self.Close()

    def INPUT_TXTClick(self, sender, e):
        global numberOfFitting
        global angles_dict
        numberOfFitting=self._Number_TXT.Text
        numberOfFitting=int(numberOfFitting)
        input=numberOfFitting-1
        # loop to create pipe
        for i in range(0,numberOfFitting-1):
            self._label4 = System.Windows.Forms.Label()
            self._textBox1 = System.Windows.Forms.TextBox()
            self._comboBox1 = System.Windows.Forms.ComboBox()
            # label4
            # 
            self._label4.Location = System.Drawing.Point(14,20+ 20*i*2)
            self._label4.Name = "label4"
            self._label4.Size = System.Drawing.Size(100, 29)
            self._label4.TabIndex = 0
            self._label4.Text = "Angle of Elbow {}".format(i+1)
            # 
            # textBox1
            # 
            self._textBox1.Location = System.Drawing.Point(140,20+ 20*i*2)
            self._textBox1.Name = "textBox1"
            self._textBox1.Size = System.Drawing.Size(100, 20)
            self._textBox1.TabIndex = 1
            self._textBox1.TextChanged += self.TextBox1TextChanged
            
            self._ElbowBox.Controls.Add(self._textBox1)
            self._ElbowBox.Controls.Add(self._label4)
            angles_dict.append(self._textBox1)
            self._textBox1.KeyDown += self.TextBox1KeyDown
            
        self._lastangle = System.Windows.Forms.Label()
        self._lasttxtbox = System.Windows.Forms.TextBox()
        # label4
        # 
        self._lastangle.Location = System.Drawing.Point(14,20+ 20*(numberOfFitting-1)*2)
        self._lastangle.Name = "label4"
        self._lastangle.Size = System.Drawing.Size(100, 29)
        self._lastangle.TabIndex = 0
        self._lastangle.Text = "Angle of Elbow {}".format((numberOfFitting-1)+1)
        # 
        # textBox1
        # 
        self._lasttxtbox.Location = System.Drawing.Point(140,20+ 20*(numberOfFitting-1)*2)
        self._lasttxtbox.Name = "textBox1"
        self._lasttxtbox.Size = System.Drawing.Size(100, 20)
        self._lasttxtbox.TabIndex = 1
        self._ElbowBox.Controls.Add(self._lasttxtbox)
        self._ElbowBox.Controls.Add(self._lastangle)
        pass
    def BasePipe_BTNClick(self, sender, e):
        global pipe1
        pipe1=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
        pipe1=doc.GetElement(pipe1)

    def SecondPipe_BTNClick(self, sender, e):
        global pipe2
        pipe2=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
        pipe2=doc.GetElement(pipe2)
        global anglePipes,A,B,C,D
        get_pipes_param(pipe1,pipe2)
        anglePipes=Angle(A.Origin,B.Origin,C.Origin,D.Origin)
        self._Angle_TXT.Text=str(anglePipes)
    def RUN_BTNClick(self, sender, e):
        global angle_list
        for i in angles_dict:
            angle_list.append(float(i.Text))
        if self._lasttxtbox.Text!="":
            for i in range(len(angle_list)/2):
                angle_list.pop()

        return_value=init()
        self.close()
    def close(self):
        self.Close()
    
    def TextBox1KeyDown(self,sender, e):
        if e.KeyCode==System.Windows.Forms.Keys.Enter:
            if len(angles_dict) ==0:
                pass
            try:
                for i in angles_dict:
                    angle_list.append(float(i.Text))
                
                self._lasttxtbox.Text=str(abs(anglePipes-float(sum(angle_list))))
                angle_list.append(abs(float(anglePipes-float(sum(angle_list)))))
            except Exception as e:
                pass
        pass
            
#endregion
form=MainForm()
form.TopMost=True
Application.Run(form)

OUT=data



