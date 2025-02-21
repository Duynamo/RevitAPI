"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
from email.errors import FirstHeaderLineIsContinuationDefect
from os import error, pipe
import select



from sqlite3 import connect
from unittest import skip
from xml.dom import minidom
import clr
import sys 
import System   
import math
import collections
from math import cos, e,sin,tan,radians

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import ISelectionFilter
clr.AddReference("System") 
from System.Collections.Generic import List

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# clr.AddReference("System.Windows.Forms")
# clr.AddReference("System.Drawing")
# clr.AddReference("System.Windows.Forms.DataVisualization")

#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion


#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
#endregion
def PlaneFromThreePoints(A,B,C):
    vector_AB=B.Subtract(A)
    Vector_AC=C.Subtract(A)
    normal=vector_AB.CrossProduct(Vector_AC)
    normal=normal.Normalize()
    D=-normal.DotProduct(A)  
    A=normal.X
    B=normal.Y
    C=normal.Z
    return A,B,C,D
def solveThreeEquations(planeEquation,angleEquation,distanceEquation):
    """
    planeEquation: Ax+By+Cz+D=0
    angleEquation: Ax+By+Cz=cos(angle)
    distanceEquation: (x-a)**2+(y-b)**2+(z-c)**2=d**2
    """
    x,y,z=1.0,1.0,1.0
    xtol=10e-6
    max_loop=100
    for _ in range(max_loop):
        x_prev,y_prev,z_prev=x,y,z
        x=(planeEquation[1]*y+planeEquation[2]*z+planeEquation[3])/planeEquation[0]
        y=(angleEquation[3]-angleEquation[0]*x-angleEquation[2]*z)/(angleEquation[1])
        z=math.sqrt(distanceEquation[3]**2-(x-distanceEquation[0])**2-(y-distanceEquation[1])**2)+distanceEquation[2]
        fz=(x-distanceEquation[0])**2+(y-distanceEquation[1])**2+(z-distanceEquation[2])**2-distanceEquation[3]**2
        dfz_dz=2*(z-2)

        if dfz_dz==0:
            break
        z=z-fz/dfz_dz
        if abs(x-x_prev) < xtol and abs(y-y_prev)<xtol and abs(z-z_prev)<xtol:
            break

    return x,y,z
pipe1=None
pipe2=None
delete_list=[]
data=[]
def distance(a,b):
    return math.sqrt((a.X-b.X)**2+(a.Y-b.Y)**2+(a.Z-b.Z)**2)
pipes_list=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
types=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()
level_ids=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
piping_system_types=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()

def select_pipes():
    global pipe1
    global pipe2
    pipe1=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
    pipe1=doc.GetElement(pipe1)
    pipe2=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
    pipe2=doc.GetElement(pipe2)
    
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
    return temp,connectors

def orthogonalVec(A,B,C):
    """
    get vector lie on plane (ABC) and orthogonal to vectorAB
    """
    A,B,C=A.Origin,B.Origin,C.Origin
    planeEquation=PlaneFromThreePoints(A,B,C)
    planeVector=XYZ(planeEquation[0],planeEquation[1],planeEquation[2])
    vector_AB=B.Subtract(A)
    orthogonal=planeVector.CrossProduct(vector_AB)
    orthogonal=orthogonal.Normalize().Multiply(3.0)
    B1=B.Add(orthogonal)
    B2=B.Add(orthogonal.Negate())
    return B1 if B1.DistanceTo(C)<B2.DistanceTo(C) else B2
def create_connectors(con,con1):
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
    if error is not "":
        data.append(error)
    return error if error is not  "" else connector
def create_temp_pipe(A,B):
    error=""
    connect=None
    
    t=Transaction(doc,"rotate")
    t.Start()
    try:
        pipe3=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
													piping_system_types[1].Id,
													pipe1.GetTypeId(),
                                                    # types[2].Id,
													level_ids[1].Id,
													B,
													A)
        diameter=pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
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
def create_false_pipe(A,B):
    error=""
    connect=None
    pipe3=None
    t=Transaction(doc,"rotate")
    t.Start()
    try:
        pipe3=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
													piping_system_types[1].Id,
													pipe1.GetTypeId(),
                                                    # types[2].Id,
													level_ids[1].Id,
													B,
													A)
        diameter=pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
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
    except:
        error=e
    finally:
        t.Commit()
    if error is not "":
        data.append(error)
    return error if error is "" else "moved"
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
def last_connect(connector,A,B,C,D):
    """
    create last connector to destination pipe
    """
    error=""
    for con in connector.MEPModel.ConnectorManager.Connectors:
        if con.Origin not  in (A.Origin,B.Origin,C.Origin,D.Origin):
            error=create_connectors(con,C)
    return error
def get_vector_by_angle(A,B,C,angle):
    """
    lie on plane (A,B,C) and form with AB by angle and closest to C
    angle in deg
    """
    angle=math.radians(angle)
    # A,B,C=A.Origin,B.Origin,C.Origin
    planeEquation=PlaneFromThreePoints(A,B,C)
    normal=XYZ(planeEquation[0],planeEquation[1],planeEquation[2]).Normalize()
    vector_AB=B.Subtract(A).Normalize()
    vector_plane=normal.CrossProduct(vector_AB)
    vector1=vector_AB.Multiply(math.cos(angle)).Add(vector_plane.Multiply(math.sin(angle))).Normalize().Multiply(3.0)
    vector2=vector1.Negate()
    b1=B.Add(vector1)
    b2=B.Add(vector2)
    return b1 if distance(b1,C)<distance(b2,C) else b2
def get_connector():
    error=""
    collector=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()
    pipeType=None
    for type in collector:
        if type.Id==pipe1.GetTypeId():
            pipeType=type
            break
    # t=Transaction(doc,"update fitting")
    # t.Start()
    # try:

    # except Exception as e:
    #     error=e
    # finally:
    #     t.Commit()
    list=pipeType.get_Parameter( 
          BuiltInParameter.RBS_CURVETYPE_DEFAULT_ELBOW_PARAM).AsValueString()
    return collector
def two_connect(A,B,C,D):
    error= ""
    try:
        orthogonal=orthogonalVec(A,B,C)
        connect3=create_temp_pipe(orthogonal,B.Origin)
        connector=create_connectors(B,connect3)
        move(connector,B.Origin)
        last_connect(connector,A,B,C,D)
    except Exception as e:
        error=e
    return error if error is not "" else "Connection created"
def three_connect(A,B,C,D,first_angle,second_angle):

    error=""
    try:
        v=get_vector_by_angle(A.Origin,B.Origin,C.Origin,first_angle)
        connect=create_temp_pipe(v,B.Origin)
        connector=create_connectors(B,connect)
        move(connector,B.Origin)
        v2=get_vector_by_angle(D.Origin,C.Origin,B.Origin,second_angle)
        connect2=create_temp_pipe(v2,C.Origin)
        connector2=create_connectors(C,connect2)
        move(connector2,C.Origin)
        con1,con2=None,None
        for con in connector.MEPModel.ConnectorManager.UnusedConnectors:
            con1=con
        for con in connector2.MEPModel.ConnectorManager.UnusedConnectors:
            con2=con
        create_connectors(con1,con2)
    except Exception as e:
        error= e
    return error if error is not "" else "Connection created"
def disconnect_connectors(A,B):
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
def create_independent_connector(A,B,C,D,angle):
    error=""
    v=get_vector_by_angle(A.Origin,B.Origin,C.Origin,angle)
    connect=create_temp_pipe(v,B.Origin)
    connector=create_connectors(B,connect)
    move(connector,B.Origin)
    for con in connector.MEPModel.ConnectorManager.Connectors:
        if con.IsConnectedTo(B):
            error=disconnect_connectors(con,B)
    return error if error is not "" else connector 
def create_independent_connectors(A,B,C,D,angles):
    connectors=[]
    for angle in angles:
        connector=create_independent_connector(A,B,C,D,angle)
        connectors.append(connector)
        data.append(connector)
    delete()
    return connectors
def connect_created_connector(connectors):
    error=""
    connect=[]
    for connector in connectors:
        for con in connector.MEPModel.ConnectorManager.Connectors:
             connect.append(con)
    t=Transaction(doc,"disconnect")
    t.Start()
    try:
        connectors[0].Location.Move(connect[2].Origin.Subtract(connect[3].Origin))
        bound=Line.CreateBound(connect[0].Origin,connect[1].Origin)
        connectors[0].Location.Rotate(bound,math.radians(180))
        connect[0].ConnectTo(connect[2])
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error,connect
# def create_pipe_angle(A,B,connector,angle):
# """
# create pipe from 2 points from the connector with angle
# """
#     error=""
#     connect=[]
#     for con in connector.MEPModel.ConnectorManager.Connectors:
#         connect.append(con)
#     v=get_vector_by_angle(A.Origin,B.Origin,connect[0].Origin,angle)
#     pipe=create_false_pipe(v,connect[0].Origin)
#     connector2=pipe.ConnectorManager.Connectors
#     # for con in connector2:
#     #     move(con,connect[1].Origin)
#     dis=10e6
#     first,second=None,None
#     for con1 in connect:
#         for con2 in pipe.ConnectorManager.Connectors:
#             if distance(con1.Origin,con2.Origin)<dis:
#                 first=con1
#                 second=con2
#                 dis=distance(con1.Origin,con2.Origin)
#     error=create_connectors(first,second)
#     return error
def get_pipe_data():
    error=""
    pipe=select_pipes()
    closest=pipe[0]
    connectors=pipe[1]
    A=connectors[0]
    B=connectors[1]
    C=connectors[2]
    D=connectors[3]
    angle_list=[45,90]
    connectors2=create_independent_connectors(A,B,C,D,angle_list )
    temp=connect_created_connector(connectors2)

    return error,temp

OUT=get_pipe_data(),data