"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

clr.AddReference("ProtoGeometry")

from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB import Line, XYZ, IntersectionResultArray, SetComparisonResult

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.UI import*

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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

class selectionFilter(ISelectionFilter):
    def __init__(self, category):
          self.category = category
    def AllowElement(self, element):
        if element.Category.Name == self.category:
            return True
        else:
            return False
def pickPipe():
    pipeFilter = selectionFilter('Pipes')
    pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter)
    pipe = doc.GetElement(pipeRef.ElementId)
    return pipe

def getNominalDiameter(pipe):
    """Lấy đường kính danh nghĩa của ống (đơn vị: feet)."""
    return pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()

def getFreeConnector(pipe, position=None):
    """Tìm connector chưa sử dụng của ống, ưu tiên gần vị trí nếu có."""
    connectors = list(pipe.ConnectorManager.Connectors)
    if position is None:
        for conn in connectors:
            if not conn.IsConnected:
                return conn
        return None
    
    min_dist = float('inf')
    closest_conn = None
    for conn in connectors:
        if not conn.IsConnected:
            dist = (conn.Origin - position).GetLength()
            if dist < min_dist:
                min_dist = dist
                closest_conn = conn
    return closest_conn

def findClosestConnectors(pipeA, pipeB):
    """Tìm cặp connector gần nhất giữa pipeA và pipeB."""
    connectorsA = list(pipeA.ConnectorManager.Connectors)
    connectorsB = list(pipeB.ConnectorManager.Connectors)
    
    min_dist = float('inf')
    best_connA = None
    best_connB = None
    
    for connA in connectorsA:
        if not connA.IsConnected:
            for connB in connectorsB:
                if not connB.IsConnected:
                    dist = (connA.Origin - connB.Origin).GetLength()
                    if dist < min_dist:
                        min_dist = dist
                        best_connA = connA
                        best_connB = connB
    
    return best_connA, best_connB, min_dist

def connectPipesWithSingleReducer(pipeA, pipeB):
    """Nối pipeA và pipeB bằng một reducer phù hợp."""
    # if not pipeA or not isinstance(pipeA, Pipe) or not pipeB or not isinstance(pipeB, Pipe):
    #     # print("Invalid pipe(s) selected.")
    #     return None
    
    # Lấy đường kính
    dn1 = getNominalDiameter(pipeA)
    dn2 = getNominalDiameter(pipeB)
    
    # Xác định ống lớn và nhỏ
    large_pipe = pipeA if dn1 >= dn2 else pipeB
    small_pipe = pipeB if dn1 >= dn2 else pipeA
    large_dia = max(dn1, dn2)
    small_dia = min(dn1, dn2)
    large_dia_mm = large_dia * 304.8
    small_dia_mm = small_dia * 304.8
    
    # Lấy PipeType và PipingSystem từ ống lớn
    pipe_type = large_pipe.PipeType
    piping_system = large_pipe.MEPSystem
    if not pipe_type:
        # print("Large pipe has no PipeType.")
        return None
    
    # Tìm cặp connector gần nhất
    connA, connB, distance = findClosestConnectors(pipeA, pipeB)
    if not connA or not connB:
        # print("No free connectors found for connection.")
        return None
    
    if distance > 1.0:  # Nếu khoảng cách lớn hơn 1 feet
        # print(f"Connectors are too far apart ({distance*304.8:.0f} mm). Please move pipes closer.")
        return None
    
    # Tìm FamilySymbol của reducer (large_dia - small_dia)
    reducer_family = None
    for family in FilteredElementCollector(doc).OfClass(Family):
        if family.FamilyCategory.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
            for symbol_id in family.GetFamilySymbolIds():
                symbol = doc.GetElement(symbol_id)
                if symbol.LookupParameter("Nominal Diameter 1") and symbol.LookupParameter("Nominal Diameter 2"):
                    nd1 = symbol.LookupParameter("Nominal Diameter 1").AsDouble()
                    nd2 = symbol.LookupParameter("Nominal Diameter 2").AsDouble()
                    if (abs(nd1 - large_dia) < 1e-6 and abs(nd2 - small_dia) < 1e-6) or \
                       (abs(nd1 - small_dia) < 1e-6 and abs(nd2 - large_dia) < 1e-6):
                        if symbol.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId() == pipe_type.Id:
                            reducer_family = symbol
                            break
        if reducer_family:
            break
    
    if not reducer_family:
        # print(f"No reducer found for DN{large_dia_mm:.0f}-DN{small_dia_mm:.0f}.")
        return None
    
    if not reducer_family.IsActive:
        reducer_family.Activate()
    
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    # Đặt reducer tại điểm giữa hai connector
    connection_point = (connA.Origin + connB.Origin) / 2
    reducer = doc.Create.NewFamilyInstance(connection_point, reducer_family, large_pipe, StructuralType.NonStructural)
    
    # Kết nối large_pipe và small_pipe với reducer
    try:
        reducer_conns = list(reducer.MEPSystem.ConnectorManager.Connectors)
        # Kết nối large_pipe
        for conn in reducer_conns:
            if abs(conn.Radius - large_dia / 2) < 1e-6:  # Connector với large_dia
                connA.ConnectTo(conn)
                break
        # Kết nối small_pipe
        for conn in reducer_conns:
            if abs(conn.Radius - small_dia / 2) < 1e-6:  # Connector với small_dia
                connB.ConnectTo(conn)
                break
    except Exception as ex:
        # print(f"Failed to connect pipes: {str(ex)}")
        TransactionManager.Instance.TransactionTaskDone()
        return None
    
    TransactionManager.Instance.TransactionTaskDone()
    return reducer

pipeA = pickPipe()
pipeB = pickPipe()
reducer = connectPipesWithSingleReducer(pipeA, pipeB)


OUT = reducer