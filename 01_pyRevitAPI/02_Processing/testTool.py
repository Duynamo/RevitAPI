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

def pickFittings():
    pipeFilter = selectionFilter( 'Pipe Fittings')
    pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter)
    pipe = doc.GetElement(pipeRef.ElementId)
    return pipe

def pickPoint():
    try:
        point = uidoc.Selection.PickPoint("Please pick a point on Pipe A")
        return point
    except Exception as e:
        # print(f"Error picking point: {str(e)}")
        return None
def get_fittings_by_pipe_type(doc, pipe_type_id, system_type_id=None):
    """
    Trả về danh sách tất cả Pipe Fitting instances (FamilyInstance) thuộc PipeType của pipeA.
    
    Args:
        doc: Autodesk.Revit.DB.Document
        pipe_type_id: ElementId của PipeType (từ pipeA.GetTypeId())
        system_type_id: ElementId của SystemType (tùy chọn, để lọc thêm)
    
    Returns:
        list: Danh sách các FamilyInstance (Pipe Fittings) thuộc PipeType
    """
    try:
        from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter
        
        fittings = []
        
        # Lấy tất cả Pipe Fittings trong dự án
        collector = FilteredElementCollector(doc).OfClass(FamilyInstance).OfCategoryId(BuiltInCategory.OST_PipeFitting)
        
        for fitting in collector:
            # Kiểm tra PipeType thông qua Connector
            connectors = list(fitting.MEPModel.ConnectorManager.Connectors) if fitting.MEPModel else []
            if not connectors:
                continue
            
            # Kiểm tra PipeType của Connector
            pipe_type_match = False
            for conn in connectors:
                if hasattr(conn, "PipeType"):
                    conn_pipe_type_id = conn.PipeType
                    if conn_pipe_type_id == pipe_type_id:
                        pipe_type_match = True
                        break
            
            # Kiểm tra thêm SystemType (nếu cung cấp)
            system_type_match = True
            if system_type_id:
                fitting_system_type_id = fitting.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
                system_type_match = (fitting_system_type_id == system_type_id)
            
            # Nếu khớp PipeType và (nếu có) SystemType, thêm vào danh sách
            if pipe_type_match and system_type_match:
                fittings.append(fitting)
        
        return fittings
    
    except Exception as e:
        # print(f"Error: {str(e)}")
        return []
def get_tee_symbol(pipe_type_id):
    """
    Tìm FamilySymbol của Tee phù hợp với PipeType.
    
    Args:
        pipe_type_id: ElementId của PipeType
    
    Returns:
        FamilySymbol: Tee symbol hoặc None
    """
    try:
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategoryId(BuiltInCategory.OST_PipeFitting)
        for symbol in collector:
            # Kiểm tra xem có phải Tee (dựa trên tên hoặc số connector sau khi kích hoạt)
            symbol.Activate()
            temp_instance = doc.Create.NewFamilyInstance(XYZ(0, 0, 0), symbol, None, None, Autodesk.Revit.DB.Structure.StructuralType.NonStructural)
            if temp_instance:
                conns = list(temp_instance.MEPModel.ConnectorManager.Connectors)
                doc.Delete(temp_instance.Id)
                if len(conns) == 3:  # Tee có 3 connector
                    return symbol
        # print("Error: No suitable Tee FamilySymbol found.")
        return None
    except Exception as e:
        # print(f"Error finding Tee symbol: {str(e)}")
        return None

def place_tee_at_point(pipeA, point_x):
    """
    Đặt Tee fitting tại point_x trên pipeA.
    
    Args:
        pipeA: Pipe chính
        point_x: Điểm XYZ trên pipeA
    
    Returns:
        bool: True nếu thành công
    """
    try:
        # Lấy đường cong của pipeA
        pipeA_curve = pipeA.Location.Curve
        P1 = pipeA_curve.GetEndPoint(0)
        P2 = pipeA_curve.GetEndPoint(1)
        
        # Kiểm tra point_x có trên pipeA không
        projection = pipeA_curve.Project(point_x)
        if projection.Distance > 1e-6:
            # print(f"Error: Point {point_x} is not on Pipe A (distance: {projection.Distance}).")
            return False
        point_x = projection.XYZPoint
        
        # Kiểm tra point_x không trùng P1, P2
        if point_x.IsAlmostEqualTo(P1, 1e-6) or point_x.IsAlmostEqualTo(P2, 1e-6):
            # print("Error: Point X is at the endpoint of Pipe A.")
            return False
        
        # Lấy thông tin pipeA
        pipe_type_id = pipeA.GetTypeId()
        system_type_id = pipeA.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        level_id = pipeA.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
        
        # Tìm Tee FamilySymbol
        tee_symbol = get_tee_symbol(pipe_type_id)
        if not tee_symbol:
            return False
        
        # Bắt đầu transaction
        with Transaction(doc, "Place Tee on Pipe A") as t:
            t.Start()
            
            # Chia pipeA tại point_x
            # pipeA1: P1 -> point_x
            # pipeA2: point_x -> P2
            pipeA.Location.Curve = Line.CreateBound(P1, point_x)
            pipeA2 = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc, system_type_id, pipe_type_id, level_id, point_x, P2)
            
            # Tạo Tee tại point_x
            tee = doc.Create.NewFamilyInstance(point_x, tee_symbol, None, None, Autodesk.Revit.DB.Structure.StructuralType.NonStructural)
            
            # Điều chỉnh hướng Tee
            pipeA_dir = (P2 - P1).Normalize()
            tee_conns = list(tee.MEPModel.ConnectorManager.Connectors)
            if len(tee_conns) != 3:
                print("Error: Tee does not have exactly 3 connectors.")
                doc.Delete(tee.Id)
                doc.Delete(pipeA2.Id)
                t.RollBack()
                return False
            
            # Tìm 2 connector chính của Tee (song song pipeA)
            main_conns = []
            for conn in tee_conns:
                conn_dir = conn.CoordinateSystem.BasisZ
                if abs(conn_dir.DotProduct(pipeA_dir)) > 0.9:  # Song song pipeA
                    main_conns.append(conn)
            
            if len(main_conns) != 2:
                print("Error: Could not identify main connectors of Tee.")
                doc.Delete(tee.Id)
                doc.Delete(pipeA2.Id)
                t.RollBack()
                return False
            
            # Tìm connector của pipeA và pipeA2 tại point_x
            pipeA_conns = list(pipeA.ConnectorManager.Connectors)
            pipeA2_conns = list(pipeA2.ConnectorManager.Connectors)
            
            pipeA_conn = min(
                pipeA_conns,
                key=lambda c: c.Origin.DistanceTo(point_x),
                default=None
            )
            pipeA2_conn = min(
                pipeA2_conns,
                key=lambda c: c.Origin.DistanceTo(point_x),
                default=None
            )
            
            if not pipeA_conn or not pipeA2_conn:
                print("Error: Could not find connectors at point X.")
                doc.Delete(tee.Id)
                doc.Delete(pipeA2.Id)
                t.RollBack()
                return False
            
            # Kết nối Tee với pipeA và pipeA2
            pipeA_conn.ConnectTo(main_conns[0])
            pipeA2_conn.ConnectTo(main_conns[1])
            
            t.Commit()
        
        # print(f"Placed Tee at {point_x} on Pipe A.")
        return True, tee_symbol
    
    except Exception as e:
        # print(f"Error: {str(e)}")
        return False

def rotate_tee_and_pipe(pipeA, teeA, angle_degrees):
    """
    Xoay teeA quanh trục của pipeA một góc angle_degrees.
    
    Args:
        pipeA: Pipe chính
        teeA: Tee fitting trên pipeA
        angle_degrees: Góc xoay (độ)
    
    Returns:
        bool: True nếu thành công
    """
    try:
        if not teeA:
            # print("Error: No Tee fitting provided.")
            return False
        
        # Lấy trục xoay từ pipeA
        pipeA_curve = pipeA.Location.Curve
        P1 = pipeA_curve.GetEndPoint(0)
        P2 = pipeA_curve.GetEndPoint(1)
        axis_dir = (P2 - P1).Normalize()
        
        # Tâm xoay là vị trí teeA
        teeA_location = teeA.Location
        if not hasattr(teeA_location, "Point"):
            # print("Error: TeeA has no location point.")
            return False
        center = teeA_location.Point
        
        # Tạo trục xoay
        axis_line = Line.CreateUnbound(center, axis_dir)
        
        # Góc xoay (radian)
        angle_radians = math.radians(angle_degrees)
        
        # Bắt đầu transaction
        with Transaction(doc, "Rotate Tee") as t:
            t.Start()
            
            # Xoay teeA
            ElementTransformUtils.RotateElement(
                doc, teeA.Id, axis_line, angle_radians
            )
            
            t.Commit()
        
        # print(f"Rotated teeA {angle_degrees} degrees around pipeA.")
        return True
    
    except Exception as e:
        # print(f"Error: {str(e)}")
        return False

pipeA = pickPipe()
listConns = list(pipeA.ConnectorManager.Connectors)
listConnsXYZ = [a.Origin for a in listConns]
cenPoint = XYZ( (listConnsXYZ[0].X + listConnsXYZ[1].X )/2, 
                (listConnsXYZ[0].Y + listConnsXYZ[1].Y )/2,
                (listConnsXYZ[0].Z + listConnsXYZ[1].Z )/2)
# newTee = place_tee_at_point(pipeA, cenPoint.ToPoint())

# teeA = pickFittings()
# pipeB = pickPipe()
angle = 22.5    
pipeATypeId = pipeA.GetTypeId()
pipingSystemId =pipeA.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
paramPipingSystem = doc.GetElement(pipingSystemId)
pipeFittings = get_fittings_by_pipe_type(doc, pipeATypeId)
# ro = rotate_tee_and_pipe(pipeA , teeA , -(90-angle))
# OUT = '', cenPoint.ToPoint(), listConnsXYZ[0].ToPoint(), listConnsXYZ[1].ToPoint()
OUT = pipeFittings