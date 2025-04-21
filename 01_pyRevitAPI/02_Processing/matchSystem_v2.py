import clr
import sys 
import System   
import math

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB.Plumbing import*
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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""
class selectionFilter(ISelectionFilter):
    def __init__(self,categories):
        self.categories = categories
    def AllowElement(self, element):
        eleName = element.Category.Name in self.categories
        return eleName
def pickPipeOrPipePart():
    categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
    desEle = doc.GetElement(ref.ElementId)
    return desEle
def pickPipeOrPipeParts():
	partList = []
	categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
	fittingFilter = selectionFilter(categories)
	refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
	for ref in refs:
		desEle = doc.GetElement(ref.ElementId)
		partList.append(desEle)
	return partList
def offsetPointAlongVector(point, vector, offsetDistance):
    direction = vector.ToRevitType().Normalize()
    scaledVector = direction.Multiply(offsetDistance).ToVector()
    offsetPoint = point.Add(scaledVector)
    return offsetPoint    
# def getPipeParameter(p):
#     if p.Category.Name == 'Pipes':
#         paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
        
#         paramPipeTypeId = p.GetTypeId()
#         paramPipeType = doc.GetElement(paramPipeTypeId)
        
#         paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
#         paramPipingSystem = doc.GetElement(paramPipingSystemId)

#         paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
#         paramLevel = doc.GetElement(paramLevelId)
#     elif p.Category.Name == 'Pipe Fittings':
#         paramDiameter = p.LookupParameter("Minimum Size").AsDouble() * 304.8
        
#         paramPipeTypeId = p.GetTypeId()
#         paramPipeType = doc.GetElement(paramPipeTypeId)
        
#         paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
#         paramPipingSystem = doc.GetElement(paramPipingSystemId)

#         paramLevelId = p.LookupParameter('LevelId').AsElementId()
#         paramLevel = doc.GetElement(paramLevelId)        
#         return [paramDiameter, paramPipingSystem, paramPipeType, paramLevel]
def findClosestLevel(z_coord):
    """Tìm Level gần nhất dựa trên tọa độ Z của fitting."""
    levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
    if not levels:
        return None
    
    closest_level = None
    min_diff = float('inf')
    
    for level in levels:
        elevation = level.Elevation  # Độ cao của Level (feet)
        diff = abs(elevation - z_coord)
        if diff < min_diff:
            min_diff = diff
            closest_level = level
    
    return closest_level
def getPipeParameter(p):
    """Lấy đường kính, reference level, system type, và pipe type của Pipe hoặc Pipe Fitting."""
    if not p or not isinstance(p, (Pipe, FamilyInstance)):
        return None
    
    # Dictionary để lưu kết quả
    params = {
        "Diameter_mm": None,
        "ReferenceLevel": None,
        "SystemType": None,
        "PipeType": None
    }
    
    # Xử lý Pipe
    if p.Category.Name == "Pipes":
        # Đường kính (mm)
        diameter_param = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if diameter_param:
            params["Diameter_mm"] = diameter_param.AsDouble() * 304.8
        
        # Reference Level
        level_id = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
        if level_id and level_id.IntegerValue != -1:
            level = doc.GetElement(level_id)
            params["ReferenceLevel"] = level.Name if level else "None"
        
        # System Type
        system_id = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        if system_id and system_id.IntegerValue != -1:
            system = doc.GetElement(system_id)
            params["SystemType"] = system.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() if system else "None"
        
        # Pipe Type
        pipe_type_id = p.GetTypeId()
        if pipe_type_id and pipe_type_id.IntegerValue != -1:
            pipe_type = doc.GetElement(pipe_type_id)
            params["PipeType"] = pipe_type
    
    # Xử lý Pipe Fitting
    # Xử lý Pipe Fitting
    elif p.Category.Name == "Pipe Fittings":
        # Đường kính (mm) - Lấy Nominal Diameter lớn nhất
        nd1 = p.LookupParameter("Nominal Diameter 1")
        nd2 = p.LookupParameter("Nominal Diameter 2")
        diameter = 0.0
        if nd1 and nd1.HasValue:
            diameter = max(diameter, nd1.AsDouble())
        if nd2 and nd2.HasValue:
            diameter = max(diameter, nd2.AsDouble())
        params["Diameter_mm"] = diameter * 304.8 if diameter > 0 else None
        
        # Reference Level - Thử các phương pháp để lấy Level
        level_name = None
        # 1. Thử tham số tùy chỉnh Reference Level hoặc Level
        level_param = p.LookupParameter("Reference Level") or p.LookupParameter("Level")
        if level_param and level_param.HasValue:
            level_name = level_param.AsString()
        
        # 2. Nếu không tìm thấy, lấy Level gần nhất dựa trên tọa độ Z
        if not level_name:
            location = p.Location
            if isinstance(location, LocationPoint):
                z_coord = location.Point.Z
                closest_level = findClosestLevel(z_coord)
                level_name = closest_level.Name if closest_level else "None"
            else:
                level_name = "None"  # Không xác định được tọa độ
        
        params["ReferenceLevel"] = level_name
        
        # System Type
        system_id = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        if system_id and system_id.IntegerValue != -1:
            system = doc.GetElement(system_id)
            params["SystemType"] = system.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() if system else "None"
        
        # Pipe Type
        # pipe_type_id = p.GetTypeId()
        # if pipe_type_id and pipe_type_id.IntegerValue != -1:
        #     pipe_type = doc.GetElement(pipe_type_id)
        #     params["PipeType"] = pipe_type.Name if pipe_type else "None"
    
    return params



"""___"""
#pick pipe part. 
basePart = pickPipeOrPipePart()
basePartParam = getPipeParameter(basePart)
# listSetPart = pickPipeOrPipeParts() 
# #get base part piping system parameter
# basePart_PipingSystem_check = basePart.LookupParameter('System Type')
# #NOTE: Duyệt qua từng phần tử trong list set part. Sau đó vẽ ống ảo.
# for ele in listSetPart:
#     conns = list(ele.ConnectorManager.Connectors)
#     connsXYZ = [c.Origin for c in conns]
#     connsPoint = [c.ToPoint() for c in connsXYZ]
#     for con in conns:
#         if con.IsConnected == True:
#             con.Disconnect
#         else: pass
#         offsetVector = connsXYZ[0]-connsXYZ[1]
#         newPoint = offsetPointAlongVector(connsXYZ[0].ToPoint(),offsetVector,100 )

# if basePart_PipingSystem_check != None:
# 	basePart_PipingSystem = basePart_PipingSystem_check.AsElementId()
# 	for part in listSetPart:
# 		updatedPart = []
# 		part_PipingSystem = part.LookupParameter('System Type')
# 		TransactionManager.Instance.EnsureInTransaction(doc)
# 		part_PipingSystem.Set(basePart_PipingSystem)
# 		TransactionManager.Instance.TransactionTaskDone()
# 		updatedPart.append(part)
# else : pass

# OUT = listSetPart, updatedPart

OUT = basePartParam