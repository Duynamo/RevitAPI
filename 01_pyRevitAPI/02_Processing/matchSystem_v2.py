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

def getPipeParameter(p):
    params = {
        "Diameter_mm": None,
        "ReferenceLevel": None,
        "SystemType": None,
        "PipeType": None
    }
    if p.Category.Name == "Pipes":
        paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
        params['Diameter_mm'] = paramDiameter
        paramPipeTypeId = p.GetTypeId()
        paramPipeType = doc.GetElement(paramPipeTypeId)
        params['PipeType'] = paramPipeType
        paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        paramPipingSystem = doc.GetElement(paramPipingSystemId)
        params['SystemType'] = paramPipingSystem
        paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
        paramLevel = doc.GetElement(paramLevelId)
        params['ReferenceLevel'] = paramLevel

    if p.Category.Name == "Pipe Fittings":
        paramDiameter = p.LookupParameter('Maximum Size').AsDouble() * 304.8
        params['Diameter_mm'] = paramDiameter
        paramPipeTypeId = p.GetTypeId()
        paramPipeType = doc.GetElement(paramPipeTypeId)
        params['PipeType'] = paramPipeType
        paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        paramPipingSystem = doc.GetElement(paramPipingSystemId)
        params['SystemType'] = paramPipingSystem
        paramLevelId = p.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId()
        paramLevel = doc.GetElement(paramLevelId)
        params['ReferenceLevel'] = paramLevel        
    return params

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

"""___"""
#pick pipe part. 
basePart = pickPipeOrPipePart()
basePartParam = getPipeParameter(basePart)
listSetPart = pickPipeOrPipeParts() 
pipe_settings = PipeSettings.GetPipeSettings(doc)
# #NOTE: Duyệt qua từng phần tử trong list set part. Sau đó vẽ ống ảo.
for ele in listSetPart:
    conns = list(ele.MEPModel.ConnectorManager.Connectors)
    connsXYZ = [c.Origin for c in conns]
    connsPoint = [c.ToPoint() for c in connsXYZ]
    # for con in conns:
        # if con.IsConnected == True:
        #     con.Disconnect
        # else: pass
        # offsetVector = connsXYZ[0]-connsXYZ[1]
        # newPoint = offsetPointAlongVector(connsXYZ[0].ToPoint(),offsetVector,100 )
    if ele.Category.Name == 'Pipe Fittings':
        baseDiameter = basePartParam['Diameter_mm']
        pipe_segment = None
        tempSeg =  FilteredElementCollector(doc).OfClass(PipeSegment).ToElements()  
        # Kiểm tra danh sách kích thước khả dụng
        # matching_segments = []
        # for segment in tempSeg:
        #     size_exists = False
        #     # Tìm tất cả Pipe sử dụng PipeSegment
        #     pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
        #     for pipe in pipes:
        #         pipe_type = doc.GetElement(pipe.GetTypeId())
        #         diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
        #         if abs(diameter - baseDiameter) < 1e-6:  # So sánh với độ chính xác nhỏ
        #             size_exists = True
        #             break
            
        #     if size_exists:
        #         matching_segments.append(segment)
            
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

# OUT = drawPipeFromFitting(basePart, pipe_length_mm=100)

OUT =  FilteredElementCollector(doc).OfClass(PipeSegment).ToElements() 
