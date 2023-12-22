# from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

# # Lấy danh sách các ống trong mô hình
# pipes_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves)
# pipes = [pipe for pipe in pipes_collector]

# # Giả sử bạn muốn tìm connector gần nhất giữa ống thứ nhất và ống thứ hai trong danh sách ống
# pipe1 = pipes[0]
# pipe2 = pipes[1]

# # Tạo một biến lưu khoảng cách gần nhất
# closest_distance = float('inf')  # Gán giá trị lớn nhất cho khoảng cách ban đầu
# closest_connector_pipe1 = None
# closest_connector_pipe2 = None

# # Lặp qua connectors của ống thứ nhất
# for connector1 in pipe1.ConnectorManager.Connectors:
#     for connector2 in pipe2.ConnectorManager.Connectors:
#         distance = connector1.Origin.DistanceTo(connector2.Origin)
#         if distance < closest_distance:
#             closest_distance = distance
#             closest_connector_pipe1 = connector1
#             closest_connector_pipe2 = connector2

# # In ra thông tin của hai connector gần nhau nhất
# if closest_connector_pipe1 and closest_connector_pipe2:
#     print(f"Closest connectors found with distance: {closest_distance}")
#     print(f"Connector 1: {closest_connector_pipe1.Id}")
#     print(f"Connector 2: {closest_connector_pipe2.Id}")
# else:
#     print("No connectors found.")



"""Copyright by: vudinhduybm@gmail.com"""
import clr 
import sys 
import System   
clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
from Autodesk.Revit.DB import*
from Autodesk.Revit.DB.Structure import*

clr.AddReference("RevitAPIUI") 
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

#Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView


def getAllPipingSystemsInActiveView(doc):
	collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
		pipingSystemsName.append(systemName)
	return pipingSystemsName
	
def getBuiltInCategoryOTS(nameCategory): # Get Name of BuiltInCategory
    return [i for i in System.Enum.GetValues(BuiltInCategory) if i.ToString() == "OST_"+nameCategory or i.ToString() == nameCategory ]

categoriesIN = ['OST_PipeCurves','PipeFitting']
categories = getBuiltInCategoryOTS(categoriesIN)

for category in categories:
    categoriesFilter = []
    collector = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType()
    categoriesFilter.append(collector)
flat_categoriesFilter = [item for sublist in categoriesFilter for item in sublist]

filteredFams = []
filteredFams.append([[doc.GetElement(item.GetTypeId()) for item in sublist] for sublist in categoriesFilter])


OUT = filteredFams(doc)


eles = [FilteredElementCollector(doc , view.id).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElement()]

OUT = eles


###############
"""Copyright by: vudinhduybm@gmail.com"""
import clr 
import System
import math 
from System.Collections.Generic import *

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import*
from  Autodesk.Revit.UI.Selection import*

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
view = doc.ActiveView
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

diameterIN = IN[1]
inSystemName = IN[2]

def getAllPipingSystemsInActiveView(doc):
	collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		pipingSystemsName.append(systemName)
	return pipingSystemsName, pipingSystems

allPipingSystemsInActiveView = getAllPipingSystemsInActiveView(doc)
allPipingSystemName = allPipingSystemsInActiveView[0]
allPipingSystem = allPipingSystemsInActiveView[1]
idOfInSystem = allPipingSystemName.index(inSystemName)
desPipingSystem = allPipingSystem[idOfInSystem]

categories = [BuiltInCategory.OST_PipeCurves]
categoriesFilter = []
for c in categories:
    collector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType()
    categoriesFilter.append(collector)
flat_categoriesFilter = [item for sublist in categoriesFilter for item in sublist]

IDS = List[ElementId]()
for i in flat_categoriesFilter:
	IDS.Add(i.Id)

paramFilter1 = ParameterValueProvider(ElementId(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
reason1 = FilterNumericGreater()
i_Check = IN[1]/304.8
fRule1 = FilterDoubleRule(paramFilter1,reason1,i_Check,0.02)
filter1 = ElementParameterFilter(fRule1)
elems1 = FilteredElementCollector(doc).WherePasses(filter1).WhereElementIsNotElementType().ToElements()

combineFilter = LogicalAndFilter(desPipingSystem , fRule1)
collectorDesPipes = FilteredElementCollector(doc, IDS).WherePasses(combineFilter).ToElements()

# OUT = [desPipingSystem],[ inSystemName]
OUT = collectorDesPipes
