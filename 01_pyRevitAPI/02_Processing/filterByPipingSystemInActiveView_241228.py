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

# def getAllPipingSystemsInActiveView(doc):
#     collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipingSystem)
#     pipingSystems = collector.ToElements()
#     pipingSystemsName = []
#     for system in pipingSystems:
#         systemName = system.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
#         pipingSystemsName = [systemName for systemName not in pipingSystemsName]
#     return pipingSystemsName

def getAllPipingSystemsInActiveView(doc):  
    # Collect all piping systems in the active view  
    collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipingSystem)  
    pipingSystems = collector.ToElements()  
    # Create a set to store unique system names  
    pipingSystemsName = set()  
    for system in pipingSystems:  
        systemName = system.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()  
        if systemName:  # Check if systemName is not None  
            pipingSystemsName.add(systemName)  # Add system name to the set  
    return list(pipingSystemsName)  # Convert the set back to a list before returning

def getBuiltInCategoryOST(nameCategory):
    # Get BuiltInCategory by name
    return [i for i in System.Enum.GetValues(BuiltInCategory) if i.ToString() == "OST_" + str(nameCategory) or i.ToString() == nameCategory]
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
"________________________________________________________________"
inSystemName = IN[0]
desPipes = []

"________________________________________________________________"

categoriesIN = ["OST_PipeCurves", "OST_PipeFitting", "OST_PipeAccessory"]
categories = [getBuiltInCategoryOST(category) for category in categoriesIN]
flat_categories = [item for sublist in categories for item in sublist]

categoriesFilter = []
for c in flat_categories:
    collector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType()
    categoriesFilter.append(collector)
lst1 = [[item for item in sublist] for sublist in categoriesFilter]

flat_categoriesFilter = flatten(lst1)

paramFilter1 = ParameterValueProvider(ElementId(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM))
reason1 = FilterStringEquals()
fRule1 = FilterStringRule(paramFilter1,reason1, inSystemName)
filter1 = ElementParameterFilter(fRule1)

filteredElements = []
for category in [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeAccessory]:
    collector = FilteredElementCollector(doc).OfCategory(category).WherePasses(filter1).WhereElementIsNotElementType()
    filteredElements.extend(collector.ToElements())
    
IDS = List[ElementId]()
for i in filteredElements:
	IDS.Add(i.Id)
# combineFilter = LogicalAndFilter(filter1 , filter2)
# desPipes = FilteredElementCollector(doc, IDS).WherePasses(filter1).WhereElementIsNotElementType().ToElements()

# TransactionManager.Instance.EnsureInTransaction(doc)
# isolatedEles = view.IsolateElementsTemporary(IDS)
# TransactionManager.Instance.TransactionTaskDone()

OUT = getAllPipingSystemsInActiveView(doc)