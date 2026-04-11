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


'''___'''
def allPipesInActiveView():
	pipesCollector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
	return pipesCollector

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
def getPipeParameter(p):
    paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()*304.8
    paramDiameter1 = p.LookupParameter('Diameter').AsDouble()*304.8
    paramPipeTypeId = p.GetTypeId()
    paramPipeType = doc.GetElement(paramPipeTypeId)
    pipeTypeName = paramPipeType.LookupParameter("Type Name").AsString()
    paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    # paramPipingSystem = doc.GetElement(paramPipingSystemId)
    pipingSystemName = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
    paramLevel = p.LookupParameter('Reference Level').AsValueString()
    paramLength = p.LookupParameter('Length').AsDouble()*304.8
    return str(paramLength),str(paramDiameter1),pipeTypeName,pipingSystemName, paramLevel
'''___'''

# #region set pipe fitting, pipeAccessory FVC_Length
# categoriesIN = ["OST_PipeFitting", "OST_PipeAccessory"]
# categories = [getBuiltInCategoryOST(category) for category in categoriesIN]
# flat_categories = [item for sublist in categories for item in sublist]
# categoriesFilter = []
# for c in flat_categories:
#     collector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType().ToElements()
#     categoriesFilter.append(collector)
# eleList = [item for sublist in categoriesFilter for item in sublist]
# for item in eleList:
#     elementSystemTypeCheck = item.LookupParameter("System Type")
#     if elementSystemTypeCheck is not None:
#         elementSystemType = elementSystemTypeCheck.AsValueString()
#         element_FVC_PipingSystem = item.LookupParameter("FVC_Piping System")
#         TransactionManager.Instance.EnsureInTransaction(doc)
#         element_FVC_PipingSystem.Set(elementSystemType)
#         TransactionManager.Instance.TransactionTaskDone()
#     else:
#         pass

# #endregion

# #region set revit pipe FVC_Length
# pipeCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()

# for pipe in pipeCollector:
#     pipeParam = getPipeParameter(pipe)
#     # pipeLength = pipe.LookupParameter("Length").AsDouble()
#     # pipePipingSystem = pipe.LookupParameter("").A
#     if pipeParam is not None:
#         pipe_FVC_Length = pipe.LookupParameter("FVC_Length")
#         pipe_FVC_Diameter = pipe.LookupParameter("FVC_Diameter")
#         pipe_FVC_PipeType = pipe.LookupParameter("FVC_Pipe Type")
#         pipe_FVC_PipingSystem = pipe.LookupParameter("FVC_Piping System")
#         pipe_FVC_ReferenceLevel = pipe.LookupParameter("FVC_Reference Level")
#         TransactionManager.Instance.EnsureInTransaction(doc)
#         pipe_FVC_Length.Set(pipeParam[0])
#         pipe_FVC_Diameter.Set(pipeParam[1])
#         pipe_FVC_PipeType.Set(pipeParam[2])
#         pipe_FVC_PipingSystem.Set(pipeParam[3])
#         pipe_FVC_ReferenceLevel.Set(pipeParam[4])
#         TransactionManager.Instance.TransactionTaskDone()
#     else: pass
# #endregion

#region set pipe fitting  FVC_Angle
categoriesIN = ["OST_PipeFitting"]
categories = [getBuiltInCategoryOST(category) for category in categoriesIN]
flat_categories = [item for sublist in categories for item in sublist]
categoriesFilter = []
for c in flat_categories:
    collector = FilteredElementCollector(doc, view.Id).OfCategory(c).WhereElementIsNotElementType().ToElements()
    categoriesFilter.append(collector)
eleList = [item for sublist in categoriesFilter for item in sublist]
for item in eleList:
    elementAngleCheck = item.LookupParameter("Angle")
    if elementAngleCheck is not None:
        elementAngle = elementAngleCheck.AsValueString()
        element_FVC_Angle = item.LookupParameter("FVC_Angle")
        TransactionManager.Instance.EnsureInTransaction(doc)
        element_FVC_Angle.Set(elementAngle)
        TransactionManager.Instance.TransactionTaskDone()
    else:
        pass

#endregion

OUT =  eleList 