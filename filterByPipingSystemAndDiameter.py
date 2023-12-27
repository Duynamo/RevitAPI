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

diameterIN = UnwrapElement(IN[1])
inSystemName = UnwrapElement(IN[2])

categories = [BuiltInCategory.OST_PipeCurves]
categoriesFilter = []
for c in categories:
    collector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType()
    categoriesFilter.append(collector)
flat_categoriesFilter = [item for sublist in categoriesFilter for item in sublist]

IDS = List[ElementId]()
for i in flat_categoriesFilter:
	IDS.Add(i.Id)


desPipes = []

paramFilter1 = ParameterValueProvider(ElementId(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM))
reason1 = FilterStringEquals()
fRule = FilterStringRule(paramFilter1,reason1, inSystemName)
filter1 = ElementParameterFilter(fRule)

i_Check = IN[1]/304.8
paramFilter2 = ParameterValueProvider(ElementId(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
reason2 = FilterNumericEquals()
reason3 = FilterNumericGreaterOrEqual()
fRule2 = FilterDoubleRule(paramFilter2,reason2,i_Check,0.02)
filter2 = ElementParameterFilter(fRule2)
combineFilter = LogicalAndFilter(filter1 , filter2)
desPipes = FilteredElementCollector(doc, IDS).WherePasses(filter2).WhereElementIsNotElementType().ToElements()

OUT = desPipes


