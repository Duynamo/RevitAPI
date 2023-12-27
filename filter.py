"""Copyright by: vudinhduybm@gmail.com"""
import clr 

import System
from System import Array
from System.Collections.Generic import *

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import*
from Autodesk.Revit.DB import Structure

#Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

filterColumns = [FilteredElementCollector(doc , view.Id).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElements()]
filterFramings = [FilteredElementCollector(doc , view.Id).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()]

####filterMultipleSystemFamily
typeList = List[System.Type]()
typeList.Add(Floor)
typeList.Add(Wall)
typeList.Add(Level)

filterList = ElementMulticlassFilter(typeList)
elements = FilteredElementCollector(doc , view.Id).WherePasses(filterList).WhereElementIsNotElementType().ToElements()

####filterMultiFamilySymbol
categories = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_StructuralColumns]

categoriesFilter = []
for c in categories:
    collector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType()
    categoriesFilter.append(collector)
flat_categoriesFilter = [[item for item in sublist] for sublist in categoriesFilter]

####filterMultiFamilySymbols
cateList = List[BuiltInCategory]()

cateList.Add(BuiltInCategory.OST_StructuralColumns)
cateList.Add(BuiltInCategory.OST_StructuralFraming)

_filter = ElementMulticategoryFilter(cateList)
elems = FilteredElementCollector(doc).WherePasses(_filter).WhereElementIsNotElementType().ToElements()

OUT = elems


####filterCurveElements
filter3 = CurveElementFilter(CurveElementType.ModelCurve)
elems3 = FilteredElementCollector(doc).WherePasses(filter3).WhereElementIsNotElementType().ToElements()

####filterElementByParameters
paramFilter1 = ParameterValueProvider(ElementId(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
reason1 = FilterNumericGreater()
reason2 = FilterNumericEquals()

i_Check = IN[1]/304.8
fRule = FilterDoubleRule(paramFilter1,reason1,i_Check,0.02)
filter4 = ElementParameterFilter(fRule)
elems4 = FilteredElementCollector(doc).WherePasses(filter4).WhereElementIsNotElementType().ToElements()

IDS = List[ElementId]()
for i in elems4:
	IDS.Add(i.Id)

collector = FilteredElementCollector(doc, IDS).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()

OUT = collector