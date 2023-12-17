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


def getBuiltInCategoryOTS(nameCategory): # Get Name of BuiltInCategory
    return [i for i in System.Enum.GetValues(BuiltInCategory) if i.ToString() == "OST_"+nameCategory or i.ToString() == nameCategory ]

categoriesIn = IN[0]
categories = []
categoriesFilter = []

for i in categoriesIn:
	cat = getBuiltInCategoryOTS(i)
	categories.append(cat)
flat_categories = [item for sublist in categories for item in sublist]

categoriesFilter.append([FilteredElementCollector(doc).OfCategory(BuitltInCategory.c).WhereElementIsNotElementType() for c in flat_categories])

desEle = categoriesFilter.ToElements()

IDS = List[Autodesk.Revit.DB.ElementId]()
for i in desEle:
	IDS.Add(i.Id)

TransactionManager.Instance.EnsureInTransaction(doc)
isolateEles = view.IsolateElementsTemporary(IDS)
TransactionManager.Instance.TransactionTaskDone()

OUT = desEle

OUT = categories