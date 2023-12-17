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

for c in flat_categories:
	collector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType()
	categoriesFilter.append(collector)
flat_categoriesFilter = [item for sublist in categoriesFilter for item in sublist]
# for sublist in categoriesFilter:
# 	for item in sublist:
# 		filteredFams = []
# 		fam = doc.GetElement(item.Symbol.Family.Id)
# 		filteredFams.append(fam)

filteredFams = []
filteredFams.append([[doc.GetElement(item.Symbol.Family.Id)] for sublist in categoriesFilter for item in sublist])

IDS = List[ElementId]()
for i in flat_categoriesFilter:
	IDS.Add(i.Id)

TransactionManager.Instance.EnsureInTransaction(doc)
isolatedEles = view.IsolateElementsTemporary(IDS)
TransactionManager.Instance.TransactionTaskDone()

OUT = filteredFams
