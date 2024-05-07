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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView


"""____"""

categories = [BuiltInCategory.OST_PipeAccessory]

categoriesFilter = []
for c in categories:
    catCollector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsElementType()
    #catName = catCollector.FamilyName
    categoriesFilter.append(catCollector)

flat_categoriesFilter = [[item for item in sublist] for sublist in categoriesFilter]

for cat in flat_categoriesFilter:
    catName = cat.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()

####filterMultiFamilySymbols
# cateList = List[BuiltInCategory]()

# cateList.Add(BuiltInCategory.OST_StructuralColumns)
# cateList.Add(BuiltInCategory.OST_StructuralFraming)

# _filter = ElementMulticategoryFilter(cateList)
# elems = FilteredElementCollector(doc).WherePasses(_filter).WhereElementIsNotElementType().ToElements()

OUT = flat_categoriesFilter