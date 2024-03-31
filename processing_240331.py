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
desFamTypes = []
key = "FU_Support"
categoriesFilter = []
for category in categories:
    elementTypes = FilteredElementCollector(doc).OfCategory(category).WhereElementIsElementType().ToElements()
    for elementType in elementTypes:
        typeName = elementType.FamilyName
        if key in typeName:
            desFamTypes.append(elementType)
categoriesFilter.append(desFamTypes)
# flat_categoriesFilter = [[item for item in sublist] for sublist in categoriesFilter]

flat_categoriesFilter = [ item for sublist in categoriesFilter for item in sublist]

OUT = flat_categoriesFilter