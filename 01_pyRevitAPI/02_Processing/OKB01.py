import clr
import sys 
import System   
import math
import os

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB import SaveAsOptions
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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#region __function
cateList = [
    "Pipe Fittings"
    ]

class selectionFilter(ISelectionFilter):
    def __init__(self, category):
        self.category = category
    def AllowElement(self, element):
        if element.Category.Name in [c.Name for c in self.category]:
            return True
        else:
            return False
    def AllowReference(reference, point):
        return False

#endregion 

#region __code here
categories = [BuiltInCategory.OST_PipeFitting]
susPipes = []
for c in categories:
    collector = FilteredElementCollector(doc, view.Id).OfCategory(c).WhereElementIsNotElementType().ToElements()
for ele in collector:
    typeId = ele.GetTypeId()
    if typeId != ElementId.InvalidElementId:  
        elementType = doc.GetElement(typeId)
        if elementType is not None:
            elementName = elementType.FamilyName 
            if elementName and elementName == 'SUS_PIPE_DN':  
                susPipes.append(ele)
    




# flat_categoriesFilter = [item for sublist in categoriesFilter for item in sublist]
# IDS = List[ElementId]()
# for i in flat_categoriesFilter:
# 	IDS.Add(i.Id)



#endregion


OUT = susPipes