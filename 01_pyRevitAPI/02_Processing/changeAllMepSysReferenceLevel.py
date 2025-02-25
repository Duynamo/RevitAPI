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
"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""
def get_level_by_name(level_name):
    # Thu thập tất cả các Level trong dự án
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    
    # Tìm Level có tên khớp với level_name
    for level in levels:
        if level.Name == level_name:
            return level
# def flatten(nestedList):
#     flatList = []
#     for item in nestedList:
#         if isinstance(item, list):
#             flatList.extend(flatten(item)) #Gọi đệ quy (flatten(item) để làm phẳng danh sách con) và thêm tất cả phần tử của ds con 
#                                         #đã được làm phẳng vào flatList.
#         else:
#             flatList.append(item)
#     return flatList


def flatten(nestedList):
    if not isinstance(nestedList, (list, tuple, List[Element], List[XYZ])):  # Hỗ trợ Revit collections
        raise ValueError("Input must be a list, tuple, or Revit collection, got {}".format(type(nestedList)))
    
    flatList = []
    for item in nestedList:
        if isinstance(item, (list, tuple, List[Element], List[XYZ])):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList

def getBuiltInCategoryOST(nameCategory):
    # Get BuiltInCategory by name
    return [i for i in System.Enum.GetValues(BuiltInCategory) if i.ToString() == "OST_" + str(nameCategory) or i.ToString() == nameCategory]


categoriesIN = ["OST_PipeCurves", "OST_PipeFitting", "OST_PipeAccessory"]
categories = [getBuiltInCategoryOST(category) for category in categoriesIN]
flat_categories = [item for sublist in categories for item in sublist]

categoriesFilter = []
for c in flat_categories:
    collector = FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType()
    categoriesFilter.append(collector)
lst1 = [[item for item in sublist] for sublist in categoriesFilter]
# lst2 = [[item for item in sublist1] for sublist1 in lst1] 
desReferenceLevelId = get_level_by_name("1FL(下部) + 5.3m")
# for item in flatten(lst1):
#     try:
#         element_referenceLevelCheck = item.LookupParameter("")
#         if element_referenceLevelCheck is not None:
#             element_referenceLevel = element_referenceLevelCheck.AsElementId()
#             TransactionManager.Instance.EnsureInTransaction(doc)
#             element_referenceLevel.Set(desReferenceLevelId)
#             TransactionManager.Instance.TransactionTaskDone()
#     except: pass        
levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
lss = []
for lev in levels:
    lss.append(lev.Name)
OUT =  desReferenceLevelId, levels,lss