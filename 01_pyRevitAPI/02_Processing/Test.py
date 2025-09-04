
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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#region __function
class selectionFilter(ISelectionFilter):
    def __init__(self,categories):
        self.categories = categories
    def AllowElement(self, element):
        eleName = element.Category.Name in self.categories
        return eleName
def pickFittingOrAccessory(doc):
    categories = ['Pipe Fittings', 'Pipe Accessories']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory')
    desEle = doc.GetElement(ref.ElementId)
    return desEle
def pickMultiFittingOrAccessory(doc):
    desEles = []
    categories = ['Pipe Fittings', 'Pipe Accessories']
    fittingFilter = selectionFilter(categories)
    refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory')
    for ref in refs:
        ele = doc.GetElement(ref.ElementId)
        desEles.append(ele)
    return desEles
def pickPipes(doc):
    desEles = []
    category = "Pipes"
    pipeFilter = selectionFilter(category)
    refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter, "Pick Pipes")
    for ref in refs:
        ele = doc.GetElement(ref.ElementId)
        desEles.append(ele)
    return desEles
#endregion
#region __code here
pipeList = pickPipes(doc)
pipeDiameter = 300
for pipe in pipeList:
    pipeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    TransactionManager.Instance.EnsureInTransaction(doc)
    pipeParam.Set(pipeDiameter/304.8)
    TransactionManager.Instance.TransactionTaskDone()
OUT = pipeParam
#endregion



