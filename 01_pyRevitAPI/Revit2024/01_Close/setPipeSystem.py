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
"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""
class selectionFilter(ISelectionFilter):
    def __init__(self,categories):
        self.categories = categories
    def AllowElement(self, element):
        eleName = element.Category.Name in self.categories
        return eleName
def pickPipeOrPipePart():
    categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
    desEle = doc.GetElement(ref.ElementId)
    return desEle
def pickPipeOrPipeParts():
	partList = []
	categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
	fittingFilter = selectionFilter(categories)
	refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
	for ref in refs:
		desEle = doc.GetElement(ref.ElementId)
		partList.append(desEle)
	return partList
"""___"""
#pick pipe part. 
basePart = pickPipeOrPipePart()
listSetPart = pickPipeOrPipeParts() 
#get base part piping system parameter
basePart_PipingSystem_check = basePart.LookupParameter('System Type')
if basePart_PipingSystem_check != None:
	basePart_PipingSystem = basePart_PipingSystem_check.AsElementId()
	for part in listSetPart:
		updatedPart = []
		part_PipingSystem = part.LookupParameter('System Type')
		TransactionManager.Instance.EnsureInTransaction(doc)
		part_PipingSystem.Set(basePart_PipingSystem)
		TransactionManager.Instance.TransactionTaskDone()
		updatedPart.append(part)
else : pass
OUT = listSetPart, updatedPart
