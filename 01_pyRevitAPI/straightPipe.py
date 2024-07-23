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
"""_________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

class selectionFilter(ISelectionFilter):
	def __init__(self, category):
		self.category = category
	def AllowElement(self, element):
		if element.Category.Name == self.category:
			return True
		else:
			return False
def pickPipe():
	pipes = []
	pipeFilter = selectionFilter("Pipes")
	pipesRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipes")
	for ref in pipesRef:
		pipe = doc.GetElement(ref.ElementId)
		pipes.append(pipe)
	return pipes
def findMidpoint(pipe):
    curve = pipe.Location.Curve
    startPoint = curve.GetEndPoint(0)
    endPoint = curve.GetEndPoint(1)
    midpoint = (startPoint + endPoint) / 2
    return midpoint
def splitPipeAtPoints(doc, pipe, points):
    newPipes = []
    currentPipe = pipe
    TransactionManager.Instance.EnsureInTransaction(doc)
    for point in points:
        currentPipe = pipe
        pipeLocation = currentPipe.Location.Curve
        newPipeIds = PlumbingUtils.BreakCurve(doc, currentPipe.Id, point)
        newPipe = doc.GetElement(newPipeIds)
        newPipes.append(newPipe)
        currentPipe = newPipe
    TransactionManager.Instance.TransactionTaskDone
    return newPipes
def placePipeUnionAtMidpoint(doc, pipe, unionType):
    midpoint = findMidpoint(pipe)
    splitPipeAtPoints(doc, pipe, [midpoint])
    pipe1 = doc.GetElement(pipe.Id)
    pipe2 = doc.GetElement(pipe1.ConnectorManager.Connectors[0].AllRefs[0].Owner.Id)
    TransactionManager.Instance.EnsureInTransaction(doc)
    union = doc.Create.NewFamilyInstance(midpoint, unionType, pipe1, pipe2, StructuralType.NonStructural)
    TransactionManager.Instance.TransactionTaskDone
    return union
def findStraightFamily():
	desUnions = []
	collectors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsElementType().ToElements()
	for fitting in collectors:
		fittingName = fitting.Name.GetValue(fitting)
		if '短管' in fittingName or '直管' in fittingName :
			desUnions.append(fitting)
	return desUnions
"""_________________________________________"""
pipe = pickPipe()
pipeLength = pipe.LookupParameter('Length').AsDouble()
straightPipeFamily = findStraightFamily()[0]
placedStraightFamily = placePipeUnionAtMidpoint(doc, pipe, straightPipeFamily)
straightPipeLength_param = placedStraightFamily.LookupParameter('Length')
straightPipeLength_param.Set(pipeLength)

OUT =  straightPipeLength_param
