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
	pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipe")
	pipe = doc.GetElement(pipeRef.ElementId)
	return pipe
def findMidpoint(pipe):
    curve = pipe.Location.Curve
    startPoint = curve.GetEndPoint(0)
    endPoint = curve.GetEndPoint(1)
    midpoint = (startPoint + endPoint) / 2
    return midpoint
def splitPipeAtPoints(doc, pipe, points):
    newPipes = []
    ids = []
    currentPipe = pipe
    TransactionManager.Instance.EnsureInTransaction(doc)
    for point in points:
        currentPipe = pipe
        pipeLocation = currentPipe.Location.Curve
        newPipeIds = PlumbingUtils.BreakCurve(doc, currentPipe.Id, point)
        newPipe = doc.GetElement(newPipeIds)
        newPipes.append(newPipe)
        newPipes.append(pipe)
    TransactionManager.Instance.TransactionTaskDone
    return newPipes
def placePipeUnionAtMidpoint(doc, pipe, unionType):
    levelId = pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    level = doc.GetElement(levelId)

    midpoint = findMidpoint(pipe)
    splittedPipes = splitPipeAtPoints(doc, pipe, [midpoint])
    pipe1 = splittedPipes[0]
    pipe2 = splittedPipes[1]
    TransactionManager.Instance.EnsureInTransaction(doc)
    union = doc.Create.NewFamilyInstance(midpoint, unionType, pipe1,level, StructuralType.NonStructural)
    TransactionManager.Instance.TransactionTaskDone
    return union
def findStraightFamily(doc):
    desUnions = []
    lst = []
    collectors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsElementType().ToElements()
    for fitting in collectors:
        fittingName = fitting.LookupParameter('Family Name').AsString()
        if '短管' in fittingName or '直管' in fittingName:
            desUnions.append(fitting)
    return None
def NearestConnector(ConnectorSet, curCurve):
    MinLength = float("inf")
    result = None  # Initialize result to None
    for n in ConnectorSet:
        distance = curCurve.Location.Curve.Distance(n.Origin)
        if distance < MinLength:
            MinLength = distance
            result = n
    return result

def closetConn(mPipe, bPipe):
    connectors1 = list(bPipe.ConnectorManager.Connectors.GetEnumerator())
    Connector1 = NearestConnector(connectors1, mPipe)
    XYZconn	= Connector1.Origin
    return Connector1
def getPipeParameter(p):
    paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()*304.8
    paramDiameter1 = p.LookupParameter('Diameter').AsDouble()*304.8
    paramPipeTypeId = p.GetTypeId()
    paramPipeType = doc.GetElement(paramPipeTypeId)
    pipeTypeName = paramPipeType.LookupParameter("Type Name").AsString()
    paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    # paramPipingSystem = doc.GetElement(paramPipingSystemId)
    pipingSystemName = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
    paramLevel = p.LookupParameter('Reference Level').AsValueString()
    return str(paramDiameter1),pipingSystemName,pipeTypeName,paramLevel
"""_________________________________________"""
pipe = pickPipe()
pipeLength = pipe.LookupParameter('Length').AsDouble()
midPoint = findMidpoint(pipe)
splitPipe = splitPipeAtPoints(doc, pipe, [midPoint])
conn1 = closetConn(splitPipe[0], splitPipe[1])
conn2 = closetConn(splitPipe[1], splitPipe[0])
insertUnion  = doc.Create.NewUnionFitting(conn1,conn2)
#region __lookup and set union pipe parameters 
TransactionManager.Instance.EnsureInTransaction(doc)
pipe_diameter_params = getPipeParameter(pipe)
union_Diameter_param = insertUnion.LookupParameter('_Diameter')
union_Diameter_param.Set(pipe_diameter_params[0])
union_PipeType_param = insertUnion.LookupParameter('_Pipe Type')
union_PipeType_param.Set(pipe_diameter_params[2])
union_PipingSystem_param = insertUnion.LookupParameter('_Piping System')
union_PipingSystem_param.Set(pipe_diameter_params[1])
union_ReferenceLevel_param = insertUnion.LookupParameter('_Reference Level')
union_ReferenceLevel_param.Set(pipe_diameter_params[3])
TransactionManager.Instance.TransactionTaskDone()
#endregion
TransactionManager.Instance.EnsureInTransaction(doc)
straightPipeConns = []
sortStraightPipeConns = []
if insertUnion is not None:
    straightPipeLength_param = insertUnion.LookupParameter('L')
    if straightPipeLength_param is not None:
        straightPipeLength_param.Set(pipeLength)
        straightPipeConns = [conn for conn in insertUnion.MEPModel.ConnectorManager.Connectors]
        connsOrigin = [c.Origin for c in straightPipeConns]
        sortStraightPipeConns = [midPoint]
        added_points = {midPoint}
        for c in connsOrigin:
            if not any(c.IsAlmostEqualTo(point) for point in added_points):
                sortStraightPipeConns.append(c)
                added_points.add(c)
        vector = sortStraightPipeConns[0] - sortStraightPipeConns[1]
        transVector = vector.Normalize().Multiply(pipeLength/2)
        translation = Transform.CreateTranslation(transVector)
        insertUnion.Location.Move(translation.Origin)
TransactionManager.Instance.TransactionTaskDone()

OUT =  insertUnion, sortStraightPipeConns, pipe_diameter_params
