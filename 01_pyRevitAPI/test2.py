"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import cos,sin,tan,radians

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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList

def divideLineSegment(doc, pipe, pointA, LengthA):
    TransactionManager.Instance.EnsureInTransaction(doc)
    lineSegment = pipe.Location.Curve
    start_point = lineSegment.GetEndPoint(0)
    end_point = lineSegment.GetEndPoint(1)
    vector = end_point - start_point
    total_length = vector.GetLength()*304.8
    direction = vector.Normalize()
    num_segments = int(total_length / LengthA)
    points = []
    desPoints =[]
    current_point = pointA
    for i in range(num_segments + 1):
        points.append(current_point)
        desPoints.append(current_point)
    TransactionManager.Instance.TransactionTaskDone()
    return desPoints
def splitPipeAtPoints(doc, pipe, points):
    newPipes = []
    points = sorted(points, key=lambda p: (pipe.Location.Curve.GetEndPoint(0) - p).DotProduct(pipe.Location.Curve.Direction))

    for point in points:
        pipeLocation = pipe.Location
        if isinstance(pipeLocation, LocationCurve):
            pipeCurve = pipeLocation.Curve
            if pipeCurve is not None:
                newPipeIds = PlumbingUtils.BreakCurve(doc, pipe.Id, point)
    return newPipes
def divideLineSegment(line, length, startPoint,endPoint):
    # Initialize the list of points
    points = []
    points.append(startPoint)
    current_point = startPoint
    total_length = line.Length
    direction = (endPoint - startPoint).Normalize()
    # Loop to add points at intervals of 'length'
    while (startPoint.DistanceTo(start_point) + length) <= total_length:
        # Calculate the next point
        current_point = current_point + direction * length

        # Add the next point to the list of points
        points.append(current_point)

    return points
#endregion
pipe   = UnwrapElement(IN[0])
splLength = 1000
children = {}
parents, childs, fittings = [], [], []
# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
pipeCurve = pipe.Location.Curve
conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
orgiginConns = list(c.Origin for c in conns)
sortConns = sorted(orgiginConns, key=lambda c : c.Y)
#splitPoints = divideLineSegment(doc, pipe, sortConns[0], splLength)
#list = list(c.ToPoint() for c in splitPoints)
points = divideLineSegment(pipeCurve, splLength, sortConns[0], sortConns[1])

TransactionManager.Instance.TransactionTaskDone()

OUT = points



def connectPartToPipe(doc, pipePart, pipe):
    TransactionManager.Instance.EnsureInTransaction(doc)
    pipePartConnectors = [conn for conn in pipePart.MEPModel.ConnectorManager.Connectors]
    pipeConnectors = [conn for conn in pipe.ConnectorManager.Connectors]
    nearestPipePartConn = findNearestConnector(pipePartConnectors, pipeConnectors[0].Origin)
    nearestPipeConn = findNearestConnector(pipeConnectors, nearestPipePartConn.Origin)
    nearestPipePartConn.ConnectTo(nearestPipeConn)
    TransactionManager.Instance.TransactionTaskDone()
    return nearestPipePartConn, nearestPipeConn

def connectPartToPart(doc, pipePart1, pipePart2):
    TransactionManager.Instance.EnsureInTransaction(doc)
    pipePart1Connectors = [conn for conn in pipePart1.MEPModel.ConnectorManager.Connectors]
    pipePart2Connectors = [conn for conn in pipePart2.MEPModel.ConnectorManager.Connectors]
    nearestPipePart1Conn = findNearestConnector(pipePart1Connectors, pipePart2Connectors[0].Origin)
    nearestPipePart2Conn = findNearestConnector(pipePart2Connectors, pipePart1Connectors[0].Origin)
    nearestPipePartConn.ConnectTo(nearestPipeConn)
    TransactionManager.Instance.TransactionTaskDone()
    return nearestPipePartConn, nearestPipeConn

def connect2Connectors(doc, pipePart1, pipePart2):
    TransactionManager.Instance.EnsureInTransaction(doc)
    if pipePart1.Category.Name != 'Pipes':
        pipePart1Connectors = [conn for conn in pipePart1.MEPModel.ConnectorManager.Connectors]   
    else:
        pipePart1Connectors = [conn for conn in pipePart1.ConnectorManager.Connectors]     
    if pipePart2.Category.Name != 'Pipes':
        pipePart2Connectors = [conn for conn in pipePart2.MEPModel.ConnectorManager.Connectors]   
    else:
        pipePart2Connectors = [conn for conn in pipePart2.ConnectorManager.Connectors] 
    nearestPipePart1Conn = findNearestConnector(pipePart1Connectors, pipePart2Connectors[0].Origin)
    nearestPipePart2Conn = findNearestConnector(pipePart2Connectors, pipePart1Connectors[0].Origin)    
    nearestPipePart1Conn.ConnectTo(nearestPipePart2Conn)
    TransactionManager.Instance.TransactionTaskDone()
    return nearestPipePart1Conn, nearestPipePart2Conn






def offsetPointAlongVector(point, vector, offsetDistance):
    direction = vector.ToRevitType().Normalize()
    scaledVector = direction.Multiply(offsetDistance).ToVector()
    offsetPoint = point.Add(scaledVector)
    return offsetPoint

def scaleCurve(pipe, offsetDistance):
    curve = pipe.Location.Curve
    point1 = curve.GetEndPoint(0)
    point2 = curve.GetEndPoint(1)
    v1 = (point1 - point2).Normalize()
    v2 = (point2 - point1).Normalize()
    scaled_v1 = point1.Multiply(offsetDistance)
    scaled_v2 = point2.Multiply(offsetDistance)
    offsetPoint1 = point1.Add(scaled_v1)
    offsetPoint2 = point2.Add(scaled_v2)
    
    return 


"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import *

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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
class selectionFilter(ISelectionFilter):
	def __init__(self, category):
		self.category = category
	def AllowElement(self, element):
		if element.Category.Name == self.category:
			return True
		else:
			return False
def pickPipe():
	pipeFilter = selectionFilter("Pipes")
	ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipe")
	pipe = doc.GetElement(ref.ElementId)
	return pipe
	
def getPipeParameter(p):
    paramDiameters = []
    paramPipingSystems = []
    paramLevels = []
    paramPipeTypes = []

    paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
    
    paramPipeTypeId = p.GetTypeId()
    paramPipeType = doc.GetElement(paramPipeTypeId)
    paramPipeTypes.append(paramPipeType)
    pipeTypeName = paramPipeType.LookupParameter("Type Name").AsString()
	
    paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    paramPipingSystem = doc.GetElement(paramPipingSystemId)
    paramPipingSystems.append(paramPipingSystem)
    pipingSystemName = paramPipingSystem.LookupParameter("System Classification").AsValueString()

    paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    paramLevel = doc.GetElement(paramLevelId)
    paramLevels.append(paramLevel)

    return paramDiameter, paramPipingSystem, paramPipeType, paramLevel
def scaleCurve(pipe, offsetDistance):
    k = offsetDistance/304.8
    curve = pipe.Location.Curve
    point1 = curve.GetEndPoint(0)
    point2 = curve.GetEndPoint(1)
    
    # Compute the direction vectors and normalize them
    v1 = (point1 - point2).Normalize()
    v2 = (point2 - point1).Normalize()
    
    # Scale the vectors by the offset distance
    scaled_v1 = v1.Multiply(k)
    scaled_v2 = v2.Multiply(k)
    
    # Compute the new endpoints
    offsetPoint1 = point1.Add(scaled_v1)
    offsetPoint2 = point2.Add(scaled_v2)
    
    # Create a new line with the scaled endpoints
    newCurve = Line.CreateBound(offsetPoint1, offsetPoint2)
    
    # Update the pipe's location curve
    pipe.Location.Curve = newCurve

    return newCurve    
#endregion


eleList   = uwList(IN[0])

# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
pipe = pickPipe()
pipeParam = getPipeParameter(pipe)
newCurve = scaleCurve(pipe, 1000)

TransactionManager.Instance.TransactionTaskDone()

OUT = newCurve


def connectElbowToPipe(doc, ele1, ele2):
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    # Get the connectors of the elbow and pipe
    if ele1.Name == "Pipes":
        ele1Conns = [conn for conn in ele1.ConnectorManager.Connectors]
    elif ele1.Name != "Pipes":
        ele1Conns = [conn for conn in ele1.MEPModel.ConnectorManager.Connectors]    
    if ele2.Name == "Pipes":
        ele2Conns = [conn for conn in ele2.ConnectorManager.Connectors]
    elif ele2.Name != "Pipes":
        ele2Conns = [conn for conn in ele2.MEPModel.ConnectorManager.Connectors] 
    
    # Find the nearest connectors between the elbow and pipe
    nearestEle1Conn = findNearestConnector(ele1Conns, ele2Conns[0].Origin)
    nearestEle2Conn = findNearestConnector(ele2Conns, nearestEle1Conn.Origin)
    
    # Connect the nearest connectors
    nearestEle1Conn.ConnectTo(nearestEle2Conn)
    
    TransactionManager.Instance.TransactionTaskDone()
    return nearestEle1Conn, nearestEle2Conn



def findConnectors(element):
    connectors = []
    if element.Category.Name == "Pipes":
        connectors = [conn for conn in element.ConnectorManager.Connectors]
    else:
        connectors = [conn for conn in element.MEPModel.ConnectorManager.Connectors]
    return connectors
def disconnectConnectors(element):
    desConn = []
    connectors = findConnectors(element)
    for connector in connectors:
        for connected_connector in connector.AllRefs:
            if connected_connector.IsConnected:
                desConn.append(connected_connector)
    return desConn