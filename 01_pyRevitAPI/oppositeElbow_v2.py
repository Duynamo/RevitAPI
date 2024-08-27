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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

#region _function
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return result
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
	pipesRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipes")
	pipe = doc.GetElement(pipesRef.ElementId)
	return pipe	
def getConnectedElement(doc, ele):
    _connectedEle = None
    desEle = []
    conns = ele.MEPModel.ConnectorManager.Connectors
    for conn in conns:
        if conn.IsConnected:
            for refConn in conn.AllRefs:
                connectedEle = refConn.Owner
                if connectedEle.Id != ele.Id:
                    _connectedEle = connectedEle
    return _connectedEle
def unconnectedConn(doc, fittingOrAccessory):
    conns = list(fittingOrAccessory.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    _connectedConn = None
    _unConnectedConn = None
    for conn in conns:
        if conn.IsConnected:
            _connectedConn = conn
        else:
            _unConnectedConn = conn
    return _unConnectedConn
def connectedConn(doc, fittingOrAccessory):
    conns = list(fittingOrAccessory.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    _connectedConn = None
    _unConnectedConn = None
    for conn in conns:
        if conn.IsConnected:
            _connectedConn = conn
        else:
            _unConnectedConn = conn
    return _connectedConn
def findNearestConnector(connectorSet, targetPoint):
    nearest_connector = None
    minDistance = float('inf')
    
    for connector in connectorSet:
        distance = (connector.Origin - targetPoint).GetLength()
        if distance < minDistance:
            minDistance = distance
            nearestConnector = connector
    return nearestConnector
def connect2Connectors(doc, pipePart1, pipePart2):
    TransactionManager.Instance.EnsureInTransaction(doc)
    nearestPipePart1Conn = None
    nearestPipePart2Conn = None    
    if pipePart1.Category.Name != 'Pipes':
        pipePart1Connectors = [conn for conn in pipePart1.MEPModel.ConnectorManager.Connectors]   
    else:
        pipePart1Connectors = [conn for conn in pipePart1.ConnectorManager.Connectors]     
    if pipePart2.Category.Name != 'Pipes':
        pipePart2Connectors = [conn for conn in pipePart2.MEPModel.ConnectorManager.Connectors]   
    else:
        pipePart1Connectors = [conn for conn in pipePart2.ConnectorManager.Connectors] 
    nearestPipePart1Conn = findNearestConnector(pipePart1Connectors, pipePart2Connectors[0].Origin)
    nearestPipePart2Conn = findNearestConnector(pipePart2Connectors, pipePart1Connectors[0].Origin)    
    nearestPipePart1Conn.ConnectTo(nearestPipePart2Conn)
    TransactionManager.Instance.TransactionTaskDone()
    return nearestPipePart1Conn, nearestPipePart2Conn
def createElbow(doc, pipes):
	connectors = []

	try:
		for pipe in pipes:
			connectors.append(pipe.ConnectorManager.Connectors)
	except:
		pass
	TransactionManager.Instance.EnsureInTransaction(doc)
	try:
		for set1, set2 in zip(connectors[:-1], connectors[1:]):
			conStart = list(set1)
			conEnd = list(set2)
			d1 = conStart[0].Origin.DistanceTo(conEnd[0].Origin)*conStart[0].Origin.DistanceTo(conEnd[1].Origin)
			d2 = conStart[1].Origin.DistanceTo(conEnd[0].Origin)*conStart[1].Origin.DistanceTo(conEnd[1].Origin)
			i = 0
			if d1 > d2:
				i = 1
			d3 = conEnd[0].Origin.DistanceTo(conStart[0].Origin)*conEnd[0].Origin.DistanceTo(conStart[1].Origin)
			d4 = conEnd[1].Origin.DistanceTo(conStart[0].Origin)*conEnd[1].Origin.DistanceTo(conStart[1].Origin)
			j = 0
			if d3 > d4:
				j = 1
			fitting = doc.Create.NewElbowFitting(conStart[i],conEnd[j])
	except:
		pass
	return fitting
def NearestConnector(ConnectorSet, curCurve):
    MinLength = float("inf")
    result = None  # Initialize result to None
    for n in ConnectorSet:
        distance = curCurve.Location.Curve.Distance(n.Origin)
        if distance < MinLength:
            MinLength = distance
            result = n
    return result
def closetConn(pipe1, pipe2):
    connLst = []
    connectors = list(pipe2.ConnectorManager.Connectors.GetEnumerator())
    nearConn1 = [NearestConnector(connectors, pipe1)]
    farConn1  = [conn for conn in connectors if conn not in nearConn1]
    shortedConnLst = [ele for ele in nearConn1] + [ele for ele in farConn1]
    connLst = [conn.Origin for conn in shortedConnLst]
    vector = connLst[0]-connLst[1]
    return connLst, vector
def offsetPointAlongVector(point, vector, offsetDistance):
    direction = vector.Normalize()
    scaledVector = direction.Multiply(offsetDistance)
    offsetPoint = point.Add(scaledVector)
    return offsetPoint
def getPipeParameter(p):
    pipeTypeId = p.PipeType.Id
    pipingSystemId = p.MEPSystem.GetTypeId()
    levelId = doc.GetElement(p.ReferenceLevel.Id)
    diam = pipe2.LookupParameter('Diameter').AsDouble()
    return diam, pipingSystemId, pipeTypeId, levelId
def rotateElbowAroundPipe(doc, elbow, pipe, connector):
    TransactionManager.Instance.EnsureInTransaction(doc)
    connOrigin = connector.Origin
    connTangent = connector.CoordinateSystem.BasisZ   
    rotationAxis = connTangent.Normalize()
    rotationAngle = math.radians(180)
    rotationTransform = Transform.CreateRotationAtPoint(rotationAxis, rotationAngle, connOrigin)
    ElementTransformUtils.RotateElement(doc, elbow.Id, Line.CreateBound(connOrigin, connOrigin + rotationAxis), rotationAngle)
    TransactionManager.Instance.TransactionTaskDone()
    return elbow
def mirrorLine(doc, line1, line2):
    start1 = line1.GetEndPoint(0)
    end1 = line1.GetEndPoint(1)
    midpoint1 = (start1 + end1) / 2
     # Calculate the direction of line1 (this is the axis of symmetry)
    direction1 = (end1 - start1).Normalize()
    # Calculate the normal vector to the mirror plane
    normal = direction1.CrossProduct(XYZ.BasisZ).Normalize()
    # Create the mirror plane (defined by midpoint1 and the normal vector)
    mirrorPlane = Plane.CreateByNormalAndOrigin(normal, midpoint1)   
    # Create the mirror transform
    mirrorTransform = Transform.CreateReflection(mirrorPlane)
    # Apply the mirror transform to line2
    ElementTransformUtils.MirrorElement(doc, line2.Id, mirrorPlane)
    return line2

def mirrorPipe(doc, pipe1, pipe2):
    line1 = pipe1.Location.Curve
    line2 = pipe2.Location.Curve
    line2_SP = line2.GetEndPoint(0)
    line2_EP = line2.GetEndPoint(1)
    tempLine = Line.CreateBound(line2_SP, line2_EP)
    start1 = line1.GetEndPoint(0)
    end1 = line1.GetEndPoint(1)
    midpoint1 = (start1 + end1) / 2
     # Calculate the direction of line1 (this is the axis of symmetry)
    direction1 = (end1 - start1).Normalize()
    # Calculate the normal vector to the mirror plane
    normal = direction1.CrossProduct(XYZ.BasisZ).Normalize()
    # Create the mirror plane (defined by midpoint1 and the normal vector)
    mirrorPlane = Plane.CreateByNormalAndOrigin(normal, midpoint1)   
    # Create the mirror transform
    mirrorTransform = Transform.CreateReflection(mirrorPlane)
    # Apply the mirror transform to line2
    newPipe = ElementTransformUtils.MirrorElement(doc, pipe2.Id, mirrorPlane)
    return newPipe
#endregion
'''___'''
pipe1 = pickPipe()
pipe2 = pickPipe()

pipeTypeId = pipe2.PipeType.Id
pipingSystemId = pipe2.MEPSystem.GetTypeId()
levelId = doc.GetElement(pipe2.ReferenceLevel.Id)
pipe2_diam = pipe2.LookupParameter('Diameter').AsDouble()
TransactionManager.Instance.EnsureInTransaction(doc)
pipeList = []
pipeList.append(pipe1)
tempPipe = mirrorPipe(doc, pipe1, pipe2)
# tempPoint1 = tempLocationCurve.GetEndPoint(0)
# tempPoint2 = tempLocationCurve.GetEndPoint(1)
# newPipe1 = Pipe.Create(doc,pipingSystemId, pipeTypeId, levelId.Id, tempPoint1, tempPoint2)
# tempPipeId = tempPipe.Id
# newPipe1_diam = newPipe1.LookupParameter("Diameter")
# newPipe1_diam.Set(pipe2_diam)
pipeList.append(tempPipe)
# elbow1 = createElbow(doc, pipeList)

# if elbow1:
    # doc.Delete(tempPipe.Id)
    # elbow1_connectedConn = connectedConn(doc, elbow1)
    # rotateElbow1 = rotateElbowAroundPipe(doc, elbow1, pipe1, elbow1_connectedConn)

TransactionManager.Instance.TransactionTaskDone()


OUT =   pipeList