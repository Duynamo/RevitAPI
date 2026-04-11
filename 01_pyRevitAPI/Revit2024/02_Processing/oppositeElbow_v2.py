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
def findFarConnector(connectorSet, targetPoint):
    farConnector = None
    minDistance = float('inf')
    for connector in connectorSet:
        distance = (connector.Origin - targetPoint).GetLength()
        if distance > minDistance:
            maxDistance = distance
            farConnector = connector
    return farConnector
def connect2Connectors(doc, elbow, pipe): #Still bug here
    TransactionManager.Instance.EnsureInTransaction(doc)
    unconnected_elbow = unconnectedConn(doc, elbow)
    pipeConns = [conn for conn in pipe.ConnectorManager.Connectors]
    nearConn_pipe = findNearestConnector(pipeConns, unconnected_elbow.Origin)
    farConn_pipe = findFarConnector(pipeConns, unconnected_elbow.Origin)
    newLocationCurve = Line.CreateBound(unconnected_elbow.Origin, farConn_pipe)   
    locationCurve = pipe.Location
    locationCurve.Curve = newLocationCurve
    nearConn_pipe.ConnectTo(unconnected_elbow)
    TransactionManager.Instance.TransactionTaskDone()
    return nearConn_pipe
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
def mirrorLocationCurve(doc, pipe1, pipe2, createCopy = True):
    line1 = pipe1.Location.Curve
    line2 = pipe2.Location.Curve
    line2_SP = line2.GetEndPoint(0)
    line2_EP = line2.GetEndPoint(1)
    start1 = line1.GetEndPoint(0)
    end1 = line1.GetEndPoint(1)
    midpoint1 = (start1 + end1) / 2
    direction1 = (end1 - start1).Normalize()
    normal = direction1.CrossProduct(XYZ.BasisZ).Normalize()
    mirrorPlane = Plane.CreateByNormalAndOrigin(normal, midpoint1)   
    vector1_to_line1 = line2_SP - line1.GetEndPoint(0)
    perpendicular_vector1 = vector1_to_line1 - vector1_to_line1.DotProduct(direction1)*direction1
    mirror_point1 = line2_SP-2*perpendicular_vector1
    vector2_to_line1 = line2_EP - line1.GetEndPoint(0)
    perpendicular_vector2 = vector2_to_line1 - vector2_to_line1.DotProduct(direction1)*direction1
    mirror_point2 = line2_EP-2*perpendicular_vector2  
    mirror_line2 = Line.CreateBound(mirror_point1, mirror_point2)
    return mirror_line2
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

tempCurve = mirrorLocationCurve(doc, pipe1, pipe2)
tempPoint1 = tempCurve.GetEndPoint(0)
tempPoint2 = tempCurve.GetEndPoint(1)
newPipe1 = Pipe.Create(doc,pipingSystemId, pipeTypeId, levelId.Id, tempPoint1, tempPoint2)
tempPipeId = newPipe1.Id
newPipe1_diam = newPipe1.LookupParameter("Diameter")
newPipe1_diam.Set(pipe2_diam)
pipeList.append(newPipe1)
elbow1 = createElbow(doc, pipeList)
if elbow1:
    doc.Delete(newPipe1.Id)
    elbow1_connectedConn = connectedConn(doc, elbow1)
    rotateElbow1 = rotateElbowAroundPipe(doc, elbow1, pipe1, elbow1_connectedConn)
    unconnected_elbow = unconnectedConn(doc, elbow1)
    pipe2Conns = [conn for conn in pipe2.ConnectorManager.Connectors]
    nearConn_pipe2 = findNearestConnector(pipe2Conns, unconnected_elbow.Origin)
    farConn_pipe = findFarConnector(pipe2Conns, unconnected_elbow.Origin)
    newLocationCurve = Line.CreateBound(unconnected_elbow.Origin, farConn_pipe)   
    locationCurve = pipe2.Location
    locationCurve.Curve = newLocationCurve
    nearConn_pipe2.ConnectTo(unconnected_elbow)    
    # connect_elbow1_pipe2 = connect2Connectors(doc, elbow1, pipe2)
TransactionManager.Instance.TransactionTaskDone()
OUT =   elbow1