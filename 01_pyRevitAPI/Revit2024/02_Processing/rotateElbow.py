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
    def __init__(self,categories):
        self.categories = categories
    def AllowElement(self, element):
        eleName = element.Category.Name in self.categories
        return eleName
def findNearestConnector(connectorSet, targetPoint):
    nearest_connector = None
    minDistance = float('inf')
    for connector in connectorSet:
        distance = (connector.Origin - targetPoint).GetLength()
        if distance < minDistance:
            minDistance = distance
            nearestConnector = connector
    return nearestConnector.Origin    
def pickFittingOrAccessory():
    pickedEles = []
    categories = ['Pipe Fittings', 'Pipe Accessories']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory')
    ele = doc.GetElement(ref.ElementId)
    return ele
def getElbowOrigin(elbow):
    conns = list(elbow.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    sumX = sum(conn.Origin.X for conn in conns)
    sumY = sum(conn.Origin.Y for conn in conns)
    sumZ = sum(conn.Origin.Z for conn in conns)
    numConnectors = len(conns)
    originX = sumX / numConnectors
    originY = sumY / numConnectors
    originZ = sumZ / numConnectors
    elbowOrigin = XYZ(originX, originY, originZ)
    return elbowOrigin
def findNearestConnectorOf2Fittings(fitting1, fitting2):
    nearest_connector = []
    minDistance = float('inf')
    conns1 = list(fitting1.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    conns2 = list(fitting2.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    for c1,c2 in zip(conns1, conns2):
        distance = (c1.Origin - c2.Origin).GetLength()
        if distance < minDistance:
            minDistance = distance
            nearestConnector = [c1,c2]
    return nearestConnector
def rotateElbow(doc, elbow, point, angle):
    elbowLocation = elbow.Location
    elbowTransform = elbow.GetTransform()
    rotateAxis = XYZ.BasisZ
    rotationTransform = Transform.CreateRotationAtPoint(rotateAxis, angle, point)
    rotationLine = Line.CreateBound(point, point + rotateAxis)
    TransactionManager.Instance.EnsureInTransaction(doc)
    rotatedElbow = ElementTransformUtils.RotateElement(doc, elbow.Id, rotationLine, angle)
    TransactionManager.Instance.TransactionTaskDone()
    return rotatedElbow
def rotateAndMoveElbow(elbow, baseElbow):
    rotateAngle = math.radians(180) - elbow.LookupParameter('Angle').AsDouble()
    elbowOrigin = getElbowOrigin(elbow)
    _rotatedElbow = rotateElbow(doc, elbow, elbowOrigin, rotateAngle)
    elbow_Conn = None
    baseElbow_conn = None
    nearConns = findNearestConnectorOf2Fittings(elbow, baseElbow)
    for c in nearConns:
        if c.Owner.Id == elbow.Id:
            elbow_Conn = c.Origin
        else:
            baseElbow_Conn = c.Origin   
    transVector = baseElbow_Conn - elbow_Conn 
    transform = Autodesk.Revit.DB.Transform.CreateTranslation(transVector)
    TransactionManager.Instance.EnsureInTransaction(doc)
    ElementTransformUtils.MoveElement(doc, elbow.Id, transVector)
    TransactionManager.Instance.TransactionTaskDone()
    return elbow_Conn, baseElbow_conn
elbow = pickFittingOrAccessory()
baseElbow = pickFittingOrAccessory()
fittingConns = list(elbow.MEPModel.ConnectorManager.Connectors.GetEnumerator())
fittingConnsOrigin = list(c.Origin for c in fittingConns)
newLocationOfElbow = rotateAndMoveElbow(elbow, baseElbow)

OUT = newLocationOfElbow