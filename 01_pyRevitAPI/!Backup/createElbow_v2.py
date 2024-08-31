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
def pickFittingOrAccessory():
    categories = ['Pipe Fittings', 'Pipe Accessories']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory')
    desEle = doc.GetElement(ref.ElementId)
    return desEle
def pickPipe():
    category = ['Pipes']
    pipeFilter = selectionFilter(category)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter, 'Pick Pipe')
    desPipe = doc.GetElement(ref.ElementId)
    return desPipe
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

def CreateElbow(doc, pipes):
	connectors = []
	elbowList = []
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
			elbowList.append(fitting)
	except:
		pass
	return elbowList[0]
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
def pickPoint(doc):
	TransactionManager.Instance.EnsureInTransaction(doc)
	activeView = doc.ActiveView
	iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
	sketchPlane = SketchPlane.Create(doc, iRefPlane)
	doc.ActiveView.SketchPlane = sketchPlane
	condition = True
	pointsList = []
	dynPList = []
	n = 0
	# msg = "Pick Points on Current Section plane, hit ESC when finished."
	# TaskDialog.Show("^------^", msg)
	while condition:
		try:
			pt=uidoc.Selection.PickPoint()
			n += 1
			pointsList.append(pt)				
		except :
			condition = False
	doc.Delete(sketchPlane.Id)	
	# for j in pointsList:
	# 	dynP = j.ToPoint()
	# 	dynPList.append(dynP)
	TransactionManager.Instance.TransactionTaskDone()			
	return pointsList[0]
def findNearestConnector(connectorSet, targetPoint):
    nearest_connector = None
    minDistance = float('inf')
    for connector in connectorSet:
        distance = (connector.Origin - targetPoint).GetLength()
        if distance < minDistance:
            minDistance = distance
            nearest_connector = connector
    return nearest_connector
#endregion
'''___'''
fittingOrAccessory = pickFittingOrAccessory()
basePipe = pickPipe()
pickedPoint = pickPoint(doc)
#region __input Angle
class MyForm(Form):
    def __init__(self):
        self.Text = ''
        self.Size = Size(300, 180)
        self.Font = System.Drawing.Font("Meiryo UI", 7.5, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self.ForeColor = System.Drawing.Color.Red   
        # Create and add label
        self.label = Label()
        self.label.Text = "Angle"
        self.label.Size = Size(260,20)
        self.label.Location = Point(10, 10)
        self.Controls.Add(self.label)       
        # Create and add text box
        self.textBox = TextBox()
        self.textBox.Location = Point(10, 40)
        self.textBox.Size = Size(260, 20)
        self.Controls.Add(self.textBox)      
        # Create and add OK button
        self.okButton = Button()
        self.okButton.Text = 'OK'
        self.okButton.Location = Point(120, 90)
        self.okButton.Click += self.okButton_Click
        self.Controls.Add(self.okButton)  
        # Create and add Cancel button
        self.cancelButton = Button()
        self.cancelButton.Text = 'Cancel'
        self.cancelButton.Location = Point(200, 90)
        self.cancelButton.Click += self.cancelButton_Click
        self.Controls.Add(self.cancelButton)
        # label
        self.fvcLabel = Label()
        self.fvcLabel.Text = "@FVC"
        self.fvcLabel.Size = Size(50, 20)
        self.fvcLabel.Location = Point(10, 110)  # Bottom left corner
        self.Controls.Add(self.fvcLabel)        
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.result = None
    def okButton_Click(self, sender, event):
        self.result = self.textBox.Text
        self.DialogResult = DialogResult.OK
        self.Close()
    def cancelButton_Click(self, sender, event):
        self.DialogResult = DialogResult.Cancel
        self.Close()
form = MyForm()
result = form.ShowDialog()
if result == DialogResult.OK:
    text_input = form.result
else:
    text_input = None
angle = float(text_input)
#endregion
#endregion
'''___'''
#region __copy parameters of base pipe to create temporary  pipe 1 and temporary pipe 2
pipeTypeId = basePipe.PipeType.Id
pipingSystemId = basePipe.MEPSystem.GetTypeId()
levelId = doc.GetElement(basePipe.ReferenceLevel.Id)
diam = basePipe.LookupParameter('Diameter').AsDouble()
#endregion
fittingOrAccessory_conns = fittingOrAccessory.MEPModel.ConnectorManager.Connectors
conn = findNearestConnector(fittingOrAccessory_conns, pickedPoint)
startPoint = conn.Origin
direction = conn.CoordinateSystem.BasisZ
length = basePipe.LookupParameter('Diameter').AsDouble()*5
endPoint1 = startPoint + direction.Multiply(length)
#region __Create temporary pipe 1
TransactionManager.Instance.EnsureInTransaction(doc)
newPipe1 = Pipe.Create(doc,pipingSystemId, pipeTypeId, levelId.Id, startPoint, endPoint1)
newPipeDiam = newPipe1.LookupParameter('Diameter')
newPipeDiam.Set(diam)
# connect = connect2Connectors(doc, newPipe1, fittingOrAccessory)
TransactionManager.Instance.TransactionTaskDone()
#endregion
#region __create temporary pipe 2 and connect temporary pipe 1 to temporary pipe 2 by elbow
rotateAxis = XYZ.BasisZ #rotation axis . asuming rotation axis is
_endPoint2 = endPoint1 + direction.Multiply(length)
rotationTransform = Transform.CreateRotationAtPoint(rotateAxis, math.radians(angle), endPoint1)
endPoint2 = rotationTransform.OfPoint(_endPoint2)
TransactionManager.Instance.EnsureInTransaction(doc)
pipes = []
newPipe2 = Pipe.Create(doc,pipingSystemId, pipeTypeId, levelId.Id, endPoint1, endPoint2)
newPipeDiam = newPipe2.LookupParameter('Diameter')
newPipeDiam.Set(diam)
pipes.append(newPipe1)
pipes.append(newPipe2)
elbow1 = CreateElbow(doc, pipes)
TransactionManager.Instance.TransactionTaskDone()
#endregion
nearConns1 = findNearestConnectorOf2Fittings(fittingOrAccessory, elbow1)
#region ___get near connector of first fitting and new fitting
fittingOrAccessory_Conn = None
elbow1_conn = None
for c in nearConns1:
    if c.Owner.Id == fittingOrAccessory.Id:
        fittingOrAccessory_Conn = c.Origin
    else:
        elbow1_conn = c.Origin
#endregion
TransactionManager.Instance.EnsureInTransaction(doc)
transVector1 = conn.Origin - elbow1_conn
transform1 = Autodesk.Revit.DB.Transform.CreateTranslation(transVector1)
ElementTransformUtils.MoveElement(doc, elbow1.Id, transVector1)
if elbow1:
    connectedEle = None
    conns = list(elbow1.MEPModel.ConnectorManager.Connectors)
    for conn in conns:
         if conn.IsConnected:
              for refConn in conn.AllRefs:
                   connectedElement = refConn.Owner
                   if connectedElement.Id != elbow1.Id:
                        connectedEle = connectedElement
                        del_connectedEle = doc.Delete(connectedEle.Id)
TransactionManager.Instance.TransactionTaskDone()
# connectElbow1_1stFitting = connect2Connectors(doc, fittingOrAccessory,elbow1[0])
'''___'''
OUT = 1