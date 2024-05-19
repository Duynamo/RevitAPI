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

if isinstance(IN[0],list):
	mainPipe = UnwrapElement(IN[0])
else:
	mainPipe = [UnwrapElement(IN[0])]
if isinstance(IN[1],list):
	branchPipe = UnwrapElement(IN[1])
else:
	branchPipe = [UnwrapElement(IN[1])]
	
	
def closest_connectors(pipe1, pipe2, pipe3):
	conn1 = pipe1.ConnectorManager.Connectors
	conn2 = pipe2.ConnectorManager.Connectors
	conn3 = pipe3.ConnectorManager.Connectors
	
	dist1 = 100000000
	dist2 = 100000000
	connSet = []
	for c in conn1:
		for d in conn3:
			connDist = c.Origin.DistanceTo(d.Origin)
			if connDist < dist1:
				dist1 = connDist
				c1 = c
				d1 = d
		for e in conn2:
			connDist = c.Origin.DistanceTo(e.Origin)
			if connDist < dist2:
				dist2 = connDist
				e1 = e
		connSet = [c1,d1,e1]
	return connSet
k=304.8
TransactionManager.Instance.EnsureInTransaction(doc)
for mainPipe,branchPipe in zip(mainPipe,branchPipe):
	width = branchPipe.Diameter
	#get startPoint and endpoint of curve1
	mainPipeLine = mainPipe.Location.Curve
	mainPipeLen  = mainPipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
	startPoint   = mainPipeLine.GetEndPoint(0)
	endPoint     = mainPipeLine.GetEndPoint(1)
	
	#get point perpendicular of curve2
	branchPipeLine = branchPipe.Location.Curve
	branchPipeSP   = branchPipeLine.GetEndPoint(0)
	pointMid       = mainPipeLine.Project(branchPipeSP).XYZPoint
	len1 = startPoint.DistanceTo(pointMid)
	len2 = len1 - width/2
	len3 = len1 + width/2
	
	midStart = mainPipeLine.Evaluate((len2/mainPipeLen),True)
	midEnd   = mainPipeLine.Evaluate((len3/mainPipeLen),True)
	
	#change startPoint and endpoints so the longest piece of mepCurve is retained
	toggle = False
	if startPoint.DistanceTo(pointMid) < endPoint.DistanceTo(pointMid):
		toggle = True
		tempPoint = startPoint
		startPoint = endPoint
		endPoint = tempPoint
		
		tempMid = midStart
		midStart = midEnd
		midEnd = tempMid
		
	mainPipeConn = mainPipe.ConnectorManager.Connectors
	for conn in mainPipeConn:
		if conn.Origin.IsAlmostEqualTo(startPoint):
			startConn = conn
		elif conn.Origin.IsAlmostEqualTo(endPoint):
			endConn = conn
	
	otherFitting = None
	for conn in endConn.AllRefs:
		if conn.IsConnectedTo(endConn):
			if conn.Owner.GetType() == FamilyInstance:
				otherFitting = conn.Owner
			endConn.DisconnectFrom(conn)
			
	#shorten existing curve and copy it
	if toggle:
		mainPipe.Location.Curve = Line.CreateBound(midStart,startPoint)
	else:
		mainPipe.Location.Curve = Line.CreateBound(startPoint,midEnd)
		
	doc.Regenerate()
	OffsetZ = (midStart.Z - endPoint.Z)*-1
	OffsetX = (midStart.X - endPoint.X)*-1
	OffsetY = (midStart.Y - endPoint.Y)*-1
	direction = XYZ(OffsetX,OffsetY,OffsetZ)
	newElem = ElementTransformUtils.CopyElement(doc,mainPipe.Id,direction)
	curve3 = doc.GetElement(newElem[0])
	doc.Regenerate()
	
TransactionManager.Instance.TransactionTaskDone()
#Assign your output to the OUT variable.
dynMidStart = Autodesk.DesignScript.Geometry.Point.ByCoordinates(midStart.X*k, midStart.Y*k, midStart.Z*k)
dynMidEnd = Autodesk.DesignScript.Geometry.Point.ByCoordinates(midEnd.X*k, midEnd.Y*k, midEnd.Z*k)






OUT = dynMidStart, dynMidEnd