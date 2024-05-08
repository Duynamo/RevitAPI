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
	curve1 = UnwrapElement(IN[0])
else:
	curve1 = [UnwrapElement(IN[0])]
if isinstance(IN[1],list):
	curve2 = UnwrapElement(IN[1])
else:
	curve2 = [UnwrapElement(IN[1])]
	
	
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
	

def newTee(conn1,conn2,conn3):
	fitting = doc.Create.NewTeeFitting(conn1,conn2,conn3)
	return fitting

fittings = []
	
TransactionManager.Instance.EnsureInTransaction(doc)
for curve1,curve2 in zip(curve1,curve2):
	try:
		width = curve2.Diameter
		#get startPoint and endpoint of curve1
		curve1Line = curve1.Location.Curve
		startPoint = curve1Line.GetEndPoint(0)
		endPoint = curve1Line.GetEndPoint(1)
		
		#get point perpendicular of curve2
		curve2Line = curve2.Location.Curve
		curve2start = curve2Line.GetEndPoint(0)
		pointMid = curve1Line.Project(curve2start).XYZPoint
		len1 = startPoint.DistanceTo(pointMid)
		len2 = len1 - width/2
		len3 = len1 + width/2
		
		midStart = curve1Line.Evaluate((len2/curve1Line.Length),True)
		midEnd = curve1Line.Evaluate((len3/curve1Line.Length),True)
		
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
		
		#disconnect curve1:
		connectors1 = curve1.ConnectorManager.Connectors
		for conn in connectors1:
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
			curve1.Location.Curve = Line.CreateBound(midStart,startPoint)
		else:
			curve1.Location.Curve = Line.CreateBound(startPoint,midStart)
		doc.Regenerate()
		OffsetZ = (midStart.Z - endPoint.Z)*-1
		OffsetX = (midStart.X - endPoint.X)*-1
		OffsetY = (midStart.Y - endPoint.Y)*-1
		direction = XYZ(OffsetX,OffsetY,OffsetZ)
		newElem = ElementTransformUtils.CopyElement(doc,curve1.Id,direction)
		curve3 = doc.GetElement(newElem[0])
		doc.Regenerate()
		
		#shorten new curve
		curve3.Location.Curve = Line.CreateBound(endPoint,midEnd)
		doc.Regenerate()
		
		connectors = closest_connectors(curve1,curve2,curve3)
	#	try:
		fitting = newTee(connectors[0],connectors[1],connectors[2])
		fittings.append(fitting.ToDSType(False))
		
		if otherFitting != None:
			connectors3 = curve3.ConnectorManager.Connectors
			for conn in connectors3:
				for conn2 in otherFitting.MEPModel.ConnectorManager.Connectors:
					if conn.Origin.IsAlmostEqualTo(conn2.Origin):
						conn.ConnectTo(conn2)
						break
	except:
		fittings.append(None)

	
TransactionManager.Instance.TransactionTaskDone()



#Assign your output to the OUT variable.
OUT = fittings