import clr
import math

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference("RevitNodes")
import Revit	
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

from System.Collections.Generic import List

if isinstance(IN[0],list):
	curve1 = UnwrapElement(IN[0])
else:
	curve1 = [UnwrapElement(IN[0])]
if isinstance(IN[1],list):
	curve2 = UnwrapElement(IN[1])
else:
	curve2 = [UnwrapElement(IN[1])]
	
	
def closest_connectors(pipe1, pipe2,pipe3):
	conn1 = pipe1.ConnectorManager.Connectors
	conn2 = pipe2.ConnectorManager.Connectors
	conn3 = pipe3.ConnectorManager.Connectors
	
	dist1 = 100000000
	dist2 = 100000000
	connset = []
	for c in conn1:
		for d in conn3:
			conndist = c.Origin.DistanceTo(d.Origin)
			if conndist < dist1:
				dist1 = conndist
				c1 = c
				d1 = d
		for e in conn2:
			conndist = c.Origin.DistanceTo(e.Origin)
			if conndist < dist2:
				dist2 = conndist
				e1 = e
		connset = [c1,d1,e1]
	return connset
	

def newTee(conn1,conn2,conn3):
	fitting = doc.Create.NewTeeFitting(conn1,conn2,conn3)
	return fitting

fittings = []
	
TransactionManager.Instance.EnsureInTransaction(doc)
for curve1,curve2 in zip(curve1,curve2):
	try:
		try:
			width = curve2.Diameter
		except:
			width = curve2.Width
		#get startpoint and endpoint of curve1
		curve1Line = curve1.Location.Curve
		startpoint = curve1Line.GetEndPoint(0)
		endpoint = curve1Line.GetEndPoint(1)
		
		#get point perpendicular of curve2
		curve2Line = curve2.Location.Curve
		curve2start = curve2Line.GetEndPoint(0)
		pointmid = curve1Line.Project(curve2start).XYZPoint
		len1 = startpoint.DistanceTo(pointmid)
		len2 = len1 - width/2
		len3 = len1 + width/2
		
		midstart = curve1Line.Evaluate((len2/curve1Line.Length),True)
		midend = curve1Line.Evaluate((len3/curve1Line.Length),True)
		
		#change startpoint and endpoints so the longest piece of mepcurve is retained
		toggle = False
		if startpoint.DistanceTo(pointmid) < endpoint.DistanceTo(pointmid):
			toggle = True
			temppoint = startpoint
			startpoint = endpoint
			endpoint = temppoint
			
			tempmid = midstart
			midstart = midend
			midend = tempmid
		
		#disconnect curve1:
		connectors1 = curve1.ConnectorManager.Connectors
		for conn in connectors1:
			if conn.Origin.IsAlmostEqualTo(startpoint):
				startconn = conn
			elif conn.Origin.IsAlmostEqualTo(endpoint):
				endconn = conn
		
		otherfitting = None
		for conn in endconn.AllRefs:
			if conn.IsConnectedTo(endconn):
				if conn.Owner.GetType() == FamilyInstance:
					otherfitting = conn.Owner
				endconn.DisconnectFrom(conn)
				
		#shorten existing curve and copy it
		if toggle:
			curve1.Location.Curve = Line.CreateBound(midstart,startpoint)
		else:
			curve1.Location.Curve = Line.CreateBound(startpoint,midstart)
		doc.Regenerate()
		OffsetZ = (midstart.Z - endpoint.Z)*-1
		OffsetX = (midstart.X - endpoint.X)*-1
		OffsetY = (midstart.Y - endpoint.Y)*-1
		direction = XYZ(OffsetX,OffsetY,OffsetZ)
		newelem = ElementTransformUtils.CopyElement(doc,curve1.Id,direction)
		curve3 = doc.GetElement(newelem[0])
		doc.Regenerate()
		
		#shorten new curve
		curve3.Location.Curve = Line.CreateBound(endpoint,midend)
		doc.Regenerate()
		
		connectors = closest_connectors(curve1,curve2,curve3)
	#	try:
		fitting = newTee(connectors[0],connectors[1],connectors[2])
		fittings.append(fitting.ToDSType(False))
		
		if otherfitting != None:
			connectors3 = curve3.ConnectorManager.Connectors
			for conn in connectors3:
				for conn2 in otherfitting.MEPModel.ConnectorManager.Connectors:
					if conn.Origin.IsAlmostEqualTo(conn2.Origin):
						conn.ConnectTo(conn2)
						break
	except:
		fittings.append(None)

	
TransactionManager.Instance.TransactionTaskDone()



#Assign your output to the OUT variable.
OUT = fittings