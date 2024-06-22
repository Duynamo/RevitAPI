import clr
import math
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

from System.Collections.Generic import List

if isinstance(IN[0], list):
	ducts = UnwrapElement(IN[0])
else:
	ducts = [UnwrapElement(IN[0])]
	
if isinstance(IN[1], list):
	points = IN[1]
else:
	points = [IN[1]]

if isinstance(IN[2], list):
	famtype = UnwrapElement(IN[2])
else:
	famtype = [UnwrapElement(IN[2])]
ftl = len(famtype)

def IsParallel(dir1,dir2):
	if dir1.Normalize().IsAlmostEqualTo(dir2.Normalize()):
		return True
	if dir1.Normalize().Negate().IsAlmostEqualTo(dir2.Normalize()):
		return True
	return False
	
def measure(startpoint, point):
	return startpoint.DistanceTo(point)
	
def copyElement(element, oldloc, loc):
	TransactionManager.Instance.EnsureInTransaction(doc)
	elementlist = List[ElementId]()
	elementlist.Add(element.Id)
	OffsetZ = (oldloc.Z - loc.Z)*-1
	OffsetX = (oldloc.X - loc.X)*-1
	OffsetY = (oldloc.Y - loc.Y)*-1
	direction = XYZ(OffsetX,OffsetY,OffsetZ)
	newelementId = ElementTransformUtils.CopyElements(doc,elementlist,direction)
	newelement = doc.GetElement(newelementId[0])
	TransactionManager.Instance.TransactionTaskDone()
	return newelement

def GetDirection(faminstance):
	for c in faminstance.MEPModel.ConnectorManager.Connectors:
		conn = c
		break
	return conn.CoordinateSystem.BasisZ

def GetClosestDirection(faminstance, lineDirection):
	conndir = None
	flat_linedir = XYZ(lineDirection.X,lineDirection.Y,0).Normalize()
	for conn in faminstance.MEPModel.ConnectorManager.Connectors:
		conndir = conn.CoordinateSystem.BasisZ
		if flat_linedir.IsAlmostEqualTo(conndir):
			return conndir
	return conndir
	


#global variables for rotating new families
tempfamtype = None
xAxis = XYZ(1,0,0)
def placeFitting(duct,point,familytype,lineDirection):
	toggle = False
	isVertical = False
	global tempfamtype
	if tempfamtype == None:
		tempfamtype = familytype
		toggle = True
	elif tempfamtype != familytype:
		toggle = True
		tempfamtype = familytype
	level = duct.ReferenceLevel
	
	width = 4
	height = 4
	radius = 2
	round = False
	connectors = duct.ConnectorManager.Connectors
	for c in connectors:
		if c.ConnectorType != ConnectorType.End:
			continue
		shape = c.Shape
		if shape == ConnectorProfileType.Round:
			radius = c.Radius
			round = True	
			break
		elif shape == ConnectorProfileType.Rectangular or shape == ConnectorProfileType.Oval:
			if abs(lineDirection.Z) == 1:
				isVertical = True
				yDir = c.CoordinateSystem.BasisY
			width = c.Width
			height = c.Height
			break
		
	TransactionManager.Instance.EnsureInTransaction(doc)		
	point = XYZ(point.X,point.Y,point.Z-level.Elevation)
	newfam = doc.Create.NewFamilyInstance(point,familytype,level,Structure.StructuralType.NonStructural)
	doc.Regenerate()
	transform = newfam.GetTransform()
	axis = Line.CreateUnbound(transform.Origin, transform.BasisZ)
	global xAxis
	if toggle:
		xAxis = GetDirection(newfam)
	zAxis = XYZ(0,0,1)
	
	if isVertical:
		angle = xAxis.AngleOnPlaneTo(yDir,zAxis)
	else:
		angle = xAxis.AngleOnPlaneTo(lineDirection,zAxis)
	
	
	ElementTransformUtils.RotateElement(doc,newfam.Id,axis,angle)
	doc.Regenerate()
	
	if lineDirection.Z != 0:
		newAxis = GetClosestDirection(newfam,lineDirection)
		yAxis = newAxis.CrossProduct(zAxis)
		angle2 = newAxis.AngleOnPlaneTo(lineDirection,yAxis)
		axis2 = Line.CreateUnbound(transform.Origin, yAxis)
		ElementTransformUtils.RotateElement(doc,newfam.Id,axis2,angle2)
	
	result = {}
	connpoints = []
	famconns = newfam.MEPModel.ConnectorManager.Connectors
	
	if round:
		for conn in famconns:
			if IsParallel(lineDirection,conn.CoordinateSystem.BasisZ) == False:
				continue
			if conn.Shape != shape:
				continue
			conn.Radius = radius
			connpoints.append(conn.Origin)
	else:
		for conn in famconns:
			if IsParallel(lineDirection,conn.CoordinateSystem.BasisZ) == False:
				continue
			if conn.Shape != shape:
				continue
			conn.Width = width
			conn.Height = height
			connpoints.append(conn.Origin)
	TransactionManager.Instance.TransactionTaskDone()
	result[newfam] = connpoints
	return result
	
	
def ConnectElements(duct, fitting):
	ductconns = duct.ConnectorManager.Connectors
	fittingconns = fitting.MEPModel.ConnectorManager.Connectors
	
	TransactionManager.Instance.EnsureInTransaction(doc)
	for conn in fittingconns:
		for ductconn in ductconns:
			result = ductconn.Origin.IsAlmostEqualTo(conn.Origin)
			if result:
				ductconn.ConnectTo(conn)
				break
	TransactionManager.Instance.TransactionTaskDone()
	return result

def SortedPoints(fittingspoints,ductStartPoint):
	sortedpoints = sorted(fittingspoints, key=lambda x: measure(ductStartPoint, x))
	return sortedpoints
	

ductsout = []
fittingsout = []
combilist = []

TransactionManager.Instance.EnsureInTransaction(doc)
for typ in famtype:
	if typ.IsActive == False:
		typ.Activate()
		doc.Regenerate()
TransactionManager.Instance.TransactionTaskDone()

for i,duct in enumerate(ducts):
	ListOfPoints = [x.ToXyz() for x in points[i]]
	familytype = famtype[i%ftl]
	#create duct location line
	ductline = duct.Location.Curve
	ductStartPoint = ductline.GetEndPoint(0)
	ductEndPoint = ductline.GetEndPoint(1)
	#get end connector to reconnect later
	endIsConnected = False
	endrefconn = None
	for ductconn in duct.ConnectorManager.Connectors:
		if ductconn.Origin.DistanceTo(ductEndPoint) < 5/304.8:
			if ductconn.IsConnected:
				endIsConnected = True
				for refconn in ductconn.AllRefs:
					if refconn.ConnectorType != ConnectorType.Logical and refconn.Owner.Id.IntegerValue != duct.Id.IntegerValue:
						endrefconn = refconn
			
	#sort the points from start of duct to end of duct
	pointlist = SortedPoints(ListOfPoints,ductStartPoint)
	
	ductlist = []
	newFittings = []
	ductlist.append(duct)
	
	tempStartPoint = None
	tempEndPoint = None
	
	lineDirection = ductline.Direction
	
	for i,p in enumerate(pointlist):		
		output = placeFitting(duct,p,familytype,lineDirection)
		newfitting = output.keys()[0]
		newFittings.append(newfitting)
		fittingpoints = output.values()[0]
		
		tempPoints = SortedPoints(fittingpoints,ductStartPoint)
		if i == 0:
			tempEndPoint = tempPoints[0]
			tempStartPoint = tempPoints[1]			
			newduct = copyElement(duct,ductStartPoint,tempStartPoint)	
			duct.Location.Curve = Line.CreateBound(ductStartPoint,tempEndPoint)
			ductlist.append(newduct)
			combilist.append([duct,newfitting])
			combilist.append([newduct,newfitting])
		else:
			combilist.append([newduct,newfitting])
			tempEndPoint = tempPoints[0]
			newduct = copyElement(duct,ductStartPoint,tempStartPoint)
			ductlist[-1].Location.Curve = Line.CreateBound(tempStartPoint,tempEndPoint)
			tempStartPoint = tempPoints[1]
			ductlist.append(newduct)
			combilist.append([newduct,newfitting])
	

	
	newline = Line.CreateBound(tempStartPoint,ductEndPoint)
	ductlist[-1].Location.Curve = newline
	
	ductsout.append(ductlist)
	fittingsout.append(newFittings)
	TransactionManager.Instance.EnsureInTransaction(doc)
	doc.Regenerate()
	TransactionManager.Instance.TransactionTaskDone()

	if endIsConnected:
		for conn in ductlist[-1].ConnectorManager.Connectors:
			if conn.Origin.DistanceTo(ductEndPoint) < 5/304.8:
				endrefconn.ConnectTo(conn)
			


TransactionManager.Instance.EnsureInTransaction(doc)
for combi in combilist:
	ConnectElements(combi[0],combi[1])
TransactionManager.Instance.TransactionTaskDone()

OUT = ductsout, fittingsout