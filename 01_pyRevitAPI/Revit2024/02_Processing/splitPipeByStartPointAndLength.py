#Thanks to Maciek Glowka and Cyril Poupin on the Dynamo Forum

import clr
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
doc = DocumentManager.Instance.CurrentDBDocument

elements = UnwrapElement(IN[0]) if isinstance(IN[0],list) else [UnwrapElement(IN[0])]
points = IN[1] if isinstance(IN[1],list) else [IN[1]]
points = [p.ToXyz() for p in points]

parents,childs,fittings = [],[],[]
children = {}

def colCurve(col):	
	loc= col.Location.Point
	bb = col.get_BoundingBox(None).Min
	height=col.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsDouble()
	start=XYZ(loc.X,loc.Y,bb.Z)
	end=XYZ(start.X,start.Y,start.Z+height)
	curve=Line.CreateBound(start,end)
	return curve

def pointOnLine(p,a,b):
	if not (p-a).CrossProduct(b-a).IsAlmostEqualTo(XYZ.Zero):
		return False
	if a.X!=b.X:
		if a.X<p.X<b.X:
			return True
		if a.X>p.X>b.X:
			return True
	else:
		if a.Y<p.Y<b.Y:
			return True
		if a.Y>p.Y>b.Y:
			return True
	return False
	
def splitWall(ec,p):
	lc = ec.Location.Curve
	v0 = lc.GetEndPoint(0)
	v1 = lc.GetEndPoint(1)
	isOnLine=False
	if pointOnLine(p,v0,v1):
		isOnLine = True
	if isOnLine:
		join_1 = WallUtils.IsWallJoinAllowedAtEnd(ec, 1)
		WallUtils.DisallowWallJoinAtEnd(ec, 1)
		e1Id=ElementTransformUtils.CopyElement(doc, ec.Id, XYZ.Zero)[0]
		e1 = doc.GetElement(e1Id)
		WallUtils.DisallowWallJoinAtEnd(e1, 0)
		nc0 = Autodesk.Revit.DB.Line.CreateBound(v0, p)
		nc1 = Autodesk.Revit.DB.Line.CreateBound(v1, p)
		ec.Location.Curve = nc0
		e1.Location.Curve = nc1
		if join_1 :
			WallUtils.DisallowWallJoinAtEnd(e1, 0)
		WallUtils.AllowWallJoinAtEnd(ec, 1)
		WallUtils.AllowWallJoinAtEnd(e1, 0)
	return e1Id

TransactionManager.Instance.EnsureInTransaction(doc)
for e, p in zip(elements,points):
	to_check = [e]
	if e.Id in children:
		to_check.extend(children[e.Id])
		
	splitId = None
	for ec in to_check:
		if isinstance(ec,Autodesk.Revit.DB.Plumbing.Pipe):
			try:
				splitId = Autodesk.Revit.DB.Plumbing.PlumbingUtils.BreakCurve(doc, ec.Id, p)
				break
			except:
				pass				
		elif isinstance(ec,Autodesk.Revit.DB.Mechanical.Duct):
			try:
				splitId = Autodesk.Revit.DB.Mechanical.MechanicalUtils.BreakCurve(doc, ec.Id, p)
				break
			except:
				pass
		elif isinstance(ec,Autodesk.Revit.DB.FamilyInstance) and ec.CanSplit :
			try:
				if ec.Location.ToString() == 'Autodesk.Revit.DB.LocationCurve':
					curvB = ec.Location.Curve
				elif ElementId(BuiltInCategory.OST_StructuralColumns) == ec.Category.Id and not ec.IsSlantedColumn :
					curvB = colCurve(ec)
				lenBeam = curvB.Length 
				param = (curvB.GetEndPoint(0).DistanceTo(p)) / lenBeam
				splitId = ec.Split(param)
				break
			except:
				pass
		elif isinstance(ec,Autodesk.Revit.DB.Wall):
			try:
				splitId = splitWall(ec,p)
				break
			except:
				pass
	if splitId:
		split = doc.GetElement(splitId)
		if hasattr(split,"ConnectorManager"):
			newPipeConnectors = split.ConnectorManager.Connectors
			#Check
			connA = None
			connB = None
			for c in ec.ConnectorManager.Connectors:
				pc = c.Origin
				#get connectorB near to connectorA
				nearest = [x for x in newPipeConnectors if pc.DistanceTo(x.Origin) < 0.01]
				#If exists assign
				if nearest:
					connA = c
					connB = nearest[0]
			#Create fitting
			try:
				fittings.append(doc.Create.NewUnionFitting(connA, connB))
			except:pass
		
		if e.Id in children:
			children[e.Id].append(split)
		else:
			children[e.Id] = [split]
		parents.append(ec)
		childs.append(split)
	else:
		parents.append(None)
		childs.append(None)
TransactionManager.Instance.TransactionTaskDone()
			
OUT = parents,childs,fittings