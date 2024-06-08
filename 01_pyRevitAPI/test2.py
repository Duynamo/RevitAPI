
import clr
import math
import duynamoLibrary as dLib

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
#The inputs to this node will be stored as a list in the IN variables.
if isinstance(IN[0], list):
	fittings = UnwrapElement(IN[0])
	toggle = 0
else:
	toggle = 1
	fittings = [UnwrapElement(IN[0])]
	
	
def getConnSysType(connector):
	domain = connector.Domain
	if domain == Domain.DomainHvac:
		return connector.DuctSystemType.ToString()
	elif domain == Domain.DomainPiping:
		return connector.PipeSystemType.ToString()
	elif domain == Domain.DomainElectrical:
		return connector.ElectricalSystemType.ToString()
	else:
		return None
	

p = []
dir = []
fd = []
ref = []
conns = []
descript = []
H = []
W = []
R = []
MEP = []
Sys = []
Shap = []
sysClass = []

for fitting in fittings:
	
	
	points = []
	direction = []
	flowdir = []
	refs = []
	connlist = []
	description = []
	height = []
	width = []
	radius = []
	MEPSystem = []
	systemType = []
	shape = []
	systemclass = []
	
	try:
		connectors = fitting.MEPModel.ConnectorManager.Connectors
	except:
		try:
			connectors = fitting.ConnectorManager.Connectors
		except:			
			p.append(None)
			dir.append(None)
			fd.append(None)
			ref.append(None)
			conns.append(None)
			descript.append(None)
			H.append(None)
			W.append(None)
			R.append(None)
			Shap.append(None)
			MEP.append(None)
			Sys.append(None)
			sysClass.append(None)
			continue
	for conn in connectors:
		connlist.append(conn)
		description.append(conn.Description)
		try:
			height.append(conn.Height)
			width.append(conn.Width)
			radius.append(None)
		except:
			try:
				radius.append(conn.Radius)
				height.append(None)
				width.append(None)
			except:
				radius.append(None)
				height.append(None)
				width.append(None)
		shape.append((conn.Shape).ToString())
		
		try:
			MEPSystem.append(conn.MEPSystem.Name)
			systype = doc.GetElement(conn.MEPSystem.GetTypeId())
			systemType.append(systype)			
		except:
			MEPSystem.append(None)
			systemType.append(None)
		
		try:		
			systemclass.append(getConnSysType(conn))
		except:
			systemclass.append(None)
		
		points.append(conn.Origin.ToPoint())
		direction.append(conn.CoordinateSystem.BasisZ.ToVector())
		
		try:
			flowdir.append(conn.Direction.ToString())
		except:
			flowdir.append(None)
		for c in conn.AllRefs:
			refs.append(c.Owner)
	p.append(points)
	dir.append(direction)
	fd.append(flowdir)
	ref.append(refs)
	conns.append(connlist)
	descript.append(description)
	H.append(height)
	W.append(width)
	R.append(radius)
	MEP.append(MEPSystem)
	Sys.append(systemType)
	Shap.append(shape)
	sysClass.append(systemclass)


#Assign your output to the OUT variable.
if toggle:
	OUT = points, flowdir, refs, direction, connlist, description, height, width, radius, MEPSystem, systemType, shape, systemclass
else:
	OUT = p,fd,ref, dir, conns, descript, H, W, R, MEP, Sys, Shap, sysClass
	


