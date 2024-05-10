#region Setup ban dau
import clr
import sys
sys.path.append('C:\Program Files (x86)\IronPython 2.7\Lib')
import System
from System import Array
from System.Collections.Generic import *
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager 
from RevitServices.Transactions import TransactionManager 

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import Autodesk 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication 
app = uiapp.Application 
uidoc = uiapp.ActiveUIDocument
#endregion

#region PipingSystem
def getAllPipingSystemsName(doc):
	collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		if systemName != None:
		    pipingSystemsName.append(systemName)
	return pipingSystemsName

#OUT = getAllPipingSystemsName(doc)

#endregion

#region PipeType
def getAllPipeTypesName(doc):
	collector1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves)
	pipeTypes = collector1.ToElements()
	pipeTypesName = []
	for pipeType in pipeTypes:
		pipeTypeName = pipeType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		if pipeTypeName != None:
		    pipeTypesName.append(pipeTypeName)
	return pipeTypesName

#OUT = getAllPipeTypesName(doc)

#endregion

#region PickPoints
def PickPoints(doc):
	TransactionManager.Instance.EnsureInTransaction(doc)
	activeView = doc.ActiveView
	iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
	sketchPlane = SketchPlane.Create(doc, iRefPlane)
	doc.ActiveView.SketchPlane = sketchPlane
	
	points = []
	condition = True
	while condition:
		try:
			pickpoint = uidoc.Selection.PickPoint()
			points.append(Point.Create(pickpoint).ToProtoType())
		except:
			condition = False
	return points
	TransactionManager.Instance.TransactionTaskDone()	

#OUT = PickPoints(doc)

#endregion

#region CreatePipe

#points = [XYZ(0, 0, 0),XYZ(10, 10, 0),XYZ(10, 0, 0)]
#systemTypeId = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()[0].Id
#pipeTypeId = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).ToElements()[0].Id
#levelId = FilteredElementCollector(doc).OfClass(Level).ToElements()[0].Id
#diameter = 200

def CreatePipe(doc, systemTypeId, pipeTypeId, levelId, diameter, points):
	TransactionManager.Instance.EnsureInTransaction(doc)
	pipeList = []
	for P1, P2 in zip(points[:-1], points[1:]):
		pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,systemTypeId,pipeTypeId,levelId,P1, P2)
		pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).SetValueString(diameter.ToString())
		pipeList.append(pipe)
	return pipeList
	TransactionManager.Instance.TransactionTaskDone()

#OUT = CreatePipe(doc, systemTypeId, pipeTypeId, levelId, diameter, points)

#endregion

#region CreateElbow

#pipes = UnwrapElement(IN[0])
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

	return elbowList

	TransactionManager.Instance.TransactionTaskDone()

#OUT = elbowList

#endregion


