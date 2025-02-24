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
"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""
class ISelectionFilter_ModelLine(Autodesk.Revit.UI.Selection.ISelectionFilter):
	def AllowElement(self, element):
		if type(element) == ModelLine:
			return True

#region pickPoints
def pickPoints(doc):
	TransactionManager.Instance.EnsureInTransaction(doc)
	activeView = doc.ActiveView
	iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
	sketchPlane = SketchPlane.Create(doc, iRefPlane)
	doc.ActiveView.SketchPlane = sketchPlane
	condition = True
	pointsList = []
	dynPList = []
	n = 0
	msg = "Pick Points on Current Section plane, hit ESC when finished."
	TaskDialog.Show("^------^", msg)
	while condition:
		try:
			pt=uidoc.Selection.PickPoint()
			n += 1
			pointsList.append(pt)				
		except :
			condition = False
	doc.Delete(sketchPlane.Id)	
	for j in pointsList:
		dynP = j.ToPoint()
		dynPList.append(dynP)
	TransactionManager.Instance.TransactionTaskDone()			
	return dynPList, pointsList
#endregion

#Create midpoints between 2 refpoints
def MidPointsBySegment(P1, P2, segment_length):
	count = P1.DistanceTo(P2)//(segment_length/304.8)
	direction = (P2-P1).Normalize()
	midPoints = []
	for i in range(int(count)):
		point = P1.Add(direction.Multiply((i+1)*segment_length/304.8))
		midPoints.append(point)
	return midPoints
#Create section at define point
def CreateViewSectionAtPoint(doc, point, direction_vector, box_height, box_width, box_depth):
	t = Transaction(doc, "Create Section")
	t.Start()
	transform = Transform.Identity
	transform.Origin = point
	transform.BasisX = XYZ(0, 0, 1)
	transform.BasisY = direction_vector.CrossProduct(transform.BasisX)
	transform.BasisZ = direction_vector

	section_box = BoundingBoxXYZ()
	section_box.Min = XYZ(-box_height/(2*304.8), -box_width/(2*304.8), 0)
	section_box.Max = XYZ(box_height/(2*304.8), box_width/(2*304.8), box_depth/304.8)
	section_box.Transform = transform

	section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
	new_section = ViewSection.CreateSection(doc, section_type_id, section_box)

	t.Commit()
	return new_section.Name

def pipeCreateFromPoints(desPointsList, sel_pipingSystem, sel_PipeType, sel_Level, diameter):
    lst_Points1 = [i for i in desPointsList]
    lst_Points2 = [i for i in desPointsList[1:]]
    linesList = []
    for pt1, pt2 in zip(lst_Points1,lst_Points2):
        line =  Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(pt1,pt2)
        linesList.append(line)
    firstPoint   = [x.StartPoint for x in linesList]
    secondPoint  = [x.EndPoint for x in linesList]
    pipesList = []
    pipesList1 = []
    TransactionManager.Instance.EnsureInTransaction(doc)
    for i,x in enumerate(firstPoint):
        try:
            levelId = sel_Level.Id
            sysTypeId = sel_pipingSystem.Id
            pipeTypeId = sel_PipeType.Id
            diam = diameter
            pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,sysTypeId,pipeTypeId,levelId,x.ToXyz(),secondPoint[i].ToXyz())
            param_Diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            param_Diameter.SetValueString(diam.ToString())
            TransactionManager.Instance.EnsureInTransaction(doc)
            pipesList.append(pipe.ToDSType(False))
            pipesList1.append(pipe)
            # param_Length.Set()
            TransactionManager.Instance.TransactionTaskDone()
        except:
            pipesList.append(None)				
    TransactionManager.Instance.TransactionTaskDone()
    return pipesList
#endregion


#Create sections along ModelLine
# segment_length = 1000
# points = pickPoints(doc)
# section_names = []
# for i in range(len(points)-1):
# 	refPoints_start = points[i]
# 	refPoints_end = points[i+1]
# 	section_points = MidPointsBySegment(refPoints_start, refPoints_end, segment_length)
# 	direction_vector = (refPoints_end-refPoints_start).Normalize()
# 	for point in section_points:
# 		new_section = CreateViewSectionAtPoint(doc, point, direction_vector, 5000, 5000, 100)
# 		section_names.append(new_section)
#region pickPipe and create elbow
class selectionFilter(ISelectionFilter):
	def __init__(self, category):
		self.category = category
	def AllowElement(self, element):
		if element.Category.Name == self.category:
			return True
		else:
			return False
def pickPipes():
	pipes = []
	pipeFilter = selectionFilter("Pipes")
	pipesRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipes")
	for ref in pipesRef:
		pipe = doc.GetElement(ref.ElementId)
		pipes.append(pipe)
	return pipes	      
#endregion
def closestConn1(pipe1, pipe2):
    if not pipe1 or not pipe2:
        return None, None
    connectors1 = pipe1.ConnectorManager.Connectors
    connectors2 = pipe2.ConnectorManager.Connectors
    min_distance = float('inf')
    closest_conn1 = None
    closest_conn2 = None
    
    for conn1 in connectors1:
        if conn1.IsConnected:
            continue
        for conn2 in connectors2:
            if conn2.IsConnected:
                continue
            distance = conn1.Origin.DistanceTo(conn2.Origin)
            if distance < min_distance:
                min_distance = distance
                closest_conn1 = conn1
                closest_conn2 = conn2
    
    return closest_conn1, closest_conn2
"""___"""
# pipingSystems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
# pipeType = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeSegment).ToElements()
# refLevel = FilteredElementCollector(doc).OfClass(Level).ToElements


pipes = pickPipes()
# conn1 = closetConn(pipes[0], pipes[1])
# conn2 = closetConn(pipes[1], pipes[0])
closestConns = closestConn1(pipes[0], pipes[1])
TransactionManager.Instance.EnsureInTransaction(doc)		
fittings = []
try:
    fitting = doc.Create.NewElbowFitting(closestConns[0], closestConns[1])
    fittings.append(fitting.ToDSType(False))
except:
    pass
TransactionManager.Instance.TransactionTaskDone()
OUT = fittings, conn2, conn1