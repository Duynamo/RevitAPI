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

"""___"""
class ISelectionFilter_ModelLine(Autodesk.Revit.UI.Selection.ISelectionFilter):
	def AllowElement(self, element):
		if type(element) == ModelLine:
			return True

#Create midpoints between 2 refpoints
def MidPointsBySegment(P1, P2, segment_length):
	count = P1.DistanceTo(P2)//(segment_length/304.8)
	direction = (P2-P1).Normalize()
	midPoints = []
	for i in range(int(count)):
		point = P1.Add(direction.Multiply((i+1)*segment_length/304.8))
		midPoints.append(point)
	return midPoints

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

#Create sections along ModelLine
class ISelectionFilter_ModelLine(Autodesk.Revit.UI.Selection.ISelectionFilter):
	def AllowElement(self, element):
		if type(element) == ModelLine:
			return True

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

#region ______to connect pipes by Elbows
def createElbows(pipes):
    fittings = []
    connectors = {}
    connlist = []    
    for pipe in pipes:
        conns = pipe.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                continue
            connectors[conn] = None
            connlist.append(conn)
    for k in connectors.keys():
        mindist = 1000000
        closest = None
        for conn in connlist:
            if conn.Owner.Id.Equals(k.Owner.Id):
                continue
            dist = k.Origin.DistanceTo(conn.Origin)
            if dist < mindist:
                mindist = dist
                closest = conn
        if mindist > margin:
            continue
        connectors[k] = closest
        connlist.remove(closest)
        try:
            del connectors[closest]
        except:
            pass
    for k,v in connectors.items():
        TransactionManager.Instance.EnsureInTransaction(doc)		
        try:
            fitting = doc.Create.NewElbowFitting(k,v)
            fittings.append(fitting.ToDSType(False))
        except:
            pass
        TransactionManager.Instance.TransactionTaskDone()
    return fittings
    #endregion

#Create sections along ModelLine
segment_length = 1000
points = pickPoints(doc)
section_names = []
for i in range(len(points)-1):
	refPoints_start = points[i]
	refPoints_end = points[i+1]
	section_points = MidPointsBySegment(refPoints_start, refPoints_end, segment_length)
	direction_vector = (refPoints_end-refPoints_start).Normalize()
	for point in section_points:
		new_section = CreateViewSectionAtPoint(doc, point, direction_vector, 5000, 5000, 100)
		section_names.append(new_section)
      
def closetConn(mPipe, bPipe):
    connectors1 = list(bPipe.ConnectorManager.Connectors.GetEnumerator())
    Connector1 = NearestConnector(connectors1, mPipe)
    XYZconn	= Connector1.Origin
    return XYZconn
#endregion
#region ___get 3 closest connectors of 3 pipe
def NearestConnector(ConnectorSet, curCurve):
    MinLength = float("inf")
    result = None  # Initialize result to None
    for n in ConnectorSet:
        distance = curCurve.Location.Curve.Distance(n.Origin)
        if distance < MinLength:
            MinLength = distance
            result = n
    return result

pipingSystems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
pipeType = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeSegment).ToElements()
refLevel = FilteredElementCollector(doc).OfClass(Level).ToElements



OUT = section_names