#region __boilerplate
import clr
import sys 
import System   
import math

from math import *

clr.AddReference("DSCoreNodes")
from DSCore.List import Flatten

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *
from Autodesk.DesignScript.Geometry import Line, Point

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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

def getDirectionVector(pipe):
    curve = pipe.Location.Curve
    directionVector = curve.GetEndPoint(1) - curve.GetEndPoint(0)
    return directionVector.Normalize()

#region ___to calculate the angle between two pipe

def getDirectionVector(pipe):
    curve = pipe.Location.Curve
    directionVector = curve.GetEndPoint(1) - curve.GetEndPoint(0)
    return directionVector.Normalize()

def calculateAngleBetweenPipes(mPipe, bPipe):
    dir1 = getDirectionVector(mPipe)
    dir2 = getDirectionVector(bPipe)
    dotProduct = dir1.DotProduct(dir2)
    mag1 = dir1.GetLength()
    mag2 = dir2.GetLength()
    cosAngle = dotProduct / (mag1 * mag2)
    angleRadians = acos(cosAngle)
    angleDegrees = degrees(angleRadians)
    angle1 = 180-angleDegrees
    minAngle = 0
    maxAngle = 0
    if angle1 < angleDegrees:
          minAngle = angle1
          maxAngle = angleDegrees
    else: 
          minAngle = angleDegrees
          maxAngle = angle1
    return minAngle,maxAngle
#endregion

def planeFromPointAndLine(point, line):
    direction = line.Direction
    up_vector = XYZ.BasisZ
    normal = direction.CrossProduct(up_vector).Normalize()
    plane = Plane.CreateByNormalAndOrigin(normal, point) 
    return plane

def offsetPointAlongVector(point, vector, offsetDistance):
    point = point.ToPoint()
    direction = vector.ToVector().Normalize()
    scaledVector = direction.Multiply(offsetDistance)
    offsetPoint = point.ToRevitType().Add(scaledVector)
    return offsetPoint

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

def getAllPipeAccessoriesInActiveView(doc):
    activeViewId = doc.ActiveView.Id
    collector = FilteredElementCollector(doc, activeViewId).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType()
    pipeAccessories = []
    for element in collector:
        pipeAccessories.append(element)
    return pipeAccessories

def pickObjects():
    from Autodesk.Revit.UI.Selection import ObjectType
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs], [ref for ref in refs]

def getAllPipingSystemsInActiveView(doc):
	collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
		pipingSystemsName.append(systemName)
	return pipingSystemsName, pipingSystems

def getPipeLocationCurve(pipe):
    lCurve = pipe.Location.Curve
    return lCurve

def allCateInPJ(doc):
    categories = doc.Settings.Categories
    modelCate = []
    for c in categories:
        if c.CategoryType == CategoryType.Model and c.CanAddSubcategory:
            modelCate.append(Revit.Elements.Category.ById(c.Id.IntegerValue))
    return modelCate

def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)

def getPipeParameter(p):
    paramDiameters = []
    paramPipingSystems = []
    paramLevels = []
    paramPipeTypes = []

    paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
    
    paramPipeTypeId = p.GetTypeId()
    paramPipeType = doc.GetElement(paramPipeTypeId)
    paramPipeTypes.append(paramPipeType)
    pipeTypeName = paramPipeType.LookupParameter("Type Name").AsString()
	
    paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    paramPipingSystem = doc.GetElement(paramPipingSystemId)
    paramPipingSystems.append(paramPipingSystem)
    pipingSystemName = paramPipingSystem.LookupParameter("System Classification").AsValueString()

    paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    paramLevel = doc.GetElement(paramLevelId)
    paramLevels.append(paramLevel)

    return [paramDiameter, paramPipingSystem, paramPipeType, paramLevel],[paramDiameter,pipingSystemName,pipeTypeName,paramLevel]

def PickPoints(doc):
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
		dynP = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X*304.8, j.Y*304.8, j.Z*304.8)	
		dynPList.append(dynP)
	TransactionManager.Instance.TransactionTaskDone()			
	return dynPList, pointsList

def pickObjects():
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs], [ref for ref in refs]

def pickEdges():
    from Autodesk.Revit.UI.Selection import ObjectType
    refs = uidoc.Selection.PickObjects(ObjectType.Edge, "Please select the Desired Faces")
    return  [doc.GetElement(i.ElementId) for i in refs], [ref for ref in refs]

def pickFaces():
    from Autodesk.Revit.UI.Selection import ObjectType
    refs = uidoc.Selection.PickObjects(ObjectType.Face, "Please select the Desired Faces")
    return  [doc.GetElement(i.ElementId) for i in refs], [ref for ref in refs]

#region ___to pick pipe only
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

def calculateAngleBetweenPipes(mPipe, bPipe):
    curve1 = mPipe.Location.Curve
    v1 = curve1.GetEndPoint(1)-curve1.GetEndPoint(0)
    dir1 = v1.Normalize()
    curve2 = bPipe.Location.Curve
    v2 = curve2.GetEndPoint(1)-curve2.GetEndPoint(0)
    dir2 = v2.Normalize()
    
    dotProduct = dir1.DotProduct(dir2)
    mag1 = dir1.GetLength()
    mag2 = dir2.GetLength()
    cosAngle = dotProduct / (mag1 * mag2)
    angleRadians = acos(cosAngle)
    angleDegrees = degrees(angleRadians)
    angle1 = 180-angleDegrees
    minAngle = 0
    maxAngle = 0
    if angle1 < angleDegrees:
          minAngle = angle1
          maxAngle = angleDegrees
    else: 
          minAngle = angleDegrees
          maxAngle = angle1
    return minAngle,maxAngle