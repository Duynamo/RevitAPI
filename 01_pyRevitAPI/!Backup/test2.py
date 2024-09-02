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
from Autodesk.Revit.DB.Plumbing import*
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
"""_____"""
#region ___def pickPipe
class selectionFilter(ISelectionFilter):
	def __init__(self, category):
		self.category = category
	def AllowElement(self, element):
		if element.Category.Name == self.category:
			return True
		else:
			return False
def pickPipe():
	pipeFilter = selectionFilter("Pipes")
	ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipe")
	pipe = doc.GetElement(ref.ElementId)
	return pipe
#endregion
#region ___def getPipeLocationCurve
def getPipeLocationCurve(pipe):
    lCurve = pipe.Location.Curve
    return lCurve
#endregion
#region ___def flatten
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
#endregion
#region ___def devideLineSegment
def divideLineSegment(line, length, startPoint, endPoint):
    points = []
    total_length = line.Length
    direction = (endPoint - startPoint).Normalize()
    current_point = startPoint
    points.append(current_point.ToPoint())
    while (current_point.DistanceTo(startPoint) + length) <= total_length:
        current_point = current_point + direction * length
        points.append(current_point.ToPoint())
    return points
#endregion
def splitPipeAtPoints(doc, pipe, points):
    newPipes = []
    currentPipe = pipe
    originalPipe = pipe
    TransactionManager.Instance.EnsureInTransaction(doc)
    for point in points:
        pipeLocation = currentPipe.Location
        if isinstance(pipeLocation, LocationCurve):
            pipeCurve = pipeLocation.Curve
            if pipeCurve is not None:
                if is_point_on_curve(pipeCurve, point):
                    newPipeIds = PlumbingUtils.BreakCurve(doc, currentPipe.Id, point)
                    if isinstance(newPipeIds, list):
                        newPipe = doc.GetElement(newPipeIds[0])  # Assuming the first element is the new pipe
                        newPipes.append(newPipe)
                        currentPipe = newPipe
                    else:
                        currentPipe = doc.GetElement(newPipeIds)
                else:
                    # Use the original pipe for splitting
                    newPipeIds = PlumbingUtils.BreakCurve(doc, originalPipe.Id, point)
                    if isinstance(newPipeIds, list):
                        newPipe = doc.GetElement(newPipeIds[0])  # Assuming the first element is the new pipe
                        newPipes.append(newPipe)
                        currentPipe = newPipe
                    else:
                        currentPipe = doc.GetElement(newPipeIds)
    TransactionManager.Instance.TransactionTaskDone()
    return newPipes

def splitElementAtPoints(doc, element, points):
    newElements = []
    currentElement = element
    originalElement = element
    TransactionManager.Instance.EnsureInTransaction(doc)
    for point in points:
        location = getElementLocationCurve(currentElement)
        if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericModel):
            # Split the generic model at the point
            curveStart = location.GetEndPoint(0)
            curveEnd = location.GetEndPoint(1)
            
            if curveStart.DistanceTo(point) > 1e-6 and curveEnd.DistanceTo(point) > 1e-6:
                # Create two new curves for the generic model
                curve1 = Line.CreateBound(curveStart, point)
                curve2 = Line.CreateBound(point, curveEnd)
                
                # Set the original generic model to the first curve
                location.Curve = curve1
                
                # Duplicate the generic model for the second curve
                newElementId = ElementTransformUtils.CopyElement(doc, currentElement.Id, XYZ.Zero).ElementAt(0)
                newElement = doc.GetElement(newElementId)
                newElement.Location.Curve = curve2
                newElements.append(newElement)
                currentElement = newElement
            else:
                # Handle the case where the split point is at or near an endpoint
                newElements.append(currentElement)
    else:
        # Handle elements without a LocationCurve (e.g., FamilyInstance)
        pass


def is_point_on_curve(curve, point):
    projected_point = curve.Project(point)
    return projected_point.Distance < 1e-6
def pickObject():
	ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
	ele = doc.GetElement(ref.ElementId)
	return  ele

def getElementLocationCurve(element):
    # Check if the element is a pipe
    if element.Category.Name == 'Pipes':
        locationCurve = element.Location.Curve
        return locationCurve

    # Check if the element is a generic model
    elif element.Category.Name == 'Generic Models':
        if isinstance(element, FamilyInstance):
            location = element.Location
            if hasattr(location, 'Curve'):
                locationCurve = location.Curve
                return locationCurve
        else:
            return None
    else:
        return None

pipe = pickObject()
new_dynPoints = []
try:
    if pipe is not None:
        splitNumber1 = 2
        if splitNumber1.strip():
            try:
                splitNumber = int(splitNumber1)
            except ValueError:
                splitNumber = None
        else:
            splitNumber = None
        '''__________'''
        splitLength1 = 1500
        if splitLength1.strip():
            try:
                splitLength = int(splitLength1)/304.8
            except ValueError:
                splitLength = None
        else:
            splitLength = None
        '''___'''
        if hasattr(pipe, 'Pipes'):
            locationCurve = pipe.Location.Curve
        if hasattr(pipe, 'GenericModel'):
            locationCurve = pipe.GenericModel.Location



        pipeCurve  = getElementLocationCurve(pipe)
        originConns = []
        originConns.append(pipeCurve.GetEndPoint(0))
        originConns.append(pipeCurve.GetEndPoint(1))

        # conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
        # originConns = list(c.Origin for c in conns)
        sortByCor_case1 = True
        sortByCor_case2 = False
        sortByCor_case3 = False
        '''___'''
        minCase = True
        maxCase = False
        '''___'''
        if sortByCor_case1 == True and minCase == True :
            sortConns = sorted(originConns, key=lambda c : c.X)			
        elif sortByCor_case1 == True and maxCase == True :
            sortConns = sorted(originConns, key=lambda c : c.X)
            sortConns.reverse()
        elif sortByCor_case2 == True and minCase == True :
            sortConns = sorted(originConns, key=lambda c : c.Y)			
        elif sortByCor_case2 == True and maxCase == True :
            sortConns = sorted(originConns, key=lambda c : c.Y)
            sortConns.reverse()
        elif sortByCor_case3 == True and minCase == True :
            sortConns = sorted(originConns, key=lambda c : c.Z)			
        elif sortByCor_case3 == True and maxCase == True :
            sortConns = sorted(originConns, key=lambda c : c.Z)
            sortConns.reverse()		
        points = divideLineSegment(pipeCurve, splitLength, sortConns[0], sortConns[1])
        
        dynPoints = list(c.ToRevitType() for c in points)
        if splitNumber <= len(dynPoints):
            new_dynPoints = dynPoints[1:splitNumber+1]
        else:
            new_dynPoints = dynPoints[1:]
except Exception as e:
    TransactionManager.Instance.ForceCloseTransaction()
    pass
TransactionManager.Instance.EnsureInTransaction(doc)
newPipes = splitElementAtPoints(doc, pipe, new_dynPoints)
TransactionManager.Instance.TransactionTaskDone
pass
locationCurve = pipe.Location

OUT = locationCurve