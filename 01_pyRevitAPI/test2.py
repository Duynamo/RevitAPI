"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import cos,sin,tan,radians

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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList

def divideLineSegment(doc, pipe, pointA, LengthA):
    TransactionManager.Instance.EnsureInTransaction(doc)
    lineSegment = pipe.Location.Curve
    start_point = lineSegment.GetEndPoint(0)
    end_point = lineSegment.GetEndPoint(1)
    vector = end_point - start_point
    total_length = vector.GetLength()*304.8
    direction = vector.Normalize()
    num_segments = int(total_length / LengthA)
    points = []
    desPoints =[]
    current_point = pointA
    for i in range(num_segments + 1):
        points.append(current_point)
        desPoints.append(current_point)
    TransactionManager.Instance.TransactionTaskDone()
    return desPoints
def splitPipeAtPoints(doc, pipe, points):
    newPipes = []
    points = sorted(points, key=lambda p: (pipe.Location.Curve.GetEndPoint(0) - p).DotProduct(pipe.Location.Curve.Direction))

    for point in points:
        pipeLocation = pipe.Location
        if isinstance(pipeLocation, LocationCurve):
            pipeCurve = pipeLocation.Curve
            if pipeCurve is not None:
                newPipeIds = PlumbingUtils.BreakCurve(doc, pipe.Id, point)
    return newPipes
def divideLineSegment(line, length, startPoint,endPoint):
    # Initialize the list of points
    points = []
    points.append(startPoint)
    current_point = startPoint
    total_length = line.Length
    direction = (endPoint - startPoint).Normalize()
    # Loop to add points at intervals of 'length'
    while (startPoint.DistanceTo(start_point) + length) <= total_length:
        # Calculate the next point
        current_point = current_point + direction * length

        # Add the next point to the list of points
        points.append(current_point)

    return points
#endregion
pipe   = UnwrapElement(IN[0])
splLength = 1000
children = {}
parents, childs, fittings = [], [], []
# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
pipeCurve = pipe.Location.Curve
conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
orgiginConns = list(c.Origin for c in conns)
sortConns = sorted(orgiginConns, key=lambda c : c.Y)
#splitPoints = divideLineSegment(doc, pipe, sortConns[0], splLength)
#list = list(c.ToPoint() for c in splitPoints)
points = divideLineSegment(pipeCurve, splLength, sortConns[0], sortConns[1])

TransactionManager.Instance.TransactionTaskDone()

OUT = points