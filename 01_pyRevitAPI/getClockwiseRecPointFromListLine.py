"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections
import duynamoLibrary as dLib

sys.path.append('')
from math import cos,sin,tan,radians

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

clr.AddReference("DSCoreNodes")
from DSCore.List import Flatten

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
def chunkList(inputList, chunkSize):
    """Chop a list into chunks of specified size."""
    return [inputList[i:i+chunkSize] for i in range(0, len(inputList), chunkSize)]


class XYZ:
    def __init__(self, x, y, z):
        self.x = x
        self.y = y
        self.z = z

    def __lt__(self, other):
        if self.x != other.x:
            return self.x < other.x
        if self.y != other.y:
            return self.y < other.y
        return self.z < other.z

    def __repr__(self):
        return f"XYZ({self.x}, {self.y}, {self.z})"

def min_Point(points):
    minPt = points[0]  # Assume the first point is the minimum initially
    for pt in points[1:]:
        if pt < minPt:
            minPt = pt
    return minPt
def min_y_point(points):
    min_y_pt = points[0]  # Assume the first point is the minimum initially
    for pt in points[1:]:
        if pt.y < min_y_pt.y:
            min_y_pt = pt
    return min_y_pt

#endregion


lineList   = uwList(IN[0])
chunkedList = chunkList(lineList, 4)
startPoints = []
for chunked4Line in chunkedList:
    for ele in chunked4Line:
        startPoints.append(ele.GetEndPoint(0))
    minPoint = min_Point(startPoints)
    startPoints.remove(minPoint)
    if startPoints:
        minYPoint = min_y_point(startPoints)
    startPoints.remove(minYPoint)
    if startPoints:
        minPointB = min_Point(startPoints)
# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)


TransactionManager.Instance.TransactionTaskDone()

OUT = chunkList(lineList),minPoint,minPointB,minYPoint










def minXPoints(points):
    if len(points) < 2:
        return points
    points = sorted(points, key=lambda pt: pt.X)

def minXPoints(points):
    if len(points) < 2:
        return points
    else:
        points = sorted(points, key=lambda pt: pt.X)
        return points[:2]
def minYPoints(points):
    if len(points) < 2:
        return points
    else:
        points = sorted(points, key=lambda pt: pt.Y)
        return points[:2]
def getRoundedDisFrom2Points(point1, point2, k):
    distance = point1.DistanceTo(point2)
    roundedDis = math.ceil(distance/k) * k
    return roundedDis
#endregion


inPoints   = uwList(IN[0])
# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
sort1 = minXPoints(inPoints)
columnB = getRoundedDisFrom2Points(sort1[0],sort1[1],5)
sort2 = minYPoints(inPoints)
columnH = getRoundedDisFrom2Points(sort2[0],sort2[1],5)

TransactionManager.Instance.TransactionTaskDone()

OUT = sort1,sort2,columnB, columnH  