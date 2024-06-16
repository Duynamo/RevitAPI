"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

sys.path.append('C:\Program Files\Autodesk\Revit 2022\AddIns\DynamoForRevit\IronPython.StdLib.2.7.9\duynamoLibrary')
from duynamoLibrary import *
clr.AddReference("ProtoGeometry")

from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

clr.AddReference('DSCoreNodes')
from DSCore.List import Flatten

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
def chunkList(inputList, chunkSize = 4):
    """Chop a list into chunks of specified size."""
    return [inputList[i:i+chunkSize] for i in range(0, len(inputList), chunkSize)]
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

def centerOf4Points(points):
    centerPoint = XYZ
    sumX = 0
    sumY = 0
    sumZ = 0
    try:
        if len(points) == 4:
           for pt in points:
               sumX += pt.X 
               sumY += pt.Y 
               sumZ += pt.Z
        centerX = sumX /4
        centerY = sumY /4
        centerZ = sumZ /4
        centerPoint = XYZ(centerX, centerY, centerZ)
    except Exception as e:
        pass
    return centerPoint
#endregion


curveList   = uwList(IN[0])

# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
chunkedList = chunkList(curveList, 4)
startPoints = []
chunked4LineSP = []
for chunked4Line in chunkedList: 
	for c in chunked4Line:
		sp = c.ToRevitType().GetEndPoint(0).ToPoint()
		chunked4LineSP.append(sp)
startPoints = chunkList(chunked4LineSP, 4)
centerPointsLst = []
for recPoints in startPoints:
	cp = centerOf4Points(recPoints)
	centerPointsLst.append(cp)
chunkedCPList = chunkList(centerPointsLst, 1)
TransactionManager.Instance.TransactionTaskDone()

OUT = chunkedList, startPoints, chunkedCPList