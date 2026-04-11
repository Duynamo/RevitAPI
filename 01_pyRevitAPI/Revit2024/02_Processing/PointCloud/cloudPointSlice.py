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

#######OK NOW YOU CAN CODE########
#TaskDialog.Show("Test","OK")

selection = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
pc = doc.GetElement(selection)

#Function to find farest point from reference point
def max_distance_point(ref_point, points):
    distances = []
    for point in points:
        distances.append(ref_point.DistanceTo(point))
    max_value = max(distances)
    max_index = distances.index(max_value)
    return points[max_index]

#Function to create model line from 2 points
def createLine(doc, startPoint, endPoint):
    t = Transaction(doc, "Test 1")
    t.Start()
    geomLine = Line.CreateBound(startPoint, endPoint)
    normal_vector = startPoint.Subtract(endPoint).CrossProduct(XYZ(1,1,1)).Normalize()
    geomPlane = Plane.CreateByNormalAndOrigin(normal_vector, startPoint)
    sketch = SketchPlane.Create(doc, geomPlane)
    line = doc.Create.NewModelCurve(geomLine, sketch)
    t.Commit()

#Get point collection from pointcloudinstance
def readPC(pointCloudInstance):
    transform = pointCloudInstance.GetTotalTransform()
    boundingBox = pointCloudInstance.get_BoundingBox(None)
    planes = []
    a = 20
    P0 = uidoc.Selection.PickPoint()
    P1 = XYZ (P0.X -0.1, P0.Y -a, P0.Z -200)
    P2 = XYZ (P0.X +0.1, P0.Y +a, P0.Z +200)
    
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisX, P1))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisX.Negate(), P2))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisY, P1))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisY.Negate(), P2))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, P1))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisZ.Negate(), P2))
    
    filter = Autodesk.Revit.DB.PointClouds.PointCloudFilterFactory.CreateMultiPlaneFilter(planes)
    points = pointCloudInstance.GetPoints(filter, 10e-2, 100)

    points = [XYZ(p.X, p.Y, p.Z) for p in points]
    points = [transform.OfPoint(p) for p in points]
    
    pointCloudInstance.SetSelectionFilter(filter)
    pointCloudInstance.FilterAction = Autodesk.Revit.DB.SelectionFilterAction.Highlight

    return points
    
t = Transaction(doc, "Test")
t.Start()
points = readPC(pc)
t.Commit()
startPoint = points[0]
for point in points[1:]:
    endPoint = point
    createLine(doc, startPoint, endPoint)



OUT = points
