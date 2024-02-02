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

#Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

#Preparing input from dynamo to revit
points = UnwrapElement(IN[0])
circles = []
modelCurves = []
circle_refs = []
TransactionManager.Instance.EnsureInTransaction(doc)
iRefPlane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
sketchPlane = SketchPlane.Create(doc, iRefPlane)
doc.ActiveView.SketchPlane = sketchPlane
red_color = Autodesk.Revit.DB.Color(255, 0, 0)


fill_pattern_elements = list(FilteredElementCollector(doc).OfClass(FillPatternElement))
hatch_pattern_names = [element.Name for element in fill_pattern_elements if element.GetFillPattern().Target == FillPatternTarget.Model]
fill_pattern = fill_pattern_elements[0]

# Create circles for each center point
TransactionManager.Instance.EnsureInTransaction(doc)

for point in points:

    XYZ_point = XYZ(point.X/304.8, point.Y/304.8, point.Z/304.8)
    radius = 1/304.8*50  # You can set your desired radius
    circle = Arc.Create(XYZ_point, radius, 0.0, 2 * 3.142,  XYZ.BasisX, XYZ.BasisY)
    model_curve = doc.Create.NewModelCurve(circle, sketchPlane)
    #filled_regionArray = FilledRegion.Create(doc, fill_pattern.Id, doc.ActiveView.Id, model_curve)
    circles.append(circle.ToProtoType())
    circle_refs.append(model_curve.GeometryCurve.Reference)
    modelCurves.append(model_curve)
# End the transaction
filled_regionArray = FilledRegion.Create(doc, fill_pattern.Id, view.Id, [modelCurves.GeometryCurve.Reference])

TransactionManager.Instance.TransactionTaskDone()

# Output: List of created circles
OUT = fill_pattern
