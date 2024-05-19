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
from Autodesk.Revit.UI.Selection import ObjectType
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

# def PickPoints(doc):
# 	TransactionManager.Instance.EnsureInTransaction(doc)
# 	activeView = doc.ActiveView
# 	iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
# 	sketchPlane = SketchPlane.Create(doc, iRefPlane)
# 	doc.ActiveView.SketchPlane = sketchPlane
# 	condition = True
# 	pointsList = []
# 	dynPList = []
# 	n = 0
# 	msg = "Pick Points on Current Section plane, hit ESC when finished."
# 	TaskDialog.Show("^------^", msg)
# 	while condition:
# 		try:
# 			pt=uidoc.Selection.PickPoint()
# 			n += 1
# 			pointsList.append(pt)				
# 		except :
# 			condition = False
# 	doc.Delete(sketchPlane.Id)	
# 	for j in pointsList:
# 		dynP = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X*304.8, j.Y*304.8, j.Z*304.8)	
# 		dynPList.append(dynP)
# 	TransactionManager.Instance.TransactionTaskDone()			
# 	return dynPList, pointsList
# pickedPoints = PickPoints(doc)

# def pickPipePoint():
#     try:
#         TransactionManager.Instance.EnsureInTransaction(doc)
#         ref = ReferenceWithContext()
#         ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick a pipe")
#         pipe = doc.GetElement(ref)
#         pipeCurve = pipe.Location.Curve
#         options = Options()
#         options.ComputeReferences = True
#         options.IncludeNonVisibleObjects = True
#         options.View = doc.ActiveView
#         intersectPoint = pipeCurve.ProjectTo(options, uidoc.Selection.PickPoint()).XYZPoint
#         TransactionManager.Instance.TransactionTaskDone()
#         return intersectPoint
#     except Exception as e:
#         return None

# pickedPoint = pickPipePoint()

# def pickObjects():
#     pipes = []
#     refs = uidoc.Selection.PickObjects(ObjectType.Element)
#     for r in refs:
#         pipe = doc.GetElement(r.ElementId)
#         lC = pipe.Location.Curve
#         options = Options()
#         options.ComputeReferences = True
#         options.IncludeNonVisibleObjects = True
#         options.View = doc.ActiveView
#         intersectPoint = lC.ProjectTo(options, uidoc.Selection.PickPoint()).XYZPoint       
#         pipes.append(pipe)
#     return  pipes, intersectPoint
# picked = pickObjects()

def pickFaceSetWorkPlaneAndPickPoints(doc):
    TransactionManager.Instance.EnsureInTransaction(doc)
    ref = uidoc.Selection.PickObject(ObjectType.Face, "Vu Dinh Duy")
    element = doc.GetElement(ref.ElementId)
    pointList =[]
    dynList = []
    if element is not None:
        face = element.GetGeometryObjectFromReference(ref)
        if face is not None:
            if isinstance(face, CylindricalFace):
                axis = face.Axis  
                normal = XYZ.BasisZ.CrossProduct(axis.Direction)
                origin = face.Origin
                #newPlane = Plane(normal, origin)
            if isinstance(face, PlanarFace):
                normal = face.FaceNormal
                origin = face.Origin
                #newPlane = Plane(normal, origin)                  
                TransactionManager.Instance.EnsureInTransaction(doc)
                newPlane = Plane(normal, origin) 
                skPlane = SketchPlane.Create(doc, newPlane)
                uidoc.ActiveView.SketchPlane = skPlane
                uidoc.ActiveView.ShowActiveWorkPlane()   
                condition = True
                while condition:
                        try:
                                p3D = uidoc.Selection.PickPoint()
                                pointList.append(p3D)
                        except:
                                condition = False
                doc.Delete(skPlane.Id)
                for p in pointList:
                    dynP = Autodesk.DesignScript.Geometry.Point.ByCoordinates(p.X*304.8, p.Y*304.8, p.Z*304.8)
                    dynList.append(dynP)
                TransactionManager.Instance.TransactionTaskDone()
                return pointList, dynList               
    TransactionManager.Instance.TransactionTaskDone()
picked = pickFaceSetWorkPlaneAndPickPoints(doc)

OUT = picked