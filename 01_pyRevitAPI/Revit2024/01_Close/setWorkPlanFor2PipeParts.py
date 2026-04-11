"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

clr.AddReference("ProtoGeometry")
from math import *
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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
class selectionFilter(ISelectionFilter):
    def __init__(self,categories):
        self.categories = categories
    def AllowElement(self, element):
        eleName = element.Category.Name in self.categories
        return eleName
def pickPipesOrPipeParts():
    categories = ['Pipe Fittings', 'Pipe Accessories', 'Pipes']
    eles = []
    fittingFilter = selectionFilter(categories)
    try:
        while len(eles) < 2:
            ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter)
            ele = doc.GetElement(ref.ElementId)
            eles.append(ele)
        return eles
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        if len(eles) < 2:
            return []
        return eles
    except Exception as e:
        return []
def create_plane_from_three_points(point1, point2, point3):
    vector1 = point2 - point1
    vector2 = point3 - point1
    normal = vector1.CrossProduct(vector2).Normalize()
    plane = Plane.CreateByNormalAndOrigin(normal, point1)  
    return plane    
    
selectedElement = pickPipesOrPipeParts()
conOriginList = []
for ele in selectedElement:
    if ele.Category.Name == 'Pipes':
        pipe_conList = list(ele.ConnectorManager.Connectors)
        for c in pipe_conList:
            # Kiểm tra khoảng cách đến các Origin hiện có
            is_unique = True
            for existing_origin in conOriginList:
                distance = c.Origin.DistanceTo(existing_origin)  # Khoảng cách (feet)
                if distance < 0.00328084:  # 1mm = 0.00328084 feet
                    is_unique = False
                    break
            if is_unique:
                conOriginList.append(c.Origin)

    if ele.Category.Name != 'Pipes':
        notPipe_conList = list(ele.MEPModel.ConnectorManager.Connectors)
        for c in notPipe_conList:
            # Kiểm tra khoảng cách đến các Origin hiện có
            is_unique = True
            for existing_origin in conOriginList:
                distance = c.Origin.DistanceTo(existing_origin)  # Khoảng cách (feet)
                if distance < 0.00328084:  # 1mm = 0.00328084 feet
                    is_unique = False
                    break
            if is_unique:
                conOriginList.append(c.Origin)
TransactionManager.Instance.EnsureInTransaction(doc)                
tempPlan = create_plane_from_three_points(conOriginList[0], conOriginList[1], conOriginList[2])
sketchPlane = SketchPlane.Create(doc, tempPlan)
view.SketchPlane = sketchPlane
TransactionManager.Instance.TransactionTaskDone()
OUT = 'Hello World'