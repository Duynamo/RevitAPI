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
def pickPipeOrPipePart():
    categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
    desEle = doc.GetElement(ref.ElementId)
    return desEle

selectedElement = pickPipeOrPipePart()
elementCheck = selectedElement.Category.Name

if elementCheck == 'Pipes':
    pipe_refLevel_ID = selectedElement.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    pipe_refLevel = doc.GetElement(pipe_refLevel_ID)
    pipe_COPLevel = selectedElement.LookupParameter("Lower End Centerline Elevation").AsDouble()
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        # Lấy tọa độ và thông tin của Level tham chiếu
        pipe_refLevelElevation = pipe_refLevel.Elevation  # Độ cao của Level (feet)
        pipe_newElevation = pipe_refLevelElevation + pipe_COPLevel
        # Tạo một mặt phẳng mới
        # Mặt phẳng nằm ngang (normal = Z-axis), gốc tại (0, 0, newElevation)
        origin = XYZ(0, 0, pipe_newElevation)
        normal = XYZ(0, 0, 1)  # Hướng pháp tuyến (Z-axis)
        plane = Plane.CreateByNormalAndOrigin(normal, origin)
        sketchPlane = SketchPlane.Create(doc, plane)
        view.SketchPlane = sketchPlane
        # Kết thúc giao dịch
        TransactionManager.Instance.TransactionTaskDone()
    except Exception as e:
        # Hủy giao dịch nếu có lỗi
        TransactionManager.Instance.ForceCloseTransaction()
        TaskDialog.Show("Error", "An error occurred: {}".format(str(e)))

if elementCheck != 'Pipes':
    fitting_ParamLevel = selectedElement.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
    fitting_RefLevel = doc.GetElement(fitting_ParamLevel.AsElementId())
    fitting_COPLevel = selectedElement.LookupParameter("Elevation from Level").AsDouble()
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        # Lấy tọa độ và thông tin của Level tham chiếu
        fitting_refLevelElevation = fitting_RefLevel.Elevation  # Độ cao của Level (feet)
        fitting_newElevation = fitting_refLevelElevation + fitting_COPLevel
        # Tạo một mặt phẳng mới
        # Mặt phẳng nằm ngang (normal = Z-axis), gốc tại (0, 0, newElevation)
        origin = XYZ(0, 0, fitting_newElevation)
        normal = XYZ(0, 0, 1)  # Hướng pháp tuyến (Z-axis)
        plane = Plane.CreateByNormalAndOrigin(normal, origin)
        sketchPlane = SketchPlane.Create(doc, plane)
        view.SketchPlane = sketchPlane
        # Kết thúc giao dịch
        TransactionManager.Instance.TransactionTaskDone()
    except Exception as e:
        # Hủy giao dịch nếu có lỗi
        TransactionManager.Instance.ForceCloseTransaction()
        TaskDialog.Show("Error", "An error occurred: {}".format(str(e)))

    pass

# OUT = selectedElement_COPLevel, selectedElement_refLevel
OUT = 'Hello World'