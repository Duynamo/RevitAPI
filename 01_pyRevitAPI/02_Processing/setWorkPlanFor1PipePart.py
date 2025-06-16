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
selectedElement_refLevel_ID = selectedElement.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
selectedElement_refLevel = doc.GetElement(selectedElement_refLevel_ID)
selectedElement_COPLevel = selectedElement.LookupParameter("Lower End Centerline Elevation").AsDouble()

TransactionManager.Instance.EnsureInTransaction(doc)

try:
    # Lấy tọa độ và thông tin của Level tham chiếu
    refLevelElevation = selectedElement_refLevel.Elevation  # Độ cao của Level (feet)

    # Tạo mặt phẳng ảo dựa trên Level
    # Mặt phẳng sẽ nằm tại độ cao của Level + Lower End Centerline Elevation
    newElevation = refLevelElevation + selectedElement_COPLevel

    # Tạo một mặt phẳng mới
    # Mặt phẳng nằm ngang (normal = Z-axis), gốc tại (0, 0, newElevation)
    origin = XYZ(0, 0, newElevation)
    normal = XYZ(0, 0, 1)  # Hướng pháp tuyến (Z-axis)
    plane = Plane.CreateByNormalAndOrigin(normal, origin)

    # Tạo SketchPlane từ mặt phẳng ảo
    sketchPlane = SketchPlane.Create(doc, plane)

    # Đặt SketchPlane làm Work Plane cho view hiện tại
    view.SketchPlane = sketchPlane

    # Kết thúc giao dịch
    TransactionManager.Instance.TransactionTaskDone()

    # Thông báo thành công
    # TaskDialog.Show("Success", "Work plane has been set at elevation: {} feet".format(newElevation))

except Exception as e:
    # Hủy giao dịch nếu có lỗi
    TransactionManager.Instance.ForceCloseTransaction()
    TaskDialog.Show("Error", "An error occurred: {}".format(str(e)))

# # Bắt đầu xử lý
# try:
#     # Lấy đối tượng được chọn
#     selectedElement = pickPipeOrPipePart()

#     # Lấy Level tham chiếu từ tham số LevelId
#     selectedElement_refLevel_ID = selectedElement.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
#     if selectedElement_refLevel_ID == ElementId.InvalidElementId:
#         raise Exception("No valid Level associated with the selected element.")
#     selectedElement_refLevel = doc.GetElement(selectedElement_refLevel_ID)
#     refLevelElevation = selectedElement_refLevel.Elevation  # Độ cao của Level (feet)

#     # Xác định khoảng cách offset dựa trên loại đối tượng
#     offsetElevation = 0.0
#     if selectedElement.Category.Name == "Pipes":
#         param = selectedElement.LookupParameter("Lower End Centerline Elevation")
#         if param is None:
#             raise Exception("Parameter 'Lower End Centerline Elevation' not found for Pipe.")
#         offsetElevation = param.AsDouble()
#     else:  # Pipe Fittings hoặc Pipe Accessories
#         param = selectedElement.LookupParameter("Elevation from Level")
#         if param is None:
#             raise Exception("Parameter 'Elevation from Level' not found for Pipe Fitting or Accessory.")
#         offsetElevation = param.AsDouble()

#     # Tính độ cao mới cho mặt phẳng
#     newElevation = refLevelElevation + offsetElevation

#     # Bắt đầu giao dịch để tạo và chỉnh sửa mô hình
#     TransactionManager.Instance.EnsureInTransaction(doc)

#     try:
#         # Tạo một mặt phẳng mới
#         origin = XYZ(0, 0, newElevation)
#         normal = XYZ(0, 0, 1)  # Hướng pháp tuyến (Z-axis)
#         plane = Plane.CreateByNormalAndOrigin(normal, origin)

#         # Tạo SketchPlane từ mặt phẳng ảo
#         sketchPlane = SketchPlane.Create(doc, plane)

#         # Kiểm tra xem view có hỗ trợ Work Plane không
#         if not view.CanBeModified():
#             raise Exception("Current view does not support setting a work plane.")

#         # Đặt SketchPlane làm Work Plane cho view hiện tại
#         view.SketchPlane = sketchPlane

#         # Kết thúc giao dịch
#         TransactionManager.Instance.TransactionTaskDone()

#         # Thông báo thành công
#         # TaskDialog.Show("Success", "Work plane has been set at elevation: {} feet".format(newElevation))

#     except Exception as e:
#         # Hủy giao dịch nếu có lỗi
#         TransactionManager.Instance.ForceCloseTransaction()
#         raise Exception("Error during plane creation: {}".format(str(e)))

# except Exception as e:
#     # Thông báo lỗi
#     TaskDialog.Show("Error", str(e))

OUT = selectedElement_COPLevel, selectedElement_refLevel