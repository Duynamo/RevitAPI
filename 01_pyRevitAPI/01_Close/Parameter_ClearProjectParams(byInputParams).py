"""Copyright by: vudinhduybm@gmail.com"""
# Import tất cả các thư viện cần thiết
import clr
import sys
import System
import math
import collections
import os

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB import Line, XYZ, IntersectionResultArray, SetComparisonResult

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.UI import *

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

clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel
import System.Runtime.InteropServices

# Khởi tạo các biến Revit cơ bản
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

# class ParaProj():
#     def __init__(self, definition, binding):
#         self.definition = definition
#         self.binding = binding
#         self.defName = definition.Name if definition else "None"
#         self.group_label = LabelUtils.GetLabelFor(definition.ParameterGroup) if definition else "Unknown"
#         self.group_id = definition.ParameterGroup if definition else None
#         self.type = "Instance" if isinstance(binding, InstanceBinding) else "Type"
#         self.thesecats = []
        
#         if binding and binding.Categories:
#             for cat in binding.Categories:
#                 try:
#                     self.thesecats.append(cat.Name)
#                 except:
#                     self.thesecats.append("Unknown Category")

#     @property
#     def isShared(self):
#         if not self.definition:
#             return False
#         coll = FilteredElementCollector(doc).OfClass(SharedParameterElement).ToElements()
#         shared_names = {sp.GetDefinition().Name for sp in coll}
#         return self.defName in shared_names


# # === XÓA CẢ NON-SHARED PROJECT PARAMETERS VÀ SHARED PARAMETERS CÓ PREFIX "Params_" ===
# # from Autodesk.Revit.DB import FilteredElementCollector, ParameterElement, Transaction, BuiltInParameterGroup

# t = Transaction(doc, "Xóa Parameters prefix Params_ (Project & Shared)")
# t.Start()

# deleted_count = 0
# ids_to_delete = []

# # Thu thập tất cả ParameterElement trong project
# param_elements = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

# for param_el in param_elements:
#     defi = param_el.GetDefinition()
#     # if defi and defi.ParameterGroup == BuiltInParameterGroup.PG_DATA:
#     if defi.Name.startswith("ISD_"):
#         ids_to_delete.append(param_el.Id)
#         deleted_count += 1

# # Xóa bằng doc.Delete (xóa cả non-shared project parameters và binding của shared)
# if ids_to_delete:
#     for i in ids_to_delete:
#         doc.Delete(i)

# t.Commit()

# # === KẾT QUẢ ===
# if deleted_count > 0:
#     OUT = "ok"
# else:
#     OUT = "Không tìm thấy parameter nào thỏa điều kiện."


class DeleteParamForm(Form):
    def __init__(self):
        # Thiết lập UI đơn giản
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width // 4
        screen_height = primary_screen.Height // 8
        
        self.Text = 'Delete Parameters By Preffix'
        self.ClientSize = Size(screen_width, screen_height)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        
        # Label
        self.label = Label()
        self.label.Text = "Preffix Value(Preffix_)"
        self.label.Location = Point(20, 20)
        self.label.Size = Size(200, 20)
        self.Controls.Add(self.label)
        
        # TextBox
        self.textBox = TextBox()
        self.textBox.Location = Point(20, 50)
        self.textBox.Size = Size(screen_width - 80, 25)
        self.textBox.Text = "Params_"  # Giá trị mặc định
        self.Controls.Add(self.textBox)
        
        # OK Button
        self.okButton = Button()
        self.okButton.Text = 'Delete'
        self.okButton.Location = Point(20, 90)
        self.okButton.Size = Size(100, 35)
        self.okButton.DialogResult = DialogResult.OK
        self.Controls.Add(self.okButton)
        
        # Cancel Button
        self.cancelButton = Button()
        self.cancelButton.Text = 'Cancle'
        self.cancelButton.Location = Point(130, 90)
        self.cancelButton.Size = Size(100, 35)
        self.cancelButton.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.cancelButton)
        
        self.AcceptButton = self.okButton
        self.CancelButton = self.cancelButton

# Hiển thị form lấy prefix
form = DeleteParamForm()
if form.ShowDialog() != DialogResult.OK:
    OUT = "Cancled"
else:
    prefix = form.textBox.Text.strip()
    if not prefix:
        OUT = "Prefix trống, không thực hiện xóa."
    else:
        t = Transaction(doc, "Xóa Parameters prefix")
        t.Start()
        
        deleted_count = 0
        ids_to_delete = []
        
        param_elements = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()
        
        for param_el in param_elements:
            defi = param_el.GetDefinition()
            if defi and defi.Name.startswith(prefix):
                ids_to_delete.append(param_el.Id)
                deleted_count += 1
        
        if ids_to_delete:
            for i in ids_to_delete:
                doc.Delete(i)  # Xóa hàng loạt đúng cách
        
        t.Commit()
        
        if deleted_count > 0:
            OUT = "OK"
        else:
            OUT = "No"