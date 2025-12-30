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

class DeleteParamForm(Form):
    def __init__(self, param_names):
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width // 3
        screen_height = primary_screen.Height // 2
        
        self.Text = 'Parameters'
        self.ClientSize = Size(screen_width, screen_height)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        
        # Label
        self.label = Label()
        self.label.Text = "Select Parameters to Delete"
        self.label.Location = Point(20, 20)
        self.label.AutoSize = True
        self.Controls.Add(self.label)
        
        # CheckedListBox
        self.checkedListBox = CheckedListBox()
        self.checkedListBox.Location = Point(20, 50)
        self.checkedListBox.Size = Size(screen_width - 60, screen_height - 150)
        self.checkedListBox.CheckOnClick = True
        for name in sorted(param_names):
            self.checkedListBox.Items.Add(name, False)
        self.Controls.Add(self.checkedListBox)
        
        # Run Button
        self.runButton = Button()
        self.runButton.Text = 'Delete'
        self.runButton.Location = Point(20, screen_height - 80)
        self.runButton.Size = Size(120, 40)
        self.runButton.DialogResult = DialogResult.OK
        self.Controls.Add(self.runButton)
        
        # Cancel Button
        self.cancelButton = Button()
        self.cancelButton.Text = 'Cancle'
        self.cancelButton.Location = Point(150, screen_height - 80)
        self.cancelButton.Size = Size(120, 40)
        self.cancelButton.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.cancelButton)
        
        self.AcceptButton = self.runButton
        self.CancelButton = self.cancelButton

# Chỉ lấy các project parameter do người dùng tạo (loại bỏ built-in)
param_elements = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

user_param_names = []
param_dict = {}  # Lưu mapping name → ParameterElement để xóa sau

for pe in param_elements:
    defi = pe.GetDefinition()
    if defi and defi.BuiltInParameter == BuiltInParameter.INVALID:
        name = defi.Name
        user_param_names.append(name)
        param_dict[name] = pe  # Lưu để lấy Id sau

if not user_param_names:
    OUT = "Không có project parameter nào do người dùng tạo."
else:
    form = DeleteParamForm(sorted(user_param_names))
    if form.ShowDialog() != DialogResult.OK:
        OUT = "Cancled"
    else:
        selected_names = [form.checkedListBox.Items[i] for i in form.checkedListBox.CheckedIndices]
        if not selected_names:
            OUT = "Selected None"
        else:
            t = Transaction(doc, "Xóa các parameter đã chọn")
            t.Start()
            
            ids_to_delete = [param_dict[name].Id for name in selected_names if name in param_dict]
            
            if ids_to_delete:
                for i in ids_to_delete:
                    doc.Delete(i)  # Xóa hàng loạt hiệu quả hơn
            
            t.Commit()
            OUT = "Done"