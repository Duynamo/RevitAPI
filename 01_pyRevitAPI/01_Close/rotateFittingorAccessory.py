"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

class selectionFilter(ISelectionFilter):
    def __init__(self, categories):
        self.categories = categories  # Danh sách các category (ví dụ: ["Pipe Fittings", "Pipe Accessories"])
    
    def AllowElement(self, element):
        return element.Category and element.Category.Name in self.categories if self.categories else False

def pickPipe():
    pipeFilter = selectionFilter(['Pipes'])
    pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter, "Select a Pipe")
    pipe = doc.GetElement(pipeRef.ElementId)
    return pipe

def pickFittingOrAccessory():
    filter = selectionFilter(['Pipe Fittings', 'Pipe Accessories'])
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, filter, "Select a Pipe Fitting or Pipe Accessory")
    element = doc.GetElement(ref.ElementId)
    return element

def rotate_fitting_or_accessory(pipeA, fitting_or_accessory, angle_degrees):
    """
    Xoay fitting hoặc accessory quanh trục của pipeA một góc angle_degrees.
    
    Args:
        pipeA: Pipe chính
        fitting_or_accessory: Pipe Fitting hoặc Pipe Accessory (FamilyInstance)
        angle_degrees: Góc xoay (độ)
    
    Returns:
        bool: True nếu thành công
    """
    try:
        if not fitting_or_accessory:
            return False
        # Lấy trục xoay từ pipeA
        pipeA_curve = pipeA.Location.Curve
        P1 = pipeA_curve.GetEndPoint(0)
        P2 = pipeA_curve.GetEndPoint(1)
        axis_dir = (P2 - P1).Normalize()
        # Tâm xoay là vị trí fitting_or_accessory
        location = fitting_or_accessory.Location
        if not hasattr(location, "Point"):
            return False
        center = location.Point
        # Tạo trục xoay
        axis_line = Line.CreateUnbound(center, axis_dir)
        # Góc xoay (radian)
        angle_radians = math.radians(angle_degrees)
        # Bắt đầu transaction
        with Transaction(doc, "Rotate Fitting or Accessory") as t:
            t.Start()
            # Xoay fitting_or_accessory
            ElementTransformUtils.RotateElement(
                doc, fitting_or_accessory.Id, axis_line, angle_radians
            )
            t.Commit()
        return True
    except Exception as e:
        return False
class MyForm(Form):
    def __init__(self, pipeA, fittingA):
        # Lưu pipeA và fittingA để sử dụng trong okButton_Click
        self.pipeA = pipeA
        self.fittingA = fittingA
        # Thiết lập UI
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width // 5
        screen_height = primary_screen.Height // 6
        self.Text = ''
        self.ClientSize = Size(screen_width, screen_height)
        self.Font = System.Drawing.Font("Meiryo UI", 7.5, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self.ForeColor = System.Drawing.Color.Red     
        # Cố định vị trí ban đầu (căn giữa màn hình)
        self.StartPosition = FormStartPosition.Manual
        self.initial_location = Point(
            (primary_screen.Width - screen_width) // 3,  # Căn giữa ngang
            (primary_screen.Height - screen_height) // 3  # Căn giữa dọc
        )
        self.Location = self.initial_location
        # Create and add label
        self.label = Label()
        self.label.Text = "ANGLE"
        self.label.Size = Size(int(screen_width * 0.5), int(screen_height * 0.15))
        self.label.Location = Point(int(screen_width * 0.05), int(screen_height * 0.1))
        self.Controls.Add(self.label)
        # Create and add text box
        self.textBox = TextBox()
        self.textBox.Location = Point(int(screen_width * 0.05), int(screen_height * 0.3))
        self.textBox.Size = Size(int(screen_width * 0.9), int(screen_height * 0.15))
        self.textBox.KeyDown += self.textBox_KeyDown
        self.Controls.Add(self.textBox)
        
        # Create and add OK button
        self.okButton = Button()
        self.okButton.Text = 'OK'
        self.okButton.Size = Size(int(screen_width * 0.3), int(screen_height * 0.2))
        self.okButton.Location = Point(int(screen_width * 0.4), int(screen_height * 0.7))
        self.okButton.Click += self.okButton_Click
        self.Controls.Add(self.okButton)
        # Create and add Cancel button
        self.cancelButton = Button()
        self.cancelButton.Text = 'CANCLE'
        self.cancelButton.Size = Size(int(screen_width * 0.3), int(screen_height * 0.2))
        self.cancelButton.Location = Point(int(screen_width * 0.7), int(screen_height * 0.7))
        self.cancelButton.Click += self.cancelButton_Click
        self.Controls.Add(self.cancelButton)
        # Create and add FVC label
        self.fvcLabel = Label()
        self.fvcLabel.Text = "@FVC"
        self.fvcLabel.Size = Size(int(screen_width * 0.3), int(screen_height * 0.15))
        self.fvcLabel.Location = Point(int(screen_width * 0.05), int(screen_height * 0.8))
        self.Controls.Add(self.fvcLabel)        
        self.result = None

    def textBox_KeyDown(self, sender, event):
        if event.KeyCode == Keys.Enter:
            self.okButton_Click(sender, None)
            event.Handled = True        

    def okButton_Click(self, sender, event):
        try:
            angle = float(self.textBox.Text)
            # Thực thi quay fitting
            success = rotate_fitting_or_accessory(self.pipeA, self.fittingA, angle)
            # Reset TextBox
            self.textBox.Text = ""
            # Không đóng form, giữ UI mở
            self.DialogResult = DialogResult.OK
        except ValueError:
            # print("Error: Please enter a valid number for Angle.")
            self.textBox.Text = ""  # Reset TextBox ngay cả khi lỗi
            self.DialogResult = DialogResult.OK

    def cancelButton_Click(self, sender, event):
        self.DialogResult = DialogResult.Cancel
        self.Close()  # Đóng form khi nhấn Cancel

# Chọn pipeA và fittingA
try:
    pipeA = pickPipe()
    fittingA = pickFittingOrAccessory()
    
    # Mở form và giữ UI mở cho đến khi nhấn Cancel
    form = MyForm(pipeA, fittingA)
    while form.ShowDialog() == DialogResult.OK:
        pass  # Lặp lại để giữ form mở cho đến khi nhấn Cancel
except Exception as e: pass
    # print(f"Error: {str(e)}")

OUT = ''