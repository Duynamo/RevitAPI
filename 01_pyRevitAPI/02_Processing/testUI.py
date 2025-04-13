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
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB import Line, XYZ, IntersectionResultArray, SetComparisonResult

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI.Selection import ISelectionFilter
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

# def create_ui(uiapp):
#     """
#     Tạo UI hiển thị ở màn hình làm việc của Revit với kích thước 1/5 màn hình.
    
#     Args:
#         uiapp: UIApplication của Revit
#     """
#     # Tạo form
#     form = Form()
#     form.Text = "Revit UI Tool"
#     form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
#     form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
#     form.MaximizeBox = False
#     form.MinimizeBox = False
    
#     # Lấy kích thước màn hình làm việc
#     primary_screen = Screen.PrimaryScreen.WorkingArea
#     screen_width = primary_screen.Width
#     screen_height = primary_screen.Height
    
#     # Đặt kích thước form bằng 1/5 màn hình
#     form_width = screen_width // 5
#     form_height = screen_height // 5
#     form.Size = Size(form_width, form_height)
    
#     # Đặt vị trí ở trung tâm màn hình (gần viewport 3D)
#     form.StartPosition = FormStartPosition.CenterScreen
    
#     # Thêm nút đóng (ví dụ)
#     close_button = Button()
#     close_button.Text = "Close"
#     close_button.Size = Size(80, 30)
#     close_button.Location = Point((form_width - 90) // 2, form_height - 70)
#     close_button.Click += lambda sender, args: form.Close()
#     form.Controls.Add(close_button)
    
#     # Hiển thị form (modal)
#     form.ShowDialog()



def create_ui(uiapp):
    """
    Tạo UI cho Revit với TextBox Angle, nút Run/Cancel, logo @FVC, kích thước điều chỉnh.
    
    Args:
        uiapp: UIApplication của Revit
    """
    # Tạo form
    form = Form()
    form.Text = "Revit Angle Tool"
    form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
    form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    form.MaximizeBox = False
    form.MinimizeBox = False
    
    # Lấy kích thước màn hình làm việc
    primary_screen = Screen.PrimaryScreen.WorkingArea
    screen_width = primary_screen.Width
    screen_height = primary_screen.Height
    
    # Đặt kích thước form ~1/5 màn hình, tối thiểu 400x250 để cân đối
    form_width = max(screen_width // 5, 400)
    # form_width = screen_width // 5
    form_height = max(screen_height // 5, 250)
    # form_height = screen_height // 5
    form.ClientSize = Size(form_width, form_height)
    
    # Đặt vị trí ở trung tâm màn hình (3D Fix)
    form.StartPosition = FormStartPosition.CenterScreen
    
    # Font chính (Meiryo, size 10)
    try:
        main_font = Font("Meiryo", 8)
    except:
        main_font = Font("Arial", 8)
    
    # Font logo (Meiryo, size 8)
    try:
        logo_font = Font("Meiryo", 6.5)
    except:
        logo_font = Font("Arial", 6.5)
    
    # Label cho TextBox Angle
    angle_label = Label()
    angle_label.Text = "Angle:"
    angle_label.Font = main_font
    angle_label.Size = Size(form_width // 2, 40)  # Rộng 50% form
    angle_label.Location = Point(20, 20)  # Margin trái 20
    form.Controls.Add(angle_label)
    
    # TextBox Angle
    angle_textbox = TextBox()
    angle_textbox.Name = "Angle"
    angle_textbox.Font = main_font
    angle_textbox.Size = Size(form_width // 1.1, 200)  # Rộng 50% form, cao hơn
    angle_textbox.Location = Point(30, 70)  # Dưới label, margin trái 20
    form.Controls.Add(angle_textbox)
    
    # Button Run
    run_button = Button()
    run_button.Text = "Run"
    run_button.Font = main_font
    run_button.Size = Size(150, 40)  # Lớn hơn, dễ nhấn
    run_button.Location = Point(form_width - 350, form_height - 70)  # Góc dưới phải
    # def run_click(sender, args):
    #     angle = angle_textbox.Text
    #     try:
    #         angle_value = float(angle)
    #         print("Running with Angle: {} degrees".format(angle_value))
    #         # TODO: Thêm mã của bạn tại đây
    #         form.Close()
    #     except ValueError:
    #         print("Error: Please enter a valid number for Angle.")
    run_button.Click += run_click
    form.Controls.Add(run_button)
    
    # Button Cancel
    cancel_button = Button()
    cancel_button.Text = "Cancel"
    cancel_button.Font = main_font
    cancel_button.Size = Size(150, 40)  # Lớn hơn
    cancel_button.Location = Point(form_width - 180, form_height - 70)  # Cạnh Run
    cancel_button.Click += lambda sender, args: form.Close()
    form.Controls.Add(cancel_button)
    
    # Logo @FVC
    logo_label = Label()
    logo_label.Text = "@FVC"
    logo_label.Font = System.Drawing.Font("Arial", 6.75, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
    logo_label.ForeColor = System.Drawing.Color.Red
    logo_label.Size = Size(80, 25)  # Rộng hơn để rõ text
    logo_label.Location = Point(20, form_height - 40)  # Góc dưới trái
    logo_label.ForeColor = Color.DarkBlue
    form.Controls.Add(logo_label)
    
    # Hiển thị form (modal)
    result = form.ShowDialog()
    if result == DialogResult.Ok:
        text_input = form.result
    else:
        text_input = None
    return text_input
    


doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
# Ví dụ sử dụng
UI  =  create_ui(uiapp)