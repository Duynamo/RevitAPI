"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
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

clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel
import System.Runtime.InteropServices
import System.Runtime.InteropServices as Marshal

"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""

try:
    clr.AddReference("IronPython.wpf")
    clr.AddReference('PresentationCore')
    clr.AddReference('PresentationFramework')
except IOError:
    raise
from System.IO import StringReader
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application, MessageBox, MessageBoxButton, MessageBoxResult

try:
    import wpf
except ImportError:
    raise

"""___"""
def select_excel_file():
    """Hiển thị hộp thoại để chọn file Excel."""
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Chọn file Excel"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

# Hàm xuất danh sách ID vào file Excel
def export_to_excel(element_ids, file_path):
    """Xuất danh sách ID vào worksheet 'ElementID' của file Excel với cột 'No' và 'ElementID'."""
    excel = None
    workbook = None
    sheet = None
    try:
        # Khởi tạo ứng dụng Excel
        excel = Excel.ApplicationClass()
        excel.Visible = False
        
        # Mở file Excel nếu tồn tại, nếu không thì tạo mới
        if os.path.exists(file_path):
            workbook = excel.Workbooks.Open(file_path)
        else:
            workbook = excel.Workbooks.Add()
        
        # Tìm hoặc tạo worksheet 'ElementID'
        sheet_name = "ElementID"
        sheet = None
        for ws in workbook.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                # Xóa dữ liệu cũ trong worksheet
                sheet.UsedRange.ClearContents()
                break
        if sheet is None:
            sheet = workbook.Worksheets.Add()
            sheet.Name = sheet_name
        
        # Ghi tiêu đề
        sheet.Cells[1, 1].Value2 = "No"
        sheet.Cells[1, 2].Value2 = "ElementID"
        
        # Ghi dữ liệu: số thứ tự (No) và ElementID
        for idx, element_id in enumerate(element_ids, start=1):
            sheet.Cells[idx + 1, 1].Value2 = idx  # Cột 'No'
            sheet.Cells[idx + 1, 2].Value2 = element_id.IntegerValue  # Cột 'ElementID'
        
        # Lưu file Excel
        workbook.SaveAs(file_path)
        return None
    
    except Exception as e:
        pass
    
    finally:
        # Đóng và giải phóng COM objects
        if sheet is not None:
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        if workbook is not None:
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        if excel is not None:
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)

# Lớp lọc lựa chọn các danh mục
class SelectionFilter(ISelectionFilter):
    """Lớp lọc để chọn các đối tượng thuộc danh mục Pipes, Pipe Fittings, Pipe Accessories, Generic Models."""
    def __init__(self, category1, category2, category3, category4):
        self.category1 = category1
        self.category2 = category2
        self.category3 = category3
        self.category4 = category4
    
    def AllowElement(self, element):
        """Kiểm tra xem phần tử có thuộc các danh mục được chỉ định không."""
        if element.Category and element.Category.Name in [self.category1, self.category2, self.category3, self.category4]:
            return True
        return False
    
    def AllowReference(self, reference, point):
        """Không cho phép chọn tham chiếu."""
        return False

# Hàm chọn các phần tử trong Revit
def pick_objects():
    """Cho phép người dùng chọn các phần tử thuộc danh mục Pipes, Pipe Fittings, Pipe Accessories, Generic Models."""
    try:
        from Autodesk.Revit.UI.Selection import ObjectType
        filter = SelectionFilter("Pipes", "Pipe Fittings", "Pipe Accessories", "Generic Models")
        ele_refs = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Chọn các đối tượng")
        elements = [doc.GetElement(ref.ElementId) for ref in ele_refs]
        return elements
    except Exception as e:
        return []
"""___"""

class MyWindow(Window):
    def __init__(self):
        self.ui = wpf.LoadComponent(self, r'C:\Users\Laptop\OneDrive\07_Python Dynamo\01_Duuynamo\Duuynamo.extension\vitaminD.tab\Other.panel\edit3.stack\Test260218.pushbutton\Form1.xaml')

    def btnSelectExcel (self,sender, event):
        global excelFile
        excelFile = select_excel_file()  # hàm chọn file của bạn, trả về full path string

        if excelFile:  # Chỉ gán nếu người dùng chọn file (không cancel)
                # Lấy TextBlock từ UI
                txt_excel_path = self.ui.FindName("ExcelPath")
                
                if txt_excel_path:
                    txt_excel_path.Text = excelFile  # Gán full path vào Text
                    # Hoặc nếu muốn rút gọn đường dẫn (chỉ tên file + extension)
                    # import os
                    # txt_excel_path.Text = os.path.basename(excelFile)
                else:
                    print("Không tìm thấy TextBlock với tên 'ExcelPath'")
            
            # self.Close()  # Đóng window nếu cần

    def btnDelete (self,sender, event):
        pass
    def btnUp (self,sender, event):
        pass
    def btnDown (self,sender, event):
        pass
    def btnClear (self,sender, event):
        pass
    def btnPickElements (self,sender, event):
        pass
    def btnCancle (self,sender, event):
        self.Close()
                  
    


myWindow = MyWindow()
myWindow.ShowDialog()

OUT = excelFile