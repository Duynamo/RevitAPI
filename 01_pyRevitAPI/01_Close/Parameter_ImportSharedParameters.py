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

# Hàm chọn file Excel
def select_excel_file():
    """Hiển thị hộp thoại để chọn file Excel."""
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Chọn file Excel"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

# Hàm đọc dữ liệu từ file Excel
def read_excel_data(file_path, sheet_name):
    """Đọc dữ liệu từ sheet 'Shared Parameter' trong file Excel và trả về danh sách tên parameter."""
    if not os.path.exists(file_path):
        return None, "Lỗi: File '%s' không tồn tại" % file_path
    
    excel = None
    workbook = None
    sheet = None
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path)
        try:
            sheet = workbook.Worksheets[sheet_name]
        except:
            return None, "Sheet '%s' không tồn tại" % sheet_name
        
        # Xác định phạm vi dữ liệu
        used_range = sheet.UsedRange
        last_row = used_range.Rows.Count
        last_col = used_range.Columns.Count
        data = []
        
        # Kiểm tra sheet có dữ liệu không
        if last_row < 1:
            return None, "Sheet '%s' rỗng" % sheet_name
        
        # Đọc tiêu đề từ dòng 1
        headers = []
        for col_idx in range(1, last_col + 1):
            cell = sheet.Cells[1, col_idx]
            header_value = str(cell.Value2) if cell.Value2 is not None else "Column_%s" % col_idx
            headers.append(header_value)
        
        # Kiểm tra có tiêu đề 'Shared Parameters Name' không
        if "Shared Parameters Name" not in headers:
            return None, "Sheet '%s' thiếu tiêu đề 'Shared Parameters Name'" % sheet_name
        
        # Đọc dữ liệu từ dòng 2, lấy giá trị từ cột 'Shared Parameters Name'
        param_names = []
        for row_idx in range(2, last_row + 1):
            cell = sheet.Cells[row_idx, headers.index("Shared Parameters Name") + 1]
            cell_value = str(cell.Value2) if cell.Value2 is not None else None
            if cell_value:
                param_names.append(cell_value)
        
        return param_names, None
    
    except Exception as e:
        return None, "Lỗi đọc file Excel: %s" % str(e)
    
    finally:
        # Đóng và giải phóng COM objects
        if sheet is not None:
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        if workbook is not None:
            workbook.Close(False)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        if excel is not None:
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)

# Hàm lấy tất cả các phần tử thuộc các danh mục Pipe, Pipe Fittings, Pipe Accessories, và Generic Model
def get_elements_in_categories():
    """Lấy tất cả các đối tượng thuộc các danh mục Pipe, Pipe Fittings, Pipe Accessories, và Generic Model trong view hiện tại."""
    all_elements = []
    all_elements_Id = []
    
    # Định nghĩa các danh mục cần lọc
    categories = [
        BuiltInCategory.OST_PipeCurves,      # Danh mục Pipe
        BuiltInCategory.OST_PipeFitting,     # Danh mục Pipe Fittings
        BuiltInCategory.OST_PipeAccessory,   # Danh mục Pipe Accessories
        BuiltInCategory.OST_GenericModel     # Danh mục Generic Model
    ]
    
    # Lấy view hiện tại
    active_view_id = doc.ActiveView.Id
    
    # Lặp qua từng danh mục và lấy các phần tử
    for category in categories:
        elements = FilteredElementCollector(doc, active_view_id)\
                    .OfCategory(category)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
        all_elements.extend(elements)
    
    # Nếu không tìm thấy phần tử, trả về danh sách rỗng
    if not all_elements:
        return [], []
    
    # Lấy ID của các phần tử
    for element in all_elements:
        all_elements_Id.append(element.Id.IntegerValue)
    
    return all_elements, all_elements_Id

# Hàm tạo và gán shared parameters cho các phần tử
def create_and_bind_shared_parameters(param_names):
    """Tạo và gán shared parameters từ danh sách tên parameter vào các danh mục."""
    # Định nghĩa các danh mục để gán parameter
    category_ids = [
        BuiltInCategory.OST_PipeCurves,      # Pipes
        BuiltInCategory.OST_PipeFitting,     # Pipe Fittings
        BuiltInCategory.OST_PipeAccessory,   # Pipe Accessories
        BuiltInCategory.OST_GenericModel     # Generic Model
    ]
    
    # Bắt đầu transaction
    t = Transaction(doc, "Add Shared Parameters")
    t.Start()
    
    try:
        # Mở hoặc tạo file shared parameter
        shared_param_file = app.OpenSharedParameterFile()
        if shared_param_file is None:
            app.SharedParameterFile = app.OpenSharedParameterFile()
            shared_param_file = app.OpenSharedParameterFile()
        
        # Tạo hoặc lấy group trong file shared parameter
        group_name = "ProjectParameters"
        group = shared_param_file.Groups.get_Item(group_name)
        if group is None:
            group = shared_param_file.Groups.Create(group_name)
        
        # Tạo CategorySet để gán parameter
        cat_set = app.Create.NewCategorySet()
        for cat_id in category_ids:
            category = doc.Settings.Categories.get_Item(cat_id)
            if category:
                cat_set.Insert(category)
        
        # Tạo hoặc cập nhật shared parameters
        for param_name in param_names:
            # Kiểm tra xem parameter đã tồn tại chưa
            def_exists = group.Definitions.get_Item(param_name)
            if def_exists is None:
                # Tạo external definition mới với định dạng Text
                external_def_creation_options = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text) # Sửa lỗi: thay ParameterType.Text thành SpecTypeId.String.Text
                external_def_creation_options.Visible = True
                external_def_creation_options.UserModifiable = True
                def_exists = group.Definitions.Create(external_def_creation_options)
            
            # Gán parameter vào các danh mục
            binding = app.Create.NewInstanceBinding(cat_set)
            doc.ParameterBindings.Insert(def_exists, binding, BuiltInParameterGroup.PG_DATA)
        
        t.Commit()
        return True, "Shared parameters added successfully."
    
    except Exception as e:
        t.RollBack()
        return False, "Error: " + str(e)

# Chọn file Excel và đọc dữ liệu từ sheet "Shared Parameter"
excelPath = select_excel_file()
if excelPath:
    sheetName = "Shared Parameter"
    param_names, error_message = read_excel_data(excelPath, sheetName)
    
    if error_message:
        OUT = error_message
    elif param_names:
        # Tạo và gán shared parameters
        success, message = create_and_bind_shared_parameters(param_names)
        OUT = message
    else:
        OUT = "Không tìm thấy shared parameters trong sheet."
else:
    OUT = "Không chọn được file Excel."