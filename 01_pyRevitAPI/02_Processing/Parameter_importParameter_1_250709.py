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

"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""
#region _function
def select_excel_file():
    """Hiển thị hộp thoại để chọn file Excel."""
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Chọn file Excel"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

def read_excel_data(file_path, sheet_name):
    """Đọc dữ liệu từ file Excel bằng Microsoft.Office.Interop.Excel."""
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path)
        try:
            sheet = workbook.Worksheets[sheet_name]
        except:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            return None, "Sheet '%s' không tồn tại" % sheet_name
        
        # Xác định phạm vi dữ liệu
        used_range = sheet.UsedRange
        last_row = used_range.Rows.Count
        last_col = used_range.Columns.Count
        data = []
        
        # Kiểm tra sheet có dữ liệu không
        if last_row < 1:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            return None, "Sheet '%s' rỗng" % sheet_name
        
        # Đọc tiêu đề từ dòng 1
        headers = []
        for col_idx in range(1, last_col + 1):
            cell = sheet.Cells[1, col_idx]
            header_value = str(cell.Value2) if cell.Value2 is not None else "Column_%s" % col_idx
            headers.append(header_value)
        
        # Kiểm tra có tiêu đề Element_ID không
        if "Element_ID" not in headers:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            return None, "Sheet '%s' thiếu tiêu đề 'Element_ID'" % sheet_name
        
        # Đọc dữ liệu từ dòng 2
        for row_idx in range(2, last_row + 1):
            row_data = {}
            for idx, header in enumerate(headers):
                cell = sheet.Cells[row_idx, idx + 1]
                cell_value = cell.Value2 if cell.Value2 is not None else None
                row_data[header] = cell_value
            if row_data.get("Element_ID"):
                data.append(row_data)
            else:
                data.append({"Error": "Dòng %s: Thiếu hoặc rỗng Element_ID" % row_idx})
        
        # Đóng Excel
        workbook.Close(False)
        excel.Quit()
        # Giải phóng COM objects
        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        return headers, data
    
    except Exception as e:
        try:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        except:
            pass
        return None, "Lỗi đọc file Excel: %s" % str(e)
    
def get_element_by_id(element_id):
    """Tìm phần tử Revit theo Element ID."""
    try:
        # Xử lý element_id có thể là số thực hoặc chuỗi
        if isinstance(element_id, float):
            element_id = int(element_id)
        elif isinstance(element_id, str):
            element_id = int(float(element_id.strip()))
        return doc.GetElement(ElementId(element_id))
    except:
        return None

def update_parameter(element, param_name, param_value):
    """Cập nhật tham số kiểu text của phần tử Revit (instance hoặc type)."""
    try:
        # Tìm instance parameter
        param = element.LookupParameter(param_name)
        if not param:
            # Tìm type parameter nếu instance không tồn tại
            element_type = doc.GetElement(element.GetTypeId())
            if element_type:
                param = element_type.LookupParameter(param_name)
        if not param:
            return False, "Tham số '%s' không tồn tại" % param_name
        if param.IsReadOnly:
            return False, "Tham số '%s' là chỉ đọc" % param_name
        TransactionManager.Instance.EnsureInTransaction(doc)
        param.Set(str(param_value) if param_value is not None else "")
        TransactionManager.Instance.TransactionTaskDone()
        return True, "Đã cập nhật tham số '%s' thành '%s'" % (param_name, param_value if param_value is not None else "")
    except Exception as e:
        return False, "Lỗi khi cập nhật tham số '%s': %s" % (param_name, str(e))

def process_excel_data(headers, data):
    """Xử lý dữ liệu Excel và cập nhật mô hình Revit."""
    results = []
    
    # Tìm hàng chứa giá trị "その他すべて" (hàng 26)
    apply_row = next((row for row in data if row.get("Element_ID") == "その他すべて"), None)
    if apply_row is None:
        results.append(None)
        return results
    
    # Lấy danh sách các Element_ID cần loại trừ (từ hàng 2 đến 24)
    exclude_ids_str = [row["Element_ID"] for row in data[:23] if row.get("Element_ID") and row.get("Element_ID")]
    exclude_ids_int = [int(id_str) for id_str in exclude_ids_str]
    
    # Lấy tất cả các phần tử trong mô hình Revit (không phải type)
    all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    all_ids_int = [ele.Id.IntegerValue for ele in all_elements]
    
    # Lọc các phần tử không có ID thuộc danh sách loại trừ
    desElesId_int = [id_int for id_int in all_ids_int if id_int not in exclude_ids_int]
    desEles = [doc.GetElement(ElementId(id_int)) for id_int in desElesId_int]
    
    if not desEles:
        results.append(None)
        return results
    
    # Bắt đầu giao dịch
    try:
        TransactionManager.Instance.EnsureInTransaction(doc)
        
        updated_elements = 0  # Đếm số phần tử được cập nhật ít nhất một tham số
        for element in desEles:
            element_updated = False  # Theo dõi xem phần tử có được cập nhật không
            for param_name in headers[1:]:  # Bỏ qua cột Element_ID
                param_value = apply_row.get(param_name)
                if param_value is not None:
                    param = element.LookupParameter(param_name)
                    if param and param.StorageType == StorageType.String and not param.IsReadOnly:
                        try:
                            param.Set(str(param_value))  # Chuyển đổi thành chuỗi cho kiểu Text
                            element_updated = True
                        except:
                            pass  # Bỏ qua lỗi cho tham số cụ thể
                    else:
                        pass  # Tham số không tồn tại, không phải kiểu Text, hoặc chỉ đọc
                else:
                    pass  # Giá trị None, bỏ qua
            if element_updated:
                updated_elements += 1
        
        # Kết thúc giao dịch
        TransactionManager.Instance.TransactionTaskDone()
        
        # Append kết quả
        if updated_elements > 0:
            results.append(updated_elements)
        else:
            results.append(None)
    
    except:
        TransactionManager.Instance.TransactionTaskDone()  # Đảm bảo kết thúc giao dịch nếu có lỗi
        results.append(None)
    
    return results, desElesId_int, desEles
def get_element_by_id(element_id):
    """Tìm phần tử Revit theo Element ID."""
    try:
        # Xử lý element_id có thể là số thực hoặc chuỗi
        if isinstance(element_id, float):
            element_id = int(element_id)
        elif isinstance(element_id, str):
            element_id = int(float(element_id.strip()))
        return doc.GetElement(ElementId(element_id))
    except:
        return None
def get_elements_in_active_view():
    """Lấy tất cả các đối tượng thuộc các danh mục Pipe, Pipe Fittings, Pipe Accessories, và Generic Model trong view hiện tại."""
    all_elements_Id = []  # Danh sách để lưu kết quả (ID của các phần tử)

    # Lấy view hiện tại từ tài liệu Revit
    active_view_id = doc.ActiveView.Id

    # Định nghĩa các danh mục cần lọc
    categories = [
        BuiltInCategory.OST_PipeCurves,      # Danh mục Pipe
        BuiltInCategory.OST_PipeFitting,     # Danh mục Pipe Fittings
        BuiltInCategory.OST_PipeAccessory,   # Danh mục Pipe Accessories
        BuiltInCategory.OST_GenericModel     # Danh mục Generic Model
    ]

    # Tạo danh sách để lưu tất cả phần tử từ các danh mục
    all_elements = []
    
    # Lặp qua từng danh mục và lấy các phần tử trong view hiện tại
    for category in categories:
        # Sử dụng FilteredElementCollector với view hiện tại và danh mục
        # WhereElementIsNotElementType để lấy instance (không phải type)
        elements = FilteredElementCollector(doc, active_view_id)\
                    .OfCategory(category)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
        # Gộp các phần tử vào danh sách chung
        all_elements.extend(elements)
    
    # Nếu không tìm thấy phần tử nào, trả về None
    if not all_elements:
        all_elements_Id.append(None)
        return all_elements_Id
    
    # Lấy ID của các phần tử và thêm vào kết quả
    for element in all_elements:
        all_elements_Id.append(element.Id.IntegerValue)
    
    return all_elements, all_elements_Id
#endregion
activeViewId = doc.ActiveView.Id
excelPath = select_excel_file()
sheetName = "250708"
excelData = read_excel_data(excelPath, sheetName)
headers = excelData[0]
data = excelData[1]
results = []

# Tìm hàng chứa giá trị "その他すべて" (hàng 26)
apply_row = next((row for row in data if row.get("Element_ID") == "その他すべて"), None)
if apply_row is None:
    results.append(None)
    

# Lấy danh sách các Element_ID cần loại trừ (từ hàng 2 đến 24)
exclude_ids_str = [row["Element_ID"] for row in data[:23] if row.get("Element_ID") and row.get("Element_ID")]
exclude_ids_int = [int(id_str) for id_str in exclude_ids_str]


all_elements_check = get_elements_in_active_view()
all_elements = all_elements_check[0]
all_ids_int = all_elements_check[1]

# Lọc các phần tử không có ID thuộc danh sách loại trừ
desElesId_int = [id_int for id_int in all_ids_int if id_int not in exclude_ids_int]
desEles = [doc.GetElement(ElementId(id_int)) for id_int in desElesId_int]
# Bắt đầu giao dịch
if not desEles:
    results.append(None)
    OUT = results
else:
    try:
        TransactionManager.Instance.EnsureInTransaction(doc)
        updated_elements = 0
        
        for element in desEles:
            element_updated = False
            for param_name in headers[1:]:
                param_value = apply_row.get(param_name)
                if param_value is not None:
                    param = element.LookupParameter(param_name)
                    if param and param.StorageType == StorageType.String and not param.IsReadOnly:
                        try:
                            TransactionManager.Instance.EnsureInTransaction(doc)
                            param.Set(str(param_value))
                            element_updated = True
                            TransactionManager.Instance.TransactionTaskDone()
                        except:
                            pass
            if element_updated:
                updated_elements += 1
        
        TransactionManager.Instance.TransactionTaskDone()
        results.append(updated_elements)
        
    except:
        TransactionManager.Instance.TransactionTaskDone()
        results.append(None)

    
    # Kết thúc giao dịch
    TransactionManager.Instance.TransactionTaskDone()


OUT = results