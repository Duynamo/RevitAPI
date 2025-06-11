import clr
import sys 
import System   
import math

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB.Plumbing import*
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

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

def select_excel_file():
    """Hiển thị hộp thoại để chọn file Excel."""
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Chọn file Excel"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

def read_excel_data(file_path, sheet_name="Sheet1"):
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
        headers = ["Element_ID", "FVC_Manufactured Day", "FVC_Delivery Day", 
                   "FVC_Installation Day", "FVC_Installer", "FVC_Cost", "FVC_Other"]
        data = []
        
        # Kiểm tra sheet có dữ liệu không
        if last_row < 2:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            return None, "Sheet '%s' rỗng hoặc chỉ có tiêu đề" % sheet_name
        
        # Đọc dữ liệu từ dòng 2 (bỏ tiêu đề)
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
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    for row in data:
        if "Error" in row:
            results.append(row["Error"])
            continue
        
        element_id = row.get("Element_ID")
        element = get_element_by_id(element_id)
        if not element:
            pass
            # results.append("Element_ID %s: Không tìm thấy phần tử" % element_id)
            # continue
        
        row_result = "Element_ID %s: " % element_id
        param_results = []
        for param_name in headers[1:]:  # Bỏ Element_ID
            TransactionManager.Instance.EnsureInTransaction(doc)
            param_value = row.get(param_name)
            success, message = update_parameter(element, param_name, param_value)
            TransactionManager.Instance.TransactionTaskDone()
            param_results.append(message)
        
        if param_results:
            row_result += "; ".join(param_results)
        else:
            row_result += "Không có tham số nào được cập nhật"
        results.append(row_result)
    
    TransactionManager.Instance.TransactionTaskDone()
    return results

# Chọn file Excel
file_path = select_excel_file()
if not file_path:
    pass
else:
    headers, data = read_excel_data(file_path)
    if headers is None:
        pass
    else:
        TransactionManager.Instance.EnsureInTransaction(doc)
        results = process_excel_data(headers, data)
        TransactionManager.Instance.TransactionTaskDone()
        OUT = results

OUT = read_excel_data(file_path)