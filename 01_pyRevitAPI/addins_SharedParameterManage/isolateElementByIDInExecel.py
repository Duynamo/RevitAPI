"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
# -*- coding: utf-8 -*-
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
try:
    from Autodesk.Revit.DB.Structure import *
except:
    pass

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
import System.Windows.Forms
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
#endregion

"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""

def alert(message, title="Notification"):
    System.Windows.Forms.MessageBox.Show(message, title, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

def get_hidden_excel(file_path):
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise Exception("Excel is not installed or COM registry is missing.")
        
    excel = System.Activator.CreateInstance(excel_type)
    excel.Visible = False
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    
    try:
        workbook = excel.Workbooks.Open(file_path)
        return excel, workbook
    except Exception as e:
        try:
            excel.Quit()
        except:
            pass
        raise e

def select_excel_file():
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Select Excel file with ElementIDs"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

def extract_element_ids_from_excel(file_path):
    if not os.path.exists(file_path):
        return None, "Error: File '%s' does not exist." % file_path
    
    excel = None
    workbook = None
    try:
        excel, workbook = get_hidden_excel(file_path)
        
        # Mở đúng Sheet tên 'ElementID'
        sheet = None
        for i in range(1, workbook.Worksheets.Count + 1):
            name = str(workbook.Worksheets(i).Name)
            if name.strip().lower() == "elementid":
                sheet = workbook.Worksheets(i)
                break
                
        # Nếu không có sheet tên ElementID, fallback sang sheet đầu tiên
        if sheet is None:
            if workbook.Worksheets.Count > 0:
                sheet = workbook.Worksheets(1)
            else:
                return None, "Excel file is completely empty."
        
        # Ưu tiên cứng cột B (Index = 2) theo yêu cầu User
        target_col_index = 2
        
        # Double check xem nếu cột B trống tiêu đề thì dò tìm lại
        header_val = sheet.Cells(1, 2).Value2
        if header_val is None or "elementid" not in str(header_val).strip().lower():
            for i in range(1, 10):
                v = sheet.Cells(1, i).Value2
                if v is not None and "elementid" in str(v).strip().lower():
                    target_col_index = i
                    break
        
        # Cào đúng dòng cuối cùng có chứa data thực thụ (End xlUp)
        try:
            last_row = sheet.Cells(sheet.Rows.Count, target_col_index).End(-4162).Row
        except:
            last_row = sheet.UsedRange.Rows.Count
        
        if last_row < 2:
            return None, "Sheet '%s' does not have any data rows." % sheet.Name
            
        ids = []
        # Chạy từ dòng 2 xuống dòng chứa data cuối cùng
        for r in range(2, last_row + 1):
            val = sheet.Cells(r, target_col_index).Value2
            if val is not None and str(val).strip() != "":
                ids.append(str(val).strip())
                
        if not ids:
            return None, "No valid IDs were found in column B (index %s)." % target_col_index
            
        return ids, None
        
    except Exception as e:
        return None, "Error reading Excel file: " + str(e)
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except:
                pass

def get_element_id_object(id_str):
    """Safe Parsing for ElementID in Revit 2024 (Int64) and below (Int32)"""
    try:
        if isinstance(id_str, float):
            id_val = int(id_str)
        elif isinstance(id_str, str):
            id_val = int(float(id_str.strip()))
        else:
            id_val = int(id_str)
            
        # Try primary creation (Revit 2024 uses Int64, 2023 uses Int32)
        try:
            return ElementId(id_val)
        except:
            import System
            return ElementId(System.Int64(id_val))
    except:
        return None

def isolate_elements(id_list):
    if not view:
        return False, "There is no active view available."

    if not view.CanEnableTemporaryViewPropertiesMode():
        return False, "This active View Type does not support Temporary Isolate/Hide."
        
    valid_ids = List[ElementId]()
    for id_str in id_list:
        eid = get_element_id_object(id_str)
        if eid is not None:
            el = doc.GetElement(eid)
            # Chỉ lấy các Element thực sự còn sống trên model
            if el is not None:
                valid_ids.Add(eid)
                
    if valid_ids.Count == 0:
        return False, "None of the Element IDs extracted from Excel actually exist in this project."
        
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        # Nếu đang có cách Isolate khác, reset đi trước
        if view.IsTemporaryHideIsolateActive():
            view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
            
        # Kích hoạt Isolate
        view.IsolateElementsTemporary(valid_ids)
        
        TransactionManager.Instance.TransactionTaskDone()
        return True, "Successfully ISOLATED {} element(s) in current view.".format(valid_ids.Count)
    except Exception as e:
        TransactionManager.Instance.ForceCloseTransaction()
        return False, "Failed to isolate elements: " + str(e)


# ================= THE MAIN RUN LOGIC ================= #
OUT = "No operation has been completed yet."
excelPath = select_excel_file()

if excelPath:
    raw_ids, error_msg = extract_element_ids_from_excel(excelPath)
    
    if error_msg:
        alert(error_msg, "System Error")
        OUT = error_msg
    elif raw_ids:
        # Thực hiện Isolate ID lên Revit Model
        success, message = isolate_elements(raw_ids)
        if success:
            alert(message, "Success")
            OUT = message
        else:
            alert(message, "Isolation Failed")
            OUT = message
    else:
        alert("Excel file read successfully but no ID data was found.", "Empty Result")
else:
    OUT = "Excel file was not selected."
