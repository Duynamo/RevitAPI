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
    if not os.path.exists(file_path):
        return None, "File not found"
    
    excel = workbook = sheet = None
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path)
        sheet = workbook.Worksheets[sheet_name]
        
        used = sheet.UsedRange
        rows, cols = used.Rows.Count, used.Columns.Count
        data = []
        
        headers = [str(sheet.Cells[1, c].Value2 or f"Col{c}") for c in range(1, cols+1)]
        name_idx = headers.index("Shared Parameters Name")
        group_idx = headers.index("Group") if "Group" in headers else -1
        type_idx = headers.index("Type") if "Type" in headers else -1
        
        for r in range(2, rows+1):
            name = str(sheet.Cells[r, name_idx+1].Value2 or "").strip()
            group = str(sheet.Cells[r, group_idx+1].Value2 or "").strip() if group_idx >= 0 else "ProjectParameters"
            typ = str(sheet.Cells[r, type_idx+1].Value2 or "").strip() if type_idx >= 0 else "Text"
            if name: data.append((name, group, typ))
        
        return data, None
    except Exception as e:
        return None, str(e)
    finally:
        for obj in (sheet, workbook, excel):
            if obj: System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)

def create_and_bind_shared_parameters(param_data):
    cats = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_GenericModel]
    
    t = Transaction(doc, "Add Shared Parameters")
    t.Start()
    try:
        spf = app.OpenSharedParameterFile() or app.Application.OpenSharedParameterFile()
        cat_set = app.Create.NewCategorySet()
        for c in cats:
            cat = doc.Settings.Categories.get_Item(c)
            if cat: cat_set.Insert(cat)
        
        for name, group_name, typ in param_data:
            if typ != "Text": continue
            group = spf.Groups.get_Item(group_name) or spf.Groups.Create(group_name)
            defn = group.Definitions.get_Item(name)
            if not defn:
                opts = ExternalDefinitionCreationOptions(name, SpecTypeId.String.Text)
                opts.Visible = opts.UserModifiable = True
                defn = group.Definitions.Create(opts)
            bind = app.Create.NewInstanceBinding(cat_set)
            doc.ParameterBindings.Insert(defn, bind, BuiltInParameterGroup.PG_DATA)
        
        t.Commit()
        return True, "Done"
    except Exception as e:
        t.RollBack()
        return False, str(e)

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