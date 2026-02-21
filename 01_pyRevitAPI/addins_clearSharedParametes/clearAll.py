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

class ParaProj():
    def __init__(self, definition, binding):
        self.definition = definition
        self.binding = binding
        self.defName = definition.Name if definition else "None"
        self.group_label = LabelUtils.GetLabelFor(definition.ParameterGroup) if definition else "Unknown"
        self.group_id = definition.ParameterGroup if definition else None
        self.type = "Instance" if isinstance(binding, InstanceBinding) else "Type"
        self.thesecats = []
        
        if binding and binding.Categories:
            for cat in binding.Categories:
                try:
                    self.thesecats.append(cat.Name)
                except:
                    self.thesecats.append("Unknown Category")

    @property
    def isShared(self):
        if not self.definition:
            return False
        coll = FilteredElementCollector(doc).OfClass(SharedParameterElement).ToElements()
        shared_names = {sp.GetDefinition().Name for sp in coll}
        return self.defName in shared_names


# === PHẦN 1: LẤY TẤT CẢ PROJECT PARAMETER ===
para_list = []
iterator = doc.ParameterBindings.ForwardIterator()
iterator.Reset()  # Bắt buộc!

while iterator.MoveNext():
    definition = iterator.Key
    binding = iterator.Current
    if definition:
        para_list.append(ParaProj(definition, binding))

# === PHẦN 2: LỌC VÀ XÓA TẤT CẢ PARAMETER THUỘC NHÓM "Data" (PG_DATA) ===
t = Transaction(doc, "Xóa tất cả Project Parameter nhóm Data")
t.Start()

deleted_count = 0
definitions_to_remove = []

for para in para_list:
    # Kiểm tra chính xác bằng BuiltInParameterGroup.PG_DATA
    if para.group_id == BuiltInParameterGroup.PG_DATA:
        definitions_to_remove.append(para.definition)
        deleted_count += 1

# Xóa binding
for defi in definitions_to_remove:
    try:
        doc.ParameterBindings.Remove(defi)
    except:
        pass  # Bỏ qua nếu Revit từ chối (rất hiếm)

t.Commit()

# === KẾT QUẢ TRẢ VỀ ===
if deleted_count > 0:
    OUT = "Hello World"
else:
    OUT = "Không tìm thấy Project Parameter nào thuộc nhóm 'Data' để xóa."
