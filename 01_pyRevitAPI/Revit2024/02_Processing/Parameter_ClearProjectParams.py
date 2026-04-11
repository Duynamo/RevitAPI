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


# === XÓA CẢ NON-SHARED PROJECT PARAMETERS VÀ SHARED PARAMETERS CÓ PREFIX "Params_" ===
# from Autodesk.Revit.DB import FilteredElementCollector, ParameterElement, Transaction, BuiltInParameterGroup

t = Transaction(doc, "Xóa Parameters prefix Params_ (Project & Shared)")
t.Start()

deleted_count = 0
ids_to_delete = []

# Thu thập tất cả ParameterElement trong project
param_elements = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

for param_el in param_elements:
    defi = param_el.GetDefinition()
    # if defi and defi.ParameterGroup == BuiltInParameterGroup.PG_DATA:
    if defi.Name.startswith("ISD_"):
        ids_to_delete.append(param_el.Id)
        deleted_count += 1

# Xóa bằng doc.Delete (xóa cả non-shared project parameters và binding của shared)
if ids_to_delete:
    for i in ids_to_delete:
        doc.Delete(i)

t.Commit()

# === KẾT QUẢ ===
if deleted_count > 0:
    OUT = "ok"
else:
    OUT = "Không tìm thấy parameter nào thỏa điều kiện."