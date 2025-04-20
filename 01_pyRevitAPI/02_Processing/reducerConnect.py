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
def getNominalDiameter(pipe):
    """Lấy đường kính danh nghĩa của ống (đơn vị: feet)."""
    return pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()

def getAvailableReducers(pipe_type):
    """Lấy tất cả reducer từ Pipe Type của ống."""
    reducers = []
    # Lấy tất cả Pipe Fitting Families
    for family in FilteredElementCollector(doc).OfClass(Family):
        if family.FamilyCategory.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
            for symbol_id in family.GetFamilySymbolIds():
                symbol = doc.GetElement(symbol_id)
                # Kiểm tra nếu là reducer (có Nominal Diameter 1 và 2)
                if symbol.LookupParameter("Nominal Diameter 1") and symbol.LookupParameter("Nominal Diameter 2"):
                    nd1 = symbol.LookupParameter("Nominal Diameter 1").AsDouble()
                    nd2 = symbol.LookupParameter("Nominal Diameter 2").AsDouble()
                    if nd1 != nd2:  # Đảm bảo là reducer (đường kính khác nhau)
                        # Kiểm tra nếu reducer thuộc Pipe Type
                        if symbol.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId() == pipe_type.Id:
                            reducers.append((max(nd1, nd2), min(nd1, nd2)))  # (lớn, nhỏ)
    return reducers

def findReducerSequence(start_dia, end_dia, available_reducers):
    """Tìm chuỗi reducer để chuyển từ start_dia xuống end_dia."""
    if start_dia <= end_dia:
        return []  # Không cần reducer nếu đường kính không giảm
    
    sequence = []
    current_dia = start_dia
    available_reducers = sorted(available_reducers, key=lambda x: x[0], reverse=True)  # Sắp xếp theo đường kính lớn
    while current_dia > end_dia:
        found = False
        for large_dia, small_dia in available_reducers:
            if abs(current_dia - large_dia) < 1e-6 and small_dia >= end_dia:
                sequence.append((large_dia, small_dia))
                current_dia = small_dia
                found = True
                break
        if not found:
            return None  # Không tìm thấy chuỗi reducer phù hợp
    return sequence

def placeReducer(doc, large_pipe, small_dia, reducer_large_dia, reducer_small_dia, connection_point):
    """Đặt một reducer tại điểm kết nối và tạo ống trung gian."""
    pipe_type = large_pipe.PipeType
    piping_system = large_pipe.MEPSystem
    
    # Tìm family reducer phù hợp
    reducer_family = None
    for family in FilteredElementCollector(doc).OfClass(Family):
        if family.FamilyCategory.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
            for symbol_id in family.GetFamilySymbolIds():
                symbol = doc.GetElement(symbol_id)
                if symbol.LookupParameter("Nominal Diameter 1") and symbol.LookupParameter("Nominal Diameter 2"):
                    nd1 = symbol.LookupParameter("Nominal Diameter 1").AsDouble()
                    nd2 = symbol.LookupParameter("Nominal Diameter 2").AsDouble()
                    if (abs(nd1 - reducer_large_dia) < 1e-6 and abs(nd2 - reducer_small_dia) < 1e-6) or \
                       (abs(nd1 - reducer_small_dia) < 1e-6 and abs(nd2 - reducer_large_dia) < 1e-6):
                        if symbol.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId() == pipe_type.Id:
                            reducer_family = symbol
                            break
        if reducer_family:
            break
    
    if not reducer_family:
        return None, None  # Không tìm thấy reducer
    
    # Kích hoạt FamilySymbol nếu chưa
    if not reducer_family.IsActive:
        reducer_family.Activate()
    
    # Tạo ống trung gian với đường kính nhỏ hơn
    pipe_curve = large_pipe.Location.Curve
    direction = (pipe_curve.GetEndPoint(1) - pipe_curve.GetEndPoint(0)).Normalize()
    mid_point = connection_point + direction * 0.5  # Ống trung gian ngắn (0.5 feet)
    new_pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc, pipe_type.Id, piping_system.Id, large_pipe.LevelId, connection_point, mid_point)
    new_pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(small_dia)
    if piping_system is not None:
        piping_system.Add(new_pipe)
    
    # Đặt reducer
    reducer = doc.Create.NewFamilyInstance(connection_point, reducer_family, large_pipe, StructuralType.NonStructural)
    
    return new_pipe, reducer


def get_mepcurvetype_properties(pipe):
    """Lấy tất cả properties, parameters và tên của các Family liên kết với PipeType từ một Pipe."""
    # if not pipe or not isinstance(pipe, Pipe):
    #     return None
    
    # Lấy PipeType (kế thừa từ MEPCurveType)
    pipe_type = pipe.PipeType
    if not pipe_type:
        return None
    
    # Dictionary để lưu properties, parameters và families
    all_properties = {"Properties": {}, "Parameters": {}, "Families": []}
    
    # 1. Lấy các properties của PipeType (sử dụng dir() để liệt kê)
    for attr in dir(pipe_type):
        try:
            # Loại bỏ các phương thức và thuộc tính riêng tư (bắt đầu bằng '_')
            if not attr.startswith('_'):
                value = getattr(pipe_type, attr)
                # Kiểm tra xem value có phải là thuộc tính (không phải phương thức)
                if not callable(value):
                    # Chuyển đổi giá trị thành string để dễ đọc
                    if isinstance(value, ElementId):
                        value = str(value.IntegerValue)
                    elif value is None:
                        value = "None"
                    else:
                        value = str(value)
                    all_properties["Properties"][attr] = value
        except:
            pass  # Bỏ qua nếu không thể truy cập thuộc tính
    
    # 2. Lấy các parameters của PipeType
    for param in pipe_type.Parameters:
        try:
            param_name = param.Definition.Name
            # Lấy giá trị của parameter dựa trên kiểu dữ liệu
            if param.StorageType == StorageType.String:
                value = param.AsString() or "None"
            elif param.StorageType == StorageType.Double:
                value = param.AsDouble()
            elif param.StorageType == StorageType.Integer:
                value = param.AsInteger()
            elif param.StorageType == StorageType.ElementId:
                value = str(param.AsElementId().IntegerValue)
            else:
                value = "Unknown"
            all_properties["Parameters"][param_name] = value
        except:
            pass  # Bỏ qua nếu không thể đọc parameter
    
    # 3. Lấy tên các Family của Pipe Fitting liên kết với PipeType
    family_names = set()
    for family in FilteredElementCollector(doc).OfClass(Family):
        if family.FamilyCategory.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
            for symbol_id in family.GetFamilySymbolIds():
                symbol = doc.GetElement(symbol_id)
                try:
                    # Kiểm tra nếu FamilySymbol thuộc PipeType
                    if symbol.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId() == pipe_type.Id:
                        family_names.add(family.Name)
                except:
                    pass  # Bỏ qua nếu không thể truy cập tham số
    all_properties["Families"] = sorted(list(family_names))  # Sắp xếp danh sách tên Family
    
    return all_properties, all_properties
# def connectPipesWithReducers(doc, pipe1, pipe2):
#     """Nối hai ống với các reducer phù hợp từ Pipe Type."""
#     dia1 = getNominalDiameter(pipe1)
#     dia2 = getNominalDiameter(pipe2)
    
#     # Đảm bảo pipe1 là ống lớn hơn
#     large_pipe = pipe1 if dia1 > dia2 else pipe2
#     small_pipe = pipe2 if dia1 > dia2 else pipe1
#     large_dia = max(dia1, dia2)
#     small_dia = min(dia1, dia2)
    
#     # Lấy danh sách reducer từ Pipe Type của ống lớn
#     available_reducers = getAvailableReducers(large_pipe.PipeType)
#     if not available_reducers:
#         return None  # Không có reducer nào khả dụng
    
#     # Tìm chuỗi reducer
#     reducer_sequence = findReducerSequence(large_dia, small_dia, available_reducers)
#     if not reducer_sequence:
#         return None  # Không thể nối
    
#     # Lấy vị trí kết nối (đầu của ống nhỏ)
#     small_curve = small_pipe.Location.Curve
#     connection_point = small_curve.GetEndPoint(0)
    
#     new_pipes = []
#     reducers = []
#     current_pipe = large_pipe
#     current_point = connection_point
    
#     TransactionManager.Instance.EnsureInTransaction(doc)
#     for large_dia, small_dia in reducer_sequence:
#         new_pipe, reducer = placeReducer(doc, current_pipe, small_dia, large_dia, small_dia, current_point)
#         if new_pipe and reducer:
#             new_pipes.append(new_pipe)
#             reducers.append(reducer)
#             current_pipe = new_pipe
#             current_point = current_pipe.Location.Curve.GetEndPoint(1)
    
#     # Kết nối ống cuối cùng với small_pipe
#     if current_pipe and small_pipe:
#         try:
#             PlumbingUtils.ConnectPipe(doc, current_pipe.Id, small_pipe.Id)
#         except:
#             pass  # Bỏ qua nếu không thể kết nối trực tiếp
    
#     TransactionManager.Instance.TransactionTaskDone()
#     return new_pipes, reducers

# Chạy hàm chính
# try:
#     pipe1 = pickPipe("Select the first pipe")
#     pipe2 = pickPipe("Select the second pipe")
#     result = connectPipesWithReducers(doc, pipe1, pipe2)
#     if result is None:
#         print("Cannot connect pipes with available reducers.")
#     else:
#         print("Pipes connected successfully with reducers.")
# except Exception as ex:
#     print(f"Error: {str(ex)}")
#     TransactionManager.Instance.ForceCloseTransaction()

# fitting = pickFittingOrAccessory()
pipe1 = pickPipe()
allProperties = get_mepcurvetype_properties(pipe1)

# OUT = connectPipesWithReducers(doc, pipe1, pipe2)
OUT = allProperties