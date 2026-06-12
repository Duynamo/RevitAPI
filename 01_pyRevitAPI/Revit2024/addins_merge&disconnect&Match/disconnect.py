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

#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
class PipeSelectionFilter(ISelectionFilter):
    """Lớp lọc lựa chọn cho phép chọn Pipes, Pipe Fittings, và Pipe Accessories."""
    def AllowElement(self, element):
        category = element.Category
        if category is None:
            return False
        return category.Name in ["Pipes", "Pipe Fittings", "Pipe Accessories"]

    def AllowReference(self, reference, position):
        # Phương thức này phải được triển khai, trả về True là một mặc định an toàn.
        return True

def pick_element(prompt_message):
    """Cho phép người dùng chọn một đối tượng với bộ lọc và thông báo tùy chỉnh."""
    try:
        pipe_filter = PipeSelectionFilter()
        ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipe_filter, prompt_message)
        return doc.GetElement(ref.ElementId)
    except:
        return None

def find_connectors(element):
    """Tìm tất cả các connector của một đối tượng (Pipe, Fitting, Accessory)."""
    connectors = []
    if element is None:
        return connectors
        
    # Đối với Pipes, dùng ConnectorManager
    if hasattr(element, 'ConnectorManager') and element.ConnectorManager is not None:
        for conn in element.ConnectorManager.Connectors:
            connectors.append(conn)
    # Đối với Fittings/Accessories, dùng MEPModel.ConnectorManager
    elif hasattr(element, 'MEPModel') and element.MEPModel is not None:
        for conn in element.MEPModel.ConnectorManager.Connectors:
            connectors.append(conn)
            
    return connectors

def find_and_disconnect_nearest(element1, element2):
    """
    Tìm cặp connector gần nhất giữa hai đối tượng,
    và ngắt kết nối nếu chúng đang được kết nối.
    """
    connectors1 = find_connectors(element1)
    connectors2 = find_connectors(element2)

    if not connectors1 or not connectors2:
        return "Lỗi: Một hoặc cả hai đối tượng không có connector."

    # Tìm cặp connector gần nhất
    min_distance = float('inf')
    nearest_conn1 = None
    nearest_conn2 = None

    for c1 in connectors1:
        for c2 in connectors2:
            distance = c1.Origin.DistanceTo(c2.Origin)
            if distance < min_distance:
                min_distance = distance
                nearest_conn1 = c1
                nearest_conn2 = c2

    if nearest_conn1 is None or nearest_conn2 is None:
        return "Không thể tìm thấy các connector gần nhất."

    # Kiểm tra xem chúng có được kết nối với nhau không
    is_connected = False
    for ref_conn in nearest_conn1.AllRefs:
        if ref_conn.Owner.Id == nearest_conn2.Owner.Id:
            is_connected = True
            break
    
    if is_connected:
        TransactionManager.Instance.EnsureInTransaction(doc)
        try:
            nearest_conn1.DisconnectFrom(nearest_conn2)
            TransactionManager.Instance.TransactionTaskDone()
            return "Đã ngắt kết nối thành công cặp connector gần nhất."
        except Exception as ex:
            TransactionManager.Instance.TransactionTaskDone()
            return "Lỗi khi ngắt kết nối: " + str(ex)
    else:
        return "Các connector gần nhất không được kết nối với nhau."

#endregion

# --- Main execution ---
try:
    ele1 = pick_element("Chọn đối tượng thứ nhất")
    if ele1:
        ele2 = pick_element("Chọn đối tượng thứ hai")
        if ele2:
            # Thực hiện hành động
            result = find_and_disconnect_nearest(ele1, ele2)
            OUT = result
        else:
            OUT = "Đã hủy chọn đối tượng thứ hai."
    else:
        OUT = "Đã hủy chọn đối tượng thứ nhất."

except Exception as e:
    # Bắt lỗi nếu người dùng nhấn Esc
    if "OperationCanceledException" in str(e):
        OUT = "Thao tác đã bị hủy bởi người dùng."
    else:
        OUT = "Đã xảy ra lỗi: " + str(e)
