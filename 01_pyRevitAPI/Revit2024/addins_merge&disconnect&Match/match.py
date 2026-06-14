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
"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""
class selectionFilter(ISelectionFilter):
    def __init__(self,categories):
        self.categories = categories
    def AllowElement(self, element):
        eleName = element.Category.Name in self.categories
        return eleName
def pickPipeOrPipePart():
    categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
    desEle = doc.GetElement(ref.ElementId)
    return desEle
def pickPipeOrPipeParts():
	partList = []
	categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
	fittingFilter = selectionFilter(categories)
	refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
	for ref in refs:
		desEle = doc.GetElement(ref.ElementId)
		partList.append(desEle)
	return partList

def is_connected(element):
    """Kiểm tra xem đối tượng có bất kỳ connector nào đang kết nối không."""
    connectors = []
    # Lấy connectors cho Pipe, Fitting, hoặc Accessory
    if hasattr(element, 'ConnectorManager') and element.ConnectorManager:
        connectors = element.ConnectorManager.Connectors
    elif hasattr(element, 'MEPModel') and element.MEPModel and element.MEPModel.ConnectorManager:
        connectors = element.MEPModel.ConnectorManager.Connectors
    else:
        return False

    for conn in connectors:
        if conn.IsConnected:
            return True # Tìm thấy một kết nối là đủ
    return False

def get_connected_pipes(element):
    """Lấy tất cả các đối tượng Pipe đang kết nối với một phụ kiện."""
    pipes = []
    connectors = []
    
    # Lấy connectors cho Fitting hoặc Accessory
    if hasattr(element, 'MEPModel') and element.MEPModel and element.MEPModel.ConnectorManager:
        connectors = element.MEPModel.ConnectorManager.Connectors
    else:
        return pipes # Không phải fitting/accessory hoặc không có connector

    for conn in connectors:
        if conn.IsConnected:
            for ref_conn in conn.AllRefs:
                owner = ref_conn.Owner
                # Chỉ lấy đối tượng là Pipe
                if isinstance(owner, Pipe):
                    # Tránh thêm trùng lặp
                    if owner.Id not in [p.Id for p in pipes]:
                        pipes.append(owner)
    return pipes
"""___"""
try:
    # 1. Chọn đối tượng gốc để lấy thông tin
    basePart = pickPipeOrPipePart()

    # 2. Chọn các đối tượng đích để gán thông tin
    listSetPart = pickPipeOrPipeParts()

    # 3. Lấy 'System Type' từ đối tượng gốc
    # Sử dụng BuiltInParameter để hoạt động trên mọi ngôn ngữ của Revit,
    # đảm bảo hoạt động cho cả Pipe, Pipe Fitting và Pipe Accessory.
    param_bip = BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
    base_param = basePart.get_Parameter(param_bip)

    if base_param and base_param.HasValue:
        base_system_type_id = base_param.AsElementId()
        
        # Lấy thông tin cần thiết để tạo ống ảo, phòng trường hợp cần dùng
        pipe_type_id_to_use = None
        level_id_to_use = None

        if isinstance(basePart, Pipe):
            pipe_type_id_to_use = basePart.PipeType.Id
            level_id_to_use = basePart.ReferenceLevel.Id
        else: # Nếu basePart là fitting, tìm fallback
            level_param = basePart.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
            if level_param:
                level_id_to_use = level_param.AsElementId()

        # Fallback nếu vẫn chưa có
        if not pipe_type_id_to_use:
            first_pipe_type = FilteredElementCollector(doc).OfClass(PipeType).FirstElement()
            if first_pipe_type:
                pipe_type_id_to_use = first_pipe_type.Id
        if not level_id_to_use:
            first_level = FilteredElementCollector(doc).OfClass(Level).FirstElement()
            if first_level:
                level_id_to_use = first_level.Id

        # 4. Gán 'System Type' cho các đối tượng đích
        updated_count = 0
        error_log = []

        TransactionManager.Instance.EnsureInTransaction(doc)
        for part in listSetPart:
            try:
                # --- XỬ LÝ ỐNG (PIPE) ---
                if isinstance(part, Pipe):
                    part_param = part.get_Parameter(param_bip)
                    if part_param and not part_param.IsReadOnly:
                        part_param.Set(base_system_type_id)
                        updated_count += 1
                    else:
                        error_log.append("Không thể cập nhật ống ID {} (tham số chỉ đọc).".format(part.Id))
                    continue # Chuyển sang đối tượng tiếp theo
                
                # --- XỬ LÝ PHỤ KIỆN (FITTING/ACCESSORY) ---
                if hasattr(part, 'MEPModel'):
                    # TRƯỜNG HỢP 1: PHỤ KIỆN ĐÃ ĐƯỢC KẾT NỐI
                    # Logic mới dựa trên phản hồi của bạn: thay đổi hệ thống của ống kết nối.
                    if is_connected(part):
                        connected_pipes = get_connected_pipes(part)
                        
                        if not connected_pipes:
                            error_log.append("Phụ kiện ID {} đã kết nối nhưng không tìm thấy ống nào để cập nhật.".format(part.Id))
                            continue

                        # Thay đổi hệ thống của tất cả các ống đang kết nối với phụ kiện
                        pipe_changed = False
                        for p in connected_pipes:
                            p_param = p.get_Parameter(param_bip)
                            if p_param and not p_param.IsReadOnly:
                                p_param.Set(base_system_type_id)
                                pipe_changed = True
                        
                        if pipe_changed:
                            updated_count += 1
                        else:
                            error_log.append("Không thể cập nhật các ống kết nối với phụ kiện ID {}.".format(part.Id))
                    # TRƯỜNG HỢP 2: PHỤ KIỆN ĐỨNG RIÊNG LẺ -> DÙNG ỐNG ẢO
                    else:
                        # Kiểm tra xem có đủ thông tin để tạo ống không
                        if not pipe_type_id_to_use or not level_id_to_use:
                            error_log.append("Không thể xử lý ID {}: Thiếu Pipe Type hoặc Level trong dự án.".format(part.Id))
                            continue

                        # Tìm một connector bất kỳ trên phụ kiện để bắt đầu
                        conn_to_use = None
                        if part.MEPModel.ConnectorManager.Connectors.Size > 0:
                            conn_to_use = list(part.MEPModel.ConnectorManager.Connectors)[0]
                        
                        if not conn_to_use:
                            error_log.append("Lỗi xử lý ID {}: Phụ kiện không có connector.".format(part.Id))
                            continue

                        # Lấy thông tin cần thiết từ connector để tạo ống ảo
                        fitting_diameter = conn_to_use.Radius * 2
                        start_point = conn_to_use.Origin
                        
                        # Xác định hướng của ống ảo (ngắn)
                        direction = XYZ.BasisX
                        if conn_to_use.CoordinateSystem and not conn_to_use.CoordinateSystem.BasisZ.IsZeroLength():
                            direction = conn_to_use.CoordinateSystem.BasisZ
                        end_point = start_point + direction.Normalize() * 0.1 # Dài 0.1 feet

                        # Bước 1 & 2: Tạo một ống ảo (virtual pipe) và gán Piping System cho nó.
                        # Ghi chú: Piping System được lấy từ đối tượng gốc (basePart) vì đó là mục tiêu của lệnh "Match".
                        # Lệnh Pipe.Create() yêu cầu Piping System ngay khi tạo, nên 2 bước này được kết hợp.
                        virtual_pipe = Pipe.Create(doc, base_system_type_id, pipe_type_id_to_use, level_id_to_use, start_point, end_point)
                        
                        # Cấu hình đường kính cho ống ảo
                        pipe_diameter_param = virtual_pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                        if pipe_diameter_param:
                            pipe_diameter_param.Set(fitting_diameter)
                        
                        doc.Regenerate() # Cần thiết để connector của ống ảo được tạo ra

                        # Tìm connector của ống ảo tại điểm kết nối
                        pipe_conns_at_start = [c for c in virtual_pipe.ConnectorManager.Connectors if c.Origin.IsAlmostEqualTo(start_point)]

                        if pipe_conns_at_start:
                            pipe_conn = pipe_conns_at_start[0]
                            
                            # Bước 3: Kết nối ống ảo đó với phụ kiện đích.
                            # Khi kết nối, phụ kiện sẽ tự động nhận Piping System từ ống ảo.
                            conn_to_use.ConnectTo(pipe_conn)
                            updated_count += 1
                            # Bước 4: Tạm thời dừng lại ở bước đó.
                            # Theo yêu cầu, ống ảo sẽ được giữ lại trong mô hình để kiểm tra.
                        else:
                            error_log.append("Lỗi xử lý ID {}: Không tìm thấy connector của ống ảo.".format(part.Id))

            except Exception as e:
                error_log.append("Lỗi với đối tượng ID {}: {}".format(part.Id, str(e)))
        TransactionManager.Instance.TransactionTaskDone()

        # 5. Thông báo kết quả
        OUT = "Đã cập nhật thành công 'System Type' cho {}/{} đối tượng.".format(updated_count, len(listSetPart))
        if error_log:
            OUT += "\n\nChi tiết lỗi:\n" + "\n".join(error_log)
    else:
        OUT = "Lỗi: Không thể xác định 'System Type' từ đối tượng gốc."

except Autodesk.Revit.Exceptions.OperationCanceledException:
    OUT = "Thao tác đã bị người dùng hủy bỏ."
except Exception as ex:
    OUT = "Đã xảy ra lỗi không mong muốn: " + str(ex)
