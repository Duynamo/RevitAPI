from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

# Lấy danh sách các ống trong mô hình
pipes_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves)
pipes = [pipe for pipe in pipes_collector]

# Giả sử bạn muốn tìm connector gần nhất giữa ống thứ nhất và ống thứ hai trong danh sách ống
pipe1 = pipes[0]
pipe2 = pipes[1]

# Tạo một biến lưu khoảng cách gần nhất
closest_distance = float('inf')  # Gán giá trị lớn nhất cho khoảng cách ban đầu
closest_connector_pipe1 = None
closest_connector_pipe2 = None

# Lặp qua connectors của ống thứ nhất
for connector1 in pipe1.ConnectorManager.Connectors:
    for connector2 in pipe2.ConnectorManager.Connectors:
        distance = connector1.Origin.DistanceTo(connector2.Origin)
        if distance < closest_distance:
            closest_distance = distance
            closest_connector_pipe1 = connector1
            closest_connector_pipe2 = connector2

# In ra thông tin của hai connector gần nhau nhất
if closest_connector_pipe1 and closest_connector_pipe2:
    print(f"Closest connectors found with distance: {closest_distance}")
    print(f"Connector 1: {closest_connector_pipe1.Id}")
    print(f"Connector 2: {closest_connector_pipe2.Id}")
else:
    print("No connectors found.")
