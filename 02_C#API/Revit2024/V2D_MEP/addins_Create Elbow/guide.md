Bạn là một chuyên gia lập trình Revit C# API cấp cao. Hãy viết một file mã nguồn C# hoàn chỉnh triển khai một IExternalCommand cho Revit đảm bảo các tiêu chí: code sạch, tư duy tường minh, dễ hiểu, dễ bảo trì cho người mới và có chú thích chi tiết bằng tiếng Việt trên từng dòng lệnh/phương thức.

### YÊU CẦU CHỨC NĂNG (Mục đích: Đặt và xoay Khuỷu ống - Elbow)
1. **Bước 1: Chọn vị trí đầu ống**
   - Sử dụng `uiDoc.Selection.PickPoint` để người dùng click chọn một vị trí mong muốn trên không gian.
   - Sử dụng `uiDoc.Selection.PickObject` kết hợp với một bộ lọc tuyển chọn đối tượng (`ISelectionFilter`) để đảm bảo người dùng chỉ chọn được đối tượng thuộc lớp `MEPCurve` (Ống).

2. **Bước 2: Xác định Connector tối ưu**
   - Duyệt qua toàn bộ danh sách `Connectors` của đoạn ống vừa chọn.
   - Tính toán khoảng cách từ điểm `PickPoint` ở Bước 1 đến từng Connector bằng hàm `DistanceTo()`.
   - Tìm ra Connector có khoảng cách gần nhất để làm điểm gốc đặt fitting.

3. **Bước 3: Tạo Fitting Elbow**
   - Tạo một Transaction để thực thi việc tạo hình.
   - Sử dụng `RoutingPreferenceManager` từ cấu hình định tuyến của ống để lấy ra `ElementId` của Family loại Elbow mặc định phù hợp với kích thước bán kính của Connector vừa tìm được.
   - Thực hiện kiểm tra, kích hoạt `FamilySymbol` nếu nó chưa được active.
   - Tạo thực thể khuỷu bằng phương thức `doc.Create.NewElbowFitting`.

4. **Bước 4: Giao diện UI xoay cấu kiện liên tục (Vòng lặp tương tác)**
   - Khởi tạo một cửa sổ Windows Form hoặc WPF (ưu tiên khởi tạo UI nhanh ngay trong code bằng thư viện System.Windows.Controls hoặc System.Windows.Forms) hiển thị ở chế độ Topmost (luôn nằm trên cùng).
   - Giao diện gồm có: 1 ô TextBox nhập góc quay (độ), 1 nút "Quay / Áp dụng", và 1 nút "Close" (hoặc nút X đóng cửa sổ).
   - Khi người dùng nhập góc và nhấn "Quay / Áp dụng":
     + Tạo một Transaction con.
     + Xác định trục xoay chính là đường thẳng đi qua tâm Connector và có Vector hướng là Vector pháp tuyến (BasisZ) của hệ tọa độ Connector đó. Sử dụng `Line.CreateUnbound`.
     + Đổi đơn vị góc nhập vào từ Độ sang Radian.
     + Sử dụng `ElementTransformUtils.RotateElement` để thực hiện lệnh xoay cấu kiện Elbow vừa tạo quanh trục.
     + Gọi `uiDoc.RefreshActiveView()` để cập nhật trực quan trên màn hình ngay lập tức cho người dùng quan sát.
   - Luồng lặp này sẽ diễn ra liên tục cho đến khi người dùng nhấn nút "Close" hoặc nút "X" để đóng cửa sổ UI, lúc này Tool mới chính thức dừng hoạt động và trả về kết quả `Result.Succeeded`.