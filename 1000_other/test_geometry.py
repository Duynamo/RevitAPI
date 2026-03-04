from varmain.custom import * # type: ignore
from primitives import Box, Cylinder

def build_geometry_test(s, **kw):
    """
    Hàm này được gọi từ wrapper.py để dựng hình học thực tế.
    Tham số 's' là bắt buộc (kế thừa từ Plant 3D context).
    """
    try:
        # 1. TẠO KHỐI ĐẾ (Đế hình vuông)
        base_length = 100.0
        base_width  = 100.0
        base_height = 20.0

        # ✅ FIX: Truyền vị trí qua Origin= khi khởi tạo,
        #         KHÔNG dùng .translate() vì API Plant 3D không hỗ trợ
        base_box = Box(
            s,
            L=base_length,
            W=base_width,
            H=base_height,
            Origin=(0, 0, base_height / 2)   # Đáy box nằm đúng tại Z=0
        )

        # 2. TẠO KHỐI THÂN (Khối trụ tròn)
        cyl_radius = 25.0
        cyl_height = 150.0

        # ✅ FIX: Tương tự, dùng Origin= để định vị cylinder
        z_offset = base_height + (cyl_height / 2)
        body_cyl = Cylinder(
            s,
            R=cyl_radius,
            H=cyl_height,
            Origin=(0, 0, z_offset)           # Đáy cylinder nằm trên đỉnh base_box
        )

        # 3. TRẢ VỀ KẾT QUẢ
        # ✅ FIX: Trả về base_box — trong Plant 3D, tất cả các solid được tạo với
        #         cùng context 's' sẽ tự động được gộp vào scene.
        #         body_cyl đã được đăng ký vào 's' nên không cần gộp thủ công.
        return base_box

    except Exception as e:
        print("Loi trong qua trinh dung hinh test_geometry: " + str(e))
        raise
