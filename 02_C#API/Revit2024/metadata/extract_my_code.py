import os

def extract_pyrevit_scripts(workspace_path, output_file):
    """
    Quét thư mục làm việc hiện tại, trích xuất tất cả các file .py
    (bỏ qua môi trường ảo và file rác) để dạy AI phong cách code cá nhân.
    """
    print(f"Đang quét thư mục: {os.path.abspath(workspace_path)}...\n")

    # QUAN TRỌNG: Danh sách các thư mục phải bỏ qua
    # Nếu không bỏ qua .venv, AI sẽ đọc hàng ngàn file thư viện hệ thống gây rác data
    ignore_dirs = {'.venv', '!Backup', '.vscode', '__pycache__', 'metadata'}
    
    total_files = 0

    with open(output_file, 'w', encoding='utf-8') as outfile:
        # Ghi System Prompt mồi cho file
        outfile.write("# Cơ sở dữ liệu: pyRevit Custom Scripts & UI Framework\n\n")
        outfile.write("Tài liệu này chứa các mã nguồn Python hoàn chỉnh và các thiết kế UI (User Interface) đã hoạt động tốt trên pyRevit. ")
        outfile.write("Khi viết mã mới, HÃY ƯU TIÊN tái sử dụng các hàm (def), cách xử lý luồng, và phong cách viết giao diện từ các file mẫu trong tài liệu này.\n\n")

        for root, dirs, files in os.walk(workspace_path):
            # Chặn không cho os.walk đi vào các thư mục trong danh sách đen
            dirs[:] = [d for d in dirs if d not in ignore_dirs]

            for file in files:
                if file.endswith('.py'):
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, workspace_path)

                    try:
                        with open(file_path, 'r', encoding='utf-8') as infile:
                            content = infile.read()

                            outfile.write(f"## Context: File `{rel_path}`\n\n")
                            outfile.write("```python\n")
                            outfile.write(content)
                            outfile.write("\n```\n\n")
                            outfile.write("---\n\n")

                            total_files += 1
                            print(f"✔️ Đã lấy: {rel_path}")
                    except Exception as e:
                        print(f"❌ Bỏ qua file (Lỗi): {rel_path} - {e}")

    print(f"\n✅ HOÀN TẤT! Đã gộp thành công {total_files} file Python cá nhân.")
    print(f"📁 Dữ liệu AI đã sẵn sàng tại: {os.path.abspath(output_file)}")

# ==========================================
# CẤU HÌNH ĐƯỜNG DẪN
# ==========================================
# Sử dụng "." để quét từ thư mục gốc 01_pyRevitAPI (nếu bạn chạy file này ở ngoài)
# Hoặc r"C:\Users\Laptop\OneDrive\Desktop\RevitAPI\01_pyRevitAPI" để chỉ định cứng
THU_MUC_GOC = r"C:\Users\Laptop\OneDrive\Desktop\RevitAPI\01_pyRevitAPI"
FILE_DAU_RA = "My_pyRevit_Knowledge.md"

if __name__ == "__main__":
    extract_pyrevit_scripts(THU_MUC_GOC, FILE_DAU_RA)