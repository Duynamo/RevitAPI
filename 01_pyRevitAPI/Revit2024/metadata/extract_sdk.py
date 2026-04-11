import os

def extract_csharp_from_sdk(sdk_samples_path, output_md_file):
    """
    Quét thư mục Samples của Revit SDK và trích xuất code C# thành file Markdown để train AI.
    """
    print(f"Đang quét thư mục: {sdk_samples_path}...\n")
    
    if not os.path.exists(sdk_samples_path):
        print(f"❌ Lỗi: Không tìm thấy thư mục {sdk_samples_path}")
        print("Vui lòng kiểm tra lại đường dẫn.")
        return

    total_files = 0
    
    # Mở file output để ghi dữ liệu (ưu tiên UTF-8)
    with open(output_md_file, 'w', encoding='utf-8') as outfile:
        # Ghi tiêu đề và mô tả cho AI hiểu Context
        outfile.write("# Cơ sở dữ liệu: Revit 2026 API C# Samples\n\n")
        outfile.write("Tài liệu này chứa các mã nguồn mẫu (sample code) từ Revit 2026 SDK. ")
        outfile.write("Hãy sử dụng cấu trúc, class, và method trong đây làm ngữ cảnh chuẩn xác để viết mã Revit API.\n\n")
        
        # Duyệt qua tất cả các thư mục con trong folder Samples
        for root, dirs, files in os.walk(sdk_samples_path):
            for file in files:
                # Chỉ lấy các file mã nguồn C#
                if file.endswith('.cs'):
                    file_path = os.path.join(root, file)
                    
                    # Trích xuất đường dẫn tương đối (VD: CreateShared/CS/Command.cs)
                    # Điều này rất quan trọng để AI biết đoạn code thuộc về Project/Tính năng nào
                    rel_path = os.path.relpath(file_path, sdk_samples_path)
                    
                    try:
                        # Đọc nội dung file
                        with open(file_path, 'r', encoding='utf-8') as infile:
                            content = infile.read()
                            
                            # Ghi vào file tổng hợp
                            outfile.write(f"## Context: File `{rel_path}`\n\n")
                            outfile.write("```csharp\n")
                            outfile.write(content)
                            outfile.write("\n```\n\n")
                            outfile.write("---\n\n") # Đường gạch ngang phân cách
                            
                            total_files += 1
                            print(f"✔️ Đã lấy: {rel_path}")
                    
                    except UnicodeDecodeError:
                        # Fallback: Một số file SDK cũ có thể không lưu ở dạng UTF-8
                        try:
                            with open(file_path, 'r', encoding='latin-1') as infile:
                                content = infile.read()
                                outfile.write(f"## Context: File `{rel_path}`\n\n")
                                outfile.write("```csharp\n")
                                outfile.write(content)
                                outfile.write("\n```\n\n")
                                outfile.write("---\n\n")
                                total_files += 1
                                print(f"⚠️ Đã lấy (latin-1): {rel_path}")
                        except Exception as e:
                            print(f"❌ Bỏ qua file (Lỗi định dạng): {rel_path}")

    print(f"\n✅ HOÀN TẤT! Đã gộp thành công {total_files} file mã nguồn.")
    print(f"📁 Dữ liệu AI đã sẵn sàng tại: {os.path.abspath(output_md_file)}")

# ==========================================
# CẤU HÌNH ĐƯỜNG DẪN (Dựa theo ảnh của bạn)
# ==========================================
# Khai báo r"..." để tránh lỗi ký tự đặc biệt trong đường dẫn Windows
THU_MUC_SDK = r"C:\Revit 2026 SDK\Samples" 
FILE_DAU_RA = "Revit2026_AI_Training_Data.md"

# Chạy lệnh
if __name__ == "__main__":
    extract_csharp_from_sdk(THU_MUC_SDK, FILE_DAU_RA)