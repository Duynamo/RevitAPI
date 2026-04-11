import os

def extract_csharp_from_sdk(sdk_samples_path, output_md_file):
    """
    Quét thư mục Samples của Revit SDK và trích xuất code C# thành file Markdown.
    Định dạng được tối ưu hóa cho AI đọc hiểu ngữ cảnh (Context).
    """
    print(f"Đang quét thư mục: {sdk_samples_path}...\n")
    
    if not os.path.exists(sdk_samples_path):
        print(f"❌ Lỗi: Không tìm thấy thư mục {sdk_samples_path}")
        print("Vui lòng kiểm tra lại đường dẫn.")
        return

    total_files = 0
    
    # Mở file output để ghi dữ liệu (ưu tiên UTF-8)
    with open(output_md_file, 'w', encoding='utf-8') as outfile:
        # Ghi System Prompt mồi để AI hiểu dữ liệu nó đang đọc là gì
        outfile.write("# Cơ sở dữ liệu: Revit API C# Samples\n\n")
        outfile.write("Tài liệu này chứa các mã nguồn mẫu (sample code) chính thức từ Autodesk Revit SDK. ")
        outfile.write("Hãy sử dụng cấu trúc, class, và method trong đây làm ngữ cảnh chuẩn xác để viết mã Revit API, thay vì tự bịa ra các hàm không tồn tại.\n\n")
        
        # Duyệt qua tất cả các thư mục con trong folder Samples
        for root, dirs, files in os.walk(sdk_samples_path):
            for file in files:
                # Chỉ lấy các file mã nguồn C# (.cs)
                if file.endswith('.cs'):
                    file_path = os.path.join(root, file)
                    
                    # Trích xuất đường dẫn tương đối (VD: CreateShared/CS/Command.cs)
                    # Giúp AI phân loại được đoạn code này thuộc về tính năng quản lý Parameter hay Geometry
                    rel_path = os.path.relpath(file_path, sdk_samples_path)
                    
                    try:
                        # Đọc nội dung file
                        with open(file_path, 'r', encoding='utf-8') as infile:
                            content = infile.read()
                            
                            # Ghi vào file tổng hợp theo định dạng Markdown chuẩn
                            outfile.write(f"## Context: File `{rel_path}`\n\n")
                            outfile.write("```csharp\n")
                            outfile.write(content)
                            outfile.write("\n```\n\n")
                            outfile.write("---\n\n") # Đường gạch ngang phân cách
                            
                            total_files += 1
                            print(f"✔️ Đã lấy: {rel_path}")
                    
                    except UnicodeDecodeError:
                        # Fallback: Một số file C# cũ của Autodesk lưu ở dạng ANSI/Latin-1
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
# CẤU HÌNH ĐƯỜNG DẪN 
# ==========================================
# Khai báo r"..." để tránh lỗi ký tự đặc biệt trong đường dẫn Windows
THU_MUC_SDK = r"C:\Revit 2026 SDK\Samples" # Cập nhật theo thư mục thực tế trên máy bạn
FILE_DAU_RA = "Revit_AI_Training_Data.md"

# Chạy lệnh
if __name__ == "__main__":
    extract_csharp_from_sdk(THU_MUC_SDK, FILE_DAU_RA)