import os
from bs4 import BeautifulSoup

def clean_html_to_markdown_with_progress(html_dir, output_file):
    print(f"Đang chuẩn bị dữ liệu từ: {os.path.abspath(html_dir)}...")
    
    if not os.path.exists(html_dir):
        print(f"❌ Không tìm thấy thư mục: {html_dir}")
        return

    print("Đang đếm tổng số file HTML... (vui lòng đợi vài giây)")
    total_files = 0
    for root, dirs, files in os.walk(html_dir):
        for file in files:
            if file.endswith('.htm') or file.endswith('.html'):
                total_files += 1
                
    if total_files == 0:
        print("❌ Không có file HTML nào để xử lý!")
        return
        
    print(f"Tìm thấy tổng cộng {total_files} file. Bắt đầu bóc tách...\n")

    processed_files = 0
    saved_files = 0
    
    with open(output_file, 'w', encoding='utf-8') as outfile:
        # Ghi chú cho AI hiểu đây là bản 2027
        outfile.write("# Cơ sở dữ liệu: Revit 2027 API Classes & Methods Reference\n\n")

        for root, dirs, files in os.walk(html_dir):
            for file in files:
                if file.endswith('.htm') or file.endswith('.html'):
                    file_path = os.path.join(root, file)
                    processed_files += 1
                    
                    try:
                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as infile:
                            soup = BeautifulSoup(infile.read(), 'html.parser')
                            
                            title = soup.find('title')
                            title_text = title.get_text(strip=True) if title else "Untitled API Concept"
                            
                            for element in soup(['script', 'style', 'nav', 'header', 'footer']):
                                element.decompose()
                            
                            main_text = soup.get_text(separator='\n', strip=True)
                            
                            if len(main_text) > 100:
                                outfile.write(f"## {title_text}\n\n")
                                outfile.write(f"{main_text}\n\n")
                                outfile.write("---\n\n")
                                saved_files += 1
                                
                    except Exception:
                        pass 

                    # Cập nhật tiến độ % trên Terminal
                    if processed_files % 50 == 0 or processed_files == total_files:
                        percent = (processed_files / total_files) * 100
                        print(f"\r⏳ Tiến độ: {percent:.1f}% ({processed_files}/{total_files} file)", end="", flush=True)

    print(f"\n\n✅ HOÀN TẤT! Đã lọc sạch và lưu thành công {saved_files} trang tài liệu API.")
    print(f"📁 Dữ liệu AI đã sẵn sàng tại: {os.path.abspath(output_file)}")

# ==========================================
# CẤU HÌNH ĐƯỜNG DẪN 
# ==========================================
THU_MUC_HTML = "RevitAPI_Extracted/html" 
# Lưu file thành tên 2027
FILE_DAU_RA = "Revit2027_API_Dictionary.md"

if __name__ == "__main__":
    clean_html_to_markdown_with_progress(THU_MUC_HTML, FILE_DAU_RA)