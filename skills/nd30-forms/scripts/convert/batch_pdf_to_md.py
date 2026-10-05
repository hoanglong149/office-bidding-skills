import os
import subprocess
import re
from datetime import datetime
from pathlib import Path

def post_process_markdown(md_path, title_suggestion=None):
    """
    Chuẩn hóa nội dung Markdown theo chuẩn Obsidian Skill.
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Thêm YAML Frontmatter (nếu chưa có)
    if not content.startswith('---'):
        title = title_suggestion or Path(md_path).stem.replace('-', ' ').title()
        date_str = datetime.now().strftime('%Y-%m-%d')
        yaml_header = f"---\ntitle: \"{title}\"\ndate: {date_str}\ntags: [converted, pdf, automated]\n---\n\n"
        content = yaml_header + content

    # Tách YAML và Body để tránh làm hỏng Frontmatter
    parts = content.split('---', 2)
    if len(parts) >= 3:
        yaml_part = parts[1]
        body_part = parts[2]
        
        # 1. Xóa Header lặp lại và thay bằng Page Marker
        # Danh sách các cụm Header lặp lại (theo tài liệu Phả Lại)
        headers_to_remove = [
            r'HỒ SƠ MỜI THẦU',
            r'PHẦN 1 - CHƯƠNG I',
            r'VIỆN NĂNG LƯỢNG',
            r'MÔ TẢ TÓM TẮT'
        ]
        
        # Regex linh hoạt cho Header (có thể có # hoặc không, có khoảng trắng)
        pattern_header = r'(?:#\s*)?(?:' + '|'.join(headers_to_remove) + r')\s*\n*'
        
        page_counter = [1]
        def page_replacer(match):
            page_num = page_counter[0]
            marker = f"\n\n%% [Trang {page_num}] %%\n\n"
            page_counter[0] += 1
            return marker

        # Thay thế các cụm header lặp lại (ít nhất 2 dòng tiêu đề liên tiếp)
        body_part = re.sub(f"({pattern_header}){{2,}}", page_replacer, body_part)

        # 2. Tự động tạo WikiLink cho các thuật ngữ (idempotent: không lồng nếu đã có link)
        body_part = re.sub(r'(?<!\|)(?<!\[)\bBDL\b(?!\])', r'[[Chương II - Bảng dữ liệu đấu thầu 29.3|BDL]]', body_part)
        body_part = re.sub(r'(?<!\|)(?<!\[)\bCDNT\b(?!\])', r'[[Chương I - Chỉ dẫn nhà thầu 29.3|CDNT]]', body_part)
        body_part = re.sub(r'(?<!\|)(?<!\[)\bHSMT\b(?!\])', r'[[Hồ sơ mời thầu|HSMT]]', body_part)

        # 3. Chuẩn hóa ngày tháng trong phần Body
        def date_replacer(match):
            d, m = match.group(1), match.group(2)
            y = match.group(3) if match.group(3) else "2026"
            if len(y) == 2: y = "20" + y
            return f"{y}-{m.zfill(2)}-{d.zfill(2)}"

        body_part = re.sub(r'\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b', date_replacer, body_part)
        
        # 4. Chuyển đổi thành Obsidian Callouts
        body_part = re.sub(r'▸ \*\* Đầu ra: \*\*', r'> [!success] Đầu ra', body_part)
        body_part = re.sub(r'▸ \*\* Không nên chứa: \*\*', r'> [!warning] Không nên chứa', body_part)
        body_part = re.sub(r'▸ \*\* Lưu ý: \*\*', r'> [!info] Lưu ý', body_part)

        # 5. Highlight tiền tệ
        body_part = re.sub(r'(\d{1,3}(,\d{3}){1,})', r'==\1==', body_part)
        
        content = f"---{yaml_part}---{body_part}"
    
    # 5. Clean whitespace
    content = re.sub(r'\n{3,}', r'\n\n', content)

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(content)

def convert_pdf_to_md(input_dir, output_dir):
    """
    Sử dụng pdfmd để chuyển đổi và sau đó hậu xử lý sang chuẩn Obsidian.
    """
    pdfmd_path = r"C:\Users\HP\AppData\Roaming\Python\Python311\Scripts\pdfmd.exe"
    
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    
    if not output_path.exists():
        output_path.mkdir(parents=True)
        
    print(f"--- Bắt đầu quy trình Obsidian Standard (pdfmd) ---")
    
    pdf_files = list(input_path.glob("*.pdf"))
    
    if not pdf_files:
        print("Không tìm thấy tệp PDF nào.")
        return

    for pdf in pdf_files:
        print(f"Đang chuyển đổi: {pdf.name}...")
        try:
            # 1. Chuyển đổi PDF -> MD
            subprocess.run(
                [pdfmd_path, str(pdf), "-o", str(output_path)],
                capture_output=True,
                text=True,
                check=True
            )
            
            # 2. Hậu xử lý MD -> Obsidian Standard
            md_file = output_path / (pdf.stem + ".md")
            if md_file.exists():
                print(f"  => Đang chuẩn hóa Obsidian Skill cho: {md_file.name}...")
                post_process_markdown(md_file)
                print(f"  => Hoàn tất: {md_file.name}")
                
        except subprocess.CalledProcessError as e:
            print(f"  => Lỗi khi xử lý {pdf.name}: {e.stderr}")

if __name__ == "__main__":
    BASE_DIR = r"D:\DATA\MEMORY\Second memory\10.Inbox\12.Random Idea\SKILL_VANPHONG_VIETNAM"
    input_examples = os.path.join(BASE_DIR, "examples")
    output_tests = os.path.join(BASE_DIR, "tests")
    
    convert_pdf_to_md(input_examples, output_tests)
