# SKILL: PDF → Obsidian Markdown (MCP Tool)

## Mô tả
Skill này hướng dẫn Claude gọi MCP tool `convert_pdf_to_obsidian` để chuyển đổi
file PDF thành Obsidian-flavored Markdown chuẩn hóa, RAG-ready, và có WikiLinks.

---

## Khi nào kích hoạt skill này?

Kích hoạt khi người dùng:
- Nói "convert PDF", "chuyển PDF sang Markdown", "đọc PDF vào Obsidian"
- Cung cấp đường dẫn `.pdf` và muốn lưu vào vault
- Muốn phân tích hoặc tóm tắt nội dung từ file PDF cục bộ

---

## MCP Tool: `convert_pdf_to_obsidian`

**Server:** `PDF-to-Markdown-Obsidian`

### Tham số

| Tham số | Kiểu | Bắt buộc | Mô tả |
| --- | --- | --- | --- |
| `pdf_path` | string | ✅ | Đường dẫn tuyệt đối tới file `.pdf` nguồn |
| `output_folder` | string | ❌ | Thư mục output (mặc định: `SKILL_VANPHONG_VIETNAM\tests`) |

### Ví dụ gọi

```json
{
  "tool": "convert_pdf_to_obsidian",
  "pdf_path": "D:\\DATA\\MEMORY\\Second memory\\30 Areas\\DEV\\pdf-to-md\\Testfile\\Chương I.pdf",
  "output_folder": "D:\\DATA\\MEMORY\\Second memory\\30 Areas\\DEV\\pdf-to-md\\output"
}
```

---

## Quy trình thực hiện khi nhận yêu cầu

```
1. Xác nhận đường dẫn PDF hợp lệ (file tồn tại)
2. Hỏi output_folder nếu chưa có
   → Gợi ý mặc định: 30 Areas/DEV/pdf-to-md/output
3. Gọi tool convert_pdf_to_obsidian
4. Báo kết quả: đường dẫn .md + số dòng
5. Tóm tắt nội dung tài liệu nếu được yêu cầu
```

---

## Output được chuẩn hóa tự động

### YAML Frontmatter
```yaml
---
title: "Tên file"
date: YYYY-MM-DD
tags: [converted, pdf, automated]
---
```

### Page Markers (Obsidian comment — không gây noise LLM)
```
%% [Trang 3] %%
```

### WikiLinks tự động (idempotent — chạy nhiều lần không bị lồng)
- `BDL` → `[[Chương II - Bảng dữ liệu đấu thầu 29.3|BDL]]`
- `CDNT` → `[[Chương I - Chỉ dẫn nhà thầu 29.3|CDNT]]`
- `HSMT` → `[[Hồ sơ mời thầu|HSMT]]`

### Obsidian Callouts
- `▸ ** Đầu ra: **` → `> [!success] Đầu ra`
- `▸ ** Lưu ý: **` → `> [!info] Lưu ý`
- `▸ ** Không nên chứa: **` → `> [!warning] Không nên chứa`

### Highlight tài chính
- `1,500,000` → `==1,500,000==`

### Ngày tháng
- `29/3/2026` → `2026-03-29` (chỉ trong body, không chạm YAML)

---

## Thông tin kỹ thuật

| Thành phần | Đường dẫn |
| --- | --- |
| pdfmd binary | `C:\Users\HP\AppData\Roaming\Python\Python311\Scripts\pdfmd.exe` |
| MCP server | `...\SKILL_VANPHONG_VIETNAM\scripts\convert\pdf_mcp_server.py` |
| Batch script | `...\SKILL_VANPHONG_VIETNAM\scripts\convert\batch_pdf_to_md.py` |

> **Giới hạn:** Chỉ hỗ trợ PDF dạng text. PDF scan/ảnh cần OCR riêng.

---

## Xử lý lỗi phổ biến

| Lỗi | Nguyên nhân | Cách xử lý |
| --- | --- | --- |
| `PDF file not found` | Đường dẫn sai | Kiểm tra lại path, viết dạng raw string |
| `Process Error` | pdfmd crash | Kiểm tra PDF dạng text hay scan |
| `Markdown not generated` | Lỗi quyền ghi | Đổi output folder, kiểm tra permissions |
| WikiLink lồng nhau | Script chạy nhiều lần | Script đã có guard idempotent, chạy lại an toàn |

---

## Cài đặt Skill này vào Claude Global

Sau khi tạo xong, copy file này vào:

```
C:\Users\HP\.claude\skills\pdf-to-obsidian\SKILL.md
```

Hoặc dùng lệnh PowerShell:

```powershell
$src = "D:\DATA\MEMORY\Second memory\10.Inbox\12.Random Idea\SKILL_VANPHONG_VIETNAM\SKILL-pdf-to-obsidian.md"
$dst = "C:\Users\HP\.claude\skills\pdf-to-obsidian\SKILL.md"
New-Item -ItemType Directory -Force -Path (Split-Path $dst)
Copy-Item $src $dst
```
