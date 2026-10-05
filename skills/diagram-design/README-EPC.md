# Diagram Design — Hermes port (EPC workflow)

Skill gốc: `cathrynlavery/diagram-design` (MIT, 2.6) — 39 loại biểu đồ editorial HTML/SVG.
Port về Hermes ngày 2026-08-23. Không chạy được plugin Claude Code marketplace, nên bóc skill ra dùng trực tiếp.

## Cách dùng

Người dùng chỉ cần bảo: "vẽ sơ đồ [loại] về [nội dung]". Skill sẽ:
1. Load skill này (`skill_view diagram-design`).
2. Đọc reference tương ứng trong `references/type-*.md`.
3. Sinh file HTML vào `/tmp` (hoặc thư mục dự án).
4. Mở browser → screenshot → gửi ảnh vào Telegram / ghép vào DOCX báo cáo thầu.

## 5 template EPC sẵn (templates-epc/)

| File | Dùng cho |
|------|----------|
| `template-architecture.html` | Sơ đồ hệ thống xử lý khí thải (FGD/SCR/ESP), mặt bằng nhà máy |
| `template-process.html` | Quy trình thi công / vận hành (multi-actor) |
| `template-gantt.html` | Tiến độ thi công (tasks + phases trên timeline) |
| `template-org-chart.html` | Cơ cấu tổ chức dự án, phân công |
| `template-data-flow.html` | Sơ đồ công nghệ / luân chuyển vật liệu |

## Tùy chỉnh màu thương hiệu IPC

Mở file HTML, sửa 4 biến trong `:root` (dòng ~10-14):

```
--color-paper:   #f5f5f5;   /* nền */
--color-ink:     #2d3142;   /* chữ chính */
--color-muted:   #4f5d75;   /* chữ phụ */
--color-accent:  #eb6c36;   /* điểm nhấn (cam) — chỉ dùng 1-2 node */
```

Anh cho Em mã màu brand IPC (hoặc URL website công ty), Skill sẽ set mặc định luôn vào `.diagram-design` profile.

## Không dùng khi nào

- Sơ đồ unicode nhanh → viết text.
- Danh sách / bảng đơn giản → table/bullets.
- "Một hình" vô nghĩa → viết câu.

## Lưu ý

- Static HTML là mặc định (không JS, mở trình duyệt là xem được).
- Mỗi biểu đồ target density 4/10 — trên 9 node thì tách 2 sơ đồ.
- Accent (coral/cam) chỉ dành 1-2 node quan trọng nhất.

## Ghép hình vào thuyết minh Word (một lệnh)

Mỗi sơ đồ là một file `.html` (từ `templates-epc/` hoặc do skill sinh). Đặt tất cả vào một thư mục,
rồi:

```bash
bash scripts/vao-thuyet-minh.sh ./hinh "Thuyet minh.docx"          # chèn vào bài có sẵn
bash scripts/vao-thuyet-minh.sh ./hinh "Thuyet minh.docx" --new    # tạo bài mới
```

Script làm hai việc:

1. Tách nút `<svg>` của sơ đồ ra file `.svg` (vector, giữ `<title>`/`<desc>`).
2. Render đúng khung sơ đồ bằng Chrome headless, hệ số 3x, chỉ lấy phần sơ đồ (bỏ tiêu đề trang,
   nền chấm, thẻ tóm tắt), rồi chèn vào Word: ảnh canh giữa rộng 16 cm, chú thích
   `Hình N. <tiêu đề>` in nghiêng 12 pt bên dưới. Số thứ tự đếm tiếp các chú thích đã có trong bài.

Vị trí chèn: đặt dòng đánh dấu `[[FIG:tên-file-không-đuôi]]` trong bài, ví dụ `[[FIG:fig-01-architecture]]`.
Không có đánh dấu thì hình được thêm vào cuối bài.

Tùy chọn: `--width 16` (cm), `--scale 3`, `--white` (nền trắng, mặc định trong suốt),
`--png-only` (chỉ render), `--chrome <đường-dẫn>`.

Yêu cầu: Chrome/Chromium trên máy, và `python-docx` (`python3 -m pip install python-docx`).

Lưu ý: script chọn **khối `<svg>` lớn nhất** trong file, không phải thẻ đầu tiên - nhiều template có
icon nhỏ đứng trước sơ đồ.
