# diagram-design

Vẽ sơ đồ cho báo cáo, thuyết minh: HTML/SVG tự bố cục (không auto-layout như mermaid), xuất SVG/PNG,
chèn thẳng vào file Word.

- 39 loại biểu đồ, 5 template dựng sẵn cho hồ sơ EPC: kiến trúc hệ thống, quy trình, tiến độ Gantt,
  cơ cấu tổ chức, luồng dữ liệu.
- Màu và font đã gắn theo bộ nhận diện IPC; sửa 4 biến trong `:root` của file HTML là đổi được.
- Xuất hình và ghép vào thuyết minh bằng một lệnh.

## Cài đặt

```bash
git clone https://github.com/hoanglong149/office-bidding-skills.git
cp -r office-bidding-skills/skills/diagram-design ~/.agents/skills/diagram-design
```

Cần Chrome hoặc Chromium trên máy (để render PNG) và `python-docx` cho bước chèn Word:

```bash
python3 -m pip install --user python-docx
```

## Dùng

```bash
cd ~/.agents/skills/diagram-design
bash scripts/vao-thuyet-minh.sh ./hinh "Thuyet minh.docx"          # chèn vào bài có sẵn
bash scripts/vao-thuyet-minh.sh ./hinh "Thuyet minh.docx" --new    # tạo bài mới
```

Chi tiết template, cách đổi màu, vị trí chèn hình: [README-EPC.md](README-EPC.md).
Nguồn gốc và trạng thái kiểm thử: [NOTICE.md](NOTICE.md).

## Tác giả

Hoàng Long - [github.com/hoanglong149](https://github.com/hoanglong149)
