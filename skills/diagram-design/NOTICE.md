# Nguồn gốc và giấy phép

- Skill gốc: [`cathrynlavery/diagram-design`](https://github.com/cathrynlavery/diagram-design) — giấy phép MIT, bản 2.6.
- Bản trong repo này: port về dùng trực tiếp (không qua marketplace plugin), thêm template EPC
  (`templates-epc/`) và script ghép hình vào Word (`scripts/figdoc.py`, `scripts/vao-thuyet-minh.sh`).
- Giữ nguyên giấy phép MIT của bản gốc. Phần thêm của repo này cũng theo MIT.

## Trạng thái

Đã kiểm: render 5 template EPC và 4 template in bằng Chrome headless, chèn vào `.docx` (kể cả hình
đặt trên trang ngang), `scripts/self_check.py` báo OK trên các template. Chạy thử trên Chương 3 thuyết
minh Off-Gas: 3 hình, chữ 7,2-10,3 pt khi in rộng 16 cm.

Chưa kiểm: dùng thật cho một thuyết minh nộp thầu. Khi dùng lần đầu nên xem lại cỡ chữ, tỉ lệ hình
và màu khi in đen trắng.
