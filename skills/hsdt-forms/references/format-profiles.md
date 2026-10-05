# Hồ sơ định dạng (format profile) cho biểu mẫu HSDT

Bộ biểu mẫu trong `templates/` **giữ nguyên định dạng E-HSMT mẫu** (profile `keep`).
Khi làm một gói thầu cụ thể, đóng dấu letterhead và (tuỳ chọn) áp profile định dạng của gói
bằng `scripts/apply_project_format.py`.

---

## 1. `keep` — giữ nguyên bản TT79 (mặc định)

| Thông số | Giá trị |
|:--|:--|
| Khổ giấy | A4 210 × 297 mm (`pgSz 11907 × 16839`) |
| Lề | trên/dưới 20 mm (`1134`) · trái 30 mm (`1701`) · phải 20 mm (`1134`) |
| Font nội dung | theo E-HSMT mẫu (Times New Roman, 14 pt trên phần lớn run) |
| Số trang | field `PAGE` trong `word/header1.xml` |
| Bảng | giữ nguyên style/gộp ô của bản gốc |

Dùng khi nộp thầu trên Hệ thống (E-HSDT) và HSMT yêu cầu đúng bản mẫu.

---

## 2. `intl-epc` — theo HSDT gói quốc tế (EDF/EPC nước ngoài)

Trích từ file HSDT thật của gói EPC Phả Lại (Nhà thầu liên danh TUNA-XINLONG-IPC).

| Thông số | Giá trị | Nguồn |
|:--|:--|:--|
| Khổ giấy | **US Letter 216 × 279 mm** (`pgSz 12240 × 15840`) | file thật |
| Lề | trên/dưới 20 mm · trái 30 mm · phải 20 mm | `pgMar 1134/1134/1134/1701` |
| Khoảng cách header / footer | 11,2 mm / 14,9 mm (`635` / `845`) | file thật |
| Chữ trong bảng | **Times New Roman 13 pt** (`sz=26`) | file thật |
| Hàng tiêu đề bảng | in đậm, canh giữa | file thật |
| Ô | `vAlign=center`, canh lề giữa cho hàng đầu, viền `single sz=4` (0,5 pt) | file thật |
| Bảng | `tblLayout=fixed`, `tblW` giữ nguyên của mẫu | file thật |

### Letterhead (chèn vào `word/header1.xml`)

Bảng **không viền**, 3 cột `5670 | 2381 | 1701` dxa, `trHeight 361`:

| Cột | Nội dung |
|:--|:--|
| 1 | `Project : <tên dự án>` (nhãn in nghiêng + gạch chân) — xuống dòng — `Bidding Package : <tên gói thầu>` |
| 2 | logo (tuỳ chọn) |
| 3 | logo (tuỳ chọn) |

Nội dung cũ trong header (field `PAGE`) được **giữ lại** bên dưới letterhead; giá trị cache của
field được đưa về `1` để bản xem trước không hiện số trang của HSMT gốc.

Điểm bắt buộc: file phải **tham chiếu** header trong `w:sectPr`
(`<w:headerReference w:type="default" r:id="…"/>`). Nhiều file tách từ E-HSMT có part
`word/header1.xml` nhưng không có reference → script tự thêm.

---

## 3. Lệnh

```bash
python3 scripts/apply_project_format.py templates/EPC-10A \
  --project "<tên dự án>" \
  --package "<tên gói thầu>" \
  --logo-right assets/logo-ipc-ec.png \
  --profile intl-epc \
  --out out/EPC-PL --suffix=-PL
```

Không truyền `--out` thì sửa tại chỗ và tạo `.bak` bên cạnh.

Sau khi áp dụng, kiểm tra lại:

```bash
python3 scripts/fix_page_setup.py --check out/EPC-PL/*.docx
```

Checker tự nhận biết biểu mẫu HSDT (`Mẫu số …`): số trang có thể nằm ở header **hoặc** footer,
không bắt buộc `pPrDefault` dày như văn bản hành chính.

---

## 4. Khuôn mẫu hoá cho gói khác

1. Mở **Chương IV** của E-HSMT thực tế, xem đầu biểu mẫu (`Mẫu số … (Webform trên Hệ thống)`),
   khổ giấy, lề, cỡ chữ trong bảng, và vị trí số trang.
2. Ghi lại thành một profile mới trong file này (bảng như §2).
3. Nếu chỉ khác letterhead → dùng ngay `apply_project_format.py --profile keep`.
4. Nếu khác cả định dạng bảng/trang → bổ sung một nhánh trong `apply_profile()` của script
   (mỗi profile là một khối rõ ràng, không trộn điều kiện).
