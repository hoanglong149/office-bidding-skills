---
name: hsdt-forms
description: >
  MUST USE khi cần lập/chuẩn hoá Hồ sơ dự thầu (HSDT) hoặc biểu mẫu đấu thầu dạng Word/Excel:
  "làm HSDT", "biểu mẫu dự thầu", "bidding form", "đơn dự thầu", "bảng giá dự thầu",
  "BOQ điền giá", "danh mục hợp đồng tương tự", "nhân sự chủ chốt", "list nhà thầu phụ",
  "deviation form", "grand summary", "price schedule", "letterhead theo gói thầu".
  Cũng dùng khi cần chuẩn hoá định dạng .docx theo Nghị định 30/2020/NĐ-CP
  ("in ra không giống format", "lỗi font", "thiếu số trang", "lệch lề", "rớt dòng tiêu đề").
  Bộ đã sinh sẵn: `templates/` gồm 11 file Word + 11 workbook Excel (28 sheet, phủ 26 form
  danh mục & biểu giá) — copy dùng ngay, không cần cài Python.
  `examples/demo_project.py` sinh bộ demo hoàn chỉnh để đối chiếu.
  NOT for: soạn thảo công văn/tờ trình hành chính thường ngày (dùng template ND30 riêng),
  bóc khối lượng/BOQ kỹ thuật (dùng skill BOQ), hay đọc HSMT để tóm tắt điều khoản.
date: 2026-10-05
tags: [skill, hsdt, dau-thau, docx, xlsx, nd30, bidding-form, ipc]
---

# HSDT Forms — Bộ biểu mẫu Hồ sơ dự thầu (Word + Excel, chuẩn NĐ30)

> **Dùng skill này KHI**: lập HSDT mới, chuẩn hoá biểu mẫu thầu, chuyển list/biểu giá sang Excel,
> hoặc sửa lỗi "in ra không giống format" của .docx.
> **KHÔNG dùng KHI**: soạn công văn/tờ trình/quyết định hành chính (dùng `00 System/Templates/ND30-*.md`),
> bóc khối lượng kỹ thuật (dùng skill BOQ).

---

## 1. Nguyên tắc phân vai Word vs Excel (BẮT BUỘC)

| Loại biểu mẫu | Định dạng | Lý do |
|:--|:--|:--|
| Đơn từ, ủy quyền, thỏa thuận, bảo đảm, xác nhận, CV | **Word (.docx)** | Có tính văn bản, cần chữ ký/đóng dấu, bố cục nghị luận |
| **Danh mục / list / biểu giá / bảng kê số liệu** | **Excel (.xlsx)** | Nhiều dòng, cần công thức, lọc, tổng, in lặp tiêu đề |

**NEVER** làm danh mục dài trong Word (rớt trang, không tổng được).
**NEVER** làm đơn dự thầu trong Excel.

---

## 2. Chạy nhanh

### 2.0. Cách nhanh nhất — dùng template có sẵn (không cần Python)

```text
templates/                     ← 22 file đã sinh sẵn, đúng định dạng NĐ30
├── HSDT-F01-Letter-of-Bid-Technical.docx      … 11 file Word
└── HSDT-F20-F23-Bieu-gia.xlsx                 … 11 workbook Excel (28 sheet)
```

Copy `templates/*` ra thư mục gói thầu là dùng được ngay.
Muốn xem **bộ thành phẩm trông thế nào**: mở `demo_output/` hoặc xem ảnh trong `examples/preview/`.

### 2.1. Sinh lại / tuỳ biến bằng script

```bash
SK="$(dirname "$(readlink -f "$0")")"   # thư mục skill
bash "$SK/scripts/build_all.sh"                         # sinh lại toàn bộ Word + Excel
python3 "$SK/scripts/build_word.py"  [thư_mục_đích]     # chỉ Word (11 file)
python3 "$SK/scripts/build_excel.py" [thư_mục_đích]     # chỉ Excel (11 workbook)
python3 "$SK/scripts/fix_nd30.py" --check  <file.docx>  # kiểm tra định dạng 1 file
python3 "$SK/scripts/fix_nd30.py"          <file.docx>  # tự chuẩn hoá (tạo .bak)
python3 "$SK/scripts/fill_letterhead.py" <thư_mục> --bidder "..." --employer "..." \
        --place "Hà Nội" --doc-no "05/2026-HSDT" --date "ngày 05 tháng 10 năm 2026"
```

Mặc định ghi ra `./hsdt_forms` (hoặc `$HSDT_OUT`). Khi làm gói thầu cụ thể:
**copy ra thư mục gói thầu rồi mới sửa** — không sửa trực tiếp trong `templates/`.

### 2.2. Demo end-to-end

```bash
python3 "$SK/examples/demo_project.py" [thư_mục_output]
```

Lấy `templates/` → điền letterhead gói giả định → chèn dữ liệu thiết bị mẫu → xuất bộ hoàn chỉnh.
Dùng để học pipeline hoặc để kiểm tra môi trường mới cài.

---

## 3. Quy cách định dạng NĐ30 (đã áp sẵn — không tự ý đổi)

| Thông số | Giá trị |
|:--|:--|
| Khổ giấy | A4 210 × 297 mm |
| Lề | trên 20 · dưới 20 · **trái 30** · **phải 20** mm |
| Font | Times New Roman, đen |
| Cỡ chữ | nội dung **13 pt**; bảng **11 pt**; chú thích 9–10 pt |
| Dãn dòng | **1,2**; khoảng cách sau đoạn **6 pt** |
| Số trang | giữa lề dưới, field `PAGE`, 12 pt |
| Khối tiêu đề | **2 cột 6,0 / 10,0 cm**, co giãn ký tự **95 %** → không rớt dòng |
| Bảng | rộng **100 %** khổ chữ, `tblLayout=fixed` → Word không co cột |
| Excel in | A4 (ngang nếu bảng rộng), `fitToWidth=1`, **lặp dòng tiêu đề 1–6**, footer `Trang &P/&N` |

Chi tiết + bảng cỡ chữ từng thành phần: `references/nd30-format.md`.

### 3 lỗi khiến "in ra không giống format" (đã sửa, gặp lại thì xử lý ngay)
1. `docDefaults` dùng **font theme** → run không set rPr sẽ in ra Calibri ⇒ ép Times New Roman.
2. **Thiếu style `Normal`** + `pPrDefault` rỗng → Word tự chế (Calibri 11, dãn 1,08, after 8pt).
3. **Thiếu footer** ⇒ mất số trang; **bảng autofit** ⇒ bị co cột, chữ wrap loạn.

---

## 4. Letterhead điền theo từng gói thầu

Không hard-code tên đơn vị. Mỗi file có 3 chế độ:

| Mode | Dùng cho | Nội dung ô trái |
|:--|:--|:--|
| `BIDDER` (mặc định) | form do nhà thầu lập | `[TÊN NHÀ THẦU / LIÊN DANH NHÀ THẦU]` + `[Địa chỉ · Điện thoại · Email]` |
| `EMPLOYER` | Form 17(a) | `[TÊN CHỦ ĐẦU TƯ / BÊN MỜI THẦU]` |
| `NONE` | Attachment 1A/1B/2 | — |

Khi phát hành cho một gói cụ thể, thay: tên nhà thầu/CĐT, địa chỉ, `Số: …/…-HSDT`,
`[Địa danh]`, `ngày … tháng … năm …`.

---

## 5. Danh mục 37 biểu mẫu (36 form theo HSMT + Attachment)

**Word (11):** F01 Đơn dự thầu KT · F02 Ủy quyền · F03 Thỏa thuận liên danh · F04 Bảo đảm dự thầu ·
F05a Thông tin nhà thầu · F05b Thông tin thành viên LD · F07B CV nhân sự · F17a/b Xác nhận ·
F19a/b Đơn dự thầu tài chính.

**Excel (11 workbook / 28 sheet):** F06A/B/C · F07A+F07C · F08 · F09A+B · F10A1+10A2+10B+11 ·
F12+13+14+15+16 · F18A+B · ATT 1A+1B+2 · F20+21.1+21.2+21.3+22+23.

Bảng ánh xạ đầy đủ form ↔ file ↔ HSMT reference: `references/form-index.md`.

---

## 6. Quy trình khi có gói thầu mới (checklist)

1. **Đọc danh mục form trong HSMT** — số Form/Attachment có thể khác bộ chuẩn; phải map lại.
2. Copy bộ template ra thư mục gói thầu.
3. Điền letterhead theo gói (tên CĐT/Bên mời thầu + tên nhà thầu).
4. Điền dữ liệu: nhân sự, thiết bị, hợp đồng tương tự, tài chính, giá.
5. Kiểm tra: `python3 scripts/fix_nd30.py --check <file>` cho mọi .docx phát hành.
6. Đối chiếu checklist nộp thầu (Điều 11 HSMT) — đủ bản gốc/bản chụp/USB, ký + đóng dấu + giáp lai.
7. Xuất PDF để nộp; giữ nguyên bản .docx/.xlsx gốc để chỉnh sửa.

---

## 7. Cấu trúc repo

```text
hsdt-forms-skill/
├── SKILL.md
├── templates/                  # 22 file đã sinh sẵn — copy dùng ngay
├── demo_output/                # bộ demo thành phẩm (letterhead + dữ liệu mẫu)
├── examples/
│   ├── demo_project.py         # script demo end-to-end
│   └── preview/*.png           # ảnh render của bộ demo
├── scripts/
│   ├── build_word.py           # sinh 11 Word   ← nguồn chính
│   ├── build_excel.py          # sinh 11 Excel  ← nguồn chính
│   ├── fill_letterhead.py      # điền letterhead/ngày/số hồ sơ vào .docx + .xlsx
│   ├── fix_nd30.py             # kiểm tra & tự chuẩn hoá .docx
│   └── build_all.sh
└── references/{nd30-format.md, form-index.md}
```

Sửa nội dung biểu mẫu ⇒ sửa `scripts/build_word.py` / `build_excel.py` rồi chạy lại
`build_all.sh` + copy kết quả vào `templates/` (giữ 2 nơi khớp nhau).
