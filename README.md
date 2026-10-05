# hsdt-forms

Bộ **biểu mẫu Hồ sơ dự thầu (HSDT / Bidding Documents)** sinh tự động — Word cho văn bản, Excel cho danh mục & biểu giá — định dạng theo **Nghị định 30/2020/NĐ-CP** (Việt Nam).

Đóng gói dưới dạng **skill** dùng được cho agent (omp / Claude Code / Cursor) và dùng độc lập như script CLI.

> 11 biểu mẫu **Word** + 11 workbook **Excel** (28 sheet), phủ **36 form + 3 Attachment** theo cấu trúc
> `Chapter IV — Bidding Forms` của HSMT EPC tiêu chuẩn.

---

## Vì sao

Lập HSDT thường mất nhiều ngày vì 3 việc lặp lại:

1. **Định dạng sai** → in ra lệch: font lẫn Calibri, dãn dòng không đều, mất số trang, bảng bị co cột.
2. **Sai định dạng file** → danh mục dài làm trong Word thì rớt trang, không tổng được.
3. **Thiếu biểu mẫu** → sót form so với danh mục HSMT.

Repo này giải cả 3: sinh sẵn bộ biểu mẫu đúng chuẩn, đúng vai trò file, kèm công cụ **kiểm tra & tự chuẩn hoá** `.docx`.

---

## Cài đặt

### Yêu cầu
- Python 3.9+
- `pip install python-docx openpyxl`

### Dùng như skill (agent)
```bash
# omp: copy vào thư mục skill của project hoặc user
git clone https://github.com/<user>/hsdt-forms-skill.git
cp -r hsdt-forms-skill ~/.agents/skills/hsdt-forms        # user-level
# hoặc: cp -r hsdt-forms-skill <workspace>/.agents/skills/hsdt-forms   # project-level
```
Agent sẽ tự đọc `SKILL.md` khi gặp task về HSDT / biểu mẫu dự thầu / lỗi format `.docx`.

### Dùng như CLI
```bash
python3 scripts/build_word.py  [thư_mục_đích]     # 11 biểu mẫu Word
python3 scripts/build_excel.py [thư_mục_đích]     # 11 workbook Excel
bash    scripts/build_all.sh   [thư_mục_đích]     # cả hai + tự kiểm tra

python3 scripts/fix_nd30.py --check file.docx     # chỉ kiểm tra lỗi định dạng
python3 scripts/fix_nd30.py         file.docx     # tự sửa (tạo file.docx.bak)
```

Mặc định ghi ra `./hsdt_forms` (hoặc `$HSDT_OUT` nếu đặt biến môi trường).

---

## Nội dung bộ biểu mẫu

### Word (11) — văn bản có chữ ký/đóng dấu
`F01` Đơn dự thầu KT · `F02` Ủy quyền · `F03` Thỏa thuận liên danh · `F04` Bảo đảm dự thầu ·
`F05a` Thông tin nhà thầu · `F05b` Thông tin thành viên liên danh · `F07B` CV nhân sự chủ chốt ·
`F17a`/`F17b` Xác nhận · `F19a`/`F19b` Đơn dự thầu tài chính

### Excel (11 workbook / 28 sheet) — danh mục, list, biểu giá
`F06A/B/C` HĐ tương tự · `F07A`+`F07C` Nhân sự & trình độ · `F08` Thiết bị ·
`F09A`+`F09B` Non-fulfilment · `F10A1`+`10A2`+`10B`+`11` Tài chính · `F12`+`13`+`14`+`15`+`16` Danh mục ·
`F18A`+`18B` Deviation · `Attachment 1A`+`1B`+`2` Nhà sản xuất & vật tư ·
`F20` Grand Summary · `F21` Price Schedules 1–3 · `F22` Ưu đãi · `F23` Điều chỉnh giá

Chi tiết ánh xạ: [`references/form-index.md`](references/form-index.md)

---

## Quy cách định dạng (đã áp sẵn)

| Thông số | Giá trị |
|:--|:--|
| Khổ giấy | A4 210 × 297 mm |
| Lề | trên 20 · dưới 20 · **trái 30** · **phải 20** mm |
| Font | Times New Roman, màu đen |
| Cỡ chữ | nội dung **13 pt** · bảng **11 pt** · chú thích 9–10 pt |
| Dãn dòng | **1,2** · khoảng cách sau đoạn **6 pt** |
| Số trang | giữa lề dưới, field `PAGE`, 12 pt |
| Khối tiêu đề | 2 cột **6,0 / 10,0 cm**, co giãn ký tự **95 %** → không rớt dòng |
| Bảng (Word) | rộng 100 %, `tblLayout=fixed` → không bị co cột |
| In Excel | A4 (tự ngang nếu bảng rộng), `fitToWidth=1`, lặp tiêu đề dòng 1–6, footer `Trang &P/&N` |

Chi tiết: [`references/nd30-format.md`](references/nd30-format.md)

### Letterhead theo từng gói thầu
Không hard-code tên đơn vị. Mỗi file hỗ trợ 3 chế độ:

| Mode | Dùng cho | Ô trái |
|:--|:--|:--|
| `BIDDER` (mặc định) | form do nhà thầu lập | `[TÊN NHÀ THẦU / LIÊN DANH NHÀ THẦU]` |
| `EMPLOYER` | Form 17(a) | `[TÊN CHỦ ĐẦU TƯ / BÊN MỜI THẦU]` |
| `NONE` | Attachment 1A/1B/2 | — |

---

## 3 lỗi format `.docx` thường gặp (và cách `fix_nd30.py` xử lý)

| # | Triệu chứng khi in | Nguyên nhân trong XML | Cách sửa |
|:--:|:--|:--|:--|
| 1 | Chữ lẫn font (Calibri xen Times New Roman) | `docDefaults` dùng `w:asciiTheme` → run thiếu `rPr` lấy font theme | ép `Times New Roman` ở `docDefaults` + style `Normal` |
| 2 | Dãn dòng không đều, đoạn cách xa | `<w:pPrDefault/>` rỗng, thiếu style `Normal` | thêm `pPrDefault` (`after=120`, `line=288`, `auto`) + style `Normal` |
| 3 | Mất số trang; bảng co cột, tiêu đề rớt dòng | thiếu `footer1.xml` / `footerReference`; bảng `autofit` | thêm footer field `PAGE` + `tblW=5000pct`, `tblLayout=fixed` |

`fix_nd30.py` **validate XML bằng `ElementTree` trước khi ghi** và luôn tạo `.bak`.

---

## Cấu trúc repo

```text
hsdt-forms-skill/
├── SKILL.md                    # hướng dẫn cho agent (when to use / workflow / rules)
├── README.md
├── LICENSE                     # MIT
├── scripts/
│   ├── build_word.py           # sinh 11 biểu mẫu Word
│   ├── build_excel.py          # sinh 11 workbook Excel
│   ├── fix_nd30.py             # kiểm tra & chuẩn hoá .docx theo NĐ30
│   └── build_all.sh
└── references/
    ├── nd30-format.md          # quy chuẩn NĐ30 + 3 lỗi thường gặp
    └── form-index.md           # ánh xạ biểu mẫu ↔ file ↔ tham chiếu HSMT
```

---

## Ghi chú pháp lý / dữ liệu

- Không chứa dữ liệu gói thầu cụ thể (tên chủ đầu tư, giá, nhân sự). Toàn bộ trường dữ liệu là placeholder `[……]`.
- Danh mục biểu mẫu dựa trên cấu trúc chuẩn của HSMT EPC; **luôn đối chiếu lại danh mục trong HSMT của gói thầu thực tế** vì số Form có thể khác nhau.
- Định dạng tham chiếu **Nghị định 30/2020/NĐ-CP**.

## License

MIT — xem [LICENSE](LICENSE).
