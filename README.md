# hsdt-forms

Bộ **biểu mẫu Hồ sơ dự thầu (HSDT / E-HSDT)** đúng chuẩn **Thông tư 79/2025/TT-BTC** — tách nguyên trạng
từ E-HSMT mẫu do Bộ Tài chính ban hành, kèm công cụ tách biểu mẫu cho **mọi loại gói thầu**.

> ⚠️ **Không phải biểu mẫu NĐ 30**. HSDT không dùng thể thức văn bản hành chính:
> không có `Số: …/…-HSDT`, không có `Hà Nội, ngày …`, không có khối tiêu đề 2 cột.
> Đầu biểu mẫu là `Mẫu số 02 (webform trên Hệ thống)` — căn phải, đậm, 14 pt.

---

## Bộ có sẵn — dùng ngay, không cần cài gì

```text
templates/EPC-10A/          ← 32 biểu mẫu HSDT gói EPC (Chương IV, Mẫu số 10A)
├── TT79-M02.docx             Đơn dự thầu
├── TT79-M03.docx             Thỏa thuận liên danh
├── TT79-M04A.docx / 04B.docx Bảo lãnh dự thầu (độc lập / liên danh)
├── TT79-M05A1..05B.docx      Hợp đồng tương tự · Năng lực sản xuất
├── TT79-M06A..06D.docx/.xlsx Nhân sự chủ chốt · Lý lịch · Kinh nghiệm · Thiết bị thi công
├── TT79-M07..09C.docx/.xlsx  Lịch sử HĐ · Tài chính · Nhà thầu phụ
├── TT79-M10A..13C.docx/.xlsx Tiến độ · Giá hàng hóa · Bảng giá dự thầu · Công nhật · Ưu đãi
└── DANH-MUC-BIEU-MAU.md      danh mục 32 mẫu

sources/                    ← E-HSMT mẫu chính thức (EPC · xây lắp · hàng hóa · EC · PC)
demo_output/                ← bộ demo: đã điền dữ liệu mẫu vào 06A / 06D / 10B
examples/preview/           ← ảnh render của Mẫu 02, 03, 06A, 11.1A
```

| Mẫu 02 — Đơn dự thầu | Mẫu 03 — Thỏa thuận liên danh |
|:--|:--|
| ![Mẫu 02](examples/preview/TT79-M02.docx.png) | ![Mẫu 03](examples/preview/TT79-M03.docx.png) |

| Mẫu 06A — Nhân sự chủ chốt (Excel) | Mẫu 11.1A — Bảng tổng hợp giá dự thầu |
|:--|:--|
| ![Mẫu 06A](examples/preview/TT79-M06A.xlsx.png) | ![Mẫu 11.1A](examples/preview/TT79-M11.1A.xlsx.png) |

---

## Dùng cho loại gói thầu khác

```bash
# EPC đã có sẵn; các loại khác tách từ nguồn trong sources/
python3 scripts/tt79_extract.py sources/TT79-M3A-E-HSMT-Xay-lap-01-tui.docx  out/xay-lap
python3 scripts/tt79_extract.py sources/TT79-M4A-E-HSMT-Hang-hoa-01-tui.docx out/hang-hoa
python3 scripts/tt79_extract.py sources/TT79-M8A-E-HSMT-EC-01-tui.docx       out/ec
python3 scripts/tt79_extract.py sources/TT79-M9A-E-HSMT-PC-01-tui.docx       out/pc

# hoặc từ E-HSMT thực tế của gói thầu (file Chương IV)
python3 scripts/tt79_extract.py "E-HSMT gói thầu X.docx" out --only 02,03,06,11
```

Ra: `TT79-M<mã>.docx` + `TT79-M<mã>.xlsx` (mỗi bảng → 1 sheet) + `DANH-MUC-BIEU-MAU.md`.
Tách 32 biểu mẫu mất **< 1 giây** (giữ nguyên 100 % định dạng gốc).

---

## Cài đặt

- **Chỉ dùng biểu mẫu** (`templates/`): không cần gì.
- **Chạy script**: Python 3.9+ và `pip install -r requirements.txt` (`lxml`, `python-docx`, `openpyxl`).

### Dùng như skill cho agent
```bash
git clone https://github.com/hoanglong149/hsdt-forms-skill.git
cp -r hsdt-forms-skill ~/.agents/skills/hsdt-forms        # user-level
# hoặc: cp -r hsdt-forms-skill <workspace>/.agents/skills/hsdt-forms
```
Agent tự đọc `SKILL.md` khi gặp task về HSDT / biểu mẫu dự thầu / TT79.

### Lệnh khác
```bash
python3 scripts/fix_page_setup.py --check file.docx   # kiểm khổ giấy / lề / font / số trang
python3 scripts/fix_page_setup.py         file.docx   # sửa (tạo .bak)
python3 examples/demo_project.py demo_output          # demo: chép mẫu + điền dữ liệu mẫu
```

---

## Vì sao tách nguyên trạng

| Cách | Rủi ro |
|:--|:--|
| Dựng lại mẫu bằng tay | thiếu cột, sai cỡ chữ 14 pt, mất ghi chú `(1) (2)`, sai số mẫu |
| **Tách từ E-HSMT mẫu** | chỉ xoá phần ngoài biểu mẫu — giữ nguyên `styles.xml`, footer, gộp ô, ký hiệu chú thích |

Script giữ lại `<w:sectPr>` (khổ A4, lề 20/20/30/20) — thiếu nó Word dùng khổ Letter, lề 1 inch.

Chi tiết: [`references/tt79-layout.md`](references/tt79-layout.md) · danh mục & căn cứ: [`references/tt79-forms.md`](references/tt79-forms.md)

---

## Cấu trúc repo

```text
hsdt-forms-skill/
├── SKILL.md                    # hướng dẫn cho agent
├── sources/                    # E-HSMT mẫu chính thức TT79 (EPC, xây lắp, hàng hóa, EC, PC)
├── templates/EPC-10A/          # 32 biểu mẫu HSDT EPC đã tách (Word + Excel)
├── demo_output/                # bộ demo có dữ liệu mẫu
├── examples/
│   ├── demo_project.py
│   └── preview/*.png
├── scripts/
│   ├── tt79_extract.py         # tách biểu mẫu từ E-HSMT mẫu  ← công cụ chính
│   └── fix_page_setup.py       # chuẩn hoá khổ giấy / lề / font / số trang
└── references/
    ├── tt79-forms.md           # danh mục 32 mẫu + căn cứ pháp lý
    └── tt79-layout.md          # đặc tả định dạng + 3 lỗi thường gặp
```

## Ghi chú pháp lý

- `sources/` chứa phụ lục E-HSMT mẫu do Bộ Tài chính ban hành kèm TT 79/2025/TT-BTC (văn bản công khai).
- Không chứa dữ liệu gói thầu cụ thể; toàn bộ trường dữ liệu là placeholder `____ [hướng dẫn]` theo nguyên bản.
- Luôn đối chiếu **Chương IV của E-HSMT thực tế** — số mẫu và tên mẫu có thể khác bộ mẫu chuẩn.

## License

MIT — xem [LICENSE](LICENSE).
