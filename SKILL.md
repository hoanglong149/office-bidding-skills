---
name: hsdt-forms
description: >
  MUST USE khi làm Hồ sơ dự thầu (HSDT / E-HSDT) hoặc cần biểu mẫu đấu thầu theo
  Thông tư 79/2025/TT-BTC: "làm HSDT", "biểu mẫu dự thầu", "Mẫu số 02", "đơn dự thầu",
  "thỏa thuận liên danh", "bảo lãnh dự thầu", "hợp đồng tương tự", "nhân sự chủ chốt",
  "thiết bị thi công chủ yếu", "nguồn lực tài chính", "nhà thầu phụ đặc biệt",
  "bảng tổng hợp giá dự thầu", "bảng kê công nhật", "hàng hóa hưởng ưu đãi",
  "E-HSMT mẫu", "Mẫu số 10A", "webform E-HSDT", "tách biểu mẫu từ HSMT".
  Cũng dùng khi cần chuẩn hoá khổ giấy / lề / font cho file .docx đấu thầu.
  NOT for: soạn công văn – tờ trình – quyết định hành chính (thể thức NĐ30, dùng template riêng);
  bóc khối lượng kỹ thuật / lập BOQ (dùng skill BOQ).
date: 2026-10-05
tags: [skill, hsdt, dau-thau, tt79, muasamcong, docx, xlsx, e-hsmt]
---

# HSDT Forms — Biểu mẫu Hồ sơ dự thầu theo **TT 79/2025/TT-BTC**

> **Dùng khi**: lập E-HSDT gói thầu EPC / xây lắp / hàng hóa / EP / EC / PC / tư vấn,
> cần đúng biểu mẫu và đúng định dạng do Bộ Tài chính ban hành.
> **KHÔNG dùng khi**: soạn văn bản hành chính (công văn, tờ trình…) — thể thức NĐ30 khác hẳn.

---

## 0. Chuẩn áp dụng — KHÔNG nhầm với NĐ30

| | **HSDT (skill này)** | Văn bản hành chính |
|:--|:--|:--|
| Chuẩn | **TT 79/2025/TT-BTC** ngày 04/8/2025 | NĐ 30/2020/NĐ-CP |
| Mẫu | `Mẫu số 02`, `Mẫu số 10A`… (Chương IV E-HSMT) | Quốc hiệu, tên loại văn bản, ký hiệu |
| Đầu trang | **`Mẫu số X (Webform trên Hệ thống)`** — phải, đậm, 14 pt | Khối tiêu đề 2 cột + `Số: …/…` |
| Nơi nhận / địa danh–ngày | **KHÔNG có** | Có |
| Khổ giấy / lề / font | A4 · 20-20-**30**-20 mm · Times New Roman **14 pt** | giống |

**NEVER** gắn khối tiêu đề kiểu công văn (Số: …/…-HSDT, "Hà Nội, ngày …") vào biểu mẫu HSDT,
**NEVER** đặt "Nơi nhận", "Kính gửi" theo thể thức hành chính vào các mẫu 05–13.

### Căn cứ pháp lý
- Luật Đấu thầu 22/2023/QH15 (sửa đổi: 57/2024/QH15, 90/2025/QH15)
- NĐ 214/2025/NĐ-CP (04/8/2025) · NĐ 349/2026/NĐ-CP (sửa đổi)
- **TT 79/2025/TT-BTC** (04/8/2025) — hướng dẫn cung cấp, đăng tải thông tin đấu thầu và
  **mẫu hồ sơ đấu thầu trên Hệ thống mạng đấu thầu quốc gia** (thay TT 22/2024/TT-BKHĐT)
- Hệ thống: https://muasamcong.mof.gov.vn

---

## 1. Nguyên tắc: lấy biểu mẫu TỪ E-HSMT mẫu, không tự chế

Biểu mẫu HSDT nằm trong **Chương IV — Biểu mẫu mời thầu và dự thầu** của E-HSMT mẫu
(Mẫu số 3A/4A/5A/6A/7A/8A/9A/10A/10B tuỳ loại gói thầu). Bản quy định số mẫu, tên mẫu,
cột dữ liệu và ghi chú chân trang — **không được tự sáng tác**.

Skill này **cắt nguyên trạng** từng biểu mẫu ra file riêng — giữ 100 % định dạng gốc
(khổ, lề, font, gộp ô, ghi chú). Không dựng lại mẫu bằng tay.

```
sources/                     ← E-HSMT mẫu chính thức (TT79)
├── TT79-M10A-E-HSMT-EPC-01-tui.docx
├── TT79-M3A-E-HSMT-Xay-lap-01-tui.docx
├── TT79-M4A-E-HSMT-Hang-hoa-01-tui.docx
├── TT79-M8A-E-HSMT-EC-01-tui.docx
└── TT79-M9A-E-HSMT-PC-01-tui.docx

templates/                   ← 5 bộ đã tách sẵn (copy là dùng ngay)
├── EPC-10A/       32 Word + 30 Excel   gói EPC
├── Xay-lap-3A/    25 Word + 22 Excel   gói xây lắp
├── Hang-hoa-4A/   32 Word + 29 Excel   gói mua sắm hàng hóa
├── EC-8A/         26 Word + 23 Excel   gói EC
├── PC-9A/         35 Word + 32 Excel   gói PC
└── DANH-MUC-TONG-HOP.md
Mỗi bộ: `TT79-M02.docx` (Word: in/ký/scan) · `TT79-M06A.xlsx` (Excel: bảng nhiều dòng) · `DANH-MUC-BIEU-MAU.md`
```

---

## 2. Chạy nhanh

### 2.0. Dùng ngay bộ có sẵn (không cần Python)
Chọn đúng bộ theo loại gói thầu trong `templates/` rồi copy ra thư mục làm việc:

| Loại gói | Bộ |
|:--|:--|
| EPC (tư vấn + hàng hóa + xây lắp) | `templates/EPC-10A/` |
| Xây lắp | `templates/Xay-lap-3A/` |
| Mua sắm hàng hóa | `templates/Hang-hoa-4A/` |
| EC (tư vấn + xây lắp) | `templates/EC-8A/` |
| PC (hàng hóa + xây lắp) | `templates/PC-9A/` |

Gói EPC hai túi hồ sơ hoặc gói thực tế: xem §2.1.

### 2.1. Tách bộ biểu mẫu cho loại gói thầu khác
```bash
python3 scripts/tt79_extract.py sources/TT79-M3A-E-HSMT-Xay-lap-01-tui.docx  out/xay-lap
python3 scripts/tt79_extract.py sources/TT79-M4A-E-HSMT-Hang-hoa-01-tui.docx  out/hang-hoa

# chỉ lấy một số mẫu (tiền tố mã mẫu)
python3 scripts/tt79_extract.py sources/TT79-M10A-E-HSMT-EPC-01-tui.docx out --only 02,03,06,11
python3 scripts/tt79_extract.py <E-HSMT.docx> out --no-excel        # chỉ Word
```

Ra: `TT79-M<mã>-<tên>.docx` + `TT79-M<mã>-<tên>.xlsx` (mỗi bảng → 1 sheet) + `DANH-MUC-BIEU-MAU.md`.

### 2.2. Chuẩn hoá khổ giấy / lề / font cho file .docx
```bash
python3 scripts/fix_page_setup.py --check file.docx     # kiểm tra
python3 scripts/fix_page_setup.py         file.docx     # sửa (tạo .bak)
```

---

## 3. Định dạng chuẩn của biểu mẫu (đã đúng trong bộ đã tách)

| Thông số | Giá trị |
|:--|:--|
| Khổ giấy | A4 210 × 297 mm |
| Lề | trên 20 · dưới 20 · **trái 30** · **phải 20** mm |
| Font | Times New Roman, chữ đen |
| Cỡ chữ | **14 pt** (đúng như E-HSMT mẫu) |
| Đầu biểu mẫu | `Mẫu số X (Webform trên Hệ thống)` — căn phải, đậm, 14 pt |
| Tên biểu mẫu | CHỮ HOA — căn giữa, đậm |
| Cuối biểu mẫu | `Ghi chú:` + các chú thích (1), (2)… đúng nguyên bản |
| Số trang | lề dưới |

Ba lỗi hay gặp khi tự dựng lại mẫu (dẫn tới "in ra không giống"):
1. `docDefaults` theo **font theme** → run thiếu `rPr` ra Calibri.
2. Thiếu style `Normal` / `pPrDefault` → Word tự chế cỡ chữ, dãn dòng.
3. Thiếu footer → mất số trang; bảng để `autofit` → co cột, rớt dòng tiêu đề.

Dùng `fix_page_setup.py` để xử lý; **không sửa tay từng file**.

---

## 4. Phân vai Word / Excel

| Loại | Định dạng | Lý do |
|:--|:--|:--|
| Đơn dự thầu, thỏa thuận liên danh, bảo lãnh, cam kết | **Word** | Văn bản cần ký/đóng dấu, scan |
| Danh mục, kê khai, bảng giá (06A, 06D, 08A–C, 10A–B, 11.x, 12A–C, 13A–C) | **Excel** | Nhiều dòng, cần công thức, lọc, tổng, in lặp tiêu đề |

Trên Hệ thống các mẫu này là **webform** — file Excel ở đây dùng để **chuẩn bị dữ liệu offline**
rồi nhập lên; file Word dùng để in/ký/scan khi HSMT yêu cầu bản giấy.

---

## 5. Danh mục biểu mẫu EPC (32 mẫu — Chương IV, Mẫu số 10A)

| Nhóm | Mẫu |
|:--|:--|
| Đơn & pháp lý | 02 Đơn dự thầu · 03 Thỏa thuận liên danh · 04A/04B Bảo lãnh dự thầu *(scan)* |
| Kinh nghiệm | 05A1 HĐ EPC/EC/EP/PC tương tự · 05A2 HĐ cung cấp hàng hóa (P) · 05A3 HĐ xây lắp (C) · 05A4 HĐ tư vấn (E) · 05B Năng lực sản xuất hàng hóa |
| Nhân sự & thiết bị | 06A Đề xuất nhân sự chủ chốt · 06B Lý lịch chuyên môn · 06C Kinh nghiệm chuyên môn · 06D Thiết bị thi công chủ yếu |
| Lịch sử & tài chính | 07 HĐ không hoàn thành do lỗi nhà thầu · 08A Tình hình tài chính · 08B Nguồn lực tài chính · 08C NLTC hàng tháng |
| Nhà thầu phụ | 09A Phạm vi công việc dùng NTP · 09B NTP đặc biệt · 09C Công ty con, thành viên |
| Tiến độ & hàng hóa | 10A Bảng tiến độ thực hiện · 10B Đề xuất về giá hàng hóa |
| Giá dự thầu | 11.1A–D (giá chưa gồm thuế) · 11.2A–D (giá đã gồm thuế, phí, lệ phí) — theo 4 loại hợp đồng |
| Chi phí khác | 12A Bảng kê công nhật · 12B Khoản tạm tính · 12C Số liệu điều chỉnh |
| Ưu đãi | 13A Hàng hóa hưởng ưu đãi · 13B/13C Chi phí sản xuất trong nước |

Bảng chi tiết + tên file: `templates/<bộ>/DANH-MUC-BIEU-MAU.md`; danh mục 5 bộ: `templates/DANH-MUC-TONG-HOP.md`.

Số mẫu khác nhau giữa các loại gói (EPC 32 · xây lắp 25 · hàng hóa 32 · EC 26 · PC 35) — gói xây lắp không có mẫu 05A2/05A3, gói hàng hóa không có 05A1/05A3…

---

## 6. Checklist khi làm HSDT

1. Mở **Chương IV của E-HSMT thực tế** — đối chiếu danh mục biểu mẫu (số Mẫu, mục nào Webform,
   mục nào Scan, mục nào phải nộp bản giấy). Số mẫu có thể khác bộ mẫu chuẩn.
2. Tách bộ mẫu đúng loại gói: `tt79_extract.py sources/<mẫu đúng>.docx out/`.
3. Chuẩn bị dữ liệu offline (Excel) cho các bảng 06–13.
4. Nhập lên Hệ thống bằng **chữ ký số**; phần Scan thì in – ký – đóng dấu – scan – đính kèm.
5. Kiểm tra: đúng mẫu, đúng cột, không thêm/bớt dòng tiêu đề, không còn chỗ trống bắt buộc.
6. Lưu bản gốc `.docx/.xlsx` để chỉnh sửa; xuất PDF chỉ khi cần đối chiếu.

---

## 7. Cấu trúc skill

```text
hsdt-forms/
├── SKILL.md
├── sources/                 # E-HSMT mẫu chính thức (TT79/2025/TT-BTC)
├── templates/               # 5 bộ biểu mẫu đã tách (EPC, xây lắp, hàng hóa, EC, PC)
├── examples/demo_project.py # điền dữ liệu mẫu vào vài bảng để xem trước
├── scripts/
│   ├── tt79_extract.py      # tách biểu mẫu từ E-HSMT mẫu  ← công cụ chính
│   └── fix_page_setup.py    # chuẩn hoá khổ giấy/lề/font .docx
└── references/
    ├── tt79-forms.md        # danh mục biểu mẫu + căn cứ pháp lý
    └── tt79-layout.md       # đặc tả định dạng + 3 lỗi thường gặp
```
