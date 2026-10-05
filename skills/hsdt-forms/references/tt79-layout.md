# Đặc tả định dạng biểu mẫu HSDT (TT 79/2025/TT-BTC)

## 1. Định dạng trang & chữ

| Thông số | Giá trị | Nguồn |
|:--|:--|:--|
| Khổ giấy | A4 210 × 297 mm (`w:pgSz w:w="11907" w:h="16839"`) | E-HSMT mẫu |
| Lề trên / dưới | 20 mm (`w:top="1134" w:bottom="1134"`) | E-HSMT mẫu |
| Lề trái / phải | **30 mm / 20 mm** (`w:left="1701" w:right="1134"`) | E-HSMT mẫu |
| Font | Times New Roman | E-HSMT mẫu |
| Cỡ chữ nội dung | **14 pt** | E-HSMT mẫu |
| Cỡ chữ bảng | 11–12 pt tuỳ bảng | E-HSMT mẫu |
| Số trang | footer, lề dưới | E-HSMT mẫu |
| Đánh số bắt đầu | `w:pgNumType w:start="1"` | — |

> Đây cũng là dải định dạng của NĐ 30/2020/NĐ-CP cho văn bản hành chính — nhưng **thể thức trình bày
> hoàn toàn khác** (xem §2).

## 2. Khác biệt cốt lõi với văn bản hành chính (NĐ 30)

| Thành phần | Biểu mẫu HSDT (TT79) | Văn bản hành chính (NĐ30) |
|:--|:--|:--|
| Dòng đầu | `Mẫu số 02 (webform trên Hệ thống)` — căn **phải**, đậm, 14 pt | Quốc hiệu + tiêu ngữ (căn giữa) |
| Khối tiêu đề 2 cột | **không có** | tên cơ quan (trái) / quốc hiệu (phải) |
| Số văn bản | **không có** | `Số: …/…-XXX` |
| Địa danh – ngày | **không có** | `Hà Nội, ngày … tháng … năm …` |
| Tên biểu mẫu | CHỮ HOA, căn giữa, đậm, 14 pt | tên loại văn bản (căn giữa, đậm) |
| Cuối văn bản | `Ghi chú:` + chú thích (1), (2)… | Nơi nhận · chữ ký · dấu |
| Chữ ký | chữ ký số trên Hệ thống, hoặc ký/đóng dấu cho bản scan | ký + đóng dấu + chức danh |

**Sai thường gặp**: dán khối `Số: …/…-HSDT` và `Hà Nội, ngày …` (kiểu công văn) lên đầu biểu mẫu HSDT.
Không được — HSDT không phải văn bản hành chính.

## 3. Ba lỗi khiến "in ra không giống mẫu"

| # | Triệu chứng | Nguyên nhân trong XML | Cách xử lý |
|:--:|:--|:--|:--|
| 1 | Chữ lẫn font (Calibri xen Times New Roman) | `docDefaults` dùng `w:asciiTheme`; run thiếu `rPr` lấy font theme | ép Times New Roman ở `docDefaults` + style `Normal` |
| 2 | Dãn dòng không đều, đoạn cách xa | `<w:pPrDefault/>` rỗng, thiếu style `Normal` → Word tự chế Calibri 11 / dãn 1,08 / after 8 pt | thêm `pPrDefault` (`after=120`, `line=288`, `auto`) + style `Normal` |
| 3 | Mất số trang; bảng bị co cột, tiêu đề rớt dòng | thiếu `footer1.xml` / `footerReference`; bảng `autofit` | thêm footer field `PAGE`; `tblW=5000pct` + `tblLayout=fixed` |

Kiểm tra & sửa tự động: `python3 scripts/fix_page_setup.py [--check] <file.docx>` (luôn tạo `.bak`,
validate XML bằng `ElementTree` trước khi ghi).

## 4. Vì sao tách nguyên trạng thay vì dựng lại mẫu

Dựng lại mẫu bằng tay (python-docx) dễ sai: thiếu cột, sai cỡ chữ, mất ghi chú chân trang, sai số mẫu.
`tt79_extract.py` sao chép file E-HSMT mẫu rồi **chỉ xoá các thành phần ngoài phạm vi biểu mẫu**,
giữ nguyên `styles.xml`, `numbering.xml`, footer, gộp ô, ký hiệu chú thích — nên bản ra khớp 100 %
với văn bản Bộ Tài chính ban hành.

Điều kiện bắt buộc khi ghi file: **giữ `<w:sectPr>`** ở cuối `w:body`; thiếu nó Word dùng khổ Letter,
lề 1 inch → in ra sai khổ.

## 5. Chuyển Word → Excel cho các bảng kê

Các mẫu 06–13 là webform nhiều dòng; bản Excel do script sinh chỉ dùng để **chuẩn bị dữ liệu offline**:

- mỗi bảng Word → 1 sheet, giữ nguyên tiêu đề cột của TT79 (3 dòng đầu in đậm, có nền);
- Times New Roman 11, wrap text, freeze `A4`, lặp tiêu đề `1:3` khi in, khổ in ngang, `fitToWidth=1`.
