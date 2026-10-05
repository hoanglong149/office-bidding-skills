# Danh mục bộ biểu mẫu HSDT — TT 79/2025/TT-BTC

Tách nguyên trạng từ E-HSMT mẫu trong `sources/` (Chương IV — Biểu mẫu mời thầu và dự thầu).

| Bộ | Loại gói thầu | Nguồn | Word | Excel | Thư mục |
|:--|:--|:--|--:|--:|:--|
| `EPC-10A` | EPC — thiết kế, cung cấp hàng hóa và xây lắp | Mẫu số 10A (một giai đoạn một túi hồ sơ) | 30 | 29 | `templates/EPC-10A/` |
| `Xay-lap-3A` | Xây lắp | Mẫu số 3A (một giai đoạn một túi hồ sơ) | 23 | 22 | `templates/Xay-lap-3A/` |
| `Hang-hoa-4A` | Mua sắm hàng hóa | Mẫu số 4A (một giai đoạn một túi hồ sơ) | 30 | 27 | `templates/Hang-hoa-4A/` |
| `EC-8A` | EC — thiết kế và xây lắp | Mẫu số 8A (một giai đoạn một túi hồ sơ) | 24 | 23 | `templates/EC-8A/` |
| `PC-9A` | PC — cung cấp hàng hóa và xây lắp | Mẫu số 9A (một giai đoạn một túi hồ sơ) | 33 | 32 | `templates/PC-9A/` |

**Không đóng gói mẫu 04A / 04B (bảo lãnh dự thầu)** — TT79 đánh dấu *Scan đính kèm*:
do tổ chức tín dụng phát hành theo mẫu của họ, không phải biểu mẫu nhà thầu tự lập.

Mỗi bộ có `DANH-MUC-BIEU-MAU.md` liệt kê chi tiết từng mẫu.

```bash
# tách lại / tách từ E-HSMT thực tế
python3 scripts/tt79_extract.py sources/TT79-M3A-E-HSMT-Xay-lap-01-tui.docx out/xay-lap
```
