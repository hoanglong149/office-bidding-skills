# Cấu hình theo gói thầu

**Tên dự án / tên gói thầu là chỗ trống `[TÊN DỰ ÁN]` / `[TÊN GÓI THẦU]`** — điền theo từng gói,
KHÔNG hard-code tên gói nào vào biểu mẫu hay vào script.

```bash
cp projects/_template.json projects/<mã-gói>.json
# điền: project, package (tên thật lấy từ HSMT/letterhead của gói), logo, paper, profile, page_number

python3 scripts/apply_project_format.py templates/EPC-10A \
  --config projects/<mã-gói>.json --out out/<mã-gói>
```

| Khoá | Mặc định | Ý nghĩa |
|:--|:--|:--|
| `project` | `[TÊN DỰ ÁN]` | tên dự án — dòng `Project : ...` trong letterhead |
| `package` | `[TÊN GÓI THẦU]` | tên gói thầu — dòng `Bidding Package : ...` |
| `logo_left` / `logo_right` | `assets/logo-ipc-ec.png` ở phải | ảnh logo (bỏ trống nếu không dùng) |
| `paper` | `a4` | `a4` hoặc `letter` |
| `profile` | `intl-epc` | `intl-epc` (áp định dạng bảng/ô) hoặc `keep` |
| `logo_height` | `18` | chiều cao logo (mm), giữ tỉ lệ ảnh |
| `page_number` | `footer` | `footer` · `header` · `keep` |

Script **dừng ngay** nếu tên còn dấu `[ ]` — tránh đóng dấu placeholder lên cả bộ biểu mẫu.

Tham số dòng lệnh (`--project`, `--paper`, ...) đè giá trị trong file cấu hình.
