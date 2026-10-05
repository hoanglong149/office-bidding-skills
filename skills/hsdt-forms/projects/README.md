# Cấu hình theo gói thầu

Mỗi gói thầu một file JSON — **không hard-code tên gói vào biểu mẫu**.

```bash
cp projects/_template.json projects/<mã-gói>.json
# sửa: project, package, logo_left/right, paper, profile, logo_height, page_number

python3 scripts/apply_project_format.py templates/EPC-10A \
  --config projects/<mã-gói>.json --out out/<mã-gói>
```

| Khoá | Ý nghĩa |
|:--|:--|
| `project` | tên dự án — hiện trong letterhead, dòng `Project : …` |
| `package` | tên gói thầu — dòng `Bidding Package : …` |
| `logo_left` / `logo_right` | đường dẫn ảnh logo (bỏ trống nếu không dùng) |
| `paper` | `a4` (mặc định) hoặc `letter` |
| `profile` | `intl-epc` (định dạng gói quốc tế) hoặc `keep` (giữ nguyên bản TT79) |
| `logo_height` | chiều cao logo (mm), mặc định 18 |
| `page_number` | `footer` (mặc định) · `header` · `keep` |

Tham số dòng lệnh (`--project`, `--paper`, …) **đè** giá trị trong file cấu hình.
