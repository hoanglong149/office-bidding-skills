# skills - Biểu mẫu & Văn bản (Việt Nam)

Monorepo chứa **2 skill độc lập** cho agent (omp / Claude Code / Cursor), phục vụ công việc đấu thầu
và văn phòng theo đúng chuẩn pháp lý Việt Nam:

| Skill | Chuẩn | Dùng cho |
|:--|:--|:--|
| [`skills/hsdt-forms`](skills/hsdt-forms/) | **Thông tư 79/2025/TT-BTC** | Biểu mẫu Hồ sơ dự thầu (E-HSDT): đơn dự thầu, thỏa thuận liên danh, bảo lãnh, nhân sự, thiết bị, tài chính, bảng giá dự thầu... |
| [`skills/nd30-forms`](skills/nd30-forms/) | **Nghị định 30/2020/NĐ-CP** | Văn bản hành chính & xử lý file văn phòng: công văn, quyết định, tờ trình, báo cáo, biên bản, thông báo... + Word/Excel/Slide/PDF |

> Hai chuẩn này **khác nhau về thể thức**, không trộn:
> HSDT **không có** `Số: .../...`, không có `Hà Nội, ngày ...`, không có khối tiêu đề 2 cột.
> Văn bản hành chính thì **bắt buộc** có.

---

## Cài đặt

### Cách 1. Cài như skill cho agent (khuyến nghị)

```bash
git clone https://github.com/hoanglong149/office-bidding-skills.git
cd office-bidding-skills

# cả hai skill, mức người dùng
cp -r skills/hsdt-forms skills/nd30-forms ~/.agents/skills/

# hoặc chỉ một skill, hoặc theo từng project
cp -r skills/hsdt-forms ~/.agents/skills/hsdt-forms
cp -r skills/nd30-forms  <workspace>/.agents/skills/nd30-forms
```

Agent tự đọc `SKILL.md` của skill phù hợp khi gặp việc tương ứng.

### Cách 2. Dùng CLI trên máy (không cần agent)

Chỉ cần Python. Bộ biểu mẫu có sẵn trong `templates/` nên phần lớn việc không cần cài gì thêm.

#### Kiểm tra Python

```bash
python3 --version     # macOS, Linux  (cần 3.9 trở lên)
python --version      # Windows PowerShell
```

| Nền tảng | Cài Python nếu chưa có |
|:--|:--|
| macOS | `brew install python` hoặc tải từ python.org |
| Debian/Ubuntu | `sudo apt update && sudo apt install -y python3 python3-pip python3-venv` |
| Fedora/RHEL | `sudo dnf install -y python3 python3-pip` |
| Windows | `winget install Python.Python.3.12` hoặc tải từ python.org. Khi cài nhớ tích "Add python.exe to PATH" |
| WSL / Git Bash | dùng như Linux |

#### Cài thư viện

Nên dùng môi trường ảo để không ảnh hưởng Python hệ thống.

macOS và Linux:

```bash
cd office-bidding-skills/skills/hsdt-forms
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows PowerShell:

```powershell
cd office-bidding-skills\skills\hsdt-forms
py -3 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Windows cmd:

```bat
cd office-bidding-skills\skills\hsdt-forms
py -3 -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

Nếu không dùng môi trường ảo, cài trực tiếp:

```bash
python3 -m pip install --user lxml python-docx openpyxl
```

Skill `nd30-forms` (tạo/sửa Word, Excel, Slide, chuyển đổi PDF):

```bash
cd ../nd30-forms
python3 -m pip install -r requirements.txt      # trong môi trường ảo ở trên
```

Chuyển đổi sang PDF trong `nd30-forms` cần LibreOffice (`soffice`) trên PATH:

| Nền tảng | Cài LibreOffice |
|:--|:--|
| macOS | `brew install --cask libreoffice` |
| Debian/Ubuntu | `sudo apt install -y libreoffice` |
| Fedora/RHEL | `sudo dnf install -y libreoffice` |
| Windows | `winget install TheDocumentFoundation.LibreOffice` |

#### Chạy

macOS và Linux:

```bash
cd skills/hsdt-forms
bash lam-ho-so.sh                                   # hỏi từng câu rồi sinh hồ sơ
python3 scripts/ask_context.py --form phieu.md      # cũng chạy được trên Windows
python3 scripts/tt79_extract.py sources/TT79-M3A-E-HSMT-Xay-lap-01-tui.docx out/xay-lap
```

Windows PowerShell (dùng `py -3` thay `python3`, và gọi trực tiếp script Python):

```powershell
cd skills\hsdt-forms
py -3 scripts\ask_context.py --form phieu.md
py -3 scripts\ask_context.py --check phieu.md
py -3 scripts\tt79_extract.py sources\TT79-M3A-E-HSMT-Xay-lap-01-tui.docx out\xay-lap
py -3 scripts\apply_project_format.py templates\EPC-10A --config projects\<ma-goi>.json --out out\<ma-goi>
py -3 scripts\fix_page_setup.py --check out\<ma-goi>\*.docx
```

`lam-ho-so.sh` là script bash: chạy được trên macOS, Linux, WSL và Git Bash; trên PowerShell thì dùng
chuỗi lệnh Python như trên.

Tất cả script đều là Python thuần, không cần biên dịch, không phụ thuộc hệ điều hành.

---

## 1. `hsdt-forms` - Biểu mẫu HSDT theo TT 79/2025/TT-BTC

Bộ biểu mẫu **tách nguyên trạng** từ E-HSMT mẫu của Bộ Tài chính (Chương IV — Biểu mẫu mời thầu và dự thầu),
giữ 100 % định dạng gốc; kèm công cụ tách cho mọi loại gói thầu.

| Bộ trong `skills/hsdt-forms/templates/` | Loại gói thầu | Word | Excel |
|:--|:--|--:|--:|
| `EPC-10A` | EPC | 32 | 30 |
| `Xay-lap-3A` | Xây lắp | 25 | 22 |
| `Hang-hoa-4A` | Mua sắm hàng hóa | 32 | 29 |
| `EC-8A` | EC | 26 | 23 |
| `PC-9A` | PC | 35 | 32 |

**Người không rành kỹ thuật — 1 lệnh:**

```bash
cd skills/hsdt-forms && bash lam-ho-so.sh     # hỏi từng câu: ra bộ hồ sơ, mở thư mục kết quả
```

Thiếu ngữ cảnh thì skill **hỏi lại**, không tự bịa:
`python3 scripts/ask_context.py --form phieu-ngu-can.md`: điền: `--check phieu-ngu-can.md`.

```bash
python3 scripts/tt79_extract.py sources/TT79-M3A-E-HSMT-Xay-lap-01-tui.docx out/xay-lap
python3 scripts/apply_project_format.py templates/EPC-10A --config projects/<mã-gói>.json --out out
python3 scripts/fix_page_setup.py --check file.docx
```

| Mẫu 02 — Đơn dự thầu | Mẫu 06A — Nhân sự chủ chốt |
|:--|:--|
| ![Mẫu 02](skills/hsdt-forms/examples/preview/TT79-M02.docx.png) | ![Mẫu 06A](skills/hsdt-forms/examples/preview/TT79-M06A.xlsx.png) |

---

## 2. `nd30-forms` - Văn bản hành chính NĐ 30 & xử lý file văn phòng

Bộ quy chuẩn trình bày + mẫu văn bản hành chính, kèm công cụ tạo/sửa/chuyển đổi file.

| Thành phần | Nội dung |
|:--|:--|
| `standards/nd30.md` + `standards/structure/` | Quy chuẩn thể thức NĐ 30: khổ giấy, lề, typography, heading, bảng, header/footer, bìa, caption |
| `standards/color/` | 4 bộ phối màu Word + palette Excel/Slide |
| `templates/` | 10 mẫu: công văn, quyết định, tờ trình, báo cáo, biên bản, thông báo, kế hoạch, giấy mời/ủy quyền, template chung, đề xuất |
| `scripts/office/` | Unpack/Pack XML (giữ format mẫu), validate, clone text |
| `scripts/format/` | Chuẩn hoá định dạng `.docx` |
| `scripts/convert/` | MD→DOCX, PDF→DOCX, batch PDF→MD |
| `resources/` | Cách dùng thư viện: docx, xlsx, pptx, pdf, office-xml, convert |

```bash
cd skills/nd30-forms
python3 scripts/office/unpack.py file.docx ./work      # bung để sửa XML
python3 scripts/office/pack.py  ./work file-moi.docx   # đóng gói lại
```

---

## Cấu trúc repo

```text
.
├── README.md
├── LICENSE
└── skills/
    ├── hsdt-forms/          # skill 1 — TT 79/2025/TT-BTC
    │   ├── SKILL.md
    │   ├── sources/         # E-HSMT mẫu chính thức
    │   ├── templates/       # 5 bộ biểu mẫu HSDT đã tách
    │   ├── scripts/         # tt79_extract.py, fix_page_setup.py
    │   ├── references/      # danh mục mẫu, đặc tả định dạng
    │   ├── examples/        # demo + ảnh render
    │   └── demo_output/
    └── nd30-forms/          # skill 2 — NĐ 30/2020/NĐ-CP
        ├── SKILL.md
        ├── standards/       # nd30.md, structure/, color/
        ├── templates/       # mẫu văn bản hành chính
        ├── scripts/         # office/, format/, convert/
        ├── resources/
        ├── examples/
        └── tests/
```

## Tác giả

Hoàng Long - [github.com/hoanglong149](https://github.com/hoanglong149)

## License

MIT, xem [LICENSE](LICENSE). Riêng `skills/hsdt-forms/sources/` chứa phụ lục E-HSMT mẫu do
Bộ Tài chính ban hành kèm TT 79/2025/TT-BTC, là văn bản công khai.
