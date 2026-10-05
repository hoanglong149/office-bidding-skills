# Cần gì để ra một bộ HSDT — danh mục ngữ cảnh đầu vào

Skill **không tự sinh nội dung**. Nó là **khuôn + máy lắp**: tách đúng biểu mẫu, đóng dấu letterhead,
tạo bảng Excel đúng cột, kiểm định dạng. Hồ sơ chỉ "ra" khi có **nguyên liệu** dưới đây.

Ký hiệu: 🤖 skill làm được (đã có công cụ) · 🧑 cần người/quyết định · 📄 cần tài liệu nguồn

> Bảng câu hỏi máy đọc được nằm ở `references/required-inputs.json`; dùng
> `python3 scripts/ask_context.py --form phieu-ngu-can.md` để xuất phiếu cho người dùng điền,
> và `--check` để biết còn thiếu gì trước khi cho phép sinh hồ sơ.

---

## A. Bắt buộc có trước khi bắt đầu (không có thì không chạy được)

| # | Ngữ cảnh | Loại | Lấy ở đâu |
|:--:|:--|:--:|:--|
| 1 | **Chương IV của E-HSMT thực tế** (file Word) | 📄 | HSMT gói thầu — để đối chiếu số mẫu, mục nào Webform / mục nào Scan |
| 2 | Loại gói thầu (EPC / xây lắp / hàng hóa / EC / PC / EP) | 🧑 | quyết định của gói |
| 3 | Tên dự án + tên gói thầu (nguyên văn như HSMT) | 📄 | HSMT / letterhead |
| 4 | Logo nhà thầu (và logo CĐT nếu HSMT yêu cầu) | 📄 | hồ sơ nhà thầu |
| 5 | Khổ giấy + lề + vị trí số trang theo HSMT | 📄 | Chương I/IV HSMT |

→ Có 1–5 là chạy được `tt79_extract.py` + `apply_project_format.py` để ra **bộ khung hồ sơ**.

## B. Dữ liệu để điền nội dung từng nhóm mẫu

| Mẫu | Cần gì | Loại | Skill làm được gì |
|:--|:--|:--:|:--|
| 02 Đơn dự thầu | ngày phát hành, tên gói, MST, giá dự thầu, tỷ lệ giảm giá, hiệu lực E-HSDT, giá trị + hiệu lực bảo đảm dự thầu | 🧑📄 | điền vào biểu mẫu nếu có số |
| 03 Thỏa thuận liên danh | thành viên, người đại diện, phân công công việc + tỷ lệ % | 📄🧑 | tạo bảng phân công đúng cột |
| 05A1–05A4 HĐ tương tự | danh mục hợp đồng: tên/số, ngày ký, hoàn thành, giá trị, CĐT, phạm vi | 📄 | đổ vào bảng Excel |
| 05B Năng lực sản xuất | nhà máy, công suất, chứng nhận | 📄 | đổ vào bảng |
| 06A/06B/06C Nhân sự | danh sách nhân sự chủ chốt + CV + bằng cấp/chứng chỉ | 📄🧑 | đổ danh sách; CV cần dữ liệu từng người |
| 06D Thiết bị thi công | loại, NSX, model, công suất, năm, sở hữu/thuê | 📄 | đổ vào bảng |
| 07 Lịch sử không hoàn thành | HĐ không hoàn thành do lỗi nhà thầu (nếu có) | 📄 | đổ vào bảng |
| 08A/08B/08C Tài chính | BCTC kiểm toán 3 năm, nguồn lực tài chính, HĐ đang thực hiện | 📄 | tính NLTC = TNL − ĐTH, đổ vào bảng |
| 09A/09B/09C Nhà thầu phụ | danh mục NTP, phạm vi, tỷ lệ | 🧑📄 | đổ vào bảng |
| 10A Tiến độ | mốc tiến độ theo Chương V HSMT | 📄🧑 | đổ vào bảng |
| 10B Giá hàng hóa | danh mục hàng hóa, NSX, xuất xứ, khối lượng, đơn giá | 📄🧑 | đổ danh mục; **giá do bộ phận thầu quyết** |
| 11.1x/11.2x Bảng tổng hợp giá | BOQ + đơn giá + loại hợp đồng | 🧑 | tạo khung đúng cột, không tự định giá |
| 12A/12B/12C | công nhật, khoản tạm tính, số liệu điều chỉnh | 🧑📄 | đổ vào bảng |
| 13A/13B/13C Ưu đãi | hàng hóa sản xuất trong nước + chi phí | 📄🧑 | đổ vào bảng |
| 04A/04B Bảo lãnh dự thầu | **không thuộc bộ mẫu** — tổ chức tín dụng phát hành | 🧑 | — |

## C. Việc con người phải làm (skill không thay được)

1. **Định giá dự thầu** — không có cơ sở để tự bịa số.
2. **Ký số + nộp trên Hệ thống** (các mẫu 02, 03, 05–13 là webform).
3. **Scan + công chứng** tài liệu đính kèm (bằng cấp, chứng chỉ, BCTC, bảo lãnh).
4. **Quyết định kỹ thuật**: Đề xuất kỹ thuật / biện pháp thi công (Chương V) — ngoài phạm vi bộ biểu mẫu này.
5. **Xác nhận với ngân hàng** về bảo lãnh dự thầu / bảo lãnh thực hiện hợp đồng.

---

## Kết luận thực dụng

| Có sẵn | Kết quả |
|:--|:--|
| A (1–5) | Bộ khung hồ sơ đúng mẫu, đóng dấu letterhead, Excel đúng cột — **dùng được ngay để phát cho các bộ môn điền** |
| A + B (dữ liệu trong vault/Docs) | Điền được phần lớn bảng: thiết bị, hàng hóa, nhân sự (danh sách), HĐ tương tự, tài chính, ưu đãi |
| A + B + C | Hồ sơ hoàn chỉnh để nộp |

Nói cách khác: **đủ ngữ cảnh = có A + dữ liệu nguồn cho từng mẫu ở B**. Phần C luôn thuộc về con người.
