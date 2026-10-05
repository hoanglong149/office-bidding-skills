# Ngữ cảnh cần có để làm một bộ HSDT

Skill không sinh nội dung. Nó tách đúng biểu mẫu, đóng dấu letterhead, tạo bảng Excel đúng cột và
kiểm định dạng. Hồ sơ chỉ ra được khi có đủ dữ liệu dưới đây.

Cột "Nguồn" ghi nơi lấy. Không có nguồn thì ghi `[……]`, không tự đoán.

Bảng câu hỏi máy đọc được nằm ở `references/required-inputs.json`. Dùng
`python3 scripts/ask_context.py --form phieu-ngu-can.md` để xuất phiếu cho người dùng điền,
rồi `--check` để biết còn thiếu mục nào trước khi cho phép sinh hồ sơ.

## A. Phải có trước khi chạy

| # | Ngữ cảnh | Nguồn |
|--:|:--|:--|
| 1 | Chương IV của E-HSMT thực tế (file Word) | HSMT gói thầu. Dùng để đối chiếu số mẫu, mục nào Webform, mục nào Scan |
| 2 | Loại gói thầu: EPC / xây lắp / hàng hóa / EP / EC / PC / tư vấn | Quyết định của gói |
| 3 | Tên dự án, nguyên văn như HSMT | HSMT |
| 4 | Tên gói thầu, nguyên văn như HSMT | HSMT |
| 5 | Logo nhà thầu, và logo chủ đầu tư nếu HSMT yêu cầu | Hồ sơ nhà thầu |
| 6 | Khổ giấy, lề, vị trí số trang theo HSMT | Chương I và IV HSMT |
| 7 | Nộp qua mạng hay nộp bản giấy; bản giấy thì in/ký/scan mẫu nào | HSMT |

Có mục 1 đến 7 là chạy được `tt79_extract.py` và `apply_project_format.py` để ra bộ khung hồ sơ.

## B. Dữ liệu để điền nội dung

| Mẫu | Cần gì | Nguồn |
|:--|:--|:--|
| 02 | Ngày phát hành, tên gói, mã số thuế, giá dự thầu, tỷ lệ giảm giá, hiệu lực E-HSDT, giá trị và hiệu lực bảo đảm dự thầu | Bộ phận thầu và ngân hàng |
| 03 | Thành viên liên danh, người đại diện, phân công công việc và tỷ lệ phần trăm | Thỏa thuận liên danh |
| 05A1 đến 05A4 | Danh mục hợp đồng tương tự: tên và số hợp đồng, ngày ký, ngày hoàn thành, giá trị, chủ đầu tư, phạm vi | Hồ sơ năng lực |
| 05B | Nhà máy, địa chỉ, tổng mức đầu tư, công suất, chứng nhận | Nhà sản xuất |
| 06A, 06B, 06C | Danh sách nhân sự chủ chốt và lý lịch từng người: căn cước, ngày sinh, bằng cấp chứng chỉ, đơn vị công tác, số năm kinh nghiệm | Bộ phận nhân sự |
| 06D | Thiết bị thi công: loại, nhà sản xuất, model, công suất, năm sản xuất, xuất xứ, hiện trạng, sở hữu hay thuê | Bộ phận kỹ thuật |
| 07 | Hợp đồng không hoàn thành do lỗi nhà thầu, nếu có | Hồ sơ năng lực |
| 08A, 08B, 08C | Báo cáo tài chính kiểm toán 3 năm gần nhất và danh sách hợp đồng đang thực hiện | Bộ phận tài chính |
| 09A, 09B, 09C | Nhà thầu phụ: danh mục công việc, tên và địa chỉ, phạm vi, tỷ lệ | Thương thảo với nhà thầu phụ |
| 10A | Mốc tiến độ theo Chương V HSMT | HSMT |
| 10B, 11.1, 11.2 | BOQ, khối lượng, đơn giá, loại hợp đồng, giá chào đã gồm hay chưa gồm thuế phí lệ phí | Bộ phận thầu |
| 12A, 12B, 12C | Công nhật, khoản tạm tính, số liệu điều chỉnh giá | Bộ phận thầu |
| 13A, 13B, 13C | Hàng hóa sản xuất trong nước và chi phí sản xuất trong nước | Nhà thầu hoặc nhà sản xuất |
| 04A, 04B | Bảo lãnh dự thầu. Không thuộc bộ mẫu, do tổ chức tín dụng phát hành | Ngân hàng |

## C. Việc con người phải làm

1. Chốt giá dự thầu. Skill không định giá.
2. Ký số và nộp trên Hệ thống. Các mẫu 02, 03 và 05 đến 13 là webform.
3. Scan và công chứng tài liệu đính kèm: bằng cấp, chứng chỉ, báo cáo tài chính, bảo lãnh.
4. Lập đề xuất kỹ thuật và biện pháp thi công theo Chương V. Việc này nằm ngoài bộ biểu mẫu.
5. Làm việc với ngân hàng về bảo lãnh dự thầu và bảo lãnh thực hiện hợp đồng.

## Kết luận

| Đã có | Kết quả |
|:--|:--|
| Mục A | Bộ khung hồ sơ đúng mẫu, có letterhead, Excel đúng cột. Phát được cho các bộ môn điền |
| Mục A và dữ liệu mục B | Điền được phần lớn bảng: thiết bị, hàng hóa, danh sách nhân sự, hợp đồng tương tự, tài chính, ưu đãi |
| Mục A, B và C | Hồ sơ hoàn chỉnh để nộp |

Đủ ngữ cảnh nghĩa là có mục A và có dữ liệu nguồn cho từng mẫu ở mục B. Mục C luôn do con người làm.
