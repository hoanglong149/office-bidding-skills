---
title: "Pdf Proposal Formal Navy"
date: 2026-04-23
tags: [converted, pdf, automated]
---

*Tài liệu tư vấn nội bộ - Không phát hành*

# BẢN ĐỀ XUẤT NÂNG CẤP

# KIẾN TRÚC HẠ TẦNG SỐ *Ứng dụng điện toán đám mây nhằm nâng cao * *năng lực vận hành cho Tổng Công ty Sigma* **Phạm vi áp dụng** **Tổng Công ty Cổ phần Hạ tầng Sigma** **Lĩnh vực trọng tâm** **Vận hành Datacenter - Hạ tầng Kỹ thuật số**

*Mở rộng sang Bảo mật, DevOps, Quản trị Dữ liệu* **I. Mục tiêu**

Giải bài toán quá tải hạ tầng khi khối lượng truy vấn dữ liệu tại Tổng Công ty tăng nhanh

theo chiến lược mở rộng 12 chi nhánh mới mỗi quý, trong bối cảnh không thể mua thêm máy chủ vật lý kịp tiến độ.

Phương án đề xuất: chuyển đổi toàn bộ hệ thống on-premise sang kiến trúc hybrid cloud,

cho phép mở rộng linh hoạt theo nhu cầu thực tế mà không cần đầu tư phần cứng lớn ban đầu. **1.1. Ba kết quả cuối cùng**

**Tăng năng lực xử lý gấp 3 lần**, chi phí vận hành chỉ tăng 15% nhờ tối ưu auto- scaling.

**Rút ngắn thời gian triển khai**, từ 6 tuần xuống 3 ngày nhờ containerization.

**Tự chủ vận hành**, đội ngũ nội bộ không phụ thuộc vendor sau khi dự án kết thúc. **II. Phạm vi triển khai**

Dự án được chia thành 3 giai đoạn chính, mỗi giai đoạn kéo dài 4 tuần. Phạm vi bao phủ

toàn bộ các hệ thống lõi bao gồm ERP, CRM, hệ thống giám sát mạng và nền tảng phân tích dữ liệu. **2.1. Giai đoạn 1 - Đánh giá và Thiết kế**

Khảo sát toàn bộ hạ tầng hiện tại tại 4 datacenter khu vực. Đánh giá mức độ sẵn sàng

chuyển đổi của từng hệ thống. Xây dựng bản thiết kế kiến trúc đích phù hợp với quy mô doanh nghiệp.

Kiểm kê toàn bộ máy chủ vật lý, virtual machines và containers hiện có.

Đo lường băng thông, latency và throughput trong giờ cao điểm.

Xác định các workload có thể migrate ngay so với nhóm cần refactor.

> [!success] Đầu ra Bản thiết kế kiến trúc đích (target architecture) được Ban Giám đốc phê duyệt. **2.2. Giai đoạn 2 - Triển khai Pilot**

Chọn 2 hệ thống ít rủi ro nhất để chạy pilot trên nền tảng cloud. Theo dõi hiệu năng trong 2 tuần trước khi mở rộng.

Hệ thống email và collaboration tools sẽ migrate trước.

Hệ thống giám sát mạng chuyển sang cloud-native stack.

***Điểm mấu chốt: *** Nếu pilot thành công với uptime trên 99.5%, toàn bộ hệ thống còn lại sẽ được phê duyệt migration trong vòng 1 tuần. **2.3. Giai đoạn 3 - Mở rộng toàn diện**

Sau khi pilot thành công, tiến hành migration cho các hệ thống còn lại theo thứ tự ưu tiên đã được phê duyệt bởi Ban Giám đốc.

**Kết quả cuối giai đoạn: ** Toàn bộ hệ thống lõi vận hành trên nền tảng hybrid cloud, đội ngũ nội bộ tự vận hành. **III. Nguồn lực và Ngân sách**

Dựa trên khảo sát sơ bộ và báo giá từ 3 nhà cung cấp dịch vụ cloud hàng đầu, chúng tôi đề xuất bảng phân bổ ngân sách như sau: **Bảng 1. Phân bổ ngân sách dự kiến**

**Hạng mục** **Chi phí (VNĐ)** **Ghi chú** Licensing Cloud Platform ==850,000,000== Hợp đồng 12 tháng Dịch vụ Migration ==420,000,000== Bao gồm đào tạo Nâng cấp Bảo mật ==280,000,000== Zero-trust architecture Dự phòng (15%) ==232,500,000== Buffer cho phát sinh

Tổng ngân sách dự kiến: ==1,782,500,000== VNĐ. Con số này thấp hơn 40% so với phương

án mua thêm máy chủ vật lý tương đương và không phát sinh chi phí bảo trì phần cứng hàng năm. **IV. Rủi ro và Biện pháp giảm thiểu**

**Gián đoạn dịch vụ**, thực hiện ngoài giờ hành chính, có rollback plan cho mỗi hệ thống.

**Thiếu nhân sự cloud**, đào tạo 8 kỹ sư nội bộ lấy chứng chỉ trong tháng đầu.

**Chi phí vượt dự toán**, thiết lập budget alert và spending cap trên cloud console. **V. Kế hoạch Đào tạo Nội bộ**

Để đảm bảo tính bền vững sau khi dự án kết thúc, toàn bộ đội ngũ kỹ thuật sẽ được đào tạo vận hành nền tảng mới: **Bảng 2. Chương trình đào tạo**

**Khóa học** **Đối tượng** **Thời lượng** **Hình thức** Cloud Foundation Toàn bộ IT 16 giờ Workshop DevOps Pipeline Team Ops 24 giờ Lab thực hành Security & Compliance Team Security 8 giờ Seminar

**Kết quả cuối giai đoạn: ** Sau đào tạo, đội ngũ Sigma hoàn toàn tự chủ vận hành mà không cần hỗ trợ từ bên ngoài.

*- Hết -*