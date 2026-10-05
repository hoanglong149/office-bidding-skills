# Ánh xạ biểu mẫu HSDT ↔ file ↔ tham chiếu HSMT

Danh mục theo cấu trúc **Chapter IV — Bidding Forms** của HSMT EPC tiêu chuẩn
(Technical Proposal I → Financial Proposal II). Số Form có thể khác giữa các gói —
**luôn map lại theo danh mục trong HSMT của chính gói thầu**.

---

## I. TECHNICAL PROPOSAL

| Form | Nội dung | Định dạng | File / Sheet |
|:--|:--|:--|:--|
| 01 | Letter of Bid (Technical Proposal) | Word | `HSDT-F01-Letter-of-Bid-Technical.docx` |
| 02 | Power of Attorney | Word | `HSDT-F02-Power-of-Attorney.docx` |
| 03 | Consortium / Joint Venture Agreement | Word | `HSDT-F03-Consortium-JV-Agreement.docx` |
| 04 | Bid Security | Word | `HSDT-F04-Bid-Security.docx` |
| 05(a) | Bidder Information | Word | `HSDT-F05a-Bidder-Information.docx` |
| 05(b) | Consortium / JV Party Information | Word | `HSDT-F05b-JV-Party-Information.docx` |
| 06A | Similar EPC/EC/EP/PC contracts — Bidder | **Excel** | `HSDT-F06A-Similar-Contracts-Bidder.xlsx` |
| 06B | Similar Engineering (E) — Special sub-contractor | **Excel** | `HSDT-F06B-Similar-Engineering-Subcontractor.xlsx` |
| 06C | Similar Construction (C) — Special sub-contractor | **Excel** | `HSDT-F06C-Similar-Construction-Subcontractor.xlsx` |
| 07A | Table of proposed personnel for Key Person | **Excel** | `HSDT-F07A-F07C-Nhan-su-chu-chot.xlsx` → sheet `Nhan su chu chot` |
| 07B | Curriculum Vitae of Key Person | Word | `HSDT-F07B-CV-Key-Personnel.docx` |
| 07C | Professional qualification of Key Person | **Excel** | `HSDT-F07A-F07C-Nhan-su-chu-chot.xlsx` → sheet `Trinh do chuyen mon` |
| 08 | List of main construction equipment | **Excel** | `HSDT-F08-Thiet-bi-thi-cong.xlsx` |
| 09A | Previous non-fulfilment — Bidder’s fault | **Excel** | `HSDT-F09A-F09B-Non-fulfilment.xlsx` → sheet `09A` |
| 09B | Previous non-fulfilment — Special sub-contractor | **Excel** | `HSDT-F09A-F09B-Non-fulfilment.xlsx` → sheet `09B` |
| 10A1 | Financial situation of Bidder | **Excel** | `HSDT-F10-F11-Tai-chinh.xlsx` → sheet `10A1` |
| 10A2 | Financial situation of Special sub-contractor | **Excel** | `HSDT-F10-F11-Tai-chinh.xlsx` → sheet `10A2` |
| 10B | Average annual revenue | **Excel** | `HSDT-F10-F11-Tai-chinh.xlsx` → sheet `10B` |
| 11 | Financial resources (TFR − MFR) | **Excel** | `HSDT-F10-F11-Tai-chinh.xlsx` → sheet `11` |
| 12 | Monthly financial resources — contracts in progress | **Excel** | `HSDT-F12-F16-Danh-muc.xlsx` → sheet `12` |
| 13 | Works item performed by subcontractors | **Excel** | `HSDT-F12-F16-Danh-muc.xlsx` → sheet `13` |
| 14 | List of special subcontractors | **Excel** | `HSDT-F12-F16-Danh-muc.xlsx` → sheet `14` |
| 15 | Subsidiary / associate company in charge of package works | **Excel** | `HSDT-F12-F16-Danh-muc.xlsx` → sheet `15` |
| 16 | Time schedule of project | **Excel** | `HSDT-F12-F16-Danh-muc.xlsx` → sheet `16` |
| 17(a) | Employer / End user’s certificate | Word (letterhead **EMPLOYER**) | `HSDT-F17a-Employer-Certificate.docx` |
| 17(b) | Employer / Contractor certificate | Word | `HSDT-F17b-Contractor-Certificate.docx` |
| 18A | Technical deviation declaration | **Excel** | `HSDT-F18A-F18B-Deviation.xlsx` → sheet `18A` |
| 18B | Commercial deviation declaration | **Excel** | `HSDT-F18A-F18B-Deviation.xlsx` → sheet `18B` |
| Att. 1A | Manufacturers & country of origin — systems | **Excel** | `HSDT-ATT1A-1B-2-Manufacturers-Origin-Equipment.xlsx` → sheet `ATT1A` |
| Att. 1B | Manufacturers & country of origin — major equipment | **Excel** | `… → sheet ATT1B` |
| Att. 2 | Equipment and materials of the project | **Excel** | `… → sheet ATT2` |

## II. FINANCIAL PROPOSAL

| Form | Nội dung | Định dạng | File / Sheet |
|:--|:--|:--|:--|
| 19(a) | Letter of Bid (Financial Proposal) | Word | `HSDT-F19a-Letter-of-Bid-Financial.docx` |
| 19(b) | Letter of Bid (Financial Proposal) — alternative | Word | `HSDT-F19b-Letter-of-Bid-Financial-Alt.docx` |
| 20 | Grand Summary | **Excel** | `HSDT-F20-F23-Bieu-gia.xlsx` → sheet `20` |
| 21 | Price Schedules | **Excel** | `… → sheets 21.1 / 21.2 / 21.3` |
| 22 | Domestic costs eligible for incentives | **Excel** | `… → sheet 22` |
| 23 | Schedule of adjustment data | **Excel** | `… → sheet 23` |

---

## Tổng kết định dạng

| | Số lượng | Ghi chú |
|:--|--:|:--|
| Word (.docx) | **11** | Văn bản có chữ ký/đóng dấu |
| Excel (.xlsx) | **11 workbook / 28 sheet** | Bao 26 form danh mục + biểu giá |
| Tổng biểu mẫu phủ | **36 + 3 Attachment** | Theo Chapter IV Bidding Forms |

---

## Cấu trúc cột tham chiếu (đã dựng sẵn trong file)

| Form | Cột |
|:--|:--|
| 06A/B/C | No. · Name and contract number · Employer/End user · Contractor/Sub-contractor · Value · Duration & completion · Scope of works · Remark |
| 07A | No. · Name · Position in charge · Years of experience · Remark |
| 07C | No. · Full name · Position · Degree/Certificate · Issuing authority · Issue date · Validity · Remark |
| 08 | No. · Equipment · Brand/Origin · Year · Quantity · Capacity · Condition · Owned/Leased · Remark |
| 12 | No. · Contract’s name · Employer contact · Finish date · Remaining months · Unpaid value · Monthly resources required |
| 18A/18B | Deviation No. · Reference to Bidding Documents · Reference to Bidder’s Proposal · Deviations/Reservations · Remark |
| Att. 2 | No. · Item · Manufacturer · Country of Origin · Mode/codes/labels · Technical specification · Remark |
| 20 | No. · Description · Reference in Schedule · Amount (local) · Amount (foreign) · Remark |
| 21.1 | No. · Descriptions · Unit · Quantity · Cost (design & manufacture) · Freight & Insurance · Inland transportation · Total |
| 21.2 / 21.3 | No. · Descriptions · Unit · Quantity · Unit price · Total |
