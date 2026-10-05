#!/usr/bin/env python3
"""
hsdt_build_excel.py — Sinh bộ biểu mẫu HSDT dạng EXCEL (.xlsx).

Dùng cho các biểu mẫu dạng DANH MỤC / LIST / BIỂU GIÁ (không phải văn bản hành chính):
  Danh mục  : 06A/06B/06C, 07A, 07C, 08, 09A/09B, 12, 13, 14, 15, 16, 18A/18B,
              Attachment 1A/1B/2
  Tài chính : 10A1, 10A2, 10B, 11
  Biểu giá  : 20 (Grand Summary), 21 (Price Schedules 1–3), 22 (ưu đãi), 23 (điều chỉnh giá)

Quy cách in: A4, Times New Roman, lề chuẩn NĐ30 (trên/dưới 20mm, trái 30mm, phải 20mm),
lặp dòng tiêu đề khi in, đánh số trang ở chân trang.

Chạy:  python3 hsdt_build_excel.py [thư_mục_đích]
"""

from __future__ import annotations
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

DEFAULT_OUT = None  # xem _default_out()
FONT = "Times New Roman"
TITLE_SZ, HDR_SZ, BODY_SZ = 13, 11, 11

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_HDR = PatternFill("solid", start_color="D9E1F2")
F_TITLE = Font(name=FONT, size=TITLE_SZ, bold=True)
F_LBL = Font(name=FONT, size=BODY_SZ, bold=True)
F_BODY = Font(name=FONT, size=BODY_SZ)
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left", vertical="center", wrap_text=True)
R = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _safe_sheet(name: str) -> str:
    """Excel cấm các ký tự \ / ? * [ ] : trong tên sheet."""
    return re.sub(r'[\\/?*\[\]:]', "-", name)[:31]


def page_setup(ws, landscape=False, title_rows="1:6"):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 1.18   # 30 mm
    ws.page_margins.right = 0.79  # 20 mm
    ws.page_margins.top = 0.79
    ws.page_margins.bottom = 0.79
    ws.page_margins.header = 0.49
    ws.page_margins.footer = 0.49
    ws.print_title_rows = title_rows
    ws.oddFooter.center.text = "Trang &P/&N"
    ws.oddFooter.center.size = 10
    ws.oddFooter.center.font = FONT
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def header_block(ws, *, code, en, vn, ncol=6, mode="BIDDER", money=None):
    """Khối tiêu đề + tên biểu mẫu (letterhead điền theo từng gói thầu)."""
    left = ("[TÊN NHÀ THẦU / LIÊN DANH NHÀ THẦU]" if mode == "BIDDER"
            else "[TÊN CHỦ ĐẦU TƯ / BÊN MỜI THẦU]" if mode == "EMPLOYER" else "")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, ncol // 2))
    ws.cell(1, 1, left).font = F_LBL
    ws.cell(1, 1).alignment = L
    ws.merge_cells(start_row=1, start_column=max(2, ncol // 2) + 1, end_row=1, end_column=ncol)
    ws.cell(1, max(2, ncol // 2) + 1,
            "Số: ……/……-HSDT          [Địa danh], ngày …… tháng …… năm 20……").font = F_BODY
    ws.cell(1, max(2, ncol // 2) + 1).alignment = R
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    ws.cell(2, 1, f"{code} — {en}").font = F_TITLE
    ws.cell(2, 1).alignment = C
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncol)
    ws.cell(3, 1, vn).font = Font(name=FONT, size=BODY_SZ, italic=True)
    ws.cell(3, 1).alignment = C
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=ncol)
    if money:
        ws.cell(4, 1, f"Đơn vị tiền: {money}").font = Font(name=FONT, size=9, italic=True)
        ws.cell(4, 1).alignment = C
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30
    return 6  # dòng bắt đầu bảng


def sheet(wb, name, *, code, en, vn, headers, widths, rows=6, money=None,
          landscape=False, sample_rows=6, formulas=None, note=None,
          mode="BIDDER", num_cols=None, preset_rows=None):
    ws = wb.create_sheet(_safe_sheet(name))
    start = header_block(ws, code=code, en=en, vn=vn, ncol=len(headers),
                         mode=mode, money=money)
    # header row
    for i, h in enumerate(headers, 1):
        c = ws.cell(start, i, h)
        c.font = Font(name=FONT, size=HDR_SZ, bold=True)
        c.fill = F_HDR; c.alignment = C; c.border = BORDER
    ws.row_dimensions[start].height = 34
    end = start + rows
    for r in range(start + 1, end + 1):
        for i in range(1, len(headers) + 1):
            c = ws.cell(r, i)
            c.border = BORDER
            c.font = F_BODY
            c.alignment = C if (num_cols is None or i not in num_cols) else R
        if i == 1:
            ws.cell(r, 1, r - start)
    # cột: STT tự động ở cột 1 nếu header bắt đầu bằng No./STT
    if headers and headers[0].strip().lower() in ("no.", "no", "stt", "tt"):
        for r in range(start + 1, end + 1):
            ws.cell(r, 1).value = r - start
    # công thức tổng
    if formulas:
        for col_letter, rng in formulas:
            rc = end + 1
            ws.cell(rc, 1, "TỔNG").font = F_LBL
            ws.cell(rc, 1).alignment = C
            cell = ws.cell(rc, ws[col_letter + "1"].column)
            cell.value = f"=SUM({col_letter}{start+1}:{col_letter}{end})"
            cell.font = F_LBL; cell.border = BORDER; cell.alignment = R
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if num_cols:
        for i in num_cols:
            for r in range(start + 1, end + 2):
                ws.cell(r, i).number_format = '#,##0'
    # dòng cố định (mô tả sẵn) — ghi đè phần đầu
    if preset_rows:
        for i, vals in enumerate(preset_rows):
            for j, v in enumerate(vals, 1):
                c = ws.cell(start + 1 + i, j, v)
                c.border = BORDER; c.font = F_BODY
                c.alignment = L if j in (1, 2, 3) else R
    # kẻ viền đầy đủ cho dòng tổng
    if formulas:
        for i in range(1, len(headers) + 1):
            ws.cell(end + 1, i).border = BORDER
    if note:
        rn = end + (2 if formulas else 2)
        ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=len(headers))
        ws.cell(rn, 1, note).font = Font(name=FONT, size=9, italic=True)
        ws.cell(rn, 1).alignment = L
    page_setup(ws, landscape=landscape)
    ws.freeze_panes = ws.cell(start + 1, 1)
    return ws


# ------------------------------------------------------------------ forms
def build(out):
    os.makedirs(out, exist_ok=True)
    made = []

    def new_wb():
        wb = Workbook(); wb.remove(wb.active); return wb

    # ============ 1) HỢP ĐỒNG TƯƠNG TỰ 06A / 06B / 06C ============
    sim_hdr = ["No.", "Name and contract number", "Name of Employer / End user",
               "Name of Contractor / Sub-contractor", "Value of contract",
               "Duration & completion status", "Scope of works", "Remark"]
    sim_w = [5, 26, 22, 24, 16, 20, 24, 12]
    for code, en, vn, fn in (
        ("Form 06A", "SIMILAR EPC, EC, EP, PC CONTRACTS PERFORMED BY THE BIDDER",
         "Danh mục hợp đồng tương tự do Nhà thầu thực hiện", "HSDT-F06A-Similar-Contracts-Bidder.xlsx"),
        ("Form 06B", "SIMILAR ENGINEERING CONTRACT (E) BY THE SPECIAL SUB-CONTRACTOR",
         "Danh mục hợp đồng thiết kế tương tự (nhà thầu phụ đặc biệt)",
         "HSDT-F06B-Similar-Engineering-Subcontractor.xlsx"),
        ("Form 06C", "SIMILAR CONSTRUCTION CONTRACT (C) BY THE SPECIAL SUB-CONTRACTOR",
         "Danh mục hợp đồng thi công tương tự (nhà thầu phụ đặc biệt)",
         "HSDT-F06C-Similar-Construction-Subcontractor.xlsx"),
    ):
        wb = new_wb()
        sheet(wb, "DS hop dong tuong tu", code=code, en=en, vn=vn,
              headers=sim_hdr, widths=sim_w, rows=15, landscape=True,
              note="Khai mỗi hợp đồng 1 dòng; đính kèm bản sao hợp đồng + biên bản nghiệm thu/thanh lý; "
                   "giá trị và tính chất công việc phải đạt yêu cầu tối thiểu của HSMT; "
                   "mức độ hoàn thành tối thiểu 80% khối lượng.")
        wb.save(os.path.join(out, fn)); made.append(fn)

    # ============ 2) NHÂN SỰ 07A / 07C ============
    wb = new_wb()
    sheet(wb, "Nhan su chu chot", code="Form 07A",
          en="TABLE OF PROPOSED PERSONNEL FOR KEY PERSON",
          vn="Bảng kê nhân sự chủ chốt đề xuất",
          headers=["No.", "Name", "Position in charge", "No. of years of experience", "Remark"],
          widths=[5, 30, 34, 16, 20], rows=12, landscape=True,
          note="Mỗi vị trí chủ chốt đề xuất 01 người; kèm CV (Form 07B), bằng cấp, chứng chỉ hành nghề; "
               "nhân sự phải thuộc biên chế nhà thầu (xác nhận BHXH nếu HSMT yêu cầu).")
    sheet(wb, "Trinh do chuyen mon", code="Form 07C",
          en="PROFESSIONAL QUALIFICATION OF PERSONNEL FOR KEY PERSON",
          vn="Trình độ chuyên môn của nhân sự chủ chốt",
          headers=["No.", "Full name", "Position", "Degree / Certificate", "Issuing authority",
                   "Issue date", "Validity", "Remark"],
          widths=[5, 24, 20, 26, 22, 12, 12, 16], rows=12, landscape=True,
          note="Đính kèm bản sao bằng tốt nghiệp, chứng chỉ hành nghề còn hiệu lực.")
    wb.save(os.path.join(out, "HSDT-F07A-F07C-Nhan-su-chu-chot.xlsx"))
    made.append("HSDT-F07A-F07C-Nhan-su-chu-chot.xlsx")

    # ============ 3) THIẾT BỊ 08 ============
    wb = new_wb()
    sheet(wb, "Thiet bi thi cong", code="Form 08",
          en="LIST OF MAIN CONSTRUCTION EQUIPMENT", vn="Danh mục thiết bị thi công chính",
          headers=["No.", "Equipment, machinery", "Brand / Origin", "Year of manufacture",
                   "Quantity", "Capacity", "Condition", "Owned / Leased", "Remark"],
          widths=[5, 30, 20, 14, 10, 14, 14, 14, 14], rows=20, landscape=True,
          note="Đính kèm đăng ký/kiểm định thiết bị; hợp đồng thuê nếu là thiết bị thuê.")
    wb.save(os.path.join(out, "HSDT-F08-Thiet-bi-thi-cong.xlsx"))
    made.append("HSDT-F08-Thiet-bi-thi-cong.xlsx")

    # ============ 4) NON-FULFILMENT 09A / 09B ============
    wb = new_wb()
    sheet(wb, "09A - Loi Nha thau", code="Form 09A",
          en="PREVIOUS NON-FULFILMENT DUE TO THE BIDDER’S FAULT",
          vn="Hợp đồng không hoàn thành do lỗi Nhà thầu",
          headers=["No.", "Contract name & number", "Employer / End user", "Value & duration",
                   "Reason for non-fulfilment", "Consequences / penalty", "Remark"],
          widths=[5, 26, 22, 18, 28, 24, 12], rows=8, landscape=True,
          note="Nếu không có, ghi “Not applicable / Không có”.")
    sheet(wb, "09B - Loi NT phu dac biet", code="Form 09B",
          en="PREVIOUS NON-FULFILMENT DUE TO SPECIAL SUB-CONTRACTOR’S FAULT",
          vn="Hợp đồng không hoàn thành do lỗi nhà thầu phụ đặc biệt",
          headers=["No.", "Contract name & number", "Special sub-contractor", "Employer / End user",
                   "Reason", "Consequences", "Remark"],
          widths=[5, 26, 22, 22, 24, 22, 12], rows=8, landscape=True)
    wb.save(os.path.join(out, "HSDT-F09A-F09B-Non-fulfilment.xlsx"))
    made.append("HSDT-F09A-F09B-Non-fulfilment.xlsx")

    # ============ 5) TÀI CHÍNH 10A1 / 10A2 / 10B / 11 ============
    wb = new_wb()
    sheet(wb, "10A1 - Tai chinh Nha thau", code="Form 10A1",
          en="FINANCIAL SITUATION OF BIDDER", vn="Tình hình tài chính của Nhà thầu",
          headers=["No.", "Description", "Unit", "Year 1", "Year 2", "Year 3"],
          widths=[5, 40, 12, 18, 18, 18], rows=8, num_cols=[4, 5, 6],
          note="Đính kèm báo cáo tài chính đã kiểm toán 03 năm gần nhất; tài sản ròng phải dương.")
    sheet(wb, "10A2 - Tai chinh NT phu", code="Form 10A2",
          en="FINANCIAL SITUATION OF SPECIAL SUB-CONTRACTOR",
          vn="Tình hình tài chính của nhà thầu phụ đặc biệt",
          headers=["No.", "Description", "Unit", "Year 1", "Year 2", "Year 3"],
          widths=[5, 40, 12, 18, 18, 18], rows=8, num_cols=[4, 5, 6])
    sheet(wb, "10B - Doanh thu BQ", code="Form 10B", en="AVERAGE ANNUAL REVENUE",
          vn="Doanh thu bình quân hằng năm",
          headers=["No.", "Annual turnover data for the last 3 years", "Amount (VND)"],
          widths=[5, 48, 22], rows=6, num_cols=[3],
          note="Ghi rõ giá trị tối thiểu theo BDS và so sánh.")
    sheet(wb, "11 - Nguon luc tai chinh", code="Form 11", en="FINANCIAL RESOURCES",
          vn="Nguồn lực tài chính (TFR – MFR)",
          headers=["No.", "Description", "Symbol", "Amount (VND)"],
          widths=[5, 52, 12, 22], rows=6, num_cols=[4],
          note="TFR = tổng nguồn lực tài chính; MFR = tổng nhu cầu tài chính tháng cho các hợp đồng "
               "đang thực hiện (Form 12). Nguồn lực khả dụng = TFR − MFR.")
    wb.save(os.path.join(out, "HSDT-F10-F11-Tai-chinh.xlsx"))
    made.append("HSDT-F10-F11-Tai-chinh.xlsx")

    # ============ 6) DANH MỤC 12 / 13 / 14 / 15 / 16 ============
    wb = new_wb()
    sheet(wb, "12 - Nguon luc theo thang", code="Form 12",
          en="MONTHLY FINANCIAL RESOURCES — CONTRACTS IN PROGRESS",
          vn="Nguồn lực tài chính theo tháng cho hợp đồng đang thực hiện",
          headers=["No.", "Contract’s name", "Contact person of the Employer", "Finish date of contract",
                   "Number of remaining months", "Unpaid contract value", "Monthly financial resources required"],
          widths=[5, 30, 24, 16, 16, 20, 22], rows=10, num_cols=[6, 7], landscape=True)
    sheet(wb, "13 - NT phu thuc hien", code="Form 13",
          en="WORKS ITEM PERFORMED BY SUBCONTRACTORS",
          vn="Khối lượng công việc do nhà thầu phụ thực hiện",
          headers=["No.", "Subcontractor’s name", "Work items", "Volume", "Estimated value",
                   "Contract/agreement content"],
          widths=[5, 26, 30, 14, 18, 28], rows=10, num_cols=[5], landscape=True)
    sheet(wb, "14 - NT phu dac biet", code="Form 14",
          en="LIST OF SPECIAL SUBCONTRACTORS", vn="Danh sách nhà thầu phụ đặc biệt",
          headers=["No.", "Name of special sub-contractor", "Work items", "Volume",
                   "Estimated value", "Contract/agreement content"],
          widths=[5, 26, 30, 14, 18, 28], rows=10, num_cols=[5], landscape=True)
    sheet(wb, "15 - Cong ty con/lien ket", code="Form 15",
          en="LIST OF SUBSIDIARY / ASSOCIATE COMPANY IN CHARGE OF PACKAGE WORKS",
          vn="Danh sách công ty con, công ty liên kết",
          headers=["No.", "Name of subsidiary/associate company", "Work performed in the Package",
                   "Proportion of value (%)", "Notes"],
          widths=[5, 32, 34, 16, 20], rows=10, num_cols=[4], landscape=True)
    sheet(wb, "16 - Tien do du an", code="Form 16",
          en="TIME SCHEDULE OF PROJECT", vn="Tiến độ thực hiện dự án",
          headers=["No.", "Task", "Time to start", "Time to finish", "Duration (days)", "Remark"],
          widths=[5, 46, 16, 16, 14, 20], rows=20, landscape=True,
          note="Đính kèm sơ đồ Gantt/đường cong tiến độ; nêu rõ đường găng và mốc bàn giao.")
    wb.save(os.path.join(out, "HSDT-F12-F16-Danh-muc.xlsx"))
    made.append("HSDT-F12-F16-Danh-muc.xlsx")

    # ============ 7) DEVIATION 18A / 18B ============
    wb = new_wb()
    sheet(wb, "18A - Sai khac ky thuat", code="Form 18A",
          en="TECHNICAL DEVIATION DECLARATION FORM", vn="Bảng kê sai khác kỹ thuật",
          headers=["Deviation No.", "Reference to Bidding Documents", "Reference to Bidder’s Proposal",
                   "Deviations/Reservations (technical)", "Remark"],
          widths=[12, 30, 30, 40, 16], rows=20, landscape=True,
          note="Nếu không có sai khác, ghi “No deviation / Không có sai khác”.")
    sheet(wb, "18B - Sai khac thuong mai", code="Form 18B",
          en="COMMERCIAL DEVIATION DECLARATION FORM", vn="Bảng kê sai khác thương mại",
          headers=["Deviation No.", "Reference to Bidding Documents", "Reference to Bidder’s Proposal",
                   "Deviations/Reservations (commercial)", "Remark"],
          widths=[12, 30, 30, 40, 16], rows=20, landscape=True)
    wb.save(os.path.join(out, "HSDT-F18A-F18B-Deviation.xlsx"))
    made.append("HSDT-F18A-F18B-Deviation.xlsx")

    # ============ 8) ATTACHMENT 1A / 1B / 2 ============
    wb = new_wb()
    sheet(wb, "ATT1A - Theo he thong", code="Attachment 1A",
          en="BIDDER’S PROPOSAL FOR MANUFACTURERS AND COUNTRY OF ORIGIN — SYSTEMS",
          vn="Kê khai nhà sản xuất và xuất xứ theo hệ thống",
          headers=["No.", "Item", "Manufacturer", "Country of Origin", "Remark"],
          widths=[5, 46, 30, 22, 20], rows=25, landscape=True, mode="NONE")
    sheet(wb, "ATT1B - Thiet bi chinh", code="Attachment 1B",
          en="BIDDER’S PROPOSAL FOR MANUFACTURERS AND COUNTRY OF ORIGIN — MAJOR EQUIPMENT",
          vn="Kê khai nhà sản xuất và xuất xứ thiết bị chính",
          headers=["No.", "Item", "Manufacturer", "Country of Origin", "Remark"],
          widths=[5, 46, 30, 22, 20], rows=35, landscape=True, mode="NONE")
    sheet(wb, "ATT2 - Thiet bi vat tu", code="Attachment 2",
          en="BIDDER’S PROPOSAL FOR EQUIPMENT AND MATERIALS OF THE PROJECT",
          vn="Kê khai thiết bị và vật tư của dự án (kèm model & thông số)",
          headers=["No.", "Item", "Manufacturer", "Country of Origin", "Mode, codes, labels/trademark",
                   "Technical specification", "Remark"],
          widths=[5, 30, 22, 16, 24, 30, 14], rows=40, landscape=True, mode="NONE")
    wb.save(os.path.join(out, "HSDT-ATT1A-1B-2-Manufacturers-Origin-Equipment.xlsx"))
    made.append("HSDT-ATT1A-1B-2-Manufacturers-Origin-Equipment.xlsx")

    # ============ 9) BIỂU GIÁ 20 / 21 / 22 / 23 ============
    wb = new_wb()
    sheet(wb, "20 - Grand Summary", code="Form 20", en="GRAND SUMMARY",
          vn="Bảng tổng hợp giá dự thầu",
          headers=["No.", "Description", "Reference in Schedule", "Amount (VND)",
                   "Amount (USD)", "Remark"],
          widths=[5, 40, 20, 22, 22, 16], rows=12, num_cols=[4, 5],
          preset_rows=[
              [1, "Supply of equipment and materials", "Schedule 1", None, None, ""],
              [2, "Local equipment, materials and construction works", "Schedule 2", None, None, ""],
              [3, "Erection, installation works", "Schedule 3", None, None, ""],
              [4, "Testing, commissioning and start-up", "Schedule 4", None, None, ""],
              [5, "Other costs (spare parts, training, documentation)", "Schedule 5", None, None, ""],
              [6, "Provisional sum (if any)", "—", None, None, ""],
              [7, "Value Added Tax (VAT)", "—", None, None, ""],
          ],
          formulas=[("D", True), ("E", True)],
          note="Grand Total = tổng các Schedule 1–3 (Form 21). Giá trị ghi rõ đã/chưa bao gồm VAT.")

    s1 = ["No.", "Descriptions", "Unit", "Quantity", "Cost (design and manufacture) — VND",
          "Freight and Insurance — USD", "Inland transportation — VND", "Total"]
    sheet(wb, "21.1 - Schedule 1", code="Form 21 — Schedule 1",
          en="PRICE SCHEDULE 1 — EQUIPMENT AND MATERIALS SUPPLIED FROM ABROAD",
          vn="Bảng giá 1 — Thiết bị và vật tư nhập khẩu",
          headers=s1, widths=[5, 40, 10, 12, 26, 22, 22, 22], rows=60, num_cols=[4, 5, 6, 7, 8],
          landscape=True, formulas=[("H", True)])
    s2 = ["No.", "Descriptions", "Unit", "Quantity", "Unit price — VND", "Total — VND"]
    sheet(wb, "21.2 - Schedule 2", code="Form 21 — Schedule 2",
          en="PRICE SCHEDULE 2 — LOCAL EQUIPMENT, MATERIALS AND CONSTRUCTION",
          vn="Bảng giá 2 — Thiết bị, vật tư trong nước và xây lắp",
          headers=s2, widths=[5, 52, 12, 14, 22, 24], rows=60, num_cols=[4, 5, 6],
          landscape=True, formulas=[("F", True)])
    s3 = ["No.", "Descriptions", "Unit", "Quantity", "Unit price — VND", "Total — VND"]
    sheet(wb, "21.3 - Schedule 3", code="Form 21 — Schedule 3",
          en="PRICE SCHEDULE 3 — ERECTION, TESTING AND COMMISSIONING",
          vn="Bảng giá 3 — Lắp đặt, chạy thử và nghiệm thu",
          headers=s3, widths=[5, 52, 12, 14, 22, 24], rows=40, num_cols=[4, 5, 6],
          landscape=True, formulas=[("F", True)])
    sheet(wb, "22 - Uu dai trong nuoc", code="Form 22",
          en="DOMESTIC COSTS ELIGIBLE FOR INCENTIVES", vn="Chi phí trong nước được hưởng ưu đãi",
          headers=["No.", "DESCRIPTION", "AMOUNT (VND)", "Remark"],
          widths=[5, 60, 24, 20], rows=20, num_cols=[3], landscape=True,
          formulas=[("C", True)],
          note="Nếu hàng hóa không thuộc đối tượng ưu đãi, nhà thầu không khai biểu mẫu này.")
    sheet(wb, "23 - Dieu chinh gia", code="Form 23",
          en="SCHEDULE OF ADJUSTMENT DATA", vn="Bảng dữ liệu điều chỉnh giá",
          headers=["No.", "Description", "Base value", "Index / source", "Weight (%)", "Remark"],
          widths=[5, 44, 18, 26, 14, 16], rows=10, num_cols=[3, 5], landscape=True,
          note="Thường áp dụng “Not applicable” nếu hợp đồng không điều chỉnh giá.")
    wb.save(os.path.join(out, "HSDT-F20-F23-Bieu-gia.xlsx"))
    made.append("HSDT-F20-F23-Bieu-gia.xlsx")
    return made


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else (DEFAULT_OUT or _default_out())
    for f in build(out):
        print(f"  ✓ {f}")
    print("Xong.")


if __name__ == "__main__":
    main()
