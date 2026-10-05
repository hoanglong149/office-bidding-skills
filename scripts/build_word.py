#!/usr/bin/env python3
"""
hsdt_build_templates.py — Sinh BỘ TEMPLATE WORD HSDT (36 biểu mẫu).

Chuẩn định dạng: Nghị định 30/2020/NĐ-CP
  A4 210x297 · lề trên/dưới 20mm · trái 30mm · phải 20mm
  Times New Roman 13pt · dãn dòng 1,2 · số trang giữa lề dưới
  Khối tiêu đề 2 cột 6,0/10,0 cm, co giãn ký tự 95% (không rớt dòng)

Danh mục biểu mẫu theo cấu trúc HSMT EPC tiêu chuẩn (Chapter IV — Bidding Forms):
  I.  TECHNICAL PROPOSAL  — Form 01 → Form 18B + Attachment 1A/1B/2
  II. FINANCIAL PROPOSAL  — Form 19a → Form 23

Chạy:  python3 hsdt_build_templates.py [thư_mục_đích]
"""

from __future__ import annotations
import os
import shutil
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

FONT = "Times New Roman"
SZ = Pt(13)
SZ_S = Pt(11)
SZ_FOOT = Pt(12)
LINE = 1.2
AFTER = Pt(6)
DEFAULT_OUT = None  # xem _default_out()

CENTER = WD_ALIGN_PARAGRAPH.CENTER
RIGHT = WD_ALIGN_PARAGRAPH.RIGHT
LEFT = WD_ALIGN_PARAGRAPH.LEFT
JUST = WD_ALIGN_PARAGRAPH.JUSTIFY


# ------------------------------------------------------------------ low level
def _bg(cell, hexcolor):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    cell._tc.get_or_add_tcPr().append(shd)


def _char_scale(run, pct=95):
    w = OxmlElement("w:w"); w.set(qn("w:val"), str(pct))
    run._r.get_or_add_rPr().append(w)


def _tbl_width(t, pct=5000):
    tblPr = t._tbl.tblPr
    w = OxmlElement("w:tblW"); w.set(qn("w:w"), str(pct)); w.set(qn("w:type"), "pct")
    tblPr.append(w)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)


def _borderless(t):
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "none"); b.append(e)
    t._tbl.tblPr.append(b)


def _page_field(par):
    for kind, txt in (("begin", None), (None, " PAGE "), ("separate", None),
                      (None, "1"), ("end", None)):
        r = par.add_run(); r.font.name = FONT; r.font.size = SZ_FOOT
        if kind:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), kind); r._r.append(fc)
        else:
            el = OxmlElement("w:instrText" if "PAGE" in str(txt) else "w:t")
            el.set(qn("xml:space"), "preserve"); el.text = txt; r._r.append(el)


def _run(par, text, *, bold=False, italic=False, size=SZ, caps=False, scale=None):
    r = par.add_run(text.upper() if caps else text)
    r.bold, r.italic = bold, italic
    r.font.name = FONT; r.font.size = size
    if scale:
        _char_scale(r, scale)
    return r


def setup_doc() -> Document:
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT; st.font.size = SZ
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    pf = st.paragraph_format
    pf.line_spacing = LINE; pf.space_after = AFTER; pf.space_before = Pt(0); pf.alignment = JUST
    s = doc.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.top_margin = s.bottom_margin = Mm(20)
    s.left_margin, s.right_margin = Mm(30), Mm(20)
    fp = s.footer.paragraphs[0]; fp.alignment = CENTER
    _page_field(fp)
    return doc


def p(doc, text="", *, bold=False, italic=False, size=SZ, align=None,
      after=AFTER, before=0, caps=False, scale=None):
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(before)
    par.paragraph_format.space_after = after
    par.paragraph_format.line_spacing = LINE
    if align:
        par.alignment = align
    if text:
        _run(par, text, bold=bold, italic=italic, size=size, caps=caps, scale=scale)
    return par


def bullets(doc, items, mark="- "):
    for it in items:
        p(doc, mark + it, after=Pt(2))


def table(doc, headers, rows, widths=None, header_fill="D9E1F2", size=SZ_S):
    n = len(headers)
    t = doc.add_table(rows=1, cols=n)
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False; _tbl_width(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ""
        par = c.paragraphs[0]; par.alignment = CENTER
        par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
        _run(par, h, bold=True, size=size)
        _bg(c, header_fill)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row[:n]):
            cells[i].text = ""
            par = cells[i].paragraphs[0]
            par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
            _run(par, str(v), size=size)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    _tbl_width(t)
    return t


def header_block(doc, *, code="", en="", vn="", mode="BIDDER",
                 ten_chu_the=None, so_ky_hieu=None, dia_danh=None, ngay=None):
    """Khối tiêu đề (letterhead) — ĐIỀN THEO TỪNG GÓI THẦU.

    mode:
      "BIDDER"   → letterhead của Nhà thầu (mặc định; dùng cho hầu hết biểu mẫu)
      "EMPLOYER" → letterhead của Chủ đầu tư / Bên mời thầu (vd Form 17a)
      "NONE"     → không có letterhead (bảng kê thuần / phụ lục)
    """
    if mode != "NONE":
        chu_the = ten_chu_the or ("[TÊN NHÀ THẦU / LIÊN DANH NHÀ THẦU]" if mode == "BIDDER"
                                  else "[TÊN CHỦ ĐẦU TƯ / BÊN MỜI THẦU]")
        so = so_ky_hieu or "……/……-HSDT"
        dd = dia_danh or "[Địa danh]"
        ng = ngay or "ngày …… tháng …… năm 20……"
        t = doc.add_table(rows=2, cols=2)
        t.autofit = False; _borderless(t); _tbl_width(t)
        c = t.rows[0].cells
        par = c[0].paragraphs[0]; par.alignment = CENTER
        par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
        _run(par, chu_the, bold=True, size=Pt(12), scale=95)
        par = c[1].paragraphs[0]; par.alignment = CENTER
        par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
        _run(par, "[Địa chỉ · Điện thoại · Email]" if mode == "BIDDER"
             else "[Địa chỉ · Điện thoại · Fax]", size=Pt(11), scale=95)
        c = t.rows[1].cells
        par = c[0].paragraphs[0]; par.alignment = CENTER
        par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
        _run(par, f"Số: {so}", size=Pt(12), scale=95)
        par = c[1].paragraphs[0]; par.alignment = CENTER
        par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
        _run(par, f"{dd}, {ng}", italic=True, size=Pt(12), scale=95)
        for cell in (t.rows[1].cells[0], t.rows[1].cells[1]):
            par = cell.add_paragraph(); par.alignment = CENTER
            par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
            _run(par, "——————————", size=Pt(9))
        for row in t.rows:
            row.cells[0].width = Cm(6.0); row.cells[1].width = Cm(10.0)
        p(doc, "", after=Pt(6))
    if code:
        p(doc, code, bold=True, size=Pt(13), align=RIGHT, after=Pt(0))
    if en:
        p(doc, en, bold=True, size=Pt(14), align=CENTER, after=Pt(2))
    if vn:
        p(doc, vn, italic=True, size=Pt(12), align=CENTER, after=Pt(10))


def signature(doc, left_label="ĐẠI DIỆN HỢP PHÁP CỦA NHÀ THẦU",
              right_label="", note="(Ký, ghi rõ họ tên, chức vụ và đóng dấu)"):
    p(doc, "", after=Pt(10))
    if right_label:
        t = table(doc, [left_label, right_label],
                  [[note, "(Ký, ghi rõ họ tên, đóng dấu)"], ["", ""], ["", ""]],
                  widths=[8.0, 8.0])
    else:
        p(doc, left_label, bold=True, align=CENTER, after=Pt(2))
        p(doc, note, italic=True, size=SZ_S, align=CENTER, after=Pt(36))
        p(doc, "[HỌ VÀ TÊN · CHỨC VỤ]", bold=True, align=CENTER)


def guide(doc, items, title="Hướng dẫn điền"):
    p(doc, title + ":", bold=True, size=SZ_S, after=Pt(2))
    bullets(doc, items)


def kv(doc, rows, w1=6.5, w2=9.5,
       h1="Item / Nội dung", h2="Bidder’s declaration / Khai báo của Nhà thầu"):
    return table(doc, [h1, h2], rows, widths=[w1, w2])


# ------------------------------------------------------------------ 36 forms
def build():
    F = []

    # ============ I. TECHNICAL PROPOSAL ============
    F.append(dict(code="Form 01", en="LETTER OF BID (Technical Proposal)",
                  vn="Đơn dự thầu — Hồ sơ đề xuất kỹ thuật", f="HSDT-F01-Letter-of-Bid-Technical.docx",
                  body=lambda d: (
                      p(d, "Date: ……/……/20……", after=Pt(2)),
                      p(d, "Package name: [tên gói thầu theo BDS 1.2]", after=Pt(2)),
                      p(d, "Project name: [tên dự án theo BDS 1.2]", after=Pt(2)),
                      p(d, "To: [Employer’s Representative theo ITB 1.1]", after=Pt(8)),
                      p(d, "After carefully examining the Bidding Documents and Addendum No. ……, we, the undersigned, "
                           "offer to perform the Package in accordance with the Bidding Documents.", after=Pt(6)),
                      kv(d, [["Validity of the bid (days)", "[≥ …… ngày kể từ ngày đóng thầu]"],
                             ["Bid security (value/currency)", "[…… đồng / USD]"],
                             ["Validity of Bid Security", "[…… ngày kể từ ngày đóng thầu]"],
                             ["Bid price (technical proposal)", "Không nêu giá ở phần kỹ thuật"]]),
                      p(d, "We hereby declare that:", bold=True, before=Pt(8), after=Pt(2)),
                      bullets(d, [
                          "We only participate in this Bid as the Contractor;",
                          "Not undergoing dissolution/bankruptcy procedures; valid business registration;",
                          "We commit no violations against regulations on assurance of competitiveness;",
                          "We have fulfilled tax declaration and payment obligations of the latest fiscal year;",
                          "We are not banned from bidding and not under criminal prosecution;",
                          "Within 03 years before the bidding deadline, no personnel have been convicted;",
                          "We are not involved in corruption, bribery, bid rigging or obstruction;",
                          "Every information provided herein is truthful to the best of our knowledge.",
                      ]),
                      signature(d),
                  )))

    F.append(dict(code="Form 02", en="POWER OF ATTORNEY",
                  vn="Giấy ủy quyền", f="HSDT-F02-Power-of-Attorney.docx",
                  body=lambda d: (
                      p(d, "[Location and date] ……/……/20……", align=RIGHT, after=Pt(8)),
                      p(d, "I am [full name, ID/passport number, position], the legal representative of "
                           "[Bidder’s full name], hereby authorize:", after=Pt(6)),
                      kv(d, [["Authorized person (full name)", "[……]"],
                             ["ID / Passport No.", "[……]"],
                             ["Position", "[……]"],
                             ["Authorizer (full name)", "[Nhà thầu — người ủy quyền]"],
                             ["Position of Authorizer", "[Tổng Giám đốc / Giám đốc]"],
                             ["Effective from", "[……/……/20……]"],
                             ["Effective to", "[……/……/20……]"],
                             ["Number of copies", "[…… bản]"]]),
                      p(d, "Scope of authorization:", bold=True, before=Pt(8), after=Pt(2)),
                      bullets(d, [
                          "Sign the Letter of Technical Proposal and Letter of Financial Proposal;",
                          "Sign the Consortium/Joint Venture agreement (if any);",
                          "Sign documents with the Employer’s Representative during bid selection, incl. clarifications;",
                          "Participate in contract negotiation and finalization;",
                          "Sign complaint letter (if any);",
                          "Sign contract with the Employer if successful.",
                      ]),
                      signature(d, "AUTHORIZER (Nhà thầu ủy quyền)", "AUTHORIZED PERSON (Người được ủy quyền)"),
                  )))

    F.append(dict(code="Form 03", en="CONSORTIUM / JOINT VENTURE AGREEMENT",
                  vn="Thỏa thuận liên danh", f="HSDT-F03-Consortium-JV-Agreement.docx",
                  body=lambda d: (
                      p(d, "[Location and date] ……/……/20……", align=RIGHT, after=Pt(6)),
                      p(d, "Package: [tên gói thầu] — Project: [tên dự án]", after=Pt(6)),
                      p(d, "Pursuant to Law on Bidding No. 22/2023/QH15;", after=Pt(2)),
                      p(d, "Pursuant to Decree No. 214/2025/ND-CP on guidelines for the Law on Bidding, as amended;",
                        after=Pt(8)),
                      p(d, "Representatives signing the Consortium/Joint Venture agreement:", bold=True, after=Pt(4)),
                      table(d, ["No.", "Names of members", "Tasks", "Proportion of value to Package (%)"],
                            [["1", "[Nhà thầu đứng đầu]", "[Phần việc đảm nhận]", "[…… %]"],
                             ["2", "[Thành viên]", "[Phần việc đảm nhận]", "[…… %]"]],
                            widths=[1.2, 5.6, 5.4, 3.8]),
                      p(d, "Legal representative of the Consortium/JV:", bold=True, before=Pt(8), after=Pt(4)),
                      kv(d, [["Full name", "[……]"], ["Position", "[……]"],
                             ["Authorization scope", "[Ký HSDT, ký hợp đồng, các thủ tục liên quan]"],
                             ["Power of attorney attached", "[Có / Không]"]]),
                      signature(d, "REPRESENTATIVE OF LEADER MEMBER", "REPRESENTATIVE OF MEMBER"),
                  )))

    F.append(dict(code="Form 04", en="BID SECURITY",
                  vn="Bảo đảm dự thầu", f="HSDT-F04-Bid-Security.docx",
                  body=lambda d: (
                      kv(d, [["Package name", "[……]"], ["Bidder’s name", "[……]"],
                             ["Form of bid security", "[Bảo lãnh ngân hàng / Thư bảo lãnh / Đặt cọc]"],
                             ["Bid security value", "[…… đồng / USD (bằng chữ: ……)]"],
                             ["Percentage of bid price", "[…… %]"],
                             ["Issuing bank", "[Tên ngân hàng — Chi nhánh]"],
                             ["Security No.", "[……]"], ["Issued date", "[……/……/20……]"],
                             ["Valid until", "[≥ Bid closing date + …… ngày]"],
                             ["Documents attached", "[Bản gốc …… tờ; bản chụp …… tờ]"]]),
                      signature(d),
                  )))

    F.append(dict(code="Form 05(a)", en="BIDDER INFORMATION",
                  vn="Thông tin nhà thầu", f="HSDT-F05a-Bidder-Information.docx",
                  body=lambda d: (
                      kv(d, [["Bidder’s name (full)", "[……]"], ["International trading name", "[……]"],
                             ["Tax code / Business registration No.", "[……]"], ["Date/Place of issue", "[……]"],
                             ["Head office address", "[……]"], ["Telephone / Fax / Email / Website", "[……]"],
                             ["Legal representative", "[Họ tên · Chức vụ]"], ["Bank account", "[……]"],
                             ["Type of enterprise", "[CP / TNHH / …]"], ["Charter capital", "[……] đồng"],
                             ["Number of employees / engineers", "[…… / ……]"],
                             ["Construction capacity certificate", "[Số … cấp ngày …]"],
                             ["ISO certificates", "[9001 / 14001 / 45001 — số, hiệu lực]"],
                             ["Joint venture member", "[Không / Có — tên thành viên]"]], w1=7.0, w2=9.0),
                      signature(d),
                  )))

    F.append(dict(code="Form 05(b)", en="CONSORTIUM / JOINT VENTURE PARTY INFORMATION FORM",
                  vn="Thông tin từng thành viên liên danh", f="HSDT-F05b-JV-Party-Information.docx",
                  body=lambda d: (
                      p(d, "Each Consortium/Joint Venture member shall declare information using this Form.", italic=True,
                        size=SZ_S, after=Pt(8)),
                      kv(d, [["Member’s name", "[……]"], ["Role in JV", "[Leader / Member]"],
                             ["Tax code", "[……]"], ["Address", "[……]"],
                             ["Legal representative", "[……]"], ["Scope of works undertaken", "[……]"],
                             ["Proportion of value (%)", "[……]"], ["Documents attached", "[……]"]], w1=7.0, w2=9.0),
                      signature(d),
                  )))

    similar_cols = ["Name and contract number", "Name of Employer / End user",
                    "Name of Contractor / Sub-contractor", "Value of contract",
                    "Duration & completion status", "Scope of works", "Remark"]
    sim_w = [3.4, 3.0, 3.0, 2.2, 2.6, 2.6, 1.2]
    F.append(dict(code="Form 07B", en="CURRICULUM VITAE OF PERSONNEL FOR KEY PERSON",
                  vn="Sơ yếu lý lịch nhân sự chủ chốt", f="HSDT-F07B-CV-Key-Personnel.docx",
                  body=lambda d: (
                      kv(d, [["Full name", "[……]"], ["Date / Place of birth", "[……]"],
                             ["Position proposed", "[……]"], ["Professional qualification", "[……]"],
                             ["Year of graduation / University", "[……]"],
                             ["Number of years of experience", "[……]"],
                             ["Certificates (No., validity)", "[……]"],
                             ["Employment confirmation", "[HĐLĐ / xác nhận của nhà thầu]"]], w1=6.0, w2=10.0),
                      p(d, "Summarize professional experience in reverse chronological order:", bold=True,
                        before=Pt(8), after=Pt(4)),
                      table(d, ["From (month/year)", "To (month/year)", "Employer", "Position",
                                "Project / Assignment", "Responsibilities"],
                            [["", "", "", "", "", ""] for _ in range(4)],
                            widths=[2.4, 2.4, 3.0, 2.6, 3.2, 2.4]),
                      signature(d),
                  )))

    F.append(dict(code="Form 17(a)", en="EMPLOYER / END USER’S CERTIFICATE",
                  vn="Xác nhận của Chủ đầu tư / Người sử dụng cuối cùng",
                  f="HSDT-F17a-Employer-Certificate.docx", head="EMPLOYER",
                  body=lambda d: (
                      kv(d, [["Employer / End user’s name", "[……]"], ["Address", "[……]"],
                             ["Contract name & number", "[……]"], ["Contract value", "[……]"],
                             ["Scope of works performed", "[……]"], ["Completion status", "[…… %]"],
                             ["Agreement", "[Xác nhận đồng ý cho nhà thầu sử dụng hợp đồng này làm kinh nghiệm]"]],
                         w1=7.0, w2=9.0),
                      signature(d, "EMPLOYER / END USER", ""),
                  )))

    F.append(dict(code="Form 17(b)", en="EMPLOYER / CONTRACTOR CERTIFICATE",
                  vn="Xác nhận của Chủ đầu tư / Nhà thầu chính",
                  f="HSDT-F17b-Contractor-Certificate.docx",
                  body=lambda d: (
                      kv(d, [["Issuing party", "[Chủ đầu tư / Nhà thầu chính]"], ["Contract name & number", "[……]"],
                             ["Sub-contract value", "[……]"], ["Works performed by sub-contractor", "[……]"],
                             ["Completion status", "[…… %]"]], w1=7.0, w2=9.0),
                      signature(d),
                  )))

    # ============ II. FINANCIAL PROPOSAL ============
    for c, ttl, vn, fn in (
        ("Form 19(a)", "LETTER OF BID (Financial Proposal)",
         "Đơn dự thầu — Hồ sơ đề xuất tài chính (a)", "HSDT-F19a-Letter-of-Bid-Financial.docx"),
        ("Form 19(b)", "LETTER OF BID (Financial Proposal) — alternative",
         "Đơn dự thầu — Hồ sơ đề xuất tài chính (b)", "HSDT-F19b-Letter-of-Bid-Financial-Alt.docx"),
    ):
        F.append(dict(code=c, en=ttl, vn=vn, f=fn, body=lambda d: (
            p(d, "Date: ……/……/20…… — Package: [tên gói thầu]", after=Pt(6)),
            p(d, "To: [Employer’s Representative]", after=Pt(8)),
            p(d, "We offer to perform the Package in accordance with the Bidding Documents at the following price:",
              after=Pt(6)),
            kv(d, [["Total bid price (exclusive of VAT)", "[……] đồng (in words: ……)"],
                   ["Total bid price (inclusive of VAT)", "[……] đồng (in words: ……)"],
                   ["Currency", "[VND / USD]"],
                   ["Validity of the bid", "[≥ …… ngày kể từ ngày đóng thầu]"],
                   ["Bid security", "[……]"], ["Contract duration", "[…… ngày]"]], w1=7.0, w2=9.0),
            signature(d),
        )))

    

    

    

    

    return F


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else (DEFAULT_OUT or _default_out())
    os.makedirs(out, exist_ok=True)
    # dọn bộ template cũ (19 file không khớp danh mục HSMT)
    for old in os.listdir(out):
        if old.startswith("HSDT-") and old.endswith(".docx") and "-F" not in old and "ATT" not in old:
            os.remove(os.path.join(out, old)); print(f"  − xoá template cũ: {old}")
    forms = build()
    print(f"Sinh {len(forms)} biểu mẫu HSDT → {out}")
    for fm in forms:
        doc = setup_doc()
        header_block(doc, code=fm["code"], en=fm["en"], vn=fm["vn"],
                     mode=fm.get("head", "BIDDER"))
        fm["body"](doc)
        doc.save(os.path.join(out, fm["f"]))
        print(f"  ✓ {fm['f']}")
    print("Xong.")


if __name__ == "__main__":
    main()
