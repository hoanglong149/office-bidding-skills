#!/usr/bin/env python3
"""
apply_project_format.py — Đóng dấu **letterhead tên dự án** và áp định dạng bảng/ô
cho một bộ biểu mẫu HSDT (.docx), theo đúng kiểu file HSDT thực tế đang dùng.

Letterhead (chèn vào `word/header1.xml`):
    ┌──────────────────────────┬──────────┬──────────┐
    │ Project : <tên dự án>    │  logo 1  │  logo 2  │   bảng KHÔNG viền, 3 cột
    │ Bidding Package : <gói>  │          │          │   5670 / 2381 / 1701 dxa
    └──────────────────────────┴──────────┴──────────┘
Nội dung cũ trong header (vd. field PAGE của TT79) được **giữ nguyên**, letterhead chèn lên trên.

Profile định dạng:
    keep      — chỉ chèn letterhead, không đổi gì khác (mặc định)
    intl-epc  — theo HSDT quốc tế (file thật Phả Lại): khổ Letter 216×279 mm, lề 30/20,
                bảng viền single 0.5 pt, tblLayout fixed, ô canh giữa dọc, chữ ô TNR 13 pt,
                hàng tiêu đề in đậm + canh giữa

Dùng:
    python3 apply_project_format.py <thư_mục_hoặc_file...> \\
        --project "Retrofit and upgrading of ..." \\
        --package "Engineering, Procurement and Construction (EPC) of ..." \\
        [--logo-left logo1.png --logo-right logo2.png] \\
        [--profile intl-epc] [--out <thư_mục_ra>] [--suffix "-PL"]

Không truyền `--out` thì sửa tại chỗ và tạo `.bak` bên cạnh.
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
import zipfile

from lxml import etree

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Mm, Twips

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# --- thông số lấy đúng từ file HSDT thật (Phả Lại — HSDT quốc tế) -------------
LETTERHEAD_COLS = (5670, 2381, 1701)      # dxa
LETTERHEAD_W = 9752
CELL_FONT = "Times New Roman"
CELL_SIZE = 13.0                          # pt
MARGINS = dict(top=1134, right=1134, bottom=1134, left=1701, header=635, footer=845)
PAPERS = {"a4": (11907, 16839), "letter": (12240, 15840)}   # twips — mặc định A4


# --------------------------------------------------------------- letterhead
def _run(text, *, italic=False, underline=False, size_cs=18, font="Times New Roman"):
    rpr = f'<w:rPr {nsdecls("w")}>'
    if font:
        rpr += f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    if italic:
        rpr += "<w:i/>"
    if underline:
        rpr += '<w:u w:val="single"/>'
    rpr += f'<w:szCs w:val="{size_cs}"/></w:rPr>'
    return (f'<w:r>{rpr}<w:t xml:space="preserve">{_esc(text)}</w:t></w:r>')


def _esc(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _cell(col_w: int, body: str) -> str:
    return (f'<w:tc><w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/><w:vAlign w:val="center"/>'
            f'</w:tcPr>{body}</w:tc>')


def _para(content: str) -> str:
    return (f'<w:p><w:pPr><w:pStyle w:val="Header"/><w:spacing w:after="0"/>'
            f'<w:rPr><w:i/><w:szCs w:val="18"/></w:rPr></w:pPr>{content}</w:p>')


def _png_size(path: str) -> tuple[int, int]:
    """(rộng, cao) pixel — đọc thẳng IHDR của PNG, không cần thư viện ảnh."""
    with open(path, "rb") as f:
        head = f.read(24)
    if head[:8] == b"\x89PNG\r\n\x1a\n":
        return int.from_bytes(head[16:20], "big"), int.from_bytes(head[20:24], "big")
    return 0, 0                                  # không phải PNG (jpg/...): đoán tỉ lệ 1:1


def _logo_emu(path: str, max_h_mm: float = 18.0, max_w_mm: float = 36.0) -> tuple[int, int]:
    """Kích thước hiển thị logo (EMU) giữ đúng tỉ lệ, không tràn ô."""
    w_px, h_px = _png_size(path)
    ratio = (w_px / h_px) if (w_px and h_px) else 1.0
    h_mm = max_h_mm
    w_mm = h_mm * ratio
    if w_mm > max_w_mm:
        w_mm, h_mm = max_w_mm, max_w_mm / ratio
    return int(w_mm * 36000), int(h_mm * 36000)   # 1 mm = 36 000 EMU


def letterhead_xml(project: str, package: str, logo_rels: list[tuple[str, str]],
                   logo_h_mm: float = 18.0) -> str:
    """Bảng letterhead 3 cột: nội dung dự án | logo | logo (bỏ cột logo nếu không có ảnh)."""
    cols = list(LETTERHEAD_COLS)
    texts = [_para(_run("Project", italic=True, underline=True)
                   + _run(" : " + project, italic=True)),
             _para(_run("Bidding Package", italic=True, underline=True)
                   + _run(" : " + package, italic=True))]
    media = ""
    for i, (rid, lp) in enumerate(logo_rels[:2]):
        cx, cy = _logo_emu(lp, logo_h_mm)
        media += (f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>'
                  f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
                  f'<wp:extent cx="{cx}" cy="{cy}"/>'
                  f'<wp:docPr id="{100 + i}" name="Logo{i + 1}"/>'
                  f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                  f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                  f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                  f'<pic:nvPicPr><pic:cNvPr id="{100 + i}" name="Logo{i + 1}"/><pic:cNvPicPr/></pic:nvPicPr>'
                  f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                  f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                  f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                  f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
    ncol = len(logo_rels[:2]) + 1
    used = cols[:ncol]
    grid = "".join(f'<w:gridCol w:w="{c}"/>' for c in used)
    cells = _cell(used[0], "".join(texts))
    for i in range(ncol - 1):
        cells += _cell(used[1 + i], media)
    return (
        '<w:tbl xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" {nsdecls("w")}>'
        f'<w:tblPr><w:tblW w:w="{LETTERHEAD_W}" w:type="dxa"/>'
        '<w:tblBorders>' + "".join(
            f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            for e in ("top", "left", "bottom", "right", "insideH", "insideV")) +
        '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'<w:tr><w:trPr><w:trHeight w:val="361"/></w:trPr>{cells}</w:tr></w:tbl>'
        + '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>')


# --------------------------------------------------------------- profile
def apply_profile(doc: Document, profile: str, paper: str = "a4") -> int:
    """Áp định dạng trang + bảng/ô. Trả về số bảng đã xử lý."""
    if profile == "intl-epc":
        s = doc.sections[0]
        w, h = PAPERS[paper]
        s.page_width, s.page_height = Twips(w), Twips(h)
        s.top_margin, s.bottom_margin = Twips(MARGINS["top"]), Twips(MARGINS["bottom"])
        s.left_margin, s.right_margin = Twips(MARGINS["left"]), Twips(MARGINS["right"])
        s.header_distance, s.footer_distance = Twips(MARGINS["header"]), Twips(MARGINS["footer"])
    n = 0
    if profile in ("intl-epc",):
        for t in doc.tables:
            tblPr = t._tbl.tblPr
            for old in tblPr.findall(qn("w:tblBorders")):
                tblPr.remove(old)
            b = OxmlElement("w:tblBorders")
            for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
                el = OxmlElement(f"w:{e}")
                el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "4")
                el.set(qn("w:space"), "0"); el.set(qn("w:color"), "auto")
                b.append(el)
            tblPr.append(b)
            for old in tblPr.findall(qn("w:tblLayout")):
                tblPr.remove(old)
            lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
            for r_i, row in enumerate(t.rows):
                for c in row.cells:
                    tcPr = c._tc.get_or_add_tcPr()
                    for old in tcPr.findall(qn("w:vAlign")):
                        tcPr.remove(old)
                    va = OxmlElement("w:vAlign"); va.set(qn("w:val"), "center"); tcPr.append(va)
                    for p in c.paragraphs:
                        if r_i == 0:
                            p.alignment = 1          # CENTER
                        for run in p.runs:
                            run.font.name = CELL_FONT
                            run.font.size = None
                            rPr = run._r.get_or_add_rPr()
                            for old in rPr.findall(qn("w:sz")) + rPr.findall(qn("w:szCs")):
                                rPr.remove(old)
                            for tag in ("w:sz", "w:szCs"):
                                el = OxmlElement(tag); el.set(qn("w:val"), str(int(CELL_SIZE * 2)))
                                rPr.append(el)
                            rF = rPr.find(qn("w:rFonts"))
                            if rF is None:
                                rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
                            for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                                rF.set(qn(a), CELL_FONT)
            n += 1
    return n


# --------------------------------------------------------------- excel
def stamp_excel(path: str, project: str, package: str, page_number: str = "footer") -> None:
    """Ghi tên dự án / gói thầu vào **print header** của mọi sheet (bản in Excel)."""
    import openpyxl

    wb = openpyxl.load_workbook(path)
    # KHÔNG nhúng mã font/size vào text — openpyxl tự thêm từ .font/.size (nhúng sẽ bị lặp)
    head = f"Project : {project}\nBidding Package : {package}"
    for ws in wb.worksheets:
        ws.oddHeader.left.text = head
        ws.oddHeader.left.size = 9
        ws.oddHeader.left.font = "Times New Roman,Italic"
        if page_number == "footer":
            ws.oddFooter.center.text = "Trang &P/&N"  # &P = số trang, &N = tổng số trang
            ws.oddFooter.center.size = 9
    wb.save(path)


# --------------------------------------------------------------- zip surgery
FOOTER_PART = "word/footerPage.xml"
FOOTER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

FOOTER_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:ftr {nsdecls("w")}><w:p><w:pPr><w:jc w:val="center"/>'
    '<w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
    '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
    '<w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
    '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
    '<w:t>1</w:t></w:r>'
    '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
    '</w:p></w:ftr>'
)


def _next_rid(rel_xml: str, base: str) -> str:
    used = set(re.findall(r'Id="([^"]+)"', rel_xml))
    rid = base
    while rid in used:
        rid += "x"
    return rid


def _strip_header_page_field(hdr_xml: str) -> str:
    """Bỏ các đoạn chứa field PAGE trong header — để số trang chỉ còn ở footer.

    Phải thao tác bằng lxml: cắt bằng regex sẽ làm đứt cấu trúc `<w:tbl>` của letterhead.
    """
    root = etree.fromstring(hdr_xml.encode("utf-8"))
    for p in list(root.iter(f"{{{W}}}p")):
        if any("PAGE" in (t.text or "") for t in p.iter(f"{{{W}}}instrText")):
            parent = p.getparent()
            if parent is not None:
                parent.remove(p)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf-8")


def set_page_number(data: dict, where: str) -> None:
    """`footer` (mặc định) — đưa số trang xuống chân trang; `header` — giữ ở đầu trang; `keep` — không đổi."""
    drels, ddoc = "word/_rels/document.xml.rels", "word/document.xml"
    if where == "keep":
        return
    rel_xml = data[drels].decode("utf-8")
    doc_xml = data[ddoc].decode("utf-8")

    if where == "header":
        if "<w:headerReference" not in doc_xml and "word/header1.xml" in data:
            rid = _next_rid(rel_xml, "rIdHdr1")
            rel_xml = rel_xml.replace(
                "</Relationships>",
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/'
                f'relationships/header" Target="header1.xml"/></Relationships>')
            doc_xml = re.sub(r"(<w:sectPr\b[^>]*>)",
                             f'\\1<w:headerReference w:type="default" r:id="{rid}"/>', doc_xml, count=1)
            data[drels], data[ddoc] = rel_xml.encode("utf-8"), doc_xml.encode("utf-8")
        return

    # ---- footer ----
    data[FOOTER_PART] = FOOTER_XML.encode("utf-8")
    rid = _next_rid(rel_xml, "rIdFtrPage")
    rel_xml = rel_xml.replace("</Relationships>",
                              f'<Relationship Id="{rid}" Type="{REL_FOOTER}" '
                              f'Target="footerPage.xml"/></Relationships>')
    data[drels] = rel_xml.encode("utf-8")

    ct = data["[Content_Types].xml"].decode("utf-8")
    if "/word/footerPage.xml" not in ct:
        ct = ct.replace("</Types>",
                        f'<Override PartName="/word/footerPage.xml" ContentType="{FOOTER_CT}"/></Types>')
        data["[Content_Types].xml"] = ct.encode("utf-8")

    doc_xml = re.sub(r"<w:footerReference[^>]*/>", "", doc_xml)
    doc_xml = re.sub(r"(<w:sectPr\b[^>]*>)",
                     f'\\1<w:footerReference w:type="default" r:id="{rid}"/>', doc_xml, count=1)
    data[ddoc] = doc_xml.encode("utf-8")

    hdr = "word/header1.xml"
    if hdr in data:
        data[hdr] = _strip_header_page_field(data[hdr].decode("utf-8")).encode("utf-8")



def inject_letterhead(path: str, project: str, package: str, logos: list[str],
                      logo_h_mm: float = 18.0, page_number: str = "footer") -> None:
    """Chèn bảng letterhead vào đầu header mặc định; thêm ảnh logo nếu có."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}

    # header mặc định
    hdr = "word/header1.xml"
    if hdr not in data:
        raise SystemExit(f"  ! {os.path.basename(path)}: không có {hdr} — bỏ qua")
    hrels = "word/_rels/header1.xml.rels"

    logo_rels = []
    if logos:
        rels = data.get(hrels, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                              b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                              b'relationships"></Relationships>')
        rel_xml = rels.decode("utf-8")
        used_ids = set(re.findall(r'Id="([^"]+)"', rel_xml))
        existing_media = [n for n in data if n.startswith("word/media/")]
        for i, lp in enumerate(logos[:2]):
            ext = os.path.splitext(lp)[1].lower().lstrip(".") or "png"
            mname = f"logo{len(existing_media) + i + 1}.{ext}"
            rid = f"rIdLogo{i + 1}"
            while rid in used_ids:
                rid += "x"
            used_ids.add(rid)
            data[f"word/media/{mname}"] = open(lp, "rb").read()
            rel_xml = rel_xml.replace("</Relationships>",
                                      f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                                      f'officeDocument/2006/relationships/image" Target="media/{mname}"/>'
                                      "</Relationships>")
            logo_rels.append((rid, lp))
        data[hrels] = rel_xml.encode("utf-8")

        # Content_Types: bảo đảm có default cho phần mở rộng ảnh
        ct = data["[Content_Types].xml"].decode("utf-8")
        for ext in {os.path.splitext(p)[1].lower().lstrip(".") or "png" for p in logos[:2]}:
            if f'Extension="{ext}"' not in ct:
                ct = ct.replace("<Types ", "<Types ", 1).replace(
                    "</Types>", f'<Default Extension="{ext}" ContentType="image/{ext}"/></Types>')
        data["[Content_Types].xml"] = ct.encode("utf-8")

    # bảo đảm sectPr THAM CHIẾU header — nhiều file TT79 có part header nhưng không reference
    drels, ddoc = "word/_rels/document.xml.rels", "word/document.xml"
    rel_xml = data[drels].decode("utf-8")
    rid = None
    for m in re.finditer(r'<Relationship\b[^>]*/>', rel_xml):
        tag = m.group(0)
        if "/header" in tag and 'Target="header1.xml"' in tag:
            rid = re.search(r'Id="([^"]+)"', tag).group(1)
            break
    if rid is None:
        used = set(re.findall(r'Id="([^"]+)"', rel_xml))
        rid = "rIdHdr1"
        while rid in used:
            rid += "x"
        rel_xml = rel_xml.replace("</Relationships>",
                                  f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                                  f'officeDocument/2006/relationships/header" Target="header1.xml"/>'
                                  "</Relationships>")
        data[drels] = rel_xml.encode("utf-8")
    doc_xml = data[ddoc].decode("utf-8")
    if "<w:headerReference" not in doc_xml:
        doc_xml = re.sub(r"(<w:sectPr\b[^>]*>)",
                         f'\\1<w:headerReference w:type="default" r:id="{rid}"/>', doc_xml, count=1)
        data[ddoc] = doc_xml.encode("utf-8")

    x = data[hdr].decode("utf-8")
    if "Project" in x and "Bidding Package" in x:
        return                                    # đã có letterhead
    # giá trị cache của field PAGE trong file gốc là số trang của HSMT (vd. 191) — đưa về 1
    x = re.sub(r"(<w:fldChar w:fldCharType=\"separate\"/>(?:<[^>]+>)*?<w:t[^>]*>)\d+(</w:t>)",
               r"\g<1>1\g<2>", x)
    tbl = letterhead_xml(project, package, logo_rels, logo_h_mm)
    m = re.search(r"(<w:hdr\b[^>]*>)", x)
    x = x[:m.end()] + tbl + x[m.end():]
    data[hdr] = x.encode("utf-8")

    set_page_number(data, page_number)

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in list(names) + [k for k in data if k not in names]:
            zo.writestr(n, data[n])
    os.replace(tmp, path)


# --------------------------------------------------------------- main
def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("targets", nargs="+", help="thư mục hoặc file .docx")
    ap.add_argument("--config", default=None,
                    help="file JSON của gói thầu: {project, package, logo_left, logo_right, paper, profile, logo_height, page_number}")
    ap.add_argument("--project", default=None)
    ap.add_argument("--package", default=None)
    ap.add_argument("--logo-left", default=None)
    ap.add_argument("--logo-right", default=None)
    ap.add_argument("--profile", choices=["keep", "intl-epc"], default=None)
    ap.add_argument("--paper", choices=["a4", "letter"], default=None, help="mặc định a4")
    ap.add_argument("--logo-height", type=float, default=None, help="chiều cao logo (mm), mặc định 18")
    ap.add_argument("--allow-placeholder", action="store_true",
                    help="cho phép tên còn [ ] — chỉ dùng khi tạo mẫu định dạng, KHÔNG để nộp thầu")
    ap.add_argument("--page-number", choices=["footer", "header", "keep"], default=None,
                    help="vị trí số trang — mặc định footer")
    ap.add_argument("--out", default=None, help="thư mục ra (mặc định: sửa tại chỗ + .bak)")
    ap.add_argument("--suffix", default="", help="hậu tố tên file khi dùng --out")
    a = ap.parse_args()

    cfg = {}
    if a.config:
        import json
        cfg = json.load(open(a.config, encoding="utf-8"))
        print(f"Nạp cấu hình gói thầu: {a.config}")
    project = a.project or cfg.get("project")
    package = a.package or cfg.get("package")
    if not (project and package):
        raise SystemExit("Thiếu tên dự án / gói thầu. Dùng --project/--package "
                         "hoặc --config <file.json> (xem projects/_template.json).")
    if a.allow_placeholder:
        print("  (--allow-placeholder: giữ nguyên tên placeholder — dùng để tạo MẪU ĐỊNH DẠNG)")
    elif any(c in (project + package) for c in "[]"):
        raise SystemExit(
            "Tên dự án / gói thầu còn là placeholder — điền TÊN THẬT của gói thầu vào "
            "file cấu hình rồi chạy lại:\n"
            "  cp projects/_template.json projects/<mã-gói>.json   # rồi sửa project/package\n"
            f"  (đang có: project={project!r}, package={package!r})")
    logo_l = a.logo_left or cfg.get("logo_left")
    logo_r = a.logo_right or cfg.get("logo_right")
    paper = a.paper or cfg.get("paper", "a4")
    profile = a.profile or cfg.get("profile", "keep")
    logo_h = a.logo_height or cfg.get("logo_height", 18.0)
    page_number = a.page_number or cfg.get("page_number", "footer")

    docx: list[str] = []
    xlsx: list[str] = []
    for t in a.targets:
        if os.path.isdir(t):
            docx += sorted(glob.glob(os.path.join(t, "*.docx")))
            xlsx += sorted(glob.glob(os.path.join(t, "*.xlsx")))
        elif t.endswith(".xlsx"):
            xlsx.append(t)
        else:
            docx.append(t)

    logos = [p for p in (logo_l, logo_r) if p]

    for f in xlsx:
        dst = f
        if a.out:
            os.makedirs(a.out, exist_ok=True)
            dst = os.path.join(a.out, os.path.basename(f))
            shutil.copy2(f, dst)
        else:
            shutil.copy2(f, f + ".bak")
        stamp_excel(dst, project, package, page_number)
        print(f"   {os.path.basename(dst)}  (excel: print header)")

    for f in docx:
        dst = f
        if a.out:
            os.makedirs(a.out, exist_ok=True)
            base = os.path.splitext(os.path.basename(f))[0]
            dst = os.path.join(a.out, f"{base}{a.suffix}.docx")
            shutil.copy2(f, dst)
        else:
            shutil.copy2(f, f + ".bak")
        doc = Document(dst)
        n = apply_profile(doc, profile, paper)
        doc.save(dst)
        inject_letterhead(dst, project, package, logos, logo_h, page_number)
        print(f"   {os.path.basename(dst)}  (profile={profile}/{paper}, số trang={page_number}, bảng={n}, logo={len(logos)})")
    print(f"\n{len(docx)} file Word + {len(xlsx)} file Excel: {a.out or 'sửa tại chỗ (.bak được tạo)'}")


if __name__ == "__main__":
    main()
