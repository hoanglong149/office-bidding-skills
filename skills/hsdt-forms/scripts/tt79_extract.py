#!/usr/bin/env python3
"""
tt79_extract.py — Tách bộ biểu mẫu dự thầu (HSDT) từ E-HSMT mẫu theo **Thông tư 79/2025/TT-BTC**.

Nguồn: file E-HSMT mẫu do Bộ Tài chính ban hành kèm TT79 (Mẫu số 3A/4A/5A/6A/7A/8A/9A/10A/10B/…),
Chương IV — "Biểu mẫu mời thầu và dự thầu". Script cắt **nguyên trạng** từng biểu mẫu của
nhà thầu ra file riêng, giữ 100 % định dạng gốc (Times New Roman 14, khổ A4, lề 20/20/30/20,
bảng biểu, gộp ô, ghi chú chân trang).

  • Word  (.docx) : 1 file / biểu mẫu — dùng để in, ký, scan, hoặc tham chiếu webform
  • Excel (.xlsx) : mỗi bảng trong biểu mẫu → 1 sheet, giữ nguyên tiêu đề cột của TT79
  • Index         : DANH-MUC-BIEU-MAU.md — mã mẫu, tên, số sheet, tên file

Cách dùng:
    python3 tt79_extract.py <E-HSMT-mau.docx> [thư_mục_ra]
    python3 tt79_extract.py "10. Mẫu số 10A_E-HSMT_EPC 01 túi.docx" out --only 02,03,06,08,10,11

Tùy chọn:
    --only 02,11,13.   chỉ tách các mẫu có mã khớp tiền tố (mặc định: tất cả)
    --no-excel         không xuất workbook Excel
    --keep-media       giữ toàn bộ ảnh gốc (mặc định xoá ảnh không dùng để giảm dung lượng)
"""

from __future__ import annotations

import argparse
import copy
import os
import re
import shutil
import unicodedata
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W, "r": R}

# Header biểu mẫu thật luôn có hậu tố "(Webform trên Hệ thống)" / "(Scan đính kèm)";
# dòng tiêu đề bìa ("Mẫu số 10A_E-HSMT_EPC 01 túi") không có hậu tố này nên bị loại.
MAU_RE = re.compile(r"^\s*Mẫu\s*(?:số|so)\s*([0-9]+(?:\.[0-9]+)?[A-Z]?)\s*\(")
CH4_BEGIN = re.compile(r"^\s*Chương\s+IV\b", re.I)
CH4_END = re.compile(r"^\s*(Chương\s+V\b|Phần\s+(?:2\b|thứ\s+hai\b))", re.I)
SKIP_NUM = {1}                      # Mẫu 01A–01E là biểu mẫu của Chủ đầu tư, không thuộc HSDT
BAD_TITLE = ("(", "[", "-", "Ghi chú", "Mẫu số", "Nhà thầu", "Liệt kê", "Đối với", "Điều ")


def slug(text: str, maxlen: int = 44) -> str:
    t = unicodedata.normalize("NFD", text)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn").replace("đ", "d").replace("Đ", "D")
    return (re.sub(r"[^A-Za-z0-9]+", "-", t).strip("-")[:maxlen].strip("-") or "bieu-mau")


def mau_number(code: str) -> int:
    m = re.match(r"(\d+)", code)
    return int(m.group(1)) if m else 0


def para_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(f"{{{W}}}t")).strip()


def is_bold(el) -> bool:
    return el.find(f".//{{{W}}}b", NS) is not None


def is_centered(el) -> bool:
    jc = el.find(f"./{{{W}}}pPr/{{{W}}}jc", NS)
    return jc is not None and jc.get(f"{{{W}}}val") in ("center", "both")


def table_rows(el):
    return el.findall(f"./{{{W}}}tr", NS)


def cell_positions(tr):
    """[(cột, text)] — giữ đúng cột theo `gridSpan` để ô không bị dồn sang trái."""
    col, out = 1, []
    for tc in tr.findall(f"./{{{W}}}tc", NS):
        txt = " ".join(para_text(p) for p in tc.findall(f"./{{{W}}}p", NS)).strip()
        gs = tc.find(f"./{{{W}}}tcPr/{{{W}}}gridSpan", NS)
        span = int(gs.get(f"{{{W}}}val")) if gs is not None and gs.get(f"{{{W}}}val") else 1
        if txt:
            out.append((col, txt))
        col += span
    return out


def cell_texts(tr):
    out = []
    for tc in tr.findall(f"./{{{W}}}tc", NS):
        txt = " ".join(para_text(p) for p in tc.findall(f"./{{{W}}}p", NS)).strip()
        if txt and (not out or out[-1] != txt):
            out.append(txt)
    return out


# ------------------------------------------------------------------ source parse
class Source:
    def __init__(self, path: str):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.names = z.namelist()
            self.xml = z.read("word/document.xml")
        self.root = etree.fromstring(self.xml)
        self.body = self.root.find(f"{{{W}}}body")
        self.sectpr = self.body.find(f"{{{W}}}sectPr")
        self.blocks = []                    # [(element, kind, text)] — KHÔNG gồm sectPr
        for el in self.body:
            tag = el.tag.split("}")[-1]
            if tag == "p":
                self.blocks.append((el, "p", para_text(el)))
            elif tag == "tbl":
                self.blocks.append((el, "tbl", ""))

    def window(self):
        """Cửa sổ Chương IV — chọn occurrence chứa nhiều header biểu mẫu nhất.

        'Chương IV' xuất hiện cả ở MỤC LỤC; mục lục không có header dạng 'Mẫu số X ('.
        """
        begins = [i for i, (_e, k, t) in enumerate(self.blocks) if k == "p" and CH4_BEGIN.match(t)]
        best, best_n = (0, len(self.blocks)), -1
        for a in begins:
            b = len(self.blocks)
            for i in range(a + 1, len(self.blocks)):
                _e, k, t = self.blocks[i]
                if k == "p" and CH4_END.match(t):
                    b = i
                    break
            n = sum(1 for i in range(a, b) if self.blocks[i][1] == "p" and MAU_RE.match(self.blocks[i][2]))
            if n > best_n:
                best, best_n = (a, b), n
        return best

    def forms(self):
        a, b = self.window()
        heads = [(MAU_RE.match(t).group(1).upper(), i)
                 for i, (_e, k, t) in enumerate(self.blocks) if k == "p" and a <= i < b and MAU_RE.match(t)]
        out = []
        for k, (code, i) in enumerate(heads):
            i_end = heads[k + 1][1] if k + 1 < len(heads) else b
            out.append((code, self.pick_title(i, i_end), i, i_end))
        return out

    def pick_title(self, i: int, i_end: int) -> str:
        """Tiêu đề biểu mẫu theo thứ tự ưu tiên:

        1. đoạn CHỮ HOA (≥ 2 từ) — tiêu đề TT79 luôn viết hoa
        2. đoạn bold hoặc căn giữa (≥ 2 từ)
        3. ô CHỮ HOA ở 3 dòng đầu của bảng (một số mẫu để tiêu đề trong bảng)
        """
        GENERIC = {"STT", "TT", "I", "II", "III", "IV", "V", "A", "B", "G"}

        def ok(t: str) -> bool:
            if len(t) < 8 or t.upper() in GENERIC or t.startswith(BAD_TITLE):
                return False
            return len(t.split()) >= 2

        def upper(t: str) -> bool:
            return t == t.upper() and any(c.isalpha() for c in t)

        p1, p2, p3 = [], [], []
        for j in range(i + 1, min(i + 40, i_end)):
            el, kind, text = self.blocks[j]
            if kind == "p":
                if not text or len(text) > 160 or not ok(text):
                    continue
                if upper(text):
                    p1.append(text)
                elif is_bold(el) or is_centered(el):
                    p2.append(text)
            elif kind == "tbl":
                for tr in table_rows(el)[:3]:
                    for txt in cell_texts(tr):
                        if ok(txt) and len(txt) <= 160 and upper(txt):
                            p3.append(txt)
        for group in (p1, p2, p3):
            if group:
                return min(group, key=len)
        return ""


# ------------------------------------------------------------------ writers
def strip_dangling_media_rels(rels_xml: bytes, keep: set[str]) -> bytes:
    """Xoá `<Relationship Target="media/...">` trỏ tới file media đã bị lược bỏ.

    Không xoá được thì Word vẫn mở (bỏ qua), nhưng python-docx báo lỗi khi đọc lại file —
    để lại quan hệ mồ côi là lỗi dữ liệu, phải dọn.
    """
    text = rels_xml.decode("utf-8")
    kept = {n.rsplit("/", 1)[-1] for n in keep}

    def keep_or_drop(m: re.Match) -> str:
        t = re.search(r'Target="([^"]+)"', m.group(0))
        if t and t.group(1).startswith("media/") and t.group(1).rsplit("/", 1)[-1] not in kept:
            return ""
        return m.group(0)

    new = re.sub(r"<Relationship\b[^>]*/>", keep_or_drop, text)
    return new.encode("utf-8") if new != text else rels_xml


def write_form(src: Source, dst: str, i0: int, i_end: int, keep_media: bool) -> None:
    """Ghi file mới = bản sao nguồn, body chỉ còn [i0, i_end) + sectPr gốc.

    Giữ nguyên styles.xml, footers, numbering, media được tham chiếu → file ra
    đúng 100 % định dạng gốc TT79 (khổ A4, lề 20/20/30/20, TNR 14).
    """
    body = etree.Element(src.body.tag, src.body.attrib, nsmap=src.body.nsmap)
    for j, (el, _kind, _t) in enumerate(src.blocks):
        if i0 <= j < i_end:
            body.append(copy.deepcopy(el))
    sect = src.sectpr                      # BẮT BUỘC: thiếu sectPr là mất khổ giấy & lề
    if sect is not None:
        body.append(copy.deepcopy(sect))

    new_root = etree.Element(src.root.tag, src.root.attrib, nsmap=src.root.nsmap)
    new_root.append(body)
    xml = etree.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    used = set(re.findall(rb'r:(?:embed|link)="([^"]+)"', xml))
    with zipfile.ZipFile(src.path) as z:
        rels = z.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in src.names else b""
    relmap = {m.group(1).decode(): m.group(2).decode()
              for m in re.finditer(rb'Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    keep = {"word/" + relmap[rid.decode()] for rid in used
            if rid.decode() in relmap and relmap[rid.decode()].startswith("media/")}

    with zipfile.ZipFile(src.path) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            n = it.filename
            if n == "word/document.xml":
                zout.writestr(it, xml)
                continue
            if (not keep_media) and n.startswith("word/media/") and n not in keep:
                continue
            blob = zin.read(n)
            # bỏ quan hệ trỏ tới media đã xoá (tránh rels "mồ côi" làm hỏng file khi mở lại)
            if (not keep_media) and n.endswith(".rels"):
                blob = strip_dangling_media_rels(blob, keep)
            zout.writestr(it, blob)


def write_excel(src: Source, dst: str, i0: int, i_end: int) -> int:
    """Mỗi bảng Word trong biểu mẫu → 1 sheet Excel, giữ nguyên tiêu đề cột TT79."""
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    tables = [el for el, kind, _t in src.blocks[i0:i_end] if kind == "tbl"]
    if not tables:
        return 0

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    head_fill = PatternFill("solid", fgColor="D9E1F2")
    for n, tb in enumerate(tables, 1):
        ws = wb.create_sheet(f"Bang {n}"[:31])
        for r, tr in enumerate(table_rows(tb), 1):
            for ci, txt in cell_positions(tr):
                c = ws.cell(r, ci, txt)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                c.font = Font(name="Times New Roman", size=11, bold=(r <= 3))
                if r <= 3:
                    c.fill = head_fill
        for ci in range(1, min(ws.max_column, 16) + 1):
            ws.column_dimensions[ws.cell(1, ci).column_letter].width = 19
        ws.freeze_panes = "A4"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = "1:3"
    wb.save(dst)
    return len(tables)


# ------------------------------------------------------------------ main
def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("out", nargs="?", default="tt79_forms")
    ap.add_argument("--only", default="", help="vd: 02,06,11 — lọc theo tiền tố mã mẫu")
    ap.add_argument("--no-excel", dest="excel", action="store_false", default=True)
    ap.add_argument("--keep-media", action="store_true", default=False)
    a = ap.parse_args()

    os.makedirs(a.out, exist_ok=True)
    src = Source(a.template)
    only = [p.strip().upper() for p in a.only.split(",") if p.strip()]

    rows, n_docx, n_xlsx = [], 0, 0
    for code, title, i0, i1 in src.forms():
        n = mau_number(code)
        if n == 0 or n in SKIP_NUM:
            continue
        if only and not any(code.startswith(p) for p in only):
            continue
        base = f"TT79-M{code}"
        dst = os.path.join(a.out, base + ".docx")
        write_form(src, dst, i0, i1, a.keep_media)
        n_docx += 1
        sheets = 0
        if a.excel:
            xdst = os.path.join(a.out, base + ".xlsx")
            sheets = write_excel(src, xdst, i0, i1)
            if sheets:
                n_xlsx += 1
            elif os.path.exists(xdst):
                os.remove(xdst)
        rows.append((code, title, sheets, os.path.basename(dst)))
        print(f"  Mẫu {code:6} {title[:58]:60} {sheets or '-':>3} sheet")

    idx = os.path.join(a.out, "DANH-MUC-BIEU-MAU.md")
    with open(idx, "w", encoding="utf-8") as f:
        f.write(f"# Danh mục biểu mẫu dự thầu — TT 79/2025/TT-BTC\n\n"
                f"Nguồn: `{os.path.basename(a.template)}` — Chương IV. Biểu mẫu mời thầu và dự thầu\n\n"
                "| Mẫu | Tên biểu mẫu | Sheet Excel | File Word |\n|:--|:--|--:|:--|\n")
        for code, title, sheets, fn in rows:
            f.write(f"| {code} | {title} | {sheets or ''} | `{fn}` |\n")
    print(f"\n{n_docx} file Word, {n_xlsx} workbook Excel → {a.out}")
    print(f"Danh mục: {idx}")


if __name__ == "__main__":
    main()
