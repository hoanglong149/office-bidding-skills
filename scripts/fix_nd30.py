#!/usr/bin/env python3
"""
nd30_fix_format.py — Chuẩn hoá .docx theo Nghị định 30/2020/NĐ-CP (bản 2, đã sửa lỗi XML).

Sửa 4 lỗi khiến bản in không giống mẫu:
  1. Thiếu style `Normal` + `<w:pPrDefault>` rỗng  -> Word tự chế Normal (Calibri 11,
     dãn 1,08, after 8pt) làm lệch format.
  2. Không có footer đánh số trang (ND30 bắt buộc).
  3. Lề trang chưa theo dải ND30 (trên/dưới 20-25mm; trái 30-35mm; phải 15-20mm).
  4. Thiếu <w:pgNumType w:start="1"/>.

An toàn: kiểm tra XML hợp lệ (ElementTree) trước khi ghi đè; luôn tạo .bak.

Dùng:
  python3 nd30_fix_format.py --check  <files...>
  python3 nd30_fix_format.py          <files...>
"""

from __future__ import annotations
import os
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

MM = 56.6929
MAR_TOP = round(20 * MM)      # 1134
MAR_BOTTOM = round(20 * MM)   # 1134
MAR_LEFT = round(30 * MM)     # 1701
MAR_RIGHT = round(20 * MM)    # 1134  (NĐ30: lề phải 20 mm)
HEADER_H = round(12.5 * MM)
FOOTER_H = round(12.5 * MM)
FONT = "Times New Roman"
SZ = 26                       # 13 pt
LINE = 288                    # 1,2 lines
AFTER = 120                   # 6 pt

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

W = f'xmlns:w="{NS_W}"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

RPR = (f'<w:rFonts w:ascii="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
       f'<w:color w:val="000000"/><w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/>')

PPRDEFAULT = (f'<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="{AFTER}" '
              f'w:line="{LINE}" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>')

NORMAL_STYLE = (
    f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    f'<w:qFormat/><w:pPr><w:spacing w:before="0" w:after="{AFTER}" w:line="{LINE}" '
    f'w:lineRule="auto"/><w:jc w:val="both"/></w:pPr><w:rPr>{RPR}</w:rPr></w:style>'
)

FOOTER_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<w:ftr {W} {R}><w:p><w:pPr><w:jc w:val="center"/>'
    '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    f'<w:r><w:rPr>{RPR}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
    f'<w:r><w:rPr>{RPR}</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    f'<w:r><w:rPr>{RPR}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
    f'<w:r><w:rPr>{RPR}</w:rPr><w:t>1</w:t></w:r>'
    f'<w:r><w:rPr>{RPR}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
    '</w:p></w:ftr>'
)

CT_FOOTER = ('<Override PartName="/word/footer1.xml" ContentType="application/vnd.'
             'openxmlformats-officedocument.wordprocessingml.footer+xml"/>')

REL_FOOTER = ('<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/'
              'officeDocument/2006/relationships/footer" Target="footer1.xml"/>')


def fix_styles(styles: str) -> str:
    if 'w:styleId="Normal"' not in styles:
        styles = styles.replace('</w:styles>', NORMAL_STYLE + '</w:styles>', 1)
    if re.search(r'<w:pPrDefault\s*/>', styles):
        styles = re.sub(r'<w:pPrDefault\s*/>', PPRDEFAULT, styles, count=1)
    elif re.search(r'<w:pPrDefault>', styles):
        styles = re.sub(r'<w:pPrDefault>.*?</w:pPrDefault>', PPRDEFAULT, styles, count=1, flags=re.S)
    elif '</w:rPrDefault>' in styles:
        styles = styles.replace('</w:rPrDefault>', '</w:rPrDefault>' + PPRDEFAULT, 1)
    return styles


def fix_sectpr(doc: str) -> str:
    m = re.search(r'<w:sectPr\b.*?</w:sectPr>', doc, re.S)
    if not m:
        return doc
    s = m.group(0)
    new = s
    # 1) lề chuẩn
    if re.search(r'<w:pgMar[^>]*/>', new):
        new = re.sub(r'<w:pgMar[^>]*/>',
                     f'<w:pgMar w:top="{MAR_TOP}" w:right="{MAR_RIGHT}" w:bottom="{MAR_BOTTOM}" '
                     f'w:left="{MAR_LEFT}" w:header="{HEADER_H}" w:footer="{FOOTER_H}" w:gutter="0"/>',
                     new, count=1)
    # 2) pgNumType start=1
    if '<w:pgNumType' not in new:
        new = re.sub(r'(<w:pgMar[^>]*/>)', r'\1<w:pgNumType w:start="1"/>', new, count=1)
    else:
        new = new.replace('<w:pgNumType/>', '<w:pgNumType w:start="1"/>')
    # 3) footerReference: chèn NGAY SAU thẻ mở <w:sectPr ...>
    if 'footerReference' not in new:
        new = re.sub(r'(<w:sectPr\b[^>]*>)',
                     r'\1<w:footerReference w:type="default" r:id="rIdFtr1"/>',
                     new, count=1)
    return doc.replace(s, new, 1)


def fix_rels(rels: str) -> str:
    if 'footer1.xml' in rels:
        return rels
    return rels.replace('</Relationships>', REL_FOOTER + '</Relationships>', 1)


def fix_ct(ct: str) -> str:
    if '/word/footer1.xml' in ct:
        return ct
    return ct.replace('</Types>', CT_FOOTER + '</Types>', 1)


def analyze(path: str) -> dict:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        styles = z.read("word/styles.xml").decode("utf-8", "ignore")
        doc = z.read("word/document.xml").decode("utf-8", "ignore")
    prob, ok = [], []
    (ok if 'w:styleId="Normal"' in styles else prob).append(
        "style Normal" if 'w:styleId="Normal"' in styles else "thiếu style Normal")
    good = bool(re.search(r'<w:pPrDefault>\s*<w:pPr>', styles))
    (ok if good else prob).append("pPrDefault có dãn dòng" if good else "pPrDefault rỗng (thiếu dãn dòng)")
    has_ftr = "word/footer1.xml" in names and "footerReference" in doc
    (ok if has_ftr else prob).append("footer số trang" if has_ftr else "không có footer đánh số trang")
    mar = re.search(r'<w:pgMar ([^/>]*)/>', doc)
    if mar and f'w:left="{MAR_LEFT}"' in mar.group(1) and f'w:right="{MAR_RIGHT}"' in mar.group(1):
        ok.append("lề trang")
    else:
        prob.append(f"lề chưa chuẩn (cần trái {MAR_LEFT} / phải {MAR_RIGHT} twips)")
    return {"prob": prob, "ok": ok}


def process(path: str, check_only=False) -> dict:
    rep = analyze(path)
    if check_only or not rep["prob"]:
        return rep
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
        order = list(z.namelist())
    styles = fix_styles(parts["word/styles.xml"].decode("utf-8"))
    doc = fix_sectpr(parts["word/document.xml"].decode("utf-8"))
    rn = "word/_rels/document.xml.rels"
    rels = fix_rels(parts[rn].decode("utf-8")) if rn in parts else \
        f'<Relationships {W}>{REL_FOOTER}</Relationships>'
    cn = "[Content_Types].xml"
    ct = fix_ct(parts[cn].decode("utf-8")) if cn in parts else ""

    # --- BẮT BUỘC: kiểm XML hợp lệ trước khi ghi ---
    for nm, xml in (("styles.xml", styles), ("document.xml", doc), ("footer1.xml", FOOTER_XML)):
        ET.fromstring(xml)   # raise nếu hỏng
    ET.fromstring(rels)

    shutil.copyfile(path, path + ".bak")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in order:
            if n == "word/styles.xml":
                zo.writestr(n, styles)
            elif n == "word/document.xml":
                zo.writestr(n, doc)
            elif n == rn:
                zo.writestr(n, rels)
            elif n == cn and ct:
                zo.writestr(n, ct)
            else:
                zo.writestr(n, parts[n])
        if "word/footer1.xml" not in parts:
            zo.writestr("word/footer1.xml", FOOTER_XML)
    os.replace(tmp, path)
    rep["backup"] = path + ".bak"
    return rep


def main() -> int:
    files = [a for a in sys.argv[1:] if not a.startswith("--")]
    check = "--check" in sys.argv
    if not files:
        print(__doc__); return 1
    for p in files:
        if not os.path.isfile(p):
            print(f"✗ không thấy {p}"); continue
        try:
            r = process(p, check_only=check)
        except Exception as e:
            print(f"✗ {os.path.basename(p)} — LỖI: {e}"); continue
        icon = "✓" if not r["prob"] else ("•" if check else "✎")
        print(f"{icon} {os.path.basename(p)}")
        if r["ok"]:
            print("    đạt:", ", ".join(r["ok"]))
        if r["prob"]:
            print(("    CẦN SỬA: " if check else "    ĐÃ SỬA: ") + "; ".join(r["prob"]))
            if r.get("backup"):
                print("    backup:", os.path.basename(r["backup"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
