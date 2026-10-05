#!/usr/bin/env python3
"""
figdoc.py - render sơ đồ HTML thành SVG/PNG và chèn vào thuyết minh .docx.

Một lệnh cho cả hai việc:

    python3 figdoc.py --figs <thư_mục_hoặc_file.html> --docx <file.docx>

Việc script làm:

1. Với mỗi file .html trong `--figs`:
   - tách nút `<svg>` đầu tiên ra file `.svg` (vector, giữ nguyên `<title>`, `<desc>`)
   - render nguyên trang HTML bằng Chrome headless -> `.png` độ phân giải cao (mặc định 3x)
2. Chèn ảnh vào `--docx`:
   - nếu bài có dòng đánh dấu `[[FIG:tên-file]]` thì chèn đúng chỗ đó
   - không có đánh dấu thì chèn vào cuối bài
   - ảnh canh giữa, rộng 16 cm; chú thích `Hình N. <tiêu đề>` bên dưới, nghiêng, 12 pt
   - số thứ tự tự đếm tiếp theo các chú thích `Hình N.` đã có trong bài

Ví dụ:

    # tạo thuyết minh mới chỉ gồm các hình
    python3 figdoc.py --figs ./hinh/ --docx thuyet-minh.docx --new

    # chèn vào thuyết minh có sẵn, chèn tại dòng [[FIG:fig-01]]
    python3 figdoc.py --figs ./hinh/ --docx "Thuyet minh.docx"

    # chỉ render, không chèn
    python3 figdoc.py --figs ./hinh/ --png-only

Tùy chọn:
    --width 16        bề rộng ảnh trong Word (cm)
    --scale 3         hệ số phóng khi render PNG
    --white           nền PNG trắng (mặc định trong suốt, hợp trang trắng)
    --chrome <path>   đường dẫn Chrome/Chromium nếu không tự tìm thấy
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "msedge",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


# ------------------------------------------------------------------ helpers
def find_chrome(explicit: str | None = None) -> str:
    if explicit:
        return explicit
    for c in CHROME_CANDIDATES:
        if os.path.isabs(c):
            if os.path.exists(c):
                return c
        elif shutil.which(c):
            return shutil.which(c)
    raise SystemExit("Không tìm thấy Chrome/Chromium. Truyền đường dẫn bằng --chrome.")


def first_svg(html: str) -> str:
    """Nút <svg> của SƠ ĐỒ: lấy khối lớn nhất theo diện tích viewBox.

    Nhiều template có thêm <svg> nhỏ (icon, sparkline) đứng trước sơ đồ; lấy "thẻ <svg> đầu tiên"
    như tài liệu gốc sẽ chụp nhầm icon 66x66 px.
    """
    blocks = re.findall(r"<svg\b.*?</svg>", html, re.S)
    if not blocks:
        raise SystemExit("File HTML không có thẻ <svg> - không phải sơ đồ của diagram-design.")

    def area(svg: str) -> int:
        w, h = svg_size(svg)
        return w * h

    return max(blocks, key=area)


def svg_size(svg: str) -> tuple[int, int]:
    """(rộng, cao) điểm ảnh, chỉ đọc thuộc tính của CHÍNH thẻ <svg>.

    Không quét cả chuỗi: bên trong sơ đồ có nhiều <rect width="66">, quét cả chuỗi sẽ lấy nhầm.
    """
    m = re.match(r"<svg\b([^>]*)>", svg, re.S)
    attrs = m.group(1) if m else svg[:400]

    def num(attr):
        mm = re.search(rf'\b{attr}="([0-9.]+)', attrs)
        return float(mm.group(1)) if mm else None

    w, h = num("width"), num("height")
    if w and h:
        return int(w), int(h)
    vb = re.search(r'viewBox="([\d.\-\s]+)"', attrs)
    if vb:
        vals = [float(x) for x in vb.group(1).split()]
        if len(vals) == 4:
            return int(vals[2]), int(vals[3])
    return 1400, 900


def svg_title(svg: str, fallback: str) -> str:
    m = re.search(r"<title[^>]*>(.*?)</title>", svg, re.S)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else fallback


def write_svg(html_path: str, svg_path: str) -> str:
    html = open(html_path, encoding="utf-8").read()
    svg = first_svg(html)
    if 'xmlns="http://www.w3.org/2000/svg"' not in svg:
        svg = svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"', 1)
    if "viewBox=" not in svg:
        w, h = svg_size(svg)
        svg = svg.replace("<svg", f'<svg viewBox="0 0 {w} {h}"', 1)
    open(svg_path, "w", encoding="utf-8").write('<?xml version="1.0" encoding="UTF-8"?>\n' + svg)
    return svg


def render_png(chrome: str, html_path: str, png_path: str, svg: str, w: int, h: int,
               scale: int, white: bool) -> None:
    """Chụp ĐÚNG khung sơ đồ, không lấy phần trang (tiêu đề, nền chấm, thẻ tóm tắt).

    Chrome headless CLI không chụp được theo phần tử, nên dựng một trang tạm chỉ chứa <svg>,
    mang theo CSS và link font của trang gốc, rồi chụp với kích thước bằng đúng sơ đồ.
    """
    src = open(html_path, encoding="utf-8").read()
    head_bits = re.findall(r"<link[^>]+stylesheet[^>]*>", src) + re.findall(r"<style[^>]*>.*?</style>", src, re.S)
    tmp_html = (
        "<!doctype html><html><head><meta charset='utf-8'>"
        + "".join(head_bits)
        + "<style>html,body{margin:0;padding:0;background:transparent;}"
          "svg{display:block;}</style></head><body>"
        + svg + "</body></html>"
    )
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as fh:
        fh.write(tmp_html)
        tmp_path = fh.name

    bg = "FFFFFFFF" if white else "00000000"
    cmd = [
        chrome, "--headless", "--no-sandbox", "--disable-gpu", "--hide-scrollbars",
        f"--force-device-scale-factor={scale}",
        f"--default-background-color={bg}",
        f"--window-size={w},{h}",
        f"--screenshot={png_path}",
        f"file://{tmp_path}",
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        if not os.path.exists(png_path):
            raise SystemExit(f"Render PNG thất bại: {html_path}\n{r.stderr[-400:]}")
    finally:
        os.unlink(tmp_path)


def png_size(path: str) -> tuple[int, int]:
    import struct
    d = open(path, "rb").read(33)
    return struct.unpack(">II", d[16:24])


# ------------------------------------------------------------------ word
def next_figure_number(doc) -> int:
    n = 0
    for p in doc.paragraphs:
        m = re.match(r"\s*Hình\s+(\d+)\s*\.", p.text)
        if m:
            n = max(n, int(m.group(1)))
    return n + 1


def insert_into_docx(docx_path: str, figures: list[tuple[str, str]], width_cm: float,
                     create_new: bool) -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt

    if create_new and not os.path.exists(docx_path):
        doc = Document()
        st = doc.styles["Normal"]
        st.font.name = "Times New Roman"
        st.font.size = Pt(13)
    else:
        if not os.path.exists(docx_path):
            raise SystemExit(f"Không thấy file Word: {docx_path} (thêm --new để tạo mới)")
        doc = Document(docx_path)

    # Thứ tự đánh số phải theo VỊ TRÍ trong bài, không theo thứ tự file đầu vào.
    ordered = []
    used = set()
    for p in doc.paragraphs:
        for idx, (png, caption) in enumerate(figures):
            marker = f"[[FIG:{os.path.splitext(os.path.basename(png))[0]}]]"
            if idx not in used and marker in p.text:
                ordered.append((idx, p))
                used.add(idx)
    tail = [i for i in range(len(figures)) if i not in used]

    start_no = next_figure_number(doc)
    seq = ordered + [(i, None) for i in tail]

    for n, (idx, par) in enumerate(seq):
        png, caption = figures[idx]
        no = start_no + n
        marker = f"[[FIG:{os.path.splitext(os.path.basename(png))[0]}]]"

        if par is None:                      # không có đánh dấu: thêm vào cuối bài
            par = doc.add_paragraph()
        else:
            for r in list(par.runs):
                r.text = r.text.replace(marker, "")

        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.add_run().add_picture(png, width=Cm(width_cm))

        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cap.add_run(f"Hình {no}. {caption}")
        run.italic = True
        run.font.size = Pt(12)
        if par is not None and par._p.getparent() is not None:
            par._p.addnext(cap._p)

    doc.save(docx_path)


# ------------------------------------------------------------------ main
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--figs", required=True, help="thư mục hoặc danh sách file .html")
    ap.add_argument("--docx", default=None, help="file Word để chèn hình")
    ap.add_argument("--new", action="store_true", help="tạo file Word mới nếu chưa có")
    ap.add_argument("--out-dir", default=None, help="nơi ghi .svg/.png (mặc định: cạnh file html)")
    ap.add_argument("--width", type=float, default=16.0, help="bề rộng ảnh trong Word (cm)")
    ap.add_argument("--scale", type=int, default=3, help="hệ số phóng PNG (mặc định 3)")
    ap.add_argument("--white", action="store_true", help="nền PNG trắng thay vì trong suốt")
    ap.add_argument("--png-only", action="store_true", help="chỉ render, không chèn Word")
    ap.add_argument("--chrome", default=None)
    a = ap.parse_args()

    files = sorted(glob.glob(os.path.join(a.figs, "*.html"))) if os.path.isdir(a.figs) \
        else [f for f in glob.glob(a.figs) if f.endswith(".html")]
    if not files:
        raise SystemExit(f"Không thấy file .html nào trong: {a.figs}")

    chrome = find_chrome(a.chrome)
    figures: list[tuple[str, str]] = []

    for f in files:
        base = os.path.splitext(os.path.basename(f))[0]
        out_dir = a.out_dir or os.path.dirname(os.path.abspath(f))
        os.makedirs(out_dir, exist_ok=True)
        svg_path = os.path.join(out_dir, base + ".svg")
        png_path = os.path.join(out_dir, base + ".png")

        svg = write_svg(f, svg_path)
        w, h = svg_size(svg)
        render_png(chrome, f, png_path, svg, w, h, a.scale, a.white)
        pw, ph = png_size(png_path)
        caption = svg_title(svg, base)
        figures.append((png_path, caption))
        print(f"  {base}: {pw}x{ph} px  |  {svg_path}  |  {png_path}")

    if a.png_only:
        print(f"\nĐã render {len(figures)} hình.")
        return 0

    if not a.docx:
        raise SystemExit("Thiếu --docx (hoặc dùng --png-only nếu chỉ muốn render).")
    insert_into_docx(a.docx, figures, a.width, a.new)
    print(f"\nĐã chèn {len(figures)} hình vào {a.docx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
