#!/usr/bin/env python3
"""
md_to_docx.py — Markdown sang DOCX: tiền xử lý, chuyển bằng Pandoc, chuẩn hoá định dạng.

    python3 scripts/convert/md_to_docx.py <file.md> [-o <file.docx>] [--no-format]

Ba bước:

1. Tiền xử lý (Pandoc không tự làm, hoặc làm không đúng ý):
   - bỏ YAML frontmatter
   - `[[wikilink]]` và `[[wikilink|nhãn]]` -> nhãn (liên kết Obsidian)
   - bỏ dòng đánh dấu trang OCR: `### Trang 18`
   - bỏ dòng phân cách rỗng kiểu `## ---`
2. `pandoc -f markdown -t docx` — Pandoc lo tiêu đề, đậm/nghiêng, danh sách, bảng, code, blockquote.
3. `scripts/format/format_docx.py` — áp font, lề, giãn dòng cho file kết quả (bỏ qua bằng `--no-format`).

Cần Pandoc trên PATH (`brew install pandoc`, `apt install pandoc`, hoặc winget).
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile


def preprocess(text: str) -> str:
    text = re.sub(r"\A---\n.*?\n---\n", "", text, flags=re.S)                 # frontmatter
    text = re.sub(r"\[\[([^\]|]+)\|([^\]]+)\]\]", r"\2", text)                # [[đích|nhãn]]
    text = re.sub(r"\[\[([^\]]+)\]\]", r"\1", text)                           # [[đích]]
    text = re.sub(r"^\s*#{2,4}\s*Trang\s+\d+\s*$", "", text, flags=re.M | re.I)
    text = re.sub(r"^\s*#{1,6}\s*[-–—]+\s*$", "", text, flags=re.M)           # '## ---'
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


def pandoc_md_to_docx(md_path: str, docx_path: str) -> None:
    exe = shutil.which("pandoc")
    if not exe:
        raise SystemExit("Không thấy pandoc trên PATH. Cài rồi chạy lại: brew install pandoc "
                         "(macOS) / sudo apt install pandoc (Linux) / winget install JohnMacFarlane.Pandoc (Windows).")
    r = subprocess.run([exe, md_path, "-o", docx_path, "-f", "markdown", "-t", "docx", "--wrap=none"],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit(f"pandoc lỗi:\n{r.stderr[-600:]}")


def format_docx(docx_path: str) -> bool:
    here = os.path.dirname(os.path.abspath(__file__))
    fmt = os.path.join(here, "..", "format", "format_docx.py")
    if not os.path.exists(fmt):
        return False
    r = subprocess.run([sys.executable, fmt, docx_path, docx_path], capture_output=True, text=True)
    return r.returncode == 0


def convert(md_file: str, output_file: str | None = None, do_format: bool = True) -> str:
    out = output_file or os.path.splitext(md_file)[0] + ".docx"
    with open(md_file, encoding="utf-8") as f:
        text = preprocess(f.read())
    with tempfile.NamedTemporaryFile("w", suffix=".md", delete=False, encoding="utf-8") as fh:
        fh.write(text)
        tmp = fh.name
    try:
        pandoc_md_to_docx(tmp, out)
    finally:
        os.unlink(tmp)
    if do_format:
        format_docx(out)
    return out


# tương thích tên hàm cũ
convert_md_to_docx = convert


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("md")
    ap.add_argument("-o", "--out", default=None, help="file .docx kết quả (mặc định: cùng tên)")
    ap.add_argument("--no-format", dest="do_format", action="store_false", default=True,
                    help="không chạy bước chuẩn hoá định dạng")
    a = ap.parse_args()
    out = convert(a.md, a.out, a.do_format)
    print(f"Xong: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
