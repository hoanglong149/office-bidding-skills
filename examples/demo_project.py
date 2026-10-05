#!/usr/bin/env python3
"""
demo_project.py — Demo đầy đủ: lấy bộ template trắng trong `templates/`,
điền letterhead của một gói thầu giả định, chèn vài dòng dữ liệu mẫu,
rồi xuất ra một thư mục output hoàn chỉnh.

    python3 examples/demo_project.py [thư_mục_output]

Không cần dữ liệu thật. Chạy xong mở thư mục output để xem bộ HSDT thành phẩm.
"""

from __future__ import annotations
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TEMPLATES = os.path.join(ROOT, "templates")
SCRIPTS = os.path.join(ROOT, "scripts")

# --------------------------------------------------------------- gói thầu giả định
DEMO = {
    "bidder": "CÔNG TY CỔ PHẦN CƠ ĐIỆN XÂY DỰNG ABC",
    "bidder_info": "Số 123 Đường Lê Lợi, Q. Hải Châu, Đà Nẵng · (0236) 3xxx xxx · bid@abc.com.vn",
    "employer": "BAN QUẢN LÝ DỰ ÁN ĐIỆN LỰC MIỀN TRUNG (EVNCPC)",
    "employer_info": "278 Trần Phú, TP. Đà Nẵng · (0236) 3xxx xxx · (0236) 3xxx xxx",
    "place": "Đà Nẵng",
    "doc_no": "07/2026-HSDT",
    "date": "ngày 05 tháng 10 năm 2026",
}

# Vài dòng dữ liệu mẫu (thiết bị) để thấy pipeline chạy end-to-end.
EQ = [
    ("T-101", "Bồn chứa nước làm mát", "Trụ đứng, SS304, ØxH=3000x6000 mm"),
    ("P-101A/B", "Bơm nước làm mát ly tâm", "Q=320 m³/h; H=55 m; động cơ 75 kW, 380 V"),
    ("E-101", "Thiết bị trao đổi nhiệt dạng bản", "Vỏ & ống, 1.200 kW, 38→45 °C"),
    ("K-101", "Máy nén khí dụng cụ", "Trục vít, Q=1.200 Nm³/h; 8 barg"),
]


def run_demo(out: str) -> str:
    os.makedirs(out, exist_ok=True)
    n = 0
    for f in sorted(os.listdir(TEMPLATES)):
        if f.endswith((".docx", ".xlsx")):
            shutil.copy2(os.path.join(TEMPLATES, f), os.path.join(out, f))
            n += 1
    print(f"1/3  Đã copy {n} template trắng → {out}")

    subprocess.run([sys.executable, os.path.join(SCRIPTS, "fill_letterhead.py"), out,
                    "--bidder", DEMO["bidder"], "--bidder-info", DEMO["bidder_info"],
                    "--employer", DEMO["employer"], "--employer-info", DEMO["employer_info"],
                    "--place", DEMO["place"], "--doc-no", DEMO["doc_no"],
                    "--date", DEMO["date"]], check=True)
    print("2/3  Đã điền letterhead")

    import openpyxl
    p = os.path.join(out, "HSDT-F20-F23-Bieu-gia.xlsx")
    wb = openpyxl.load_workbook(p)
    ws = wb["21.1 - Schedule 1"]
    for i, (tag, name, spec) in enumerate(EQ):
        r = 7 + i
        ws.cell(r, 1).value = i + 1
        ws.cell(r, 2).value = f"{tag} — {name}\n{spec}"
    ws.cell(4, 1).value = "DỮ LIỆU MẪU (demo) — thay bằng dữ liệu thật của gói thầu"
    wb.save(p)

    p2 = os.path.join(out, "HSDT-ATT1A-1B-2-Manufacturers-Origin-Equipment.xlsx")
    wb = openpyxl.load_workbook(p2)
    w2, w1 = wb["ATT2 - Thiet bi vat tu"], wb["ATT1B - Thiet bi chinh"]
    for i, (tag, name, spec) in enumerate(EQ):
        r = 7 + i
        w2.cell(r, 1).value = i + 1
        w2.cell(r, 2).value = f"{tag} — {name}"
        w2.cell(r, 6).value = spec
        w2.cell(r, 7).value = "[NC cung cấp]"
        w1.cell(r, 1).value = i + 1
        w1.cell(r, 2).value = f"{tag} — {name}"
        w1.cell(r, 5).value = "[NC cung cấp]"
    wb.save(p2)
    print(f"3/3  Đã chèn {len(EQ)} dòng thiết bị mẫu vào Form 21.1 / Attachment 1B / Attachment 2")
    return out


if __name__ == "__main__":
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "demo_output")
    run_demo(out_dir)
    print(f"\nXong. Mở thư mục: {out_dir}")
    print("Kiểm tra định dạng:  python3 scripts/fix_nd30.py --check "
          f"{out_dir}/*.docx")
