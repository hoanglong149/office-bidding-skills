#!/usr/bin/env python3
"""
ask_context.py — Quy trình HỎI LẠI ngữ cảnh trước khi sinh hồ sơ HSDT.

Skill không tự bịa thông tin gói thầu. Trước khi chạy, phải có đủ ngữ cảnh bắt buộc;
thiếu thì phải hỏi lại người dùng.

Ba chế độ:

  1) In câu hỏi ra màn hình (để agent hỏi trong chat)
        python3 ask_context.py

  2) Xuất PHIẾU để người dùng điền (không cần biết kỹ thuật)
        python3 ask_context.py --form phieu-ngu-can.md
        #  → mở file bằng Word/Notepad, điền sau chữ "Trả lời:"

  3) Kiểm phiếu đã điền — thiếu mục bắt buộc thì DỪNG và nói rõ còn thiếu gì
        python3 ask_context.py --check phieu-ngu-can.md
        #  → exit 0 nếu đủ, exit 1 nếu còn thiếu (in danh sách thiếu)

Bảng câu hỏi nằm ở `references/required-inputs.json` (một nguồn duy nhất).
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "references", "required-inputs.json")

EMPTY = {"", "-", "--", "n/a", "na", "chưa", "chua", "?", "..."}


def load() -> dict:
    with open(DATA, encoding="utf-8") as f:
        return json.load(f)


def all_items(data: dict):
    for g in data["groups"]:
        for it in g["items"]:
            yield g, it


# ------------------------------------------------------------------ 1. in câu hỏi
def print_questions(data: dict) -> None:
    for g in data["groups"]:
        print(f"\n### {g['id']}. {g['name']}")
        for it in g["items"]:
            mark = " (*BẮT BUỘC)" if it.get("blocking") else ""
            print(f"  [{it['id']}]{mark} {it['question']}")
            if it.get("where"):
                print(f"        ↳ lấy ở: {it['where']}")
    nb = sum(1 for _, it in all_items(data) if it.get("blocking"))
    print(f"\nTổng: {sum(1 for _ in all_items(data))} mục — trong đó {nb} mục BẮT BUỘC.")


# ------------------------------------------------------------------ 2. xuất phiếu
def write_form(data: dict, path: str) -> None:
    L = [
        "# PHIẾU NGỮ CẢNH — LÀM HỒ SƠ DỰ THẦU (HSDT)",
        "",
        "**Cách điền**: mở file này, gõ câu trả lời ngay sau chữ `Trả lời:` ở mỗi câu.",
        "Câu có dấu **(*)** là bắt buộc — thiếu là chưa làm hồ sơ được.",
        "Không biết thì ghi `chưa có` để người lập hồ sơ biết mà bổ sung.",
        "",
        "Điền xong, lưu file rồi báo lại: `python3 scripts/ask_context.py --check phieu-ngu-can.md`",
        "",
    ]
    for g in data["groups"]:
        L += [f"## {g['id']}. {g['name']}", ""]
        for it in g["items"]:
            mark = " **(*)**" if it.get("blocking") else ""
            L.append(f"### {it['id']}.{mark} {it['question']}")
            if it.get("where"):
                L.append(f"*(lấy ở: {it['where']})*")
            L += ["Trả lời:", "", ""]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print(f"Đã tạo phiếu: {path}")
    print("→ Mở bằng Word/Notepad, điền sau chữ 'Trả lời:', lưu lại rồi chạy --check.")


# ------------------------------------------------------------------ 3. kiểm phiếu
def parse_answers(path: str) -> dict[str, str]:
    with open(path, encoding="utf-8") as f:
        text = f.read()
    ans: dict[str, str] = {}
    blocks = re.split(r"^###\s+", text, flags=re.M)[1:]
    for b in blocks:
        m = re.match(r"([A-Z]\d+)\.", b)
        if not m:
            continue
        key = m.group(1)
        v = ""
        mv = re.search(r"Trả lời:[ \t]*(.*)", b)   # \s* sẽ ăn cả dòng mới → lọt heading nhóm sau
        if mv:
            v = mv.group(1).strip()
            # nếu câu trả lời nằm ở dòng dưới, gom các dòng tiếp theo tới dòng trống thứ hai
            # câu trả lời có thể xuống dòng — gom tới khi gặp tiêu đề/ghi chú kế tiếp
            for line in b[mv.end():].splitlines():
                if line.startswith("#") or line.startswith("*(") or line.startswith("### "):
                    break
                if line.strip():
                    v += (" " if v else "") + line.strip()
        ans[key] = v
    return ans


def check_answers(data: dict, path: str) -> int:
    ans = parse_answers(path)
    missing, done, warned = [], 0, []
    for g, it in all_items(data):
        v = (ans.get(it["id"], "") or "").strip()
        filled = v.lower() not in EMPTY and len(v) > 0
        if it.get("blocking") and not filled:
            missing.append((g["id"], it["id"], it["question"], it.get("where", "")))
        elif filled:
            done += 1
    print(f"Phiếu: {os.path.basename(path)} — đã trả lời {done} mục")
    if missing:
        print(f"\n❌ CÒN THIẾU {len(missing)} MỤC BẮT BUỘC — chưa sinh hồ sơ được:\n")
        for gid, iid, q, where in missing:
            print(f"  [{gid}·{iid}] {q}")
            if where:
                print(f"          ↳ lấy ở: {where}")
        print("\n→ Bổ sung các mục trên vào phiếu rồi chạy lại. "
              "Không tự đoán/không bịa số.")
        return 1
    print("\n✅ Đủ ngữ cảnh bắt buộc — được phép chạy bước sinh hồ sơ.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--form", metavar="FILE", help="xuất phiếu để người dùng điền")
    ap.add_argument("--check", metavar="FILE", help="kiểm phiếu đã điền (exit 1 nếu thiếu)")
    a = ap.parse_args()
    data = load()
    if a.form:
        write_form(data, a.form)
        return 0
    if a.check:
        return check_answers(data, a.check)
    print_questions(data)
    return 0


if __name__ == "__main__":
    sys.exit(main())
