#!/usr/bin/env bash
# Sinh lại toàn bộ bộ biểu mẫu HSDT (Word + Excel) và tự kiểm tra định dạng.
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT="${1:-${HSDT_OUT:-$DIR/../hsdt_forms}}"
python3 "$DIR/build_word.py"  "$OUT"
python3 "$DIR/build_excel.py" "$OUT"
echo "== Kiểm tra định dạng NĐ30 =="
python3 "$DIR/fix_nd30.py" --check "$OUT"/*.docx
