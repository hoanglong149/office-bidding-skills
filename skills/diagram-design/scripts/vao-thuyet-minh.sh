#!/usr/bin/env bash
# vao-thuyet-minh.sh — render sơ đồ HTML trong một thư mục rồi chèn vào thuyết minh Word.
#
#   bash scripts/vao-thuyet-minh.sh <thư-mục-hình> "<file.docx>"
#
# Lần đầu chưa có file Word thì thêm --new:
#   bash scripts/vao-thuyet-minh.sh ./hinh "Thuyet minh.docx" --new
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY="${PYTHON:-python3}"

if [ $# -lt 2 ]; then
  echo "Dùng: bash scripts/vao-thuyet-minh.sh <thư-mục-hình> \"<file.docx>\" [--new]"
  exit 1
fi
exec "$PY" "$DIR/figdoc.py" --figs "$1" --docx "$2" --white "${@:3}"
