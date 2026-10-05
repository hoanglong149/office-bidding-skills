#!/usr/bin/env bash
# lam-ho-so.sh — Chạy một lệnh để ra bộ hồ sơ dự thầu (dành cho người không rành kỹ thuật).
#
#   bash lam-ho-so.sh
#
# Script sẽ hỏi từng câu bằng tiếng Việt (Enter = dùng giá trị mặc định), rồi:
#   1. ghi cấu hình gói thầu vào projects/<mã-gói>.json
#   2. tách/đóng dấu letterhead + định dạng cho toàn bộ biểu mẫu
#   3. kiểm tra lại định dạng và mở thư mục kết quả
set -euo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY="${PYTHON:-python3}"

echo "=================================================="
echo " LÀM HỒ SƠ DỰ THẦU (HSDT) — TT 79/2025/TT-BTC"
echo " Enter = dùng giá trị trong [ ]"
echo "=================================================="
echo

ask() {  # ask "câu hỏi" "mặc định"  -> in ra giá trị
  local q="$1" d="${2:-}" v=""
  # câu hỏi phải in ra STDERR: stdout được dùng để trả giá trị qua $(ask ...)
  if [ -n "$d" ]; then printf '%s [%s]: ' "$q" "$d" >&2; else printf '%s: ' "$q" >&2; fi
  read -r v || true
  [ -z "$v" ] && v="$d"
  printf '%s' "$v"
}

# --- 1. mã gói ---
while :; do
  MA_GOI="$(ask "Mã gói thầu (viết liền, không dấu — vd OFFGAS-EPC, PL-EPC)")"
  [ -n "$MA_GOI" ] && break
  echo "  ! Phải nhập mã gói."
done

# --- 2. loại gói ---
echo
echo "Loại gói thầu:"
echo "  1) EPC  (thiết kế + cung cấp hàng hóa + xây lắp)"
echo "  2) Xây lắp"
echo "  3) Mua sắm hàng hóa"
echo "  4) EC   (thiết kế + xây lắp)"
echo "  5) PC   (cung cấp hàng hóa + xây lắp)"
LT="$(ask "Chọn 1-5" "1")"
case "$LT" in
  1) BO="EPC-10A";  LOAI="EPC" ;;
  2) BO="Xay-lap-3A"; LOAI="Xây lắp" ;;
  3) BO="Hang-hoa-4A"; LOAI="Mua sắm hàng hóa" ;;
  4) BO="EC-8A";    LOAI="EC" ;;
  5) BO="PC-9A";    LOAI="PC" ;;
  *) echo "  ! Không hợp lệ, dùng mặc định EPC."; BO="EPC-10A"; LOAI="EPC" ;;
esac
echo "  → Bộ biểu mẫu: $BO"

# --- 3-4. tên dự án / gói thầu ---
echo
echo "Nhập NGUYÊN VĂN như trong HSMT (sẽ in lên letterhead mọi trang):"
while :; do
  DU_AN="$(ask "Tên dự án")"; case "$DU_AN" in *"["*|"") echo "  ! Chưa điền.";; *) break;; esac
done
while :; do
  GOI="$(ask "Tên gói thầu")"; case "$GOI" in *"["*|"") echo "  ! Chưa điền.";; *) break;; esac
done

# --- 5. logo ---
echo
echo "Logo:  Enter = dùng logo IPC E&C có sẵn trong skill"
LOGO="$(ask "Đường dẫn file logo (PNG/JPG)" "assets/logo-ipc-ec.png")"
[ -f "$DIR/$LOGO" ] || [ -f "$LOGO" ] || { echo "  ! Không thấy file logo — bỏ logo."; LOGO=""; }

# --- 6. khổ giấy ---
echo
GIU="$(ask "Khổ giấy: a = A4 (mặc định) / l = Letter" "a")"
if [ "$GIU" = "l" ] || [ "$GIU" = "L" ]; then PAPER="letter"; else PAPER="a4"; fi

# --- 7. nơi lưu ---
echo
OUT="$(ask "Thư mục lưu kết quả" "$HOME/Desktop/HSDT-$MA_GOI")"

# --- ghi cấu hình ---
CFG="$DIR/projects/$MA_GOI.json"
mkdir -p "$DIR/projects"
cat > "$CFG" << EOF
{
  "project": "$DU_AN",
  "package": "$GOI",
  "logo_left": null,
  "logo_right": "$LOGO",
  "paper": "$PAPER",
  "profile": "intl-epc",
  "logo_height": 18,
  "page_number": "footer"
}
EOF
echo
echo "①  Đã ghi cấu hình: projects/$MA_GOI.json"

# --- chạy ---
echo "②  Đang tạo hồ sơ ($LOAI) …"
rm -rf "$OUT"; mkdir -p "$OUT"
"$PY" "$DIR/scripts/apply_project_format.py" "$DIR/templates/$BO" --config "$CFG" --out "$OUT"

echo "③  Kiểm tra định dạng …"
"$PY" "$DIR/scripts/fix_page_setup.py" --check "$OUT"/*.docx | tail -20

echo
echo "=================================================="
echo " XONG. Hồ sơ ở: $OUT"
echo "=================================================="
echo "Việc tiếp theo:"
echo "  • Điền nội dung từng biểu mẫu (nhân sự, thiết bị, tài chính, giá…)"
echo "  • Kiểm ngữ cảnh còn thiếu:  python3 scripts/ask_context.py --check phieu-ngu-can.md"
echo "  • Chưa có phiếu?  python3 scripts/ask_context.py --form phieu-ngu-can.md"
echo
open "$OUT" 2>/dev/null || true
