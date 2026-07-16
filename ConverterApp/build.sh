#!/usr/bin/env bash
###############################################################################
# build.sh
# Đóng gói ứng dụng Flask (pdftoexcel.py + launcher.py) thành 1 file .exe
# duy nhất bằng PyInstaller, để có thể gửi cho người khác dùng mà không cần
# cài Python.
#
# Cách chạy (trong Git Bash, tại thư mục ConverterApp/):
#   chmod +x build.sh
#   ./build.sh
#
# Kết quả: dist/PDFtoExcelConverter.exe
###############################################################################

set -e  # dừng ngay nếu có lệnh nào lỗi

APP_NAME="PDFtoExcelConverter"
ENTRY_POINT="launcher.py"
ICON_PATH="static/icon.ico"

echo "==> [1/5] Kiểm tra Python..."
if ! command -v python &> /dev/null; then
    echo "Không tìm thấy lệnh 'python'. Hãy cài Python và thêm vào PATH."
    exit 1
fi
python --version

echo "==> [2/5] Tạo/kích hoạt virtual environment (.venv)..."
# Nếu máy có nhiều bản Python (qua Python Launcher "py"), có thể ép dùng
# bản ổn định (vd 3.12) thay vì bản quá mới chưa đủ wheel dựng sẵn cho các
# gói C-extension như PyMuPDF/pandas, gây build từ source cực chậm hoặc treo.
# Bỏ comment dòng dưới và đổi số version nếu máy bạn có "py -3.12":
# PYTHON_CMD="py -3.12"
PYTHON_CMD="${PYTHON_CMD:-python}"

if [ ! -d ".venv" ]; then
    $PYTHON_CMD -m venv .venv
fi
source .venv/Scripts/activate   # Git Bash trên Windows dùng Scripts/, không phải bin/

echo "==> [3/5] Cài dependencies..."
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo "==> [4/5] Dọn build cũ (nếu có)..."
rm -rf build dist "${APP_NAME}.spec"

echo "==> [5/5] Chạy PyInstaller..."
# --add-data "nguồn;đích" : trên Windows dùng dấu ; để ngăn cách
# templates/ và static/ cần được đóng gói kèm vì code đọc qua BASE_DIR (_MEIPASS)
pyinstaller \
    --name "$APP_NAME" \
    --onefile \
    --icon "$ICON_PATH" \
    --add-data "templates;templates" \
    --add-data "static;static" \
    --hidden-import "pdfplumber" \
    --hidden-import "fitz" \
    --hidden-import "openpyxl.cell._writer" \
    --collect-submodules "pdfplumber" \
    --collect-data "pdfplumber" \
    --collect-data "fitz" \
    --collect-submodules "waitress" \
    "$ENTRY_POINT"

echo ""
echo "=================================================================="
echo " BUILD XONG!"
echo " File thực thi nằm tại: dist/${APP_NAME}.exe"
echo " Có thể gửi trực tiếp file này cho người khác (không cần cài Python)."
echo "=================================================================="