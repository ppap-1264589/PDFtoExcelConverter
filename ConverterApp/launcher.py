"""
launcher.py
Chạy Flask app (pdftoexcel.py) dưới dạng desktop app:
- Tự tìm port trống (tránh xung đột nếu 5000 đang bận)
- Tự mở trình duyệt sau khi server sẵn sàng
- Bắt sự kiện đóng cửa sổ console (nút X) trên Windows để dọn dẹp gọn gàng
- Khi process kết thúc (bằng bất kỳ cách nào), OS tự giải phóng port -> không cần
  tự "đóng port" thủ công, chỉ cần đảm bảo process thoát sạch.
"""

import os
import sys
import time
import socket
import atexit
import logging
import threading
import webbrowser

from waitress import serve

# ---------------------------------------------------------------------------
# Logging: ép UTF-8 để không bị UnicodeEncodeError trên console cp1252 (Windows)
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def find_free_port(start: int = 5000, end: int = 5100) -> int:
    """Quét một dải port, trả về port đầu tiên đang trống."""
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError("No free port found in the range 5000-5100")


def resource_path(relative_path: str) -> str:
    """
    Trả về đường dẫn tuyệt đối tới resource, hoạt động đúng cả khi
    chạy bằng `python launcher.py` lẫn khi đã đóng gói bằng PyInstaller.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def open_browser_when_ready(url: str, port: int, timeout: float = 10.0):
    """Đợi server thật sự lắng nghe rồi mới mở browser, tránh mở quá sớm."""
    start = time.time()
    while time.time() - start < timeout:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) == 0:
                webbrowser.open(url)
                return
        time.sleep(0.2)
    logger.warning("Server is not ready after %.0f seconds, browser will not be opened automatically.", timeout)


def cleanup():
    """Chỗ để dọn file tạm, đóng kết nối, v.v. trước khi thoát hẳn."""
    logger.info("Cleaning up resources before exit...")
    # TODO: xoá file tạm trong checked_status/ hoặc thư mục tạm nếu cần
    # ví dụ: shutil.rmtree(resource_path("tmp"), ignore_errors=True)


atexit.register(cleanup)


def register_windows_console_handler():
    """
    Trên Windows, khi người dùng bấm nút X ở cửa sổ console, Python KHÔNG
    tự nhận được KeyboardInterrupt như Ctrl+C. Cần đăng ký handler riêng
    qua Win32 API để bắt sự kiện đó và dọn dẹp trước khi process bị kill.
    """
    if sys.platform != "win32":
        return
    import ctypes

    kernel32 = ctypes.windll.kernel32
    PHANDLER_ROUTINE = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_uint)

    def console_ctrl_handler(ctrl_type):
        # 0=CTRL_C, 2=CTRL_CLOSE (nút X), 5=CTRL_LOGOFF, 6=CTRL_SHUTDOWN
        logger.info("Received application shutdown signal (ctrl_type=%s)", ctrl_type)
        cleanup()
        os._exit(0)
        return 1

    # Giữ reference để tránh bị garbage collect
    register_windows_console_handler._handler_ref = PHANDLER_ROUTINE(console_ctrl_handler)
    kernel32.SetConsoleCtrlHandler(register_windows_console_handler._handler_ref, True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    register_windows_console_handler()

    # Import app SAU khi logging đã cấu hình xong, để log của Flask cũng đúng UTF-8
    from pdftoexcel import app  # app = Flask(__name__) trong pdftoexcel.py

    port = find_free_port()
    url = f"http://127.0.0.1:{port}"

    threading.Thread(target=open_browser_when_ready, args=(url, port), daemon=True).start()

    logger.info("Starting server at %s", url)
    try:
        # waitress.serve() là production WSGI server thuần Python, không cần
        # compile C nên rất hợp để đóng gói bằng PyInstaller trên Windows.
        # Thay cho app.run() của Flask (vốn chỉ dành cho dev, luôn in cảnh
        # báo "This is a development server...").
        serve(app, host="127.0.0.1", port=port, threads=8)
    except KeyboardInterrupt:
        logger.info("Received Ctrl+C, exiting...")
    finally:
        cleanup()


if __name__ == "__main__":
    main()