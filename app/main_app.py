import sys
import threading
import time
import socket
import urllib.request
import uvicorn
from .api import app as fastapi_app

from PyQt6.QtWidgets import QApplication, QMainWindow, QSplashScreen, QLabel, QVBoxLayout, QWidget, QStatusBar
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineSettings
from PyQt6.QtCore import QUrl, QTimer, Qt, QFileInfo
from PyQt6.QtGui import QIcon, QColor, QPixmap, QFont
import webbrowser
import os
import json


# Keep global references so Python doesn't garbage-collect the window
_window = None
_splash = None
_public_url = None
_temp_views = [] # Storage for target="_blank" views


def run_server():
    """Runs the FastAPI server accessible from all network interfaces."""
    uvicorn.run(fastapi_app, host="0.0.0.0", port=8000, log_level="error")


def start_ngrok():
    """Start a ngrok tunnel and return the public URL (or None on failure)."""
    global _public_url
    try:
        from pyngrok import ngrok, conf

        # Load token from config.json if available
        token = ""
        if os.path.exists("config.json"):
            with open("config.json", "r") as f:
                cfg = json.load(f)
                token = cfg.get("NGROK_AUTHTOKEN", "")

        if token:
            ngrok.set_auth_token(token)

        # Try to open the tunnel
        tunnel = ngrok.connect(8000, "http")
        _public_url = tunnel.public_url
        print(f"\n  [GLOBE]  Remote Access URL: {_public_url}")
        print(f"  [CLIPBOARD]  Share this with anyone - works from home, phone, anywhere!\n")
        return _public_url
    except Exception as e:
        print(f"  [WARNING]  ngrok tunnel not started: {e}")
        return None


def wait_for_server(max_tries=40):
    """Block until the server responds or we give up."""
    for _ in range(max_tries):
        try:
            urllib.request.urlopen("http://127.0.0.1:8000", timeout=0.5)
            return True
        except Exception:
            time.sleep(0.15)
    return False


class CustomWebEngineView(QWebEngineView):
    def createWindow(self, type):
        """Handle target='_blank' links by opening them in the default system browser."""
        # Create a temporary view to catch the URL, then open in system browser.
        new_view = QWebEngineView()
        new_view.urlChanged.connect(lambda url: self.handle_external_url(url, new_view))
        _temp_views.append(new_view)
        return new_view

    def handle_external_url(self, url, view):
        url_str = url.toString()
        if url_str != "about:blank":
            webbrowser.open(url_str)
            if view in _temp_views:
                _temp_views.remove(view)
            view.deleteLater()


@fastapi_app.get("/api/info")
async def get_server_info():
    """Returns the current LAN and Remote URLs for connectivity."""
    global _public_url
    try:
        lan_ip = socket.gethostbyname(socket.gethostname())
        lan_url = f"http://{lan_ip}:8000"
    except Exception:
        lan_url = "http://127.0.0.1:8000"
    return {"lan": lan_url, "remote": _public_url or ""}


def launch_window():
    global _window, _splash, _public_url
    try:
        _window = WimpyWindow()

        # Embed the browser with custom link handler
        browser = CustomWebEngineView()
        browser.setUrl(QUrl("http://127.0.0.1:8000"))

        from PyQt6.QtWidgets import QFileDialog

        def handle_download(download_item):
            try:
                # We must use a safe name
                suggested = download_item.suggestedFileName() or "download"
                path, _ = QFileDialog.getSaveFileName(_window, "Save File", suggested)
                if path:
                    # In PyQt6, we must set Directory and Name separately (setPath is gone)
                    info = QFileInfo(path)
                    download_item.setDownloadDirectory(info.absolutePath())
                    download_item.setDownloadFileName(info.fileName())
                    download_item.accept()
                else:
                    download_item.cancel()
            except Exception as e:
                print(f"Download Error: {e}")

        def handle_print_request(frame):
            from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            dialog = QPrintDialog(printer, _window)
            if dialog.exec() == QPrintDialog.DialogCode.Accepted:
                frame.print(printer)

        browser.page().printRequested.connect(handle_print_request)
        browser.page().profile().downloadRequested.connect(handle_download)

        # Allow JS window.close() to shut down the app (Safe check)
        try:
            if hasattr(QWebEngineSettings.WebAttribute, "JavascriptCanCloseWindows"):
                browser.settings().setAttribute(
                    QWebEngineSettings.WebAttribute.JavascriptCanCloseWindows, True
                )
            browser.page().windowCloseRequested.connect(_window.close)
        except Exception as se:
            print(f"Settings Error: {se}")

        _window.setCentralWidget(browser)

        # Status bar showing LAN & Remote info
        status = QStatusBar()
        _window.setStatusBar(status)
        try:
            lan_ip = socket.gethostbyname(socket.gethostname())
            lan_text = f" [COMPUTER]  Local: http://{lan_ip}:8000"
        except Exception:
            lan_text = " [COMPUTER]  Local: http://127.0.0.1:8000"

        remote_text = ""
        if _public_url:
            remote_text = f"   |    [GLOBE]  Remote Login: {_public_url}"
        else:
            remote_text = "   |    [WARNING]  Remote Login: (Wait or Check config)"

        status.showMessage(
            f"{lan_text}{remote_text}      (Staff can access this URL from home/phone)"
        )

        # Periodically check for ngrok status update if not yet online
        def update_status():
            global _public_url
            if _public_url:
                try:
                    host_ip = socket.gethostbyname(socket.gethostname())
                    l_text = f" [COMPUTER]  Local: http://{host_ip}:8000"
                except Exception:
                    l_text = " [COMPUTER]  Local: http://127.0.0.1:8000"
                r_text = f"   |    [GLOBE]  Remote Login: {_public_url}"
                status.showMessage(
                    f"{l_text}{r_text}      (Staff can access this URL from home/phone)"
                )
                status_timer.stop()

        status_timer = QTimer(_window)
        status_timer.timeout.connect(update_status)
        if not _public_url:
            status_timer.start(2000)  # Every 2 seconds until it appears

        # Center on screen
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - _window.width()) // 2
        y = (screen.height() - _window.height()) // 2
        _window.move(x, y)

        # Start maximized (Safer than full screen on some setups)
        _window.showMaximized()
        if _splash:
            _splash.finish(_window)

    except Exception as ge:
        print(f"Global Launch Error: {ge}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


class WimpyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Wimpy De Ville Manager v1.0")
        self.setMinimumSize(1000, 700)

        # Set window icon if exists
        try:
            icon_path = os.path.join(os.getcwd(), "data", "templates", "logo.png")
            self.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_F11:
            if self.isFullScreen():
                self.showNormal()
            else:
                self.showFullScreen()
        super().keyPressEvent(event)


def main():
    global _splash, _public_url

    # 1. Start server in background
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()

    # 2. Start Ngrok tunnel in a separate thread so it doesn't block startup
    def bg_ngrok():
        global _public_url
        _public_url = start_ngrok()

    threading.Thread(target=bg_ngrok, daemon=True).start()

    # 3. Print LAN info immediately
    try:
        lan_ip = socket.gethostbyname(socket.gethostname())
        print(f"\n   [COMPUTER]  LAN Access: http://{lan_ip}:8000  (same WiFi network)")
    except Exception:
        pass

    # 4. Create Qt app
    qt_app = QApplication(sys.argv)
    qt_app.setApplicationName("Wimpy De Ville Manager")

    # 5. Splash screen
    pix = QPixmap(500, 200)
    pix.fill(QColor("#0f172a"))
    _splash = QSplashScreen(pix, Qt.WindowType.WindowStaysOnTopHint)

    lbl = QLabel("   [HOTEL]  Wimpy De Ville Manager\n  Starting up...", _splash)
    font = QFont("Arial", 14)
    font.setBold(True)
    lbl.setFont(font)
    lbl.setStyleSheet("color: white;")
    lbl.move(40, 60)
    lbl.resize(420, 80)
    _splash.show()
    qt_app.processEvents()

    # 6. Launch window after short delay
    QTimer.singleShot(300, launch_window)

    sys.exit(qt_app.exec())


if __name__ == "__main__":
    main()
