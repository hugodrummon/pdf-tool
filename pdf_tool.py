"""
PDF Tool — Compress, Merge, Rename, Redact, Flatten, and OCR PDFs locally.
Built for non-technical users in legal/admin environments.
No internet, no cloud, no third-party services. Everything stays on this machine.
"""

APP_VERSION = "1.5.9"
GITHUB_REPO = "hugodrummon/pdf-tool"
UPDATE_PUBLIC_KEY = "sw613yM42XKzroyOPRE19tMKJEqHQf2Ycne7S1rOMpU="
import sys

# Built exe: only ship these tabs. Running from source: show all tabs for development.
ENABLED_TABS = ["Compress", "Merge"] if getattr(sys, 'frozen', False) else None
import atexit
import os
import subprocess
import shutil
import tempfile

# Track temp dirs for cleanup on exit
_temp_dirs = []

def _tracked_mkdtemp():
    d = tempfile.mkdtemp()
    _temp_dirs.append(d)
    return d

def _secure_delete_file(path):
    """Overwrite file contents before deleting — prevents recovery of legal docs."""
    try:
        size = os.path.getsize(path)
        with open(path, "wb") as f:
            f.write(b'\x00' * size)
            f.flush()
            os.fsync(f.fileno())
        os.remove(path)
    except OSError:
        try:
            os.remove(path)
        except OSError:
            pass

def _cleanup_temp_dirs():
    for d in _temp_dirs:
        try:
            for root, dirs, files in os.walk(d):
                for fname in files:
                    _secure_delete_file(os.path.join(root, fname))
            shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass

atexit.register(_cleanup_temp_dirs)
import json
import webbrowser
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.error import URLError

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QProgressBar, QTabWidget,
    QListWidget, QListWidgetItem, QFileDialog, QMessageBox,
    QSizePolicy, QFrame, QSpacerItem, QAbstractItemView, QDialog,
    QTextEdit, QScrollArea, QSlider
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMimeData, QSize, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon, QDragEnterEvent, QDropEvent

from PyPDF2 import PdfMerger
import fitz  # PyMuPDF — for redaction and flattening

try:
    import pytesseract
    _bundle_dir = os.path.dirname(os.path.abspath(__file__))
    if getattr(sys, 'frozen', False):
        _bundle_dir = sys._MEIPASS
    _bundled_tess = os.path.join(_bundle_dir, "tesseract", "tesseract.exe")
    if os.path.isfile(_bundled_tess):
        pytesseract.pytesseract.tesseract_cmd = _bundled_tess
    HAS_TESSERACT = bool(shutil.which(pytesseract.pytesseract.tesseract_cmd or "tesseract") or os.path.isfile(_bundled_tess))
except ImportError:
    HAS_TESSERACT = False


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

TARGET_SIZE_MB = 10
TARGET_SIZE_BYTES = TARGET_SIZE_MB * 1024 * 1024


def human_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} bytes"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def get_bundle_dir() -> str:
    """Get the directory where bundled files are stored.
    Works both in development and when packaged with PyInstaller."""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def find_ghostscript() -> str:
    """Locate the Ghostscript executable — checks bundled location first."""
    bundle_dir = get_bundle_dir()

    # 1. Check bundled Ghostscript (inside the exe)
    bundled_gs = os.path.join(bundle_dir, "gs", "bin", "gswin64c.exe")
    if os.path.isfile(bundled_gs):
        return bundled_gs

    bundled_gs32 = os.path.join(bundle_dir, "gs", "bin", "gswin32c.exe")
    if os.path.isfile(bundled_gs32):
        return bundled_gs32

    # 2. Check PATH
    for name in ("gswin64c", "gswin32c", "gs"):
        path = shutil.which(name)
        if path:
            return path

    # 3. Check common Windows install locations
    for prog_dir in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        if not prog_dir:
            continue
        gs_root = os.path.join(prog_dir, "gs")
        if os.path.isdir(gs_root):
            for version_dir in sorted(os.listdir(gs_root), reverse=True):
                for exe in ("gswin64c.exe", "gswin32c.exe"):
                    candidate = os.path.join(gs_root, version_dir, "bin", exe)
                    if os.path.isfile(candidate):
                        return candidate

    # 4. macOS/Linux fallback for development
    path = shutil.which("gs")
    if path:
        return path

    return ""


def compress_pdf(input_path: str, output_path: str, gs_exe: str,
                 quality: str = "/ebook") -> bool:
    """Compress a PDF using Ghostscript. Returns True on success."""
    # Set GS_LIB so bundled Ghostscript can find its resource files
    env = os.environ.copy()
    bundle_dir = get_bundle_dir()
    gs_lib_path = os.path.join(bundle_dir, "gs", "lib")
    gs_resource_path = os.path.join(bundle_dir, "gs", "Resource")
    if os.path.isdir(gs_lib_path):
        env["GS_LIB"] = f"{gs_lib_path};{gs_resource_path}"

    # Get CPU count for multi-threaded rendering
    num_threads = min(os.cpu_count() or 2, 8)

    args = [
        gs_exe,
        "-dSAFER",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={quality}",
        "-dNOPAUSE",
        "-dBATCH",
        "-dQUIET",
        "-dNOGC",
        "-dBandBufferSpace=500000000",
        "-sBandListStorage=memory",
        f"-dNumRenderingThreads={num_threads}",
        "-dDetectDuplicateImages=true",
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        "-dColorImageDownsampleType=/Average",
        "-dGrayImageDownsampleType=/Average",
        "-dMonoImageDownsampleType=/Subsample",
        "-dPassThroughJPEGImages=true",
        "-dFastWebView=false",
        "-dColorConversionStrategy=/LeaveColorUnchanged",
        "-dAutoFilterColorImages=false",
        "-dAutoFilterGrayImages=false",
        "-dColorImageFilter=/DCTEncode",
        "-dGrayImageFilter=/DCTEncode",
        f"-sOutputFile={output_path}",
        input_path,
    ]

    startupinfo = None
    creationflags = 0
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        creationflags = subprocess.CREATE_NO_WINDOW | subprocess.ABOVE_NORMAL_PRIORITY_CLASS

    result = subprocess.run(
        args,
        startupinfo=startupinfo,
        capture_output=True,
        creationflags=creationflags,
        env=env,
    )
    return result.returncode == 0


# ---------------------------------------------------------------------------
# Quality levels to try (best quality first, increasingly aggressive)
# ---------------------------------------------------------------------------

QUALITY_LEVELS = ["/ebook", "/screen"]


# ---------------------------------------------------------------------------
# Worker threads
# ---------------------------------------------------------------------------

class CompressWorker(QThread):
    finished = pyqtSignal(bool, str, int, int)

    def __init__(self, input_path: str, gs_exe: str):
        super().__init__()
        self.input_path = input_path
        self.gs_exe = gs_exe

    def run(self):
        orig_size = os.path.getsize(self.input_path)

        if orig_size <= TARGET_SIZE_BYTES:
            self.finished.emit(True, self.input_path, orig_size, orig_size)
            return

        tmp_dir = _tracked_mkdtemp()
        output_path = os.path.join(tmp_dir, "compressed.pdf")

        # Pick quality: go straight to /screen for files well above target
        # to avoid running Ghostscript twice (/ebook then /screen fallback)
        if orig_size > TARGET_SIZE_BYTES * 2:
            quality = "/screen"
        else:
            quality = "/ebook"

        ok = compress_pdf(self.input_path, output_path, self.gs_exe, quality)
        if ok and os.path.isfile(output_path):
            new_size = os.path.getsize(output_path)
            # If /ebook wasn't enough and we haven't tried /screen yet, try it
            if new_size > TARGET_SIZE_BYTES and quality == "/ebook":
                ok2 = compress_pdf(self.input_path, output_path, self.gs_exe, "/screen")
                if ok2 and os.path.isfile(output_path):
                    new_size = os.path.getsize(output_path)
            self.finished.emit(True, output_path, orig_size, new_size)
        else:
            self.finished.emit(False, "", orig_size, 0)


class MergeWorker(QThread):
    finished = pyqtSignal(bool, str, int, int)

    def __init__(self, file_paths: list, gs_exe: str):
        super().__init__()
        self.file_paths = file_paths
        self.gs_exe = gs_exe

    def run(self):
        tmp_dir = _tracked_mkdtemp()
        merged_path = os.path.join(tmp_dir, "merged.pdf")
        combined_size = sum(os.path.getsize(p) for p in self.file_paths)

        try:
            merger = PdfMerger()
            for f in self.file_paths:
                merger.append(f)
            merger.write(merged_path)
            merger.close()
        except Exception:
            self.finished.emit(False, "", combined_size, 0)
            return

        merged_size = os.path.getsize(merged_path)

        if merged_size <= TARGET_SIZE_BYTES:
            self.finished.emit(True, merged_path, combined_size, merged_size)
            return

        compressed_path = os.path.join(tmp_dir, "compressed.pdf")

        # Pick quality: go straight to /screen for files well above target
        if merged_size > TARGET_SIZE_BYTES * 2:
            quality = "/screen"
        else:
            quality = "/ebook"

        ok = compress_pdf(merged_path, compressed_path, self.gs_exe, quality)
        if ok and os.path.isfile(compressed_path):
            new_size = os.path.getsize(compressed_path)
            if new_size > TARGET_SIZE_BYTES and quality == "/ebook":
                ok2 = compress_pdf(merged_path, compressed_path, self.gs_exe, "/screen")
                if ok2 and os.path.isfile(compressed_path):
                    new_size = os.path.getsize(compressed_path)
            self.finished.emit(True, compressed_path, combined_size, new_size)
        else:
            self.finished.emit(True, merged_path, combined_size, merged_size)


class UpdateChecker(QThread):
    """Check GitHub Releases for a newer version. Runs in background, silent on failure."""
    update_available = pyqtSignal(str, str, str)  # (latest_version, download_url, sig_url)

    def run(self):
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            req = Request(url, headers={"Accept": "application/vnd.github.v3+json"})
            with urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode())

            latest = data.get("tag_name", "").lstrip("v")
            if not latest:
                return

            # Compare versions
            current_parts = [int(x) for x in APP_VERSION.split(".")]
            latest_parts = [int(x) for x in latest.split(".")]
            if latest_parts > current_parts:
                # Find the installer and signature asset URLs
                download_url = data.get("html_url", "")
                sig_url = ""
                for asset in data.get("assets", []):
                    name = asset["name"]
                    if name.lower().endswith(".exe.sig"):
                        sig_url = asset["browser_download_url"]
                    elif "install" in name.lower() and name.endswith(".exe"):
                        download_url = asset["browser_download_url"]
                if download_url.startswith("https://github.com/"):
                    self.update_available.emit(latest, download_url, sig_url)
        except Exception:
            pass  # Silent fail — no internet, no problem


class UpdateDownloader(QThread):
    """Downloads the installer in the background and verifies its signature."""
    progress = pyqtSignal(int)  # percentage
    finished = pyqtSignal(bool, str)  # success, file_path

    def __init__(self, download_url: str, sig_url: str = ""):
        super().__init__()
        self.download_url = download_url
        self.sig_url = sig_url

    def run(self):
        try:
            tmp_dir = _tracked_mkdtemp()
            installer_path = os.path.join(tmp_dir, "Install PDF Tool.exe")
            req = Request(self.download_url)
            with urlopen(req, timeout=60) as resp:
                total = int(resp.headers.get("Content-Length", 0))
                downloaded = 0
                chunk_size = 256 * 1024  # 256 KB chunks
                with open(installer_path, "wb") as f:
                    while True:
                        chunk = resp.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total > 0:
                            self.progress.emit(int(downloaded * 100 / total))
            self.progress.emit(100)

            # Verify signature before trusting the installer
            if self.sig_url:
                try:
                    import hashlib
                    import base64
                    from nacl.signing import VerifyKey
                    from nacl.exceptions import BadSignatureError

                    sig_req = Request(self.sig_url)
                    with urlopen(sig_req, timeout=15) as resp:
                        lines = resp.read().decode().strip().splitlines()
                    if len(lines) < 2:
                        self.finished.emit(False, "")
                        return
                    sig_hex, expected_hash = lines[0], lines[1]

                    actual_hash = hashlib.sha256(
                        open(installer_path, "rb").read()
                    ).hexdigest()
                    if actual_hash != expected_hash:
                        self.finished.emit(False, "")
                        return

                    vk = VerifyKey(base64.b64decode(UPDATE_PUBLIC_KEY))
                    vk.verify(expected_hash.encode(), bytes.fromhex(sig_hex))
                except (BadSignatureError, Exception):
                    self.finished.emit(False, "")
                    return

            self.finished.emit(True, installer_path)
        except Exception:
            self.finished.emit(False, "")


class UpdateBanner(QFrame):
    """In-app banner: auto-downloads update, shows 'Restart to update' button."""

    def __init__(self, parent, latest_version, download_url, sig_url=""):
        super().__init__(parent)
        self.download_url = download_url
        self.sig_url = sig_url
        self.latest_version = latest_version
        self.installer_path = ""
        self.downloader = None

        self.setStyleSheet(
            "UpdateBanner { background-color: #e3f2fd; border: 1px solid #90caf9; "
            "border-radius: 8px; }")
        banner_layout = QHBoxLayout(self)
        banner_layout.setContentsMargins(16, 8, 16, 8)
        banner_layout.setSpacing(12)

        self.status_label = QLabel("Downloading update...")
        self.status_label.setFont(QFont("Segoe UI", 11))
        self.status_label.setStyleSheet("color: #1565C0; border: none; background: transparent;")
        banner_layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setMaximumWidth(150)
        self.progress.setMaximumHeight(14)
        banner_layout.addWidget(self.progress)

        self.restart_btn = QPushButton("Restart to update")
        self.restart_btn.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.restart_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; border: none; "
            "padding: 6px 16px; border-radius: 6px; }"
            "QPushButton:hover { background-color: #43A047; }")
        self.restart_btn.setCursor(Qt.PointingHandCursor)
        self.restart_btn.clicked.connect(self._do_restart_update)
        self.restart_btn.hide()
        banner_layout.addWidget(self.restart_btn)

        # Start downloading immediately
        self._start_download()

    def _start_download(self):
        self.downloader = UpdateDownloader(self.download_url, self.sig_url)
        self.downloader.progress.connect(self.progress.setValue)
        self.downloader.finished.connect(self._on_download_finished)
        self.downloader.start()

    def _on_download_finished(self, success, installer_path):
        if not success:
            self.status_label.setText("Update download failed — will retry next launch")
            self.status_label.setStyleSheet(
                "color: #c62828; border: none; background: transparent;")
            self.progress.hide()
            return

        self.installer_path = installer_path
        self.progress.hide()
        self.status_label.setText(f"v{self.latest_version} ready!")
        self.restart_btn.show()

    def _do_restart_update(self):
        self.restart_btn.setEnabled(False)
        self.status_label.setText("Restarting...")

        # Find the path to the current app executable and its install dir
        if getattr(sys, 'frozen', False):
            app_exe = sys.executable
            app_dir = os.path.dirname(app_exe)
        else:
            app_exe = os.path.abspath(sys.argv[0])
            app_dir = os.path.dirname(app_exe)

        # Create a batch script that:
        # 1. Waits for this app to fully close
        # 2. Runs the installer silently
        # 3. Waits for installer to finish, then relaunches
        bat_dir = _tracked_mkdtemp()
        bat_path = os.path.join(bat_dir, "update.bat")
        with open(bat_path, "w") as f:
            f.write(f'@echo off\n')
            f.write(f':waitloop\n')
            f.write(f'tasklist /FI "PID eq {os.getpid()}" 2>nul | find /I "PDF" >nul && (\n')
            f.write(f'  ping 127.0.0.1 -n 2 > nul\n')
            f.write(f'  goto waitloop\n')
            f.write(f')\n')
            f.write(f'ping 127.0.0.1 -n 5 > nul\n')
            f.write(f'"{self.installer_path}" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS /FORCECLOSEAPPLICATIONS\n')
            f.write(f'ping 127.0.0.1 -n 6 > nul\n')
            f.write(f'start "" "{app_exe}"\n')
            f.write(f'del "%~f0"\n')

        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        subprocess.Popen(
            ["cmd.exe", "/c", bat_path],
            startupinfo=startupinfo,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        QApplication.instance().quit()


# ---------------------------------------------------------------------------
# Custom drop zone widget
# ---------------------------------------------------------------------------

class DropZone(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self, label_text: str, accept_multiple: bool = False,
                 file_extensions: list = None, file_filter_name: str = "PDF"):
        super().__init__()
        self.accept_multiple = accept_multiple
        self._extensions = [e.lower() for e in (file_extensions or [".pdf"])]
        self._filter_name = file_filter_name
        self.setAcceptDrops(True)
        self.setMinimumHeight(160)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self._default_style = """
            DropZone {
                border: 3px dashed #aaaaaa;
                border-radius: 16px;
                background-color: #f9f9f9;
            }
        """
        self._hover_style = """
            DropZone {
                border: 3px dashed #4CAF50;
                border-radius: 16px;
                background-color: #e8f5e9;
            }
        """
        self.setStyleSheet(self._default_style)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        icon_label = QLabel("\U0001F4C4")
        icon_label.setFont(QFont("Segoe UI", 36))
        icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon_label)

        self.text_label = QLabel(label_text)
        self.text_label.setFont(QFont("Segoe UI", 14))
        self.text_label.setAlignment(Qt.AlignCenter)
        self.text_label.setWordWrap(True)
        self.text_label.setStyleSheet("color: #555555;")
        layout.addWidget(self.text_label)

        browse_btn = QPushButton("or click here to browse")
        browse_btn.setFont(QFont("Segoe UI", 11))
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.setStyleSheet("""
            QPushButton {
                border: none;
                color: #1976D2;
                text-decoration: underline;
                background: transparent;
            }
            QPushButton:hover { color: #0D47A1; }
        """)
        browse_btn.clicked.connect(self._browse)
        layout.addWidget(browse_btn, alignment=Qt.AlignCenter)

    def _browse(self):
        ext_str = " ".join(f"*{e}" for e in self._extensions)
        filter_str = f"{self._filter_name} Files ({ext_str})"
        if self.accept_multiple:
            paths, _ = QFileDialog.getOpenFileNames(
                self, f"Select {self._filter_name} files", "", filter_str)
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, f"Select a {self._filter_name} file", "", filter_str)
            paths = [path] if path else []
        if paths:
            self.files_dropped.emit(paths)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(self._hover_style)

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self._default_style)

    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet(self._default_style)
        paths = []
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if any(p.lower().endswith(ext) for ext in self._extensions):
                paths.append(p)
        if paths:
            if not self.accept_multiple:
                paths = paths[:1]
            self.files_dropped.emit(paths)
        else:
            ext_list = ", ".join(self._extensions)
            QMessageBox.warning(self, "Wrong file type",
                                f"Please drop a supported file ({ext_list}).")


# ---------------------------------------------------------------------------
# Shared styles
# ---------------------------------------------------------------------------

GLOBAL_STYLE = """
    QMainWindow, QWidget {
        background-color: #ffffff;
        font-family: "Segoe UI", Arial, sans-serif;
    }
    QTabWidget::pane {
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        background: white;
    }
    QTabBar::tab {
        padding: 10px 16px;
        font-size: 13px;
        font-weight: 500;
        border: none;
        border-bottom: 3px solid transparent;
        color: #666666;
        background: transparent;
        min-width: 70px;
    }
    QTabBar::tab:selected {
        color: #1976D2;
        border-bottom: 3px solid #1976D2;
    }
    QTabBar::tab:hover {
        color: #333333;
    }
    QPushButton {
        font-size: 14px;
        padding: 12px 32px;
        border-radius: 8px;
        font-weight: 500;
    }
    QLineEdit {
        font-size: 13px;
        padding: 8px 12px;
        border: 2px solid #e0e0e0;
        border-radius: 8px;
        background: white;
    }
    QLineEdit:focus {
        border-color: #1976D2;
    }
    QProgressBar {
        border: none;
        border-radius: 8px;
        background-color: #e0e0e0;
        height: 18px;
        text-align: center;
        font-size: 11px;
    }
    QProgressBar::chunk {
        background-color: #4CAF50;
        border-radius: 8px;
    }
    QListWidget {
        font-size: 13px;
        border: 2px solid #e0e0e0;
        border-radius: 8px;
        padding: 4px;
        background: white;
    }
    QListWidget::item {
        padding: 8px;
        border-radius: 4px;
    }
    QListWidget::item:selected {
        background-color: #e3f2fd;
        color: black;
    }
"""

BTN_PRIMARY = """
    QPushButton {
        background-color: #1976D2;
        color: white;
        border: none;
    }
    QPushButton:hover { background-color: #1565C0; }
    QPushButton:pressed { background-color: #0D47A1; }
    QPushButton:disabled { background-color: #BDBDBD; color: #888888; }
"""

BTN_SUCCESS = """
    QPushButton {
        background-color: #4CAF50;
        color: white;
        border: none;
    }
    QPushButton:hover { background-color: #43A047; }
    QPushButton:pressed { background-color: #388E3C; }
    QPushButton:disabled { background-color: #BDBDBD; color: #888888; }
"""

BTN_SECONDARY = """
    QPushButton {
        background-color: #f5f5f5;
        color: #333333;
        border: 1px solid #e0e0e0;
    }
    QPushButton:hover { background-color: #eeeeee; }
"""


# ---------------------------------------------------------------------------
# Compress tab
# ---------------------------------------------------------------------------

IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif"]

class CompressTab(QWidget):
    ALL_EXTENSIONS = [".pdf"] + IMAGE_EXTENSIONS

    def __init__(self, gs_exe: str):
        super().__init__()
        self.gs_exe = gs_exe
        self.input_path = ""
        self.output_tmp_path = ""
        self.output_ext = ".pdf"  # tracks output file extension
        self._is_image = False
        self.worker = None
        self._saved_files = []  # list of (display_name, full_path)

        # Smooth progress animation
        self._progress_timer = QTimer()
        self._progress_timer.setInterval(150)
        self._progress_timer.timeout.connect(self._tick_progress)
        self._progress_value = 0.0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        self.drop_zone = DropZone(
            "Drop your file here to compress it\n(PDF, JPEG, PNG, WebP, BMP, TIFF)",
            file_extensions=self.ALL_EXTENSIONS,
            file_filter_name="PDF / Image")
        self.drop_zone.files_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_zone)

        self.file_info = QLabel("")
        self.file_info.setFont(QFont("Segoe UI", 12))
        self.file_info.setAlignment(Qt.AlignCenter)
        self.file_info.setWordWrap(True)
        self.file_info.hide()
        layout.addWidget(self.file_info)

        # Quality slider (only shown for images)
        self.quality_frame = QFrame()
        q_layout = QHBoxLayout(self.quality_frame)
        q_layout.setContentsMargins(0, 0, 0, 0)
        q_label = QLabel("Quality:")
        q_label.setFont(QFont("Segoe UI", 12))
        q_layout.addWidget(q_label)
        self.quality_slider = QSlider(Qt.Horizontal)
        self.quality_slider.setRange(20, 95)
        self.quality_slider.setValue(75)
        self.quality_slider.setTickPosition(QSlider.TicksBelow)
        self.quality_slider.setTickInterval(15)
        q_layout.addWidget(self.quality_slider)
        self.quality_val = QLabel("75%")
        self.quality_val.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.quality_val.setMinimumWidth(40)
        q_layout.addWidget(self.quality_val)
        self.quality_slider.valueChanged.connect(
            lambda v: self.quality_val.setText(f"{v}%"))
        self.quality_frame.hide()
        layout.addWidget(self.quality_frame)

        self.compress_img_btn = QPushButton("Compress")
        self.compress_img_btn.setStyleSheet(BTN_PRIMARY)
        self.compress_img_btn.setCursor(Qt.PointingHandCursor)
        self.compress_img_btn.clicked.connect(self._start_image_compress)
        self.compress_img_btn.hide()
        layout.addWidget(self.compress_img_btn, alignment=Qt.AlignCenter)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.result_frame = QFrame()
        result_layout = QVBoxLayout(self.result_frame)
        result_layout.setSpacing(12)

        self.result_icon = QLabel()
        self.result_icon.setFont(QFont("Segoe UI", 48))
        self.result_icon.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(self.result_icon)

        self.result_text = QLabel("")
        self.result_text.setFont(QFont("Segoe UI", 13))
        self.result_text.setAlignment(Qt.AlignCenter)
        self.result_text.setWordWrap(True)
        result_layout.addWidget(self.result_text)

        self.size_warning = QLabel("")
        self.size_warning.setFont(QFont("Segoe UI", 12))
        self.size_warning.setAlignment(Qt.AlignCenter)
        self.size_warning.setStyleSheet("color: #e65100; font-weight: bold;")
        self.size_warning.setWordWrap(True)
        self.size_warning.hide()
        result_layout.addWidget(self.size_warning)

        save_label = QLabel("What would you like to call this file?")
        save_label.setFont(QFont("Segoe UI", 13))
        save_label.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(save_label)

        name_row = QHBoxLayout()
        name_row.setSpacing(8)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type a name for the file")
        self.name_input.setMinimumHeight(36)
        name_row.addWidget(self.name_input)
        self.ext_label = QLabel(".pdf")
        self.ext_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.ext_label.setStyleSheet("color: #888;")
        name_row.addWidget(self.ext_label)
        result_layout.addLayout(name_row)

        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet(BTN_SUCCESS)
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self._save)
        result_layout.addWidget(self.save_btn, alignment=Qt.AlignCenter)

        self.result_frame.hide()
        layout.addWidget(self.result_frame)

        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Segoe UI", 13))
        self.error_label.setStyleSheet("color: #d32f2f;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)

        # --- History list of completed files ---
        self.history_label = QLabel("Completed files")
        self.history_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.history_label.setStyleSheet("color: #555; margin-top: 8px;")
        self.history_label.hide()
        layout.addWidget(self.history_label)

        self.history_list = QListWidget()
        self.history_list.setFont(QFont("Segoe UI", 11))
        self.history_list.setAlternatingRowColors(True)
        self.history_list.setStyleSheet(
            "QListWidget { border: 1px solid #e0e0e0; border-radius: 6px; background: #fafafa; }"
            "QListWidget::item { padding: 6px 10px; }"
            "QListWidget::item:selected { background-color: #e3f2fd; color: black; }"
            "QListWidget::item:hover { background-color: #f0f0f0; }"
        )
        self.history_list.setMaximumHeight(140)
        self.history_list.itemDoubleClicked.connect(self._open_history_item)
        self.history_list.hide()
        layout.addWidget(self.history_list)

        layout.addStretch()

    def _is_image_file(self, path):
        return any(path.lower().endswith(e) for e in IMAGE_EXTENSIONS)

    def _reset(self):
        self._progress_timer.stop()
        self.result_frame.hide()
        self.error_label.hide()
        self.progress.hide()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self.size_warning.hide()
        self.quality_frame.hide()
        self.compress_img_btn.hide()
        self.save_btn.setEnabled(True)
        self.name_input.setEnabled(True)

    def _on_file_dropped(self, paths):
        self._reset()
        self.input_path = paths[0]
        self._is_image = self._is_image_file(self.input_path)
        fname = os.path.basename(self.input_path)
        fsize = human_size(os.path.getsize(self.input_path))
        self.file_info.setText(f"Selected: {fname} ({fsize})")
        self.file_info.show()

        if self._is_image:
            # Show quality slider and compress button for images
            self.quality_frame.show()
            self.compress_img_btn.show()
            return

        # --- PDF flow ---
        if os.path.getsize(self.input_path) <= TARGET_SIZE_BYTES:
            self.output_ext = ".pdf"
            self.ext_label.setText(".pdf")
            self.result_icon.setText("\u2705")
            self.result_text.setText(
                f'<div style="text-align:center;">'
                f'This file is already under {TARGET_SIZE_MB} MB!<br>'
                f'<span style="color:#888;">Size:</span> <b style="color:#4CAF50;">{fsize}</b>'
                f'</div>')
            self.name_input.setText(Path(self.input_path).stem)
            self.output_tmp_path = self.input_path
            self.result_frame.show()
            return

        self.progress.show()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self._progress_timer.start()
        self.drop_zone.setEnabled(False)

        self.worker = CompressWorker(self.input_path, self.gs_exe)
        self.worker.finished.connect(self._on_pdf_finished)
        self.worker.start()

    def _start_image_compress(self):
        if not self.input_path:
            return
        self.quality_frame.hide()
        self.compress_img_btn.hide()
        self.progress.show()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self._progress_timer.start()
        self.drop_zone.setEnabled(False)

        self.worker = ImageCompressWorker(self.input_path, self.quality_slider.value())
        self.worker.finished.connect(self._on_image_finished)
        self.worker.start()

    def _tick_progress(self):
        """Smoothly advance progress bar — never stops moving."""
        if self._progress_value < 70:
            self._progress_value += 1.2
        elif self._progress_value < 90:
            self._progress_value += 0.4
        elif self._progress_value < 99:
            self._progress_value += 0.05
        self.progress.setValue(int(self._progress_value))

    def _on_pdf_finished(self, success, output_path, orig_size, new_size):
        self._progress_timer.stop()
        self.progress.setValue(100)
        self.drop_zone.setEnabled(True)
        self.progress.hide()

        if not success:
            self.error_label.setText(
                "Something went wrong \u2014 please try again or contact your IT team.")
            self.error_label.show()
            return

        self.output_tmp_path = output_path
        self.output_ext = ".pdf"
        self.ext_label.setText(".pdf")
        self.result_icon.setText("\u2705")
        self.result_text.setText(
            f'<div style="text-align:center;">'
            f'<span style="color:#888;">Original size:</span> <b>{human_size(orig_size)}</b><br>'
            f'<span style="color:#888;">Compressed size:</span> <b style="color:#4CAF50;">{human_size(new_size)}</b>'
            f'</div>')

        if new_size > TARGET_SIZE_BYTES:
            self.size_warning.setText(
                f"Note: The compressed file is {human_size(new_size)}, still over {TARGET_SIZE_MB} MB.\n"
                "The original PDF may contain high-resolution scans.")
            self.size_warning.show()

        self.name_input.setText(Path(self.input_path).stem + " - Compressed")
        self.result_frame.show()

    def _on_image_finished(self, success, output_path, orig_size, new_size, fmt):
        self._progress_timer.stop()
        self.progress.setValue(100)
        self.drop_zone.setEnabled(True)
        self.progress.hide()

        if not success:
            self.error_label.setText(
                "Something went wrong \u2014 please try again or contact your IT team.")
            self.error_label.show()
            self.quality_frame.show()
            self.compress_img_btn.show()
            return

        self.output_tmp_path = output_path
        self.output_ext = ".png" if fmt == "PNG" else ".jpg"
        self.ext_label.setText(self.output_ext)
        reduction = 0 if orig_size == 0 else int((1 - new_size / orig_size) * 100)
        self.result_icon.setText("\u2705")
        self.result_text.setText(
            f'<div style="text-align:center;">'
            f'<span style="color:#888;">Original:</span> <b>{human_size(orig_size)}</b><br>'
            f'<span style="color:#888;">Compressed:</span> <b style="color:#4CAF50;">'
            f'{human_size(new_size)}</b>  '
            f'<span style="color:#888;">({reduction}% smaller)</span>'
            f'</div>')

        self.name_input.setText(Path(self.input_path).stem + " - Compressed")
        self.result_frame.show()

    def _save(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Name needed", "Please type a name for the file.")
            return
        if not name.lower().endswith(self.output_ext):
            name += self.output_ext

        dest_dir = os.path.dirname(self.input_path)
        ext = self.output_ext
        if ext == ".pdf":
            file_filter = "PDF Files (*.pdf)"
            title = "Save compressed PDF"
        elif ext == ".png":
            file_filter = "PNG Files (*.png)"
            title = "Save compressed image"
        else:
            file_filter = "JPEG Files (*.jpg *.jpeg)"
            title = "Save compressed image"

        dest, _ = QFileDialog.getSaveFileName(
            self, title, os.path.join(dest_dir, name), file_filter)
        if not dest:
            return

        try:
            shutil.copy2(self.output_tmp_path, dest)
            saved_name = os.path.basename(dest)
            saved_size = human_size(os.path.getsize(dest))

            # Add to history list
            self._saved_files.append((saved_name, dest))
            self.history_list.addItem(f"\u2705  {saved_name}  ({saved_size})")
            self.history_label.show()
            self.history_list.show()

            # Reset the working area for the next file
            self.result_frame.hide()
            self.file_info.hide()
            self.size_warning.hide()
            self.drop_zone.setEnabled(True)
        except Exception:
            self.error_label.setText(
                "Something went wrong while saving \u2014 please try again or contact your IT team.")
            self.error_label.show()

    def _open_history_item(self, item):
        """Double-click a completed file to open its folder in Explorer."""
        idx = self.history_list.row(item)
        if 0 <= idx < len(self._saved_files):
            path = self._saved_files[idx][1]
            if os.path.exists(path):
                import subprocess
                subprocess.Popen(["explorer", "/select,", path])


# ---------------------------------------------------------------------------
# Merge tab
# ---------------------------------------------------------------------------

class MergeTab(QWidget):
    def __init__(self, gs_exe: str):
        super().__init__()
        self.gs_exe = gs_exe
        self.file_paths = []
        self.output_tmp_path = ""
        self.worker = None
        self._saved_files = []  # list of (display_name, full_path)

        # Smooth progress animation
        self._progress_timer = QTimer()
        self._progress_timer.setInterval(150)
        self._progress_timer.timeout.connect(self._tick_progress)
        self._progress_value = 0.0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        self.drop_zone = DropZone("Drop your PDF files here\n(you can drop several at once)",
                                  accept_multiple=True)
        self.drop_zone.files_dropped.connect(self._on_files_dropped)
        layout.addWidget(self.drop_zone)

        list_label = QLabel("Files to merge (drag to reorder, or use the buttons):")
        list_label.setFont(QFont("Segoe UI", 12))
        list_label.hide()
        self.list_label = list_label
        layout.addWidget(list_label)

        list_row = QHBoxLayout()
        self.file_list = QListWidget()
        self.file_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.file_list.setMinimumHeight(100)
        self.file_list.hide()
        list_row.addWidget(self.file_list)

        btn_col = QVBoxLayout()
        self.up_btn = QPushButton("\u25B2 Up")
        self.up_btn.setStyleSheet(BTN_SECONDARY)
        self.up_btn.setCursor(Qt.PointingHandCursor)
        self.up_btn.clicked.connect(self._move_up)
        self.up_btn.hide()
        btn_col.addWidget(self.up_btn)

        self.down_btn = QPushButton("\u25BC Down")
        self.down_btn.setStyleSheet(BTN_SECONDARY)
        self.down_btn.setCursor(Qt.PointingHandCursor)
        self.down_btn.clicked.connect(self._move_down)
        self.down_btn.hide()
        btn_col.addWidget(self.down_btn)

        self.remove_btn = QPushButton("\u2715 Remove")
        self.remove_btn.setStyleSheet(BTN_SECONDARY)
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        self.remove_btn.clicked.connect(self._remove_selected)
        self.remove_btn.hide()
        btn_col.addWidget(self.remove_btn)

        btn_col.addStretch()
        list_row.addLayout(btn_col)
        layout.addLayout(list_row)

        self.merge_btn = QPushButton("Merge into one PDF")
        self.merge_btn.setStyleSheet(BTN_PRIMARY)
        self.merge_btn.setCursor(Qt.PointingHandCursor)
        self.merge_btn.clicked.connect(self._start_merge)
        self.merge_btn.hide()
        layout.addWidget(self.merge_btn, alignment=Qt.AlignCenter)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.result_frame = QFrame()
        result_layout = QVBoxLayout(self.result_frame)
        result_layout.setSpacing(12)

        self.result_icon = QLabel()
        self.result_icon.setFont(QFont("Segoe UI", 48))
        self.result_icon.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(self.result_icon)

        self.result_text = QLabel("")
        self.result_text.setFont(QFont("Segoe UI", 13))
        self.result_text.setAlignment(Qt.AlignCenter)
        self.result_text.setWordWrap(True)
        result_layout.addWidget(self.result_text)

        self.size_warning = QLabel("")
        self.size_warning.setFont(QFont("Segoe UI", 12))
        self.size_warning.setAlignment(Qt.AlignCenter)
        self.size_warning.setStyleSheet("color: #e65100; font-weight: bold;")
        self.size_warning.setWordWrap(True)
        self.size_warning.hide()
        result_layout.addWidget(self.size_warning)

        save_label = QLabel("What would you like to call this file?")
        save_label.setFont(QFont("Segoe UI", 13))
        save_label.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(save_label)

        name_row = QHBoxLayout()
        name_row.setSpacing(8)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type a name for the file")
        self.name_input.setMinimumHeight(36)
        name_row.addWidget(self.name_input)
        pdf_label = QLabel(".pdf")
        pdf_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        pdf_label.setStyleSheet("color: #888;")
        name_row.addWidget(pdf_label)
        result_layout.addLayout(name_row)

        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet(BTN_SUCCESS)
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self._save)
        result_layout.addWidget(self.save_btn, alignment=Qt.AlignCenter)

        self.result_frame.hide()
        layout.addWidget(self.result_frame)

        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Segoe UI", 13))
        self.error_label.setStyleSheet("color: #d32f2f;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)

        # --- History list of completed files ---
        self.history_label = QLabel("Completed files")
        self.history_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.history_label.setStyleSheet("color: #555; margin-top: 8px;")
        self.history_label.hide()
        layout.addWidget(self.history_label)

        self.history_list = QListWidget()
        self.history_list.setFont(QFont("Segoe UI", 11))
        self.history_list.setAlternatingRowColors(True)
        self.history_list.setStyleSheet(
            "QListWidget { border: 1px solid #e0e0e0; border-radius: 6px; background: #fafafa; }"
            "QListWidget::item { padding: 6px 10px; }"
            "QListWidget::item:selected { background-color: #e3f2fd; color: black; }"
            "QListWidget::item:hover { background-color: #f0f0f0; }"
        )
        self.history_list.setMaximumHeight(140)
        self.history_list.itemDoubleClicked.connect(self._open_history_item)
        self.history_list.hide()
        layout.addWidget(self.history_list)

        layout.addStretch()

    def _show_list_controls(self, show=True):
        for w in (self.list_label, self.file_list, self.up_btn,
                  self.down_btn, self.remove_btn, self.merge_btn):
            w.setVisible(show)

    def _refresh_list(self):
        self.file_list.clear()
        for p in self.file_paths:
            name = os.path.basename(p)
            size = human_size(os.path.getsize(p))
            self.file_list.addItem(f"{name}  ({size})")

    def _on_files_dropped(self, paths):
        self.result_frame.hide()
        self.error_label.hide()
        self.size_warning.hide()
        for p in paths:
            if p not in self.file_paths:
                self.file_paths.append(p)
        self._refresh_list()
        if self.file_paths:
            self._show_list_controls(True)
            self.merge_btn.setEnabled(len(self.file_paths) >= 2)

    def _move_up(self):
        row = self.file_list.currentRow()
        if row > 0:
            self.file_paths[row], self.file_paths[row - 1] = \
                self.file_paths[row - 1], self.file_paths[row]
            self._refresh_list()
            self.file_list.setCurrentRow(row - 1)

    def _move_down(self):
        row = self.file_list.currentRow()
        if 0 <= row < len(self.file_paths) - 1:
            self.file_paths[row], self.file_paths[row + 1] = \
                self.file_paths[row + 1], self.file_paths[row]
            self._refresh_list()
            self.file_list.setCurrentRow(row + 1)

    def _remove_selected(self):
        row = self.file_list.currentRow()
        if 0 <= row < len(self.file_paths):
            self.file_paths.pop(row)
            self._refresh_list()
            if not self.file_paths:
                self._show_list_controls(False)
            else:
                self.merge_btn.setEnabled(len(self.file_paths) >= 2)

    def _start_merge(self):
        if len(self.file_paths) < 2:
            return
        self.result_frame.hide()
        self.error_label.hide()
        self.size_warning.hide()
        self.progress.show()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self._progress_timer.start()
        self.merge_btn.setEnabled(False)
        self.drop_zone.setEnabled(False)

        self.worker = MergeWorker(list(self.file_paths), self.gs_exe)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _tick_progress(self):
        """Smoothly advance progress bar — never stops moving."""
        if self._progress_value < 70:
            self._progress_value += 1.2
        elif self._progress_value < 90:
            self._progress_value += 0.4
        elif self._progress_value < 99:
            self._progress_value += 0.05
        self.progress.setValue(int(self._progress_value))

    def _on_finished(self, success, output_path, combined_size, final_size):
        self._progress_timer.stop()
        self.progress.setValue(100)
        self.progress.hide()
        self.merge_btn.setEnabled(True)
        self.drop_zone.setEnabled(True)

        if not success:
            self.error_label.setText(
                "Something went wrong \u2014 please try again or contact your IT team.")
            self.error_label.show()
            return

        self.output_tmp_path = output_path
        self.result_icon.setText("\u2705")
        self.result_text.setText(
            f'<div style="text-align:center;">'
            f'<span style="color:#888;">Combined original size:</span> <b>{human_size(combined_size)}</b><br>'
            f'<span style="color:#888;">Final size:</span> <b style="color:#4CAF50;">{human_size(final_size)}</b>'
            f'</div>')

        if final_size > TARGET_SIZE_BYTES:
            self.size_warning.setText(
                f"Note: The final file is {human_size(final_size)}, still over {TARGET_SIZE_MB} MB.\n"
                "The original PDFs may contain high-resolution scans.")
            self.size_warning.show()

        self.name_input.setText("Merged Document")
        self.save_btn.setEnabled(True)
        self.name_input.setEnabled(True)
        self.result_frame.show()

    def _save(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Name needed", "Please type a name for the file.")
            return
        if not name.lower().endswith(".pdf"):
            name += ".pdf"

        dest_dir = os.path.dirname(self.file_paths[0]) if self.file_paths else ""
        dest, _ = QFileDialog.getSaveFileName(
            self, "Save merged PDF", os.path.join(dest_dir, name),
            "PDF Files (*.pdf)")
        if not dest:
            return

        try:
            shutil.copy2(self.output_tmp_path, dest)
            saved_name = os.path.basename(dest)
            saved_size = human_size(os.path.getsize(dest))

            # Add to history list
            self._saved_files.append((saved_name, dest))
            self.history_list.addItem(f"\u2705  {saved_name}  ({saved_size})")
            self.history_label.show()
            self.history_list.show()

            # Reset the working area for the next merge
            self.result_frame.hide()
            self.size_warning.hide()
            self.file_paths.clear()
            self._refresh_list()
            self._show_list_controls(False)
            self.drop_zone.setEnabled(True)
        except Exception:
            self.error_label.setText(
                "Something went wrong while saving \u2014 please try again or contact your IT team.")
            self.error_label.show()

    def _open_history_item(self, item):
        """Double-click a completed file to open its folder in Explorer."""
        idx = self.history_list.row(item)
        if 0 <= idx < len(self._saved_files):
            path = self._saved_files[idx][1]
            if os.path.exists(path):
                import subprocess
                subprocess.Popen(["explorer", "/select,", path])


# ---------------------------------------------------------------------------
# Rename tab
# ---------------------------------------------------------------------------

class RenameTab(QWidget):
    def __init__(self):
        super().__init__()
        self.input_path = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        self.drop_zone = DropZone("Drop a PDF here to rename it")
        self.drop_zone.files_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_zone)

        self.rename_frame = QFrame()
        rename_layout = QVBoxLayout(self.rename_frame)
        rename_layout.setSpacing(12)

        self.current_name = QLabel("")
        self.current_name.setFont(QFont("Segoe UI", 13))
        self.current_name.setAlignment(Qt.AlignCenter)
        self.current_name.setWordWrap(True)
        rename_layout.addWidget(self.current_name)

        new_label = QLabel("New name:")
        new_label.setFont(QFont("Segoe UI", 13))
        rename_layout.addWidget(new_label)

        name_row = QHBoxLayout()
        name_row.setSpacing(8)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type the new name")
        self.name_input.setMinimumHeight(36)
        name_row.addWidget(self.name_input)
        pdf_label = QLabel(".pdf")
        pdf_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        pdf_label.setStyleSheet("color: #888;")
        name_row.addWidget(pdf_label)
        rename_layout.addLayout(name_row)

        self.rename_btn = QPushButton("Rename")
        self.rename_btn.setStyleSheet(BTN_PRIMARY)
        self.rename_btn.setCursor(Qt.PointingHandCursor)
        self.rename_btn.clicked.connect(self._rename)
        rename_layout.addWidget(self.rename_btn, alignment=Qt.AlignCenter)

        self.rename_frame.hide()
        layout.addWidget(self.rename_frame)

        self.result_icon = QLabel()
        self.result_icon.setFont(QFont("Segoe UI", 48))
        self.result_icon.setAlignment(Qt.AlignCenter)
        self.result_icon.hide()
        layout.addWidget(self.result_icon)

        self.result_text = QLabel("")
        self.result_text.setFont(QFont("Segoe UI", 13))
        self.result_text.setAlignment(Qt.AlignCenter)
        self.result_text.setWordWrap(True)
        self.result_text.hide()
        layout.addWidget(self.result_text)

        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Segoe UI", 13))
        self.error_label.setStyleSheet("color: #d32f2f;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)

        layout.addStretch()

    def _on_file_dropped(self, paths):
        self.input_path = paths[0]
        self.result_icon.hide()
        self.result_text.hide()
        self.error_label.hide()

        fname = os.path.basename(self.input_path)
        self.current_name.setText(f"Current name: {fname}")
        self.name_input.setText(Path(self.input_path).stem)
        self.name_input.setEnabled(True)
        self.rename_btn.setEnabled(True)
        self.rename_frame.show()

    def _rename(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Name needed", "Please type a name for the file.")
            return
        if not name.lower().endswith(".pdf"):
            name += ".pdf"

        dest_dir = os.path.dirname(self.input_path)
        dest = os.path.join(dest_dir, name)

        if os.path.abspath(dest) == os.path.abspath(self.input_path):
            self.result_icon.setText("\u2705")
            self.result_icon.show()
            self.result_text.setText("The name is the same \u2014 no change needed.")
            self.result_text.show()
            return

        if os.path.exists(dest):
            reply = QMessageBox.question(
                self, "File already exists",
                f"A file called \"{name}\" already exists.\n\nDo you want to replace it?",
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return

        try:
            os.rename(self.input_path, dest)
            self.input_path = dest
            self.result_icon.setText("\u2705")
            self.result_icon.show()
            self.result_text.setText(f"Renamed to: {name}")
            self.result_text.show()
            self.rename_btn.setEnabled(False)
            self.name_input.setEnabled(False)
        except Exception:
            self.error_label.setText(
                "Something went wrong \u2014 please try again or contact your IT team.")
            self.error_label.show()


# ---------------------------------------------------------------------------
# Flatten tab
# ---------------------------------------------------------------------------

class FlattenWorker(QThread):
    finished = pyqtSignal(bool, str, int, int)

    def __init__(self, input_path: str):
        super().__init__()
        self.input_path = input_path

    def run(self):
        orig_size = os.path.getsize(self.input_path)
        try:
            tmp_dir = _tracked_mkdtemp()
            output_path = os.path.join(tmp_dir, "flattened.pdf")
            doc = fitz.open(self.input_path)
            for page in doc:
                annots = list(page.annots()) if page.annots() else []
                for annot in annots:
                    annot.set_flags(fitz.PDF_ANNOT_IS_PRINT)
                if page.first_redact_annot:
                    page.apply_redactions()
            doc.save(output_path, garbage=4, deflate=True)
            doc.close()
            new_size = os.path.getsize(output_path)
            self.finished.emit(True, output_path, orig_size, new_size)
        except Exception:
            self.finished.emit(False, "", orig_size, 0)


class FlattenTab(QWidget):
    def __init__(self):
        super().__init__()
        self.input_path = ""
        self.output_tmp_path = ""
        self.worker = None
        self._progress_timer = QTimer()
        self._progress_timer.setInterval(150)
        self._progress_timer.timeout.connect(self._tick_progress)
        self._progress_value = 0.0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        self.drop_zone = DropZone("Drop your PDF here to flatten it")
        self.drop_zone.files_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_zone)

        info = QLabel("Flattening makes annotations, stamps, form fields,\nand comments permanent and uneditable.")
        info.setFont(QFont("Segoe UI", 11))
        info.setAlignment(Qt.AlignCenter)
        info.setStyleSheet("color: #888888;")
        info.setWordWrap(True)
        layout.addWidget(info)

        self.file_info = QLabel("")
        self.file_info.setFont(QFont("Segoe UI", 12))
        self.file_info.setAlignment(Qt.AlignCenter)
        self.file_info.setWordWrap(True)
        self.file_info.hide()
        layout.addWidget(self.file_info)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.result_frame = QFrame()
        result_layout = QVBoxLayout(self.result_frame)
        result_layout.setSpacing(12)

        self.result_icon = QLabel()
        self.result_icon.setFont(QFont("Segoe UI", 48))
        self.result_icon.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(self.result_icon)

        self.result_text = QLabel("")
        self.result_text.setFont(QFont("Segoe UI", 13))
        self.result_text.setAlignment(Qt.AlignCenter)
        self.result_text.setWordWrap(True)
        result_layout.addWidget(self.result_text)

        save_label = QLabel("What would you like to call this file?")
        save_label.setFont(QFont("Segoe UI", 13))
        save_label.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(save_label)

        name_row = QHBoxLayout()
        name_row.setSpacing(8)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type a name for the file")
        self.name_input.setMinimumHeight(36)
        name_row.addWidget(self.name_input)
        pdf_label = QLabel(".pdf")
        pdf_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        pdf_label.setStyleSheet("color: #888;")
        name_row.addWidget(pdf_label)
        result_layout.addLayout(name_row)

        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet(BTN_SUCCESS)
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self._save)
        result_layout.addWidget(self.save_btn, alignment=Qt.AlignCenter)

        self.result_frame.hide()
        layout.addWidget(self.result_frame)

        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Segoe UI", 13))
        self.error_label.setStyleSheet("color: #d32f2f;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)
        layout.addStretch()

    def _on_file_dropped(self, paths):
        self._progress_timer.stop()
        self.result_frame.hide()
        self.error_label.hide()
        self.progress.hide()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self.input_path = paths[0]
        fname = os.path.basename(self.input_path)
        fsize = human_size(os.path.getsize(self.input_path))
        self.file_info.setText(f"Selected: {fname} ({fsize})")
        self.file_info.show()
        self.progress.show()
        self._progress_timer.start()
        self.drop_zone.setEnabled(False)
        self.worker = FlattenWorker(self.input_path)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _tick_progress(self):
        remaining = 90.0 - self._progress_value
        self._progress_value += remaining * 0.05
        self.progress.setValue(int(self._progress_value))

    def _on_finished(self, success, output_path, orig_size, new_size):
        self._progress_timer.stop()
        self.progress.setValue(100)
        self.drop_zone.setEnabled(True)
        self.progress.hide()
        if not success:
            self.error_label.setText("Something went wrong \u2014 please try again or contact your IT team.")
            self.error_label.show()
            return
        self.output_tmp_path = output_path
        self.result_icon.setText("\u2705")
        self.result_text.setText(
            f'<div style="text-align:center;">'
            f'Flattened successfully!<br>'
            f'<span style="color:#888;">Size:</span> <b style="color:#4CAF50;">{human_size(new_size)}</b>'
            f'</div>')
        self.name_input.setText(Path(self.input_path).stem + " - Flattened")
        self.save_btn.setEnabled(True)
        self.name_input.setEnabled(True)
        self.result_frame.show()

    def _save(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Name needed", "Please type a name for the file.")
            return
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        dest_dir = os.path.dirname(self.input_path)
        dest, _ = QFileDialog.getSaveFileName(self, "Save flattened PDF", os.path.join(dest_dir, name), "PDF Files (*.pdf)")
        if not dest:
            return
        try:
            shutil.copy2(self.output_tmp_path, dest)
            saved_name = os.path.basename(dest)
            self.result_icon.setText("\u2705")
            self.result_text.setText(
                f'<div style="text-align:center;">'
                f'<span style="color:#888;">Saved as:</span> <b>{saved_name}</b><br>'
                f'<span style="color:#888;">Size:</span> <b style="color:#4CAF50;">{human_size(os.path.getsize(dest))}</b>'
                f'</div>')
            self.save_btn.setEnabled(False)
            self.name_input.setEnabled(False)
        except Exception:
            self.error_label.setText("Something went wrong while saving \u2014 please try again or contact your IT team.")
            self.error_label.show()


# ---------------------------------------------------------------------------
# Redact tab
# ---------------------------------------------------------------------------

class RedactWorker(QThread):
    finished = pyqtSignal(bool, str, int, int, int)

    def __init__(self, input_path: str, search_terms: list):
        super().__init__()
        self.input_path = input_path
        self.search_terms = search_terms

    def run(self):
        orig_size = os.path.getsize(self.input_path)
        try:
            tmp_dir = _tracked_mkdtemp()
            output_path = os.path.join(tmp_dir, "redacted.pdf")
            doc = fitz.open(self.input_path)
            total_redactions = 0
            for page in doc:
                for term in self.search_terms:
                    areas = page.search_for(term)
                    for area in areas:
                        page.add_redact_annot(area, fill=(0, 0, 0))
                        total_redactions += 1
                page.apply_redactions()
            doc.save(output_path, garbage=4, deflate=True)
            doc.close()
            new_size = os.path.getsize(output_path)
            self.finished.emit(True, output_path, orig_size, new_size, total_redactions)
        except Exception:
            self.finished.emit(False, "", orig_size, 0, 0)


class RedactTab(QWidget):
    def __init__(self):
        super().__init__()
        self.input_path = ""
        self.output_tmp_path = ""
        self.worker = None
        self._progress_timer = QTimer()
        self._progress_timer.setInterval(150)
        self._progress_timer.timeout.connect(self._tick_progress)
        self._progress_value = 0.0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        self.drop_zone = DropZone("Drop your PDF here to redact it")
        self.drop_zone.files_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_zone)

        info = QLabel("Redaction permanently removes sensitive text from the PDF.\nThe original content cannot be recovered.")
        info.setFont(QFont("Segoe UI", 11))
        info.setAlignment(Qt.AlignCenter)
        info.setStyleSheet("color: #888888;")
        info.setWordWrap(True)
        layout.addWidget(info)

        self.file_info = QLabel("")
        self.file_info.setFont(QFont("Segoe UI", 12))
        self.file_info.setAlignment(Qt.AlignCenter)
        self.file_info.setWordWrap(True)
        self.file_info.hide()
        layout.addWidget(self.file_info)

        self.search_frame = QFrame()
        search_layout = QVBoxLayout(self.search_frame)
        search_layout.setSpacing(12)

        search_label = QLabel("Enter text to redact (one per line):")
        search_label.setFont(QFont("Segoe UI", 13))
        search_layout.addWidget(search_label)

        self.search_input = QTextEdit()
        self.search_input.setFont(QFont("Segoe UI", 12))
        self.search_input.setPlaceholderText("e.g.\nJohn Smith\n555-123-4567\nConfidential")
        self.search_input.setMaximumHeight(120)
        self.search_input.setStyleSheet("QTextEdit { border: 2px solid #e0e0e0; border-radius: 8px; padding: 8px; background: white; } QTextEdit:focus { border-color: #1976D2; }")
        search_layout.addWidget(self.search_input)

        self.redact_btn = QPushButton("Redact")
        self.redact_btn.setStyleSheet("QPushButton { background-color: #d32f2f; color: white; border: none; padding: 12px 32px; border-radius: 8px; font-size: 14px; font-weight: 500; } QPushButton:hover { background-color: #c62828; } QPushButton:pressed { background-color: #b71c1c; } QPushButton:disabled { background-color: #BDBDBD; color: #888888; }")
        self.redact_btn.setCursor(Qt.PointingHandCursor)
        self.redact_btn.clicked.connect(self._start_redact)
        search_layout.addWidget(self.redact_btn, alignment=Qt.AlignCenter)

        self.search_frame.hide()
        layout.addWidget(self.search_frame)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.result_frame = QFrame()
        result_layout = QVBoxLayout(self.result_frame)
        result_layout.setSpacing(12)

        self.result_icon = QLabel()
        self.result_icon.setFont(QFont("Segoe UI", 48))
        self.result_icon.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(self.result_icon)

        self.result_text = QLabel("")
        self.result_text.setFont(QFont("Segoe UI", 13))
        self.result_text.setAlignment(Qt.AlignCenter)
        self.result_text.setWordWrap(True)
        result_layout.addWidget(self.result_text)

        save_label = QLabel("What would you like to call this file?")
        save_label.setFont(QFont("Segoe UI", 13))
        save_label.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(save_label)

        name_row = QHBoxLayout()
        name_row.setSpacing(8)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type a name for the file")
        self.name_input.setMinimumHeight(36)
        name_row.addWidget(self.name_input)
        pdf_label = QLabel(".pdf")
        pdf_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        pdf_label.setStyleSheet("color: #888;")
        name_row.addWidget(pdf_label)
        result_layout.addLayout(name_row)

        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet(BTN_SUCCESS)
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self._save)
        result_layout.addWidget(self.save_btn, alignment=Qt.AlignCenter)

        self.result_frame.hide()
        layout.addWidget(self.result_frame)

        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Segoe UI", 13))
        self.error_label.setStyleSheet("color: #d32f2f;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)
        layout.addStretch()

    def _on_file_dropped(self, paths):
        self._progress_timer.stop()
        self.result_frame.hide()
        self.error_label.hide()
        self.progress.hide()
        self._progress_value = 0.0
        self.input_path = paths[0]
        fname = os.path.basename(self.input_path)
        fsize = human_size(os.path.getsize(self.input_path))
        self.file_info.setText(f"Selected: {fname} ({fsize})")
        self.file_info.show()
        self.search_frame.show()
        self.redact_btn.setEnabled(True)
        self.search_input.setEnabled(True)

    def _start_redact(self):
        text = self.search_input.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "No search terms", "Please enter the text you want to redact.")
            return
        terms = [t.strip() for t in text.split("\n") if t.strip()]
        reply = QMessageBox.warning(self, "Confirm Redaction",
            f"This will permanently remove {len(terms)} term(s) from the PDF.\n\nThe original content cannot be recovered from the redacted file.\n\nContinue?",
            QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.No:
            return
        self.result_frame.hide()
        self.error_label.hide()
        self.progress.show()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self._progress_timer.start()
        self.redact_btn.setEnabled(False)
        self.drop_zone.setEnabled(False)
        self.worker = RedactWorker(self.input_path, terms)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _tick_progress(self):
        remaining = 90.0 - self._progress_value
        self._progress_value += remaining * 0.05
        self.progress.setValue(int(self._progress_value))

    def _on_finished(self, success, output_path, orig_size, new_size, redaction_count):
        self._progress_timer.stop()
        self.progress.setValue(100)
        self.drop_zone.setEnabled(True)
        self.redact_btn.setEnabled(True)
        self.progress.hide()
        if not success:
            self.error_label.setText("Something went wrong \u2014 please try again or contact your IT team.")
            self.error_label.show()
            return
        self.output_tmp_path = output_path
        self.result_icon.setText("\u2705")
        if redaction_count > 0:
            self.result_text.setText(
                f'<div style="text-align:center;">'
                f'Redacted {redaction_count} instance(s) across the document.<br>'
                f'<span style="color:#888;">Size:</span> <b style="color:#4CAF50;">{human_size(new_size)}</b>'
                f'</div>')
        else:
            self.result_text.setText("No matches found for the search terms.<br>The document was not modified.")
        self.name_input.setText(Path(self.input_path).stem + " - Redacted")
        self.save_btn.setEnabled(True)
        self.name_input.setEnabled(True)
        self.result_frame.show()

    def _save(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Name needed", "Please type a name for the file.")
            return
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        dest_dir = os.path.dirname(self.input_path)
        dest, _ = QFileDialog.getSaveFileName(self, "Save redacted PDF", os.path.join(dest_dir, name), "PDF Files (*.pdf)")
        if not dest:
            return
        try:
            shutil.copy2(self.output_tmp_path, dest)
            saved_name = os.path.basename(dest)
            self.result_icon.setText("\u2705")
            self.result_text.setText(
                f'<div style="text-align:center;">'
                f'<span style="color:#888;">Saved as:</span> <b>{saved_name}</b><br>'
                f'<span style="color:#888;">Size:</span> <b style="color:#4CAF50;">{human_size(os.path.getsize(dest))}</b>'
                f'</div>')
            self.save_btn.setEnabled(False)
            self.name_input.setEnabled(False)
        except Exception:
            self.error_label.setText("Something went wrong while saving \u2014 please try again or contact your IT team.")
            self.error_label.show()


# ---------------------------------------------------------------------------
# OCR tab
# ---------------------------------------------------------------------------

class OCRWorker(QThread):
    finished = pyqtSignal(bool, str, int, int)

    def __init__(self, input_path: str):
        super().__init__()
        self.input_path = input_path

    def run(self):
        orig_size = os.path.getsize(self.input_path)
        try:
            from PIL import Image
            import io
            tmp_dir = _tracked_mkdtemp()
            output_path = os.path.join(tmp_dir, "ocr_output.pdf")
            doc = fitz.open(self.input_path)
            out_doc = fitz.open()
            for page_num in range(len(doc)):
                page = doc[page_num]
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
                ocr_page_doc = fitz.open("pdf", pdf_bytes)
                out_doc.insert_pdf(ocr_page_doc)
                ocr_page_doc.close()
            out_doc.save(output_path, garbage=4, deflate=True)
            out_doc.close()
            doc.close()
            new_size = os.path.getsize(output_path)
            self.finished.emit(True, output_path, orig_size, new_size)
        except Exception:
            self.finished.emit(False, "", orig_size, 0)


class OCRTab(QWidget):
    def __init__(self):
        super().__init__()
        self.input_path = ""
        self.output_tmp_path = ""
        self.worker = None
        self._progress_timer = QTimer()
        self._progress_timer.setInterval(150)
        self._progress_timer.timeout.connect(self._tick_progress)
        self._progress_value = 0.0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        self.drop_zone = DropZone("Drop a scanned PDF here for OCR")
        self.drop_zone.files_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_zone)

        info = QLabel("OCR converts scanned documents into searchable,\nselectable text while keeping the original appearance.")
        info.setFont(QFont("Segoe UI", 11))
        info.setAlignment(Qt.AlignCenter)
        info.setStyleSheet("color: #888888;")
        info.setWordWrap(True)
        layout.addWidget(info)

        if not HAS_TESSERACT:
            warn = QLabel("\u26A0 Tesseract OCR engine was not found.\nPlease ask your IT team to install it.")
            warn.setFont(QFont("Segoe UI", 12))
            warn.setStyleSheet("color: #d32f2f; padding: 12px;")
            warn.setAlignment(Qt.AlignCenter)
            warn.setWordWrap(True)
            layout.addWidget(warn)
            self.drop_zone.setEnabled(False)

        self.file_info = QLabel("")
        self.file_info.setFont(QFont("Segoe UI", 12))
        self.file_info.setAlignment(Qt.AlignCenter)
        self.file_info.setWordWrap(True)
        self.file_info.hide()
        layout.addWidget(self.file_info)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.result_frame = QFrame()
        result_layout = QVBoxLayout(self.result_frame)
        result_layout.setSpacing(12)

        self.result_icon = QLabel()
        self.result_icon.setFont(QFont("Segoe UI", 48))
        self.result_icon.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(self.result_icon)

        self.result_text = QLabel("")
        self.result_text.setFont(QFont("Segoe UI", 13))
        self.result_text.setAlignment(Qt.AlignCenter)
        self.result_text.setWordWrap(True)
        result_layout.addWidget(self.result_text)

        save_label = QLabel("What would you like to call this file?")
        save_label.setFont(QFont("Segoe UI", 13))
        save_label.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(save_label)

        name_row = QHBoxLayout()
        name_row.setSpacing(8)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type a name for the file")
        self.name_input.setMinimumHeight(36)
        name_row.addWidget(self.name_input)
        pdf_label = QLabel(".pdf")
        pdf_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        pdf_label.setStyleSheet("color: #888;")
        name_row.addWidget(pdf_label)
        result_layout.addLayout(name_row)

        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet(BTN_SUCCESS)
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self._save)
        result_layout.addWidget(self.save_btn, alignment=Qt.AlignCenter)

        self.result_frame.hide()
        layout.addWidget(self.result_frame)

        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Segoe UI", 13))
        self.error_label.setStyleSheet("color: #d32f2f;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)
        layout.addStretch()

    def _on_file_dropped(self, paths):
        self._progress_timer.stop()
        self.result_frame.hide()
        self.error_label.hide()
        self.progress.hide()
        self._progress_value = 0.0
        self.input_path = paths[0]
        fname = os.path.basename(self.input_path)
        fsize = human_size(os.path.getsize(self.input_path))
        self.file_info.setText(f"Selected: {fname} ({fsize})")
        self.file_info.show()
        self.progress.show()
        self.progress.setValue(0)
        self._progress_timer.start()
        self.drop_zone.setEnabled(False)
        self.worker = OCRWorker(self.input_path)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _tick_progress(self):
        remaining = 90.0 - self._progress_value
        self._progress_value += remaining * 0.02
        self.progress.setValue(int(self._progress_value))

    def _on_finished(self, success, output_path, orig_size, new_size):
        self._progress_timer.stop()
        self.progress.setValue(100)
        self.drop_zone.setEnabled(True)
        self.progress.hide()
        if not success:
            self.error_label.setText("Something went wrong \u2014 please try again or contact your IT team.")
            self.error_label.show()
            return
        self.output_tmp_path = output_path
        self.result_icon.setText("\u2705")
        self.result_text.setText(
            f'<div style="text-align:center;">'
            f'OCR complete! Text is now searchable and selectable.<br>'
            f'<span style="color:#888;">Size:</span> <b style="color:#4CAF50;">{human_size(new_size)}</b>'
            f'</div>')
        self.name_input.setText(Path(self.input_path).stem + " - OCR")
        self.save_btn.setEnabled(True)
        self.name_input.setEnabled(True)
        self.result_frame.show()

    def _save(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Name needed", "Please type a name for the file.")
            return
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        dest_dir = os.path.dirname(self.input_path)
        dest, _ = QFileDialog.getSaveFileName(self, "Save OCR PDF", os.path.join(dest_dir, name), "PDF Files (*.pdf)")
        if not dest:
            return
        try:
            shutil.copy2(self.output_tmp_path, dest)
            saved_name = os.path.basename(dest)
            self.result_icon.setText("\u2705")
            self.result_text.setText(
                f'<div style="text-align:center;">'
                f'<span style="color:#888;">Saved as:</span> <b>{saved_name}</b><br>'
                f'<span style="color:#888;">Size:</span> <b style="color:#4CAF50;">{human_size(os.path.getsize(dest))}</b>'
                f'</div>')
            self.save_btn.setEnabled(False)
            self.name_input.setEnabled(False)
        except Exception:
            self.error_label.setText("Something went wrong while saving \u2014 please try again or contact your IT team.")
            self.error_label.show()


# ---------------------------------------------------------------------------
# Image compress worker + tab
# ---------------------------------------------------------------------------


class ImageCompressWorker(QThread):
    finished = pyqtSignal(bool, str, int, int, str)  # success, output_path, orig_size, new_size, format_used

    def __init__(self, input_path: str, quality: int = 80):
        super().__init__()
        self.input_path = input_path
        self.quality = quality

    def run(self):
        try:
            from PIL import Image
            orig_size = os.path.getsize(self.input_path)
            img = Image.open(self.input_path)

            # Convert HEIC/HEIF and other modes to RGB for JPEG output
            if img.mode in ("RGBA", "P", "LA"):
                # Keep PNG for images with transparency
                has_alpha = img.mode in ("RGBA", "LA") or \
                    (img.mode == "P" and "transparency" in img.info)
                if has_alpha:
                    img = img.convert("RGBA")
                else:
                    img = img.convert("RGB")
            elif img.mode != "RGB":
                img = img.convert("RGB")

            tmp_dir = _tracked_mkdtemp()
            ext = os.path.splitext(self.input_path)[1].lower()

            # For images with transparency, save as optimised PNG
            if img.mode == "RGBA":
                output_path = os.path.join(tmp_dir, "compressed.png")
                img.save(output_path, "PNG", optimize=True)
                fmt = "PNG"
            else:
                # Save as JPEG with quality setting
                output_path = os.path.join(tmp_dir, "compressed.jpg")
                img.save(output_path, "JPEG", quality=self.quality, optimize=True,
                         subsampling=1)  # 4:2:2 chroma subsampling
                fmt = "JPEG"

                # If original was PNG without transparency, also try optimised PNG
                # and keep whichever is smaller
                if ext == ".png":
                    png_path = os.path.join(tmp_dir, "compressed.png")
                    img.save(png_path, "PNG", optimize=True)
                    if os.path.getsize(png_path) < os.path.getsize(output_path):
                        output_path = png_path
                        fmt = "PNG"

            new_size = os.path.getsize(output_path)
            self.finished.emit(True, output_path, orig_size, new_size, fmt)
        except Exception:
            self.finished.emit(False, "", 0, 0, "")


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    BASE_WIDTH = 780  # Design width for scale factor = 1.0

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Tool")
        self.setMinimumSize(500, 400)
        self.resize(780, 650)

        self.gs_exe = find_ghostscript()

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(16, 16, 16, 16)

        header = QLabel("PDF Tool")
        header.setFont(QFont("Segoe UI", 22, QFont.Bold))
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("color: #1976D2; margin-bottom: 4px;")
        main_layout.addWidget(header)

        subtitle = QLabel("Compress, merge, rename, redact, flatten, and OCR your PDFs and images \u2014 everything stays on this computer")
        subtitle.setFont(QFont("Segoe UI", 11))
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #888888; margin-bottom: 12px;")
        subtitle.setWordWrap(True)
        main_layout.addWidget(subtitle)

        tabs = QTabWidget()
        all_tabs = [
            (CompressTab(self.gs_exe), "Compress"),
            (MergeTab(self.gs_exe), "Merge"),
            (RenameTab(), "Rename"),
            (RedactTab(), "Redact"),
            (FlattenTab(), "Flatten"),
            (OCRTab(), "OCR"),
        ]
        for widget, label in all_tabs:
            if ENABLED_TABS is not None and label not in ENABLED_TABS:
                continue
            scroll = QScrollArea()
            scroll.setWidget(widget)
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
            tabs.addTab(scroll, label)
        main_layout.addWidget(tabs)

        if not self.gs_exe:
            warn = QLabel(
                "\u26A0 Ghostscript was not found on this computer.\n"
                "Compress and Merge features will not work.\n"
                "Please ask your IT team to install Ghostscript.")
            warn.setFont(QFont("Segoe UI", 12))
            warn.setStyleSheet("color: #d32f2f; padding: 12px;")
            warn.setAlignment(Qt.AlignCenter)
            warn.setWordWrap(True)
            main_layout.addWidget(warn)

        # Store references for responsive scaling
        self._header = header
        self._subtitle = subtitle
        self._tabs = tabs

        # Collect all scalable widgets from tabs
        self._font_map = []  # list of (widget, base_size, bold)
        self._font_map.append((header, 22, True))
        self._font_map.append((subtitle, 11, False))

        for i in range(tabs.count()):
            tab = tabs.widget(i)
            self._collect_scalable_widgets(tab)

        # Apply initial scale
        self._apply_scale()

        # Update banner placeholder (inserted above tabs when update is ready)
        self._update_banner = None
        self._main_layout = main_layout

        # Check for updates in background
        self._update_checker = UpdateChecker()
        self._update_checker.update_available.connect(self._on_update_available)
        self._update_checker.start()

    def _on_update_available(self, latest_version, download_url, sig_url):
        """Auto-download update and show an in-app banner."""
        if self._update_banner is not None:
            return  # already showing
        self._update_banner = UpdateBanner(self, latest_version, download_url, sig_url)
        # Insert banner after subtitle (index 2: header=0, subtitle=1)
        self._main_layout.insertWidget(2, self._update_banner)

    def _collect_scalable_widgets(self, widget):
        """Walk the widget tree and record every widget's base font size."""
        for child in widget.findChildren(QWidget):
            font = child.font()
            size = font.pointSize()
            if size <= 0:
                size = font.pixelSize()
            if size > 0:
                bold = font.weight() >= QFont.Bold
                self._font_map.append((child, size, bold))

    BASE_HEIGHT = 650  # Design height for scale factor = 1.0

    def _get_scale(self):
        w_scale = self.width() / self.BASE_WIDTH
        h_scale = self.height() / self.BASE_HEIGHT
        scale = min(w_scale, h_scale)  # use the more constraining axis
        return max(0.55, min(scale, 1.5))

    def _apply_scale(self):
        scale = self._get_scale()
        for widget, base_size, bold in self._font_map:
            try:
                new_size = max(8, int(base_size * scale))
                weight = QFont.Bold if bold else QFont.Normal
                widget.setFont(QFont("Segoe UI", new_size, weight))
            except RuntimeError:
                pass  # widget may have been deleted

        # Scale layout spacing and padding within tabs
        spacing = max(6, int(12 * scale))
        tab_margin = max(10, int(24 * scale))
        for i in range(self._tabs.count()):
            scroll = self._tabs.widget(i)
            tab = scroll.widget() if isinstance(scroll, QScrollArea) else scroll
            if tab and tab.layout():
                tab.layout().setSpacing(max(6, int(16 * scale)))
                tab.layout().setContentsMargins(tab_margin, tab_margin, tab_margin, tab_margin)
            for frame in tab.findChildren(QFrame):
                if frame.layout():
                    frame.layout().setSpacing(spacing)

        # Scale QLineEdit padding
        le_pad_v = max(4, int(8 * scale))
        le_pad_h = max(6, int(12 * scale))
        le_font = max(10, int(13 * scale))
        for i in range(self._tabs.count()):
            tab = self._tabs.widget(i)
            for le in tab.findChildren(QLineEdit):
                try:
                    le.setStyleSheet(f"""
                        QLineEdit {{
                            font-size: {le_font}px;
                            padding: {le_pad_v}px {le_pad_h}px;
                            border: 2px solid #e0e0e0;
                            border-radius: 8px;
                            background: white;
                        }}
                        QLineEdit:focus {{
                            border-color: #1976D2;
                        }}
                    """)
                except RuntimeError:
                    pass

        # Scale tab bar font via stylesheet
        tab_size = max(10, int(13 * scale))
        tab_padding_h = max(10, int(16 * scale))
        tab_padding_v = max(6, int(10 * scale))
        tab_min_w = max(50, int(70 * scale))
        self._tabs.tabBar().setStyleSheet(f"""
            QTabBar::tab {{
                padding: {tab_padding_v}px {tab_padding_h}px;
                font-size: {tab_size}px;
                font-weight: 500;
                border: none;
                border-bottom: 3px solid transparent;
                color: #666666;
                background: transparent;
                min-width: {tab_min_w}px;
            }}
            QTabBar::tab:selected {{
                color: #1976D2;
                border-bottom: 3px solid #1976D2;
            }}
            QTabBar::tab:hover {{
                color: #333333;
            }}
        """)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_scale()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(GLOBAL_STYLE)

    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(255, 255, 255))
    palette.setColor(QPalette.WindowText, QColor(33, 33, 33))
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
    palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ToolTipText, QColor(33, 33, 33))
    palette.setColor(QPalette.Text, QColor(33, 33, 33))
    palette.setColor(QPalette.Button, QColor(245, 245, 245))
    palette.setColor(QPalette.ButtonText, QColor(33, 33, 33))
    palette.setColor(QPalette.Highlight, QColor(25, 118, 210))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
