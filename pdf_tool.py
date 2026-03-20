"""
PDF Tool — Compress, Merge, and Rename PDFs locally.
Built for non-technical users in legal/admin environments.
No internet, no cloud, no third-party services. Everything stays on this machine.
"""

APP_VERSION = "1.0.0"
GITHUB_REPO = "hugodrummon/pdf-tool"

import sys
import os
import subprocess
import shutil
import tempfile
import json
import webbrowser
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.error import URLError

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QProgressBar, QTabWidget,
    QListWidget, QListWidgetItem, QFileDialog, QMessageBox,
    QSizePolicy, QFrame, QSpacerItem, QAbstractItemView, QDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMimeData, QSize, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon, QDragEnterEvent, QDropEvent

from PyPDF2 import PdfMerger


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
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={quality}",
        "-dNOPAUSE",
        "-dBATCH",
        "-dQUIET",
        f"-dNumRenderingThreads={num_threads}",
        "-dDetectDuplicateImages=true",
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        "-dColorImageDownsampleType=/Bicubic",
        "-dGrayImageDownsampleType=/Bicubic",
        f"-sOutputFile={output_path}",
        input_path,
    ]

    startupinfo = None
    creationflags = 0
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        creationflags = subprocess.CREATE_NO_WINDOW

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

        tmp_dir = tempfile.mkdtemp()
        output_path = os.path.join(tmp_dir, "compressed.pdf")

        # Pick quality based on file size — skip gentle /ebook for large files
        # since it's slow and often not aggressive enough anyway
        if orig_size > 50 * 1024 * 1024:  # > 50 MB: go straight to aggressive
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
        tmp_dir = tempfile.mkdtemp()
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

        # Pick quality based on merged size
        if merged_size > 50 * 1024 * 1024:
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
    update_available = pyqtSignal(str, str)  # (latest_version, download_url)

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
                # Find the installer asset download URL, fall back to release page
                download_url = data.get("html_url", "")
                for asset in data.get("assets", []):
                    if "install" in asset["name"].lower() and asset["name"].endswith(".exe"):
                        download_url = asset["browser_download_url"]
                        break
                self.update_available.emit(latest, download_url)
        except Exception:
            pass  # Silent fail — no internet, no problem


class UpdateDialog(QDialog):
    """Dialog shown when a new version is available."""
    def __init__(self, parent, latest_version, download_url):
        super().__init__(parent)
        self.download_url = download_url
        self.setWindowTitle("Update Available")
        self.setFixedSize(420, 200)

        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("A new version is available!")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        info = QLabel(
            f"Your version: {APP_VERSION}\n"
            f"Latest version: {latest_version}")
        info.setFont(QFont("Segoe UI", 11))
        info.setAlignment(Qt.AlignCenter)
        layout.addWidget(info)

        btn_row = QHBoxLayout()

        download_btn = QPushButton("Download Update")
        download_btn.setFont(QFont("Segoe UI", 12))
        download_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50; color: white;
                border: none; padding: 10px 24px; border-radius: 8px;
            }
            QPushButton:hover { background-color: #43A047; }
        """)
        download_btn.setCursor(Qt.PointingHandCursor)
        download_btn.clicked.connect(self._download)
        btn_row.addWidget(download_btn)

        later_btn = QPushButton("Later")
        later_btn.setFont(QFont("Segoe UI", 12))
        later_btn.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5; color: #333;
                border: 1px solid #e0e0e0; padding: 10px 24px; border-radius: 8px;
            }
            QPushButton:hover { background-color: #eee; }
        """)
        later_btn.setCursor(Qt.PointingHandCursor)
        later_btn.clicked.connect(self.close)
        btn_row.addWidget(later_btn)

        layout.addLayout(btn_row)

    def _download(self):
        webbrowser.open(self.download_url)
        self.close()


# ---------------------------------------------------------------------------
# Custom drop zone widget
# ---------------------------------------------------------------------------

class DropZone(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self, label_text: str, accept_multiple: bool = False):
        super().__init__()
        self.accept_multiple = accept_multiple
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
        if self.accept_multiple:
            paths, _ = QFileDialog.getOpenFileNames(
                self, "Select PDF files", "", "PDF Files (*.pdf)")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select a PDF file", "", "PDF Files (*.pdf)")
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
            if p.lower().endswith(".pdf"):
                paths.append(p)
        if paths:
            if not self.accept_multiple:
                paths = paths[:1]
            self.files_dropped.emit(paths)
        else:
            QMessageBox.warning(self, "Wrong file type",
                                "Please drop a PDF file.")


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
        padding: 12px 28px;
        font-size: 14px;
        font-weight: 500;
        border: none;
        border-bottom: 3px solid transparent;
        color: #666666;
        background: transparent;
        min-width: 140px;
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
        font-size: 14px;
        padding: 10px 14px;
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

class CompressTab(QWidget):
    def __init__(self, gs_exe: str):
        super().__init__()
        self.gs_exe = gs_exe
        self.input_path = ""
        self.output_tmp_path = ""
        self.worker = None

        # Smooth progress animation
        self._progress_timer = QTimer()
        self._progress_timer.setInterval(150)
        self._progress_timer.timeout.connect(self._tick_progress)
        self._progress_value = 0.0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        self.drop_zone = DropZone("Drop your PDF here to compress it")
        self.drop_zone.files_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_zone)

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

        self.size_warning = QLabel("")
        self.size_warning.setFont(QFont("Segoe UI", 12))
        self.size_warning.setAlignment(Qt.AlignCenter)
        self.size_warning.setStyleSheet("color: #e65100; font-weight: bold;")
        self.size_warning.hide()
        result_layout.addWidget(self.size_warning)

        save_label = QLabel("What would you like to call this file?")
        save_label.setFont(QFont("Segoe UI", 13))
        save_label.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(save_label)

        name_row = QHBoxLayout()
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type a name for the file")
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

    def _reset(self):
        self._progress_timer.stop()
        self.result_frame.hide()
        self.error_label.hide()
        self.progress.hide()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self.size_warning.hide()
        self.save_btn.setEnabled(True)
        self.name_input.setEnabled(True)

    def _on_file_dropped(self, paths):
        self._reset()
        self.input_path = paths[0]
        fname = os.path.basename(self.input_path)
        fsize = human_size(os.path.getsize(self.input_path))
        self.file_info.setText(f"Selected: {fname} ({fsize})")
        self.file_info.show()

        if os.path.getsize(self.input_path) <= TARGET_SIZE_BYTES:
            self.result_icon.setText("\u2705")
            self.result_text.setText(
                f"This file is already under {TARGET_SIZE_MB} MB!\nSize: {fsize}")
            stem = Path(self.input_path).stem
            self.name_input.setText(stem)
            self.output_tmp_path = self.input_path
            self.result_frame.show()
            return

        self.progress.show()
        self.progress.setValue(0)
        self._progress_value = 0.0
        self._progress_timer.start()
        self.drop_zone.setEnabled(False)

        self.worker = CompressWorker(self.input_path, self.gs_exe)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _tick_progress(self):
        """Smoothly advance progress bar, slowing as it approaches 90%."""
        remaining = 90.0 - self._progress_value
        self._progress_value += remaining * 0.03
        self.progress.setValue(int(self._progress_value))

    def _on_finished(self, success, output_path, orig_size, new_size):
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
        self.result_icon.setText("\u2705")
        self.result_text.setText(
            f"Original size: {human_size(orig_size)}\nCompressed size: {human_size(new_size)}")

        if new_size > TARGET_SIZE_BYTES:
            self.size_warning.setText(
                f"Note: The compressed file is {human_size(new_size)}, still over {TARGET_SIZE_MB} MB.\n"
                "The original PDF may contain high-resolution scans.")
            self.size_warning.show()

        self.name_input.setText(Path(self.input_path).stem + " - Compressed")
        self.result_frame.show()

    def _save(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Name needed", "Please type a name for the file.")
            return
        if not name.lower().endswith(".pdf"):
            name += ".pdf"

        # Open a Save As dialog so the user picks where to save
        dest_dir = os.path.dirname(self.input_path)
        dest, _ = QFileDialog.getSaveFileName(
            self, "Save compressed PDF", os.path.join(dest_dir, name),
            "PDF Files (*.pdf)")
        if not dest:
            return

        try:
            shutil.copy2(self.output_tmp_path, dest)
            saved_name = os.path.basename(dest)
            self.result_icon.setText("\u2705")
            self.result_text.setText(f"Saved as: {saved_name}\nSize: {human_size(os.path.getsize(dest))}")
            self.save_btn.setEnabled(False)
            self.name_input.setEnabled(False)
        except Exception:
            self.error_label.setText(
                "Something went wrong while saving \u2014 please try again or contact your IT team.")
            self.error_label.show()


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
        self.size_warning.hide()
        result_layout.addWidget(self.size_warning)

        save_label = QLabel("What would you like to call this file?")
        save_label.setFont(QFont("Segoe UI", 13))
        save_label.setAlignment(Qt.AlignCenter)
        result_layout.addWidget(save_label)

        name_row = QHBoxLayout()
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type a name for the file")
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
        """Smoothly advance progress bar, slowing as it approaches 90%."""
        remaining = 90.0 - self._progress_value
        self._progress_value += remaining * 0.03
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
            f"Combined original size: {human_size(combined_size)}\nFinal size: {human_size(final_size)}")

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
            self.result_icon.setText("\u2705")
            self.result_text.setText(f"Saved as: {saved_name}\nSize: {human_size(os.path.getsize(dest))}")
            self.save_btn.setEnabled(False)
            self.name_input.setEnabled(False)
        except Exception:
            self.error_label.setText(
                "Something went wrong while saving \u2014 please try again or contact your IT team.")
            self.error_label.show()


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
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Type the new name")
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

        subtitle = QLabel("Compress, merge, and rename your PDF files \u2014 everything stays on this computer")
        subtitle.setFont(QFont("Segoe UI", 11))
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #888888; margin-bottom: 12px;")
        subtitle.setWordWrap(True)
        main_layout.addWidget(subtitle)

        tabs = QTabWidget()
        tabs.addTab(CompressTab(self.gs_exe), "  Compress  ")
        tabs.addTab(MergeTab(self.gs_exe), "  Merge  ")
        tabs.addTab(RenameTab(), "  Rename  ")
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

        # Check for updates in background
        self._update_checker = UpdateChecker()
        self._update_checker.update_available.connect(self._show_update_dialog)
        self._update_checker.start()

    def _show_update_dialog(self, latest_version, download_url):
        dialog = UpdateDialog(self, latest_version, download_url)
        dialog.exec_()

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

    def _get_scale(self):
        width = self.width()
        scale = width / self.BASE_WIDTH
        return max(0.65, min(scale, 1.5))

    def _apply_scale(self):
        scale = self._get_scale()
        for widget, base_size, bold in self._font_map:
            try:
                new_size = max(8, int(base_size * scale))
                weight = QFont.Bold if bold else QFont.Normal
                widget.setFont(QFont("Segoe UI", new_size, weight))
            except RuntimeError:
                pass  # widget may have been deleted

        # Scale tab bar font via stylesheet
        tab_size = max(10, int(14 * scale))
        tab_padding_h = max(12, int(28 * scale))
        tab_padding_v = max(8, int(12 * scale))
        tab_min_w = max(80, int(140 * scale))
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
