"""
PDF Tool — Compress, Merge, Rename, Redact, Flatten, and OCR PDFs locally.
Built for non-technical users in legal/admin environments.
No internet, no cloud, no third-party services. Everything stays on this machine.
"""

APP_VERSION = "2.1.3"
GITHUB_REPO = "hugodrummon/pdf-tool"
UPDATE_PUBLIC_KEY = "sw613yM42XKzroyOPRE19tMKJEqHQf2Ycne7S1rOMpU="
import sys

# Prevent DLL hijacking — remove current directory from DLL search order
if sys.platform == "win32":
    import ctypes
    ctypes.windll.kernel32.SetDllDirectoryW("")

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
            # Skip if the path is a symlink/junction — prevents symlink attack
            if os.path.islink(d) or (sys.platform == "win32" and os.path.isdir(d) and os.stat(d).st_file_attributes & 0x400):
                continue
            for root, dirs, files in os.walk(d, followlinks=False):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    if not os.path.islink(fpath):
                        _secure_delete_file(fpath)
            shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass

atexit.register(_cleanup_temp_dirs)

def _cleanup_orphaned_temp():
    """Clean up temp dirs from previous crashed sessions."""
    try:
        temp_root = tempfile.gettempdir()
        for name in os.listdir(temp_root):
            d = os.path.join(temp_root, name)
            if not name.startswith("tmp") or not os.path.isdir(d):
                continue
            if os.path.islink(d):
                continue
            # Check if it contains our signature files
            has_our_files = any(
                f in ("compressed.pdf", "merged.pdf")
                for f in os.listdir(d) if os.path.isfile(os.path.join(d, f))
            )
            if has_our_files:
                for fname in os.listdir(d):
                    fpath = os.path.join(d, fname)
                    if os.path.isfile(fpath) and not os.path.islink(fpath):
                        _secure_delete_file(fpath)
                shutil.rmtree(d, ignore_errors=True)
    except OSError:
        pass

import json
import webbrowser
from pathlib import Path
import hashlib
import base64
from urllib.request import urlopen, Request
from urllib.error import URLError

try:
    from nacl.signing import VerifyKey
    from nacl.exceptions import BadSignatureError
    HAS_NACL = True
except ImportError:
    HAS_NACL = False

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QProgressBar, QTabWidget,
    QListWidget, QListWidgetItem, QFileDialog, QMessageBox,
    QSizePolicy, QFrame, QSpacerItem, QAbstractItemView, QDialog,
    QTextEdit, QScrollArea, QSlider, QSpinBox, QStackedWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMimeData, QSize, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon, QDragEnterEvent, QDropEvent, QPixmap, QImage
from PyQt5.QtWidgets import QGraphicsDropShadowEffect

from PyPDF2 import PdfMerger, PdfReader, PdfWriter
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


def compress_pdf_aggressive(input_path: str, output_path: str, gs_exe: str,
                            dpi: int = 72, qfactor: float = 2.4,
                            grayscale: bool = False) -> bool:
    """Aggressive compression — forces image downsampling and recompression at given DPI/quality."""
    env = os.environ.copy()
    bundle_dir = get_bundle_dir()
    gs_lib_path = os.path.join(bundle_dir, "gs", "lib")
    gs_resource_path = os.path.join(bundle_dir, "gs", "Resource")
    if os.path.isdir(gs_lib_path):
        env["GS_LIB"] = f"{gs_lib_path};{gs_resource_path}"

    num_threads = min(os.cpu_count() or 2, 8)

    args = [
        gs_exe,
        "-dSAFER",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/screen",
        "-dNOPAUSE",
        "-dBATCH",
        "-dQUIET",
        f"-dNumRenderingThreads={num_threads}",
        "-dDetectDuplicateImages=true",
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        "-dDownsampleColorImages=true",
        "-dDownsampleGrayImages=true",
        "-dDownsampleMonoImages=true",
        f"-dColorImageResolution={dpi}",
        f"-dGrayImageResolution={dpi}",
        f"-dMonoImageResolution={dpi}",
        "-dColorImageDownsampleType=/Bicubic",
        "-dGrayImageDownsampleType=/Bicubic",
        "-dMonoImageDownsampleType=/Subsample",
        "-dPassThroughJPEGImages=false",
        "-dAutoFilterColorImages=false",
        "-dAutoFilterGrayImages=false",
        "-dColorImageFilter=/DCTEncode",
        "-dGrayImageFilter=/DCTEncode",
    ]
    if grayscale:
        args += ["-dColorConversionStrategy=/Gray", "-dProcessColorModel=/DeviceGray"]
    else:
        args += ["-dColorConversionStrategy=/LeaveColorUnchanged"]
    args += [
        f"-sOutputFile={output_path}",
        "-c",
        f"<< /ColorACSImageDict << /QFactor {qfactor} /Blend 1 /HSamples [2 1 1 2] /VSamples [2 1 1 2] >> /GrayACSImageDict << /QFactor {qfactor} /Blend 1 /HSamples [2 1 1 2] /VSamples [2 1 1 2] >> >> setdistillerparams",
        "-f",
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
        args, startupinfo=startupinfo, capture_output=True,
        creationflags=creationflags, env=env,
    )
    return result.returncode == 0


def rasterize_pdf(input_path: str, output_path: str, target_bytes: int) -> bool:
    """Last resort: render each page as a grayscale JPEG image and rebuild the PDF.
    Progressively lowers quality until under target size. Destroys editability.
    Minimum 72 DPI / 50% JPEG to keep text readable."""
    try:
        doc = fitz.open(input_path)

        for dpi, jpeg_quality in [(150, 70), (100, 60), (72, 50)]:
            out_doc = fitz.open()
            for page in doc:
                pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
                img_data = pix.tobytes(output="jpeg", jpg_quality=jpeg_quality)
                rect = page.rect
                out_page = out_doc.new_page(width=rect.width, height=rect.height)
                out_page.insert_image(rect, stream=img_data)
            out_doc.save(output_path, deflate=True, garbage=4)
            out_doc.close()
            if os.path.getsize(output_path) <= target_bytes:
                doc.close()
                return True
        doc.close()
        # Only return the file if rasterizing actually reduced size vs input
        if os.path.isfile(output_path):
            return os.path.getsize(output_path) < os.path.getsize(input_path)
        return False
    except Exception:
        return False


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
            # If still over target, progressively lower DPI until under 10 MB
            if new_size > TARGET_SIZE_BYTES:
                best_size = new_size
                best_path = output_path
                tmp_dir2 = _tracked_mkdtemp()
                trial_path = os.path.join(tmp_dir2, "trial.pdf")
                for dpi, qf, gray in [
                    (72, 2.4, False), (50, 2.4, False), (36, 3.0, False),
                    (72, 2.4, True), (50, 2.4, True), (36, 3.0, True), (24, 4.0, True),
                ]:
                    ok3 = compress_pdf_aggressive(self.input_path, trial_path, self.gs_exe, dpi=dpi, qfactor=qf, grayscale=gray)
                    if ok3 and os.path.isfile(trial_path):
                        trial_size = os.path.getsize(trial_path)
                        if trial_size < best_size:
                            best_size = trial_size
                            shutil.copy2(trial_path, output_path)
                        if best_size <= TARGET_SIZE_BYTES:
                            break
                new_size = best_size
            # Last resort: rasterize pages as JPEG images
            if new_size > TARGET_SIZE_BYTES:
                raster_path = os.path.join(tmp_dir, "rasterized.pdf")
                if rasterize_pdf(self.input_path, raster_path, TARGET_SIZE_BYTES):
                    raster_size = os.path.getsize(raster_path)
                    if raster_size < new_size:
                        shutil.copy2(raster_path, output_path)
                        new_size = raster_size
            # Re-compress output — a second pass often shrinks further
            if new_size > TARGET_SIZE_BYTES:
                for _pass in range(2):
                    repass_path = os.path.join(tmp_dir, f"repass{_pass}.pdf")
                    ok_r = compress_pdf(output_path, repass_path, self.gs_exe, "/screen")
                    if ok_r and os.path.isfile(repass_path):
                        repass_size = os.path.getsize(repass_path)
                        if repass_size < new_size:
                            shutil.copy2(repass_path, output_path)
                            new_size = repass_size
                        if new_size <= TARGET_SIZE_BYTES:
                            break
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
            # If still over target, progressively lower DPI until under 10 MB
            if new_size > TARGET_SIZE_BYTES:
                best_size = new_size
                tmp_dir2 = _tracked_mkdtemp()
                trial_path = os.path.join(tmp_dir2, "trial.pdf")
                for dpi, qf, gray in [
                    (72, 2.4, False), (50, 2.4, False), (36, 3.0, False),
                    (72, 2.4, True), (50, 2.4, True), (36, 3.0, True), (24, 4.0, True),
                ]:
                    ok3 = compress_pdf_aggressive(merged_path, trial_path, self.gs_exe, dpi=dpi, qfactor=qf, grayscale=gray)
                    if ok3 and os.path.isfile(trial_path):
                        trial_size = os.path.getsize(trial_path)
                        if trial_size < best_size:
                            best_size = trial_size
                            shutil.copy2(trial_path, compressed_path)
                        if best_size <= TARGET_SIZE_BYTES:
                            break
                new_size = best_size
            # Last resort: rasterize pages as JPEG images
            if new_size > TARGET_SIZE_BYTES:
                raster_path = os.path.join(tmp_dir, "rasterized.pdf")
                if rasterize_pdf(merged_path, raster_path, TARGET_SIZE_BYTES):
                    raster_size = os.path.getsize(raster_path)
                    if raster_size < new_size:
                        shutil.copy2(raster_path, compressed_path)
                        new_size = raster_size
            # Re-compress output — a second pass often shrinks further
            if new_size > TARGET_SIZE_BYTES:
                for _pass in range(2):
                    repass_path = os.path.join(tmp_dir, f"repass{_pass}.pdf")
                    ok_r = compress_pdf(compressed_path, repass_path, self.gs_exe, "/screen")
                    if ok_r and os.path.isfile(repass_path):
                        repass_size = os.path.getsize(repass_path)
                        if repass_size < new_size:
                            shutil.copy2(repass_path, compressed_path)
                            new_size = repass_size
                        if new_size <= TARGET_SIZE_BYTES:
                            break
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
                    if sig_url and not sig_url.startswith("https://github.com/"):
                        sig_url = ""
                    self.update_available.emit(latest, download_url, sig_url)
        except (URLError, ValueError, KeyError, OSError):
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
            # If nacl is available, REQUIRE a valid signature — never fail-open
            if HAS_NACL and not self.sig_url:
                self.finished.emit(False, "")
                return
            if self.sig_url and HAS_NACL:
                try:
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
                except (BadSignatureError, URLError, ValueError, OSError):
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
            "UpdateBanner { background-color: #252527; border: 1px solid #3a3a3c; "
            "border-radius: 6px; }")
        banner_layout = QHBoxLayout(self)
        banner_layout.setContentsMargins(16, 8, 16, 8)
        banner_layout.setSpacing(12)

        self.status_label = QLabel("Downloading update...")
        self.status_label.setFont(QFont("Segoe UI", 11))
        self.status_label.setStyleSheet("color: #3b82f6; border: none; background: transparent;")
        banner_layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setMaximumWidth(150)
        self.progress.setMaximumHeight(14)
        banner_layout.addWidget(self.progress)

        self.restart_btn = QPushButton("Install and close")
        self.restart_btn.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.restart_btn.setStyleSheet(
            "QPushButton { background-color: #22c55e; color: #1c1c1e; border: none; "
            "padding: 6px 16px; border-radius: 6px; }"
            "QPushButton:hover { background-color: #16a34a; }")
        self.restart_btn.setCursor(Qt.PointingHandCursor)
        self.restart_btn.clicked.connect(self._do_restart_update)
        self.restart_btn.hide()
        banner_layout.addWidget(self.restart_btn)

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
                "color: #ef4444; border: none; background: transparent;")
            self.progress.hide()
            return

        self.installer_path = installer_path
        self.progress.hide()
        self.status_label.setText(f"v{self.latest_version} ready!")
        self.restart_btn.show()

    def _do_restart_update(self):
        self.restart_btn.setEnabled(False)
        self.status_label.setText("Closing and installing update...")

        app_exe = sys.executable
        # Write a temp batch script to run installer then relaunch — avoids shell=True
        bat_path = os.path.join(tempfile.gettempdir(), "_pdf_tool_update.bat")
        with open(bat_path, "w") as bat:
            bat.write("@echo off\r\n")
            bat.write(f'"{self.installer_path}" /SILENT /CLOSEAPPLICATIONS /FORCECLOSEAPPLICATIONS\r\n')
            bat.write(f'start "" "{app_exe}"\r\n')
            bat.write('del "%~f0"\r\n')  # self-delete
        subprocess.Popen(
            ["cmd", "/c", bat_path],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
        )

        QApplication.instance().quit()


# ---------------------------------------------------------------------------
# Custom drop zone widget
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Additional workers (kept from original)
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


class ImageCompressWorker(QThread):
    finished = pyqtSignal(bool, str, int, int, str)

    def __init__(self, input_path: str, quality: int = 80):
        super().__init__()
        self.input_path = input_path
        self.quality = quality

    def run(self):
        try:
            from PIL import Image
            orig_size = os.path.getsize(self.input_path)
            img = Image.open(self.input_path)

            if img.mode in ("RGBA", "P", "LA"):
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

            if img.mode == "RGBA":
                output_path = os.path.join(tmp_dir, "compressed.png")
                img.save(output_path, "PNG", optimize=True)
                fmt = "PNG"
            else:
                output_path = os.path.join(tmp_dir, "compressed.jpg")
                img.save(output_path, "JPEG", quality=self.quality, optimize=True,
                         subsampling=1)
                fmt = "JPEG"

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


IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif"]


# ---------------------------------------------------------------------------
# Dark theme
# ---------------------------------------------------------------------------

DARK_STYLE = """
QWidget {
    background-color: #1c1c1e;
    color: #f2f2f7;
    font-family: "Segoe UI", sans-serif;
    font-size: 13px;
}
QMainWindow { background-color: #1c1c1e; }
QScrollArea { background-color: transparent; border: none; }
QScrollArea > QWidget > QWidget { background-color: transparent; }
QScrollBar:vertical {
    background: #252527; width: 8px; border: none; border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #48484a; min-height: 24px; border-radius: 4px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background: #252527; height: 8px; border: none; border-radius: 4px;
}
QScrollBar::handle:horizontal {
    background: #48484a; min-width: 24px; border-radius: 4px;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QProgressBar {
    background: #3a3a3c; border: none; border-radius: 4px;
    height: 6px; text-align: center; color: transparent;
}
QProgressBar::chunk { background: #3b82f6; border-radius: 4px; }
QLineEdit {
    background: #2c2c2e; border: 1px solid #3a3a3c; border-radius: 5px;
    padding: 7px 9px; color: #f2f2f7; font-size: 12px;
}
QLineEdit:focus { border-color: #3b82f6; }
QTextEdit {
    background: #2c2c2e; border: 1px solid #3a3a3c; border-radius: 5px;
    padding: 6px; color: #f2f2f7; font-size: 12px;
}
QListWidget {
    background: #2c2c2e; border: 1px solid #3a3a3c; border-radius: 5px;
    color: #f2f2f7; font-size: 12px; padding: 2px;
}
QListWidget::item { padding: 4px 6px; border-radius: 3px; }
QListWidget::item:selected { background: rgba(59,130,246,0.25); color: #f2f2f7; }
QSpinBox {
    background: #2c2c2e; border: 1px solid #3a3a3c; border-radius: 5px;
    padding: 4px 8px; color: #f2f2f7;
}
QSpinBox:focus { border-color: #3b82f6; }
QSlider::groove:horizontal {
    background: #3a3a3c; height: 4px; border-radius: 2px;
}
QSlider::handle:horizontal {
    background: #3b82f6; width: 14px; height: 14px; margin: -5px 0;
    border-radius: 7px;
}
QToolTip {
    background: #252527; color: #f2f2f7; border: 1px solid #3a3a3c;
    padding: 4px 8px; border-radius: 4px;
}
QMessageBox { background-color: #252527; }
QMessageBox QLabel { color: #f2f2f7; }
QMessageBox QPushButton {
    background: #3a3a3c; color: #f2f2f7; border: none;
    padding: 6px 16px; border-radius: 5px;
}
QMessageBox QPushButton:hover { background: #48484a; }
"""

BTN_PRIMARY = (
    "QPushButton { background-color: #3b82f6; color: #ffffff; border: none; "
    "padding: 9px 0; border-radius: 6px; font-size: 13px; font-weight: 500; }"
    "QPushButton:hover { background-color: #2563eb; }"
    "QPushButton:pressed { background-color: #1d4ed8; }"
    "QPushButton:disabled { background-color: #3a3a3c; color: #636366; }")

BTN_SUCCESS = (
    "QPushButton { background-color: #22c55e; color: #1c1c1e; border: none; "
    "padding: 9px 0; border-radius: 6px; font-size: 13px; font-weight: 500; }"
    "QPushButton:hover { background-color: #16a34a; }"
    "QPushButton:pressed { background-color: #15803d; }"
    "QPushButton:disabled { background-color: #3a3a3c; color: #636366; }")

BTN_DANGER = (
    "QPushButton { background-color: #ef4444; color: #ffffff; border: none; "
    "padding: 9px 0; border-radius: 6px; font-size: 13px; font-weight: 500; }"
    "QPushButton:hover { background-color: #dc2626; }"
    "QPushButton:pressed { background-color: #b91c1c; }"
    "QPushButton:disabled { background-color: #3a3a3c; color: #636366; }")

BTN_SECONDARY = (
    "QPushButton { background-color: #3a3a3c; color: #f2f2f7; border: none; "
    "padding: 7px 0; border-radius: 5px; font-size: 12px; }"
    "QPushButton:hover { background-color: #48484a; }"
    "QPushButton:pressed { background-color: #505052; }")


# ---------------------------------------------------------------------------
# Compact drop zone for right panel
# ---------------------------------------------------------------------------

class CompactDropZone(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self, label_text="Drop PDF here or browse",
                 accept_multiple=False, file_extensions=None, file_filter_name="PDF"):
        super().__init__()
        self._accept_multiple = accept_multiple
        self._extensions = [e.lower() for e in (file_extensions or [".pdf"])]
        self._filter_name = file_filter_name
        self.setAcceptDrops(True)
        self.setFixedHeight(64)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(
            "CompactDropZone { background: #252527; border: 1.5px dashed #3a3a3c; "
            "border-radius: 8px; }"
            "CompactDropZone:hover { border-color: #3b82f6; background: #2a2a2c; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(2)
        layout.setAlignment(Qt.AlignCenter)
        self._label = QLabel(label_text)
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setStyleSheet("color: #aeaeb2; font-size: 11px; border: none; background: transparent;")
        layout.addWidget(self._label)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._browse()

    def _browse(self):
        ext_filter = " ".join(f"*{e}" for e in self._extensions)
        title = "Select files" if self._accept_multiple else "Select file"
        if self._accept_multiple:
            paths, _ = QFileDialog.getOpenFileNames(
                self, title, "", f"{self._filter_name} Files ({ext_filter})")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, title, "", f"{self._filter_name} Files ({ext_filter})")
            paths = [path] if path else []
        if paths:
            self.files_dropped.emit(paths)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(
                "CompactDropZone { background: #2a2a2c; border: 1.5px dashed #3b82f6; "
                "border-radius: 8px; }")

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            "CompactDropZone { background: #252527; border: 1.5px dashed #3a3a3c; "
            "border-radius: 8px; }"
            "CompactDropZone:hover { border-color: #3b82f6; background: #2a2a2c; }")

    def dropEvent(self, event: QDropEvent):
        self.dragLeaveEvent(None)
        paths = []
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isfile(p) and os.path.splitext(p)[1].lower() in self._extensions:
                paths.append(p)
                if not self._accept_multiple:
                    break
        if paths:
            self.files_dropped.emit(paths)


# ---------------------------------------------------------------------------
# Thumbnail panel (left sidebar)
# ---------------------------------------------------------------------------

class ThumbnailPanel(QWidget):
    page_clicked = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.setFixedWidth(88)
        self.setStyleSheet("background: #252527;")
        self._thumbnails = []
        self._current = -1
        self._doc = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = QLabel("Pages")
        header.setStyleSheet(
            "padding: 8px 8px 6px; font-size: 10px; color: #636366; "
            "text-transform: uppercase; letter-spacing: 1px; "
            "border-bottom: 1px solid #3a3a3c; background: #252527;")
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { border: none; background: #252527; }")
        self._container = QWidget()
        self._container.setStyleSheet("background: #252527;")
        self._container_layout = QVBoxLayout(self._container)
        self._container_layout.setContentsMargins(6, 6, 6, 6)
        self._container_layout.setSpacing(8)
        self._container_layout.setAlignment(Qt.AlignTop)
        scroll.setWidget(self._container)
        layout.addWidget(scroll)

    def load_pdf(self, pdf_path):
        self.clear()
        try:
            self._doc = fitz.open(pdf_path)
            for i in range(len(self._doc)):
                page = self._doc[i]
                pix = page.get_pixmap(dpi=72)
                # Scale to fit 60px width
                scale = 60.0 / pix.width if pix.width > 0 else 1
                w = int(pix.width * scale)
                h = int(pix.height * scale)
                img = QImage(pix.samples, pix.width, pix.height,
                             pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img).scaled(
                    w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)

                frame = QFrame()
                frame.setStyleSheet("background: transparent;")
                frame.setCursor(Qt.PointingHandCursor)
                fl = QVBoxLayout(frame)
                fl.setContentsMargins(4, 4, 4, 2)
                fl.setSpacing(2)
                fl.setAlignment(Qt.AlignCenter)

                thumb_label = QLabel()
                thumb_label.setPixmap(pixmap)
                thumb_label.setAlignment(Qt.AlignCenter)
                thumb_label.setStyleSheet(
                    "border: 1.5px solid #48484a; border-radius: 2px; "
                    "background: #f0ede8; padding: 1px;")
                fl.addWidget(thumb_label)

                num_label = QLabel(str(i + 1))
                num_label.setAlignment(Qt.AlignCenter)
                num_label.setStyleSheet("color: #636366; font-size: 10px; background: transparent;")
                fl.addWidget(num_label)

                page_num = i + 1
                frame.mousePressEvent = lambda e, n=page_num: self.page_clicked.emit(n)
                self._container_layout.addWidget(frame)
                self._thumbnails.append((frame, thumb_label, num_label))

            if self._thumbnails:
                self.set_current_page(1)
        except Exception:
            pass

    def set_current_page(self, page_num):
        idx = page_num - 1
        if self._current >= 0 and self._current < len(self._thumbnails):
            _, tl, nl = self._thumbnails[self._current]
            tl.setStyleSheet(
                "border: 1.5px solid #48484a; border-radius: 2px; "
                "background: #f0ede8; padding: 1px;")
            nl.setStyleSheet("color: #636366; font-size: 10px; background: transparent;")
        if 0 <= idx < len(self._thumbnails):
            self._current = idx
            _, tl, nl = self._thumbnails[idx]
            tl.setStyleSheet(
                "border: 1.5px solid #3b82f6; border-radius: 2px; "
                "background: #f0ede8; padding: 1px;")
            nl.setStyleSheet("color: #3b82f6; font-size: 10px; background: transparent;")

    def clear(self):
        self._thumbnails = []
        self._current = -1
        if self._doc:
            self._doc.close()
            self._doc = None
        while self._container_layout.count():
            item = self._container_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()


# ---------------------------------------------------------------------------
# PDF Viewer (center panel)
# ---------------------------------------------------------------------------

class PdfViewer(QWidget):
    page_changed = pyqtSignal(int, int)

    def __init__(self, dpi=150):
        super().__init__()
        self._dpi = dpi
        self._doc = None
        self._page_widgets = []
        self._total_pages = 0
        self._zoom = 1.0
        self._rotation = 0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Toolbar
        toolbar = QFrame()
        toolbar.setFixedHeight(40)
        toolbar.setStyleSheet(
            "QFrame { background: #2c2c2e; border-bottom: 1px solid #3a3a3c; }")
        tb_layout = QHBoxLayout(toolbar)
        tb_layout.setContentsMargins(12, 0, 12, 0)
        tb_layout.setSpacing(4)

        btn_style = (
            "QPushButton { background: none; border: none; color: #aeaeb2; "
            "font-size: 14px; padding: 5px 8px; border-radius: 5px; }"
            "QPushButton:hover { background: #3a3a3c; color: #f2f2f7; }"
            "QPushButton:disabled { color: #48484a; }")

        self._prev_btn = QPushButton("\u25C0")
        self._prev_btn.setStyleSheet(btn_style)
        self._prev_btn.setFixedSize(30, 30)
        self._prev_btn.clicked.connect(self._prev_page)
        tb_layout.addWidget(self._prev_btn)

        self._page_label = QLabel("No file loaded")
        self._page_label.setStyleSheet("color: #aeaeb2; font-size: 12px; background: transparent;")
        self._page_label.setAlignment(Qt.AlignCenter)
        self._page_label.setMinimumWidth(80)
        tb_layout.addWidget(self._page_label)

        self._next_btn = QPushButton("\u25B6")
        self._next_btn.setStyleSheet(btn_style)
        self._next_btn.setFixedSize(30, 30)
        self._next_btn.clicked.connect(self._next_page)
        tb_layout.addWidget(self._next_btn)

        sep1 = QFrame()
        sep1.setFixedSize(1, 18)
        sep1.setStyleSheet("background: #48484a;")
        tb_layout.addWidget(sep1)

        self._zoom_out_btn = QPushButton("\u2212")
        self._zoom_out_btn.setStyleSheet(btn_style)
        self._zoom_out_btn.setFixedSize(30, 30)
        self._zoom_out_btn.clicked.connect(self._zoom_out)
        tb_layout.addWidget(self._zoom_out_btn)

        self._zoom_label = QLabel("100%")
        self._zoom_label.setStyleSheet("color: #aeaeb2; font-size: 12px; background: transparent;")
        self._zoom_label.setAlignment(Qt.AlignCenter)
        self._zoom_label.setMinimumWidth(44)
        tb_layout.addWidget(self._zoom_label)

        self._zoom_in_btn = QPushButton("+")
        self._zoom_in_btn.setStyleSheet(btn_style)
        self._zoom_in_btn.setFixedSize(30, 30)
        self._zoom_in_btn.clicked.connect(self._zoom_in)
        tb_layout.addWidget(self._zoom_in_btn)

        sep2 = QFrame()
        sep2.setFixedSize(1, 18)
        sep2.setStyleSheet("background: #48484a;")
        tb_layout.addWidget(sep2)

        self._rotate_btn = QPushButton("\u21BB")
        self._rotate_btn.setStyleSheet(btn_style)
        self._rotate_btn.setFixedSize(30, 30)
        self._rotate_btn.setToolTip("Rotate")
        self._rotate_btn.clicked.connect(self._rotate)
        tb_layout.addWidget(self._rotate_btn)

        self._fit_btn = QPushButton("\u2B1C")
        self._fit_btn.setStyleSheet(btn_style)
        self._fit_btn.setFixedSize(30, 30)
        self._fit_btn.setToolTip("Fit page")
        self._fit_btn.clicked.connect(self._fit_page)
        tb_layout.addWidget(self._fit_btn)

        tb_layout.addStretch()
        layout.addWidget(toolbar)

        # Scroll area for pages
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setStyleSheet("QScrollArea { background: #525252; border: none; }")
        self._scroll.verticalScrollBar().valueChanged.connect(self._on_scroll)

        self._pages_container = QWidget()
        self._pages_container.setStyleSheet("background: #525252;")
        self._pages_layout = QVBoxLayout(self._pages_container)
        self._pages_layout.setContentsMargins(24, 24, 24, 24)
        self._pages_layout.setSpacing(20)
        self._pages_layout.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self._scroll.setWidget(self._pages_container)
        layout.addWidget(self._scroll)

        # Drop zone (shown when no file loaded)
        self._drop_zone_label = QLabel("Open a PDF to get started")
        self._drop_zone_label.setAlignment(Qt.AlignCenter)
        self._drop_zone_label.setStyleSheet(
            "color: #636366; font-size: 16px; background: transparent;")
        self._pages_layout.addWidget(self._drop_zone_label)

    def load_pdf(self, pdf_path):
        self._clear()
        try:
            self._doc = fitz.open(pdf_path)
            self._total_pages = len(self._doc)
            self._drop_zone_label.hide()
            self._fit_page()  # auto-fit to viewport width on load
            self._page_label.setText(f"Page 1 of {self._total_pages}")
            self.page_changed.emit(1, self._total_pages)
        except Exception:
            self._page_label.setText("Failed to load PDF")

    def _render_pages(self):
        for w in self._page_widgets:
            w.deleteLater()
        self._page_widgets = []
        if not self._doc:
            return

        mat = fitz.Matrix(self._dpi * self._zoom / 72.0, self._dpi * self._zoom / 72.0)
        mat = mat.prerotate(self._rotation)

        for i in range(self._total_pages):
            page = self._doc[i]
            pix = page.get_pixmap(matrix=mat)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride,
                         QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img)

            frame = QFrame()
            frame.setStyleSheet(
                "QFrame { background: white; border-radius: 2px; }")
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(20)
            shadow.setOffset(0, 4)
            shadow.setColor(QColor(0, 0, 0, 128))
            frame.setGraphicsEffect(shadow)

            fl = QVBoxLayout(frame)
            fl.setContentsMargins(0, 0, 0, 0)
            label = QLabel()
            label.setPixmap(pixmap)
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet("background: white;")
            fl.addWidget(label)

            self._pages_layout.addWidget(frame)
            self._page_widgets.append(frame)

    def _clear(self):
        for w in self._page_widgets:
            w.deleteLater()
        self._page_widgets = []
        self._total_pages = 0
        self._rotation = 0
        self._zoom = 1.0
        self._zoom_label.setText("100%")
        if self._doc:
            self._doc.close()
            self._doc = None
        self._drop_zone_label.show()
        self._page_label.setText("No file loaded")

    def _zoom_in(self):
        if self._zoom < 3.0:
            self._zoom = min(self._zoom + 0.25, 3.0)
            self._zoom_label.setText(f"{int(self._zoom * 100)}%")
            self._render_pages()

    def _zoom_out(self):
        if self._zoom > 0.25:
            self._zoom = max(self._zoom - 0.25, 0.25)
            self._zoom_label.setText(f"{int(self._zoom * 100)}%")
            self._render_pages()

    def _rotate(self):
        self._rotation = (self._rotation + 90) % 360
        self._render_pages()

    def _fit_page(self):
        """Fit page width to available viewport width."""
        if not self._doc or self._total_pages == 0:
            return
        page = self._doc[0]
        if self._rotation in (90, 270):
            page_width_pts = page.rect.height
        else:
            page_width_pts = page.rect.width
        # Available width = scroll area width minus padding (24 each side) minus scrollbar
        available = self._scroll.viewport().width() - 48 - 20
        if available < 100:
            available = 400
        # page_width at zoom=1.0 is page_width_pts * dpi / 72
        base_pixel_width = page_width_pts * self._dpi / 72.0
        if base_pixel_width > 0:
            self._zoom = available / base_pixel_width
            self._zoom = max(0.25, min(self._zoom, 3.0))
        else:
            self._zoom = 1.0
        self._zoom_label.setText(f"{int(self._zoom * 100)}%")
        self._render_pages()

    def _prev_page(self):
        current = self._get_current_page()
        if current > 1:
            self.scroll_to_page(current - 1)

    def _next_page(self):
        current = self._get_current_page()
        if current < self._total_pages:
            self.scroll_to_page(current + 1)

    def _get_current_page(self):
        if not self._page_widgets:
            return 1
        scroll_y = self._scroll.verticalScrollBar().value()
        for i, w in enumerate(self._page_widgets):
            if w.geometry().bottom() > scroll_y:
                return i + 1
        return self._total_pages

    def _on_scroll(self):
        if not self._page_widgets:
            return
        current = self._get_current_page()
        self._page_label.setText(f"Page {current} of {self._total_pages}")
        self.page_changed.emit(current, self._total_pages)

    def scroll_to_page(self, page_num):
        idx = page_num - 1
        if 0 <= idx < len(self._page_widgets):
            self._scroll.ensureWidgetVisible(self._page_widgets[idx])

    def resizeEvent(self, event):
        """Re-fit pages when viewer is resized."""
        super().resizeEvent(event)
        if self._doc and self._total_pages > 0:
            self._fit_page()

    def get_total_pages(self):
        return self._total_pages


# ---------------------------------------------------------------------------
# Tool panels (right sidebar, one per tab)
# ---------------------------------------------------------------------------

class _ToolPanelBase(QWidget):
    """Base class for right-panel tool widgets. Provides common patterns."""

    def __init__(self, main_window=None):
        super().__init__()
        self._main_window = main_window
        self.input_path = ""
        self.output_tmp_path = ""
        self.worker = None
        self._progress_timer = QTimer()
        self._progress_timer.setInterval(150)
        self._progress_value = 0

        self.setStyleSheet("background: #252527;")
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(14, 14, 14, 14)
        self._layout.setSpacing(10)

    def _add_header(self, title):
        lbl = QLabel(title)
        lbl.setStyleSheet(
            "font-size: 11px; font-weight: 500; color: #aeaeb2; "
            "text-transform: uppercase; letter-spacing: 1px; background: transparent; "
            "padding-bottom: 8px; border-bottom: 1px solid #3a3a3c;")
        self._layout.addWidget(lbl)

    def _add_file_info(self):
        self._file_info = QLabel("")
        self._file_info.setWordWrap(True)
        self._file_info.setStyleSheet("color: #f2f2f7; font-size: 12px; background: transparent;")
        self._file_info.hide()
        self._layout.addWidget(self._file_info)

    def _add_progress(self):
        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setFixedHeight(6)
        self._progress.hide()
        self._layout.addWidget(self._progress)
        self._progress_timer.timeout.connect(self._tick_progress)

    def _add_result_section(self):
        self._result_frame = QFrame()
        self._result_frame.setStyleSheet("background: transparent;")
        self._result_frame.hide()
        rl = QVBoxLayout(self._result_frame)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(8)

        self._result_text = QLabel("")
        self._result_text.setWordWrap(True)
        self._result_text.setStyleSheet("color: #f2f2f7; font-size: 12px; background: transparent;")
        rl.addWidget(self._result_text)

        self._size_warning = QLabel("")
        self._size_warning.setWordWrap(True)
        self._size_warning.setStyleSheet("color: #f59e0b; font-weight: bold; font-size: 11px; background: transparent;")
        self._size_warning.hide()
        rl.addWidget(self._size_warning)

        name_row = QHBoxLayout()
        name_row.setSpacing(4)
        self._name_input = QLineEdit()
        self._name_input.setPlaceholderText("File name")
        name_row.addWidget(self._name_input)
        self._ext_label = QLabel(".pdf")
        self._ext_label.setStyleSheet("color: #aeaeb2; font-size: 12px; background: transparent;")
        name_row.addWidget(self._ext_label)
        rl.addLayout(name_row)

        self._save_btn = QPushButton("Save As")
        self._save_btn.setStyleSheet(BTN_SUCCESS)
        self._save_btn.setCursor(Qt.PointingHandCursor)
        self._save_btn.clicked.connect(self._save)
        rl.addWidget(self._save_btn)

        self._layout.addWidget(self._result_frame)

    def _add_error_label(self):
        self._error_label = QLabel("")
        self._error_label.setWordWrap(True)
        self._error_label.setStyleSheet("color: #ef4444; font-size: 12px; background: transparent;")
        self._error_label.hide()
        self._layout.addWidget(self._error_label)

    def _tick_progress(self):
        if self._progress_value < 70:
            self._progress_value += 1.2
        elif self._progress_value < 90:
            self._progress_value += 0.4
        else:
            self._progress_value = min(self._progress_value + 0.05, 99)
        self._progress.setValue(int(self._progress_value))

    def _start_progress(self):
        self._progress_value = 0
        self._progress.setValue(0)
        self._progress.show()
        self._progress_timer.start()

    def _stop_progress(self):
        self._progress_timer.stop()
        self._progress.setValue(100)

    def _load_in_viewer(self, path):
        if self._main_window and hasattr(self._main_window, 'load_pdf'):
            self._main_window.load_pdf(path)

    def _reset_result(self):
        self._result_frame.hide()
        self._size_warning.hide()
        self._error_label.hide()
        self._progress.hide()
        self._progress_value = 0
        self._save_btn.setEnabled(True)

    def _save(self):
        if not self.output_tmp_path or not os.path.isfile(self.output_tmp_path):
            return
        ext = self._ext_label.text()
        name = self._name_input.text().strip()
        if not name:
            name = "output"
        if not name.endswith(ext):
            name += ext
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save file", name, f"Files (*{ext})")
        if save_path:
            shutil.copy2(self.output_tmp_path, save_path)
            self._save_btn.setEnabled(False)
            self._save_btn.setText("Saved!")
            self._reset_for_next()

    def _reset_for_next(self):
        """Override in subclasses to prepare for next file."""
        pass


class CompressToolPanel(_ToolPanelBase):
    def __init__(self, gs_exe, main_window=None):
        super().__init__(main_window)
        self.gs_exe = gs_exe
        self._is_image = False
        self.output_ext = ".pdf"

        self._add_header("Compress")

        self._drop = CompactDropZone(
            "Drop PDF or image here",
            file_extensions=[".pdf"] + IMAGE_EXTENSIONS,
            file_filter_name="PDF/Image")
        self._drop.files_dropped.connect(self._on_file_dropped)
        self._layout.addWidget(self._drop)

        self._add_file_info()

        # Quality slider for images
        self._quality_frame = QFrame()
        self._quality_frame.setStyleSheet("background: transparent;")
        self._quality_frame.hide()
        ql = QVBoxLayout(self._quality_frame)
        ql.setContentsMargins(0, 0, 0, 0)
        ql.setSpacing(4)
        ql_lbl = QLabel("Image quality")
        ql_lbl.setStyleSheet("color: #aeaeb2; font-size: 11px; background: transparent;")
        ql.addWidget(ql_lbl)
        row = QHBoxLayout()
        row.setSpacing(8)
        self._quality_slider = QSlider(Qt.Horizontal)
        self._quality_slider.setRange(20, 95)
        self._quality_slider.setValue(75)
        self._quality_slider.valueChanged.connect(
            lambda v: self._quality_val.setText(f"{v}%"))
        row.addWidget(self._quality_slider)
        self._quality_val = QLabel("75%")
        self._quality_val.setStyleSheet("color: #aeaeb2; font-size: 12px; background: transparent; min-width: 32px;")
        self._quality_val.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        row.addWidget(self._quality_val)
        ql.addLayout(row)
        self._layout.addWidget(self._quality_frame)

        self._compress_btn = QPushButton("Compress")
        self._compress_btn.setStyleSheet(BTN_PRIMARY)
        self._compress_btn.setCursor(Qt.PointingHandCursor)
        self._compress_btn.clicked.connect(self._start_compress)
        self._compress_btn.hide()
        self._layout.addWidget(self._compress_btn)

        self._add_progress()
        self._add_result_section()
        self._add_error_label()
        self._layout.addStretch()

    def _on_file_dropped(self, paths):
        self._reset_result()
        self.input_path = paths[0]
        name = Path(self.input_path).name
        size = human_size(os.path.getsize(self.input_path))
        self._file_info.setText(f"{name}\n{size}")
        self._file_info.show()

        ext = os.path.splitext(self.input_path)[1].lower()
        self._is_image = ext in IMAGE_EXTENSIONS
        self._quality_frame.setVisible(self._is_image)
        self._compress_btn.show()

        if not self._is_image:
            self._load_in_viewer(self.input_path)
            orig_size = os.path.getsize(self.input_path)
            if orig_size <= TARGET_SIZE_BYTES:
                self._file_info.setText(f"{name}\n{size} (already under {TARGET_SIZE_MB} MB)")

    def _start_compress(self):
        if not self.input_path:
            return
        self._reset_result()
        self._compress_btn.setEnabled(False)
        self._drop.setEnabled(False)
        self._start_progress()

        if self._is_image:
            self.worker = ImageCompressWorker(self.input_path, self._quality_slider.value())
            self.worker.finished.connect(self._on_image_finished)
        else:
            self.worker = CompressWorker(self.input_path, self.gs_exe)
            self.worker.finished.connect(self._on_pdf_finished)
        self.worker.start()

    def _on_pdf_finished(self, success, output_path, orig_size, new_size):
        self._stop_progress()
        self._compress_btn.setEnabled(True)
        self._drop.setEnabled(True)

        if success:
            self.output_tmp_path = output_path
            self.output_ext = ".pdf"
            self._ext_label.setText(".pdf")
            pct = int((1 - new_size / orig_size) * 100) if orig_size > 0 else 0
            self._result_text.setText(
                f"{human_size(orig_size)} \u2192 {human_size(new_size)}  ({pct}% reduction)")

            if new_size > TARGET_SIZE_BYTES:
                self._size_warning.setText(
                    f"Note: Still {human_size(new_size)}, over {TARGET_SIZE_MB} MB.\n"
                    "The original may contain high-resolution scans.")
                self._size_warning.show()
            elif orig_size > TARGET_SIZE_BYTES * 5 and new_size <= TARGET_SIZE_BYTES:
                self._size_warning.setText(
                    "Image quality was reduced to meet the 10 MB limit.\n"
                    "Text is still readable but images may appear lower quality.")
                self._size_warning.setStyleSheet("color: #f59e0b; font-weight: bold; font-size: 11px; background: transparent;")
                self._size_warning.show()

            self._name_input.setText(Path(self.input_path).stem + " - Compressed")
            self._result_frame.show()
            self._save_btn.setEnabled(True)
            self._save_btn.setText("Save As")
        else:
            self._error_label.setText("Compression failed. The file may be corrupted or protected.")
            self._error_label.show()

    def _on_image_finished(self, success, output_path, orig_size, new_size, fmt):
        self._stop_progress()
        self._compress_btn.setEnabled(True)
        self._drop.setEnabled(True)

        if success:
            self.output_tmp_path = output_path
            self.output_ext = ".png" if fmt == "PNG" else ".jpg"
            self._ext_label.setText(self.output_ext)
            pct = int((1 - new_size / orig_size) * 100) if orig_size > 0 else 0
            self._result_text.setText(
                f"{human_size(orig_size)} \u2192 {human_size(new_size)}  ({pct}% reduction)")
            self._name_input.setText(Path(self.input_path).stem + " - Compressed")
            self._result_frame.show()
            self._save_btn.setEnabled(True)
            self._save_btn.setText("Save As")
        else:
            self._error_label.setText("Image compression failed.")
            self._error_label.show()


class MergeToolPanel(_ToolPanelBase):
    def __init__(self, gs_exe, main_window=None):
        super().__init__(main_window)
        self.gs_exe = gs_exe
        self.file_paths = []

        self._add_header("Merge")

        self._drop = CompactDropZone(
            "Drop PDFs here (multiple)",
            accept_multiple=True)
        self._drop.files_dropped.connect(self._on_files_dropped)
        self._layout.addWidget(self._drop)

        self._file_list = QListWidget()
        self._file_list.setMaximumHeight(150)
        self._file_list.setDragDropMode(QAbstractItemView.InternalMove)
        self._file_list.hide()
        self._layout.addWidget(self._file_list)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(4)
        for text, handler in [("\u2191", self._move_up), ("\u2193", self._move_down), ("\u2715", self._remove_selected)]:
            b = QPushButton(text)
            b.setFixedSize(30, 26)
            b.setStyleSheet(BTN_SECONDARY)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(handler)
            btn_row.addWidget(b)
        btn_row.addStretch()
        self._btn_row_widget = QWidget()
        self._btn_row_widget.setLayout(btn_row)
        self._btn_row_widget.setStyleSheet("background: transparent;")
        self._btn_row_widget.hide()
        self._layout.addWidget(self._btn_row_widget)

        self._merge_btn = QPushButton("Merge PDFs")
        self._merge_btn.setStyleSheet(BTN_PRIMARY)
        self._merge_btn.setCursor(Qt.PointingHandCursor)
        self._merge_btn.clicked.connect(self._start_merge)
        self._merge_btn.hide()
        self._layout.addWidget(self._merge_btn)

        self._add_progress()
        self._add_result_section()
        self._add_error_label()
        self._layout.addStretch()

    def _on_files_dropped(self, paths):
        for p in paths:
            if p not in self.file_paths:
                self.file_paths.append(p)
        self._refresh_list()
        if self.file_paths:
            self._load_in_viewer(self.file_paths[0])

    def _refresh_list(self):
        self._file_list.clear()
        for p in self.file_paths:
            name = Path(p).name
            size = human_size(os.path.getsize(p))
            self._file_list.addItem(f"{name}  ({size})")
        visible = len(self.file_paths) > 0
        self._file_list.setVisible(visible)
        self._btn_row_widget.setVisible(visible)
        self._merge_btn.setVisible(len(self.file_paths) >= 2)
        self._merge_btn.setEnabled(len(self.file_paths) >= 2)

    def _move_up(self):
        row = self._file_list.currentRow()
        if row > 0:
            self.file_paths[row], self.file_paths[row - 1] = self.file_paths[row - 1], self.file_paths[row]
            self._refresh_list()
            self._file_list.setCurrentRow(row - 1)

    def _move_down(self):
        row = self._file_list.currentRow()
        if 0 <= row < len(self.file_paths) - 1:
            self.file_paths[row], self.file_paths[row + 1] = self.file_paths[row + 1], self.file_paths[row]
            self._refresh_list()
            self._file_list.setCurrentRow(row + 1)

    def _remove_selected(self):
        row = self._file_list.currentRow()
        if 0 <= row < len(self.file_paths):
            self.file_paths.pop(row)
            self._refresh_list()

    def _start_merge(self):
        if len(self.file_paths) < 2:
            return
        self._reset_result()
        self._merge_btn.setEnabled(False)
        self._drop.setEnabled(False)
        self._start_progress()
        self.worker = MergeWorker(list(self.file_paths), self.gs_exe)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_finished(self, success, output_path, combined_size, new_size):
        self._stop_progress()
        self._merge_btn.setEnabled(True)
        self._drop.setEnabled(True)

        if success:
            self.output_tmp_path = output_path
            self._ext_label.setText(".pdf")
            self._result_text.setText(
                f"Combined {human_size(combined_size)} \u2192 {human_size(new_size)}")
            if new_size > TARGET_SIZE_BYTES:
                self._size_warning.setText(
                    f"Note: Still {human_size(new_size)}, over {TARGET_SIZE_MB} MB.")
                self._size_warning.show()
            self._name_input.setText("Merged Document")
            self._result_frame.show()
            self._save_btn.setEnabled(True)
            self._save_btn.setText("Save As")
        else:
            self._error_label.setText("Merge failed. One of the files may be corrupted.")
            self._error_label.show()

    def _reset_for_next(self):
        self.file_paths = []
        self._refresh_list()


class SplitToolPanel(_ToolPanelBase):
    def __init__(self, main_window=None):
        super().__init__(main_window)
        self.total_pages = 0

        self._add_header("Split")

        self._drop = CompactDropZone()
        self._drop.files_dropped.connect(self._on_file_dropped)
        self._layout.addWidget(self._drop)

        self._add_file_info()

        self._split_frame = QFrame()
        self._split_frame.setStyleSheet("background: transparent;")
        self._split_frame.hide()
        sf_layout = QVBoxLayout(self._split_frame)
        sf_layout.setContentsMargins(0, 0, 0, 0)
        sf_layout.setSpacing(8)

        row = QHBoxLayout()
        row.setSpacing(8)
        lbl = QLabel("Split after page:")
        lbl.setStyleSheet("color: #aeaeb2; font-size: 12px; background: transparent;")
        row.addWidget(lbl)
        self._page_spin = QSpinBox()
        self._page_spin.setMinimum(1)
        self._page_spin.setMaximum(1)
        self._page_spin.valueChanged.connect(self._update_preview)
        row.addWidget(self._page_spin)
        row.addStretch()
        sf_layout.addLayout(row)

        self._preview_label = QLabel("")
        self._preview_label.setWordWrap(True)
        self._preview_label.setStyleSheet("color: #aeaeb2; font-size: 11px; background: transparent;")
        sf_layout.addWidget(self._preview_label)

        self._split_btn = QPushButton("Split PDF")
        self._split_btn.setStyleSheet(BTN_PRIMARY)
        self._split_btn.setCursor(Qt.PointingHandCursor)
        self._split_btn.clicked.connect(self._split)
        sf_layout.addWidget(self._split_btn)

        self._layout.addWidget(self._split_frame)

        self._result_label = QLabel("")
        self._result_label.setWordWrap(True)
        self._result_label.setStyleSheet("color: #22c55e; font-size: 12px; background: transparent;")
        self._result_label.hide()
        self._layout.addWidget(self._result_label)

        self._add_error_label()
        self._layout.addStretch()

    def _on_file_dropped(self, paths):
        self.input_path = paths[0]
        self._result_label.hide()
        self._error_label.hide()
        try:
            reader = PdfReader(self.input_path)
            self.total_pages = len(reader.pages)
        except Exception:
            self._error_label.setText("Could not read PDF.")
            self._error_label.show()
            return

        if self.total_pages < 2:
            self._error_label.setText("PDF has only 1 page — cannot split.")
            self._error_label.show()
            return

        name = Path(self.input_path).name
        self._file_info.setText(f"{name}  ({self.total_pages} pages)")
        self._file_info.show()
        self._page_spin.setMaximum(self.total_pages - 1)
        self._page_spin.setValue(1)
        self._update_preview()
        self._split_frame.show()
        self._load_in_viewer(self.input_path)

    def _update_preview(self):
        p = self._page_spin.value()
        self._preview_label.setText(
            f"Part 1: pages 1\u2013{p}  |  Part 2: pages {p+1}\u2013{self.total_pages}")

    def _split(self):
        if not self.input_path or self.total_pages < 2:
            return
        try:
            reader = PdfReader(self.input_path)
            split_at = self._page_spin.value()
            stem = Path(self.input_path).stem
            parent = str(Path(self.input_path).parent)

            w1 = PdfWriter()
            for i in range(split_at):
                w1.add_page(reader.pages[i])
            p1 = os.path.join(parent, f"{stem} - Part 1.pdf")
            with open(p1, "wb") as f:
                w1.write(f)

            w2 = PdfWriter()
            for i in range(split_at, len(reader.pages)):
                w2.add_page(reader.pages[i])
            p2 = os.path.join(parent, f"{stem} - Part 2.pdf")
            with open(p2, "wb") as f:
                w2.write(f)

            self._result_label.setText(
                f"Split into:\n{Path(p1).name} (pages 1\u2013{split_at})\n"
                f"{Path(p2).name} (pages {split_at+1}\u2013{self.total_pages})")
            self._result_label.show()
        except Exception:
            self._error_label.setText("Split failed.")
            self._error_label.show()


class FlattenToolPanel(_ToolPanelBase):
    def __init__(self, main_window=None):
        super().__init__(main_window)

        self._add_header("Flatten")

        self._drop = CompactDropZone()
        self._drop.files_dropped.connect(self._on_file_dropped)
        self._layout.addWidget(self._drop)

        info = QLabel(
            "Flattening merges all form fields, annotations, and layers "
            "into static page content. The PDF becomes non-editable.")
        info.setWordWrap(True)
        info.setStyleSheet("color: #aeaeb2; font-size: 11px; background: transparent; line-height: 1.5;")
        self._layout.addWidget(info)

        self._add_file_info()

        self._flatten_btn = QPushButton("Flatten PDF")
        self._flatten_btn.setStyleSheet(BTN_PRIMARY)
        self._flatten_btn.setCursor(Qt.PointingHandCursor)
        self._flatten_btn.clicked.connect(self._start_flatten)
        self._flatten_btn.hide()
        self._layout.addWidget(self._flatten_btn)

        self._add_progress()
        self._add_result_section()
        self._add_error_label()
        self._layout.addStretch()

    def _on_file_dropped(self, paths):
        self._reset_result()
        self.input_path = paths[0]
        name = Path(self.input_path).name
        size = human_size(os.path.getsize(self.input_path))
        self._file_info.setText(f"{name}\n{size}")
        self._file_info.show()
        self._flatten_btn.show()
        self._load_in_viewer(self.input_path)

    def _start_flatten(self):
        if not self.input_path:
            return
        self._reset_result()
        self._flatten_btn.setEnabled(False)
        self._drop.setEnabled(False)
        self._start_progress()
        self.worker = FlattenWorker(self.input_path)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_finished(self, success, output_path, orig_size, new_size):
        self._stop_progress()
        self._flatten_btn.setEnabled(True)
        self._drop.setEnabled(True)

        if success:
            self.output_tmp_path = output_path
            self._ext_label.setText(".pdf")
            self._result_text.setText(
                f"Flattened: {human_size(orig_size)} \u2192 {human_size(new_size)}")
            self._name_input.setText(Path(self.input_path).stem + " - Flattened")
            self._result_frame.show()
            self._save_btn.setEnabled(True)
            self._save_btn.setText("Save As")
        else:
            self._error_label.setText("Flatten failed.")
            self._error_label.show()


class RedactToolPanel(_ToolPanelBase):
    def __init__(self, main_window=None):
        super().__init__(main_window)

        self._add_header("Redact")

        self._drop = CompactDropZone()
        self._drop.files_dropped.connect(self._on_file_dropped)
        self._layout.addWidget(self._drop)

        self._add_file_info()

        self._search_frame = QFrame()
        self._search_frame.setStyleSheet("background: transparent;")
        self._search_frame.hide()
        sl = QVBoxLayout(self._search_frame)
        sl.setContentsMargins(0, 0, 0, 0)
        sl.setSpacing(6)

        lbl = QLabel("Text to redact (one per line):")
        lbl.setStyleSheet("color: #aeaeb2; font-size: 11px; background: transparent;")
        sl.addWidget(lbl)

        self._search_input = QTextEdit()
        self._search_input.setPlaceholderText("e.g.\nJohn Smith\n555-1234")
        self._search_input.setMaximumHeight(80)
        sl.addWidget(self._search_input)

        self._redact_btn = QPushButton("Redact")
        self._redact_btn.setStyleSheet(BTN_DANGER)
        self._redact_btn.setCursor(Qt.PointingHandCursor)
        self._redact_btn.clicked.connect(self._start_redact)
        sl.addWidget(self._redact_btn)

        self._layout.addWidget(self._search_frame)

        warn = QLabel("Redactions are permanent and cannot be undone.")
        warn.setWordWrap(True)
        warn.setStyleSheet("color: #636366; font-size: 10px; background: transparent;")
        self._layout.addWidget(warn)

        self._add_progress()
        self._add_result_section()
        self._add_error_label()
        self._layout.addStretch()

    def _on_file_dropped(self, paths):
        self._reset_result()
        self.input_path = paths[0]
        name = Path(self.input_path).name
        size = human_size(os.path.getsize(self.input_path))
        self._file_info.setText(f"{name}\n{size}")
        self._file_info.show()
        self._search_frame.show()
        self._load_in_viewer(self.input_path)

    def _start_redact(self):
        if not self.input_path:
            return
        text = self._search_input.toPlainText().strip()
        if not text:
            self._error_label.setText("Enter at least one term to redact.")
            self._error_label.show()
            return
        terms = [t.strip() for t in text.splitlines() if t.strip()]
        if not terms:
            return

        reply = QMessageBox.warning(
            self, "Confirm Redaction",
            f"This will permanently redact {len(terms)} term(s) from the PDF.\n\nContinue?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        self._reset_result()
        self._redact_btn.setEnabled(False)
        self._drop.setEnabled(False)
        self._start_progress()
        self.worker = RedactWorker(self.input_path, terms)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_finished(self, success, output_path, orig_size, new_size, count):
        self._stop_progress()
        self._redact_btn.setEnabled(True)
        self._drop.setEnabled(True)

        if success:
            self.output_tmp_path = output_path
            self._ext_label.setText(".pdf")
            if count == 0:
                self._result_text.setText("No matches found — nothing was redacted.")
            else:
                self._result_text.setText(f"Redacted {count} occurrence(s).")
            self._name_input.setText(Path(self.input_path).stem + " - Redacted")
            self._result_frame.show()
            self._save_btn.setEnabled(True)
            self._save_btn.setText("Save As")
        else:
            self._error_label.setText("Redaction failed.")
            self._error_label.show()


class OCRToolPanel(_ToolPanelBase):
    def __init__(self, main_window=None):
        super().__init__(main_window)

        self._add_header("OCR")

        self._drop = CompactDropZone()
        self._drop.files_dropped.connect(self._on_file_dropped)
        self._layout.addWidget(self._drop)

        info = QLabel("Extract text from scanned PDFs using OCR.\nMakes the PDF searchable and copyable.")
        info.setWordWrap(True)
        info.setStyleSheet("color: #aeaeb2; font-size: 11px; background: transparent;")
        self._layout.addWidget(info)

        if not HAS_TESSERACT:
            warn = QLabel(
                "\u26A0 Tesseract OCR engine was not found.\n"
                "Please ask your IT team to install it.")
            warn.setWordWrap(True)
            warn.setStyleSheet("color: #f59e0b; font-size: 11px; background: transparent;")
            self._layout.addWidget(warn)
            self._drop.setEnabled(False)

        self._add_file_info()

        self._ocr_btn = QPushButton("Run OCR")
        self._ocr_btn.setStyleSheet(BTN_PRIMARY)
        self._ocr_btn.setCursor(Qt.PointingHandCursor)
        self._ocr_btn.clicked.connect(self._start_ocr)
        self._ocr_btn.hide()
        self._layout.addWidget(self._ocr_btn)

        self._add_progress()
        self._add_result_section()
        self._add_error_label()
        self._layout.addStretch()

    def _on_file_dropped(self, paths):
        self._reset_result()
        self.input_path = paths[0]
        name = Path(self.input_path).name
        size = human_size(os.path.getsize(self.input_path))
        self._file_info.setText(f"{name}\n{size}")
        self._file_info.show()
        self._ocr_btn.show()
        self._load_in_viewer(self.input_path)

    def _start_ocr(self):
        if not self.input_path:
            return
        self._reset_result()
        self._ocr_btn.setEnabled(False)
        self._drop.setEnabled(False)
        self._start_progress()
        self.worker = OCRWorker(self.input_path)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_finished(self, success, output_path, orig_size, new_size):
        self._stop_progress()
        self._ocr_btn.setEnabled(True)
        self._drop.setEnabled(True)

        if success:
            self.output_tmp_path = output_path
            self._ext_label.setText(".pdf")
            self._result_text.setText(
                f"OCR complete: {human_size(new_size)}")
            self._name_input.setText(Path(self.input_path).stem + " - OCR")
            self._result_frame.show()
            self._save_btn.setEnabled(True)
            self._save_btn.setText("Save As")
        else:
            self._error_label.setText("OCR failed. Tesseract may not be installed.")
            self._error_label.show()


# ---------------------------------------------------------------------------
# Main window — three-panel layout
# ---------------------------------------------------------------------------

class LandingDropZone(QFrame):
    """Full-screen landing drop zone shown before any file is loaded."""
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setStyleSheet("background: #1c1c1e;")
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(16)

        icon = QLabel("\U0001F4C4")
        icon.setAlignment(Qt.AlignCenter)
        icon.setStyleSheet("font-size: 48px; background: transparent;")
        layout.addWidget(icon)

        title = QLabel("Open a PDF to get started")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #f2f2f7; font-size: 18px; font-weight: 500; background: transparent;")
        layout.addWidget(title)

        sub = QLabel("Drag and drop a PDF or image file here, or click Browse")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet("color: #636366; font-size: 13px; background: transparent;")
        layout.addWidget(sub)

        browse_btn = QPushButton("Browse Files")
        browse_btn.setFixedWidth(160)
        browse_btn.setStyleSheet(BTN_PRIMARY)
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._browse)
        layout.addWidget(browse_btn, alignment=Qt.AlignCenter)

    def _browse(self):
        exts = " ".join(f"*{e}" for e in [".pdf"] + IMAGE_EXTENSIONS)
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select files", "", f"PDF/Image Files ({exts})")
        if paths:
            self.files_dropped.emit(paths)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = []
        valid_exts = [".pdf"] + IMAGE_EXTENSIONS
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isfile(p) and os.path.splitext(p)[1].lower() in valid_exts:
                paths.append(p)
        if paths:
            self.files_dropped.emit(paths)


class FileInfoBar(QFrame):
    """Persistent bar showing current file info."""

    def __init__(self):
        super().__init__()
        self.setFixedHeight(32)
        self.setStyleSheet(
            "QFrame { background: #252527; border-bottom: 1px solid #3a3a3c; }")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 0, 12, 0)
        layout.setSpacing(16)

        self._name_label = QLabel("")
        self._name_label.setStyleSheet("color: #f2f2f7; font-size: 12px; font-weight: 500; background: transparent;")
        layout.addWidget(self._name_label)

        self._pages_label = QLabel("")
        self._pages_label.setStyleSheet("color: #aeaeb2; font-size: 11px; background: transparent;")
        layout.addWidget(self._pages_label)

        self._size_label = QLabel("")
        self._size_label.setStyleSheet("color: #aeaeb2; font-size: 11px; background: transparent;")
        layout.addWidget(self._size_label)

        layout.addStretch()

    def update_info(self, name="", pages=0, size_bytes=0):
        self._name_label.setText(name)
        self._pages_label.setText(f"{pages} pages" if pages else "")
        self._size_label.setText(human_size(size_bytes) if size_bytes else "")



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Combiner")
        self.resize(1200, 800)
        self.setMinimumSize(1000, 700)

        self.gs_exe = find_ghostscript()
        self._current_pdf = ""
        self._file_loaded = False

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Update banner placeholder
        self._update_banner = None
        self._banner_container = QVBoxLayout()
        self._banner_container.setContentsMargins(8, 4, 8, 0)
        main_layout.addLayout(self._banner_container)

        # Error banner for missing dependencies (shown as top-of-screen banner)
        self._error_banner = QFrame()
        self._error_banner.setStyleSheet(
            "QFrame { background: #3b1818; border-bottom: 1px solid #ef4444; }")
        eb_layout = QHBoxLayout(self._error_banner)
        eb_layout.setContentsMargins(12, 8, 12, 8)
        self._error_banner_label = QLabel("")
        self._error_banner_label.setWordWrap(True)
        self._error_banner_label.setStyleSheet("color: #fca5a5; font-size: 12px; background: transparent;")
        eb_layout.addWidget(self._error_banner_label)
        self._error_banner.hide()
        main_layout.addWidget(self._error_banner)

        if not self.gs_exe:
            self._error_banner_label.setText(
                "\u26A0  Ghostscript was not found. Compression and merge will not work. "
                "Please ask your IT team to install Ghostscript.")
            self._error_banner.show()

        # Operation toolbar (hidden until file loaded)
        self._toolbar = QFrame()
        self._toolbar.setFixedHeight(48)
        self._toolbar.setStyleSheet(
            "QFrame { background: #141416; border-bottom: 1px solid #3a3a3c; }")
        self._toolbar.hide()
        tb_layout = QHBoxLayout(self._toolbar)
        tb_layout.setContentsMargins(8, 0, 8, 0)
        tb_layout.setSpacing(0)

        self._op_buttons = []
        self._op_labels = []  # store labels for conditional thumbnail lookup
        self._panel_stack = QStackedWidget()
        self._panel_stack.setFixedWidth(236)
        self._panel_stack.setStyleSheet("background: #252527;")

        all_ops = [
            ("Compress", "\u2B07", CompressToolPanel(self.gs_exe, self)),
            ("Merge", "\u2795", MergeToolPanel(self.gs_exe, self)),
            ("Split", "\u2702", SplitToolPanel(self)),
            ("Flatten", "\u25A3", FlattenToolPanel(self)),
            ("OCR", "\U0001F50D", OCRToolPanel(self)),
            ("Redact", "\u2588", RedactToolPanel(self)),
        ]

        op_idx = 0
        for label, icon, panel in all_ops:
            if ENABLED_TABS is not None and label not in ENABLED_TABS:
                continue

            btn = QPushButton(f" {icon}  {label}")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(
                "QPushButton { background: none; border: none; border-bottom: 2px solid transparent; "
                "color: #636366; font-size: 13px; padding: 0 16px; min-height: 46px; }"
                "QPushButton:hover { color: #aeaeb2; }"
                "QPushButton:checked { color: #f2f2f7; border-bottom-color: #3b82f6; font-weight: 500; }")
            idx = op_idx
            btn.clicked.connect(lambda checked, i=idx: self._switch_op(i))
            tb_layout.addWidget(btn)
            self._op_buttons.append(btn)
            self._op_labels.append(label)

            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setStyleSheet("QScrollArea { background: #252527; border: none; }")
            scroll.setWidget(panel)
            self._panel_stack.addWidget(scroll)
            op_idx += 1

        tb_layout.addStretch()
        main_layout.addWidget(self._toolbar)

        # File info bar (hidden until file loaded)
        self._file_info_bar = FileInfoBar()
        self._file_info_bar.hide()
        main_layout.addWidget(self._file_info_bar)

        # Landing drop zone (visible when no file loaded)
        self._landing = LandingDropZone()
        self._landing.files_dropped.connect(self._on_landing_drop)
        main_layout.addWidget(self._landing, 1)

        # Body: three-panel layout (hidden until file loaded)
        self._body_widget = QWidget()
        self._body_widget.hide()
        body = QHBoxLayout(self._body_widget)
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        # Left: thumbnails (conditionally shown)
        self._thumb_panel = ThumbnailPanel()
        body.addWidget(self._thumb_panel)

        self._sep_left = QFrame()
        self._sep_left.setFixedWidth(1)
        self._sep_left.setStyleSheet("background: #3a3a3c;")
        body.addWidget(self._sep_left)

        # Center: viewer
        self._viewer = PdfViewer()
        self._viewer.page_changed.connect(self._on_page_changed)
        self._thumb_panel.page_clicked.connect(self._viewer.scroll_to_page)
        body.addWidget(self._viewer, 1)

        # Separator
        sep_right = QFrame()
        sep_right.setFixedWidth(1)
        sep_right.setStyleSheet("background: #3a3a3c;")
        body.addWidget(sep_right)

        # Right: operation panels
        body.addWidget(self._panel_stack)

        main_layout.addWidget(self._body_widget, 1)

        # Check for updates
        self._update_checker = UpdateChecker()
        self._update_checker.update_available.connect(self._on_update_available)
        self._update_checker.start()

    def _on_landing_drop(self, paths):
        """Handle file drop on landing screen — load first file and show workspace."""
        if not paths:
            return
        self._show_workspace()
        path = paths[0]
        if path.lower().endswith(".pdf"):
            self.load_pdf(path)
        # If multiple PDFs dropped, auto-select Merge and populate
        if len(paths) > 1 and all(p.lower().endswith(".pdf") for p in paths):
            for i, lbl in enumerate(self._op_labels):
                if lbl == "Merge":
                    self._switch_op(i)
                    panel = self._panel_stack.widget(i).widget()
                    if hasattr(panel, '_on_files_dropped'):
                        panel._on_files_dropped(paths)
                    break
        elif self._op_buttons:
            self._switch_op(0)

    def _show_workspace(self):
        """Transition from landing to three-panel workspace."""
        if self._file_loaded:
            return
        self._file_loaded = True
        self._landing.hide()
        self._toolbar.show()
        self._file_info_bar.show()
        self._body_widget.show()

    def _switch_op(self, index):
        for i, btn in enumerate(self._op_buttons):
            btn.setChecked(i == index)
        self._panel_stack.setCurrentIndex(index)

        # Thumbnail panel always visible for easier page navigation
        self._thumb_panel.setVisible(True)
        self._sep_left.setVisible(True)

    def load_pdf(self, path):
        if not os.path.isfile(path) or not path.lower().endswith(".pdf"):
            return
        self._current_pdf = path
        if not self._file_loaded:
            self._show_workspace()
            if self._op_buttons:
                self._switch_op(0)

        self._viewer.load_pdf(path)
        self._thumb_panel.load_pdf(path)

        # Update file info bar
        try:
            reader = PdfReader(path)
            pages = len(reader.pages)
        except Exception:
            pages = self._viewer.get_total_pages()
        size = os.path.getsize(path)
        self._file_info_bar.update_info(Path(path).name, pages, size)

    def _on_page_changed(self, current, total):
        self._thumb_panel.set_current_page(current)

    def _on_update_available(self, latest_version, download_url, sig_url):
        if self._update_banner is not None:
            return
        self._update_banner = UpdateBanner(self, latest_version, download_url, sig_url)
        self._banner_container.addWidget(self._update_banner)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    _cleanup_orphaned_temp()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Dark palette
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(28, 28, 30))
    palette.setColor(QPalette.WindowText, QColor(242, 242, 247))
    palette.setColor(QPalette.Base, QColor(37, 37, 39))
    palette.setColor(QPalette.AlternateBase, QColor(44, 44, 46))
    palette.setColor(QPalette.Text, QColor(242, 242, 247))
    palette.setColor(QPalette.Button, QColor(58, 58, 60))
    palette.setColor(QPalette.ButtonText, QColor(242, 242, 247))
    palette.setColor(QPalette.Highlight, QColor(59, 130, 246))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    palette.setColor(QPalette.ToolTipBase, QColor(37, 37, 39))
    palette.setColor(QPalette.ToolTipText, QColor(242, 242, 247))
    palette.setColor(QPalette.Link, QColor(59, 130, 246))
    app.setPalette(palette)

    app.setStyleSheet(DARK_STYLE)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
