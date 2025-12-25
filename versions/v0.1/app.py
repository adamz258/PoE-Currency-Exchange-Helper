import faulthandler
import logging
import os
import re
import sys
import threading
from dataclasses import dataclass
import json
from typing import Optional, Tuple

from PySide6 import QtCore, QtGui, QtWidgets
import mss
from PIL import Image, ImageFilter, ImageOps
import pytesseract
from pytesseract import TesseractNotFoundError


def _enable_dpi_awareness() -> None:
    return


_crash_log = None


def _configure_logging() -> None:
    global _crash_log
    try:
        _crash_log = open("crash.log", "a")
        faulthandler.enable(_crash_log)
    except Exception:
        pass
    logging.basicConfig(
        filename="app.log",
        filemode="a",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    def _excepthook(exc_type, exc_value, exc_traceback):
        logging.exception("Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))
        sys.__excepthook__(exc_type, exc_value, exc_traceback)

    sys.excepthook = _excepthook

    if hasattr(threading, "excepthook"):
        def _thread_hook(args):
            logging.exception("Unhandled thread exception", exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        threading.excepthook = _thread_hook


@dataclass(frozen=True)
class Region:
    left: int
    top: int
    width: int
    height: int

    def to_mss(self) -> dict:
        return {"left": self.left, "top": self.top, "width": self.width, "height": self.height}

    def to_display(self) -> str:
        return f"{self.width}x{self.height} @ {self.left},{self.top}"


@dataclass(frozen=True)
class ParseResult:
    items: Optional[int]
    listing_price: Optional[int]
    ratio_num: Optional[int]
    ratio_den: Optional[int]
    raw_text: str
    left_value: Optional[int] = None
    right_value: Optional[int] = None

    def key(self) -> tuple:
        return (self.ratio_num, self.ratio_den, self.left_value, self.right_value)

    @property
    def has_ratio(self) -> bool:
        return self.ratio_num is not None and self.ratio_den is not None and self.ratio_den != 0


class OcrParser:
    ratio_re = re.compile(r"(\d{1,6})\s*[:/]\s*(\d{1,6})", re.IGNORECASE)
    items_re = re.compile(r"(\d{1,6})\s*items?", re.IGNORECASE)
    price_re = re.compile(r"(\d{1,8})\s*(?:orbs?|orb)", re.IGNORECASE)
    number_re = re.compile(r"(\d{1,8})")

    def parse(self, text: str) -> ParseResult:
        raw_text = (text or "").strip()
        if not raw_text:
            return ParseResult(None, None, None, None, raw_text)

        lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
        ratio_num, ratio_den = self._find_ratio(lines)
        items = self._find_items(lines)
        price = self._find_price(lines)

        if (items is None or price is None) and ratio_num and ratio_den:
            numbers = self._extract_numbers(raw_text)
            for value in self._extract_ratio_numbers(raw_text):
                self._remove_once(numbers, value)
            if numbers:
                if items is None:
                    items = min(numbers)
                if price is None:
                    price = max(numbers)

        return ParseResult(items, price, ratio_num, ratio_den, raw_text)

    def _find_ratio(self, lines) -> Tuple[Optional[int], Optional[int]]:
        for i, line in enumerate(lines):
            if "ratio" not in line.lower():
                continue
            match = self.ratio_re.search(line)
            if not match and i + 1 < len(lines):
                match = self.ratio_re.search(lines[i + 1])
            if match:
                return int(match.group(1)), int(match.group(2))

        ratios = self._collect_ratios(" ".join(lines))
        if ratios:
            return max(ratios, key=lambda pair: (pair[0], pair[1]))
        return None, None

    def _find_items(self, lines) -> Optional[int]:
        for line in lines:
            match = self.items_re.search(line)
            if match:
                return int(match.group(1))
        return self._find_number_near_keyword(lines, "item")

    def _find_price(self, lines) -> Optional[int]:
        for line in lines:
            match = self.price_re.search(line)
            if match:
                return int(match.group(1))
        return self._find_number_near_keyword(lines, "price")

    def _find_number_near_keyword(self, lines, keyword: str) -> Optional[int]:
        for i, line in enumerate(lines):
            if keyword not in line.lower():
                continue
            match = self.number_re.search(line)
            if not match and i + 1 < len(lines):
                match = self.number_re.search(lines[i + 1])
            if match:
                return int(match.group(1))
        return None

    def _extract_numbers(self, text: str) -> list[int]:
        return [int(match.group(1)) for match in self.number_re.finditer(text)]

    def _collect_ratios(self, text: str) -> list[Tuple[int, int]]:
        results = []
        for match in self.ratio_re.finditer(text):
            results.append((int(match.group(1)), int(match.group(2))))
        return results

    def _extract_ratio_numbers(self, text: str) -> list[int]:
        values = []
        for num, den in self._collect_ratios(text):
            values.extend([num, den])
        return values

    @staticmethod
    def _remove_once(values: list[int], target: int) -> None:
        try:
            values.remove(target)
        except ValueError:
            return


class RegionSelector(QtWidgets.QWidget):
    region_selected = QtCore.Signal(object)
    canceled = QtCore.Signal()

    def __init__(self):
        super().__init__()
        self._start = None
        self._end = None
        self._virtual_rect = QtGui.QGuiApplication.primaryScreen().virtualGeometry()

        self.setWindowFlags(
            QtCore.Qt.WindowType.FramelessWindowHint
            | QtCore.Qt.WindowType.WindowStaysOnTopHint
            | QtCore.Qt.WindowType.Tool
        )
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setCursor(QtCore.Qt.CursorShape.CrossCursor)
        self.setGeometry(self._virtual_rect)

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() != QtCore.Qt.MouseButton.LeftButton:
            return
        self._start = event.position()
        self._end = None
        self.update()

    def mouseMoveEvent(self, event: QtGui.QMouseEvent) -> None:
        if self._start is None:
            return
        self._end = event.position()
        self.update()

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent) -> None:
        if self._start is None or self._end is None:
            self.canceled.emit()
            self.close()
            return
        rect = self._normalize_rect(self._start, self._end)
        self._start = None
        self._end = None
        if rect.width() < 10 or rect.height() < 10:
            self.canceled.emit()
            self.close()
            return

        global_left = self._virtual_rect.left() + rect.left()
        global_top = self._virtual_rect.top() + rect.top()
        center_point = QtCore.QPoint(int(global_left + rect.width() / 2), int(global_top + rect.height() / 2))
        screen = QtGui.QGuiApplication.screenAt(center_point) or QtGui.QGuiApplication.primaryScreen()
        scale = float(screen.devicePixelRatio()) if screen else 1.0

        region = Region(
            int(global_left * scale),
            int(global_top * scale),
            max(1, int(rect.width() * scale)),
            max(1, int(rect.height() * scale)),
        )
        self.region_selected.emit(region)
        self.close()

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:
        if event.key() == QtCore.Qt.Key.Key_Escape:
            self.canceled.emit()
            self.close()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        overlay = QtGui.QColor(0, 0, 0, 120)
        painter.fillRect(self.rect(), overlay)

        if self._start and self._end:
            rect = self._normalize_rect(self._start, self._end)
            fill = QtGui.QColor(242, 179, 90, 60)
            stroke = QtGui.QPen(QtGui.QColor(242, 179, 90), 2)
            painter.setPen(stroke)
            painter.setBrush(fill)
            painter.drawRect(rect)

        hint = "Drag to select OCR region. Press Esc to cancel."
        painter.setPen(QtGui.QColor(246, 242, 233))
        painter.setFont(QtGui.QFont("Segoe UI", 11, QtGui.QFont.Weight.Medium))
        painter.drawText(20, 30, hint)

    @staticmethod
    def _normalize_rect(start: QtCore.QPointF, end: QtCore.QPointF) -> QtCore.QRectF:
        left = min(start.x(), end.x())
        top = min(start.y(), end.y())
        right = max(start.x(), end.x())
        bottom = max(start.y(), end.y())
        return QtCore.QRectF(QtCore.QPointF(left, top), QtCore.QPointF(right, bottom))


class MainWindow(QtWidgets.QMainWindow):
    result_ready = QtCore.Signal(object)
    error_ready = QtCore.Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Currency Exchange Helper")
        self.resize(1120, 720)

        self._parser = OcrParser()
        self._ratio_region: Optional[Region] = None
        self._left_region: Optional[Region] = None
        self._right_region: Optional[Region] = None
        self._busy = False
        self._paused = False
        self._locked = False
        self._selector: Optional[RegionSelector] = None
        self._select_target: Optional[str] = None
        self._last_good: Optional[ParseResult] = None
        self._bad_streak = 0
        self._pending_key: Optional[tuple] = None
        self._pending_count = 0
        self._swap_sides = False

        self._tesseract_config = "--psm 6"
        self._ratio_configs = [
            "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789:/ -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
            "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789:/ -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
        ]
        self._box_configs = [
            "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789",
            "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789",
            "--oem 3 --psm 10 -c tessedit_char_whitelist=0123456789",
        ]
        self._apply_tesseract_path()

        self._build_ui()
        self._apply_styles()

        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(1200)
        self._timer.timeout.connect(self._tick)

        self.result_ready.connect(self._handle_result)
        self.error_ready.connect(self._handle_error)

        if not self._tesseract_available():
            self._set_status("Tesseract OCR not found. Install it or set TESSERACT_PATH.")
            logging.warning("Tesseract OCR not available on PATH or default install paths.")

        self._load_regions()
        self._update_region_display()
        if self._regions_ready():
            self._timer.start()

    def _apply_tesseract_path(self) -> None:
        path = os.environ.get("TESSERACT_PATH")
        candidates = [
            path,
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]

        for candidate in candidates:
            if candidate and os.path.exists(candidate):
                pytesseract.pytesseract.tesseract_cmd = candidate
                return

    @staticmethod
    def _tesseract_available() -> bool:
        try:
            pytesseract.get_tesseract_version()
            return True
        except Exception:
            return False

    def _build_ui(self) -> None:
        central = QtWidgets.QWidget()
        grid = QtWidgets.QGridLayout(central)
        grid.setSpacing(18)
        grid.setContentsMargins(24, 24, 24, 24)
        grid.setColumnStretch(0, 3)
        grid.setColumnStretch(1, 2)

        self._title_panel = self._panel()
        title_layout = QtWidgets.QVBoxLayout(self._title_panel)
        self._title_label = QtWidgets.QLabel("Currency Exchange Helper")
        self._title_label.setObjectName("titleLabel")
        self._subtitle_label = QtWidgets.QLabel(
            "Reads the market ratio and listing price from your screen to suggest the correct ask."
        )
        self._subtitle_label.setWordWrap(True)
        self._subtitle_label.setObjectName("mutedText")
        title_layout.addWidget(self._title_label)
        title_layout.addWidget(self._subtitle_label)

        status_row = QtWidgets.QHBoxLayout()
        self._status_pill = QtWidgets.QLabel("LIVE OCR")
        self._status_pill.setObjectName("statusPill")
        self._status_label = QtWidgets.QLabel("Pick a region to start OCR.")
        self._status_label.setObjectName("mutedText")
        status_row.addWidget(self._status_pill)
        status_row.addWidget(self._status_label, 1)
        title_layout.addLayout(status_row)

        self._ratio_panel = self._panel()
        ratio_layout = QtWidgets.QVBoxLayout(self._ratio_panel)
        ratio_title = QtWidgets.QLabel("Expected Ratio")
        ratio_title.setObjectName("sectionTitle")
        self._ratio_value = QtWidgets.QLabel("--")
        self._ratio_value.setObjectName("ratioValue")
        ratio_sub = QtWidgets.QLabel("Market ratio detected")
        ratio_sub.setObjectName("mutedText")
        self._last_update = QtWidgets.QLabel("Last update: --")
        self._last_update.setObjectName("mutedText")
        ratio_layout.addWidget(ratio_title)
        ratio_layout.addWidget(self._ratio_value)
        ratio_layout.addWidget(ratio_sub)
        ratio_layout.addWidget(self._last_update)

        self._capture_panel = self._panel()
        capture_layout = QtWidgets.QVBoxLayout(self._capture_panel)
        capture_title = QtWidgets.QLabel("Screen Capture")
        capture_title.setObjectName("sectionTitle")
        capture_layout.addWidget(capture_title)

        stats_grid = QtWidgets.QGridLayout()
        self._items_value = self._stat_block("Items listed", "--")
        self._listing_value = self._stat_block("Listing price", "--")
        self._confidence_value = self._stat_block("Confidence", "--")
        self._region_value = self._stat_block("Regions", "Not set", large=False)
        stats_grid.addLayout(self._items_value[0], 0, 0)
        stats_grid.addLayout(self._listing_value[0], 0, 1)
        stats_grid.addLayout(self._confidence_value[0], 1, 0)
        stats_grid.addLayout(self._region_value[0], 1, 1)
        capture_layout.addLayout(stats_grid)

        button_row = QtWidgets.QHBoxLayout()
        self._pick_ratio_button = QtWidgets.QPushButton("Pick Ratio")
        self._pick_ratio_button.clicked.connect(lambda: self._pick_region("ratio"))
        self._pick_left_button = QtWidgets.QPushButton("Pick Left Box")
        self._pick_left_button.clicked.connect(lambda: self._pick_region("left"))
        self._pick_right_button = QtWidgets.QPushButton("Pick Right Box")
        self._pick_right_button.clicked.connect(lambda: self._pick_region("right"))
        self._lock_button = QtWidgets.QPushButton("Lock Regions")
        self._lock_button.clicked.connect(self._toggle_lock)
        self._pause_button = QtWidgets.QPushButton("Pause OCR")
        self._pause_button.clicked.connect(self._toggle_pause)
        button_row.addWidget(self._pick_ratio_button)
        button_row.addWidget(self._pick_left_button)
        button_row.addWidget(self._pick_right_button)
        button_row.addWidget(self._lock_button)
        button_row.addWidget(self._pause_button)
        button_row.addStretch(1)
        capture_layout.addLayout(button_row)

        swap_row = QtWidgets.QHBoxLayout()
        self._swap_checkbox = QtWidgets.QCheckBox("Swap Left/Right")
        self._swap_checkbox.stateChanged.connect(self._toggle_swap)
        swap_row.addWidget(self._swap_checkbox)
        swap_row.addStretch(1)
        capture_layout.addLayout(swap_row)

        text_label = QtWidgets.QLabel("Detected Text")
        text_label.setObjectName("sectionTitle")
        capture_layout.addWidget(text_label)
        self._raw_text = QtWidgets.QPlainTextEdit()
        self._raw_text.setReadOnly(True)
        self._raw_text.setObjectName("rawText")
        capture_layout.addWidget(self._raw_text)

        self._suggest_panel = self._panel()
        suggest_layout = QtWidgets.QVBoxLayout(self._suggest_panel)
        suggest_title = QtWidgets.QLabel("Suggested Ask")
        suggest_title.setObjectName("sectionTitle")
        suggest_layout.addWidget(suggest_title)

        highlight = QtWidgets.QFrame()
        highlight.setObjectName("highlight")
        highlight_layout = QtWidgets.QVBoxLayout(highlight)
        highlight_label = QtWidgets.QLabel("Recommended price")
        highlight_label.setObjectName("mutedText")
        self._recommended_value = QtWidgets.QLabel("--")
        self._recommended_value.setObjectName("recommendedValue")
        highlight_layout.addWidget(highlight_label)
        highlight_layout.addWidget(self._recommended_value)
        suggest_layout.addWidget(highlight)

        self._ratio_summary = self._stat_row("Market ratio", "--")
        self._items_summary = self._stat_row("Items to sell", "--")
        suggest_layout.addLayout(self._ratio_summary[0])
        suggest_layout.addLayout(self._items_summary[0])

        grid.addWidget(self._title_panel, 0, 0)
        grid.addWidget(self._ratio_panel, 0, 1)
        grid.addWidget(self._capture_panel, 1, 0)
        grid.addWidget(self._suggest_panel, 1, 1)

        self.setCentralWidget(central)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background: #0b151c;
                color: #f6f2e9;
                font-family: "Segoe UI";
            }
            QLabel#titleLabel {
                font-size: 26px;
                font-weight: 700;
            }
            QLabel#sectionTitle {
                font-size: 12px;
                text-transform: uppercase;
                letter-spacing: 1px;
                color: #c9c3b8;
            }
            QLabel#mutedText {
                color: #c9c3b8;
            }
            QLabel#ratioValue {
                font-size: 34px;
                font-weight: 600;
            }
            QLabel#recommendedValue {
                font-size: 28px;
                font-weight: 700;
                color: #f2b35a;
            }
            QLabel#statusPill {
                background: #223843;
                color: #7cd0d4;
                border: 1px solid #7cd0d4;
                border-radius: 12px;
                padding: 4px 10px;
                font-size: 11px;
                font-weight: 700;
            }
            QFrame[panel="true"] {
                background: #141b20;
                border: 1px solid #2a343c;
                border-radius: 16px;
            }
            QFrame#highlight {
                background: #2a2012;
                border: 1px solid #f2b35a;
                border-radius: 12px;
                padding: 12px;
            }
            QPlainTextEdit#rawText {
                background: #10161b;
                border: 1px solid #2b353d;
                color: #c9c3b8;
            }
            QPushButton {
                background: #1c252b;
                border: 1px solid #2e3a42;
                padding: 8px 12px;
                border-radius: 6px;
            }
            QPushButton:hover {
                border-color: #7cd0d4;
            }
            """
        )

    def _panel(self) -> QtWidgets.QFrame:
        panel = QtWidgets.QFrame()
        panel.setProperty("panel", True)
        panel.setContentsMargins(16, 16, 16, 16)
        return panel

    def _stat_block(self, label: str, value: str, large: bool = True):
        layout = QtWidgets.QVBoxLayout()
        label_widget = QtWidgets.QLabel(label)
        label_widget.setObjectName("mutedText")
        value_widget = QtWidgets.QLabel(value)
        value_widget.setFont(QtGui.QFont("Segoe UI", 20 if large else 14, QtGui.QFont.Weight.Medium))
        if not large:
            value_widget.setWordWrap(True)
        layout.addWidget(label_widget)
        layout.addWidget(value_widget)
        return layout, value_widget

    def _stat_row(self, label: str, value: str):
        layout = QtWidgets.QHBoxLayout()
        label_widget = QtWidgets.QLabel(label)
        label_widget.setObjectName("mutedText")
        value_widget = QtWidgets.QLabel(value)
        value_widget.setFont(QtGui.QFont("Segoe UI", 18, QtGui.QFont.Weight.Medium))
        layout.addWidget(label_widget)
        layout.addStretch(1)
        layout.addWidget(value_widget)
        return layout, value_widget

    def _pick_region(self, target: str) -> None:
        if self._locked:
            self._set_status("Regions are locked. Unlock to reselect.")
            return
        self._select_target = target
        self._selector = RegionSelector()
        self._selector.region_selected.connect(self._set_region)
        self._selector.canceled.connect(self._handle_selection_canceled)
        self._selector.show()
        self._selector.raise_()
        self._selector.activateWindow()

    def _set_region(self, region: Region) -> None:
        if self._select_target == "ratio":
            self._ratio_region = region
            self._set_status("Ratio region set.")
        elif self._select_target == "left":
            self._left_region = region
            self._set_status("Left box region set.")
        elif self._select_target == "right":
            self._right_region = region
            self._set_status("Right box region set.")
        else:
            self._set_status("Region set.")
        self._selector = None
        self._select_target = None
        self._last_good = None
        self._pending_key = None
        self._pending_count = 0
        self._save_regions()
        self._update_region_display()
        if self._regions_ready():
            self._set_status("Regions selected. Running OCR...")
            if not self._timer.isActive():
                self._timer.start()
            QtCore.QTimer.singleShot(100, self._tick)

    def _handle_selection_canceled(self) -> None:
        self._selector = None
        self._select_target = None
        self._set_status("Region selection canceled.")

    def _toggle_pause(self) -> None:
        self._paused = not self._paused
        self._pause_button.setText("Resume OCR" if self._paused else "Pause OCR")
        self._set_status("OCR paused." if self._paused else "OCR resumed.")

    def _toggle_lock(self) -> None:
        self._locked = not self._locked
        self._lock_button.setText("Unlock Regions" if self._locked else "Lock Regions")
        self._set_status("Regions locked." if self._locked else "Regions unlocked.")

    def _toggle_swap(self, state: int) -> None:
        self._swap_sides = state == QtCore.Qt.CheckState.Checked
        self._last_good = None
        self._pending_key = None
        self._pending_count = 0
        self._save_regions()
        self._set_status("Swapped left/right inputs." if self._swap_sides else "Using left/right inputs.")
        if self._regions_ready():
            QtCore.QTimer.singleShot(100, self._tick)

    def _tick(self) -> None:
        if self._paused or self._busy:
            return
        if not self._regions_ready():
            self._set_status(f"Set regions: {', '.join(self._missing_regions())}.")
            return
        self._busy = True
        thread = threading.Thread(
            target=self._run_ocr_task,
            args=(self._ratio_region, self._left_region, self._right_region),
            daemon=True,
        )
        thread.start()

    def _run_ocr_task(self, ratio_region: Region, left_region: Region, right_region: Region) -> None:
        try:
            raw_text, result = self._capture_and_ocr(ratio_region, left_region, right_region)
            self.result_ready.emit(result)
        except TesseractNotFoundError:
            self.error_ready.emit("Tesseract OCR not found. Install it or add to PATH.")
        except Exception as exc:
            self.error_ready.emit(f"OCR error: {exc}")

    def _capture_and_ocr(
        self,
        ratio_region: Region,
        left_region: Region,
        right_region: Region,
    ) -> tuple[str, ParseResult]:
        with mss.mss() as sct:
            ratio_image = self._grab_region(sct, ratio_region)
            left_image = self._grab_region(sct, left_region)
            right_image = self._grab_region(sct, right_region)

        ratio_num, ratio_den, ratio_text = self._ocr_ratio(ratio_image)
        left_value, left_text = self._ocr_box_value(left_image)
        right_value, right_text = self._ocr_box_value(right_image)

        items, price = self._assign_items_price(left_value, right_value)
        raw_text = f"RATIO: {ratio_text}\nLEFT: {left_text}\nRIGHT: {right_text}".strip()
        result = ParseResult(items, price, ratio_num, ratio_den, raw_text, left_value, right_value)
        return raw_text, result

    @staticmethod
    def _grab_region(sct: mss.mss, region: Region) -> Image.Image:
        bounds = sct.monitors[0]
        left = max(bounds["left"], region.left)
        top = max(bounds["top"], region.top)
        right = min(bounds["left"] + bounds["width"], region.left + region.width)
        bottom = min(bounds["top"] + bounds["height"], region.top + region.height)

        width = max(1, right - left)
        height = max(1, bottom - top)
        monitor = {"left": left, "top": top, "width": width, "height": height}
        screenshot = sct.grab(monitor)
        return Image.frombytes("RGB", screenshot.size, screenshot.rgb)

    @staticmethod
    def _preprocess_variants(image: Image.Image, scale: int = 2, threshold_value: int = 160) -> list[Image.Image]:
        base = ImageOps.grayscale(image)
        base = ImageOps.autocontrast(base)
        width, height = base.size
        scaled = base.resize((width * scale, height * scale), resample=Image.Resampling.LANCZOS)
        sharpened = scaled.filter(ImageFilter.UnsharpMask(radius=1, percent=180, threshold=3))
        median = sharpened.filter(ImageFilter.MedianFilter(size=3))
        threshold = median.point(lambda p: 255 if p > threshold_value else 0)
        inverted = ImageOps.invert(threshold)
        return [sharpened, median, threshold, inverted]

    def _ocr_ratio(self, image: Image.Image) -> tuple[Optional[int], Optional[int], str]:
        candidates = []
        best_text = ""
        for threshold in (170, 150, 130):
            for processed in self._preprocess_variants(image, scale=4, threshold_value=threshold):
                for config in self._ratio_configs:
                    text = pytesseract.image_to_string(processed, config=config, lang="eng").strip()
                    best_text = text or best_text
                    ratios = self._parser._collect_ratios(text)
                    for ratio in ratios:
                        score = self._score_ratio(ratio[0], ratio[1])
                        candidates.append((score, ratio[0], ratio[1], text))

        if candidates:
            best = max(candidates, key=lambda item: item[0])
            return best[1], best[2], f"{best[1]}:{best[2]}"

        ratios = self._parser._collect_ratios(best_text)
        if ratios:
            best_ratio = max(ratios, key=lambda pair: (pair[1], pair[0]))
            return best_ratio[0], best_ratio[1], f"{best_ratio[0]}:{best_ratio[1]}"
        return None, None, best_text

    @staticmethod
    def _score_ratio(num: int, den: int) -> int:
        digits = len(str(num)) + len(str(den))
        score = digits * 10
        if num > 0 and den > 0:
            score += 4
        if num <= 10000 and den <= 10000:
            score += 2
        if num != den:
            score += 1
        return score

    def _ocr_box_value(self, image: Image.Image) -> tuple[Optional[int], str]:
        best_text = ""
        for threshold in (150, 120):
            for processed in self._preprocess_variants(image, scale=3, threshold_value=threshold):
                for config in self._box_configs:
                    text = pytesseract.image_to_string(processed, config=config, lang="eng").strip()
                    best_text = text or best_text
                    value = self._extract_best_int(text)
                    if value is not None:
                        return value, text
        return None, best_text

    @staticmethod
    def _extract_best_int(text: str) -> Optional[int]:
        matches = re.findall(r"\d{1,8}", text)
        if not matches:
            return None
        values = [int(match) for match in matches]
        return max(values)

    def _assign_items_price(
        self,
        left_value: Optional[int],
        right_value: Optional[int],
    ) -> tuple[Optional[int], Optional[int]]:
        if left_value is None and right_value is None:
            return None, None
        if left_value is None:
            return (right_value, None) if self._swap_sides else (None, right_value)
        if right_value is None:
            return (None, left_value) if self._swap_sides else (left_value, None)

        if self._swap_sides:
            return right_value, left_value
        return left_value, right_value

    def _handle_result(self, result: ParseResult) -> None:
        self._raw_text.setPlainText(result.raw_text or "")

        is_valid = self._is_valid_result(result)
        if is_valid:
            key = result.key()
            if key == self._pending_key:
                self._pending_count += 1
            else:
                self._pending_key = key
                self._pending_count = 1
            if self._pending_count >= 2:
                self._last_good = result
                self._bad_streak = 0
        else:
            self._pending_key = None
            self._pending_count = 0
            self._bad_streak += 1

        display = self._last_good if self._last_good else result
        if is_valid and self._pending_count >= 2:
            display = result

        self._items_value[1].setText(str(display.items) if display.items is not None else "--")
        self._listing_value[1].setText(
            f"{display.listing_price} orbs" if display.listing_price is not None else "--"
        )
        self._ratio_value.setText(
            f"{display.ratio_num} : {display.ratio_den}" if display.has_ratio else "--"
        )
        self._ratio_summary[1].setText(self._ratio_value.text())
        self._items_summary[1].setText(self._items_value[1].text())

        if display.has_ratio and display.items is not None:
            recommended = round(display.items * (display.ratio_num / display.ratio_den))
            self._recommended_value.setText(f"{recommended} orbs")
        else:
            self._recommended_value.setText("--")

        confidence = self._compute_confidence(display)
        self._confidence_value[1].setText(f"{confidence}%")
        self._last_update.setText(f"Last update: {QtCore.QTime.currentTime().toString('HH:mm:ss')}")
        if not is_valid and self._last_good and self._bad_streak <= 3:
            self._set_status("OCR unstable; holding last good values.")
        else:
            self._set_status(self._build_status(display))
        logging.info(
            "OCR updated: ratio=%s:%s items=%s price=%s left=%s right=%s",
            display.ratio_num,
            display.ratio_den,
            display.items,
            display.listing_price,
            display.left_value,
            display.right_value,
        )
        self._busy = False

    def _handle_error(self, message: str) -> None:
        self._set_status(message)
        logging.error(message)
        self._busy = False

    def _set_status(self, message: str) -> None:
        self._status_label.setText(message)

    @staticmethod
    def _is_valid_result(result: ParseResult) -> bool:
        if not result.has_ratio:
            return False
        if result.left_value is None or result.right_value is None:
            return False
        if result.left_value <= 0 or result.right_value <= 0:
            return False
        return True

    def _regions_ready(self) -> bool:
        return self._ratio_region is not None and self._left_region is not None and self._right_region is not None

    def _missing_regions(self) -> list[str]:
        missing = []
        if self._ratio_region is None:
            missing.append("ratio")
        if self._left_region is None:
            missing.append("left")
        if self._right_region is None:
            missing.append("right")
        return missing

    def _update_region_display(self) -> None:
        def region_text(name: str, region: Optional[Region]) -> str:
            return f"{name}: {region.to_display()}" if region else f"{name}: not set"

        lines = [
            region_text("Ratio", self._ratio_region),
            region_text("Left", self._left_region),
            region_text("Right", self._right_region),
        ]
        self._region_value[1].setText("\n".join(lines))

    def _save_regions(self) -> None:
        data = {
            "ratio_region": self._region_to_dict(self._ratio_region),
            "left_region": self._region_to_dict(self._left_region),
            "right_region": self._region_to_dict(self._right_region),
            "swap_sides": self._swap_sides,
        }
        try:
            with open("config.json", "w", encoding="utf-8") as handle:
                json.dump(data, handle, indent=2)
        except Exception as exc:
            logging.error("Failed to save config: %s", exc)

    def _load_regions(self) -> None:
        if not os.path.exists("config.json"):
            return
        try:
            with open("config.json", "r", encoding="utf-8") as handle:
                data = json.load(handle)
            self._ratio_region = self._region_from_dict(data.get("ratio_region"))
            self._left_region = self._region_from_dict(data.get("left_region"))
            self._right_region = self._region_from_dict(data.get("right_region"))
            self._swap_sides = bool(data.get("swap_sides", False))
            if hasattr(self, "_swap_checkbox"):
                self._swap_checkbox.setChecked(self._swap_sides)
        except Exception as exc:
            logging.error("Failed to load config: %s", exc)

    @staticmethod
    def _region_to_dict(region: Optional[Region]) -> Optional[dict]:
        if region is None:
            return None
        return {
            "left": region.left,
            "top": region.top,
            "width": region.width,
            "height": region.height,
        }

    @staticmethod
    def _region_from_dict(payload: Optional[dict]) -> Optional[Region]:
        if not payload:
            return None
        return Region(
            int(payload.get("left", 0)),
            int(payload.get("top", 0)),
            int(payload.get("width", 0)),
            int(payload.get("height", 0)),
        )

    @staticmethod
    def _compute_confidence(result: ParseResult) -> int:
        score = 0
        if result.has_ratio:
            score += 50
        if result.items is not None:
            score += 25
        if result.listing_price is not None:
            score += 25
        return score

    @staticmethod
    def _build_status(result: ParseResult) -> str:
        if not result.has_ratio:
            return "Looking for market ratio..."
        if result.items is None:
            return "Looking for item count..."
        if result.listing_price is None:
            return "Looking for listing price..."
        return "OCR updated."


def main() -> None:
    _configure_logging()
    _enable_dpi_awareness()
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("Currency Exchange Helper")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
