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

_win32gui = None
_win32con = None

if sys.platform == "win32":
    try:
        import win32gui as _win32gui  # type: ignore
        import win32con as _win32con  # type: ignore
    except Exception:
        _win32gui = None
        _win32con = None


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
    ratio_num: Optional[float]
    ratio_den: Optional[float]
    raw_text: str
    left_value: Optional[int] = None
    right_value: Optional[int] = None

    def key(self) -> tuple:
        ratio_num = round(self.ratio_num, 2) if self.ratio_num is not None else None
        ratio_den = round(self.ratio_den, 2) if self.ratio_den is not None else None
        return (ratio_num, ratio_den, self.left_value, self.right_value)

    @property
    def has_ratio(self) -> bool:
        return (
            self.ratio_num is not None
            and self.ratio_den is not None
            and abs(self.ratio_den) > 1e-9
        )


class OcrParser:
    ratio_re = re.compile(r"(\d{1,6}(?:[.,]\d{1,2})?)\s*[:/]\s*(\d{1,6}(?:[.,]\d{1,2})?)", re.IGNORECASE)
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

    def _find_ratio(self, lines) -> Tuple[Optional[float], Optional[float]]:
        for i, line in enumerate(lines):
            if "ratio" not in line.lower():
                continue
            match = self.ratio_re.search(line)
            if not match and i + 1 < len(lines):
                match = self.ratio_re.search(lines[i + 1])
            if match:
                try:
                    return float(match.group(1).replace(",", ".")), float(match.group(2).replace(",", "."))
                except ValueError:
                    continue

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

    def _collect_ratios(self, text: str) -> list[Tuple[float, float]]:
        results = []
        for match in self.ratio_re.finditer(text):
            num = match.group(1).replace(",", ".")
            den = match.group(2).replace(",", ".")
            try:
                results.append((float(num), float(den)))
            except ValueError:
                continue
        return results

    def _extract_ratio_numbers(self, text: str) -> list[int]:
        values = []
        for num, den in self._collect_ratios(text):
            values.extend([int(num), int(den)])
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
        self._topmost_enabled = False
        self._closing = False
        self._minimal_mode = False
        self._full_geometry: Optional[QtCore.QRect] = None
        self._full_min_size: Optional[QtCore.QSize] = None
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
        self._calc_mode = "auto"
        self._auto_direction = "from_right"
        self._last_inputs: Optional[tuple[int, int]] = None

        self._ratio_configs = [
            "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789:.,/ -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
            "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789:.,/ -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
            "--oem 1 --psm 7 -c tessedit_char_whitelist=0123456789:.,/ -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
        ]
        self._ratio_digit_configs = [
            "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789., -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
            "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789., -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
            "--oem 1 --psm 7 -c tessedit_char_whitelist=0123456789., -c classify_bln_numeric_mode=1 -c load_system_dawg=0 -c load_freq_dawg=0",
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
        self._timer.setInterval(600)
        self._timer.timeout.connect(self._tick)

        self._topmost_timer = QtCore.QTimer(self)
        self._topmost_timer.setInterval(500)
        self._topmost_timer.timeout.connect(self._reassert_topmost)

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
        self._stack = QtWidgets.QStackedWidget()
        self._full_root = QtWidgets.QWidget()
        grid = QtWidgets.QGridLayout(self._full_root)
        grid.setSpacing(18)
        grid.setContentsMargins(24, 24, 24, 24)
        grid.setColumnStretch(0, 3)
        grid.setColumnStretch(1, 2)

        self._title_panel = self._panel()
        title_layout = QtWidgets.QVBoxLayout(self._title_panel)
        self._title_label = QtWidgets.QLabel("Currency Exchange Helper")
        self._title_label.setObjectName("titleLabel")
        self._subtitle_label = QtWidgets.QLabel(
            "Reads the market ratio and your inputs to suggest the correct side to match the ratio."
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
        self._left_value = self._stat_block("I want", "--")
        self._right_value = self._stat_block("I have", "--")
        self._confidence_value = self._stat_block("Confidence", "--")
        self._region_value = self._stat_block("Regions", "Not set", large=False)
        stats_grid.addLayout(self._left_value[0], 0, 0)
        stats_grid.addLayout(self._right_value[0], 0, 1)
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

        mode_row = QtWidgets.QHBoxLayout()
        mode_label = QtWidgets.QLabel("Mode")
        mode_label.setObjectName("mutedText")
        self._mode_combo = QtWidgets.QComboBox()
        self._mode_combo.addItems(
            [
                "Auto",
                "I have -> calc I want",
                "I want -> calc I have",
            ]
        )
        self._mode_combo.currentIndexChanged.connect(self._change_mode)
        mode_row.addWidget(mode_label)
        mode_row.addWidget(self._mode_combo)
        mode_row.addStretch(1)
        capture_layout.addLayout(mode_row)

        top_row = QtWidgets.QHBoxLayout()
        self._top_checkbox = QtWidgets.QCheckBox("Always on top")
        self._top_checkbox.setTristate(False)
        self._top_checkbox.toggled.connect(self._toggle_topmost)
        self._minimal_checkbox = QtWidgets.QCheckBox("Minimal mode")
        self._minimal_checkbox.setTristate(False)
        self._minimal_checkbox.toggled.connect(self._toggle_minimal_mode)
        top_row.addWidget(self._top_checkbox)
        top_row.addWidget(self._minimal_checkbox)
        top_row.addStretch(1)
        capture_layout.addLayout(top_row)

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
        self._recommended_label = QtWidgets.QLabel("Recommended value")
        self._recommended_label.setObjectName("mutedText")
        self._recommended_value = QtWidgets.QLabel("--")
        self._recommended_value.setObjectName("recommendedValue")
        highlight_layout.addWidget(self._recommended_label)
        highlight_layout.addWidget(self._recommended_value)
        suggest_layout.addWidget(highlight)

        self._expected_left = self._stat_row("Expected I want", "--")
        self._expected_right = self._stat_row("Expected I have", "--")
        self._ratio_summary = self._stat_row("Market ratio", "--")
        suggest_layout.addLayout(self._expected_left[0])
        suggest_layout.addLayout(self._expected_right[0])
        suggest_layout.addLayout(self._ratio_summary[0])

        grid.addWidget(self._title_panel, 0, 0)
        grid.addWidget(self._ratio_panel, 0, 1)
        grid.addWidget(self._capture_panel, 1, 0)
        grid.addWidget(self._suggest_panel, 1, 1)

        self._stack.addWidget(self._full_root)

        self._minimal_root = QtWidgets.QWidget()
        self._minimal_root.setObjectName("minimalRoot")
        self._minimal_root.setToolTip("Double-click or press Esc to return to full view.")
        mini_layout = QtWidgets.QVBoxLayout(self._minimal_root)
        mini_layout.setContentsMargins(6, 6, 6, 6)
        mini_layout.setSpacing(6)
        mini_panel = QtWidgets.QFrame()
        mini_panel.setProperty("panel", True)
        mini_panel.setContentsMargins(8, 8, 8, 8)
        mini_panel_layout = QtWidgets.QVBoxLayout(mini_panel)
        mini_panel_layout.setSpacing(4)

        mini_ratio_label = QtWidgets.QLabel("Market ratio")
        mini_ratio_label.setObjectName("miniTitle")
        self._mini_ratio_value = QtWidgets.QLabel("--")
        self._mini_ratio_value.setObjectName("miniRatioValue")

        self._mini_recommended_label = QtWidgets.QLabel("Recommended")
        self._mini_recommended_label.setObjectName("miniTitle")
        self._mini_recommended_value = QtWidgets.QLabel("--")
        self._mini_recommended_value.setObjectName("miniRecommendedValue")

        mini_panel_layout.addWidget(mini_ratio_label)
        mini_panel_layout.addWidget(self._mini_ratio_value)
        mini_panel_layout.addSpacing(6)
        mini_panel_layout.addWidget(self._mini_recommended_label)
        mini_panel_layout.addWidget(self._mini_recommended_value)
        mini_panel_layout.addSpacing(2)

        mini_controls = QtWidgets.QHBoxLayout()
        self._mini_top_checkbox = QtWidgets.QCheckBox("Always on top")
        self._mini_top_checkbox.setTristate(False)
        self._mini_top_checkbox.setObjectName("miniCheckbox")
        self._mini_top_checkbox.toggled.connect(self._toggle_topmost)
        self._mini_minimal_checkbox = QtWidgets.QCheckBox("Minimal mode")
        self._mini_minimal_checkbox.setTristate(False)
        self._mini_minimal_checkbox.setObjectName("miniCheckbox")
        self._mini_minimal_checkbox.toggled.connect(self._toggle_minimal_mode)
        mini_controls.addWidget(self._mini_top_checkbox)
        mini_controls.addWidget(self._mini_minimal_checkbox)
        mini_controls.addStretch(1)
        mini_panel_layout.addLayout(mini_controls)

        mini_layout.addWidget(mini_panel)
        self._stack.addWidget(self._minimal_root)
        self._stack.setCurrentWidget(self._full_root)

        for widget in (
            self._minimal_root,
            mini_panel,
            self._mini_ratio_value,
            self._mini_recommended_label,
            self._mini_recommended_value,
            self._mini_top_checkbox,
            self._mini_minimal_checkbox,
        ):
            widget.installEventFilter(self)

        self.setCentralWidget(self._stack)

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
            QLabel#miniTitle {
                font-size: 10px;
                text-transform: uppercase;
                color: #c9c3b8;
            }
            QLabel#miniRatioValue {
                font-size: 16px;
                font-weight: 600;
            }
            QLabel#miniRecommendedValue {
                font-size: 16px;
                font-weight: 700;
                color: #f2b35a;
            }
            QCheckBox#miniCheckbox {
                font-size: 10px;
                color: #c9c3b8;
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
        self._last_inputs = None
        self._save_regions()
        self._set_status("Swapped left/right inputs." if self._swap_sides else "Using left/right inputs.")
        if self._regions_ready():
            QtCore.QTimer.singleShot(100, self._tick)

    def _toggle_topmost(self, state) -> None:
        sender = self.sender()
        if sender is self._top_checkbox:
            is_top = self._top_checkbox.isChecked()
        elif hasattr(self, "_mini_top_checkbox") and sender is self._mini_top_checkbox:
            is_top = self._mini_top_checkbox.isChecked()
        elif isinstance(state, bool):
            is_top = state
        else:
            is_top = state == QtCore.Qt.CheckState.Checked
        self._topmost_enabled = is_top
        logging.info("Topmost toggle: %s", is_top)
        self._apply_topmost(is_top)
        if is_top:
            self._topmost_timer.start()
        else:
            self._topmost_timer.stop()

        if hasattr(self, "_top_checkbox") and self._top_checkbox.isChecked() != is_top:
            blocked = self._top_checkbox.blockSignals(True)
            self._top_checkbox.setChecked(is_top)
            self._top_checkbox.blockSignals(blocked)
        if hasattr(self, "_mini_top_checkbox") and self._mini_top_checkbox.isChecked() != is_top:
            blocked = self._mini_top_checkbox.blockSignals(True)
            self._mini_top_checkbox.setChecked(is_top)
            self._mini_top_checkbox.blockSignals(blocked)

    def _toggle_minimal_mode(self, state) -> None:
        sender = self.sender()
        if sender is self._minimal_checkbox:
            is_minimal = self._minimal_checkbox.isChecked()
        elif hasattr(self, "_mini_minimal_checkbox") and sender is self._mini_minimal_checkbox:
            is_minimal = self._mini_minimal_checkbox.isChecked()
        elif isinstance(state, bool):
            is_minimal = state
        else:
            is_minimal = state == QtCore.Qt.CheckState.Checked
        self._set_minimal_mode(is_minimal)

    def _set_minimal_mode(self, enabled: bool) -> None:
        if enabled == self._minimal_mode:
            return
        self._minimal_mode = enabled
        if enabled:
            self._full_geometry = self.geometry()
            self._full_min_size = self.minimumSize()
            self._stack.setCurrentWidget(self._minimal_root)
            self._minimal_root.adjustSize()
            size = self._minimal_root.sizeHint()
            if size.isValid():
                self.setMinimumSize(size)
                self.resize(size)
            self.setWindowTitle("Currency Exchange Helper (Minimal)")
        else:
            self._stack.setCurrentWidget(self._full_root)
            if self._full_min_size is not None:
                self.setMinimumSize(self._full_min_size)
            if self._full_geometry is not None:
                self.setGeometry(self._full_geometry)
            else:
                self.resize(1120, 720)
            self.setWindowTitle("Currency Exchange Helper")

        if hasattr(self, "_minimal_checkbox") and self._minimal_checkbox.isChecked() != enabled:
            blocked = self._minimal_checkbox.blockSignals(True)
            self._minimal_checkbox.setChecked(enabled)
            self._minimal_checkbox.blockSignals(blocked)
        if hasattr(self, "_mini_minimal_checkbox") and self._mini_minimal_checkbox.isChecked() != enabled:
            blocked = self._mini_minimal_checkbox.blockSignals(True)
            self._mini_minimal_checkbox.setChecked(enabled)
            self._mini_minimal_checkbox.blockSignals(blocked)

    def changeEvent(self, event: QtCore.QEvent) -> None:
        if not hasattr(self, "_topmost_enabled"):
            super().changeEvent(event)
            return
        if (
            self._topmost_enabled
            and event.type() in (QtCore.QEvent.Type.WindowDeactivate, QtCore.QEvent.Type.ActivationChange)
        ):
            QtCore.QTimer.singleShot(0, self._reassert_topmost)
        super().changeEvent(event)

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:
        if self._minimal_mode and event.type() == QtCore.QEvent.Type.MouseButtonDblClick:
            self._set_minimal_mode(False)
            return True
        return super().eventFilter(obj, event)

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:
        if self._minimal_mode and event.key() == QtCore.Qt.Key.Key_Escape:
            self._set_minimal_mode(False)
            event.accept()
            return
        super().keyPressEvent(event)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._closing = True
        self._topmost_enabled = False
        if hasattr(self, "_topmost_timer"):
            self._topmost_timer.stop()
        if hasattr(self, "_timer"):
            self._timer.stop()
        super().closeEvent(event)

    def _get_hwnd(self) -> Optional[int]:
        try:
            hwnd = int(self.winId())
        except Exception:
            hwnd = 0
        if hwnd == 0 and self.windowHandle() is not None:
            try:
                hwnd = int(self.windowHandle().winId())
            except Exception:
                hwnd = 0
        return hwnd or None

    def _reassert_topmost(self) -> None:
        if self._closing or not self._topmost_enabled:
            return
        self._apply_topmost(True)

    def _apply_topmost(self, is_top: bool) -> None:
        if self._closing:
            return
        if not self.isVisible():
            self.show()
        if is_top:
            self.raise_()
        if sys.platform != "win32":
            return
        hwnd = self._get_hwnd()
        if hwnd is None:
            logging.error("Topmost failed: unable to resolve window handle.")
            return

        if _win32gui is None or _win32con is None:
            logging.error("Topmost unavailable: pywin32 not loaded.")
            return
        try:
            if not _win32gui.IsWindow(hwnd):
                logging.error("Topmost failed: invalid window handle %s.", hwnd)
                return
            insert_after = _win32con.HWND_TOPMOST if is_top else _win32con.HWND_NOTOPMOST
            flags = _win32con.SWP_NOMOVE | _win32con.SWP_NOSIZE | _win32con.SWP_SHOWWINDOW
            _win32gui.SetWindowPos(hwnd, insert_after, 0, 0, 0, 0, flags)
            logging.info("Topmost set (win32gui): %s", is_top)
        except Exception as exc:
            logging.error("Topmost win32gui failed: %s", exc)


    def _change_mode(self, index: int) -> None:
        if index == 1:
            self._calc_mode = "from_right"
        elif index == 2:
            self._calc_mode = "from_left"
        else:
            self._calc_mode = "auto"
        self._last_inputs = None
        self._save_regions()
        self._set_status(f"Mode set to {self._mode_combo.currentText()}.")
        if self._regions_ready():
            QtCore.QTimer.singleShot(100, self._tick)

    def _set_mode_combo(self, mode: str) -> None:
        if mode == "from_right":
            self._mode_combo.setCurrentIndex(1)
        elif mode == "from_left":
            self._mode_combo.setCurrentIndex(2)
        else:
            self._mode_combo.setCurrentIndex(0)

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

        ratio_num, ratio_den, ratio_text = self._ocr_ratio(ratio_image, fast=True)
        left_value, left_text = self._ocr_box_value(left_image, fast=True)
        right_value, right_text = self._ocr_box_value(right_image, fast=True)

        if ratio_num is None or ratio_den is None:
            ratio_num, ratio_den, ratio_text = self._ocr_ratio(ratio_image, fast=False)
        if left_value is None:
            left_value, left_text = self._ocr_box_value(left_image, fast=False)
        if right_value is None:
            right_value, right_text = self._ocr_box_value(right_image, fast=False)

        left_input, right_input = self._normalize_inputs(left_value, right_value)
        raw_text = f"RATIO: {ratio_text}\nLEFT: {left_text}\nRIGHT: {right_text}".strip()
        result = ParseResult(right_input, left_input, ratio_num, ratio_den, raw_text, left_input, right_input)
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
    def _preprocess_variants(
        image: Image.Image,
        scale: int = 2,
        threshold_value: int = 160,
        mode: str = "full",
    ) -> list[Image.Image]:
        base = ImageOps.grayscale(image)
        base = ImageOps.autocontrast(base)
        width, height = base.size
        scaled = base.resize((width * scale, height * scale), resample=Image.Resampling.LANCZOS)
        sharpened = scaled.filter(ImageFilter.UnsharpMask(radius=1, percent=180, threshold=3))
        median = sharpened.filter(ImageFilter.MedianFilter(size=3))
        threshold = median.point(lambda p: 255 if p > threshold_value else 0)
        inverted = ImageOps.invert(threshold)
        if mode == "fast":
            return [threshold, inverted]
        return [sharpened, median, threshold, inverted]

    @staticmethod
    def _is_binary_image(image: Image.Image) -> bool:
        extrema = image.getextrema()
        return extrema in ((0, 255), (0, 1))

    @staticmethod
    def _binarize_for_ratio(image: Image.Image) -> Image.Image:
        gray = image.convert("L")
        pixels = list(gray.getdata())
        if not pixels:
            return gray
        mean = sum(pixels) / len(pixels)
        if mean > 127:
            gray = ImageOps.invert(gray)
        return gray.point(lambda p: 255 if p > 128 else 0)

    @staticmethod
    def _auto_crop_ratio(image: Image.Image) -> Image.Image:
        gray = ImageOps.grayscale(image)
        mask = gray.point(lambda p: 255 if p > 150 else 0)
        bbox = mask.getbbox()
        if not bbox:
            return image
        left, top, right, bottom = bbox
        pad = 2
        left = max(0, left - pad)
        top = max(0, top - pad)
        right = min(image.width, right + pad)
        bottom = min(image.height, bottom + pad)
        if right - left < 2 or bottom - top < 2:
            return image
        return image.crop((left, top, right, bottom))

    @staticmethod
    def _split_ratio_by_gap(image: Image.Image) -> Optional[tuple[Image.Image, Image.Image]]:
        binary = MainWindow._binarize_for_ratio(image)
        width, height = binary.size
        if width < 6 or height < 4:
            return None
        pixels = binary.load()
        counts = []
        for x in range(width):
            count = 0
            for y in range(height):
                if pixels[x, y] > 0:
                    count += 1
            counts.append(count)
        start = int(width * 0.25)
        end = int(width * 0.75)
        min_count = None
        min_idx = None
        for x in range(start, end):
            count = counts[x]
            if min_count is None or count < min_count:
                min_count = count
                min_idx = x
        if min_idx is None:
            return None
        if min_count is not None and min_count > max(2, int(height * 0.15)):
            return None
        left_img = image.crop((0, 0, min_idx, height))
        right_img = image.crop((min_idx + 1, 0, width, height))
        if left_img.size[0] < 2 or right_img.size[0] < 2:
            return None
        return left_img, right_img

    def _ocr_ratio_side(self, image: Image.Image, configs: list[str]) -> Optional[float]:
        best_text = ""
        for config in configs:
            text = pytesseract.image_to_string(image, config=config, lang="eng").strip()
            best_text = text or best_text
            match = re.search(r"\d{1,6}(?:[.,]\d{1,2})?", text)
            if match:
                value, _, _ = self._parse_ratio_token(match.group(0))
                if value is not None:
                    return value
        if best_text:
            match = re.search(r"\d{1,6}(?:[.,]\d{1,2})?", best_text)
            if match:
                value, _, _ = self._parse_ratio_token(match.group(0))
                return value
        return None

    @staticmethod
    def _classify_digit_1_7(image: Image.Image) -> Optional[str]:
        binary = MainWindow._binarize_for_ratio(image)
        width, height = binary.size
        if width < 2 or height < 2:
            return None
        pixels = binary.load()

        top_rows = max(1, int(height * 0.25))
        top_count = 0
        for y in range(top_rows):
            for x in range(width):
                if pixels[x, y] > 0:
                    top_count += 1
        top_ratio = top_count / (top_rows * width)

        col_start = max(0, width // 2 - 1)
        col_end = min(width, col_start + 3)
        vert_count = 0
        for y in range(height):
            for x in range(col_start, col_end):
                if pixels[x, y] > 0:
                    vert_count += 1
        vert_ratio = vert_count / (height * max(1, col_end - col_start))

        upper_right_count = 0
        upper_right_area = max(1, (width - width // 2) * (height // 2))
        for y in range(height // 2):
            for x in range(width // 2, width):
                if pixels[x, y] > 0:
                    upper_right_count += 1
        upper_right_ratio = upper_right_count / upper_right_area

        if top_ratio >= 0.5 and vert_ratio < 0.7:
            return "7"
        if top_ratio <= 0.35 and vert_ratio >= 0.65:
            return "1"
        if upper_right_ratio > 0.35 and top_ratio > 0.4:
            return "7"
        if vert_ratio > upper_right_ratio + 0.2:
            return "1"
        return None

    @staticmethod
    def _ratio_text_from_boxes(image: Image.Image, config: str) -> str:
        try:
            data = pytesseract.image_to_boxes(image, config=config, lang="eng")
        except Exception:
            return ""
        items = []
        height = image.size[1]
        for line in data.splitlines():
            parts = line.split()
            if len(parts) < 5:
                continue
            char = parts[0]
            if char not in "0123456789:.,/":
                continue
            try:
                x1 = int(parts[1])
                y1 = int(parts[2])
                x2 = int(parts[3])
                y2 = int(parts[4])
            except ValueError:
                continue
            width = max(1, x2 - x1)
            box_height = max(1, y2 - y1)
            aspect = width / box_height
            if char in ("1", "7"):
                roi = image.crop((x1, height - y2, x2, height - y1))
                corrected = MainWindow._classify_digit_1_7(roi)
                if corrected is not None:
                    char = corrected
                elif char == "7" and aspect < 0.35:
                    char = "1"
                elif char == "1" and aspect > 0.6:
                    char = "7"
            items.append((x1, char))
        if not items:
            return ""
        items.sort(key=lambda item: item[0])
        return "".join(char for _, char in items)

    @staticmethod
    def _parse_ratio_token(token: str) -> tuple[Optional[float], bool, bool]:
        normalized = token.replace(",", ".")
        explicit_decimal = "." in normalized
        if explicit_decimal:
            try:
                return float(normalized), True, False
            except ValueError:
                return None, False, False
        try:
            value = int(normalized)
        except ValueError:
            return None, False, False
        if value >= 1000 and value % 100 != 0:
            return value / 100.0, False, True
        return float(value), False, False

    @staticmethod
    def _format_ratio_value(value: Optional[float]) -> str:
        if value is None:
            return "--"
        if abs(value - round(value)) < 1e-6:
            return str(int(round(value)))
        return f"{value:.2f}".rstrip("0").rstrip(".")

    def _ocr_ratio(self, image: Image.Image, fast: bool) -> tuple[Optional[float], Optional[float], str]:
        candidate_map: dict[tuple[float, float], dict[str, int]] = {}
        best_text = ""
        thresholds = (170,) if fast else (170, 150, 130, 110)
        configs = self._ratio_configs[:1] if fast else self._ratio_configs
        scale = 3 if fast else 4
        mode = "fast" if fast else "full"
        boxes_checked = False
        split_checked = False
        ratio_image = self._auto_crop_ratio(image)
        for threshold in thresholds:
            for processed in self._preprocess_variants(
                ratio_image, scale=scale, threshold_value=threshold, mode=mode
            ):
                for config in configs:
                    text = pytesseract.image_to_string(processed, config=config, lang="eng").strip()
                    best_text = text or best_text
                    for match in self._parser.ratio_re.finditer(text):
                        num_value, num_explicit, num_inferred = self._parse_ratio_token(match.group(1))
                        den_value, den_explicit, den_inferred = self._parse_ratio_token(match.group(2))
                        if num_value is None or den_value is None:
                            continue
                        if num_value <= 0 or den_value <= 0:
                            continue
                        explicit = num_explicit or den_explicit
                        inferred = num_inferred or den_inferred
                        score = self._score_ratio(num_value, den_value, explicit, inferred)
                        key = (round(num_value, 2), round(den_value, 2))
                        entry = candidate_map.get(key)
                        if entry is None:
                            candidate_map[key] = {"count": 1, "score": score}
                        else:
                            entry["count"] += 1
                            entry["score"] = max(entry["score"], score)
                if not boxes_checked and self._is_binary_image(processed):
                    box_text = self._ratio_text_from_boxes(processed, configs[0])
                    boxes_checked = True
                    if box_text and box_text != best_text:
                        for match in self._parser.ratio_re.finditer(box_text):
                            num_value, num_explicit, num_inferred = self._parse_ratio_token(match.group(1))
                            den_value, den_explicit, den_inferred = self._parse_ratio_token(match.group(2))
                            if num_value is None or den_value is None:
                                continue
                            if num_value <= 0 or den_value <= 0:
                                continue
                            explicit = num_explicit or den_explicit
                            inferred = num_inferred or den_inferred
                            score = self._score_ratio(num_value, den_value, explicit, inferred) + 2
                            key = (round(num_value, 2), round(den_value, 2))
                            entry = candidate_map.get(key)
                            if entry is None:
                                candidate_map[key] = {"count": 1, "score": score}
                            else:
                                entry["count"] += 1
                                entry["score"] = max(entry["score"], score)
                if not split_checked and self._is_binary_image(processed):
                    split_checked = True
                    split = self._split_ratio_by_gap(processed)
                    if split:
                        left_img, right_img = split
                        left_value = self._ocr_ratio_side(left_img, self._ratio_digit_configs)
                        right_value = self._ocr_ratio_side(right_img, self._ratio_digit_configs)
                        if left_value is not None and right_value is not None:
                            score = self._score_ratio(left_value, right_value, False, False) + 3
                            key = (round(left_value, 2), round(right_value, 2))
                            entry = candidate_map.get(key)
                            if entry is None:
                                candidate_map[key] = {"count": 1, "score": score}
                            else:
                                entry["count"] += 1
                                entry["score"] = max(entry["score"], score)
        if candidate_map:
            best_key, meta = max(
                candidate_map.items(), key=lambda item: (item[1]["count"], item[1]["score"])
            )
            num_value, den_value = best_key
            return (
                num_value,
                den_value,
                f"{self._format_ratio_value(num_value)}:{self._format_ratio_value(den_value)}",
            )

        ratios = self._parser._collect_ratios(best_text)
        if ratios:
            best_ratio = max(ratios, key=lambda pair: (pair[1], pair[0]))
            return (
                best_ratio[0],
                best_ratio[1],
                f"{self._format_ratio_value(best_ratio[0])}:{self._format_ratio_value(best_ratio[1])}",
            )
        return None, None, best_text

    @staticmethod
    def _score_ratio(num: float, den: float, explicit_decimal: bool, inferred_decimal: bool) -> int:
        score = 0
        if num > 0 and den > 0:
            score += 10
        if abs(num - 1) < 1e-6 or abs(den - 1) < 1e-6:
            score += 4
        if num <= 10000 and den <= 10000:
            score += 4
        if num <= 1000 and den <= 1000:
            score += 2
        if explicit_decimal:
            score += 3
        if inferred_decimal:
            score += 1
        if abs(num - den) > 1e-6:
            score += 1
        return score

    def _ocr_box_value(self, image: Image.Image, fast: bool) -> tuple[Optional[int], str]:
        best_text = ""
        thresholds = (150,) if fast else (150, 120)
        configs = self._box_configs[:1] if fast else self._box_configs
        scale = 3
        mode = "fast" if fast else "full"
        for threshold in thresholds:
            for processed in self._preprocess_variants(image, scale=scale, threshold_value=threshold, mode=mode):
                for config in configs:
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

    def _normalize_inputs(
        self,
        left_value: Optional[int],
        right_value: Optional[int],
    ) -> tuple[Optional[int], Optional[int]]:
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

        if is_valid:
            display = result
        else:
            display = self._last_good if self._last_good else result

        left_input = display.left_value
        right_input = display.right_value
        self._left_value[1].setText(str(left_input) if left_input is not None else "--")
        self._right_value[1].setText(str(right_input) if right_input is not None else "--")
        if display.has_ratio:
            ratio_text = f"{self._format_ratio_value(display.ratio_num)} : {self._format_ratio_value(display.ratio_den)}"
        else:
            ratio_text = "--"
        self._ratio_value.setText(ratio_text)
        self._ratio_summary[1].setText(ratio_text)
        self._mini_ratio_value.setText(ratio_text)

        expected_left, expected_right = self._compute_expected(left_input, right_input, display)
        self._expected_left[1].setText(str(expected_left) if expected_left is not None else "--")
        self._expected_right[1].setText(str(expected_right) if expected_right is not None else "--")

        mode = self._resolve_mode(left_input, right_input, display.ratio_num, display.ratio_den)
        if mode == "from_right":
            self._recommended_label.setText("Recommended I want")
            recommended_text = str(expected_left) if expected_left is not None else "--"
            self._recommended_value.setText(recommended_text)
            self._mini_recommended_label.setText("Recommended I want")
            self._mini_recommended_value.setText(recommended_text)
        else:
            self._recommended_label.setText("Recommended I have")
            recommended_text = str(expected_right) if expected_right is not None else "--"
            self._recommended_value.setText(recommended_text)
            self._mini_recommended_label.setText("Recommended I have")
            self._mini_recommended_value.setText(recommended_text)

        if is_valid and left_input is not None and right_input is not None:
            self._last_inputs = (left_input, right_input)

        confidence = self._compute_confidence(display)
        self._confidence_value[1].setText(f"{confidence}%")
        self._last_update.setText(f"Last update: {QtCore.QTime.currentTime().toString('HH:mm:ss')}")
        if not is_valid and self._last_good and self._bad_streak <= 3:
            self._set_status("OCR unstable; holding last good values.")
        else:
            self._set_status(self._build_status(display))
        logging.info(
            "OCR updated: ratio=%s:%s left=%s right=%s expected_left=%s expected_right=%s mode=%s",
            display.ratio_num,
            display.ratio_den,
            left_input,
            right_input,
            expected_left,
            expected_right,
            mode,
        )
        self._busy = False

    def _handle_error(self, message: str) -> None:
        self._set_status(message)
        logging.error(message)
        self._busy = False

    def _set_status(self, message: str) -> None:
        self._status_label.setText(message)

    def _resolve_mode(
        self,
        left_input: Optional[int],
        right_input: Optional[int],
        ratio_num: Optional[float],
        ratio_den: Optional[float],
    ) -> str:
        if self._calc_mode != "auto":
            return self._calc_mode
        if left_input is None and right_input is not None:
            return "from_right"
        if right_input is None and left_input is not None:
            return "from_left"
        if left_input is None or right_input is None:
            return self._auto_direction
        if ratio_num is not None and ratio_den is not None and ratio_den != 0:
            if ratio_num < ratio_den:
                self._auto_direction = "from_left"
            elif ratio_num > ratio_den:
                self._auto_direction = "from_right"
        if self._last_inputs is None:
            return self._auto_direction

        last_left, last_right = self._last_inputs
        left_delta = abs(left_input - last_left)
        right_delta = abs(right_input - last_right)
        if left_delta == 0 and right_delta == 0:
            return self._auto_direction
        if left_delta >= right_delta:
            self._auto_direction = "from_left"
        else:
            self._auto_direction = "from_right"
        return self._auto_direction

    @staticmethod
    def _compute_expected(
        left_input: Optional[int],
        right_input: Optional[int],
        result: ParseResult,
    ) -> tuple[Optional[int], Optional[int]]:
        if not result.has_ratio or result.ratio_num is None or abs(result.ratio_num) < 1e-9:
            return None, None
        expected_left = None
        expected_right = None
        if right_input is not None:
            expected_left = round(right_input * (result.ratio_num / result.ratio_den))
        if left_input is not None:
            expected_right = round(left_input * (result.ratio_den / result.ratio_num))
        return expected_left, expected_right

    @staticmethod
    def _is_valid_result(result: ParseResult) -> bool:
        if not result.has_ratio:
            return False
        if result.left_value is None and result.right_value is None:
            return False
        if result.left_value is not None and result.left_value <= 0:
            return False
        if result.right_value is not None and result.right_value <= 0:
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
            "calc_mode": self._calc_mode,
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
            self._calc_mode = data.get("calc_mode", "auto")
            if hasattr(self, "_swap_checkbox"):
                self._swap_checkbox.setChecked(self._swap_sides)
            if hasattr(self, "_mode_combo"):
                self._set_mode_combo(self._calc_mode)
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
        if result.left_value is not None:
            score += 25
        if result.right_value is not None:
            score += 25
        return score

    @staticmethod
    def _build_status(result: ParseResult) -> str:
        if not result.has_ratio:
            return "Looking for market ratio..."
        if result.left_value is None:
            return "Looking for I want..."
        if result.right_value is None:
            return "Looking for I have..."
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
