"""
MultiProfile Launcher — single-file edition.

Run with: python main.py
Package with PyInstaller: pyinstaller --onefile --windowed main.py

Optional — Web Checks tab (Snapp vouchers / Okala discounts):
    pip install playwright
No separate browser download is needed: the feature attaches (via CDP) to the
user's own installed Chromium-based browser and only Chromium-engine profiles
are selectable. Firefox-family (Gecko) profiles are not attachable.
"""

from __future__ import annotations

import base64
import configparser
import datetime
import json
import logging
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.parse
import uuid
from dataclasses import dataclass, field
from enum import Enum
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Optional

from PySide6.QtCore import (
    QPoint,
    QRect,
    QSize,
    Qt,
    QMetaObject,
    QObject,
    QTimer,
    QThreadPool,
    QRunnable,
    QtMsgType,
    Signal,
    qInstallMessageHandler,
)
from PySide6 import QtGui
from PySide6.QtGui import (
    QAction, QActionGroup, QCloseEvent, QCursor, QColor, QIcon, QKeySequence,
    QPixmap, QShortcut,
)

# Window Layout Manager (window tiling tool) is bundled as a sibling module
# and merged into this app as an extra tool accessible from the header bar.
# It needs pywin32 and only makes sense on Windows, so both the import and
# every use of it are guarded — if either is missing, the browser/profile
# launcher keeps working exactly as before, just without that one button.
_LAYOUT_MANAGER_AVAILABLE = False
if sys.platform == "win32":
    try:
        from window_layout_manager import (
            WindowLayoutManagerApp,
            GlobalHotkeyEventFilter,
        )
        _LAYOUT_MANAGER_AVAILABLE = True
    except Exception:  # pragma: no cover - e.g. pywin32 not installed
        logging.getLogger(__name__).warning(
            "Window Layout Manager unavailable (missing pywin32?); "
            "its button will be hidden.", exc_info=True,
        )
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QLayout,
    QLayoutItem,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QSystemTrayIcon,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


APP_VERSION = "1.0.0"
APP_NAME = "MultiProfile Launcher"

# pywin32 is optional. It powers (a) real .lnk shortcut icon extraction and
# (b) the Window Layout Manager. If it's missing, the app still works: those
# features degrade gracefully to their fallbacks. Detect availability up front
# so we can branch without paying the cost of a failing import each time.
_PYWIN32_AVAILABLE = False
if sys.platform == "win32":
    try:  # pragma: no cover - pywin32 present on most Windows dev boxes
        import win32com.client  # noqa: F401
        import win32gui  # noqa: F401
        import win32ui  # noqa: F401
        import win32con  # noqa: F401
        _PYWIN32_AVAILABLE = True
    except Exception:  # pragma: no cover - pywin32 not installed
        _PYWIN32_AVAILABLE = False

# Playwright powers the Web Checks tab (automating the user's already-open
# Chromium browser via CDP to detect login session and scrape vouchers). It is
# an optional dependency: if it's not installed the app still runs, and the
# Web Checks tab is hidden with a hint instead of crashing.
_WEB_AUTOMATION_AVAILABLE = False
try:
    from playwright.sync_api import sync_playwright  # type: ignore
    _WEB_AUTOMATION_AVAILABLE = True
except Exception:
    logging.getLogger(__name__).warning(
        "Playwright not installed; the Web Checks tab is hidden. "
        "Install with: pip install playwright",
    )


def build_app_icon() -> QIcon:
    """Programmatically drawn app icon (no asset file needed), used for the
    window title bar, taskbar and system tray. If an app.ico exists next to
    the script it is preferred so a packaged build can supply a real icon."""
    candidate = Path(sys.argv[0]).with_name("app.ico")
    if candidate.is_file():
        icon = QIcon(str(candidate))
        if not icon.isNull():
            return icon
    # Fallback: draw a rounded square with a bright accent and a browser glyph.
    pm = QPixmap(64, 64)
    pm.fill(Qt.GlobalColor.transparent)
    from PySide6.QtGui import QPainter, QPen, QFont
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setBrush(QColor(Tokens.AMBER))
    painter.setPen(Qt.PenStyle.NoPen)
    painter.drawRoundedRect(4, 4, 56, 56, 14, 14)
    painter.setPen(QPen(QColor(Tokens.ON_AMBER), 3))
    font = QFont()
    font.setPixelSize(34)
    font.setBold(True)
    painter.setFont(font)
    painter.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, "M")
    painter.end()
    return QIcon(pm)


# ══════════════════════════════════════════════════════════════════════
#  SECTION 1 — THEMES, DESIGN TOKENS & STYLESHEET
# ══════════════════════════════════════════════════════════════════════

# Every colour the UI paints comes from one of these palettes. A theme is a
# name -> colour map, so adding one is a data change, never a code change:
# append an entry and it shows up in Settings and the tray menu.
#
# ON_AMBER / ON_TEAL / ON_TAG are the foreground colours for text drawn *on
# top of* an accent fill or a tag chip. They are tokens rather than literals
# because one theme may use a pale accent (Nord yellow) and another a mid-tone
# one (Solarized cyan), which need opposite text colours to stay readable.

DEFAULT_THEME = "dark"

THEMES: dict[str, dict[str, Any]] = {
    "dark": {
        "label": "Dark (default)",
        "colors": {
            "BG_BASE":         "#0F1117",
            "BG_PANEL":        "#161923",
            "BG_PANEL_RAISED": "#1D2130",
            "BG_CARD":         "#1A1E2B",
            "BG_CARD_HOVER":   "#20253A",
            "BG_INPUT":        "#0C0F18",
            "BORDER_SUBTLE":   "#252A3A",
            "BORDER_MID":      "#2E3447",
            "BORDER_FOCUS":    "#3D4A6B",
            "TEXT_PRIMARY":    "#E4E6F0",
            "TEXT_SECONDARY":  "#7C84A0",
            "TEXT_MUTED":      "#4A5068",
            "TEXT_DISABLED":   "#363A4F",
            "AMBER":           "#FF8A3D",
            "AMBER_HOVER":     "#FFA05E",
            "AMBER_PRESSED":   "#E67228",
            "AMBER_SUBTLE":    "#2A1E10",
            "ON_AMBER":        "#1A0D00",
            "TEAL":            "#3DDFC8",
            "TEAL_HOVER":      "#5EEBD4",
            "TEAL_PRESSED":    "#28C4A8",
            "TEAL_SUBTLE":     "#0E2825",
            "ON_TEAL":         "#061A18",
            "SUCCESS":         "#4DD9A0",
            "DANGER":          "#FF5F57",
            "DANGER_HOVER":    "#FF7B73",
            "WARNING":         "#FFD166",
            "ON_TAG":          "#0C0F18",
        },
        "tags": (
            "#FF8A3D", "#3DDFC8", "#6FA8FF", "#D88CFF",
            "#FF6F9C", "#FFD166", "#7EE8A2",
        ),
    },
    "midnight": {
        "label": "Midnight (OLED)",
        "colors": {
            "BG_BASE":         "#05070C",
            "BG_PANEL":        "#0A0E16",
            "BG_PANEL_RAISED": "#121725",
            "BG_CARD":         "#0D1119",
            "BG_CARD_HOVER":   "#151B29",
            "BG_INPUT":        "#020407",
            "BORDER_SUBTLE":   "#1A2130",
            "BORDER_MID":      "#242C40",
            "BORDER_FOCUS":    "#3A4A70",
            "TEXT_PRIMARY":    "#EAECF5",
            "TEXT_SECONDARY":  "#8A93AD",
            "TEXT_MUTED":      "#575F79",
            "TEXT_DISABLED":   "#3A4055",
            "AMBER":           "#FFA24D",
            "AMBER_HOVER":     "#FFB673",
            "AMBER_PRESSED":   "#E8893A",
            "AMBER_SUBTLE":    "#2E1F0E",
            "ON_AMBER":        "#1A0D00",
            "TEAL":            "#4DEBD5",
            "TEAL_HOVER":      "#6FF3E0",
            "TEAL_PRESSED":    "#33CDB8",
            "TEAL_SUBTLE":     "#0C2E2A",
            "ON_TEAL":         "#04191A",
            "SUCCESS":         "#58E3AC",
            "DANGER":          "#FF6B62",
            "DANGER_HOVER":    "#FF877F",
            "WARNING":         "#FFD97A",
            "ON_TAG":          "#06090E",
        },
        "tags": (
            "#FFA24D", "#4DEBD5", "#7FB6FF", "#E09CFF",
            "#FF7EA8", "#FFD97A", "#8BF0B0",
        ),
    },
    "nord": {
        "label": "Nord",
        "colors": {
            "BG_BASE":         "#2E3440",
            "BG_PANEL":        "#3B4252",
            "BG_PANEL_RAISED": "#434C5E",
            "BG_CARD":         "#3B4252",
            "BG_CARD_HOVER":   "#4C566A",
            "BG_INPUT":        "#272C36",
            "BORDER_SUBTLE":   "#434C5E",
            "BORDER_MID":      "#4C566A",
            "BORDER_FOCUS":    "#81A1C1",
            "TEXT_PRIMARY":    "#ECEFF4",
            "TEXT_SECONDARY":  "#A6B1C2",
            "TEXT_MUTED":      "#7C8799",
            "TEXT_DISABLED":   "#5A6474",
            "AMBER":           "#EBCB8B",
            "AMBER_HOVER":     "#F3DCAB",
            "AMBER_PRESSED":   "#D6B476",
            "AMBER_SUBTLE":    "#3E3A2D",
            "ON_AMBER":        "#2E2A18",
            "TEAL":            "#88C0D0",
            "TEAL_HOVER":      "#A3D4E1",
            "TEAL_PRESSED":    "#6FA8BB",
            "TEAL_SUBTLE":     "#2B3B44",
            "ON_TEAL":         "#1B2A32",
            "SUCCESS":         "#A3BE8C",
            "DANGER":          "#BF616A",
            "DANGER_HOVER":    "#CF7C84",
            "WARNING":         "#EBCB8B",
            "ON_TAG":          "#2E3440",
        },
        "tags": (
            "#D08770", "#88C0D0", "#81A1C1", "#B48EAD",
            "#BF616A", "#EBCB8B", "#A3BE8C",
        ),
    },
    "solarized": {
        "label": "Solarized Dark",
        "colors": {
            "BG_BASE":         "#002B36",
            "BG_PANEL":        "#073642",
            "BG_PANEL_RAISED": "#0B4551",
            "BG_CARD":         "#073642",
            "BG_CARD_HOVER":   "#0F4E5C",
            "BG_INPUT":        "#002129",
            "BORDER_SUBTLE":   "#124E5C",
            "BORDER_MID":      "#1C6070",
            "BORDER_FOCUS":    "#268BD2",
            "TEXT_PRIMARY":    "#EEE8D5",
            "TEXT_SECONDARY":  "#93A1A1",
            "TEXT_MUTED":      "#6A848C",
            "TEXT_DISABLED":   "#48707A",
            "AMBER":           "#B58900",
            "AMBER_HOVER":     "#C99A1F",
            "AMBER_PRESSED":   "#9B7400",
            "AMBER_SUBTLE":    "#12303A",
            "ON_AMBER":        "#1A1500",
            "TEAL":            "#2AA198",
            "TEAL_HOVER":      "#35BDB3",
            "TEAL_PRESSED":    "#21867F",
            "TEAL_SUBTLE":     "#06333B",
            "ON_TEAL":         "#05191B",
            "SUCCESS":         "#859900",
            "DANGER":          "#DC322F",
            "DANGER_HOVER":    "#E5544F",
            "WARNING":         "#B58900",
            "ON_TAG":          "#00252E",
        },
        "tags": (
            "#CB4B16", "#2AA198", "#268BD2", "#6C71C4",
            "#D33682", "#B58900", "#859900",
        ),
    },
    "light": {
        "label": "Light",
        "colors": {
            "BG_BASE":         "#F2F4F8",
            "BG_PANEL":        "#FFFFFF",
            "BG_PANEL_RAISED": "#FFFFFF",
            "BG_CARD":         "#FFFFFF",
            "BG_CARD_HOVER":   "#EEF1F7",
            "BG_INPUT":        "#FFFFFF",
            "BORDER_SUBTLE":   "#E3E6EE",
            "BORDER_MID":      "#CFD4E0",
            "BORDER_FOCUS":    "#8E9BC4",
            "TEXT_PRIMARY":    "#1A1F2B",
            "TEXT_SECONDARY":  "#5B6377",
            "TEXT_MUTED":      "#8C93A6",
            "TEXT_DISABLED":   "#B6BCC9",
            "AMBER":           "#E0701F",
            "AMBER_HOVER":     "#F08434",
            "AMBER_PRESSED":   "#C55F14",
            "AMBER_SUBTLE":    "#FCEEE2",
            "ON_AMBER":        "#2A1503",
            "TEAL":            "#0E9E8A",
            "TEAL_HOVER":      "#12B8A1",
            "TEAL_PRESSED":    "#0A8574",
            "TEAL_SUBTLE":     "#E1F6F2",
            "ON_TEAL":         "#04241F",
            "SUCCESS":         "#16A34A",
            "DANGER":          "#D92D20",
            "DANGER_HOVER":    "#EA4A3E",
            "WARNING":         "#B25E09",
            "ON_TAG":          "#FFFFFF",
        },
        "tags": (
            "#D9540F", "#0E9E8A", "#2563EB", "#9333EA",
            "#DB2777", "#B45309", "#15803D",
        ),
    },
}

# A theme that is missing (or inventing) a colour key would leave some token
# pointing at the previous palette after a switch. Fail loudly at import
# instead of shipping a half-painted UI.
_THEME_COLOR_KEYS = frozenset(THEMES[DEFAULT_THEME]["colors"])
for _theme_name, _theme_data in THEMES.items():
    _missing = _THEME_COLOR_KEYS - set(_theme_data["colors"])
    _extra = set(_theme_data["colors"]) - _THEME_COLOR_KEYS
    if _missing or _extra:
        raise RuntimeError(
            f"Theme '{_theme_name}' colour keys do not match "
            f"'{DEFAULT_THEME}' (missing={sorted(_missing)}, extra={sorted(_extra)})"
        )
del _theme_name, _theme_data, _missing, _extra


class Tokens:
    """The active palette.

    Theme-independent tokens (fonts, sizes, radii) are plain class constants.
    The colour tokens are installed by :meth:`apply`, which runs once at import
    time and again on every live theme switch, so reading ``Tokens.BG_CARD``
    always yields the *current* theme's colour.
    """

    FONT_UI   = "Segoe UI Variable Text, Segoe UI, Inter, sans-serif"
    FONT_MONO = "JetBrains Mono, Cascadia Mono, Consolas, monospace"

    SIZE_XS  = 8
    SIZE_SM  = 9
    SIZE_MD  = 10
    SIZE_LG  = 12
    SIZE_XL  = 15
    SIZE_2XL = 18

    R_CARD = 10
    R_BTN  = 6
    R_CHIP = 4
    R_TAG  = 3

    _active = DEFAULT_THEME

    @classmethod
    def apply(cls, theme_key: str) -> str:
        """Install *theme_key*'s colours onto the class and return the key that
        was actually applied. Unknown keys fall back to the default theme
        rather than raising, so a hand-edited config.json can never break
        startup.
        """
        key = theme_key if theme_key in THEMES else DEFAULT_THEME
        theme = THEMES[key]
        for name, value in theme["colors"].items():
            setattr(cls, name, value)
        cls.TAG_PALETTE = tuple(theme["tags"])
        cls._active = key
        return key

    @classmethod
    def theme_key(cls) -> str:
        """Key of the theme currently installed."""
        return cls._active

    @classmethod
    def theme_choices(cls) -> list[tuple[str, str]]:
        """``[(key, label)]`` for the current theme, in menu order."""
        return [(key, data["label"]) for key, data in THEMES.items()]


# Install the default palette immediately so every ``Tokens.<COLOR>`` read
# below (including the ones in widget constructors) sees a real colour.
Tokens.apply(DEFAULT_THEME)


def build_stylesheet() -> str:
    t = Tokens
    return f"""
        QWidget {{
            background-color: {t.BG_BASE};
            color: {t.TEXT_PRIMARY};
            font-family: {t.FONT_UI};
            font-size: {t.SIZE_MD}pt;
        }}
        QMainWindow {{ background-color: {t.BG_BASE}; }}

        #HeaderBar {{
            background-color: {t.BG_PANEL};
            border-bottom: 1px solid {t.BORDER_SUBTLE};
        }}
        #AppTitle {{
            font-size: {t.SIZE_XL}pt;
            font-weight: 700;
            color: {t.TEXT_PRIMARY};
            letter-spacing: -0.3px;
        }}
        #StatusLabel {{ color: {t.TEXT_MUTED}; font-size: {t.SIZE_SM}pt; }}
        #StatusLabelSuccess {{ color: {t.SUCCESS}; font-size: {t.SIZE_SM}pt; }}
        #StatusLabelDanger  {{ color: {t.DANGER};  font-size: {t.SIZE_SM}pt; }}

        #ModeTabBar {{
            background-color: {t.BG_PANEL};
            border-bottom: 1px solid {t.BORDER_SUBTLE};
        }}
        #ModeTab {{
            background-color: transparent;
            color: {t.TEXT_SECONDARY};
            border: none;
            border-bottom: 3px solid transparent;
            border-radius: 0px;
            padding: 10px 24px;
            font-size: {t.SIZE_LG}pt;
            font-weight: 600;
        }}
        #ModeTab:hover {{ color: {t.TEXT_PRIMARY}; background-color: {t.BG_PANEL_RAISED}; }}
        #ModeTabActive {{
            background-color: transparent;
            color: {t.AMBER};
            border: none;
            border-bottom: 3px solid {t.AMBER};
            border-radius: 0px;
            padding: 10px 24px;
            font-size: {t.SIZE_LG}pt;
            font-weight: 700;
        }}
        #Sidebar {{
            background-color: {t.BG_PANEL};
            border-right: 1px solid {t.BORDER_SUBTLE};
        }}
        #SidebarTitle {{
            color: {t.TEXT_MUTED};
            font-size: {t.SIZE_SM}pt;
            font-weight: 700;
            letter-spacing: 1.2px;
            text-transform: uppercase;
            padding: 8px 12px 4px 12px;
        }}
        #SidebarDivider {{ background-color: {t.BORDER_SUBTLE}; }}
        QListWidget#WorkspaceList,
        QListWidget#BrowserList {{ background: transparent; border: none; outline: none; }}
        QListWidget#WorkspaceList::item,
        QListWidget#BrowserList::item {{
            padding: 9px 10px;
            border-radius: {t.R_BTN}px;
            margin: 1px 6px;
            color: {t.TEXT_SECONDARY};
        }}
        QListWidget#WorkspaceList::item:selected,
        QListWidget#BrowserList::item:selected {{
            background-color: {t.AMBER_SUBTLE};
            color: {t.AMBER};
        }}
        QListWidget#WorkspaceList::item:hover:!selected,
        QListWidget#BrowserList::item:hover:!selected {{
            background-color: {t.BG_PANEL_RAISED};
            color: {t.TEXT_PRIMARY};
        }}

        #ProfileCard {{
            background-color: {t.BG_CARD};
            border: 1px solid {t.BORDER_SUBTLE};
            border-radius: {t.R_CARD}px;
        }}
        #ProfileCard:hover {{
            background-color: {t.BG_CARD_HOVER};
            border: 1px solid {t.BORDER_MID};
        }}
        #ProfileCardTag {{ border-radius: {t.R_TAG}px; }}
        #ProfileCardBrowser {{
            color: {t.TEXT_MUTED};
            font-size: {t.SIZE_SM}pt;
            font-weight: 600;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }}
        #ProfileCardName {{ font-size: {t.SIZE_LG}pt; font-weight: 700; color: {t.TEXT_PRIMARY}; }}
        #ProfileCardArg {{
            color: {t.TEXT_SECONDARY};
            font-family: {t.FONT_MONO};
            font-size: {t.SIZE_SM}pt;
            background-color: {t.BG_INPUT};
            border: 1px solid {t.BORDER_SUBTLE};
            border-radius: {t.R_CHIP}px;
            padding: 2px 6px;
        }}

        QCheckBox {{ spacing: 0px; background: transparent; }}
        QCheckBox::indicator {{
            width: 16px; height: 16px;
            border: 1.5px solid {t.BORDER_MID};
            border-radius: 4px;
            background-color: {t.BG_INPUT};
        }}
        QCheckBox::indicator:hover {{ border-color: {t.BORDER_FOCUS}; }}
        QCheckBox::indicator:checked {{
            background-color: {t.AMBER};
            border-color: {t.AMBER};
            image: none;
        }}
        QPushButton {{
            background-color: {t.BG_PANEL_RAISED};
            color: {t.TEXT_PRIMARY};
            border: 1px solid {t.BORDER_MID};
            border-radius: {t.R_BTN}px;
            padding: 6px 14px;
            font-size: {t.SIZE_MD}pt;
        }}
        QPushButton:hover {{ background-color: {t.BG_CARD_HOVER}; border-color: {t.BORDER_FOCUS}; }}
        QPushButton:pressed {{ background-color: {t.BG_BASE}; }}
        QPushButton:disabled {{ color: {t.TEXT_DISABLED}; border-color: {t.BORDER_SUBTLE}; }}

        QPushButton#LaunchButton {{
            background-color: {t.AMBER};
            color: {t.ON_AMBER};
            border: none;
            font-weight: 700;
            font-size: {t.SIZE_SM}pt;
            padding: 5px 14px;
        }}
        QPushButton#LaunchButton:hover {{ background-color: {t.AMBER_HOVER}; }}
        QPushButton#LaunchButton:pressed {{ background-color: {t.AMBER_PRESSED}; }}

        QPushButton#BulkLaunchButton {{
            background-color: {t.AMBER};
            color: {t.ON_AMBER};
            border: none;
            font-weight: 700;
            font-size: {t.SIZE_MD}pt;
            padding: 7px 18px;
            border-radius: {t.R_BTN}px;
        }}
        QPushButton#BulkLaunchButton:hover {{ background-color: {t.AMBER_HOVER}; }}
        QPushButton#BulkLaunchButton:disabled {{
            background-color: {t.BG_PANEL_RAISED};
            color: {t.TEXT_DISABLED};
        }}

        QPushButton#RunCheckButton {{
            background-color: {t.TEAL};
            color: {t.ON_TEAL};
            border: none;
            font-weight: 700;
            font-size: {t.SIZE_SM}pt;
            padding: 4px 14px;
            border-radius: {t.R_BTN}px;
        }}
        QPushButton#RunCheckButton:hover {{ background-color: {t.TEAL_HOVER}; }}
        QPushButton#RunCheckButton:disabled {{
            background-color: {t.BG_PANEL_RAISED};
            color: {t.TEXT_DISABLED};
        }}

        QPushButton#SnappExportButton {{
            background-color: #059669;
            color: #FFFFFF;
            border: none;
            font-weight: 700;
            font-size: {t.SIZE_SM}pt;
            padding: 5px 13px;
            border-radius: {t.R_BTN}px;
        }}
        QPushButton#SnappExportButton:hover {{ background-color: #10B981; }}
        QPushButton#SnappExportButton:pressed {{ background-color: #047857; }}
        QPushButton#DigikalaAllButton, QPushButton#DigikalaSelectedButton {{
            background-color: #EF394E;
            color: #FFFFFF;
            border: none;
            font-weight: 700;
            font-size: {t.SIZE_SM}pt;
            padding: 5px 13px;
            border-radius: {t.R_BTN}px;
        }}
        QPushButton#DigikalaAllButton:hover, QPushButton#DigikalaSelectedButton:hover {{ background-color: #F43F5E; }}
        QPushButton#DigikalaAllButton:pressed, QPushButton#DigikalaSelectedButton:pressed {{ background-color: #DC2626; }}
        QPushButton#DigikalaAllButton:disabled, QPushButton#DigikalaSelectedButton:disabled {{
            background-color: {t.BG_PANEL_RAISED};
            color: {t.TEXT_DISABLED};
        }}

        QPushButton#UniversalCheckButton {{
            background-color: #2563EB;
            color: #FFFFFF;
            border: none;
            font-weight: 700;
            font-size: {t.SIZE_SM}pt;
            padding: 5px 13px;
            border-radius: {t.R_BTN}px;
        }}
        QPushButton#UniversalCheckButton:hover {{ background-color: #3B82F6; }}
        QPushButton#UniversalCheckButton:pressed {{ background-color: #1D4ED8; }}

        QPushButton#SnappExportButton:disabled {{
            background-color: {t.BG_PANEL_RAISED};
            color: {t.TEXT_DISABLED};
        }}

        QPushButton#DangerButton {{
            background-color: transparent;
            color: {t.DANGER};
            border: 1px solid {t.DANGER};
        }}
        QPushButton#DangerButton:hover {{ background-color: {t.DANGER}; color: {t.BG_BASE}; }}

        QPushButton#AddInstanceButton {{
            background-color: {t.TEAL_SUBTLE};
            color: {t.TEAL};
            border: 1px solid {t.TEAL};
            border-radius: {t.R_BTN}px;
            padding: 6px 14px;
            font-weight: 600;
        }}
        QPushButton#AddInstanceButton:hover {{ background-color: {t.TEAL}; color: {t.ON_TEAL}; }}

        QLineEdit, QComboBox, QTextEdit, QPlainTextEdit {{
            background-color: {t.BG_INPUT};
            border: 1px solid {t.BORDER_MID};
            border-radius: {t.R_BTN}px;
            padding: 6px 8px;
            color: {t.TEXT_PRIMARY};
            font-size: {t.SIZE_MD}pt;
        }}
        QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {{
            border-color: {t.AMBER};
        }}
        QComboBox::drop-down {{ border: none; padding-right: 8px; }}
        QComboBox QAbstractItemView {{
            background-color: {t.BG_PANEL};
            border: 1px solid {t.BORDER_MID};
            selection-background-color: {t.AMBER_SUBTLE};
            selection-color: {t.AMBER};
        }}

        QScrollArea {{ background-color: transparent; border: none; }}
        QScrollBar:vertical {{ background: transparent; width: 8px; margin: 0px; }}
        QScrollBar::handle:vertical {{
            background: {t.BORDER_MID};
            border-radius: 4px;
            min-height: 20px;
        }}
        QScrollBar::handle:vertical:hover {{ background: {t.BORDER_FOCUS}; }}
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {{ height: 0px; }}
        QScrollBar:horizontal {{ background: transparent; height: 8px; }}
        QScrollBar::handle:horizontal {{
            background: {t.BORDER_MID};
            border-radius: 4px;
            min-width: 20px;
        }}
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {{ width: 0px; }}

        QDialog {{ background-color: {t.BG_PANEL}; }}
        QLabel#DialogTitle {{ font-size: {t.SIZE_LG}pt; font-weight: 700; color: {t.TEXT_PRIMARY}; }}
        QLabel#FormLabel {{ color: {t.TEXT_SECONDARY}; font-size: {t.SIZE_SM}pt; font-weight: 600; }}
        QToolTip {{
            background-color: {t.BG_PANEL_RAISED};
            color: {t.TEXT_PRIMARY};
            border: 1px solid {t.BORDER_FOCUS};
            padding: 4px 8px;
            border-radius: {t.R_CHIP}px;
        }}
        QMessageBox {{ background-color: {t.BG_PANEL}; }}

        #EmptyStateLabel {{ color: {t.TEXT_MUTED}; font-size: {t.SIZE_LG}pt; }}
        #EmptyStateHint {{ color: {t.TEXT_DISABLED}; font-size: {t.SIZE_MD}pt; }}

        #GridActionBar {{
            background-color: {t.BG_PANEL};
            border-bottom: 1px solid {t.BORDER_SUBTLE};
        }}
        #GridActionBarCount {{ color: {t.TEXT_MUTED}; font-size: {t.SIZE_SM}pt; }}
        #GridSearchEdit {{
            background-color: {t.BG_INPUT};
            border: 1px solid {t.BORDER_MID};
            border-radius: {t.R_BTN}px;
            padding: 4px 8px;
            color: {t.TEXT_PRIMARY};
            font-size: {t.SIZE_SM}pt;
        }}
        #GridSearchEdit:focus {{ border-color: {t.AMBER}; }}

        QListWidget#CompactList {{
            background-color: {t.BG_BASE};
            border: none;
            padding: 8px;
        }}
        QListWidget#CompactList::item {{
            background-color: {t.BG_CARD};
            border: 1px solid {t.BORDER_SUBTLE};
            border-radius: {t.R_BTN}px;
            padding: 8px 12px;
            margin: 2px 4px;
            color: {t.TEXT_PRIMARY};
        }}
        QListWidget#CompactList::item:hover {{
            background-color: {t.BG_CARD_HOVER};
            border-color: {t.BORDER_MID};
        }}
        QListWidget#CompactList::item:selected {{
            background-color: {t.AMBER_SUBTLE};
            color: {t.AMBER};
            border-color: {t.AMBER};
        }}

        QPushButton#UndoButton {{
            background-color: {t.TEAL_SUBTLE};
            color: {t.TEAL};
            border: 1px solid {t.TEAL};
            border-radius: {t.R_BTN}px;
            padding: 2px 10px;
            font-size: {t.SIZE_SM}pt;
            font-weight: 600;
        }}
        QPushButton#UndoButton:hover {{ background-color: {t.TEAL}; color: {t.ON_TEAL}; }}
    """


# ══════════════════════════════════════════════════════════════════════
#  SECTION 2 — DATA MODELS
# ══════════════════════════════════════════════════════════════════════

class BrowserEngine(str, Enum):
    CHROMIUM = "chromium"
    GECKO = "gecko"


@dataclass
class BrowserDefinition:
    key: str
    display_name: str
    executable_path: Path
    engine: BrowserEngine
    icon_path: Optional[Path] = None
    auto_discovered: bool = False

    def is_valid(self) -> bool:
        return self.executable_path.is_file()

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "display_name": self.display_name,
            "executable_path": str(self.executable_path),
            "engine": self.engine.value,
            "icon_path": str(self.icon_path) if self.icon_path else None,
            "auto_discovered": self.auto_discovered,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BrowserDefinition:
        return cls(
            key=data["key"],
            display_name=data["display_name"],
            executable_path=Path(data["executable_path"]),
            engine=BrowserEngine(data["engine"]),
            icon_path=Path(data["icon_path"]) if data.get("icon_path") else None,
            auto_discovered=bool(data.get("auto_discovered", False)),
        )


@dataclass
class BrowserProfile:
    id: str
    browser_key: str
    display_name: str
    profile_arg: str
    color_tag: Optional[str] = None
    notes: str = ""
    # Absolute filesystem path to the profile folder. When set for a Gecko
    # (Firefox-family) profile, the launcher uses "-profile <path>" instead
    # of "-P <name>", which works even for profiles that only exist in the
    # browser's newer "selectable profiles" datastore and never got a
    # matching [ProfileN] section in profiles.ini.
    profile_path: Optional[str] = None

    @staticmethod
    def new_id() -> str:
        return uuid.uuid4().hex[:12]

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "browser_key": self.browser_key,
            "display_name": self.display_name,
            "profile_arg": self.profile_arg,
            "color_tag": self.color_tag,
            "notes": self.notes,
            "profile_path": self.profile_path,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BrowserProfile:
        return cls(
            id=data["id"],
            browser_key=data["browser_key"],
            display_name=data["display_name"],
            profile_arg=data["profile_arg"],
            color_tag=data.get("color_tag"),
            notes=data.get("notes", ""),
            profile_path=data.get("profile_path"),
        )


@dataclass
class Workspace:
    id: str
    name: str
    profile_ids: list[str] = field(default_factory=list)
    launch_stagger_seconds: float = 0.4

    @staticmethod
    def new_id() -> str:
        return uuid.uuid4().hex[:12]

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "profile_ids": list(self.profile_ids),
            "launch_stagger_seconds": self.launch_stagger_seconds,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Workspace:
        return cls(
            id=data["id"],
            name=data["name"],
            profile_ids=list(data.get("profile_ids", [])),
            launch_stagger_seconds=float(data.get("launch_stagger_seconds", 0.4)),
        )


# ══════════════════════════════════════════════════════════════════════
#  SECTION 2.5 — WEB CHECK MODELS
# ══════════════════════════════════════════════════════════════════════

class WebCheckKind(str, Enum):
    SNAPP_VOUCHERS = "snapp_vouchers"
    OKALA_DISCOUNTS = "okala_discounts"


@dataclass
class WebCheckTask:
    """A single recurring Web Check that can target MULTIPLE Chromium profiles.

    For every selected profile the check opens an isolated, persistent automation
    data folder (see _automation_user_data_dir), detects whether the user's
    session is logged in, and (if so) scrapes the visible vouchers / discount
    codes. Results are stored per profile in `results`.

    This is for the user's own accounts only. It must not be used for mass
    scraping, rate-limit bypass, or anything that violates the sites' terms.
    """
    id: str
    kind: WebCheckKind
    profile_ids: list[str] = field(default_factory=list)   # Chromium profiles
    display_name: str = ""
    # Per-profile last result: profile_id -> {"session_active": bool|None,
    # "summary": str, "details": [dict], "run_at": ISO str or None, "error": str}
    results: dict[str, dict] = field(default_factory=dict)

    @staticmethod
    def new_id() -> str:
        return uuid.uuid4().hex[:12]

    @property
    def url(self) -> str:
        if self.kind == WebCheckKind.SNAPP_VOUCHERS:
            return "https://app.snapp.taxi/?superapp_service=vouchers"
        if self.kind == WebCheckKind.OKALA_DISCOUNTS:
            return "https://www.okala.com/profile"
        return ""

    def result_for(self, profile_id: str) -> dict:
        return self.results.get(profile_id, {})

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "kind": self.kind.value,
            "profile_ids": list(self.profile_ids),
            "display_name": self.display_name,
            "results": {
                pid: dict(entry) for pid, entry in self.results.items()
            },
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> WebCheckTask:
        # Migration from the single-profile schema (profile_id + last_* fields).
        profile_ids = list(data.get("profile_ids", []))
        results: dict[str, dict] = {
            pid: dict(entry) for pid, entry in data.get("results", {}).items()
        }
        legacy_profile = data.get("profile_id", "")
        if legacy_profile and legacy_profile not in profile_ids:
            profile_ids.insert(0, legacy_profile)
            legacy_result: dict = {}
            if data.get("last_run_at") is not None:
                legacy_result["run_at"] = data["last_run_at"]
            if data.get("last_session_active") is not None:
                legacy_result["session_active"] = data["last_session_active"]
            if data.get("last_result_summary"):
                legacy_result["summary"] = data["last_result_summary"]
            if data.get("last_result_details"):
                legacy_result["details"] = list(data["last_result_details"])
            if legacy_result:
                results[legacy_profile] = legacy_result
        return cls(
            id=data["id"],
            kind=WebCheckKind(data["kind"]),
            profile_ids=profile_ids,
            display_name=data.get("display_name", data["kind"]),
            results=results,
        )


# ══════════════════════════════════════════════════════════════════════
#  SECTION 3 — CONFIG STORE
# ══════════════════════════════════════════════════════════════════════

CONFIG_SCHEMA_VERSION = 2
APP_DIR_NAME = "MultiProfileLauncher"
CONFIG_FILENAME = "config.json"


def get_app_data_dir() -> Path:
    if sys.platform == "win32":
        base = os.environ.get("APPDATA")
        if not base:
            base = str(Path.home() / "AppData" / "Roaming")
        return Path(base) / APP_DIR_NAME
    return Path.home() / ".config" / APP_DIR_NAME


def get_config_path() -> Path:
    return get_app_data_dir() / CONFIG_FILENAME


class ConfigStore:
    def __init__(self, config_path: Path | None = None) -> None:
        self.config_path: Path = config_path or get_config_path()
        self.browsers: dict[str, BrowserDefinition] = {}
        self.profiles: dict[str, BrowserProfile] = {}
        self.workspaces: dict[str, Workspace] = {}
        self.web_check_tasks: dict[str, WebCheckTask] = {}
        self.settings: dict[str, Any] = {
            "start_minimized_to_tray": False,
            "close_to_tray": True,
            "theme": "dark",
            "bulk_launch_confirm_threshold": 10,
            "bulk_launch_stagger_seconds": 2.0,
            "view_mode": "grid",
            "web_check_sort": "session",
            "launch_open_url": "",
        }

    def load(self) -> None:
        if not self.config_path.is_file():
            return
        try:
            raw_text = self.config_path.read_text(encoding="utf-8")
            data = json.loads(raw_text)
        except (OSError, json.JSONDecodeError) as exc:
            self._quarantine_corrupt_file(exc)
            return
        if not isinstance(data, dict):
            self._quarantine_corrupt_file(
                ValueError(f"config root is {type(data).__name__}, expected an object")
            )
            return
        self._populate_from_payload(data)
        # One-time migration: 0.3 s was the old default bulk-launch stagger
        # (near-instant with many profiles). Give configs still carrying the
        # old default the new 2 s default exactly once; any other value the
        # user deliberately chose is left untouched.
        if self.settings.get("bulk_launch_stagger_seconds") == 0.3:
            self.settings["bulk_launch_stagger_seconds"] = 2.0

    def _deserialize_map(self, raw_items: Any, label: str, from_dict, key_of) -> dict:
        """Deserialize a list of raw dicts into {key: object}, skipping (and
        logging) any entry that fails to parse instead of crashing.

        A single malformed entry in config.json (hand-edited, imported from an
        older schema, ...) used to raise out of load() and abort the whole
        process at startup. Skipping the bad entry keeps everything else
        usable.
        """
        result: dict = {}
        if not isinstance(raw_items, list):
            logging.error("Config section '%s' is not a list; ignoring it.", label)
            return result
        for item in raw_items:
            if not isinstance(item, dict):
                logging.error("Skipping malformed %s entry (not an object): %r", label, item)
                continue
            try:
                obj = from_dict(item)
                result[key_of(obj, item)] = obj
            except Exception as exc:
                logging.error("Skipping malformed %s entry: %s (%r)", label, exc, item)
        return result

    def _populate_from_payload(self, payload: dict[str, Any]) -> None:
        """Replace all in-memory state from a payload dict shaped like the
        JSON config file. Shared by load(), import_from() and load_payload()
        so all three tolerate malformed entries the same way."""
        self.browsers = self._deserialize_map(
            payload.get("browsers", []), "browser",
            BrowserDefinition.from_dict, lambda obj, item: obj.key,
        )
        self.profiles = self._deserialize_map(
            payload.get("profiles", []), "profile",
            BrowserProfile.from_dict, lambda obj, item: obj.id,
        )
        self.workspaces = self._deserialize_map(
            payload.get("workspaces", []), "workspace",
            Workspace.from_dict, lambda obj, item: obj.id,
        )
        self.web_check_tasks = self._deserialize_map(
            payload.get("web_check_tasks", []), "web check task",
            WebCheckTask.from_dict, lambda obj, item: obj.id,
        )
        # Merge loaded settings over the defaults so any key missing from an
        # older config gets a sensible value instead of KeyError.
        loaded_settings = payload.get("settings", {})
        if not isinstance(loaded_settings, dict):
            logging.error("Config 'settings' section is not an object; ignoring it.")
            loaded_settings = {}
        self.settings = {**self.settings, **loaded_settings}

    def save(self) -> None:
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "schema_version": CONFIG_SCHEMA_VERSION,
            "browsers": [b.to_dict() for b in self.browsers.values()],
            "profiles": [p.to_dict() for p in self.profiles.values()],
            "workspaces": [w.to_dict() for w in self.workspaces.values()],
            "web_check_tasks": [w.to_dict() for w in self.web_check_tasks.values()],
            "settings": self.settings,
        }
        fd, tmp_path_str = tempfile.mkstemp(
            dir=self.config_path.parent, prefix=".config_", suffix=".tmp"
        )
        tmp_path = Path(tmp_path_str)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)
            os.replace(tmp_path, self.config_path)
        except OSError:
            tmp_path.unlink(missing_ok=True)
            raise

    def _backup_config(self) -> Path | None:
        """Copy the current config.json next to itself with a timestamp."""
        if not self.config_path.is_file():
            return None
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = self.config_path.parent / f"config.backup-{stamp}.json"
        try:
            shutil.copy2(self.config_path, backup)
            return backup
        except OSError as exc:
            logging.warning("Could not back up config: %s", exc)
            return None

    def _quarantine_corrupt_file(self, exc: Exception) -> None:
        logging.warning("Config file is corrupt (%s); quarantining it.", exc)
        try:
            corrupt_path = self.config_path.with_suffix(".corrupt.json")
            self.config_path.replace(corrupt_path)
            logging.warning("Corrupt config moved to: %s", corrupt_path)
        except OSError as move_exc:
            logging.warning("Could not quarantine corrupt config: %s", move_exc)

    def to_payload(self) -> dict[str, Any]:
        """Serialise the whole config to a plain dict (the same shape as the
        JSON file). Used for both export and in-memory snapshots."""
        return {
            "schema_version": CONFIG_SCHEMA_VERSION,
            "browsers": [b.to_dict() for b in self.browsers.values()],
            "profiles": [p.to_dict() for p in self.profiles.values()],
            "workspaces": [w.to_dict() for w in self.workspaces.values()],
            "web_check_tasks": [w.to_dict() for w in self.web_check_tasks.values()],
            "settings": dict(self.settings),
        }

    def export_to(self, path: Path) -> None:
        payload = self.to_payload()
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)

    def import_from(self, path: Path) -> None:
        """Replace the current config with the contents of `path`."""
        data = json.loads(path.read_text(encoding="utf-8"))
        missing = {"browsers", "profiles"} - set(data.keys())
        if missing:
            raise ValueError(f"Not a valid MultiProfile Launcher config (missing: {', '.join(sorted(missing))})")
        self._populate_from_payload(data)

    def load_payload(self, payload: dict[str, Any]) -> None:
        """Replace in-memory state from a payload dict (used by undo/restore)."""
        self._populate_from_payload(payload)

    def upsert_browser(self, browser: BrowserDefinition) -> None:
        self.browsers[browser.key] = browser

    def remove_browser(self, key: str) -> None:
        self.browsers.pop(key, None)
        orphaned_ids = {
            pid for pid, p in self.profiles.items() if p.browser_key == key
        }
        for pid in orphaned_ids:
            self.remove_profile(pid)  # also cleans web-check selections
        for ws in self.workspaces.values():
            ws.profile_ids = [pid for pid in ws.profile_ids if pid not in orphaned_ids]

    def upsert_profile(self, profile: BrowserProfile) -> None:
        self.profiles[profile.id] = profile

    def remove_profile(self, profile_id: str) -> None:
        self.profiles.pop(profile_id, None)
        for ws in self.workspaces.values():
            ws.profile_ids = [pid for pid in ws.profile_ids if pid != profile_id]
        # Keep Web Checks consistent: drop the profile from selections and
        # forget its stored result, so cards never show '(deleted profile)'.
        for task in self.web_check_tasks.values():
            if profile_id in task.profile_ids:
                task.profile_ids = [p for p in task.profile_ids if p != profile_id]
            task.results.pop(profile_id, None)

    def profiles_for_browser(self, browser_key: str) -> list[BrowserProfile]:
        return [p for p in self.profiles.values() if p.browser_key == browser_key]

    def upsert_workspace(self, workspace: Workspace) -> None:
        self.workspaces[workspace.id] = workspace

    def remove_workspace(self, workspace_id: str) -> None:
        self.workspaces.pop(workspace_id, None)

    def upsert_web_check_task(self, task: WebCheckTask) -> None:
        self.web_check_tasks[task.id] = task

    def remove_web_check_task(self, task_id: str) -> None:
        self.web_check_tasks.pop(task_id, None)


# ══════════════════════════════════════════════════════════════════════
#  SECTION 4 — BROWSER DISCOVERY
# ══════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class BrowserCatalogEntry:
    key: str
    display_name: str
    engine: BrowserEngine
    relative_exe_candidates: tuple[str, ...]


BROWSER_CATALOG: tuple[BrowserCatalogEntry, ...] = (
    BrowserCatalogEntry("chrome", "Google Chrome", BrowserEngine.CHROMIUM,
                        ("Google\\Chrome\\Application\\chrome.exe",)),
    BrowserCatalogEntry("edge", "Microsoft Edge", BrowserEngine.CHROMIUM,
                        ("Microsoft\\Edge\\Application\\msedge.exe",)),
    BrowserCatalogEntry("brave", "Brave", BrowserEngine.CHROMIUM,
                        ("BraveSoftware\\Brave-Browser\\Application\\brave.exe",)),
    BrowserCatalogEntry("yandex", "Yandex Browser", BrowserEngine.CHROMIUM,
                        ("Yandex\\YandexBrowser\\Application\\browser.exe",)),
    BrowserCatalogEntry("vivaldi", "Vivaldi", BrowserEngine.CHROMIUM,
                        ("Vivaldi\\Application\\vivaldi.exe",)),
    BrowserCatalogEntry("firefox", "Mozilla Firefox", BrowserEngine.GECKO,
                        ("Mozilla Firefox\\firefox.exe",)),
    BrowserCatalogEntry("waterfox", "Waterfox", BrowserEngine.GECKO,
                        ("Waterfox\\waterfox.exe",)),
    BrowserCatalogEntry("floorp", "Floorp", BrowserEngine.GECKO,
                        ("Floorp\\floorp.exe", "Ablaze Floorp\\floorp.exe")),
    BrowserCatalogEntry("librewolf", "LibreWolf", BrowserEngine.GECKO,
                        ("LibreWolf\\librewolf.exe",)),
    BrowserCatalogEntry("opera", "Opera", BrowserEngine.CHROMIUM,
                        ("Opera\\opera.exe", "Opera Software\\Opera\\opera.exe")),
    BrowserCatalogEntry("opera_gx", "Opera GX", BrowserEngine.CHROMIUM,
                        ("Opera Software\\Opera GX\\opera.exe",)),
)

_GECKO_ROAMING_FOLDER: dict[str, tuple[str, ...]] = {
    "firefox": ("Mozilla", "Firefox"),
    "waterfox": ("Waterfox",),
    "floorp": ("Floorp",),
    "librewolf": ("LibreWolf",),
}

_CHROMIUM_USER_DATA_FOLDER: dict[str, tuple[str, ...]] = {
    "chrome": ("Google", "Chrome", "User Data"),
    "edge": ("Microsoft", "Edge", "User Data"),
    "brave": ("BraveSoftware", "Brave-Browser", "User Data"),
    "yandex": ("Yandex", "YandexBrowser", "User Data"),
    "vivaldi": ("Vivaldi", "User Data"),
    "opera": ("Opera Software", "Opera", "User Data"),
    "opera_gx": ("Opera Software", "Opera GX Stable", "User Data"),
}


def _candidate_base_dirs() -> list[Path]:
    bases: list[Path] = []
    for env_key in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        val = os.environ.get(env_key)
        if val:
            bases.append(Path(val))
    return bases


def _resolve_first_existing(entry: BrowserCatalogEntry, base_dirs: list[Path]) -> Path | None:
    for base in base_dirs:
        for rel in entry.relative_exe_candidates:
            candidate = base / rel
            if candidate.is_file():
                return candidate
    return None


def find_installed_browsers() -> list[BrowserDefinition]:
    found: list[BrowserDefinition] = []
    base_dirs = _candidate_base_dirs()
    for entry in BROWSER_CATALOG:
        exe_path = _resolve_first_existing(entry, base_dirs)
        if exe_path is not None:
            found.append(BrowserDefinition(
                key=entry.key, display_name=entry.display_name,
                executable_path=exe_path, engine=entry.engine,
                auto_discovered=True,
            ))
    return found


def _discover_chromium_profiles(browser: BrowserDefinition) -> list[BrowserProfile]:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        return []
    user_data_rel = _CHROMIUM_USER_DATA_FOLDER.get(browser.key)
    if not user_data_rel:
        return []
    local_state_path = Path(local_app_data).joinpath(*user_data_rel, "Local State")
    if not local_state_path.is_file():
        return []
    data = json.loads(local_state_path.read_text(encoding="utf-8"))
    info_cache: dict[str, dict] = data.get("profile", {}).get("info_cache", {})
    profiles: list[BrowserProfile] = []
    for folder_name, info in info_cache.items():
        display_name = info.get("name") or folder_name
        profiles.append(BrowserProfile(
            id=BrowserProfile.new_id(), browser_key=browser.key,
            display_name=f"{display_name} ({folder_name})",
            profile_arg=folder_name,
        ))
    return profiles


def _read_pref_string(prefs_path: Path, pref_key: str) -> Optional[str]:
    if not prefs_path.is_file():
        return None
    try:
        text = prefs_path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return None
    marker = f'user_pref("{pref_key}", "'
    idx = text.find(marker)
    if idx == -1:
        return None
    end = text.find('"', idx + len(marker))
    if end == -1:
        return None
    value = text[idx + len(marker):end].strip()
    return value or None


def _read_profile_friendly_name(profile_dir: Path) -> Optional[str]:
    """Fallback nickname lookup used only when the Profile Groups datastore
    (see _read_profile_group_names) isn't available for some reason."""
    return _read_pref_string(profile_dir / "prefs.js", "browser.profiles.profileName")


def _read_profile_group_names(vendor_root: Path, store_id: str) -> dict[str, str]:
    """Reads {profile_path_leaf: exact_display_name} from the Firefox/
    Waterfox "Selectable Profiles" SQLite datastore for one storeID.

    This is the ONLY place the exact nickname the user typed in the
    browser's own profile switcher (the "alef", "asdasd", "be", ... names)
    actually lives. profiles.ini's Name= field is a different, older value
    and does NOT reflect what the user renamed the profile to. See:
    https://firefox-source-docs.mozilla.org/toolkit/profile/index.html

    Returns {} on any failure (db locked/missing/unexpected schema) so
    callers always have a safe fallback.
    """
    if not store_id:
        return {}
    db_path = vendor_root / "Profile Groups" / f"{store_id}.sqlite"
    if not db_path.is_file():
        return {}
    names: dict[str, str] = {}
    try:
        # Open read-only so we never fight the browser for a write lock,
        # and so a currently-running browser doesn't block us.
        uri = f"file:{db_path.as_posix()}?mode=ro"
        conn = sqlite3.connect(uri, uri=True, timeout=1.0)
        try:
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
            for (table,) in cur.fetchall():
                cur.execute(f'PRAGMA table_info("{table}")')
                cols = [row[1] for row in cur.fetchall()]
                lower = {c.lower(): c for c in cols}
                if "name" in lower and "path" in lower:
                    cur.execute(f'SELECT "{lower["name"]}", "{lower["path"]}" FROM "{table}"')
                    for name_val, path_val in cur.fetchall():
                        if name_val and path_val:
                            leaf = Path(str(path_val).replace("\\", "/")).name
                            if leaf:
                                names[leaf] = str(name_val)
        finally:
            conn.close()
    except (sqlite3.Error, OSError):
        return {}
    return names


def _discover_gecko_profiles(browser: BrowserDefinition) -> list[BrowserProfile]:
    roaming = os.environ.get("APPDATA")
    if not roaming:
        return []
    vendor_folder = _GECKO_ROAMING_FOLDER.get(browser.key)
    if not vendor_folder:
        return []
    vendor_root = Path(roaming).joinpath(*vendor_folder)
    profiles_ini_path = vendor_root / "profiles.ini"

    # Pass 1: read profiles.ini. Under the newer "selectable profiles"
    # system this typically has only ONE entry per profile *group* (its
    # Path just points at whichever profile in the group was used last),
    # plus it records the group's StoreID -- which we need for pass 1.5.
    known_by_path: dict[str, BrowserProfile] = {}
    store_id = ""
    if profiles_ini_path.is_file():
        parser = configparser.ConfigParser()
        parser.optionxform = str  # type: ignore[assignment]
        parser.read(profiles_ini_path, encoding="utf-8")
        for section in parser.sections():
            if not section.startswith("Profile"):
                continue
            if not parser.has_option(section, "Name"):
                continue
            name = parser.get(section, "Name")
            path_hint = parser.get(section, "Path", fallback="")
            is_relative = parser.get(section, "IsRelative", fallback="1") == "1"
            store_id = store_id or parser.get(section, "StoreID", fallback="")
            abs_path = (vendor_root / path_hint) if is_relative else Path(path_hint)
            key = str(abs_path.resolve()) if path_hint else f"__noname__{name}"
            known_by_path[key] = BrowserProfile(
                id=BrowserProfile.new_id(), browser_key=browser.key,
                display_name=name, profile_arg=name,
                profile_path=str(abs_path) if path_hint else None,
                notes=f"profiles.ini path: {path_hint}" if path_hint else "",
            )

    # Pass 1.5: pull the *real*, exact nicknames from the Profile Groups
    # datastore (keyed by profile folder leaf name), if this browser uses
    # the selectable-profiles feature.
    group_names = _read_profile_group_names(vendor_root, store_id)

    # Pass 2: scan the actual "Profiles" folder on disk. This is what
    # catches profiles created through a browser's newer built-in profile
    # picker (e.g. Firefox/Waterfox "selectable profiles"), which are kept
    # in a separate per-group datastore and often never get their own
    # [ProfileN] section in profiles.ini — profiles.ini alone is not a
    # reliable source of truth for how many profiles actually exist.
    profiles_root = vendor_root / "Profiles"
    if profiles_root.is_dir():
        for entry in profiles_root.iterdir():
            if not entry.is_dir():
                continue
            if not (entry / "prefs.js").is_file():
                continue
            key = str(entry.resolve())
            exact_name = group_names.get(entry.name)
            if key in known_by_path:
                # Already have it from profiles.ini -- if the datastore
                # knows its *exact* current nickname, prefer that over
                # profiles.ini's (possibly stale) Name= value.
                if exact_name:
                    existing = known_by_path[key]
                    known_by_path[key] = BrowserProfile(
                        id=existing.id, browser_key=existing.browser_key,
                        display_name=exact_name, profile_arg=exact_name,
                        profile_path=existing.profile_path, notes=existing.notes,
                    )
                continue
            friendly_name = (
                exact_name
                or _read_profile_friendly_name(entry)
                or entry.name
            )
            known_by_path[key] = BrowserProfile(
                id=BrowserProfile.new_id(), browser_key=browser.key,
                display_name=friendly_name, profile_arg=friendly_name,
                profile_path=str(entry),
                notes="discovered from Profiles folder (not in profiles.ini)",
            )

    return list(known_by_path.values())


def discover_profiles_for_browser(browser: BrowserDefinition) -> list[BrowserProfile]:
    try:
        if browser.engine == BrowserEngine.CHROMIUM:
            return _discover_chromium_profiles(browser)
        return _discover_gecko_profiles(browser)
    except (OSError, json.JSONDecodeError, configparser.Error, KeyError, ValueError):
        return []


# ══════════════════════════════════════════════════════════════════════
#  SECTION 4.5 — PROFILE CREATION
# ══════════════════════════════════════════════════════════════════════

_CREATE_LOG = logging.getLogger(__name__)

CREATE_MIN_COUNT = 1
CREATE_MAX_COUNT = 100
CREATE_CONFIRM_THRESHOLD = 25

# Chromium-family browsers all number their built-in profile avatars, but each
# vendor ships its own resource prefix:
#   Chrome / Edge / Brave : chrome://theme/IDR_PROFILE_AVATAR_N
#   Vivaldi               : chrome://theme/IDR_PROFILE_VIVALDI_AVATAR_N
#   Yandex                : chrome://theme/IDR_PROFILE_AVATAR_YANDEX_N
# The prefix is always read back out of the browser's OWN profile registry when
# it has one; these are only the fallbacks for a browser whose registry happens
# to contain no avatar-bearing profile yet.
_AVATAR_ICON_PREFIX_FALLBACK: dict[str, str] = {
    "vivaldi": "chrome://theme/IDR_PROFILE_VIVALDI_AVATAR_",
    "yandex": "chrome://theme/IDR_PROFILE_AVATAR_YANDEX_",
}
_AVATAR_ICON_PREFIX_DEFAULT = "chrome://theme/IDR_PROFILE_AVATAR_"

# The avatar set has a hole at index 26. Vivaldi's own profile picker declares
# AVATAR_COUNT = 38 with AVATAR_SKIP_INDEX = 26 ("Unused index"), and stock
# Chromium keeps the same reserved slot. A profile parked on it shows a blank
# picture, so it is always skipped.
_AVATAR_SKIP_INDEX = 26
_AVATAR_MAX_INDEX = 37

_WINDOWS_RESERVED_NAMES = frozenset((
    "con", "prn", "aux", "nul",
    "com1", "com2", "com3", "com4", "com5", "com6", "com7", "com8", "com9",
    "lpt1", "lpt2", "lpt3", "lpt4", "lpt5", "lpt6", "lpt7", "lpt8", "lpt9",
))

# Keeps the tasklist/taskkill helper windows from flashing up on every poll.
_NO_WINDOW_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0


@dataclass
class ProfileCreationPlan:
    """One profile that is about to be created (or was just created)."""

    browser_key: str
    display_name: str
    folder_name: str
    profile_arg: str
    registry_key: str = ""
    profile_path: Optional[str] = None
    avatar_index: Optional[int] = None
    color_tag: Optional[str] = None
    notes: str = ""

    def disk_path(self, browser: BrowserDefinition) -> Optional[Path]:
        if self.profile_path:
            return Path(self.profile_path)
        if browser.engine == BrowserEngine.CHROMIUM:
            user_data = chromium_user_data_dir(browser)
            if user_data is not None:
                return user_data / self.folder_name
        return None


@dataclass
class ProfileCreationResult:
    browser_key: str
    created: list[ProfileCreationPlan] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    backup_paths: list[str] = field(default_factory=list)


# ---------------------------------------------------------------- locations

def chromium_user_data_dir(browser: BrowserDefinition) -> Optional[Path]:
    """The folder that holds a Chromium browser's Local State + profile dirs."""
    if browser.engine != BrowserEngine.CHROMIUM:
        return None
    segments = _CHROMIUM_USER_DATA_FOLDER.get(browser.key)
    if not segments:
        return None
    base = os.environ.get("LOCALAPPDATA")
    if not base:
        return None
    return Path(base).joinpath(*segments)


def chromium_local_state_path(browser: BrowserDefinition) -> Optional[Path]:
    user_data = chromium_user_data_dir(browser)
    if user_data is None:
        return None
    return user_data / "Local State"


def gecko_vendor_root(browser: BrowserDefinition) -> Optional[Path]:
    """The folder that holds a Firefox-family browser's profiles.ini."""
    if browser.engine != BrowserEngine.GECKO:
        return None
    segments = _GECKO_ROAMING_FOLDER.get(browser.key)
    if not segments:
        return None
    base = os.environ.get("APPDATA")
    if not base:
        return None
    return Path(base).joinpath(*segments)


def profile_creation_support(browser: BrowserDefinition) -> tuple[bool, str]:
    """Can we create profiles for this browser? Returns (supported, reason)."""
    if browser.engine == BrowserEngine.CHROMIUM:
        if chromium_user_data_dir(browser) is None:
            return False, "Could not locate this browser's User Data folder."
        local_state = chromium_local_state_path(browser)
        if local_state is None or not local_state.is_file():
            return False, (
                "This browser keeps no profile registry (Local State), so it "
                "cannot hold several named profiles."
            )
        return True, ""
    if browser.engine == BrowserEngine.GECKO:
        if gecko_vendor_root(browser) is None:
            return False, "Could not locate this browser's profile folder."
        return True, ""
    return False, "Unsupported browser engine."


# --------------------------------------------------------------- processes

def _run_hidden(args: list[str], timeout: float = 20.0):
    """Run a console tool without flashing a window; None on any failure."""
    try:
        return subprocess.run(
            args, capture_output=True, text=True, errors="ignore",
            timeout=timeout, creationflags=_NO_WINDOW_FLAGS,
        )
    except (OSError, subprocess.SubprocessError, ValueError):
        return None


def is_browser_running(browser: BrowserDefinition) -> bool:
    if sys.platform != "win32":
        return False
    exe_name = browser.executable_path.name.strip().lower()
    if not exe_name:
        return False
    proc = _run_hidden(["tasklist", "/FI", f"IMAGENAME eq {exe_name}", "/NH"], timeout=15.0)
    if proc is None:
        return False
    # Only match a real process row. The "no tasks match" notice is localized,
    # so substring-matching the whole output would be unsafe.
    for line in (proc.stdout or "").splitlines():
        if line.strip().lower().startswith(exe_name):
            return True
    return False


def close_browser_gracefully(browser: BrowserDefinition, timeout: float = 25.0) -> bool:
    """Ask the browser to close (WM_CLOSE, never /F) and wait for it to exit.

    A plain WM_CLOSE is sometimes swallowed -- e.g. while a page is showing a
    beforeunload prompt -- so the request is re-sent a few times before giving
    up. Returns False if the browser is still running at the deadline.
    """
    if sys.platform != "win32":
        return False
    exe_name = browser.executable_path.name.strip()
    if not exe_name or not is_browser_running(browser):
        return True
    deadline = time.monotonic() + max(2.0, float(timeout))
    next_attempt = time.monotonic()
    while time.monotonic() < deadline:
        if time.monotonic() >= next_attempt:
            _run_hidden(["taskkill", "/IM", exe_name], timeout=15.0)
            next_attempt = time.monotonic() + 5.0
        if not is_browser_running(browser):
            return True
        time.sleep(0.4)
    return not is_browser_running(browser)


# ------------------------------------------------------------- file writes

def _backup_registry_file(path: Path) -> Optional[Path]:
    """Copy a registry file next to itself, timestamped. Never raises."""
    if not path.is_file():
        return None
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = path.with_name(f"{path.name}.backup-{stamp}")
    try:
        shutil.copy2(path, backup)
    except OSError:
        _CREATE_LOG.exception("Could not back up %s", path)
        return None
    return backup


def _atomic_write_text(path: Path, text: str) -> None:
    """Write via a temp file + os.replace so a crash can never leave a
    half-written profile registry behind."""
    path.parent.mkdir(parents=True, exist_ok=True)
    handle_fd, tmp_name = tempfile.mkstemp(
        dir=str(path.parent), prefix=path.name + ".", suffix=".tmp"
    )
    tmp_path = Path(tmp_name)
    try:
        with os.fdopen(handle_fd, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
        os.replace(tmp_path, path)
    except BaseException:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except OSError:
            pass
        raise


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    _atomic_write_text(path, json.dumps(payload, ensure_ascii=False))


# ------------------------------------------------------------------- names

def _sanitize_name(text: str, fallback: str = "Profile") -> str:
    cleaned = "".join(
        char for char in (text or "") if char.isprintable() and char not in r'\/:*?"<>|'
    )
    cleaned = cleaned.strip().strip(".").strip()
    if not cleaned:
        cleaned = fallback
    cleaned = cleaned[:64].strip()
    if cleaned.lower() in _WINDOWS_RESERVED_NAMES:
        cleaned = f"{cleaned}-profile"
    return cleaned or fallback


def _unique_folder_name(desired: str, taken: set[str]) -> str:
    lowered = {name.lower() for name in taken}
    candidate = desired
    suffix = 2
    while candidate.lower() in lowered and suffix < 9999:
        candidate = f"{desired} {suffix}"
        suffix += 1
    return candidate


def _unique_display_name(desired: str, used: set[str]) -> str:
    candidate = desired
    suffix = 2
    while candidate.lower() in used and suffix < 9999:
        candidate = f"{desired} {suffix}"
        suffix += 1
    return candidate


def _format_profile_name(pattern: str, number: int, browser: BrowserDefinition) -> str:
    try:
        text = pattern.format(
            n=number, number=number, browser=browser.display_name, key=browser.key
        )
    except (IndexError, KeyError, ValueError):
        text = f"{pattern} {number}"
    return _sanitize_name(text, fallback=f"Profile {number}")


def _cycle_tag(offset: int) -> str:
    palette = Tokens.TAG_PALETTE
    if not palette:
        return "#6FA8FF"
    return palette[offset % len(palette)]


# ------------------------------------------------------------- avatars

def _avatar_prefix_from_records(info_cache: dict) -> tuple[str, set[int]]:
    """Read the avatar resource prefix and the indices already in use straight
    out of the browser's own records, so this keeps working when a browser
    changes its prefix or adds avatars."""
    prefix = ""
    used: set[int] = set()
    for record in info_cache.values():
        if not isinstance(record, dict):
            continue
        raw = record.get("avatar_icon")
        if not isinstance(raw, str):
            continue
        head, _, tail = raw.rpartition("_")
        if not head or not tail.isdigit():
            continue
        candidate = head + "_"
        if not prefix:
            prefix = candidate
        if candidate == prefix:
            used.add(int(tail))
    return prefix, used


def avatar_prefix_for_browser(browser: BrowserDefinition, info_cache: Optional[dict] = None) -> str:
    if info_cache:
        prefix, _used = _avatar_prefix_from_records(info_cache)
        if prefix:
            return prefix
    return _AVATAR_ICON_PREFIX_FALLBACK.get(browser.key, _AVATAR_ICON_PREFIX_DEFAULT)


def _avatar_icon_url(index: Optional[int], prefix: str) -> Optional[str]:
    if index is None:
        return None
    return f"{prefix}{int(index)}"


def _read_chromium_info_cache(browser: BrowserDefinition) -> dict:
    local_state = chromium_local_state_path(browser)
    if local_state is None or not local_state.is_file():
        return {}
    try:
        data = json.loads(local_state.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    info_cache = data.get("profile", {}).get("info_cache", {})
    return info_cache if isinstance(info_cache, dict) else {}


def _spread_indices(pool: list[int], count: int) -> list[int]:
    """Pick `count` entries out of `pool`, spread evenly rather than clumped
    at the front."""
    if not pool or count <= 0:
        return []
    if count <= len(pool):
        step = len(pool) / float(count)
        return [pool[min(int(index * step), len(pool) - 1)] for index in range(count)]
    return [pool[index % len(pool)] for index in range(count)]


def _next_chromium_folder_index(info_cache: dict) -> int:
    highest = 0
    for key in info_cache:
        if not isinstance(key, str) or not key.startswith("Profile "):
            continue
        tail = key[len("Profile "):].strip()
        if tail.isdigit():
            highest = max(highest, int(tail))
    return highest + 1


# ------------------------------------------------------- gecko registry io

def _read_gecko_profiles_ini(ini_path: Path) -> configparser.ConfigParser:
    parser = configparser.ConfigParser()
    parser.optionxform = str  # type: ignore[assignment]
    if ini_path.is_file():
        try:
            parser.read(ini_path, encoding="utf-8")
        except (configparser.Error, OSError):
            _CREATE_LOG.exception("Could not parse %s", ini_path)
            return configparser.ConfigParser()
    return parser


def _existing_gecko_profile_names(ini_path: Path) -> set[str]:
    parser = _read_gecko_profiles_ini(ini_path)
    names: set[str] = set()
    for section in parser.sections():
        if section.startswith("Profile") and parser.has_option(section, "Name"):
            names.add(parser.get(section, "Name").strip().lower())
    return names


def _existing_gecko_folder_leaves(vendor_root: Path, ini_path: Path) -> set[str]:
    leaves: set[str] = set()
    parser = _read_gecko_profiles_ini(ini_path)
    for section in parser.sections():
        if not section.startswith("Profile"):
            continue
        path_hint = parser.get(section, "Path", fallback="")
        if path_hint:
            leaves.add(Path(path_hint.replace("\\", "/")).name)
    profiles_root = vendor_root / "Profiles"
    if profiles_root.is_dir():
        try:
            for entry in profiles_root.iterdir():
                if entry.is_dir():
                    leaves.add(entry.name)
        except OSError:
            pass
    return leaves


def _next_gecko_section_index(parser: configparser.ConfigParser) -> int:
    highest = -1
    for section in parser.sections():
        if section.startswith("Profile") and section[len("Profile"):].isdigit():
            highest = max(highest, int(section[len("Profile"):]))
    return highest + 1


# ------------------------------------------------------------------ planning

def build_creation_plan(
    browser: BrowserDefinition,
    *,
    count: int,
    name_pattern: str = "Profile {n}",
    start_number: int = 1,
    avatar_mode: str = "auto",
    avatar_start: int = 0,
    color_cycle: bool = False,
    notes: str = "",
) -> tuple[list[ProfileCreationPlan], list[str]]:
    """Work out exactly which profiles WOULD be created. Touches nothing on
    disk. Returns (plans, warnings)."""
    warnings: list[str] = []
    try:
        wanted = int(count)
    except (TypeError, ValueError):
        wanted = CREATE_MIN_COUNT
    wanted = max(CREATE_MIN_COUNT, min(wanted, CREATE_MAX_COUNT))

    supported, reason = profile_creation_support(browser)
    if not supported:
        return [], [reason]

    pattern = (name_pattern or "Profile {n}").strip() or "Profile {n}"
    plans: list[ProfileCreationPlan] = []

    if browser.engine == BrowserEngine.CHROMIUM:
        info_cache = _read_chromium_info_cache(browser)
        _prefix, used_indices = _avatar_prefix_from_records(info_cache)
        start_index = _next_chromium_folder_index(info_cache)
        taken = set(info_cache)

        if avatar_mode == "start":
            pool = [
                index for index in range(max(0, int(avatar_start)), _AVATAR_MAX_INDEX + 1)
                if index != _AVATAR_SKIP_INDEX
            ]
            avatar_indices: list[Optional[int]] = list(_spread_indices(pool, wanted))
        elif avatar_mode == "default":
            avatar_indices = [None] * wanted
        else:
            pool = [
                index for index in range(0, _AVATAR_MAX_INDEX + 1)
                if index != _AVATAR_SKIP_INDEX and index not in used_indices
            ]
            if pool:
                avatar_indices = list(_spread_indices(pool, wanted))
            else:
                warnings.append(
                    "Every built-in picture is already in use, so the new "
                    "profiles will keep the browser's default one."
                )
                avatar_indices = [None] * wanted

        for offset in range(wanted):
            number = start_number + offset
            folder = _unique_folder_name(f"Profile {start_index + offset}", taken)
            taken.add(folder)
            plans.append(ProfileCreationPlan(
                browser_key=browser.key,
                display_name=_format_profile_name(pattern, number, browser),
                folder_name=folder,
                profile_arg=folder,
                registry_key=folder,
                avatar_index=avatar_indices[offset] if offset < len(avatar_indices) else None,
                color_tag=_cycle_tag(offset) if color_cycle else None,
                notes=notes,
            ))
        return plans, warnings

    # Gecko: profiles.ini + a folder under Profiles/
    vendor_root = gecko_vendor_root(browser)
    if vendor_root is None:
        return [], ["Could not locate this browser's profile folder."]
    ini_path = vendor_root / "profiles.ini"
    taken = _existing_gecko_folder_leaves(vendor_root, ini_path)
    used_names = _existing_gecko_profile_names(ini_path)

    for offset in range(wanted):
        number = start_number + offset
        display_name = _format_profile_name(pattern, number, browser)
        if display_name.lower() in used_names:
            original = display_name
            display_name = _unique_display_name(original, used_names)
            warnings.append(
                f"'{original}' already exists in {browser.display_name}; "
                f"using '{display_name}' instead."
            )
        used_names.add(display_name.lower())
        folder = _unique_folder_name(_sanitize_name(display_name), taken)
        taken.add(folder)
        plans.append(ProfileCreationPlan(
            browser_key=browser.key,
            display_name=display_name,
            folder_name=folder,
            profile_arg=display_name,
            profile_path=str(vendor_root / "Profiles" / folder),
            avatar_index=None,
            color_tag=_cycle_tag(offset) if color_cycle else None,
            notes=notes,
        ))
    return plans, warnings


# ------------------------------------------------------------------ creating

def create_chromium_profiles(
    browser: BrowserDefinition, plans: list[ProfileCreationPlan]
) -> ProfileCreationResult:
    result = ProfileCreationResult(browser_key=browser.key)
    user_data = chromium_user_data_dir(browser)
    local_state = chromium_local_state_path(browser)
    if user_data is None or local_state is None or not local_state.is_file():
        result.errors.append("Could not locate this browser's profile registry (Local State).")
        return result

    try:
        data = json.loads(local_state.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        result.errors.append(f"Could not read Local State: {exc}")
        return result

    profile_section = data.get("profile")
    if not isinstance(profile_section, dict):
        result.errors.append("This browser's Local State has no profile section.")
        return result
    info_cache = profile_section.get("info_cache")
    if not isinstance(info_cache, dict) or not info_cache:
        result.errors.append("This browser's Local State has no profile registry (info_cache).")
        return result

    # Clone a record the browser itself wrote. Every vendor stores its own
    # extra fields (Edge's edge_*, Chrome's gaia_picture_file_name and
    # profile_color_seed, ...) and building a record from scratch would risk
    # handing a browser fields it does not understand.
    template_key = "Default" if "Default" in info_cache else next(iter(info_cache))
    template = info_cache.get(template_key)
    if not isinstance(template, dict) or not template:
        result.errors.append("This browser's profile registry is malformed.")
        return result

    # Pre-flight: refuse rather than half-apply if any target is already taken.
    clashes = [
        plan.folder_name for plan in plans
        if plan.folder_name in info_cache or (user_data / plan.folder_name).exists()
    ]
    if clashes:
        result.errors.append(
            "These profile folders already exist, so nothing was changed: "
            + ", ".join(sorted(clashes))
        )
        return result

    backup = _backup_registry_file(local_state)
    if backup is not None:
        result.backup_paths.append(str(backup))

    prefix, _used = _avatar_prefix_from_records(info_cache)
    if not prefix:
        prefix = _AVATAR_ICON_PREFIX_FALLBACK.get(browser.key, _AVATAR_ICON_PREFIX_DEFAULT)

    order = profile_section.get("profiles_order")
    if not isinstance(order, list):
        order = list(info_cache)
    metrics = profile_section.get("metrics")
    if not isinstance(metrics, dict):
        metrics = {}
    bucket = metrics.get("next_bucket_index")
    if not isinstance(bucket, int):
        bucket = len(info_cache) + 1
    now = time.time()

    for plan in plans:
        record = dict(template)
        record["name"] = plan.display_name
        record["shortcut_name"] = plan.display_name
        record["metrics_bucket_index"] = bucket
        record["active_time"] = now
        record["is_using_default_name"] = False
        record["is_ephemeral"] = False
        record["background_apps"] = False
        icon_url = _avatar_icon_url(plan.avatar_index, prefix)
        if icon_url is not None:
            record["avatar_icon"] = icon_url
            # False, else Chrome/Edge/Brave draw the letter-circle placeholder
            # instead of the picture we just picked.
            record["is_using_default_avatar"] = False
        info_cache[plan.folder_name] = record
        if plan.folder_name not in order:
            order.append(plan.folder_name)
        bucket += 1

        profile_dir = user_data / plan.folder_name
        try:
            profile_dir.mkdir(parents=True, exist_ok=True)
            preferences: dict[str, Any] = {}
            if plan.avatar_index is not None:
                preferences["profile"] = {"avatar_index": int(plan.avatar_index)}
            _atomic_write_json(profile_dir / "Preferences", preferences)
        except OSError as exc:
            result.errors.append(f"Could not prepare the folder for '{plan.display_name}': {exc}")

    profile_section["info_cache"] = info_cache
    profile_section["profiles_order"] = order
    metrics["next_bucket_index"] = bucket
    profile_section["metrics"] = metrics
    if "profiles_created" in profile_section:
        profile_section["profiles_created"] = len(order)
    data["profile"] = profile_section

    try:
        _atomic_write_json(local_state, data)
    except OSError as exc:
        result.errors.append(f"Could not write Local State: {exc}")
        return result

    verified = _verify_chromium_folders(local_state, [p.folder_name for p in plans])
    for plan in plans:
        if plan.folder_name in verified:
            result.created.append(plan)
        else:
            result.errors.append(
                f"'{plan.display_name}' was written but could not be verified afterwards."
            )
    return result


def _verify_chromium_folders(local_state: Path, folder_names: list[str]) -> set[str]:
    """Re-read the file we just wrote and confirm the entries really landed."""
    try:
        data = json.loads(local_state.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return set()
    info_cache = data.get("profile", {}).get("info_cache", {})
    if not isinstance(info_cache, dict):
        return set()
    return {name for name in folder_names if name in info_cache}


def create_gecko_profiles(
    browser: BrowserDefinition, plans: list[ProfileCreationPlan]
) -> ProfileCreationResult:
    result = ProfileCreationResult(browser_key=browser.key)
    vendor_root = gecko_vendor_root(browser)
    if vendor_root is None:
        result.errors.append("Could not locate this browser's profile folder.")
        return result
    ini_path = vendor_root / "profiles.ini"
    parser = _read_gecko_profiles_ini(ini_path)

    existing_names = _existing_gecko_profile_names(ini_path)
    clashes = [plan.display_name for plan in plans if plan.display_name.lower() in existing_names]
    if clashes:
        result.errors.append(
            "These profile names already exist, so nothing was changed: "
            + ", ".join(sorted(clashes))
        )
        return result

    try:
        (vendor_root / "Profiles").mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        result.errors.append(f"Could not create the Profiles folder: {exc}")
        return result

    # Folders first: a profiles.ini entry pointing at a missing folder makes
    # the browser fail loudly on launch, the reverse is harmless.
    ready: list[ProfileCreationPlan] = []
    for plan in plans:
        folder = vendor_root / "Profiles" / plan.folder_name
        try:
            folder.mkdir(parents=True, exist_ok=True)
            ready.append(plan)
        except OSError as exc:
            result.errors.append(f"Could not create the folder for '{plan.display_name}': {exc}")
    if not ready:
        return result

    backup = _backup_registry_file(ini_path)
    if backup is not None:
        result.backup_paths.append(str(backup))

    section_index = _next_gecko_section_index(parser)
    block: list[str] = []
    for plan in ready:
        while f"Profile{section_index}" in parser:
            section_index += 1
        section = f"Profile{section_index}"
        section_index += 1
        plan.registry_key = section
        block.extend((
            f"[{section}]",
            f"Name={plan.display_name}",
            "IsRelative=1",
            f"Path=Profiles/{plan.folder_name}",
            "",
        ))

    try:
        original = ini_path.read_text(encoding="utf-8", errors="ignore") if ini_path.is_file() else ""
    except OSError as exc:
        result.errors.append(f"Could not read profiles.ini: {exc}")
        return result
    if original and not original.endswith("\n"):
        original += "\n"
    # Appending rather than rewriting: not one existing line is touched.
    prefix_text = original + ("\n" if original.strip() else "")
    try:
        _atomic_write_text(ini_path, prefix_text + "\n".join(block).rstrip("\n") + "\n")
    except OSError as exc:
        result.errors.append(f"Could not write profiles.ini: {exc}")
        return result

    verified = _verify_gecko_sections(ini_path, [p.registry_key for p in ready])
    for plan in ready:
        if plan.registry_key in verified:
            result.created.append(plan)
        else:
            result.errors.append(
                f"'{plan.display_name}' was written but could not be verified afterwards."
            )
    return result


def _verify_gecko_sections(ini_path: Path, section_keys: list[str]) -> set[str]:
    parser = _read_gecko_profiles_ini(ini_path)
    return {
        key for key in section_keys
        if key and parser.has_section(key) and parser.has_option(key, "Path")
    }


def create_profiles(
    browser: BrowserDefinition, plans: list[ProfileCreationPlan]
) -> ProfileCreationResult:
    if not plans:
        return ProfileCreationResult(browser_key=browser.key, errors=["Nothing to create."])
    if browser.engine == BrowserEngine.CHROMIUM:
        return create_chromium_profiles(browser, plans)
    return create_gecko_profiles(browser, plans)


def remove_created_profiles(
    browser: BrowserDefinition, plans: list[ProfileCreationPlan]
) -> list[str]:
    """Best-effort rollback of create_profiles(). Drops the registry entries
    and deletes the folders. Returns notes about anything left behind."""
    notes: list[str] = []
    if not plans:
        return notes
    if is_browser_running(browser):
        return [f"{browser.display_name} is running — close it before undoing."]

    keys = {plan.folder_name for plan in plans}

    if browser.engine == BrowserEngine.CHROMIUM:
        local_state = chromium_local_state_path(browser)
        user_data = chromium_user_data_dir(browser)
        if local_state is None or user_data is None or not local_state.is_file():
            return ["Could not locate this browser's profile registry."]
        try:
            data = json.loads(local_state.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return [f"Could not read Local State: {exc}"]
        profile_section = data.get("profile", {})
        info_cache = profile_section.get("info_cache", {})
        if isinstance(info_cache, dict):
            for name in keys:
                info_cache.pop(name, None)
            order = profile_section.get("profiles_order")
            if isinstance(order, list):
                profile_section["profiles_order"] = [
                    name for name in order if name not in keys
                ]
            try:
                _atomic_write_json(local_state, data)
            except OSError as exc:
                notes.append(f"Could not update Local State: {exc}")
        for plan in plans:
            folder = user_data / plan.folder_name
            try:
                if folder.is_dir():
                    shutil.rmtree(folder)
            except OSError as exc:
                notes.append(f"Could not delete {plan.folder_name}: {exc}")
        return notes

    vendor_root = gecko_vendor_root(browser)
    if vendor_root is None:
        return ["Could not locate this browser's profile folder."]
    ini_path = vendor_root / "profiles.ini"
    section_keys = {plan.registry_key for plan in plans if plan.registry_key}
    try:
        text = ini_path.read_text(encoding="utf-8", errors="ignore") if ini_path.is_file() else ""
    except OSError as exc:
        return [f"Could not read profiles.ini: {exc}"]

    kept: list[str] = []
    dropping = False
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            dropping = stripped[1:-1] in section_keys
        if not dropping:
            kept.append(line)
    try:
        _atomic_write_text(ini_path, "\n".join(kept).rstrip("\n") + "\n")
    except OSError as exc:
        notes.append(f"Could not update profiles.ini: {exc}")

    for plan in plans:
        folder = vendor_root / "Profiles" / plan.folder_name
        try:
            if folder.is_dir():
                shutil.rmtree(folder)
        except OSError as exc:
            notes.append(f"Could not delete {plan.folder_name}: {exc}")
    return notes


# ---------------------------------------------------------------- threading
#
# `signals` is always parented to a long-lived GUI-thread QObject (the dialog
# that started the work). QThreadPool auto-deletes the QRunnable on the worker
# thread the instant run() returns, so the QObject carrying the result must not
# be owned by the task -- see the note on LaunchEngine.launch_profile.

class _CloseBrowserSignals(QObject):
    finished = Signal(bool, str)


class _CloseBrowserTask(QRunnable):
    def __init__(self, signals: _CloseBrowserSignals, browser: BrowserDefinition,
                 timeout: float = 25.0) -> None:
        super().__init__()
        self._signals = signals
        self._browser = browser
        self._timeout = timeout

    def run(self) -> None:
        ok = False
        try:
            ok = close_browser_gracefully(self._browser, self._timeout)
        except Exception:
            _CREATE_LOG.exception("Closing %s failed", self._browser.display_name)
        try:
            self._signals.finished.emit(ok, self._browser.display_name)
        except RuntimeError:
            # The dialog went away underneath us; nothing left to notify.
            pass


class _CreateProfilesSignals(QObject):
    finished = Signal(object)


class _CreateProfilesTask(QRunnable):
    def __init__(self, signals: _CreateProfilesSignals, browser: BrowserDefinition,
                 plans: list[ProfileCreationPlan]) -> None:
        super().__init__()
        self._signals = signals
        self._browser = browser
        self._plans = plans

    def run(self) -> None:
        try:
            result = create_profiles(self._browser, self._plans)
        except Exception as exc:
            _CREATE_LOG.exception("Creating profiles for %s failed", self._browser.display_name)
            result = ProfileCreationResult(
                browser_key=self._browser.key,
                errors=[f"Unexpected error: {exc}"],
            )
        try:
            self._signals.finished.emit(result)
        except RuntimeError:
            pass


# ══════════════════════════════════════════════════════════════════════
#  SECTION 5 — LAUNCH ENGINE
# ══════════════════════════════════════════════════════════════════════

if sys.platform == "win32":
    _DETACHED_PROCESS = 0x00000008
    _CREATE_NEW_PROCESS_GROUP = 0x00000200
else:
    _DETACHED_PROCESS = 0
    _CREATE_NEW_PROCESS_GROUP = 0


@dataclass
class LaunchResult:
    item_id: str
    display_name: str
    success: bool
    error_message: str = ""
    pid: int | None = None

    @property
    def profile_id(self) -> str:
        return self.item_id

    @property
    def profile_display_name(self) -> str:
        return self.display_name


class _LaunchSignals(QObject):
    result_ready = Signal(object)


class _BrowserLaunchTask(QRunnable):
    def __init__(self, browser: BrowserDefinition, profile: BrowserProfile,
                 signals: _LaunchSignals, url: str | None = None) -> None:
        super().__init__()
        self._browser = browser
        self._profile = profile
        self._signals = signals
        self._url = url
        self.setAutoDelete(True)

    def run(self) -> None:
        result = _do_browser_launch(self._browser, self._profile, url=self._url)
        self._signals.result_ready.emit(result)


def normalize_launch_url(text: str | None) -> str | None:
    """Turn the 'URL on launch' box text into something a browser can open,
    or None when the feature is off (box empty).

    - already has a scheme (https://, ...)  -> unchanged
    - looks like a domain (no spaces, has a dot) -> https:// prefix
    - anything else (plain words) -> a Google search for those words
    """
    if not text:
        return None
    t = text.strip()
    if not t:
        return None
    if "://" in t:
        return t
    if " " not in t and "." in t:
        return "https://" + t
    return "https://www.google.com/search?q=" + urllib.parse.quote_plus(t)


def build_launch_args(browser: BrowserDefinition, profile: BrowserProfile,
                      remote_debugging_port: int | None = None,
                      url: str | None = None) -> list[str]:
    """Build the command line for launching a browser profile.

    When `remote_debugging_port` is given (Web Checks), the Chromium instance
    is launched with --remote-debugging-port so Playwright can attach over CDP
    and see the user's existing login session. This is only meaningful for
    Chromium engines; for Gecko we simply ignore it (returning the normal args)
    because Web Checks doesn't support Gecko.

    When `url` is given, it is appended so the browser opens it — as a new tab
    in an already-running instance, or as the start page on a cold launch.
    """
    exe = str(browser.executable_path)
    if browser.engine == BrowserEngine.CHROMIUM:
        args = [exe, f"--profile-directory={profile.profile_arg}"]
        if remote_debugging_port is not None:
            args.append(f"--remote-debugging-port={remote_debugging_port}")
        if url:
            args.append(url)
        return args
    if browser.engine == BrowserEngine.GECKO:
        if profile.profile_path:
            args = [exe, "-no-remote", "-profile", profile.profile_path]
        else:
            args = [exe, "-no-remote", "-P", profile.profile_arg]
        if url:
            args.append(url)
        return args
    raise ValueError(f"Unsupported engine: {browser.engine!r}")


def _popen_kwargs() -> dict:
    if sys.platform == "win32":
        return {
            "creationflags": _DETACHED_PROCESS | _CREATE_NEW_PROCESS_GROUP,
            "close_fds": True,
            "stdin": subprocess.DEVNULL,
            "stdout": subprocess.DEVNULL,
            "stderr": subprocess.DEVNULL,
        }
    return {
        "start_new_session": True,
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.DEVNULL,
        "stderr": subprocess.DEVNULL,
    }


def _automation_user_data_dir(browser_key: str, profile_id: str) -> Path:
    """Dedicated persistent user-data dir for Web-Check automation of one
    browser profile.

    Modern Chromium (136+) ignores --remote-debugging-port when launched on the
    DEFAULT user-data dir, so automation must use its own folder. This one is
    persistent (kept in the app-data dir) so cookies/logins survive between
    runs: the user logs in once per site inside the automation window, and
    every later Run reuses that session.
    """
    root = get_app_data_dir() / "webcheck_data"
    safe = f"{browser_key}_{profile_id}".replace(" ", "_")
    path = root / safe
    try:
        path.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return path


def find_free_port() -> int:
    """Return a free TCP port on localhost. Used for a per-run distinct
    --remote-debugging-port so concurrent Web Checks never collide."""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def cd_endpoint_is_ready(port: int, timeout_seconds: float = 8.0) -> bool:
    """Wait until the browser's DevTools endpoint on `port` serves /json/version.

    A bare TCP connect is not enough: any process that happens to hold the port
    satisfies it, and attaching to the wrong one fails much later inside
    Playwright with a confusing error. The endpoint has to answer with a real
    DevTools version document, so we know it is actually the browser.
    """
    import json as _json
    import urllib.request
    deadline = time.time() + timeout_seconds
    url = f"http://127.0.0.1:{port}/json/version"
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1.5) as response:
                payload = _json.loads(response.read().decode("utf-8", "replace"))
            if isinstance(payload, dict) and payload.get("webSocketDebuggerUrl"):
                return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def _do_browser_launch(browser: BrowserDefinition, profile: BrowserProfile,
                       url: str | None = None) -> LaunchResult:
    log = logging.getLogger(__name__)
    if not browser.is_valid():
        log.error("Browser executable not found: %s", browser.executable_path)
        return LaunchResult(item_id=profile.id, display_name=profile.display_name,
                            success=False, error_message=f"Browser executable not found: {browser.executable_path}")
    try:
        args = build_launch_args(browser, profile, url=url)
        process = subprocess.Popen(args, **_popen_kwargs())
        log.info("Launched browser %s (pid=%s) for profile %s",
                 browser.display_name, process.pid, profile.display_name)
        return LaunchResult(item_id=profile.id, display_name=profile.display_name,
                            success=True, pid=process.pid)
    except FileNotFoundError:
        log.error("Executable not found for profile %s: %s", profile.display_name, browser.executable_path)
        return LaunchResult(item_id=profile.id, display_name=profile.display_name,
                            success=False, error_message=f"Executable not found: {browser.executable_path}")
    except PermissionError:
        log.error("Permission denied launching %s (profile %s)", browser.executable_path, profile.display_name)
        return LaunchResult(item_id=profile.id, display_name=profile.display_name,
                            success=False, error_message="Permission denied. Check antivirus or run as the same user who installed the browser.")
    except OSError as exc:
        log.error("OSError launching %s (profile %s): %s", browser.executable_path, profile.display_name, exc)
        return LaunchResult(item_id=profile.id, display_name=profile.display_name,
                            success=False, error_message=str(exc))


class LaunchEngine(QObject):
    launch_completed = Signal(object)

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._pool = QThreadPool.globalInstance()
        # Optional URL every browser-profile launch should open (set from the
        # browser grid's 'URL on launch' box). Read at each launch_profile()
        # call, so single, bulk and workspace launches all pick it up live.
        self.launch_url: str | None = None

    def launch_profile(self, browser: BrowserDefinition, profile: BrowserProfile) -> None:
        # `signals` is parented to `self` (LaunchEngine), a QObject that lives
        # in the GUI thread for the whole app lifetime. This is what actually
        # keeps it alive. QThreadPool auto-deletes `task` on the *worker*
        # thread the instant run() returns, so a reference stashed only on
        # `task` (e.g. the old task._signals_ref) can vanish right then — and
        # destroying that QObject from the wrong thread while its
        # QueuedConnection emit is still in flight to the GUI thread is what
        # caused the silent/native crashes during multi-profile bulk launches.
        signals = _LaunchSignals(self)
        signals.result_ready.connect(self.launch_completed.emit, Qt.ConnectionType.QueuedConnection)
        signals.result_ready.connect(lambda _r, s=signals: s.deleteLater())
        task = _BrowserLaunchTask(browser, profile, signals, url=self.launch_url)
        self._pool.start(task)

    def launch_profiles_batch(self, pairs: list[tuple[BrowserDefinition, BrowserProfile]]) -> None:
        for browser, profile in pairs:
            self.launch_profile(browser, profile)


# ══════════════════════════════════════════════════════════════════════
#  SECTION 5.5 — AUTO DISCOVERY (background task)
# ══════════════════════════════════════════════════════════════════════

@dataclass
class _DiscoveryResult:
    found_browsers: list[BrowserDefinition]
    profiles_by_browser: dict[str, list[BrowserProfile]]


class _DiscoverySignals(QObject):
    result_ready = Signal(object)


class _DiscoveryTask(QRunnable):
    """Scans for installed browsers and their profiles on a worker thread so
    the UI thread never blocks on disk I/O / SQLite while the app opens."""

    def __init__(self, signals: _DiscoverySignals) -> None:
        super().__init__()
        self._signals = signals
        self.setAutoDelete(True)

    def run(self) -> None:
        found_browsers = find_installed_browsers()
        profiles_by_browser: dict[str, list[BrowserProfile]] = {}
        for browser in found_browsers:
            try:
                profiles_by_browser[browser.key] = discover_profiles_for_browser(browser)
            except Exception:  # pragma: no cover - defensive
                profiles_by_browser[browser.key] = []
        try:
            self._signals.result_ready.emit(
                _DiscoveryResult(found_browsers, profiles_by_browser)
            )
        except RuntimeError:  # pragma: no cover - shutdown race
            # The app is going away and the receiver is already gone; there is
            # nobody left to hand the scan result to.
            logging.debug("Discovery finished during shutdown; result dropped.")


# ══════════════════════════════════════════════════════════════════════
#  SECTION 5.6 — WEB CHECKS AUTOMATION
# ══════════════════════════════════════════════════════════════════════
#
# IMPORTANT — PERSONAL USE ONLY
# ------------------------------
# These routines automate the user's OWN browser session to check for logged-in
# vouchers (Snapp) and discount codes (Okala). They're deliberately kept simple
# and are NOT intended for mass scraping, rate-limit bypass, resale, or anything
# that breaks the sites' terms of service. The user is responsible for how they
# use this.
#
# NOTE ON SELECTORS: The DOM selectors in the *_extract / *_detect helpers below
# are HEURISTICS. Snapp and Okala change markup and will likely need periodic
# tuning. They are written with multiple fallbacks and strict error handling so
# that a layout change produces a readable "page structure not recognized"
# message instead of a crash or a silent false-negative. These functions WILL
# need real-site testing and adjustment.

@dataclass
class WebCheckResult:
    task_id: str
    kind: WebCheckKind
    profile_id: str
    display_name: str
    success: bool
    session_active: bool | None = None
    summary: str = ""
    details: list[dict] = field(default_factory=list)
    error_message: str = ""


class _WebCheckSignals(QObject):
    result_ready = Signal(object)


class WebCheckPageNotReady(RuntimeError):
    """Raised when a target page never finished rendering.

    Both target sites are client-rendered: the HTML that arrives with the
    document is an empty shell and the real content is painted by JavaScript
    afterwards. Reading that shell and reporting the result is worse than
    reporting nothing -- it told a logged-OUT user they were logged in. So this
    is surfaced as a check failure instead of a guess.

    See the render-wait below; `test_webcheck_behavior.py` pins the behaviour
    down with a local page that paints after the load event, exactly like the
    real sites do.
    """


# A page counts as rendered once its visible text is non-empty and has stopped
# changing, OR once it is clearly substantial. Two independent escape hatches
# because a page with a live countdown (voucher timers) never goes fully still.
_PAGE_READY_POLL_S = 0.2
_PAGE_READY_STABLE_POLLS = 3
_PAGE_READY_BIG_CHARS = 300
_PAGE_READY_BIG_POLLS = 2
# Hard ceiling for the WHOLE wait, including the networkidle attempt below.
_PAGE_READY_BUDGET_S = 20.0
_PAGE_READY_NETWORKIDLE_MS = 6000


def _visible_body_text(page) -> str:
    """Rendered text of <body>, or "" if the page is not answering yet."""
    try:
        return (page.inner_text("body", timeout=3000) or "").strip()
    except Exception:
        return ""


def _wait_for_page_ready(page, log=None) -> str:
    """Block until the page has painted and return its visible body text.

    Waits for the network to go quiet first (bounded and best-effort: a page
    that polls never reaches `networkidle`), then polls the body until the text
    is non-empty AND unchanged for a few polls in a row, or is clearly
    substantial. Raises :class:`WebCheckPageNotReady` if that never happens.
    """
    started = time.time()
    deadline = started + _PAGE_READY_BUDGET_S

    # A short, bounded chance for the network to go quiet first: an SSR shell
    # that hydrates later would otherwise look "stable" the moment it arrives.
    # Pages with pollers/timers never reach networkidle, so this is capped (and
    # counted against the total budget) rather than allowed to run long.
    try:
        page.wait_for_load_state(
            "networkidle",
            timeout=max(500, min(_PAGE_READY_NETWORKIDLE_MS,
                                 (deadline - time.time()) * 1000)),
        )
    except Exception:
        pass                      # polling pages never go idle; not fatal

    text = ""
    stable = big = 0
    while time.time() < deadline:
        current = _visible_body_text(page)
        if current and current == text:
            stable += 1
        elif current:
            stable = 1 if current.startswith(text[:40]) else 0
        else:
            stable = 0
        big = big + 1 if len(current) >= _PAGE_READY_BIG_CHARS else 0
        if stable >= _PAGE_READY_STABLE_POLLS or big >= _PAGE_READY_BIG_POLLS:
            if log is not None:
                log.info("Web check: page painted after %.1fs (%d chars of text)",
                         time.time() - started, len(current))
            return current
        text = current
        time.sleep(_PAGE_READY_POLL_S)

    raise WebCheckPageNotReady(
        f"The page never finished rendering ({time.time() - started:.0f}s, "
        f"{len(text)} chars of text). The site may be slow, blocked, or stuck "
        f"on a loading screen -- see app.log."
    )


def _leaf_matches(page, selector: str) -> list:
    """Matches of *selector* that do NOT themselves contain another match.

    `[class*='voucher']` matches the list wrapper as well as every card inside
    it; keeping the wrapper produced a single "voucher" whose text was the whole
    panel -- a false positive. Only leaf matches are treated as cards.
    """
    try:
        candidates = page.locator(selector).all()
    except Exception:
        return []
    leaves = []
    for candidate in candidates:
        try:
            if candidate.locator(selector).count() > 0:
                continue          # a container, not a card
        except Exception:
            pass
        leaves.append(candidate)
    return leaves


def _first_cards(page, selectors) -> tuple[list, str]:
    """First selector in *selectors* that yields at least one leaf card."""
    for selector in selectors:
        cards = _leaf_matches(page, selector)
        if cards:
            return cards, selector
    return [], ""


SNAPP_VOUCHER_SELECTORS = (
    "[data-testid*='voucher']",
    "[class*='voucher']",
    "[class*='Voucher']",
    "[class*='coupon']",
    "[class*='discount']",
)

OKALA_COUPON_SELECTORS = (
    "[class*='discount']",
    "[class*='coupon']",
    "[class*='off']",
)

# Texts that only appear inside Okala's signed-in account area. Any of them
# means the session is alive; their absence means it is not.
OKALA_ACCOUNT_TEXTS = (
    "کدهای تخفیف من",
    "کد تخفیف من",
    "کدهای تخفیف",
    "خروج از حساب",
    "تخفیف‌های من",
    "سفارش‌های من",
    "اطلاعات حساب",
)


def _okala_account_text(body: str) -> str:
    """The account-area text present on the page, or "" if there is none.

    Deliberately NOT "is there a mobile number on the page": an online shop
    prints its own support numbers in the footer of every page, so that test
    reported a signed-out visitor as logged in. Only account-specific UI counts.
    """
    for text in OKALA_ACCOUNT_TEXTS:
        if text in (body or ""):
            return text
    return ""


def _run_snapp_vouchers(page) -> tuple[bool, list[dict], str]:
    """Detect login and scrape Snapp vouchers.

    Heuristic: if the page shows a login button / phone-number form the session
    is inactive. Otherwise we assume the user is logged in (post-logout Snapp
    typically redirects to an auth screen, and the vouchers page is only shown
    to authenticated users) and collect visible voucher cards.

    Each voucher is recorded with a `service_name` read from the DOM label near
    the card (e.g. "اسنپ‌فود") when present, otherwise "عمومی / نامشخص".
    """
    log = logging.getLogger(__name__)

    # Pattern A: explicit login affordance. Phrases verified against the real
    # logged-out Snapp vouchers page (2026-09); generic "ورود" is deliberately
    # excluded to avoid false logout reports on logged-in pages.
    login_texts = (
        "موبایلتان",          # "شماره‌‌ٔ موبایلتان را وارد کنید"
        "شمارهٔ موبایل",       # input label
        "وارد کنید تا",       # login prompt
        "ثبت‌نام در اسنپ",     # signup note under the form
        "وارد شو",            # login button (older layout)
    )
    # The page is a client-rendered app: read it only once it has actually
    # painted, otherwise every signal below is read off an empty shell (which
    # reported a logged-OUT user as logged in).
    body = _wait_for_page_ready(page, log)
    login_hit = False
    for t in login_texts:
        if t in body:
            login_hit = True
            break
    # Pattern B: phone-input field present.
    try:
        if page.locator("input[type='tel']").count() > 0 or \
           page.locator("input[placeholder*='موبایل']").count() > 0 or \
           page.locator("input[placeholder*='شماره']").count() > 0:
            login_hit = True
    except Exception:
        pass

    if login_hit:
        return False, [], "Session inactive (login required)"

    details: list[dict] = []
    # Try a handful of likely selector patterns for voucher cards. Leaf matches
    # only: a container that holds every card would otherwise be reported as a
    # single "voucher" whose text is the whole panel.
    found_cards, matched_selector = _first_cards(page, SNAPP_VOUCHER_SELECTORS)
    if matched_selector:
        log.info("Web check: %d voucher card(s) matched via %r",
                 len(found_cards), matched_selector)

    if not found_cards:
        return True, [], "Logged in, but no voucher cards were recognized"

    for card in found_cards:
        try:
            title = (card.inner_text(timeout=2000) or "").strip()
        except Exception:
            title = ""
        service = _snapp_service_name(card, title)
        details.append({
            "title": title or "(unlabeled)",
            "value": title or "",
            "expiry": "",
            "service_name": service,
        })
    # If nothing useful came out, degrade gracefully.
    if not any(d["title"] and d["title"] != "(unlabeled)" for d in details):
        return True, [], "Logged in, but could not parse vouchers (page structure unknown)"
    return True, details, f"{len(details)} voucher(s)"


def _snapp_service_name(card, fallback_text: str) -> str:
    """Best-effort attempt to read the service a Snapp voucher applies to.
    Looks for a short label near the card containing a known service keyword.
    Returns 'عمومی / نامشخص' when the site doesn't make it explicit (never guess)."""
    services = {
        "تاکسی": "اسنپ (تاکسی)",
        "فود": "اسنپ‌فود",
        "مارکت": "اسنپ‌مارکت",
        "باکس": "اسنپ‌باکس",
        "دیجی": "دیجی‌کالا",
        "پیک": "اسنپ‌پیک",
    }
    try:
        text = card.inner_text(timeout=1500) or ""
    except Exception:
        text = fallback_text
    low = text
    for key, name in services.items():
        if key in low:
            return name
    return "عمومی / نامشخص"


def _run_okala_discounts(page) -> tuple[bool, list[dict], str]:
    """Detect Okala login (via a Persian mobile number pattern) and scrape the
    user's discount codes by clicking the "کدهای تخفیف من" row on /profile.

    Heuristics only. If the selector for the menu row or the discount list
    changes, this needs re-tuning.
    """
    import re
    log = logging.getLogger(__name__)
    # Client-rendered app: read it only once it has actually painted.
    body = _wait_for_page_ready(page, log)

    # The "My discount codes" menu row is part of the signed-in profile area.
    row = None
    for t in OKALA_ACCOUNT_TEXTS[:3]:
        try:
            loc = page.get_by_text(t, exact=False)
            if loc.count() > 0:
                row = loc.first
                break
        except Exception:
            continue

    # Login detection: only account-area UI proves a live session. Everything
    # else -- including a mobile number, which the shop also prints in its own
    # footer -- is treated as signed out. Under-claiming sends the user to log
    # in again; over-claiming hides a dead session behind a green tick.
    account_text = _okala_account_text(body)
    if row is None and not account_text:
        log.info("Web check: no Okala account UI on the page; treating as signed out")
        return False, [], "Session inactive (no account menu found; login required)"
    if row is None:
        log.info("Web check: Okala account UI %r found, but no discount row", account_text)
        return True, [], "Logged in, but 'کدهای تخفیف من' row was not found"

    try:
        row.click(timeout=5000)
    except Exception as exc:
        log.warning("Web check: could not open the discount-code row: %s", exc)

    # The codes are fetched after the click, so wait for one to appear instead
    # of reading the page immediately.
    try:
        page.wait_for_selector(
            ", ".join(OKALA_COUPON_SELECTORS), state="attached", timeout=10000
        )
    except Exception:
        pass                      # reported explicitly below

    details: list[dict] = []
    candidates, matched_selector = _first_cards(page, OKALA_COUPON_SELECTORS)
    # Only elements that contain something that looks like a code.
    cards = [c for c in candidates if _looks_like_coupon(c)]
    if matched_selector:
        log.info("Web check: %d/%d coupon candidate(s) looked like codes via %r",
                 len(cards), len(candidates), matched_selector)

    for card in cards:
        try:
            text = (card.inner_text(timeout=2000) or "").strip()
        except Exception:
            text = ""
        details.append({
            "title": text or "(unlabeled)",
            "value": text or "",
            "expiry": "",
            "status": "",
        })
    if not details:
        return True, [], "Logged in, but could not parse discount codes (page structure unknown)"
    return True, details, f"{len(details)} discount code(s)"


def _looks_like_coupon(card) -> bool:
    """True if the element's text contains something that resembles a code
    (a short alphanumeric token), to avoid grabbing every stray div."""
    import re
    try:
        text = (card.inner_text(timeout=1500) or "").strip()
    except Exception:
        return False
    if not text:
        return False
    # Okala codes are typically short uppercase tokens or numeric sequences.
    if re.search(r"\b[A-Z0-9]{4,}\b", text):
        return True
    if re.search(r"\b\d{3,}\b", text):
        return True
    return False


class _WebCheckTask(QRunnable):
    """Runs one Web Check (for one task × one profile) on a worker thread.

    The profile is launched with --remote-debugging-port against a DEDICATED
    persistent automation data folder (NOT the browser's default one). Modern
    Chromium (136+) ignores --remote-debugging-port on the default user-data
    dir for security reasons, which is why attaching to the user's normal
    profile silently fails. With our own folder the flag always works, the
    browser does NOT need to be closed first, and cookies persist across runs
    (the user logs in once per site and future runs reuse the session).

    The browser window is left OPEN on purpose — we only detach Playwright.
    """

    def __init__(self, config_store: ConfigStore, task: WebCheckTask,
                 browser: BrowserDefinition, profile: BrowserProfile,
                 signals: _WebCheckSignals) -> None:
        super().__init__()
        self._config_store = config_store
        self._task = task
        self._browser = browser
        self._profile = profile
        self._signals = signals
        self.setAutoDelete(True)

    def run(self) -> None:
        """Run the check and ALWAYS report back.

        If _execute() escaped with an exception no result would be emitted, and
        MainWindow would keep `_web_check_running = True` forever: every later
        Run would only be queued, never executed, and Stop cannot clear the flag
        either -- i.e. the whole Web Checks feature wedged until a restart.
        """
        log = logging.getLogger(__name__)
        try:
            result = self._execute()
        except BaseException as exc:  # noqa: BLE001 - a worker must report back
            log.error("Web check crashed (task=%s, profile=%s): %s",
                      self._task.id, self._profile.id, exc, exc_info=True)
            result = WebCheckResult(
                task_id=self._task.id, kind=self._task.kind,
                profile_id=self._profile.id,
                display_name=self._profile.display_name,
                success=False, error_message=f"Check crashed: {exc}",
            )
        try:
            self._signals.result_ready.emit(result)
        except RuntimeError:  # pragma: no cover - app quitting mid-check
            # The window (and with it this signal object) is gone; there is
            # nobody left to hand the result to.
            log.debug("Web check finished during shutdown; result dropped.")

    def _execute(self) -> WebCheckResult:
        log = logging.getLogger(__name__)
        base = WebCheckResult(
            task_id=self._task.id, kind=self._task.kind,
            profile_id=self._profile.id, display_name=self._profile.display_name,
            success=False, error_message="",
        )
        log.info("Web check start: task=%s profile=%s browser=%s",
                 self._task.display_name, self._profile.display_name,
                 self._browser.display_name)
        if not self._browser.is_valid():
            log.error("Web check: browser exe missing: %s", self._browser.executable_path)
            return self._with_error(
                base, f"Browser executable not found: {self._browser.executable_path}"
            )
        port = find_free_port()
        data_dir = _automation_user_data_dir(self._browser.key, self._profile.id)
        args = [
            str(self._browser.executable_path),
            f"--user-data-dir={data_dir}",
            f"--remote-debugging-port={port}",
            "--no-first-run", "--no-default-browser-check",
            "--window-size=1100,800",
        ]
        log.info("Web check: launching %s (data dir=%s, port=%d)",
                 self._browser.display_name, data_dir, port)
        try:
            subprocess.Popen(args, **_popen_kwargs())
        except FileNotFoundError:
            return self._with_error(base, "Executable not found; is the browser installed?")
        except OSError as exc:
            return self._with_error(base, f"Could not start browser: {exc}")

        if not cd_endpoint_is_ready(port):
            log.error("Web check: CDP endpoint never came up on port %d", port)
            return self._with_error(
                base,
                "Could not reach the browser's debugging endpoint (timeout). "
                "See app.log for details.",
            )
        log.info("Web check: CDP endpoint up on port %d", port)

        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
                contexts = browser.contexts
                if not contexts:
                    return self._with_error(base, "No browser context found via CDP")
                context = contexts[0]
                try:
                    log.info("Web check: navigating to %s", self._task.url)
                    page = context.new_page()
                    page.set_default_timeout(15000)
                    page.goto(self._task.url, wait_until="domcontentloaded", timeout=25000)
                    log.info("Web check: page loaded, running extractor")
                    if self._task.kind == WebCheckKind.SNAPP_VOUCHERS:
                        session, details, summary = _run_snapp_vouchers(page)
                    else:
                        session, details, summary = _run_okala_discounts(page)
                    log.info("Web check: done — session=%s, %s", session, summary)
                    return WebCheckResult(
                        task_id=self._task.id, kind=self._task.kind,
                        profile_id=self._profile.id, display_name=self._profile.display_name,
                        success=True, session_active=session, summary=summary, details=details,
                    )
                finally:
                    # Detach from Playwright WITHOUT closing the browser window
                    # (the user may want to keep browsing).
                    try:
                        browser.close()
                    except Exception:
                        pass
        except WebCheckPageNotReady as exc:
            log.error("Web check %s: page not ready: %s", self._task.display_name, exc)
            return self._with_error(base, str(exc))
        except Exception as exc:
            log.error("Web Check %s failed: %s", self._task.display_name, exc)
            return self._with_error(base, f"Automation failed: {exc}")

    def _with_error(self, base: WebCheckResult, message: str) -> WebCheckResult:
        base.error_message = message
        base.summary = message
        return base


# ══════════════════════════════════════════════════════════════════════
#  SECTION 6 — WORKSPACE RUNNER
# ══════════════════════════════════════════════════════════════════════

class WorkspaceRunner(QObject):
    profile_launched = Signal(object)
    finished = Signal(str)
    progress = Signal(int, int)   # (launched_so_far, total)

    def __init__(self, config_store: ConfigStore, launch_engine: LaunchEngine,
                 parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self._launch_engine = launch_engine
        self._pending_profile_ids: list[str] = []
        self._workspace_id: str = ""
        self._total: int = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._launch_next)

    def run(self, workspace: Workspace) -> None:
        self._timer.stop()
        self._workspace_id = workspace.id
        self._pending_profile_ids = list(workspace.profile_ids)
        self._total = len(self._pending_profile_ids)
        # Never a 0 ms interval: Qt would fire the timer on every event-loop
        # iteration and busy-loop the UI thread.
        interval_ms = max(30, int(workspace.launch_stagger_seconds * 1000))
        self._timer.setInterval(interval_ms)
        self._timer.start()
        self._launch_next()

    def _launch_next(self) -> None:
        while self._pending_profile_ids:
            profile_id = self._pending_profile_ids.pop(0)
            profile = self._config_store.profiles.get(profile_id)
            if profile is None:
                self.profile_launched.emit(LaunchResult(
                    item_id=profile_id, display_name="(deleted profile)",
                    success=False, error_message="This profile no longer exists in the config.",
                ))
                continue
            browser = self._config_store.browsers.get(profile.browser_key)
            if browser is None:
                self.profile_launched.emit(LaunchResult(
                    item_id=profile.id, display_name=profile.display_name,
                    success=False, error_message=f"Browser '{profile.browser_key}' is not configured.",
                ))
                continue
            self.progress.emit(self._total - len(self._pending_profile_ids), self._total)
            self._launch_engine.launch_profile(browser, profile)
            return
        self._timer.stop()
        self.finished.emit(self._workspace_id)


class StaggeredLauncher(QObject):
    """Launches a list of profiles one at a time on a timer so a large bulk
    launch doesn't flood the thread pool and thrash disk I/O.

    The caller passes a list of zero-arg callables (each of which starts one
    launch via LaunchEngine) and a stagger interval in seconds."""

    finished = Signal()
    progress = Signal(int, int)   # (launched_so_far, total)

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._pending: list = []
        self._total: int = 0
        self._timer = QTimer(self)
        self._timer.setSingleShot(False)
        self._timer.timeout.connect(self._launch_next)

    def run(self, calls: list, stagger_seconds: float) -> None:
        self._timer.stop()
        self._pending = list(calls)
        self._total = len(self._pending)
        interval_ms = max(30, int(stagger_seconds * 1000))
        self._timer.setInterval(interval_ms)
        self._timer.start()
        self._launch_next()

    def _launch_next(self) -> None:
        if not self._pending:
            self._timer.stop()
            self.finished.emit()
            return
        call = self._pending.pop(0)
        self.progress.emit(self._total - len(self._pending), self._total)
        call()


# ══════════════════════════════════════════════════════════════════════
#  SECTION 6.5 — SNAPP, DIGIKALA & UNIVERSAL BACKGROUND SESSION ENGINE (v2.0)
# ══════════════════════════════════════════════════════════════════════

def _snappy_decompress(data: bytes) -> bytes:
    pos = 0
    uncompressed_len = 0
    shift = 0
    while pos < len(data):
        b = data[pos]
        pos += 1
        uncompressed_len |= (b & 0x7f) << shift
        if not (b & 0x80):
            break
        shift += 7

    out = bytearray()
    while pos < len(data):
        tag = data[pos]
        pos += 1
        element_type = tag & 0x03
        if element_type == 0:
            length = (tag >> 2) + 1
            if length > 60:
                extra_bytes = length - 60
                length = 1 + int.from_bytes(data[pos:pos+extra_bytes], 'little')
                pos += extra_bytes
            out.extend(data[pos:pos+length])
            pos += length
        elif element_type == 1:
            length = ((tag >> 2) & 0x07) + 4
            offset = ((tag >> 5) << 8) | data[pos]
            pos += 1
            for _ in range(length):
                out.append(out[-offset])
        elif element_type == 2:
            length = (tag >> 2) + 1
            offset = int.from_bytes(data[pos:pos+2], 'little')
            pos += 2
            for _ in range(length):
                out.append(out[-offset])
        elif element_type == 3:
            length = (tag >> 2) + 1
            offset = int.from_bytes(data[pos:pos+4], 'little')
            pos += 4
            for _ in range(length):
                out.append(out[-offset])
    return bytes(out)


def _decode_jwt_payload(jwt_str: str) -> dict | None:
    try:
        parts = jwt_str.split('.')
        if len(parts) >= 2:
            p = parts[1]
            p += '=' * ((4 - len(p) % 4) % 4)
            return json.loads(base64.urlsafe_b64decode(p).decode('utf-8', errors='ignore'))
    except Exception:
        pass
    return None


_master_keys_cache: dict[str, bytes | None] = {}

def _get_chromium_master_key(user_data_root: Path) -> bytes | None:
    ud_str = str(user_data_root)
    if ud_str in _master_keys_cache:
        return _master_keys_cache[ud_str]
    local_state_path = user_data_root / 'Local State'
    if not local_state_path.exists():
        _master_keys_cache[ud_str] = None
        return None
    try:
        with open(local_state_path, 'r', encoding='utf-8') as f:
            local_state = json.load(f)
        enc_key_b64 = local_state.get('os_crypt', {}).get('encrypted_key')
        if not enc_key_b64:
            _master_keys_cache[ud_str] = None
            return None
        encrypted_key = base64.b64decode(enc_key_b64)[5:]
        from ctypes import windll, byref, c_char, Structure, POINTER
        import ctypes
        class DATA_BLOB(Structure):
            _fields_ = [('cbData', ctypes.c_ulong), ('pbData', POINTER(c_char))]
        blob_in = DATA_BLOB(len(encrypted_key), ctypes.cast(ctypes.create_string_buffer(encrypted_key), POINTER(c_char)))
        blob_out = DATA_BLOB()
        if windll.crypt32.CryptUnprotectData(byref(blob_in), None, None, None, None, 0, byref(blob_out)):
            key = ctypes.string_at(blob_out.pbData, blob_out.cbData)
            windll.kernel32.LocalFree(blob_out.pbData)
            _master_keys_cache[ud_str] = key
            return key
    except Exception:
        pass
    _master_keys_cache[ud_str] = None
    return None


def _decrypt_chromium_cookie_value(enc_val: bytes, plain_val: str, master_key: bytes | None, db_version: int = 24) -> str:
    """Decrypts Chromium v10/v20 AES-GCM cookie value and strips the 32-byte
    SHA-256 domain-binding prefix introduced in Chromium 127+ (db_version >= 24)."""
    if plain_val:
        return plain_val
    if not enc_val:
        return ""
    if not (enc_val.startswith(b'v10') or enc_val.startswith(b'v11') or enc_val.startswith(b'v20')):
        try:
            from ctypes import windll, byref, c_char, Structure, POINTER
            import ctypes
            class DATA_BLOB(Structure):
                _fields_ = [('cbData', ctypes.c_ulong), ('pbData', POINTER(c_char))]
            blob_in = DATA_BLOB(len(enc_val), ctypes.cast(ctypes.create_string_buffer(enc_val), POINTER(c_char)))
            blob_out = DATA_BLOB()
            if windll.crypt32.CryptUnprotectData(byref(blob_in), None, None, None, None, 0, byref(blob_out)):
                raw = ctypes.string_at(blob_out.pbData, blob_out.cbData)
                windll.kernel32.LocalFree(blob_out.pbData)
                return raw.decode('utf-8', errors='ignore')
        except Exception:
            return ""
    if not master_key:
        return ""
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        nonce = enc_val[3:15]
        ciphertext = enc_val[15:]
        dec = AESGCM(master_key).decrypt(nonce, ciphertext, None)
        # In Chromium 127+ (cookie DB version >= 24), the first 32 bytes are the SHA256(host_key) digest!
        if db_version >= 24 and len(dec) > 32:
            # Check if stripping 32 bytes yields clean printable ASCII/UTF-8
            candidate = dec[32:]
            try:
                return candidate.decode('utf-8')
            except UnicodeDecodeError:
                pass
        return dec.decode('utf-8', errors='ignore')
    except Exception:
        return ""


def _extract_chromium_domain_cookies(user_data: Path, prof_dir: Path, domain_keywords: tuple[str, ...]) -> list[dict]:
    """Extracts and decrypts cookies matching any of `domain_keywords` from a Chromium profile.
    Handles locked SQLite files (when browser is running) and WAL journals."""
    cookie_db = prof_dir / 'Network' / 'Cookies'
    if not cookie_db.exists():
        cookie_db = prof_dir / 'Cookies'
    if not cookie_db.exists():
        return []

    master_key = _get_chromium_master_key(user_data)
    cookies_out: list[dict] = []
    tmp_dir = tempfile.mkdtemp(prefix="mpl_ck_")
    try:
        tmp_db = Path(tmp_dir) / "Cookies"
        try:
            shutil.copyfile(cookie_db, tmp_db)
            for ext in ("-wal", "-shm", "-journal"):
                wal_src = Path(str(cookie_db) + ext)
                if wal_src.exists():
                    shutil.copyfile(wal_src, Path(str(tmp_db) + ext))
            conn = sqlite3.connect(str(tmp_db), timeout=1.5)
        except Exception:
            # Fallback to immutable read-only URI if file is locked by an open browser
            uri = f"file:{cookie_db.as_posix()}?mode=ro&immutable=1"
            conn = sqlite3.connect(uri, uri=True, timeout=1.5)

        try:
            c = conn.cursor()
            db_version = 24
            try:
                c.execute("SELECT value FROM meta WHERE key='version'")
                row_v = c.fetchone()
                if row_v and str(row_v[0]).isdigit():
                    db_version = int(row_v[0])
            except Exception:
                pass

            where_clauses = " OR ".join(["host_key LIKE ?" for _ in domain_keywords])
            params = tuple(f"%{kw}%" for kw in domain_keywords)
            c.execute(
                f"SELECT name, value, encrypted_value, host_key, path, expires_utc, is_secure, is_httponly "
                f"FROM cookies WHERE {where_clauses}",
                params
            )
            for name, val, enc_val, host, path, exp_utc, sec, httponly in c.fetchall():
                dec_val = _decrypt_chromium_cookie_value(enc_val, val, master_key, db_version)
                # Convert Chromium WebKit timestamp (microseconds since 1601-01-01) to Unix epoch seconds
                unix_exp = int((exp_utc / 1_000_000) - 11644473600) if exp_utc and exp_utc > 11644473600000000 else int(time.time() + 86400 * 30)
                cookies_out.append({
                    'name': name,
                    'value': dec_val,
                    'domain': host,
                    'path': path or '/',
                    'expirationDate': unix_exp,
                    'secure': bool(sec),
                    'httpOnly': bool(httponly),
                    '_encrypted_present': bool(enc_val),
                })
        finally:
            conn.close()
    except Exception:
        pass
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return cookies_out


def _read_chromium_leveldb_texts(leveldb_dir: Path, keyword_bytes: bytes) -> list[str]:
    """Reads both raw and UTF-16LE decoded text chunks from Chromium LevelDB .ldb/.log files."""
    texts: list[str] = []
    if not leveldb_dir.exists():
        return texts
    try:
        files = sorted(
            list(leveldb_dir.glob('*.ldb')) + list(leveldb_dir.glob('*.log')),
            key=lambda p: p.stat().st_mtime if p.exists() else 0,
            reverse=True
        )
        for f in files:
            try:
                b = f.read_bytes()
                low = b.lower()
                # Check both ASCII and UTF-16LE representations of keyword_bytes
                kw_utf16 = keyword_bytes.decode('ascii', errors='ignore').encode('utf-16le')
                if keyword_bytes not in low and kw_utf16 not in low:
                    continue
                texts.append(b.decode('utf-8', errors='ignore'))
                texts.append(b.decode('utf-16le', errors='ignore'))
            except Exception:
                pass
    except Exception:
        pass
    return texts


def check_profile_snapp_session(profile: BrowserProfile, browser: BrowserDefinition | None) -> dict:
    """Inspect profile disk storage (LevelDB / SQLite / Cookies) in background without launching browser.
    Fixed v2.0:
    - Reads both LocalStorage AND Cookies on Chromium and Gecko
    - Decodes UTF-16LE & Snappy blocks
    - Never falsely marks a logged-in Snapp session as 'expired' just because initial JWT exp passed
      (Snapp's PWA keeps the session alive via refreshToken / cookies / persist:persisted)."""
    now = int(time.time())
    res = {
        'profile_id': profile.id,
        'status': 'none',  # 'active', 'expired', 'none'
        'phone': None,
        'exp': None,
        'token': None,
        'has_snapp': False,
        'has_refresh_or_cookie': False,
        'error': None
    }
    bk = (profile.browser_key or '').lower()
    engine = browser.engine if browser else BrowserEngine.CHROMIUM
    is_gecko = (engine == BrowserEngine.GECKO) or bk in ('waterfox', 'floorp', 'firefox', 'librewolf')

    if is_gecko:
        p_path = profile.profile_path
        if not p_path or not os.path.exists(p_path):
            return res
        storage_dir = Path(p_path) / 'storage' / 'default'
        if storage_dir.exists():
            try:
                snapp_dirs = [d for d in storage_dir.iterdir() if 'snapp' in d.name.lower()]
                for sd in snapp_dirs:
                    ls_db = sd / 'ls' / 'data.sqlite'
                    if ls_db.exists():
                        res['has_snapp'] = True
                        try:
                            conn = sqlite3.connect(f'file:{ls_db}?mode=ro', uri=True, timeout=1.0)
                            c = conn.cursor()
                            c.execute("SELECT key, compression_type, value FROM data")
                            for k, comp, val in c.fetchall():
                                raw_bytes = None
                                if comp == 1:
                                    try:
                                        raw_bytes = _snappy_decompress(val)
                                    except Exception:
                                        pass
                                else:
                                    raw_bytes = val
                                if not raw_bytes:
                                    continue
                                txt = raw_bytes.decode('utf-8', errors='ignore')
                                if k in ('refreshToken', 'persist:persisted', 'user', 'profile'):
                                    res['has_refresh_or_cookie'] = True
                                if not res['phone']:
                                    ph = re.findall(r'09\d{9}', txt)
                                    if ph:
                                        res['phone'] = ph[0]
                                jwts = re.findall(r'eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}', txt)
                                for j in jwts:
                                    p = _decode_jwt_payload(j)
                                    if p and ('passenger' in str(p.get('aud', '')) or 'snapp' in str(p).lower()):
                                        exp_val = p.get('exp', 0)
                                        if not res['token'] or exp_val >= (res['exp'] or 0):
                                            res['token'] = j
                                            res['exp'] = exp_val
                                        if not res['phone']:
                                            res['phone'] = p.get('phone') or p.get('cellphone')
                            conn.close()
                        except Exception as e:
                            res['error'] = str(e)
            except Exception:
                pass
        # Also check Gecko cookies.sqlite for snapp.taxi / snapp.ir
        cookies_db = Path(p_path) / 'cookies.sqlite'
        if cookies_db.exists():
            try:
                conn = sqlite3.connect(f'file:{cookies_db}?mode=ro', uri=True, timeout=1.0)
                c = conn.cursor()
                c.execute("SELECT name, value FROM moz_cookies WHERE host LIKE '%snapp.taxi%' OR host LIKE '%snapp.ir%'")
                rows = c.fetchall()
                if rows:
                    res['has_snapp'] = True
                    res['has_refresh_or_cookie'] = True
                conn.close()
            except Exception:
                pass
    else:
        # Chromium engine
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        rel = _CHROMIUM_USER_DATA_FOLDER.get(bk)
        if rel and local_appdata:
            user_data = Path(local_appdata).joinpath(*rel)
            arg = profile.profile_arg or 'Default'
            prof_dir = user_data / arg
            leveldb_dir = prof_dir / 'Local Storage' / 'leveldb'

            texts = _read_chromium_leveldb_texts(leveldb_dir, b'snapp')
            for txt in texts:
                res['has_snapp'] = True
                if 'persist:persisted' in txt or 'refreshToken' in txt or 'cellphone' in txt:
                    res['has_refresh_or_cookie'] = True
                if not res['phone']:
                    pm = re.findall(r'09[0-9]{9}', txt)
                    if pm:
                        res['phone'] = pm[0]
                jwts = re.findall(r'eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}', txt)
                for jstr in jwts:
                    p = _decode_jwt_payload(jstr)
                    if p and ('passenger' in str(p.get('aud', '')) or 'snapp' in str(p).lower()):
                        exp_val = p.get('exp', 0)
                        if not res['token'] or exp_val >= (res['exp'] or 0):
                            res['token'] = jstr
                            res['exp'] = exp_val
                        if not res['phone']:
                            res['phone'] = p.get('phone') or p.get('cellphone')

            # Also inspect Chromium Cookies for snapp.taxi / snapp.ir
            ck_list = _extract_chromium_domain_cookies(user_data, prof_dir, ('snapp.taxi', 'snapp.ir'))
            if ck_list:
                res['has_snapp'] = True
                res['has_refresh_or_cookie'] = True
                for ck in ck_list:
                    v = ck.get('value', '')
                    if v.startswith('eyJ') and not res['token']:
                        p = _decode_jwt_payload(v)
                        if p:
                            res['token'] = v
                            res['exp'] = p.get('exp', 0)

    # Status determination:
    # In Snapp PWA, as long as a valid token + phone/refresh/cookies exist in storage,
    # the session remains logged in even after the initial JWT 'exp' timestamp passes!
    if res['token']:
        if (res['exp'] and res['exp'] > now) or res['phone'] or res['has_refresh_or_cookie']:
            res['status'] = 'active'
        else:
            res['status'] = 'expired'
    elif res['has_snapp'] and res['phone'] and res['has_refresh_or_cookie']:
        res['status'] = 'active'
    elif res['has_snapp']:
        res['status'] = 'expired'

    return res


class _SnappCheckSignals(QObject):
    profile_checked = Signal(str, dict)  # (profile_id, result_dict)
    progress = Signal(int, int)          # (done, total)
    finished = Signal(dict)              # summary stats


class _SnappCheckTask(QRunnable):
    def __init__(self, profiles: list[BrowserProfile], browsers: dict[str, BrowserDefinition],
                 signals: _SnappCheckSignals) -> None:
        super().__init__()
        self._profiles = list(profiles)
        self._browsers = dict(browsers)
        self._signals = signals
        self.setAutoDelete(True)

    def run(self) -> None:
        counts = {'active': 0, 'expired': 0, 'none': 0, 'total': len(self._profiles)}
        for idx, profile in enumerate(self._profiles):
            browser = self._browsers.get(profile.browser_key)
            try:
                res = check_profile_snapp_session(profile, browser)
            except Exception as e:
                res = {'profile_id': profile.id, 'status': 'none', 'error': str(e)}

            st = res.get('status', 'none')
            counts[st] = counts.get(st, 0) + 1
            try:
                self._signals.profile_checked.emit(profile.id, res)
                self._signals.progress.emit(idx + 1, len(self._profiles))
            except RuntimeError:
                return
            time.sleep(0.003)

        try:
            self._signals.finished.emit(counts)
        except RuntimeError:
            pass


def _parse_digikala_text(text: str) -> tuple[str | None, str | None]:
    phone = None
    name = None
    if not text:
        return phone, name
    m_user = re.search(r'addressUserMobile[\\"\s:]+(09\d{9})', text)
    if m_user:
        phone = m_user.group(1)
    if not phone:
        m_phone = re.search(r'(?:cellphone|phone|mobile|userPhone)[\\"\s:]+(09\d{9})', text)
        if m_phone:
            phone = m_phone.group(1)

    m_name = re.search(r'addressUserFullName[\\"\s:]+([^\\",}{]+)', text)
    if m_name:
        candidate = m_name.group(1).strip()
        if candidate and len(candidate) > 2 and not candidate.isdigit():
            name = candidate
    if not name:
        m_fn = re.search(r'(?:first_name|firstName)[\\"\s:]+([^\\",}{]+)', text)
        m_ln = re.search(r'(?:last_name|lastName)[\\"\s:]+([^\\",}{]+)', text)
        if m_fn and m_ln:
            fn = m_fn.group(1).strip()
            ln = m_ln.group(1).strip()
            if fn and ln and not fn.isdigit():
                name = f"{fn} {ln}"
        elif m_fn:
            fn = m_fn.group(1).strip()
            if fn and not fn.isdigit():
                name = fn

    if not phone:
        all_phones = re.findall(r'09\d{9}', text)
        if all_phones:
            phone = all_phones[0]
    return phone, name


def check_profile_digikala_session(profile: BrowserProfile, browser: BrowserDefinition | None) -> dict:
    """Inspect profile disk storage (LevelDB / SQLite / Cookies) for Digikala session in background."""
    now = int(time.time())
    res = {
        'profile_id': profile.id,
        'status': 'none',  # 'active', 'expired', 'none'
        'phone': None,
        'name': None,
        'exp': None,
        'token': None,
        'has_digikala': False,
        'has_auth_cookie': False,
        'error': None
    }
    bk = (profile.browser_key or '').lower()
    engine = browser.engine if browser else BrowserEngine.CHROMIUM
    is_gecko = (engine == BrowserEngine.GECKO) or bk in ('waterfox', 'floorp', 'firefox', 'librewolf')

    if is_gecko:
        p_path = profile.profile_path
        if not p_path or not os.path.exists(p_path):
            return res
        storage_dir = Path(p_path) / 'storage' / 'default'
        if storage_dir.exists():
            try:
                dk_dirs = [d for d in storage_dir.iterdir() if 'digikala.com' in d.name.lower()]
                for dd in dk_dirs:
                    ls_db = dd / 'ls' / 'data.sqlite'
                    if ls_db.exists():
                        res['has_digikala'] = True
                        try:
                            conn = sqlite3.connect(f'file:{ls_db}?mode=ro', uri=True, timeout=1.0)
                            c = conn.cursor()
                            c.execute("SELECT key, compression_type, value FROM data")
                            for k, comp, val in c.fetchall():
                                raw_bytes = _snappy_decompress(val) if comp == 1 else val
                                if not raw_bytes:
                                    continue
                                for enc in ('utf-8', 'utf-16le'):
                                    txt = raw_bytes.decode(enc, errors='ignore')
                                    ph, nm = _parse_digikala_text(txt)
                                    if ph and not res['phone']:
                                        res['phone'] = ph
                                    if nm and not res['name']:
                                        res['name'] = nm
                                    jwts = re.findall(r'eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}', txt)
                                    for j in jwts:
                                        p = _decode_jwt_payload(j)
                                        if p:
                                            iss = str(p.get('iss', '')).lower()
                                            azp = str(p.get('azp', '')).lower()
                                            if 'snapp' in iss or 'passenger' in str(p.get('aud', '')):
                                                continue
                                            if 'digikala' in iss or 'dk-group' in iss or 'digikala' in azp:
                                                exp = p.get('exp', 0)
                                                if exp >= (res['exp'] or 0):
                                                    res['token'] = j
                                                    res['exp'] = exp
                            conn.close()
                        except Exception as e:
                            res['error'] = str(e)
            except Exception:
                pass

        # Also check Gecko cookies.sqlite for Digikala auth cookies!
        cookies_db = Path(p_path) / 'cookies.sqlite'
        if cookies_db.exists():
            try:
                conn = sqlite3.connect(f'file:{cookies_db}?mode=ro', uri=True, timeout=1.0)
                c = conn.cursor()
                c.execute("SELECT name, value FROM moz_cookies WHERE host LIKE '%digikala.com%'")
                for cname, cval in c.fetchall():
                    res['has_digikala'] = True
                    cn_low = str(cname or '').lower()
                    if 'token' in cn_low or 'session' in cn_low or 'auth' in cn_low or 'user' in cn_low:
                        res['has_auth_cookie'] = True
                    m_ph = re.search(r'09\d{9}', str(cval or ''))
                    if m_ph and not res['phone']:
                        res['phone'] = m_ph.group(0)
                conn.close()
            except Exception:
                pass
    else:
        # Chromium engine
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        rel = _CHROMIUM_USER_DATA_FOLDER.get(bk)
        if rel and local_appdata:
            user_data = Path(local_appdata).joinpath(*rel)
            arg = profile.profile_arg or 'Default'
            prof_dir = user_data / arg

            # 1. Read Chromium Cookies using Chrome 130+ aware decryptor
            dk_cookies = _extract_chromium_domain_cookies(user_data, prof_dir, ('digikala.com',))
            for ck in dk_cookies:
                res['has_digikala'] = True
                cname = str(ck.get('name') or '').lower()
                cval = str(ck.get('value') or '')
                if 'token' in cname or 'auth' in cname or 'user' in cname or ck.get('_encrypted_present'):
                    res['has_auth_cookie'] = True
                m_ph = re.search(r'09\d{9}', cval)
                if m_ph and not res['phone']:
                    res['phone'] = m_ph.group(0)
                jwts = re.findall(r'eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}', cval)
                for jstr in jwts:
                    p = _decode_jwt_payload(jstr)
                    if p:
                        iss = str(p.get('iss', '')).lower()
                        azp = str(p.get('azp', '')).lower()
                        if 'digikala' in iss or 'dk-group' in iss or 'digikala' in azp:
                            exp = p.get('exp', 0)
                            if exp >= (res['exp'] or 0):
                                res['token'] = jstr
                                res['exp'] = exp

            # 2. Read Local Storage LevelDB (both UTF-8 & UTF-16LE + Snappy blocks)
            leveldb_dir = prof_dir / 'Local Storage' / 'leveldb'
            if leveldb_dir.exists():
                try:
                    files = list(leveldb_dir.glob('*.ldb')) + list(leveldb_dir.glob('*.log'))
                    for f in files:
                        try:
                            content = f.read_bytes()
                            if b'digikala' not in content.lower():
                                continue
                            res['has_digikala'] = True

                            pos = 0
                            while True:
                                idx = content.find(b'persist:DIGIKALA', pos)
                                if idx == -1:
                                    break
                                raw_chunk = content[idx:idx+8000]
                                for off in range(12, min(40, len(raw_chunk))):
                                    try:
                                        d = _snappy_decompress(raw_chunk[off:])
                                        if d and len(d) > 60:
                                            for enc in ('utf-16le', 'utf-8'):
                                                t_dec = d[1:].decode(enc, errors='ignore') if d[0] == 0 else d.decode(enc, errors='ignore')
                                                ph, nm = _parse_digikala_text(t_dec)
                                                if ph and not res['phone']:
                                                    res['phone'] = ph
                                                if nm and not res['name']:
                                                    res['name'] = nm
                                            break
                                    except Exception:
                                        pass
                                pos = idx + 16

                            for enc in ('utf-8', 'utf-16le'):
                                t_all = content.decode(enc, errors='ignore')
                                ph, nm = _parse_digikala_text(t_all)
                                if ph and not res['phone']:
                                    res['phone'] = ph
                                if nm and not res['name']:
                                    res['name'] = nm

                            jwts = re.findall(rb'eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}', content)
                            for jb in jwts:
                                jstr = jb.decode('ascii', errors='ignore')
                                p = _decode_jwt_payload(jstr)
                                if p:
                                    iss = str(p.get('iss', '')).lower()
                                    azp = str(p.get('azp', '')).lower()
                                    if 'snapp' in iss or 'passenger' in str(p.get('aud', '')):
                                        continue
                                    if 'digikala' in iss or 'dk-group' in iss or 'digikala' in azp:
                                        exp = p.get('exp', 0)
                                        if exp >= (res['exp'] or 0):
                                            res['token'] = jstr
                                            res['exp'] = exp
                        except Exception:
                            pass
                except Exception:
                    pass

    if res['phone'] or res['token'] or res['has_auth_cookie']:
        res['status'] = 'active'
    elif res['has_digikala']:
        res['status'] = 'none'

    return res


class _DigikalaCheckSignals(QObject):
    profile_checked = Signal(str, dict)  # (profile_id, result_dict)
    progress = Signal(int, int)          # (done, total)
    finished = Signal(dict)              # summary stats


class _DigikalaCheckTask(QRunnable):
    def __init__(self, profiles: list[BrowserProfile], browsers: dict[str, BrowserDefinition],
                 signals: _DigikalaCheckSignals) -> None:
        super().__init__()
        self._profiles = list(profiles)
        self._browsers = dict(browsers)
        self._signals = signals
        self.setAutoDelete(True)

    def run(self) -> None:
        counts = {'active': 0, 'expired': 0, 'none': 0, 'total': len(self._profiles)}
        for idx, profile in enumerate(self._profiles):
            browser = self._browsers.get(profile.browser_key)
            try:
                res = check_profile_digikala_session(profile, browser)
            except Exception as e:
                res = {'profile_id': profile.id, 'status': 'none', 'error': str(e)}

            st = res.get('status', 'none')
            counts[st] = counts.get(st, 0) + 1
            try:
                self._signals.profile_checked.emit(profile.id, res)
                self._signals.progress.emit(idx + 1, len(self._profiles))
            except RuntimeError:
                return
            time.sleep(0.003)

        try:
            self._signals.finished.emit(counts)
        except RuntimeError:
            pass


def extract_full_digikala_session(profile: BrowserProfile, browser: BrowserDefinition | None) -> dict:
    """Extract complete Digikala session payload (storage, cookies, token, user info)."""
    info = check_profile_digikala_session(profile, browser)
    bk = (profile.browser_key or '').lower()
    browser_display = browser.display_name if browser else bk.title()
    phone_clean = info.get('phone') or '09000000000'

    payload = {
        "version": 1,
        "app": "digikala",
        "phone": phone_clean,
        "name": info.get('name'),
        "token": info.get('token'),
        "exp": info.get('exp'),
        "status": info.get('status', 'none'),
        "userName": f"{browser_display} - {profile.display_name}",
        "browser": bk,
        "profileName": profile.display_name,
        "exportedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "expiresAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(info['exp'])) if info.get('exp') else None
    }
    b64 = base64.b64encode(json.dumps(payload, ensure_ascii=False).encode('utf-8')).decode('ascii')
    return {
        'success': info.get('has_digikala', False) and (info.get('status') == 'active' or info.get('phone') is not None),
        'profile_id': profile.id,
        'display_name': profile.display_name,
        'phone': info.get('phone'),
        'name': info.get('name'),
        'token': info.get('token'),
        'exp': info.get('exp'),
        'is_active': info.get('status') == 'active',
        'payload': payload,
        'session_code': f"DK_{b64}",
        'error': info.get('error') or ("سشن دیجی‌کالا در این پروفایل یافت نشد." if not info.get('has_digikala') else None)
    }


def check_profile_universal_domain(profile: BrowserProfile, browser: BrowserDefinition | None, domain_or_url: str) -> dict:
    """Offline disk inspection checking BOTH Cookies and LocalStorage across Chromium & Gecko for ANY domain."""
    netloc = domain_or_url.strip().lower()
    if '://' in netloc:
        netloc = urllib.parse.urlparse(netloc).netloc
    netloc = netloc.split(':')[0]
    domain_core = netloc[4:] if netloc.startswith('www.') else netloc

    res = {
        'profile_id': profile.id,
        'domain': netloc,
        'status': 'none',
        'details': '',
        'phone': None,
        'has_data': False
    }

    if 'digikala' in domain_core:
        dk_res = check_profile_digikala_session(profile, browser)
        if dk_res.get('status') == 'active' or dk_res.get('phone'):
            res['has_data'] = True
            res['status'] = 'active'
            res['phone'] = dk_res.get('phone')
            ph = dk_res.get('phone', '')
            nm = dk_res.get('name', '')
            details = "سشن فعال دیجی‌کالا"
            if ph:
                details += f" (شماره: {ph})"
            if nm:
                details += f" [نام: {nm}]"
            res['details'] = details
            return res

    if 'snapp.taxi' in domain_core or domain_core == 'snapp.ir':
        sn_res = check_profile_snapp_session(profile, browser)
        if sn_res.get('status') == 'active':
            res['has_data'] = True
            res['status'] = 'active'
            res['phone'] = sn_res.get('phone')
            ph = sn_res.get('phone') or 'دارای توکن'
            res['details'] = f"سشن فعال اسنپ (شماره: {ph})"
            return res

    bk = (profile.browser_key or '').lower()
    engine = browser.engine if browser else BrowserEngine.CHROMIUM
    is_gecko = (engine == BrowserEngine.GECKO) or bk in ('waterfox', 'floorp', 'firefox', 'librewolf')

    if is_gecko:
        p_path = profile.profile_path
        if p_path and os.path.exists(p_path):
            storage_dir = Path(p_path) / 'storage' / 'default'
            if storage_dir.exists():
                matched = [d for d in storage_dir.iterdir() if domain_core in d.name.lower()]
                if matched:
                    res['has_data'] = True
                    res['status'] = 'active'
                    res['details'] = f"فضای LocalStorage ({len(matched)} بخش)"
                    for md in matched:
                        ls_db = md / 'ls' / 'data.sqlite'
                        if ls_db.exists():
                            try:
                                conn = sqlite3.connect(f'file:{ls_db}?mode=ro', uri=True, timeout=1.0)
                                c = conn.cursor()
                                c.execute("SELECT compression_type, value FROM data")
                                for comp, val in c.fetchall():
                                    rb = _snappy_decompress(val) if comp == 1 else val
                                    if rb:
                                        m_ph = re.search(r'09\d{9}', rb.decode('utf-8', errors='ignore'))
                                        if m_ph:
                                            res['phone'] = m_ph.group(0)
                                            res['details'] += f" | شماره: {res['phone']}"
                                            break
                                conn.close()
                            except Exception:
                                pass
            cookies_db = Path(p_path) / 'cookies.sqlite'
            if cookies_db.exists():
                try:
                    conn = sqlite3.connect(f'file:{cookies_db}?mode=ro', uri=True, timeout=1.0)
                    c = conn.cursor()
                    c.execute("SELECT count(*) FROM moz_cookies WHERE host LIKE ?", (f"%{domain_core}%",))
                    cnt = c.fetchone()[0]
                    if cnt > 0:
                        res['has_data'] = True
                        res['status'] = 'active'
                        res['details'] = (res['details'] + f" + {cnt} کوکی" if res['details'] else f"{cnt} کوکی فعال")
                    conn.close()
                except Exception:
                    pass
    else:
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        rel = _CHROMIUM_USER_DATA_FOLDER.get(bk)
        if rel and local_appdata:
            user_data = Path(local_appdata).joinpath(*rel)
            arg = profile.profile_arg or 'Default'
            prof_dir = user_data / arg

            # 1. Check Chromium Cookies DB (Crucial for Divar, Basalam, SnappFood, SnappExpress, etc.)
            ck_matches = _extract_chromium_domain_cookies(user_data, prof_dir, (domain_core,))
            if ck_matches:
                res['has_data'] = True
                res['status'] = 'active'
                res['details'] = f"{len(ck_matches)} کوکی فعال"
                for ck in ck_matches:
                    m_ph = re.search(r'09\d{9}', str(ck.get('value') or ''))
                    if m_ph and not res['phone']:
                        res['phone'] = m_ph.group(0)
                        res['details'] += f" (شماره: {res['phone']})"
                        break

            # 2. Check Chromium LocalStorage LevelDB
            leveldb_dir = prof_dir / 'Local Storage' / 'leveldb'
            texts = _read_chromium_leveldb_texts(leveldb_dir, domain_core.encode('ascii', errors='ignore'))
            if texts:
                res['has_data'] = True
                res['status'] = 'active'
                res['details'] = (res['details'] + " + LocalStorage") if res['details'] else "داده‌های LocalStorage فعال"
                if not res['phone']:
                    for t in texts:
                        m_ph = re.search(r'09\d{9}', t)
                        if m_ph:
                            res['phone'] = m_ph.group(0)
                            res['details'] += f" (شماره: {res['phone']})"
                            break
    return res


class UniversalSessionCheckDialog(QDialog):
    """Dialog allowing users to enter ANY URL to inspect sessions across profiles and launch active ones."""
    def __init__(self, profiles: list[BrowserProfile], browsers: dict[str, BrowserDefinition],
                 launch_callback=None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("🌐 رادار جامع بررسی سشن سایت‌ها (اسنپ، فود، اکسپرس، دیجی‌کالا، دیوار و ...)")
        self.resize(860, 560)
        self._profiles = profiles
        self._browsers = browsers
        self._launch_callback = launch_callback
        self._active_profiles: list[BrowserProfile] = []

        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(18, 18, 18, 18)

        lbl_info = QLabel(
            "<b>آدرس سایت یا سرویس مورد نظر را انتخاب یا وارد کنید:</b><br>"
            "<span style='color: #9CA3AF; font-size: 11px;'>"
            "برنامه بدون باز کردن مرورگر، دیتابیس کوکی‌های رمزگشایی‌شده (AES-GCM) و LocalStorage تمام پروفایل‌ها را اسکن می‌کند."
            "</span>"
        )
        lay.addWidget(lbl_info)

        url_box = QHBoxLayout()
        url_box.setSpacing(8)
        self._url_input = QLineEdit("https://www.digikala.com/profile/")
        self._url_input.setPlaceholderText("https://... (مثلاً: https://snappfood.ir یا https://divar.ir/my-divar)")
        self._url_input.setStyleSheet("padding: 7px 10px; font-size: 12px; border-radius: 6px;")
        url_box.addWidget(self._url_input, stretch=1)

        self._btn_run = QPushButton("🔍 شروع اسکن دیسک")
        self._btn_run.setStyleSheet("background-color: #2563EB; color: white; font-weight: bold; padding: 7px 18px; border-radius: 6px;")
        self._btn_run.clicked.connect(self._run_check)
        url_box.addWidget(self._btn_run)
        lay.addLayout(url_box)

        presets_lay = QHBoxLayout()
        presets_lay.setSpacing(6)
        presets_lay.addWidget(QLabel("سرویس‌های آماده:"))
        presets = [
            ("🚗 اسنپ", "https://app.snapp.taxi/"),
            ("🍔 اسنپ‌فود", "https://snappfood.ir/"),
            ("🛒 اسنپ‌اکسپرس", "https://snapp.express/"),
            ("🛍️ دیجی‌کالا", "https://www.digikala.com/profile/"),
            ("📱 دیوار", "https://divar.ir/my-divar"),
            ("🧺 باسلام", "https://basalam.com/user"),
            ("🥦 اکالا", "https://www.okala.com/profile"),
            ("✈️ تلگرام وب", "https://web.telegram.org/"),
        ]
        for name, u in presets:
            btn_p = QPushButton(name)
            btn_p.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_p.setStyleSheet("background: #1F2937; color: #E5E7EB; border: 1px solid #374151; border-radius: 4px; padding: 3px 8px; font-size: 11px;")
            btn_p.clicked.connect(lambda _, url_val=u: (self._url_input.setText(url_val), self._run_check()))
            presets_lay.addWidget(btn_p)
        presets_lay.addStretch(1)
        lay.addLayout(presets_lay)

        self._table = QTableWidget(0, 5)
        self._table.setHorizontalHeaderLabels(["ردیف", "نام پروفایل", "مرورگر", "وضعیت سشن", "جزئیات کوکی و استوریج"])
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self._table.verticalHeader().setVisible(False)
        self._table.setAlternatingRowColors(True)
        lay.addWidget(self._table, stretch=1)

        bottom_box = QHBoxLayout()
        self._lbl_status = QLabel("آماده بررسی…")
        self._lbl_status.setStyleSheet("color: #9CA3AF; font-size: 11px;")
        bottom_box.addWidget(self._lbl_status, stretch=1)

        self._btn_launch_active = QPushButton("🚀 اجرای پروفایل‌های دارای سشن فعال (۰)")
        self._btn_launch_active.setEnabled(False)
        self._btn_launch_active.setStyleSheet("background-color: #059669; color: white; font-weight: bold; padding: 6px 14px; border-radius: 6px;")
        self._btn_launch_active.clicked.connect(self._launch_active_profiles)
        bottom_box.addWidget(self._btn_launch_active)

        btn_copy = QPushButton("📋 کپی گزارش")
        btn_copy.clicked.connect(self._copy_report)
        bottom_box.addWidget(btn_copy)

        btn_close = QPushButton("بستن")
        btn_close.clicked.connect(self.accept)
        bottom_box.addWidget(btn_close)
        lay.addLayout(bottom_box)

    def _run_check(self) -> None:
        raw_url = self._url_input.text().strip()
        if not raw_url:
            QMessageBox.warning(self, "خطا", "لطفاً ابتدا یک آدرس یا لینک وارد کنید.")
            return

        self._btn_run.setEnabled(False)
        self._lbl_status.setText("در حال بررسی پروفایل‌ها…")
        self._table.setRowCount(0)
        self._active_profiles.clear()

        active_count = 0
        none_count = 0

        for idx, prof in enumerate(self._profiles):
            browser = self._browsers.get(prof.browser_key)
            res = check_profile_universal_domain(prof, browser, raw_url)

            row = self._table.rowCount()
            self._table.insertRow(row)

            item_idx = QTableWidgetItem(str(idx + 1))
            item_idx.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(row, 0, item_idx)

            item_name = QTableWidgetItem(prof.display_name)
            self._table.setItem(row, 1, item_name)

            item_bk = QTableWidgetItem(prof.browser_key or "Unknown")
            item_bk.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(row, 2, item_bk)

            if res.get('status') == 'active':
                active_count += 1
                self._active_profiles.append(prof)
                item_st = QTableWidgetItem("🟢 سشن فعال")
                item_st.setForeground(QColor("#34D399"))
            else:
                none_count += 1
                item_st = QTableWidgetItem("⚪ بدون سشن")
                item_st.setForeground(QColor("#9CA3AF"))
            item_st.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(row, 3, item_st)

            item_det = QTableWidgetItem(res.get('details') or "داده‌ای در دیسک یافت نشد.")
            self._table.setItem(row, 4, item_det)

        self._btn_run.setEnabled(True)
        self._btn_launch_active.setEnabled(active_count > 0)
        self._btn_launch_active.setText(f"🚀 اجرای پروفایل‌های دارای سشن فعال ({active_count})")
        self._lbl_status.setText(f"پایان بررسی: {active_count} پروفایل دارای سشن فعال 🟢 | {none_count} بدون سشن ⚪")

    def _launch_active_profiles(self) -> None:
        if not self._active_profiles or not self._launch_callback:
            return
        self._launch_callback([p.id for p in self._active_profiles])
        self.accept()

    def _copy_report(self) -> None:
        lines = [f"گزارش بررسی سشن برای دامنه: {self._url_input.text()}"]
        lines.append("-" * 50)
        for r in range(self._table.rowCount()):
            pname = self._table.item(r, 1).text() if self._table.item(r, 1) else ""
            bname = self._table.item(r, 2).text() if self._table.item(r, 2) else ""
            st = self._table.item(r, 3).text() if self._table.item(r, 3) else ""
            det = self._table.item(r, 4).text() if self._table.item(r, 4) else ""
            lines.append(f"{r+1}. {pname} ({bname}) -> {st} | {det}")
        txt = "\n".join(lines)
        QApplication.clipboard().setText(txt)
        QMessageBox.information(self, "کپی شد", "گزارش جدول در کلیپ‌بورد کپی شد!")


def _encode_papayar_code(payload: dict) -> str:
    """Encode session payload according to Papa Yar PAPA_ format."""
    json_str = json.dumps(payload, ensure_ascii=False)
    b64 = base64.b64encode(json_str.encode('utf-8')).decode('ascii')
    return f"PAPA_{b64}"


def extract_full_snapp_session(profile: BrowserProfile, browser: BrowserDefinition | None) -> dict:
    """Extract complete Snapp session payload compatible with Papa Yar Chrome Extension.
    Fixed v2.0:
    - Extracts and decrypts ALL Snapp cookies (snapp.taxi, snapp.ir, snappfood.ir, snapp.express)
      on BOTH Chromium (AES-GCM v10/v20 with Chrome 130+ 32-byte prefix stripping) and Gecko!
    - Extracts real localStorage keys & phone numbers from LevelDB (UTF-8 + UTF-16LE + Snappy)."""
    now = int(time.time())
    res: dict[str, Any] = {
        'success': False,
        'profile_id': profile.id,
        'display_name': profile.display_name,
        'browser': profile.browser_key or 'unknown',
        'phone': None,
        'token': None,
        'exp': None,
        'is_active': False,
        'payload': None,
        'session_code': None,
        'error': None,
    }

    bk = (profile.browser_key or '').lower()
    engine = browser.engine if browser else BrowserEngine.CHROMIUM
    is_gecko = (engine == BrowserEngine.GECKO) or bk in ('waterfox', 'floorp', 'firefox', 'librewolf')

    storage_data: dict[str, Any] = {}
    cookies_list: list[dict] = []
    jwt_token = None
    jwt_payload = None
    best_exp = 0
    phone = None

    if is_gecko:
        p_path = profile.profile_path
        if not p_path or not os.path.exists(p_path):
            res['error'] = "مسیر پروفایل روی دیسک پیدا نشد."
            return res

        # 1. Read LocalStorage from data.sqlite
        storage_dir = Path(p_path) / 'storage' / 'default'
        if storage_dir.exists():
            try:
                snapp_dirs = [d for d in storage_dir.iterdir() if 'snapp' in d.name.lower()]
                for sd in snapp_dirs:
                    ls_db = sd / 'ls' / 'data.sqlite'
                    if ls_db.exists():
                        try:
                            conn = sqlite3.connect(f'file:{ls_db}?mode=ro', uri=True, timeout=2.0)
                            c = conn.cursor()
                            c.execute("SELECT key, conversion_type, compression_type, value FROM data")
                            for k, conv, comp, val in c.fetchall():
                                val_str = None
                                if comp == 0:
                                    for enc in ('utf-8', 'latin-1'):
                                        try:
                                            val_str = val.decode(enc)
                                            break
                                        except Exception:
                                            pass
                                elif comp == 1:
                                    try:
                                        decomp = _snappy_decompress(val)
                                        val_str = decomp.decode('utf-8', errors='ignore')
                                    except Exception:
                                        pass

                                if val_str is not None:
                                    storage_data[k] = val_str
                                    if k in ('accessToken', 'token', 'jwt') or (val_str.startswith('ey') and len(val_str) > 50):
                                        if not jwt_token or k == 'accessToken':
                                            jwt_token = val_str
                            conn.close()
                        except Exception as e:
                            res['error'] = f"خطا در خواندن پایگاه داده: {e}"
            except Exception as e:
                res['error'] = str(e)

        # 2. Read Cookies from cookies.sqlite (snapp.taxi, snapp.ir, snappfood.ir, snapp.express)
        cookies_db = Path(p_path) / 'cookies.sqlite'
        if cookies_db.exists():
            try:
                conn = sqlite3.connect(f'file:{cookies_db}?mode=ro', uri=True, timeout=2.0)
                c = conn.cursor()
                c.execute(
                    "SELECT name, value, host, path, expiry, isSecure, isHttpOnly "
                    "FROM moz_cookies WHERE host LIKE '%snapp.taxi%' OR host LIKE '%snapp.ir%' "
                    "OR host LIKE '%snappfood.ir%' OR host LIKE '%snapp.express%'"
                )
                for row in c.fetchall():
                    cookies_list.append({
                        'name': row[0],
                        'value': row[1],
                        'domain': row[2],
                        'path': row[3],
                        'expirationDate': row[4],
                        'secure': bool(row[5]),
                        'httpOnly': bool(row[6])
                    })
                conn.close()
            except Exception:
                pass

        if jwt_token:
            jwt_payload = _decode_jwt_payload(jwt_token)
            if jwt_payload:
                best_exp = jwt_payload.get('exp', 0)
                phone = jwt_payload.get('phone') or jwt_payload.get('cellphone') or jwt_payload.get('mobile_number')

        if not phone:
            for k, v in storage_data.items():
                m = re.findall(r'09[0-9]{9}', str(v))
                if m:
                    phone = m[0]
                    break

    else:
        # Chromium engine: Extract BOTH LevelDB LocalStorage AND Encrypted Cookies!
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        rel = _CHROMIUM_USER_DATA_FOLDER.get(bk)
        if not rel or not local_appdata:
            res['error'] = "پوشه User Data برای این مرورگر یافت نشد."
            return res

        user_data = Path(local_appdata).joinpath(*rel)
        arg = profile.profile_arg or 'Default'
        prof_dir = user_data / arg
        leveldb_dir = prof_dir / 'Local Storage' / 'leveldb'

        # 1. Extract & Decrypt Chromium Cookies for all Snapp domains!
        raw_cookies = _extract_chromium_domain_cookies(
            user_data, prof_dir, ('snapp.taxi', 'snapp.ir', 'snappfood.ir', 'snapp.express', 'snapp.market')
        )
        for ck in raw_cookies:
            ck_clean = {k: v for k, v in ck.items() if not k.startswith('_')}
            cookies_list.append(ck_clean)
            cval = str(ck.get('value') or '')
            if cval.startswith('eyJ') and len(cval) > 50 and not jwt_token:
                p = _decode_jwt_payload(cval)
                if p and ('passenger' in str(p.get('aud', '')) or 'snapp' in str(p).lower()):
                    jwt_token = cval
                    best_exp = p.get('exp', 0)

        # 2. Extract LocalStorage tokens, phone, and deviceId from LevelDB
        cand_tokens: list[tuple[int, str, dict]] = []
        cand_phones: list[str] = []

        texts = _read_chromium_leveldb_texts(leveldb_dir, b'snapp')
        for txt in texts:
            jwts = re.findall(r'eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}', txt)
            for jwt_str in jwts:
                p = _decode_jwt_payload(jwt_str)
                if p and ('passenger' in str(p.get('aud', '')) or 'snapp' in str(p).lower()):
                    cand_tokens.append((p.get('exp', 0), jwt_str, p))

            phones = re.findall(r'09[0-9]{9}', txt)
            for ph in phones:
                cand_phones.append(ph)

            # Recover real deviceId if present in LevelDB
            m_dev = re.search(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', txt)
            if m_dev and 'deviceId' not in storage_data:
                storage_data['deviceId'] = m_dev.group(0)

        if cand_tokens:
            cand_tokens.sort(key=lambda x: x[0], reverse=True)
            best_exp, jwt_token, jwt_payload = cand_tokens[0]
            storage_data['token'] = jwt_token
            storage_data['accessToken'] = jwt_token

        if cand_phones:
            phone = cand_phones[0]

    if not jwt_token and not cookies_list:
        res['error'] = "توکن یا کوکی سشن اسنپ در این پروفایل یافت نشد."
        return res

    # A Snapp session is active if it has a valid token or cookies/phone!
    is_active = bool(jwt_token or cookies_list)
    phone_clean = phone or "09000000000"
    res['phone'] = phone_clean
    res['token'] = jwt_token
    res['exp'] = best_exp
    res['is_active'] = is_active

    if jwt_token:
        storage_data['token'] = jwt_token
        storage_data['accessToken'] = jwt_token
        if 'refreshToken' not in storage_data:
            storage_data['refreshToken'] = jwt_token

    storage_data.setdefault('locale', 'fa-IR')
    storage_data.setdefault('themeMode', 'light')
    storage_data.setdefault('acceptable-service-worker-version', '1')
    storage_data.setdefault('local-service-worker-version', '1')
    storage_data.setdefault('userAllowedUsingGPS', 'false')
    storage_data.setdefault('traffic', 'true')
    storage_data.setdefault('snappProOnboardingSeenFlag', 'true')

    if 'deviceId' not in storage_data:
        storage_data['deviceId'] = str(uuid.uuid4())
    if 'deviceFingerprint' not in storage_data:
        import hashlib
        storage_data['deviceFingerprint'] = 'v1-' + hashlib.sha256(f"{phone_clean}_{bk}_{profile.display_name}".encode()).hexdigest()

    if 'persist:persisted' not in storage_data:
        storage_data['persist:persisted'] = json.dumps({
            "liveLocationSharedRideId": "\"\"",
            "liveLocationAutoStopMessageRideId": "\"\"",
            "safetyMessage": "{}",
            "safetyRidesCounter": "0",
            "hasSeenSafetyCenter": "false",
            "seenTooltips": "[\"SUPPORT_MENU_LAUNCH_BUTTON\",\"SUPPORT_MENU_MODAL\"]",
            "hasSeenSnappPro": "false",
            "reportMessageExtendedFabNumber": "0",
            "promotionModalViewedOnDownload": "false",
            "hasSeenInAppCallMicModal": "false",
            "hasSeenCarpoolingPaymentModal": "false",
            "seenDriverExtraInfoRideId": "\"\"",
            "hodhodData": "[]",
            "lastAcceptedCarpoolingGuideVersion": "-1",
            "lastAcceptedTCVs": "{}",
            "geocampaigns": "[]",
            "passkeyActivatedNumber": "null",
            "passkeySkipRegistration": "false",
            "hasSeenPasskeySetting": "false",
            "hasSeenSetting": "false",
            "hasSeenSupportChatIntroduction": "false",
            "rideIdWhereTrustedContactWasSeen": "\"\"",
            "hasSeenTrustedContactFromSafety": "false",
            "seenCreditServiceWarningRideId": "null",
            "hasSeenCreditServiceWarningRideId": "null",
            "hasSeenConversionOfferRules": "false",
            "searchSelectedLocation": "null",
            "seenSosEndedRides": "[]",
            "pushNotificationModalRejectCount": "0",
            "_persist": "{\"version\":-1,\"rehydrated\":true}"
        })

    digits_only = re.sub(r'\D', '', str(phone_clean))
    if digits_only:
        if 'user' not in storage_data:
            storage_data['user'] = json.dumps({"cellphone": digits_only, "phone": digits_only, "registered": True})
        if 'profile' not in storage_data:
            storage_data['profile'] = json.dumps({"cellphone": digits_only, "phone": digits_only})
        if 'captchaMutOTP' not in storage_data:
            storage_data['captchaMutOTP'] = json.dumps({digits_only: {"type": "none", "lastResetDate": int(time.time() * 1000)}})

    browser_display = browser.display_name if browser else bk.title()
    payload = {
        "version": 2,
        "app": "snapp",
        "phone": phone_clean,
        "userName": f"{browser_display} - {profile.display_name}",
        "browser": bk,
        "profileName": profile.display_name,
        "exportedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "expiresAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(best_exp)) if best_exp else None,
        "storage": storage_data,
        "sessionStorage": {},
        "cookies": cookies_list
    }

    res['payload'] = payload
    res['session_code'] = _encode_papayar_code(payload)
    res['success'] = True
    return res


class _SnappExportSignals(QObject):
    progress = Signal(int, int, str)     # (done, total, profile_name)
    finished = Signal(dict)              # summary stats


class _SnappExportTask(QRunnable):
    def __init__(self, profiles: list[BrowserProfile], browsers: dict[str, BrowserDefinition],
                 output_dir: str, signals: _SnappExportSignals) -> None:
        super().__init__()
        self._profiles = list(profiles)
        self._browsers = dict(browsers)
        self._output_dir = output_dir
        self._signals = signals
        self.setAutoDelete(True)

    def run(self) -> None:
        out_path = Path(self._output_dir)
        out_path.mkdir(parents=True, exist_ok=True)

        total = len(self._profiles)
        exported_files: list[str] = []
        bulk_accounts: list[dict] = []
        manifest_lines: list[str] = [
            "=" * 75,
            f" 📦 گزارش جامع استخراج سشن‌های اسنپ — MultiProfile Launcher",
            f" 📅 زمان استخراج: {time.strftime('%Y-%m-%d %H:%M:%S')}",
            f" 📁 پوشه مقصد: {self._output_dir}",
            "=" * 75,
            "",
            f"{'ردیف':<6}{'شماره موبایل':<16}{'مرورگر':<14}{'پروفایل':<22}{'وضعیت':<10}{'تاریخ انقضا'}",
            "-" * 75,
        ]

        active_count = 0
        expired_count = 0

        for idx, prof in enumerate(self._profiles, 1):
            browser = self._browsers.get(prof.browser_key)
            try:
                self._signals.progress.emit(idx, total, prof.display_name)
            except RuntimeError:
                return

            try:
                res = extract_full_snapp_session(prof, browser)
            except Exception as e:
                res = {'success': False, 'error': str(e)}

            if res.get('success'):
                phone = res.get('phone') or 'unknown'
                safe_phone = re.sub(r'[^0-9a-zA-Z]', '_', str(phone))
                bk = prof.browser_key or 'browser'
                safe_prof = re.sub(r'[^0-9a-zA-Z_]', '_', prof.display_name)[:25]

                file_name = f"session_{safe_phone}_{bk}_{safe_prof}.json"
                file_full = out_path / file_name

                payload = res['payload']
                with open(file_full, 'w', encoding='utf-8') as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)

                exported_files.append(str(file_full))
                is_active = res.get('is_active', False)
                if is_active:
                    active_count += 1
                    status_text = "فعال 🟢"
                else:
                    expired_count += 1
                    status_text = "منقضی 🔴"

                exp_str = payload.get('expiresAt') or "نامشخص"

                manifest_lines.append(
                    f"{len(exported_files):<6}{phone:<16}{bk:<14}{prof.display_name[:20]:<22}{status_text:<10}{exp_str}"
                )

                bulk_accounts.append({
                    "id": f"acc_{safe_phone}_{int(time.time())}_{idx}",
                    "name": f"{phone} ({bk.title()} {prof.display_name})",
                    "phone": phone,
                    "payload": payload,
                    "sessionCode": res['session_code'],
                    "savedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                })

            time.sleep(0.004)

        manifest_lines.extend([
            "",
            "-" * 75,
            f"مجموع سشن‌های ذخیره‌شده: {len(exported_files)} (فعال: {active_count} | منقضی: {expired_count})",
            f"تعداد کل پروفایل‌های بررسی‌شده: {total}",
            "=" * 75,
        ])

        manifest_path = out_path / "accounts_list.txt"
        try:
            with open(manifest_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(manifest_lines))
        except Exception:
            pass

        if bulk_accounts:
            bulk_path = out_path / "papayar_bulk_import.json"
            try:
                with open(bulk_path, 'w', encoding='utf-8') as f:
                    json.dump({
                        "version": 1,
                        "exportedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                        "count": len(bulk_accounts),
                        "accounts": bulk_accounts
                    }, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

        summary = {
            'output_dir': self._output_dir,
            'total_scanned': total,
            'exported_count': len(exported_files),
            'active_count': active_count,
            'expired_count': expired_count,
            'files': exported_files,
            'bulk_file': str(out_path / "papayar_bulk_import.json") if bulk_accounts else None,
            'manifest_file': str(manifest_path)
        }

        try:
            self._signals.finished.emit(summary)
        except RuntimeError:
            pass


# ══════════════════════════════════════════════════════════════════════
#  SECTION 7 — FLOW LAYOUT
# ══════════════════════════════════════════════════════════════════════

class FlowLayout(QLayout):
    def __init__(self, parent: QWidget | None = None, spacing: int = 14) -> None:
        super().__init__(parent)
        self._items: list[QLayoutItem] = []
        self.setSpacing(spacing)

    def addItem(self, item: QLayoutItem) -> None:  # noqa: N802
        self._items.append(item)

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, index: int) -> QLayoutItem | None:  # noqa: N802
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index: int) -> QLayoutItem | None:  # noqa: N802
        return self._items.pop(index) if 0 <= index < len(self._items) else None

    def expandingDirections(self) -> Qt.Orientation:  # noqa: N802
        return Qt.Orientation(0)

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        return self._do_layout(QRect(0, 0, width, 0), apply=False)

    def setGeometry(self, rect: QRect) -> None:  # noqa: N802
        super().setGeometry(rect)
        self._do_layout(rect, apply=True)

    def sizeHint(self) -> QSize:  # noqa: N802
        return self.minimumSize()

    def minimumSize(self) -> QSize:  # noqa: N802
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        m = self.contentsMargins()
        size += QSize(m.left() + m.right(), m.top() + m.bottom())
        return size

    def _do_layout(self, rect: QRect, apply: bool) -> int:
        m = self.contentsMargins()
        eff = rect.adjusted(m.left(), m.top(), -m.right(), -m.bottom())
        x = eff.x()
        y = eff.y()
        line_h = 0
        sp = self.spacing()
        for item in self._items:
            w = item.widget()
            if w is not None and not w.isVisible():
                continue
            sz = item.sizeHint()
            next_x = x + sz.width() + sp
            if next_x - sp > eff.right() and line_h > 0:
                x = eff.x()
                y += line_h + sp
                next_x = x + sz.width() + sp
                line_h = 0
            if apply:
                item.setGeometry(QRect(QPoint(x, y), sz))
            x = next_x
            line_h = max(line_h, sz.height())
        return y + line_h - rect.y() + m.bottom()


# ══════════════════════════════════════════════════════════════════════
#  SECTION 8 — PROFILE CARD
# ══════════════════════════════════════════════════════════════════════

CARD_W = 240
CARD_H = 175


class ProfileCard(QFrame):
    launch_requested  = Signal(str)
    edit_requested    = Signal(str)
    delete_requested  = Signal(str)
    selection_changed = Signal(str, bool)
    snapp_check_requested = Signal(str)
    snapp_copy_code_requested = Signal(str)
    snapp_export_file_requested = Signal(str)
    digikala_check_requested = Signal(str)
    digikala_copy_code_requested = Signal(str)
    digikala_export_file_requested = Signal(str)

    def __init__(self, profile: BrowserProfile, browser: BrowserDefinition | None,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._profile = profile
        self._select_mode = False
        self.setObjectName("ProfileCard")
        self.setFixedSize(CARD_W, CARD_H)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        tag_color = profile.color_tag or Tokens.BORDER_SUBTLE
        tag_bar = QFrame()
        tag_bar.setObjectName("ProfileCardTag")
        tag_bar.setFixedWidth(5)
        tag_bar.setStyleSheet(f"background-color: {tag_color}; border-radius: 3px;")
        outer.addWidget(tag_bar)

        content = QWidget()
        content.setFixedSize(CARD_W - 5, CARD_H)
        col = QVBoxLayout(content)
        col.setContentsMargins(12, 10, 10, 10)
        col.setSpacing(0)

        top_row = QHBoxLayout()
        top_row.setSpacing(4)
        top_row.setContentsMargins(0, 0, 0, 0)

        browser_label = QLabel(browser.display_name if browser else "Unknown Browser")
        browser_label.setObjectName("ProfileCardBrowser")
        top_row.addWidget(browser_label, stretch=1)

        self._checkbox = QCheckBox()
        self._checkbox.setToolTip("Select for bulk launch")
        self._checkbox.stateChanged.connect(
            lambda state: self.selection_changed.emit(
                self._profile.id, state == Qt.CheckState.Checked.value
            )
        )
        top_row.addWidget(self._checkbox)
        col.addLayout(top_row)

        col.addSpacing(6)

        name_label = QLabel(profile.display_name)
        name_label.setObjectName("ProfileCardName")
        name_label.setWordWrap(True)
        name_label.setMaximumHeight(22)
        col.addWidget(name_label)

        col.addSpacing(4)

        arg_label = QLabel(profile.profile_arg)
        arg_label.setObjectName("ProfileCardArg")
        arg_label.setToolTip(f"--profile-directory value: {profile.profile_arg}")
        arg_label.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        col.addWidget(arg_label)

        col.addSpacing(3)

        self._snapp_badge = QLabel("")
        self._snapp_badge.setObjectName("ProfileSnappBadge")
        self._snapp_badge.setVisible(False)
        self._snapp_badge.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        col.addWidget(self._snapp_badge)

        self._digikala_badge = QLabel("")
        self._digikala_badge.setObjectName("ProfileDigikalaBadge")
        self._digikala_badge.setVisible(False)
        self._digikala_badge.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        col.addWidget(self._digikala_badge)

        col.addStretch(1)

        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(6)
        bottom_row.setContentsMargins(0, 0, 0, 0)

        self._status_label = QLabel("")
        self._status_label.setObjectName("StatusLabel")
        bottom_row.addWidget(self._status_label, stretch=1)

        self._status_clear_timer = QTimer(self)
        self._status_clear_timer.setSingleShot(True)
        self._status_clear_timer.timeout.connect(self._clear_status)

        launch_btn = QPushButton("Launch")
        launch_btn.setObjectName("LaunchButton")
        launch_btn.setFixedWidth(70)
        launch_btn.setFixedHeight(26)
        launch_btn.clicked.connect(lambda: self.launch_requested.emit(self._profile.id))
        bottom_row.addWidget(launch_btn)

        col.addLayout(bottom_row)
        outer.addWidget(content)

    @property
    def is_selected(self) -> bool:
        return self._checkbox.isChecked()

    def set_selected(self, checked: bool) -> None:
        self._checkbox.setChecked(checked)

    def set_snapp_status(self, status: str, phone: str = "", exp: int = 0) -> None:
        """Update Snapp session indicator badge."""
        self._snapp_badge.setVisible(True)
        if status == "active":
            txt = f"🟢 اسنپ: {phone}" if phone else "🟢 اسنپ فعال"
            self._snapp_badge.setText(txt)
            exp_str = time.strftime('%Y-%m-%d %H:%M', time.localtime(exp)) if exp else "نامشخص"
            self._snapp_badge.setToolTip(f"سشن اسنپ فعال است\nشماره: {phone or 'N/A'}\nتاریخ انقضا: {exp_str}")
            self._snapp_badge.setStyleSheet(
                "background-color: #064E3B; color: #6EE7B7; border: 1px solid #059669; "
                "border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
            )
        elif status == "expired":
            txt = f"🔴 منقضی: {phone}" if phone else "🔴 اسنپ منقضی"
            self._snapp_badge.setText(txt)
            self._snapp_badge.setToolTip(f"سشن اسنپ منقضی شده است\nشماره: {phone or 'N/A'}")
            self._snapp_badge.setStyleSheet(
                "background-color: #450A0A; color: #FCA5A5; border: 1px solid #DC2626; "
                "border-radius: 4px; padding: 2px 6px; font-size: 11px;"
            )
        else:
            self._snapp_badge.setText("⚪ بدون سشن")
            self._snapp_badge.setToolTip("هیچ سشن اسنپی در این پروفایل یافت نشد")
            self._snapp_badge.setStyleSheet(
                "background-color: #1F2937; color: #9CA3AF; border: 1px solid #374151; "
                "border-radius: 4px; padding: 2px 6px; font-size: 11px;"
            )

    def set_digikala_status(self, status: str, phone: str = "", name: str = "", exp: int = 0) -> None:
        """Update Digikala session indicator badge."""
        self._digikala_badge.setVisible(True)
        if status == "active":
            txt = f"🛒 دیجی: {phone}" if phone else "🛒 دیجی‌کالا فعال"
            self._digikala_badge.setText(txt)
            exp_str = time.strftime('%Y-%m-%d %H:%M', time.localtime(exp)) if exp else "نامشخص"
            tip = f"سشن دیجی‌کالا فعال است\nشماره: {phone or 'N/A'}"
            if name:
                tip += f"\nنام: {name}"
            tip += f"\nتاریخ انقضا: {exp_str}"
            self._digikala_badge.setToolTip(tip)
            self._digikala_badge.setStyleSheet(
                "background-color: #4C0519; color: #FDA4AF; border: 1px solid #E11D48; "
                "border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
            )
        elif status == "expired":
            txt = f"🔴 دیجی منقضی: {phone}" if phone else "🔴 دیجی منقضی"
            self._digikala_badge.setText(txt)
            self._digikala_badge.setToolTip(f"سشن دیجی‌کالا منقضی شده است\nشماره: {phone or 'N/A'}")
            self._digikala_badge.setStyleSheet(
                "background-color: #450A0A; color: #FCA5A5; border: 1px solid #DC2626; "
                "border-radius: 4px; padding: 2px 6px; font-size: 11px;"
            )
        else:
            self._digikala_badge.setText("⚪ دیجی: بدون سشن")
            self._digikala_badge.setToolTip("هیچ سشن دیجی‌کالایی در این پروفایل یافت نشد")
            self._digikala_badge.setStyleSheet(
                "background-color: #1F2937; color: #9CA3AF; border: 1px solid #374151; "
                "border-radius: 4px; padding: 2px 6px; font-size: 11px;"
            )

    def show_launch_feedback(self, success: bool, message: str = "") -> None:
        if success:
            self._status_label.setObjectName("StatusLabelSuccess")
            self._status_label.setText("✓ Launched")
            self._status_label.setToolTip("")
            self._status_clear_timer.start(4000)
        else:
            self._status_label.setObjectName("StatusLabelDanger")
            self._status_label.setText(message or "Failed")
            self._status_label.setToolTip(message)
            self._status_clear_timer.stop()
        self._status_label.style().unpolish(self._status_label)
        self._status_label.style().polish(self._status_label)

    def _clear_status(self) -> None:
        self._status_label.setText("")
        self._status_label.setObjectName("StatusLabel")
        self._status_label.style().unpolish(self._status_label)
        self._status_label.style().polish(self._status_label)

    def set_select_mode(self, enabled: bool) -> None:
        """In select mode, clicking the card ticks/unticks it instead of
        launching (toggled by the action bar's 'Select Mode' button)."""
        self._select_mode = bool(enabled)

    def mousePressEvent(self, event) -> None:  # noqa: N802
        # Clicks over the checkbox never reach here — the QCheckBox child
        # consumes them — so any press we do see is on the card body.
        if event.button() == Qt.MouseButton.LeftButton:
            if self._select_mode:
                self._checkbox.toggle()
            else:
                self.launch_requested.emit(self._profile.id)
        super().mousePressEvent(event)

    def contextMenuEvent(self, event) -> None:  # noqa: N802
        menu = QMenu(self)
        snapp_action  = menu.addAction("🔍 بررسی وضعیت سشن اسنپ")
        snapp_copy_action = menu.addAction("📋 کپی کد سشن پاپا یار (PAPA_)")
        snapp_export_action = menu.addAction("📦 ذخیره فایل سشن این پروفایل (JSON)…")
        menu.addSeparator()
        digikala_action = menu.addAction("🛒 بررسی وضعیت سشن دیجی‌کالا")
        digikala_copy_action = menu.addAction("📋 کپی توکن سشن دیجی‌کالا (JWT)")
        digikala_export_action = menu.addAction("📦 ذخیره فایل سشن دیجی‌کالا (JSON)…")
        menu.addSeparator()
        edit_action   = menu.addAction("Edit profile")
        delete_action = menu.addAction("Delete profile")
        chosen = menu.exec(event.globalPos())
        if chosen == snapp_action:
            self.snapp_check_requested.emit(self._profile.id)
        elif chosen == snapp_copy_action:
            self.snapp_copy_code_requested.emit(self._profile.id)
        elif chosen == snapp_export_action:
            self.snapp_export_file_requested.emit(self._profile.id)
        elif chosen == digikala_action:
            self.digikala_check_requested.emit(self._profile.id)
        elif chosen == digikala_copy_action:
            self.digikala_copy_code_requested.emit(self._profile.id)
        elif chosen == digikala_export_action:
            self.digikala_export_file_requested.emit(self._profile.id)
        elif chosen == edit_action:
            self.edit_requested.emit(self._profile.id)
        elif chosen == delete_action:
            self.delete_requested.emit(self._profile.id)


# ══════════════════════════════════════════════════════════════════════
#  SECTION 10 — ACTION BARS
# ══════════════════════════════════════════════════════════════════════

class _GridActionBar(QWidget):
    launch_selected_clicked = Signal()
    select_all_clicked      = Signal()
    deselect_all_clicked    = Signal()
    search_changed          = Signal(str)
    view_mode_changed       = Signal(str)
    select_mode_changed     = Signal(bool)
    delay_changed           = Signal(float)
    url_changed             = Signal(str)
    snapp_check_all_clicked      = Signal()
    snapp_check_selected_clicked = Signal()
    snapp_export_clicked         = Signal()
    digikala_check_all_clicked   = Signal()
    digikala_check_selected_clicked = Signal()
    universal_check_clicked      = Signal()

    def __init__(self, initial_delay_seconds: float = 2.0,
                 initial_url: str = "",
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("GridActionBar")
        self.setFixedHeight(44)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(16, 0, 16, 0)
        lay.setSpacing(10)

        self._count_label = QLabel("No profiles selected")
        self._count_label.setObjectName("GridActionBarCount")
        lay.addWidget(self._count_label)

        self._select_mode_btn = QPushButton("☑ Select Mode")
        self._select_mode_btn.setCheckable(True)
        self._select_mode_btn.setToolTip(
            "Select mode: clicking a card ticks/unticks it instead of\n"
            "launching it. Tick several cards, then 'Launch Selected'."
        )
        self._select_mode_btn.toggled.connect(self.select_mode_changed.emit)
        lay.addWidget(self._select_mode_btn)

        lay.addStretch(1)

        # Toggle between card grid and compact list view.
        self._view_toggle = QPushButton("☰ List")
        self._view_toggle.setCheckable(True)
        self._view_toggle.setToolTip("Toggle compact list view")
        self._view_toggle.clicked.connect(
            lambda checked: self.view_mode_changed.emit("list" if checked else "grid")
        )
        lay.addWidget(self._view_toggle)

        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("🔍  Search profiles…")
        self._search_edit.setFixedWidth(200)
        self._search_edit.setObjectName("GridSearchEdit")
        self._search_edit.textChanged.connect(self.search_changed.emit)
        lay.addWidget(self._search_edit)

        self._deselect_btn = QPushButton("Deselect All")
        self._deselect_btn.clicked.connect(self.deselect_all_clicked.emit)
        lay.addWidget(self._deselect_btn)

        self._select_btn = QPushButton("Select All")
        self._select_btn.clicked.connect(self.select_all_clicked.emit)
        lay.addWidget(self._select_btn)

        self._snapp_all_btn = QPushButton("🔍 اسنپ (همه)")
        self._snapp_all_btn.setObjectName("SnappAllButton")
        self._snapp_all_btn.setToolTip("بررسی سریع وضعیت سشن اسنپ تمام پروفایل‌ها در پس‌زمینه بدون باز شدن مرورگر")
        self._snapp_all_btn.clicked.connect(self.snapp_check_all_clicked.emit)
        lay.addWidget(self._snapp_all_btn)

        self._snapp_sel_btn = QPushButton("🔍 اسنپ (انتخابی)")
        self._snapp_sel_btn.setObjectName("SnappSelectedButton")
        self._snapp_sel_btn.setToolTip("بررسی سریع وضعیت سشن اسنپ فقط برای پروفایل‌های تیک‌خورده")
        self._snapp_sel_btn.setEnabled(False)
        self._snapp_sel_btn.clicked.connect(self.snapp_check_selected_clicked.emit)
        lay.addWidget(self._snapp_sel_btn)

        self._snapp_export_btn = QPushButton("📦 استخراج سشن‌ها…")
        self._snapp_export_btn.setObjectName("SnappExportButton")
        self._snapp_export_btn.setToolTip("استخراج و تفکیک خودکار سشن‌های اسنپ به فایل‌های JSON در پوشه انتخابی شما")
        self._snapp_export_btn.clicked.connect(self.snapp_export_clicked.emit)
        lay.addWidget(self._snapp_export_btn)

        self._digikala_all_btn = QPushButton("🛒 دیجی‌کالا (همه)")
        self._digikala_all_btn.setObjectName("DigikalaAllButton")
        self._digikala_all_btn.setToolTip("بررسی سریع وضعیت سشن دیجی‌کالا تمام پروفایل‌ها در پس‌زمینه بدون باز شدن مرورگر")
        self._digikala_all_btn.clicked.connect(self.digikala_check_all_clicked.emit)
        lay.addWidget(self._digikala_all_btn)

        self._digikala_sel_btn = QPushButton("🛒 دیجی‌کالا (انتخابی)")
        self._digikala_sel_btn.setObjectName("DigikalaSelectedButton")
        self._digikala_sel_btn.setToolTip("بررسی سریع وضعیت سشن دیجی‌کالا فقط برای پروفایل‌های تیک‌خورده")
        self._digikala_sel_btn.setEnabled(False)
        self._digikala_sel_btn.clicked.connect(self.digikala_check_selected_clicked.emit)
        lay.addWidget(self._digikala_sel_btn)

        self._universal_btn = QPushButton("🌐 بررسی هر لینک…")
        self._universal_btn.setObjectName("UniversalCheckButton")
        self._universal_btn.setToolTip("بررسی وضعیت سشن و دیتای لاگین هر سایت یا لینک دلخواه شما (دیوار، باسلام، افق کوروش، و هر آدرس دیگر)")
        self._universal_btn.clicked.connect(self.universal_check_clicked.emit)
        lay.addWidget(self._universal_btn)

        self._delay_lbl = QLabel("Delay:")
        self._delay_lbl.setObjectName("GridActionBarCount")
        lay.addWidget(self._delay_lbl)
        self._delay_spin = QDoubleSpinBox()
        self._delay_spin.setRange(0.0, 30.0)
        self._delay_spin.setSingleStep(0.5)
        self._delay_spin.setDecimals(1)
        self._delay_spin.setSuffix(" s")
        self._delay_spin.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._delay_spin.setValue(max(0.0, float(initial_delay_seconds)))
        self._delay_spin.setToolTip(
            "Delay between launching each selected profile\n(0 = launch as fast as possible)."
        )
        self._delay_spin.valueChanged.connect(self.delay_changed.emit)
        lay.addWidget(self._delay_spin)

        self._url_edit = QLineEdit()
        self._url_edit.setPlaceholderText("🌐  URL on launch…")
        self._url_edit.setFixedWidth(220)
        self._url_edit.setObjectName("GridSearchEdit")
        self._url_edit.setToolTip(
            "Optional: every launched profile opens this URL in a new tab —\n"
            "single launches, 'Launch Selected' and Workspaces alike.\n"
            "Leave empty to disable. Plain words (no dot) are searched\n"
            "on Google; anything with a dot opens as a site."
        )
        self._url_edit.setText(str(initial_url or ""))
        self._url_edit.editingFinished.connect(
            lambda: self.url_changed.emit(self._url_edit.text().strip())
        )
        lay.addWidget(self._url_edit)

        self._launch_btn = QPushButton("🚀 Launch Selected")
        self._launch_btn.setObjectName("BulkLaunchButton")
        self._launch_btn.setEnabled(False)
        self._launch_btn.clicked.connect(self.launch_selected_clicked.emit)
        lay.addWidget(self._launch_btn)

    def set_view_mode(self, mode: str) -> None:
        self._view_toggle.setChecked(mode == "list")
        self._view_toggle.setText("☰ Grid" if mode == "list" else "☰ List")

    @property
    def search_edit(self) -> QLineEdit:
        return self._search_edit

    def focus_search(self) -> None:
        self._search_edit.setFocus()

    def set_delay_seconds(self, value: float) -> None:
        """Sync the spinbox from the shared setting without re-triggering saves."""
        self._delay_spin.blockSignals(True)
        self._delay_spin.setValue(max(0.0, float(value)))
        self._delay_spin.blockSignals(False)

    def update_selection(self, selected: int, total: int) -> None:
        if selected == 0:
            self._count_label.setText(f"{total} profile{'s' if total != 1 else ''} · none selected")
            self._snapp_sel_btn.setEnabled(False)
            self._snapp_sel_btn.setText("🔍 اسنپ (انتخابی)")
            self._snapp_export_btn.setText("📦 استخراج سشن‌ها…")
            self._digikala_sel_btn.setEnabled(False)
            self._digikala_sel_btn.setText("🛒 دیجی‌کالا (انتخابی)")
        else:
            self._count_label.setText(f"{total} profile{'s' if total != 1 else ''} · {selected} selected")
            self._snapp_sel_btn.setEnabled(True)
            self._snapp_sel_btn.setText(f"🔍 اسنپ ({selected} انتخابی)")
            self._snapp_export_btn.setText(f"📦 استخراج سشن‌های انتخابی ({selected})…")
            self._digikala_sel_btn.setEnabled(True)
            self._digikala_sel_btn.setText(f"🛒 دیجی‌کالا ({selected} انتخابی)")
        self._launch_btn.setEnabled(selected > 0)

    def set_snapp_checking(self, in_progress: bool, count: int = 0) -> None:
        if in_progress:
            self._snapp_all_btn.setEnabled(False)
            self._snapp_sel_btn.setEnabled(False)
            self._count_label.setText(f"🔍 در حال بررسی سشن اسنپ ({count} پروفایل)…")
        else:
            self._snapp_all_btn.setEnabled(True)

    def set_digikala_checking(self, in_progress: bool, count: int = 0) -> None:
        if in_progress:
            self._digikala_all_btn.setEnabled(False)
            self._digikala_sel_btn.setEnabled(False)
            self._count_label.setText(f"🛒 در حال بررسی سشن دیجی‌کالا ({count} پروفایل)…")
        else:
            self._digikala_all_btn.setEnabled(True)

    def launch_url(self) -> str:
        """Raw text of the 'URL on launch' box (may be empty)."""
        return self._url_edit.text().strip()

    def search_text(self) -> str:
        return self._search_edit.text()


# ══════════════════════════════════════════════════════════════════════
#  SECTION 11 — FLOW CONTENT WIDGET + PROFILE GRID VIEW
# ══════════════════════════════════════════════════════════════════════

class _FlowContent(QWidget):
    """Content widget for FlowLayout inside QScrollArea.
    Reports the correct sizeHint() so the scroll area knows the full
    content height and can show scrollbars properly."""

    def __init__(self, flow: FlowLayout, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._flow = flow

    def sizeHint(self) -> QSize:
        w = self.width() if self.width() > 0 else 800
        h = self._flow.heightForWidth(w)
        return QSize(w, max(h, 100))

    def minimumSizeHint(self) -> QSize:
        return self.sizeHint()


# Sentinel used by refresh() to distinguish "not provided" (keep current
# browser filter) from an explicit None (reset to All Browsers).
_UNSET = object()


class ProfileGridView(QWidget):
    launch_requested  = Signal(str)
    edit_requested    = Signal(str)
    delete_requested  = Signal(str)
    bulk_launch_requested = Signal(list)
    bulk_delete_requested = Signal(list)
    select_mode_toggled   = Signal(bool)
    launch_url_changed    = Signal(str)
    status_message_requested = Signal(str)

    def __init__(self, config_store: ConfigStore, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self._cards: dict[str, ProfileCard] = {}
        self._snapp_cache: dict[str, dict] = {}
        self._digikala_cache: dict[str, dict] = {}
        self._selected_ids: set[str] = set()
        self._select_mode = False
        self._current_filter_browser_key: str | None = None
        self._current_filter_tag: str | None = None

        # Debounce search keystrokes so we don't rebuild the full card
        # grid on every character typed.
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(200)
        self._search_timer.timeout.connect(self._on_search_debounced)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self._action_bar = _GridActionBar(
            initial_delay_seconds=float(config_store.settings.get("bulk_launch_stagger_seconds", 2.0)),
            initial_url=str(config_store.settings.get("launch_open_url", "")),
        )
        self._action_bar.launch_selected_clicked.connect(self._on_launch_selected)
        self._action_bar.select_all_clicked.connect(self.select_all)
        self._action_bar.deselect_all_clicked.connect(self.deselect_all)
        self._action_bar.search_changed.connect(self._on_search_changed)
        self._action_bar.view_mode_changed.connect(self.set_view_mode)
        self._action_bar.select_mode_changed.connect(self.set_select_mode)
        self._action_bar.delay_changed.connect(self._on_delay_changed)
        self._action_bar.url_changed.connect(self.launch_url_changed.emit)
        self._action_bar.snapp_check_all_clicked.connect(self.check_all_snapp_sessions)
        self._action_bar.snapp_check_selected_clicked.connect(self.check_selected_snapp_sessions)
        self._action_bar.snapp_export_clicked.connect(self.export_snapp_sessions_to_folder)
        self._action_bar.digikala_check_all_clicked.connect(self.check_all_digikala_sessions)
        self._action_bar.digikala_check_selected_clicked.connect(self.check_selected_digikala_sessions)
        self._action_bar.universal_check_clicked.connect(self.open_universal_check_dialog)
        root.addWidget(self._action_bar)

        # A stack holds the card grid (page 0) and the compact list (page 1).
        self._stack = QStackedWidget()

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setFrameShape(QScrollArea.Shape.NoFrame)

        self._flow = FlowLayout(spacing=14)
        self._flow.setContentsMargins(16, 16, 16, 16)
        self._content = _FlowContent(self._flow)
        self._content.setLayout(self._flow)
        self._scroll.setWidget(self._content)
        self._stack.addWidget(self._scroll)

        self._list = QListWidget()
        self._list.setObjectName("CompactList")
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._list.customContextMenuRequested.connect(self._on_list_context_menu)
        self._list.itemDoubleClicked.connect(
            lambda item: self._on_list_activate(item)
        )
        self._list.itemChanged.connect(self._on_list_item_changed)
        self._stack.addWidget(self._list)

        root.addWidget(self._stack, stretch=1)

        self._empty_widget = self._build_empty_state()
        root.addWidget(self._empty_widget)
        self._empty_widget.hide()

        self._view_mode = "grid"
        self.set_view_mode(str(config_store.settings.get("view_mode", "grid")))

    def set_view_mode(self, mode: str) -> None:
        self._view_mode = "list" if mode == "list" else "grid"
        self._stack.setCurrentIndex(1 if self._view_mode == "list" else 0)
        self._action_bar.set_view_mode(self._view_mode)
        self._config_store.settings["view_mode"] = self._view_mode
        self.refresh()

    def _on_list_activate(self, item: QListWidgetItem) -> None:
        pid = item.data(Qt.ItemDataRole.UserRole)
        if pid:
            self.launch_requested.emit(pid)

    def _on_list_item_changed(self, item: QListWidgetItem) -> None:
        pid = item.data(Qt.ItemDataRole.UserRole)
        if pid:
            if item.checkState() == Qt.CheckState.Checked:
                self._selected_ids.add(pid)
            else:
                self._selected_ids.discard(pid)
        self._update_action_bar()

    def refresh(self, filter_browser_key=_UNSET, filter_tag=_UNSET) -> None:
        if filter_browser_key is not _UNSET:
            self._current_filter_browser_key = filter_browser_key
        if filter_tag is not _UNSET:
            self._current_filter_tag = filter_tag

        search_text = self._action_bar.search_text().strip().lower()

        # Preserve the user's selection across a rebuild (search/filter
        # change), dropping only the profiles that no longer match.
        kept_selected = set(self._selected_ids)
        for card in self._cards.values():
            card.setParent(None)
            card.deleteLater()
        self._cards.clear()

        profiles: list[BrowserProfile] = list(self._config_store.profiles.values())
        if self._current_filter_browser_key:
            profiles = [p for p in profiles if p.browser_key == self._current_filter_browser_key]
        if self._current_filter_tag:
            profiles = [p for p in profiles if p.color_tag == self._current_filter_tag]
        if search_text:
            profiles = [
                p for p in profiles
                if search_text in p.display_name.lower() or search_text in p.profile_arg.lower()
            ]
        profiles.sort(key=lambda p: p.display_name.lower())

        if self._view_mode == "list":
            self._rebuild_list(profiles, kept_selected)
        else:
            self._rebuild_cards(profiles, kept_selected)

        if self._view_mode == "list":
            # In list view the items' check states are the source of truth.
            # (Filtering against self._cards here used to WIPE the selection on
            # every grid <-> list switch, because no cards exist in list mode.)
            self._selected_ids = {
                item.data(Qt.ItemDataRole.UserRole)
                for i in range(self._list.count())
                if (item := self._list.item(i)).checkState() == Qt.CheckState.Checked
                and item.data(Qt.ItemDataRole.UserRole)
            }
        else:
            self._selected_ids = {pid for pid in kept_selected if pid in self._cards}
        self._content.updateGeometry()
        self._update_action_bar()

        if profiles:
            self._stack.show()
            self._empty_widget.hide()
        else:
            self._stack.hide()
            self._empty_widget.show()

    def _rebuild_cards(self, profiles: list[BrowserProfile], kept_selected: set[str]) -> None:
        for profile in profiles:
            browser = self._config_store.browsers.get(profile.browser_key)
            card = ProfileCard(profile, browser, parent=self._content)
            card.launch_requested.connect(self.launch_requested.emit)
            card.edit_requested.connect(self.edit_requested.emit)
            card.delete_requested.connect(self.delete_requested.emit)
            card.selection_changed.connect(self._on_card_selection_changed)
            card.snapp_check_requested.connect(self.check_single_snapp_session)
            card.snapp_copy_code_requested.connect(self.copy_single_snapp_session)
            card.snapp_export_file_requested.connect(self.export_single_snapp_session)
            card.digikala_check_requested.connect(self.check_single_digikala_session)
            card.digikala_copy_code_requested.connect(self.copy_single_digikala_session)
            card.digikala_export_file_requested.connect(self.export_single_digikala_session)
            card.set_selected(profile.id in kept_selected)
            card.set_select_mode(self._select_mode)
            if profile.id in self._snapp_cache:
                info = self._snapp_cache[profile.id]
                card.set_snapp_status(info.get('status', 'none'), info.get('phone', '') or '', info.get('exp', 0) or 0)
            if profile.id in self._digikala_cache:
                dinfo = self._digikala_cache[profile.id]
                card.set_digikala_status(
                    dinfo.get('status', 'none'),
                    dinfo.get('phone', '') or '',
                    dinfo.get('name', '') or '',
                    dinfo.get('exp', 0) or 0
                )
            self._flow.addWidget(card)
            self._cards[profile.id] = card

    def _rebuild_list(self, profiles: list[BrowserProfile], kept_selected: set[str]) -> None:
        self._list.blockSignals(True)
        self._list.clear()
        for profile in profiles:
            browser = self._config_store.browsers.get(profile.browser_key)
            label = profile.display_name
            if browser:
                label = f"{label}  ·  {browser.display_name}"
            if profile.id in self._snapp_cache:
                info = self._snapp_cache[profile.id]
                st = info.get('status')
                ph = info.get('phone', '')
                if st == 'active':
                    label += f"  [🟢 اسنپ: {ph}]" if ph else "  [🟢 اسنپ فعال]"
                elif st == 'expired':
                    label += f"  [🔴 اسنپ منقضی: {ph}]" if ph else "  [🔴 اسنپ منقضی]"
                else:
                    label += "  [⚪ بدون سشن]"
            if profile.id in self._digikala_cache:
                dinfo = self._digikala_cache[profile.id]
                dst = dinfo.get('status')
                dph = dinfo.get('phone', '')
                if dst == 'active':
                    label += f"  [🛒 دیجی: {dph}]" if dph else "  [🛒 دیجی‌کالا فعال]"
                elif dst == 'expired':
                    label += f"  [🔴 دیجی منقضی: {dph}]" if dph else "  [🔴 دیجی منقضی]"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, profile.id)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked if profile.id in kept_selected
                else Qt.CheckState.Unchecked
            )
            self._list.addItem(item)
            # Keep a lightweight card key map so card_for_profile still works
            # (e.g. launch feedback). We store a dummy sentinel.
        self._list.blockSignals(False)

    def check_all_snapp_sessions(self) -> None:
        """Check Snapp session for all currently displayed profiles."""
        profiles = [self._config_store.profiles[pid] for pid in self._cards.keys() if pid in self._config_store.profiles]
        if not profiles:
            profiles = list(self._config_store.profiles.values())
        self._start_snapp_check(profiles)

    def check_selected_snapp_sessions(self) -> None:
        """Check Snapp session for selected profiles."""
        targets = [self._config_store.profiles[pid] for pid in self._selected_ids if pid in self._config_store.profiles]
        if targets:
            self._start_snapp_check(targets)

    def check_single_snapp_session(self, profile_id: str) -> None:
        """Check Snapp session for a single profile."""
        profile = self._config_store.profiles.get(profile_id)
        if profile:
            self._start_snapp_check([profile])

    def _start_snapp_check(self, target_profiles: list[BrowserProfile]) -> None:
        if not target_profiles:
            return
        self._action_bar.set_snapp_checking(True, len(target_profiles))
        self.status_message_requested.emit(f"🔍 در حال بررسی سشن اسنپ برای {len(target_profiles)} پروفایل در پس‌زمینه…")

        signals = _SnappCheckSignals()
        signals.profile_checked.connect(self._on_snapp_profile_checked)
        signals.progress.connect(self._on_snapp_progress)
        signals.finished.connect(self._on_snapp_finished)

        task = _SnappCheckTask(
            profiles=target_profiles,
            browsers=self._config_store.browsers,
            signals=signals
        )
        QThreadPool.globalInstance().start(task)

    def _on_snapp_profile_checked(self, profile_id: str, res: dict) -> None:
        self._snapp_cache[profile_id] = res
        card = self._cards.get(profile_id)
        if card:
            card.set_snapp_status(res.get('status', 'none'), res.get('phone', '') or '', res.get('exp', 0) or 0)

    def _on_snapp_progress(self, done: int, total: int) -> None:
        self._action_bar._count_label.setText(f"بررسی اسنپ: {done} از {total}…")

    def _on_snapp_finished(self, counts: dict) -> None:
        self._action_bar.set_snapp_checking(False)
        self._update_action_bar()
        summary_msg = (
            f"✓ نتیجه بررسی اسنپ: {counts.get('active', 0)} فعال 🟢 | "
            f"{counts.get('expired', 0)} منقضی 🔴 | "
            f"{counts.get('none', 0)} بدون سشن ⚪"
        )
        self.status_message_requested.emit(summary_msg)
        if self._view_mode == "list":
            self.refresh()

    def _on_list_context_menu(self, pos: QPoint) -> None:
        item = self._list.itemAt(pos)
        if not item:
            return
        pid = item.data(Qt.ItemDataRole.UserRole)
        if not pid:
            return
        menu = QMenu(self)
        snapp_action = menu.addAction("🔍 بررسی وضعیت سشن اسنپ")
        snapp_copy_action = menu.addAction("📋 کپی کد سشن پاپا یار (PAPA_)")
        snapp_export_action = menu.addAction("📦 ذخیره فایل سشن این پروفایل (JSON)…")
        menu.addSeparator()
        digikala_action = menu.addAction("🛒 بررسی وضعیت سشن دیجی‌کالا")
        digikala_copy_action = menu.addAction("📋 کپی توکن سشن دیجی‌کالا (JWT)")
        digikala_export_action = menu.addAction("📦 ذخیره فایل سشن دیجی‌کالا (JSON)…")
        menu.addSeparator()
        edit_action = menu.addAction("Edit profile")
        delete_action = menu.addAction("Delete profile")
        chosen = menu.exec(self._list.mapToGlobal(pos))
        if chosen == snapp_action:
            self.check_single_snapp_session(pid)
        elif chosen == snapp_copy_action:
            self.copy_single_snapp_session(pid)
        elif chosen == snapp_export_action:
            self.export_single_snapp_session(pid)
        elif chosen == digikala_action:
            self.check_single_digikala_session(pid)
        elif chosen == digikala_copy_action:
            self.copy_single_digikala_session(pid)
        elif chosen == digikala_export_action:
            self.export_single_digikala_session(pid)
        elif chosen == edit_action:
            self.edit_requested.emit(pid)
        elif chosen == delete_action:
            self.delete_requested.emit(pid)

    def copy_single_snapp_session(self, profile_id: str) -> None:
        """Copy Papa Yar PAPA_ encoded session code for a profile directly to clipboard."""
        profile = self._config_store.profiles.get(profile_id)
        if not profile:
            return
        browser = self._config_store.browsers.get(profile.browser_key)
        res = extract_full_snapp_session(profile, browser)
        if not res.get('success'):
            QMessageBox.warning(
                self,
                "عدم وجود سشن",
                f"سشن یا توکن اسنپ در پروفایل «{profile.display_name}» یافت نشد.\n\n"
                f"علت: {res.get('error') or 'هیچ داده‌ای یافت نشد.'}"
            )
            return

        code = res.get('session_code')
        if code:
            QApplication.clipboard().setText(code)
            status = "فعال 🟢" if res.get('is_active') else "منقضی 🔴"
            phone = res.get('phone', 'N/A')
            msg = f"📋 کد سشن اسنپ ({phone} | {status}) در کلیپ‌بورد کپی شد!"
            self.status_message_requested.emit(msg)
            QMessageBox.information(
                self,
                "کپی کد سشن پاپا یار",
                f"کد سشن پاپا یار برای پروفایل «{profile.display_name}» با موفقیت در کلیپ‌بورد کپی شد!\n\n"
                f"📱 شماره: {phone}\n"
                f"⚡ وضعیت سشن: {status}\n\n"
                f"می‌توانید این کد (PAPA_...) را مستقیماً در کادر سشن تب «حساب» افزونه پاپا یار Paste کنید."
            )

    def export_single_snapp_session(self, profile_id: str) -> None:
        """Export single profile session to a standalone JSON file."""
        profile = self._config_store.profiles.get(profile_id)
        if not profile:
            return
        browser = self._config_store.browsers.get(profile.browser_key)
        res = extract_full_snapp_session(profile, browser)
        if not res.get('success'):
            QMessageBox.warning(
                self,
                "عدم وجود سشن",
                f"سشن یا توکن اسنپ در پروفایل «{profile.display_name}» یافت نشد.\n\n"
                f"علت: {res.get('error') or 'هیچ داده‌ای یافت نشد.'}"
            )
            return

        phone = res.get('phone') or 'unknown'
        safe_phone = re.sub(r'[^0-9a-zA-Z]', '_', str(phone))
        safe_prof = re.sub(r'[^0-9a-zA-Z_]', '_', profile.display_name)[:25]
        bk = profile.browser_key or 'browser'
        default_name = f"session_{safe_phone}_{bk}_{safe_prof}.json"

        last_dir = self._config_store.settings.get("snapp_export_last_dir", "")
        start_path = os.path.join(last_dir, default_name) if last_dir else default_name

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "ذخیره فایل سشن JSON اسنپ",
            start_path,
            "JSON Files (*.json);;All Files (*)"
        )
        if not file_path:
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(res['payload'], f, ensure_ascii=False, indent=2)

            self._config_store.settings["snapp_export_last_dir"] = os.path.dirname(file_path)
            self._config_store.save()

            msg = f"📦 سشن پروفایل {profile.display_name} در فایل {os.path.basename(file_path)} ذخیره شد."
            self.status_message_requested.emit(msg)
            QMessageBox.information(
                self,
                "ذخیره سشن",
                f"فایل سشن با موفقیت ذخیره شد:\n\n{file_path}\n\n"
                f"📱 شماره: {res.get('phone')}\n"
                f"⚡ وضعیت: {'فعال 🟢' if res.get('is_active') else 'منقضی 🔴'}"
            )
        except Exception as e:
            QMessageBox.critical(self, "خطا در ذخیره", f"ذخیره فایل سشن با خطا مواجه شد:\n{e}")

    def export_snapp_sessions_to_folder(self) -> None:
        """Export Snapp sessions as individual JSON files to a selected folder."""
        if self._selected_ids:
            targets = [self._config_store.profiles[pid] for pid in self._selected_ids if pid in self._config_store.profiles]
            desc = f"{len(targets)} پروفایل انتخابی"
        else:
            targets = [self._config_store.profiles[pid] for pid in self._cards.keys() if pid in self._config_store.profiles]
            if not targets:
                targets = list(self._config_store.profiles.values())
            desc = f"تمام پروفایل‌ها ({len(targets)} پروفایل)"

        if not targets:
            QMessageBox.warning(self, "خطا", "هیچ پروفایلی برای استخراج سشن یافت نشد.")
            return

        last_dir = self._config_store.settings.get("snapp_export_last_dir", "")
        folder = QFileDialog.getExistingDirectory(
            self,
            f"انتخاب پوشه برای ذخیره تفکیک‌شده سشن‌ها ({desc})",
            last_dir or ""
        )
        if not folder:
            return

        self._config_store.settings["snapp_export_last_dir"] = folder
        self._config_store.save()

        self._action_bar._snapp_export_btn.setEnabled(False)
        self._action_bar._snapp_all_btn.setEnabled(False)
        self._action_bar._snapp_sel_btn.setEnabled(False)
        self._action_bar._count_label.setText("📦 در حال استخراج سشن‌ها به پوشه…")
        self.status_message_requested.emit(f"📦 در حال استخراج تفکیک‌شده سشن‌ها برای {len(targets)} پروفایل…")

        signals = _SnappExportSignals()
        signals.progress.connect(self._on_snapp_export_progress)
        signals.finished.connect(self._on_snapp_export_finished)

        task = _SnappExportTask(
            profiles=targets,
            browsers=self._config_store.browsers,
            output_dir=folder,
            signals=signals
        )
        QThreadPool.globalInstance().start(task)

    def _on_snapp_export_progress(self, done: int, total: int, prof_name: str) -> None:
        self._action_bar._count_label.setText(f"استخراج سشن: {done}/{total} ({prof_name[:15]})…")

    def _on_snapp_export_finished(self, summary: dict) -> None:
        self._action_bar._snapp_export_btn.setEnabled(True)
        self._action_bar._snapp_all_btn.setEnabled(True)
        self._update_action_bar()

        out_dir = summary.get('output_dir', '')
        exported_count = summary.get('exported_count', 0)
        active_count = summary.get('active_count', 0)
        expired_count = summary.get('expired_count', 0)
        total_scanned = summary.get('total_scanned', 0)

        msg = (
            f"✓ استخراج سشن‌ها پایان یافت: {exported_count} سشن در پوشه ذخیره شد "
            f"({active_count} فعال 🟢 | {expired_count} منقضی 🔴)"
        )
        self.status_message_requested.emit(msg)

        box = QMessageBox(self)
        box.setWindowTitle("استخراج سشن‌های اسنپ")
        box.setIcon(QMessageBox.Icon.Information)
        box.setText(
            f"<b>عملیات استخراج سشن‌ها با موفقیت پایان یافت!</b><br><br>"
            f"📂 <b>پوشه مقصد:</b><br><code>{out_dir}</code><br><br>"
            f"📊 <b>آمار:</b><br>"
            f"• کل پروفایل‌های بررسی‌شده: {total_scanned}<br>"
            f"• سشن‌های استخراج‌شده: <b>{exported_count} فایل JSON تفکیک‌شده</b><br>"
            f"• سشن‌های معتبر و فعال: <b>{active_count}</b> 🟢<br>"
            f"• سشن‌های منقضی‌شده: {expired_count} 🔴<br><br>"
            f"📄 فایل‌های زیر در پوشه ایجاد شدند:<br>"
            f"۱. فایل‌های مستقل <code>session_*.json</code> برای هر اکانت<br>"
            f"۲. فایل لیست و گزارش <code>accounts_list.txt</code><br>"
            f"۳. فایل ورود دسته‌جمعی <code>papayar_bulk_import.json</code> برای افزونه پاپا یار"
        )
        open_folder_btn = box.addButton("باز کردن پوشه در ویندوز", QMessageBox.ButtonRole.ActionRole)
        box.addButton("بستن", QMessageBox.ButtonRole.RejectRole)
        box.exec()

        if box.clickedButton() == open_folder_btn:
            try:
                os.startfile(out_dir)
            except Exception:
                subprocess.Popen(['explorer', out_dir])

    def check_all_digikala_sessions(self) -> None:
        """Check Digikala session for all currently displayed profiles."""
        profiles = [self._config_store.profiles[pid] for pid in self._cards.keys() if pid in self._config_store.profiles]
        if not profiles:
            profiles = list(self._config_store.profiles.values())
        self._start_digikala_check(profiles)

    def check_selected_digikala_sessions(self) -> None:
        """Check Digikala session for selected profiles."""
        targets = [self._config_store.profiles[pid] for pid in self._selected_ids if pid in self._config_store.profiles]
        if targets:
            self._start_digikala_check(targets)

    def check_single_digikala_session(self, profile_id: str) -> None:
        """Check Digikala session for a single profile."""
        profile = self._config_store.profiles.get(profile_id)
        if profile:
            self._start_digikala_check([profile])

    def _start_digikala_check(self, target_profiles: list[BrowserProfile]) -> None:
        if not target_profiles:
            return
        self._action_bar.set_digikala_checking(True, len(target_profiles))
        self.status_message_requested.emit(f"🛒 در حال بررسی سشن دیجی‌کالا برای {len(target_profiles)} پروفایل در پس‌زمینه…")

        signals = _DigikalaCheckSignals()
        signals.profile_checked.connect(self._on_digikala_profile_checked)
        signals.progress.connect(self._on_digikala_progress)
        signals.finished.connect(self._on_digikala_finished)

        task = _DigikalaCheckTask(
            profiles=target_profiles,
            browsers=self._config_store.browsers,
            signals=signals
        )
        QThreadPool.globalInstance().start(task)

    def _on_digikala_profile_checked(self, profile_id: str, res: dict) -> None:
        self._digikala_cache[profile_id] = res
        card = self._cards.get(profile_id)
        if card:
            card.set_digikala_status(
                res.get('status', 'none'),
                res.get('phone', '') or '',
                res.get('name', '') or '',
                res.get('exp', 0) or 0
            )

    def _on_digikala_progress(self, done: int, total: int) -> None:
        self._action_bar._count_label.setText(f"بررسی دیجی‌کالا: {done} از {total}…")

    def _on_digikala_finished(self, counts: dict) -> None:
        self._action_bar.set_digikala_checking(False)
        self._update_action_bar()
        summary_msg = (
            f"✓ نتیجه بررسی دیجی‌کالا: {counts.get('active', 0)} فعال 🟢 | "
            f"{counts.get('expired', 0)} منقضی 🔴 | "
            f"{counts.get('none', 0)} بدون سشن ⚪"
        )
        self.status_message_requested.emit(summary_msg)
        if self._view_mode == "list":
            self.refresh()

    def copy_single_digikala_session(self, profile_id: str) -> None:
        """Copy Digikala JWT token directly to clipboard."""
        profile = self._config_store.profiles.get(profile_id)
        if not profile:
            return
        browser = self._config_store.browsers.get(profile.browser_key)
        res = extract_full_digikala_session(profile, browser)
        if not res.get('success'):
            QMessageBox.warning(
                self,
                "عدم وجود سشن دیجی‌کالا",
                f"سشن یا توکن دیجی‌کالا در پروفایل «{profile.display_name}» یافت نشد.\n\n"
                f"علت: {res.get('error') or 'هیچ دیتایی یافت نشد.'}"
            )
            return

        token = res.get('token')
        if token:
            QApplication.clipboard().setText(token)
            status = "فعال 🟢" if res.get('is_active') else "منقضی 🔴"
            phone = res.get('phone', 'N/A')
            name = res.get('name', '')
            msg = f"📋 توکن دیجی‌کالا ({phone} | {status}) در کلیپ‌بورد کپی شد!"
            self.status_message_requested.emit(msg)
            name_line = f"\n👤 نام کاربر: {name}" if name else ""
            QMessageBox.information(
                self,
                "کپی توکن دیجی‌کالا",
                f"توکن دیجی‌کالا (JWT) برای پروفایل «{profile.display_name}» در کلیپ‌بورد کپی شد!\n\n"
                f"📱 شماره: {phone}{name_line}\n"
                f"⚡ وضعیت سشن: {status}\n\n"
                f"می‌توانید از این توکن برای درخواست‌های API یا اسکریپت‌ها استفاده کنید."
            )

    def export_single_digikala_session(self, profile_id: str) -> None:
        """Export single profile Digikala session to a JSON file."""
        profile = self._config_store.profiles.get(profile_id)
        if not profile:
            return
        browser = self._config_store.browsers.get(profile.browser_key)
        res = extract_full_digikala_session(profile, browser)
        if not res.get('success'):
            QMessageBox.warning(
                self,
                "عدم وجود سشن دیجی‌کالا",
                f"سشن دیجی‌کالا در پروفایل «{profile.display_name}» یافت نشد.\n\n"
                f"علت: {res.get('error') or 'هیچ دیتایی یافت نشد.'}"
            )
            return

        phone = res.get('phone') or 'unknown'
        safe_phone = re.sub(r'[^0-9a-zA-Z]', '_', str(phone))
        safe_prof = re.sub(r'[^0-9a-zA-Z_]', '_', profile.display_name)[:25]
        bk = profile.browser_key or 'browser'
        default_name = f"digikala_session_{safe_phone}_{bk}_{safe_prof}.json"

        last_dir = self._config_store.settings.get("digikala_export_last_dir", "")
        start_path = os.path.join(last_dir, default_name) if last_dir else default_name

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "ذخیره فایل سشن JSON دیجی‌کالا",
            start_path,
            "JSON Files (*.json);;All Files (*)"
        )
        if not file_path:
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(res.get('payload', {}), f, ensure_ascii=False, indent=2)

            self._config_store.settings["digikala_export_last_dir"] = os.path.dirname(file_path)
            self._config_store.save()

            msg = f"📦 سشن دیجی‌کالا پروفایل {profile.display_name} در فایل {os.path.basename(file_path)} ذخیره شد."
            self.status_message_requested.emit(msg)
            QMessageBox.information(
                self,
                "ذخیره سشن دیجی‌کالا",
                f"فایل سشن با موفقیت ذخیره شد:\n\n{file_path}\n\n"
                f"📱 شماره: {res.get('phone')}\n"
                f"⚡ وضعیت: {'فعال 🟢' if res.get('is_active') else 'منقضی 🔴'}"
            )
        except Exception as e:
            QMessageBox.critical(self, "خطا در ذخیره", f"ذخیره فایل سشن با خطا مواجه شد:\n{e}")

    def open_universal_check_dialog(self) -> None:
        """Open the Universal Session & Storage Checker Dialog for custom links."""
        if self._selected_ids:
            targets = [self._config_store.profiles[pid] for pid in self._selected_ids if pid in self._config_store.profiles]
        else:
            targets = [self._config_store.profiles[pid] for pid in self._cards.keys() if pid in self._config_store.profiles]
            if not targets:
                targets = list(self._config_store.profiles.values())

        if not targets:
            QMessageBox.warning(self, "خطا", "هیچ پروفایلی برای بررسی یافت نشد.")
            return

        dlg = UniversalSessionCheckDialog(
            profiles=targets,
            browsers=self._config_store.browsers,
            launch_callback=self.bulk_launch_requested.emit,
            parent=self
        )
        dlg.exec()

    def current_launch_url(self) -> str:
        """Raw text of the grid's 'URL on launch' box (may be empty)."""
        return self._action_bar.launch_url()

    def card_for_profile(self, profile_id: str) -> ProfileCard | None:
        return self._cards.get(profile_id)

    def set_select_mode(self, enabled: bool) -> None:
        """Select mode: clicking anywhere on a card ticks/unticks it instead of
        launching. Pushed to every current (and future) card."""
        self._select_mode = bool(enabled)
        for card in self._cards.values():
            card.set_select_mode(self._select_mode)
        self.select_mode_toggled.emit(self._select_mode)

    def set_delay_seconds(self, value: float) -> None:
        self._action_bar.set_delay_seconds(value)

    def _on_delay_changed(self, value: float) -> None:
        self._config_store.settings["bulk_launch_stagger_seconds"] = float(value)
        self._config_store.save()

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self._content.updateGeometry()

    def _on_card_selection_changed(self, profile_id: str, checked: bool) -> None:
        if checked:
            self._selected_ids.add(profile_id)
        else:
            self._selected_ids.discard(profile_id)
        self._update_action_bar()

    def _on_search_changed(self, text: str) -> None:
        self._search_timer.start()

    def _on_search_debounced(self) -> None:
        self.refresh()

    def select_all(self) -> None:
        if self._view_mode == "list":
            for i in range(self._list.count()):
                self._list.item(i).setCheckState(Qt.CheckState.Checked)
        else:
            for pid, card in self._cards.items():
                card.set_selected(True)
                self._selected_ids.add(pid)
        self._update_action_bar()

    def deselect_all(self) -> None:
        if self._view_mode == "list":
            for i in range(self._list.count()):
                self._list.item(i).setCheckState(Qt.CheckState.Unchecked)
        else:
            for card in self._cards.values():
                card.set_selected(False)
        self._selected_ids.clear()
        self._update_action_bar()

    def focus_search(self) -> None:
        self._action_bar.focus_search()
        self._action_bar.search_edit.selectAll()

    def _on_launch_selected(self) -> None:
        ids = list(self._selected_ids)
        if ids:
            self.bulk_launch_requested.emit(ids)

    def _on_delete_selected(self) -> None:
        ids = list(self._selected_ids)
        if ids:
            self.bulk_delete_requested.emit(ids)

    def _update_action_bar(self) -> None:
        total = len(self._cards) if self._view_mode == "grid" else self._list.count()
        self._action_bar.update_selection(
            selected=len(self._selected_ids), total=total,
        )

    def _build_empty_state(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.setSpacing(8)
        icon = QLabel("🌐")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setStyleSheet("font-size: 48px;")
        lay.addWidget(icon)
        title = QLabel("No browser profiles")
        title.setObjectName("EmptyStateLabel")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)
        hint = QLabel(
            "Click 'Re-scan Browsers', 'Add Profile' or 'Create Profiles…' to get started."
        )
        hint.setObjectName("EmptyStateHint")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(hint)
        return w


# ══════════════════════════════════════════════════════════════════════
#  SECTION 13 — DIALOGS
# ══════════════════════════════════════════════════════════════════════

class ProfileEditDialog(QDialog):
    def __init__(self, config_store: ConfigStore, existing: BrowserProfile | None = None,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self._existing = existing
        self.setWindowTitle("Edit Profile" if existing else "Add Profile")
        self.setMinimumWidth(380)

        form = QFormLayout()

        self._browser_combo = QComboBox()
        for browser in sorted(config_store.browsers.values(), key=lambda b: b.display_name.lower()):
            self._browser_combo.addItem(browser.display_name, userData=browser.key)
        form.addRow("Browser:", self._browser_combo)

        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("e.g. Work - Shopping")
        form.addRow("Display name:", self._name_edit)

        arg_row = QHBoxLayout()
        self._arg_edit = QLineEdit()
        self._arg_edit.setPlaceholderText('e.g. "Profile 2" or "default-release"')
        arg_row.addWidget(self._arg_edit)
        discover_button = QPushButton("Discover...")
        discover_button.setToolTip("Scan this browser's profile registry and pick from the results")
        discover_button.clicked.connect(self._on_discover_clicked)
        arg_row.addWidget(discover_button)
        form.addRow("Profile argument:", arg_row)

        self._color_combo = QComboBox()
        self._color_combo.addItem("None", userData=None)
        for hex_color in Tokens.TAG_PALETTE:
            px = QPixmap(14, 14)
            px.fill(QColor(hex_color))
            icon = QIcon(px)
            self._color_combo.addItem(icon, hex_color, userData=hex_color)
        form.addRow("Color tag:", self._color_combo)

        self._notes_edit = QLineEdit()
        form.addRow("Notes:", self._notes_edit)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

        if existing:
            self._populate_from_existing(existing)

    def _populate_from_existing(self, profile: BrowserProfile) -> None:
        index = self._browser_combo.findData(profile.browser_key)
        if index >= 0:
            self._browser_combo.setCurrentIndex(index)
        self._name_edit.setText(profile.display_name)
        self._arg_edit.setText(profile.profile_arg)
        if profile.color_tag:
            idx = self._color_combo.findData(profile.color_tag)
            if idx >= 0:
                self._color_combo.setCurrentIndex(idx)
        self._notes_edit.setText(profile.notes)

    def _on_discover_clicked(self) -> None:
        browser_key = self._browser_combo.currentData()
        browser = self._config_store.browsers.get(browser_key)
        if browser is None:
            QMessageBox.warning(self, "No browser selected", "Choose a browser first.")
            return
        discovered = discover_profiles_for_browser(browser)
        if not discovered:
            QMessageBox.information(
                self, "No profiles found",
                f"Couldn't auto-detect any profiles for {browser.display_name}. "
                "You can still type a profile argument manually.",
            )
            return
        picked = _DiscoveredProfilePicker.pick(discovered, self)
        if picked is not None:
            self._arg_edit.setText(picked.profile_arg)
            if not self._name_edit.text():
                self._name_edit.setText(picked.display_name)

    def _on_accept(self) -> None:
        browser_key = self._browser_combo.currentData()
        name = self._name_edit.text().strip()
        profile_arg = self._arg_edit.text().strip()
        if not browser_key:
            QMessageBox.warning(self, "Missing browser", "Select a browser.")
            return
        if not name:
            QMessageBox.warning(self, "Missing name", "Enter a display name.")
            return
        if not profile_arg:
            QMessageBox.warning(self, "Missing profile argument",
                                "Enter the profile directory/name, or use Discover.")
            return

        # Duplicate detection: a profile is a duplicate if it targets the same
        # browser and the same profile arg OR the same on-disk profile path, not
        # counting the profile being edited itself.
        dup = self._find_duplicate(browser_key, profile_arg)
        if dup is not None and (self._existing is None or dup.id != self._existing.id):
            reply = QMessageBox.question(
                self, "Duplicate profile",
                f"A similar profile already exists: '{dup.display_name}'."
                "\n\nDo you want to add it again anyway?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        self.result_profile = BrowserProfile(
            id=self._existing.id if self._existing else BrowserProfile.new_id(),
            browser_key=browser_key, display_name=name, profile_arg=profile_arg,
            color_tag=self._color_combo.currentData(),
            notes=self._notes_edit.text().strip(),
        )
        self.accept()

    def _find_duplicate(self, browser_key: str, profile_arg: str) -> BrowserProfile | None:
        for p in self._config_store.profiles.values():
            if p.browser_key != browser_key:
                continue
            if p.profile_arg == profile_arg:
                return p
            if p.profile_path and p.profile_path == (self._arg_edit.text().strip()):
                return p
        return None


class CreateProfilesDialog(QDialog):
    """Create N brand-new browser profiles in one go.

    Chromium-family browsers get real entries in their own profile registry
    (Local State -> profile.info_cache) plus a Profile N folder. Firefox-family
    browsers get a [ProfileN] section appended to profiles.ini plus a
    Profiles/<name> folder. Both paths are strictly append-only: nothing that
    already exists is ever modified or removed.
    """

    def __init__(self, config_store: ConfigStore, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self.setWindowTitle("Create Profiles")
        self.setMinimumWidth(780)

        self.creation_result: ProfileCreationResult | None = None
        self.register_in_launcher = True

        self._busy = False
        self._browser_running = False
        self._close_signals: _CloseBrowserSignals | None = None
        self._create_signals: _CreateProfilesSignals | None = None

        root = QVBoxLayout(self)
        form = QFormLayout()

        self._browser_combo = QComboBox()
        for browser in sorted(config_store.browsers.values(), key=lambda b: b.display_name.lower()):
            self._browser_combo.addItem(browser.display_name, userData=browser.key)
        self._browser_combo.currentIndexChanged.connect(self._on_browser_changed)
        form.addRow("Browser:", self._browser_combo)

        self._engine_label = QLabel("")
        self._engine_label.setStyleSheet(f"color: {Tokens.TEXT_SECONDARY};")
        form.addRow("", self._engine_label)

        self._count_spin = QSpinBox()
        self._count_spin.setRange(CREATE_MIN_COUNT, CREATE_MAX_COUNT)
        self._count_spin.setValue(8)
        self._count_spin.valueChanged.connect(self._refresh_preview)
        form.addRow("How many profiles:", self._count_spin)

        self._pattern_edit = QLineEdit("Profile {n}")
        self._pattern_edit.setPlaceholderText("e.g. Shopping {n}   —   {n} counts up, {browser} is the browser")
        self._pattern_edit.textChanged.connect(self._refresh_preview)
        form.addRow("Name pattern:", self._pattern_edit)

        self._start_spin = QSpinBox()
        self._start_spin.setRange(0, 9999)
        self._start_spin.setValue(1)
        self._start_spin.valueChanged.connect(self._refresh_preview)
        form.addRow("Start number:", self._start_spin)

        self._avatar_combo = QComboBox()
        self._avatar_combo.addItem("Auto — spread over the built-in pictures", userData="auto")
        self._avatar_combo.addItem("Start at a specific picture…", userData="start")
        self._avatar_combo.addItem("Browser default (no picture)", userData="default")
        self._avatar_combo.currentIndexChanged.connect(self._on_avatar_mode_changed)
        form.addRow("Profile picture:", self._avatar_combo)

        self._avatar_index_spin = QSpinBox()
        self._avatar_index_spin.setRange(0, _AVATAR_MAX_INDEX)
        self._avatar_index_spin.setValue(1)
        self._avatar_index_spin.setEnabled(False)
        self._avatar_index_spin.valueChanged.connect(self._refresh_preview)
        form.addRow("", self._avatar_index_spin)

        self._color_check = QCheckBox("Cycle through the color tags")
        self._color_check.toggled.connect(self._refresh_preview)
        form.addRow("", self._color_check)

        self._notes_edit = QLineEdit()
        self._notes_edit.setPlaceholderText("Optional note stored on every profile created")
        form.addRow("Notes:", self._notes_edit)

        self._register_check = QCheckBox("Add the new profiles to this launcher automatically")
        self._register_check.setChecked(True)
        form.addRow("", self._register_check)

        root.addLayout(form)

        self._notice_label = QLabel("")
        self._notice_label.setWordWrap(True)
        root.addWidget(self._notice_label)

        self._running_row = QWidget()
        running_lay = QHBoxLayout(self._running_row)
        running_lay.setContentsMargins(0, 0, 0, 0)
        self._running_label = QLabel("")
        self._running_label.setWordWrap(True)
        running_lay.addWidget(self._running_label, 1)
        self._close_browser_btn = QPushButton("Close Browser && Continue")
        self._close_browser_btn.clicked.connect(self._on_close_browser)
        running_lay.addWidget(self._close_browser_btn)
        self._running_row.setVisible(False)
        root.addWidget(self._running_row)

        root.addWidget(QLabel("These profiles will be created:"))

        self._preview = QTableWidget(0, 4)
        self._preview.setHorizontalHeaderLabels(
            ["Display name", "Profile argument", "Picture", "Folder on disk"]
        )
        self._preview.verticalHeader().setVisible(False)
        self._preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._preview.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self._preview.setMinimumHeight(190)
        preview_header = self._preview.horizontalHeader()
        preview_header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        preview_header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        preview_header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        preview_header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        root.addWidget(self._preview, 1)

        buttons = QDialogButtonBox()
        self._create_btn = buttons.addButton("Create", QDialogButtonBox.ButtonRole.AcceptRole)
        self._cancel_btn = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        self._create_btn.clicked.connect(self._on_create_clicked)
        self._cancel_btn.clicked.connect(self._on_cancel_clicked)
        root.addWidget(buttons)

        self._apply_engine_state()
        self._refresh_running_state()
        self._refresh_preview()

    # -- helpers -----------------------------------------------------------

    def _current_browser(self) -> BrowserDefinition | None:
        key = self._browser_combo.currentData()
        if not key:
            return None
        return self._config_store.browsers.get(key)

    def _avatar_mode(self) -> str:
        return str(self._avatar_combo.currentData() or "auto")

    def _build_plans(self) -> tuple[list[ProfileCreationPlan], list[str]]:
        browser = self._current_browser()
        if browser is None:
            return [], ["Select a browser first."]
        try:
            return build_creation_plan(
                browser,
                count=self._count_spin.value(),
                name_pattern=self._pattern_edit.text(),
                start_number=self._start_spin.value(),
                avatar_mode=self._avatar_mode(),
                avatar_start=self._avatar_index_spin.value(),
                color_cycle=self._color_check.isChecked(),
                notes=self._notes_edit.text().strip(),
            )
        except Exception:
            _CREATE_LOG.exception("Could not work out a creation plan")
            return [], ["Could not work out a plan — see app.log for details."]

    def _apply_engine_state(self) -> None:
        browser = self._current_browser()
        is_gecko = browser is not None and browser.engine == BrowserEngine.GECKO
        if browser is None:
            self._engine_label.setText("")
        elif is_gecko:
            self._engine_label.setText(
                "Firefox-family engine — new profiles are added to profiles.ini"
            )
        else:
            self._engine_label.setText(
                "Chromium engine — new profiles are added to the browser's own profile list"
            )
        self._avatar_combo.setEnabled(not is_gecko)
        self._avatar_combo.setToolTip(
            "Firefox-family browsers have no built-in profile pictures." if is_gecko else ""
        )
        self._avatar_index_spin.setEnabled(not is_gecko and self._avatar_mode() == "start")

    def _set_busy(self, busy: bool, message: str = "") -> None:
        self._busy = busy
        if message:
            self._notice_label.setText(message)
        self._create_btn.setText("Working…" if busy else "Create")
        if busy:
            for widget in (
                self._browser_combo, self._count_spin, self._pattern_edit,
                self._start_spin, self._avatar_combo, self._avatar_index_spin,
                self._color_check, self._notes_edit, self._register_check,
                self._create_btn, self._cancel_btn, self._close_browser_btn,
            ):
                widget.setEnabled(False)
        else:
            for widget in (
                self._browser_combo, self._count_spin, self._pattern_edit,
                self._start_spin, self._avatar_combo, self._color_check,
                self._notes_edit, self._register_check, self._cancel_btn,
            ):
                widget.setEnabled(True)
            self._apply_engine_state()
            self._close_browser_btn.setEnabled(not self._browser_running)

    # -- state refresh -----------------------------------------------------

    def _refresh_running_state(self) -> None:
        browser = self._current_browser()
        running = False
        if browser is not None:
            try:
                running = is_browser_running(browser)
            except Exception:
                _CREATE_LOG.exception("Could not check whether the browser is running")
                running = False
        self._browser_running = running
        if running and browser is not None:
            self._running_label.setText(
                f"⚠ {browser.display_name} is running. Creating profiles writes its "
                "profile registry, so the browser has to be closed first — otherwise "
                "it would overwrite the change the moment it exits."
            )
            self._running_row.setVisible(True)
            self._close_browser_btn.setEnabled(not self._busy)
        else:
            self._running_row.setVisible(False)

    def _refresh_preview(self) -> None:
        if self._busy:
            return
        plans, warnings = self._build_plans()
        browser = self._current_browser()

        self._preview.setRowCount(0)
        for row, plan in enumerate(plans):
            self._preview.insertRow(row)
            picture = "browser default" if plan.avatar_index is None else f"#{plan.avatar_index}"
            location: Optional[Path] = None
            if browser is not None:
                try:
                    location = plan.disk_path(browser)
                except Exception:
                    location = None
            cells = (
                plan.display_name,
                plan.profile_arg,
                picture,
                str(location) if location else plan.folder_name,
            )
            for column, text in enumerate(cells):
                self._preview.setItem(row, column, QTableWidgetItem(text))

        supported, reason = (True, "")
        if browser is None:
            supported, reason = False, "Select a browser first."
        else:
            try:
                supported, reason = profile_creation_support(browser)
            except Exception:
                _CREATE_LOG.exception("Could not check creation support")
                supported, reason = False, "Could not check whether this browser supports it."

        if not supported:
            self._notice_label.setText(f"⚠ {reason}")
        elif warnings:
            self._notice_label.setText("⚠ " + " ".join(warnings))
        elif plans:
            self._notice_label.setText(
                f"{len(plans)} profile(s) will be added to {browser.display_name if browser else ''}."
            )
        else:
            self._notice_label.setText("")

        self._create_btn.setEnabled(
            bool(plans) and supported and not self._busy and not self._browser_running
        )

    # -- slots -------------------------------------------------------------

    def _on_browser_changed(self) -> None:
        self._apply_engine_state()
        self._refresh_running_state()
        self._refresh_preview()

    def _on_avatar_mode_changed(self) -> None:
        self._apply_engine_state()
        self._refresh_preview()

    def _on_close_browser(self) -> None:
        browser = self._current_browser()
        if browser is None or self._busy:
            return
        self._set_busy(True, f"Closing {browser.display_name}…")
        self._close_signals = _CloseBrowserSignals(self)
        self._close_signals.finished.connect(
            self._on_close_finished, Qt.ConnectionType.QueuedConnection
        )
        QThreadPool.globalInstance().start(
            _CloseBrowserTask(self._close_signals, browser)
        )

    def _on_close_finished(self, ok: bool, display_name: str) -> None:
        if self._close_signals is not None:
            self._close_signals.deleteLater()
            self._close_signals = None
        self._set_busy(False)
        self._refresh_running_state()
        self._refresh_preview()
        if not ok:
            QMessageBox.warning(
                self, "Browser still running",
                f"{display_name} did not close. Close it by hand (look for a "
                "'do you want to close all tabs?' prompt) and try again.",
            )

    def _on_create_clicked(self) -> None:
        browser = self._current_browser()
        if browser is None or self._busy:
            return
        plans, warnings = self._build_plans()
        if not plans:
            QMessageBox.warning(
                self, "Nothing to create",
                warnings[0] if warnings else "No profiles could be worked out.",
            )
            return
        try:
            if is_browser_running(browser):
                self._refresh_running_state()
                return
        except Exception:
            _CREATE_LOG.exception("Could not check whether the browser is running")

        if len(plans) >= CREATE_CONFIRM_THRESHOLD:
            reply = QMessageBox.question(
                self, "Create many profiles",
                f"Create {len(plans)} profiles for {browser.display_name}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        self._set_busy(True, f"Creating {len(plans)} profile(s)…")
        self._create_signals = _CreateProfilesSignals(self)
        self._create_signals.finished.connect(
            self._on_create_finished, Qt.ConnectionType.QueuedConnection
        )
        QThreadPool.globalInstance().start(
            _CreateProfilesTask(self._create_signals, browser, plans)
        )

    def _on_create_finished(self, result: ProfileCreationResult) -> None:
        if self._create_signals is not None:
            self._create_signals.deleteLater()
            self._create_signals = None
        self._set_busy(False)
        self.creation_result = result
        self.register_in_launcher = self._register_check.isChecked()

        if not isinstance(result, ProfileCreationResult):
            QMessageBox.warning(self, "Could not create profiles",
                                "The worker returned an unexpected result — see app.log.")
            self._refresh_preview()
            return

        if result.errors and not result.created:
            _CREATE_LOG.error("Profile creation failed: %s", "; ".join(result.errors))
            QMessageBox.warning(self, "Could not create profiles", "\n".join(result.errors))
            self._refresh_preview()
            return
        for message in result.errors:
            _CREATE_LOG.warning("Profile creation: %s", message)

        self.accept()

    def _on_cancel_clicked(self) -> None:
        if self._busy:
            return
        self.reject()

    def reject(self) -> None:
        # Never let a dialog disappear while its worker thread still holds a
        # signal object parented to it.
        if self._busy:
            return
        super().reject()


class _DiscoveredProfilePicker(QDialog):
    def __init__(self, profiles: list[BrowserProfile], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Discovered Profiles")
        self._profiles = profiles
        self._chosen: BrowserProfile | None = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Select a profile to use:"))

        self._list = QListWidget()
        for p in profiles:
            item = QListWidgetItem(f"{p.display_name}  →  {p.profile_arg}")
            self._list.addItem(item)
        layout.addWidget(self._list)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_accept(self) -> None:
        row = self._list.currentRow()
        if 0 <= row < len(self._profiles):
            self._chosen = self._profiles[row]
        self.accept()

    @classmethod
    def pick(cls, profiles: list[BrowserProfile], parent: QWidget | None = None) -> BrowserProfile | None:
        dialog = cls(profiles, parent)
        dialog.exec()
        return dialog._chosen


class BrowserEditDialog(QDialog):
    def __init__(self, existing: BrowserDefinition | None = None,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._existing = existing
        self.setWindowTitle("Edit Browser" if existing else "Add Browser")
        self.setMinimumWidth(420)

        form = QFormLayout()

        self._key_edit = QLineEdit()
        self._key_edit.setPlaceholderText("internal id, e.g. 'chrome_beta'")
        form.addRow("Key:", self._key_edit)

        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("e.g. Chrome Beta")
        form.addRow("Display name:", self._name_edit)

        path_row = QHBoxLayout()
        self._path_edit = QLineEdit()
        path_row.addWidget(self._path_edit)
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self._on_browse_clicked)
        path_row.addWidget(browse_button)
        form.addRow("Executable path:", path_row)

        self._engine_combo = QComboBox()
        self._engine_combo.addItem("Chromium-based (--profile-directory)", BrowserEngine.CHROMIUM)
        self._engine_combo.addItem("Gecko-based (-P / -no-remote)", BrowserEngine.GECKO)
        form.addRow("Engine:", self._engine_combo)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

        if existing:
            self._key_edit.setText(existing.key)
            self._key_edit.setEnabled(False)
            self._name_edit.setText(existing.display_name)
            self._path_edit.setText(str(existing.executable_path))
            idx = self._engine_combo.findData(existing.engine)
            if idx >= 0:
                self._engine_combo.setCurrentIndex(idx)

    def _on_browse_clicked(self) -> None:
        path_str, _ = QFileDialog.getOpenFileName(
            self, "Select browser executable", "", "Executable (*.exe);;All files (*)"
        )
        if path_str:
            self._path_edit.setText(path_str)

    def _on_accept(self) -> None:
        key = self._key_edit.text().strip()
        name = self._name_edit.text().strip()
        path_str = self._path_edit.text().strip()
        if not key or not name or not path_str:
            QMessageBox.warning(self, "Missing fields", "Fill in key, name, and path.")
            return
        # Validate the executable path. A missing/stale path is only a warning
        # (the browser may live on a disconnected drive), so the user can
        # still choose to save it deliberately.
        if not Path(path_str).is_file():
            reply = QMessageBox.warning(
                self, "Executable not found",
                f"The file does not exist:\n{path_str}\n\n"
                "It may be on a drive that is currently disconnected, or the path "
                "may have changed. Do you want to save it anyway?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        self.result_browser = BrowserDefinition(
            key=key, display_name=name, executable_path=Path(path_str),
            engine=self._engine_combo.currentData(), auto_discovered=False,
        )
        self.accept()


class WorkspaceEditDialog(QDialog):
    def __init__(self, config_store: ConfigStore, existing: Workspace | None = None,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self._existing = existing
        self.setWindowTitle("Edit Workspace" if existing else "New Workspace")
        self.setMinimumSize(380, 460)

        layout = QVBoxLayout(self)

        form = QFormLayout()
        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("e.g. Social Media Management")
        form.addRow("Name:", self._name_edit)

        self._stagger_spin = QDoubleSpinBox()
        self._stagger_spin.setRange(0.0, 30.0)
        self._stagger_spin.setSingleStep(0.5)
        self._stagger_spin.setSuffix(" s")
        self._stagger_spin.setValue(0.4)
        self._stagger_spin.setToolTip(
            "Delay between launching each profile in this workspace. "
            "Helps avoid all browsers cold-starting at the exact same instant."
        )
        form.addRow("Launch stagger:", self._stagger_spin)
        layout.addLayout(form)

        layout.addWidget(QLabel("Profiles in this workspace:"))
        self._profile_list = QListWidget()
        self._populate_profile_checklist()
        layout.addWidget(self._profile_list, stretch=1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        if existing:
            self._name_edit.setText(existing.name)
            self._stagger_spin.setValue(existing.launch_stagger_seconds)

    def _populate_profile_checklist(self) -> None:
        selected_ids = set(self._existing.profile_ids) if self._existing else set()
        profiles = sorted(self._config_store.profiles.values(), key=lambda p: p.display_name.lower())
        for profile in profiles:
            browser = self._config_store.browsers.get(profile.browser_key)
            browser_name = browser.display_name if browser else "(unknown)"
            item = QListWidgetItem(f"{profile.display_name}  ·  {browser_name}")
            item.setData(Qt.ItemDataRole.UserRole, profile.id)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked if profile.id in selected_ids else Qt.CheckState.Unchecked
            )
            self._profile_list.addItem(item)

    def _on_accept(self) -> None:
        name = self._name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing name", "Enter a workspace name.")
            return
        selected_ids: list[str] = []
        for i in range(self._profile_list.count()):
            item = self._profile_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_ids.append(item.data(Qt.ItemDataRole.UserRole))
        if not selected_ids:
            QMessageBox.warning(self, "No profiles selected", "Check at least one profile to include.")
            return
        self.result_workspace = Workspace(
            id=self._existing.id if self._existing else Workspace.new_id(),
            name=name, profile_ids=selected_ids,
            launch_stagger_seconds=self._stagger_spin.value(),
        )
        self.accept()


class SettingsDialog(QDialog):
    """Application settings bound directly to ConfigStore.settings so the tray
    menu (which reads/writes the same keys) and this dialog stay in sync."""

    def __init__(self, config_store: ConfigStore, main_window: "MainWindow",
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self._main_window = main_window
        self.setWindowTitle("Settings")
        self.setFixedWidth(440)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("Settings")
        title.setObjectName("DialogTitle")
        layout.addWidget(title)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self._theme_combo = QComboBox()
        for theme_key, theme_label in Tokens.theme_choices():
            self._theme_combo.addItem(theme_label, theme_key)
        saved_theme = str(config_store.settings.get("theme", DEFAULT_THEME))
        theme_index = self._theme_combo.findData(saved_theme)
        self._theme_combo.setCurrentIndex(theme_index if theme_index >= 0 else 0)
        self._theme_combo.setToolTip(
            "Applies instantly so you can preview it.\n"
            "Cancel restores the theme you had before."
        )
        self._theme_combo.currentIndexChanged.connect(self._on_theme_preview)
        form.addRow("Theme", self._theme_combo)

        self._close_to_tray = QCheckBox("Close to tray")
        self._close_to_tray.setChecked(bool(config_store.settings.get("close_to_tray", True)))
        form.addRow("Close to tray", self._close_to_tray)

        self._start_min = QCheckBox("Start minimized to tray")
        self._start_min.setChecked(bool(config_store.settings.get("start_minimized_to_tray", False)))
        form.addRow("Start minimized", self._start_min)

        self._threshold_spin = QSpinBox()
        self._threshold_spin.setRange(0, 10000)
        self._threshold_spin.setValue(int(config_store.settings.get("bulk_launch_confirm_threshold", 10)))
        self._threshold_spin.setToolTip("0 disables the bulk-launch confirmation.")
        form.addRow("Confirm bulk launch above", self._threshold_spin)

        self._stagger_spin = QDoubleSpinBox()
        self._stagger_spin.setRange(0.0, 30.0)
        self._stagger_spin.setSingleStep(0.5)
        self._stagger_spin.setSuffix(" s")
        self._stagger_spin.setValue(float(config_store.settings.get("bulk_launch_stagger_seconds", 2.0)))
        self._stagger_spin.setToolTip(
            "Delay between launching each profile in a bulk launch.\n"
            "Also editable right next to 'Launch Selected' in the action bar."
        )
        form.addRow("Bulk launch stagger", self._stagger_spin)

        layout.addLayout(form)

        div = QFrame()
        div.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(div)

        backup_row = QHBoxLayout()
        export_btn = QPushButton("Export Config…")
        export_btn.clicked.connect(main_window._on_export_config)
        backup_row.addWidget(export_btn)
        import_btn = QPushButton("Import Config…")
        import_btn.clicked.connect(main_window._on_import_config)
        backup_row.addWidget(import_btn)
        layout.addLayout(backup_row)

        open_row = QHBoxLayout()
        open_cfg = QPushButton("Open Config Folder")
        open_cfg.clicked.connect(main_window._on_open_config_folder)
        open_row.addWidget(open_cfg)
        open_log = QPushButton("Open Log Folder")
        open_log.clicked.connect(main_window._on_open_log_folder)
        open_row.addWidget(open_log)
        layout.addLayout(open_row)

        # Danger zone: wipe the isolated Web-Check data folders. Each profile
        # checked via Web Checks keeps its own cookie store here; deleting them
        # logs the user out of every site inside those automation windows.
        clear_row = QHBoxLayout()
        self._clear_webdata_btn = QPushButton("Clear Web-Check Logins…")
        self._clear_webdata_btn.setObjectName("DangerButton")
        self._clear_webdata_btn.setToolTip(
            "Deletes the isolated browser data used by Web Checks\n"
            "(you will be logged out everywhere in those windows)."
        )
        self._clear_webdata_btn.clicked.connect(self._on_clear_webdata)
        clear_row.addWidget(self._clear_webdata_btn)
        clear_row.addStretch(1)
        layout.addLayout(clear_row)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_theme_preview(self) -> None:
        """Repaint the app in the highlighted theme without saving it, so the
        user can flip through the options and judge the real UI."""
        self._main_window.apply_theme(self._theme_combo.currentData(), persist=False)

    def reject(self) -> None:
        """Cancel/Esc/close must undo a live preview, or the window would keep
        a palette the user never saved."""
        self._main_window.apply_theme(
            str(self._config_store.settings.get("theme", DEFAULT_THEME)),
            persist=False,
        )
        super().reject()

    def _on_accept(self) -> None:
        s = self._config_store.settings
        s["theme"] = self._theme_combo.currentData()
        s["close_to_tray"] = self._close_to_tray.isChecked()
        s["start_minimized_to_tray"] = self._start_min.isChecked()
        s["bulk_launch_confirm_threshold"] = self._threshold_spin.value()
        s["bulk_launch_stagger_seconds"] = self._stagger_spin.value()
        self.accept()

    def _on_clear_webdata(self) -> None:
        """Delete all Web-Check automation data folders after confirmation."""
        root = get_app_data_dir() / "webcheck_data"
        try:
            n = len([d for d in root.iterdir() if d.is_dir()]) if root.is_dir() else 0
        except OSError:
            n = 0
        if n == 0:
            QMessageBox.information(
                self, "Nothing to clear",
                "No Web-Check data stored yet.\n"
                "(Data is created the first time a check runs.)"
            )
            return
        reply = QMessageBox.question(
            self, "Clear Web-Check logins",
            f"Delete saved logins for {n} profile data folder(s)?\n\n"
            "You will be logged out of Snapp/Okala inside the automation\n"
            "windows and will have to log in again on the next run.",
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        import shutil as _shutil
        removed = 0
        for d in root.iterdir():
            if not d.is_dir():
                continue
            try:
                _shutil.rmtree(d, ignore_errors=True)
                removed += 1
            except OSError:
                continue
        QMessageBox.information(
            self, "Cleared",
            f"Removed {removed} Web-Check data folder(s)."
        )


# ══════════════════════════════════════════════════════════════════════
#  SECTION 14 — SIDEBAR
# ══════════════════════════════════════════════════════════════════════

SIDEBAR_WIDTH = 220
_WORKSPACE_ID_ROLE = 1001
_BROWSER_KEY_ROLE  = 1002


class Sidebar(QFrame):
    workspace_launch_requested = Signal(str)
    workspace_edit_requested   = Signal(str)
    workspace_delete_requested = Signal(str)
    new_workspace_requested    = Signal()
    browser_filter_changed     = Signal(object)
    tag_filter_changed         = Signal(object)
    browser_edit_requested     = Signal(str)
    browser_remove_requested   = Signal(str)
    web_check_run_requested     = Signal(str)   # task_id

    def __init__(self, config_store: ConfigStore, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self.setObjectName("Sidebar")
        self.setFixedWidth(SIDEBAR_WIDTH)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self._stack = QStackedWidget()
        self._stack.addWidget(self._build_browser_panel())
        self._stack.addWidget(self._build_web_panel())
        root.addWidget(self._stack, stretch=1)

    def set_mode(self, mode: str) -> None:
        self._stack.setCurrentIndex({"browsers": 0}.get(mode, 1))
        self.refresh()

    def rebuild(self) -> None:
        """Recreate both panels from scratch.

        The panels bake theme colours into per-widget inline styles at
        construction time, so after a live theme switch they have to be built
        again — the app-wide stylesheet alone cannot repaint them.
        """
        index = self._stack.currentIndex()
        while self._stack.count():
            panel = self._stack.widget(0)
            self._stack.removeWidget(panel)
            panel.deleteLater()
        self._stack.addWidget(self._build_browser_panel())
        self._stack.addWidget(self._build_web_panel())
        self._stack.setCurrentIndex(index if 0 <= index < 2 else 0)
        self.refresh()

    def _build_browser_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 12, 0, 12)
        layout.setSpacing(4)

        header_row = QHBoxLayout()
        header_row.setContentsMargins(12, 0, 8, 0)
        ws_title = QLabel("Workspaces")
        ws_title.setObjectName("SidebarTitle")
        header_row.addWidget(ws_title)
        header_row.addStretch(1)
        add_btn = QPushButton("+")
        add_btn.setFixedSize(22, 22)
        add_btn.setToolTip("New workspace")
        add_btn.clicked.connect(self.new_workspace_requested.emit)
        header_row.addWidget(add_btn)
        hw = QWidget()
        hw.setLayout(header_row)
        layout.addWidget(hw)

        self._workspace_list = QListWidget()
        self._workspace_list.setObjectName("WorkspaceList")
        self._workspace_list.itemDoubleClicked.connect(self._on_ws_double_clicked)
        self._workspace_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._workspace_list.customContextMenuRequested.connect(self._on_workspace_context_menu)
        layout.addWidget(self._workspace_list, stretch=1)

        launch_ws_btn = QPushButton("Launch Workspace")
        launch_ws_btn.setObjectName("LaunchButton")
        launch_ws_btn.clicked.connect(self._on_launch_ws_clicked)
        lm = QWidget()
        lm.setLayout(self._pad(launch_ws_btn))
        layout.addWidget(lm)

        div = QFrame()
        div.setObjectName("SidebarDivider")
        div.setFrameShape(QFrame.Shape.HLine)
        div.setFixedHeight(1)
        layout.addWidget(div)

        filter_title = QLabel("Filter by Browser")
        filter_title.setObjectName("SidebarTitle")
        filter_title.setContentsMargins(12, 8, 0, 0)
        layout.addWidget(filter_title)

        self._browser_list = QListWidget()
        self._browser_list.setObjectName("BrowserList")
        self._browser_list.setMaximumHeight(180)
        self._browser_list.itemClicked.connect(self._on_browser_filter_clicked)
        self._browser_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._browser_list.customContextMenuRequested.connect(self._on_browser_context_menu)
        layout.addWidget(self._browser_list)

        tag_title = QLabel("Filter by Tag")
        tag_title.setObjectName("SidebarTitle")
        tag_title.setContentsMargins(12, 8, 0, 0)
        layout.addWidget(tag_title)

        tag_container = QWidget()
        self._tag_layout = QHBoxLayout(tag_container)
        self._tag_layout.setContentsMargins(12, 4, 12, 4)
        self._tag_layout.setSpacing(6)
        self._tag_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(tag_container)

        self._tag_buttons: dict[str, QPushButton] = {}

        return panel

    def _pad(self, widget: QWidget) -> QVBoxLayout:
        lay = QVBoxLayout()
        lay.setContentsMargins(8, 4, 8, 4)
        lay.addWidget(widget)
        return lay

    def _build_web_panel(self) -> QWidget:
        """Sidebar for the Web Checks mode: one row per check (Snapp / Okala)
        with a Chromium-only multi-profile picker and a Run button."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 16, 0, 16)
        layout.setSpacing(10)

        icon_lbl = QLabel("\U0001f50d")
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setStyleSheet(f"font-size: 32px; color: {Tokens.AMBER}; padding: 8px;")
        layout.addWidget(icon_lbl)

        title = QLabel("Web Checks")
        title.setObjectName("SidebarTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 11pt; text-transform: none; padding: 0 12px;")
        layout.addWidget(title)

        info = QLabel(
            "Pick one or more Chromium profiles\nper check — each is opened\nand checked one-by-one.\nLog in once per site; the\nsession is reused next runs."
        )
        info.setObjectName("SidebarTitle")
        info.setStyleSheet("padding: 0 16px; line-height: 160%;")
        info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info.setWordWrap(True)
        layout.addWidget(info)

        self._web_task_pickers: dict[str, QPushButton] = {}
        self._web_task_status: dict[str, QLabel] = {}
        self._web_panel_box: QWidget = panel
        self._web_task_card_ids: set[str] = set(self._config_store.web_check_tasks.keys())

        for task in self._config_store.web_check_tasks.values():
            box = QFrame()
            box.setStyleSheet(
                f"QFrame {{ background: {Tokens.BG_CARD}; border: 1px solid {Tokens.BORDER_SUBTLE};"
                f" border-radius: {Tokens.R_CARD}px; }}"
            )
            blay = QVBoxLayout(box)
            blay.setContentsMargins(10, 8, 10, 10)
            blay.setSpacing(6)

            name_lbl = QLabel(task.display_name)
            name_lbl.setStyleSheet(
                f"color: {Tokens.TEXT_PRIMARY}; font-weight: 700; border: none;"
            )
            blay.addWidget(name_lbl)

            picker = QPushButton("No profiles selected")
            picker.setObjectName("WebProfileCombo")
            picker.setCursor(Qt.CursorShape.PointingHandCursor)
            picker.setToolTip(
                "Click to select which Chromium profiles to check.\n"
                "Firefox-family profiles can't be used (no CDP support)."
            )
            picker.clicked.connect(
                lambda checked=False, tid=task.id: self._on_pick_profiles(tid)
            )
            blay.addWidget(picker)

            run_btn = QPushButton("Run Check")
            run_btn.setObjectName("AddInstanceButton")
            run_btn.clicked.connect(
                lambda checked=False, tid=task.id: self.web_check_run_requested.emit(tid)
            )
            blay.addWidget(run_btn)

            status_lbl = QLabel("")
            status_lbl.setWordWrap(True)
            status_lbl.setStyleSheet(
                f"color: {Tokens.TEXT_MUTED}; font-size: {Tokens.SIZE_SM}pt; border: none;"
            )
            blay.addWidget(status_lbl)

            self._web_task_pickers[task.id] = picker
            self._web_task_status[task.id] = status_lbl
            layout.addWidget(box)

        layout.addStretch(1)
        return panel

    def _on_pick_profiles(self, task_id: str) -> None:
        """Open a checkbox dialog listing Chromium profiles; save the selection."""
        task = self._config_store.web_check_tasks.get(task_id)
        if task is None:
            return
        chromium_profiles = sorted(
            (p for p in self._config_store.profiles.values()
             if self._config_store.browsers.get(p.browser_key) is not None
             and self._config_store.browsers[p.browser_key].engine == BrowserEngine.CHROMIUM),
            key=lambda p: p.display_name.lower(),
        )
        if not chromium_profiles:
            QMessageBox.warning(
                self, "No Chromium profiles",
                "No Chromium-based profiles found.\n"
                "Firefox-family (Waterfox/Firefox/Floorp/LibreWolf) profiles\n"
                "cannot be automated — only Chrome/Edge/Brave/Yandex/Vivaldi/Opera."
            )
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Select profiles — {task.display_name}")
        dlg.setMinimumWidth(380)
        vlay = QVBoxLayout(dlg)
        hint = QLabel(
            "Only Chromium-based browsers are listed (CDP automation).\n"
            "Tick every profile you want checked, in any number."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {Tokens.TEXT_SECONDARY};")
        vlay.addWidget(hint)

        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.Shape.NoFrame)
        host = QWidget()
        hlay = QVBoxLayout(host)
        hlay.setContentsMargins(2, 2, 2, 2)
        checkboxes: dict[str, QCheckBox] = {}
        for p in chromium_profiles:
            browser = self._config_store.browsers[p.browser_key]
            cb = QCheckBox(f"{p.display_name}   ({browser.display_name})")
            cb.setChecked(p.id in task.profile_ids)
            hlay.addWidget(cb)
            checkboxes[p.id] = cb
        hlay.addStretch(1)
        area.setWidget(host)
        vlay.addWidget(area, stretch=1)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        vlay.addWidget(btns)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        task.profile_ids = [pid for pid, cb in checkboxes.items() if cb.isChecked()]
        self._config_store.save()
        self._refresh_web_panel()

    def _rebuild_web_panel(self) -> None:
        """Replace the web panel with a fresh one (used when the task set changes
        after an import)."""
        old = self._web_panel_box
        idx = self._stack.currentIndex()
        self._stack.removeWidget(old)
        old.deleteLater()
        self._web_task_pickers.clear()
        self._web_task_status.clear()
        self._web_task_card_ids = set(self._config_store.web_check_tasks.keys())
        self._stack.insertWidget(idx, self._build_web_panel())
        self._stack.setCurrentIndex(idx)

    def _refresh_web_panel(self) -> None:
        """Update the picker buttons (selected counts) and the last-run labels."""
        if not hasattr(self, "_web_task_pickers"):
            return
        # Rebuild the whole panel if the set of tasks changed (import etc.).
        current_ids = getattr(self, "_web_task_card_ids", None)
        wanted_ids = set(self._config_store.web_check_tasks.keys())
        if current_ids != wanted_ids:
            self._rebuild_web_panel()
            return
        for task_id, picker in self._web_task_pickers.items():
            task = self._config_store.web_check_tasks.get(task_id)
            if task is None:
                continue
            n = len(task.profile_ids)
            picker.setText(
                f"{n} profile{'s' if n != 1 else ''} selected" if n
                else "No profiles selected"
            )
            lbl = self._web_task_status.get(task_id)
            if lbl is None:
                continue
            runs = [task.result_for(pid) for pid in task.profile_ids]
            done = [r for r in runs if r.get("run_at")]
            if not done:
                lbl.setText("⌛ never run")
            else:
                ok_n = sum(1 for r in done if r.get("session_active") is True)
                latest = max(r["run_at"] for r in done)
                lbl.setText(
                    f"🟢 {ok_n}/{len(done)} logged in · last: "
                    f"{latest[:16].replace('T', ' ')}"
                )

    def refresh(self) -> None:
        if hasattr(self, "_workspace_list"):
            self._workspace_list.clear()
            for ws in sorted(self._config_store.workspaces.values(), key=lambda w: w.name.lower()):
                item = QListWidgetItem(f"{ws.name}  ({len(ws.profile_ids)})")
                item.setData(_WORKSPACE_ID_ROLE, ws.id)
                self._workspace_list.addItem(item)

        if hasattr(self, "_browser_list"):
            self._browser_list.clear()
            all_item = QListWidgetItem("All Browsers")
            all_item.setData(_BROWSER_KEY_ROLE, None)
            self._browser_list.addItem(all_item)
            self._browser_list.setCurrentItem(all_item)
            for browser in sorted(self._config_store.browsers.values(), key=lambda b: b.display_name.lower()):
                item = QListWidgetItem(browser.display_name)
                item.setData(_BROWSER_KEY_ROLE, browser.key)
                self._browser_list.addItem(item)

        self._rebuild_tag_filter()

        self._refresh_web_panel()

    def _rebuild_tag_filter(self) -> None:
        """Show one chip per color tag actually in use, plus an "All Tags" chip.
        Only color tags that appear on at least one profile are shown."""
        if not hasattr(self, "_tag_layout"):
            return
        # Clear existing chips.
        while self._tag_layout.count():
            item = self._tag_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._tag_buttons.clear()

        used_tags = sorted({p.color_tag for p in self._config_store.profiles.values() if p.color_tag})

        def make_chip(color: str | None, label: str, checked: bool):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setChecked(checked)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            if color:
                btn.setStyleSheet(
                    f"QPushButton {{ border: 1px solid {color}; color: {color};"
                    f" background: {Tokens.BG_INPUT}; border-radius: 4px;"
                    f" padding: 3px 10px; font-size: {Tokens.SIZE_SM}pt; }}"
                    f"QPushButton:checked {{ background: {color}; color: {Tokens.ON_TAG}; }}"
                )
            else:
                btn.setStyleSheet(
                    f"QPushButton {{ border: 1px solid {Tokens.BORDER_MID}; color: {Tokens.TEXT_SECONDARY};"
                    f" background: {Tokens.BG_INPUT}; border-radius: 4px;"
                    f" padding: 3px 10px; font-size: {Tokens.SIZE_SM}pt; }}"
                    f"QPushButton:checked {{ background: {Tokens.AMBER}; color: {Tokens.ON_AMBER}; }}"
                )
            btn.clicked.connect(lambda c, col=color: self._on_tag_chip_clicked(c, col))
            self._tag_layout.addWidget(btn)
            self._tag_buttons[color] = btn
            return btn

        make_chip(None, "All Tags", True)
        for color in used_tags:
            make_chip(color, "●", False)

    def _on_tag_chip_clicked(self, checked: bool, color: str | None) -> None:
        if not checked:
            return
        # Enforce a single active chip.
        for col, btn in self._tag_buttons.items():
            if col != color and btn.isChecked():
                btn.setChecked(False)
        self.tag_filter_changed.emit(color)

    def current_tag_filter(self) -> str | None:
        """Return the color tag the user currently has selected, or None for All."""
        for color, btn in self._tag_buttons.items():
            if btn.isChecked():
                return color
        return None

    def selected_workspace_id(self) -> str | None:
        if not hasattr(self, "_workspace_list"):
            return None
        item = self._workspace_list.currentItem()
        return item.data(_WORKSPACE_ID_ROLE) if item else None

    def _on_ws_double_clicked(self, item: QListWidgetItem) -> None:
        self.workspace_edit_requested.emit(item.data(_WORKSPACE_ID_ROLE))

    def _on_launch_ws_clicked(self) -> None:
        wid = self.selected_workspace_id()
        if wid:
            self.workspace_launch_requested.emit(wid)

    def _on_workspace_context_menu(self, pos: QPoint) -> None:
        item = self._workspace_list.itemAt(pos)
        if item is None:
            return
        wid = item.data(_WORKSPACE_ID_ROLE)
        if not wid:
            return
        menu = QMenu(self)
        launch_action = menu.addAction("Launch Workspace")
        edit_action   = menu.addAction("Edit Workspace")
        delete_action = menu.addAction("Delete Workspace")
        chosen = menu.exec(self._workspace_list.mapToGlobal(pos))
        if chosen == launch_action:
            self.workspace_launch_requested.emit(wid)
        elif chosen == edit_action:
            self.workspace_edit_requested.emit(wid)
        elif chosen == delete_action:
            self.workspace_delete_requested.emit(wid)

    def _on_browser_filter_clicked(self, item: QListWidgetItem) -> None:
        self.browser_filter_changed.emit(item.data(_BROWSER_KEY_ROLE))

    def _on_browser_context_menu(self, pos: QPoint) -> None:
        item = self._browser_list.itemAt(pos)
        if item is None:
            return
        key = item.data(_BROWSER_KEY_ROLE)
        if key is None:
            return  # "All Browsers" heading has no browser behind it
        menu = QMenu(self)
        edit_action   = menu.addAction("Edit browser…")
        remove_action = menu.addAction("Remove browser")
        chosen = menu.exec(self._browser_list.mapToGlobal(pos))
        if chosen == edit_action:
            self.browser_edit_requested.emit(key)
        elif chosen == remove_action:
            self.browser_remove_requested.emit(key)


# ══════════════════════════════════════════════════════════════════════
#  SECTION 15 — SYSTEM TRAY
# ══════════════════════════════════════════════════════════════════════

class TrayController:
    def __init__(self, main_window: QWidget, config_store: ConfigStore,
                 on_launch_workspace, callbacks: dict | None = None) -> None:
        self._main_window = main_window
        self._config_store = config_store
        self._on_launch_workspace = on_launch_workspace
        self._callbacks = callbacks or {}

        self.tray_icon = QSystemTrayIcon(build_app_icon(), main_window)
        self.tray_icon.setToolTip("MultiProfile Launcher")
        self.tray_icon.activated.connect(self._on_activated)

        self._menu = QMenu()
        self._rebuild_menu()
        self.tray_icon.setContextMenu(self._menu)

    def _cb(self, name: str):
        return self._callbacks.get(name)

    def show(self) -> None:
        self.tray_icon.show()

    def rebuild_menu(self) -> None:
        self._rebuild_menu()

    def set_icon(self, icon: QIcon) -> None:
        """Swap the tray icon (used after a live theme switch)."""
        self.tray_icon.setIcon(icon)

    def _apply_theme(self, theme_key: str) -> None:
        self._main_window.apply_theme(theme_key, persist=True)
        # Tick marks in the Theme submenu have to follow the new selection.
        self._rebuild_menu()

    def notify(self, title: str, message: str, is_error: bool = False) -> None:
        icon = (
            QSystemTrayIcon.MessageIcon.Critical if is_error
            else QSystemTrayIcon.MessageIcon.Information
        )
        self.tray_icon.showMessage(title, message, icon, 4000)

    def _rebuild_menu(self) -> None:
        self._menu.clear()
        show_action = QAction("Show Launcher", self._menu)
        show_action.triggered.connect(self._restore_window)
        self._menu.addAction(show_action)
        self._menu.addSeparator()

        workspaces = sorted(self._config_store.workspaces.values(), key=lambda w: w.name.lower())
        if workspaces:
            for workspace in workspaces:
                action = QAction(f"Launch: {workspace.name}", self._menu)
                action.triggered.connect(
                    lambda checked=False, wid=workspace.id: self._on_launch_workspace(wid)
                )
                self._menu.addAction(action)
            self._menu.addSeparator()
        else:
            placeholder = QAction("No workspaces yet", self._menu)
            placeholder.setEnabled(False)
            self._menu.addAction(placeholder)
            self._menu.addSeparator()

        self._menu.addSeparator()
        # Keep references to the submenus instead of relying on walking the
        # menu tree later: QAction.menu() hands back a *Python-owned* wrapper,
        # and letting a temporary one go out of scope deletes the underlying
        # C++ QMenu. Holding the menus here means nothing ever has to walk.
        self._settings_menu = self._menu.addMenu("Settings")
        settings_menu = self._settings_menu

        close_to_tray = QAction("Close to tray", settings_menu)
        close_to_tray.setCheckable(True)
        close_to_tray.setChecked(bool(self._config_store.settings.get("close_to_tray", True)))
        close_to_tray.triggered.connect(self._toggle_close_to_tray)
        settings_menu.addAction(close_to_tray)

        start_min = QAction("Start minimized to tray", settings_menu)
        start_min.setCheckable(True)
        start_min.setChecked(bool(self._config_store.settings.get("start_minimized_to_tray", False)))
        start_min.triggered.connect(self._toggle_start_minimized)
        settings_menu.addAction(start_min)

        self._theme_menu = settings_menu.addMenu("Theme")
        theme_menu = self._theme_menu
        self._theme_group = QActionGroup(theme_menu)
        theme_group = self._theme_group
        theme_group.setExclusive(True)      # one theme ticked at a time
        active_theme = Tokens.theme_key()
        for theme_key, theme_label in Tokens.theme_choices():
            action = QAction(theme_label, theme_menu)
            action.setCheckable(True)
            action.setChecked(theme_key == active_theme)
            action.triggered.connect(
                lambda checked=False, key=theme_key: self._apply_theme(key)
            )
            theme_group.addAction(action)
            theme_menu.addAction(action)

        self._menu.addSeparator()
        export_action = QAction("Export Config…", self._menu)
        export_action.triggered.connect(lambda: self._cb("export") if self._cb("export") else None)
        self._menu.addAction(export_action)
        import_action = QAction("Import Config…", self._menu)
        import_action.triggered.connect(lambda: self._cb("import") if self._cb("import") else None)
        self._menu.addAction(import_action)

        open_config = QAction("Open Config Folder", self._menu)
        open_config.triggered.connect(lambda: self._cb("open_config_folder")() if self._cb("open_config_folder") else None)
        self._menu.addAction(open_config)
        open_log = QAction("Open Log Folder", self._menu)
        open_log.triggered.connect(lambda: self._cb("open_log_folder")() if self._cb("open_log_folder") else None)
        self._menu.addAction(open_log)

        self._menu.addSeparator()
        about_action = QAction(f"About {APP_NAME}", self._menu)
        about_action.triggered.connect(lambda: self._cb("about")() if self._cb("about") else None)
        self._menu.addAction(about_action)

        self._menu.addSeparator()
        quit_action = QAction("Quit", self._menu)
        quit_action.triggered.connect(QApplication.instance().quit)
        self._menu.addAction(quit_action)

    def _toggle_close_to_tray(self, checked: bool) -> None:
        self._config_store.settings["close_to_tray"] = checked
        self._config_store.save()

    def _toggle_start_minimized(self, checked: bool) -> None:
        self._config_store.settings["start_minimized_to_tray"] = checked
        self._config_store.save()

    def _on_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self._restore_window()

    def _restore_window(self) -> None:
        self._main_window.showNormal()
        self._main_window.raise_()
        self._main_window.activateWindow()


# ══════════════════════════════════════════════════════════════════════
#  SECTION 16 — MODE TAB BAR
# ══════════════════════════════════════════════════════════════════════

class _ModeTabBar(QWidget):
    mode_changed = Signal(str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("ModeTabBar")
        self.setFixedHeight(46)
        self._mode = "browsers"

        lay = QHBoxLayout(self)
        lay.setContentsMargins(16, 0, 16, 0)
        lay.setSpacing(0)

        self._browser_tab = QPushButton("🌐  Browsers")
        self._browser_tab.setObjectName("ModeTabActive")
        self._browser_tab.setCheckable(False)
        self._browser_tab.clicked.connect(lambda: self._activate("browsers"))
        lay.addWidget(self._browser_tab)

        self._web_tab = QPushButton("🔍  Web Checks")
        self._web_tab.setObjectName("ModeTab")
        self._web_tab.setCheckable(False)
        self._web_tab.clicked.connect(lambda: self._activate("web_checks"))
        lay.addWidget(self._web_tab)
        if not _WEB_AUTOMATION_AVAILABLE:
            self._web_tab.setEnabled(False)
            self._web_tab.setToolTip("Install Playwright to enable: pip install playwright")

        lay.addStretch(1)

    def _activate(self, mode: str) -> None:
        if mode == self._mode:
            return
        self._mode = mode
        if mode == "browsers":
            self._browser_tab.setObjectName("ModeTabActive")
            self._web_tab.setObjectName("ModeTab")
        else:
            self._browser_tab.setObjectName("ModeTab")
            self._web_tab.setObjectName("ModeTabActive")
        for btn in (self._browser_tab, self._web_tab):
            btn.style().unpolish(btn)
            btn.style().polish(btn)
        self.mode_changed.emit(mode)

    @property
    def current_mode(self) -> str:
        return self._mode


# ════════════════════════════════════════════════════════════════════
#  SECTION 16.5 — WEB CHECKS PANEL
# ══════════════════════════════════════════════════════════════════════

class _WebCheckCard(QFrame):
    """One result card per Web Check: a status row for EVERY selected profile
    (session state + summary), found items (grouped by service for Snapp), and
    a Run button. Each profile row can be run individually."""

    run_requested = Signal(str, str)          # (task_id, profile_id or "" = all)

    def __init__(self, task: WebCheckTask, config_store: "ConfigStore",
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._task_id = task.id
        self._config_ref = config_store
        self._running_profile_id: str | None = None
        self.setObjectName("ProfileCard")
        self.setMinimumWidth(360)
        self.setMaximumWidth(560)
        self.setMinimumHeight(150)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(8)

        header = QHBoxLayout()
        header.setSpacing(8)
        self._title_lbl = QLabel(task.display_name)
        self._title_lbl.setObjectName("ProfileCardName")
        header.addWidget(self._title_lbl)
        header.addStretch(1)
        self._run_btn = QPushButton("Run All")
        self._run_btn.setObjectName("LaunchButton")
        self._run_btn.clicked.connect(
            lambda checked=False: self.run_requested.emit(self._task_id, "")
        )
        header.addWidget(self._run_btn)
        lay.addLayout(header)

        # Per-profile status rows + found items live in one scroll area.
        self._details_area = QScrollArea()
        self._details_area.setWidgetResizable(True)
        self._details_area.setFrameShape(QFrame.Shape.NoFrame)
        self._details_area.setMinimumHeight(90)
        self._details_area.setMaximumHeight(320)
        self._details_host = QWidget()
        self._details_lay = QVBoxLayout(self._details_host)
        self._details_lay.setContentsMargins(0, 0, 0, 0)
        self._details_lay.setSpacing(4)
        self._details_area.setWidget(self._details_host)
        lay.addWidget(self._details_area)

        self._time_lbl = QLabel("")
        self._time_lbl.setObjectName("ProfileCardBrowser")
        lay.addWidget(self._time_lbl)

        self.apply_task(task)

    @property
    def task_id(self) -> str:
        return self._task_id

    def apply_task(self, task: WebCheckTask) -> None:
        self._time_lbl.setText(f"{len(task.profile_ids)} profile(s) assigned")
        self._rebuild_details(task)

    def set_running(self, profile_id: str | None, running: bool) -> None:
        self._running_profile_id = profile_id if running else None
        self._run_btn.setEnabled(not running)

    def _rebuild_details(self, task: WebCheckTask) -> None:
        while self._details_lay.count():
            item = self._details_lay.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
            elif item.layout() is not None:
                pass

        if not task.profile_ids:
            empty = QLabel("No profiles selected — use the sidebar picker.")
            empty.setStyleSheet(f"color: {Tokens.TEXT_MUTED};")
            self._details_lay.addWidget(empty)
            self._details_lay.addStretch(1)
            return

        config = self._config_ref
        for pid in task.profile_ids:
            profile = config.profiles.get(pid) if config else None
            res = task.result_for(pid)
            row_widget = QWidget()
            row_lay = QVBoxLayout(row_widget)
            row_lay.setContentsMargins(0, 6, 0, 2)
            row_lay.setSpacing(3)

            head = QHBoxLayout()
            pname = QLabel(profile.display_name if profile else "(deleted profile)")
            pname.setStyleSheet(
                f"color: {Tokens.TEXT_PRIMARY}; font-weight: 700;"
                f" font-size: {Tokens.SIZE_SM}pt;"
            )
            head.addWidget(pname)
            head.addStretch(1)
            run_one = QPushButton("Run")
            run_one.setObjectName("RunCheckButton")
            run_one.clicked.connect(
                lambda checked=False, p=pid: self.run_requested.emit(self._task_id, p)
            )
            head.addWidget(run_one)
            row_lay.addLayout(head)

            if pid == getattr(self, "_running_profile_id", None):
                state = QLabel("⌳ Checking…")
                state.setStyleSheet(f"color: {Tokens.TEXT_SECONDARY};")
            elif res.get("error"):
                state = QLabel(f"⚠️ {res['error']}")
                state.setWordWrap(True)
                state.setStyleSheet(f"color: {Tokens.WARNING};")
            elif res.get("session_active") is True:
                state = QLabel(f"🟢 {res.get('summary') or 'Session active'}")
                state.setWordWrap(True)
                state.setStyleSheet(f"color: {Tokens.SUCCESS};")
            elif res.get("session_active") is False:
                state = QLabel("🔴 Session inactive — login required")
                state.setWordWrap(True)
                state.setStyleSheet(f"color: {Tokens.DANGER};")
            elif res.get("run_at"):
                state = QLabel(f"⚠️ {res.get('summary') or 'No data'}")
                state.setWordWrap(True)
                state.setStyleSheet(f"color: {Tokens.WARNING};")
            else:
                state = QLabel("⌳ never run")
                state.setStyleSheet(f"color: {Tokens.TEXT_MUTED};")
            row_lay.addWidget(state)

            if res.get("run_at"):
                when = QLabel(f"      last run: {res['run_at'][:16].replace('T', ' ')}")
                when.setStyleSheet(
                    f"color: {Tokens.TEXT_MUTED}; font-size: {Tokens.SIZE_XS}pt;"
                )
                row_lay.addWidget(when)

            details = res.get("details") or []
            if details:
                if task.kind == WebCheckKind.SNAPP_VOUCHERS:
                    groups: dict[str, list[dict]] = {}
                    for d in details:
                        groups.setdefault(
                            d.get("service_name") or "عمومی / نامشخص", []
                        ).append(d)
                    for service, items in groups.items():
                        head_lbl = QLabel(f"   ▸ {service}  ({len(items)})")
                        head_lbl.setStyleSheet(
                            f"color: {Tokens.AMBER}; font-weight: 700;"
                            f" font-size: {Tokens.SIZE_SM}pt;"
                        )
                        row_lay.addWidget(head_lbl)
                        for d in items:
                            item_lbl = QLabel("   " + self._item_text(d))
                            item_lbl.setWordWrap(True)
                            row_lay.addWidget(item_lbl)
                else:
                    for d in details:
                        item_lbl = QLabel("   " + self._item_text(d))
                        item_lbl.setWordWrap(True)
                        row_lay.addWidget(item_lbl)
            self._details_lay.addWidget(row_widget)
        self._details_lay.addStretch(1)

    @staticmethod
    def _item_text(d: dict) -> str:
        bits = [d.get("title") or "(بدون عنوان)"]
        if d.get("value"):
            bits.append(d["value"])
        if d.get("expiry"):
            bits.append(f"exp: {d['expiry']}")
        if d.get("status"):
            bits.append(str(d["status"]))
        return "  ·  ".join(bits)


class WebChecksPanel(QWidget):
    """Main view for the Web Checks mode: a Run All bar plus one result card
    per configured check. Each card shows a per-profile status row and can run
    all assigned profiles (sequentially) or a single one."""

    run_task_requested = Signal(str, str)   # (task_id, profile_id or "")
    run_all_requested = Signal()
    stop_requested = Signal()
    sort_changed = Signal(str)

    def __init__(self, config_store: ConfigStore, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._config_store = config_store
        self._cards: dict[str, _WebCheckCard] = {}
        self._running_keys: set[tuple[str, str]] = set()

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        bar = QWidget()
        bar.setObjectName("GridActionBar")
        blay = QHBoxLayout(bar)
        blay.setContentsMargins(16, 10, 16, 10)
        blay.setSpacing(10)

        self._run_all_btn = QPushButton("▶  Run All Checks")
        self._run_all_btn.setObjectName("BulkLaunchButton")
        self._run_all_btn.clicked.connect(self.run_all_requested.emit)
        blay.addWidget(self._run_all_btn)

        self._stop_btn = QPushButton("■  Stop")
        self._stop_btn.setObjectName("DangerButton")
        self._stop_btn.setEnabled(False)
        self._stop_btn.clicked.connect(self.stop_requested.emit)
        self._stop_btn.setToolTip("Cancel all queued checks (the currently opening browser stays open)")
        blay.addWidget(self._stop_btn)

        blay.addStretch(1)
        lbl = QLabel("Sort by:")
        lbl.setObjectName("GridActionBarCount")
        blay.addWidget(lbl)
        self._sort_combo = QComboBox()
        self._sort_combo.addItem("Session / discounts", "session")
        self._sort_combo.addItem("Has discount", "has_discount")
        saved_sort = str(config_store.settings.get("web_check_sort", "session"))
        idx = self._sort_combo.findData(saved_sort)
        if idx >= 0:
            self._sort_combo.setCurrentIndex(idx)
        self._sort_combo.currentIndexChanged.connect(
            lambda _i: self.sort_changed.emit(self._sort_combo.currentData())
        )
        blay.addWidget(self._sort_combo)
        root.addWidget(bar)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        host = QWidget()
        self._grid_lay = FlowLayout(host)
        self._scroll.setWidget(host)
        root.addWidget(self._scroll, stretch=1)

    # -- public API -------------------------------------------------------

    def refresh(self) -> None:
        """Sync cards with ConfigStore.web_check_tasks (create/remove as needed)."""
        for task_id, card in list(self._cards.items()):
            if task_id not in self._config_store.web_check_tasks:
                self._grid_lay.removeWidget(card)
                card.deleteLater()
                del self._cards[task_id]
        for task in self._config_store.web_check_tasks.values():
            card = self._cards.get(task.id)
            if card is None:
                card = _WebCheckCard(task, self._config_store)
                card.run_requested.connect(self.run_task_requested.emit)
                self._grid_lay.addWidget(card)
                self._cards[task.id] = card
            else:
                card.apply_task(task)
        self.sort_cards(self._sort_combo.currentData())

    def set_running(self, task_id: str, profile_id: str | None, running: bool) -> None:
        key = (task_id, profile_id or "")
        if running:
            self._running_keys.add(key)
        else:
            self._running_keys.discard(key)
        card = self._cards.get(task_id)
        if card is not None:
            card.set_running(profile_id, running)
            task = self._config_store.web_check_tasks.get(task_id)
            if task is not None:
                card.apply_task(task)
        self._run_all_btn.setEnabled(not self._running_keys)
        self._stop_btn.setEnabled(bool(self._running_keys))

    def sort_cards(self, mode: str) -> None:
        tasks = list(self._config_store.web_check_tasks.values())

        def active_count(t: WebCheckTask) -> int:
            return sum(
                1 for r in t.results.values() if r.get("session_active") is True
            )

        def discount_count(t: WebCheckTask) -> int:
            return sum(
                len(r.get("details") or [])
                for r in t.results.values()
                if r.get("session_active") is True
            )

        if mode == "has_discount":
            tasks.sort(key=lambda t: (discount_count(t) == 0, -discount_count(t),
                                      t.display_name.lower()))
        else:  # "session"
            # Cards with any active session first, then by active count, then
            # by number of found discounts.
            tasks.sort(key=lambda t: (active_count(t) == 0, -active_count(t),
                                      -discount_count(t), t.display_name.lower()))

        # FlowLayout keeps insertion order — re-add in the desired order.
        for t in tasks:
            card = self._cards.get(t.id)
            if card is not None:
                self._grid_lay.removeWidget(card)
                self._grid_lay.addWidget(card)


# ══════════════════════════════════════════════════════════════════════
#  SECTION 17 — MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MultiProfile Launcher")
        self.resize(1100, 700)

        self._config_store = ConfigStore()
        self._config_store.load()
        # Resolve the saved theme *before* any widget is built so inline token
        # colours are baked with the right palette from the very first paint.
        Tokens.apply(str(self._config_store.settings.get("theme", DEFAULT_THEME)))
        self.setStyleSheet(build_stylesheet())
        self.setWindowIcon(build_app_icon())

        self._ensure_default_web_checks()

        self._launch_engine = LaunchEngine(parent=self)
        self._launch_engine.launch_completed.connect(self._on_launch_result)

        self._workspace_runner = WorkspaceRunner(
            self._config_store, self._launch_engine, self
        )
        self._workspace_runner.profile_launched.connect(self._on_workspace_profile_launched)
        self._workspace_runner.finished.connect(self._on_workspace_finished)
        self._workspace_runner.progress.connect(self._on_bulk_progress)

        self._staggered_launcher = StaggeredLauncher(self)
        self._staggered_launcher.finished.connect(
            lambda: self._set_status("✓ Bulk launch finished")
        )
        self._staggered_launcher.progress.connect(self._on_bulk_progress)

        self._status_clear_timer = QTimer(self)
        self._status_clear_timer.setSingleShot(True)
        self._status_clear_timer.timeout.connect(self._on_status_timer_timeout)

        self._layout_manager_window: QMainWindow | None = None
        self._discovery_in_flight = False
        self._undo_snapshot: dict[str, Any] | None = None
        self._status_undo_cb = None
        self._web_check_running = False
        self._web_check_queue: list[tuple[str, str]] = []

        self._build_ui()
        self._install_shortcuts()

        self._tray = TrayController(
            self, self._config_store,
            on_launch_workspace=self._launch_workspace_by_id,
            callbacks={
                "export": self._on_export_config,
                "import": self._on_import_config,
                "open_config_folder": self._on_open_config_folder,
                "open_log_folder": self._on_open_log_folder,
                "about": self._on_about,
            },
        )
        self._tray.show()

        # Quitting must not tear down widgets while pool tasks are still
        # running -- their queued signal delivery would hit deleted objects.
        # Bounded so a wedged task can never block shutdown for good.
        QApplication.instance().aboutToQuit.connect(self._wait_for_background_work)

        if not self._config_store.browsers:
            self._start_auto_discovery(show_summary=False)

    @staticmethod
    def _wait_for_background_work() -> None:
        """Let in-flight thread-pool tasks finish before the widgets die."""
        QThreadPool.globalInstance().waitForDone(3000)

    def apply_theme(self, theme_key: str, persist: bool = True) -> None:
        """Repaint the whole app in *theme_key* at runtime.

        Everything styled through :func:`build_stylesheet` repaints on its own;
        the widgets that bake token colours into per-widget inline styles
        (profile cards, web-check rows, the sidebar, the tray icon) are rebuilt
        or redrawn explicitly.

        ``persist=False`` previews the theme without touching config.json, so
        the Settings dialog can revert on Cancel.
        """
        resolved = Tokens.apply(theme_key)
        self.setStyleSheet(build_stylesheet())
        self.setWindowIcon(build_app_icon())
        self._tray.set_icon(build_app_icon())
        self._sidebar.rebuild()
        self._browser_grid.refresh()
        self._web_panel.refresh()
        if persist:
            self._config_store.settings["theme"] = resolved
            self._config_store.save()

    def _build_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_header_bar())

        self._mode_tab_bar = _ModeTabBar()
        self._mode_tab_bar.mode_changed.connect(self._on_mode_changed)
        root.addWidget(self._mode_tab_bar)

        body = QWidget()
        body_lay = QHBoxLayout(body)
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(0)

        self._sidebar = Sidebar(self._config_store)
        self._sidebar.workspace_launch_requested.connect(self._launch_workspace_by_id)
        self._sidebar.workspace_edit_requested.connect(self._on_edit_workspace)
        self._sidebar.workspace_delete_requested.connect(self._on_delete_workspace)
        self._sidebar.new_workspace_requested.connect(self._on_new_workspace)
        self._sidebar.browser_filter_changed.connect(self._on_browser_filter_changed)
        self._sidebar.tag_filter_changed.connect(self._on_tag_filter_changed)
        self._sidebar.browser_edit_requested.connect(self._on_edit_browser)
        self._sidebar.browser_remove_requested.connect(self._on_remove_browser)
        body_lay.addWidget(self._sidebar)

        self._content_stack = QStackedWidget()

        self._browser_grid = ProfileGridView(self._config_store)
        self._browser_grid.launch_requested.connect(self._launch_profile_by_id)
        self._browser_grid.select_mode_toggled.connect(self._on_select_mode_toggled)
        self._browser_grid.edit_requested.connect(self._on_edit_profile)
        self._browser_grid.delete_requested.connect(self._on_delete_profile)
        self._browser_grid.bulk_launch_requested.connect(self._launch_profiles_bulk)
        self._browser_grid.bulk_delete_requested.connect(self._on_delete_profiles_bulk)
        self._browser_grid.launch_url_changed.connect(self._on_launch_url_changed)
        self._browser_grid.status_message_requested.connect(self._set_status)
        self._launch_engine.launch_url = normalize_launch_url(
            self._browser_grid.current_launch_url()
        )
        self._browser_grid.refresh()
        self._content_stack.addWidget(self._browser_grid)

        # Web Checks mode (only reachable when Playwright is installed).
        self._web_panel = WebChecksPanel(self._config_store)
        self._web_panel.run_task_requested.connect(self._run_web_check)
        self._web_panel.run_all_requested.connect(self._run_all_web_checks)
        self._web_panel.stop_requested.connect(self._stop_web_checks)
        self._web_panel.sort_changed.connect(self._on_web_sort_changed)
        self._web_panel.refresh()
        self._content_stack.addWidget(self._web_panel)

        self._sidebar.web_check_run_requested.connect(self._run_web_check)
        self._sidebar.refresh()

        body_lay.addWidget(self._content_stack, stretch=1)
        root.addWidget(body, stretch=1)
        self.setCentralWidget(central)

    def _build_header_bar(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("HeaderBar")
        bar.setFixedHeight(54)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(16, 8, 16, 8)
        lay.setSpacing(8)

        title = QLabel("MultiProfile Launcher")
        title.setObjectName("AppTitle")
        lay.addWidget(title)
        lay.addStretch(1)

        # Status area with an inline Undo button (shown for a few seconds after
        # a delete so the user can revert it).
        status_host = QWidget()
        status_host.setStyleSheet("background-color: transparent;")
        status_host.setObjectName("StatusArea")
        status_lay = QHBoxLayout(status_host)
        status_lay.setContentsMargins(0, 0, 0, 0)
        status_lay.setSpacing(6)

        self._status_label = QLabel("")
        self._status_label.setObjectName("StatusLabel")
        self._status_label.setStyleSheet("background-color: transparent;")
        status_lay.addWidget(self._status_label)

        self._status_undo_btn = QPushButton("Undo")
        self._status_undo_btn.setObjectName("UndoButton")
        self._status_undo_btn.setToolTip("Undo the last delete")
        self._status_undo_btn.setVisible(False)
        self._status_undo_btn.clicked.connect(self._on_undo_clicked)
        status_lay.addWidget(self._status_undo_btn)

        lay.addWidget(status_host)

        settings_btn = QPushButton("⚙ Settings")
        settings_btn.setToolTip("Open settings")
        settings_btn.clicked.connect(self._on_open_settings)
        lay.addWidget(settings_btn)

        self._rescan_btn = QPushButton("Re-scan Browsers")
        self._rescan_btn.clicked.connect(lambda: self._start_auto_discovery(show_summary=True))
        lay.addWidget(self._rescan_btn)

        if _LAYOUT_MANAGER_AVAILABLE:
            layout_manager_btn = QPushButton("🪟 Window Layout Manager")
            layout_manager_btn.clicked.connect(self._open_layout_manager)
            lay.addWidget(layout_manager_btn)

        add_browser_btn = QPushButton("Add Browser…")
        add_browser_btn.clicked.connect(self._on_add_browser)
        lay.addWidget(add_browser_btn)

        add_profile_btn = QPushButton("Add Profile")
        add_profile_btn.setObjectName("LaunchButton")
        add_profile_btn.clicked.connect(self._on_add_profile)
        lay.addWidget(add_profile_btn)

        create_profiles_btn = QPushButton("➕ Create Profiles…")
        create_profiles_btn.setToolTip(
            "Create several brand-new browser profiles at once"
        )
        create_profiles_btn.clicked.connect(self._on_create_profiles)
        lay.addWidget(create_profiles_btn)

        return bar

    def _open_layout_manager(self) -> None:
        if not _LAYOUT_MANAGER_AVAILABLE:
            return
        if self._layout_manager_window is None:
            self._layout_manager_window = WindowLayoutManagerApp()
            # Global Ctrl+Alt+L hotkey filter, installed once on the shared
            # QApplication instance so it works no matter which window has
            # focus (same behavior as when the tool ran standalone).
            hotkey_filter = GlobalHotkeyEventFilter(
                self._layout_manager_window.apply_equal_layout
            )
            app = QApplication.instance()
            app.installNativeEventFilter(hotkey_filter)
            app._wlm_hotkey_filter_ref = hotkey_filter  # keep a strong ref alive
            self._layout_manager_window._hotkey_filter_ref = hotkey_filter
        self._layout_manager_window.show()
        self._layout_manager_window.raise_()
        self._layout_manager_window.activateWindow()

    def _on_mode_changed(self, mode: str) -> None:
        self._content_stack.setCurrentIndex({"browsers": 0}.get(mode, 1))
        self._sidebar.set_mode(mode)
        current_delay = float(self._config_store.settings.get("bulk_launch_stagger_seconds", 2.0))
        self._browser_grid.set_delay_seconds(current_delay)
        if mode == "web_checks":
            self._web_panel.refresh()
            self._sidebar.refresh()

    # -- Web Checks -------------------------------------------------------

    def _ensure_default_web_checks(self) -> None:
        """Create the two built-in checks (Snapp / Okala) the first time the app
        runs; existing configs keep whatever the user already has."""
        if self._config_store.web_check_tasks:
            return
        defaults = (
            (WebCheckKind.SNAPP_VOUCHERS, "Snapp Vouchers"),
            (WebCheckKind.OKALA_DISCOUNTS, "Okala Discounts"),
        )
        for kind, name in defaults:
            self._config_store.upsert_web_check_task(WebCheckTask(
                id=WebCheckTask.new_id(), kind=kind, display_name=name,
            ))

    def _run_web_check(self, task_id: str, profile_id: str = "") -> None:
        """Queue one check × profile. A FIFO queue keeps runs strictly
        sequential — each run opens its own browser instance with a dedicated
        debug port and data dir, so parallel runs would be heavy and racy.

        profile_id == "" means: run ALL profiles assigned to this task.
        """
        task = self._config_store.web_check_tasks.get(task_id)
        if task is None:
            return
        if not _WEB_AUTOMATION_AVAILABLE:
            QMessageBox.warning(
                self, "Playwright missing",
                "This feature needs Playwright.\n\n"
                "Run:  pip install playwright\n"
                "(No separate browser download is needed — the check opens your\n"
                "own installed Chromium browser with an isolated data folder.)",
            )
            return
        if not task.profile_ids:
            QMessageBox.warning(
                self, "No profiles selected",
                "Pick one or more Chromium profiles for this check first\n"
                "(sidebar → click the profile picker)."
            )
            return
        wanted = task.profile_ids if profile_id == "" else [profile_id]
        if getattr(self, "_web_check_running", False):
            for pid in wanted:
                self._web_check_queue.append((task_id, pid))
            self._set_status(f"Queued: {task.display_name} × {len(wanted)}")
            return
        self._web_check_queue.extend((task_id, pid) for pid in wanted)
        self._web_check_running = False
        self._dequeue_web_check()

    def _run_all_web_checks(self) -> None:
        """Run every profile of every configured check, one-by-one.

        Validates up front so a user with nothing selected gets ONE dialog
        instead of one stacked modal per check (the per-task handler also
        validates, and those messages would pile up).
        """
        tasks = list(self._config_store.web_check_tasks.values())
        runnable = [t for t in tasks if t.profile_ids]
        if not runnable:
            QMessageBox.warning(
                self, "No profiles selected",
                "None of the checks has any profile assigned yet.\n\n"
                "Pick profiles per check first (sidebar → click the profile\n"
                "picker on the check card)."
            )
            return
        skipped = len(tasks) - len(runnable)
        for task in runnable:
            self._run_web_check(task.id)
        if skipped:
            self._set_status(
                f"Queued {len(runnable)} check(s); {skipped} had no profiles selected"
            )

    def _stop_web_checks(self) -> None:
        """Empty the pending queue. The currently-running check (if any) is
        allowed to finish — its browser window stays open either way."""
        n = len(self._web_check_queue)
        self._web_check_queue.clear()
        if n:
            self._set_status(f"■ Stopped — {n} queued check(s) cancelled")
        else:
            self._set_status("■ Stopped — nothing was queued")

    def _dequeue_web_check(self) -> None:
        """Start the next (task, profile) run from the queue if none is active."""
        if getattr(self, "_web_check_running", False):
            return
        if not self._web_check_queue:
            self._set_status("✓ All web checks finished")
            return
        task_id, profile_id = self._web_check_queue.pop(0)
        task = self._config_store.web_check_tasks.get(task_id)
        profile = self._config_store.profiles.get(profile_id)
        if task is None or profile is None:
            QTimer.singleShot(0, self._dequeue_web_check)
            return
        browser = self._config_store.browsers.get(profile.browser_key)
        if browser is None or browser.engine != BrowserEngine.CHROMIUM:
            task.results[profile_id] = {
                "run_at": datetime.datetime.now().isoformat(timespec="seconds"),
                "session_active": None,
                "summary": "",
                "details": [],
                "error": f"'{profile.browser_key}' is not a Chromium browser",
            }
            self._config_store.save()
            # Show the refusal right away instead of waiting for some unrelated
            # refresh to reveal it.
            self._web_panel.refresh()
            self._set_status(
                f"⚠ {profile.display_name}: "
                f"'{profile.browser_key}' is not a Chromium browser"
            )
            QTimer.singleShot(0, self._dequeue_web_check)
            return
        self._web_check_running = True
        self._web_panel.set_running(task_id, profile_id, True)
        remaining = len(self._web_check_queue)
        suffix = f"   ({remaining} left)" if remaining else ""
        self._set_status(f"⌳ Checking {task.display_name} — {profile.display_name} …{suffix}")
        # Parented to `self` (MainWindow, a permanent GUI-thread QObject) so it
        # survives past the worker task's own (cross-thread) auto-deletion —
        # see the comment on LaunchEngine.launch_profile for why that matters.
        signals = _WebCheckSignals(self)
        signals.result_ready.connect(
            self._on_web_check_result, Qt.ConnectionType.QueuedConnection
        )
        signals.result_ready.connect(lambda _r, s=signals: s.deleteLater())
        web_task = _WebCheckTask(self._config_store, task, browser, profile, signals)
        QThreadPool.globalInstance().start(web_task)

    def _on_web_check_result(self, result: "WebCheckResult") -> None:
        self._web_check_running = False
        self._web_panel.set_running(result.task_id, result.profile_id, False)
        task = self._config_store.web_check_tasks.get(result.task_id)
        if task is not None:
            entry: dict = {
                "run_at": datetime.datetime.now().isoformat(timespec="seconds"),
            }
            if result.success:
                entry["session_active"] = result.session_active
                entry["summary"] = result.summary
                entry["details"] = list(result.details)
                entry["error"] = ""
                self._set_status(
                    f"{'🟢' if result.session_active else '🔴'} "
                    f"{result.display_name}: {result.summary}"
                )
            else:
                entry["session_active"] = None
                entry["summary"] = ""
                entry["error"] = result.error_message or "Check failed"
                self._set_status(f"⚠ {result.display_name}: {entry['error']}")
            task.results[result.profile_id] = entry
            self._config_store.save()
        self._web_panel.refresh()
        self._sidebar.refresh()
        self._web_panel.sort_cards(
            str(self._config_store.settings.get("web_check_sort", "session"))
        )
        QTimer.singleShot(800, self._dequeue_web_check)

    def _on_web_sort_changed(self, mode: str) -> None:
        self._config_store.settings["web_check_sort"] = mode
        self._config_store.save()
        self._web_panel.sort_cards(mode)

    def _start_auto_discovery(self, show_summary: bool) -> None:
        """Kick off browser/profile scanning on a worker thread."""
        if self._discovery_in_flight:
            return
        self._discovery_in_flight = True
        if hasattr(self, "_rescan_btn"):
            self._rescan_btn.setEnabled(False)
            self._rescan_btn.setText("Scanning…")
        self._set_status("Scanning browsers…")

        # Deliberately NOT parented to this window: the scan may still be
        # running when the app quits, and emitting into a QObject that died
        # with the window logs a bogus "Signal source has been deleted"
        # traceback. The worker holds its own reference via the connection
        # lambdas below, so it stays valid as long as it is needed.
        signals = _DiscoverySignals()
        signals.result_ready.connect(
            lambda result, s=show_summary: self._on_discovery_done(result, s),
            Qt.ConnectionType.QueuedConnection,
        )
        signals.result_ready.connect(lambda _r, s=signals: s.deleteLater())
        task = _DiscoveryTask(signals)
        QThreadPool.globalInstance().start(task)

    def _on_discovery_done(self, result: "_DiscoveryResult", show_summary: bool) -> None:
        self._discovery_in_flight = False
        if hasattr(self, "_rescan_btn"):
            self._rescan_btn.setEnabled(True)
            self._rescan_btn.setText("Re-scan Browsers")

        new_browser_count = 0
        new_profile_count = 0
        renamed_profile_count = 0
        for browser in result.found_browsers:
            if browser.key not in self._config_store.browsers:
                self._config_store.upsert_browser(browser)
                new_browser_count += 1

            def identity(p: BrowserProfile) -> str:
                # A profile's real identity is where it lives on disk, not
                # whatever name is currently attached to it -- names can be
                # (and are) renamed by the user from inside the browser.
                return p.profile_path or f"arg::{p.profile_arg}"

            existing_by_identity = {
                identity(p): p
                for p in self._config_store.profiles_for_browser(browser.key)
            }
            for profile in result.profiles_by_browser.get(browser.key, []):
                existing = existing_by_identity.get(identity(profile))
                if existing is None:
                    self._config_store.upsert_profile(profile)
                    new_profile_count += 1
                elif existing.display_name != profile.display_name:
                    # Same underlying profile folder, name changed since
                    # last scan -- update in place so workspaces/launch
                    # history tied to this profile's id keep working.
                    existing.display_name = profile.display_name
                    existing.profile_arg = profile.profile_arg
                    self._config_store.upsert_profile(existing)
                    renamed_profile_count += 1
        self._config_store.save()
        self._sidebar.refresh()
        self._browser_grid.refresh()
        if show_summary:
            QMessageBox.information(
                self, "Scan complete",
                f"Found {new_browser_count} new browser(s), "
                f"{new_profile_count} new profile(s), and updated "
                f"{renamed_profile_count} renamed profile(s).\n\n"
                "If a browser wasn't detected, use 'Add Browser…' to "
                "point at it manually.",
            )

    def _on_create_profiles(self) -> None:
        if not self._config_store.browsers:
            QMessageBox.warning(
                self, "No browsers configured",
                "Add a browser first via 'Add Browser…' or 'Re-scan Browsers'.",
            )
            return
        dialog = CreateProfilesDialog(self._config_store, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        result = dialog.creation_result
        if result is None or not result.created:
            return
        browser = self._config_store.browsers.get(result.browser_key)
        if browser is None:
            return

        snapshot = self._config_store.to_payload()

        if dialog.register_in_launcher:
            for plan in result.created:
                if self._existing_profile_for(result.browser_key, plan) is not None:
                    continue
                self._config_store.upsert_profile(BrowserProfile(
                    id=BrowserProfile.new_id(),
                    browser_key=result.browser_key,
                    display_name=plan.display_name,
                    profile_arg=plan.profile_arg,
                    color_tag=plan.color_tag,
                    notes=plan.notes or "Created by 'Create Profiles'",
                    profile_path=plan.profile_path,
                ))
            self._config_store.save()
            self._sidebar.refresh()
            self._browser_grid.refresh()

        if result.backup_paths:
            _CREATE_LOG.info("Profile registry backed up to: %s", "; ".join(result.backup_paths))
        if result.warnings:
            for message in result.warnings:
                _CREATE_LOG.warning("Profile creation: %s", message)

        created_count = len(result.created)
        self._set_status(
            f"✓ Created {created_count} profile(s) for {browser.display_name}",
            on_undo=lambda: self._undo_profile_creation(snapshot, browser, result.created),
        )

    def _existing_profile_for(self, browser_key: str,
                              plan: ProfileCreationPlan) -> BrowserProfile | None:
        for profile in self._config_store.profiles_for_browser(browser_key):
            if profile.profile_arg == plan.profile_arg:
                return profile
            if plan.profile_path and profile.profile_path == plan.profile_path:
                return profile
        return None

    def _undo_profile_creation(self, snapshot: dict[str, Any], browser: BrowserDefinition,
                               plans: list[ProfileCreationPlan]) -> None:
        try:
            notes = remove_created_profiles(browser, plans)
        except Exception:
            _CREATE_LOG.exception("Undoing profile creation failed")
            notes = ["Something went wrong while undoing — see app.log."]
        self._config_store.load_payload(snapshot)
        self._config_store.save()
        self._sidebar.refresh()
        self._browser_grid.refresh()
        if notes:
            QMessageBox.warning(self, "Profile creation partly undone", "\n".join(notes))
            self._set_status("⚠ Profile creation could not be fully undone")
        else:
            self._set_status("✓ Profile creation undone")

    def _on_add_profile(self) -> None:
        if not self._config_store.browsers:
            QMessageBox.warning(self, "No browsers configured",
                                "Add a browser first via 'Add Browser…' or 'Re-scan Browsers'.")
            return
        dialog = ProfileEditDialog(self._config_store, parent=self)
        if dialog.exec() == ProfileEditDialog.DialogCode.Accepted:
            self._config_store.upsert_profile(dialog.result_profile)
            self._config_store.save()
            self._browser_grid.refresh()

    def _on_edit_profile(self, profile_id: str) -> None:
        profile = self._config_store.profiles.get(profile_id)
        if profile is None:
            return
        dialog = ProfileEditDialog(self._config_store, existing=profile, parent=self)
        if dialog.exec() == ProfileEditDialog.DialogCode.Accepted:
            self._config_store.upsert_profile(dialog.result_profile)
            self._config_store.save()
            self._browser_grid.refresh()

    def _on_delete_profile(self, profile_id: str) -> None:
        profile = self._config_store.profiles.get(profile_id)
        if profile is None:
            return
        reply = QMessageBox.question(
            self, "Delete profile",
            f"Delete '{profile.display_name}'? This also removes it from workspaces.",
        )
        if reply == QMessageBox.StandardButton.Yes:
            snapshot = self._config_store.to_payload()
            self._config_store.remove_profile(profile_id)
            self._config_store.save()
            self._browser_grid.refresh()
            self._sidebar.refresh()
            self._set_status(
                f"✓ Deleted '{profile.display_name}'",
                on_undo=lambda: self._undo_delete(snapshot),
            )

    def _on_add_browser(self) -> None:
        dialog = BrowserEditDialog(parent=self)
        if dialog.exec() == BrowserEditDialog.DialogCode.Accepted:
            browser = dialog.result_browser
            if browser.key in self._config_store.browsers:
                QMessageBox.warning(self, "Duplicate", f"Browser '{browser.key}' already exists.")
                return
            self._config_store.upsert_browser(browser)
            self._config_store.save()
            self._sidebar.refresh()

    def _on_edit_browser(self, browser_key: str) -> None:
        browser = self._config_store.browsers.get(browser_key)
        if browser is None:
            return
        dialog = BrowserEditDialog(existing=browser, parent=self)
        if dialog.exec() == BrowserEditDialog.DialogCode.Accepted:
            self._config_store.upsert_browser(dialog.result_browser)
            self._config_store.save()
            self._sidebar.refresh()
            self._browser_grid.refresh()

    def _on_remove_browser(self, browser_key: str) -> None:
        browser = self._config_store.browsers.get(browser_key)
        if browser is None:
            return
        n = len(self._config_store.profiles_for_browser(browser_key))
        reply = QMessageBox.question(
            self, "Remove browser",
            f"Remove '{browser.display_name}' and its {n} profile(s)? "
            "Profiles are also removed from workspaces.",
        )
        if reply == QMessageBox.StandardButton.Yes:
            snapshot = self._config_store.to_payload()
            self._config_store.remove_browser(browser_key)
            self._config_store.save()
            self._sidebar.refresh()
            self._browser_grid.refresh()
            self._set_status(
                f"✓ Removed browser '{browser.display_name}'",
                on_undo=lambda: self._undo_delete(snapshot),
            )

    def _on_new_workspace(self) -> None:
        if not self._config_store.profiles:
            QMessageBox.warning(self, "No profiles yet",
                                "Add at least one browser profile before creating a workspace.")
            return
        dialog = WorkspaceEditDialog(self._config_store, parent=self)
        if dialog.exec() == WorkspaceEditDialog.DialogCode.Accepted:
            self._config_store.upsert_workspace(dialog.result_workspace)
            self._config_store.save()
            self._sidebar.refresh()
            self._tray.rebuild_menu()

    def _on_edit_workspace(self, workspace_id: str) -> None:
        ws = self._config_store.workspaces.get(workspace_id)
        if ws is None:
            return
        dialog = WorkspaceEditDialog(self._config_store, existing=ws, parent=self)
        if dialog.exec() == WorkspaceEditDialog.DialogCode.Accepted:
            self._config_store.upsert_workspace(dialog.result_workspace)
            self._config_store.save()
            self._sidebar.refresh()
            self._tray.rebuild_menu()

    def _on_delete_workspace(self, workspace_id: str) -> None:
        ws = self._config_store.workspaces.get(workspace_id)
        if ws is None:
            return
        reply = QMessageBox.question(
            self, "Delete workspace",
            f"Delete workspace '{ws.name}'? Its profiles are not deleted.",
        )
        if reply == QMessageBox.StandardButton.Yes:
            snapshot = self._config_store.to_payload()
            self._config_store.remove_workspace(workspace_id)
            self._config_store.save()
            self._sidebar.refresh()
            self._tray.rebuild_menu()
            self._set_status(
                f"✓ Deleted workspace '{ws.name}'",
                on_undo=lambda: self._undo_delete(snapshot),
            )

    def _launch_workspace_by_id(self, workspace_id: str) -> None:
        ws: Workspace | None = self._config_store.workspaces.get(workspace_id)
        if ws is None:
            return
        if not self._confirm_bulk_launch(len(ws.profile_ids), "profiles"):
            return
        self._set_status(f"Launching workspace '{ws.name}'…")
        self._workspace_runner.run(ws)

    def _on_launch_url_changed(self, text: str) -> None:
        """'URL on launch' box changed: sync the engine (used by every
        single/bulk/workspace launch) and persist the raw text."""
        self._launch_engine.launch_url = normalize_launch_url(text)
        self._config_store.settings["launch_open_url"] = text.strip()
        self._config_store.save()

    def _launch_profile_by_id(self, profile_id: str) -> None:
        profile = self._config_store.profiles.get(profile_id)
        if profile is None:
            return
        browser = self._config_store.browsers.get(profile.browser_key)
        if browser is None:
            self._set_status(f"✗ Browser '{profile.browser_key}' not configured")
            return
        self._launch_engine.launch_profile(browser, profile)

    def _confirm_bulk_launch(self, count: int, kind: str) -> bool:
        """Ask for confirmation before a bulk launch above the configured
        threshold. Smaller batches launch immediately."""
        threshold = int(self._config_store.settings.get("bulk_launch_confirm_threshold", 10))
        if count <= 0:
            return False
        if count <= threshold:
            return True
        reply = QMessageBox.question(
            self, "Confirm bulk launch",
            f"You're about to launch {count} {kind}. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return reply == QMessageBox.StandardButton.Yes

    def _on_bulk_progress(self, done: int, total: int) -> None:
        self._set_status(f"Launching {done}/{total}…")

    def _on_select_mode_toggled(self, on: bool) -> None:
        if on:
            self._set_status(
                "☑ Select mode ON — click cards to tick/untick them, then 'Launch Selected'"
            )
        else:
            self._set_status("Select mode OFF — clicking a card launches it again")

    def _launch_profiles_bulk(self, profile_ids: list[str]) -> None:
        pairs = []
        for pid in profile_ids:
            profile = self._config_store.profiles.get(pid)
            browser = self._config_store.browsers.get(profile.browser_key) if profile else None
            if profile and browser:
                pairs.append((browser, profile))
        if not pairs:
            return
        if not self._confirm_bulk_launch(len(pairs), "profiles"):
            return
        self._set_status(f"Launching {len(pairs)} profile(s)…")
        stagger = float(self._config_store.settings.get("bulk_launch_stagger_seconds", 2.0))
        self._staggered_launcher.run(
            [lambda b=brow, p=prof: self._launch_engine.launch_profile(b, p) for brow, prof in pairs],
            stagger,
        )

    def _on_launch_result(self, result: LaunchResult) -> None:
        card = self._browser_grid.card_for_profile(result.item_id)
        if card is not None:
            try:
                card.show_launch_feedback(result.success, result.error_message)
            except RuntimeError:
                # The card was rebuilt away (C++ object deleted) while the
                # launch was in flight — nothing to show feedback on.
                pass
        if result.success:
            self._set_status(f"✓ Launched {result.display_name}")
        else:
            self._set_status(f"✗ Failed: {result.display_name}")
            self._tray.notify("Launch failed", f"{result.display_name}: {result.error_message}",
                              is_error=True)

    def _on_workspace_profile_launched(self, result: LaunchResult) -> None:
        # Per-profile feedback emitted by WorkspaceRunner for profiles it could
        # not launch (missing browser config, deleted profile, etc.).
        card = self._browser_grid.card_for_profile(result.item_id)
        if card is not None:
            try:
                card.show_launch_feedback(result.success, result.error_message)
            except RuntimeError:
                # The card was rebuilt away (C++ object deleted) meanwhile.
                pass
        if result.success:
            self._tray.notify("Workspace launch", f"Launched {result.display_name}")
        else:
            self._set_status(f"✗ {result.display_name}: {result.error_message}")
            self._tray.notify("Workspace launch failed",
                              f"{result.display_name}: {result.error_message}", is_error=True)

    def _on_workspace_finished(self, workspace_id: str) -> None:
        ws = self._config_store.workspaces.get(workspace_id)
        self._set_status(f"✓ Workspace '{ws.name if ws else 'workspace'}' finished")

    def _on_delete_profiles_bulk(self, profile_ids: list[str]) -> None:
        names = [
            self._config_store.profiles[pid].display_name
            for pid in profile_ids
            if pid in self._config_store.profiles
        ]
        reply = QMessageBox.question(
            self, "Delete profiles",
            f"Delete {len(names)} profile(s)? This also removes them from workspaces.\n\n"
            + "\n".join(f"• {n}" for n in names[:12])
            + ("\n…" if len(names) > 12 else ""),
        )
        if reply == QMessageBox.StandardButton.Yes:
            snapshot = self._config_store.to_payload()
            for pid in profile_ids:
                self._config_store.remove_profile(pid)
            self._config_store.save()
            self._browser_grid.refresh()
            self._sidebar.refresh()
            self._set_status(
                f"✓ Deleted {len(names)} profile(s)",
                on_undo=lambda: self._undo_delete(snapshot),
            )

    def _on_browser_filter_changed(self, browser_key: object) -> None:
        self._browser_grid.refresh(filter_browser_key=browser_key)  # type: ignore[arg-type]

    def _on_tag_filter_changed(self, tag: object) -> None:
        self._browser_grid.refresh(filter_tag=tag)  # type: ignore[arg-type]

    def _install_shortcuts(self) -> None:
        # Ctrl+F focuses the active search box. The remaining shortcuts live in
        # keyPressEvent so they don't steal Ctrl+A / Delete / Enter from an
        # editable field (e.g. the search box itself).
        self._focus_search_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        self._focus_search_shortcut.activated.connect(self._focus_active_search)
        # Ctrl+1/2 switch the main tabs.
        for seq, mode in (("Ctrl+1", "browsers"), ("Ctrl+2", "web_checks")):
            sc = QShortcut(QKeySequence(seq), self)
            sc.activated.connect(lambda m=mode: self._mode_tab_bar._activate(m))

    def _focus_active_search(self) -> None:
        if self._content_stack.currentIndex() == 0:
            self._browser_grid.focus_search()

    def _is_text_edit_focused(self) -> bool:
        fw = QApplication.instance().focusWidget()
        return isinstance(fw, (QLineEdit,))

    def _select_all_active(self) -> None:
        self._browser_grid.select_all()

    def _deselect_all_active(self) -> None:
        self._browser_grid.deselect_all()

    def _launch_active_selected(self) -> None:
        grid = self._browser_grid
        if grid._selected_ids:
            self._launch_profiles_bulk(list(grid._selected_ids))

    def _delete_active_selected(self) -> None:
        grid = self._browser_grid
        if grid._selected_ids:
            self._on_delete_profiles_bulk(list(grid._selected_ids))

    def keyPressEvent(self, event) -> None:  # noqa: N802
        # Let focused editable fields (search boxes, dialogs) handle their own
        # Ctrl+A / Delete / Enter before we treat them as grid shortcuts.
        if self._is_text_edit_focused():
            super().keyPressEvent(event)
            return
        key = event.key()
        mods = event.modifiers()
        ctrl = bool(mods & Qt.KeyboardModifier.ControlModifier)
        shift = bool(mods & Qt.KeyboardModifier.ShiftModifier)
        if ctrl and not shift and key == Qt.Key.Key_A:
            self._select_all_active()
            event.accept()
            return
        if ctrl and shift and key == Qt.Key.Key_A:
            self._deselect_all_active()
            event.accept()
            return
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self._launch_active_selected()
            event.accept()
            return
        if key == Qt.Key.Key_Delete:
            self._delete_active_selected()
            event.accept()
            return
        super().keyPressEvent(event)

    def _set_status(self, text: str, on_undo=None) -> None:
        self._status_label.setText(text)
        self._status_clear_timer.start(5000)
        if on_undo is not None:
            self._status_undo_cb = on_undo
            self._status_undo_btn.setVisible(True)
        else:
            self._hide_undo()

    def _hide_undo(self) -> None:
        self._status_undo_cb = None
        self._status_undo_btn.setVisible(False)

    def _on_undo_clicked(self) -> None:
        callback = self._status_undo_cb
        self._hide_undo()
        if callback is not None:
            callback()

    def _on_status_timer_timeout(self) -> None:
        self._status_label.setText("")
        self._hide_undo()

    def _undo_delete(self, snapshot: dict[str, Any]) -> None:
        """Restore a full config snapshot captured just before a delete."""
        self._config_store.load_payload(snapshot)
        self._config_store.save()
        self._sidebar.refresh()
        self._browser_grid.refresh()
        if hasattr(self, "_web_panel"):
            self._web_panel.refresh()
        self._tray.rebuild_menu()
        self._set_status("✓ Undone")

    def _on_open_settings(self) -> None:
        dialog = SettingsDialog(self._config_store, self, parent=self)
        if dialog.exec() == SettingsDialog.DialogCode.Accepted:
            self._config_store.save()
            self._tray.rebuild_menu()
            self._browser_grid.refresh()
    
    def _on_export_config(self) -> None:
        path_str, _ = QFileDialog.getSaveFileName(
            self, "Export config", "config.json",
            "JSON config (*.json);;All files (*)",
        )
        if not path_str:
            return
        try:
            self._config_store.export_to(Path(path_str))
        except OSError as exc:
            QMessageBox.warning(self, "Export failed", str(exc))
            return
        self._set_status(f"✓ Config exported to {path_str}")

    def _on_import_config(self) -> None:
        path_str, _ = QFileDialog.getOpenFileName(
            self, "Import config", "", "JSON config (*.json);;All files (*)",
        )
        if not path_str:
            return
        try:
            candidate = json.loads(Path(path_str).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            QMessageBox.warning(self, "Import failed", f"Could not read file: {exc}")
            return
        missing = {"browsers", "profiles"} - set(candidate.keys())
        if missing:
            QMessageBox.warning(
                self, "Import failed",
                "That file isn't a MultiProfile Launcher config (missing keys: "
                + ", ".join(sorted(missing)) + ").",
            )
            return
        reply = QMessageBox.question(
            self, "Import config",
            "Importing will REPLACE your current browsers, profiles, workspaces and instances. "
            "A backup of the current config will be saved first. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            self._config_store._backup_config()
            self._config_store.import_from(Path(path_str))
            self._config_store.save()
        except Exception as exc:
            QMessageBox.warning(self, "Import failed", str(exc))
            return
        self._sidebar.refresh()
        self._browser_grid.refresh()
        self._tray.rebuild_menu()
        self._set_status("✓ Config imported")

    def _on_open_config_folder(self) -> None:
        try:
            os.startfile(get_app_data_dir())  # type: ignore[attr-defined]
        except AttributeError:
            self._set_status("Open config folder only works on Windows")
        except OSError as exc:
            QMessageBox.warning(self, "Open config folder", str(exc))

    def _on_open_log_folder(self) -> None:
        try:
            os.startfile(get_app_data_dir())  # type: ignore[attr-defined]
        except AttributeError:
            self._set_status("Open log folder only works on Windows")
        except OSError as exc:
            QMessageBox.warning(self, "Open log folder", str(exc))

    def _on_about(self) -> None:
        QMessageBox.about(
            self, f"About {APP_NAME}",
            f"<b>{APP_NAME}</b> v{APP_VERSION}<br><br>"
            "Launch many isolated browser profiles with one click, grouped "
            "into workspaces, with search, tags, bulk launch and a built-in "
            "profile creator.<br><br>"
            "Config is stored in <code>%APPDATA%\\MultiProfileLauncher</code>.",
        )

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        tray_ok = QSystemTrayIcon.isSystemTrayAvailable()
        close_to_tray = bool(self._config_store.settings.get("close_to_tray", True))
        if tray_ok and close_to_tray:
            event.ignore()
            self.hide()
            self._tray.notify(
                "MultiProfile Launcher",
                "Still running in the background. Right-click the tray icon to quit.",
            )
        else:
            event.accept()
            QApplication.instance().quit()


# ══════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════


def setup_logging() -> None:
    """Rotating file log in the app data dir. With a --windowed PyInstaller
    build there is no stdout/stderr, so errors are only visible in this file."""
    try:
        log_dir = get_app_data_dir()
        log_dir.mkdir(parents=True, exist_ok=True)
        handler = RotatingFileHandler(
            log_dir / "app.log", maxBytes=2 * 1024 * 1024, backupCount=2,
            encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)-7s %(name)s: %(message)s"
        ))
        root = logging.getLogger()
        root.setLevel(logging.INFO)
        # Avoid duplicate handlers if main() runs more than once (tests).
        if not any(isinstance(h, RotatingFileHandler) for h in root.handlers):
            root.addHandler(handler)
    except OSError as exc:
        logging.warning("Could not set up file logging: %s", exc)


def install_crash_capture() -> None:
    """Route unhandled Python exceptions and Qt log messages into app.log.

    With PySide6 >= 6.5, an unhandled exception inside a slot invoked from C++
    ABORTS the whole process — and in a --windowed app there is no stderr, so
    the crash reason used to be silently lost (the app 'just closes'). These
    hooks make every future crash leave an exact traceback in the rotating
    log file inside the MultiProfileLauncher app-data folder (app.log).
    """
    root = logging.getLogger()

    def _log_exc(prefix: str, exc_type, exc_value, exc_tb) -> None:
        try:
            root.critical(
                "%s%s", prefix,
                "".join(traceback.format_exception(exc_type, exc_value, exc_tb)),
            )
        except Exception:
            pass

    def _sys_excepthook(exc_type, exc_value, exc_tb) -> None:
        _log_exc("Unhandled exception: ", exc_type, exc_value, exc_tb)
        # Keep the default behavior too (prints to stderr where one exists).
        sys.__excepthook__(exc_type, exc_value, exc_tb)

    def _thread_excepthook(args: "threading.ExceptHookArgs") -> None:
        thread_name = getattr(args.thread, "name", "?") if args.thread else "?"
        _log_exc(
            f"Unhandled exception in thread '{thread_name}': ",
            args.exc_type, args.exc_value, args.exc_traceback,
        )

    def _qt_message_handler(mode, context, message) -> None:
        if mode in (QtMsgType.QtFatalMsg, QtMsgType.QtCriticalMsg):
            level = root.critical
        elif mode == QtMsgType.QtWarningMsg:
            level = root.warning
        elif mode == QtMsgType.QtInfoMsg:
            level = root.info
        else:
            level = root.debug
        where = ""
        try:
            if context is not None and context.file:
                where = f" ({context.file}:{context.line})"
        except Exception:
            pass
        try:
            level("Qt: %s%s", message, where)
        except Exception:
            pass

    sys.excepthook = _sys_excepthook
    threading.excepthook = _thread_excepthook
    qInstallMessageHandler(_qt_message_handler)


def main() -> int:
    setup_logging()
    install_crash_capture()
    logging.getLogger(__name__).info("Starting %s v%s", APP_NAME, APP_VERSION)

    app = QApplication(sys.argv)
    app.setApplicationName("MultiProfile Launcher")
    app.setOrganizationName("MultiProfileLauncher")
    app.setQuitOnLastWindowClosed(False)
    window = MainWindow()
    # MainWindow resolves the saved theme while constructing, so take the icon
    # from it instead of drawing one with the default palette here.
    app.setWindowIcon(window.windowIcon())
    start_hidden = (
        window._config_store.settings.get("start_minimized_to_tray", False)
        and QSystemTrayIcon.isSystemTrayAvailable()
    )
    if not start_hidden:
        window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
