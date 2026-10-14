#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Window Layout Manager
======================================================================
A production-ready Windows desktop application (64-bit) for managing,
previewing, and tiling open application windows across one or more
monitors, with a premium dark-themed three-panel UI:

    LEFT    : searchable, checkable list of every real top-level window
              currently open (each row is a distinct HWND -- including
              multiple separate windows of the same browser/app --
              with a strict fixed-height row layout and the window's
              REAL extracted .exe icon, falling back to an emoji glyph
              only when no native icon can be resolved).
    CENTER  : an interactive "Layout Preview" canvas that draws the
              numbered grid cells the checked windows will be tiled
              into, with REAL LIVE DWM thumbnails of each window's
              actual on-screen content rendered inside each cell.
    RIGHT   : a control panel for applying the equal grid layout,
              saving/loading named layouts to disk, and resetting
              window positions back to where they were before tiling.

A monitor switcher across the top lets the user pick which physical
monitor (enumerated live via EnumDisplayMonitors) the layout should be
calculated and applied against.

The right panel's Layout Template selector offers four tiling modes:
Equal Grid, Master+Stack, Columns, and Vertical Stack. The latter three
let the user drag the divider line between any two adjacent windows in
the preview to resize one larger and shrink its neighbor(s) to fit --
e.g. drag a Telegram window's bottom edge down to make it taller while
every other stacked window shrinks to make room. A Gap slider in the
same panel (0 by default) adds a uniform pixel gutter between every
pair of tiled windows across all four templates; at 0 the windows tile
edge-to-edge with no gap, exactly as before. A defensive overlap check
(ensure_no_overlap) runs immediately before any window is actually
moved, guaranteeing two tiled windows can never end up stacked on top
of one another regardless of which template or gap setting is active.

When "Apply Equal Layout" runs, every tiled window is forced through a
strict restore -> HWND_TOPMOST -> HWND_NOTOPMOST -> SetForegroundWindow
sequence (see force_window_to_foreground()) so it is guaranteed to land
on top of the desktop instead of staying correctly resized but hidden
behind other windows. Once at least one window has been tiled
successfully, the manager window then auto-minimizes itself to the
taskbar so it doesn't sit on top of the workspace it just organized.

Requirements:
    Windows 7 / 8 / 8.1 / 10 / 11, 64-bit
    Python 3.7+ (a 64-bit interpreter is recommended on 64-bit Windows)
    PySide6          (already required by MultiProfile Launcher)
    pywin32           (pip install pywin32)

Run:
    python window_layout_manager.py

----------------------------------------------------------------------
64-BIT SAFETY NOTES
----------------------------------------------------------------------
- HWNDs are opaque pointers. On 64-bit Windows they are 64-bit values.
  Every place in this file that touches a HWND does so through pywin32
  (win32gui / win32con / win32process / win32api) or through ctypes
  calls that are declared with ctypes.c_void_p / wintypes.HWND argument
  types -- never through a fixed-width 32-bit integer cast. This avoids
  the classic 64-bit truncation/casting bug where a HWND above the 4GB
  boundary gets silently mangled by code that assumes a 32-bit pointer.
- DWM thumbnail handles (HTHUMBNAIL) are likewise passed around as
  plain Python ints produced directly by ctypes' own marshaling of the
  dwmapi.dll calls -- never reinterpreted or packed/unpacked by hand.
- Monitor handles (HMONITOR) are passed straight through from
  win32api.EnumDisplayMonitors / MonitorFromPoint without modification.
"""

import sys
import os
import json
import time
import ctypes
import ctypes.wintypes as wintypes
import traceback
import math

import win32api
import win32con
import win32gui
import win32process

from PySide6.QtCore import (
    Qt,
    QTimer,
    QRect,
    QSize,
    QObject,
    QAbstractNativeEventFilter,
    Signal,
)
from PySide6.QtGui import (
    QFont,
    QColor,
    QPainter,
    QPen,
    QBrush,
    QImage,
    QPixmap,
    QIcon,
)
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QMainWindow,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QToolButton,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QStatusBar,
    QAbstractItemView,
    QFrame,
    QSizePolicy,
    QInputDialog,
    QSlider,
    QSpinBox,
    QScrollArea,
)


# ======================================================================
# SECTION 0: DPI Awareness (must run before any Qt window is created)
# ======================================================================

def _set_dpi_awareness():
    """
    Requests the highest DPI awareness mode the running OS supports so
    that all GetWindowRect / GetMonitorInfo / SetWindowPos coordinates
    are reported and accepted in real physical pixels, not pixels that
    have been silently scaled by Windows' compatibility shims. Without
    this, every rectangle calculation in this file would be wrong by
    the active display scaling factor (125%, 150%, etc.) on most modern
    laptops and 4K monitors.
    """
    try:
        # Windows 10 1703+: DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        # Windows 8.1+: PROCESS_PER_MONITOR_DPI_AWARE = 2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


_set_dpi_awareness()


# ======================================================================
# SECTION 1: Dynamic icon/emoji resolution by executable name
# ======================================================================

# Maps a lowercase substring of the process executable name to a
# representative emoji glyph. Matching is "first substring found wins"
# in declaration order below, so more specific entries are listed before
# more generic catch-alls. This purely cosmetic layer never affects
# which HWND gets selected or tiled -- it only decorates the list row.
_EXE_EMOJI_RULES = [
    # Browsers
    ("chrome", "🌐"),
    ("msedge", "🌐"),
    ("firefox", "🦊"),
    ("waterfox", "🌐"),
    ("brave", "🦁"),
    ("opera", "🌐"),
    ("iexplore", "🌐"),
    ("vivaldi", "🌐"),
    # Chat / communication
    ("telegram", "💬"),
    ("whatsapp", "💬"),
    ("slack", "💬"),
    ("discord", "💬"),
    ("teams", "💬"),
    ("skype", "💬"),
    ("signal", "💬"),
    ("messenger", "💬"),
    # Terminals / shells
    ("windowsterminal", "💻"),
    ("powershell", "💻"),
    ("pwsh", "💻"),
    ("cmd.exe", "💻"),
    ("conhost", "💻"),
    ("wsl", "💻"),
    ("git-bash", "💻"),
    ("mintty", "💻"),
    # Code editors / IDEs
    ("code.exe", "🧩"),
    ("code - insiders", "🧩"),
    ("pycharm", "🧩"),
    ("idea64", "🧩"),
    ("idea.exe", "🧩"),
    ("webstorm", "🧩"),
    ("clion", "🧩"),
    ("rider64", "🧩"),
    ("sublime_text", "🧩"),
    ("notepad++", "📝"),
    ("notepad", "📝"),
    ("devenv", "🧩"),  # Visual Studio
    # File management
    ("explorer", "📂"),
    ("totalcmd", "📂"),
    ("filezilla", "📂"),
    # Office / productivity
    ("winword", "📄"),
    ("excel", "📊"),
    ("powerpnt", "📽️"),
    ("outlook", "📧"),
    ("onenote", "🗒️"),
    ("acrobat", "📕"),
    ("acrord32", "📕"),
    ("foxitreader", "📕"),
    # Media
    ("spotify", "🎵"),
    ("vlc", "🎬"),
    ("wmplayer", "🎬"),
    ("itunes", "🎵"),
    ("musicbee", "🎵"),
    ("obs64", "🎥"),
    ("obs32", "🎥"),
    ("obs.exe", "🎥"),
    # Creative / design
    ("photoshop", "🎨"),
    ("illustrator", "🎨"),
    ("figma", "🎨"),
    ("blender", "🎨"),
    ("premiere", "🎬"),
    ("aftereffects", "🎬"),
    # Dev tools / databases
    ("docker", "🐳"),
    ("postman", "📡"),
    ("mysql", "🗄️"),
    ("ssms", "🗄️"),
    ("dbeaver", "🗄️"),
    ("github", "🐙"),
    ("sourcetree", "🐙"),
    # Games / launchers
    ("steam", "🎮"),
    ("epicgameslauncher", "🎮"),
    ("battle.net", "🎮"),
    # System utilities
    ("taskmgr", "⚙️"),
    ("control", "⚙️"),
    ("regedit", "⚙️"),
    ("mmc", "⚙️"),
    ("calculator", "🧮"),
    ("calc.exe", "🧮"),
    ("snippingtool", "✂️"),
    ("paint", "🎨"),
    ("zoom", "📹"),
]

_DEFAULT_EMOJI = "🗔"  # generic window glyph fallback


def get_emoji_for_process(process_name):
    """
    Returns a single emoji glyph representative of the given executable
    name (case-insensitive substring match against a curated rule
    table), falling back to a generic window glyph for anything
    unrecognized so every row in the list always has *some* icon.
    """
    if not process_name:
        return _DEFAULT_EMOJI
    lowered = process_name.lower()
    for needle, emoji in _EXE_EMOJI_RULES:
        if needle in lowered:
            return emoji
    return _DEFAULT_EMOJI


# ======================================================================
# SECTION 2: Monitor enumeration
# ======================================================================

class MonitorInfo(object):
    """
    Plain data holder describing one physical monitor attached to the
    system, as reported live by EnumDisplayMonitors + GetMonitorInfo.
    """

    __slots__ = (
        "handle",       # HMONITOR
        "device_name",  # e.g. "\\\\.\\DISPLAY1"
        "is_primary",
        "full_rect",    # (left, top, right, bottom) -- entire monitor
        "work_rect",    # (left, top, right, bottom) -- excludes taskbar
        "index",        # 1-based display order, left-to-right by left edge
    )

    def __init__(self, handle, device_name, is_primary, full_rect, work_rect, index):
        self.handle = handle
        self.device_name = device_name
        self.is_primary = is_primary
        self.full_rect = full_rect
        self.work_rect = work_rect
        self.index = index

    @property
    def width(self):
        return self.full_rect[2] - self.full_rect[0]

    @property
    def height(self):
        return self.full_rect[3] - self.full_rect[1]

    @property
    def label(self):
        return "Monitor {}{}".format(
            self.index, " (Primary)" if self.is_primary else ""
        )

    @property
    def resolution_label(self):
        return "{} x {}".format(self.width, self.height)


def enumerate_monitors():
    """
    Enumerates every physical monitor currently attached to the system
    using the real Win32 multi-monitor API (EnumDisplayMonitors). Each
    returned MonitorInfo carries both the monitor's full rectangle and
    its *work area* rectangle (i.e. with the taskbar and any other
    docked appbars subtracted out), exactly as GetMonitorInfo reports
    it -- this is what every layout calculation in this file targets so
    tiled windows never get hidden behind the taskbar.

    Monitors are returned sorted left-to-right by their horizontal
    position, and numbered starting at 1, which matches how Windows'
    own Display Settings page numbers them in the common case.
    """
    raw_monitors = win32api.EnumDisplayMonitors()
    collected = []

    for monitor_handle, _hdc, _rect in raw_monitors:
        try:
            info = win32api.GetMonitorInfo(monitor_handle)
        except Exception:
            continue
        device_name = info.get("Device", "")
        is_primary = bool(info.get("Flags", 0) & win32con.MONITORINFOF_PRIMARY)
        full_rect = info["Monitor"]
        work_rect = info["Work"]
        collected.append(
            {
                "handle": monitor_handle,
                "device_name": device_name,
                "is_primary": is_primary,
                "full_rect": full_rect,
                "work_rect": work_rect,
            }
        )

    # Sort left-to-right by the left edge of the full monitor rect, which
    # is how a user visually scans multi-monitor setups arranged side by
    # side (the common arrangement).
    collected.sort(key=lambda m: m["full_rect"][0])

    monitors = []
    for idx, m in enumerate(collected, start=1):
        monitors.append(
            MonitorInfo(
                handle=m["handle"],
                device_name=m["device_name"],
                is_primary=m["is_primary"],
                full_rect=m["full_rect"],
                work_rect=m["work_rect"],
                index=idx,
            )
        )
    return monitors


def get_primary_monitor():
    monitors = enumerate_monitors()
    for m in monitors:
        if m.is_primary:
            return m
    # Defensive fallback -- should never happen on a real Windows system,
    # but guarantees we always return *something* usable.
    if monitors:
        return monitors[0]
    work_left, work_top, work_right, work_bottom = (
        0, 0,
        win32api.GetSystemMetrics(win32con.SM_CXSCREEN),
        win32api.GetSystemMetrics(win32con.SM_CYSCREEN),
    )
    return MonitorInfo(
        handle=None,
        device_name="",
        is_primary=True,
        full_rect=(work_left, work_top, work_right, work_bottom),
        work_rect=(work_left, work_top, work_right, work_bottom),
        index=1,
    )


# ======================================================================
# SECTION 2B: Global hotkey support (Ctrl+Alt+L -> Apply Equal Layout)
# ======================================================================
#
# Implemented with a direct ctypes wrapper around the native Win32
# RegisterHotKey/UnregisterHotKey entry points -- never an external
# "keyboard"-style package -- so the whole application stays a single,
# dependency-free script. RegisterHotKey asks the OS itself to deliver a
# WM_HOTKEY message to our own window's message queue the instant the
# chosen key combination is pressed ANYWHERE on the system, including
# while this app is minimized or a completely different application has
# focus -- exactly the "global" behavior required. That WM_HOTKEY
# message is intercepted via a QAbstractNativeEventFilter installed on
# the QApplication instance, which lets Qt's own native Win32 message
# loop hand us the raw MSG struct without running a second message pump
# of our own (the classic way this kind of integration either misses
# messages or ends up blocking the UI thread).
# ======================================================================

WM_HOTKEY = 0x0312

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000

VK_L = 0x4C

HOTKEY_ID_APPLY_LAYOUT = 1

_user32_hotkey = ctypes.windll.user32
_user32_hotkey.RegisterHotKey.argtypes = [
    wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT,
]
_user32_hotkey.RegisterHotKey.restype = wintypes.BOOL
_user32_hotkey.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
_user32_hotkey.UnregisterHotKey.restype = wintypes.BOOL


def register_global_hotkey(hwnd, hotkey_id=HOTKEY_ID_APPLY_LAYOUT,
                            modifiers=MOD_CONTROL | MOD_ALT | MOD_NOREPEAT,
                            vk=VK_L):
    """
    Registers a single system-wide hotkey (default: Ctrl+Alt+L) bound to
    `hwnd`'s message queue. Returns True on success. Safe to call more
    than once for the same id -- Windows will simply report failure on
    a second attempt (e.g. a previous instance of this app already
    holds the registration), which we treat as non-fatal.
    """
    try:
        return bool(
            _user32_hotkey.RegisterHotKey(
                wintypes.HWND(hwnd), hotkey_id, modifiers, vk
            )
        )
    except Exception:
        return False


def unregister_global_hotkey(hwnd, hotkey_id=HOTKEY_ID_APPLY_LAYOUT):
    try:
        _user32_hotkey.UnregisterHotKey(wintypes.HWND(hwnd), hotkey_id)
    except Exception:
        pass


class GlobalHotkeyEventFilter(QAbstractNativeEventFilter):
    """
    Installed once on the QApplication instance (see main() in Section
    13). Watches every native Win32 message Qt's own event loop
    receives and fires `callback()` whenever it sees our specific
    WM_HOTKEY id, regardless of which widget -- if any -- currently has
    focus. This is what makes the hotkey work even while the manager
    window is minimized or another application is in the foreground.
    """

    def __init__(self, callback, hotkey_id=HOTKEY_ID_APPLY_LAYOUT):
        super().__init__()
        self._callback = callback
        self._hotkey_id = hotkey_id

    def nativeEventFilter(self, event_type, message):
        try:
            if event_type in (b"windows_generic_MSG", "windows_generic_MSG"):
                msg = wintypes.MSG.from_address(int(message))
                if msg.message == WM_HOTKEY and msg.wParam == self._hotkey_id:
                    self._callback()
        except Exception:
            pass
        return False, 0


# ======================================================================
# SECTION 2C: DWM Extended Frame Bounds (anti-shadow-gap)
# ======================================================================
#
# Windows 10/11 adds an invisible drop-shadow border around most
# top-level windows. win32gui.GetWindowRect() reports a rect that
# INCLUDES this shadow, but the actual visible frame is smaller.
# DwmGetWindowAttribute(DWMWA_EXTENDED_FRAME_BOUNDS) returns the true
# visible rect. When tiling windows edge-to-edge (Gap=0), the raw
# GetWindowRect shadow creates a visible gap between tiles. This
# section provides helpers to compute the shadow offset per-window and
# to adjust SetWindowPos coordinates so the VISIBLE edges land exactly
# where we want them.
# ======================================================================

DWMWA_EXTENDED_FRAME_BOUNDS = 9


def get_window_extended_frame_rect(hwnd):
    """
    Returns (left, top, right, bottom) of the ACTUAL VISIBLE frame
    bounds of `hwnd`, excluding the invisible Windows 10/11 drop-shadow
    borders that win32gui.GetWindowRect incorrectly includes. Falls
    back to GetWindowRect if DWM fails (pre-8.1 OS, protected windows).
    """
    try:
        rect = wintypes.RECT()
        hr = ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rect),
            ctypes.sizeof(rect),
        )
        if hr == 0:
            return (rect.left, rect.top, rect.right, rect.bottom)
    except Exception:
        pass
    return win32gui.GetWindowRect(hwnd)


def get_shadow_offsets(hwnd):
    """
    Returns (left, top, right, bottom) shadow pixel counts -- the
    difference between the raw GetWindowRect (includes invisible shadow)
    and the DWM visible frame (excludes shadow) for `hwnd`. These
    offsets are subtracted from target coordinates when calling
    SetWindowPos so the VISIBLE edges of the window land exactly at the
    intended position, eliminating the inter-tile gap at Gap=0.
    """
    try:
        raw = win32gui.GetWindowRect(hwnd)
    except Exception:
        return (0, 0, 0, 0)
    try:
        visible = get_window_extended_frame_rect(hwnd)
    except Exception:
        return (0, 0, 0, 0)
    return (
        raw[0] - visible[0],    # left shadow width
        raw[1] - visible[1],    # top shadow height
        visible[2] - raw[2],    # right shadow width
        visible[3] - raw[3],    # bottom shadow height
    )


# ======================================================================
# SECTION 3: Top-level window enumeration (per-HWND, never bundled)
# ======================================================================
#
# IMPORTANT DESIGN NOTE ON "TAB SEPARATION":
#
# Win32 has no native concept of browser tabs -- a tab inside Chrome is
# an internal renderer-process construct that Windows itself cannot see;
# the OS only ever sees Chrome's actual top-level HWNDs. What Chrome
# *does* expose as separate, real, independently movable top-level
# windows are: (a) separate browser windows opened via "New Window", and
# (b) tabs the user has dragged out into their own window. Each of those
# is a genuinely distinct HWND with its own window title, and this
# enumerator lists and tiles every one of them completely independently
# -- selecting one never touches any sibling window's HWND. This is the
# correct and complete interpretation of "isolated per-window targeting"
# that is achievable through the public Win32 API, and it is exactly how
# the Windows Taskbar itself treats multiple Chrome windows (each gets
# its own thumbnail/flyout entry, keyed by HWND).
# ======================================================================

EXCLUDED_CLASS_NAMES = {
    "Shell_TrayWnd",
    "Shell_SecondaryTrayWnd",
    "Progman",
    "WorkerW",
    "Button",
    "DV2ControlHost",
    "MsgrIMEWindowClass",
    "SysShadow",
    "Windows.UI.Core.CoreWindow",
    "ApplicationFrameWindow_DesktopBackgroundWindow",
    "tooltips_class32",
    "NotifyIconOverflowWindow",
    "TaskListThumbnailWnd",     # taskbar's own live-thumbnail flyout window
    "TaskListOverlayWnd",
    "XamlExplorerHostIslandWindow",
}

EXCLUDED_TITLES = {
    "",
    "Program Manager",
    "Windows Input Experience",
}

# Our own application's window-class prefix; used as a second layer of
# protection (in addition to the live own_hwnd check) against ever
# enumerating, previewing, or tiling our own GUI -- which would risk a
# feedback loop (a window trying to tile/preview itself).
_OWN_APP_MARKER_TITLE_PREFIX = "Window Layout Manager"


def is_window_cloaked(hwnd):
    """
    True if DWM has cloaked the window (used for many invisible UWP host
    frames that otherwise still report IsWindowVisible() == True).
    """
    try:
        DWMWA_CLOAKED = 14
        cloaked_val = ctypes.c_int(0)
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, DWMWA_CLOAKED, ctypes.byref(cloaked_val), ctypes.sizeof(cloaked_val)
        )
        return cloaked_val.value != 0
    except Exception:
        return False


def is_taskbar_window(hwnd, own_hwnd=None):
    """
    True if `hwnd` is a real, independent, taskbar-visible top-level
    window that should be listed as its own row and is safe to tile on
    its own, without affecting any other window.
    """
    if own_hwnd is not None and hwnd == own_hwnd:
        return False

    try:
        if not win32gui.IsWindow(hwnd):
            return False
        if not win32gui.IsWindowVisible(hwnd):
            return False
        if win32gui.GetParent(hwnd) != 0:
            return False

        title = win32gui.GetWindowText(hwnd)
        if not title or not title.strip():
            return False
        if title in EXCLUDED_TITLES:
            return False
        if title.startswith(_OWN_APP_MARKER_TITLE_PREFIX):
            return False

        owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
        if owner != 0:
            return False

        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        is_tool_window = bool(ex_style & win32con.WS_EX_TOOLWINDOW)
        is_app_window = bool(ex_style & win32con.WS_EX_APPWINDOW)
        if is_tool_window and not is_app_window:
            return False

        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if style & win32con.WS_DISABLED:
            return False

        class_name = win32gui.GetClassName(hwnd)
        if class_name in EXCLUDED_CLASS_NAMES:
            return False

        if is_window_cloaked(hwnd):
            return False

        return True
    except Exception:
        return False


def get_process_full_path(pid):
    """
    Best-effort resolution of the FULL executable path for a PID (e.g.
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"), using
    only the Win32 API. Returns "" if the process cannot be opened
    (protected/elevated process, already exited, etc.) rather than
    raising, so one inaccessible process never aborts enumeration of
    the rest. This is the same lookup get_process_name() uses, just
    returning the full path instead of the trailing file name -- kept
    as its own function because the icon-extraction code path needs
    the full path (to fall back to extracting an icon straight from
    the .exe on disk when the live window doesn't hand us one).
    """
    if not pid:
        return ""
    h_process = None
    try:
        h_process = win32api.OpenProcess(
            win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid
        )
        return win32process.GetModuleFileNameEx(h_process, 0)
    except Exception:
        try:
            h_process = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
                False,
                pid,
            )
            return win32process.GetModuleFileNameEx(h_process, 0)
        except Exception:
            return ""
    finally:
        if h_process is not None:
            try:
                win32api.CloseHandle(h_process)
            except Exception:
                pass


def get_process_name(pid):
    """Returns just the trailing executable file name (e.g. 'chrome.exe')."""
    full_path = get_process_full_path(pid)
    return full_path.split("\\")[-1] if full_path else ""


def get_monitor_for_window(hwnd):
    """
    Returns the HMONITOR that the given window currently has the
    greatest overlap with, used to show which physical monitor each
    listed window currently lives on (informational only).
    """
    try:
        return win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
    except Exception:
        return None


def enumerate_taskbar_windows(own_hwnd=None):
    """
    Enumerates every real top-level, taskbar-eligible window currently
    open on the system. Each result is a fully independent record keyed
    by its own unique HWND -- two separate Chrome windows (or two
    separate windows of any other app) always produce two separate
    entries here, each with its own title text, never merged or
    deduplicated by process.

    Returns a list of dicts:
        {
            'hwnd': int,
            'title': str,
            'pid': int,
            'process_name': str,
            'is_minimized': bool,
            'monitor_handle': HMONITOR or None,
        }
    """
    results = []

    def _enum_callback(hwnd, _extra):
        if is_taskbar_window(hwnd, own_hwnd=own_hwnd):
            title = win32gui.GetWindowText(hwnd)
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception:
                pid = 0

            exe_path = get_process_full_path(pid)
            process_name = exe_path.split("\\")[-1] if exe_path else ""
            is_minimized = bool(win32gui.IsIconic(hwnd))
            monitor_handle = get_monitor_for_window(hwnd)

            results.append(
                {
                    "hwnd": hwnd,
                    "title": title,
                    "pid": pid,
                    "process_name": process_name,
                    "exe_path": exe_path,
                    "is_minimized": is_minimized,
                    "monitor_handle": monitor_handle,
                }
            )
        return True

    win32gui.EnumWindows(_enum_callback, None)

    # Stable, predictable ordering: group by process name, then by title,
    # so repeated refreshes don't visually shuffle unrelated rows.
    results.sort(key=lambda w: (w["process_name"].lower(), w["title"].lower()))
    return results


# ======================================================================
# SECTION 4: Smart equal-grid layout algorithm
# ======================================================================

def calculate_grid_dimensions(n):
    """
    Given N selected windows, returns (rows, cols) per the required
    rules:
        N = 1       -> (1, 1)   fills the screen
        N = 2       -> (1, 2)   two equal vertical columns
        N = 3 or 4  -> (2, 2)   balanced 2x2 grid
        N > 4       -> the most BALANCED rows x cols grid that can hold
                        all N windows (rows*cols >= N), preferring a
                        near-square shape over minimizing empty cells --
                        e.g. N=5 -> 2x3 (one empty cell, balanced) rather
                        than 1x5 (zero empty cells, but a degenerate
                        single-row strip that is not useful for tiling).
    """
    if n <= 0:
        return (0, 0)
    if n == 1:
        return (1, 1)
    if n == 2:
        return (1, 2)
    if n in (3, 4):
        return (2, 2)

    best = None
    sqrt_n = int(math.ceil(math.sqrt(n)))
    for rows in range(1, sqrt_n + 2):
        cols = math.ceil(n / rows)
        cells = rows * cols
        waste = cells - n
        imbalance = abs(cols - rows)
        score = (imbalance, waste, rows)
        if best is None or score < best[0]:
            best = (score, rows, cols)

    _, rows, cols = best
    return (rows, cols)


def compute_tile_rects(work_area, n):
    """
    Given a working-area rectangle (left, top, right, bottom) and a
    window count N, returns a list of N (x, y, width, height) tuples --
    one slot per window, in row-major order -- that perfectly tile the
    working area edge-to-edge with zero gaps and zero overlap. Any
    leftover pixels from integer division are distributed one-by-one
    across the first few columns/rows so every pixel of the working
    area is covered exactly once.
    """
    left, top, right, bottom = work_area
    total_width = right - left
    total_height = bottom - top

    rows, cols = calculate_grid_dimensions(n)
    if rows == 0 or cols == 0:
        return []

    base_col_width = total_width // cols
    extra_width_px = total_width - (base_col_width * cols)

    base_row_height = total_height // rows
    extra_height_px = total_height - (base_row_height * rows)

    col_widths = [
        base_col_width + (1 if c < extra_width_px else 0) for c in range(cols)
    ]
    row_heights = [
        base_row_height + (1 if r < extra_height_px else 0) for r in range(rows)
    ]

    col_x_offsets = [left]
    for c in range(1, cols):
        col_x_offsets.append(col_x_offsets[-1] + col_widths[c - 1])

    row_y_offsets = [top]
    for r in range(1, rows):
        row_y_offsets.append(row_y_offsets[-1] + row_heights[r - 1])

    rects = []
    for index in range(n):
        row = index // cols
        col = index % cols
        x = col_x_offsets[col]
        y = row_y_offsets[row]
        w = col_widths[col]
        h = row_heights[row]
        rects.append((x, y, w, h))

    return rects


# ----------------------------------------------------------------------
# Custom / asymmetric template math (Feature 3)
# ----------------------------------------------------------------------
#
# These two templates are computed directly against a work-area
# rectangle in real monitor pixel space -- exactly like
# compute_tile_rects() above -- so the SAME function drives both the
# scaled-down preview drawing in LayoutPreviewCanvas and the real,
# final SetWindowPos rectangles applied to actual HWNDs. There is only
# ever one source of truth for "where does window i end up"; the
# preview never uses a separate approximate copy of this math.

MIN_RATIO_FRACTION = 0.08  # no manually-resized slot may shrink below 8%

# Absolute pixel floor for a single dragged slot -- keeps a window
# usably wide/tall (e.g. Telegram's own minimum chat-list width) no
# matter how many windows are checked, separate from the percentage-
# based MIN_RATIO_FRACTION used by Master+Stack (which only ever has
# two regions -- the master and the stack -- so an 8% floor there is
# always satisfiable).
MIN_SLOT_PX = 24


def min_ratio_for_count(count):
    """
    Returns a safe per-slot minimum ratio for an N-way Columns or Rows
    split being resized by divider drag. A flat 8% floor (as used by
    the always-two-region Master+Stack split) becomes mathematically
    impossible once N more than ~12 -- N slots each needing >= 8% would
    sum to over 100%. Instead, scale the floor down with N so it is
    always comfortably satisfiable (here, 40% of an exactly-equal
    share), which keeps every divider smoothly draggable at any window
    count -- including the 16+ Telegram windows this template exists
    for -- while still preventing a slot from being dragged down to
    zero width/height.
    """
    if count <= 0:
        return MIN_RATIO_FRACTION
    return min(MIN_RATIO_FRACTION, 0.4 / count)


def normalize_ratios(ratios, count):
    """
    Returns a list of exactly `count` positive floats summing to 1.0.
    Falls back to a perfectly equal split if `ratios` is missing, the
    wrong length, or degenerate (sums to <= 0) -- which is exactly the
    state the very first time a given window count is seen, before the
    user has dragged any divider.
    """
    if count <= 0:
        return []
    if not ratios or len(ratios) != count:
        return [1.0 / count] * count
    total = sum(ratios)
    if total <= 0:
        return [1.0 / count] * count
    return [r / total for r in ratios]


def compute_master_stack_rects(work_area, n, master_ratio, stack_ratios):
    """
    "Master + Stack" template: window 0 (the "master") occupies a full-
    height column on the left whose width is `master_ratio` (0..1) of
    the total working width (64% by default, per the classic
    master/stack tiling convention). Every remaining window is stacked
    vertically in the rest of the width, with each stacked slot's
    height controlled by `stack_ratios` (normalized to sum to 1 across
    the N-1 stack slots). Returns N (x, y, w, h) tuples, index 0 always
    being the master.
    """
    left, top, right, bottom = work_area
    total_w = right - left
    total_h = bottom - top

    if n <= 0:
        return []
    if n == 1:
        return [(left, top, total_w, total_h)]

    master_ratio = max(MIN_RATIO_FRACTION, min(1.0 - MIN_RATIO_FRACTION, master_ratio))
    master_w = max(1, int(round(total_w * master_ratio)))
    master_w = min(master_w, total_w - 1)
    stack_w = total_w - master_w

    rects = [(left, top, master_w, total_h)]

    stack_count = n - 1
    ratios = normalize_ratios(stack_ratios, stack_count)

    stack_x = left + master_w
    y_cursor = top
    for i, ratio in enumerate(ratios):
        if i == len(ratios) - 1:
            h = bottom - y_cursor  # absorb rounding remainder on the last slot
        else:
            h = max(1, int(round(total_h * ratio)))
        rects.append((stack_x, y_cursor, stack_w, h))
        y_cursor += h

    return rects


def compute_columns_rects(work_area, n, column_ratios):
    """
    "Custom Columns" template: all N windows sit side-by-side in a
    single full-height row, each column's width independently
    controlled by `column_ratios` (normalized to sum to 1 across N
    columns). With all ratios equal this looks identical to the N==2
    equal-grid case, but every divider remains independently draggable
    for arbitrary "customized column splits".
    """
    left, top, right, bottom = work_area
    total_w = right - left
    total_h = bottom - top

    if n <= 0:
        return []

    ratios = normalize_ratios(column_ratios, n)

    rects = []
    x_cursor = left
    for i, ratio in enumerate(ratios):
        if i == len(ratios) - 1:
            w = right - x_cursor  # absorb rounding remainder on the last column
        else:
            w = max(1, int(round(total_w * ratio)))
        rects.append((x_cursor, top, w, total_h))
        x_cursor += w

    return rects


def compute_rows_rects(work_area, n, row_ratios):
    """
    "Vertical Stack" template: all N windows sit one above another in a
    single full-width column, each row's height independently
    controlled by `row_ratios` (normalized to sum to 1 across N rows).
    This is the vertical mirror of compute_columns_rects() above --
    same independent-divider-drag behavior, just stacked top-to-bottom
    instead of side-by-side, which is the natural fit for tall, narrow
    windows such as Telegram/WhatsApp/Slack columns.
    """
    left, top, right, bottom = work_area
    total_w = right - left
    total_h = bottom - top

    if n <= 0:
        return []

    ratios = normalize_ratios(row_ratios, n)

    rects = []
    y_cursor = top
    for i, ratio in enumerate(ratios):
        if i == len(ratios) - 1:
            h = bottom - y_cursor  # absorb rounding remainder on the last row
        else:
            h = max(1, int(round(total_h * ratio)))
        rects.append((left, y_cursor, total_w, h))
        y_cursor += h

    return rects


# Minimum window dimensions (pixels) that most applications -- and
# Telegram in particular -- will honour via WM_GETMINMAXINFO. Windows
# below these dimensions are resized by the OS back up to the minimum
# instead of landing where we asked, which produces visual overlap.
# These defaults match Telegram Desktop's observed minimum:
# ~380 px wide × ~300 px tall. They are used by the Custom Grid
# template to warn the user before they apply a layout that would
# produce cells smaller than the application can actually render.
DEFAULT_MIN_WIN_W = 380
DEFAULT_MIN_WIN_H = 300


def compute_custom_grid_rects(work_area, n, grid_rows, grid_cols):
    """
    "Custom Grid" template: places all N windows into a grid of exactly
    `grid_rows` rows × `grid_cols` columns. Every cell is the same size
    (work-area width / grid_cols, work-area height / grid_rows). If
    grid_rows * grid_cols > n, the last row may be partially empty;
    those trailing empty cells are simply not returned, so exactly N
    rects are always produced and no window is ever placed off-screen.

    Unlike compute_tile_rects() -- which auto-selects a near-square
    grid for you -- this function uses the EXACT row/column counts the
    user specified, so a user who wants 6 columns × 2 rows always gets
    that layout even for 12 windows where auto-sizing would have chosen
    something different.
    """
    left, top, right, bottom = work_area
    total_w = right - left
    total_h = bottom - top

    if n <= 0 or grid_rows <= 0 or grid_cols <= 0:
        return []

    # integer-division with remainder distributed across first columns/rows
    base_w = total_w // grid_cols
    extra_w = total_w - base_w * grid_cols
    col_widths = [base_w + (1 if c < extra_w else 0) for c in range(grid_cols)]

    base_h = total_h // grid_rows
    extra_h = total_h - base_h * grid_rows
    row_heights = [base_h + (1 if r < extra_h else 0) for r in range(grid_rows)]

    col_x = [left + sum(col_widths[:c]) for c in range(grid_cols)]
    row_y = [top + sum(row_heights[:r]) for r in range(grid_rows)]

    rects = []
    for index in range(n):
        r = index // grid_cols
        c = index % grid_cols
        if r >= grid_rows:
            break  # more windows than cells: stop (shouldn't happen with sane inputs)
        rects.append((col_x[c], row_y[r], col_widths[c], row_heights[r]))

    return rects


def check_grid_fits_min_size(work_area, n, grid_rows, grid_cols,
                              min_w=DEFAULT_MIN_WIN_W, min_h=DEFAULT_MIN_WIN_H):
    """
    Returns (fits, cell_w, cell_h, needed_cols, needed_rows) where:
    - fits         : True if every cell is >= (min_w, min_h)
    - cell_w/cell_h: actual cell dimensions for the requested grid
    - needed_cols  : smallest column count that fits without going below min_w
    - needed_rows  : smallest row count that fits without going below min_h

    Use this before calling Apply to warn the user when the requested
    grid would produce windows smaller than Telegram (or any other app)
    can actually render.
    """
    left, top, right, bottom = work_area
    total_w = right - left
    total_h = bottom - top

    cell_w = total_w // max(1, grid_cols)
    cell_h = total_h // max(1, grid_rows)
    fits = (cell_w >= min_w) and (cell_h >= min_h)

    import math
    needed_cols = min(grid_cols, math.floor(total_w / max(1, min_w)))
    needed_rows = min(grid_rows, math.floor(total_h / max(1, min_h)))

    return fits, cell_w, cell_h, max(1, needed_cols), max(1, needed_rows)


# supports. "equal", "columns", and "rows" have NO cap -- they always
# grow to fit every checked window, exactly like calculate_grid_dimensions()
# does for "equal" (e.g. 16 windows in Columns -> 16 thin side-by-side
# slots, 16 windows in Vertical Stack -> 16 thin stacked rows). Only
# "master_stack" intentionally caps out at a sane, readable slot count
# (a master pane plus more than a handful of tiny stacked slivers stops
# being useful), past which Window Cycling (Feature 4) takes over to
# rotate the overflow windows through the available slots instead of
# cramming them in.
TEMPLATE_MAX_SLOTS = {
    "equal": None,
    "master_stack": 6,
    "columns": None,
    "rows": None,
    "custom_grid": None,    # user controls rows×cols; no artificial cap
}

TEMPLATE_DISPLAY_LABELS = {
    "equal": "Equal Grid",
    "master_stack": "Master+Stack",
    "columns": "Columns",
    "rows": "Vertical Stack",
    "custom_grid": "Custom Grid",
}


# ----------------------------------------------------------------------
# Inter-window gap (spacing) support
# ----------------------------------------------------------------------
#
# All four templates above compute perfectly edge-to-edge rectangles
# (zero gap, zero overlap) against a work-area. apply_gap_to_rects()
# is a single, template-agnostic post-processing pass that shrinks
# every rectangle towards its own center by `gap_px` / 2 on each side,
# carving a uniform visual gutter of exactly `gap_px` pixels between
# every pair of adjacent windows (and leaving a half-gap margin against
# the screen edge, so the whole tiled group still stays centered in
# the work area). gap_px == 0 reproduces the original edge-to-edge
# behavior exactly -- nothing changes for users who don't touch the
# slider. Because this runs as the very last step before rectangles
# leave get_monitor_space_tile_rects(), it is applied identically to
# the live preview and to the real SetWindowPos calls, and it can
# never introduce overlap: every cell only ever shrinks, never grows.

MAX_GAP_PX = 60


def rects_overlap(a, b):
    """True if (x, y, w, h) rectangles a and b share any pixel area."""
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return ax < bx + bw and bx < ax + aw and ay < by + bh and by < ay + ah


def any_overlap(rects):
    """True if any two rectangles in `rects` overlap."""
    for i in range(len(rects)):
        for j in range(i + 1, len(rects)):
            if rects_overlap(rects[i], rects[j]):
                return True
    return False


def ensure_no_overlap(rects, work_area=None):
    """
    Defensive safety net, run once just before any rectangle list is
    actually applied to real windows via SetWindowPos. Every template's
    own math already produces strictly non-overlapping rectangles by
    construction (this is checked first and is a true no-op pass-
    through in normal operation), but if `rects` is ever found to
    contain an overlap -- from a future template, an edge-case monitor
    geometry, a corrupted saved-layout file, or any other unforeseen
    source -- this discards it entirely and falls back to the
    always-non-overlapping equal grid (compute_tile_rects) against
    `work_area`, so two tiled windows can never end up visually
    stacked on top of one another no matter what produced the input.
    """
    if not any_overlap(rects):
        return list(rects)
    if work_area is None:
        left = min(r[0] for r in rects)
        top = min(r[1] for r in rects)
        right = max(r[0] + r[2] for r in rects)
        bottom = max(r[1] + r[3] for r in rects)
        work_area = (left, top, right, bottom)
    return compute_tile_rects(work_area, len(rects))



def apply_gap_to_rects(rects, gap_px):
    """
    Shrinks each (x, y, w, h) rectangle in `rects` by `gap_px` / 2 on
    every side (clamped so a rectangle can never invert/shrink past
    zero size), producing a uniform gap between every pair of tiles
    that were previously touching edge-to-edge. Pass gap_px <= 0 to
    get the original rectangles back unchanged.
    """
    if gap_px <= 0 or not rects:
        return list(rects)

    half = gap_px / 2.0
    out = []
    for (x, y, w, h) in rects:
        inset_w = min(half, max(0.0, (w - 1) / 2.0))
        inset_h = min(half, max(0.0, (h - 1) / 2.0))
        nx = x + inset_w
        ny = y + inset_h
        nw = max(1, w - 2 * inset_w)
        nh = max(1, h - 2 * inset_h)
        out.append((int(round(nx)), int(round(ny)), int(round(nw)), int(round(nh))))
    return out


# ======================================================================
# SECTION 5: Window moving / state utilities
# ======================================================================

def restore_if_minimized_or_maximized(hwnd):
    """
    Restores a window out of the minimized or maximized state before we
    attempt to move/resize it -- SetWindowPos on a minimized window has
    no visible effect, and a maximized window will snap back to
    full-screen on its own the moment it regains focus, undoing our
    layout. Safe to call on a window that is already in the normal
    state (no-op).
    """
    try:
        if win32gui.IsIconic(hwnd) or win32gui.IsZoomed(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    except Exception:
        pass


def is_window_topmost(hwnd):
    """
    Returns True if `hwnd` currently has the WS_EX_TOPMOST extended
    style set, i.e. it is already pinned "Always on Top" at the OS
    level. Used purely to initialize the sidebar pin button's checked
    state to match whatever the window's real, current Win32 state
    already is (which may have been set by some other application, or
    by a previous run of this one).
    """
    try:
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        return bool(ex_style & win32con.WS_EX_TOPMOST)
    except Exception:
        return False


def set_window_topmost(hwnd, topmost):
    """
    Toggles `hwnd` between HWND_TOPMOST and HWND_NOTOPMOST without
    moving, resizing, or stealing focus -- this is the real, immediate
    Win32 state change behind the per-row pin toggle (Feature 5).
    SWP_NOMOVE | SWP_NOSIZE keeps the window exactly where it already
    is; SWP_NOACTIVATE keeps the user's current foreground window from
    changing just because they clicked a pin icon in our list.
    """
    try:
        target = win32con.HWND_TOPMOST if topmost else win32con.HWND_NOTOPMOST
        win32gui.SetWindowPos(
            hwnd, target, 0, 0, 0, 0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
        )
        return True
    except Exception:
        return False


def get_window_rect_safe(hwnd):
    """
    Returns the VISIBLE (extended frame bounds) rect for hwnd, or None
    on failure. Uses DwmGetWindowAttribute(DWMWA_EXTENDED_FRAME_BOUNDS)
    instead of raw GetWindowRect so pre-layout snapshots and animation
    start positions reflect the actual visible area, not the shadow-
    inflated bounding box.
    """
    try:
        return get_window_extended_frame_rect(hwnd)
    except Exception:
        try:
            return win32gui.GetWindowRect(hwnd)
        except Exception:
            return None


def apply_window_rect(hwnd, x, y, width, height):
    """
    Moves and resizes a single window so its VISIBLE area (after DWM
    shadow removal) lands at exactly (x, y, width, height). Adjusts
    the raw SetWindowPos coordinates by the window's shadow offsets to
    eliminate the gap between tiled windows at Gap=0. All values passed
    through are plain Python ints -- no manual pointer-width casting is
    performed anywhere in this call path, keeping this safe on 64-bit
    HWNDs.
    """
    x, y, width, height = int(x), int(y), int(width), int(height)
    sl, st, sr, sb = get_shadow_offsets(hwnd)
    adj_x = x - sl
    adj_y = y - st
    adj_w = width + sl + sr
    adj_h = height + st + sb
    flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_SHOWWINDOW
    try:
        win32gui.SetWindowPos(hwnd, 0, adj_x, adj_y, adj_w, adj_h, flags)
    except Exception:
        pass


def bring_to_front_safely(hwnd):
    """
    Brings a window to the foreground, working around the well-known
    Windows restriction that silently blocks SetForegroundWindow calls
    from background processes by falling back to the classic
    AttachThreadInput technique.
    """
    try:
        win32gui.SetForegroundWindow(hwnd)
        return
    except Exception:
        pass
    try:
        current_thread_id = win32api.GetCurrentThreadId()
        target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
        if target_thread_id != current_thread_id:
            ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, True)
            try:
                win32gui.SetForegroundWindow(hwnd)
            finally:
                ctypes.windll.user32.AttachThreadInput(target_thread_id, current_thread_id, False)
        else:
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def force_window_to_foreground(hwnd, x, y, width, height):
    """
    Moves/resizes `hwnd` into the given VISIBLE rectangle AND forcibly
    pulls it to the absolute foreground of the desktop, fixing the
    common bug where freshly-tiled windows resize correctly but remain
    trapped behind other already-open windows. Shadow offsets are
    applied to the raw SetWindowPos calls so the visible frame lands
    exactly at (x, y, width, height).

    This runs the strict 4-step activation sequence required to defeat
    that z-order/focus-stealing protection:
        1. Restore the window first if it is minimized -- SetWindowPos
           has no visible effect on an iconic window.
        2. Briefly pin it HWND_TOPMOST while moving/resizing -- this
           unconditionally pulls it above every other top-level window
           on the desktop, including windows that are themselves
           topmost, guaranteeing it is no longer hidden behind anything.
        3. Immediately release the pin with HWND_NOTOPMOST -- this
           demotes it back to a normal top-level window so it won't
           permanently float above everything else (which would be
           obnoxious), while *keeping* the elevated z-order position it
           just earned in step 2 (it does not fall back behind whatever
           it was behind before).
        4. Force real OS input focus onto it via SetForegroundWindow
           (with the AttachThreadInput fallback for the case where
           Windows blocks a background process from stealing focus
           outright), so the window is both visually on top AND the
           one that will actually receive keyboard input.
    """
    x, y, width, height = int(x), int(y), int(width), int(height)

    # Compute shadow offsets so the VISIBLE frame lands at (x,y,w,h).
    sl, st, sr, sb = get_shadow_offsets(hwnd)
    adj_x = x - sl
    adj_y = y - st
    adj_w = width + sl + sr
    adj_h = height + st + sb

    try:
        restore_if_minimized_or_maximized(hwnd)
    except Exception:
        pass

    # Step 2: pin TOPMOST while moving -- forces it above everything,
    # including other topmost windows.
    try:
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOPMOST,
            adj_x, adj_y, adj_w, adj_h,
            win32con.SWP_SHOWWINDOW,
        )
    except Exception:
        pass

    # Step 3: release the hard-lock so it remains a normal top-level
    # window afterwards instead of staying permanently pinned above
    # every other window on the system.
    try:
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_NOTOPMOST,
            adj_x, adj_y, adj_w, adj_h,
            win32con.SWP_SHOWWINDOW,
        )
    except Exception:
        pass

    # Step 4: force real OS focus onto it.
    bring_to_front_safely(hwnd)


# ======================================================================
# SECTION 5B: Smooth animated window-move engine (Feature 6)
# ======================================================================
#
# Replaces an instantaneous SetWindowPos "snap" with a short, eased
# interpolation driven entirely by one shared QTimer ticking on the Qt
# UI thread's own event loop -- there is no second thread and no
# blocking sleep anywhere in this class, so the main window stays fully
# responsive (still redraws, still accepts clicks) for the whole
# duration of the animation. Every intermediate frame still goes
# through apply_window_rect(), i.e. the same 64-bit-safe SetWindowPos
# call path used everywhere else in this file.
# ======================================================================

def _ease_out_cubic(t):
    """Standard 'decelerate towards the end' easing curve, t in [0, 1]."""
    t = max(0.0, min(1.0, t))
    inv = 1.0 - t
    return 1.0 - (inv * inv * inv)


class WindowMoveAnimator(QObject):
    """
    Owns one shared QTimer that, while active, advances every in-flight
    window move by one eased step per tick. animate() can be called
    with any number of simultaneous (hwnd, start_rect, end_rect) jobs --
    every window in a freshly-applied grid layout moves together in the
    same pass, which is what makes the whole grid feel like one
    coherent motion instead of windows arriving one at a time.
    """

    TICK_MS = 15  # ~60 fps; cheap enough to never compete with the UI

    def __init__(self, parent=None):
        super().__init__(parent)
        self._timer = QTimer(self)
        self._timer.setInterval(self.TICK_MS)
        self._timer.timeout.connect(self._on_tick)
        self._jobs = []
        self._start_time = None
        self._duration_s = 0.26
        self._on_complete = None

    def animate(self, moves, duration_ms=260, on_complete=None):
        """
        moves: iterable of (hwnd, (sx, sy, sw, sh), (ex, ey, ew, eh)).
        Any animation already in flight is replaced outright -- this
        app only ever runs one layout-apply animation at a time.
        """
        self._jobs = [
            {"hwnd": hwnd, "start": start_rect, "end": end_rect}
            for hwnd, start_rect, end_rect in moves
        ]
        self._duration_s = max(0.05, duration_ms / 1000.0)
        self._on_complete = on_complete
        self._start_time = time.time()

        if not self._jobs:
            if on_complete:
                on_complete()
            return

        if not self._timer.isActive():
            self._timer.start()

    def stop(self):
        self._timer.stop()
        self._jobs = []
        self._on_complete = None

    def _on_tick(self):
        elapsed = time.time() - self._start_time
        t = min(1.0, elapsed / self._duration_s)
        eased = _ease_out_cubic(t)

        for job in self._jobs:
            hwnd = job["hwnd"]
            try:
                if not win32gui.IsWindow(hwnd):
                    continue
                sx, sy, sw, sh = job["start"]
                ex, ey, ew, eh = job["end"]
                x = int(round(sx + (ex - sx) * eased))
                y = int(round(sy + (ey - sy) * eased))
                w = max(1, int(round(sw + (ew - sw) * eased)))
                h = max(1, int(round(sh + (eh - sh) * eased)))
                apply_window_rect(hwnd, x, y, w, h)
            except Exception:
                pass

        if t >= 1.0:
            self._timer.stop()
            callback = self._on_complete
            self._jobs = []
            self._on_complete = None
            if callback:
                try:
                    callback()
                except Exception:
                    pass


# ======================================================================
# SECTION 6: DWM live thumbnail hosting
# ======================================================================
#
# DwmRegisterThumbnail/DwmUpdateThumbnailProperties/DwmUnregisterThumbnail
# are raw dwmapi.dll entry points with no pywin32 wrapper, so we declare
# their ctypes signatures explicitly with correctly 64-bit-safe argument
# types (wintypes.HWND, ctypes.c_void_p for the opaque HTHUMBNAIL,
# ctypes.POINTER(...) for the in/out structures). DWM can only stream a
# live thumbnail of one window onto another *real native HWND* -- it
# cannot render into an arbitrary Qt-painted QWidget surface -- so each
# preview cell in the center canvas gets its own tiny native child
# window (created via win32gui.CreateWindowEx, parented to the Qt
# canvas's native HWND) that exists purely as a DWM thumbnail target and
# is repositioned every time the canvas is resized or the layout
# changes.
# ======================================================================

dwmapi = ctypes.windll.dwmapi


class _DWM_THUMBNAIL_PROPERTIES(ctypes.Structure):
    _fields_ = [
        ("dwFlags", wintypes.DWORD),
        ("rcDestination", wintypes.RECT),
        ("rcSource", wintypes.RECT),
        ("opacity", ctypes.c_byte),
        ("fVisible", wintypes.BOOL),
        ("fSourceClientAreaOnly", wintypes.BOOL),
    ]


DWM_TNP_RECTDESTINATION = 0x00000001
DWM_TNP_RECTSOURCE = 0x00000002
DWM_TNP_OPACITY = 0x00000004
DWM_TNP_VISIBLE = 0x00000008
DWM_TNP_SOURCECLIENTAREAONLY = 0x00000010

dwmapi.DwmRegisterThumbnail.argtypes = [
    wintypes.HWND,
    wintypes.HWND,
    ctypes.POINTER(ctypes.c_void_p),
]
dwmapi.DwmRegisterThumbnail.restype = ctypes.c_long

dwmapi.DwmUnregisterThumbnail.argtypes = [ctypes.c_void_p]
dwmapi.DwmUnregisterThumbnail.restype = ctypes.c_long

dwmapi.DwmUpdateThumbnailProperties.argtypes = [
    ctypes.c_void_p,
    ctypes.POINTER(_DWM_THUMBNAIL_PROPERTIES),
]
dwmapi.DwmUpdateThumbnailProperties.restype = ctypes.c_long

dwmapi.DwmQueryThumbnailSourceSize.argtypes = [
    ctypes.c_void_p,
    ctypes.POINTER(wintypes.SIZE),
]
dwmapi.DwmQueryThumbnailSourceSize.restype = ctypes.c_long


_HOST_WNDCLASS_NAME = "WLM_ThumbnailHostClass"
_host_wndclass_atom = None

# ----------------------------------------------------------------------
# Explicit argtypes/restype declarations for every raw user32.dll /
# kernel32.dll / gdi32.dll entry point used below. Declaring these
# exactly (rather than relying on ctypes' default "treat as C int"
# argument guessing) is what guarantees HWND values -- 64-bit pointers
# on 64-bit Windows -- are marshaled correctly in every call, whether we
# pass a plain Python int or a wintypes.HWND wrapper.
# ----------------------------------------------------------------------
_user32 = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32
_gdi32 = ctypes.windll.gdi32

_user32.DefWindowProcW.argtypes = [
    wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM
]
_user32.DefWindowProcW.restype = ctypes.c_long

_user32.RegisterClassW.argtypes = [ctypes.c_void_p]
_user32.RegisterClassW.restype = wintypes.ATOM

_user32.CreateWindowExW.argtypes = [
    wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, ctypes.c_void_p,
]
_user32.CreateWindowExW.restype = wintypes.HWND

_user32.DestroyWindow.argtypes = [wintypes.HWND]
_user32.DestroyWindow.restype = wintypes.BOOL

_user32.SetWindowPos.argtypes = [
    wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_uint,
]
_user32.SetWindowPos.restype = wintypes.BOOL

_user32.LoadCursorW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR]
_user32.LoadCursorW.restype = wintypes.HANDLE

_gdi32.GetStockObject.argtypes = [ctypes.c_int]
_gdi32.GetStockObject.restype = wintypes.HANDLE

_kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
_kernel32.GetModuleHandleW.restype = wintypes.HINSTANCE

_kernel32.GetLastError.argtypes = []
_kernel32.GetLastError.restype = wintypes.DWORD


def _ensure_host_window_class_registered():
    """
    Registers (once per process) the minimal native window class used
    purely as a DWM thumbnail rendering surface. We use the simplest
    possible WNDPROC (DefWindowProc for everything) since these host
    windows have no UI logic of their own -- DWM paints directly into
    them via window composition.
    """
    global _host_wndclass_atom
    if _host_wndclass_atom is not None:
        return _host_wndclass_atom

    wndproc = ctypes.WINFUNCTYPE(
        ctypes.c_long, wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM
    )(lambda hwnd, msg, wparam, lparam: ctypes.windll.user32.DefWindowProcW(
        hwnd, msg, wparam, lparam
    ))

    # Keep a module-level reference so the ctypes callback is never
    # garbage-collected while the window class is registered.
    global _host_wndproc_ref
    _host_wndproc_ref = wndproc

    class WNDCLASS(ctypes.Structure):
        _fields_ = [
            ("style", ctypes.c_uint),
            ("lpfnWndProc", wndproc.__class__),
            ("cbClsExtra", ctypes.c_int),
            ("cbWndExtra", ctypes.c_int),
            ("hInstance", wintypes.HINSTANCE),
            ("hIcon", wintypes.HICON),
            ("hCursor", wintypes.HANDLE),
            ("hbrBackground", wintypes.HBRUSH),
            ("lpszMenuName", wintypes.LPCWSTR),
            ("lpszClassName", wintypes.LPCWSTR),
        ]

    hinstance = _kernel32.GetModuleHandleW(None)
    wc = WNDCLASS()
    wc.style = 0
    wc.lpfnWndProc = wndproc
    wc.cbClsExtra = 0
    wc.cbWndExtra = 0
    wc.hInstance = hinstance
    wc.hIcon = None
    wc.hCursor = _user32.LoadCursorW(None, 32512)  # IDC_ARROW
    wc.hbrBackground = _gdi32.GetStockObject(4)  # BLACK_BRUSH
    wc.lpszMenuName = None
    # Build an explicit, separately-held wide-string buffer for the
    # class name rather than assigning the Python str directly to the
    # LPCWSTR struct field. ctypes constructs a throwaway internal
    # buffer for that implicit str -> LPCWSTR conversion and does not
    # keep it alive once the assignment statement finishes, which can
    # leave wc.lpszClassName pointing at already-freed memory by the
    # time RegisterClassW actually reads it. Holding our own
    # create_unicode_buffer reference for the lifetime of this call
    # avoids that hazard entirely.
    class_name_buf = ctypes.create_unicode_buffer(_HOST_WNDCLASS_NAME)
    wc.lpszClassName = ctypes.cast(class_name_buf, wintypes.LPCWSTR)

    atom = _user32.RegisterClassW(ctypes.byref(wc))
    if atom == 0:
        last_err = _kernel32.GetLastError()
        # 1410 == ERROR_CLASS_ALREADY_EXISTS -- harmless, treat as success.
        if last_err != 1410:
            raise OSError(
                "RegisterClassW failed for thumbnail host (error {})".format(last_err)
            )
    _host_wndclass_atom = _HOST_WNDCLASS_NAME
    return _host_wndclass_atom


WS_CHILD = 0x40000000
WS_VISIBLE = 0x10000000


class ThumbnailSlot(object):
    """
    Owns exactly one native child HWND used as a DWM thumbnail render
    target for exactly one source window (an entry from the open
    windows list), plus the associated HTHUMBNAIL registration. Created
    lazily by ThumbnailManager and torn down explicitly via close() --
    DWM thumbnail handles are a limited system resource and MUST be
    unregistered when no longer needed, so this class is careful to
    always clean up, including on error paths.
    """

    def __init__(self, parent_hwnd, source_hwnd):
        self.parent_hwnd = parent_hwnd
        self.source_hwnd = source_hwnd
        self.host_hwnd = None
        self.thumbnail_handle = None
        self._create_host_window()
        self._register_thumbnail()

    def _create_host_window(self):
        _ensure_host_window_class_registered()
        hinstance = _kernel32.GetModuleHandleW(None)
        self.host_hwnd = _user32.CreateWindowExW(
            0,
            _HOST_WNDCLASS_NAME,
            "WLM_Thumb",
            WS_CHILD | WS_VISIBLE,
            0, 0, 1, 1,
            self.parent_hwnd,
            None,
            hinstance,
            None,
        )
        if not self.host_hwnd:
            raise OSError("Failed to create thumbnail host child window")

    def _register_thumbnail(self):
        handle_out = ctypes.c_void_p()
        hr = dwmapi.DwmRegisterThumbnail(
            wintypes.HWND(self.host_hwnd),
            wintypes.HWND(self.source_hwnd),
            ctypes.byref(handle_out),
        )
        if hr != 0 or not handle_out.value:
            # Source window may be elevated/protected, or may have closed
            # between enumeration and preview build -- non-fatal, this
            # slot simply renders nothing and the cell falls back to its
            # plain placeholder paint.
            self.thumbnail_handle = None
            return
        self.thumbnail_handle = handle_out.value

    def get_source_size(self):
        """Returns (width, height) of the live source window, or None."""
        if not self.thumbnail_handle:
            return None
        size = wintypes.SIZE()
        hr = dwmapi.DwmQueryThumbnailSourceSize(
            self.thumbnail_handle, ctypes.byref(size)
        )
        if hr != 0:
            return None
        return (size.cx, size.cy)

    def update(self, dest_rect_in_parent, visible=True, opacity=255,
               client_area_only=True):
        """
        Repositions the native host window to `dest_rect_in_parent`
        (left, top, right, bottom -- parent-client-relative coordinates)
        and tells DWM to render the live thumbnail to fill that same
        rectangle, letterboxed to preserve the source window's aspect
        ratio. Safe to call every frame/resize; DWM handles this very
        cheaply since it is simply re-pointing an existing composition
        surface.
        """
        if self.host_hwnd is None:
            return

        left, top, right, bottom = dest_rect_in_parent
        width = max(1, right - left)
        height = max(1, bottom - top)

        _user32.SetWindowPos(
            self.host_hwnd,
            None,
            int(left), int(top), int(width), int(height),
            0x0004 | 0x0010,  # SWP_NOZORDER | SWP_NOACTIVATE
        )

        if not self.thumbnail_handle:
            return

        props = _DWM_THUMBNAIL_PROPERTIES()
        props.dwFlags = (
            DWM_TNP_RECTDESTINATION
            | DWM_TNP_VISIBLE
            | DWM_TNP_OPACITY
            | DWM_TNP_SOURCECLIENTAREAONLY
        )

        # Letterbox the source window's actual aspect ratio inside the
        # destination cell instead of stretching it, so app content
        # never looks squashed.
        dest_w, dest_h = width, height
        source_size = self.get_source_size()
        if source_size and source_size[0] > 0 and source_size[1] > 0:
            src_w, src_h = source_size
            src_aspect = src_w / float(src_h)
            dest_aspect = dest_w / float(dest_h)
            if src_aspect > dest_aspect:
                fit_w = dest_w
                fit_h = max(1, int(round(dest_w / src_aspect)))
            else:
                fit_h = dest_h
                fit_w = max(1, int(round(dest_h * src_aspect)))
            off_x = (dest_w - fit_w) // 2
            off_y = (dest_h - fit_h) // 2
            rect = wintypes.RECT(off_x, off_y, off_x + fit_w, off_y + fit_h)
        else:
            rect = wintypes.RECT(0, 0, dest_w, dest_h)

        props.rcDestination = rect
        props.fVisible = bool(visible)
        props.opacity = int(opacity) & 0xFF
        props.fSourceClientAreaOnly = bool(client_area_only)

        dwmapi.DwmUpdateThumbnailProperties(self.thumbnail_handle, ctypes.byref(props))

    def close(self):
        if self.thumbnail_handle:
            try:
                dwmapi.DwmUnregisterThumbnail(self.thumbnail_handle)
            except Exception:
                pass
            self.thumbnail_handle = None
        if self.host_hwnd:
            try:
                _user32.DestroyWindow(self.host_hwnd)
            except Exception:
                pass
            self.host_hwnd = None


class ThumbnailManager(object):
    """
    Owns the full set of ThumbnailSlot objects currently active in the
    preview canvas, keyed by source HWND, and provides simple
    set/update/clear-all semantics so the canvas widget never has to
    deal with raw DWM handles directly. Designed so that calling
    sync(parent_hwnd, {hwnd: rect, ...}) is always safe to call
    repeatedly (e.g. on every selection change or resize) -- it diffs
    against what is currently registered, creating new slots, updating
    moved ones, and closing removed ones.
    """

    def __init__(self):
        self._slots = {}  # hwnd -> ThumbnailSlot

    def sync(self, parent_hwnd, hwnd_to_rect):
        """
        hwnd_to_rect: dict mapping source HWND -> (left, top, right,
        bottom) destination rectangle in the parent canvas's client
        coordinates. Any previously-registered HWND not present in this
        dict gets its thumbnail slot torn down.
        """
        wanted_hwnds = set(hwnd_to_rect.keys())
        current_hwnds = set(self._slots.keys())

        for stale_hwnd in current_hwnds - wanted_hwnds:
            self._slots[stale_hwnd].close()
            del self._slots[stale_hwnd]

        for hwnd, rect in hwnd_to_rect.items():
            if hwnd not in self._slots:
                try:
                    self._slots[hwnd] = ThumbnailSlot(parent_hwnd, hwnd)
                except Exception:
                    continue
            try:
                self._slots[hwnd].update(rect)
            except Exception:
                pass

    def clear(self):
        for slot in self._slots.values():
            slot.close()
        self._slots.clear()


# ======================================================================
# SECTION 6B: Real process icon extraction (HICON -> QPixmap)
# ======================================================================
#
# Replaces the cosmetic emoji glyphs with the application's REAL icon,
# extracted live from the native Win32 HICON the window itself reports
# (the same icon Windows shows in the taskbar / Alt-Tab for that exact
# window), falling back to an icon pulled straight from the .exe file
# on disk if the window doesn't hand one back. Every native handle is
# declared through ctypes with explicit, pointer-sized argument types
# (wintypes.HWND / wintypes.HICON / ctypes.c_void_p) -- never a fixed
# 32-bit cast -- consistent with the 64-bit safety policy noted at the
# top of this file. In particular GetClassLongPtrW (NOT the legacy,
# pointer-truncating GetClassLongW) is used to read a window class's
# icon handles.
# ======================================================================

WM_GETICON = 0x007F
ICON_SMALL = 0
ICON_BIG = 1
ICON_SMALL2 = 2
GCL_HICON = -14
GCL_HICONSM = -34
DI_NORMAL = 0x0003
_SMTO_ABORTIFHUNG = getattr(win32con, "SMTO_ABORTIFHUNG", 0x0002)

# ---- Extra ctypes signatures needed only for icon extraction --------
# (_user32 / _gdi32 / _kernel32 are already bound to their respective
# DLLs by Section 6 above; we just attach a few more entry points.)

_user32.GetDC.argtypes = [wintypes.HWND]
_user32.GetDC.restype = wintypes.HDC

_user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
_user32.ReleaseDC.restype = ctypes.c_int

_user32.DrawIconEx.argtypes = [
    wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.HICON,
    ctypes.c_int, ctypes.c_int, ctypes.c_uint, wintypes.HBRUSH, ctypes.c_uint,
]
_user32.DrawIconEx.restype = wintypes.BOOL

_user32.DestroyIcon.argtypes = [wintypes.HICON]
_user32.DestroyIcon.restype = wintypes.BOOL

try:
    _user32.GetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
    _user32.GetClassLongPtrW.restype = ctypes.c_void_p
    _HAS_GETCLASSLONGPTR = True
except AttributeError:
    # Extremely old (pre-XP-64) systems only -- GetClassLongPtrW has
    # shipped on every Windows version this app actually targets.
    _HAS_GETCLASSLONGPTR = False

_gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
_gdi32.CreateCompatibleDC.restype = wintypes.HDC

_gdi32.DeleteDC.argtypes = [wintypes.HDC]
_gdi32.DeleteDC.restype = wintypes.BOOL

_gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
_gdi32.SelectObject.restype = wintypes.HGDIOBJ

_gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
_gdi32.DeleteObject.restype = wintypes.BOOL


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class _BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", _BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


_gdi32.CreateDIBSection.argtypes = [
    wintypes.HDC, ctypes.POINTER(_BITMAPINFO), ctypes.c_uint,
    ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, wintypes.DWORD,
]
_gdi32.CreateDIBSection.restype = wintypes.HBITMAP


def _get_class_long_ptr(hwnd, index):
    """
    64-bit-safe replacement for win32gui.GetClassLong(), which wraps
    the legacy GetClassLongW and silently truncates pointer-sized
    values (like a window class's HICON) on 64-bit Windows. Uses
    GetClassLongPtrW instead, which is pointer-width-correct on both
    32-bit and 64-bit builds.
    """
    try:
        if _HAS_GETCLASSLONGPTR:
            val = _user32.GetClassLongPtrW(wintypes.HWND(hwnd), index)
            return int(val) if val else 0
    except Exception:
        pass
    try:
        return int(win32gui.GetClassLong(hwnd, index))
    except Exception:
        return 0


def _send_wm_geticon(hwnd, icon_type):
    """
    Sends WM_GETICON to `hwnd` with a short timeout (SendMessageTimeout
    rather than plain SendMessage) so a single hung/unresponsive window
    can never freeze this app's UI thread while we're just trying to
    fetch its icon for display.
    """
    try:
        result = win32gui.SendMessageTimeout(
            hwnd, WM_GETICON, icon_type, 0, _SMTO_ABORTIFHUNG, 200,
        )
        lresult = result[-1] if isinstance(result, tuple) else result
        return int(lresult) if lresult else 0
    except Exception:
        return 0


def get_hicon_for_window(hwnd):
    """
    Best-effort extraction of the native HICON that represents `hwnd`'s
    real application icon -- the exact same icon Windows itself shows
    for that window in the taskbar/Alt-Tab -- tried in order of
    fidelity: WM_GETICON(ICON_BIG) -> WM_GETICON(ICON_SMALL2) ->
    WM_GETICON(ICON_SMALL) -> the window class's large icon -> the
    window class's small icon. Returns 0 if the live window doesn't
    hand back any icon at all, in which case the caller falls back to
    pulling one straight from the owning .exe on disk. This handle is
    OWNED by the target window/class -- the caller must NEVER call
    DestroyIcon on it.
    """
    for icon_type in (ICON_BIG, ICON_SMALL2, ICON_SMALL):
        hicon = _send_wm_geticon(hwnd, icon_type)
        if hicon:
            return hicon
    hicon = _get_class_long_ptr(hwnd, GCL_HICON)
    if hicon:
        return hicon
    hicon = _get_class_long_ptr(hwnd, GCL_HICONSM)
    if hicon:
        return hicon
    return 0


def get_hicon_for_exe(exe_path):
    """
    Extracts the first icon embedded in the given .exe file on disk via
    ExtractIconEx, used as a fallback when a live window doesn't report
    any icon of its own. Unlike get_hicon_for_window(), this handle IS
    owned by the caller and MUST be destroyed (via DestroyIcon) once
    converted to a QPixmap, to avoid leaking a GDI icon handle.
    """
    if not exe_path or not os.path.isfile(exe_path):
        return 0
    try:
        large_handles, small_handles = win32gui.ExtractIconEx(exe_path, 0, 1)
    except Exception:
        return 0

    hicon = 0
    if large_handles:
        hicon = large_handles[0]
        for extra in large_handles[1:]:
            try:
                win32gui.DestroyIcon(extra)
            except Exception:
                pass
    if small_handles:
        for extra in small_handles:
            try:
                win32gui.DestroyIcon(extra)
            except Exception:
                pass
    return hicon


def hicon_to_qimage(hicon, size=32):
    """
    Converts a native HICON into a QImage with a real alpha channel by
    drawing it (DrawIconEx, DI_NORMAL) into an off-screen 32bpp
    top-down DIB section, then copying that section's raw pixel buffer
    straight into a QImage -- the standard GDI technique for extracting
    full-fidelity icon pixels (including per-pixel alpha for modern
    32-bit icons) without depending on the optional/legacy QtWinExtras
    Qt module. A 32bpp top-down BGRA DIB's byte layout is exactly what
    QImage.Format_ARGB32 expects on a little-endian machine (which is
    every machine Windows runs on), so no channel reordering is needed.
    """
    if not hicon:
        return None
    hdc_screen = None
    hdc_mem = None
    hbitmap = None
    try:
        hdc_screen = _user32.GetDC(None)
        hdc_mem = _gdi32.CreateCompatibleDC(hdc_screen)
        if not hdc_mem:
            return None

        bmi = _BITMAPINFO()
        ctypes.memset(ctypes.byref(bmi), 0, ctypes.sizeof(bmi))
        bmi.bmiHeader.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = size
        bmi.bmiHeader.biHeight = -size  # negative height == top-down DIB
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0  # BI_RGB

        bits_ptr = ctypes.c_void_p()
        hbitmap = _gdi32.CreateDIBSection(
            hdc_mem, ctypes.byref(bmi), 0, ctypes.byref(bits_ptr), None, 0
        )
        if not hbitmap or not bits_ptr.value:
            return None

        old_obj = _gdi32.SelectObject(hdc_mem, hbitmap)
        try:
            _user32.DrawIconEx(
                hdc_mem, 0, 0, hicon, size, size, 0, None, DI_NORMAL
            )
            buf_size = size * size * 4
            raw = ctypes.string_at(bits_ptr, buf_size)
        finally:
            _gdi32.SelectObject(hdc_mem, old_obj)

        image = QImage(raw, size, size, QImage.Format_ARGB32)
        return image.copy()  # detach from the raw buffer before it's freed
    except Exception:
        return None
    finally:
        if hbitmap:
            try:
                _gdi32.DeleteObject(hbitmap)
            except Exception:
                pass
        if hdc_mem:
            try:
                _gdi32.DeleteDC(hdc_mem)
            except Exception:
                pass
        if hdc_screen:
            try:
                _user32.ReleaseDC(None, hdc_screen)
            except Exception:
                pass


class IconCache(object):
    """
    Resolves and caches a real, high-resolution QPixmap of each open
    window's actual application icon, extracted live via the native
    HICON the window (or its owning .exe) reports, instead of a
    generic emoji glyph. Cached by executable path so every window
    belonging to the same .exe (e.g. two separate Chrome windows) only
    ever pays the GDI extraction cost once. Shared as a single
    module-level instance (see get_window_icon_pixmap() below) so the
    sidebar list and the center preview canvas reuse the same cache.
    """

    def __init__(self):
        self._cache = {}  # cache_key -> QPixmap (never None; a failed
                           # lookup is cached as a null QPixmap so we
                           # don't retry the expensive GDI calls every
                           # single refresh tick for icon-less windows)

    def get_pixmap(self, window_info, size=32):
        exe_path = (window_info.get("exe_path") or "").lower()
        cache_key = exe_path or "hwnd:{}".format(window_info.get("hwnd"))
        full_key = (cache_key, size)

        cached = self._cache.get(full_key)
        if cached is not None:
            return None if cached.isNull() else cached

        pixmap = self._resolve_pixmap(window_info, size)
        self._cache[full_key] = pixmap if pixmap is not None else QPixmap()
        return pixmap

    @staticmethod
    def _resolve_pixmap(window_info, size):
        hwnd = window_info.get("hwnd")
        hicon = 0
        try:
            if hwnd and win32gui.IsWindow(hwnd):
                hicon = get_hicon_for_window(hwnd)
        except Exception:
            hicon = 0

        # HICONs returned by get_hicon_for_window() are owned by the
        # target window/class and must NEVER be destroyed by us --
        # only handles WE pulled from the .exe file ourselves need
        # cleanup after conversion.
        owned_by_us = False
        if not hicon:
            hicon = get_hicon_for_exe(window_info.get("exe_path", ""))
            owned_by_us = bool(hicon)

        if not hicon:
            return None

        image = hicon_to_qimage(hicon, size)

        if owned_by_us:
            try:
                _user32.DestroyIcon(hicon)
            except Exception:
                pass

        if image is None or image.isNull():
            return None
        return QPixmap.fromImage(image)

    def invalidate(self, window_info):
        exe_path = (window_info.get("exe_path") or "").lower()
        cache_key = exe_path or "hwnd:{}".format(window_info.get("hwnd"))
        for key in [k for k in self._cache if k[0] == cache_key]:
            del self._cache[key]


# Single shared cache instance used everywhere an icon needs resolving,
# so the sidebar rows and the preview canvas's cell headers never pay
# to extract the same .exe's icon twice.
_GLOBAL_ICON_CACHE = IconCache()


def get_window_icon_pixmap(window_info, size=32):
    """Convenience wrapper around the shared IconCache instance."""
    try:
        return _GLOBAL_ICON_CACHE.get_pixmap(window_info, size)
    except Exception:
        return None


# ======================================================================
# SECTION 7: Layout persistence (Save Layout / Load Layout)
# ======================================================================

def get_layouts_directory():
    """
    Returns (creating it if necessary) the directory layouts are saved
    into: %APPDATA%\\WindowLayoutManager\\layouts on a normal Windows
    install, falling back to a local "layouts" folder next to the
    script if APPDATA isn't set (extremely unusual, but defensive).
    """
    appdata = os.environ.get("APPDATA")
    if appdata:
        base = os.path.join(appdata, "WindowLayoutManager", "layouts")
    else:
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "layouts")
    try:
        os.makedirs(base, exist_ok=True)
    except Exception:
        pass
    return base


def save_layout_to_disk(layout_name, window_entries):
    """
    Persists the CURRENT position/size of each given window to a JSON
    file named after `layout_name`. Each entry is matched at load time
    by (process_name, title) rather than by HWND, since HWNDs are only
    valid for the lifetime of the process that owns them and will not
    still refer to the same window in a future session.

    window_entries: list of dicts with keys 'hwnd', 'title',
    'process_name' (as produced by enumerate_taskbar_windows).

    Returns the full path written to.
    """
    directory = get_layouts_directory()
    safe_name = "".join(
        c for c in layout_name if c.isalnum() or c in (" ", "-", "_")
    ).strip()
    if not safe_name:
        safe_name = "Layout_{}".format(int(time.time()))
    file_path = os.path.join(directory, safe_name + ".json")

    windows_payload = []
    for entry in window_entries:
        hwnd = entry["hwnd"]
        rect = get_window_rect_safe(hwnd)
        if rect is None:
            continue
        left, top, right, bottom = rect
        windows_payload.append(
            {
                "title": entry["title"],
                "process_name": entry["process_name"],
                "x": left,
                "y": top,
                "width": right - left,
                "height": bottom - top,
            }
        )

    payload = {
        "layout_name": safe_name,
        "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "windows": windows_payload,
    }

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    return file_path


def list_saved_layouts():
    """Returns a sorted list of layout names (without the .json extension)."""
    directory = get_layouts_directory()
    try:
        names = [
            fn[:-5] for fn in os.listdir(directory)
            if fn.lower().endswith(".json")
        ]
    except Exception:
        names = []
    names.sort(key=str.lower)
    return names


def load_layout_from_disk(layout_name):
    """
    Loads a previously saved layout JSON file and returns its parsed
    payload dict, or None if the file doesn't exist / can't be parsed.
    """
    directory = get_layouts_directory()
    file_path = os.path.join(directory, layout_name + ".json")
    if not os.path.isfile(file_path):
        return None
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def apply_saved_layout(payload, currently_open_windows):
    """
    Matches each saved window entry against the currently open window
    list (by exact title + process_name match first, falling back to
    process_name-only if no exact title match is found -- handling the
    common case where a window's title has changed slightly, e.g. a
    browser tab title, since the layout was saved) and repositions any
    matches to their saved rectangle.

    Returns (applied_count, skipped_count).
    """
    applied = 0
    skipped = 0

    # Build lookup structures from the live window list.
    by_title_and_proc = {}
    by_proc_only = {}
    for w in currently_open_windows:
        key_exact = (w["title"], w["process_name"])
        by_title_and_proc.setdefault(key_exact, []).append(w)
        by_proc_only.setdefault(w["process_name"], []).append(w)

    used_hwnds = set()

    for saved_window in payload.get("windows", []):
        title = saved_window.get("title", "")
        process_name = saved_window.get("process_name", "")
        x = saved_window.get("x")
        y = saved_window.get("y")
        width = saved_window.get("width")
        height = saved_window.get("height")
        if None in (x, y, width, height):
            skipped += 1
            continue

        candidates = by_title_and_proc.get((title, process_name), [])
        candidates = [c for c in candidates if c["hwnd"] not in used_hwnds]

        if not candidates:
            candidates = by_proc_only.get(process_name, [])
            candidates = [c for c in candidates if c["hwnd"] not in used_hwnds]

        if not candidates:
            skipped += 1
            continue

        target = candidates[0]
        try:
            restore_if_minimized_or_maximized(target["hwnd"])
            apply_window_rect(target["hwnd"], x, y, width, height)
            used_hwnds.add(target["hwnd"])
            applied += 1
        except Exception:
            skipped += 1

    return applied, skipped


# ======================================================================
# SECTION 8: Dark theme stylesheet (matches the reference design)
# ======================================================================

COLOR_BG_APP = "#0f1419"
COLOR_BG_PANEL = "#161b22"
COLOR_BG_PANEL_ALT = "#1c2228"
COLOR_BG_ROW = "#1a2029"
COLOR_BG_ROW_HOVER = "#222a35"
COLOR_BORDER = "#2a3240"
COLOR_BORDER_LIGHT = "#363f4d"
COLOR_TEXT_PRIMARY = "#e6e9ef"
COLOR_TEXT_SECONDARY = "#8b94a3"
COLOR_TEXT_MUTED = "#5e6773"
COLOR_ACCENT = "#3b82f6"
COLOR_ACCENT_HOVER = "#5a93f7"
COLOR_ACCENT_DIM = "#1d3a63"
COLOR_GREEN = "#22c55e"
COLOR_ORANGE = "#f59e0b"
COLOR_RED = "#ef4444"
COLOR_RED_HOVER = "#dc4444"

APP_STYLESHEET = """
QMainWindow {{
    background-color: {bg_app};
}}
QWidget {{
    color: {text_primary};
    font-family: "Segoe UI", "Segoe UI Variable", Arial, sans-serif;
    font-size: 13px;
}}
QFrame#TopBar {{
    background-color: {bg_panel};
    border-bottom: 1px solid {border};
}}
QFrame#LeftPanel {{
    background-color: {bg_panel};
    border-right: 1px solid {border};
}}
QFrame#RightPanel {{
    background-color: {bg_panel};
    border-left: 1px solid {border};
}}
QFrame#CenterPanel {{
    background-color: {bg_app};
}}
QFrame#CanvasFrame {{
    background-color: {bg_panel_alt};
    border: 1px solid {border};
    border-radius: 10px;
}}
QFrame#SectionCard {{
    background-color: {bg_panel_alt};
    border: 1px solid {border};
    border-radius: 10px;
}}
QLabel#TitleLabel {{
    font-size: 16px;
    font-weight: 600;
    color: {text_primary};
}}
QLabel#SectionHeader {{
    font-size: 13px;
    font-weight: 600;
    color: {text_primary};
}}
QLabel#SubText {{
    color: {text_secondary};
    font-size: 12px;
}}
QLabel#MutedText {{
    color: {text_muted};
    font-size: 11px;
}}
QLabel#WarningText {{
    color: #e8a317;
    font-size: 11px;
    font-weight: 600;
}}
QLabel#CountBadge {{
    background-color: {accent_dim};
    color: {accent_hover};
    border-radius: 9px;
    padding: 1px 9px;
    font-weight: 600;
    font-size: 12px;
}}
QLineEdit#SearchBox {{
    background-color: {bg_row};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 7px 10px;
    color: {text_primary};
    font-size: 13px;
}}
QLineEdit#SearchBox:focus {{
    border: 1px solid {accent};
}}
QListWidget#WindowList {{
    background-color: transparent;
    border: none;
    outline: none;
}}
QListWidget#WindowList::item {{
    background-color: transparent;
    border: none;
    padding: 0px;
    margin: 0px;
}}
QListWidget#WindowList::item:hover {{
    background-color: transparent;
}}
QListWidget#WindowList::item:selected {{
    background-color: transparent;
}}
QFrame#WindowRowCard {{
    background-color: {bg_row};
    border: 1px solid {border};
    border-radius: 8px;
}}
QFrame#WindowRowCard:hover {{
    background-color: {bg_row_hover};
    border: 1px solid {border_light};
}}
QFrame#WindowRowCard[checked="true"] {{
    background-color: {bg_row_hover};
    border: 1px solid {accent};
}}
QPushButton {{
    background-color: {bg_row};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 8px 14px;
    color: {text_primary};
    font-size: 13px;
}}
QPushButton:hover {{
    background-color: {bg_row_hover};
    border: 1px solid {border_light};
}}
QPushButton:pressed {{
    background-color: {bg_panel};
}}
QPushButton:disabled {{
    color: {text_muted};
    background-color: {bg_row};
    border: 1px solid {border};
}}
QPushButton#PrimaryButton {{
    background-color: {accent};
    border: 1px solid {accent};
    color: white;
    font-weight: 600;
    padding: 10px 14px;
    border-radius: 8px;
}}
QPushButton#PrimaryButton:hover {{
    background-color: {accent_hover};
    border: 1px solid {accent_hover};
}}
QPushButton#DangerButton {{
    background-color: transparent;
    border: 1px solid #4a2a2a;
    color: {red};
    padding: 10px 14px;
    border-radius: 8px;
}}
QPushButton#DangerButton:hover {{
    background-color: #2a1818;
    border: 1px solid {red_hover};
}}
QPushButton#TopBarButton {{
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 6px 12px;
    color: {text_secondary};
}}
QPushButton#TopBarButton:hover {{
    background-color: {bg_row_hover};
    color: {text_primary};
}}
QToolButton#MonitorTab {{
    background-color: {bg_row};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 6px 14px;
    color: {text_secondary};
    text-align: left;
}}
QToolButton#MonitorTab:hover {{
    border: 1px solid {border_light};
}}
QToolButton#MonitorTab:checked {{
    background-color: {accent_dim};
    border: 1px solid {accent};
    color: {text_primary};
}}
QToolButton#QuickLayoutButton {{
    background-color: {bg_row};
    border: 1px solid {border};
    border-radius: 8px;
}}
QToolButton#QuickLayoutButton:hover {{
    border: 1px solid {border_light};
}}
QToolButton#QuickLayoutButton:checked {{
    background-color: {accent_dim};
    border: 1px solid {accent};
}}
QComboBox {{
    background-color: {bg_row};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 7px 10px;
    color: {text_primary};
}}
QComboBox:hover {{
    border: 1px solid {border_light};
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox QAbstractItemView {{
    background-color: {bg_panel_alt};
    border: 1px solid {border};
    color: {text_primary};
    selection-background-color: {accent_dim};
    outline: none;
}}
QStatusBar {{
    background-color: {bg_panel};
    color: {text_secondary};
    border-top: 1px solid {border};
}}
QScrollBar:vertical {{
    background: transparent;
    width: 8px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {border_light};
    border-radius: 4px;
    min-height: 20px;
}}
QScrollBar::handle:vertical:hover {{
    background: {text_muted};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent;
}}
QScrollArea {{
    background-color: transparent;
    border: none;
}}
QCheckBox {{
    spacing: 8px;
}}
"""

def build_stylesheet():
    return APP_STYLESHEET.format(
        bg_app=COLOR_BG_APP,
        bg_panel=COLOR_BG_PANEL,
        bg_panel_alt=COLOR_BG_PANEL_ALT,
        bg_row=COLOR_BG_ROW,
        bg_row_hover=COLOR_BG_ROW_HOVER,
        border=COLOR_BORDER,
        border_light=COLOR_BORDER_LIGHT,
        text_primary=COLOR_TEXT_PRIMARY,
        text_secondary=COLOR_TEXT_SECONDARY,
        text_muted=COLOR_TEXT_MUTED,
        accent=COLOR_ACCENT,
        accent_hover=COLOR_ACCENT_HOVER,
        accent_dim=COLOR_ACCENT_DIM,
        green=COLOR_GREEN,
        orange=COLOR_ORANGE,
        red=COLOR_RED,
        red_hover=COLOR_RED_HOVER,
    )


# ======================================================================
# SECTION 9: Custom list row widget for each open window
# ======================================================================

class WindowListRow(QWidget):
    """
    One row in the "Open Windows" list: a checkbox, the application's
    REAL extracted icon (falling back to an emoji glyph only if no
    native icon could be resolved), the window title, the executable
    name as a subtitle, and a small colored status dot (green = normal
    / restored, orange = minimized). Each instance is bound to exactly
    one HWND for its entire lifetime and never represents more than one
    window, which is what guarantees that checking/unchecking or tiling
    this row can only ever affect that single HWND.

    LAYOUT-COLLAPSE FIX: every row has an explicit, strict, FIXED pixel
    height (ROW_HEIGHT). All of the row's visible chrome -- background,
    border, rounded corners, hover/checked highlight -- is drawn by an
    inner QFrame ("card") that this widget owns and sizes itself,
    rather than via QListWidget's ::item style rules. This matters
    because when a QListWidgetItem has a custom widget attached via
    setItemWidget(), any padding/border/margin a stylesheet applies to
    ::item is carved OUT of the item's already-allocated size-hint
    rectangle before our widget is placed inside it -- which is exactly
    what was squashing the checkbox/icon/text below their natural size
    and causing adjacent rows to visually overlap. By keeping ::item
    completely bare (no padding, no border, no margin -- see
    build_stylesheet()) and overriding sizeHint() below to always
    return exactly ROW_HEIGHT regardless of font metrics or window
    title length, the row widget is guaranteed the full pixel rectangle
    it asks for, every time, for every title length.
    """

    # NOTE: these MUST NOT use the plain `int` PyQt argument type. PyQt5
    # marshals a Signal(int, ...) argument as a C++ 32-bit int, but
    # HWNDs are 64-bit pointers on 64-bit Windows -- any HWND value above
    # 0x7FFFFFFF (entirely common; HWNDs are not small sequential numbers)
    # would overflow that 32-bit slot during Qt's signal/slot argument
    # marshaling and crash the process (a hard native abort, not a
    # catchable Python exception) the moment such a window's checkbox or
    # pin button was clicked. Using `object` passes the raw Python int
    # through unboxed, with no width limit, consistent with the 64-bit
    # safety policy documented at the top of this file.
    toggled = Signal(object, bool)  # (hwnd, checked)
    pin_toggled = Signal(object, bool)  # (hwnd, pinned) -- Feature 5

    ROW_HEIGHT = 50
    ICON_SIZE = 30

    def __init__(self, window_info, parent=None):
        super().__init__(parent)
        self.hwnd = window_info["hwnd"]
        self.window_info = window_info

        self.setFixedHeight(self.ROW_HEIGHT)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        outer = QVBoxLayout(self)
        # Small top/bottom inset bakes a visual gap between cards
        # directly into this widget's own fixed height, independent of
        # whatever spacing the QListWidget itself adds between items.
        outer.setContentsMargins(0, 3, 0, 3)
        outer.setSpacing(0)

        self.card = QFrame()
        self.card.setObjectName("WindowRowCard")
        self.card.setProperty("checked", "false")
        outer.addWidget(self.card)

        layout = QHBoxLayout(self.card)
        layout.setContentsMargins(8, 0, 10, 0)
        layout.setSpacing(10)

        self.checkbox = QPushButton()
        self.checkbox.setCheckable(True)
        self.checkbox.setFixedSize(20, 20)
        self.checkbox.setCursor(Qt.PointingHandCursor)
        self.checkbox.setStyleSheet(self._checkbox_style(False))
        self.checkbox.clicked.connect(self._on_checkbox_clicked)
        layout.addWidget(self.checkbox, 0, Qt.AlignVCenter)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(self.ICON_SIZE, self.ICON_SIZE)
        self.icon_label.setAlignment(Qt.AlignCenter)
        self.icon_label.setStyleSheet(
            "background-color: {}; border-radius: 6px;".format(COLOR_BG_PANEL_ALT)
        )
        self._apply_icon(window_info)
        layout.addWidget(self.icon_label, 0, Qt.AlignVCenter)

        text_col = QVBoxLayout()
        text_col.setSpacing(1)
        text_col.setContentsMargins(0, 0, 0, 0)

        title_text = window_info["title"]
        if len(title_text) > 42:
            title_text = title_text[:39] + "..."
        self.title_label = QLabel(title_text)
        self.title_label.setStyleSheet(
            "color: {}; font-size: 12.5px; font-weight: 500; background: transparent;".format(
                COLOR_TEXT_PRIMARY
            )
        )
        self.title_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.title_label.setMinimumWidth(0)
        text_col.addWidget(self.title_label)

        subtitle = window_info["process_name"] or "unknown.exe"
        self.subtitle_label = QLabel(subtitle)
        self.subtitle_label.setStyleSheet(
            "color: {}; font-size: 11px; background: transparent;".format(COLOR_TEXT_MUTED)
        )
        self.subtitle_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.subtitle_label.setMinimumWidth(0)
        text_col.addWidget(self.subtitle_label)

        layout.addLayout(text_col, stretch=1)

        # FEATURE 5: per-row "Always on Top" pin toggle. Reflects (and
        # immediately edits) the window's REAL Win32 WS_EX_TOPMOST
        # state -- this is genuine OS-level pin management from the
        # sidebar, not just a cosmetic UI flag.
        self.pin_button = QToolButton()
        self.pin_button.setCheckable(True)
        self.pin_button.setCursor(Qt.PointingHandCursor)
        self.pin_button.setFixedSize(22, 22)
        self.pin_button.setText("📌")
        self.pin_button.setToolTip("Keep this window Always on Top")
        try:
            self.pin_button.setChecked(is_window_topmost(self.hwnd))
        except Exception:
            pass
        self._apply_pin_style(self.pin_button.isChecked())
        self.pin_button.clicked.connect(self._on_pin_clicked)
        layout.addWidget(self.pin_button, 0, Qt.AlignVCenter)

        self.status_dot = QLabel()
        self.status_dot.setFixedSize(8, 8)
        is_minimized = window_info.get("is_minimized", False)
        dot_color = COLOR_ORANGE if is_minimized else COLOR_GREEN
        self.status_dot.setStyleSheet(
            "background-color: {}; border-radius: 4px;".format(dot_color)
        )
        layout.addWidget(self.status_dot, 0, Qt.AlignVCenter)

        self.setLayout(outer)

    def sizeHint(self):
        # Always exactly ROW_HEIGHT tall -- see class docstring. Width
        # follows the current widget width so QListWidget can still
        # stretch rows to fill the viewport horizontally.
        width = self.width() if self.width() > 0 else 280
        return QSize(width, self.ROW_HEIGHT)

    def minimumSizeHint(self):
        return QSize(120, self.ROW_HEIGHT)

    def _apply_icon(self, window_info):
        pixmap = get_window_icon_pixmap(window_info, size=32)
        if pixmap is not None and not pixmap.isNull():
            scaled = pixmap.scaled(
                self.ICON_SIZE - 4, self.ICON_SIZE - 4,
                Qt.KeepAspectRatio, Qt.SmoothTransformation,
            )
            self.icon_label.setPixmap(scaled)
            self.icon_label.setText("")
        else:
            # No real icon could be resolved (protected process, etc.)
            # -- fall back to the cosmetic emoji glyph so the row never
            # has a blank/empty icon slot.
            emoji = get_emoji_for_process(window_info["process_name"])
            self.icon_label.setPixmap(QPixmap())
            self.icon_label.setText(emoji)
            self.icon_label.setStyleSheet(
                "font-size: 16px; background-color: {}; border-radius: 6px;".format(
                    COLOR_BG_PANEL_ALT
                )
            )

    @staticmethod
    def _checkbox_style(checked):
        if checked:
            return (
                "QPushButton {{ background-color: {accent}; border: 1px solid {accent}; "
                "border-radius: 5px; }}"
            ).format(accent=COLOR_ACCENT)
        return (
            "QPushButton {{ background-color: transparent; border: 1.5px solid {border}; "
            "border-radius: 5px; }}"
        ).format(border=COLOR_BORDER_LIGHT)

    @staticmethod
    def _pin_style(pinned):
        if pinned:
            return (
                "QToolButton {{ background-color: {dim}; border: 1px solid {accent}; "
                "border-radius: 5px; font-size: 11px; color: {accent_hover}; }}"
            ).format(
                dim=COLOR_ACCENT_DIM, accent=COLOR_ACCENT, accent_hover=COLOR_ACCENT_HOVER
            )
        return (
            "QToolButton {{ background-color: transparent; border: 1px solid transparent; "
            "border-radius: 5px; font-size: 11px; color: {muted}; }}"
        ).format(muted=COLOR_TEXT_MUTED)

    def _apply_pin_style(self, pinned):
        self.pin_button.setStyleSheet(self._pin_style(pinned))
        self.pin_button.setToolTip(
            "Pinned Always on Top -- click to unpin"
            if pinned else "Keep this window Always on Top"
        )

    def _on_pin_clicked(self):
        # FEATURE 5: clicking the pin immediately edits the real Win32
        # topmost state of exactly this one HWND -- nothing else on
        # screen is affected, and the change takes effect instantly,
        # with no need to apply or re-tile any layout.
        want_pinned = self.pin_button.isChecked()
        set_window_topmost(self.hwnd, want_pinned)
        self._apply_pin_style(want_pinned)
        self.pin_toggled.emit(self.hwnd, want_pinned)

    def _refresh_card_checked_state(self, checked):
        self.card.setProperty("checked", "true" if checked else "false")
        # Force the stylesheet engine to re-evaluate the [checked="..."]
        # selector now that the dynamic property changed.
        self.card.style().unpolish(self.card)
        self.card.style().polish(self.card)
        self.card.update()

    def _on_checkbox_clicked(self):
        checked = self.checkbox.isChecked()
        self.checkbox.setStyleSheet(self._checkbox_style(checked))
        self._refresh_card_checked_state(checked)
        self.toggled.emit(self.hwnd, checked)

    def set_checked(self, checked):
        self.checkbox.setChecked(checked)
        self.checkbox.setStyleSheet(self._checkbox_style(checked))
        self._refresh_card_checked_state(checked)

    def is_checked(self):
        return self.checkbox.isChecked()


# ======================================================================
# SECTION 10: Layout Preview Canvas (numbered grid + live thumbnails)
# ======================================================================

class LayoutPreviewCanvas(QFrame):
    """
    The center "Layout Preview" panel. Given the currently checked
    window list and the currently selected monitor, draws a scaled-down
    representation of that monitor's working area subdivided into the
    numbered grid cells the active template will produce, and hosts
    one live DWM thumbnail per cell showing that window's real,
    currently-on-screen content. Purely a preview -- nothing here moves
    any real application window; only "Apply Equal Layout" actually
    does that.

    Two new interactive behaviors live here:

    FEATURE 1 (drag & drop reorder): clicking and dragging the BODY of
    a numbered cell onto another cell instantly swaps which checked
    window occupies which slot, both in this preview and in the order
    Apply Equal Layout will use.

    FEATURE 3 (custom asymmetric templates): when the active template
    is "master_stack" or "columns" (see set_template_mode()), thin
    divider lines are drawn between adjacent slots and are themselves
    draggable, letting the user manually resize the master/stack split
    or any column width before applying. The exact same ratio-driven
    math used to draw the preview (get_monitor_space_tile_rects) is
    also what Apply Equal Layout calls to compute the real, final
    window rectangles -- there is only ever one source of truth.
    """

    order_changed = Signal(list)  # emits the new [hwnd, ...] order after a drag swap

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("CanvasFrame")
        self.setMinimumHeight(420)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)

        self._monitor = None
        self._checked_windows = []  # list of window_info dicts, in display order
        self._cell_rects_widget = []  # QRect list, widget-local coordinates
        self._thumbnail_manager = ThumbnailManager()
        self._own_hwnd = None

        # --- Feature 3: custom template state -----------------------
        self._template_mode = "equal"  # "equal"|"master_stack"|"columns"|"rows"|"custom_grid"
        self._master_ratio = 0.64
        self._stack_ratios = []   # length n-1, normalized, master_stack mode
        self._column_ratios = []  # length n, normalized, columns mode
        self._row_ratios = []     # length n, normalized, rows mode

        # Custom Grid template: user-specified row × column counts.
        # Defaults to 2×3 (6 cells) which works well for many Telegram
        # multi-account setups without going below the app's min size.
        self._grid_rows = 2
        self._grid_cols = 3

        # --- Inter-window gap (spacing), 0..MAX_GAP_PX, shared by all
        # templates -- see apply_gap_to_rects(). Defaults to 0 so the
        # original edge-to-edge behavior is unchanged until the user
        # moves the Gap slider in the right-hand panel.
        self._gap_px = 0

        # --- Feature 1 / 3: live drag-interaction state --------------
        self._drag_mode = None       # None | "reorder" | "divider"
        self._drag_source_index = -1
        self._drag_hover_index = -1
        self._divider_info = None

        # Cached widget-space transform from the last repaint, used to
        # convert mouse positions back into ratios while dragging.
        self._scale = 1.0
        self._offset_x = 0
        self._offset_y = 0
        self._fit_w = 0
        self._fit_h = 0

        self._header_label = QLabel("Layout Preview")
        self._header_label.setObjectName("SectionHeader")
        self._dims_label = QLabel("")
        self._dims_label.setObjectName("MutedText")

        header_row = QHBoxLayout()
        header_row.addWidget(self._header_label)
        header_row.addStretch(1)
        header_row.addWidget(self._dims_label)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 14, 16, 16)
        outer.addLayout(header_row)
        outer.addSpacing(8)
        outer.addStretch(1)
        self.setLayout(outer)

    def set_own_hwnd(self, hwnd):
        self._own_hwnd = hwnd

    def set_monitor(self, monitor_info):
        self._monitor = monitor_info
        if monitor_info is not None:
            self._dims_label.setText(monitor_info.resolution_label)
        self._recompute_and_repaint()

    def set_checked_windows(self, window_info_list):
        self._checked_windows = list(window_info_list)
        self._ensure_ratio_defaults()
        self._drag_mode = None
        self._drag_source_index = -1
        self._drag_hover_index = -1
        self._divider_info = None
        self._recompute_and_repaint()

    def set_template_mode(self, mode):
        """FEATURE 3: switches between Equal Grid / Master+Stack / Columns / Vertical Stack / Custom Grid."""
        if mode not in ("equal", "master_stack", "columns", "rows", "custom_grid"):
            mode = "equal"
        self._template_mode = mode
        self._drag_mode = None
        self._divider_info = None
        self._ensure_ratio_defaults()
        self._recompute_and_repaint()

    def set_custom_grid_dims(self, rows, cols):
        """Updates the Custom Grid row×col counts and repaints immediately."""
        self._grid_rows = max(1, int(rows))
        self._grid_cols = max(1, int(cols))
        if self._template_mode == "custom_grid":
            self._recompute_and_repaint()

    def get_template_mode(self):
        return self._template_mode

    def set_gap_px(self, gap_px):
        """Updates the shared inter-window gap and repaints immediately."""
        self._gap_px = max(0, min(MAX_GAP_PX, int(gap_px)))
        self._recompute_and_repaint()

    def get_monitor_space_tile_rects(self, work_area, n):
        """
        Returns the real, final (x, y, w, h) rectangles -- in monitor
        work-area pixel space -- for the currently active template,
        using whatever manual divider ratios the user has dragged so
        far, with the current gap setting (see apply_gap_to_rects())
        applied as the very last step. This is the SAME math used to
        draw the scaled-down preview, so what the user sees in the
        canvas is always exactly what Apply Equal Layout will produce.
        """
        if n <= 0:
            return []
        if self._template_mode == "master_stack":
            rects = compute_master_stack_rects(
                work_area, n, self._master_ratio, self._stack_ratios
            )
        elif self._template_mode == "columns":
            rects = compute_columns_rects(work_area, n, self._column_ratios)
        elif self._template_mode == "rows":
            rects = compute_rows_rects(work_area, n, self._row_ratios)
        elif self._template_mode == "custom_grid":
            rects = compute_custom_grid_rects(
                work_area, n, self._grid_rows, self._grid_cols
            )
        else:
            rects = compute_tile_rects(work_area, n)
        return apply_gap_to_rects(rects, self._gap_px)

    def _ensure_ratio_defaults(self):
        """
        (Re)initializes the custom-template ratio lists to a perfectly
        equal split whenever the checked window count changes -- e.g.
        the first time N windows are seen in a given template, or right
        after a selection change -- without disturbing ratios the user
        has already manually dragged for the CURRENT window count.
        """
        n = len(self._checked_windows)
        stack_count = max(0, n - 1)
        if len(self._stack_ratios) != stack_count:
            self._stack_ratios = (
                [1.0 / stack_count] * stack_count if stack_count else []
            )
        if len(self._column_ratios) != n:
            self._column_ratios = [1.0 / n] * n if n else []
        if len(self._row_ratios) != n:
            self._row_ratios = [1.0 / n] * n if n else []

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._recompute_and_repaint()

    def showEvent(self, event):
        super().showEvent(event)
        self._recompute_and_repaint()

    # ------------------------------------------------------------------
    def _canvas_drawing_rect(self):
        """
        The pixel rectangle, in this widget's own local coordinates,
        available for drawing the scaled monitor preview -- i.e. our
        full client area minus the header row and margins.
        """
        margin = 16
        top_offset = 46
        return QRect(
            margin,
            top_offset,
            max(10, self.width() - margin * 2),
            max(10, self.height() - top_offset - margin),
        )

    def _recompute_and_repaint(self):
        self._cell_rects_widget = []

        if self._monitor is None or not self._checked_windows:
            self._thumbnail_manager.clear()
            self.update()
            return

        n = len(self._checked_windows)
        work_left, work_top, work_right, work_bottom = self._monitor.work_rect
        mon_w = work_right - work_left
        mon_h = work_bottom - work_top
        if mon_w <= 0 or mon_h <= 0:
            self._thumbnail_manager.clear()
            self.update()
            return

        draw_rect = self._canvas_drawing_rect()

        # Fit the monitor's aspect ratio inside draw_rect, centered.
        mon_aspect = mon_w / float(mon_h)
        draw_aspect = draw_rect.width() / float(draw_rect.height())
        if mon_aspect > draw_aspect:
            fit_w = draw_rect.width()
            fit_h = max(1, int(round(fit_w / mon_aspect)))
        else:
            fit_h = draw_rect.height()
            fit_w = max(1, int(round(fit_h * mon_aspect)))

        offset_x = draw_rect.left() + (draw_rect.width() - fit_w) // 2
        offset_y = draw_rect.top() + (draw_rect.height() - fit_h) // 2

        scale = fit_w / float(mon_w)

        # Cache the transform so mouse-driven divider dragging can map
        # pointer positions back into monitor-space ratios.
        self._scale = scale
        self._offset_x = offset_x
        self._offset_y = offset_y
        self._fit_w = fit_w
        self._fit_h = fit_h

        # Compute the active template's cells against the monitor's
        # actual work-area pixel space, then scale those rectangles
        # down into widget-local preview coordinates. This is the exact
        # same function Apply Equal Layout calls against the REAL
        # monitor work area, so the preview can never drift from reality.
        tile_rects = self.get_monitor_space_tile_rects((0, 0, mon_w, mon_h), n)

        for (rx, ry, rw, rh) in tile_rects:
            wx = offset_x + int(round(rx * scale))
            wy = offset_y + int(round(ry * scale))
            ww = max(1, int(round(rw * scale)))
            wh = max(1, int(round(rh * scale)))
            self._cell_rects_widget.append(QRect(wx, wy, ww, wh))

        self.update()
        self._sync_thumbnails()

    def _sync_thumbnails(self):
        if self._own_hwnd is None:
            return
        # The DWM thumbnail host child windows are parented to the
        # top-level window's HWND (self._own_hwnd), not to this canvas
        # widget's own (non-existent, since QWidget children are not
        # separate native HWNDs by default) handle -- so every
        # destination rectangle must be translated out of this widget's
        # own local coordinate space and into the top-level window's
        # client coordinate space before being handed to ThumbnailSlot,
        # or every thumbnail renders offset by however far this canvas
        # sits from the top-level window's top-left corner (i.e. it
        # would appear shifted into the left/right panels).
        top_level = self.window()
        origin_in_top_level = self.mapTo(top_level, self.rect().topLeft())
        dx, dy = origin_in_top_level.x(), origin_in_top_level.y()

        hwnd_to_rect = {}
        inset = 3  # keep the live thumbnail slightly inside the drawn border
        for window_info, rect in zip(self._checked_windows, self._cell_rects_widget):
            content_rect = rect.adjusted(inset, inset + 22, -inset, -inset)
            if content_rect.width() < 4 or content_rect.height() < 4:
                continue
            hwnd_to_rect[window_info["hwnd"]] = (
                content_rect.left() + dx,
                content_rect.top() + dy,
                content_rect.right() + dx,
                content_rect.bottom() + dy,
            )
        self._thumbnail_manager.sync(self._own_hwnd, hwnd_to_rect)

    def clear_thumbnails(self):
        self._thumbnail_manager.clear()

    # ------------------------------------------------------------------
    # FEATURE 1 / FEATURE 3: mouse-driven drag interaction
    # ------------------------------------------------------------------
    def _cell_index_at_pos(self, pos):
        for idx, rect in enumerate(self._cell_rects_widget):
            if rect.contains(pos):
                return idx
        return -1

    def _build_divider_segments(self):
        """
        Returns the draggable divider lines for the active custom
        template, derived directly from the already-computed widget-
        space cell rectangles so the drawn dividers and the cell
        borders can never visually disagree. Each segment carries both
        a thin `hit_rect` (used for both hover-cursor and drag hit
        testing) and a `p1`/`p2` pair (used purely for drawing the line).
        """
        segments = []
        if self._template_mode == "master_stack" and len(self._cell_rects_widget) >= 2:
            master_rect = self._cell_rects_widget[0]
            stack_rects = self._cell_rects_widget[1:]
            x = master_rect.right()
            segments.append({
                "kind": "master", "orientation": "v",
                "hit_rect": QRect(x - 4, master_rect.top(), 8, master_rect.height()),
                "p1": (x, master_rect.top()), "p2": (x, master_rect.bottom()),
            })
            for i in range(len(stack_rects) - 1):
                top_rect = stack_rects[i]
                y = top_rect.bottom()
                segments.append({
                    "kind": "stack", "index": i, "orientation": "h",
                    "hit_rect": QRect(top_rect.left(), y - 4, top_rect.width(), 8),
                    "p1": (top_rect.left(), y), "p2": (top_rect.right(), y),
                })
        elif self._template_mode == "columns" and len(self._cell_rects_widget) >= 2:
            for i in range(len(self._cell_rects_widget) - 1):
                left_rect = self._cell_rects_widget[i]
                x = left_rect.right()
                segments.append({
                    "kind": "column", "index": i, "orientation": "v",
                    "hit_rect": QRect(x - 4, left_rect.top(), 8, left_rect.height()),
                    "p1": (x, left_rect.top()), "p2": (x, left_rect.bottom()),
                })
        elif self._template_mode == "rows" and len(self._cell_rects_widget) >= 2:
            for i in range(len(self._cell_rects_widget) - 1):
                top_rect = self._cell_rects_widget[i]
                y = top_rect.bottom()
                segments.append({
                    "kind": "row", "index": i, "orientation": "h",
                    "hit_rect": QRect(top_rect.left(), y - 4, top_rect.width(), 8),
                    "p1": (top_rect.left(), y), "p2": (top_rect.right(), y),
                })
        return segments

    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton or self._monitor is None or not self._checked_windows:
            return super().mousePressEvent(event)

        pos = event.pos()

        # Dividers take priority over cell-body dragging so resizing a
        # custom template is always possible even though its thin hit
        # zone slightly overlaps the edge of the adjacent cell.
        for seg in self._build_divider_segments():
            if seg["hit_rect"].contains(pos):
                self._drag_mode = "divider"
                self._divider_info = seg
                self.setCursor(
                    Qt.SizeHorCursor if seg["orientation"] == "v" else Qt.SizeVerCursor
                )
                return

        idx = self._cell_index_at_pos(pos)
        if idx != -1:
            self._drag_mode = "reorder"
            self._drag_source_index = idx
            self._drag_hover_index = idx
            self.setCursor(Qt.ClosedHandCursor)
            self.update()
            return

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        pos = event.pos()

        if self._drag_mode == "divider" and self._divider_info is not None:
            self._update_divider_drag(pos)
            return

        if self._drag_mode == "reorder":
            self._drag_hover_index = self._cell_index_at_pos(pos)
            self.update()
            return

        # Not currently dragging anything -- update the hover cursor so
        # a divider is visually discoverable before the user commits.
        segments = self._build_divider_segments()
        hovered = next((s for s in segments if s["hit_rect"].contains(pos)), None)
        if hovered is not None:
            self.setCursor(
                Qt.SizeHorCursor if hovered["orientation"] == "v" else Qt.SizeVerCursor
            )
        elif self._cell_index_at_pos(pos) != -1:
            self.setCursor(Qt.OpenHandCursor)
        else:
            self.setCursor(Qt.ArrowCursor)

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._drag_mode == "reorder" and self._drag_source_index != -1:
            target_idx = self._cell_index_at_pos(event.pos())
            if target_idx != -1 and target_idx != self._drag_source_index:
                # FEATURE 1: swap the two windows' positions -- this is
                # the single authoritative reorder operation; Apply
                # Equal Layout always tiles in this exact resulting
                # order afterwards.
                src, dst = self._drag_source_index, target_idx
                self._checked_windows[src], self._checked_windows[dst] = (
                    self._checked_windows[dst], self._checked_windows[src]
                )
                self.order_changed.emit([w["hwnd"] for w in self._checked_windows])
                self._recompute_and_repaint()

        self._drag_mode = None
        self._drag_source_index = -1
        self._drag_hover_index = -1
        self._divider_info = None
        self.setCursor(Qt.ArrowCursor)
        self.update()
        super().mouseReleaseEvent(event)

    def _update_divider_drag(self, pos):
        """
        FEATURE 3: live divider resize. Recomputes ratios from the
        divider's new pixel position while holding every OTHER
        boundary's pixel position fixed, so dragging one divider only
        ever changes the two slots that actually touch it.
        """
        seg = self._divider_info

        if seg["kind"] == "master":
            if self._fit_w:
                local_x = pos.x() - self._offset_x
                ratio = local_x / float(self._fit_w)
                self._master_ratio = max(
                    MIN_RATIO_FRACTION, min(1.0 - MIN_RATIO_FRACTION, ratio)
                )
                self._recompute_and_repaint()

        elif seg["kind"] == "stack":
            stack_rects = self._cell_rects_widget[1:]
            if not stack_rects:
                return
            top_of_stack = stack_rects[0].top()
            bottom_of_stack = stack_rects[-1].bottom()
            total_px = max(1, bottom_of_stack - top_of_stack)

            i = seg["index"]
            boundaries = [top_of_stack]
            cursor = top_of_stack
            for r in self._stack_ratios:
                cursor += r * total_px
                boundaries.append(cursor)
            boundaries[-1] = bottom_of_stack

            min_px = max(MIN_SLOT_PX, min_ratio_for_count(len(self._stack_ratios)) * total_px)
            local_y = max(top_of_stack, min(bottom_of_stack, pos.y()))
            new_y = max(boundaries[i] + min_px, min(boundaries[i + 2] - min_px, local_y))
            boundaries[i + 1] = new_y

            new_ratios = [
                (boundaries[k + 1] - boundaries[k]) / float(total_px)
                for k in range(len(boundaries) - 1)
            ]
            self._stack_ratios = normalize_ratios(new_ratios, len(self._stack_ratios))
            self._recompute_and_repaint()

        elif seg["kind"] == "column":
            if not self._cell_rects_widget:
                return
            left_of_all = self._cell_rects_widget[0].left()
            right_of_all = self._cell_rects_widget[-1].right()
            total_px = max(1, right_of_all - left_of_all)

            i = seg["index"]
            boundaries = [left_of_all]
            cursor = left_of_all
            for r in self._column_ratios:
                cursor += r * total_px
                boundaries.append(cursor)
            boundaries[-1] = right_of_all

            min_px = max(MIN_SLOT_PX, min_ratio_for_count(len(self._column_ratios)) * total_px)
            local_x = max(left_of_all, min(right_of_all, pos.x()))
            new_x = max(boundaries[i] + min_px, min(boundaries[i + 2] - min_px, local_x))
            boundaries[i + 1] = new_x

            new_ratios = [
                (boundaries[k + 1] - boundaries[k]) / float(total_px)
                for k in range(len(boundaries) - 1)
            ]
            self._column_ratios = normalize_ratios(new_ratios, len(self._column_ratios))
            self._recompute_and_repaint()

        elif seg["kind"] == "row":
            if not self._cell_rects_widget:
                return
            top_of_all = self._cell_rects_widget[0].top()
            bottom_of_all = self._cell_rects_widget[-1].bottom()
            total_px = max(1, bottom_of_all - top_of_all)

            i = seg["index"]
            boundaries = [top_of_all]
            cursor = top_of_all
            for r in self._row_ratios:
                cursor += r * total_px
                boundaries.append(cursor)
            boundaries[-1] = bottom_of_all

            min_px = max(MIN_SLOT_PX, min_ratio_for_count(len(self._row_ratios)) * total_px)
            local_y = max(top_of_all, min(bottom_of_all, pos.y()))
            new_y = max(boundaries[i] + min_px, min(boundaries[i + 2] - min_px, local_y))
            boundaries[i + 1] = new_y

            new_ratios = [
                (boundaries[k + 1] - boundaries[k]) / float(total_px)
                for k in range(len(boundaries) - 1)
            ]
            self._row_ratios = normalize_ratios(new_ratios, len(self._row_ratios))
            self._recompute_and_repaint()

    # ------------------------------------------------------------------
    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        if self._monitor is None:
            painter.end()
            return

        if not self._checked_windows:
            painter.setPen(QColor(COLOR_TEXT_MUTED))
            font = QFont("Segoe UI", 11)
            painter.setFont(font)
            draw_rect = self._canvas_drawing_rect()
            painter.drawText(
                draw_rect,
                Qt.AlignCenter,
                "Check windows on the left to preview the layout here.",
            )
            painter.end()
            return

        accent = QColor(COLOR_ACCENT)
        badge_bg = QColor(COLOR_ACCENT)
        cell_border_pen = QPen(accent)
        cell_border_pen.setWidth(2)

        for idx, rect in enumerate(self._cell_rects_widget):
            window_info = self._checked_windows[idx]

            # Cell background + border.
            painter.setPen(cell_border_pen)
            painter.setBrush(QBrush(QColor("#10141a")))
            painter.drawRoundedRect(rect, 6, 6)

            # Header strip with number badge + title, drawn above where
            # the live thumbnail content will be placed.
            header_height = 24
            header_rect = QRect(rect.left(), rect.top(), rect.width(), header_height)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QBrush(QColor(COLOR_BG_PANEL_ALT)))
            painter.drawRoundedRect(header_rect, 6, 6)
            # Square off the bottom corners of the header strip so it
            # reads as a flat divider rather than a separate pill.
            painter.drawRect(
                QRect(rect.left(), rect.top() + header_height - 8, rect.width(), 8)
            )

            badge_size = 18
            badge_rect = QRect(
                rect.left() + 4, rect.top() + 3, badge_size, badge_size
            )
            painter.setBrush(QBrush(badge_bg))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(badge_rect, 4, 4)
            painter.setPen(QColor("white"))
            font = QFont("Segoe UI", 9, QFont.Bold)
            painter.setFont(font)
            painter.drawText(badge_rect, Qt.AlignCenter, str(idx + 1))

            icon_size = 14
            icon_x = rect.left() + badge_size + 8
            icon_y = rect.top() + (header_height - icon_size) // 2
            icon_pixmap = get_window_icon_pixmap(window_info, size=28)
            title_text_x = icon_x
            if icon_pixmap is not None and not icon_pixmap.isNull():
                scaled_icon = icon_pixmap.scaled(
                    icon_size, icon_size, Qt.KeepAspectRatio, Qt.SmoothTransformation
                )
                painter.drawPixmap(icon_x, icon_y, scaled_icon)
                title_text_x = icon_x + icon_size + 5
                label_text = window_info["title"]
            else:
                emoji = get_emoji_for_process(window_info["process_name"])
                label_text = "{}  {}".format(emoji, window_info["title"])

            text_rect = QRect(
                title_text_x,
                rect.top() + 2,
                max(10, rect.right() - title_text_x - 6),
                header_height - 4,
            )
            painter.setPen(QColor(COLOR_TEXT_PRIMARY))
            font = QFont("Segoe UI", 9)
            painter.setFont(font)
            metrics_text = painter.fontMetrics().elidedText(
                label_text, Qt.ElideRight, text_rect.width()
            )
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, metrics_text)

        # FEATURE 3: draw the draggable divider handles for custom
        # templates so they are visually discoverable even before the
        # user hovers over one.
        if self._template_mode in ("master_stack", "columns", "rows"):
            divider_pen = QPen(QColor(COLOR_BORDER_LIGHT))
            divider_pen.setWidth(3)
            painter.setPen(divider_pen)
            for seg in self._build_divider_segments():
                x1, y1 = seg["p1"]
                x2, y2 = seg["p2"]
                painter.drawLine(x1, y1, x2, y2)

        # FEATURE 1: drag-in-progress visual feedback -- a dashed,
        # translucent outline on the cell being dragged, and a solid
        # green outline on whichever cell it is currently hovering over
        # as the swap target.
        if self._drag_mode == "reorder" and self._drag_source_index != -1:
            source_rect = self._cell_rects_widget[self._drag_source_index]
            ghost_pen = QPen(QColor(COLOR_ACCENT_HOVER))
            ghost_pen.setWidth(3)
            ghost_pen.setStyle(Qt.DashLine)
            painter.setPen(ghost_pen)
            painter.setBrush(QColor(255, 255, 255, 18))
            painter.drawRoundedRect(source_rect, 6, 6)

            if self._drag_hover_index not in (-1, self._drag_source_index):
                hover_rect = self._cell_rects_widget[self._drag_hover_index]
                hover_pen = QPen(QColor(COLOR_GREEN))
                hover_pen.setWidth(3)
                painter.setPen(hover_pen)
                painter.setBrush(Qt.NoBrush)
                painter.drawRoundedRect(hover_rect, 6, 6)

        painter.end()


# ======================================================================
# SECTION 11: Quick layout icon widget (small SVG-free grid glyph)
# ======================================================================

class QuickLayoutGlyph(QToolButton):
    """
    One of the small grid-pattern buttons in the right panel's "Quick
    Layouts" section (1x1, 1x2, 2x1, 2x2, etc.) Drawn procedurally
    rather than from an image file so the whole app stays a single
    script with zero external asset dependencies. Purely a visual
    shortcut -- clicking one sets the "preview window count" hint state
    used only to illustrate that shape; the REAL grid actually applied
    is always computed live from however many windows are checked, per
    the Smart Equal Grid algorithm (this mirrors the reference design,
    where Quick Layouts are a visual style picker, while the live N
    count still drives the literal math).
    """

    def __init__(self, rows, cols, label_n, parent=None):
        super().__init__(parent)
        self.rows = rows
        self.cols = cols
        self.label_n = label_n
        self.setObjectName("QuickLayoutButton")
        self.setCheckable(True)
        self.setFixedSize(54, 40)
        self.setToolTip("{} window{}".format(label_n, "" if label_n == 1 else "s"))

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        pad = 9
        rect = self.rect().adjusted(pad, pad, -pad, -pad)
        color = COLOR_ACCENT if self.isChecked() else COLOR_TEXT_SECONDARY
        pen = QPen(QColor(color))
        pen.setWidth(2)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)

        cell_w = rect.width() / float(self.cols)
        cell_h = rect.height() / float(self.rows)
        for r in range(self.rows):
            for c in range(self.cols):
                cx = rect.left() + c * cell_w
                cy = rect.top() + r * cell_h
                cell_rect = QRect(
                    int(round(cx)), int(round(cy)),
                    max(1, int(round(cell_w)) - 2),
                    max(1, int(round(cell_h)) - 2),
                )
                painter.drawRoundedRect(cell_rect, 2, 2)
        painter.end()


# ======================================================================
# SECTION 12: Main Application Window
# ======================================================================

class WindowLayoutManagerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Window Layout Manager")
        self.resize(1440, 900)
        self.setMinimumSize(1100, 700)

        self._own_hwnd = None
        self._all_windows = []          # full enumerated list, list[dict]
        self._row_widgets = {}          # hwnd -> WindowListRow
        self._checked_hwnds = set()     # set of currently checked HWNDs
        self._monitors = []             # list[MonitorInfo]
        self._selected_monitor = None   # MonitorInfo
        self._monitor_buttons = []      # list[QToolButton]
        self._pre_layout_rects = {}     # hwnd -> (l,t,r,b) saved just before Apply, for Reset Positions

        # --- Feature 1 (drag reorder) / Feature 4 (window cycling) ---
        self._manual_order_hwnds = []   # user drag-reordered hwnd order, full checked set
        self._last_visible_hwnds = []   # hwnds actually visible/tiled after cycle capping
        self._cycle_offset = 0
        self._currently_tiled_hwnds = set()
        self._cycle_timer = QTimer(self)
        self._cycle_timer.setInterval(5000)
        self._cycle_timer.timeout.connect(self._on_cycle_tick)

        # --- Feature 6 (smooth animated grid transitions) ------------
        self._window_animator = WindowMoveAnimator(self)

        # --- Feature 2 (global hotkey) --------------------------------
        self._hotkey_registered = False

        self.setStyleSheet(build_stylesheet())

        self._build_ui()
        self._refresh_monitors()
        self.refresh_window_list()

        # Periodic light-weight refresh of minimized/restored status dots
        # and the live preview (in case the user manually moves/resizes
        # a checked window while this app is open) without doing a full,
        # heavier re-enumeration. A full re-enumeration only happens on
        # explicit "Refresh Windows" clicks or window-count drift checks.
        self._light_refresh_timer = QTimer(self)
        self._light_refresh_timer.timeout.connect(self._on_light_refresh_tick)
        self._light_refresh_timer.start(1500)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_top_bar())

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        body.addWidget(self._build_left_panel(), stretch=0)
        body.addWidget(self._build_center_panel(), stretch=1)
        body.addWidget(self._build_right_panel(), stretch=0)

        body_widget = QWidget()
        body_widget.setLayout(body)
        root.addWidget(body_widget, stretch=1)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready.")

    # ------------------------------------------------------------------
    def _build_top_bar(self):
        bar = QFrame()
        bar.setObjectName("TopBar")
        bar.setFixedHeight(64)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(18, 10, 18, 10)
        layout.setSpacing(14)

        logo = QLabel("🧮")
        logo.setStyleSheet("font-size: 20px;")
        layout.addWidget(logo)

        title = QLabel("Window Layout Manager")
        title.setObjectName("TitleLabel")
        layout.addWidget(title)

        layout.addSpacing(16)

        self._monitor_tabs_layout = QHBoxLayout()
        self._monitor_tabs_layout.setSpacing(8)
        layout.addLayout(self._monitor_tabs_layout)

        layout.addStretch(1)

        refresh_all_btn = QPushButton("⟲  Refresh")
        refresh_all_btn.setObjectName("TopBarButton")
        refresh_all_btn.clicked.connect(self.refresh_window_list)
        layout.addWidget(refresh_all_btn)

        auto_arrange_btn = QPushButton("▦  Auto Arrange")
        auto_arrange_btn.setObjectName("TopBarButton")
        auto_arrange_btn.clicked.connect(self.select_all_and_apply)
        layout.addWidget(auto_arrange_btn)

        detect_btn = QPushButton("🖥  Detect Windows")
        detect_btn.setObjectName("TopBarButton")
        detect_btn.clicked.connect(self._refresh_monitors_and_windows)
        layout.addWidget(detect_btn)

        return bar

    # ------------------------------------------------------------------
    def _build_left_panel(self):
        panel = QFrame()
        panel.setObjectName("LeftPanel")
        panel.setFixedWidth(330)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        header_row = QHBoxLayout()
        header_label = QLabel("Open Windows")
        header_label.setObjectName("SectionHeader")
        header_row.addWidget(header_label)
        header_row.addStretch(1)
        self.count_badge = QLabel("0")
        self.count_badge.setObjectName("CountBadge")
        header_row.addWidget(self.count_badge)
        layout.addLayout(header_row)

        self.search_box = QLineEdit()
        self.search_box.setObjectName("SearchBox")
        self.search_box.setPlaceholderText("Search open windows...")
        self.search_box.textChanged.connect(self._apply_search_filter)
        layout.addWidget(self.search_box)

        self.list_widget = QListWidget()
        self.list_widget.setObjectName("WindowList")
        self.list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        self.list_widget.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.list_widget.setFrameShape(QFrame.NoFrame)
        self.list_widget.setSpacing(4)
        self.list_widget.setUniformItemSizes(True)
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.list_widget.setResizeMode(QListWidget.Adjust)
        layout.addWidget(self.list_widget, stretch=1)

        selection_row = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self.select_all)
        select_none_btn = QPushButton("Select None")
        select_none_btn.clicked.connect(self.select_none)
        selection_row.addWidget(select_all_btn)
        selection_row.addWidget(select_none_btn)
        layout.addLayout(selection_row)

        refresh_row = QHBoxLayout()
        self.refresh_button = QPushButton("⟲  Refresh Windows")
        self.refresh_button.clicked.connect(self.refresh_window_list)
        refresh_row.addWidget(self.refresh_button)
        layout.addLayout(refresh_row)

        self.last_refreshed_label = QLabel("")
        self.last_refreshed_label.setObjectName("MutedText")
        layout.addWidget(self.last_refreshed_label)

        return panel

    # ------------------------------------------------------------------
    def _build_center_panel(self):
        panel = QFrame()
        panel.setObjectName("CenterPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        self.preview_canvas = LayoutPreviewCanvas()
        self.preview_canvas.order_changed.connect(self._on_canvas_order_changed)
        layout.addWidget(self.preview_canvas, stretch=1)

        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(10)

        grid_label = QLabel("Grid:")
        grid_label.setObjectName("MutedText")
        bottom_row.addWidget(grid_label)

        self.grid_size_label = QLabel("0 x 0")
        self.grid_size_label.setStyleSheet(
            "color: {}; font-weight: 600;".format(COLOR_TEXT_PRIMARY)
        )
        bottom_row.addWidget(self.grid_size_label)

        bottom_row.addStretch(1)

        self.usage_label = QLabel("")
        self.usage_label.setObjectName("MutedText")
        bottom_row.addWidget(self.usage_label)

        layout.addLayout(bottom_row)

        return panel

    # ------------------------------------------------------------------
    def _build_right_panel(self):
        panel = QFrame()
        panel.setObjectName("RightPanel")
        panel.setFixedWidth(300)
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(0)

        # --- Scrollable content area for all cards -----------------------
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setObjectName("RightPanelScroll")

        content = QWidget()
        content.setObjectName("RightPanelContent")
        content.setStyleSheet("background-color: transparent;")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(16, 16, 16, 8)
        content_layout.setSpacing(12)

        # --- Quick Layouts card -----------------------------------------
        quick_card = QFrame()
        quick_card.setObjectName("SectionCard")
        quick_layout = QVBoxLayout(quick_card)
        quick_layout.setContentsMargins(14, 12, 14, 14)
        quick_layout.setSpacing(8)

        quick_header = QLabel("Quick Layouts")
        quick_header.setObjectName("SectionHeader")
        quick_layout.addWidget(quick_header)

        quick_grid = QGridLayout()
        quick_grid.setSpacing(6)
        quick_shapes = [
            (1, 1, 1), (1, 2, 2), (2, 1, 2), (1, 3, 3),
            (2, 2, 4), (2, 3, 6), (3, 3, 9), (3, 4, 12),
        ]
        self._quick_layout_buttons = []
        for i, (rows, cols, n) in enumerate(quick_shapes):
            glyph = QuickLayoutGlyph(rows, cols, n)
            glyph.clicked.connect(
                lambda _checked, n=n: self._apply_quick_layout_hint(n)
            )
            quick_grid.addWidget(glyph, i // 4, i % 4)
            self._quick_layout_buttons.append(glyph)
        quick_layout.addLayout(quick_grid)

        content_layout.addWidget(quick_card)

        # --- Layout Template card (Feature 3) -----------------------------
        template_card = QFrame()
        template_card.setObjectName("SectionCard")
        template_layout = QVBoxLayout(template_card)
        template_layout.setContentsMargins(14, 12, 14, 14)
        template_layout.setSpacing(8)

        template_header = QLabel("Layout Template")
        template_header.setObjectName("SectionHeader")
        template_layout.addWidget(template_header)

        # Vertical stack of full-width buttons -- avoids text clipping
        # that occurs in a tight multi-column grid at 300px panel width.
        template_vbox = QVBoxLayout()
        template_vbox.setSpacing(4)
        self._template_buttons = {}
        for mode_key, label in [
            ("equal",        "Equal Grid"),
            ("master_stack", "Master+Stack"),
            ("columns",      "Columns"),
            ("rows",         "Vertical Stack"),
            ("custom_grid",  "Custom Grid"),
        ]:
            btn = QToolButton()
            btn.setObjectName("MonitorTab")
            btn.setCheckable(True)
            btn.setText(label)
            btn.setChecked(mode_key == "equal")
            btn.clicked.connect(lambda _checked, m=mode_key: self._select_template_mode(m))
            template_vbox.addWidget(btn)
            self._template_buttons[mode_key] = btn
        template_layout.addLayout(template_vbox)

        template_hint = QLabel(
            "Master+Stack, Columns, and Vertical Stack let you drag "
            "divider lines in the preview to resize slots. "
            "Custom Grid lets you set exact rows × columns."
        )
        template_hint.setObjectName("MutedText")
        template_hint.setWordWrap(True)
        template_layout.addWidget(template_hint)

        content_layout.addWidget(template_card)

        # --- Custom Grid config card (shown/hidden based on template) ---
        self.custom_grid_card = QFrame()
        self.custom_grid_card.setObjectName("SectionCard")
        cg_layout = QVBoxLayout(self.custom_grid_card)
        cg_layout.setContentsMargins(14, 12, 14, 14)
        cg_layout.setSpacing(8)

        cg_header = QLabel("Custom Grid — Rows × Columns")
        cg_header.setObjectName("SectionHeader")
        cg_layout.addWidget(cg_header)

        cg_spinbox_row = QHBoxLayout()
        cg_spinbox_row.setSpacing(8)

        cg_spinbox_row.addWidget(QLabel("Rows:"))
        self.grid_rows_spin = QSpinBox()
        self.grid_rows_spin.setMinimum(1)
        self.grid_rows_spin.setMaximum(20)
        self.grid_rows_spin.setValue(2)
        self.grid_rows_spin.setFixedWidth(58)
        cg_spinbox_row.addWidget(self.grid_rows_spin)

        cg_spinbox_row.addSpacing(10)
        cg_spinbox_row.addWidget(QLabel("Cols:"))
        self.grid_cols_spin = QSpinBox()
        self.grid_cols_spin.setMinimum(1)
        self.grid_cols_spin.setMaximum(20)
        self.grid_cols_spin.setValue(3)
        self.grid_cols_spin.setFixedWidth(58)
        cg_spinbox_row.addWidget(self.grid_cols_spin)
        cg_spinbox_row.addStretch(1)

        cg_layout.addLayout(cg_spinbox_row)

        # min-size warning label -- hidden until a min-size violation is detected
        self.cg_warning_label = QLabel("")
        self.cg_warning_label.setObjectName("WarningText")
        self.cg_warning_label.setWordWrap(True)
        self.cg_warning_label.setVisible(False)
        cg_layout.addWidget(self.cg_warning_label)

        cg_hint = QLabel(
            "e.g. 2 rows × 6 cols = 12 Telegram windows, each full "
            "height of half the monitor. If cells would be narrower "
            "than Telegram's minimum size a warning will appear."
        )
        cg_hint.setObjectName("MutedText")
        cg_hint.setWordWrap(True)
        cg_layout.addWidget(cg_hint)

        self.grid_rows_spin.valueChanged.connect(self._on_custom_grid_changed)
        self.grid_cols_spin.valueChanged.connect(self._on_custom_grid_changed)

        # start hidden; shown only when "Custom Grid" template is active
        self.custom_grid_card.setVisible(False)
        content_layout.addWidget(self.custom_grid_card)

        # --- Gap (window spacing) card --------------------------------
        gap_card = QFrame()
        gap_card.setObjectName("SectionCard")
        gap_layout = QVBoxLayout(gap_card)
        gap_layout.setContentsMargins(14, 12, 14, 14)
        gap_layout.setSpacing(8)

        gap_header_row = QHBoxLayout()
        gap_header = QLabel("Gap Between Windows")
        gap_header.setObjectName("SectionHeader")
        gap_header_row.addWidget(gap_header)
        gap_header_row.addStretch(1)
        self.gap_value_label = QLabel("0 px")
        self.gap_value_label.setObjectName("MutedText")
        gap_header_row.addWidget(self.gap_value_label)
        gap_layout.addLayout(gap_header_row)

        self.gap_slider = QSlider(Qt.Horizontal)
        self.gap_slider.setMinimum(0)
        self.gap_slider.setMaximum(MAX_GAP_PX)
        self.gap_slider.setValue(0)
        self.gap_slider.setSingleStep(1)
        self.gap_slider.setPageStep(4)
        self.gap_slider.valueChanged.connect(self._on_gap_slider_changed)
        gap_layout.addWidget(self.gap_slider)

        gap_hint = QLabel(
            "0 = windows touch edge-to-edge with no gap. Applies to "
            "every template, including the live preview."
        )
        gap_hint.setObjectName("MutedText")
        gap_hint.setWordWrap(True)
        gap_layout.addWidget(gap_hint)

        content_layout.addWidget(gap_card)

        # --- Window Cycling card (Feature 4) ------------------------------
        cycle_card = QFrame()
        cycle_card.setObjectName("SectionCard")
        cycle_layout = QVBoxLayout(cycle_card)
        cycle_layout.setContentsMargins(14, 12, 14, 14)
        cycle_layout.setSpacing(10)

        cycle_header = QLabel("Window Cycling")
        cycle_header.setObjectName("SectionHeader")
        cycle_layout.addWidget(cycle_header)

        self.cycle_status_label = QLabel("No overflow windows queued.")
        self.cycle_status_label.setObjectName("SubText")
        self.cycle_status_label.setWordWrap(True)
        cycle_layout.addWidget(self.cycle_status_label)

        cycle_buttons_row = QHBoxLayout()
        cycle_buttons_row.setSpacing(8)

        self.cycle_next_button = QPushButton("▶  Cycle Next")
        self.cycle_next_button.setEnabled(False)
        self.cycle_next_button.clicked.connect(self._cycle_next_manual)
        cycle_buttons_row.addWidget(self.cycle_next_button)

        self.cycle_auto_button = QToolButton()
        self.cycle_auto_button.setObjectName("QuickLayoutButton")
        self.cycle_auto_button.setText("Auto")
        self.cycle_auto_button.setCheckable(True)
        self.cycle_auto_button.setEnabled(False)
        self.cycle_auto_button.setFixedHeight(34)
        self.cycle_auto_button.setToolTip(
            "Automatically rotate overflow windows into view every 5 seconds"
        )
        self.cycle_auto_button.clicked.connect(self._toggle_auto_cycle)
        cycle_buttons_row.addWidget(self.cycle_auto_button)

        cycle_layout.addLayout(cycle_buttons_row)
        content_layout.addWidget(cycle_card)

        # --- Layout Information card -------------------------------------
        info_card = QFrame()
        info_card.setObjectName("SectionCard")
        info_layout = QVBoxLayout(info_card)
        info_layout.setContentsMargins(14, 12, 14, 14)
        info_layout.setSpacing(6)

        info_header = QLabel("Layout Information")
        info_header.setObjectName("SectionHeader")
        info_layout.addWidget(info_header)

        self._info_rows = {}
        for key, label in [
            ("selected", "Selected Windows"),
            ("grid", "Grid Size"),
            ("monitor", "Monitor"),
            ("resolution", "Resolution"),
            ("workarea", "Work Area"),
            ("usage", "Usage"),
        ]:
            row = QHBoxLayout()
            name_label = QLabel(label)
            name_label.setObjectName("SubText")
            value_label = QLabel("—")
            value_label.setStyleSheet(
                "color: {}; font-weight: 600;".format(COLOR_TEXT_PRIMARY)
            )
            value_label.setAlignment(Qt.AlignRight)
            row.addWidget(name_label)
            row.addStretch(1)
            row.addWidget(value_label)
            info_layout.addLayout(row)
            self._info_rows[key] = value_label

        content_layout.addWidget(info_card)

        content_layout.addStretch(1)
        scroll.setWidget(content)
        panel_layout.addWidget(scroll, stretch=1)

        # --- Action buttons (pinned at bottom, always visible) -----------
        btn_container = QFrame()
        btn_container.setObjectName("RightPanelButtons")
        btn_container.setStyleSheet(
            "QFrame#RightPanelButtons {{ background-color: {}; border-top: 1px solid {}; }}".format(
                COLOR_BG_PANEL, COLOR_BORDER
            )
        )
        btn_layout = QVBoxLayout(btn_container)
        btn_layout.setContentsMargins(16, 10, 16, 16)
        btn_layout.setSpacing(8)

        self.apply_button = QPushButton("▦  Apply Equal Layout")
        self.apply_button.setObjectName("PrimaryButton")
        self.apply_button.clicked.connect(self.apply_equal_layout)
        btn_layout.addWidget(self.apply_button)

        self.save_button = QPushButton("💾  Save Layout")
        self.save_button.clicked.connect(self.save_current_layout)
        btn_layout.addWidget(self.save_button)

        self.load_button = QPushButton("📂  Load Layout")
        self.load_button.clicked.connect(self.load_saved_layout)
        btn_layout.addWidget(self.load_button)

        self.reset_button = QPushButton("⟲  Reset Positions")
        self.reset_button.setObjectName("DangerButton")
        self.reset_button.clicked.connect(self.reset_positions)
        btn_layout.addWidget(self.reset_button)

        panel_layout.addWidget(btn_container)

        return panel

    # ------------------------------------------------------------------
    # Native HWND bookkeeping
    # ------------------------------------------------------------------
    def showEvent(self, event):
        super().showEvent(event)
        try:
            self._own_hwnd = int(self.winId())
            self.preview_canvas.set_own_hwnd(self._own_hwnd)
        except Exception:
            self._own_hwnd = None

        # FEATURE 2: register the Ctrl+Alt+L global hotkey against our
        # own window's message queue the first time we have a real
        # HWND to bind it to. The native event filter that actually
        # reacts to the resulting WM_HOTKEY message is installed once,
        # app-wide, in main() (Section 13).
        if self._own_hwnd and not self._hotkey_registered:
            if register_global_hotkey(self._own_hwnd):
                self._hotkey_registered = True
                self.status_bar.showMessage(
                    "Ready.  Tip: Ctrl+Alt+L applies the current layout "
                    "from anywhere, even while this window is minimized."
                )

    def closeEvent(self, event):
        try:
            self.preview_canvas.clear_thumbnails()
        except Exception:
            pass
        try:
            self._light_refresh_timer.stop()
        except Exception:
            pass
        try:
            self._cycle_timer.stop()
        except Exception:
            pass
        try:
            self._window_animator.stop()
        except Exception:
            pass
        if self._hotkey_registered and self._own_hwnd:
            unregister_global_hotkey(self._own_hwnd)
        super().closeEvent(event)

    # ------------------------------------------------------------------
    # Monitor handling
    # ------------------------------------------------------------------
    def _refresh_monitors(self):
        previously_selected_index = (
            self._selected_monitor.index if self._selected_monitor else None
        )

        try:
            self._monitors = enumerate_monitors()
        except Exception as exc:
            QMessageBox.critical(
                self, "Monitor Detection Error",
                "Failed to enumerate monitors:\n{}".format(exc),
            )
            self._monitors = [get_primary_monitor()]

        # Clear old tab buttons.
        while self._monitor_tabs_layout.count():
            item = self._monitor_tabs_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._monitor_buttons = []

        for monitor in self._monitors:
            btn = QToolButton()
            btn.setObjectName("MonitorTab")
            btn.setCheckable(True)
            btn.setText("🖥  {}\n{}".format(monitor.label, monitor.resolution_label))
            btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
            btn.clicked.connect(
                lambda _checked, m=monitor: self._select_monitor(m)
            )
            self._monitor_tabs_layout.addWidget(btn)
            self._monitor_buttons.append((monitor, btn))

        # Re-select the same physical monitor index if it still exists,
        # otherwise fall back to the primary monitor.
        target = None
        if previously_selected_index is not None:
            for m in self._monitors:
                if m.index == previously_selected_index:
                    target = m
                    break
        if target is None:
            for m in self._monitors:
                if m.is_primary:
                    target = m
                    break
        if target is None and self._monitors:
            target = self._monitors[0]

        if target is not None:
            self._select_monitor(target)

    def _refresh_monitors_and_windows(self):
        self._refresh_monitors()
        self.refresh_window_list()

    def _select_monitor(self, monitor):
        self._selected_monitor = monitor
        for m, btn in self._monitor_buttons:
            btn.setChecked(m.index == monitor.index)
        self.preview_canvas.set_monitor(monitor)
        self._update_preview_and_info()

    # ------------------------------------------------------------------
    # List population
    # ------------------------------------------------------------------
    def refresh_window_list(self):
        try:
            self._all_windows = enumerate_taskbar_windows(own_hwnd=self._own_hwnd)
        except Exception as exc:
            QMessageBox.critical(
                self, "Enumeration Error",
                "Failed to enumerate open windows:\n{}".format(exc),
            )
            self._all_windows = []

        # Drop checked-state for hwnds that no longer exist.
        live_hwnds = {w["hwnd"] for w in self._all_windows}
        self._checked_hwnds &= live_hwnds

        self.list_widget.clear()
        self._row_widgets = {}

        for window_info in self._all_windows:
            item = QListWidgetItem()
            row_widget = WindowListRow(window_info)
            row_widget.set_checked(window_info["hwnd"] in self._checked_hwnds)
            row_widget.toggled.connect(self._on_row_toggled)
            item.setSizeHint(row_widget.sizeHint())
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, row_widget)
            self._row_widgets[window_info["hwnd"]] = (item, row_widget)

        self.count_badge.setText(str(len(self._all_windows)))
        self._apply_search_filter(self.search_box.text())
        self._update_preview_and_info()

        self.last_refreshed_label.setText(
            "Last refreshed: {}".format(time.strftime("%I:%M:%S %p").lstrip("0"))
        )
        self.status_bar.showMessage(
            "Found {} window{} eligible for tiling.".format(
                len(self._all_windows), "" if len(self._all_windows) == 1 else "s"
            )
        )

    def _on_light_refresh_tick(self):
        """
        Lightweight per-tick refresh: re-checks IsWindow()/IsIconic() for
        already-listed HWNDs to keep status dots accurate and to drop
        rows for windows that the user closed outside this app, WITHOUT
        doing a full EnumWindows pass (cheaper, avoids visual list
        reshuffling while the user is actively browsing the list).
        """
        stale_hwnds = []
        any_minimized_state_changed = False
        for hwnd, (item, row_widget) in list(self._row_widgets.items()):
            try:
                if not win32gui.IsWindow(hwnd):
                    stale_hwnds.append(hwnd)
                    continue
                is_minimized = bool(win32gui.IsIconic(hwnd))
                if is_minimized != row_widget.window_info.get("is_minimized"):
                    row_widget.window_info["is_minimized"] = is_minimized
                    dot_color = COLOR_ORANGE if is_minimized else COLOR_GREEN
                    row_widget.status_dot.setStyleSheet(
                        "background-color: {}; border-radius: 4px;".format(dot_color)
                    )
                    any_minimized_state_changed = True
            except Exception:
                stale_hwnds.append(hwnd)

        if stale_hwnds:
            # A tracked window disappeared -- do a full, proper refresh so
            # the list and all derived state stay perfectly consistent.
            self.refresh_window_list()
            return

        if any_minimized_state_changed:
            self._update_preview_and_info()

    def _apply_search_filter(self, query):
        query = (query or "").strip().lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            row_widget = self.list_widget.itemWidget(item)
            if row_widget is None:
                continue
            info = row_widget.window_info
            haystack = (info["title"] + " " + info["process_name"]).lower()
            item.setHidden(bool(query) and query not in haystack)

    # ------------------------------------------------------------------
    # Selection handling
    # ------------------------------------------------------------------
    def _on_row_toggled(self, hwnd, checked):
        if checked:
            self._checked_hwnds.add(hwnd)
        else:
            self._checked_hwnds.discard(hwnd)
        self._update_preview_and_info()

    def select_all(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.isHidden():
                continue
            row_widget = self.list_widget.itemWidget(item)
            if row_widget is not None:
                row_widget.set_checked(True)
                self._checked_hwnds.add(row_widget.hwnd)
        self._update_preview_and_info()

    def select_none(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            row_widget = self.list_widget.itemWidget(item)
            if row_widget is not None:
                row_widget.set_checked(False)
        self._checked_hwnds.clear()
        self._update_preview_and_info()

    def select_all_and_apply(self):
        self.select_all()
        self.apply_equal_layout()

    def _apply_quick_layout_hint(self, n):
        """
        Quick Layout buttons are a visual shortcut: clicking one simply
        checks the first N currently-listed windows (in list order) so
        the preview and Apply action have something concrete to work
        with, mirroring how a user would otherwise manually check N
        boxes. The actual grid shape used at Apply time is still always
        derived live from calculate_grid_dimensions(N), guaranteeing the
        Smart Equal Grid rules are the single source of truth.
        """
        self.select_none()
        count = 0
        for i in range(self.list_widget.count()):
            if count >= n:
                break
            item = self.list_widget.item(i)
            if item.isHidden():
                continue
            row_widget = self.list_widget.itemWidget(item)
            if row_widget is not None:
                row_widget.set_checked(True)
                self._checked_hwnds.add(row_widget.hwnd)
                count += 1
        self._update_preview_and_info()

    # ------------------------------------------------------------------
    def _get_checked_window_infos_in_order(self):
        """
        Returns the list of window_info dicts for every currently
        checked HWND. Normally this is the same stable top-to-bottom
        order the windows appear in the sidebar list, but FEATURE 1
        (drag & drop reorder in the preview) lets the user override
        that with a manual order, captured in self._manual_order_hwnds.
        The manual order is honored as long as it still names exactly
        the currently-checked set of windows; the moment the selection
        changes (a box gets (un)checked, a window closes, etc.) it is
        discarded and we fall back to plain list order again.
        """
        list_order = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            row_widget = self.list_widget.itemWidget(item)
            if row_widget is not None and row_widget.hwnd in self._checked_hwnds:
                list_order.append(row_widget.window_info)

        if self._manual_order_hwnds:
            checked_set = {info["hwnd"] for info in list_order}
            manual_set = set(self._manual_order_hwnds)
            if manual_set == checked_set:
                info_by_hwnd = {info["hwnd"]: info for info in list_order}
                return [info_by_hwnd[h] for h in self._manual_order_hwnds]
            # The checked selection changed since the last manual
            # reorder -- it no longer applies cleanly to a different
            # set of windows, so drop it and use plain list order.
            self._manual_order_hwnds = []

        return list_order

    def _get_visible_window_infos(self):
        """
        Returns the subset (and order) of currently checked windows
        that should actually be drawn in the preview AND tiled by
        Apply right now -- i.e. the full checked list, capped to
        however many slots the active template supports, rotated by
        the current Window Cycling offset whenever there is overflow
        (FEATURE 4). Also refreshes the Window Cycling card's status
        text/button states as a side effect, since both always need to
        agree on the same overflow calculation.
        """
        full = self._get_checked_window_infos_in_order()
        max_slots = TEMPLATE_MAX_SLOTS.get(self.preview_canvas.get_template_mode())

        if max_slots is None or len(full) <= max_slots:
            self._cycle_offset = 0
            visible = full
        else:
            offset = self._cycle_offset % len(full)
            rotated = full[offset:] + full[:offset]
            visible = rotated[:max_slots]

        self._last_visible_hwnds = [info["hwnd"] for info in visible]
        self._update_cycle_status_label(len(full), max_slots)
        return visible

    def _update_cycle_status_label(self, total, max_slots):
        if max_slots is None or total <= max_slots:
            self.cycle_status_label.setText("No overflow windows queued.")
            self.cycle_next_button.setEnabled(False)
            self.cycle_auto_button.setEnabled(False)
            if self.cycle_auto_button.isChecked():
                self.cycle_auto_button.setChecked(False)
                self._cycle_timer.stop()
        else:
            overflow = total - max_slots
            self.cycle_status_label.setText(
                "{} window{} queued beyond the {}-slot template -- "
                "cycle to rotate {} into view.".format(
                    overflow, "" if overflow == 1 else "s", max_slots,
                    "it" if overflow == 1 else "them",
                )
            )
            self.cycle_next_button.setEnabled(True)
            self.cycle_auto_button.setEnabled(True)

    def _select_template_mode(self, mode):
        """FEATURE 3: switches the active layout template."""
        for key, btn in self._template_buttons.items():
            btn.setChecked(key == mode)
        self.preview_canvas.set_template_mode(mode)
        self._cycle_offset = 0
        # Show the Custom Grid config card only when that template is active.
        self.custom_grid_card.setVisible(mode == "custom_grid")
        if mode == "custom_grid":
            self._update_custom_grid_warning()
        self._update_preview_and_info()

    def _on_custom_grid_changed(self):
        """Called whenever the rows or cols spinbox value changes."""
        rows = self.grid_rows_spin.value()
        cols = self.grid_cols_spin.value()
        self.preview_canvas.set_custom_grid_dims(rows, cols)
        self._update_custom_grid_warning()
        self._update_preview_and_info()

    def _update_custom_grid_warning(self):
        """
        Shows a warning if the chosen rows×cols would produce cells
        smaller than Telegram's (or any app's) minimum window size,
        which would cause windows to overlap after Apply because the OS
        refuses to shrink them below their minimum. The warning names
        the problematic dimension and suggests a safe alternative.
        """
        if self._selected_monitor is None:
            self.cg_warning_label.setVisible(False)
            return
        work_area = self._selected_monitor.work_rect
        rows = self.grid_rows_spin.value()
        cols = self.grid_cols_spin.value()
        fits, cell_w, cell_h, safe_cols, safe_rows = check_grid_fits_min_size(
            work_area, 0, rows, cols
        )
        if fits:
            self.cg_warning_label.setVisible(False)
        else:
            problems = []
            if cell_w < DEFAULT_MIN_WIN_W:
                problems.append(
                    "width {}px < {}px min — try ≤ {} cols".format(
                        cell_w, DEFAULT_MIN_WIN_W, safe_cols
                    )
                )
            if cell_h < DEFAULT_MIN_WIN_H:
                problems.append(
                    "height {}px < {}px min — try ≤ {} rows".format(
                        cell_h, DEFAULT_MIN_WIN_H, safe_rows
                    )
                )
            self.cg_warning_label.setText(
                "⚠ Cell too small: {}. "
                "Windows will overlap if applied as-is.".format("; ".join(problems))
            )
            self.cg_warning_label.setVisible(True)

    def _on_gap_slider_changed(self, value):
        """Live-updates the inter-window gap as the slider is dragged."""
        self.gap_value_label.setText("{} px".format(value))
        self.preview_canvas.set_gap_px(value)

    def _on_canvas_order_changed(self, new_visible_order):
        """
        FEATURE 1 callback: the preview canvas just swapped two cells.
        `new_visible_order` is the new hwnd order of whatever subset of
        windows is currently VISIBLE (post cycle-capping). We splice
        that back into the full checked-window order at exactly the
        positions the previously-visible hwnds occupied, so a reorder
        made while overflow windows are cycled out still produces a
        sensible, stable full ordering.
        """
        full_infos = self._get_checked_window_infos_in_order()
        full_hwnds = [info["hwnd"] for info in full_infos]
        visible_before = self._last_visible_hwnds

        if len(visible_before) != len(new_visible_order):
            return  # selection drifted mid-drag -- ignore defensively

        try:
            indices = [full_hwnds.index(h) for h in visible_before]
        except ValueError:
            return

        for idx, new_hwnd in zip(indices, new_visible_order):
            full_hwnds[idx] = new_hwnd

        self._manual_order_hwnds = full_hwnds
        self._update_preview_and_info()

    # ------------------------------------------------------------------
    # Window Cycling (Feature 4)
    # ------------------------------------------------------------------
    def _cycle_next_manual(self):
        full = self._get_checked_window_infos_in_order()
        if not full:
            return
        self._cycle_offset = (self._cycle_offset + 1) % len(full)
        self._apply_cycle_rotation_to_real_windows()
        self._update_preview_and_info()

    def _toggle_auto_cycle(self):
        if self.cycle_auto_button.isChecked():
            self._cycle_timer.start()
            self.status_bar.showMessage(
                "Auto-cycling overflow windows into view every 5 seconds."
            )
        else:
            self._cycle_timer.stop()

    def _on_cycle_tick(self):
        self._cycle_next_manual()

    def _apply_cycle_rotation_to_real_windows(self):
        """
        Quietly repositions whatever windows just rotated INTO a
        visible slot onto that slot's rectangle, and minimizes whatever
        windows just rotated OUT of view -- without the aggressive
        TOPMOST foreground-forcing pass Apply Equal Layout uses, since
        a background rotation stealing focus every few seconds would be
        far more annoying than helpful. A no-op until a layout has
        actually been applied at least once and there is genuine
        overflow to rotate through.
        """
        if self._selected_monitor is None:
            return

        new_visible = self._get_visible_window_infos()
        new_visible_hwnds = {info["hwnd"] for info in new_visible}

        rotated_out = self._currently_tiled_hwnds - new_visible_hwnds
        for hwnd in rotated_out:
            try:
                if win32gui.IsWindow(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            except Exception:
                pass

        if new_visible:
            work_area = self._selected_monitor.work_rect
            rects = self.preview_canvas.get_monitor_space_tile_rects(
                work_area, len(new_visible)
            )
            for info, (x, y, w, h) in zip(new_visible, rects):
                try:
                    if win32gui.IsWindow(info["hwnd"]):
                        restore_if_minimized_or_maximized(info["hwnd"])
                        apply_window_rect(info["hwnd"], x, y, w, h)
                except Exception:
                    pass

        self._currently_tiled_hwnds = new_visible_hwnds

    # ------------------------------------------------------------------
    def _update_preview_and_info(self):
        visible_infos = self._get_visible_window_infos()
        self.preview_canvas.set_checked_windows(visible_infos)
        self._update_info_panel()

    def _update_info_panel(self):
        full_infos = self._get_checked_window_infos_in_order()
        n_total = len(full_infos)
        n_visible = len(self._last_visible_hwnds)
        template_mode = self.preview_canvas.get_template_mode()

        self._info_rows["selected"].setText(str(n_total))

        if n_visible == 0:
            self._info_rows["grid"].setText("—")
            self.grid_size_label.setText("0 x 0")
        elif template_mode == "equal":
            rows, cols = calculate_grid_dimensions(n_visible)
            self._info_rows["grid"].setText("{} x {}".format(rows, cols))
            self.grid_size_label.setText("{} x {}".format(rows, cols))
        else:
            label = TEMPLATE_DISPLAY_LABELS.get(template_mode, template_mode)
            self._info_rows["grid"].setText(label)
            self.grid_size_label.setText(label)

        if self._selected_monitor is not None:
            self._info_rows["monitor"].setText(self._selected_monitor.label)
            self._info_rows["resolution"].setText(
                self._selected_monitor.resolution_label
            )
            wl, wt, wr, wb = self._selected_monitor.work_rect
            self._info_rows["workarea"].setText(
                "{} x {}".format(wr - wl, wb - wt)
            )
        else:
            self._info_rows["monitor"].setText("—")
            self._info_rows["resolution"].setText("—")
            self._info_rows["workarea"].setText("—")

        if n_visible and template_mode == "equal":
            rows, cols = calculate_grid_dimensions(n_visible)
            usage_pct = min(100, int(round((n_visible / float(rows * cols)) * 100)))
            usage_text = "{}%".format(usage_pct)
        elif n_visible:
            usage_text = "100%"
        else:
            usage_text = "—"
        self._info_rows["usage"].setText(usage_text)
        self.usage_label.setText(
            "Cells used: {}".format(usage_text) if n_visible else ""
        )

    # ------------------------------------------------------------------
    # Apply Equal Layout
    # ------------------------------------------------------------------
    def apply_equal_layout(self):
        # FEATURE 4: only the currently VISIBLE (post cycle-capping)
        # subset of checked windows is ever tiled in one Apply pass --
        # any overflow stays queued and reachable via Cycle Next/Auto.
        checked_infos = self._get_visible_window_infos()

        if not checked_infos:
            QMessageBox.information(
                self, "No Windows Selected",
                "Please check at least one window in the list before "
                "applying a layout.",
            )
            return

        if self._selected_monitor is None:
            QMessageBox.warning(
                self, "No Monitor Selected",
                "No monitor is currently selected. Click \"Detect Windows\" "
                "to re-scan monitors and try again.",
            )
            return

        # For Custom Grid, warn the user if cells would be smaller than
        # the minimum window size. Unlike the preview warning label (which
        # updates live), this is an explicit confirmation so the user
        # can't accidentally apply a layout that will cause windows to
        # overlap because the OS refuses to resize them below their minimum.
        if self.preview_canvas.get_template_mode() == "custom_grid":
            work_area = self._selected_monitor.work_rect
            rows = self.grid_rows_spin.value()
            cols = self.grid_cols_spin.value()
            fits, cell_w, cell_h, safe_cols, safe_rows = check_grid_fits_min_size(
                work_area, 0, rows, cols
            )
            if not fits:
                problems = []
                if cell_w < DEFAULT_MIN_WIN_W:
                    problems.append(
                        "width {}px (minimum is {}px; try ≤{} columns)".format(
                            cell_w, DEFAULT_MIN_WIN_W, safe_cols
                        )
                    )
                if cell_h < DEFAULT_MIN_WIN_H:
                    problems.append(
                        "height {}px (minimum is {}px; try ≤{} rows)".format(
                            cell_h, DEFAULT_MIN_WIN_H, safe_rows
                        )
                    )
                reply = QMessageBox.warning(
                    self,
                    "Cell Too Small — Windows May Overlap",
                    "The {}×{} grid would create cells that are too small:\n\n"
                    "  • {}\n\n"
                    "Telegram and other apps will refuse to resize below their "
                    "minimum dimensions, causing windows to overlap.\n\n"
                    "Apply anyway?".format(
                        rows, cols, "\n  • ".join(problems)
                    ),
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if reply != QMessageBox.Yes:
                    return

        # Re-validate liveness right before moving anything -- a window
        # may have closed between the last refresh and this click.
        live_infos = []
        for info in checked_infos:
            try:
                if win32gui.IsWindow(info["hwnd"]):
                    live_infos.append(info)
            except Exception:
                continue

        if not live_infos:
            QMessageBox.warning(
                self, "Windows No Longer Available",
                "None of the previously checked windows are still open. "
                "Click \"Refresh Windows\" and try again.",
            )
            self.refresh_window_list()
            return

        skipped_count = len(checked_infos) - len(live_infos)

        # Snapshot current positions BEFORE moving anything, so "Reset
        # Positions" can put every tiled window back exactly where it
        # was. Only windows we are about to move are snapshotted, and
        # only HWNDs we actually touch are ever recorded here -- no
        # other open window is affected in any way.
        for info in live_infos:
            rect = get_window_rect_safe(info["hwnd"])
            if rect is not None:
                self._pre_layout_rects[info["hwnd"]] = rect

        try:
            work_area = self._selected_monitor.work_rect

            # FEATURE 3: the active template (Equal Grid / Master+Stack
            # / Columns / Vertical Stack) -- including whatever divider
            # ratios the user has manually dragged and the current gap
            # setting -- decides the real, final rectangles. This is
            # the exact same call the preview canvas itself uses to
            # draw, so what was shown is precisely what gets applied.
            end_rects = self.preview_canvas.get_monitor_space_tile_rects(
                work_area, len(live_infos)
            )

            # Safety net: guarantee zero overlap no matter what produced
            # end_rects (see ensure_no_overlap() docstring). Under
            # normal operation every template already produces
            # non-overlapping rectangles, so this is a no-op pass-
            # through in practice.
            end_rects = ensure_no_overlap(end_rects, work_area=work_area)

            moves = []
            for info, end_rect in zip(live_infos, end_rects):
                hwnd = info["hwnd"]
                try:
                    restore_if_minimized_or_maximized(hwnd)
                except Exception:
                    pass
                start_rect_raw = get_window_rect_safe(hwnd)
                if start_rect_raw is not None:
                    sl, st, sr, sb = start_rect_raw
                    start_rect = (sl, st, sr - sl, sb - st)
                else:
                    start_rect = end_rect
                moves.append((hwnd, start_rect, end_rect))

            self._currently_tiled_hwnds = {hwnd for hwnd, _s, _e in moves}

            def _on_animation_complete():
                self._finish_apply_equal_layout(moves, skipped_count)

            # FEATURE 6: instead of an instantaneous SWP snap, every
            # window glides from wherever it currently is to its new
            # grid cell over a short, eased animation -- see
            # WindowMoveAnimator (Section 5B). This runs entirely on a
            # non-blocking QTimer, so the UI thread never freezes. The
            # strict restore -> TOPMOST -> NOTOPMOST -> SetForeground-
            # Window z-order fix (force_window_to_foreground) still
            # runs exactly as before, just deferred until the animation
            # settles -- see _finish_apply_equal_layout() below.
            self._window_animator.animate(
                moves, duration_ms=260, on_complete=_on_animation_complete
            )

        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(
                self, "Layout Error",
                "An error occurred while applying the layout:\n{}".format(exc),
            )

    def _finish_apply_equal_layout(self, moves, skipped_count):
        """
        Runs once the Feature 6 grid-transition animation has settled
        every window into its final rectangle. Performs the original
        strict restore -> HWND_TOPMOST -> HWND_NOTOPMOST ->
        SetForegroundWindow sequence (force_window_to_foreground) at
        each window's already-final rectangle -- i.e. this pass never
        moves anything further, it only fixes z-order/focus -- then
        reports the result and auto-minimizes the manager window exactly
        as the original, non-animated Apply Equal Layout always did.
        """
        move_failures = 0
        for hwnd, _start_rect, end_rect in moves:
            try:
                x, y, w, h = end_rect
                force_window_to_foreground(hwnd, x, y, w, h)
            except Exception as inner_exc:
                move_failures += 1
                print(
                    "Warning: failed to move/resize HWND {}: {}".format(hwnd, inner_exc)
                )

        n = len(moves)
        template_mode = self.preview_canvas.get_template_mode()
        if template_mode == "equal":
            rows, cols = calculate_grid_dimensions(n)
            message = "Tiled {} window{} on {} in a {}x{} grid.".format(
                n, "" if n == 1 else "s", self._selected_monitor.label, rows, cols,
            )
        else:
            template_label = TEMPLATE_DISPLAY_LABELS.get(template_mode, template_mode)
            message = "Tiled {} window{} on {} using the {} template.".format(
                n, "" if n == 1 else "s", self._selected_monitor.label, template_label,
            )
        if skipped_count:
            message += "  ({} no longer open, skipped.)".format(skipped_count)
        if move_failures:
            message += "  ({} could not be moved -- may be a protected window.)".format(
                move_failures
            )
        self.status_bar.showMessage(message)

        # UX refinement: once at least one window has actually been
        # tiled, get this manager window itself out of the way by
        # auto-minimizing it to the taskbar, so it doesn't sit on top of
        # the workspace it just organized. Skipped entirely if every
        # move failed, so the user can still see what went wrong.
        if move_failures < len(moves):
            QTimer.singleShot(150, self.showMinimized)

        # Give everything a brief moment to finish settling, then
        # refresh the list's minimized/restored status dots and the
        # live preview thumbnails to reflect the new state.
        QTimer.singleShot(250, self._update_preview_and_info)

    # ------------------------------------------------------------------
    # Save Layout
    # ------------------------------------------------------------------
    def save_current_layout(self):
        checked_infos = self._get_checked_window_infos_in_order()
        if not checked_infos:
            QMessageBox.information(
                self, "Nothing To Save",
                "Check at least one window before saving a layout, so "
                "there is something to record positions for.",
            )
            return

        default_name = "Layout_{}".format(time.strftime("%Y%m%d_%H%M%S"))
        name, ok = QInputDialog.getText(
            self, "Save Layout", "Layout name:", QLineEdit.Normal, default_name
        )
        if not ok or not name.strip():
            return

        try:
            file_path = save_layout_to_disk(name.strip(), checked_infos)
            self.status_bar.showMessage(
                "Saved layout \"{}\" ({} windows) to {}".format(
                    name.strip(), len(checked_infos), file_path
                )
            )
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(
                self, "Save Failed",
                "Could not save the layout:\n{}".format(exc),
            )

    # ------------------------------------------------------------------
    # Load Layout
    # ------------------------------------------------------------------
    def load_saved_layout(self):
        layout_names = list_saved_layouts()
        if not layout_names:
            QMessageBox.information(
                self, "No Saved Layouts",
                "There are no saved layouts yet. Use \"Save Layout\" first.",
            )
            return

        name, ok = QInputDialog.getItem(
            self, "Load Layout", "Choose a saved layout:",
            layout_names, 0, False,
        )
        if not ok or not name:
            return

        payload = load_layout_from_disk(name)
        if payload is None:
            QMessageBox.warning(
                self, "Load Failed",
                "Could not read the layout file for \"{}\".".format(name),
            )
            return

        # Always match against a FRESH enumeration so we are positioning
        # against exactly what is open right now.
        try:
            current_windows = enumerate_taskbar_windows(own_hwnd=self._own_hwnd)
        except Exception as exc:
            QMessageBox.critical(
                self, "Enumeration Error",
                "Failed to enumerate open windows:\n{}".format(exc),
            )
            return

        for info in current_windows:
            rect = get_window_rect_safe(info["hwnd"])
            if rect is not None:
                self._pre_layout_rects[info["hwnd"]] = rect

        applied, skipped = apply_saved_layout(payload, current_windows)

        self.refresh_window_list()

        message = "Loaded \"{}\": repositioned {} window{}.".format(
            name, applied, "" if applied == 1 else "s"
        )
        if skipped:
            message += "  ({} from the saved layout were not found among " \
                       "currently open windows.)".format(skipped)
        self.status_bar.showMessage(message)

    # ------------------------------------------------------------------
    # Reset Positions
    # ------------------------------------------------------------------
    def reset_positions(self):
        if not self._pre_layout_rects:
            QMessageBox.information(
                self, "Nothing To Reset",
                "No layout has been applied yet in this session, so "
                "there are no prior positions to restore.",
            )
            return

        restored = 0
        missing = 0
        for hwnd, (left, top, right, bottom) in list(self._pre_layout_rects.items()):
            try:
                if not win32gui.IsWindow(hwnd):
                    missing += 1
                    continue
                restore_if_minimized_or_maximized(hwnd)
                apply_window_rect(hwnd, left, top, right - left, bottom - top)
                restored += 1
            except Exception:
                missing += 1

        self._pre_layout_rects.clear()

        message = "Restored {} window{} to their previous position{}.".format(
            restored, "" if restored == 1 else "s", "" if restored == 1 else "s"
        )
        if missing:
            message += "  ({} no longer available.)".format(missing)
        self.status_bar.showMessage(message)

        QTimer.singleShot(250, self._update_preview_and_info)


# ======================================================================
# SECTION 13: Entry point
# ======================================================================

def main():
    if not sys.platform.startswith("win"):
        print("This application only runs on Windows.", file=sys.stderr)
        sys.exit(1)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = WindowLayoutManagerApp()

    # FEATURE 2: install the global-hotkey native event filter at the
    # QApplication level so Ctrl+Alt+L reaches apply_equal_layout() no
    # matter which window currently has OS focus, including while this
    # manager window itself is minimized. PyQt5 does not keep its own
    # strong reference to installed native event filters, so we hold
    # one explicitly on both `app` and `window` for the whole process
    # lifetime to guarantee it is never garbage-collected mid-session.
    hotkey_filter = GlobalHotkeyEventFilter(window.apply_equal_layout)
    app.installNativeEventFilter(hotkey_filter)
    app._wlm_hotkey_filter_ref = hotkey_filter
    window._hotkey_filter_ref = hotkey_filter

    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
