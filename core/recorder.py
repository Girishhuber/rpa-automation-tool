from __future__ import annotations

import ctypes
import ctypes.wintypes
import platform
import re
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from utils.logger import logger
from utils.errors import RecorderError, HookInstallError
from models.session import Session, SessionStatus, SystemInfo
from models.event import (
    Event,
    MouseClickEvent, MouseDoubleClickEvent, MouseRightClickEvent,
    MouseMiddleClickEvent, MouseScrollEvent, MouseDragEvent,
    KeyPressEvent, KeyComboEvent, TypeTextEvent,
    ClipboardCopyEvent, ClipboardCutEvent, ClipboardPasteEvent,
    WindowFocusEvent,
    ExcelCellSelectEvent, ExcelRangeSelectEvent, ExcelSheetSwitchEvent,
    BrowserNavigateEvent,
    ScreenshotCheckpointEvent,
    ProcessLaunchEvent,
)
from models.target import UITarget, TargetBackend
from .uia_enricher import UIAEnricher, detect_excel_cell, BROWSER_PROCS, _ELECTRON_CLASS
from .browser_bridge import BrowserBridge
from .overlay import RecordingOverlay
from .screenshot import ScreenCapture

try:
    from pynput import mouse, keyboard
    PYNPUT_OK = True
except ImportError:
    PYNPUT_OK = False
    logger.warning("[RECORD] pynput not installed - cannot record")

try:
    import win32clipboard, win32con
    WIN32_OK = True
except ImportError:
    WIN32_OK = False

try:
    from pywinauto import Application
    from pywinauto.findwindows import find_windows
    try:
        from pywinauto import Desktop
    except ImportError:
        from pywinauto import Desktop
    UIA_OK = True
except ImportError:
    UIA_OK = False

try:
    import psutil
    PSUTIL_OK = True
except ImportError:
    PSUTIL_OK = False


_MOD_NORMALIZE = {
    "ctrl_l": "ctrl", "ctrl_r": "ctrl",
    "shift_l": "shift", "shift_r": "shift",
    "alt_l": "alt", "alt_r": "alt", "alt_gr": "alt",
    "cmd": "cmd", "cmd_r": "cmd",
    "ctrl": "ctrl", "shift": "shift", "alt": "alt",
}

_SPECIAL_KEYS = {
    "enter", "return", "tab", "backspace", "delete", "escape", "insert",
    "home", "end", "page_up", "page_down", "space",
    "left", "right", "up", "down",
    "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8",
    "f9", "f10", "f11", "f12",
}

_CLIPBOARD_COMBOS = {
    frozenset({"ctrl", "c"}): "copy",
    frozenset({"ctrl", "x"}): "cut",
    frozenset({"ctrl", "v"}): "paste",
}

_EXCEL_PROCS = {"excel.exe"}
_CELL_RE = re.compile(r"^([A-Z]{1,3})([0-9]{1,7})$")
_RANGE_RE = re.compile(r"^([A-Z]{1,3}[0-9]{1,7}):([A-Z]{1,3}[0-9]{1,7})$")
_SHEET_CELL_RE = re.compile(r"^(.+)!([A-Z]{1,3}[0-9]{1,7})$")

_EXCEL_CONFIRM_KEYS = {"enter", "return", "tab", "escape"}
_EXCEL_NAV_KEYS = {
    "left", "right", "up", "down", "page_up", "page_down",
    "home", "end",
}

_SYSTEM_UI_WINDOWS = {"Taskbar", "Search", "Action center", "Start"}
_SYSTEM_UI_PROCS = {
    "explorer.exe", "searchhost.exe", "searchapp.exe",
    "shellexperiencehost.exe", "startmenuexperiencehost.exe",
}

_LAUNCH_HINTS = {
    "excel": "EXCEL.EXE",
    "word": "WINWORD.EXE",
    "powerpoint": "POWERPNT.EXE",
    "ppt": "POWERPNT.EXE",
    "chrome": "chrome.exe",
    "edge": "msedge.exe",
    "notepad": "notepad.exe",
    "outlook": "OUTLOOK.EXE",
}

_EDITABLE_CONTROL_TYPES = {
    "Edit", "Document", "DataItem", "SpreadsheetItem", "Cell",
    "RichEdit", "Text", "TextBox", "ComboBox",
}

_NON_EDITABLE_CONTROL_TYPES = {
    "Button", "SplitButton", "MenuItem", "TabItem", "ListItem",
    "TreeItem", "Pane", "ToolBar", "StatusBar", "ScrollBar",
    "TitleBar", "MenuBar", "Menu",
}

_TRANSIENT_TYPES = {"ToolTip", "Notification", "Toast", "Popup"}
_TRANSIENT_TITLE_RE = re.compile(
    r"(tooltip|notification|toast|snackbar|bubble)", re.IGNORECASE
)

_APP_ACTION_GROUPS = {
    "winword.exe": "word",
    "excel.exe": "excel",
    "powerpnt.exe": "powerpoint",
    "chrome.exe": "browser",
    "msedge.exe": "browser",
    "explorer.exe": "explorer",
}

SCROLL_MERGE_WINDOW_MS = 1500


def _classify_intent(payload) -> str:
    ptype = str(getattr(payload, "type", ""))
    if "navigate" in ptype or "browser_back" in ptype or "process_launch" in ptype:
        return "navigation"
    if "clipboard" in ptype or "copy" in ptype or "paste" in ptype:
        return "clipboard"
    if "type" in ptype or "key" in ptype:
        return "input"
    if "click" in ptype or "drag" in ptype:
        return "selection"
    if "scroll" in ptype:
        return "scroll"
    if "window_focus" in ptype:
        return "system"
    if "excel" in ptype:
        return "input"
    if "screenshot" in ptype or "wait" in ptype:
        return "checkpoint"
    return "system"


def _excel_get_name_box_value(hwnd: int) -> Optional[str]:
    if not UIA_OK:
        return None
    try:
        app = Application(backend="uia").connect(handle=hwnd)
        win = app.window(handle=hwnd)

        try:
            elem = win.child_window(auto_id="Box", control_type="Edit")
            if elem.exists(timeout=0.3):
                val = elem.wrapper_object().window_text() or ""
                if val:
                    return val.strip()
        except Exception:
            pass

        try:
            descs = win.descendants(auto_id="Box")
            for d in descs[:3]:
                try:
                    wrapper = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                    val = wrapper.window_text() or ""
                    clean = val.strip()
                    if val and (
                        _CELL_RE.match(clean.upper())
                        or _RANGE_RE.match(clean.upper())
                        or _SHEET_CELL_RE.match(clean)
                    ):
                        return clean
                except Exception:
                    continue
        except Exception:
            pass

        try:
            descs = win.descendants(class_name="NameBox")
            for d in descs[:2]:
                try:
                    wrapper = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                    val = wrapper.window_text() or ""
                    if val:
                        return val.strip()
                except Exception:
                    continue
        except Exception:
            pass
    except Exception:
        pass
    return None


def _excel_get_active_sheet(hwnd: int) -> Optional[str]:
    if not UIA_OK:
        return None
    try:
        app = Application(backend="uia").connect(handle=hwnd)
        win = app.window(handle=hwnd)

        try:
            tabs = win.descendants(control_type="TabItem")
            for tab in tabs[:20]:
                try:
                    wrapper = tab.wrapper_object() if hasattr(tab, "wrapper_object") else tab

                    try:
                        state = wrapper.get_toggle_state()
                        if state == 1:
                            return wrapper.window_text() or None
                    except Exception:
                        pass

                    try:
                        if wrapper.is_selected():
                            return wrapper.window_text() or None
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception:
            pass
    except Exception:
        pass
    return None


def _excel_get_cell_value(hwnd: int, cell_ref: str) -> Optional[str]:
    if not UIA_OK:
        return None
    try:
        app = Application(backend="uia").connect(handle=hwnd)
        win = app.window(handle=hwnd)

        for aid in ("FormulaBar", "formulaBar"):
            try:
                fb = win.child_window(auto_id=aid)
                if fb.exists(timeout=0.2):
                    val = fb.wrapper_object().window_text() or ""
                    return val if val else None
            except Exception:
                pass

        try:
            descs = win.descendants(class_name="EXCEL71")
            for d in descs[:2]:
                try:
                    wrapper = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                    val = wrapper.window_text() or ""
                    if val:
                        return val
                except Exception:
                    continue
        except Exception:
            pass
    except Exception:
        pass
    return None


def _excel_find_hwnd() -> Optional[int]:
    if not UIA_OK:
        return None
    try:
        handles = find_windows(title_re=".*Microsoft Excel.*")
        if not handles:
            handles = find_windows(class_name="XLMAIN")
        return handles[0] if handles else None
    except Exception:
        return None


def _is_excel_process(proc_name: Optional[str]) -> bool:
    return (proc_name or "").lower() in _EXCEL_PROCS


def _parse_cell_ref(text: str) -> Optional[str]:
    if not text:
        return None

    raw = text.strip()
    clean = raw.upper()

    m = _SHEET_CELL_RE.match(raw)
    if m:
        return m.group(2).upper()

    if _CELL_RE.match(clean):
        return clean

    return None


def _parse_range_ref(text: str) -> Optional[str]:
    if not text:
        return None

    clean = text.strip().upper()
    if _RANGE_RE.match(clean):
        return clean

    return None


class Recorder:
    _DOUBLE_CLICK_MS = 400
    _DRAG_THRESHOLD = 8
    _WINDOW_DEBOUNCE = 350
    _CLICK_DEBOUNCE = 300
    _FOCUS_SAMPLE_DELAY_MS = 150

    def __init__(self, config, sessions_dir: Path, session_name: str = "Untitled"):
        self._config = config
        self._sessions_dir = sessions_dir
        self._name = session_name

        self._enricher = UIAEnricher()
        self._browser = BrowserBridge(config.recorder.browser_cdp_port)
        self._overlay: Optional[RecordingOverlay] = None
        self._capture: Optional[ScreenCapture] = None

        self._session: Optional[Session] = None
        self._events: list[Event] = []
        self._lock = threading.Lock()
        self._running = False
        self._event_id = 0
        self._start_ms = 0

        self._mouse_down_pos: Optional[tuple[int, int]] = None
        self._mouse_down_time: float = 0.0
        self._last_click_pos: tuple[int, int] = (0, 0)
        self._last_click_time: float = 0.0
        self._last_click_target_key: str = ""
        self._last_click_target_ms: int = 0

        self._pressed_mods: set[str] = set()
        self._text_buffer: str = ""
        self._last_key_time: float = 0.0
        self._text_buffer_target: Optional[UITarget] = None
        self._text_buffer_excel_context: Optional[dict] = None
        self._text_buffer_in_excel: bool = False

        self._last_target: Optional[UITarget] = None
        self._last_typing_target: Optional[UITarget] = None

        self._last_scroll_ms: int = 0
        self._pending_scroll: Optional[dict] = None
        self._last_scroll_dir: int = 0

        self._search_mode: bool = False
        self._current_action_group: Optional[str] = None

        self._pending_hwnd: int = 0
        self._pending_time: float = 0.0
        self._emitted_hwnd: int = 0

        self._excel_hwnd: Optional[int] = None
        self._excel_active_sheet: Optional[str] = None
        self._excel_active_cell: Optional[str] = None
        self._excel_in_edit_mode: bool = False
        self._excel_context_lock = threading.Lock()

        self._excel_shift_anchor: Optional[str] = None
        self._excel_namebox_mode: bool = False

        # Source of truth for Excel typing. This is written by clicks and used
        # before polling is trusted.
        self._excel_clicked_cell: Optional[str] = None
        self._excel_clicked_sheet: Optional[str] = None
        self._excel_clicked_target: Optional[UITarget] = None
        self._excel_current_edit_key: Optional[str] = None
        self._excel_cell_source: Optional[str] = None  


        self._mouse_listener = None
        self._kbd_listener = None
        self._clipboard_thread: Optional[threading.Thread] = None
        self._window_thread: Optional[threading.Thread] = None
        self._excel_thread: Optional[threading.Thread] = None
        self._last_clipboard: str = ""

    def start(self, session_name: Optional[str] = None) -> Session:
        if self._running:
            raise RecorderError("Already recording")
        if not PYNPUT_OK:
            raise HookInstallError("pynput is required")

        self._name = session_name or self._name
        self._session = self._make_session()
        self._events = []
        self._event_id = 0
        self._start_ms = self._now_ms()

        scr_dir = self._sessions_dir / self._session.id / "screenshots"
        self._capture = ScreenCapture(scr_dir)

        browser_ok = self._browser.connect()
        logger.info(
            "[RECORD] Session started: id={} name='{}' browser={}",
            self._session.id,
            self._name,
            "connected" if browser_ok else "NOT connected",
        )

        self._overlay = RecordingOverlay(self._config)
        self._overlay.start()
        self._overlay.set_recording(True)

        self._clipboard_thread = threading.Thread(target=self._monitor_clipboard, daemon=True)
        self._clipboard_thread.start()

        self._window_thread = threading.Thread(target=self._monitor_windows, daemon=True)
        self._window_thread.start()

        self._excel_thread = threading.Thread(target=self._monitor_excel_context, daemon=True)
        self._excel_thread.start()

        self._install_hooks()
        self._running = True
        self._overlay.log_event("Recording started")
        return self._session

    def stop(self) -> Session:
        if not self._running:
            raise RecorderError("Not recording")

        self._flush_pending_scroll()
        self._flush_text_buffer()
        self._running = False
        self._remove_hooks()
        self._browser.disconnect()

        if self._overlay:
            self._overlay.set_recording(False)
            self._overlay.log_event(f"Stopped - {len(self._events)} events")
            time.sleep(0.4)
            self._overlay.stop()

        duration = self._now_ms() - self._start_ms

        with self._lock:
            self._session.events = [e.model_dump(mode="json") for e in self._events]

        self._session.duration_ms = duration
        self._session.status = SessionStatus.COMPLETE
        self._session.updated_at = datetime.now(timezone.utc).isoformat()

        logger.info("[RECORD] Stopped: {} events in {:.1f}s", len(self._events), duration / 1000)
        return self._session

    def add_checkpoint(self) -> None:
        if not self._running or not self._capture:
            return

        path = self._capture.capture_full(0)
        if path:
            self._push(ScreenshotCheckpointEvent(path=str(path.name), label="manual"), "Checkpoint")

    @property
    def is_recording(self) -> bool:
        return self._running

    @property
    def event_count(self) -> int:
        return len(self._events)

    def _install_hooks(self) -> None:
        self._mouse_listener = mouse.Listener(
            on_move=self._on_mouse_move,
            on_click=self._on_mouse_click,
            on_scroll=self._on_mouse_scroll,
        )
        self._kbd_listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release,
        )
        self._mouse_listener.start()
        self._kbd_listener.start()
        logger.info("[RECORD] Input hooks installed")

    def _remove_hooks(self) -> None:
        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._kbd_listener:
            self._kbd_listener.stop()

    def _monitor_excel_context(self) -> None:
        while self._running:
            try:
                time.sleep(0.5)

                hwnd = self._excel_hwnd
                if not hwnd:
                    hwnd = _excel_find_hwnd()
                    if hwnd:
                        with self._excel_context_lock:
                            self._excel_hwnd = hwnd

                if not hwnd:
                    continue

                sheet = _excel_get_active_sheet(hwnd)
                if sheet:
                    with self._excel_context_lock:
                        self._excel_active_sheet = sheet
                        if self._excel_clicked_cell and not self._excel_clicked_sheet:
                            self._excel_clicked_sheet = sheet

                if self._text_buffer or self._excel_in_edit_mode or self._excel_clicked_cell:
                    continue

                name_val = _excel_get_name_box_value(hwnd)
                if not name_val:
                    continue

                parsed_range = _parse_range_ref(name_val)
                parsed_cell = _parse_cell_ref(name_val)

                with self._excel_context_lock:
                    if parsed_range:
                        self._excel_active_cell = parsed_range.split(":", 1)[0]
                        self._excel_cell_source = "poll"
                    elif parsed_cell:
                        self._excel_active_cell = parsed_cell
                        self._excel_cell_source = "poll"
                    elif "!" in name_val:
                        parts = name_val.strip().split("!", 1)
                        if len(parts) == 2:
                            sheet_name, cell = parts
                            parsed = _parse_cell_ref(cell)
                            if parsed:
                                self._excel_active_cell = parsed
                                self._excel_active_sheet = sheet_name.strip("'")
                                self._excel_cell_source = "poll"

            except Exception as exc:
                logger.debug("[RECORD] excel_context thread: {}", exc)
                time.sleep(1.0)


    def _get_excel_context(self) -> dict:
        with self._excel_context_lock:
            return {
                "sheet_name": self._excel_active_sheet,
                "cell_ref": self._excel_active_cell,
            }

    def _lock_excel_cell_from_click(
    self,
    cell_ref: str,
    sheet_name: Optional[str],
    target: Optional[UITarget],
) -> None:
        cell_ref = cell_ref.upper()

        with self._excel_context_lock:
            self._excel_active_cell = cell_ref
            if sheet_name:
                self._excel_active_sheet = sheet_name

            self._excel_clicked_cell = cell_ref
            self._excel_clicked_sheet = sheet_name or self._excel_active_sheet
            self._excel_clicked_target = target
            self._excel_in_edit_mode = False
            self._excel_cell_source = "click"

        self._excel_current_edit_key = None


    def _clear_clicked_excel_cell(self) -> None:
        self._excel_clicked_cell = None
        self._excel_clicked_sheet = None
        self._excel_clicked_target = None
        self._excel_current_edit_key = None

        with self._excel_context_lock:
            if self._excel_cell_source == "click":
                self._excel_cell_source = None


    def _get_excel_typing_context(self) -> dict:
        with self._excel_context_lock:
            if self._excel_clicked_cell:
                return {
                    "cell_ref": self._excel_clicked_cell,
                    "sheet_name": self._excel_clicked_sheet or self._excel_active_sheet,
                    "target": self._excel_clicked_target,
                    "source": "click",
                }

            if self._excel_cell_source in {"click", "namebox", "nav"} and self._excel_active_cell:
                return {
                    "cell_ref": self._excel_active_cell,
                    "sheet_name": self._excel_active_sheet,
                    "target": None,
                    "source": self._excel_cell_source,
                }

        return {
            "cell_ref": None,
            "sheet_name": None,
            "target": None,
            "source": None,
        }


    def _on_mouse_move(self, x: int, y: int) -> None:
        pass

    def _on_mouse_click(self, x: int, y: int, button, pressed: bool) -> None:
        if not self._running:
            return

        btn = button.name if hasattr(button, "name") else str(button)

        if pressed:
            self._mouse_down_pos = (x, y)
            self._mouse_down_time = time.perf_counter()
            return

        now = time.perf_counter()
        start_pos = self._mouse_down_pos
        self._mouse_down_pos = None

        if start_pos:
            dx = abs(x - start_pos[0])
            dy = abs(y - start_pos[1])
            is_drag = (dx + dy) > self._DRAG_THRESHOLD
        else:
            is_drag = True

        self._flush_pending_scroll()

        if is_drag and start_pos:
            self._flush_text_buffer()

            t_start = self._enricher.get_target_at(*start_pos)
            t_end = self._enricher.get_target_at(x, y)

            if t_start and _is_excel_process(t_start.process_name):
                start_cell = _parse_cell_ref(t_start.name or "")
                end_cell = _parse_cell_ref(t_end.name or "") if t_end else None

                if start_cell and end_cell and start_cell != end_cell:
                    ctx = self._get_excel_context()
                    range_ref = f"{start_cell}:{end_cell}"
                    logger.info("[RECORD] Excel range drag: {}", range_ref)
                    self._push(
                        ExcelRangeSelectEvent(
                            range_ref=range_ref,
                            sheet_name=ctx.get("sheet_name"),
                        ),
                        f"Excel range: {range_ref}",
                    )
                    return

            self._push(
                MouseDragEvent(
                    start_x=start_pos[0],
                    start_y=start_pos[1],
                    end_x=x,
                    end_y=y,
                    button=btn,
                    start_target=t_start,
                    end_target=t_end,
                ),
                f"Drag ({start_pos[0]},{start_pos[1]})->({x},{y})",
            )
            return

        target = self._build_target_at(x, y)

        # Flush previous typed text before recording this new click. This is safe
        # because Excel typing context is locked from the previous cell click.
        self._flush_text_buffer()

        if target and self._is_transient(target):
            return

        if target and self._is_system_ui(target):
            self._handle_system_ui_click(target, x, y, btn)
            return

        if target and _is_excel_process(target.process_name) and target.control_type == "TabItem":
            sheet_name = target.name or ""
            if sheet_name:
                idx = self._excel_get_sheet_index(sheet_name, target)
                logger.info("[RECORD] Excel sheet tab click: '{}'", sheet_name)
                with self._excel_context_lock:
                    self._excel_active_sheet = sheet_name
                self._clear_clicked_excel_cell()
                self._push(
                    ExcelSheetSwitchEvent(sheet_name=sheet_name, sheet_index=idx),
                    f"Excel sheet: {sheet_name}",
                )
                self._sample_focused_element()
                return

        self._search_mode = False
        self._last_target = target

        tkey = self._target_key(target)
        now_ms = self._now_ms()
        if (
            tkey
            and tkey == self._last_click_target_key
            and (now_ms - self._last_click_target_ms) < self._CLICK_DEBOUNCE
        ):
            return

        self._last_click_target_key = tkey
        self._last_click_target_ms = now_ms

        dx2 = abs(x - self._last_click_pos[0])
        dy2 = abs(y - self._last_click_pos[1])
        if (
            dx2 < 6
            and dy2 < 6
            and (now - self._last_click_time) * 1000 < self._DOUBLE_CLICK_MS
            and btn == "left"
        ):
            self._last_click_time = 0.0
            self._push(
                MouseDoubleClickEvent(x=x, y=y, target=target),
                f"Dbl-click {self._tlabel(target)}",
            )
            self._maybe_screenshot(target)
            self._sample_focused_element()
            return

        self._last_click_pos = (x, y)
        self._last_click_time = now

        if self._overlay:
            self._overlay.flash_click(x, y, is_replay=False)

        if target and _is_excel_process(target.process_name) and target.backend == TargetBackend.UIA:
            if self._try_record_excel_cell_click(target, x, y, btn):
                self._maybe_screenshot(target)
                return

        if btn == "right":
            self._push(
                MouseRightClickEvent(x=x, y=y, target=target),
                f"Right-click {self._tlabel(target)}",
            )
        elif btn == "middle":
            self._push(MouseMiddleClickEvent(x=x, y=y, target=target), "Middle-click")
        else:
            self._push(
                MouseClickEvent(x=x, y=y, button=btn, target=target),
                f"Click {self._tlabel(target)}",
            )
            self._sample_focused_element()

        self._maybe_screenshot(target)

    def _try_record_excel_cell_click(self, target: UITarget, x: int, y: int, btn: str) -> bool:
        ctrl = target.control_type or ""
        name = target.name or ""

        cell = detect_excel_cell(name, ctrl)
        if cell:
            ctx = self._get_excel_context()
            sheet_name = ctx.get("sheet_name")

            hwnd = self._excel_hwnd or _excel_find_hwnd()
            value = _excel_get_cell_value(hwnd, cell) if hwnd else None

            logger.info(
                "[RECORD] Excel cell click: {} sheet={} value={}",
                cell,
                sheet_name,
                (value or "")[:30],
            )

            excel_target = self._make_excel_cell_target(cell, sheet_name, target)
            self._lock_excel_cell_from_click(cell, sheet_name, excel_target)

            self._push(
                ExcelCellSelectEvent(
                    cell_ref=cell,
                    sheet_name=sheet_name,
                    target=excel_target,
                ),
                f"Excel cell: {cell}",
            )
            self._sample_focused_element()
            return True

        auto_id = target.automation_id or ""
        if (auto_id == "Box" or "name" in name.lower()) and ctrl == "Edit":
            logger.info("[RECORD] Excel Name Box click")
            self._excel_namebox_mode = True
            self._last_target = target
            return False

        if "formula" in name.lower() or ctrl in ("Edit", "Document"):
            logger.info("[RECORD] Excel formula bar click")
            self._excel_in_edit_mode = True
            self._last_target = target
            return False

        return False

    def _excel_get_sheet_index(self, sheet_name: str, target: UITarget) -> int:
        if not UIA_OK or not self._excel_hwnd:
            return 0

        try:
            app = Application(backend="uia").connect(handle=self._excel_hwnd)
            win = app.window(handle=self._excel_hwnd)
            tabs = win.descendants(control_type="TabItem")
            for i, tab in enumerate(tabs):
                try:
                    wrapper = tab.wrapper_object() if hasattr(tab, "wrapper_object") else tab
                    if wrapper.window_text() == sheet_name:
                        return i
                except Exception:
                    continue
        except Exception:
            pass

        return 0

    def _on_mouse_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        if not self._running or not self._config.recorder.capture_scroll:
            return

        now_ms = self._now_ms()
        direction = 1 if dy > 0 else -1

        if (
            self._pending_scroll is not None
            and direction == self._last_scroll_dir
            and (now_ms - self._last_scroll_ms) < SCROLL_MERGE_WINDOW_MS
        ):
            self._pending_scroll["dx"] += dx
            self._pending_scroll["dy"] += dy
            self._last_scroll_ms = now_ms
            return

        self._flush_pending_scroll()

        self._pending_scroll = {
            "x": x,
            "y": y,
            "dx": dx,
            "dy": dy,
            "target": self._enricher.get_target_at(x, y),
        }
        self._last_scroll_dir = direction
        self._last_scroll_ms = now_ms

        def _auto_flush():
            time.sleep(SCROLL_MERGE_WINDOW_MS / 1000)
            self._flush_pending_scroll()

        threading.Thread(target=_auto_flush, daemon=True).start()

    def _flush_pending_scroll(self) -> None:
        if not self._pending_scroll:
            return

        s = self._pending_scroll
        self._pending_scroll = None

        self._push(
            MouseScrollEvent(
                x=s["x"],
                y=s["y"],
                dx=s["dx"],
                dy=s["dy"],
                target=s["target"],
            ),
            f"Scroll {'down' if s['dy'] < 0 else 'up'} (merged)",
        )

    def _on_key_press(self, key) -> None:
        if not self._running:
            return

        raw = self._raw_key_str(key)
        canon = _MOD_NORMALIZE.get(raw)

        if canon:
            self._pressed_mods.add(canon)
            return

        char = self._printable_char(key)
        if char:
            self._append_text_char(char)
            return

        if raw in _SPECIAL_KEYS:
            if raw == "space" and not self._pressed_mods:
                self._append_text_char(" ")
                return

            self._flush_text_buffer()

            in_excel = self._is_in_excel_context()

            if in_excel and raw in _EXCEL_CONFIRM_KEYS:
                self._handle_excel_confirm_key(raw)
                return

            if in_excel and raw in _EXCEL_NAV_KEYS and not self._pressed_mods:
                self._clear_clicked_excel_cell()
                with self._excel_context_lock:
                    self._excel_cell_source = "nav"

                self._push(
                    KeyPressEvent(key=raw, target=self._last_target),
                    f"Excel nav: {raw}",
                )
                return

            if self._pressed_mods:
                combo = sorted(self._pressed_mods) + [raw]

                if frozenset(combo) == frozenset({"ctrl", "c"}) and in_excel:
                    self._handle_excel_copy()
                    return

                if "shift" in combo and raw in _EXCEL_NAV_KEYS and in_excel:
                    self._handle_excel_shift_nav(raw)
                    return

                logger.info("[RECORD] Key combo: {}", "+".join(combo))
                self._emit_combo(combo)
            else:
                self._push(KeyPressEvent(key=raw, target=self._last_target), f"Key: {raw}")
            return

        if self._pressed_mods:
            self._flush_text_buffer()

            mods = sorted(self._pressed_mods)
            combo = mods + [raw]
            action = _CLIPBOARD_COMBOS.get(frozenset(combo))

            if action == "copy":
                if self._is_in_excel_context():
                    self._handle_excel_copy()
                    return

                content = self._read_clipboard()
                self._push(
                    ClipboardCopyEvent(content=content, target=self._last_target),
                    "Copy",
                )
                return

            if action == "cut":
                self._push(
                    ClipboardCutEvent(
                        content=self._read_clipboard(),
                        target=self._last_target,
                    ),
                    "Cut",
                )
                return

            if action == "paste":
                content = self._read_clipboard()
                self._push(
                    ClipboardPasteEvent(content=content, target=self._last_target),
                    f"Paste: {(content or '')[:30]}",
                )
                return

            self._emit_combo(combo)

    def _append_text_char(self, char: str) -> None:
        if not self._text_buffer:
            in_excel = self._is_in_excel_context()
            self._text_buffer_in_excel = in_excel

            if in_excel:
                ctx = self._get_excel_typing_context()
                cell_ref = ctx.get("cell_ref")
                sheet_name = ctx.get("sheet_name")

                self._text_buffer_excel_context = (
                    {"cell_ref": cell_ref, "sheet_name": sheet_name}
                    if cell_ref
                    else None
                )

                self._text_buffer_target = (
                    self._make_excel_cell_target(cell_ref, sheet_name)
                    if cell_ref
                    else None
                )
            else:
                self._text_buffer_target = self._get_typing_target() or self._last_target
                self._text_buffer_excel_context = None

        self._text_buffer += char
        self._last_key_time = time.perf_counter()


    def _reset_text_context(self) -> None:
        self._text_buffer_target = None
        self._text_buffer_excel_context = None
        self._text_buffer_in_excel = False

    def _make_excel_cell_target(
    self,
    cell_ref: Optional[str],
    sheet_name: Optional[str],
    source: Optional[UITarget] = None,
) -> UITarget:
        cell_name = (cell_ref or "").upper()

        base = source
        if base is None:
            base = self._excel_clicked_target

        if base is not None:
            base_cell = detect_excel_cell(base.name or "", base.control_type or "")
            if (
                cell_name
                and base_cell == cell_name
                and _is_excel_process(base.process_name)
            ):
                try:
                    target = base.model_copy(deep=True)
                except Exception:
                    target = base

                target.backend = TargetBackend.UIA
                target.name = cell_name
                target.control_type = target.control_type or "SpreadsheetItem"
                target.is_editable = True

                try:
                    target.build_selectors()
                except Exception:
                    pass

                return target

        target = UITarget(
            backend=TargetBackend.UIA,
            process_name="EXCEL.EXE",
            control_type="SpreadsheetItem",
            name=cell_name or "UNKNOWN_CELL",
            is_editable=True,
        )

        try:
            target.build_selectors()
        except Exception:
            pass

        return target


    def _is_in_excel_context(self) -> bool:
        if self._last_target and _is_excel_process(self._last_target.process_name):
            return True

        if self._excel_clicked_cell:
            return True

        try:
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            if hwnd:
                info = self._enricher.get_window_info(hwnd)
                proc = (info.get("process") or "").lower()
                if proc in _EXCEL_PROCS:
                    with self._excel_context_lock:
                        self._excel_hwnd = hwnd
                    return True
        except Exception:
            pass

        return False

    def _handle_excel_confirm_key(self, key: str) -> None:
        in_edit = self._excel_in_edit_mode
        self._excel_in_edit_mode = False
        self._clear_clicked_excel_cell()

        logger.info("[RECORD] Excel confirm key: {} (was_editing={})", key, in_edit)
        self._push(
            KeyPressEvent(key=key, target=self._last_target),
            f"Excel confirm: {key}",
        )

    def _handle_excel_shift_nav(self, key: str) -> None:
        ctx = self._get_excel_context()
        current = ctx.get("cell_ref")

        if not current:
            self._emit_combo(sorted(self._pressed_mods) + [key])
            return

        if self._excel_shift_anchor is None:
            self._excel_shift_anchor = current

        self._clear_clicked_excel_cell()
        anchor = self._excel_shift_anchor

        def _check_range():
            time.sleep(0.25)
            new_ctx = self._get_excel_context()
            new_cell = new_ctx.get("cell_ref")
            if new_cell and new_cell != anchor:
                range_ref = f"{anchor}:{new_cell}"
                if not _parse_range_ref(range_ref):
                    range_ref = f"{new_cell}:{anchor}"

                logger.info("[RECORD] Excel shift-nav range: {}", range_ref)
                self._push(
                    ExcelRangeSelectEvent(
                        range_ref=range_ref,
                        sheet_name=new_ctx.get("sheet_name"),
                    ),
                    f"Excel range: {range_ref}",
                )

        threading.Thread(target=_check_range, daemon=True).start()

    def _handle_excel_copy(self) -> None:
        hwnd = self._excel_hwnd or _excel_find_hwnd()
        range_val = _excel_get_name_box_value(hwnd) if hwnd else None
        content = self._read_clipboard()

        logger.info(
            "[RECORD] Excel copy: range={} clipboard={}",
            range_val,
            (content or "")[:30],
        )

        self._push(
            ClipboardCopyEvent(content=content, target=self._last_target),
            f"Excel copy: {range_val or '(range)'}",
        )

    def _on_key_release(self, key) -> None:
        raw = self._raw_key_str(key)
        canon = _MOD_NORMALIZE.get(raw)

        if canon:
            self._pressed_mods.discard(canon)
            if canon == "shift":
                self._excel_shift_anchor = None

        if self._text_buffer:
            t = threading.Timer(
                self._config.recorder.text_flush_idle_ms / 1000,
                self._flush_if_idle,
            )
            t.daemon = True
            t.start()

    def _flush_if_idle(self) -> None:
        idle_ms = (time.perf_counter() - self._last_key_time) * 1000
        if idle_ms >= self._config.recorder.text_flush_idle_ms:
            self._flush_text_buffer()

    def _flush_text_buffer(self) -> None:
        if not self._text_buffer:
            return

        text = self._text_buffer
        self._text_buffer = ""

        try:
            if self._excel_namebox_mode:
                self._excel_namebox_mode = False

                clean = text.strip().upper()
                cell_ref = _parse_cell_ref(clean)
                range_ref = _parse_range_ref(clean)
                ctx = self._get_excel_context()

                if cell_ref:
                    target = self._make_excel_cell_target(cell_ref, ctx.get("sheet_name"))

                    with self._excel_context_lock:
                        self._excel_active_cell = cell_ref
                        self._excel_active_sheet = ctx.get("sheet_name") or self._excel_active_sheet
                        self._excel_clicked_cell = cell_ref
                        self._excel_clicked_sheet = ctx.get("sheet_name") or self._excel_active_sheet
                        self._excel_clicked_target = target
                        self._excel_cell_source = "namebox"

                    self._push(
                        ExcelCellSelectEvent(
                            cell_ref=cell_ref,
                            sheet_name=ctx.get("sheet_name"),
                            target=target,
                        ),
                        f"Excel nav: {cell_ref}",
                    )
                    return

                if range_ref:
                    first_cell = range_ref.split(":", 1)[0]
                    with self._excel_context_lock:
                        self._excel_active_cell = first_cell
                    self._clear_clicked_excel_cell()

                    self._push(
                        ExcelRangeSelectEvent(
                            range_ref=range_ref,
                            sheet_name=ctx.get("sheet_name"),
                        ),
                        f"Excel range: {range_ref}",
                    )
                    return

            if self._search_mode:
                self._search_mode = False
                target = self._text_buffer_target or self._get_typing_target()
                if target is None:
                    try:
                        target = self._enricher.get_focused_element()
                    except Exception:
                        target = None

                self._push(
                    TypeTextEvent(text=text, target=target),
                    f"Search: '{text[:40]}'",
                )
                return

            in_excel = self._text_buffer_in_excel or self._is_in_excel_context()
            typing_target = self._text_buffer_target or self._get_typing_target()

            if in_excel:
                ctx = self._text_buffer_excel_context or self._get_excel_typing_context()
                cell_ref = ctx.get("cell_ref")
                sheet_name = ctx.get("sheet_name")

                if not cell_ref:
                    logger.error(
                        "[RECORD] Refusing to record Excel text without trusted cell context: '{}'",
                        text[:40],
                    )
                    return

                typing_target = self._make_excel_cell_target(cell_ref, sheet_name)

                cell_key = f"{sheet_name or ''}!{cell_ref}"
                clear_first = self._excel_current_edit_key != cell_key
                self._excel_current_edit_key = cell_key

                is_formula = text.startswith("=")

                logger.info(
                    "[RECORD] Excel text: '{}' in {}!{}",
                    text[:40],
                    sheet_name or "?",
                    cell_ref,
                )

                self._push(
                    TypeTextEvent(
                        text=text,
                        target=typing_target,
                        clear_first=clear_first,
                        cell_ref=cell_ref,
                        sheet_name=sheet_name,
                        force_plain_text=not is_formula,
                    ),
                    f"Excel type: '{text[:30]}' in {cell_ref}",
                )
                return


            preview = text[:40] + ("..." if len(text) > 40 else "")
            typing_target = typing_target or self._get_typing_target()

            logger.info(
                "[RECORD] TypeText: '{}' into {}",
                preview,
                self._tlabel(typing_target),
            )

            self._push(
                TypeTextEvent(text=text, target=typing_target),
                f"Type: '{preview}'",
            )

        finally:
            self._reset_text_context()

    def _emit_combo(self, combo: list[str]) -> None:
        self._push(
            KeyComboEvent(keys=combo, target=self._last_target),
            f"Combo: {'+'.join(combo)}",
        )

    def _sample_focused_element(self) -> None:
        def _do_sample():
            time.sleep(self._FOCUS_SAMPLE_DELAY_MS / 1000)
            if not self._running:
                return

            try:
                focused = self._enricher.get_focused_element()
                if focused:
                    ctrl = focused.control_type or ""
                    if ctrl in _EDITABLE_CONTROL_TYPES or focused.is_editable:
                        self._last_typing_target = focused

                        if _is_excel_process(focused.process_name):
                            hwnd = _excel_find_hwnd()
                            if hwnd:
                                with self._excel_context_lock:
                                    self._excel_hwnd = hwnd
            except Exception:
                pass

        threading.Thread(target=_do_sample, daemon=True).start()

    def _get_typing_target(self) -> Optional[UITarget]:
        if self._last_typing_target:
            ctrl = self._last_typing_target.control_type or ""
            if ctrl in _EDITABLE_CONTROL_TYPES or self._last_typing_target.is_editable:
                return self._last_typing_target

        if self._last_target:
            ctrl = self._last_target.control_type or ""
            if ctrl in _EDITABLE_CONTROL_TYPES or self._last_target.is_editable:
                return self._last_target

        return None

    @staticmethod
    def _is_system_ui(target: UITarget) -> bool:
        win = (target.window_title or "").strip()
        proc = (target.process_name or "").lower()
        return win in _SYSTEM_UI_WINDOWS or proc in _SYSTEM_UI_PROCS

    def _handle_system_ui_click(self, target: UITarget, x: int, y: int, btn: str) -> None:
        name = (target.name or "").lower()
        win = target.window_title or ""

        for keyword, exe in _LAUNCH_HINTS.items():
            if keyword in name:
                self._push(
                    ProcessLaunchEvent(
                        executable=exe,
                        arguments=[],
                        wait_for_window_title=keyword.capitalize(),
                    ),
                    f"Launch: {exe}",
                )
                return

        if "search" in name or win in ("Taskbar", "Search"):
            self._search_mode = True
            self._push(
                MouseClickEvent(x=x, y=y, button=btn, target=target),
                "Click Search (system UI)",
            )
            return

        self._push(
            MouseClickEvent(x=x, y=y, button=btn, target=target),
            f"SysUI click @ ({x},{y})",
        )

    def _monitor_clipboard(self) -> None:
        while self._running:
            time.sleep(0.4)
            try:
                c = self._read_clipboard()
                if c and c != self._last_clipboard:
                    self._last_clipboard = c
            except Exception:
                pass

    def _read_clipboard(self) -> Optional[str]:
        if not WIN32_OK:
            return None

        try:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        except Exception:
            pass
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass

        return None

    def _monitor_windows(self) -> None:
        while self._running:
            try:
                time.sleep(0.15)

                hwnd = ctypes.windll.user32.GetForegroundWindow()
                if not hwnd:
                    continue

                now = time.perf_counter()

                if hwnd != self._pending_hwnd:
                    self._pending_hwnd = hwnd
                    self._pending_time = now
                    continue

                if (now - self._pending_time) * 1000 < self._WINDOW_DEBOUNCE:
                    continue

                if hwnd == self._emitted_hwnd:
                    continue

                self._emitted_hwnd = hwnd

                info = self._enricher.get_window_info(hwnd)
                title = info.get("title", "")
                proc = info.get("process", "").lower()

                if title and title not in _SYSTEM_UI_WINDOWS:
                    self._current_action_group = _APP_ACTION_GROUPS.get(proc)

                    if proc in _EXCEL_PROCS:
                        with self._excel_context_lock:
                            self._excel_hwnd = hwnd
                    else:
                        self._clear_clicked_excel_cell()

                    self._push(
                        WindowFocusEvent(
                            window_title=title,
                            process_name=info.get("process", ""),
                            x=info.get("x", 0),
                            y=info.get("y", 0),
                            width=info.get("w", 0),
                            height=info.get("h", 0),
                        ),
                        f"Focus: {title[:40]}",
                    )

                    self._last_typing_target = None
                    self._excel_in_edit_mode = False
                    self._excel_namebox_mode = False

            except Exception as exc:
                logger.debug("[RECORD] Window monitor: {}", exc)
                time.sleep(0.5)

    def _build_target_at(self, x: int, y: int) -> Optional[UITarget]:
        target = self._enricher.get_target_at(x, y)
        is_browser = False

        if target:
            proc = (target.process_name or "").lower()
            cls = target.class_name or ""
            is_browser = proc in BROWSER_PROCS or cls in _ELECTRON_CLASS

        if is_browser and self._browser.is_connected:
            win_rect = self._get_browser_window_rect(x, y)
            vx, vy = self._browser.screen_to_viewport(x, y, win_rect)
            bt = self._browser.get_element_at(vx, vy)

            if bt:
                if target is None:
                    target = UITarget(backend=TargetBackend.BROWSER)
                target.backend = TargetBackend.BROWSER
                target.browser = bt

        if target:
            target.build_selectors()

            if target.control_type in _EDITABLE_CONTROL_TYPES:
                target.is_editable = True

            try:
                rich_sel = self._enricher.get_selector_at(x, y)
                if rich_sel is not None:
                    target.rich_selectors = [rich_sel]
            except Exception:
                pass

        return target

    def _get_browser_window_rect(self, x: int, y: int) -> dict:
        try:
            hwnd = ctypes.windll.user32.WindowFromPoint(ctypes.wintypes.POINT(x, y))

            while True:
                parent = ctypes.windll.user32.GetParent(hwnd)
                if not parent:
                    break
                hwnd = parent

            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))

            return {
                "left": rect.left,
                "top": rect.top,
                "width": rect.right - rect.left,
                "height": rect.bottom - rect.top,
            }
        except Exception:
            return {"left": 0, "top": 0, "width": 1920, "height": 1080}

    @staticmethod
    def _is_transient(target: UITarget) -> bool:
        if target.control_type in _TRANSIENT_TYPES:
            return True

        if target.window_title and _TRANSIENT_TITLE_RE.search(target.window_title):
            return True

        return False

    def _maybe_screenshot(self, target: Optional[UITarget] = None) -> None:
        if not self._config.recorder.screenshot_on_every_click:
            return

        if not self._capture:
            return

        if target and self._is_system_ui(target):
            return

        path = self._capture.capture_full(0)
        if path:
            self._push(ScreenshotCheckpointEvent(path=str(path.name)), None)

    def _push(self, payload, log_text: Optional[str]) -> None:
        self._event_id += 1

        event = Event(
            id=self._event_id,
            timestamp_ms=self._now_ms() - self._start_ms,
            wall_time=datetime.now(timezone.utc).isoformat(),
            payload=payload,
            intent=_classify_intent(payload),
            action_group=self._current_action_group,
        )

        with self._lock:
            self._events.append(event)
            self._update_stats(payload)

        if log_text and self._overlay:
            self._overlay.log_event(log_text)

    def _update_stats(self, payload) -> None:
        if self._session is None:
            return

        s = self._session.stats
        s.total_events += 1

        ptype = str(getattr(payload, "type", ""))
        if "click" in ptype or "drag" in ptype:
            s.click_count += 1
        elif "type" in ptype or "key" in ptype:
            s.typing_count += 1
        elif "scroll" in ptype:
            s.scroll_count += 1

    def _make_session(self) -> Session:
        monitors = ScreenCapture.monitor_info()
        primary = monitors[0] if monitors else {"width": 1920, "height": 1080}
        now = datetime.now(timezone.utc).isoformat()

        return Session(
            name=self._name,
            created_at=now,
            updated_at=now,
            status=SessionStatus.RECORDING,
            system_info=SystemInfo(
                os_version=platform.version(),
                screen_width=primary.get("width", 1920),
                screen_height=primary.get("height", 1080),
                dpi_scale=UIAEnricher._get_dpi(),
                monitor_count=len(monitors),
                python_version=sys.version,
            ),
            screenshot_on_click=self._config.recorder.screenshot_on_every_click,
            capture_scroll_events=self._config.recorder.capture_scroll,
        )

    @staticmethod
    def _raw_key_str(key) -> str:
        try:
            c = key.char
            return c if c else ""
        except AttributeError:
            return str(key).replace("Key.", "").lower()

    @staticmethod
    def _printable_char(key) -> Optional[str]:
        try:
            c = key.char
            if c and c.isprintable() and ord(c) >= 32:
                return c
        except AttributeError:
            pass

        return None

    @staticmethod
    def _target_key(target: Optional[UITarget]) -> str:
        if not target:
            return ""

        parts = []

        if target.automation_id:
            parts.append(f"id:{target.automation_id}")
        elif target.name:
            parts.append(f"name:{target.name[:30]}")

        if target.control_type:
            parts.append(f"ctrl:{target.control_type}")

        if target.window_title:
            parts.append(f"win:{target.window_title[:20]}")

        return "|".join(parts)

    @staticmethod
    def _tlabel(target: Optional[UITarget]) -> str:
        if not target:
            return "(none)"

        if target.browser and target.browser.inner_text:
            return f"browser:'{target.browser.inner_text[:30]}'"

        if target.name:
            return f"{target.control_type or '?'}:'{target.name[:30]}'"

        if target.automation_id:
            return f"id:{target.automation_id}"

        return target.control_type or "(element)"

    @staticmethod
    def _now_ms() -> int:
        return int(time.perf_counter() * 1000)
