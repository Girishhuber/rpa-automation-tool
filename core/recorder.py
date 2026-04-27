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
from .selector import _is_unstable

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
    import psutil
    PSUTIL_OK = True
except ImportError:
    PSUTIL_OK = False

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

_MOD_NORMALIZE = {
    "ctrl_l": "ctrl",  "ctrl_r": "ctrl",
    "shift_l": "shift", "shift_r": "shift",
    "alt_l": "alt",    "alt_r": "alt",   "alt_gr": "alt",
    "cmd": "cmd",      "cmd_r": "cmd",
    "ctrl": "ctrl",    "shift": "shift", "alt": "alt",
}

_SPECIAL_KEYS = {
    "enter", "return", "tab", "backspace", "delete", "escape", "insert",
    "home", "end", "page_up", "page_down", "space",
    "left", "right", "up", "down",
    "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8",
    "f9", "f10", "f11", "f12",
    "print_screen", "scroll_lock", "pause", "num_lock", "caps_lock",
}

_CLIPBOARD_COMBOS = {
    frozenset({"ctrl", "c"}): "copy",
    frozenset({"ctrl", "x"}): "cut",
    frozenset({"ctrl", "v"}): "paste",
}

_TRANSIENT_TYPES = {"ToolTip", "Notification", "Toast", "Popup"}
_TRANSIENT_TITLE_RE = re.compile(
    r"(tooltip|notification|toast|snackbar|bubble)", re.IGNORECASE
)

_SYSTEM_UI_WINDOWS = {"Taskbar", "Search", "Action center", "Start"}
_SYSTEM_UI_PROCS = {
    "explorer.exe", "searchhost.exe", "searchapp.exe",
    "shellexperiencehost.exe", "startmenuexperiencehost.exe",
}

_LAUNCH_HINTS = {
    "excel":       "EXCEL.EXE",
    "word":        "WINWORD.EXE",
    "powerpoint":  "POWERPNT.EXE",
    "ppt":         "POWERPNT.EXE",
    "chrome":      "chrome.exe",
    "edge":        "msedge.exe",
    "notepad":     "notepad.exe",
    "paint":       "mspaint.exe",
    "calculator":  "calc.exe",
    "outlook":     "OUTLOOK.EXE",
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

_APP_ACTION_GROUPS = {
    "winword.exe":   "word",
    "excel.exe":     "excel",
    "powerpnt.exe":  "powerpoint",
    "chrome.exe":    "browser",
    "msedge.exe":    "browser",
    "explorer.exe":  "explorer",
}

# Scroll events within this window are merged into a single event
SCROLL_MERGE_WINDOW_MS = 1500

_WEAK_RECORDING_TYPES = {
    "Pane", "Group", "ToolBar", "StatusBar", "ScrollBar", "Window",
    "Text", "Label", "StaticText",
}
_REPEATED_RECORDING_TYPES = {"ListItem", "DataItem", "Custom"}


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


# ─────────────────────────────────────────────────────────────────────────────
# Recorder
# ─────────────────────────────────────────────────────────────────────────────

class Recorder:

    _DOUBLE_CLICK_MS = 400
    _DRAG_THRESHOLD  = 8
    _WINDOW_DEBOUNCE = 350   # ms
    _CLICK_DEBOUNCE  = 300   # ms same-target dedup
    _FOCUS_SAMPLE_DELAY_MS = 150

    def __init__(self, config, sessions_dir: Path, session_name: str = "Untitled"):
        self._config       = config
        self._sessions_dir = sessions_dir
        self._name         = session_name

        self._enricher = UIAEnricher()
        self._browser  = BrowserBridge(config.recorder.browser_cdp_port)
        self._overlay: Optional[RecordingOverlay] = None
        self._capture: Optional[ScreenCapture]    = None

        self._session:   Optional[Session] = None
        self._events:    list[Event]       = []
        self._lock       = threading.Lock()
        self._running    = False
        self._event_id   = 0
        self._start_ms   = 0

        # ── Mouse state ──────────────────────────────────────────────
        self._mouse_down_pos:  Optional[tuple[int, int]] = None
        self._mouse_down_time: float = 0.0
        self._last_click_pos:  tuple[int, int] = (0, 0)
        self._last_click_time: float = 0.0
        # Dedup same-target clicks (LOG-9)
        self._last_click_target_key: str = ""
        self._last_click_target_ms:  int = 0

        # ── Scroll merge (v2.0) ──────────────────────────────────────
        self._last_scroll_ms:  int = 0
        self._pending_scroll:  Optional[dict] = None
        self._last_scroll_dir: int = 0

        # ── Keyboard / text ──────────────────────────────────────────
        self._pressed_mods:  set[str] = set()
        self._text_buffer:   str      = ""
        self._last_key_time: float    = 0.0
        # Target captured on the FIRST character of a typing run (v2.0)
        self._text_buffer_target: Optional[UITarget] = None

        self._last_target:        Optional[UITarget] = None
        self._last_typing_target: Optional[UITarget] = None

        # ── Window focus debounce ────────────────────────────────────
        self._pending_hwnd: int   = 0
        self._pending_time: float = 0.0
        self._emitted_hwnd: int   = 0

        # ── v2.0 extras ──────────────────────────────────────────────
        self._search_mode:          bool            = False
        self._current_action_group: Optional[str]  = None

        # Track last Excel cell clicked so TypeText can carry it (FIX: Excel cell recording)
        self._last_excel_cell_ref:  Optional[str]  = None

        # ── Background threads ───────────────────────────────────────
        self._mouse_listener    = None
        self._kbd_listener      = None
        self._clipboard_thread: Optional[threading.Thread] = None
        self._window_thread:    Optional[threading.Thread] = None
        self._last_clipboard:   str = ""

  
    def start(self, session_name: Optional[str] = None) -> Session:
        if self._running:
            raise RecorderError("Already recording")
        if not PYNPUT_OK:
            raise HookInstallError("pynput is required")

        self._name    = session_name or self._name
        self._session = self._make_session()
        self._events  = []
        self._event_id = 0
        self._start_ms = self._now_ms()

        scr_dir = self._sessions_dir / self._session.id / "screenshots"
        self._capture = ScreenCapture(scr_dir)

        browser_ok = self._browser.connect()
        logger.info(
            "[RECORD] Session started: id={} name='{}' browser={}",
            self._session.id, self._name,
            "connected" if browser_ok else "NOT connected",
        )

        self._overlay = RecordingOverlay(self._config)
        self._overlay.start()
        self._overlay.set_recording(True)

        self._clipboard_thread = threading.Thread(
            target=self._monitor_clipboard, daemon=True)
        self._clipboard_thread.start()

        self._window_thread = threading.Thread(
            target=self._monitor_windows, daemon=True)
        self._window_thread.start()

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
            self._overlay.log_event(f"Stopped — {len(self._events)} events")
            time.sleep(0.4)
            self._overlay.stop()

        duration = self._now_ms() - self._start_ms

        with self._lock:
            self._session.events = [e.model_dump(mode="json") for e in self._events]

        self._session.duration_ms = duration
        self._session.status      = SessionStatus.COMPLETE
        self._session.updated_at  = datetime.now(timezone.utc).isoformat()
        logger.info(
            "[RECORD] Stopped: {} events in {:.1f}s", len(self._events), duration / 1000
        )
        return self._session

    def add_checkpoint(self) -> None:
        if not self._running or not self._capture:
            return
        path = self._capture.capture_full(0)
        if path:
            self._push(
                ScreenshotCheckpointEvent(path=str(path.name), label="manual"),
                "Checkpoint",
            )

    @property
    def is_recording(self) -> bool:
        return self._running

    @property
    def event_count(self) -> int:
        return len(self._events)

    # ──────────────────────────────────────────────────────────────────
    # Hooks
    # ──────────────────────────────────────────────────────────────────

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

    # ──────────────────────────────────────────────────────────────────
    # Mouse callbacks
    # ──────────────────────────────────────────────────────────────────

    def _on_mouse_move(self, x: int, y: int) -> None:
        pass  # intentionally empty

    def _on_mouse_click(self, x: int, y: int, button, pressed: bool) -> None:
        if not self._running:
            return

        btn = button.name if hasattr(button, "name") else str(button)

        if pressed:
            self._mouse_down_pos  = (x, y)
            self._mouse_down_time = time.perf_counter()
            return

        now  = time.perf_counter()
        start_pos = self._mouse_down_pos
        self._mouse_down_pos = None

        if start_pos:
            dx = abs(x - start_pos[0])
            dy = abs(y - start_pos[1])
            is_drag = (dx + dy) > self._DRAG_THRESHOLD
        else:
            is_drag = True

        self._flush_pending_scroll()

        # ── Drag ─────────────────────────────────────────────────────
        if is_drag and start_pos:
            self._flush_text_buffer()
            t_start = self._enricher.get_target_at(*start_pos)
            t_end   = self._enricher.get_target_at(x, y)

           
            if t_start and self._is_excel_target(t_start):
                start_cell = detect_excel_cell(t_start.name or "", t_start.control_type or "")
                end_cell   = (
                    detect_excel_cell(t_end.name or "", t_end.control_type or "") if t_end else None
                )
                if start_cell and end_cell and start_cell != end_cell:
                    range_ref = f"{start_cell}:{end_cell}"
                    logger.info("[RECORD] Excel range drag: {}", range_ref)
                    self._push(
                        ExcelRangeSelectEvent(range_ref=range_ref, sheet_name=None),
                        f"Excel range: {range_ref}",
                    )
                    return

            self._log_capture("DRAG_START", *start_pos, t_start)
            self._push(
                MouseDragEvent(
                    start_x=start_pos[0], start_y=start_pos[1],
                    end_x=x, end_y=y, button=btn,
                    start_target=t_start, end_target=t_end,
                ),
                f"Drag ({start_pos[0]},{start_pos[1]})→({x},{y})",
            )
            return

        # Build enriched target (single UIA call)
        target = self._build_target_at(x, y)

        # ── Pre-flush: detect the cell being clicked ──────────────────────────
        # Two distinct cases:
        #
        # Case A — No current cell (first click ever, or typing before any cell click):
        #   The buffered text has no cell anchor. Assign the NEW cell's ref so the
        #   text goes to the cell the user is about to click. E.g. user types "Name"
        #   then clicks E8 — "Name" should land in E8.
        #
        # Case B — Already in a cell (user typed in cell X, now clicking cell Y):
        #   The buffered text was typed in cell X. Keep the existing ref so the flush
        #   records it in X. The new ref (Y) will be set AFTER the flush by the
        #   ExcelCellSelectEvent block below.
        #   E.g. user types "Age" in E8, clicks G8 — "Age" should stay in E8.
        _clicked_cell = None
        if (target
                and self._config.recorder.detect_excel_cells
                and self._is_excel_target(target)
                and target.backend == TargetBackend.UIA):
            _clicked_cell = detect_excel_cell(target.name or "", target.control_type or "")
            if _clicked_cell and not self._last_excel_cell_ref:
                # Case A: no existing anchor — assign new cell so buffered text lands here
                self._last_excel_cell_ref = _clicked_cell
                self._last_typing_target  = target

        # Flush text buffer — uses existing _last_excel_cell_ref (correct for Case B)
        self._flush_text_buffer()

        # ── Filters ──────────────────────────────────────────────────
        if target and self._is_transient(target):
            logger.debug("[RECORD] Skipping transient: ctrl={}", target.control_type)
            return

        # ── System UI (taskbar, search, launcher) ────────────────────
        if target and self._is_system_ui(target):
            self._handle_system_ui_click(target, x, y, btn)
            return

        # ── Excel sheet tab click (v2.0) ─────────────────────────────
        if (target and self._is_excel_target(target)
                and target.control_type == "TabItem"):
            sheet_name = target.name or ""
            if sheet_name:
                logger.info("[RECORD] Excel sheet tab click: '{}'", sheet_name)
                self._push(
                    ExcelSheetSwitchEvent(sheet_name=sheet_name, sheet_index=0),
                    f"Excel sheet: {sheet_name}",
                )
                self._sample_focused_element()
                return

        self._search_mode = False
        self._last_target = target

        # ── Dedup same-target clicks (LOG-9) ─────────────────────────
        tkey   = self._target_key(target)
        now_ms = self._now_ms()
        if (tkey
                and tkey == self._last_click_target_key
                and (now_ms - self._last_click_target_ms) < self._CLICK_DEBOUNCE):
            logger.debug("[RECORD] Dedup: same target '{}' within {}ms",
                         tkey[:30], self._CLICK_DEBOUNCE)
            return

        self._last_click_target_key = tkey
        self._last_click_target_ms  = now_ms

        # ── Double-click detection ────────────────────────────────────
        dx2 = abs(x - self._last_click_pos[0])
        dy2 = abs(y - self._last_click_pos[1])
        if (dx2 < 6 and dy2 < 6
                and (now - self._last_click_time) * 1000 < self._DOUBLE_CLICK_MS
                and btn == "left"):
            self._last_click_time = 0.0
            self._log_capture("DBL-CLICK", x, y, target)
            self._push(
                MouseDoubleClickEvent(x=x, y=y, target=target),
                f"Dbl-click {self._tlabel(target)}",
            )
            self._maybe_screenshot(target)
            self._sample_focused_element()
            return

        self._last_click_pos  = (x, y)
        self._last_click_time = now

        if self._overlay:
            self._overlay.flash_click(x, y, is_replay=False)

        self._log_capture("CLICK", x, y, target)

        # ── Excel cell click — v1.0 logic (simple, no polling) ───────
        if (target
                and self._config.recorder.detect_excel_cells
                and self._is_excel_target(target)
                and target.backend == TargetBackend.UIA):
            cell = detect_excel_cell(target.name or "", target.control_type or "")
            if cell:
                logger.info("[RECORD] Excel cell click: {}", cell)
                # FIX: update cell ref on every new cell click; _flush_text_buffer
                # will now carry forward the ref for consecutive flushes in the same cell.
                self._last_excel_cell_ref = cell
                self._last_typing_target  = target
                self._push(
                    ExcelCellSelectEvent(cell_ref=cell, target=target),
                    f"Excel cell: {cell}",
                )
                self._maybe_screenshot(target)
                self._sample_focused_element()
                return

        # ── Standard mouse events ─────────────────────────────────────
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

    # ──────────────────────────────────────────────────────────────────
    # Scroll — merged (v2.0)
    # ──────────────────────────────────────────────────────────────────

    def _on_mouse_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        if not self._running or not self._config.recorder.capture_scroll:
            return

        now_ms    = self._now_ms()
        direction = 1 if dy > 0 else -1

        if (self._pending_scroll is not None
                and direction == self._last_scroll_dir
                and (now_ms - self._last_scroll_ms) < SCROLL_MERGE_WINDOW_MS):
            # Accumulate into running scroll event
            self._pending_scroll["dx"] += dx
            self._pending_scroll["dy"] += dy
            self._last_scroll_ms = now_ms
            return

        # Different direction or window expired — flush old, start new
        self._flush_pending_scroll()

        self._pending_scroll = {
            "x":  x, "y": y,
            "dx": dx, "dy": dy,
            "target": self._enricher.get_target_at(x, y),
        }
        self._last_scroll_dir = direction
        self._last_scroll_ms  = now_ms

        # Auto-flush after the merge window
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
                x=s["x"], y=s["y"],
                dx=s["dx"], dy=s["dy"],
                target=s["target"],
            ),
            f"Scroll {'▼' if s['dy'] < 0 else '▲'} (merged)",
        )

    def _on_key_press(self, key) -> None:
        if not self._running:
            return

        raw   = self._raw_key_str(key)
        canon = _MOD_NORMALIZE.get(raw)

        if canon:
            self._pressed_mods.add(canon)
            return

        # Printable character → accumulate text buffer
        char = self._printable_char(key)
        if char:
            # Capture the target on the very FIRST char of a typing run (v2.0)
            if not self._text_buffer:
                self._text_buffer_target = (
                    self._get_typing_target() or self._last_target
                )
            self._text_buffer += char
            self._last_key_time = time.perf_counter()
            return

        if raw == "space" and not self._pressed_mods:
            if not self._text_buffer:
                self._text_buffer_target = (
                    self._get_typing_target() or self._last_target
                )
            self._text_buffer += " "
            self._last_key_time = time.perf_counter()
            return

        if raw in _SPECIAL_KEYS:
            self._flush_text_buffer()
            if self._pressed_mods:
                combo = sorted(self._pressed_mods) + [raw]
                logger.info("[RECORD] Key combo: {}", "+".join(combo))
                self._emit_combo(combo)
            else:
                logger.info("[RECORD] Key: {}", raw)
                self._push(KeyPressEvent(key=raw, target=self._last_target), f"Key: {raw}")
            return

        # Key + modifiers (Ctrl+something, etc.)
        if self._pressed_mods:
            self._flush_text_buffer()
            mods  = sorted(self._pressed_mods)
            combo = mods + [raw]
            action = _CLIPBOARD_COMBOS.get(frozenset(combo))

            if action == "copy":
                content = self._read_clipboard()
                logger.info("[RECORD] Clipboard COPY: '{}'", (content or "")[:40])
                self._push(ClipboardCopyEvent(content=content, target=self._last_target), "Copy")
                return
            if action == "cut":
                logger.info("[RECORD] Clipboard CUT")
                self._push(
                    ClipboardCutEvent(
                        content=self._read_clipboard(), target=self._last_target
                    ),
                    "Cut",
                )
                return
            if action == "paste":
                content = self._read_clipboard()
                logger.info("[RECORD] Clipboard PASTE: '{}'", (content or "")[:40])
                self._push(
                    ClipboardPasteEvent(content=content, target=self._last_target),
                    f"Paste: {(content or '')[:30]}",
                )
                return

            logger.info("[RECORD] Key combo: {}", "+".join(combo))
            self._emit_combo(combo)

    def _on_key_release(self, key) -> None:
        raw   = self._raw_key_str(key)
        canon = _MOD_NORMALIZE.get(raw)
        if canon:
            self._pressed_mods.discard(canon)

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
            # FIX: If we're in an Excel context (last target is EXCEL.EXE) but no
            # cell_ref has been set yet, the user typed before clicking any cell.
            # Deferring the flush lets the upcoming cell click set the correct cell_ref,
            # so the TypeText gets anchored to the right cell instead of being emitted
            # as a non-editable ListItem event (which the replayer then skips entirely).
            #
            # We detect this as: last_target is Excel, control_type is non-editable
            # (e.g. ListItem = splash screen), and no cell ref is known.
            last = self._last_target
            if (last
                    and self._is_excel_target(last)
                    and not self._last_excel_cell_ref
                    and (last.control_type or "") in ("ListItem", "Pane", "Window")):
                logger.debug("[RECORD] Idle flush deferred — Excel context, awaiting cell click")
                return

            self._flush_text_buffer()

    def _flush_text_buffer(self) -> None:
     
        if not self._text_buffer:
            return

        text   = self._text_buffer
        target = self._text_buffer_target or self._get_typing_target() or self._last_target

        self._text_buffer        = ""
        self._text_buffer_target = None

        preview = text[:40] + ("…" if len(text) > 40 else "")
        logger.info("[RECORD] TypeText: '{}' into {}", preview, self._tlabel(target))

        # v2.0: search-mode label
        log_label = f"Search: '{preview}'" if self._search_mode else f"Type: '{preview}'"
        self._search_mode = False

        # FIX: For Excel, attach the cell_ref so the replayer navigates correctly.
        # Do NOT reset _last_excel_cell_ref here — consecutive typing in the same cell
        # produces multiple flush events (e.g. 'A' then 'ge' for 'Age'), all in the same cell.
        # The ref is only reset when a new Excel cell is clicked (in _on_mouse_click).
        type_event = TypeTextEvent(text=text, target=target)
        if target and self._is_excel_target(target) and self._last_excel_cell_ref:
            type_event.cell_ref = self._last_excel_cell_ref  # type: ignore[attr-defined]
            logger.info("[RECORD] TypeText Excel cell_ref={}", self._last_excel_cell_ref)

        self._push(type_event, log_label)

    def _emit_combo(self, combo: list[str]) -> None:
        self._push(
            KeyComboEvent(keys=combo, target=self._last_target),
            f"Combo: {'+'.join(combo)}",
        )

    # ──────────────────────────────────────────────────────────────────
    # Clipboard monitor
    # ──────────────────────────────────────────────────────────────────

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

    # ──────────────────────────────────────────────────────────────────
    # Window monitor
    # ──────────────────────────────────────────────────────────────────

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
                info  = self._enricher.get_window_info(hwnd)
                title = info.get("title", "")
                proc  = info.get("process", "").lower()

                if title and title not in _SYSTEM_UI_WINDOWS:
                    self._current_action_group = _APP_ACTION_GROUPS.get(proc)

                    logger.info(
                        "[RECORD] Window focus → app={} title='{}'",
                        proc, title[:50],
                    )
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

                    if proc in {"chrome.exe", "msedge.exe"} and self._browser.is_connected:
                        url = self._browser.get_page_url()
                        if url:
                            self._push(
                                BrowserNavigateEvent(url=url, wait_for_load=False),
                                f"URL: {url[:60]}",
                            )

            except Exception as exc:
                logger.debug("[RECORD] Window monitor exception: {}", exc)
                time.sleep(0.5)



    def _sanitize_recorded_target(self, target: Optional[UITarget]) -> Optional[UITarget]:
        if not target:
            return None
        if target.automation_id and _is_unstable(target.automation_id):
            logger.warning("[RECORD] Discarding unstable automation_id='{}'", target.automation_id)
            target.automation_id = None
        if target.class_name and _is_unstable(target.class_name):
            logger.debug("[RECORD] Discarding unstable class_name='{}'", target.class_name)
            target.class_name = None
        return target

    @staticmethod
    def _target_anchor_count(target: Optional[UITarget]) -> int:
        if not target:
            return 0
        count = 0
        for rich in getattr(target, "rich_selectors", None) or []:
            count += len(getattr(rich, "anchor_elements", None) or [])
        return count

    def _recording_quality_score(self, target: Optional[UITarget]) -> int:
        if not target:
            return 0
        if target.backend == TargetBackend.BROWSER and target.browser:
            score = 0
            bt = target.browser
            if bt.xpath: score += 35
            if bt.css_selector: score += 30
            if bt.aria_label: score += 25
            if bt.inner_text and len(bt.inner_text.strip()) > 2: score += 15
            return min(score, 100)

        score = 0
        if target.automation_id and not _is_unstable(target.automation_id):
            score += 55
        if target.name and len(target.name.strip()) > 3:
            score += 25
        if target.control_type:
            score += 10
        if target.window_title or target.process_name:
            score += 10
        if target.ancestor_chain:
            score += 10
        if self._target_anchor_count(target):
            score += 15
        if (target.control_type or "") in _WEAK_RECORDING_TYPES:
            score -= 25
        if (target.control_type or "") in _REPEATED_RECORDING_TYPES and not target.automation_id:
            score -= 10
        return max(0, min(score, 100))

    def _apply_recording_quality(self, target: Optional[UITarget]) -> None:
        if not target:
            return
        score = self._recording_quality_score(target)
        target.confidence_score = max(target.confidence_score or 0.0, score / 100)
        if score >= 75:
            target.confidence_level = "high"
            target.confidence_reason = "stable_selector"
        elif score >= 50:
            target.confidence_level = "medium"
            target.confidence_reason = "anchored_selector"
        else:
            target.confidence_level = "low"
            target.confidence_reason = "weak_selector"
            logger.warning(
                "[RECORD] Weak target captured: score={} ctrl={} auto_id={} name='{}' anchors={}",
                score, target.control_type or "?", target.automation_id or "(none)",
                (target.name or "")[:40], self._target_anchor_count(target),
            )

    def _build_target_at(self, x: int, y: int) -> Optional[UITarget]:
        target = self._enricher.get_target_at(x, y)

        is_browser = False
        if target:
            proc      = (target.process_name or "").lower()
            cls       = target.class_name or ""
            is_browser = proc in BROWSER_PROCS or cls in _ELECTRON_CLASS

        if is_browser and self._browser.is_connected:
            win_rect = self._get_browser_window_rect(x, y)
            vx, vy   = self._browser.screen_to_viewport(x, y, win_rect)
            bt = self._browser.get_element_at(vx, vy)
            if bt:
                if target is None:
                    target = UITarget(backend=TargetBackend.BROWSER)
                target.backend = TargetBackend.BROWSER
                target.browser = bt

        # v2.0: sanitize unstable selectors, enrich, then score target quality.
        if target:
            target = self._sanitize_recorded_target(target)

            if target.control_type in _EDITABLE_CONTROL_TYPES:
                target.is_editable = True

            try:
                rich_sel = self._enricher.get_selector_at(x, y)
                if rich_sel is not None:
                    target.rich_selectors = [rich_sel]
            except Exception:
                pass

            target.build_selectors()
            self._apply_recording_quality(target)

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
                "left":   rect.left,
                "top":    rect.top,
                "width":  rect.right - rect.left,
                "height": rect.bottom - rect.top,
            }
        except Exception:
            return {"left": 0, "top": 0, "width": 1920, "height": 1080}

 
    @staticmethod
    def _is_system_ui(target: UITarget) -> bool:
        win  = (target.window_title or "").strip()
        proc = (target.process_name or "").lower()
        return win in _SYSTEM_UI_WINDOWS or proc in _SYSTEM_UI_PROCS

    def _handle_system_ui_click(
        self, target: UITarget, x: int, y: int, btn: str
    ) -> None:
   
        name = (target.name or "").lower()
        win  = target.window_title or ""

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

        # Search bar / Taskbar search
        if "search" in name or win in ("Taskbar", "Search"):
            self._search_mode = True
            self._push(
                MouseClickEvent(x=x, y=y, button=btn, target=target),
                "Click Search (system UI)",
            )
            return

        # Generic system UI — record as coord click
        logger.debug("[RECORD] System UI click @ ({},{}) — recording as coord click", x, y)
        self._push(
            MouseClickEvent(x=x, y=y, button=btn, target=target),
            f"SysUI click @ ({x},{y})",
        )

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

    def _sample_focused_element(self) -> None:
        def _do_sample():
            time.sleep(self._FOCUS_SAMPLE_DELAY_MS / 1000)
            if not self._running:
                return
            try:
                focused = self._enricher.get_focused_element()
                if focused:
                    focused = self._sanitize_recorded_target(focused)
                    self._apply_recording_quality(focused)
                    ctrl = focused.control_type or ""
                    if ctrl in _EDITABLE_CONTROL_TYPES or focused.is_editable:
                        self._last_typing_target = focused
            except Exception:
                pass

        threading.Thread(target=_do_sample, daemon=True).start()

    def _get_typing_target(self) -> Optional[UITarget]:
        """Return the best known editable element to attach typed text to."""
        if self._last_typing_target:
            ctrl = self._last_typing_target.control_type or ""
            if ctrl in _EDITABLE_CONTROL_TYPES or self._last_typing_target.is_editable:
                return self._last_typing_target
        if self._last_target:
            ctrl = self._last_target.control_type or ""
            if ctrl in _EDITABLE_CONTROL_TYPES or self._last_target.is_editable:
                return self._last_target
        return None

  
    def _log_capture(
        self, action: str, x: int, y: int, target: Optional[UITarget]
    ) -> None:
        if not target:
            logger.info("[RECORD] {} @ ({},{}) → NO TARGET", action, x, y)
            return

        backend = (
            target.backend.value
            if hasattr(target.backend, "value")
            else str(target.backend)
        )

        if backend == "browser" and target.browser:
            bt = target.browser
            logger.info(
                "[RECORD] {} @ ({},{}) → BROWSER app={} tag={} "
                "xpath='{}' css='{}' text='{}'",
                action, x, y,
                target.process_name or "?",
                bt.tag_name or "?",
                (bt.xpath or "")[:70],
                (bt.css_selector or "")[:50],
                (bt.inner_text or "")[:40],
            )
        else:
            bbox_str = ""
            if target.bbox:
                b = target.bbox
                bbox_str = (
                    f"bbox=({b.left},{b.top},"
                    f"{b.right - b.left}x{b.bottom - b.top})"
                )
            logger.info(
                "[RECORD] {} @ ({},{}) → UIA app={} window='{}' "
                "ctrl={} auto_id={} name='{}' class={} anc={} {}",
                action, x, y,
                target.process_name or "?",
                (target.window_title or "")[:40],
                target.control_type or "?",
                target.automation_id or "(none)",
                (target.name or "")[:40],
                target.class_name or "?",
                len(target.ancestor_chain or []),
                bbox_str,
            )


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
        primary  = monitors[0] if monitors else {"width": 1920, "height": 1080}
        now      = datetime.now(timezone.utc).isoformat()
        return Session(
            name=self._name,
            created_at=now, updated_at=now,
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
    def _is_excel_target(target: UITarget) -> bool:
        return (target.process_name or "").lower() == "excel.exe"

    @staticmethod
    def _raw_key_str(key) -> str:
        try:
            c = key.char
            return c if c else ""
        except AttributeError:
            return str(key).replace("Key.", "").lower()

    @staticmethod
    def _printable_char(key) -> Optional[str]:
        """Reject control chars (ord < 32) so Ctrl+Shift+Z → combo, not text."""
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