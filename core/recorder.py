
from __future__ import annotations
import ctypes
import platform
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
    ExcelCellSelectEvent,
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
    logger.warning("[RECORD] pynput not installed")

try:
    import win32clipboard, win32con
    WIN32_OK = True
except ImportError:
    WIN32_OK = False

# ─────────────────────────────────────────────────────────────────────────────

_MOD_NORMALIZE = {
    "ctrl_l":"ctrl",  "ctrl_r":"ctrl",
    "shift_l":"shift","shift_r":"shift",
    "alt_l":"alt",    "alt_r":"alt",  "alt_gr":"alt",
    "cmd":"cmd",      "cmd_r":"cmd",
    "ctrl":"ctrl",    "shift":"shift","alt":"alt",
}
_SPECIAL_KEYS = {
    "enter","return","tab","backspace","delete","escape","insert",
    "home","end","page_up","page_down","space",
    "left","right","up","down",
    "f1","f2","f3","f4","f5","f6","f7","f8","f9","f10","f11","f12",
}
_CLIPBOARD_COMBOS = {
    frozenset({"ctrl","c"}): "copy",
    frozenset({"ctrl","x"}): "cut",
    frozenset({"ctrl","v"}): "paste",
}

# BUG-6: scroll merge window (ms) — scrolls in same direction within this → merged
SCROLL_MERGE_WINDOW_MS = 1500

# BUG-7: System UI / Taskbar detection
_SYSTEM_UI_WINDOWS = {"Taskbar", "Search", "Action center", "Start"}
_SYSTEM_UI_PROCS   = {"explorer.exe", "searchhost.exe", "searchapp.exe",
                       "shellexperiencehost.exe", "startmenuexperiencehost.exe"}

# App hints for launch detection
_LAUNCH_HINTS = {
    "excel":         "EXCEL.EXE",
    "word":          "WINWORD.EXE",
    "powerpoint":    "POWERPNT.EXE",
    "ppt":           "POWERPNT.EXE",
    "chrome":        "chrome.exe",
    "edge":          "msedge.exe",
    "notepad":       "notepad.exe",
    "paint":         "mspaint.exe",
    "calculator":    "calc.exe",
    "outlook":       "OUTLOOK.EXE",
    "teams":         "teams.exe",
    "visual studio": "devenv.exe",
}

# BUG-1: control types that CAN accept text input
_EDITABLE_CONTROL_TYPES = {
    "Edit", "Document", "DataItem", "SpreadsheetItem", "Cell",
    "RichEdit", "Text", "TextBox", "ComboBox",
}
# Control types that definitely CANNOT accept text
_NON_EDITABLE_CONTROL_TYPES = {
    "Button", "SplitButton", "MenuItem", "TabItem", "ListItem",
    "TreeItem", "Pane", "ToolBar", "StatusBar", "ScrollBar",
    "TitleBar", "MenuBar", "Menu",
}

_TRANSIENT_TYPES    = {"ToolTip","Notification","Toast","Popup"}
_TRANSIENT_TITLE_RE = __import__("re").compile(
    r"(tooltip|notification|toast|snackbar|bubble)", __import__("re").IGNORECASE
)

# BUG-4: action group mapping
_APP_ACTION_GROUPS = {
    "winword.exe":   "word",
    "excel.exe":     "excel",
    "powerpnt.exe":  "powerpoint",
    "chrome.exe":    "browser",
    "msedge.exe":    "browser",
    "explorer.exe":  "explorer",
    "searchhost.exe":"taskbar_search",
}


# ─────────────────────────────────────────────────────────────────────────────
# BUG-2: Intent classification (moved from event_pipeline.py into recorder)
# ─────────────────────────────────────────────────────────────────────────────

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


class Recorder:

    _DOUBLE_CLICK_MS  = 400
    _DRAG_THRESHOLD   = 8
    _WINDOW_DEBOUNCE  = 350   # ms
    _CLICK_DEBOUNCE   = 300   # ms same-target dedup
    # BUG-1: delay after click before sampling focused element
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
        self._events:    list[Event]        = []
        self._lock       = threading.Lock()
        self._running    = False
        self._event_id   = 0
        self._start_ms   = 0

        # Mouse state
        self._mouse_down_pos:  Optional[tuple[int,int]] = None
        self._mouse_down_time: float = 0.0
        self._last_click_pos:  tuple[int,int] = (0, 0)
        self._last_click_time: float = 0.0
        self._last_click_target_key: str = ""
        self._last_click_target_ms:  int = 0

        # Keyboard state
        self._pressed_mods: set[str] = set()
        self._text_buffer:  str  = ""
        self._last_key_time: float = 0.0

        # BUG-1: separate targets for clicks vs typing
        self._last_target:         Optional[UITarget] = None   # last clicked element (combos/scroll)
        self._last_typing_target:  Optional[UITarget] = None   # actual focused editable element

        # BUG-6: scroll state
        self._last_scroll_ms:  int = 0                          # FIXED: was getattr fallback
        self._pending_scroll:  Optional[dict] = None
        self._last_scroll_dir: int = 0                          # +1 down, -1 up

        # BUG-7: system search state
        self._search_mode: bool = False    # True after Taskbar Search clicked

        # BUG-4: action group tracking
        self._current_action_group: Optional[str] = None

        # Window debounce
        self._pending_hwnd: int   = 0
        self._pending_time: float = 0.0
        self._emitted_hwnd: int   = 0

        self._mouse_listener = None
        self._kbd_listener   = None
        self._clipboard_thread: Optional[threading.Thread] = None
        self._window_thread:    Optional[threading.Thread] = None
        self._last_clipboard: str = ""

    # ──────────────────────────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────────────────────────

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
            "[RECORD] Session started: id={} name='{}' browser_cdp={}",
            self._session.id, self._name,
            "connected" if browser_ok else "NOT connected",
        )

        self._overlay = RecordingOverlay(self._config)
        self._overlay.start()
        self._overlay.set_recording(True)

        self._clipboard_thread = threading.Thread(target=self._monitor_clipboard, daemon=True)
        self._clipboard_thread.start()
        self._window_thread = threading.Thread(target=self._monitor_windows, daemon=True)
        self._window_thread.start()

        self._install_hooks()
        self._running = True
        self._overlay.log_event("Recording started")
        return self._session

    def stop(self) -> Session:
        if not self._running:
            raise RecorderError("Not recording")
        # Flush pending scroll before stopping (BUG-6)
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
            self._session.events      = [e.model_dump(mode="json") for e in self._events]
        self._session.duration_ms  = duration
        self._session.status       = SessionStatus.COMPLETE
        self._session.updated_at   = datetime.now(timezone.utc).isoformat()
        logger.info("[RECORD] Stopped: {} events in {:.1f}s", len(self._events), duration/1000)
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
        if self._mouse_listener: self._mouse_listener.stop()
        if self._kbd_listener:   self._kbd_listener.stop()

    # ──────────────────────────────────────────────────────────────────
    # Mouse callbacks
    # ──────────────────────────────────────────────────────────────────

    def _on_mouse_move(self, x: int, y: int) -> None:
        pass

    def _on_mouse_click(self, x: int, y: int, button, pressed: bool) -> None:
        if not self._running:
            return
        btn = button.name if hasattr(button, "name") else str(button)

        if pressed:
            self._mouse_down_pos  = (x, y)
            self._mouse_down_time = time.perf_counter()
            return

        now       = time.perf_counter()
        start_pos = self._mouse_down_pos
        self._mouse_down_pos = None

        if start_pos:
            dx = abs(x - start_pos[0])
            dy = abs(y - start_pos[1])
            is_drag = (dx + dy) > self._DRAG_THRESHOLD
        else:
            is_drag = True

        # Flush any pending scroll before a click (BUG-6)
        self._flush_pending_scroll()

        if is_drag and start_pos:
            self._flush_text_buffer()
            t_start = self._enricher.get_target_at(*start_pos)
            t_end   = self._enricher.get_target_at(x, y)
            self._push(
                MouseDragEvent(start_x=start_pos[0], start_y=start_pos[1],
                               end_x=x, end_y=y, button=btn,
                               start_target=t_start, end_target=t_end),
                f"Drag ({start_pos[0]},{start_pos[1]})→({x},{y})"
            )
            return

        self._flush_text_buffer()
        target = self._build_target_at(x, y)

        if target and self._is_transient(target):
            logger.debug("[RECORD] Skipping transient: ctrl={}", target.control_type)
            return

        # BUG-7: Detect Taskbar Search / system UI clicks
        if target and self._is_system_ui(target):
            self._handle_system_ui_click(target, x, y, btn)
            return

        # Non-system click: exit search mode
        self._search_mode = False
        self._last_target = target

        # Same-target click dedup
        tkey = self._target_key(target)
        now_ms = self._now_ms()
        if (tkey and tkey == self._last_click_target_key
                and (now_ms - self._last_click_target_ms) < self._CLICK_DEBOUNCE):
            logger.debug("[RECORD] Dedup click on same target '{}' within {}ms", tkey[:30], self._CLICK_DEBOUNCE)
            return
        self._last_click_target_key = tkey
        self._last_click_target_ms  = now_ms

        # Double-click detection
        dx2 = abs(x - self._last_click_pos[0])
        dy2 = abs(y - self._last_click_pos[1])
        if dx2 < 6 and dy2 < 6 and (now - self._last_click_time)*1000 < self._DOUBLE_CLICK_MS and btn == "left":
            self._last_click_time = 0.0
            self._push(MouseDoubleClickEvent(x=x, y=y, target=target),
                       f"Dbl-click {self._tlabel(target)}")
            self._maybe_screenshot(target)
            # BUG-1: sample focused element after double-click
            self._sample_focused_element()
            return

        self._last_click_pos  = (x, y)
        self._last_click_time = now

        if self._overlay:
            self._overlay.flash_click(x, y, is_replay=False)

        self._log_capture("CLICK", x, y, target)

        # Excel cell
        if (target and self._config.recorder.detect_excel_cells
                and target.backend == TargetBackend.UIA):
            cell = detect_excel_cell(target.name, target.control_type)
            if cell:
                self._push(ExcelCellSelectEvent(cell_ref=cell, target=target),
                           f"Excel cell: {cell}")
                self._maybe_screenshot(target)
                self._sample_focused_element()
                return

        if btn == "right":
            self._push(MouseRightClickEvent(x=x, y=y, target=target),
                       f"Right-click {self._tlabel(target)}")
        elif btn == "middle":
            self._push(MouseMiddleClickEvent(x=x, y=y, target=target), "Middle-click")
        else:
            self._push(MouseClickEvent(x=x, y=y, button=btn, target=target),
                       f"Click {self._tlabel(target)}")
            # BUG-1: sample focused element 150ms after click to get real typing target
            self._sample_focused_element()

        self._maybe_screenshot(target)

    # ──────────────────────────────────────────────────────────────────
    # BUG-1: Focus sampling — get the actual editable element
    # ──────────────────────────────────────────────────────────────────

    def _sample_focused_element(self) -> None:
        """
        BUG-1 FIX: After a click, wait _FOCUS_SAMPLE_DELAY_MS then query
        get_focused_element() to find what actually received keyboard focus.
        This is the correct target for subsequent TypeTextEvent / KeyPressEvent.
        """
        def _do_sample():
            time.sleep(self._FOCUS_SAMPLE_DELAY_MS / 1000)
            if not self._running:
                return
            try:
                focused = self._enricher.get_focused_element()
                if focused:
                    ctrl = focused.control_type or ""
                    # Only use it as typing target if it's an editable element
                    if ctrl in _EDITABLE_CONTROL_TYPES or focused.is_editable:
                        self._last_typing_target = focused
                        logger.info("[RECORD] Focused editable: ctrl={} name='{}' window='{}'",
                                    ctrl, (focused.name or "")[:30],
                                    (focused.window_title or "")[:30])
                    else:
                        # Element not editable — keep last known typing target or None
                        logger.debug("[RECORD] Focused element not editable (ctrl={}) — keeping previous typing target", ctrl)
            except Exception as exc:
                logger.debug("[RECORD] _sample_focused_element error: {}", exc)

        t = threading.Thread(target=_do_sample, daemon=True)
        t.start()

    def _get_typing_target(self) -> Optional[UITarget]:
        """
        BUG-1 FIX: Return the correct target for text input.
        Prefers _last_typing_target (focused editable element).
        Falls back to _last_target only if it's editable.
        Never returns a Button, ListItem, etc.
        """
        # Prefer the sampled focused editable element
        if self._last_typing_target:
            ctrl = self._last_typing_target.control_type or ""
            if ctrl in _EDITABLE_CONTROL_TYPES or self._last_typing_target.is_editable:
                return self._last_typing_target

        # Fall back to last clicked element only if it's actually editable
        if self._last_target:
            ctrl = self._last_target.control_type or ""
            if ctrl in _EDITABLE_CONTROL_TYPES or self._last_target.is_editable:
                return self._last_target

        # No editable target known — None means "type at current focus"
        logger.debug("[RECORD] No editable typing target — using current OS focus")
        return None

    # ──────────────────────────────────────────────────────────────────
    # BUG-6: Scroll merging
    # ──────────────────────────────────────────────────────────────────

    def _on_mouse_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        if not self._running or not self._config.recorder.capture_scroll:
            return

        now_ms = self._now_ms()
        direction = 1 if dy > 0 else -1

        # Same direction within merge window → accumulate
        if (self._pending_scroll is not None
                and direction == self._last_scroll_dir
                and (now_ms - self._last_scroll_ms) < SCROLL_MERGE_WINDOW_MS):
            self._pending_scroll["dx"] += dx
            self._pending_scroll["dy"] += dy
            self._last_scroll_ms = now_ms
            return

        # Different direction or window expired → flush previous
        self._flush_pending_scroll()

        self._pending_scroll   = {"x": x, "y": y, "dx": dx, "dy": dy,
                                   "target": self._enricher.get_target_at(x, y)}
        self._last_scroll_dir  = direction
        self._last_scroll_ms   = now_ms

        # Schedule auto-flush after merge window
        def _auto_flush():
            time.sleep(SCROLL_MERGE_WINDOW_MS / 1000)
            self._flush_pending_scroll()
        threading.Thread(target=_auto_flush, daemon=True).start()

    def _flush_pending_scroll(self) -> None:
        if self._pending_scroll:
            s = self._pending_scroll
            self._pending_scroll = None
            self._push(
                MouseScrollEvent(x=s["x"], y=s["y"], dx=s["dx"], dy=s["dy"], target=s["target"]),
                f"Scroll {'▼' if s['dy'] > 0 else '▲'} (merged)"
            )

    # ──────────────────────────────────────────────────────────────────
    # Keyboard callbacks
    # ──────────────────────────────────────────────────────────────────

    def _on_key_press(self, key) -> None:
        if not self._running:
            return
        raw   = self._raw_key_str(key)
        canon = _MOD_NORMALIZE.get(raw)

        if canon:
            self._pressed_mods.add(canon)
            return

        # BUG-8: reject control chars (ord < 32 that pass isprintable)
        char = self._printable_char(key)
        if char:
            # BUG-7: if in search mode, buffer the search query
            self._text_buffer += char
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

        if self._pressed_mods:
            self._flush_text_buffer()
            mods  = sorted(self._pressed_mods)
            combo = mods + [raw]
            action = _CLIPBOARD_COMBOS.get(frozenset(combo))
            if action == "copy":
                content = self._read_clipboard()
                self._push(ClipboardCopyEvent(content=content, target=self._last_target), "Copy")
                return
            if action == "cut":
                self._push(ClipboardCutEvent(content=self._read_clipboard(),
                                              target=self._last_target), "Cut")
                return
            if action == "paste":
                content = self._read_clipboard()
                self._push(ClipboardPasteEvent(content=content, target=self._last_target),
                           f"Paste: {(content or '')[:30]}")
                return
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
        if (time.perf_counter() - self._last_key_time)*1000 >= self._config.recorder.text_flush_idle_ms:
            self._flush_text_buffer()

    def _flush_text_buffer(self) -> None:
        if not self._text_buffer:
            return
        text = self._text_buffer
        self._text_buffer = ""

        # BUG-7: if in search mode, this is a search query → different handling
        if self._search_mode:
            self._search_mode = False
            preview = text[:40] + ("…" if len(text) > 40 else "")
            logger.info("[RECORD] Search query: '{}'", preview)
            # Emit as TypeTextEvent with None target — replayer will send_keys at focus
            self._push(TypeTextEvent(text=text, target=None),
                       f"Search: '{preview}'")
            return

        # BUG-1: use the correct typing target (focused editable element)
        typing_target = self._get_typing_target()
        preview = text[:40] + ("…" if len(text) > 40 else "")
        logger.info("[RECORD] TypeText: '{}' into target={}",
                    preview, self._tlabel(typing_target))
        self._push(TypeTextEvent(text=text, target=typing_target),
                   f"Type: '{preview}'")

    def _emit_combo(self, combo: list[str]) -> None:
        self._push(KeyComboEvent(keys=combo, target=self._last_target),
                   f"Combo: {'+'.join(combo)}")

    # ──────────────────────────────────────────────────────────────────
    # BUG-7: System UI classification
    # ──────────────────────────────────────────────────────────────────

    @staticmethod
    def _is_system_ui(target: UITarget) -> bool:
        win  = (target.window_title or "").strip()
        proc = (target.process_name or "").lower()
        return win in _SYSTEM_UI_WINDOWS or proc in _SYSTEM_UI_PROCS

    def _handle_system_ui_click(self, target: UITarget, x: int, y: int, btn: str) -> None:
        """
        BUG-7 FIX: Handle Taskbar / Search / Start clicks intelligently.
        Clicking the Search button → enter search_mode so subsequent typing
        becomes the search query, not typed into the button.
        If we can infer a launch (app name in button text) → ProcessLaunchEvent.
        """
        name = (target.name or "").lower()
        win  = (target.window_title or "")

        # Check if this click opens an app (taskbar icon)
        for keyword, exe in _LAUNCH_HINTS.items():
            if keyword in name:
                logger.info("[RECORD] Taskbar launch detected: {} → {}", name, exe)
                self._push(ProcessLaunchEvent(executable=exe, arguments=[],
                                               wait_for_window_title=keyword.capitalize()),
                           f"Launch: {exe}")
                return

        # Check if this is the Search button → next typing = search query
        if "search" in name or win in ("Taskbar", "Search"):
            logger.info("[RECORD] Search mode activated — next text is search query")
            self._search_mode = True
            # Still record the click so replayer knows to click Search first
            self._push(MouseClickEvent(x=x, y=y, button=btn, target=target),
                       f"Click Search (system UI)")
            return

        # Generic system UI click — record as coord click
        logger.debug("[RECORD] System UI click @ ({},{}) — raw click", x, y)
        self._push(MouseClickEvent(x=x, y=y, button=btn, target=target),
                   f"SysUI click @ ({x},{y})")

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
    # Window monitor (BUG-4: action group tracking)
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
                if (now - self._pending_time)*1000 < self._WINDOW_DEBOUNCE:
                    continue
                if hwnd == self._emitted_hwnd:
                    continue
                self._emitted_hwnd = hwnd
                info = self._enricher.get_window_info(hwnd)
                title = info.get("title","")
                proc  = info.get("process","").lower()
                if title and title not in _SYSTEM_UI_WINDOWS:
                    # BUG-4: update action group
                    self._current_action_group = _APP_ACTION_GROUPS.get(proc)
                    logger.info("[RECORD] Window focus → app={} title='{}' group={}",
                                proc, title[:50], self._current_action_group)
                    self._push(WindowFocusEvent(
                        window_title=title, process_name=info.get("process",""),
                        x=info.get("x",0), y=info.get("y",0),
                        width=info.get("w",0), height=info.get("h",0),
                    ), f"Focus: {title[:40]}")
                    # Reset typing target on window change (different window = different focus)
                    self._last_typing_target = None
            except Exception as exc:
                logger.debug("[RECORD] Window monitor exception: {}", exc)
                time.sleep(0.5)

    # ──────────────────────────────────────────────────────────────────
    # Target building (BUG-5: selector generation)
    # ──────────────────────────────────────────────────────────────────

    def _build_target_at(self, x: int, y: int) -> Optional[UITarget]:
        target = self._enricher.get_target_at(x, y)

        # Browser path
        is_browser = False
        if target:
            proc = (target.process_name or "").lower()
            cls  = (target.class_name   or "")
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

        # BUG-5: generate selectors + set confidence
        if target:
            target.build_selectors()
            # is_editable based on control type for correct typing target detection
            if target.control_type in _EDITABLE_CONTROL_TYPES:
                target.is_editable = True

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
                "left": rect.left, "top": rect.top,
                "width": rect.right - rect.left, "height": rect.bottom - rect.top,
            }
        except Exception:
            return {"left": 0, "top": 0, "width": 1920, "height": 1080}

    # ──────────────────────────────────────────────────────────────────
    # Filters
    # ──────────────────────────────────────────────────────────────────

    @staticmethod
    def _is_transient(target: UITarget) -> bool:
        if target.control_type in _TRANSIENT_TYPES:
            return True
        if target.window_title and _TRANSIENT_TITLE_RE.search(target.window_title):
            return True
        return False

    # ──────────────────────────────────────────────────────────────────
    # Screenshot
    # ──────────────────────────────────────────────────────────────────

    def _maybe_screenshot(self, target: Optional[UITarget] = None) -> None:
        if not self._config.recorder.screenshot_on_every_click:
            return
        if not self._capture:
            return
        # Skip screenshots for system UI (they change too fast, waste space)
        if target and self._is_system_ui(target):
            return
        path = self._capture.capture_full(0)
        if path:
            self._push(ScreenshotCheckpointEvent(path=str(path.name)), None)

    # ──────────────────────────────────────────────────────────────────
    # Event emission (BUG-2: intent, BUG-3: stats, BUG-4: action_group)
    # ──────────────────────────────────────────────────────────────────

    def _push(self, payload, log_text: Optional[str]) -> None:
        self._event_id += 1

        # BUG-2: classify intent from payload type
        intent = _classify_intent(payload)

        # BUG-4: current action group
        action_group = self._current_action_group

        event = Event(
            id=self._event_id,
            timestamp_ms=self._now_ms() - self._start_ms,
            wall_time=datetime.now(timezone.utc).isoformat(),
            payload=payload,
            intent=intent,
            action_group=action_group,
        )

        with self._lock:
            self._events.append(event)
            # BUG-3: update session stats
            self._update_stats(payload)

        if log_text and self._overlay:
            self._overlay.log_event(log_text)

    def _update_stats(self, payload) -> None:
        """BUG-3 FIX: Increment session stats on every event."""
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

    # ──────────────────────────────────────────────────────────────────
    # Logging
    # ──────────────────────────────────────────────────────────────────

    def _log_capture(self, action: str, x: int, y: int,
                     target: Optional[UITarget]) -> None:
        if not target:
            logger.info("[RECORD] {} @ ({},{}) → NO TARGET", action, x, y)
            return

        backend = target.backend.value if hasattr(target.backend, "value") else str(target.backend)
        sels    = len(target.selectors) if hasattr(target, "selectors") else 0
        conf    = getattr(target, "confidence_score", 0.0)

        if backend == "browser" and target.browser:
            bt = target.browser
            logger.info(
                "[RECORD] {} @ ({},{}) → BROWSER app={} tag={} xpath='{}' text='{}' sels={} conf={:.2f}",
                action, x, y,
                target.process_name or "?",
                bt.tag_name or "?",
                (bt.xpath or "")[:60],
                (bt.inner_text or "")[:30],
                sels, conf,
            )
        else:
            bbox_str = ""
            if target.bbox:
                b = target.bbox
                bbox_str = f"bbox=({b.left},{b.top},{b.right-b.left}x{b.bottom-b.top})"
            editable = " EDITABLE" if target.control_type in _EDITABLE_CONTROL_TYPES else ""
            logger.info(
                "[RECORD] {} @ ({},{}) → UIA app={} window='{}' ctrl={}{} "
                "auto_id={} name='{}' sels={} conf={:.2f} {}",
                action, x, y,
                target.process_name or "?",
                (target.window_title or "")[:30],
                target.control_type or "?",
                editable,
                target.automation_id or "(none)",
                (target.name or "")[:30],
                sels, conf, bbox_str,
            )

    # ──────────────────────────────────────────────────────────────────
    # Session factory
    # ──────────────────────────────────────────────────────────────────

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

    # ──────────────────────────────────────────────────────────────────
    # Key helpers
    # ──────────────────────────────────────────────────────────────────

    @staticmethod
    def _raw_key_str(key) -> str:
        try:
            c = key.char
            return c if c else ""
        except AttributeError:
            return str(key).replace("Key.", "").lower()

    @staticmethod
    def _printable_char(key) -> Optional[str]:
        """BUG-8 FIX: reject control chars with ord < 32."""
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
