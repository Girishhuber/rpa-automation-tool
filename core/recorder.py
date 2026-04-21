

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
)
from models.target import UITarget, TargetBackend
from .uia_enricher import UIAEnricher, detect_excel_cell
from .browser_bridge import BrowserBridge
from .overlay import RecordingOverlay
from .screenshot import ScreenCapture
from .selector import SelectorBuilder

try:
    from pynput import mouse, keyboard
    PYNPUT_OK = True
except ImportError:
    PYNPUT_OK = False
    logger.warning("[RECORD] pynput not installed — cannot record")

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
    "ctrl":"ctrl",    "shift":"shift", "alt":"alt",
}
_SPECIAL_KEYS = {
    "enter","return","tab","backspace","delete","escape","insert",
    "home","end","page_up","page_down","space",
    "left","right","up","down",
    "f1","f2","f3","f4","f5","f6","f7","f8","f9","f10","f11","f12",
    "print_screen","scroll_lock","pause","num_lock","caps_lock",
}
_CLIPBOARD_COMBOS = {
    frozenset({"ctrl","c"}): "copy",
    frozenset({"ctrl","x"}): "cut",
    frozenset({"ctrl","v"}): "paste",
}
_TRANSIENT_TYPES    = {"ToolTip","Notification","Toast","Popup"}
_TRANSIENT_TITLE_RE = __import__("re").compile(
    r"(tooltip|notification|toast|snackbar|bubble)", __import__("re").IGNORECASE
)


class Recorder:

    _DOUBLE_CLICK_MS = 400
    _DRAG_THRESHOLD  = 8
    _WINDOW_DEBOUNCE = 350

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

        # Mouse
        self._mouse_down_pos:  Optional[tuple[int,int]] = None
        self._mouse_down_time: float = 0.0
        self._last_click_pos:  tuple[int,int] = (0, 0)
        self._last_click_time: float = 0.0

        # Keyboard
        self._pressed_mods: set[str] = set()
        self._text_buffer:  str  = ""
        self._last_key_time: float = 0.0
        self._last_target:   Optional[UITarget] = None

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
            "connected" if browser_ok else "NOT connected (no browser automation)",
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
        logger.info("[RECORD] Input hooks installed (mouse + keyboard)")

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

        if is_drag and start_pos:
            self._flush_text_buffer()
            t_start = self._enricher.get_target_at(*start_pos)
            t_end   = self._enricher.get_target_at(x, y)
            self._log_element_capture("DRAG_START", *start_pos, t_start)
            self._log_element_capture("DRAG_END",   x, y, t_end)
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
            logger.debug("[RECORD] Skipping transient element ctrl={} title={}",
                         target.control_type, target.window_title)
            return

        self._last_target = target

        # Double-click
        dx2 = abs(x - self._last_click_pos[0])
        dy2 = abs(y - self._last_click_pos[1])
        if dx2 < 6 and dy2 < 6 and (now - self._last_click_time)*1000 < self._DOUBLE_CLICK_MS and btn == "left":
            self._last_click_time = 0.0
            self._log_element_capture("DBL-CLICK", x, y, target)
            self._push(MouseDoubleClickEvent(x=x, y=y, target=target),
                       f"Dbl-click {self._tlabel(target)}")
            self._maybe_screenshot()
            return

        self._last_click_pos  = (x, y)
        self._last_click_time = now

        if self._overlay:
            self._overlay.flash_click(x, y, is_replay=False)

        self._log_element_capture("CLICK", x, y, target)

        # Excel cell
        if (target and self._config.recorder.detect_excel_cells
                and target.backend == TargetBackend.UIA):
            cell = detect_excel_cell(target.name, target.control_type)
            if cell:
                self._push(ExcelCellSelectEvent(cell_ref=cell, target=target),
                           f"Excel cell: {cell}")
                self._maybe_screenshot()
                return

        if btn == "right":
            self._push(MouseRightClickEvent(x=x, y=y, target=target),
                       f"Right-click {self._tlabel(target)}")
        elif btn == "middle":
            self._push(MouseMiddleClickEvent(x=x, y=y, target=target), "Middle-click")
        else:
            self._push(MouseClickEvent(x=x, y=y, button=btn, target=target),
                       f"Click {self._tlabel(target)}")

        self._maybe_screenshot()

    def _on_mouse_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        if not self._running or not self._config.recorder.capture_scroll:
            return
        now = self._now_ms()
        if (now - getattr(self, "_last_scroll_ms", 0)) < self._config.recorder.debounce_scroll_ms:
            return
        self._last_scroll_ms = now
        target = self._enricher.get_target_at(x, y)
        self._push(MouseScrollEvent(x=x, y=y, dx=dx, dy=dy, target=target),
                   f"Scroll {'▼' if dy>0 else '▲'}")

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

        char = self._printable_char(key)
        if char:
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
                logger.info("[RECORD] Special key: {} target={}", raw, self._tlabel(self._last_target))
                self._push(KeyPressEvent(key=raw, target=self._last_target), f"Key: {raw}")
            return

        if self._pressed_mods:
            self._flush_text_buffer()
            mods  = sorted(self._pressed_mods)
            combo = mods + [raw]
            action = _CLIPBOARD_COMBOS.get(frozenset(combo))
            if action == "copy":
                content = self._read_clipboard()
                logger.info("[RECORD] Clipboard COPY: '{}...'", (content or "")[:40])
                self._push(ClipboardCopyEvent(content=content, target=self._last_target), "Copy")
                return
            if action == "cut":
                content = self._read_clipboard()
                logger.info("[RECORD] Clipboard CUT")
                self._push(ClipboardCutEvent(content=content, target=self._last_target), "Cut")
                return
            if action == "paste":
                content = self._read_clipboard()
                logger.info("[RECORD] Clipboard PASTE: '{}...'", (content or "")[:40])
                self._push(ClipboardPasteEvent(content=content, target=self._last_target),
                           f"Paste: {(content or '')[:30]}")
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
        if (time.perf_counter() - self._last_key_time)*1000 >= self._config.recorder.text_flush_idle_ms:
            self._flush_text_buffer()

    def _flush_text_buffer(self) -> None:
        if self._text_buffer:
            text = self._text_buffer
            self._text_buffer = ""
            preview = text[:40] + ("…" if len(text) > 40 else "")
            logger.info("[RECORD] TypeText: '{}' into target={}",
                        preview, self._tlabel(self._last_target))
            self._push(TypeTextEvent(text=text, target=self._last_target),
                       f"Type: '{preview}'")

    def _emit_combo(self, combo: list[str]) -> None:
        self._push(KeyComboEvent(keys=combo, target=self._last_target),
                   f"Combo: {'+'.join(combo)}")

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
            time.sleep(0.15)
            try:
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
                if info.get("title"):
                    logger.info("[RECORD] Window focus → app={} title='{}' pos=({},{}) size={}x{}",
                                info.get("process","?"), info["title"][:50],
                                info.get("x",0), info.get("y",0),
                                info.get("w",0), info.get("h",0))
                    self._push(WindowFocusEvent(
                        window_title=info["title"], process_name=info.get("process",""),
                        x=info.get("x",0), y=info.get("y",0),
                        width=info.get("w",0), height=info.get("h",0),
                    ), f"Focus: {info['title'][:40]}")
                    proc = info.get("process","").lower()
                    if proc in {"chrome.exe","msedge.exe"} and self._browser.is_connected:
                        url = self._browser.get_page_url()
                        if url:
                            self._push(BrowserNavigateEvent(url=url, wait_for_load=False),
                                       f"URL: {url[:60]}")
            except Exception:
                pass

    # ──────────────────────────────────────────────────────────────────
    # Target building
    # ──────────────────────────────────────────────────────────────────

    def _build_target_at(self, x: int, y: int) -> Optional[UITarget]:
        target = self._enricher.get_target_at(x, y)

        if self._browser.is_connected and self._enricher.is_browser_window(x, y):
            win_rect = self._get_browser_window_rect(x, y)
            vx, vy   = self._browser.screen_to_viewport(x, y, win_rect)
            bt = self._browser.get_element_at(vx, vy)
            if bt:
                if target is None:
                    target = UITarget(backend=TargetBackend.BROWSER)
                target.backend = TargetBackend.BROWSER
                target.browser = bt

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
    # Structured capture logging
    # ──────────────────────────────────────────────────────────────────

    def _log_element_capture(self, action: str, x: int, y: int,
                              target: Optional[UITarget]) -> None:
        """Emit one structured INFO line showing exactly what was captured."""
        if not target:
            logger.info("[RECORD] {} @ ({},{}) → NO TARGET (UIA unavailable)", action, x, y)
            return

        backend = target.backend.value if hasattr(target.backend, "value") else str(target.backend)

        if backend == "browser" and target.browser:
            bt = target.browser
            logger.info(
                "[RECORD] {} @ ({},{}) → backend=BROWSER app={} "
                "tag={} xpath='{}' css='{}' aria='{}' text='{}'",
                action, x, y,
                target.process_name or "?",
                bt.tag_name or "?",
                (bt.xpath or "")[:70],
                (bt.css_selector or "")[:50],
                bt.aria_label or "",
                (bt.inner_text or "")[:40],
            )
        else:
            bbox_str = ""
            if target.bbox:
                b = target.bbox
                bbox_str = f"bbox=({b.left},{b.top},{b.right-b.left}x{b.bottom-b.top})"
            logger.info(
                "[RECORD] {} @ ({},{}) → backend=UIA app={} window='{}' "
                "ctrl={} auto_id={} name='{}' class={} ancestors={} {}",
                action, x, y,
                target.process_name or "?",
                (target.window_title or "")[:40],
                target.control_type or "?",
                target.automation_id or "(none)",
                (target.name or "")[:40],
                target.class_name or "?",
                len(target.ancestor_chain),
                bbox_str,
            )

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

    def _maybe_screenshot(self) -> None:
        if self._config.recorder.screenshot_on_every_click and self._capture:
            path = self._capture.capture_full(0)
            if path:
                self._push(ScreenshotCheckpointEvent(path=str(path.name)), None)

    # ──────────────────────────────────────────────────────────────────
    # Event emission
    # ──────────────────────────────────────────────────────────────────

    def _push(self, payload, log_text: Optional[str]) -> None:
        self._event_id += 1
        event = Event(
            id=self._event_id,
            timestamp_ms=self._now_ms() - self._start_ms,
            wall_time=datetime.now(timezone.utc).isoformat(),
            payload=payload,
        )
        with self._lock:
            self._events.append(event)
        if log_text and self._overlay:
            self._overlay.log_event(log_text)

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
        try:
            c = key.char
            if c and c.isprintable():
                return c
        except AttributeError:
            pass
        return None

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
