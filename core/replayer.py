from __future__ import annotations
import ctypes
import re
import time
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Callable

from utils.logger import logger
from utils.errors import (
    ReplayError, ElementNotFoundError, ElementNotInteractableError,
    ReplayTimeoutError, WindowNotFoundError,
)
from models.session import Session, ReplayResult
from models.event import (
    Event,
    MouseClickEvent, MouseDoubleClickEvent, MouseRightClickEvent,
    MouseMiddleClickEvent, MouseScrollEvent, MouseDragEvent,
    KeyPressEvent, KeyComboEvent, TypeTextEvent,
    ClipboardCopyEvent, ClipboardCutEvent, ClipboardPasteEvent,
    ClipboardPasteSpecialEvent,
    BrowserNavigateEvent, BrowserTabSwitchEvent, BrowserBackEvent,
    BrowserForwardEvent, BrowserRefreshEvent, BrowserWaitLoadEvent,
    WindowFocusEvent, DialogResponseEvent, FileDialogEvent,
    DropdownSelectEvent, CheckboxToggleEvent,
    ExcelCellSelectEvent, ExcelRangeSelectEvent, ExcelSheetSwitchEvent,
    ScreenshotCheckpointEvent, ExplicitWaitEvent, ProcessLaunchEvent,
)
from models.target import TargetBackend
from .matcher import ElementMatcher
from .browser_bridge import BrowserBridge
from .overlay import RecordingOverlay
from .screenshot import ScreenCapture

try:
    from pywinauto import Application
    from pywinauto.keyboard import send_keys
    from pywinauto.findwindows import find_windows
    try:
        from pywinauto import Desktop
    except ImportError:
        from pywinauto.desktop import Desktop
    UIA_OK = True
except ImportError:
    UIA_OK = False

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

_MOD_NORMALIZE = {
    "ctrl_l":"ctrl","ctrl_r":"ctrl",
    "shift_l":"shift","shift_r":"shift",
    "alt_l":"alt","alt_r":"alt","alt_gr":"alt",
    "ctrl":"ctrl","shift":"shift","alt":"alt",
}
_MOD_PREFIX = {"ctrl":"^","alt":"%","shift":"+"}

# Verified against pywinauto 0.6.x keyboard.CODES dict
_KEY_MAP = {
    "enter":     "{ENTER}",
    "return":    "{ENTER}",
    "tab":       "{TAB}",
    "space":     " ",
    "escape":    "{ESC}",       # NOTE: must be {ESC} not {ESCAPE}
    "backspace": "{BACKSPACE}",
    "delete":    "{DELETE}",
    "insert":    "{INSERT}",
    "home":      "{HOME}",
    "end":       "{END}",
    "page_up":   "{PGUP}",
    "page_down": "{PGDN}",
    "left":      "{LEFT}",
    "right":     "{RIGHT}",
    "up":        "{UP}",
    "down":      "{DOWN}",
    **{f"f{i}": f"{{F{i}}}" for i in range(1, 25)},
}

_EXCEL_PROCS = {"excel.exe"}

# BUG-9: control types that can accept text input
_EDITABLE_CONTROL_TYPES = {
    "Edit", "Document", "DataItem", "SpreadsheetItem", "Cell",
    "RichEdit", "Text", "TextBox", "ComboBox",
}
# Definitely non-editable — never attempt set_edit_text on these
_NON_EDITABLE_CONTROL_TYPES = {
    "Button", "SplitButton", "MenuItem", "TabItem", "ListItem",
    "TreeItem", "Pane", "ToolBar", "StatusBar", "ScrollBar",
    "TitleBar", "MenuBar", "Menu",
}


class ReplayEngine:

    def __init__(
        self,
        config,
        screenshot_base_dir: Path,
        on_progress: Optional[Callable[[int,int], None]] = None,
        overlay: Optional[RecordingOverlay] = None,
    ):
        self._config      = config
        self._scr_dir     = screenshot_base_dir
        self._on_progress = on_progress
        self._overlay     = overlay
        self._abort       = threading.Event()
        self._browser: Optional[BrowserBridge] = None
        self._matcher: Optional[ElementMatcher] = None
        self._capture: Optional[ScreenCapture]  = None
        self._current_hwnd: int = 0
        self._last_mouse_pos: Optional[tuple[int,int]] = None

    # ──────────────────────────────────────────────────────────────────
    # Main entry
    # ──────────────────────────────────────────────────────────────────

    def replay(self, session: Session) -> ReplayResult:
        self._abort.clear()

        self._browser = BrowserBridge(self._config.recorder.browser_cdp_port)
        connected = self._browser.connect()
        logger.info("[REPLAY] Browser CDP: {}", "connected" if connected else "NOT connected")

        scr_dir = self._scr_dir / session.id / "screenshots"
        self._capture = ScreenCapture(scr_dir)

        self._matcher = ElementMatcher(
            screenshot_base_dir=scr_dir,
            browser=self._browser,
        )

        if self._overlay:
            self._overlay.set_replaying(True)

        events_raw = session.events
        total      = len(events_raw)
        completed  = 0
        start_ms   = self._now_ms()

        logger.info("[REPLAY] Starting: '{}' ({} events)", session.name, total)

        for i, raw in enumerate(events_raw):
            if self._abort.is_set():
                logger.warning("[REPLAY] Aborted at {}/{}", completed, total)
                break

            try:
                event = Event.model_validate(raw)
            except Exception as exc:
                logger.warning("[REPLAY] Cannot parse event {}: {}", i, exc)
                continue

            is_screenshot = isinstance(event.payload, ScreenshotCheckpointEvent)

            if self._detect_interference():
                logger.warning("[REPLAY] Mouse interference — pausing 1s")
                time.sleep(1.0)

            logger.info("[REPLAY] Event #{}/{} type={} intent={} group={} ts={}ms",
                        event.id, total, event.payload.type,
                        event.intent or "?", event.action_group or "?",
                        event.timestamp_ms)

            success, used_fallback = self._execute_with_adaptive_retry(event)
            if not success:
                self._cleanup()
                return ReplayResult(
                    replayed_at=datetime.now(timezone.utc).isoformat(),
                    success=False,
                    events_total=total,
                    events_completed=completed,
                    failed_event_id=event.id,
                    error_message=f"Event #{event.id} ({event.payload.type}) failed after all retries",
                    duration_ms=self._now_ms() - start_ms,
                )

            if used_fallback:
                logger.warning("[REPLAY] Event #{} succeeded via COORD FALLBACK (element not found by selectors)",
                               event.id)

            if not is_screenshot:
                completed += 1
                if self._on_progress:
                    self._on_progress(completed, total)
                self._inter_event_delay(event, events_raw, i + 1)

        self._cleanup()
        duration_ms = self._now_ms() - start_ms
        logger.info("[REPLAY] Done: {}/{} in {:.1f}s", completed, total, duration_ms/1000)
        return ReplayResult(
            replayed_at=datetime.now(timezone.utc).isoformat(),
            success=True,
            events_total=total,
            events_completed=completed,
            duration_ms=duration_ms,
        )

    def abort(self) -> None:
        self._abort.set()

    def _cleanup(self) -> None:
        self._browser.disconnect()
        if self._overlay:
            self._overlay.set_replaying(False)

    # ──────────────────────────────────────────────────────────────────
    # BUG-11: Adaptive retry — returns (success, used_fallback)
    # ──────────────────────────────────────────────────────────────────

    def _execute_with_adaptive_retry(self, event: Event) -> tuple[bool, bool]:
        """
        BUG-11 FIX: Returns (success, used_coord_fallback).
        attempt 1 → normal element match
        attempt 2 → force window focus + re-match
        attempt 3 → coordinate click (coord fallback)
        """
        attempts     = self._config.replay.retry_attempts
        used_fallback = False

        for attempt in range(1, attempts + 1):
            try:
                if attempt == 1:
                    self._dispatch(event)
                elif attempt == 2:
                    # Focus first then retry
                    p = event.payload
                    if hasattr(p, "target") and p.target and p.target.window_title:
                        self._smart_focus_window(p.target.window_title, event.id)
                        time.sleep(0.3)
                    self._dispatch(event)
                else:
                    # BUG-11: coord fallback — mark it
                    used_fallback = True
                    self._dispatch_coord_fallback(event)

                logger.info("[REPLAY] Event #{} ✓ attempt={} fallback={}",
                            event.id, attempt, used_fallback)
                return True, used_fallback

            except (ElementNotFoundError, ElementNotInteractableError, ReplayTimeoutError) as exc:
                logger.warning("[REPLAY] Event #{} attempt {}/{}: {}",
                               event.id, attempt, attempts, exc)
                if attempt < attempts:
                    time.sleep(0.5 * attempt)
            except Exception as exc:
                logger.error("[REPLAY] Event #{} unexpected error: {}", event.id, exc)
                if attempt < attempts:
                    time.sleep(0.5 * attempt)
                else:
                    return False, used_fallback

        logger.error("[REPLAY] Event #{} FAILED after {} attempts", event.id, attempts)
        return False, used_fallback

    def _dispatch_coord_fallback(self, event: Event) -> None:
        """Direct coordinate-based execution when element match fails."""
        p = event.payload
        if isinstance(p, (MouseClickEvent, MouseDoubleClickEvent, MouseRightClickEvent)):
            x, y = p.x, p.y
            if isinstance(p, MouseDoubleClickEvent):
                self._sendinput_click(x, y, double=True)
            elif isinstance(p, MouseRightClickEvent):
                self._sendinput_right_click(x, y)
            else:
                self._sendinput_click(x, y)
            self._flash(x, y)
        elif isinstance(p, TypeTextEvent):
            self._type_at_current_focus(p.text)
        else:
            self._dispatch(event)

    # ──────────────────────────────────────────────────────────────────
    # Dispatcher
    # ──────────────────────────────────────────────────────────────────

    def _dispatch(self, event: Event) -> None:
        p = event.payload

        if isinstance(p, ScreenshotCheckpointEvent):
            return
        if isinstance(p, ExplicitWaitEvent):
            time.sleep(max(p.duration_ms/1000/self._config.replay.speed, 0.05))
            return
        if isinstance(p, ProcessLaunchEvent):
            import subprocess
            logger.info("[REPLAY] Launch: {} {}", p.executable, p.arguments)
            subprocess.Popen([p.executable] + p.arguments)
            if p.wait_for_window_title:
                self._wait_for_window(p.wait_for_window_title, event.id)
            return

        if isinstance(p, WindowFocusEvent):
            self._smart_focus_window(p.window_title, event.id)
            return

        # Clipboard
        if isinstance(p, ClipboardCopyEvent):
            self._send_combo(["ctrl","c"]); return
        if isinstance(p, ClipboardCutEvent):
            self._send_combo(["ctrl","x"]); return
        if isinstance(p, ClipboardPasteEvent):
            if p.content is not None and WIN32_OK:
                self._set_clipboard(p.content)
            self._send_combo(["ctrl","v"]); return
        if isinstance(p, ClipboardPasteSpecialEvent):
            self._send_combo(["ctrl","v"]); return

        # Excel
        if isinstance(p, ExcelCellSelectEvent):
            self._excel_navigate_to_cell(p.cell_ref, event.id); return
        if isinstance(p, ExcelRangeSelectEvent):
            self._excel_navigate_to_cell(p.range_ref, event.id); return
        if isinstance(p, ExcelSheetSwitchEvent):
            self._excel_switch_sheet(p.sheet_name, event.id); return

        # Dialogs
        if isinstance(p, DialogResponseEvent):
            self._handle_dialog(p, event.id); return
        if isinstance(p, FileDialogEvent):
            self._handle_file_dialog(p, event.id); return
        if isinstance(p, DropdownSelectEvent):
            self._handle_dropdown(p, event.id); return
        if isinstance(p, CheckboxToggleEvent):
            self._handle_checkbox(p, event.id); return

        # Keyboard
        if isinstance(p, KeyPressEvent):
            key_str = _KEY_MAP.get(p.key.lower())
            if key_str is None:
                logger.warning("[REPLAY] Unknown key '{}' — skipping", p.key)
                return
            if UIA_OK:
                send_keys(key_str)
            return
        if isinstance(p, KeyComboEvent):
            self._send_combo(p.keys); return
        if isinstance(p, TypeTextEvent):
            logger.info("[REPLAY] Type: '{}...' into target={}",
                        p.text[:30], self._tlabel(p.target))
            self._do_type(p, event.id); return

        # Mouse
        if isinstance(p, MouseClickEvent):
            logger.info("[REPLAY] Click @ ({},{}) target={}", p.x, p.y, self._tlabel(p.target))
            self._do_click(p, event.id); return
        if isinstance(p, MouseDoubleClickEvent):
            self._do_double_click(p, event.id); return
        if isinstance(p, MouseRightClickEvent):
            self._do_right_click(p, event.id); return
        if isinstance(p, MouseScrollEvent):
            self._move_mouse(p.x, p.y); self._scroll(p.dy); return
        if isinstance(p, MouseDragEvent):
            self._drag(p.start_x, p.start_y, p.end_x, p.end_y); return

    # ──────────────────────────────────────────────────────────────────
    # BUG-9: Type with target validation
    # ──────────────────────────────────────────────────────────────────

    def _do_type(self, p: TypeTextEvent, event_id: int) -> None:
        """
        BUG-9 FIX: Validate target is editable before attempting element-based input.
        If target is Button/ListItem/Pane etc. → type at current OS focus instead.
        If target is None (e.g. search query) → type at current OS focus.
        """
        # No target or explicitly None → type at current focus
        if p.target is None:
            logger.info("[REPLAY] TypeText: target=None → typing at current OS focus")
            self._type_at_current_focus(p.text)
            return

        # BUG-9: Check if target is actually editable
        ctrl = getattr(p.target, "control_type", None) or ""
        if ctrl in _NON_EDITABLE_CONTROL_TYPES:
            logger.warning(
                "[REPLAY] BUG-9: TypeText target is non-editable {} '{}' — "
                "typing at current OS focus instead",
                ctrl, (p.target.name or "")[:30]
            )
            self._type_at_current_focus(p.text)
            return

        # Excel cell
        if p.target and self._is_excel_target(p.target):
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._excel_type_into_cell(p.text, elem, event_id)
                    return
            except ElementNotFoundError:
                pass
            self._ensure_window_focus(p.target)
            send_keys("{F2}")
            time.sleep(0.05)
            if WIN32_OK:
                self._set_clipboard(p.text)
                send_keys("^v")
            else:
                send_keys(self._escape_sk(p.text), with_spaces=True)
            send_keys("{TAB}")
            return

        # Standard UIA (editable element)
        if p.target:
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._wait_ready(elem, event_id)
                    if p.clear_first:
                        try:
                            elem.set_text("")
                        except Exception:
                            elem.triple_click_input(); time.sleep(0.04)
                    try:
                        elem.set_edit_text(p.text)
                        return
                    except Exception:
                        pass
                    if WIN32_OK:
                        self._set_clipboard(p.text)
                        elem.click_input(); time.sleep(0.05)
                        send_keys("^a^v")
                        return
                    elem.click_input(); time.sleep(0.04)
                    send_keys(self._escape_sk(p.text), with_spaces=True)
                    return
            except ElementNotFoundError:
                logger.debug("[REPLAY] Element not found for type — using OS focus fallback")

        # Final fallback: type at current focus
        self._type_at_current_focus(p.text)

    def _type_at_current_focus(self, text: str) -> None:
        """Type text at whatever currently has OS keyboard focus."""
        if not UIA_OK:
            return
        logger.debug("[REPLAY] Typing {} chars at current focus", len(text))
        if WIN32_OK and len(text) > 3:
            self._set_clipboard(text)
            send_keys("^v")
        else:
            send_keys(self._escape_sk(text), with_spaces=True)

    # ──────────────────────────────────────────────────────────────────
    # BUG-10: Click with visual validation
    # ──────────────────────────────────────────────────────────────────

    def _do_click(self, p: MouseClickEvent, event_id: int) -> None:
        """BUG-10: Capture visual hash before click, validate after."""
        # Capture pre-click hash for validation
        pre_hash = self._capture_visual_hash()

        # Perform the click
        if p.target and p.target.backend != TargetBackend.BROWSER:
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._ensure_window_focus(p.target)
                    self._wait_ready(elem, event_id)
                    elem.click_input()
                    self._flash(p.x, p.y)
                    time.sleep(0.06)
                    # BUG-10: post-click visual validation
                    self._validate_visual_change(pre_hash, event_id, "click")
                    return
                self._sendinput_click(*elem)
                self._flash(*elem)
                self._validate_visual_change(pre_hash, event_id, "click_coord")
                return
            except ElementNotFoundError:
                logger.debug("[REPLAY] Element not found for click — coord fallback")

        # Coord fallback
        self._sendinput_click(p.x, p.y)
        self._flash(p.x, p.y)
        self._validate_visual_change(pre_hash, event_id, "click_raw_coord")

    def _capture_visual_hash(self) -> Optional[str]:
        """BUG-10: Capture current screen hash for change detection."""
        if not self._capture:
            return None
        try:
            path = self._capture.capture_full(0)
            if path:
                return self._capture.visual_hash(path)
        except Exception:
            pass
        return None

    def _validate_visual_change(self, pre_hash: Optional[str], event_id: int, action: str) -> None:
        """
        BUG-10 FIX: Compare pre/post hashes. If UI didn't change after a click,
        log a warning — the click may have missed or had no effect.
        """
        if pre_hash is None:
            return
        time.sleep(0.15)   # brief wait for UI to update
        post_hash = self._capture_visual_hash()
        if post_hash is None:
            return
        changed = not self._capture.compare_visual_hash(pre_hash, post_hash)
        if changed:
            logger.debug("[REPLAY] Event #{} {}: UI changed ✓", event_id, action)
        else:
            logger.warning(
                "[REPLAY] Event #{} {}: UI DID NOT CHANGE — click may have missed target",
                event_id, action,
            )

    def _do_double_click(self, p: MouseDoubleClickEvent, event_id: int) -> None:
        if p.target:
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._wait_ready(elem, event_id)
                    elem.double_click_input()
                    self._flash(p.x, p.y)
                    return
                self._sendinput_click(*elem, double=True); self._flash(*elem); return
            except ElementNotFoundError:
                pass
        self._sendinput_click(p.x, p.y, double=True)
        self._flash(p.x, p.y)

    def _do_right_click(self, p: MouseRightClickEvent, event_id: int) -> None:
        if p.target:
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._wait_ready(elem, event_id)
                    elem.right_click_input(); return
            except ElementNotFoundError:
                pass
        self._sendinput_right_click(p.x, p.y)

    # ──────────────────────────────────────────────────────────────────
    # Excel helpers
    # ──────────────────────────────────────────────────────────────────

    def _excel_navigate_to_cell(self, cell_ref: str, event_id: int) -> None:
        if not UIA_OK:
            return
        self._focus_window_by_process("excel.exe", event_id)
        time.sleep(0.1)
        send_keys("{ESC}")
        time.sleep(0.05)
        try:
            for hwnd in find_windows(title_re=".*"):
                try:
                    app  = Application(backend="uia").connect(handle=hwnd)
                    win  = app.window(handle=hwnd)
                    descs = win.descendants(auto_id="Box", control_type="Edit")
                    if descs:
                        name_box = descs[0].wrapper_object() if hasattr(descs[0], "wrapper_object") else descs[0]
                        name_box.click_input(); time.sleep(0.08)
                        if WIN32_OK:
                            self._set_clipboard(cell_ref)
                            send_keys("^a^v")
                        else:
                            send_keys("^a{DELETE}")
                            send_keys(self._cell_ref_safe(cell_ref))
                        send_keys("{ENTER}"); time.sleep(0.1)
                        return
                except Exception:
                    continue
        except Exception as exc:
            logger.warning("[REPLAY] Excel Name Box failed: {}", exc)
        send_keys("^g"); time.sleep(0.3)
        if WIN32_OK:
            self._set_clipboard(cell_ref)
            send_keys("^a^v")
        else:
            send_keys(self._cell_ref_safe(cell_ref))
        send_keys("{ENTER}")

    def _excel_type_into_cell(self, text: str, elem, event_id: int) -> None:
        send_keys("{ESC}"); time.sleep(0.05)
        try:
            elem.click_input(); time.sleep(0.08)
        except Exception:
            pass
        send_keys("{F2}"); time.sleep(0.05)
        send_keys("^a{DELETE}"); time.sleep(0.03)
        if WIN32_OK and text:
            self._set_clipboard(text)
            send_keys("^v")
        else:
            send_keys(self._escape_sk(text), with_spaces=True)
        send_keys("{TAB}"); time.sleep(0.05)

    def _excel_switch_sheet(self, sheet_name: str, event_id: int) -> None:
        if not UIA_OK:
            return
        self._focus_window_by_process("excel.exe", event_id)
        time.sleep(0.1)
        try:
            for hwnd in find_windows(title_re=".*"):
                try:
                    app  = Application(backend="uia").connect(handle=hwnd)
                    win  = app.window(handle=hwnd)
                    tabs = win.descendants(title=sheet_name, control_type="TabItem")
                    if tabs:
                        wrapper = tabs[0].wrapper_object() if hasattr(tabs[0], "wrapper_object") else tabs[0]
                        wrapper.click_input()
                        return
                except Exception:
                    continue
        except Exception:
            pass

    @staticmethod
    def _cell_ref_safe(ref: str) -> str:
        return "".join(f"{{{c}}}" if c.isupper() else c for c in ref)

    # ──────────────────────────────────────────────────────────────────
    # Window helpers
    # ──────────────────────────────────────────────────────────────────

    def _smart_focus_window(self, title: str, event_id: int) -> None:
        if not UIA_OK:
            return
        try:
            current = ctypes.windll.user32.GetForegroundWindow()
            handles = find_windows(title_re=f".*{re.escape(title[:30])}.*")
            if not handles:
                return
            if current == handles[0]:
                return
            app = Application(backend="uia").connect(handle=handles[0])
            app.window(handle=handles[0]).set_focus()
            self._current_hwnd = handles[0]
            time.sleep(0.15)
        except Exception as exc:
            logger.debug("[REPLAY] smart_focus_window '{}': {}", title, exc)

    def _ensure_window_focus(self, target) -> None:
        if not UIA_OK or not target or not target.window_title:
            return
        try:
            current = ctypes.windll.user32.GetForegroundWindow()
            buf = ctypes.create_unicode_buffer(512)
            ctypes.windll.user32.GetWindowTextW(current, buf, 512)
            if target.window_title.lower() in (buf.value or "").lower():
                return
            handles = find_windows(title_re=f".*{re.escape(target.window_title[:30])}.*")
            if handles:
                app = Application(backend="uia").connect(handle=handles[0])
                app.window(handle=handles[0]).set_focus()
                time.sleep(0.12)
        except Exception:
            pass

    def _focus_window_by_process(self, process_name: str, event_id: int) -> None:
        if not UIA_OK:
            return
        try:
            for hwnd in find_windows(title_re=".*"):
                try:
                    app = Application(backend="uia").connect(handle=hwnd)
                    win = app.window(handle=hwnd)
                    if PSUTIL_OK:
                        pid   = win.process_id()
                        pname = psutil.Process(pid).name().lower()
                        if pname == process_name.lower():
                            win.set_focus()
                            self._current_hwnd = hwnd
                            return
                except Exception:
                    continue
        except Exception:
            pass

    def _wait_for_window(self, title: str, event_id: int, timeout_ms: int = 10000) -> None:
        if not UIA_OK:
            return
        deadline = time.time() + timeout_ms/1000
        while time.time() < deadline:
            if find_windows(title_re=f".*{re.escape(title)}.*"):
                return
            time.sleep(0.3)
        raise ReplayTimeoutError(f"Window '{title}' did not appear", event_id)

    # ──────────────────────────────────────────────────────────────────
    # Dialog helpers
    # ──────────────────────────────────────────────────────────────────

    def _handle_dialog(self, p, event_id: int) -> None:
        if not UIA_OK:
            return
        try:
            handles = find_windows(title_re=f".*{re.escape(p.dialog_title)}.*")
            if not handles:
                return
            app = Application(backend="uia").connect(handle=handles[0])
            win = app.window(handle=handles[0])
            try:
                btn = win.child_window(title=p.response, control_type="Button")
                if btn.exists(timeout=2):
                    btn.click_input(); return
            except Exception:
                pass
            descs = win.descendants(title=p.response, control_type="Button")
            if descs:
                wrapper = descs[0].wrapper_object() if hasattr(descs[0], "wrapper_object") else descs[0]
                wrapper.click_input()
        except Exception as exc:
            logger.warning("[REPLAY] Dialog '{}' failed: {}", p.dialog_title, exc)

    def _handle_file_dialog(self, p, event_id: int) -> None:
        if not UIA_OK:
            return
        time.sleep(0.4)
        deadline = time.time() + 5
        handles  = []
        while time.time() < deadline:
            handles = find_windows(class_name="#32770")
            if handles:
                break
            time.sleep(0.2)
        if not handles:
            return
        try:
            app = Application(backend="uia").connect(handle=handles[0])
            win = app.window(handle=handles[0])
            for aid in ("1148", "1001"):
                try:
                    descs = win.descendants(auto_id=aid, control_type="Edit")
                    if descs:
                        wrapper = descs[0].wrapper_object() if hasattr(descs[0], "wrapper_object") else descs[0]
                        wrapper.set_edit_text(p.path)
                        time.sleep(0.1)
                        send_keys("{ENTER}")
                        return
                except Exception:
                    pass
        except Exception as exc:
            logger.warning("[REPLAY] File dialog failed: {}", exc)

    def _handle_dropdown(self, p, event_id: int) -> None:
        if not p.target or not UIA_OK:
            return
        try:
            elem = self._matcher.find(p.target, event_id)
            if not isinstance(elem, tuple):
                try:
                    elem.select(p.selected_text)
                except Exception:
                    elem.click_input()
        except ElementNotFoundError:
            pass

    def _handle_checkbox(self, p, event_id: int) -> None:
        if not p.target or not UIA_OK:
            return
        try:
            elem = self._matcher.find(p.target, event_id)
            if not isinstance(elem, tuple):
                try:
                    current = elem.get_toggle_state()
                    if bool(current == 1) != p.checked:
                        elem.toggle()
                except Exception:
                    elem.click_input()
        except ElementNotFoundError:
            pass

    # ──────────────────────────────────────────────────────────────────
    # Keyboard
    # ──────────────────────────────────────────────────────────────────

    def _send_combo(self, keys: list[str]) -> None:
        if not UIA_OK:
            return
        mods_str, key_str = "", ""
        for k in keys:
            k_norm = _MOD_NORMALIZE.get(k.lower(), k.lower())
            if k_norm in _MOD_PREFIX:
                mods_str += _MOD_PREFIX[k_norm]
            else:
                key_str = _KEY_MAP.get(k_norm, k_norm)
        if mods_str or key_str:
            send_keys(mods_str + key_str)

    @staticmethod
    def _escape_sk(text: str) -> str:
        return "".join(f"{{{c}}}" if c in set("{}()[]^+%~") else c for c in text)

    # ──────────────────────────────────────────────────────────────────
    # SendInput
    # ──────────────────────────────────────────────────────────────────

    def _sendinput_click(self, x: int, y: int, double: bool = False) -> None:
        self._move_mouse(x, y); time.sleep(0.04)
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0); time.sleep(0.025)
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)
        if double:
            time.sleep(0.06)
            ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0); time.sleep(0.025)
            ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)

    def _sendinput_right_click(self, x: int, y: int) -> None:
        self._move_mouse(x, y); time.sleep(0.04)
        ctypes.windll.user32.mouse_event(0x0008, 0, 0, 0, 0); time.sleep(0.025)
        ctypes.windll.user32.mouse_event(0x0010, 0, 0, 0, 0)

    def _move_mouse(self, x: int, y: int) -> None:
        ctypes.windll.user32.SetCursorPos(x, y)
        self._last_mouse_pos = (x, y)
        time.sleep(0.02)

    def _scroll(self, dy: int) -> None:
        ctypes.windll.user32.mouse_event(0x0800, 0, 0, dy * 120, 0)

    def _drag(self, sx: int, sy: int, ex: int, ey: int) -> None:
        self._move_mouse(sx, sy); time.sleep(0.05)
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
        for i in range(21):
            nx = sx + (ex-sx)*i//20; ny = sy + (ey-sy)*i//20
            ctypes.windll.user32.SetCursorPos(nx, ny); time.sleep(0.01)
        time.sleep(0.05)
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)

    # ──────────────────────────────────────────────────────────────────
    # Clipboard
    # ──────────────────────────────────────────────────────────────────

    def _set_clipboard(self, text: str) -> None:
        if not WIN32_OK:
            return
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        except Exception as exc:
            logger.warning("[REPLAY] set_clipboard failed: {}", exc)

    # ──────────────────────────────────────────────────────────────────
    # Wait / validation
    # ──────────────────────────────────────────────────────────────────

    def _wait_ready(self, elem, event_id: int) -> None:
        timeout_ms = self._config.replay.wait_timeout_ms
        deadline   = time.perf_counter() + timeout_ms / 1000
        while time.perf_counter() < deadline:
            try:
                if elem.is_enabled() and elem.is_visible():
                    return
            except Exception:
                pass
            time.sleep(0.08)
        raise ReplayTimeoutError(
            f"Event #{event_id}: element not ready after {timeout_ms}ms", event_id)

    def _detect_interference(self) -> bool:
        if self._last_mouse_pos is None:
            return False
        try:
            cur = ctypes.wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(cur))
            dx = abs(cur.x - self._last_mouse_pos[0])
            dy = abs(cur.y - self._last_mouse_pos[1])
            if dx + dy > 15:
                self._last_mouse_pos = (cur.x, cur.y)
                return True
        except Exception:
            pass
        return False

    # ──────────────────────────────────────────────────────────────────
    # Misc
    # ──────────────────────────────────────────────────────────────────

    def _flash(self, x: int, y: int) -> None:
        if self._overlay:
            self._overlay.flash_click(x, y, is_replay=True)

    @staticmethod
    def _is_excel_target(target) -> bool:
        return (target.process_name or "").lower() in _EXCEL_PROCS

    @staticmethod
    def _tlabel(target) -> str:
        if not target:
            return "(none)"
        if hasattr(target, "browser") and target.browser and target.browser.inner_text:
            return f"browser:'{target.browser.inner_text[:30]}'"
        if target.name:
            return f"{target.control_type or '?'}:'{target.name[:30]}'"
        if target.automation_id:
            return f"id:{target.automation_id}"
        return getattr(target, "control_type", "(element)") or "(element)"

    def _inter_event_delay(self, event: Event, all_raw: list, next_idx: int) -> None:
        if next_idx >= len(all_raw):
            return
        try:
            next_ev = Event.model_validate(all_raw[next_idx])
            gap_ms  = max(0, next_ev.timestamp_ms - event.timestamp_ms)
        except Exception:
            gap_ms = 150
        extra = 0
        etype = str(event.payload.type).lower()
        if "excel" in etype or "cell" in etype:
            extra = self._config.replay.excel_action_delay_ms
        delay = max(
            (gap_ms + extra) / 1000 / self._config.replay.speed,
            self._config.replay.min_delay_ms / 1000,
        )
        time.sleep(min(delay, 5.0))

    @staticmethod
    def _now_ms() -> int:
        return int(time.perf_counter() * 1000)
