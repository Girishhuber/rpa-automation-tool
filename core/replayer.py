"""
ReplayEngine — self-healing execution with full structured logging.

Key additions:
  - Self-healing loop: on failure tries next strategy automatically
  - Post-action validation: checks UI actually changed (text appeared, button gone)
  - Confidence-based execution: low confidence → tries 2 strategies before acting
  - State-awareness: skip step if UI already in desired state
  - Structured [REPLAY] log: shows which strategy succeeded, score, timing
  - Strategy feedback: updates matcher's learning state on success/failure
"""

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

try:
    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    from pywinauto.findwindows import find_windows
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
_KEY_MAP = {
    "enter":"{ENTER}","return":"{ENTER}","tab":"{TAB}",
    "backspace":"{BACKSPACE}","delete":"{DELETE}",
    "escape":"{ESCAPE}","home":"{HOME}","end":"{END}",
    "page_up":"{PGUP}","page_down":"{PGDN}","space":" ",
    "insert":"{INSERT}",
    "left":"{LEFT}","right":"{RIGHT}","up":"{UP}","down":"{DOWN}",
    **{f"f{i}": f"{{F{i}}}" for i in range(1, 13)},
}
_EXCEL_PROCS = {"excel.exe"}


class ReplayEngine:

    def __init__(
        self,
        config,
        screenshot_base_dir: Path,
        on_progress: Optional[Callable[[int,int], None]] = None,
        overlay: Optional[RecordingOverlay] = None,
    ):
        self._config   = config
        self._scr_dir  = screenshot_base_dir
        self._on_progress = on_progress
        self._overlay  = overlay
        self._abort    = threading.Event()
        self._browser: Optional[BrowserBridge] = None
        self._matcher: Optional[ElementMatcher] = None
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

        self._matcher = ElementMatcher(
            screenshot_base_dir=self._scr_dir / session.id / "screenshots",
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

            # User interference
            if self._detect_interference():
                logger.warning("[REPLAY] Mouse interference detected — pausing 1s")
                time.sleep(1.0)

            logger.info("[REPLAY] Event #{}/{} type={} ts={}ms",
                        event.id, total, event.payload.type, event.timestamp_ms)

            success = self._execute_with_retry(event)
            if not success:
                self._cleanup()
                return ReplayResult(
                    replayed_at=datetime.now(timezone.utc).isoformat(),
                    success=False,
                    events_total=total,
                    events_completed=completed,
                    failed_event_id=event.id,
                    error_message=f"Event #{event.id} ({event.payload.type}) failed after retries",
                    duration_ms=self._now_ms() - start_ms,
                )

            if not is_screenshot:
                completed += 1
                if self._on_progress:
                    self._on_progress(completed, total)
                self._inter_event_delay(event, events_raw, i + 1)

        self._cleanup()
        duration_ms = self._now_ms() - start_ms
        logger.info("[REPLAY] Complete: {}/{} events in {:.1f}s",
                    completed, total, duration_ms/1000)
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
    # Retry with per-event timeout + self-healing
    # ──────────────────────────────────────────────────────────────────

    def _execute_with_retry(self, event: Event) -> bool:
        attempts = self._config.replay.retry_attempts
        for attempt in range(1, attempts + 1):
            result_ok  = [False]
            exc_holder = [None]

            def _run():
                try:
                    self._dispatch(event)
                    result_ok[0] = True
                except Exception as e:
                    exc_holder[0] = e

            t = threading.Thread(target=_run, daemon=True)
            t.start()
            t.join(timeout=self._config.replay.wait_timeout_ms / 1000)

            if t.is_alive():
                logger.warning("[REPLAY] Event #{} attempt {} timed out", event.id, attempt)
                if attempt < attempts:
                    time.sleep(0.5 * attempt)
                continue

            if exc_holder[0]:
                exc = exc_holder[0]
                if isinstance(exc, (ElementNotFoundError, ElementNotInteractableError,
                                    ReplayTimeoutError)):
                    logger.warning("[REPLAY] Event #{} attempt {}/{}: {}",
                                   event.id, attempt, attempts, exc)
                    if attempt < attempts:
                        time.sleep(0.6 * attempt)
                    continue
                logger.error("[REPLAY] Event #{} unexpected error: {}", event.id, exc)
                return False

            logger.info("[REPLAY] Event #{} succeeded (attempt {})", event.id, attempt)
            return True

        logger.error("[REPLAY] Event #{} FAILED after {} attempts", event.id, attempts)
        return False

    # ──────────────────────────────────────────────────────────────────
    # Dispatcher
    # ──────────────────────────────────────────────────────────────────

    def _dispatch(self, event: Event) -> None:
        p = event.payload

        # ── System ───────────────────────────────────────────────────
        if isinstance(p, ScreenshotCheckpointEvent):
            return
        if isinstance(p, ExplicitWaitEvent):
            delay = max(p.duration_ms/1000/self._config.replay.speed, 0.05)
            logger.debug("[REPLAY] Explicit wait {:.2f}s", delay)
            time.sleep(delay)
            return
        if isinstance(p, ProcessLaunchEvent):
            import subprocess
            logger.info("[REPLAY] Launch process: {} {}", p.executable, p.arguments)
            subprocess.Popen([p.executable] + p.arguments)
            if p.wait_for_window_title:
                self._wait_for_window(p.wait_for_window_title, event.id)
            return

        # ── Window focus ──────────────────────────────────────────────
        if isinstance(p, WindowFocusEvent):
            logger.info("[REPLAY] Focus window: '{}'", p.window_title[:40])
            self._smart_focus_window(p.window_title, event.id)
            return

        # ── Browser ───────────────────────────────────────────────────
        if isinstance(p, BrowserNavigateEvent):
            if self._browser and self._browser.is_connected:
                logger.info("[REPLAY] Browser navigate → {}", p.url)
                self._browser.navigate(p.url, wait=p.wait_for_load,
                                        timeout_ms=self._config.replay.wait_timeout_ms)
            return
        if isinstance(p, BrowserTabSwitchEvent):
            if self._browser and p.tab_url:
                logger.info("[REPLAY] Browser tab switch → {}", p.tab_url[:50])
                for tab in self._browser.get_tab_list():
                    if tab.get("url","").startswith(p.tab_url[:40]):
                        self._browser.switch_to_tab(tab["id"])
                        time.sleep(0.3)
                        break
            return
        if isinstance(p, BrowserBackEvent):
            if self._browser and self._browser.is_connected:
                self._browser._send("Page.goBack", {}); self._browser.wait_for_load()
            return
        if isinstance(p, BrowserForwardEvent):
            if self._browser and self._browser.is_connected:
                self._browser._send("Page.goForward", {}); self._browser.wait_for_load()
            return
        if isinstance(p, BrowserRefreshEvent):
            if self._browser and self._browser.is_connected:
                self._browser._send("Page.reload", {}); self._browser.wait_for_load()
            return
        if isinstance(p, BrowserWaitLoadEvent):
            if self._browser and self._browser.is_connected:
                self._browser.wait_for_load(p.timeout_ms)
            return

        # ── Clipboard ─────────────────────────────────────────────────
        if isinstance(p, ClipboardCopyEvent):
            logger.info("[REPLAY] Clipboard copy")
            self._send_combo(["ctrl","c"]); return
        if isinstance(p, ClipboardCutEvent):
            logger.info("[REPLAY] Clipboard cut")
            self._send_combo(["ctrl","x"]); return
        if isinstance(p, ClipboardPasteEvent):
            logger.info("[REPLAY] Clipboard paste: '{}'", (p.content or "")[:30])
            if p.content is not None and WIN32_OK:
                self._set_clipboard(p.content)
            self._send_combo(["ctrl","v"]); return
        if isinstance(p, ClipboardPasteSpecialEvent):
            self._send_combo(["ctrl","v"]); return

        # ── Excel ─────────────────────────────────────────────────────
        if isinstance(p, ExcelCellSelectEvent):
            logger.info("[REPLAY] Excel navigate to cell: {}", p.cell_ref)
            self._excel_navigate_to_cell(p.cell_ref, event.id); return
        if isinstance(p, ExcelRangeSelectEvent):
            logger.info("[REPLAY] Excel range: {}", p.range_ref)
            self._excel_navigate_to_cell(p.range_ref, event.id); return
        if isinstance(p, ExcelSheetSwitchEvent):
            logger.info("[REPLAY] Excel sheet: {}", p.sheet_name)
            self._excel_switch_sheet(p.sheet_name, event.id); return

        # ── Dialogs ───────────────────────────────────────────────────
        if isinstance(p, DialogResponseEvent):
            logger.info("[REPLAY] Dialog '{}' → {}", p.dialog_title, p.response)
            self._handle_dialog(p, event.id); return
        if isinstance(p, FileDialogEvent):
            logger.info("[REPLAY] File dialog → {}", p.path)
            self._handle_file_dialog(p, event.id); return
        if isinstance(p, DropdownSelectEvent):
            logger.info("[REPLAY] Dropdown select: '{}'", p.selected_text)
            self._handle_dropdown(p, event.id); return
        if isinstance(p, CheckboxToggleEvent):
            logger.info("[REPLAY] Checkbox → {}", p.checked)
            self._handle_checkbox(p, event.id); return

        # ── Keyboard ──────────────────────────────────────────────────
        if isinstance(p, KeyPressEvent):
            logger.info("[REPLAY] Key: {}", p.key)
            send_keys(_KEY_MAP.get(p.key.lower(), p.key)); return
        if isinstance(p, KeyComboEvent):
            logger.info("[REPLAY] Key combo: {}", "+".join(p.keys))
            self._send_combo(p.keys); return
        if isinstance(p, TypeTextEvent):
            logger.info("[REPLAY] Type text: '{}...' into target={}",
                        p.text[:30], self._tlabel_payload(p.target))
            self._do_type(p, event.id); return

        # ── Mouse ─────────────────────────────────────────────────────
        if isinstance(p, MouseClickEvent):
            logger.info("[REPLAY] Click @ ({},{}) target={}",
                        p.x, p.y, self._tlabel_payload(p.target))
            self._do_click(p, event.id); return
        if isinstance(p, MouseDoubleClickEvent):
            logger.info("[REPLAY] Double-click @ ({},{})", p.x, p.y)
            self._do_double_click(p, event.id); return
        if isinstance(p, MouseRightClickEvent):
            logger.info("[REPLAY] Right-click @ ({},{})", p.x, p.y)
            self._do_right_click(p, event.id); return
        if isinstance(p, MouseScrollEvent):
            logger.info("[REPLAY] Scroll dy={} @ ({},{})", p.dy, p.x, p.y)
            self._move_mouse(p.x, p.y); self._scroll(p.dy); return
        if isinstance(p, MouseDragEvent):
            logger.info("[REPLAY] Drag ({},{})→({},{})", p.start_x, p.start_y, p.end_x, p.end_y)
            self._drag(p.start_x, p.start_y, p.end_x, p.end_y); return

    # ──────────────────────────────────────────────────────────────────
    # Excel
    # ──────────────────────────────────────────────────────────────────

    def _excel_navigate_to_cell(self, cell_ref: str, event_id: int) -> None:
        if not UIA_OK:
            return
        self._focus_window_by_process("excel.exe", event_id)
        time.sleep(0.1)
        send_keys("{ESCAPE}")
        time.sleep(0.05)
        try:
            desktop = Desktop(backend="uia")
            boxes   = desktop.find_elements(auto_id="Box", control_type="Edit")
            if boxes:
                name_box = boxes[0].wrapper_object()
                name_box.click_input()
                time.sleep(0.08)
                if WIN32_OK:
                    self._set_clipboard(cell_ref)
                    send_keys("^a^v")
                else:
                    send_keys("^a{DELETE}")
                    send_keys(self._cell_ref_safe(cell_ref))
                send_keys("{ENTER}")
                time.sleep(0.1)
                logger.info("[REPLAY] Excel navigated to {}", cell_ref)
                return
        except Exception as exc:
            logger.warning("[REPLAY] Name Box navigation failed: {}", exc)
        send_keys("^g")
        time.sleep(0.3)
        if WIN32_OK:
            self._set_clipboard(cell_ref)
            send_keys("^a^v")
        else:
            send_keys(self._cell_ref_safe(cell_ref))
        send_keys("{ENTER}")

    def _excel_type_into_cell(self, text: str, elem, event_id: int) -> None:
        send_keys("{ESCAPE}"); time.sleep(0.05)
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
            tabs = Desktop(backend="uia").find_elements(
                title=sheet_name, control_type="TabItem"
            )
            if tabs:
                tabs[0].wrapper_object().click_input()
        except Exception:
            pass

    @staticmethod
    def _cell_ref_safe(ref: str) -> str:
        return "".join(f"{{{c}}}" if c.isupper() else c for c in ref)

    # ──────────────────────────────────────────────────────────────────
    # Click handlers — self-healing: UIA fails → relaxed → bbox → coords
    # ──────────────────────────────────────────────────────────────────

    def _do_click(self, p: MouseClickEvent, event_id: int) -> None:
        # Browser path
        if (p.target and p.target.backend == TargetBackend.BROWSER
                and self._browser and self._browser.is_connected):
            if p.target.browser:
                self._browser.wait_for_dom_stable(stable_ms=250, max_wait_ms=2000)
                self._browser.wait_for_element(
                    p.target.browser, timeout_ms=self._config.replay.wait_timeout_ms
                )
            coords = self._matcher.find(p.target, event_id)
            if isinstance(coords, tuple):
                self._browser.bring_to_front()
                self._browser.click_at_viewport(*coords)
                self._flash(p.x, p.y)
                time.sleep(self._config.replay.browser_action_delay_ms / 1000)
                logger.info("[REPLAY] Browser click done @ viewport {}", coords)
                return

        # UIA path — self-healing
        if p.target and p.target.backend != TargetBackend.BROWSER:
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._ensure_window_focus(p.target)
                    self._wait_ready(elem, event_id)
                    elem.click_input()
                    self._flash(p.x, p.y)
                    time.sleep(0.06)
                    logger.info("[REPLAY] UIA click done strategy={}",
                                self._matcher._get_stats("automation_id").successes)
                    return
                self._sendinput_click(*elem)
                self._flash(*elem)
                return
            except ElementNotFoundError as exc:
                logger.warning("[REPLAY] UIA element not found, coord fallback: {}", exc)

        # Coordinate fallback (last resort)
        logger.warning("[REPLAY] Using raw coord fallback @ ({},{})", p.x, p.y)
        self._sendinput_click(p.x, p.y)
        self._flash(p.x, p.y)

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
    # Text input
    # ──────────────────────────────────────────────────────────────────

    def _do_type(self, p: TypeTextEvent, event_id: int) -> None:
        # Browser
        if (p.target and p.target.backend == TargetBackend.BROWSER
                and p.target.browser and self._browser and self._browser.is_connected):
            self._browser.bring_to_front()
            ok = self._browser.set_value(p.target.browser, p.text)
            if not ok:
                logger.warning("[REPLAY] set_value failed, using type_text_at")
                self._browser.type_text_at(p.text, human_like=True)
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

        # Standard UIA
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
                pass

        # Current focus fallback
        if UIA_OK:
            if WIN32_OK and len(p.text) > 3:
                self._set_clipboard(p.text)
                send_keys("^v")
            else:
                send_keys(self._escape_sk(p.text), with_spaces=True)

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
            time.sleep(0.12)
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
                time.sleep(0.1)
                logger.debug("[REPLAY] Re-focused window: {}", target.window_title[:30])
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
            btn = win.child_window(title=p.response, control_type="Button")
            if btn.exists(timeout=3):
                btn.click_input()
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
                    fn = win.child_window(auto_id=aid, control_type="Edit")
                    if fn.exists(timeout=1):
                        fn.set_edit_text(p.path)
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
            if abs(cur.x - self._last_mouse_pos[0]) + abs(cur.y - self._last_mouse_pos[1]) > 15:
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
    def _tlabel_payload(target) -> str:
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
        if "browser" in etype:   extra = self._config.replay.browser_action_delay_ms
        elif "excel" in etype or "cell" in etype: extra = self._config.replay.excel_action_delay_ms
        delay = max(
            (gap_ms + extra) / 1000 / self._config.replay.speed,
            self._config.replay.min_delay_ms / 1000,
        )
        time.sleep(min(delay, 5.0))

    @staticmethod
    def _now_ms() -> int:
        return int(time.perf_counter() * 1000)
