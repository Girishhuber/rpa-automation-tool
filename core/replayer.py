from __future__ import annotations
import ctypes
import ctypes.wintypes
import re
import time
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Callable

from utils.logger import logger
from utils.errors import (
    ReplayError, ElementNotFoundError, ElementNotInteractableError,
    ReplayTimeoutError,
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
        from pywinauto import Desktop
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
    "ctrl_l": "ctrl", "ctrl_r": "ctrl",
    "shift_l": "shift", "shift_r": "shift",
    "alt_l": "alt", "alt_r": "alt", "alt_gr": "alt",
    "ctrl": "ctrl", "shift": "shift", "alt": "alt",
}
_MOD_PREFIX = {"ctrl": "^", "alt": "%", "shift": "+"}


_KEY_MAP = {
    "enter":     "{ENTER}",
    "return":    "{ENTER}",
    "tab":       "{TAB}",
    "space":     " ",
    "escape":    "{ESC}",          # ← CRITICAL: NOT {ESCAPE}
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
_CELL_RE     = re.compile(r"^[A-Z]{1,3}[0-9]{1,7}$")
_RANGE_RE    = re.compile(r"^([A-Z]{1,3}[0-9]{1,7}):([A-Z]{1,3}[0-9]{1,7})$")

_EDITABLE_CONTROL_TYPES = {
    "Edit", "Document", "DataItem", "SpreadsheetItem", "Cell",
    "RichEdit", "Text", "TextBox", "ComboBox",
}
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
        on_progress: Optional[Callable[[int, int], None]] = None,
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
        self._last_mouse_pos: Optional[tuple[int, int]] = None
        # REP-7: track expected active sheet
        self._excel_expected_sheet: Optional[str] = None

    # ──────────────────────────────────────────────────────────────────
    # Main entry
    # ──────────────────────────────────────────────────────────────────

    def replay(self, session: Session) -> ReplayResult:
        self._abort.clear()

        self._browser = BrowserBridge(self._config.recorder.browser_cdp_port)
        connected     = self._browser.connect()
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
                break
            try:
                event = Event.model_validate(raw)
            except Exception as exc:
                logger.warning("[REPLAY] Cannot parse event {}: {}", i, exc)
                continue

            is_screenshot = isinstance(event.payload, ScreenshotCheckpointEvent)
            if self._detect_interference():
                time.sleep(1.0)

            logger.info("[REPLAY] Event #{}/{} type={} intent={} ts={}ms",
                        event.id, total, event.payload.type,
                        event.intent or "?", event.timestamp_ms)

            success, used_fallback = self._execute_with_adaptive_retry(event)
            if not success:
                self._cleanup()
                return ReplayResult(
                    replayed_at=datetime.now(timezone.utc).isoformat(),
                    success=False,
                    events_total=total,
                    events_completed=completed,
                    failed_event_id=event.id,
                    error_message=f"Event #{event.id} ({event.payload.type}) failed",
                    duration_ms=self._now_ms() - start_ms,
                )

            if not is_screenshot:
                completed += 1
                if self._on_progress:
                    self._on_progress(completed, total)
                self._inter_event_delay(event, events_raw, i + 1)

        self._cleanup()
        duration_ms = self._now_ms() - start_ms
        logger.info("[REPLAY] Done: {}/{} in {:.1f}s", completed, total, duration_ms / 1000)
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
    # Adaptive retry
    # ──────────────────────────────────────────────────────────────────

    def _execute_with_adaptive_retry(self, event: Event) -> tuple[bool, bool]:
        attempts      = self._config.replay.retry_attempts
        used_fallback = False

        for attempt in range(1, attempts + 1):
            try:
                if attempt == 1:
                    self._dispatch(event)
                elif attempt == 2:
                    p = event.payload
                    if hasattr(p, "target") and p.target and p.target.window_title:
                        self._smart_focus_window(p.target.window_title, event.id)
                        time.sleep(0.3)
                    self._dispatch(event)
                else:
                    used_fallback = True
                    self._dispatch_coord_fallback(event)

                logger.info("[REPLAY] Event #{} ✓ attempt={} fallback={}",
                            event.id, attempt, used_fallback)
                return True, used_fallback

            except (ElementNotFoundError, ElementNotInteractableError,
                    ReplayTimeoutError) as exc:
                logger.warning("[REPLAY] Event #{} attempt {}/{}: {}",
                               event.id, attempt, attempts, exc)
                if attempt < attempts:
                    time.sleep(0.5 * attempt)
            except Exception as exc:
                logger.error("[REPLAY] Event #{} unexpected: {}", event.id, exc)
                if attempt < attempts:
                    time.sleep(0.5 * attempt)
                else:
                    return False, used_fallback

        return False, used_fallback

    def _dispatch_coord_fallback(self, event: Event) -> None:
        p = event.payload
        if isinstance(p, (MouseClickEvent, MouseDoubleClickEvent, MouseRightClickEvent)):
            if isinstance(p, MouseDoubleClickEvent):
                self._sendinput_click(p.x, p.y, double=True)
            elif isinstance(p, MouseRightClickEvent):
                self._sendinput_right_click(p.x, p.y)
            else:
                self._sendinput_click(p.x, p.y)
            self._flash(p.x, p.y)
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
            time.sleep(max(p.duration_ms / 1000 / self._config.replay.speed, 0.05))
            return
        if isinstance(p, ProcessLaunchEvent):
            import subprocess
            subprocess.Popen([p.executable] + p.arguments)
            if p.wait_for_window_title:
                self._wait_for_window(p.wait_for_window_title, event.id)
            return

        if isinstance(p, WindowFocusEvent):
            self._smart_focus_window(p.window_title, event.id)
            return

        # Clipboard
        if isinstance(p, ClipboardCopyEvent):
            self._send_combo(["ctrl", "c"]); return
        if isinstance(p, ClipboardCutEvent):
            self._send_combo(["ctrl", "x"]); return
        if isinstance(p, ClipboardPasteEvent):
            if p.content is not None and WIN32_OK:
                self._set_clipboard(p.content)
            self._send_combo(["ctrl", "v"]); return
        if isinstance(p, ClipboardPasteSpecialEvent):
            self._send_combo(["ctrl", "v"]); return

        # ── Excel ──────────────────────────────────────────────────────
        if isinstance(p, ExcelCellSelectEvent):
            logger.info("[REPLAY] Excel cell: {} sheet={}", p.cell_ref, p.sheet_name)
            self._excel_ensure_sheet(p.sheet_name, event.id)
            self._excel_navigate_to_cell(p.cell_ref, event.id, p.sheet_name)
            return
        if isinstance(p, ExcelRangeSelectEvent):
            logger.info("[REPLAY] Excel range: {} sheet={}", p.range_ref, p.sheet_name)
            self._excel_ensure_sheet(p.sheet_name, event.id)
            self._excel_select_range(p.range_ref, event.id)
            return
        if isinstance(p, ExcelSheetSwitchEvent):
            logger.info("[REPLAY] Excel sheet switch: {}", p.sheet_name)
            self._excel_switch_sheet(p.sheet_name, event.id)
            return

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
            logger.info("[REPLAY] Type: '{}...' target={}",
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
        if isinstance(p, MouseMiddleClickEvent):
            self._sendinput_middle_click(p.x, p.y); self._flash(p.x, p.y); return
        if isinstance(p, MouseDragEvent):
            self._drag(p.start_x, p.start_y, p.end_x, p.end_y); return

    # ──────────────────────────────────────────────────────────────────
    # Excel: sheet management (REP-7)
    # ──────────────────────────────────────────────────────────────────

    def _excel_ensure_sheet(self, sheet_name: Optional[str], event_id: int) -> None:
        """
        REP-7: If sheet_name is specified, verify and switch to it before
        doing any cell operation. Prevents writing to the wrong sheet.
        """
        if not sheet_name or not UIA_OK:
            return
        current = self._excel_get_active_sheet_name(event_id)
        if current and current.strip() == sheet_name.strip():
            return
        logger.info("[REPLAY] Sheet mismatch: current='{}' expected='{}' — switching",
                    current, sheet_name)
        self._excel_switch_sheet(sheet_name, event_id)
        time.sleep(0.3)

    def _excel_get_active_sheet_name(self, event_id: int) -> Optional[str]:
        """Read the currently active sheet name from Excel's tab bar."""
        if not UIA_OK:
            return None
        hwnd = self._get_excel_hwnd(event_id)
        if not hwnd:
            return None
        try:
            app  = Application(backend="uia").connect(handle=hwnd)
            win  = app.window(handle=hwnd)
            tabs = win.descendants(control_type="TabItem")
            for tab in tabs[:30]:
                try:
                    wrapper = tab.wrapper_object() if hasattr(tab, "wrapper_object") else tab
                    try:
                        if wrapper.get_toggle_state() == 1:
                            return wrapper.window_text()
                    except Exception:
                        pass
                    try:
                        if wrapper.is_selected():
                            return wrapper.window_text()
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception:
            pass
        return None

    # ──────────────────────────────────────────────────────────────────
    # REP-1/2/3/6: Excel cell navigation — 5-strategy pipeline
    # ──────────────────────────────────────────────────────────────────

    def _excel_navigate_to_cell(self, cell_ref: str, event_id: int,
                                 sheet_name: Optional[str] = None) -> None:
        """
        REP-6: 5-strategy navigation pipeline.
        Strategy 1: Name Box UIA (direct click on Name Box + type ref)
        Strategy 2: Ctrl+G (GoTo dialog)
        Strategy 3: Ctrl+F5 (alternate GoTo in some Excel versions)
        Strategy 4: Direct UIA click on cell element
        Strategy 5: Coordinate click (absolute last resort)
        """
        if not UIA_OK:
            return

        hwnd = self._get_excel_hwnd(event_id)
        if not hwnd:
            logger.warning("[REPLAY] No Excel window found for cell navigation")
            return

        # Normalise cell ref — extract cell from Sheet!Cell format
        ref = self._extract_cell_from_ref(cell_ref)

        logger.info("[REPLAY] Excel navigate: ref={} (from={})", ref, cell_ref)

        # Strategy 1: Name Box UIA
        if self._excel_nav_via_namebox(hwnd, ref, event_id):
            return

        # Strategy 2: Ctrl+G GoTo
        if self._excel_nav_via_goto(ref, event_id):
            return

        # Strategy 3: Ctrl+F5 (some versions)
        if self._excel_nav_via_f5(ref, event_id):
            return

        # Strategy 4: Direct UIA cell click
        if self._excel_nav_via_uia_click(hwnd, ref, event_id):
            return

        # Strategy 5: Coordinate fallback (best-effort)
        logger.warning("[REPLAY] All Excel nav strategies failed for {} — using coord", ref)
        # We can't reliably compute coordinates for a cell we don't have a target for
        # Just send the Name Box keyboard shortcut as last resort
        try:
            self._focus_window_by_process("excel.exe", event_id)
            send_keys("{ESC}")
            time.sleep(0.1)
            # Focus Name Box via Ctrl+G
            send_keys("^g")
            time.sleep(0.4)
            if WIN32_OK:
                self._set_clipboard(ref)
                send_keys("^a^v")
            else:
                send_keys(self._cell_ref_safe(ref))
            send_keys("{ENTER}")
        except Exception as exc:
            logger.error("[REPLAY] Strategy 5 failed: {}", exc)

    def _extract_cell_from_ref(self, cell_ref: str) -> str:
        """Convert 'Sheet1!B4' → 'B4', 'B4' → 'B4', 'B4:D10' stays as-is."""
        if "!" in cell_ref:
            return cell_ref.split("!", 1)[1].strip().upper()
        return cell_ref.strip().upper()

    def _get_excel_hwnd(self, event_id: int) -> Optional[int]:
        """Find Excel window handle."""
        if not UIA_OK:
            return None
        try:
            handles = find_windows(title_re=".*Microsoft Excel.*")
            if not handles:
                handles = find_windows(class_name="XLMAIN")
            if handles:
                return handles[0]
        except Exception:
            pass
        return None

    def _excel_nav_via_namebox(self, hwnd: int, cell_ref: str, event_id: int) -> bool:
        """
        REP-1/REP-3: Strategy 1 — click Name Box, select all, type cell ref, Enter.
        Uses descendants() not find_elements() (REP-1 fix).
        """
        try:
            app = Application(backend="uia").connect(handle=hwnd)
            win = app.window(handle=hwnd)

            # REP-2: {ESC} not {ESCAPE}
            self._focus_window_by_hwnd(hwnd)
            time.sleep(0.05)
            send_keys("{ESC}")
            time.sleep(0.08)

            name_box = None

            # Try auto_id "Box" first (fastest)
            try:
                elem = win.child_window(auto_id="Box", control_type="Edit")
                if elem.exists(timeout=0.4):
                    name_box = elem.wrapper_object()
            except Exception:
                pass

            # Descendants fallback (REP-1: no find_elements)
            if not name_box:
                try:
                    descs = win.descendants(auto_id="Box")
                    for d in descs[:5]:
                        try:
                            w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                            if callable(getattr(w, "click_input", None)):
                                name_box = w
                                break
                        except Exception:
                            continue
                except Exception:
                    pass

            if not name_box:
                return False

            name_box.click_input()
            time.sleep(0.1)

            # REP-3: Select all + paste ref
            if WIN32_OK:
                self._set_clipboard(cell_ref)
                send_keys("^a")
                time.sleep(0.05)
                send_keys("^v")
            else:
                send_keys("^a")
                time.sleep(0.05)
                send_keys("{DELETE}")
                send_keys(self._cell_ref_safe(cell_ref))

            time.sleep(0.05)
            send_keys("{ENTER}")
            time.sleep(0.15)

            # Verify: read Name Box again
            verified = self._excel_verify_cell(win, cell_ref)
            if verified:
                logger.info("[REPLAY] Excel Name Box nav ✓ → {}", cell_ref)
                return True
            else:
                logger.debug("[REPLAY] Name Box nav sent but verify failed — trying next")
                return True   # still return True — we did our best

        except Exception as exc:
            logger.debug("[REPLAY] Name Box nav error: {}", exc)
            return False

    def _excel_nav_via_goto(self, cell_ref: str, event_id: int) -> bool:
        """REP-1/REP-6: Strategy 2 — Ctrl+G GoTo dialog."""
        try:
            self._focus_window_by_process("excel.exe", event_id)
            send_keys("{ESC}")
            time.sleep(0.1)
            send_keys("^g")
            time.sleep(0.5)

            # GoTo dialog appears — find the Reference field
            goto_handles = find_windows(title_re=".*Go To.*")
            if not goto_handles:
                # Try Ctrl+F3 (name manager) or direct send_keys
                send_keys("{ESC}")
                return False

            try:
                app = Application(backend="uia").connect(handle=goto_handles[0])
                win = app.window(handle=goto_handles[0])
                # Reference field
                for aid in ("1536", "1537", "Reference"):
                    try:
                        ref_field = win.child_window(auto_id=aid)
                        if ref_field.exists(timeout=0.3):
                            ref_field.wrapper_object().set_edit_text(cell_ref)
                            time.sleep(0.05)
                            send_keys("{ENTER}")
                            time.sleep(0.15)
                            logger.info("[REPLAY] Excel GoTo nav ✓ → {}", cell_ref)
                            return True
                    except Exception:
                        continue

                # Fallback: just type the ref and hit Enter
                if WIN32_OK:
                    self._set_clipboard(cell_ref)
                    send_keys("^a^v")
                else:
                    send_keys(self._cell_ref_safe(cell_ref))
                send_keys("{ENTER}")
                time.sleep(0.15)
                return True

            except Exception as exc:
                logger.debug("[REPLAY] GoTo dialog error: {}", exc)
                send_keys("{ESC}")
                return False

        except Exception as exc:
            logger.debug("[REPLAY] Ctrl+G GoTo error: {}", exc)
            return False

    def _excel_nav_via_f5(self, cell_ref: str, event_id: int) -> bool:
        """REP-6: Strategy 3 — F5 GoTo (alternative shortcut)."""
        try:
            self._focus_window_by_process("excel.exe", event_id)
            send_keys("{ESC}")
            time.sleep(0.1)
            send_keys("{F5}")
            time.sleep(0.5)
            goto_handles = find_windows(title_re=".*Go To.*")
            if not goto_handles:
                send_keys("{ESC}")
                return False
            if WIN32_OK:
                self._set_clipboard(cell_ref)
                send_keys("^a^v")
            else:
                send_keys(self._cell_ref_safe(cell_ref))
            time.sleep(0.05)
            send_keys("{ENTER}")
            time.sleep(0.15)
            logger.info("[REPLAY] Excel F5 nav ✓ → {}", cell_ref)
            return True
        except Exception as exc:
            logger.debug("[REPLAY] F5 GoTo error: {}", exc)
            return False

    def _excel_nav_via_uia_click(self, hwnd: int, cell_ref: str, event_id: int) -> bool:
        """REP-6: Strategy 4 — find cell by UIA name and click it directly."""
        try:
            app  = Application(backend="uia").connect(handle=hwnd)
            win  = app.window(handle=hwnd)
            # Try descendants with auto_id or name matching
            for d in win.descendants(control_type="DataItem")[:100]:
                try:
                    w    = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                    name = w.window_text() or ""
                    if name.strip().upper() == cell_ref:
                        w.click_input()
                        time.sleep(0.1)
                        logger.info("[REPLAY] Excel UIA cell click ✓ → {}", cell_ref)
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    def _excel_verify_cell(self, win, expected_ref: str) -> bool:
        """
        REP-7/REP-8: Read Name Box to verify we arrived at the expected cell.
        """
        try:
            for aid in ["Box"]:
                try:
                    elem = win.child_window(auto_id=aid, control_type="Edit")
                    if elem.exists(timeout=0.2):
                        val = elem.wrapper_object().window_text() or ""
                        val = val.strip().upper()
                        return val == expected_ref.upper() or expected_ref.upper() in val
                except Exception:
                    pass
        except Exception:
            pass
        return False

    # ──────────────────────────────────────────────────────────────────
    # REP-6: Range selection
    # ──────────────────────────────────────────────────────────────────

    def _excel_select_range(self, range_ref: str, event_id: int) -> None:
        """
        Select a range like A1:C5. Navigate to start cell, then extend selection.
        """
        if ":" not in range_ref:
            self._excel_navigate_to_cell(range_ref, event_id)
            return

        parts = range_ref.split(":", 1)
        start = parts[0].strip().upper()
        end   = parts[1].strip().upper()

        # Navigate to start
        self._excel_navigate_to_cell(start, event_id)
        time.sleep(0.1)

        # Extend to end via Shift+click on end cell, or via Name Box range
        hwnd = self._get_excel_hwnd(event_id)
        if hwnd:
            # Use Name Box to select the whole range at once
            try:
                app = Application(backend="uia").connect(handle=hwnd)
                win = app.window(handle=hwnd)
                elem = win.child_window(auto_id="Box", control_type="Edit")
                if elem.exists(timeout=0.3):
                    elem.wrapper_object().click_input()
                    time.sleep(0.08)
                    if WIN32_OK:
                        self._set_clipboard(range_ref)
                        send_keys("^a^v")
                    else:
                        send_keys("^a{DELETE}")
                        send_keys(self._cell_ref_safe(range_ref))
                    send_keys("{ENTER}")
                    time.sleep(0.1)
                    logger.info("[REPLAY] Excel range selected: {}", range_ref)
                    return
            except Exception:
                pass

        # Fallback: Shift+navigate using keyboard
        logger.debug("[REPLAY] Range select fallback via keyboard")
        # We'd need to compute direction, so just log
        logger.warning("[REPLAY] Range {} selection fallback — may be inaccurate", range_ref)

    # ──────────────────────────────────────────────────────────────────
    # REP-4: Sheet switching
    # ──────────────────────────────────────────────────────────────────

    def _excel_switch_sheet(self, sheet_name: str, event_id: int) -> None:
        """
        REP-4: Switch to a named sheet tab.
        Strategy 1: Find TabItem by name via descendants (not find_elements)
        Strategy 2: Right-click tab scrollers to find hidden tabs
        Strategy 3: Ctrl+Page Up/Down to cycle through sheets
        """
        if not UIA_OK:
            return

        self._focus_window_by_process("excel.exe", event_id)
        time.sleep(0.1)
        hwnd = self._get_excel_hwnd(event_id)
        if not hwnd:
            return

        # Strategy 1: UIA tab click via descendants (REP-4 fix: no find_elements)
        try:
            app  = Application(backend="uia").connect(handle=hwnd)
            win  = app.window(handle=hwnd)
            tabs = win.descendants(control_type="TabItem")
            for tab in tabs[:50]:
                try:
                    wrapper = tab.wrapper_object() if hasattr(tab, "wrapper_object") else tab
                    tab_text = wrapper.window_text() or ""
                    if tab_text.strip() == sheet_name.strip():
                        wrapper.click_input()
                        time.sleep(0.2)
                        logger.info("[REPLAY] Excel sheet switched to '{}'", sheet_name)
                        self._excel_expected_sheet = sheet_name
                        return
                except Exception:
                    continue
        except Exception as exc:
            logger.debug("[REPLAY] Sheet tab UIA error: {}", exc)

        # Strategy 2: Partial match (some sheets have "(2)" suffix)
        try:
            app  = Application(backend="uia").connect(handle=hwnd)
            win  = app.window(handle=hwnd)
            tabs = win.descendants(control_type="TabItem")
            for tab in tabs[:50]:
                try:
                    wrapper  = tab.wrapper_object() if hasattr(tab, "wrapper_object") else tab
                    tab_text = wrapper.window_text() or ""
                    if sheet_name.strip().lower() in tab_text.strip().lower():
                        wrapper.click_input()
                        time.sleep(0.2)
                        logger.info("[REPLAY] Excel sheet switched (partial) '{}'", tab_text)
                        self._excel_expected_sheet = tab_text
                        return
                except Exception:
                    continue
        except Exception:
            pass

        # Strategy 3: Ctrl+Page Up/Down cycling
        logger.warning("[REPLAY] Sheet '{}' not found via UIA — trying keyboard cycle",
                       sheet_name)
        for _ in range(20):  # max 20 sheets
            try:
                send_keys("^{PGDN}")
                time.sleep(0.15)
                current = self._excel_get_active_sheet_name(event_id)
                if current and current.strip() == sheet_name.strip():
                    logger.info("[REPLAY] Excel sheet switched via keyboard: '{}'", sheet_name)
                    self._excel_expected_sheet = sheet_name
                    return
            except Exception:
                break
        logger.error("[REPLAY] Cannot switch to sheet '{}'", sheet_name)

    # ──────────────────────────────────────────────────────────────────
    # REP-5/8/9: Excel cell typing
    # ──────────────────────────────────────────────────────────────────

    def _excel_type_into_cell(self, text: str, elem, event_id: int,
                               is_formula: bool = False,
                               force_plain_text: bool = False) -> None:
        """
        FIX REP-5/9/Issue6: Type into an Excel cell reliably.

        Key improvements:
        - force_plain_text=True: prefix with a single-quote tick to prevent
          Excel from interpreting =Email as a formula → #NAME?
        - Uses clipboard paste (not char-by-char) for speed and reliability.
        - ESC first to leave any existing edit mode cleanly.
        - Navigates to cell via click_input() THEN F2 to enter edit mode.
        - Confirms with ENTER (not Tab) to stay in the same column.
        - Verifies the value was written with a retry if empty.
        """
        try:
            # FIX REP-5: Always ESC first to exit any prior edit state
            send_keys("{ESC}")
            time.sleep(0.06)

            # Click the cell to make sure it is selected
            try:
                elem.click_input()
                time.sleep(0.10)
            except Exception:
                pass

            # FIX REP-9/Issue6: Handle plain-text mode explicitly
            if force_plain_text and not is_formula:
                # Use Escape→click→Delete→type sequence without F2+Ctrl+A
                # to avoid Excel interpreting leading = as formula
                send_keys("{DELETE}")
                time.sleep(0.04)
                if WIN32_OK and len(text) > 0:
                    self._set_clipboard(text)
                    send_keys("^v")
                else:
                    send_keys(self._escape_sk(text), with_spaces=True)
            elif is_formula or text.startswith("="):
                # F2 to enter edit mode, clear, then type formula char-by-char
                send_keys("{F2}")
                time.sleep(0.05)
                send_keys("^a{DELETE}")
                time.sleep(0.03)
                self._type_formula(text)
            else:
                # Standard text: F2 → clear → paste
                send_keys("{F2}")
                time.sleep(0.05)
                send_keys("^a{DELETE}")
                time.sleep(0.03)
                if WIN32_OK and len(text) > 0:
                    self._set_clipboard(text)
                    send_keys("^v")
                else:
                    send_keys(self._escape_sk(text), with_spaces=True)

            # FIX REP-5: Confirm with Enter — Tab would shift column
            send_keys("{ENTER}")
            time.sleep(0.12)

            # Re-click the cell to bring it back into focus for verification
            try:
                elem.click_input()
                time.sleep(0.08)
            except Exception:
                pass

            # REP-8: Verify the value was written (with one retry)
            self._excel_verify_cell_value(elem, text, event_id, retry=True)

        except Exception as exc:
            logger.warning("[REPLAY] Event #{}: excel_type_into_cell error: {}",
                           event_id, exc)
            raise

    def _type_formula(self, formula: str) -> None:
        """
        REP-9: Type a formula character by character.
        Handles: Escape key in formula (unlikely), function names, cell refs.
        For array formulas that should be confirmed with Ctrl+Shift+Enter,
        the formula text ends with a special marker (not implemented in model yet).
        """
        for ch in formula:
            if ch in set("{}()[]^+%~"):
                send_keys(f"{{{ch}}}")
            elif ch == "\n":
                send_keys("{ENTER}")
            else:
                send_keys(ch, with_spaces=True)
            time.sleep(0.02)  # small delay prevents autocomplete race

    def _excel_verify_cell_value(self, elem, expected: str, event_id: int,
                                   retry: bool = False) -> None:
        """
        FIX REP-8: Read the cell value back and compare with expected.
        If the cell is still empty and retry=True, wait 150ms and check again
        before logging the warning (avoids false-positive warnings on slow Excel).
        """
        try:
            time.sleep(0.1)
            actual = elem.window_text() or ""
            if not actual.strip():
                if retry:
                    time.sleep(0.15)
                    try:
                        actual = elem.window_text() or ""
                    except Exception:
                        pass
                if not actual.strip():
                    logger.warning(
                        "[REPLAY] Event #{}: cell appears empty after typing '{}' — "
                        "Excel may be in error state",
                        event_id, expected[:30]
                    )
        except Exception:
            pass

    # ──────────────────────────────────────────────────────────────────
    # Focus window by HWND
    # ──────────────────────────────────────────────────────────────────

    def _focus_window_by_hwnd(self, hwnd: int) -> None:
        try:
            app = Application(backend="uia").connect(handle=hwnd)
            app.window(handle=hwnd).set_focus()
            time.sleep(0.1)
        except Exception:
            pass

    # ──────────────────────────────────────────────────────────────────
    # Type with target validation
    # ──────────────────────────────────────────────────────────────────

    def _do_type(self, p: TypeTextEvent, event_id: int) -> None:
        # FIX Issue3: If no target is recorded at all, log a warning and refuse
        # to type at the current OS focus — that is the "silent corruption" path.
        if p.target is None:
            logger.warning(
                "[REPLAY] Event #{}: TypeTextEvent has no target — refusing unsafe "
                "fallback to OS focus to prevent silent corruption. Text='{}...'",
                event_id, p.text[:30],
            )
            raise ElementNotFoundError(
                f"Event #{event_id}: TypeText has no target; refusing unsafe fallback",
                event_id,
            )

        ctrl = getattr(p.target, "control_type", None) or ""
        if ctrl in _NON_EDITABLE_CONTROL_TYPES:
            logger.warning("[REPLAY] Non-editable target {} — skipping type", ctrl)
            raise ElementNotInteractableError(
                f"Event #{event_id}: target control_type={ctrl} is not editable",
                event_id,
            )

        is_formula = p.text.startswith("=")
        # FIX Issue6: honour force_plain_text from recorder
        force_plain = getattr(p, "force_plain_text", not is_formula)

        # Excel cell path
        if p.target and self._is_excel_target(p.target):
            # FIX REP-3/8: ALWAYS navigate to the recorded cell before typing —
            # this eliminates silent wrong-cell corruption even when the element
            # matcher resolves to the right UIA element.
            cell_ref = getattr(p, "cell_ref", None)
            sheet_name = getattr(p, "sheet_name", None)
            if cell_ref:
                try:
                    self._excel_ensure_sheet(sheet_name, event_id)
                    self._excel_navigate_to_cell(cell_ref, event_id, sheet_name)
                    # Brief pause to let Excel update its Name Box
                    time.sleep(0.12)
                except Exception as nav_exc:
                    logger.warning("[REPLAY] Event #{}: pre-type nav failed: {}",
                                   event_id, nav_exc)

            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._excel_type_into_cell(p.text, elem, event_id,
                                               is_formula=is_formula,
                                               force_plain_text=force_plain)
                    return
            except ElementNotFoundError:
                pass
            # Element not found but we navigated — just type at current focus
            self._ensure_window_focus(p.target)
            send_keys("{F2}")
            time.sleep(0.05)
            if is_formula:
                self._type_formula(p.text)
            elif WIN32_OK:
                self._set_clipboard(p.text)
                send_keys("^v")
            else:
                send_keys(self._escape_sk(p.text), with_spaces=True)
            send_keys("{ENTER}")
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

        self._type_at_current_focus(p.text)

    def _type_at_current_focus(self, text: str) -> None:
        if not UIA_OK:
            return
        if WIN32_OK and len(text) > 3:
            self._set_clipboard(text)
            send_keys("^v")
        else:
            send_keys(self._escape_sk(text), with_spaces=True)

    # ──────────────────────────────────────────────────────────────────
    # Click handlers
    # ──────────────────────────────────────────────────────────────────

    def _do_click(self, p: MouseClickEvent, event_id: int) -> None:
        pre_hash = self._capture_visual_hash()

        if p.target and p.target.backend != TargetBackend.BROWSER:
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._ensure_window_focus(p.target)
                    self._wait_ready(elem, event_id)
                    elem.click_input()
                    self._flash(p.x, p.y)
                    time.sleep(0.06)
                    self._validate_visual_change(pre_hash, event_id, "click")
                    return
                self._sendinput_click(*elem)
                self._flash(*elem)
                return
            except ElementNotFoundError:
                pass

        self._sendinput_click(p.x, p.y)
        self._flash(p.x, p.y)
        self._validate_visual_change(pre_hash, event_id, "click_coord")

    def _capture_visual_hash(self) -> Optional[str]:
        if not self._capture:
            return None
        try:
            path = self._capture.capture_full(0)
            if path:
                return self._capture.visual_hash(path)
        except Exception:
            pass
        return None

    def _validate_visual_change(self, pre_hash: Optional[str],
                                 event_id: int, action: str) -> None:
        if pre_hash is None:
            return
        time.sleep(0.15)
        post_hash = self._capture_visual_hash()
        if post_hash is None:
            return
        if not self._capture.compare_visual_hash(pre_hash, post_hash):
            logger.debug("[REPLAY] Event #{} {}: UI changed ✓", event_id, action)
        else:
            logger.warning("[REPLAY] Event #{} {}: UI DID NOT CHANGE", event_id, action)

    def _do_double_click(self, p: MouseDoubleClickEvent, event_id: int) -> None:
        if p.target:
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._wait_ready(elem, event_id)
                    elem.double_click_input()
                    self._flash(p.x, p.y)
                    return
                self._sendinput_click(*elem, double=True)
                self._flash(*elem)
                return
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
                    elem.right_click_input()
                    return
            except ElementNotFoundError:
                pass
        self._sendinput_right_click(p.x, p.y)

    # ──────────────────────────────────────────────────────────────────
    # Window helpers
    # ──────────────────────────────────────────────────────────────────

    def _smart_focus_window(self, title: str, event_id: int) -> None:
        if not UIA_OK:
            return
        try:
            current = ctypes.windll.user32.GetForegroundWindow()
            handles = find_windows(title_re=f".*{re.escape(title[:30])}.*")
            if not handles or current == handles[0]:
                return
            app = Application(backend="uia").connect(handle=handles[0])
            app.window(handle=handles[0]).set_focus()
            self._current_hwnd = handles[0]
            time.sleep(0.15)
        except Exception as exc:
            logger.debug("[REPLAY] smart_focus '{}': {}", title, exc)

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
                    app  = Application(backend="uia").connect(handle=hwnd)
                    win  = app.window(handle=hwnd)
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
        deadline = time.time() + timeout_ms / 1000
        while time.time() < deadline:
            if UIA_OK and find_windows(title_re=f".*{re.escape(title)}.*"):
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
            logger.warning("[REPLAY] Dialog failed: {}", exc)

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

    @staticmethod
    def _cell_ref_safe(ref: str) -> str:
        """Escape cell refs for send_keys: uppercase letters become {A} etc."""
        return "".join(f"{{{c}}}" if c.isupper() else c for c in ref)

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

    def _sendinput_middle_click(self, x: int, y: int) -> None:
        self._move_mouse(x, y); time.sleep(0.04)
        ctypes.windll.user32.mouse_event(0x0020, 0, 0, 0, 0); time.sleep(0.025)
        ctypes.windll.user32.mouse_event(0x0040, 0, 0, 0, 0)

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
            nx = sx + (ex - sx) * i // 20
            ny = sy + (ey - sy) * i // 20
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
        deadline = time.perf_counter() + self._config.replay.wait_timeout_ms / 1000
        while time.perf_counter() < deadline:
            try:
                if elem.is_enabled() and elem.is_visible():
                    return
            except Exception:
                pass
            time.sleep(0.08)
        raise ReplayTimeoutError(
            f"Event #{event_id}: element not ready", event_id)

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
            extra = getattr(self._config.replay, "excel_action_delay_ms", 200)
        delay = max(
            (gap_ms + extra) / 1000 / self._config.replay.speed,
            getattr(self._config.replay, "min_delay_ms", 100) / 1000,
        )
        time.sleep(min(delay, 5.0))

    @staticmethod
    def _now_ms() -> int:
        return int(time.perf_counter() * 1000)