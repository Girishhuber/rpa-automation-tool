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
    DropdownSelectEvent, CheckboxToggleEvent, RadioSelectEvent,
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


_WM_GETTEXT  = 0x000D
_WM_SETTEXT  = 0x000C
_WM_KEYDOWN  = 0x0100
_WM_KEYUP    = 0x0101
_VK_RETURN   = 0x0D

_NB_HWND_CACHE: dict[int, int] = {}
_FB_HWND_CACHE: dict[int, int] = {}
_HWND_CACHE_LOCK = threading.Lock()

EnumChildProc = ctypes.WINFUNCTYPE(
    ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
)


def _winapi_sendmsg_get(hwnd: int, max_chars: int = 512) -> str:
   
    if not hwnd:
        return ""
    buf = ctypes.create_unicode_buffer(max_chars)
    n = ctypes.windll.user32.SendMessageW(hwnd, _WM_GETTEXT, max_chars, buf)
    return buf.value if n > 0 else ""


def _winapi_sendmsg_set(hwnd: int, text: str) -> bool:
   
    if not hwnd:
        return False
    return bool(ctypes.windll.user32.SendMessageW(hwnd, _WM_SETTEXT, 0, text))


def _winapi_press_enter(hwnd: int) -> None:
    
    ctypes.windll.user32.PostMessageW(hwnd, _WM_KEYDOWN, _VK_RETURN, 0)
    ctypes.windll.user32.PostMessageW(hwnd, _WM_KEYUP,   _VK_RETURN, 0)


def _enum_excel71_children(excel_hwnd: int) -> list[tuple[int, int]]:
    result: list[tuple[int, int]] = []

    def _cb(hwnd, _lp):
        buf = ctypes.create_unicode_buffer(64)
        ctypes.windll.user32.GetClassNameW(hwnd, buf, 64)
        if buf.value == "EXCEL71":
            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
            result.append((hwnd, rect.right - rect.left))
        return True

    cb = EnumChildProc(_cb)
    ctypes.windll.user32.EnumChildWindows(excel_hwnd, cb, 0)
    return result


def _get_namebox_hwnd(excel_hwnd: int) -> int:

    with _HWND_CACHE_LOCK:
        cached = _NB_HWND_CACHE.get(excel_hwnd, 0)
    if cached and _winapi_sendmsg_get(cached, 4) is not None:
        return cached

    children = _enum_excel71_children(excel_hwnd)
    if not children:
        return 0
    children.sort(key=lambda t: t[1])   # ascending width → narrowest = Name Box
    hwnd = children[0][0]
    with _HWND_CACHE_LOCK:
        _NB_HWND_CACHE[excel_hwnd] = hwnd
    return hwnd


def _get_formulabar_hwnd(excel_hwnd: int) -> int:
    """Return the Formula Bar HWND (widest EXCEL71 child, cached)."""
    with _HWND_CACHE_LOCK:
        cached = _FB_HWND_CACHE.get(excel_hwnd, 0)
    if cached and _winapi_sendmsg_get(cached, 4) is not None:
        return cached

    children = _enum_excel71_children(excel_hwnd)
    if not children:
        return 0
    children.sort(key=lambda t: t[1], reverse=True)   # widest = Formula Bar
    hwnd = children[0][0]
    with _HWND_CACHE_LOCK:
        _FB_HWND_CACHE[excel_hwnd] = hwnd
    return hwnd


def _read_namebox(excel_hwnd: int) -> str:
 
    return _winapi_sendmsg_get(_get_namebox_hwnd(excel_hwnd)).strip()


def _read_formulabar(excel_hwnd: int) -> str:

    return _winapi_sendmsg_get(_get_formulabar_hwnd(excel_hwnd), 2048).strip()


def _navigate_namebox_winapi(excel_hwnd: int, cell_ref: str) -> bool:
  
    nb_hwnd = _get_namebox_hwnd(excel_hwnd)
    if not nb_hwnd:
        return False

    # Focus Excel first
    ctypes.windll.user32.SetForegroundWindow(excel_hwnd)
    time.sleep(0.04)

    # Write ref and confirm with Enter
    ref_upper = cell_ref.strip().upper()
    if not _winapi_sendmsg_set(nb_hwnd, ref_upper):
        return False
    time.sleep(0.02)
    _winapi_press_enter(nb_hwnd)
    time.sleep(0.08)

    # Verify
    actual = _read_namebox(excel_hwnd).upper()
    ok = actual == ref_upper or ref_upper in actual
    return ok



_KEY_MAP = {
    "enter":     "{ENTER}",
    "return":    "{ENTER}",
    "tab":       "{TAB}",
    "space":     " ",
    "escape":    "{ESC}",         
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
        self._allow_unanchored_typing_once: bool = False
        # REP-7: track expected active sheet
        self._excel_expected_sheet: Optional[str] = None
        # FIX: track last navigated Excel cell to skip redundant re-navigation
        self._excel_last_cell_ref: Optional[str] = None


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
            strict_targeting=self._config.replay.strict_targeting,
            allow_coordinate_fallback=self._config.replay.allow_coordinate_fallback,
        )

        if self._overlay:
            self._overlay.set_replaying(True)

        events_raw = session.events
        total      = len(events_raw)
        completed  = 0
        start_ms   = self._now_ms()

        logger.info("[REPLAY] Starting: '{}' ({} events)", session.name, total)

        events_raw = self._patch_excel_typetext_lookahead(events_raw)
        # ─────────────────────────────────────────────────────────────────────

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

    def _patch_excel_typetext_lookahead(self, events_raw: list) -> list:

        try:
            # Parse all events once
            parsed: list[tuple[int, object, object]] = []  # (index, raw, payload_or_None)
            for i, raw in enumerate(events_raw):
                try:
                    ev = Event.model_validate(raw)
                    parsed.append((i, raw, ev))
                except Exception:
                    parsed.append((i, raw, None))

            # ── Pass 1: Lookahead cell_ref patch ──────────────────────────────
            for idx, (i, raw, ev) in enumerate(parsed):
                if ev is None:
                    continue
                p = ev.payload
                if not isinstance(p, TypeTextEvent):
                    continue
                cell_ref = getattr(p, "cell_ref", None)
                if cell_ref:
                    continue
                target = getattr(p, "target", None)
                if not target or not self._is_excel_target(target):
                    continue
                ctrl = getattr(target, "control_type", "") or ""
                if ctrl not in ("ListItem", "Pane", "Window", ""):
                    continue
                # Look ahead up to 5 events for ExcelCellSelectEvent
                for j in range(idx + 1, min(idx + 6, len(parsed))):
                    _, _, next_ev = parsed[j]
                    if next_ev is None:
                        continue
                    np = next_ev.payload
                    if isinstance(np, ExcelCellSelectEvent) and np.cell_ref:
                        p.cell_ref = np.cell_ref  # type: ignore[attr-defined]
                        logger.info(
                            "[REPLAY] Pre-patch lookahead: TypeText #{} '{}' → cell_ref={}",
                            ev.id, (p.text or "")[:20], np.cell_ref,
                        )
                        break
                    if isinstance(np, TypeTextEvent):
                        break

            # ── Pass 2: Merge consecutive same-cell TypeText fragments ─────────
            # Merge ONLY when fragments are truly continuations of the same cell:
            # - Same cell_ref
            # - No intervening KeyPressEvent (backspace/delete/enter breaks the chain)
            # - No intervening ExcelCellSelectEvent for a DIFFERENT cell
            result_raw: list = []
            skip_indices: set = set()

            for idx, (i, raw, ev) in enumerate(parsed):
                if idx in skip_indices:
                    continue
                if ev is None:
                    result_raw.append(raw)
                    continue
                p = ev.payload
                if not isinstance(p, TypeTextEvent):
                    result_raw.append(raw)
                    continue

                cell_ref = getattr(p, "cell_ref", None)
                if not cell_ref:
                    result_raw.append(raw)
                    continue

                # Is this Excel?
                target = getattr(p, "target", None)
                if not target or not self._is_excel_target(target):
                    result_raw.append(raw)
                    continue

                merged_text = p.text or ""
                merged_indices = [idx]
                cell_ref_norm = cell_ref.strip().upper()

                j = idx + 1
                while j < len(parsed):
                    _, _, next_ev = parsed[j]
                    if next_ev is None:
                        break
                    np = next_ev.payload

                    if isinstance(np, KeyPressEvent):
                        break

                    # KeyComboEvent — also a hard stop
                    if isinstance(np, KeyComboEvent):
                        break

                    if isinstance(np, ExcelCellSelectEvent):
                        break

                   
                    if isinstance(np, ScreenshotCheckpointEvent):
                        merged_indices.append(j)
                        j += 1
                        continue

                    # Another TypeText for the SAME cell — merge its text
                    if isinstance(np, TypeTextEvent):
                        next_cell = (getattr(np, "cell_ref", None) or "").strip().upper()
                        next_target = getattr(np, "target", None)
                        if (next_cell == cell_ref_norm
                                and next_target
                                and self._is_excel_target(next_target)):
                            merged_text += (np.text or "")
                            merged_indices.append(j)
                            j += 1
                            continue

                    break  # Any other event type — stop

                if len(merged_indices) > 1:
                    type_count = sum(
                        1 for mi in merged_indices
                        if parsed[mi][2] is not None
                        and isinstance(parsed[mi][2].payload, TypeTextEvent)
                    )
                    logger.info(
                        "[REPLAY] Pre-merge: {} TypeText fragments for {} → '{}' ({} events merged)",
                        type_count, cell_ref_norm, merged_text[:40], len(merged_indices),
                    )
                    p.text = merged_text
                    for skip_idx in merged_indices[1:]:
                        skip_indices.add(skip_idx)

                result_raw.append(raw)

            logger.info("[REPLAY] Pre-processing: {} events → {} after Excel merge",
                        len(events_raw), len(result_raw))
            return result_raw

        except Exception as exc:
            logger.warning("[REPLAY] _patch_excel_typetext_lookahead failed: {}", exc)
            return events_raw

    def abort(self) -> None:
        self._abort.set()

    def _cleanup(self) -> None:
        self._browser.disconnect()
        if self._overlay:
            self._overlay.set_replaying(False)

   
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
                    if not self._config.replay.allow_coordinate_fallback:
                        raise ElementNotFoundError(
                            f"Event #{event.id}: coordinate fallback disabled",
                            event.id,
                        )
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


    def _dispatch(self, event: Event) -> None:
        
        p = event.payload

        if isinstance(p, ScreenshotCheckpointEvent):
            return
        if isinstance(p, ExplicitWaitEvent):
            time.sleep(max(p.duration_ms / 1000 / self._config.replay.speed, 0.05))
            return
        if isinstance(p, ProcessLaunchEvent):
            if hasattr(self, "_browser") and self._browser and self._browser.is_connected:
                logger.info("[REPLAY] Using existing browser session")

                if hasattr(p, "arguments") and p.arguments:
                    for arg in p.arguments:
                        if isinstance(arg, str) and arg.startswith("http"):
                            logger.info("[REPLAY] Navigating to {}", arg)
                            self._browser.navigate(arg)
                            return

                logger.warning("[REPLAY] No URL found in ProcessLaunchEvent, skipping launch")
                return

            else:
                logger.error("[REPLAY] Browser not connected — cannot launch Chrome deterministically")
                raise ReplayTimeoutError("Browser not connected")

        if isinstance(p, WindowFocusEvent):
            self._smart_focus_window(p.window_title, event.id)
            return

        # Clipboard
        if isinstance(p, ClipboardCopyEvent):
            if p.content is not None and WIN32_OK:
                self._set_clipboard(p.content)
                logger.info("[REPLAY] Clipboard COPY restored recorded content ({} chars)", len(p.content))
                return
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
            norm = (p.cell_ref or "").strip().upper()
            if norm and norm == self._excel_last_cell_ref:
                logger.info("[REPLAY] Excel cell select {} — already there, skipping nav", norm)
            else:
                self._excel_ensure_sheet(p.sheet_name, event.id)
                self._excel_navigate_to_cell(p.cell_ref, event.id, p.sheet_name)
                self._excel_last_cell_ref = norm
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
        if isinstance(p, RadioSelectEvent):
            self._handle_radio(p, event.id); return

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
            self._allow_unanchored_typing_once = bool(
                p.target is not None and self._is_system_search_target(p.target)
            )
            logger.info("[REPLAY] Click @ ({},{}) target={}", p.x, p.y, self._tlabel(p.target))
            self._do_click(p, event.id); return
        if isinstance(p, MouseDoubleClickEvent):
            self._do_double_click(p, event.id); return
        if isinstance(p, MouseRightClickEvent):
            self._do_right_click(p, event.id); return
        if isinstance(p, MouseScrollEvent):
            self._do_scroll(p, event.id); return
        if isinstance(p, MouseMiddleClickEvent):
            self._sendinput_middle_click(p.x, p.y); self._flash(p.x, p.y); return
        if isinstance(p, MouseDragEvent):
            self._do_drag(p, event.id); return


    def _excel_ensure_sheet(self, sheet_name: Optional[str], event_id: int) -> None:

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

  
    def _excel_navigate_to_cell(self, cell_ref: str, event_id: int,
                                 sheet_name: Optional[str] = None) -> None:
        
        if not UIA_OK:
            return

        hwnd = self._get_excel_hwnd(event_id)
        if not hwnd:
            logger.warning("[REPLAY] No Excel window found for cell navigation")
            return

        ref = self._extract_cell_from_ref(cell_ref)
        logger.info("[REPLAY] Excel navigate: ref={} (from={})", ref, cell_ref)

        # Strategy 1: WinAPI direct (fastest — no UIA at all)
        if _navigate_namebox_winapi(hwnd, ref):
            logger.info("[REPLAY] Excel navigate ✓ WinAPI → {}", ref)
            return

        # Strategy 2: UIA Name Box (with clipboard paste)
        if self._excel_nav_via_namebox(hwnd, ref, event_id):
            return

        # Strategy 3: Ctrl+G GoTo dialog
        if self._excel_nav_via_goto(ref, event_id):
            return

        
        if self._excel_nav_via_f5(ref, event_id):
            return

        if self._excel_nav_via_uia_click(hwnd, ref, event_id):
            return

        logger.warning("[REPLAY] All Excel nav strategies failed for {} — keyboard last resort", ref)
        try:
            self._focus_window_by_hwnd(hwnd)
            send_keys("{ESC}")
            time.sleep(0.08)
            send_keys("^g")
            time.sleep(0.4)
            if WIN32_OK:
                _clip = self._clipboard_guard_begin()
                try:
                    self._set_clipboard(ref)
                    send_keys("^a^v")
                finally:
                    self._clipboard_guard_end(_clip)
            else:
                send_keys(self._cell_ref_safe(ref))
            send_keys("{ENTER}")
        except Exception as exc:
            logger.error("[REPLAY] Last-resort nav failed: {}", exc)

    def _extract_cell_from_ref(self, cell_ref: str) -> str:
        """Convert 'Sheet1!B4' → 'B4', 'B4' → 'B4', 'B4:D10' stays as-is."""
        if "!" in cell_ref:
            return cell_ref.split("!", 1)[1].strip().upper()
        return cell_ref.strip().upper()

    def _get_excel_hwnd(self, event_id: int) -> Optional[int]:
        
        if not UIA_OK:
            return None

        # Class-name lookup is fastest and most reliable
        try:
            handles = find_windows(class_name="XLMAIN")
            if handles:
                return handles[0]
        except Exception:
            pass

        # Title regex fallback (covers WPS Spreadsheets which uses a different title)
        for pattern in (r".*Microsoft Excel.*", r".*Excel.*", r".*WPS.*"):
            try:
                handles = find_windows(title_re=pattern)
                if handles:
                    return handles[0]
            except Exception:
                continue

        return None

    def _excel_nav_via_namebox(self, hwnd: int, cell_ref: str, event_id: int) -> bool:
       

        ref = cell_ref.strip().upper()

        # ── Fast path: WinAPI direct write (no COM, no clipboard) ──────────────
        if _navigate_namebox_winapi(hwnd, ref):
            logger.info("[REPLAY] Excel WinAPI Name Box nav ✓ → {}", ref)
            return True

        # ── Fallback: pywinauto UIA approach ───────────────────────────────────
        try:
            app = Application(backend="uia").connect(handle=hwnd)
            win = app.window(handle=hwnd)

            self._focus_window_by_hwnd(hwnd)
            time.sleep(0.05)
            send_keys("{ESC}")
            time.sleep(0.06)

            name_box = None
            try:
                elem = win.child_window(auto_id="Box", control_type="Edit")
                if elem.exists(timeout=0.3):
                    name_box = elem.wrapper_object()
            except Exception:
                pass

            if not name_box:
                try:
                    for d in win.descendants(auto_id="Box")[:5]:
                        try:
                            w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                            if callable(getattr(w, "click_input", None)):
                                name_box = w; break
                        except Exception:
                            continue
                except Exception:
                    pass

            if not name_box:
                return False

            name_box.click_input()
            time.sleep(0.08)

            # Paste ref via clipboard (avoids send_keys escaping issues)
            if WIN32_OK:
                _clip = self._clipboard_guard_begin()
                try:
                    self._set_clipboard(ref)
                    send_keys("^a^v")
                finally:
                    self._clipboard_guard_end(_clip)
            else:
                send_keys("^a")
                time.sleep(0.04)
                send_keys(self._cell_ref_safe(ref))

            time.sleep(0.04)
            send_keys("{ENTER}")
            time.sleep(0.12)
            logger.info("[REPLAY] Excel UIA Name Box nav ✓ → {}", ref)
            return True

        except Exception as exc:
            logger.debug("[REPLAY] Name Box UIA fallback error: {}", exc)
            return False

    def _excel_nav_via_goto(self, cell_ref: str, event_id: int) -> bool:
    
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
                    _clip = self._clipboard_guard_begin()
                    try:
                        self._set_clipboard(cell_ref)
                        send_keys("^a^v")
                    finally:
                        self._clipboard_guard_end(_clip)
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
                _clip = self._clipboard_guard_begin()
                try:
                    self._set_clipboard(cell_ref)
                    send_keys("^a^v")
                finally:
                    self._clipboard_guard_end(_clip)
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

  
    def _excel_select_range(self, range_ref: str, event_id: int) -> None:

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
                        _clip = self._clipboard_guard_begin()
                        try:
                            self._set_clipboard(range_ref)
                            send_keys("^a^v")
                        finally:
                            self._clipboard_guard_end(_clip)
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

    def _excel_type_into_cell(self, text: str, elem, event_id: int,
                               is_formula: bool = False,
                               force_plain_text: bool = False,
                               excel_hwnd: int = 0) -> None:
        try:
            send_keys("{ESC}")
            time.sleep(0.05)

            # ── Strategy A: direct Formula Bar WinAPI (plain text only) ─────────
            if not is_formula and excel_hwnd:
                fb_hwnd = _get_formulabar_hwnd(excel_hwnd)
                if fb_hwnd:
                    # Click cell to select it
                    try:
                        elem.click_input()
                        time.sleep(0.06)
                    except Exception:
                        pass
                    if _winapi_sendmsg_set(fb_hwnd, text):
                        time.sleep(0.02)
                        _winapi_press_enter(fb_hwnd)
                        time.sleep(0.10)
                    
                        actual = _read_formulabar(excel_hwnd)
                        if actual.strip() == text.strip() or not actual:
                            logger.info("[REPLAY] Excel WinAPI formula-bar write ✓: '{}'", text[:30])
                            return

            try:
                elem.click_input()
                time.sleep(0.08)
            except Exception:
                pass

            send_keys("{DELETE}")
            time.sleep(0.04)

            if is_formula or text.startswith("="):
                # Formulas: type char-by-char to let Excel process them
                send_keys("{F2}")
                time.sleep(0.05)
                send_keys("^a{DELETE}")
                time.sleep(0.03)
                self._type_formula(text)
            elif WIN32_OK:
                self._set_clipboard(text)
                send_keys("^v")
            else:
                send_keys(self._escape_sk(text), with_spaces=True)

            send_keys("{ENTER}")
            time.sleep(0.10)

        except Exception as exc:
            logger.warning("[REPLAY] Event #{}: excel_type_into_cell error: {}",
                           event_id, exc)
            raise

    def _type_formula(self, formula: str) -> None:
     
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

    def _focus_window_by_hwnd(self, hwnd: int) -> None:
        try:
            app = Application(backend="uia").connect(handle=hwnd)
            app.window(handle=hwnd).set_focus()
            time.sleep(0.1)
        except Exception:
            pass

    def _do_type(self, p: TypeTextEvent, event_id: int) -> None:
     
        if p.target is None:
            if self._allow_unanchored_typing_once:
                logger.warning(
                    "[REPLAY] Event #{}: unanchored TypeText allowed for system search. Text='{}'",
                    event_id, p.text[:30],
                )
                self._type_at_current_focus(p.text)
                self._allow_unanchored_typing_once = False
                return
            logger.warning(
                "[REPLAY] Event #{}: TypeTextEvent has no target — blocking unsafe typing. Text='{}'",
                event_id, p.text[:30],
            )
            raise ElementNotFoundError(
                f"Event #{event_id}: TypeTextEvent has no target",
                event_id,
            )

      
        cell_ref_direct = getattr(p, "cell_ref", None)
        if cell_ref_direct and (p.target is None or self._is_excel_target(p.target)):
            logger.info("[REPLAY] Event #{}: TypeText routed via cell_ref={} text='{}'",
                        event_id, cell_ref_direct, p.text[:30])
            excel_hwnd = self._get_excel_hwnd(event_id) or 0
            sheet_name = getattr(p, "sheet_name", None)
            is_formula = p.text.startswith("=")

            # Navigate only if cell changed (skip redundant nav for same-cell fragments)
            norm_ref = cell_ref_direct.strip().upper()
            if norm_ref != self._excel_last_cell_ref:
                try:
                    self._excel_ensure_sheet(sheet_name, event_id)
                    self._excel_navigate_to_cell(cell_ref_direct, event_id, sheet_name)
                    time.sleep(0.05)
                    self._excel_last_cell_ref = norm_ref
                except Exception as nav_exc:
                    logger.warning("[REPLAY] Event #{}: pre-type nav failed: {}", event_id, nav_exc)
            else:
                logger.info("[REPLAY] Event #{}: Excel early-path already at {} — skipping nav",
                            event_id, norm_ref)

            if not is_formula and excel_hwnd:
                fb_hwnd = _get_formulabar_hwnd(excel_hwnd)
                if fb_hwnd and _winapi_sendmsg_set(fb_hwnd, p.text):
                    time.sleep(0.02)
                    _winapi_press_enter(fb_hwnd)
                    time.sleep(0.08)
                    logger.info("[REPLAY] Event #{}: Excel WinAPI type ✓ '{}' → {}",
                                event_id, p.text[:30], cell_ref_direct)
                    return
            # Fallback: type at current focus (navigation already moved us to the right cell)
            self._ensure_window_focus(p.target)
            if is_formula:
                self._type_formula(p.text)
            elif WIN32_OK:
                self._set_clipboard(p.text)
                send_keys("^v")
            else:
                send_keys(self._escape_sk(p.text), with_spaces=True)
            send_keys("{ENTER}")
            return

        ctrl = getattr(p.target, "control_type", None) or ""
        if ctrl in _NON_EDITABLE_CONTROL_TYPES:
            if self._is_system_search_target(p.target):
                logger.warning(
                    "[REPLAY] Event #{}: non-editable system-search target '{}' — typing at focused search box",
                    event_id, ctrl,
                )
                self._type_at_current_focus(p.text)
                self._allow_unanchored_typing_once = False
                return
            logger.warning("[REPLAY] Non-editable target {} — skipping type", ctrl)
            raise ElementNotInteractableError(
                f"Event #{event_id}: target control_type={ctrl} is not editable",
                event_id,
            )

        is_formula  = p.text.startswith("=")
        force_plain = getattr(p, "force_plain_text", not is_formula)

        # ── Excel cell path ──────────────────────────────────────────────────────
        if p.target and self._is_excel_target(p.target):
            cell_ref   = getattr(p, "cell_ref",   None)
            sheet_name = getattr(p, "sheet_name", None)
            excel_hwnd = self._get_excel_hwnd(event_id) or 0

            if cell_ref:
                # FIX: Skip re-navigation if we're already at this cell (consecutive flushes
                # of the same cell produce multiple TypeText events with the same cell_ref).
                norm_ref = cell_ref.strip().upper()
                if norm_ref != self._excel_last_cell_ref:
                    try:
                        self._excel_ensure_sheet(sheet_name, event_id)
                        self._excel_navigate_to_cell(cell_ref, event_id, sheet_name)
                        time.sleep(0.05)
                        self._excel_last_cell_ref = norm_ref
                    except Exception as nav_exc:
                        logger.warning("[REPLAY] Event #{}: pre-type nav failed: {}",
                                       event_id, nav_exc)
                else:
                    logger.info("[REPLAY] Event #{}: Excel already at {} — skipping nav",
                                event_id, norm_ref)
            else:
                # No cell_ref — use the last navigated cell as anchor if available,
                # otherwise read the active cell from the Name Box.
                if self._excel_last_cell_ref:
                    logger.info(
                        "[REPLAY] Event #{}: TypeTextEvent has no cell_ref — "
                        "using last navigated cell '{}' as anchor.",
                        event_id, self._excel_last_cell_ref,
                    )
                    # Re-navigate to ensure we're still at the right cell
                    try:
                        self._excel_navigate_to_cell(
                            self._excel_last_cell_ref, event_id, sheet_name)
                        time.sleep(0.04)
                    except Exception:
                        pass
                elif excel_hwnd:
                    active_cell = _read_namebox(excel_hwnd)
                    logger.warning(
                        "[REPLAY] Event #{}: TypeTextEvent has no cell_ref and no nav history — "
                        "writing to currently active cell '{}'. Check recorder Excel tracking.",
                        event_id, active_cell or "?"
                    )

            # Try WinAPI formula bar write first (fastest, most reliable for plain text)
            # FIX: NEVER call matcher.find() for Excel — UIA cell lookup always times out.
            # After navigation, the correct cell is already active; just write to it.
            if not is_formula and excel_hwnd:
                fb_hwnd = _get_formulabar_hwnd(excel_hwnd)
                if fb_hwnd and _winapi_sendmsg_set(fb_hwnd, p.text):
                    time.sleep(0.02)
                    _winapi_press_enter(fb_hwnd)
                    time.sleep(0.08)
                    logger.info("[REPLAY] Event #{}: Excel WinAPI type ✓ '{}' → {}",
                                event_id, p.text[:30], cell_ref or "active")
                    return

            # Fallback: SendKeys at current focus (cell already navigated to above)
            self._ensure_window_focus(p.target)
            send_keys("{F2}")
            time.sleep(0.04)
            if is_formula:
                self._type_formula(p.text)
            elif WIN32_OK:
                self._set_clipboard(p.text)
                send_keys("^a^v")
            else:
                send_keys(self._escape_sk(p.text), with_spaces=True)
            send_keys("{ENTER}")
            time.sleep(0.06)
            logger.info("[REPLAY] Event #{}: Excel SendKeys type ✓ '{}' → {}",
                        event_id, p.text[:30], cell_ref or "active")
            return

        # ── Standard UIA (non-Excel) ─────────────────────────────────────────────
        if p.target:
            try:
                elem = self._matcher.find(p.target, event_id)
                if isinstance(elem, tuple):
                    # Browser target: matcher returned viewport (cx,cy) coords.
                    # Click to focus the element, then type via CDP key events.
                    if p.target.backend == TargetBackend.BROWSER:
                        vx, vy = elem
                        sx, sy = self._browser.viewport_to_screen(vx, vy)
                        self._browser.bring_to_front()
                        time.sleep(0.08)
                        self._sendinput_click(sx, sy)
                        # FIX: 120ms was insufficient after Gmail compose field click;
                        # Gmail's React re-render can steal focus back briefly.
                        # 200ms lets the DOM settle before CDP key injection.
                        time.sleep(0.20)
                        self._browser.type_text_at(p.text, human_like=False)
                        logger.info("[REPLAY] Event #{}: Browser type ✓ '{}' at viewport({},{})",
                                    event_id, p.text[:30], vx, vy)
                        return
                    # Non-browser tuple fallback — type at raw coords
                    self._sendinput_click(*elem)
                    time.sleep(0.05)
                    self._type_at_current_focus(p.text)
                    return
                # UIA element
                self._wait_ready(elem, event_id)
                try:
                    elem.set_focus()
                    time.sleep(0.1)
                except Exception:
                    pass
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
                # Fallback: click + send_keys (NO clipboard — avoids polluting Win+V history)
                # FIX: increased settle time from 40ms→120ms to prevent cursor-race
                # that causes characters to type backwards or mid-word.
                elem.click_input(); time.sleep(0.12)
                send_keys(self._escape_sk(p.text), with_spaces=True)
                return
            except ElementNotFoundError:
                pass

        # Last resort: type at current OS focus
        self._type_at_current_focus(p.text)

    def _type_at_current_focus(self, text: str) -> None:
        """Type text at the currently focused OS element without touching the clipboard.

        Using clipboard (Ctrl+V) for ordinary typing pollutes Windows clipboard history,
        causing the clipboard viewer (Win+V) to show intermediate strings like cell
        references and application names alongside the real content.  We now send keys
        directly via pywinauto's send_keys for all plain-text typing.  The clipboard is
        still used inside the Excel and formula-bar paths where it is the only reliable
        way to insert content, and for ClipboardPasteEvent replays.

        FIX: Added 80 ms settle delay before sending keys to prevent cursor position
        races that cause characters to be inserted mid-word or in reverse order when
        focus has just changed (e.g. after a click or window switch).
        """
        if not UIA_OK:
            return
        # Ensure any pending focus-change animations have settled before typing.
        time.sleep(0.08)
        send_keys(self._escape_sk(text), with_spaces=True)


    def _do_click(self, p: MouseClickEvent, event_id: int) -> None:
        pre_hash = self._capture_visual_hash()

        if p.target:
            # ── Browser element: resolve via CDP, then SendInput at viewport coords ──
            if p.target.backend == TargetBackend.BROWSER:
                try:
                    coords = self._matcher.find(p.target, event_id)
                    if isinstance(coords, tuple):
                        # coords are viewport-relative; convert to screen coords
                        cx, cy = self._browser.viewport_to_screen(coords[0], coords[1])
                        self._browser.bring_to_front()
                        time.sleep(0.08)
                        self._sendinput_click(cx, cy)
                        self._flash(cx, cy)
                        if not self._validate_visual_change(pre_hash, event_id, "browser_click", require_change=True):
                            raise ElementNotInteractableError(
                                f"Event #{event_id}: browser click produced no detectable UI change",
                                event_id,
                            )
                        return
                except ElementNotFoundError:
                    pass
                # Browser fallback: use recorded raw coords (already screen coords)
                self._browser.bring_to_front()
                time.sleep(0.08)
                self._sendinput_click(p.x, p.y)
                self._flash(p.x, p.y)
                self._validate_visual_change(pre_hash, event_id, "click_coord")
                return

            # ── UIA / native element ──────────────────────────────────────────────
            try:
                elem = self._matcher.find(p.target, event_id)
                if not isinstance(elem, tuple):
                    self._ensure_window_focus(p.target)
                    self._wait_ready(elem, event_id)
                    elem.click_input()
                    self._flash(p.x, p.y)
                    time.sleep(0.06)
                    if not self._validate_visual_change(pre_hash, event_id, "click", require_change=True):
                        raise ElementNotInteractableError(
                            f"Event #{event_id}: click produced no detectable UI change",
                            event_id,
                        )
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
                                 event_id: int, action: str,
                                 require_change: bool = False) -> bool:
        if pre_hash is None:
            return True
        # FIX: was 150ms — some apps (Gmail, Windows Explorer) animate for 200-400ms
        # after a click before the DOM/UI tree actually updates, causing false
        # "UI DID NOT CHANGE" verdicts and false-positive retry clicks.
        time.sleep(0.30)
        post_hash = self._capture_visual_hash()
        if post_hash is None:
            return True
        changed = not self._capture.compare_visual_hash(pre_hash, post_hash)
        if changed:
            logger.debug("[REPLAY] Event #{} {}: UI changed", event_id, action)
            return True
        if require_change:
            logger.warning("[REPLAY] Event #{} {}: UI DID NOT CHANGE — may be false click",
                           event_id, action)
        else:
            logger.debug("[REPLAY] Event #{} {}: UI unchanged (expected for this action)",
                         event_id, action)
        return False

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
        if not p.target:
            return
        if p.target.backend == TargetBackend.BROWSER:
            try:
                elem = self._matcher.find(p.target, event_id)
                if isinstance(elem, tuple):
                    sx, sy = self._browser.viewport_to_screen(elem[0], elem[1])
                    self._browser.bring_to_front()
                    time.sleep(0.08)
                    self._sendinput_click(sx, sy)
                    return
            except ElementNotFoundError:
                return
        if not UIA_OK:
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
        if not p.target:
            return
        if p.target.backend == TargetBackend.BROWSER:
            try:
                elem = self._matcher.find(p.target, event_id)
                if isinstance(elem, tuple):
                    sx, sy = self._browser.viewport_to_screen(elem[0], elem[1])
                    self._browser.bring_to_front()
                    time.sleep(0.08)
                    self._sendinput_click(sx, sy)
                    return
            except ElementNotFoundError:
                return
        if not UIA_OK:
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

    def _handle_radio(self, p, event_id: int) -> None:
        if not p.target:
            return
        if p.target.backend == TargetBackend.BROWSER:
            try:
                elem = self._matcher.find(p.target, event_id)
                if isinstance(elem, tuple):
                    sx, sy = self._browser.viewport_to_screen(elem[0], elem[1])
                    self._browser.bring_to_front()
                    time.sleep(0.08)
                    self._sendinput_click(sx, sy)
                    return
            except ElementNotFoundError:
                return
        if not UIA_OK:
            return
        try:
            elem = self._matcher.find(p.target, event_id)
            if not isinstance(elem, tuple):
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

    def _do_scroll(self, p: MouseScrollEvent, event_id: int) -> None:
        """Replay a scroll event by re-issuing the accumulated wheel delta as
        individual WHEEL_DELTA notches (120 units each).  Splitting the merged
        delta back into per-notch SendInput calls gives applications the same
        stream of WM_MOUSEWHEEL messages they saw at record time, rather than
        one enormous single event that many apps either ignore or mis-handle.

        Focus is ensured on the scroll target window before the first notch so
        that scroll events land in the right window even when the replayer has
        just finished interacting with a different application.
        """
        # Bring the target window into focus before scrolling
        if p.target:
            self._ensure_window_focus(p.target)
            if p.target.backend == TargetBackend.BROWSER and self._browser:
                self._browser.bring_to_front()
                time.sleep(0.08)

        self._move_mouse(p.x, p.y)
        time.sleep(0.04)

        # Split accumulated delta back into individual notches so applications
        # receive the same stream of WM_MOUSEWHEEL messages as at record time.
        dy = p.dy  # accumulated notches (positive = up, negative = down)
        if dy == 0:
            return

        notches = abs(int(dy))
        if notches == 0:
            notches = 1
        direction = 1 if dy > 0 else -1
        wheel_delta = direction * 120  # WHEEL_DELTA per notch

        # Use SendInput (INPUT struct) — more reliable than deprecated mouse_event,
        # works with elevated processes and modern UWP applications.
        INPUT_MOUSE   = 0
        MOUSEEVENTF_WHEEL = 0x0800

        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [
                ("dx",          ctypes.c_long),
                ("dy",          ctypes.c_long),
                ("mouseData",   ctypes.c_ulong),
                ("dwFlags",     ctypes.c_ulong),
                ("time",        ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
            ]

        class _INPUT_UNION(ctypes.Union):
            _fields_ = [("mi", MOUSEINPUT)]

        class INPUT(ctypes.Structure):
            _fields_ = [("type", ctypes.c_ulong), ("_input", _INPUT_UNION)]

        for _ in range(notches):
            inp = INPUT()
            inp.type = INPUT_MOUSE
            inp._input.mi.dx = 0
            inp._input.mi.dy = 0
            inp._input.mi.mouseData = ctypes.c_ulong(wheel_delta & 0xFFFFFFFF)
            inp._input.mi.dwFlags   = MOUSEEVENTF_WHEEL
            inp._input.mi.time      = 0
            inp._input.mi.dwExtraInfo = ctypes.cast(
                ctypes.pointer(ctypes.c_ulong(0)),
                ctypes.POINTER(ctypes.c_ulong)
            )
            ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
            time.sleep(0.02)  # ~50 notches/sec max — natural scroll pace

        logger.debug("[REPLAY] Scroll: {} notches direction={} at ({},{})",
                     notches, direction, p.x, p.y)

    def _do_drag(self, p: MouseDragEvent, event_id: int) -> None:
        """Replay a drag gesture using SendInput MOUSEEVENTF_MOVE events so that
        each intermediate position generates a proper WM_MOUSEMOVE message.

        Improvements over the old implementation:
        - Window focus / browser bring-to-front before the drag starts.
        - Step count scales with distance so both short and long drags have
          smooth, proportional movement (min 10, max 60 steps).
        - Per-step delay is derived from the recorded duration_ms so the drag
          replays at roughly the same speed it was performed.
        - Uses SendInput for mousedown/mouseup instead of deprecated mouse_event.
        - A small settle delay after mousedown lets apps register the press
          (needed for text-selection drags in Word/Chrome).
        """
        sx, sy = p.start_x, p.start_y
        ex, ey = p.end_x,   p.end_y

        # --- Focus the target window first ---
        if p.start_target:
            self._ensure_window_focus(p.start_target)
            if p.start_target.backend == TargetBackend.BROWSER and self._browser:
                self._browser.bring_to_front()
                time.sleep(0.10)

        # --- Compute steps and per-step delay ---
        dist    = ((ex - sx) ** 2 + (ey - sy) ** 2) ** 0.5
        steps   = max(10, min(60, int(dist / 8)))  # ~8px per step, clamped 10-60
        dur_ms  = max(200, getattr(p, "duration_ms", 0) or 300)
        step_ms = dur_ms / steps / 1000           # seconds per step

        # SendInput structures
        INPUT_MOUSE          = 0
        MOUSEEVENTF_MOVE     = 0x0001
        MOUSEEVENTF_LEFTDOWN = 0x0002
        MOUSEEVENTF_LEFTUP   = 0x0004
        MOUSEEVENTF_ABSOLUTE = 0x8000

        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [
                ("dx",          ctypes.c_long),
                ("dy",          ctypes.c_long),
                ("mouseData",   ctypes.c_ulong),
                ("dwFlags",     ctypes.c_ulong),
                ("time",        ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
            ]

        class _INPUT_UNION(ctypes.Union):
            _fields_ = [("mi", MOUSEINPUT)]

        class INPUT(ctypes.Structure):
            _fields_ = [("type", ctypes.c_ulong), ("_input", _INPUT_UNION)]

        def _extra():
            return ctypes.cast(
                ctypes.pointer(ctypes.c_ulong(0)),
                ctypes.POINTER(ctypes.c_ulong)
            )

        def _send_mouse_flag(flag: int, data: int = 0) -> None:
            inp = INPUT()
            inp.type = INPUT_MOUSE
            inp._input.mi.dx          = 0
            inp._input.mi.dy          = 0
            inp._input.mi.mouseData   = ctypes.c_ulong(data)
            inp._input.mi.dwFlags     = flag
            inp._input.mi.time        = 0
            inp._input.mi.dwExtraInfo = _extra()
            ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

        # Get screen dimensions for ABSOLUTE coordinate normalisation
        SM_CXSCREEN, SM_CYSCREEN = 0, 1
        screen_w = ctypes.windll.user32.GetSystemMetrics(SM_CXSCREEN) or 1920
        screen_h = ctypes.windll.user32.GetSystemMetrics(SM_CYSCREEN) or 1080

        def _send_move_absolute(x: int, y: int) -> None:
            # MOUSEEVENTF_ABSOLUTE coords must be in [0, 65535] normalised space
            ax = int(x * 65535 / screen_w)
            ay = int(y * 65535 / screen_h)
            inp = INPUT()
            inp.type = INPUT_MOUSE
            inp._input.mi.dx          = ax
            inp._input.mi.dy          = ay
            inp._input.mi.mouseData   = 0
            inp._input.mi.dwFlags     = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE
            inp._input.mi.time        = 0
            inp._input.mi.dwExtraInfo = _extra()
            ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

        # --- Execute drag ---
        # 1. Move to start without button held
        self._move_mouse(sx, sy)
        time.sleep(0.06)

        # 2. Press left button down
        _send_mouse_flag(MOUSEEVENTF_LEFTDOWN)
        time.sleep(0.08)   # settle — lets text editors register selection start

        # 3. Glide to end position
        for i in range(1, steps + 1):
            nx = int(sx + (ex - sx) * i / steps)
            ny = int(sy + (ey - sy) * i / steps)
            _send_move_absolute(nx, ny)
            time.sleep(max(step_ms, 0.008))

        time.sleep(0.06)   # hold at end briefly

        # 4. Release left button
        _send_mouse_flag(MOUSEEVENTF_LEFTUP)
        time.sleep(0.05)

        self._last_mouse_pos = (ex, ey)
        logger.debug("[REPLAY] Drag ({},{})→({},{}) steps={} dur={}ms",
                     sx, sy, ex, ey, steps, dur_ms)

    # kept for backward compat (called nowhere now but guard against subclass use)
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

    def _get_clipboard_text(self) -> Optional[str]:
        if not WIN32_OK:
            return None
        try:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        except Exception:
            return None
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
        return None

    def _clipboard_guard_begin(self) -> Optional[str]:
        return self._get_clipboard_text()

    def _clipboard_guard_end(self, original: Optional[str]) -> None:
        if original is None:
            return
        self._set_clipboard(original)

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
    def _is_system_search_target(target) -> bool:
        if not target:
            return False
        proc = (getattr(target, "process_name", "") or "").lower()
        win = (getattr(target, "window_title", "") or "").lower()
        name = (getattr(target, "name", "") or "").lower()
        if proc in {"explorer.exe", "searchhost.exe", "searchapp.exe"}:
            return True
        if win in {"taskbar", "search"}:
            return True
        return "search" in name

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