
from __future__ import annotations
import ctypes
import re
import threading
import time
from collections import OrderedDict
from typing import Optional, Any

from utils.logger import logger
from models.target import UITarget, BoundingBox, TargetBackend
from .selector import Selector, SelectorBuilder, AnchorElement

try:
    import pywinauto
    from pywinauto import Desktop
    PYWINAUTO_OK = True
except ImportError:
    PYWINAUTO_OK = False

try:
    import psutil
    PSUTIL_OK = True
except ImportError:
    PSUTIL_OK = False

# Process classifications
BROWSER_PROCS = {"chrome.exe", "msedge.exe", "firefox.exe",
                 "brave.exe", "opera.exe", "vivaldi.exe"}
EXCEL_PROCS   = {"excel.exe"}
OFFICE_PROCS  = {"excel.exe", "winword.exe", "powerpnt.exe", "outlook.exe"}

_ELECTRON_CLASS = {"Chrome_WidgetWin_1", "CefBrowserWindow"}
_CELL_RE        = re.compile(r"^[A-Z]{1,3}[0-9]{1,7}$")

# UIA-1: wrapper cache config
_CACHE_RADIUS_PX  = 5     # pixels — if new point within this of cached → reuse
_CACHE_MAX_ENTRIES = 32   # LRU size
_CACHE_TTL_S       = 0.5  # seconds — cached wrapper expires after this

# UIA-4: per-call timeout
_UIA_CALL_TIMEOUT  = 0.8  # seconds


# ─────────────────────────────────────────────────────────────────────────────
# Per-monitor DPI (UIA-2)
# ─────────────────────────────────────────────────────────────────────────────

def _get_dpi_for_point(x: int, y: int) -> float:
    """
    UIA-2: Query DPI for the monitor containing point (x, y).
    Uses GetDpiForMonitor (Windows 8.1+). Falls back to primary monitor DPI.
    """
    try:
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        MonitorFromPoint = ctypes.windll.user32.MonitorFromPoint
        GetDpiForMonitor = ctypes.windll.shcore.GetDpiForMonitor
        hmon  = MonitorFromPoint(POINT(x, y), 2)  # MONITOR_DEFAULTTONEAREST
        dpi_x = ctypes.c_uint()
        dpi_y = ctypes.c_uint()
        if GetDpiForMonitor(hmon, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y)) == 0:
            return dpi_x.value / 96.0
    except Exception:
        pass
    # Fallback: primary monitor DPI
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return round(dpi / 96.0, 2)
    except Exception:
        return 1.0


# ─────────────────────────────────────────────────────────────────────────────
# Wrapper cache entry
# ─────────────────────────────────────────────────────────────────────────────

class _WrapperCacheEntry:
    __slots__ = ("x", "y", "wrapper", "ts")

    def __init__(self, x: int, y: int, wrapper):
        self.x       = x
        self.y       = y
        self.wrapper = wrapper
        self.ts      = time.monotonic()

    def is_fresh(self) -> bool:
        return (time.monotonic() - self.ts) < _CACHE_TTL_S

    def near(self, x: int, y: int) -> bool:
        return abs(self.x - x) <= _CACHE_RADIUS_PX and abs(self.y - y) <= _CACHE_RADIUS_PX


class UIAEnricher:

    def __init__(self):
        if not PYWINAUTO_OK:
            logger.warning("pywinauto not installed — UIA enrichment disabled")

        # Desktop cache
        self._desktop_cache: Optional[Any] = None
        self._desktop_lock   = threading.Lock()

        # UIA-1: wrapper LRU cache
        self._wrapper_cache: OrderedDict[int, _WrapperCacheEntry] = OrderedDict()
        self._cache_lock     = threading.Lock()
        self._cache_counter  = 0

    # ──────────────────────────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────────────────────────

    def get_target_at(self, x: int, y: int) -> Optional[UITarget]:
        """Build enriched UITarget at screen coords. Returns None on failure."""
        if not PYWINAUTO_OK:
            return None
        wrapper = self._get_wrapper_at(x, y)
        if wrapper is None:
            return None
        return self._build_uia_target(wrapper, x, y)

    def get_selector_at(self, x: int, y: int) -> Optional[Selector]:
        """Build a full Selector (UiPath-style) at screen coords."""
        if not PYWINAUTO_OK:
            return None
        wrapper = self._get_wrapper_at(x, y)
        if wrapper is None:
            return None
        return self._build_selector(wrapper, x, y)

    def get_focused_element(self) -> Optional[UITarget]:
        """UIA-5: Return UITarget for the currently focused element."""
        if not PYWINAUTO_OK:
            return None
        try:
            desktop = self._get_desktop()
            focused = desktop.get_active()
            if focused:
                wrapper = (focused.wrapper_object()
                           if hasattr(focused, "wrapper_object") else focused)
                if self._is_visible(wrapper):
                    return self._build_uia_target(wrapper, 0, 0)
        except Exception:
            pass
        return None

    def get_window_info(self, hwnd: int) -> dict:
        info = {"title": "", "process": "", "x": 0, "y": 0, "w": 0, "h": 0}
        try:
            buf = ctypes.create_unicode_buffer(512)
            ctypes.windll.user32.GetWindowTextW(hwnd, buf, 512)
            info["title"] = buf.value

            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
            info["x"] = rect.left
            info["y"] = rect.top
            info["w"] = rect.right  - rect.left
            info["h"] = rect.bottom - rect.top

            pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if PSUTIL_OK:
                try:
                    info["process"] = psutil.Process(pid.value).name()
                except Exception:
                    pass
        except Exception:
            pass
        return info

    def is_browser_window(self, x: int, y: int) -> bool:
        """Check via wrapper already in cache if possible (no extra UIA call)."""
        entry = self._cache_lookup(x, y)
        wrapper = entry.wrapper if entry else self._get_wrapper_at(x, y)
        if not wrapper:
            return False
        proc = self._get_process_name(wrapper)
        cls  = self._safe(wrapper, "class_name") or ""
        return (proc and proc.lower() in BROWSER_PROCS) or cls in _ELECTRON_CLASS

    def is_excel_window(self, x: int, y: int) -> bool:
        entry = self._cache_lookup(x, y)
        wrapper = entry.wrapper if entry else self._get_wrapper_at(x, y)
        if not wrapper:
            return False
        proc = self._get_process_name(wrapper)
        return proc.lower() in EXCEL_PROCS if proc else False

    # ──────────────────────────────────────────────────────────────────
    # DPI (UIA-2)
    # ──────────────────────────────────────────────────────────────────

    @staticmethod
    def _get_dpi() -> float:
        """Primary monitor DPI (used as fallback)."""
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return round(dpi / 96.0, 2)
        except Exception:
            return 1.0

    # ──────────────────────────────────────────────────────────────────
    # Wrapper retrieval (UIA-1, UIA-4, UIA-5)
    # ──────────────────────────────────────────────────────────────────

    def _get_wrapper_at(self, x: int, y: int):
        """
        UIA-1: Return cached wrapper if recent + close enough, else fetch new.
        UIA-4: If fetch blocks > timeout, try focused element fallback.
        UIA-5: Focused element as last resort when from_point fails.
        """
        if not PYWINAUTO_OK:
            return None

        # UIA-1: cache lookup
        entry = self._cache_lookup(x, y)
        if entry:
            logger.debug("UIA cache hit @ ({},{})", x, y)
            return entry.wrapper

        result     = [None]
        exc_holder = [None]
        done_event = threading.Event()

        def _fetch():
            try:
                desktop = self._get_desktop()
                element = desktop.from_point(x, y)
                if element:
                    wrapper = (element.wrapper_object()
                               if hasattr(element, "wrapper_object") else element)
                    result[0] = wrapper
            except Exception as e:
                exc_holder[0] = e
            finally:
                done_event.set()

        t = threading.Thread(target=_fetch, daemon=True)
        t.start()
        finished = done_event.wait(timeout=_UIA_CALL_TIMEOUT)

        if not finished:
            # UIA-4: thread hung (UIA deadlock) — do not join, use focused fallback
            logger.debug("UIA from_point({},{}) timed out — trying focused element", x, y)
            return self._focused_element_fallback()

        if exc_holder[0]:
            logger.debug("UIA from_point({},{}) error: {}", x, y, exc_holder[0])
            # UIA-5: on error, try focused element
            focused = self._focused_element_fallback()
            if focused:
                return focused

        wrapper = result[0]
        if wrapper:
            # UIA-1: store in cache
            self._cache_put(x, y, wrapper)

        return wrapper

    def _focused_element_fallback(self):
        """UIA-5: Try to get the currently focused element."""
        try:
            desktop = self._get_desktop()
            focused = desktop.get_active()
            if focused:
                wrapper = (focused.wrapper_object()
                           if hasattr(focused, "wrapper_object") else focused)
                if self._is_visible(wrapper):
                    logger.debug("UIA-5: using focused element fallback")
                    return wrapper
        except Exception:
            pass
        return None

    # ──────────────────────────────────────────────────────────────────
    # Cache helpers (UIA-1)
    # ──────────────────────────────────────────────────────────────────

    def _cache_lookup(self, x: int, y: int) -> Optional[_WrapperCacheEntry]:
        with self._cache_lock:
            for key, entry in list(self._wrapper_cache.items()):
                if not entry.is_fresh():
                    del self._wrapper_cache[key]
                    continue
                if entry.near(x, y):
                    # Move to end (LRU)
                    self._wrapper_cache.move_to_end(key)
                    return entry
        return None

    def _cache_put(self, x: int, y: int, wrapper) -> None:
        with self._cache_lock:
            self._cache_counter += 1
            key = self._cache_counter
            if len(self._wrapper_cache) >= _CACHE_MAX_ENTRIES:
                self._wrapper_cache.popitem(last=False)   # evict LRU
            self._wrapper_cache[key] = _WrapperCacheEntry(x, y, wrapper)

    # ──────────────────────────────────────────────────────────────────
    # Target / selector builders
    # ──────────────────────────────────────────────────────────────────

    def _build_uia_target(self, wrapper, x: int, y: int) -> UITarget:
        auto_id    = self._safe(wrapper, "automation_id")
        name       = self._safe(wrapper, "window_text")
        ctrl_type  = self._safe(wrapper, "friendly_class_name")
        class_name = self._safe(wrapper, "class_name")
        win_title  = self._get_window_title(wrapper)
        proc_name  = self._get_process_name(wrapper)

        bbox        = self._get_normalised_bbox(wrapper, x, y)
        ancestors   = self._get_rich_ancestors(wrapper, depth=4)
        caps        = self._get_capabilities(wrapper)

        backend = TargetBackend.UIA
        if proc_name and proc_name.lower() in BROWSER_PROCS:
            backend = TargetBackend.BROWSER
        elif class_name in _ELECTRON_CLASS:
            backend = TargetBackend.BROWSER

        # UIA-2: use per-element DPI
        dpi = _get_dpi_for_point(x, y)

        return UITarget(
            backend       = backend,
            automation_id = (auto_id or None) if not _is_unstable(auto_id or "") else None,
            name          = (name or "")[:200] or None,
            control_type  = ctrl_type or None,
            class_name    = class_name or None,
            window_title  = win_title or None,
            process_name  = proc_name or None,
            bbox          = bbox,
            screen_x      = x,
            screen_y      = y,
            dpi_scale     = dpi,
            ancestor_chain = ancestors,
        )

    def _build_selector(self, wrapper, x: int, y: int) -> Selector:
        auto_id    = self._safe(wrapper, "automation_id")
        name       = self._safe(wrapper, "window_text")
        ctrl_type  = self._safe(wrapper, "friendly_class_name")
        class_name = self._safe(wrapper, "class_name")
        win_title  = self._get_window_title(wrapper)
        proc_name  = self._get_process_name(wrapper)
        bbox        = self._get_normalised_bbox(wrapper, x, y)
        ancestors   = self._get_rich_ancestors(wrapper, depth=4)
        sibling_idx = self._get_sibling_index(wrapper)
        caps        = self._get_capabilities(wrapper)
        # SEL-3: build anchor elements from siblings/parent
        anchors     = self._get_anchor_elements(wrapper)
        dpi         = _get_dpi_for_point(x, y)

        return SelectorBuilder.from_uia(
            automation_id   = auto_id,
            name            = name,
            control_type    = ctrl_type,
            class_name      = class_name,
            window_title    = win_title,
            process_name    = proc_name,
            bbox            = bbox,
            ancestor_chain  = ancestors,
            sibling_index   = sibling_idx,
            screen_x        = x,
            screen_y        = y,
            dpi_scale       = dpi,
            capabilities    = caps,
            anchor_elements = anchors,
        )

    def _get_normalised_bbox(self, wrapper, elem_x: int = 0, elem_y: int = 0) -> Optional[BoundingBox]:
        """
        UIA-2: Normalise bbox using per-element DPI (not global primary DPI).
        """
        try:
            r   = wrapper.rectangle()
            # Use element centre to determine which monitor it's on
            cx  = (r.left + r.right)  // 2 or elem_x
            cy  = (r.top  + r.bottom) // 2 or elem_y
            dpi = _get_dpi_for_point(cx, cy)
            return BoundingBox(
                left   = round(r.left   / dpi),
                top    = round(r.top    / dpi),
                right  = round(r.right  / dpi),
                bottom = round(r.bottom / dpi),
            )
        except Exception:
            return None

    def _get_rich_ancestors(self, wrapper, depth: int = 4) -> list[str]:
        """Returns ancestor chain with colon-safe format."""
        chain = []
        try:
            current = wrapper
            for _ in range(depth):
                parent = current.parent()
                if not parent:
                    break
                ptype = self._safe(parent, "friendly_class_name") or "?"
                ptext = (self._safe(parent, "window_text") or "")[:30]
                paid  = self._safe(parent, "automation_id") or ""
                entry = f"{ptype}:{ptext}"
                if paid and not _is_unstable(paid):
                    entry += f":{paid}"
                chain.append(entry)
                current = parent
        except Exception:
            pass
        return chain

    def _get_sibling_index(self, wrapper) -> Optional[int]:
        """Position among same-type siblings (0-based)."""
        try:
            parent    = wrapper.parent()
            if not parent:
                return None
            ctrl_type = self._safe(wrapper, "friendly_class_name")
            idx       = 0
            for child in parent.children():
                if child == wrapper:
                    return idx
                if self._safe(child, "friendly_class_name") == ctrl_type:
                    idx += 1
        except Exception:
            pass
        return None

    def _get_anchor_elements(self, wrapper) -> list[AnchorElement]:
        """
        SEL-3: Find up to 2 nearby stable anchor elements (label, heading)
        that can help the matcher disambiguate this element.
        """
        anchors = []
        try:
            parent = wrapper.parent()
            if not parent:
                return anchors
            rect_self = wrapper.rectangle()
            cx_self   = (rect_self.left + rect_self.right)  / 2
            cy_self   = (rect_self.top  + rect_self.bottom) / 2

            for sibling in parent.children():
                try:
                    if sibling == wrapper:
                        continue
                    sib_type = self._safe(sibling, "friendly_class_name") or ""
                    # Only consider stable, visible label-like elements as anchors
                    if sib_type.lower() not in ("text", "label", "statictext",
                                                "textblock", "static"):
                        continue
                    sib_name = self._safe(sibling, "window_text") or ""
                    if not sib_name or len(sib_name) < 2 or _is_unstable(sib_name):
                        continue
                    r    = sibling.rectangle()
                    cx_s = (r.left + r.right)  / 2
                    cy_s = (r.top  + r.bottom) / 2
                    dx   = int(cx_s - cx_self)
                    dy   = int(cy_s - cy_self)
                    # Determine relative direction
                    if abs(dx) > abs(dy):
                        direction = "left" if dx < 0 else "right"
                    else:
                        direction = "above" if dy < 0 else "below"
                    anchors.append(AnchorElement(
                        direction    = direction,
                        name         = sib_name[:60],
                        control_type = sib_type,
                        offset_x     = dx,
                        offset_y     = dy,
                    ))
                    if len(anchors) >= 2:
                        break
                except Exception:
                    continue
        except Exception:
            pass
        return anchors

    def _get_capabilities(self, wrapper) -> dict:
        caps = {
            "is_editable":   False,
            "is_clickable":  True,
            "is_toggleable": False,
            "is_selectable": False,
            "is_invokable":  False,
        }
        try:
            patterns = wrapper.get_patterns()
            pnames   = [p.lower() for p in (patterns or [])]
            caps["is_editable"]   = any("value" in p or "text" in p   for p in pnames)
            caps["is_toggleable"] = any("toggle" in p                  for p in pnames)
            caps["is_selectable"] = any("select" in p                  for p in pnames)
            caps["is_invokable"]  = any("invoke" in p                  for p in pnames)
            caps["is_clickable"]  = caps["is_invokable"] or True
        except Exception:
            pass
        try:
            caps["is_editable"] = caps["is_editable"] or (
                wrapper.is_editable() if hasattr(wrapper, "is_editable") else False
            )
        except Exception:
            pass
        return caps

    def _is_visible(self, wrapper) -> bool:
        """
        UIA-3: Strict visibility — checks is_visible(), non-zero rect,
        AND that the element is actually on-screen (not offscreen/virtual).
        """
        try:
            if not wrapper.is_visible():
                return False
            r = wrapper.rectangle()
            # Non-zero size
            if r.right - r.left <= 0 or r.bottom - r.top <= 0:
                return False
            # Not fully offscreen to the left/above
            if r.right < -50 or r.bottom < -50:
                return False
            # Not in virtual/hidden space far off-screen
            if r.left > 16000 or r.top > 16000:
                return False
            # Not the degenerate (0,0,0,0) rectangle
            if r.left == 0 and r.top == 0 and r.right == 0 and r.bottom == 0:
                return False
            return True
        except Exception:
            return True   # assume visible on error

    # ──────────────────────────────────────────────────────────────────
    # Desktop cache
    # ──────────────────────────────────────────────────────────────────

    def _get_desktop(self):
        with self._desktop_lock:
            if self._desktop_cache is None:
                self._desktop_cache = Desktop(backend="uia")
            return self._desktop_cache

    # ──────────────────────────────────────────────────────────────────
    # Low-level helpers
    # ──────────────────────────────────────────────────────────────────

    def _get_window_title(self, wrapper) -> Optional[str]:
        try:
            return wrapper.top_level_parent().window_text() or None
        except Exception:
            return None

    def _get_process_name(self, wrapper) -> Optional[str]:
        try:
            if PSUTIL_OK:
                pid = wrapper.process_id()
                return psutil.Process(pid).name()
        except Exception:
            pass
        return None

    @staticmethod
    def _safe(wrapper, attr: str) -> Optional[str]:
        try:
            v   = getattr(wrapper, attr)
            val = v() if callable(v) else v
            return str(val) if val is not None else None
        except Exception:
            return None


# ─────────────────────────────────────────────────────────────────────────────
# Module-level helpers used by recorder / selector
# ─────────────────────────────────────────────────────────────────────────────

def _is_unstable(value: str) -> bool:
    """Shared unstable-ID detector (same logic as selector.py)."""
    if not value:
        return True
    _KNOWN = {"1148", "1001", "1000", "100", "101", "1", "2", "3", "4"}
    if value in _KNOWN:
        return False
    if len(value) <= 3 and value.isalnum():
        return False
    patterns = [
        re.compile(r"^[0-9]{6,}$"),
        re.compile(r"\b[a-f0-9]{8,}\b"),   # SEL-4 fix: word-boundary hex
        re.compile(r"^_\w+\d{4,}$"),
        re.compile(r"-[a-f0-9]{6,}$"),
        re.compile(r"_[a-f0-9]{6,}$"),
    ]
    return any(p.search(value) for p in patterns)


def detect_excel_cell(name: Optional[str], ctrl_type: Optional[str]) -> Optional[str]:
    if not name or not ctrl_type:
        return None
    if ctrl_type in ("DataItem", "Cell", "SpreadsheetItem", "Edit"):
        clean = name.strip().upper()
        if _CELL_RE.match(clean):
            return clean
    return None
