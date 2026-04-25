
from __future__ import annotations

import ctypes
import ctypes.wintypes
import hashlib
import re
import threading
import time
from collections import OrderedDict
from enum import Enum
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


BROWSER_PROCS   = {"chrome.exe", "msedge.exe", "firefox.exe",
                    "brave.exe", "opera.exe", "vivaldi.exe"}
EXCEL_PROCS     = {"excel.exe"}
OFFICE_PROCS    = {"excel.exe", "winword.exe", "powerpnt.exe", "outlook.exe"}
_ELECTRON_CLASS = {"Chrome_WidgetWin_1", "CefBrowserWindow"}
_CELL_RE        = re.compile(r"^[A-Z]{1,3}[0-9]{1,7}$")

# ── Cache config ───────────────────────────────────────────────────────────
# FIX: Use region-based caching (8 px radius) instead of pixel-exact.
#      TTL increased to 1.5 s (was 0.5 s).
_CACHE_RADIUS_PX   = 8     # was 5
_CACHE_MAX_ENTRIES = 48    # slightly larger LRU
_CACHE_TTL_S       = 1.5   # was 0.5 — too short, causing excessive recomputation
_UIA_CALL_TIMEOUT  = 0.8

# ── PID → process-name cache ───────────────────────────────────────────────
_PID_NAME_CACHE: dict[int, str] = {}
_PID_NAME_LOCK  = threading.Lock()
_PID_CACHE_MAX  = 300     # was 200

# ── Element type priority ──────────────────────────────────────────────────
_ELEMENT_TYPE_PRIORITY: dict[str, int] = {
    "Button":          90,
    "SplitButton":     88,
    "Edit":            85,
    "Document":        85,
    "ComboBox":        80,
    "CheckBox":        80,
    "RadioButton":     80,
    "SpreadsheetItem": 78,   # Excel cells — high priority
    "DataItem":        75,
    "ListItem":        70,
    "TreeItem":        70,
    "TabItem":         65,
    "MenuItem":        65,
    "Text":            40,
    "Label":           40,
    "StaticText":      40,
    "Pane":            30,
    "Group":           25,
    "ToolBar":         20,
    "StatusBar":       15,
    "ScrollBar":       10,
}

_ANCESTOR_STOP_TYPES = {
    "Window", "Dialog", "Pane", "Frame", "CustomControl",
}


class ConfidenceLevel(str, Enum):
    HIGH   = "high"
    MEDIUM = "medium"
    LOW    = "low"


class VisibilityScore(str, Enum):
    FULL      = "full"
    PARTIAL   = "partial"
    OFFSCREEN = "offscreen"
    HIDDEN    = "hidden"


# ── DPI helpers ───────────────────────────────────────────────────────────

def _get_dpi_for_point(x: int, y: int) -> float:
    try:
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        hmon  = ctypes.windll.user32.MonitorFromPoint(POINT(x, y), 2)
        dpi_x = ctypes.c_uint()
        dpi_y = ctypes.c_uint()
        if ctypes.windll.shcore.GetDpiForMonitor(hmon, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y)) == 0:
            return dpi_x.value / 96.0
    except Exception:
        pass
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return round(dpi / 96.0, 2)
    except Exception:
        return 1.0


# ── SafeElement ───────────────────────────────────────────────────────────

class SafeElement:
    """Thin wrapper that auto-refreshes stale pywinauto wrappers."""

    def __init__(self, wrapper, hwnd: int = 0, auto_id: str = "", name: str = ""):
        self._wrapper      = wrapper
        self._hwnd         = hwnd
        self._auto_id      = auto_id
        self._name         = name
        self._refresh_lock = threading.Lock()

    def _refresh(self) -> bool:
        if not PYWINAUTO_OK or not self._hwnd:
            return False
        try:
            from pywinauto import Application
            app = Application(backend="uia").connect(handle=self._hwnd)
            win = app.window(handle=self._hwnd)
            if self._auto_id:
                elem = win.child_window(auto_id=self._auto_id)
                if elem.exists(timeout=0.5):
                    fresh = elem.wrapper_object()
                    if fresh:
                        self._wrapper = fresh
                        return True
            if self._name:
                elem = win.child_window(title_re=f".*{re.escape(self._name[:30])}.*")
                if elem.exists(timeout=0.5):
                    fresh = elem.wrapper_object()
                    if fresh:
                        self._wrapper = fresh
                        return True
        except Exception:
            pass
        return False

    def __getattr__(self, name: str):
        attr = getattr(self._wrapper, name)
        if not callable(attr):
            return attr

        def _call(*args, **kwargs):
            try:
                return attr(*args, **kwargs)
            except Exception:
                with self._refresh_lock:
                    if self._refresh():
                        try:
                            return getattr(self._wrapper, name)(*args, **kwargs)
                        except Exception:
                            pass
                raise

        return _call

    def raw(self):
        return self._wrapper


# ── Cache entry ───────────────────────────────────────────────────────────

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
        """FIX: Region-based proximity (8 px) instead of pixel-exact (5 px)."""
        return abs(self.x - x) <= _CACHE_RADIUS_PX and abs(self.y - y) <= _CACHE_RADIUS_PX

    def is_valid(self) -> bool:
        if not self.is_fresh():
            return False
        try:
            r = self.wrapper.rectangle()
            return r is not None
        except Exception:
            return False


# ── UIAEnricher ───────────────────────────────────────────────────────────

class UIAEnricher:

    def __init__(self):
        if not PYWINAUTO_OK:
            logger.warning("pywinauto not installed — UIA enrichment disabled")

        self._desktop_cache: Optional[Any] = None
        self._desktop_lock   = threading.Lock()

        self._wrapper_cache: OrderedDict[int, _WrapperCacheEntry] = OrderedDict()
        self._cache_lock     = threading.Lock()
        self._cache_counter  = 0

        # Last-known wrapper (multi-layer fallback layer 3)
        self._last_wrapper: Optional[Any] = None

    def get_target_at(self, x: int, y: int) -> Optional[UITarget]:
        if not PYWINAUTO_OK:
            return None
        wrapper = self._get_wrapper_at(x, y)
        if wrapper is None:
            return None
        return self._build_uia_target(wrapper, x, y)

    def get_selector_at(self, x: int, y: int) -> Optional[Selector]:
        if not PYWINAUTO_OK:
            return None
        wrapper = self._get_wrapper_at(x, y)
        if wrapper is None:
            return None
        return self._build_selector(wrapper, x, y)

    def get_focused_element(self, context_target: Optional[UITarget] = None) -> Optional[UITarget]:
        """
        Return UITarget for the currently focused element.
        Validates against context_target's window/process if provided.
        """
        if not PYWINAUTO_OK:
            return None
        try:
            desktop = self._get_desktop()
            focused = desktop.get_active()
            if not focused:
                return None
            wrapper = (focused.wrapper_object()
                       if hasattr(focused, "wrapper_object") else focused)
            if not self._is_visible(wrapper):
                return None
            if context_target:
                if not self._same_context(wrapper, context_target):
                    logger.debug("[ENRICHER] Focused element not in expected window — discarding")
                    return None
            return self._build_uia_target(wrapper, 0, 0)
        except Exception:
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
            info["process"] = _cached_process_name(pid.value)
        except Exception:
            pass
        return info

    def is_browser_window(self, x: int, y: int) -> bool:
        entry   = self._cache_lookup(x, y)
        wrapper = entry.wrapper if entry else self._get_wrapper_at(x, y)
        if not wrapper:
            return False
        proc = self._get_process_name(wrapper)
        cls  = self._safe(wrapper, "class_name", critical=False) or ""
        return (proc and proc.lower() in BROWSER_PROCS) or cls in _ELECTRON_CLASS

    def is_excel_window(self, x: int, y: int) -> bool:
        entry   = self._cache_lookup(x, y)
        wrapper = entry.wrapper if entry else self._get_wrapper_at(x, y)
        if not wrapper:
            return False
        proc = self._get_process_name(wrapper)
        return proc.lower() in EXCEL_PROCS if proc else False

    # ── DPI ───────────────────────────────────────────────────────────────

    @staticmethod
    def _get_dpi() -> float:
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return round(dpi / 96.0, 2)
        except Exception:
            return 1.0

    # ── Wrapper retrieval (multi-layer) ───────────────────────────────────

    def _get_wrapper_at(self, x: int, y: int):
  
        if not PYWINAUTO_OK:
            return None

        # Layer 1: region-based cache
        entry = self._cache_lookup(x, y)
        if entry:
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
            logger.debug("UIA from_point({},{}) timed out — layer 3 focused fallback", x, y)
            fb = self._focused_element_fallback()
            return fb if fb else self._last_wrapper

        if exc_holder[0]:
            logger.debug("UIA from_point({},{}) error: {} — layer 3", x, y, exc_holder[0])
            fb = self._focused_element_fallback()
            if fb:
                return fb
            return self._last_wrapper

        wrapper = result[0]
        if wrapper:
            self._cache_put(x, y, wrapper)
            self._last_wrapper = wrapper   # update last-known

        return wrapper

    def _focused_element_fallback(self):
        """Layer 3: Return currently focused element wrapper."""
        try:
            desktop = self._get_desktop()
            focused = desktop.get_active()
            if focused:
                wrapper = (focused.wrapper_object()
                           if hasattr(focused, "wrapper_object") else focused)
                if self._is_visible(wrapper):
                    logger.debug("UIA-5: using focused-element fallback")
                    return wrapper
        except Exception:
            pass
        return None

    # ── Cache helpers ─────────────────────────────────────────────────────

    def _cache_lookup(self, x: int, y: int) -> Optional[_WrapperCacheEntry]:
        with self._cache_lock:
            dead = []
            for key, entry in list(self._wrapper_cache.items()):
                if not entry.is_valid():
                    dead.append(key)
                    continue
                if entry.near(x, y):
                    self._wrapper_cache.move_to_end(key)
                    for k in dead:
                        self._wrapper_cache.pop(k, None)
                    return entry
            for k in dead:
                self._wrapper_cache.pop(k, None)
        return None

    def _cache_put(self, x: int, y: int, wrapper) -> None:
        with self._cache_lock:
            self._cache_counter += 1
            key = self._cache_counter
            if len(self._wrapper_cache) >= _CACHE_MAX_ENTRIES:
                self._wrapper_cache.popitem(last=False)
            self._wrapper_cache[key] = _WrapperCacheEntry(x, y, wrapper)

    def invalidate_cache(self) -> None:
        """Force full cache clear (call on window focus change)."""
        with self._cache_lock:
            self._wrapper_cache.clear()

    # ── Target / selector builders ────────────────────────────────────────

    def _build_uia_target(self, wrapper, x: int, y: int) -> UITarget:
        auto_id    = self._safe(wrapper, "automation_id",       critical=True)
        name       = self._safe(wrapper, "window_text",         critical=True)
        ctrl_type  = self._safe(wrapper, "friendly_class_name", critical=False)
        class_name = self._safe(wrapper, "class_name",          critical=False)
        win_title  = self._get_window_title(wrapper)
        proc_name  = self._get_process_name(wrapper)

        raw_bbox  = self._get_raw_bbox(wrapper)
        norm_bbox = self._get_normalised_bbox(wrapper, x, y)
        ancestors = self._get_rich_ancestors(wrapper)
        caps      = self._get_capabilities(wrapper)
        hwnd      = self._get_window_handle(wrapper)
        is_active = self._is_foreground_window(hwnd)
        rel_window = self._get_relative_to_window(wrapper, raw_bbox)
        rel_parent = self._get_relative_to_parent(wrapper, raw_bbox)
        elem_hash  = _element_hash(auto_id, name, ctrl_type, class_name)
        role       = _classify_role(ctrl_type, caps)
        confidence_level = _confidence_level(auto_id, name, ctrl_type, norm_bbox)
        vis_score  = self._visibility_score(wrapper)

        # Backend routing
        backend = TargetBackend.UIA
        if proc_name and proc_name.lower() in BROWSER_PROCS:
            backend = TargetBackend.BROWSER
        elif class_name in _ELECTRON_CLASS:
            backend = TargetBackend.BROWSER

        dpi = _get_dpi_for_point(x, y)

        # FIX: Filter unstable automation_ids at capture time
        from .selector import _is_unstable as _sel_unstable
        stable_aid = (auto_id or None) if not _sel_unstable(auto_id or "") else None

        return UITarget(
            backend            = backend,
            automation_id      = stable_aid,
            name               = (name or "")[:200] or None,
            control_type       = ctrl_type or None,
            class_name         = class_name or None,
            window_title       = win_title or None,
            process_name       = proc_name or None,
            window_handle      = hwnd,
            is_active_window   = is_active,
            bbox               = norm_bbox,
            raw_bbox           = raw_bbox,
            screen_x           = x,
            screen_y           = y,
            dpi_scale          = dpi,
            ancestor_chain     = ancestors,
            element_hash       = elem_hash,
            element_role       = role,
            confidence_level   = confidence_level.value,
            visibility_score   = vis_score.value,
            relative_to_window = rel_window,
            relative_to_parent = rel_parent,
            is_editable        = caps.get("is_editable", False),
        )

    def _build_selector(self, wrapper, x: int, y: int) -> Selector:
        auto_id    = self._safe(wrapper, "automation_id",       critical=True)
        name       = self._safe(wrapper, "window_text",         critical=True)
        ctrl_type  = self._safe(wrapper, "friendly_class_name", critical=False)
        class_name = self._safe(wrapper, "class_name",          critical=False)
        win_title  = self._get_window_title(wrapper)
        proc_name  = self._get_process_name(wrapper)
        norm_bbox  = self._get_normalised_bbox(wrapper, x, y)
        ancestors  = self._get_rich_ancestors(wrapper)
        sibling_idx = self._get_sibling_index(wrapper)
        caps       = self._get_capabilities(wrapper)
        anchors    = self._get_anchor_elements(wrapper)
        dpi        = _get_dpi_for_point(x, y)
        hwnd       = self._get_window_handle(wrapper)
        is_active  = self._is_foreground_window(hwnd)
        raw_bbox   = self._get_raw_bbox(wrapper)
        rel_window = self._get_relative_to_window(wrapper, raw_bbox)
        rel_parent = self._get_relative_to_parent(wrapper, raw_bbox)
        elem_hash  = _element_hash(auto_id, name, ctrl_type, class_name)
        role       = _classify_role(ctrl_type, caps)
        conf_level = _confidence_level(auto_id, name, ctrl_type, norm_bbox)
        vis_score  = self._visibility_score(wrapper)

        return SelectorBuilder.from_uia(
            automation_id       = auto_id,
            name                = name,
            control_type        = ctrl_type,
            class_name          = class_name,
            window_title        = win_title,
            process_name        = proc_name,
            bbox                = norm_bbox,
            ancestor_chain      = ancestors,
            sibling_index       = sibling_idx,
            screen_x            = x,
            screen_y            = y,
            dpi_scale           = dpi,
            capabilities        = caps,
            anchor_elements     = anchors,
            element_hash        = elem_hash,
            element_role        = role,
            confidence_level    = conf_level.value,
            visibility_score    = vis_score.value,
            window_handle       = hwnd,
            raw_bbox            = raw_bbox,
            relative_to_window  = rel_window,
            relative_to_parent  = rel_parent,
        )

    # ── BBox helpers ──────────────────────────────────────────────────────

    def _get_raw_bbox(self, wrapper) -> Optional[BoundingBox]:
        try:
            r = wrapper.rectangle()
            return BoundingBox(left=r.left, top=r.top, right=r.right, bottom=r.bottom)
        except Exception:
            return None

    def _get_normalised_bbox(self, wrapper, elem_x: int = 0, elem_y: int = 0) -> Optional[BoundingBox]:
        try:
            r   = wrapper.rectangle()
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

    # ── Relative position ─────────────────────────────────────────────────

    def _get_relative_to_window(self, wrapper, raw_bbox: Optional[BoundingBox]) -> Optional[dict]:
        if not raw_bbox:
            return None
        try:
            win = wrapper.top_level_parent()
            wr  = win.rectangle()
            return {
                "x":     raw_bbox.left - wr.left,
                "y":     raw_bbox.top  - wr.top,
                "w":     raw_bbox.right  - raw_bbox.left,
                "h":     raw_bbox.bottom - raw_bbox.top,
                "win_w": wr.right  - wr.left,
                "win_h": wr.bottom - wr.top,
            }
        except Exception:
            return None

    def _get_relative_to_parent(self, wrapper, raw_bbox: Optional[BoundingBox]) -> Optional[dict]:
        if not raw_bbox:
            return None
        try:
            parent = wrapper.parent()
            if not parent:
                return None
            pr = parent.rectangle()
            return {
                "x": raw_bbox.left - pr.left,
                "y": raw_bbox.top  - pr.top,
                "w": raw_bbox.right  - raw_bbox.left,
                "h": raw_bbox.bottom - raw_bbox.top,
            }
        except Exception:
            return None

    # ── Ancestor chain ────────────────────────────────────────────────────

    def _get_rich_ancestors(self, wrapper, max_depth: int = 8) -> list[str]:
        """
        Adaptive depth — stops at Window/Dialog/Pane boundaries.
        Never exceeds max_depth.
        """
        chain = []
        try:
            current = wrapper
            for _ in range(max_depth):
                parent = current.parent()
                if not parent:
                    break
                ptype = self._safe(parent, "friendly_class_name", critical=False) or "?"
                ptext = (self._safe(parent, "window_text", critical=False) or "")[:30]
                paid  = self._safe(parent, "automation_id", critical=False) or ""
                entry = f"{ptype}:{ptext}"
                if paid and not _is_unstable_id(paid):
                    entry += f":{paid}"
                chain.append(entry)
                if ptype in _ANCESTOR_STOP_TYPES:
                    break
                current = parent
        except Exception:
            pass
        return chain

    # ── Anchor elements ───────────────────────────────────────────────────

    def _get_anchor_elements(self, wrapper, radius_px: int = 150) -> list[AnchorElement]:
        """
        Collects nearby stable sibling/parent elements for disambiguation.
        Up to 4 anchors, limited to label-like or within radius.
        """
        anchors: list[AnchorElement] = []
        try:
            raw_rect = wrapper.rectangle()
            cx_self  = (raw_rect.left + raw_rect.right)  / 2
            cy_self  = (raw_rect.top  + raw_rect.bottom) / 2
        except Exception:
            return anchors

        try:
            parent = wrapper.parent()
            if parent:
                try:
                    for sibling in parent.children():
                        if len(anchors) >= 4:
                            break
                        try:
                            if sibling == wrapper:
                                continue
                            sib_type = self._safe(sibling, "friendly_class_name", critical=False) or ""
                            sib_name = self._safe(sibling, "window_text", critical=False) or ""
                            if not sib_name or len(sib_name) < 2 or _is_unstable_id(sib_name):
                                continue
                            r    = sibling.rectangle()
                            cx_s = (r.left + r.right)  / 2
                            cy_s = (r.top  + r.bottom) / 2
                            dx   = int(cx_s - cx_self)
                            dy   = int(cy_s - cy_self)
                            dist = (dx**2 + dy**2) ** 0.5
                            is_label  = sib_type.lower() in ("text", "label", "statictext",
                                                               "textblock", "static")
                            in_radius = dist <= radius_px
                            if not (is_label or in_radius):
                                continue
                            direction = ("left"  if dx < -10 else
                                         "right" if dx >  10 else
                                         "above" if dy <   0 else "below")
                            anchors.append(AnchorElement(
                                direction    = direction,
                                name         = sib_name[:60],
                                control_type = sib_type,
                                offset_x     = dx,
                                offset_y     = dy,
                            ))
                        except Exception:
                            continue
                except Exception:
                    pass

                # Parent label/group name
                if len(anchors) < 4:
                    try:
                        p_name = self._safe(parent, "window_text", critical=False) or ""
                        p_type = self._safe(parent, "friendly_class_name", critical=False) or ""
                        if p_name and len(p_name) >= 2 and p_type.lower() in (
                            "group", "groupbox", "label", "text", "statictext",
                            "pane", "tabitem", "header",
                        ):
                            anchors.append(AnchorElement(
                                direction    = "parent",
                                name         = p_name[:60],
                                control_type = p_type,
                                offset_x     = 0,
                                offset_y     = 0,
                            ))
                    except Exception:
                        pass
        except Exception:
            pass
        return anchors[:4]

    def _get_sibling_index(self, wrapper) -> Optional[int]:
        try:
            parent    = wrapper.parent()
            if not parent:
                return None
            ctrl_type = self._safe(wrapper, "friendly_class_name", critical=False)
            idx       = 0
            for child in parent.children():
                if child == wrapper:
                    return idx
                if self._safe(child, "friendly_class_name", critical=False) == ctrl_type:
                    idx += 1
        except Exception:
            pass
        return None

    # ── Capabilities ──────────────────────────────────────────────────────

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
            caps["is_editable"]   = any("value" in p or "text" in p for p in pnames)
            caps["is_toggleable"] = any("toggle" in p for p in pnames)
            caps["is_selectable"] = any("select" in p for p in pnames)
            caps["is_invokable"]  = any("invoke" in p for p in pnames)
        except Exception:
            pass
        try:
            caps["is_editable"] = caps["is_editable"] or (
                wrapper.is_editable() if hasattr(wrapper, "is_editable") else False
            )
        except Exception:
            pass
        return caps

    # ── Visibility ────────────────────────────────────────────────────────

    def _visibility_score(self, wrapper) -> VisibilityScore:
        try:
            vis = getattr(wrapper, "is_visible", None)
            if callable(vis) and not vis():
                return VisibilityScore.HIDDEN

            r = wrapper.rectangle()
            if (r.right - r.left <= 0 or r.bottom - r.top <= 0 or
                    (r.left == 0 and r.top == 0 and r.right == 0 and r.bottom == 0)):
                return VisibilityScore.HIDDEN

            if r.right < 0 or r.bottom < 0 or r.left > 16000 or r.top > 16000:
                return VisibilityScore.OFFSCREEN

            if r.left < -2 or r.top < -2:
                return VisibilityScore.PARTIAL

            return VisibilityScore.FULL
        except Exception:
            return VisibilityScore.FULL

    def _is_visible(self, wrapper) -> bool:
        score = self._visibility_score(wrapper)
        return score in (VisibilityScore.FULL, VisibilityScore.PARTIAL)

    # ── Window helpers ────────────────────────────────────────────────────

    def _get_window_handle(self, wrapper) -> int:
        try:
            return wrapper.top_level_parent().handle
        except Exception:
            return 0

    def _is_foreground_window(self, hwnd: int) -> bool:
        if not hwnd:
            return False
        try:
            return ctypes.windll.user32.GetForegroundWindow() == hwnd
        except Exception:
            return False

    def _same_context(self, wrapper, target: UITarget) -> bool:
        try:
            win_title = wrapper.top_level_parent().window_text() or ""
            if target.window_title and target.window_title.lower() not in win_title.lower():
                return False
            if target.process_name:
                pid   = wrapper.process_id()
                pname = _cached_process_name(pid)
                if target.process_name.lower() not in (pname or "").lower():
                    return False
        except Exception:
            pass
        return True

    def _get_desktop(self):
        with self._desktop_lock:
            if self._desktop_cache is None:
                self._desktop_cache = Desktop(backend="uia")
            return self._desktop_cache

    def _get_window_title(self, wrapper) -> Optional[str]:
        try:
            return wrapper.top_level_parent().window_text() or None
        except Exception:
            return None

    def _get_process_name(self, wrapper) -> Optional[str]:
        try:
            if PSUTIL_OK:
                pid = wrapper.process_id()
                return _cached_process_name(pid)
        except Exception:
            pass
        return None

    def _safe(self, wrapper, attr: str, critical: bool = False) -> Optional[str]:
        try:
            v   = getattr(wrapper, attr)
            val = v() if callable(v) else v
            return str(val) if val is not None else None
        except Exception as exc:
            if critical:
                logger.debug("[ENRICHER] _safe('{}') failed: {}", attr, exc)
            return None

def _cached_process_name(pid: int) -> Optional[str]:
    """Cache psutil lookups to avoid repeated OS calls."""
    with _PID_NAME_LOCK:
        if pid in _PID_NAME_CACHE:
            return _PID_NAME_CACHE[pid]
    if not PSUTIL_OK:
        return None
    try:
        name = psutil.Process(pid).name()
        with _PID_NAME_LOCK:
            if len(_PID_NAME_CACHE) >= _PID_CACHE_MAX:
                oldest = next(iter(_PID_NAME_CACHE))
                del _PID_NAME_CACHE[oldest]
            _PID_NAME_CACHE[pid] = name
        return name
    except Exception:
        return None


def _element_hash(auto_id: Optional[str], name: Optional[str],
                  ctrl_type: Optional[str], class_name: Optional[str]) -> str:
    """Stable 16-char hex hash for element identity tracking."""
    sig = f"{auto_id or ''}|{name or ''}|{ctrl_type or ''}|{class_name or ''}"
    return hashlib.md5(sig.encode()).hexdigest()[:16]


def _is_unstable_id(value: str) -> bool:
    """Proxy to selector._is_unstable for use within enricher."""
    from .selector import _is_unstable
    return _is_unstable(value)


# Keep old alias for backward compatibility
_is_unstable = _is_unstable_id


def _classify_role(ctrl_type: Optional[str], caps: dict) -> str:
    if not ctrl_type:
        return "unknown"
    ct = ctrl_type.lower()
    if "button" in ct or "splitbutton" in ct:       return "button"
    if ct in ("edit", "document", "richtext", "textbox"): return "input"
    if caps.get("is_editable"):                       return "input"
    if "checkbox" in ct:                              return "checkbox"
    if "radio" in ct:                                 return "radio"
    if "combobox" in ct:                              return "dropdown"
    if ct in ("text", "label", "statictext", "textblock"): return "label"
    if "list" in ct:                                  return "list"
    if "tree" in ct:                                  return "tree"
    if "menu" in ct:                                  return "menu"
    if "tab" in ct:                                   return "tab"
    if ct in ("dataitem", "cell", "spreadsheetitem"): return "cell"
    if ct in ("pane", "group", "toolbar", "statusbar"): return "container"
    return "unknown"


def _confidence_level(auto_id: Optional[str], name: Optional[str],
                      ctrl_type: Optional[str], bbox) -> ConfidenceLevel:
    if auto_id and not _is_unstable_id(auto_id):
        return ConfidenceLevel.HIGH
    if name and ctrl_type:
        return ConfidenceLevel.MEDIUM
    return ConfidenceLevel.LOW


def detect_excel_cell(name: Optional[str], ctrl_type: Optional[str]) -> Optional[str]:
    if not name:
        return None

    clean = name.strip()

    # Strip sheet prefix: "Sheet1!B4" → "B4"
    if "!" in clean:
        clean = clean.split("!", 1)[1].strip()

    clean_up = clean.upper()

    if not _CELL_RE.match(clean_up):
        return None

    # Accept from any of these control types
    if ctrl_type in ("DataItem", "Cell", "SpreadsheetItem", "Edit", "Custom"):
        return clean_up

    return None