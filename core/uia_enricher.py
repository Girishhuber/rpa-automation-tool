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


BROWSER_PROCS  = {"chrome.exe", "msedge.exe", "firefox.exe",
                   "brave.exe", "opera.exe", "vivaldi.exe"}
EXCEL_PROCS    = {"excel.exe"}
OFFICE_PROCS   = {"excel.exe", "winword.exe", "powerpnt.exe", "outlook.exe"}
_ELECTRON_CLASS = {"Chrome_WidgetWin_1", "CefBrowserWindow"}
_CELL_RE        = re.compile(r"^[A-Z]{1,3}[0-9]{1,7}$")

# Cache config
_CACHE_RADIUS_PX   = 5
_CACHE_MAX_ENTRIES = 32
_CACHE_TTL_S       = 0.5
_UIA_CALL_TIMEOUT  = 0.8

# PID cache (FIX-14)
_PID_NAME_CACHE: dict[int, str]     = {}
_PID_NAME_LOCK  = threading.Lock()
_PID_CACHE_MAX  = 200

# FIX-5: Element type priority (higher = preferred match when scores are close)
_ELEMENT_TYPE_PRIORITY: dict[str, int] = {
    "Button":      90,
    "SplitButton": 88,
    "Edit":        85,
    "Document":    85,
    "ComboBox":    80,
    "CheckBox":    80,
    "RadioButton": 80,
    "ListItem":    70,
    "TreeItem":    70,
    "TabItem":     65,
    "MenuItem":    65,
    "Text":        40,
    "Label":       40,
    "StaticText":  40,
    "Pane":        30,
    "Group":       25,
    "ToolBar":     20,
    "StatusBar":   15,
    "ScrollBar":   10,
}


_ANCESTOR_STOP_TYPES = {
    "Window", "Dialog", "Pane", "Frame",
    "CustomControl",   # usually top-level custom windows
}


class ConfidenceLevel(str, Enum):
    HIGH   = "high"    # automation_id or id-anchored xpath
    MEDIUM = "medium"  # name + control_type
    LOW    = "low"     # bbox, screenshot, coords


class VisibilityScore(str, Enum):
    FULL      = "full"       # fully on-screen and visible
    PARTIAL   = "partial"    # partially clipped by window edge
    OFFSCREEN = "offscreen"  # outside screen bounds
    HIDDEN    = "hidden"     # is_visible() returned False or zero-size rect



def _get_dpi_for_point(x: int, y: int) -> float:
    try:
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        MonitorFromPoint = ctypes.windll.user32.MonitorFromPoint
        GetDpiForMonitor = ctypes.windll.shcore.GetDpiForMonitor
        hmon  = MonitorFromPoint(POINT(x, y), 2)
        dpi_x = ctypes.c_uint()
        dpi_y = ctypes.c_uint()
        if GetDpiForMonitor(hmon, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y)) == 0:
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


class SafeElement:
 

    def __init__(self, wrapper, hwnd: int = 0, auto_id: str = "", name: str = ""):
        self._wrapper  = wrapper
        self._hwnd     = hwnd
        self._auto_id  = auto_id
        self._name     = name
        self._refresh_lock = threading.Lock()

    def _refresh(self) -> bool:
        """Try to re-fetch the element from its window."""
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
            except Exception as exc:
                with self._refresh_lock:
                    if self._refresh():
                        try:
                            return getattr(self._wrapper, name)(*args, **kwargs)
                        except Exception:
                            pass
                raise

        return _call

    def raw(self):
        """Return the underlying pywinauto wrapper."""
        return self._wrapper


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

    def is_valid(self) -> bool:
       
        if not self.is_fresh():
            return False
        try:
            r = self.wrapper.rectangle()
            return r is not None
        except Exception:
            return False



class UIAEnricher:

    def __init__(self):
        if not PYWINAUTO_OK:
            logger.warning("pywinauto not installed — UIA enrichment disabled")

        self._desktop_cache: Optional[Any] = None
        self._desktop_lock   = threading.Lock()

        # FIX-1/13: wrapper LRU cache with validity check
        self._wrapper_cache: OrderedDict[int, _WrapperCacheEntry] = OrderedDict()
        self._cache_lock     = threading.Lock()
        self._cache_counter  = 0

   
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
        FIX-2: Return UITarget for the focused element.
        If context_target provided, validates focused element is in same window/process.
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

            # FIX-2: validate same window/process if context provided
            if context_target:
                if not self._same_context(wrapper, context_target):
                    logger.debug("[ENRICHER] Focused element not in expected window/process — discarding")
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
        entry = self._cache_lookup(x, y)
        wrapper = entry.wrapper if entry else self._get_wrapper_at(x, y)
        if not wrapper:
            return False
        proc = self._get_process_name(wrapper)
        cls  = self._safe(wrapper, "class_name", critical=False) or ""
        return (proc and proc.lower() in BROWSER_PROCS) or cls in _ELECTRON_CLASS

    def is_excel_window(self, x: int, y: int) -> bool:
        entry = self._cache_lookup(x, y)
        wrapper = entry.wrapper if entry else self._get_wrapper_at(x, y)
        if not wrapper:
            return False
        proc = self._get_process_name(wrapper)
        return proc.lower() in EXCEL_PROCS if proc else False

    # ──────────────────────────────────────────────────────────────────
    # DPI
    # ──────────────────────────────────────────────────────────────────

    @staticmethod
    def _get_dpi() -> float:
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return round(dpi / 96.0, 2)
        except Exception:
            return 1.0

    # ──────────────────────────────────────────────────────────────────
    # Wrapper retrieval
    # ──────────────────────────────────────────────────────────────────

    def _get_wrapper_at(self, x: int, y: int):
        """UIA-1: cache-first; UIA-4: Event-based timeout; UIA-5: focused fallback."""
        if not PYWINAUTO_OK:
            return None

        # FIX-1/13: cache lookup with validity check
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
            logger.debug("UIA from_point({},{}) timed out — trying focused element", x, y)
            return self._focused_element_fallback()

        if exc_holder[0]:
            logger.debug("UIA from_point({},{}) error: {}", x, y, exc_holder[0])
            focused = self._focused_element_fallback()
            if focused:
                return focused

        wrapper = result[0]
        if wrapper:
            self._cache_put(x, y, wrapper)

        return wrapper

    def _focused_element_fallback(self):
        """UIA-5: Return the currently focused element wrapper."""
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
    # Cache helpers (FIX-1, FIX-13)
    # ──────────────────────────────────────────────────────────────────

    def _cache_lookup(self, x: int, y: int) -> Optional[_WrapperCacheEntry]:
        with self._cache_lock:
            dead = []
            for key, entry in list(self._wrapper_cache.items()):
                # FIX-1/13: validate by rectangle(), not just TTL
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

    # ──────────────────────────────────────────────────────────────────
    # Target / selector builders
    # ──────────────────────────────────────────────────────────────────

    def _build_uia_target(self, wrapper, x: int, y: int) -> UITarget:
        auto_id    = self._safe(wrapper, "automation_id",       critical=True)
        name       = self._safe(wrapper, "window_text",         critical=True)
        ctrl_type  = self._safe(wrapper, "friendly_class_name", critical=False)
        class_name = self._safe(wrapper, "class_name",          critical=False)
        win_title  = self._get_window_title(wrapper)
        proc_name  = self._get_process_name(wrapper)

        # FIX-8: both raw and normalised bbox
        raw_bbox  = self._get_raw_bbox(wrapper)
        norm_bbox = self._get_normalised_bbox(wrapper, x, y)

        # FIX-15: adaptive ancestor depth
        ancestors  = self._get_rich_ancestors(wrapper)
        caps       = self._get_capabilities(wrapper)

        # FIX-9: window handle + is_active
        hwnd      = self._get_window_handle(wrapper)
        is_active = self._is_foreground_window(hwnd)

        # FIX-11: relative positions
        rel_window = self._get_relative_to_window(wrapper, raw_bbox)
        rel_parent = self._get_relative_to_parent(wrapper, raw_bbox)

        # FIX-10: element hash
        elem_hash = _element_hash(auto_id, name, ctrl_type, class_name)

        # FIX-6: role classification
        role = _classify_role(ctrl_type, caps)

        # FIX-7: confidence level
        confidence_level = _confidence_level(auto_id, name, ctrl_type, norm_bbox)

        # FIX-17: visibility score
        vis_score = self._visibility_score(wrapper)

        # Backend routing
        backend = TargetBackend.UIA
        if proc_name and proc_name.lower() in BROWSER_PROCS:
            backend = TargetBackend.BROWSER
        elif class_name in _ELECTRON_CLASS:
            backend = TargetBackend.BROWSER

        dpi = _get_dpi_for_point(x, y)

        return UITarget(
            backend           = backend,
            automation_id     = (auto_id or None) if not _is_unstable(auto_id or "") else None,
            name              = (name or "")[:200] or None,
            control_type      = ctrl_type or None,
            class_name        = class_name or None,
            window_title      = win_title or None,
            process_name      = proc_name or None,
            window_handle     = hwnd,
            is_active_window  = is_active,
            bbox              = norm_bbox,
            raw_bbox          = raw_bbox,
            screen_x          = x,
            screen_y          = y,
            dpi_scale         = dpi,
            ancestor_chain    = ancestors,
            element_hash      = elem_hash,
            element_role      = role,
            confidence_level  = confidence_level.value,
            visibility_score  = vis_score.value,
            relative_to_window = rel_window,
            relative_to_parent = rel_parent,
            is_editable       = caps.get("is_editable", False),
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
        # FIX-4: extended anchors
        anchors    = self._get_anchor_elements(wrapper)
        dpi        = _get_dpi_for_point(x, y)

        # FIX-9/10/11: gather extended context for selector
        hwnd        = self._get_window_handle(wrapper)
        is_active   = self._is_foreground_window(hwnd)
        raw_bbox    = self._get_raw_bbox(wrapper)
        rel_window  = self._get_relative_to_window(wrapper, raw_bbox)
        rel_parent  = self._get_relative_to_parent(wrapper, raw_bbox)
        elem_hash   = _element_hash(auto_id, name, ctrl_type, class_name)
        role        = _classify_role(ctrl_type, caps)
        conf_level  = _confidence_level(auto_id, name, ctrl_type, norm_bbox)
        vis_score   = self._visibility_score(wrapper)

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

    # ──────────────────────────────────────────────────────────────────
    # BBox helpers (FIX-8)
    # ──────────────────────────────────────────────────────────────────

    def _get_raw_bbox(self, wrapper) -> Optional[BoundingBox]:
        """FIX-8: Absolute pixel bbox — no DPI normalisation."""
        try:
            r = wrapper.rectangle()
            return BoundingBox(left=r.left, top=r.top, right=r.right, bottom=r.bottom)
        except Exception:
            return None

    def _get_normalised_bbox(self, wrapper, elem_x: int = 0, elem_y: int = 0) -> Optional[BoundingBox]:
        """DPI-normalised bbox for cross-resolution matching."""
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

    # ──────────────────────────────────────────────────────────────────
    # Relative position (FIX-11)
    # ──────────────────────────────────────────────────────────────────

    def _get_relative_to_window(self, wrapper, raw_bbox: Optional[BoundingBox]) -> Optional[dict]:
        """FIX-11: Element position relative to its top-level window."""
        if not raw_bbox:
            return None
        try:
            win = wrapper.top_level_parent()
            wr  = win.rectangle()
            return {
                "x":    raw_bbox.left - wr.left,
                "y":    raw_bbox.top  - wr.top,
                "w":    raw_bbox.right  - raw_bbox.left,
                "h":    raw_bbox.bottom - raw_bbox.top,
                "win_w": wr.right  - wr.left,
                "win_h": wr.bottom - wr.top,
            }
        except Exception:
            return None

    def _get_relative_to_parent(self, wrapper, raw_bbox: Optional[BoundingBox]) -> Optional[dict]:
        """FIX-11: Element position relative to its direct parent."""
        if not raw_bbox:
            return None
        try:
            parent = wrapper.parent()
            if not parent:
                return None
            pr = parent.rectangle()
            return {
                "x":    raw_bbox.left - pr.left,
                "y":    raw_bbox.top  - pr.top,
                "w":    raw_bbox.right  - raw_bbox.left,
                "h":    raw_bbox.bottom - raw_bbox.top,
            }
        except Exception:
            return None

    # ──────────────────────────────────────────────────────────────────
    # Ancestor chain (FIX-15: adaptive depth)
    # ──────────────────────────────────────────────────────────────────

    def _get_rich_ancestors(self, wrapper, max_depth: int = 8) -> list[str]:
        """
        FIX-15: Adaptive depth — stops when we hit a Window/Dialog/Pane type,
        since those are the natural boundaries of element context. Never exceeds max_depth.
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
                if paid and not _is_unstable(paid):
                    entry += f":{paid}"
                chain.append(entry)

                # FIX-15: stop at meaningful boundary
                if ptype in _ANCESTOR_STOP_TYPES:
                    break

                current = parent
        except Exception:
            pass
        return chain

    # ──────────────────────────────────────────────────────────────────
    # Anchor elements (FIX-4: extended)
    # ──────────────────────────────────────────────────────────────────

    def _get_anchor_elements(self, wrapper, radius_px: int = 150) -> list[AnchorElement]:
        """
        FIX-4: Extended anchors — collects:
          1. Sibling label elements (original)
          2. Parent label / group name
          3. Nearby elements within radius_px
        Up to 4 anchors total.
        """
        anchors: list[AnchorElement] = []
        try:
            raw_rect = wrapper.rectangle()
            cx_self  = (raw_rect.left + raw_rect.right)  / 2
            cy_self  = (raw_rect.top  + raw_rect.bottom) / 2
        except Exception:
            return anchors

        # 1 & 3: Sibling-based + radius-based
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
                            if not sib_name or len(sib_name) < 2 or _is_unstable(sib_name):
                                continue

                            r    = sibling.rectangle()
                            cx_s = (r.left + r.right)  / 2
                            cy_s = (r.top  + r.bottom) / 2
                            dx   = int(cx_s - cx_self)
                            dy   = int(cy_s - cy_self)
                            dist = (dx**2 + dy**2) ** 0.5

                            # Only include label-like siblings OR anything within radius
                            is_label = sib_type.lower() in ("text", "label", "statictext",
                                                              "textblock", "static")
                            in_radius = dist <= radius_px

                            if not (is_label or in_radius):
                                continue

                            direction = "left"  if dx < -10 else \
                                        "right" if dx > 10  else \
                                        "above" if dy < 0   else "below"

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

                # 2: Parent label / group name (FIX-4)
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

    def _visibility_score(self, wrapper) -> VisibilityScore:
    
        try:
            vis = getattr(wrapper, "is_visible", None)
            if callable(vis) and not vis():
                return VisibilityScore.HIDDEN

            r = wrapper.rectangle()

            # Degenerate or zero-size
            if (r.right - r.left <= 0 or r.bottom - r.top <= 0 or
                    (r.left == 0 and r.top == 0 and r.right == 0 and r.bottom == 0)):
                return VisibilityScore.HIDDEN

            # Entirely offscreen
            if r.right < 0 or r.bottom < 0 or r.left > 16000 or r.top > 16000:
                return VisibilityScore.OFFSCREEN

            # Partially clipped
            if r.left < -2 or r.top < -2:
                return VisibilityScore.PARTIAL

            return VisibilityScore.FULL

        except Exception:
            return VisibilityScore.FULL  # assume visible on error

    def _is_visible(self, wrapper) -> bool:
        """Boolean shortcut — True unless HIDDEN or OFFSCREEN."""
        score = self._visibility_score(wrapper)
        return score in (VisibilityScore.FULL, VisibilityScore.PARTIAL)

  
    def _get_window_handle(self, wrapper) -> int:
        """FIX-9: Get HWND of top-level parent window."""
        try:
            return wrapper.top_level_parent().handle
        except Exception:
            return 0

    def _is_foreground_window(self, hwnd: int) -> bool:
        """FIX-9: True if this window is the current foreground window."""
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
    """FIX-14: Cache psutil lookups to avoid repeated OS calls."""
    with _PID_NAME_LOCK:
        if pid in _PID_NAME_CACHE:
            return _PID_NAME_CACHE[pid]

    if not PSUTIL_OK:
        return None
    try:
        name = psutil.Process(pid).name()
        with _PID_NAME_LOCK:
            if len(_PID_NAME_CACHE) >= _PID_CACHE_MAX:
                # Evict oldest (first inserted)
                oldest = next(iter(_PID_NAME_CACHE))
                del _PID_NAME_CACHE[oldest]
            _PID_NAME_CACHE[pid] = name
        return name
    except Exception:
        return None


def _element_hash(auto_id: Optional[str], name: Optional[str],
                   ctrl_type: Optional[str], class_name: Optional[str]) -> str:
    """
    FIX-10: Stable 16-char hex hash for element identity tracking.
    Same element across sessions produces same hash.
    """
    sig = f"{auto_id or ''}|{name or ''}|{ctrl_type or ''}|{class_name or ''}"
    return hashlib.md5(sig.encode()).hexdigest()[:16]


def _classify_role(ctrl_type: Optional[str], caps: dict) -> str:
  
    if not ctrl_type:
        return "unknown"
    ct = ctrl_type.lower()
    if "button" in ct or "splitbutton" in ct:
        return "button"
    if ct in ("edit", "document", "richtext", "textbox"):
        return "input"
    if caps.get("is_editable"):
        return "input"
    if "checkbox" in ct:
        return "checkbox"
    if "radio" in ct:
        return "radio"
    if "combobox" in ct:
        return "dropdown"
    if ct in ("text", "label", "statictext", "textblock"):
        return "label"
    if "list" in ct:
        return "list"
    if "tree" in ct:
        return "tree"
    if "menu" in ct:
        return "menu"
    if "tab" in ct:
        return "tab"
    if ct in ("dataitem", "cell", "spreadsheetitem"):
        return "cell"
    if ct in ("pane", "group", "toolbar", "statusbar"):
        return "container"
    return "unknown"


def _confidence_level(auto_id: Optional[str], name: Optional[str],
                       ctrl_type: Optional[str], bbox: Optional[BoundingBox]) -> ConfidenceLevel:
   
    if auto_id and not _is_unstable(auto_id):
        return ConfidenceLevel.HIGH
    if name and ctrl_type:
        return ConfidenceLevel.MEDIUM
    return ConfidenceLevel.LOW


def _is_unstable(value: str) -> bool:
    """Shared unstable-ID detector."""
    if not value:
        return True
    _KNOWN = {"1148", "1001", "1000", "100", "101", "1", "2", "3", "4"}
    if value in _KNOWN:
        return False
    if len(value) <= 3 and value.isalnum():
        return False
    patterns = [
        re.compile(r"^[0-9]{6,}$"),
        re.compile(r"\b[a-f0-9]{8,}\b"),
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
