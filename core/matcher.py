from __future__ import annotations

import concurrent.futures
import copy
import ctypes
import re
import threading
import time
import hashlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from utils.logger import logger
from utils.errors import ElementNotFoundError
from models.target import UITarget, TargetBackend

try:
    from pywinauto import Application
    from pywinauto.findwindows import find_windows
    try:
        from pywinauto import Desktop
    except ImportError:
        from pywinauto import Desktop
    UIA_OK = True
except ImportError:
    UIA_OK = False

try:
    import cv2, numpy as np
    CV2_OK = True
except ImportError:
    CV2_OK = False

try:
    import mss
    MSS_OK = True
except ImportError:
    MSS_OK = False

_ELECTRON_CLASSES  = {"Chrome_WidgetWin_1", "CefBrowserWindow"}
_ELECTRON_PROCS    = {"teams.exe", "slack.exe", "notion.exe", "code.exe",
                      "discord.exe", "figma.exe", "obsidian.exe"}
_BROWSER_PROCS     = {"chrome.exe", "msedge.exe", "firefox.exe", "brave.exe"}
_SYSTEM_PROCS      = {"explorer.exe", "searchhost.exe", "searchapp.exe",
                      "shellexperiencehost.exe", "startmenuexperiencehost.exe"}
_EXCEL_PROCS       = {"excel.exe"}

_ELEMENT_TYPE_PRIORITY: dict[str, int] = {
    "Button":          90,
    "SplitButton":     88,
    "Edit":            85,
    "Document":        85,
    "ComboBox":        80,
    "CheckBox":        80,
    "RadioButton":     80,
    "SpreadsheetItem": 78,
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

_ROLE_SCORE_BOOST: dict[str, float] = {
    "button":    2.0,
    "input":     3.0,
    "checkbox":  2.0,
    "radio":     2.0,
    "dropdown":  2.0,
    "label":    -5.0,
    "container": -8.0,
    "unknown":   0.0,
}

MAX_WINDOW_SCAN      = 25
MAX_DESC_SEARCH      = 50
MAX_STATS            = 300
STRATEGY_TIMEOUT_S   = 2.0    # Default timeout
EXCEL_STRATEGY_TIMEOUT_S = 0.8  # FIX: Reduced timeout for Excel (was 2.0, now 0.8)
STABILITY_WAIT_MS    = 80
STABILITY_CHECK_ENABLED = True
MAX_REFRESH_ATTEMPTS = 2

# FIX: Excel cell cache
_EXCEL_CELL_CACHE: dict[str, tuple[object, float]] = {}
_EXCEL_CELL_CACHE_TTL = 3.0  # 3 seconds
_EXCEL_CELL_CACHE_LOCK = threading.Lock()


# ── DPI ───────────────────────────────────────────────────────────────────

_DPI_PRIMARY_CACHE: Optional[float] = None
_DPI_LOCK = threading.Lock()


def _primary_dpi() -> float:
    global _DPI_PRIMARY_CACHE
    with _DPI_LOCK:
        if _DPI_PRIMARY_CACHE is None:
            try:
                hdc = ctypes.windll.user32.GetDC(0)
                dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
                ctypes.windll.user32.ReleaseDC(0, hdc)
                _DPI_PRIMARY_CACHE = dpi / 96.0
            except Exception:
                _DPI_PRIMARY_CACHE = 1.0
        return _DPI_PRIMARY_CACHE


def _dpi_for_point(x: int, y: int) -> float:
    try:
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        hmon  = ctypes.windll.user32.MonitorFromPoint(POINT(x, y), 2)
        dpi_x = ctypes.c_uint()
        dpi_y = ctypes.c_uint()
        if ctypes.windll.shcore.GetDpiForMonitor(
                hmon, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y)) == 0:
            return dpi_x.value / 96.0
    except Exception:
        pass
    return _primary_dpi()


def _is_excel_target(target: UITarget) -> bool:
    """Helper to check if target is Excel."""
    if not target:
        return False
    proc = (target.process_name or "").lower()
    return proc in _EXCEL_PROCS


def _get_cached_excel_cell(cell_ref: str, hwnd: int) -> Optional[object]:
    """Get cached Excel cell wrapper."""
    cache_key = f"{hwnd}:{cell_ref}"
    with _EXCEL_CELL_CACHE_LOCK:
        entry = _EXCEL_CELL_CACHE.get(cache_key)
        if entry:
            wrapper, ts = entry
            if time.monotonic() - ts < _EXCEL_CELL_CACHE_TTL:
                return wrapper
            else:
                del _EXCEL_CELL_CACHE[cache_key]
    return None


def _cache_excel_cell(cell_ref: str, hwnd: int, wrapper: object) -> None:
    """Cache Excel cell wrapper."""
    cache_key = f"{hwnd}:{cell_ref}"
    with _EXCEL_CELL_CACHE_LOCK:
        # Keep cache reasonable size
        if len(_EXCEL_CELL_CACHE) > 200:
            oldest = min(_EXCEL_CELL_CACHE.items(), key=lambda x: x[1][1])[0]
            del _EXCEL_CELL_CACHE[oldest]
        _EXCEL_CELL_CACHE[cache_key] = (wrapper, time.monotonic())


# ── Data classes ──────────────────────────────────────────────────────────

@dataclass
class MatchResult:
    element:   object
    score:     float
    strategy:  str
    is_unique: bool = True
    elem_hash: str  = ""

    @property
    def is_wrapper(self) -> bool:
        return not isinstance(self.element, tuple)


@dataclass
class StrategyStats:
    successes:      int   = 0
    failures:       int   = 0
    total_ms:       float = 0.0
    priority_boost: float = 0.0
    last_used_ts:   float = 0.0

    @property
    def success_rate(self) -> float:
        total = self.successes + self.failures
        return self.successes / total if total > 0 else 0.5

    @property
    def avg_ms(self) -> float:
        total = self.successes + self.failures
        return self.total_ms / total if total > 0 else 999.0

    def record_success(self, elapsed_ms: float = 0.0) -> None:
        self.successes     += 1
        self.total_ms      += elapsed_ms
        self.priority_boost = min(self.priority_boost + 5.0, 25.0)
        self.last_used_ts   = time.monotonic()

    def record_failure(self, elapsed_ms: float = 0.0) -> None:
        self.failures      += 1
        self.total_ms      += elapsed_ms
        self.priority_boost = max(self.priority_boost - 3.0, -10.0)
        self.last_used_ts   = time.monotonic()


# ── Scoring helpers ───────────────────────────────────────────────────────

def _text_similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    def norm(s: str) -> str:
        return re.sub(r"[^\w\s]", " ", s.lower()).strip()
    a, b = norm(a), norm(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.85
    ta, tb = set(a.split()), set(b.split())
    if ta and tb:
        overlap = len(ta & tb) / len(ta | tb)
        if overlap >= 0.6:
            return overlap
    def bigrams(s: str):
        return set(s[i:i+2] for i in range(len(s)-1))
    ba, bb = bigrams(a), bigrams(b)
    return len(ba & bb) / len(ba | bb) if ba and bb else 0.0


def _spatial_score(elem, orig_bbox, weight: float = 1.0) -> float:
    if not orig_bbox:
        return 0.0
    try:
        rect = elem.rectangle()
        cx   = (rect.left   + rect.right)  / 2
        cy   = (rect.top    + rect.bottom) / 2
        ocx  = (orig_bbox.left + orig_bbox.right)  / 2
        ocy  = (orig_bbox.top  + orig_bbox.bottom) / 2
        dist = ((cx - ocx)**2 + (cy - ocy)**2) ** 0.5
        return max(0.0, 20.0 * weight - dist / 10.0)
    except Exception:
        return 0.0


def _composite_score(text_sim, spatial, strategy_rate, boost,
                     sel_conf=0.5, rb=0.0, tp=0.0) -> float:
    return (text_sim * 95 * 0.45 + (spatial / 20) * 95 * 0.20 +
            strategy_rate * 95 * 0.15 + sel_conf * 95 * 0.15 +
            tp * 95 * 0.05 + boost + rb)


def _tp(ctrl: Optional[str]) -> float:
    return _ELEMENT_TYPE_PRIORITY.get(ctrl, 50) / 100.0 if ctrl else 0.5


def _role_boost(ctrl_type: Optional[str], role: Optional[str]) -> float:
    if role:
        return _ROLE_SCORE_BOOST.get(role, 0.0)
    if not ctrl_type:
        return 0.0
    ct = ctrl_type.lower()
    if "button" in ct:                              return _ROLE_SCORE_BOOST["button"]
    if "edit" in ct:                                return _ROLE_SCORE_BOOST["input"]
    if "document" in ct:                            return _ROLE_SCORE_BOOST["input"]
    if "label" in ct or "text" == ct or "statictext" in ct: return _ROLE_SCORE_BOOST["label"]
    if "pane" in ct or "group" in ct:              return _ROLE_SCORE_BOOST["container"]
    return 0.0


# ── Element validation ────────────────────────────────────────────────────

def _validate_wrapper(wrapper) -> bool:
    if wrapper is None:
        return False
    if not callable(getattr(wrapper, "rectangle", None)):
        return False
    try:
        wrapper.rectangle()
        return True
    except Exception:
        return False


def _elem_visible(wrapper) -> bool:
    vis = getattr(wrapper, "visibility_score", None)
    if vis is not None:
        return vis not in ("hidden", "offscreen")
    try:
        if not wrapper.is_visible():
            return False
        r = wrapper.rectangle()
        if r.right - r.left <= 0 or r.bottom - r.top <= 0:
            return False
        if r.right < -50 or r.bottom < -50:
            return False
        if r.left > 16000 or r.top > 16000:
            return False
        if r.left == 0 and r.top == 0 and r.right == 0 and r.bottom == 0:
            return False
        return True
    except Exception:
        return True


def _elem_enabled(wrapper) -> bool:
    try:
        return wrapper.is_enabled()
    except Exception:
        return True


def _elem_hash(w) -> str:
    try:
        aid  = ""
        name = ""
        try:
            aid = w.automation_id() if callable(getattr(w, "automation_id", None)) else ""
        except Exception:
            pass
        try:
            name = w.window_text() or ""
        except Exception:
            pass
        try:
            ct = w.friendly_class_name() if callable(getattr(w, "friendly_class_name", None)) else ""
        except Exception:
            ct = ""
        key = f"{aid}|{name}|{ct}"
        return hashlib.md5(key.encode()).hexdigest()[:12]
    except Exception:
        return str(id(w))


def _refresh_wrapper(wrapper, app_cache: dict, hwnd: int,
                     auto_id: Optional[str] = None, name: Optional[str] = None):
    if _validate_wrapper(wrapper):
        return wrapper
    try:
        app = app_cache.get(hwnd) or Application(backend="uia").connect(handle=hwnd)
        app_cache[hwnd] = app
        win = app.window(handle=hwnd)
        if auto_id:
            elem = win.child_window(auto_id=auto_id)
            if elem.exists(timeout=0.5):
                fresh = elem.wrapper_object()
                if _validate_wrapper(fresh):
                    return fresh
        if name:
            elem = win.child_window(title_re=f".*{re.escape(name[:30])}.*")
            if elem.exists(timeout=0.5):
                fresh = elem.wrapper_object()
                if _validate_wrapper(fresh):
                    return fresh
    except Exception:
        pass
    return wrapper


def _title_variants(window_title: str) -> list[str]:
    if not window_title:
        return []
    variants = [window_title]
    parts = [p.strip() for p in window_title.split(" - ")]
    for p in parts:
        if p and p not in variants:
            variants.append(p)
    if len(window_title) > 30:
        prefix = window_title[:30]
        if prefix not in variants:
            variants.append(prefix)
    return variants


def _find_window_handles(window_title: str, process_name: Optional[str] = None,
                         max_results: int = 10) -> list[int]:
    handles = []
    for variant in _title_variants(window_title):
        try:
            found = find_windows(title_re=f".*{re.escape(variant[:40])}.*")
            for h in found:
                if h not in handles:
                    handles.append(h)
            if handles:
                break
        except Exception:
            continue
    return handles[:max_results]


def _active_window_bonus(target: UITarget) -> float:
    is_active = getattr(target, "is_active_window", None)
    if is_active is True:
        return 5.0
    if is_active is False:
        return -10.0
    return 0.0


_UNSAFE_CONTAINER_TYPES = {"Pane", "Group", "ToolBar", "StatusBar", "ScrollBar", "Window", "TitleBar", "MenuBar"}
_PASSIVE_TEXT_TYPES = {"Text", "Label", "StaticText"}


def _wrapper_text(wrapper) -> str:
    try:
        return wrapper.window_text() or ""
    except Exception:
        return ""


def _wrapper_ctrl_type(wrapper) -> str:
    try:
        return (wrapper.friendly_class_name()
                if callable(getattr(wrapper, "friendly_class_name", None)) else "") or ""
    except Exception:
        return ""


def _target_anchor_names(target: UITarget) -> list[str]:
    anchors = list(getattr(target, "anchor_elements", None) or [])
    for rich in getattr(target, "rich_selectors", None) or []:
        anchors.extend(getattr(rich, "anchor_elements", None) or [])

    names: list[str] = []
    for anchor in anchors:
        name = getattr(anchor, "name", None)
        if name is None and isinstance(anchor, dict):
            name = anchor.get("name")
        if name and name not in names:
            names.append(str(name))
    return names


# ── App-cache staleness check ─────────────────────────────────────────────

def _app_cache_valid(app_cache: dict, hwnd: int) -> bool:
    app = app_cache.get(hwnd)
    if app is None:
        return False
    try:
        app.window(handle=hwnd).rectangle()
        return True
    except Exception:
        return False


# ── ElementMatcher ────────────────────────────────────────────────────────

class ElementMatcher:

    STRICT_THRESHOLD          = 85.0
    MIN_ACCEPT_SCORE         = 65.0
    RELAXED_THRESHOLD        = 65.0
    SELF_HEAL_ACCEPT_SCORE   = 65.0
    SPATIAL_ACCEPT_SCORE     = 55.0
    BROWSER_MIN_ACCEPT_SCORE = 70.0
    AMBIGUITY_MARGIN         = 10.0
    REJECT_ON_AMBIGUOUS      = True

    def __init__(
        self,
        screenshot_base_dir: Optional[Path] = None,
        browser = None,
    ):
        self._scr_dir  = screenshot_base_dir
        self._browser  = browser
        self._stats:      dict[str, StrategyStats] = {}
        self._stats_lock  = threading.RLock()
        self._app_cache:  dict[int, object] = {}
        self._app_cache_lock = threading.Lock()
        self._last_found_excel_cell: Optional[tuple[str, int, object]] = None  # FIX: cache last found Excel cell

    # ── Stats helpers ─────────────────────────────────────────────────────

    def _get_stats(self, strategy: str) -> StrategyStats:
        with self._stats_lock:
            if strategy not in self._stats:
                if len(self._stats) >= MAX_STATS:
                    oldest = min(self._stats, key=lambda k: self._stats[k].last_used_ts)
                    del self._stats[oldest]
                self._stats[strategy] = StrategyStats()
            s = self._stats[strategy]
            s.last_used_ts = time.monotonic()
            return s

    def _get_stats_snapshot(self, strategy: str) -> tuple[float, float]:
        with self._stats_lock:
            s = self._stats.get(strategy)
            if s is None:
                return (0.5, 999.0)
            return (s.success_rate, s.avg_ms)

    def _ordered_strategies(self, candidates: list[str], target: Optional[UITarget] = None) -> list[str]:
        """FIX: Excel gets reduced strategy set"""
        if target and _is_excel_target(target):
            # Excel: only use automation_id and semantic (fastest)
            excel_candidates = [c for c in candidates if c in ("automation_id", "semantic")]
            if excel_candidates:
                return excel_candidates + ["relaxed"] if "relaxed" in candidates else excel_candidates
        
        def sort_key(name: str):
            rate, ms = self._get_stats_snapshot(name)
            return (-rate, ms)
        return sorted(candidates, key=sort_key)

    # ── Selector helpers ──────────────────────────────────────────────────

    def _sel_conf(self, target: UITarget) -> float:
        rich = getattr(target, "rich_selectors", None)
        if rich:
            try:
                val = max((s.effective_confidence() for s in rich
                           if hasattr(s, "effective_confidence")), default=None)
                if val is not None:
                    return val
            except Exception:
                pass
        sels = getattr(target, "selectors", None)
        if not sels:
            return 0.5
        try:
            return max((s.effective_confidence() for s in sels
                        if hasattr(s, "effective_confidence")), default=0.5)
        except Exception:
            return 0.5

    def _notify_selector(self, target: UITarget, strategy: str, success: bool):
        rich = getattr(target, "rich_selectors", None)
        sels = rich if rich else getattr(target, "selectors", None)
        if not sels:
            return
        mapping = {
            "automation_id": "strict", "automation_id_wide": "strict",
            "semantic":      "semantic", "semantic_desc":     "semantic",
            "relaxed":       "relaxed",  "relaxed_desc":      "relaxed",
            "classname":     "classname", "ancestor":          "ancestor",
            "coord":         "positional", "bbox":             "positional",
        }
        sel_name = mapping.get(strategy, strategy)
        try:
            for sel in sels:
                if hasattr(sel, "record_replay"):
                    sel.record_replay(sel_name, success)
        except Exception:
            pass

    def _selector_order(self, target: UITarget) -> list[str]:
        rich = getattr(target, "rich_selectors", None)
        sels = rich if rich else getattr(target, "selectors", None)
        if not sels:
            return []
        try:
            best = max(
                (s for s in sels if hasattr(s, "ordered_strategies")),
                key=lambda s: (s.effective_confidence()
                               if hasattr(s, "effective_confidence") else 0),
                default=None,
            )
            return [st.name for st in best.ordered_strategies()] if best else []
        except Exception:
            return []

    def _combined_order(self, names: list[str], pref: list[str]) -> list[str]:
        pref_boost = {n: (len(pref) - i) * 0.1 for i, n in enumerate(pref)}
        def key(n):
            rate, ms = self._get_stats_snapshot(n)
            return (-(rate + pref_boost.get(n, 0.0)), ms)
        return sorted(names, key=key)

    # ── App cache helpers ─────────────────────────────────────────────────

    def _get_app(self, hwnd: int):
        with self._app_cache_lock:
            if hwnd in self._app_cache:
                if _app_cache_valid(self._app_cache, hwnd):
                    return self._app_cache[hwnd]
                else:
                    del self._app_cache[hwnd]
        app = Application(backend="uia").connect(handle=hwnd)
        with self._app_cache_lock:
            self._app_cache[hwnd] = app
        return app

    # ── Public API ────────────────────────────────────────────────────────

    def find(self, target: UITarget, event_id: int = 0,
             action_intent: Optional[str] = None):
        t = copy.copy(target)
        if t.backend == TargetBackend.BROWSER:
            return self._find_browser(t, event_id)
        return self._find_uia(t, event_id, action_intent)

    def find_with_wait(
        self,
        target:        UITarget,
        event_id:      int = 0,
        timeout_ms:    int = 10_000,
        poll_ms:       int = 300,
        action_intent: Optional[str] = None,
    ):
        deadline = time.monotonic() + timeout_ms / 1000
        last_exc = None
        attempt  = 0
        while time.monotonic() < deadline:
            attempt += 1
            try:
                result = self.find(target, event_id, action_intent)
                if not isinstance(result, tuple):
                    if _elem_visible(result):
                        if attempt > 1:
                            logger.debug("[MATCH] Event #{} found after {} attempts",
                                         event_id, attempt)
                        return result
                    else:
                        logger.debug("[MATCH] Event #{} found but not visible (attempt {})",
                                     event_id, attempt)
                else:
                    return result
            except ElementNotFoundError as exc:
                last_exc = exc
            time.sleep(poll_ms / 1000)
        raise ElementNotFoundError(
            f"Event #{event_id}: not found within {timeout_ms}ms. Last: {last_exc}", event_id
        )

    def record_outcome(self, strategy: str, success: bool, elapsed_ms: float = 0.0) -> None:
        s = self._get_stats(strategy)
        if success:
            s.record_success(elapsed_ms)
            logger.debug("[MATCH] '{}' ✓ rate={:.0%} boost={:.1f}",
                         strategy, s.success_rate, s.priority_boost)
        else:
            s.record_failure(elapsed_ms)
            logger.debug("[MATCH] '{}' ✗ rate={:.0%} boost={:.1f}",
                         strategy, s.success_rate, s.priority_boost)

    # ── Browser ───────────────────────────────────────────────────────────

    def _find_browser(self, target: UITarget, event_id: int) -> tuple[int, int]:
        if not self._browser or not self._browser.is_connected:
            raise ElementNotFoundError(f"Event #{event_id}: browser not connected", event_id)

        bt = target.browser
        if not bt:
            raise ElementNotFoundError(f"Event #{event_id}: no browser target data", event_id)

        logger.info(
            "[MATCH] Event #{} BROWSER search → xpath={} css={} aria={} text='{}'",
            event_id, bt.xpath, bt.css_selector, bt.aria_label,
            (bt.inner_text or "")[:30],
        )

        self._browser.wait_for_dom_stable(stable_ms=250, max_wait_ms=2000)
        candidates = self._browser.find_candidates(bt, timeout_ms=8000)
        if not candidates:
            raise ElementNotFoundError(
                f"Event #{event_id}: browser element not found "
                f"xpath={bt.xpath!r} css={bt.css_selector!r} text={bt.inner_text!r}",
                event_id,
            )

        candidates = [
            c for c in candidates
            if c.visible and c.score >= self.BROWSER_MIN_ACCEPT_SCORE
        ]
        if not candidates:
            raise ElementNotFoundError(
                f"Event #{event_id}: browser candidates below confidence threshold "
                f"({self.BROWSER_MIN_ACCEPT_SCORE:.0f})",
                event_id,
            )

        best = candidates[0]
        if len(candidates) >= 2:
            gap = candidates[0].score - candidates[1].score
            if gap < 10:
                logger.warning(
                    "[MATCH] Event #{} AMBIGUOUS browser candidates ({:.0f} vs {:.0f})",
                    event_id, candidates[0].score, candidates[1].score,
                )
                priority = {"xpath": 0, "css": 1, "aria": 2, "exact_text": 3, "partial_text": 4}
                candidates.sort(key=lambda c: (priority.get(c.strategy, 9), -c.score))
                best = candidates[0]

        # If element is invisible or at (0,0) — likely an off-screen/async-loading element.
        # Wait up to 3s for it to become visible (e.g. Gmail compose body loads in an iframe).
        if not best.visible or (best.cx == 0 and best.cy == 0):
            logger.info(
                "[MATCH] Event #{} browser element not yet visible (pos=({},{}) visible={}) — waiting",
                event_id, best.cx, best.cy, best.visible,
            )
            deadline = time.time() + 3.0
            while time.time() < deadline:
                time.sleep(0.3)
                retry_c = self._browser.find_candidates(bt, timeout_ms=2000)
                if retry_c:
                    rc = retry_c[0]
                    if rc.visible and not (rc.cx == 0 and rc.cy == 0):
                        best = rc
                        logger.info(
                            "[MATCH] Event #{} browser element now visible pos=({},{}) strategy={}",
                            event_id, best.cx, best.cy, best.strategy,
                        )
                        break
            else:
                logger.warning(
                    "[MATCH] Event #{} browser element still not visible after wait — "
                    "using recorded screen coords as fallback",
                    event_id,
                )
                # Return the recorded coords from the target (screen coords stored at record time)
                # Let _do_click fall back to raw coords rather than clicking at (-9,230)
                raise ElementNotFoundError(
                    f"Event #{event_id}: browser element found but not visible/reachable",
                    event_id,
                )

        verification = self._browser.verify_element(bt)
        if not verification.get("found"):
            logger.warning("[MATCH] Event #{} browser element stale — retrying query", event_id)
            # One more attempt after DOM settle before giving up
            self._browser.wait_for_dom_stable(stable_ms=300, max_wait_ms=2000)
            retry_candidates = self._browser.find_candidates(bt, timeout_ms=4000)
            if retry_candidates:
                best = retry_candidates[0]
                logger.info("[MATCH] Event #{} browser stale retry succeeded strategy={} pos=({},{})",
                            event_id, best.strategy, best.cx, best.cy)
            else:
                logger.warning("[MATCH] Event #{} browser stale retry failed — using original coords", event_id)

        logger.info(
            "[MATCH] Event #{} BROWSER → strategy={} score={:.0f} pos=({},{}) visible={} text='{}'",
            event_id, best.strategy, best.score, best.cx, best.cy, best.visible,
            best.text[:20],
        )
        self._get_stats(f"browser_{best.strategy}").record_success()
        return (best.cx, best.cy)

    # ── UIA matching — two-phase ──────────────────────────────────────────

    def _find_uia(self, target: UITarget, event_id: int,
                  action_intent: Optional[str] = None):
        if not UIA_OK:
            return self._coord_fallback(target, event_id)

        proc        = (target.process_name or "").lower()
        cls         = (target.class_name   or "")
        is_electron = proc in _ELECTRON_PROCS or cls in _ELECTRON_CLASSES
        is_system   = proc in _SYSTEM_PROCS
        is_excel    = proc in _EXCEL_PROCS

        if is_electron:
            logger.info("[MATCH] Event #{} ELECTRON → image/coord", event_id)
            coords = self._by_screenshot_cropped(target)
            if coords:
                return coords
            return self._coord_fallback(target, event_id)

        confidence_level = getattr(target, "confidence_level", "medium")
        active_bonus = _active_window_bonus(target)
        sel_conf     = self._sel_conf(target)
        pref_order   = self._selector_order(target)

        strict_th = self.STRICT_THRESHOLD + (5 if confidence_level == "high" else 0)

        logger.info(
            "[MATCH] Event #{} UIA → auto_id={} name='{}' ctrl={} win='{}' "
            "intent={} sel_conf={:.2f}",
            event_id, target.automation_id or "(none)",
            (target.name or "")[:30], target.control_type or "?",
            (target.window_title or "")[:30],
            action_intent or "any", sel_conf,
        )

        results: list[MatchResult] = []
        errors:  list[str]         = []

        # FIX: Excel gets reduced strategy set
        active_strategies = self._ordered_strategies(
            ["automation_id", "semantic", "relaxed", "classname", "ancestor"],
            target
        )

        # FIX: Excel gets reduced timeout
        strategy_timeout = EXCEL_STRATEGY_TIMEOUT_S if is_excel else STRATEGY_TIMEOUT_S

        # FIX: Check cache for Excel cell first
        if is_excel and target.automation_id and target.window_handle:
            cached_wrapper = _get_cached_excel_cell(target.automation_id, target.window_handle)
            if cached_wrapper and _validate_wrapper(cached_wrapper) and _elem_visible(cached_wrapper):
                logger.info("[MATCH] Event #{} EXCEL CELL CACHE HIT: {}", event_id, target.automation_id)
                return self._wrap_safe(cached_wrapper, target)

        def run(name: str, fn):
            """Execute fn with per-strategy timeout guard."""
            t0 = time.monotonic()
            result_holder = [None]
            exc_holder    = [None]
            done          = threading.Event()
            early_result  = [None]

            def _execute():
                try:
                    result_holder[0] = fn()
                except Exception as exc:
                    exc_holder[0] = exc
                finally:
                    done.set()

            th = threading.Thread(target=_execute, daemon=True)
            th.start()
            timed_out = not done.wait(timeout=strategy_timeout)
            ms = (time.monotonic() - t0) * 1000

            if timed_out:
                logger.warning("[MATCH] Event #{} strategy '{}' TIMED OUT ({:.0f}ms) from Excel={}",
                               event_id, name, ms, is_excel)
                with self._stats_lock:
                    self._get_stats(name).record_failure(ms)
                errors.append(f"{name}: timeout")
                return []

            if exc_holder[0]:
                with self._stats_lock:
                    self._get_stats(name).record_failure(ms)
                errors.append(f"{name}: {exc_holder[0]}")
                logger.debug("[MATCH] Event #{} {} error: {}", event_id, name, exc_holder[0])
                return []

            r = result_holder[0]
            if r:
                rs = r if isinstance(r, list) else [r]
                with self._stats_lock:
                    for _ in rs:
                        self._get_stats(name).record_success(ms)
                return rs
            return []

        # ── Phase 1: Fast strategies ───────────────────────────────────────

        if target.automation_id:
            for r in run("automation_id", lambda: self._by_automation_id(target)):
                r.score += active_bonus
                results.append(r)
                # FIX: Cache Excel cell result
                if is_excel and target.automation_id and target.window_handle and r.is_wrapper:
                    _cache_excel_cell(target.automation_id, target.window_handle, r.element)
                logger.info("[MATCH] Event #{} auto_id='{}' → score={:.0f}",
                            event_id, target.automation_id, r.score)

        if target.name and target.control_type:
            for r in run("semantic", lambda: self._by_name_type(target, exact=True)):
                r.score += active_bonus
                results.append(r)
                logger.info("[MATCH] Event #{} semantic '{}' ctrl={} → score={:.0f}",
                            event_id, (target.name or "")[:30], target.control_type, r.score)

        # Deduplicate before picking. Strict strategies must pass full acceptance gates.
        best_strict = self._select_accepted(results, target, event_id, strict_th, "STRICT")
        if best_strict:
            return self._wrap_safe(best_strict.element, target)

        # Phase 2: true fallback chain. Each tier gets a chance to pass its own gate
        # before the next, weaker strategy is allowed to run.
        if target.name and "relaxed" in active_strategies:
            tier_results: list[MatchResult] = []
            for r in run("relaxed", lambda: self._by_name_type(target, exact=False)):
                r.score += active_bonus
                tier_results.append(r)
                logger.info("[MATCH] Event #{} relaxed '{}' -> score={:.0f}",
                            event_id, (target.name or "")[:30], r.score)
            best = self._select_accepted(tier_results, target, event_id,
                                         self.RELAXED_THRESHOLD, "RELAXED")
            if best:
                return self._wrap_safe(best.element, target)

        if target.class_name and target.window_title and "classname" in active_strategies:
            tier_results = []
            for r in run("classname", lambda: self._by_classname(target)):
                r.score += active_bonus
                tier_results.append(r)
            best = self._select_accepted(tier_results, target, event_id,
                                         self.MIN_ACCEPT_SCORE, "CLASSNAME")
            if best:
                return self._wrap_safe(best.element, target)

        if target.ancestor_chain and "ancestor" in active_strategies:
            tier_results = []
            for r in run("ancestor", lambda: self._by_ancestor(target)):
                r.score += active_bonus
                tier_results.append(r)
            best = self._select_accepted(tier_results, target, event_id,
                                         self.MIN_ACCEPT_SCORE, "ANCESTOR")
            if best:
                return self._wrap_safe(best.element, target)

        # Self-heal
        # FIX: For Excel, use shorter self-heal timeout
        healed = self._self_heal(target, event_id, active_bonus, is_excel=is_excel)
        if healed is not None:
            return healed

        # ── BBox ───────────────────────────────────────────────────────────
        if target.bbox:
            r = self._by_bbox(target)
            if r:
                best = self._select_accepted([r], target, event_id,
                                             self.SPATIAL_ACCEPT_SCORE, "BBOX")
                if best:
                    logger.info("[MATCH] Event #{} BBOX fallback accepted", event_id)
                    return self._wrap_safe(best.element, target)

        # ── Screenshot ─────────────────────────────────────────────────────
        if getattr(target, "screenshot_ref", None) and self._scr_dir:
            coords = self._by_screenshot_cropped(target)
            if coords:
                logger.info("[MATCH] Event #{} SCREENSHOT fallback pos={}", event_id, coords)
                return coords

        # ── Coordinate fallback ────────────────────────────────────────────
        if target.screen_x is not None:
            logger.warning(
                "[MATCH] Event #{} RAW COORD fallback pos=({},{}) — tried: {}",
                event_id, target.screen_x, target.screen_y, " | ".join(errors[:4]),
            )
            return (target.screen_x, target.screen_y)

        logger.error("[MATCH] Event #{} TOTAL FAILURE: {}", event_id, " | ".join(errors))
        raise ElementNotFoundError(
            f"Event #{event_id}: all strategies failed", event_id
        )

    # ── Deduplication ─────────────────────────────────────────────────────

    def _deduplicate(self, results: list[MatchResult]) -> list[MatchResult]:
        seen: dict[str, MatchResult] = {}
        for r in results:
            key = str(r.element) if not r.is_wrapper else (r.elem_hash or _elem_hash(r.element))
            if key not in seen or r.score > seen[key].score:
                seen[key] = r
        return list(seen.values())

    # ── Intent filter ─────────────────────────────────────────────────────

    def _intent_filter(self, results: list[MatchResult],
                       intent: str) -> list[MatchResult]:
        if intent not in ("type", "click", "select"):
            return results
        out = []
        for r in results:
            score = r.score
            if r.is_wrapper:
                try:
                    raw = r.element.raw() if hasattr(r.element, "raw") else r.element
                    ct  = (raw.friendly_class_name()
                           if callable(getattr(raw, "friendly_class_name", None)) else "") or ""
                    ct = ct.lower()
                    if intent == "type":
                        if any(x in ct for x in ("edit", "document", "richtext")): score += 10
                        elif any(x in ct for x in ("button", "label", "pane")):    score -= 15
                    elif intent == "click":
                        if any(x in ct for x in ("button", "menuitem", "tabitem")): score += 5
                        elif any(x in ct for x in ("label", "statictext")):         score -= 8
                    elif intent == "select":
                        if any(x in ct for x in ("combobox", "listitem", "list")):  score += 10
                except Exception:
                    pass
            out.append(MatchResult(r.element, score, r.strategy, r.is_unique, r.elem_hash))
        return out


    def _filter_candidates(
        self,
        results: list[MatchResult],
        target: UITarget,
        event_id: int,
        phase: str,
    ) -> list[MatchResult]:
        filtered: list[MatchResult] = []
        for r in results:
            if not r.is_wrapper:
                filtered.append(r)
                continue

            raw = r.element.raw() if hasattr(r.element, "raw") else r.element
            if not _validate_wrapper(raw):
                logger.warning("[MATCH] Event #{} {} reject {}: invalid wrapper", event_id, phase, r.strategy)
                continue
            if not _elem_visible(raw):
                logger.warning("[MATCH] Event #{} {} reject {}: not visible", event_id, phase, r.strategy)
                continue

            score = r.score
            ctrl = _wrapper_ctrl_type(raw)
            text = _wrapper_text(raw)

            if not _elem_enabled(raw):
                score -= 35.0
            if ctrl in _UNSAFE_CONTAINER_TYPES:
                score -= 30.0
            elif ctrl in _PASSIVE_TEXT_TYPES and getattr(target, "element_role", None) not in ("label", "text"):
                score -= 15.0

            target_ctrl = target.control_type or ""
            if target_ctrl and ctrl and target_ctrl.lower() != ctrl.lower():
                score -= 12.0

            target_name = target.name or ""
            if target_name and r.strategy not in ("automation_id", "automation_id_wide", "heal_hash"):
                sim = _text_similarity(target_name, text)
                if sim < 0.25:
                    score -= 25.0
                elif sim < 0.50:
                    score -= 10.0

            if score != r.score:
                logger.debug(
                    "[MATCH] Event #{} {} {} adjusted {:.0f}->{:.0f} ctrl={} text='{}'",
                    event_id, phase, r.strategy, r.score, score, ctrl or "?", text[:30],
                )

            filtered.append(MatchResult(r.element, max(0.0, score), r.strategy, r.is_unique, r.elem_hash))
        return filtered

    def _select_accepted(
        self,
        results: list[MatchResult],
        target: UITarget,
        event_id: int,
        threshold: float,
        phase: str,
    ) -> Optional[MatchResult]:
        if not results:
            return None
        candidates = self._filter_candidates(results, target, event_id, phase)
        candidates = self._deduplicate(candidates)
        candidates = self._context_gate(candidates, target)
        best = self._pick_best(candidates, threshold)
        if not best:
            if candidates:
                top = max(candidates, key=lambda r: r.score)
                logger.warning(
                    "[MATCH] Event #{} {} low confidence top={:.0f} strategy={} threshold={:.0f}",
                    event_id, phase, top.score, top.strategy, threshold,
                )
            return None

        if best.score < threshold:
            logger.warning(
                "[MATCH] Event #{} {} rejected low score {:.0f} < {:.0f} strategy={}",
                event_id, phase, best.score, threshold, best.strategy,
            )
            return None

        if best.is_wrapper:
            raw = best.element.raw() if hasattr(best.element, "raw") else best.element
            if not _validate_wrapper(raw) or not _elem_visible(raw):
                logger.warning("[MATCH] Event #{} {} rejected invalid/invisible wrapper", event_id, phase)
                return None
            if not _elem_enabled(raw):
                logger.warning("[MATCH] Event #{} {} rejected disabled element", event_id, phase)
                return None
            if not self._cross_validate(best, target):
                logger.warning("[MATCH] Event #{} {} rejected by cross-validation", event_id, phase)
                return None
            if not self._stability_check(best.element):
                logger.warning("[MATCH] Event #{} {} rejected unstable element", event_id, phase)
                return None

        logger.info(
            "[MATCH] Event #{} {} accepted strategy={} score={:.0f} unique={}",
            event_id, phase, best.strategy, best.score, best.is_unique,
        )
        return best

    def _wrap_safe(self, element, target: UITarget):
        if isinstance(element, tuple):
            return element
        try:
            from .uia_enricher import SafeElement
            hwnd = getattr(target, "window_handle", 0) or 0
            return SafeElement(
                wrapper  = element,
                hwnd     = hwnd,
                auto_id  = target.automation_id or "",
                name     = target.name or "",
            )
        except ImportError:
            return element

    # ── Strategy implementations ──────────────────────────────────────────

    def _by_automation_id(self, t: UITarget, sc: float = 0.5) -> list[MatchResult]:
        boost = self._get_stats("automation_id").priority_boost
        found: list[MatchResult] = []

        for hwnd in _find_window_handles(t.window_title or "", t.process_name):
            try:
                app  = self._get_app(hwnd)
                elem = app.window(handle=hwnd).child_window(auto_id=t.automation_id)
                if elem.exists(timeout=0.3):  # FIX: reduced timeout from 0.5 to 0.3
                    w = elem.wrapper_object()
                    if _validate_wrapper(w) and _elem_visible(w):
                        sp  = _spatial_score(w, t.bbox)
                        rb  = _role_boost(t.control_type, getattr(t, "element_role", None))
                        tp  = _tp(t.control_type)
                        sc_ = _composite_score(1.0, sp,
                              self._get_stats("automation_id").success_rate,
                              boost, sc, rb, tp)
                        found.append(MatchResult(w, sc_, "automation_id",
                                                  is_unique=True, elem_hash=_elem_hash(w)))
            except Exception:
                continue

        if not found:
            try:
                for hwnd in find_windows(title_re=".*")[:MAX_WINDOW_SCAN]:
                    try:
                        app = self._get_app(hwnd)
                        win = app.window(handle=hwnd)
                        if t.process_name:
                            try:
                                import psutil
                                pname = psutil.Process(win.process_id()).name().lower()
                                if t.process_name.lower() not in pname:
                                    continue
                            except Exception:
                                pass
                        descs = win.descendants(auto_id=t.automation_id)
                        for d in descs[:5]:
                            w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                            if not _validate_wrapper(w) or not _elem_visible(w):
                                continue
                            if t.window_title:
                                try:
                                    wt = w.top_level_parent().window_text() or ""
                                    if not any(v.lower() in wt.lower()
                                               for v in _title_variants(t.window_title)):
                                        continue
                                except Exception:
                                    pass
                            sp  = _spatial_score(w, t.bbox)
                            rb  = _role_boost(t.control_type, getattr(t, "element_role", None))
                            tp  = _tp(t.control_type)
                            uniq = len(descs) == 1
                            sc_ = _composite_score(1.0, sp,
                                  self._get_stats("automation_id_wide").success_rate,
                                  boost - 10, sc, rb, tp)
                            found.append(MatchResult(w, sc_, "automation_id_wide",
                                                      is_unique=uniq, elem_hash=_elem_hash(w)))
                    except Exception:
                        continue
            except Exception:
                pass

        return found

    def _by_name_type(self, t: UITarget, exact: bool = True,
                      sc: float = 0.5) -> list[MatchResult]:
        results  = []
        strategy = "semantic" if exact else "relaxed"
        boost    = self._get_stats(strategy).priority_boost
        name     = t.name or ""
        conf_min = 0.85 if exact else 0.60
        name_re  = re.escape(name)

        for hwnd in _find_window_handles(t.window_title or "", t.process_name):
            try:
                app = self._get_app(hwnd)
                win = app.window(handle=hwnd)

                # child_window
                try:
                    kw   = {"control_type": t.control_type} if t.control_type else {}
                    elem = win.child_window(title_re=f".*{name_re}.*", **kw)
                    if elem.exists(timeout=0.3):  # FIX: reduced timeout
                        w = elem.wrapper_object()
                        if _validate_wrapper(w) and _elem_visible(w):
                            sim = _text_similarity(name, w.window_text() or "")
                            if sim >= conf_min:
                                sp  = _spatial_score(w, t.bbox)
                                rb  = _role_boost(t.control_type, getattr(t, "element_role", None))
                                tp  = _tp(t.control_type)
                                sc_ = _composite_score(sim, sp,
                                      self._get_stats(strategy).success_rate, boost, sc, rb, tp)
                                results.append(MatchResult(w, sc_, strategy,
                                    is_unique=True, elem_hash=_elem_hash(w)))
                except Exception:
                    pass

                # Descendants - limit search for Excel
                max_search = 20 if _is_excel_target(t) else MAX_DESC_SEARCH
                try:
                    kw    = {"control_type": t.control_type} if t.control_type else {}
                    descs = win.descendants(**kw) if kw else win.descendants()
                    for d in descs[:max_search]:
                        try:
                            w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                            if not _validate_wrapper(w) or not _elem_visible(w):
                                continue
                            sim = _text_similarity(name, w.window_text() or "")
                            if sim < conf_min:
                                continue
                            sp  = _spatial_score(w, t.bbox)
                            rb  = _role_boost(t.control_type, getattr(t, "element_role", None))
                            tp  = _tp(t.control_type)
                            sc_ = _composite_score(sim, sp,
                                  self._get_stats(strategy).success_rate, boost, sc, rb, tp)
                            results.append(MatchResult(w, sc_, f"{strategy}_desc",
                                is_unique=False, elem_hash=_elem_hash(w)))
                        except Exception:
                            continue
                except Exception:
                    pass
            except Exception:
                continue

        return results[:10]

    def _by_classname(self, t: UITarget, sc: float = 0.5) -> list[MatchResult]:
        results = []
        boost   = self._get_stats("classname").priority_boost
        for hwnd in _find_window_handles(t.window_title or ""):
            try:
                app = self._get_app(hwnd)
                win = app.window(handle=hwnd)
                try:
                    elem = win.child_window(class_name=t.class_name)
                    if elem.exists(timeout=0.3):
                        w = elem.wrapper_object()
                        if _validate_wrapper(w) and _elem_visible(w):
                            sp = _spatial_score(w, t.bbox)
                            rb = _role_boost(t.control_type, getattr(t, "element_role", None))
                            results.append(MatchResult(w, 65 + boost + sp + rb, "classname",
                                                        elem_hash=_elem_hash(w)))
                except Exception:
                    pass
                try:
                    descs = win.descendants(class_name=t.class_name)
                    for d in descs[:10]:
                        w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                        if _validate_wrapper(w) and _elem_visible(w):
                            sp = _spatial_score(w, t.bbox)
                            rb = _role_boost(t.control_type, getattr(t, "element_role", None))
                            results.append(MatchResult(w, 60 + boost + sp + rb, "classname_desc",
                                is_unique=len(descs) == 1, elem_hash=_elem_hash(w)))
                except Exception:
                    pass
            except Exception:
                continue
        return results

    def _by_ancestor(self, t: UITarget, sc: float = 0.5) -> list[MatchResult]:
        if not t.ancestor_chain:
            return []
        parts    = t.ancestor_chain[0].split(":", 2)
        anc_text = parts[1].strip() if len(parts) > 1 else ""
        if not anc_text:
            return []

        results    = []
        boost      = self._get_stats("ancestor").priority_boost
        anch_names = _target_anchor_names(t)
        try:
            for hwnd in find_windows(title_re=f".*{re.escape(anc_text[:20])}.*"):
                try:
                    app = self._get_app(hwnd)
                    win = app.window(handle=hwnd)
                    kw  = {}
                    if t.control_type: kw["control_type"] = t.control_type
                    if t.class_name:   kw["class_name"]   = t.class_name
                    if not kw:
                        continue
                    descs = win.descendants(**kw)
                    for d in descs[:20]:
                        w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                        if not _validate_wrapper(w) or not _elem_visible(w):
                            continue
                        sp  = _spatial_score(w, t.bbox)
                        ab  = self._anchor_confirmation(w, anch_names)
                        rb  = _role_boost(t.control_type, getattr(t, "element_role", None))
                        results.append(MatchResult(w, 55 + boost + sp + rb + ab, "ancestor",
                                                    elem_hash=_elem_hash(w)))
                except Exception:
                    continue
        except Exception:
            pass
        return results

    def _anchor_confirmation(self, wrapper, anchor_names: list[str]) -> float:
        if not anchor_names:
            return 0.0
        bonus = 0.0
        try:
            parent = wrapper.parent()
            if not parent:
                return 0.0
            for sibling in parent.children():
                try:
                    sib_text = sibling.window_text() or ""
                    for anchor_name in anchor_names:
                        if anchor_name.lower() in sib_text.lower():
                            bonus += 5.0
                            break
                except Exception:
                    continue
        except Exception:
            pass
        return min(bonus, 15.0)

    def _by_bbox(self, t: UITarget) -> Optional[MatchResult]:
        try:
            dpi   = _primary_dpi() / (t.dpi_scale or 1.0)
            cx    = int(((t.bbox.left + t.bbox.right)  / 2) * dpi)
            cy    = int(((t.bbox.top  + t.bbox.bottom) / 2) * dpi)
            boost = self._get_stats("bbox").priority_boost
            try:
                desktop = Desktop(backend="uia")
                elem    = desktop.from_point(cx, cy)
                if elem:
                    wrapper = elem.wrapper_object() if hasattr(elem, "wrapper_object") else elem
                    if _validate_wrapper(wrapper) and _elem_visible(wrapper):
                        return MatchResult(wrapper, 55.0 + boost, "bbox")
            except Exception:
                pass
        except Exception:
            pass
        return None

    def _by_screenshot_cropped(self, t: UITarget) -> Optional[tuple[int, int]]:
        if not CV2_OK or not MSS_OK:
            return None
        ref_path = None
        if getattr(t, "screenshot_ref", None) and self._scr_dir:
            ref_path = self._scr_dir / t.screenshot_ref
        if ref_path is None or not ref_path.exists():
            return None
        template = cv2.imread(str(ref_path), cv2.IMREAD_COLOR)
        if template is None:
            return None
        th, tw = template.shape[:2]
        try:
            with mss.mss() as sct:
                monitors     = sct.monitors[1:]
                capture_mon  = monitors[0]
                bbox = getattr(t, "raw_bbox", None) or t.bbox
                if bbox:
                    scale = _primary_dpi() / (t.dpi_scale or 1.0)
                    bcx   = int(((bbox.left + bbox.right)  / 2) * scale)
                    bcy   = int(((bbox.top  + bbox.bottom) / 2) * scale)
                    for mon in monitors:
                        if (mon["left"] <= bcx < mon["left"] + mon["width"] and
                                mon["top"] <= bcy < mon["top"] + mon["height"]):
                            capture_mon = mon
                            break
                if bbox:
                    scale  = _primary_dpi() / (t.dpi_scale or 1.0)
                    margin = 200
                    region = {
                        "left":   max(capture_mon["left"], int(bbox.left  * scale) - margin),
                        "top":    max(capture_mon["top"],  int(bbox.top   * scale) - margin),
                        "width":  tw + margin * 2 + int((bbox.right - bbox.left) * scale),
                        "height": th + margin * 2 + int((bbox.bottom - bbox.top) * scale),
                    }
                    region["width"]  = min(region["width"],
                        capture_mon["left"] + capture_mon["width"]  - region["left"])
                    region["height"] = min(region["height"],
                        capture_mon["top"]  + capture_mon["height"] - region["top"])
                else:
                    region = capture_mon
                if region["width"] <= tw or region["height"] <= th:
                    region = capture_mon
                shot = sct.grab(region)
                bgr  = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)
            res = cv2.matchTemplate(bgr, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            logger.debug("[MATCH] screenshot confidence={:.2%}", max_val)
            if max_val < 0.80:
                return None
            return (region["left"] + max_loc[0] + tw // 2,
                    region["top"]  + max_loc[1] + th // 2)
        except Exception as exc:
            logger.debug("[MATCH] screenshot error: {}", exc)
            return None

    # ── Cross-validation ──────────────────────────────────────────────────

    def _cross_validate(self, match: MatchResult, target: UITarget) -> bool:
        if not match.is_wrapper:
            return True
        try:
            elem = match.element
            if hasattr(elem, "raw"):
                elem = elem.raw()
            if not _validate_wrapper(elem):
                return False
            if target.name:
                found_text = elem.window_text() or ""
                sim  = _text_similarity(target.name, found_text)
            
                threshold = 0.50 if _is_excel_target(target) else 0.65
                ok  = sim >= threshold
                logger.debug("[MATCH] cross-validate name sim={:.2f} ok={}", sim, ok)
                return ok
            if target.control_type:
                try:
                    found_ctrl = (elem.friendly_class_name()
                                  if callable(getattr(elem, "friendly_class_name", None))
                                  else str(getattr(elem, "friendly_class_name", "")))
                    return target.control_type.lower() in found_ctrl.lower()
                except Exception:
                    return True
        except Exception:
            pass
        return True

    def _stability_check_fast(self, element) -> bool:
        if not STABILITY_CHECK_ENABLED:
            return True
        if isinstance(element, tuple):
            return True
        raw = element.raw() if hasattr(element, "raw") else element
        try:
            r1 = raw.rectangle()
            try:
                if raw.is_enabled():
                    time.sleep(STABILITY_WAIT_MS / 1000)
            except Exception:
                time.sleep(STABILITY_WAIT_MS / 1000)
            r2 = raw.rectangle()
            if abs(r1.left - r2.left) > 2 or abs(r1.top - r2.top) > 2:
                return False
            return _elem_visible(raw)
        except Exception:
            return False

    _stability_check = _stability_check_fast



    def _context_gate(self, results: list[MatchResult], target: UITarget) -> list[MatchResult]:
        if not target.window_title and not target.process_name:
            return results
        out = []
        for r in results:
            if not r.is_wrapper:
                out.append(r)
                continue
            bonus = 0.0
            try:
                raw  = r.element.raw() if hasattr(r.element, "raw") else r.element
                win_text = raw.top_level_parent().window_text() or ""
                if target.window_title:
                    if any(v.lower() in win_text.lower()
                           for v in _title_variants(target.window_title)):
                        bonus += 20.0
                    else:
                        bonus -= 25.0
            except Exception:
                pass
            out.append(MatchResult(r.element, r.score + bonus,
                                   r.strategy, r.is_unique, r.elem_hash))
        return out

    # ── Self-healing (parallel) ───────────────────────────────────────────

    def _self_heal(self, target: UITarget, event_id: int,
                   active_bonus: float = 0.0, is_excel: bool = False) -> Optional[object]:
        """
        Three-strategy self-healing, run in parallel for speed.
        FIX: Shorter timeout for Excel (2s instead of 4s)
        """
        logger.info("[MATCH] Event #{} SELF-HEAL starting", event_id)
        results: list[MatchResult] = []
        
        # FIX: Shorter timeout for Excel
        heal_timeout = STRATEGY_TIMEOUT_S if not is_excel else 2.0

        def _heal_a():
            out = []
            if not target.name:
                return out
            try:
                t_loose = copy.copy(target)
                t_loose.control_type = None
                for r in self._by_name_type(t_loose, exact=False):
                    out.append(MatchResult(
                        r.element, r.score * 0.75 + active_bonus,
                        "heal_relaxed", r.is_unique, r.elem_hash))
            except Exception:
                pass
            return out

        def _heal_b():
            out = []
            target_hash = getattr(target, "element_hash", None)
            if not target_hash or not UIA_OK:
                return out
            try:
                from .uia_enricher import _element_hash
                max_scan = 15 if is_excel else MAX_WINDOW_SCAN
                for hwnd in find_windows(title_re=".*")[:max_scan]:
                    try:
                        app = self._get_app(hwnd)
                        win = app.window(handle=hwnd)
                        kw  = {"control_type": target.control_type} if target.control_type else {}
                        max_desc = 20 if is_excel else MAX_DESC_SEARCH
                        descs = win.descendants(**kw) if kw else win.descendants()
                        for d in descs[:max_desc]:
                            try:
                                w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                                if not _validate_wrapper(w) or not _elem_visible(w):
                                    continue
                                aid, nm, ct, cn = None, None, None, None
                                try: aid = w.automation_id()
                                except: pass
                                try: nm  = w.window_text()
                                except: pass
                                try: ct  = w.friendly_class_name()
                                except: pass
                                try: cn  = w.class_name()
                                except: pass
                                if _element_hash(aid, nm, ct, cn) == target_hash:
                                    sp = _spatial_score(w, target.bbox)
                                    out.append(MatchResult(
                                        w, 65 + sp + active_bonus,
                                        "heal_hash", is_unique=True,
                                        elem_hash=target_hash))
                            except Exception:
                                continue
                    except Exception:
                        continue
            except Exception:
                pass
            return out

        def _heal_c():
            out = []
            anchor_names = _target_anchor_names(target)
            if not anchor_names or not target.window_title or not UIA_OK:
                return out
            try:
                for hwnd in _find_window_handles(target.window_title):
                    try:
                        app = self._get_app(hwnd)
                        win = app.window(handle=hwnd)
                        kw  = {"control_type": target.control_type} if target.control_type else {}
                        max_desc = 20 if is_excel else MAX_DESC_SEARCH
                        descs = win.descendants(**kw) if kw else win.descendants()
                        for d in descs[:max_desc]:
                            try:
                                w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                                if not _validate_wrapper(w) or not _elem_visible(w):
                                    continue
                                ab = self._anchor_confirmation(w, anchor_names)
                                if ab > 0:
                                    sp = _spatial_score(w, target.bbox)
                                    out.append(MatchResult(
                                        w, 45 + ab + sp + active_bonus,
                                        "heal_anchor", is_unique=False,
                                        elem_hash=_elem_hash(w)))
                            except Exception:
                                continue
                    except Exception:
                        continue
            except Exception:
                pass
            return out

        def _run_heal_tier(name: str, fn, threshold: float) -> Optional[object]:
            holder = [[]]
            done = threading.Event()

            def _execute():
                try:
                    holder[0] = fn()
                except Exception as exc:
                    logger.debug("[MATCH] Event #{} self-heal {} error: {}", event_id, name, exc)
                finally:
                    done.set()

            th = threading.Thread(target=_execute, daemon=True)
            th.start()
            if not done.wait(timeout=heal_timeout):
                logger.warning("[MATCH] Event #{} self-heal {} TIMED OUT", event_id, name)
                return None

            tier_results = self._context_gate(self._deduplicate(holder[0] or []), target)
            best = self._select_accepted(tier_results, target, event_id, threshold, f"SELF_HEAL:{name}")
            if best:
                logger.info("[MATCH] Event #{} SELF-HEAL accepted strategy={} score={:.0f}",
                            event_id, best.strategy, best.score)
                return self._wrap_safe(best.element, target)
            return None

        for name, fn, threshold in (
            ("hash", _heal_b, self.SELF_HEAL_ACCEPT_SCORE),
            ("anchor", _heal_c, self.SELF_HEAL_ACCEPT_SCORE),
            ("relaxed", _heal_a, self.RELAXED_THRESHOLD),
        ):
            healed = _run_heal_tier(name, fn, threshold)
            if healed is not None:
                return healed

        return None


    # ── Uniqueness check ──────────────────────────────────────────────────

    def _check_uniqueness(self, wrapper, target: UITarget, descs: list) -> bool:
        target_hash = getattr(target, "element_hash", None)
        if not target_hash:
            return len(descs) == 1
        try:
            from .uia_enricher import _element_hash
            aid = nm = ct = cn = None
            try: aid = wrapper.automation_id()
            except: pass
            try: nm  = wrapper.window_text()
            except: pass
            try: ct  = wrapper.friendly_class_name()
            except: pass
            try: cn  = wrapper.class_name()
            except: pass
            return _element_hash(aid, nm, ct, cn) == target_hash
        except Exception:
            pass
        return len(descs) == 1

    # ── Pick best result ──────────────────────────────────────────────────

    def _pick_best(self, results: list[MatchResult], threshold: float) -> Optional[MatchResult]:
        viable = [r for r in results if r.score >= threshold]
        if not viable:
            return None

        viable.sort(key=lambda r: r.score, reverse=True)

        if len(viable) >= 2:
            gap = viable[0].score - viable[1].score
            if gap < self.AMBIGUITY_MARGIN:
                logger.warning(
                    "[MATCH] AMBIGUOUS: {:.0f} vs {:.0f} (gap={:.0f}) {} vs {}",
                    viable[0].score, viable[1].score, gap,
                    viable[0].strategy, viable[1].strategy,
                )
                unique = [r for r in viable if r.is_unique]
                if len(unique) == 1:
                    logger.info("[MATCH] Ambiguity resolved by uniqueness")
                    return unique[0]

                # Resolve by element type priority
                if len(viable) >= 2:
                    elem0 = viable[0].element
                    elem1 = viable[1].element
                    try:
                        raw0 = elem0.raw() if hasattr(elem0, "raw") else elem0
                        raw1 = elem1.raw() if hasattr(elem1, "raw") else elem1
                        ct0  = (raw0.friendly_class_name()
                                if callable(getattr(raw0, "friendly_class_name", None)) else "") or ""
                        ct1  = (raw1.friendly_class_name()
                                if callable(getattr(raw1, "friendly_class_name", None)) else "") or ""
                        p0   = _ELEMENT_TYPE_PRIORITY.get(ct0, 50)
                        p1   = _ELEMENT_TYPE_PRIORITY.get(ct1, 50)
                        if abs(p0 - p1) >= 15:
                            chosen = viable[0] if p0 > p1 else viable[1]
                            logger.info("[MATCH] Ambiguity resolved by type priority ({} vs {})",
                                        ct0, ct1)
                            return chosen
                    except Exception:
                        pass

                if self.REJECT_ON_AMBIGUOUS and not unique:
                    logger.error("[MATCH] Ambiguity unresolved — REJECT_ON_AMBIGUOUS set")
                    return None

        unique_pool = [r for r in viable if r.is_unique]
        if unique_pool:
            best_unique  = unique_pool[0]
            best_overall = viable[0]
            if best_overall.is_unique:
                return best_overall
            if best_overall.score - best_unique.score > 20:
                return best_overall
            return best_unique

        return viable[0]

    # ── Coordinate fallback ───────────────────────────────────────────────

    def _coord_fallback(self, t: UITarget, event_id: int):
        
        if t.bbox:
            scale = _primary_dpi() / (t.dpi_scale or 1.0)
            cx    = int(((t.bbox.left + t.bbox.right)  / 2) * scale)
            cy    = int(((t.bbox.top  + t.bbox.bottom) / 2) * scale)
            # Adjust for per-monitor DPI
            actual = _dpi_for_point(cx, cy)
            if abs(actual - _primary_dpi()) > 0.05:
                cx = int(cx * actual / _primary_dpi())
                cy = int(cy * actual / _primary_dpi())
            return (cx, cy)
        if t.screen_x is not None:
            return (t.screen_x, t.screen_y)
        raise ElementNotFoundError(f"Event #{event_id}: no fallback coordinates", event_id)