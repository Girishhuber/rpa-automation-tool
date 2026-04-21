
from __future__ import annotations
import copy
import ctypes
import re
import threading
import time
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

MAX_WINDOW_SCAN    = 25     # max windows 
MAX_DESC_SEARCH    = 50     # max descendants to inspect per window
MAX_STATS          = 300    # LRU cap on _stats dict (BUG-3)
DESC_TIMEOUT_S     = 1.5    # per-window descendants timeout


STABILITY_WAIT_MS  = 80     



_DPI_PRIMARY_CACHE: Optional[float] = None
_DPI_LOCK = threading.Lock()


def _primary_dpi() -> float:
    """Cached primary-monitor DPI."""
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
        # MonitorFromPoint + GetDpiForMonitor (Windows 8.1+)
        MonitorFromPoint = ctypes.windll.user32.MonitorFromPoint
        GetDpiForMonitor = ctypes.windll.shcore.GetDpiForMonitor

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        hmon = MonitorFromPoint(POINT(x, y), 2)  # MONITOR_DEFAULTTONEAREST
        dpi_x = ctypes.c_uint()
        dpi_y = ctypes.c_uint()
        hr    = GetDpiForMonitor(hmon, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y))
        if hr == 0:
            return dpi_x.value / 96.0
    except Exception:
        pass
    return _primary_dpi()

@dataclass
class MatchResult:
    element:   object
    score:     float
    strategy:  str
    is_unique: bool = True

    @property
    def is_wrapper(self) -> bool:
        return not isinstance(self.element, tuple)


@dataclass
class StrategyStats:
    """Per-strategy statistics for adaptive ordering (ADV-1 / BUG-3)."""
    successes:       int   = 0
    failures:        int   = 0
    total_ms:        float = 0.0   
    priority_boost:  float = 0.0
    last_used_ts:    float = 0.0   

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

def _text_similarity(a: str, b: str) -> float:

    if not a or not b:
        return 0.0
    # Normalise
    def norm(s: str) -> str:
        return re.sub(r"[^\w\s]", " ", s.lower()).strip()
    a, b = norm(a), norm(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.85
    # Token overlap
    ta, tb = set(a.split()), set(b.split())
    if ta and tb:
        overlap = len(ta & tb) / len(ta | tb)
        if overlap >= 0.6:
            return overlap
    # Bigram fallback
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


def _composite_score(
    text_sim: float,
    spatial:  float,
    strategy_conf: float,
    boost:    float,
) -> float:
 
    return (
        text_sim      * 95.0 * 0.50 +
        (spatial/20.0)* 95.0 * 0.30 +
        strategy_conf * 95.0 * 0.20 +
        boost
    )

def _validate_wrapper(wrapper) -> bool:
    """True if wrapper is a real, usable UIA element."""
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
 
    try:
        if not wrapper.is_visible():
            return False
        r = wrapper.rectangle()
        if r.right  - r.left <= 0 or r.bottom - r.top <= 0:
            return False
        # Offscreen check — not entirely to left or above screen
        if r.right < -50 or r.bottom < -50:
            return False
        # Very large coords suggest hidden/virtual elements
        if r.left > 16000 or r.top > 16000:
            return False
        return True
    except Exception:
        return True   # assume visible on exception


def _refresh_wrapper(wrapper, app_cache: dict, hwnd: int, auto_id: Optional[str] = None,
                      name: Optional[str] = None):
  
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
    # Split on " - " which is Windows standard app title separator
    parts = [p.strip() for p in window_title.split(" - ")]
    for p in parts:
        if p and p not in variants:
            variants.append(p)
    # Also add a 30-char prefix for very long titles
    if len(window_title) > 30:
        prefix = window_title[:30]
        if prefix not in variants:
            variants.append(prefix)
    return variants


def _find_window_handles(window_title: str, process_name: Optional[str] = None,
                          max_results: int = 10) -> list[int]:
   
    handles = []

    # Try each title variant (CRIT-6)
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

class ElementMatcher:

    STRICT_THRESHOLD   = 85.0
    RELAXED_THRESHOLD  = 50.0
    AMBIGUITY_MARGIN   = 10.0  
    REJECT_ON_AMBIGUOUS = False  

    def __init__(
        self,
        screenshot_base_dir: Optional[Path] = None,
        browser = None,
    ):
        self._scr_dir  = screenshot_base_dir
        self._browser  = browser
        self._stats:   dict[str, StrategyStats] = {}
        self._stats_lock = threading.Lock()       # BUG-2
        self._app_cache: dict[int, object] = {}  

    def _get_stats(self, strategy: str) -> StrategyStats:
        with self._stats_lock:
            if strategy not in self._stats:
                # BUG-3: evict LRU entry if at cap
                if len(self._stats) >= MAX_STATS:
                    oldest = min(self._stats, key=lambda k: self._stats[k].last_used_ts)
                    del self._stats[oldest]
                self._stats[strategy] = StrategyStats()
            s = self._stats[strategy]
            s.last_used_ts = time.monotonic()
            return s

    def _ordered_strategies(self, candidates: list[str]) -> list[str]:
       
        def sort_key(name: str):
            with self._stats_lock:
                s = self._stats.get(name)
            if s is None:
                return (0.5, 999.0)
            return (-s.success_rate, s.avg_ms)
        return sorted(candidates, key=sort_key)

    def find(self, target: UITarget, event_id: int = 0):

        t = copy.copy(target)
        if t.backend == TargetBackend.BROWSER:
            return self._find_browser(t, event_id)
        return self._find_uia(t, event_id)

    def find_with_wait(
        self,
        target: UITarget,
        event_id: int = 0,
        timeout_ms: int = 10_000,
        poll_ms:    int = 300,
    ):
        deadline = time.monotonic() + timeout_ms / 1000
        last_exc = None
        attempt  = 0
        while time.monotonic() < deadline:
            attempt += 1
            try:
                result = self.find(target, event_id)
                # Check visibility for UIA wrappers
                if not isinstance(result, tuple):
                    if _elem_visible(result):
                        if attempt > 1:
                            logger.debug("[MATCH] Event #{} found after {} attempts",
                                         event_id, attempt)
                        return result
                    else:
                        logger.debug("[MATCH] Event #{} found but not visible yet (attempt {})",
                                     event_id, attempt)
                else:
                    # coord result — just return it
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

        # Wait for DOM to settle before searching
        self._browser.wait_for_dom_stable(stable_ms=250, max_wait_ms=2000)

        candidates = self._browser.find_candidates(bt, timeout_ms=8000)
        if not candidates:
            raise ElementNotFoundError(
                f"Event #{event_id}: browser element not found "
                f"xpath={bt.xpath!r} css={bt.css_selector!r} text={bt.inner_text!r}",
                event_id,
            )

        best = candidates[0]
        if len(candidates) >= 2:
            gap = candidates[0].score - candidates[1].score
            if gap < 10:
                logger.warning(
                    "[MATCH] Event #{} AMBIGUOUS: top-2 candidates close "
                    "(scores {:.0f} vs {:.0f}) — using highest-priority strategy",
                    event_id, candidates[0].score, candidates[1].score,
                )
                # Prefer id-anchored xpath over text
                priority = {"xpath": 0, "css": 1, "aria": 2, "exact_text": 3, "partial_text": 4}
                candidates.sort(key=lambda c: priority.get(c.strategy, 9))
                best = candidates[0]

        # Verify the chosen element is still valid
        verification = self._browser.verify_element(bt)
        if not verification.get("found"):
            logger.warning("[MATCH] Event #{} VERIFY FAILED — element stale, using coords anyway", event_id)

        logger.info(
            "[MATCH] Event #{} BROWSER → strategy={} score={:.0f} pos=({},{}) visible={} text='{}'",
            event_id, best.strategy, best.score, best.cx, best.cy, best.visible, best.text[:20],
        )
        self._get_stats(f"browser_{best.strategy}").record_success()
        return (best.cx, best.cy)


    def _find_uia(self, target: UITarget, event_id: int):
        if not UIA_OK:
            return self._coord_fallback(target, event_id)

        logger.info(
            "[MATCH] Event #{} UIA search → auto_id={} name='{}' ctrl_type={} "
            "class={} window='{}' app={}",
            event_id,
            target.automation_id, (target.name or "")[:30], target.control_type,
            target.class_name, (target.window_title or "")[:30], target.process_name,
        )

        results: list[MatchResult] = []
        errors   = []
        active_strategies = self._ordered_strategies(
            ["automation_id", "semantic", "relaxed", "classname", "ancestor"]
        )
        def run(name: str, fn):
            t0 = time.monotonic()
            try:
                r = fn()
                ms = (time.monotonic() - t0) * 1000
                if r:
                    rs = r if isinstance(r, list) else [r]
                    for x in rs:
                        self._get_stats(name).record_success(ms)
                    return rs
            except Exception as exc:
                ms = (time.monotonic() - t0) * 1000
                self._get_stats(name).record_failure(ms)
                errors.append(f"{name}: {exc}")
                logger.debug("[MATCH] Event #{} {} error: {}", event_id, name, exc)
            return []

        if target.automation_id and "automation_id" in active_strategies[:3]:
            for r in run("automation_id", lambda: self._by_automation_id(target)):
                results.append(r)
                logger.info("[MATCH] Event #{} auto_id='{}' → score={:.0f}",
                            event_id, target.automation_id, r.score)

        # ── 2. Semantic (name + ctrl) ─────────────────────────────────
        if target.name and target.control_type and "semantic" in active_strategies[:4]:
            for r in run("semantic", lambda: self._by_name_type(target, exact=True)):
                results.append(r)
                logger.info("[MATCH] Event #{} semantic '{}' ctrl={} → score={:.0f}",
                            event_id, (target.name or "")[:30], target.control_type, r.score)

        # Early exit: strict unique + cross-validated
        best_strict = self._pick_best(results, self.STRICT_THRESHOLD)
        if best_strict and best_strict.is_unique:
            if self._cross_validate(best_strict, target):
                # ADV-2: stability check — re-validate after brief wait
                if self._stability_check(best_strict.element):
                    logger.info("[MATCH] Event #{} STRICT ✓ strategy={} score={:.0f}",
                                event_id, best_strict.strategy, best_strict.score)
                    return best_strict.element
                else:
                    logger.warning("[MATCH] Event #{} stability check FAILED — retrying",
                                   event_id)
                    results.clear()   # element moved/changed — start fresh
            else:
                logger.warning("[MATCH] Event #{} cross-validation FAILED — relaxing", event_id)

        # ── 3. Relaxed (name fuzzy) ───────────────────────────────────
        if target.name and "relaxed" in active_strategies:
            for r in run("relaxed", lambda: self._by_name_type(target, exact=False)):
                results.append(r)
                logger.info("[MATCH] Event #{} relaxed '{}' → score={:.0f}",
                            event_id, (target.name or "")[:30], r.score)

        # ── 4. ClassName ──────────────────────────────────────────────
        if target.class_name and target.window_title and "classname" in active_strategies:
            for r in run("classname", lambda: self._by_classname(target)):
                results.append(r)

        # ── 5. Ancestor ───────────────────────────────────────────────
        if target.ancestor_chain and "ancestor" in active_strategies:
            for r in run("ancestor", lambda: self._by_ancestor(target)):
                results.append(r)

        # Pick best relaxed match
        best = self._pick_best(results, self.RELAXED_THRESHOLD)
        if best:
            logger.info("[MATCH] Event #{} RELAXED ✓ strategy={} score={:.0f} unique={}",
                        event_id, best.strategy, best.score, best.is_unique)
            return best.element

        # ── 6. BBox ───────────────────────────────────────────────────
        if target.bbox:
            r = self._by_bbox(target)
            if r:
                logger.info("[MATCH] Event #{} BBOX fallback", event_id)
                return r.element

        # ── 7. Screenshot (cropped region) ────────────────────────────
        if getattr(target, "screenshot_ref", None) and self._scr_dir:
            coords = self._by_screenshot_cropped(target)
            if coords:
                logger.info("[MATCH] Event #{} SCREENSHOT fallback pos={}", event_id, coords)
                return coords

        # ── 8. Raw coordinates ────────────────────────────────────────
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

    def _by_automation_id(self, t: UITarget) -> list[MatchResult]:
        boost = self._get_stats("automation_id").priority_boost
        found: list[MatchResult] = []

        # Window-scoped first (PERF-1, CRIT-6)
        handles = _find_window_handles(t.window_title or "", t.process_name)
        for hwnd in handles:
            try:
                app = self._app_cache.get(hwnd) or Application(backend="uia").connect(handle=hwnd)
                self._app_cache[hwnd] = app
                win  = app.window(handle=hwnd)
                elem = win.child_window(auto_id=t.automation_id)
                if elem.exists(timeout=0.5):
                    wrapper = elem.wrapper_object()
                    if not _validate_wrapper(wrapper):
                        continue
                    if not _elem_visible(wrapper):   # CRIT-7
                        logger.debug("[MATCH] auto_id found but not visible — skipping")
                        continue
                    spatial = _spatial_score(wrapper, t.bbox)
                    score   = 100.0 + boost + spatial
                    found.append(MatchResult(wrapper, score, "automation_id", is_unique=True))
                    return found   # window-scoped unique match — return immediately
            except Exception:
                continue

        
        if not found:
            try:
                for hwnd in find_windows(title_re=".*")[:MAX_WINDOW_SCAN]:
                    try:
                        app = (self._app_cache.get(hwnd) or
                               Application(backend="uia").connect(handle=hwnd))
                        self._app_cache[hwnd] = app
                        win  = app.window(handle=hwnd)
                        # Filter by process if available (PERF-1)
                        if t.process_name:
                            try:
                                import psutil
                                pid   = win.process_id()
                                pname = psutil.Process(pid).name().lower()
                                if t.process_name.lower() not in pname:
                                    continue
                            except Exception:
                                pass
                        descs = win.descendants(auto_id=t.automation_id)
                        for d in descs[:5]:
                            wrapper = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                            if not _validate_wrapper(wrapper) or not _elem_visible(wrapper):
                                continue
                            # Window title filter (CRIT-6 — partial match)
                            if t.window_title:
                                try:
                                    win_text = wrapper.top_level_parent().window_text() or ""
                                    if not any(v.lower() in win_text.lower()
                                               for v in _title_variants(t.window_title)):
                                        continue
                                except Exception:
                                    pass
                            spatial   = _spatial_score(wrapper, t.bbox)
                            is_unique = len(descs) == 1
                            score     = 90.0 + boost + spatial - (0 if is_unique else 15)
                            found.append(MatchResult(wrapper, score, "automation_id_wide", is_unique=is_unique))
                        if found:
                            break
                    except Exception:
                        continue
            except Exception:
                pass

        return found

    def _by_name_type(self, t: UITarget, exact: bool = True) -> list[MatchResult]:

        results = []
        boost   = self._get_stats("semantic" if exact else "relaxed").priority_boost
        name    = t.name or ""
        strategy = "semantic" if exact else "relaxed"
        conf_min = 0.85 if exact else 0.60
        name_re  = re.escape(name)

        handles = _find_window_handles(t.window_title or "", t.process_name)

        for hwnd in handles:
            try:
                app = (self._app_cache.get(hwnd) or
                       Application(backend="uia").connect(handle=hwnd))
                self._app_cache[hwnd] = app
                win = app.window(handle=hwnd)

                # child_window path (fast)
                try:
                    kwargs = {}
                    if t.control_type:
                        kwargs["control_type"] = t.control_type
                    elem = win.child_window(title_re=f".*{name_re}.*", **kwargs)
                    if elem.exists(timeout=0.4):
                        wrapper = elem.wrapper_object()
                        if _validate_wrapper(wrapper) and _elem_visible(wrapper):
                            sim  = _text_similarity(name, wrapper.window_text() or "")
                            if sim >= conf_min:
                                spatial = _spatial_score(wrapper, t.bbox)
                                score   = _composite_score(sim, spatial, self._get_stats(strategy).success_rate, boost)
                                results.append(MatchResult(wrapper, score, strategy, is_unique=True))
                                return results   # first window-scoped hit
                except Exception:
                    pass

                # descendants fallback (PERF-2: capped)
                try:
                    kwargs = {}
                    if t.control_type:
                        kwargs["control_type"] = t.control_type
                    descs = win.descendants(**kwargs) if kwargs else win.descendants()
                    for d in descs[:MAX_DESC_SEARCH]:
                        try:
                            wrapper = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                            if not _validate_wrapper(wrapper) or not _elem_visible(wrapper):
                                continue
                            txt = wrapper.window_text() or ""
                            sim = _text_similarity(name, txt)
                            if sim < conf_min:
                                continue
                            spatial = _spatial_score(wrapper, t.bbox)
                            score   = _composite_score(sim, spatial, self._get_stats(strategy).success_rate, boost)
                            results.append(MatchResult(wrapper, score, f"{strategy}_desc",
                                                        is_unique=False))
                        except Exception:
                            continue
                except Exception:
                    pass

            except Exception:
                continue

        return results[:5]

    def _by_classname(self, t: UITarget) -> Optional[MatchResult]:
        for hwnd in find_windows(title_re=f".*{re.escape(t.window_title[:30])}.*"):
            try:
                app  = Application(backend="uia").connect(handle=hwnd)
                elem = app.window(handle=hwnd).child_window(class_name=t.class_name)
                if elem.exists(timeout=0.5):
                    w = elem.wrapper_object()
                    spatial = _spatial_score(w, t.bbox)
                    boost   = self._get_stats("classname").priority_boost
                    return MatchResult(w, 65.0 + boost + spatial, "classname")
            except Exception:
                continue
        return None

    def _by_ancestor(self, t: UITarget) -> list[MatchResult]:
        """M-3: split with maxsplit=2 to handle colons in element text."""
        if not t.ancestor_chain:
            return []
        first    = t.ancestor_chain[0]
        parts    = first.split(":", 2)
        anc_text = parts[1].strip() if len(parts) > 1 else ""
        if not anc_text:
            return []
        results = []
        boost   = self._get_stats("ancestor").priority_boost
        try:
            for hwnd in find_windows(title_re=f".*{re.escape(anc_text[:20])}.*"):
                try:
                    app = (self._app_cache.get(hwnd) or
                           Application(backend="uia").connect(handle=hwnd))
                    self._app_cache[hwnd] = app
                    win    = app.window(handle=hwnd)
                    kwargs = {}
                    if t.control_type: kwargs["control_type"] = t.control_type
                    if t.class_name:   kwargs["class_name"]   = t.class_name
                    if kwargs:
                        descs = win.descendants(**kwargs)
                        for d in descs[:20]:
                            wrapper = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                            if _validate_wrapper(wrapper) and _elem_visible(wrapper):
                                spatial = _spatial_score(wrapper, t.bbox)
                                results.append(MatchResult(wrapper, 55.0 + boost + spatial, "ancestor"))
                                return results
                except Exception:
                    continue
        except Exception:
            pass
        return results

    def _by_bbox(self, t: UITarget) -> Optional[MatchResult]:
        try:
            dpi   = _primary_dpi() / (t.dpi_scale or 1.0)
            cx    = int(((t.bbox.left + t.bbox.right)  / 2) * dpi)
            cy    = int(((t.bbox.top  + t.bbox.bottom) / 2) * dpi)
            elem  = Desktop(backend="uia").from_point(cx, cy)
            boost = self._get_stats("bbox").priority_boost
            if elem:
                return MatchResult(elem.wrapper_object(), 30.0 + boost, "bbox")
            try:
                desktop = Desktop(backend="uia")
                elem    = desktop.from_point(cx, cy)
                if elem:
                    wrapper = elem.wrapper_object() if hasattr(elem, "wrapper_object") else elem
                    if _validate_wrapper(wrapper) and _elem_visible(wrapper):
                        return MatchResult(wrapper, 30.0 + boost, "bbox")
            except Exception:
                pass
            
        except Exception:
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
                monitors = sct.monitors[1:]  # skip [0] which is virtual all-monitors

                # BUG-4: find the monitor containing the recorded bbox
                capture_mon = monitors[0]  # default to primary
                if t.bbox:
                    scale = _primary_dpi() / (t.dpi_scale or 1.0)
                    bcx   = int(((t.bbox.left + t.bbox.right)  / 2) * scale)
                    bcy   = int(((t.bbox.top  + t.bbox.bottom) / 2) * scale)
                    for mon in monitors:
                        if (mon["left"] <= bcx < mon["left"] + mon["width"] and
                                mon["top"] <= bcy < mon["top"] + mon["height"]):
                            capture_mon = mon
                            break

                if t.bbox:
                    scale  = _primary_dpi() / (t.dpi_scale or 1.0)
                    margin = 200
                    region = {
                        "left":   max(capture_mon["left"], int(t.bbox.left   * scale) - margin),
                        "top":    max(capture_mon["top"],  int(t.bbox.top    * scale) - margin),
                        "width":  tw + margin * 2 + int((t.bbox.right  - t.bbox.left)  * scale),
                        "height": th + margin * 2 + int((t.bbox.bottom - t.bbox.top) * scale),
                    }
                    # Clamp to monitor bounds
                    region["width"]  = min(region["width"],
                                           capture_mon["left"] + capture_mon["width"]  - region["left"])
                    region["height"] = min(region["height"],
                                           capture_mon["top"]  + capture_mon["height"] - region["top"])
                else:
                    region = capture_mon

                if region["width"] <= tw or region["height"] <= th:
                    region = capture_mon   # fallback to full monitor

                shot = sct.grab(region)
                bgr  = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)

            res = cv2.matchTemplate(bgr, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            logger.debug("[MATCH] screenshot confidence={:.2%}", max_val)
            if max_val < 0.80:
                return None
            # Convert back to screen coordinates
            screen_x = region["left"] + max_loc[0] + tw // 2
            screen_y = region["top"]  + max_loc[1] + th // 2
            return (screen_x, screen_y)
        except Exception as exc:
            logger.debug("[MATCH] screenshot error: {}", exc)
            return None


    def _cross_validate(self, match: MatchResult, target: UITarget) -> bool:
        """M-4: compare name (primary) or ctrl_type (fallback)."""
        if not match.is_wrapper:
            return True
        try:
            elem = match.element
            if not _validate_wrapper(elem):
                return False
            if target.name:
                found_text = elem.window_text() or ""
                sim        = _text_similarity(target.name, found_text)
                ok         = sim >= 0.65
                logger.debug("[MATCH] cross-validate name sim={:.2f} ok={}", sim, ok)
                return ok
            if target.control_type:
                try:
                    found_ctrl = (elem.friendly_class_name()
                                  if callable(getattr(elem, "friendly_class_name", None))
                                  else str(getattr(elem, "friendly_class_name", "")))
                    ok = target.control_type.lower() in found_ctrl.lower()
                    return ok
                except Exception:
                    return True
        except Exception:
            pass
        return True

    def _stability_check(self, element) -> bool:
      
        if isinstance(element, tuple):
            return True  # coords don't go stale
        try:
            rect1 = element.rectangle()
            time.sleep(STABILITY_WAIT_MS / 1000)
            rect2 = element.rectangle()
            # Allow 2px movement tolerance (sub-pixel rendering differences)
            moved = (abs(rect1.left - rect2.left) > 2 or
                     abs(rect1.top  - rect2.top)  > 2)
            if moved:
                logger.debug("[MATCH] stability check: element moved ({},{})→({},{})",
                             rect1.left, rect1.top, rect2.left, rect2.top)
                return False
            return _elem_visible(element)
        except Exception:
            return False

  
    def _pick_best(
        self,
        results: list[MatchResult],
        threshold: float,
    ) -> Optional[MatchResult]:
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
                # Attempt resolution: prefer unique
                unique = [r for r in viable if r.is_unique]
                if len(unique) == 1:
                    logger.info("[MATCH] Ambiguity resolved by uniqueness")
                    return unique[0]
                # If REJECT_ON_AMBIGUOUS set and still unresolved — caller handles
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
