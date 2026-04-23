
from __future__ import annotations
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

MAX_WINDOW_SCAN    = 25      
MAX_DESC_SEARCH    = 50     
MAX_STATS          = 300    
DESC_TIMEOUT_S     = 1.5    


STABILITY_WAIT_MS  = 80  
STABILITY_CHECK_ENABLED = True   
MAX_REFRESH_ATTEMPTS   = 2     



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
        MonitorFromPoint = ctypes.windll.user32.MonitorFromPoint
        GetDpiForMonitor = ctypes.windll.shcore.GetDpiForMonitor

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        hmon  = ctypes.windll.user32.MonitorFromPoint(POINT(x, y), 2)
        dpi_x = ctypes.c_uint(); dpi_y = ctypes.c_uint()
        if ctypes.windll.shcore.GetDpiForMonitor(
                hmon, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y)) == 0:
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
    elem_hash:str = ""

    @property
    def is_wrapper(self) -> bool:
        return not isinstance(self.element, tuple)


@dataclass
class StrategyStats:
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
    return (text_sim * 95*0.45 + (spatial/20)*95*0.20 +
            strategy_rate*95*0.15 + sel_conf*95*0.15 +
            tp*95*0.05 + boost + rb)

def _tp(ctrl: Optional[str]) -> float:
    return _ELEMENT_TYPE_PRIORITY.get(ctrl, 50) / 100.0 if ctrl else 0.5


def _role_boost(ctrl_type: Optional[str], role: Optional[str]) -> float:
    if role:
        return _ROLE_SCORE_BOOST.get(role, 0.0)
    if not ctrl_type:
        return 0.0
    ct = ctrl_type.lower()
    if "button" in ct:    return _ROLE_SCORE_BOOST["button"]
    if "edit" in ct:      return _ROLE_SCORE_BOOST["input"]
    if "document" in ct:  return _ROLE_SCORE_BOOST["input"]
    if "label" in ct or "text" == ct or "statictext" in ct:
        return _ROLE_SCORE_BOOST["label"]
    if "pane" in ct or "group" in ct:
        return _ROLE_SCORE_BOOST["container"]
    return 0.0


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
    
def _elem_hash(w) -> str:
    try:
        r = w.rectangle()
        key = f"{r.left},{r.top},{r.right},{r.bottom}"
        try:
            aid = w.automation_id() if callable(getattr(w,"automation_id",None)) else ""
            if aid: key += f"|{aid}"
        except Exception: pass
        return hashlib.md5(key.encode()).hexdigest()[:12]
    except Exception:
        return str(id(w))


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
    def _sel_conf(self, target: UITarget) -> float:
        # Prefer full Selector objects (selector.py) over stub selectors
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
        # Prefer rich_selectors; fall back to stub selectors
        rich = getattr(target, "rich_selectors", None)
        sels = rich if rich else getattr(target, "selectors", None)
        if not sels:
            return
        mapping = {
            "automation_id": "strict", "automation_id_wide": "strict",
            "semantic":      "semantic", "semantic_desc":    "semantic",
            "relaxed":       "relaxed",  "relaxed_desc":     "relaxed",
            "classname":     "classname", "ancestor":         "ancestor",
            "coord":         "positional", "bbox":            "positional",
        }
        sel_name = mapping.get(strategy, strategy)
        try:
            for sel in sels:
                if hasattr(sel, "record_replay"):
                    sel.record_replay(sel_name, success)
        except Exception:
            pass

    def _selector_order(self, target: UITarget) -> list[str]:
        """M11: preferred strategy order from selector history (rich_selectors first)."""
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
        pref_boost = {n: (len(pref)-i)*0.1 for i,n in enumerate(pref)}
        def key(n):
            with self._stats_lock: s = self._stats.get(n)
            rate = s.success_rate if s else 0.5
            ms   = s.avg_ms       if s else 999.0
            return (-(rate + pref_boost.get(n,0.0)), ms)
        return sorted(names, key=key)


    def find(self, target: UITarget, event_id: int = 0, action_intent: Optional[str]=None):

        t = copy.copy(target)
        if t.backend == TargetBackend.BROWSER:
            return self._find_browser(t, event_id)
        return self._find_uia(t, event_id,action_intent)

    def find_with_wait(
        self,
        target: UITarget,
        event_id: int = 0,
        timeout_ms: int = 10_000,
        poll_ms:    int = 300,
        action_intent: Optional[str]=None
    ):
        deadline = time.monotonic() + timeout_ms / 1000
        last_exc = None
        attempt  = 0
        while time.monotonic() < deadline:
            attempt += 1
            try:
                result = self.find(target, event_id,action_intent)
                if not isinstance(result, tuple):
                    if _elem_visible(result):
                        if attempt > 1:
                            logger.debug("[MATCH] Event #{} found after {} attempts",
                                         event_id, attempt)
                        return result
                    else:
                        logger.debug("[MATCH] Event #{} found but not visible (attempt {})",                               event_id, attempt)
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


    def _find_uia(self, target: UITarget, event_id: int,action_intent: Optional[str]=None):
        if not UIA_OK:
            return self._coord_fallback(target, event_id)
        proc  = (target.process_name or "").lower()
        cls   = (target.class_name   or "")
        is_electron = proc in _ELECTRON_PROCS or cls in _ELECTRON_CLASSES
        is_system   = proc in _SYSTEM_PROCS
        
        if is_electron:
            logger.info("[MATCH] Event #{} ELECTRON → image/coord", event_id)
            coords = self._by_screenshot_cropped(target)
            if coords:
                return coords
            return self._coord_fallback(target, event_id)
        
        confidence_level = getattr(target, "confidence_level", "medium")
        strict_thresh = (
            self.STRICT_THRESHOLD + 5 if confidence_level == "high"
            else self.STRICT_THRESHOLD
        )

        active_bonus = _active_window_bonus(target)
        if active_bonus < 0:
            logger.debug("[MATCH] Event #{} target window not active (penalty {:.0f})",
                         event_id, active_bonus)
        sel_conf     = self._sel_conf(target)          # M7
        pref_order   = self._selector_order(target)    # M11
       
        conf_level   = getattr(target, "confidence_level", "medium")
        strict_th    = self.STRICT_THRESHOLD + (5 if conf_level == "high" else 0)
        logger.info("[MATCH] Event #{} UIA → auto_id={} name='{}' ctrl={} win='{}' "
                    "intent={} sel_conf={:.2f}",
                    event_id, target.automation_id or "(none)",
                    (target.name or "")[:30], target.control_type or "?",
                    (target.window_title or "")[:30],
                    action_intent or "any", sel_conf)

        results: list[MatchResult] = []
        errors :list[str]  = []
        active_strategies = self._ordered_strategies(
            ["automation_id", "semantic", "relaxed", "classname", "ancestor"]
        )
        def run(name: str, fn):
            t0 = time.monotonic()
            try:
                r = fn(); ms = (time.monotonic()-t0)*1000
                if r:
                    rs = r if isinstance(r,list) else [r]
                    for x in rs: self._get_stats(name).record_success(ms)
                    return rs
            except Exception as exc:
                ms = (time.monotonic()-t0)*1000
                self._get_stats(name).record_failure(ms)
                errors.append(f"{name}: {exc}")
                logger.debug("[MATCH] Event #{} {} error: {}", event_id, name, exc)
            return []

        if target.automation_id and "automation_id":
            for r in run("automation_id", lambda: self._by_automation_id(target)):
                r.score += active_bonus
                results.append(r)
                logger.info("[MATCH] Event #{} auto_id='{}' → score={:.0f}",
                            event_id, target.automation_id, r.score)

        # ── 2. Semantic (name + ctrl) ─────────────────────────────────
        if target.name and target.control_type:
            for r in run("semantic", lambda: self._by_name_type(target, exact=True)):
                r.score += active_bonus
                results.append(r)
                logger.info("[MATCH] Event #{} semantic '{}' ctrl={} → score={:.0f}",
                            event_id, (target.name or "")[:30], target.control_type, r.score)

        best_strict = self._pick_best(results, self.STRICT_THRESHOLD)
        if best_strict and best_strict.is_unique:
            if self._cross_validate(best_strict, target):
               
                if self._stability_check(best_strict.element):
                    logger.info("[MATCH] Event #{} STRICT ✓ strategy={} score={:.0f}",
                                event_id, best_strict.strategy, best_strict.score)
                    return self._wrap_safe(best_strict.element, target)
                else:
                    logger.warning("[MATCH] Event #{} stability check FAILED — retrying",
                                   event_id)
                    results.clear()   
            else:
                logger.warning("[MATCH] Event #{} cross-validation FAILED — relaxing", event_id)

        # ── 3. Relaxed (name fuzzy) ───────────────────────────────────
        if target.name and "relaxed" in active_strategies:
            for r in run("relaxed", lambda: self._by_name_type(target, exact=False)):
                r.score += active_bonus
                results.append(r)
                logger.info("[MATCH] Event #{} relaxed '{}' → score={:.0f}",
                            event_id, (target.name or "")[:30], r.score)

        # ── 4. ClassName ──────────────────────────────────────────────
        if target.class_name and target.window_title and "classname" in active_strategies:
            for r in run("classname", lambda: self._by_classname(target)):
                r.score += active_bonus
                results.append(r)

        # ── 5. Ancestor ───────────────────────────────────────────────
        if target.ancestor_chain and "ancestor" in active_strategies:
            for r in run("ancestor", lambda: self._by_ancestor(target)):
                r.score += active_bonus
                results.append(r)

        # ── Apply intent filter to focus on role-appropriate elements ─────
        if action_intent:
            results = self._intent_filter(results, action_intent)

        best = self._pick_best(results, self.RELAXED_THRESHOLD)
        if best:
            logger.info("[MATCH] Event #{} RELAXED ✓ strategy={} score={:.0f} unique={}",
                        event_id, best.strategy, best.score, best.is_unique)
            return self._wrap_safe(best.element, target)

        # ── Self-heal before falling to geometry ──────────────────────────
        healed = self._self_heal(target, event_id, active_bonus)
        if healed is not None:
            return healed

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
        
    def _deduplicate(self, results: list[MatchResult]) -> list[MatchResult]:
        seen: dict[str, MatchResult] = {}
        for r in results:
            key = str(r.element) if not r.is_wrapper else (r.elem_hash or _elem_hash(r.element))
            if key not in seen or r.score > seen[key].score:
                seen[key] = r
        return list(seen.values())

    # ── Intent filter ────────────────────────────────────────────────

    def _intent_filter(self, results: list[MatchResult],
                        intent: str) -> list[MatchResult]:
        if intent not in ("type","click","select"):
            return results
        out = []
        for r in results:
            score = r.score
            if r.is_wrapper:
                try:
                    raw = r.element.raw() if hasattr(r.element,"raw") else r.element
                    ct  = (raw.friendly_class_name()
                           if callable(getattr(raw,"friendly_class_name",None)) else "") or ""
                    ct = ct.lower()
                    if intent == "type":
                        if any(x in ct for x in ("edit","document","richtext")): score += 10
                        elif any(x in ct for x in ("button","label","pane")):    score -= 15
                    elif intent == "click":
                        if any(x in ct for x in ("button","menuitem","tabitem")): score += 5
                        elif any(x in ct for x in ("label","statictext")):        score -= 8
                    elif intent == "select":
                        if any(x in ct for x in ("combobox","listitem","list")):  score += 10
                except Exception:
                    pass
            out.append(MatchResult(r.element, score, r.strategy, r.is_unique, r.elem_hash))
        return out
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

    def _by_automation_id(self, t: UITarget, sc: float=0.5) -> list[MatchResult]:
        boost = self._get_stats("automation_id").priority_boost
        found: list[MatchResult] = []

        for hwnd in _find_window_handles(t.window_title or "", t.process_name):
            try:
                app = (self._app_cache.get(hwnd) or
                       Application(backend="uia").connect(handle=hwnd))
                self._app_cache[hwnd] = app
                elem = app.window(handle=hwnd).child_window(auto_id=t.automation_id)
                if elem.exists(timeout=0.5):
                    w = elem.wrapper_object()
                    if _validate_wrapper(w) and _elem_visible(w):
                        sp  = _spatial_score(w, t.bbox)
                        rb  = _role_boost(t.control_type, getattr(t,"element_role",None))
                        tp  = _tp(t.control_type)
                        sc_ = _composite_score(1.0, sp,
                              self._get_stats("automation_id").success_rate,
                              boost, sc, rb, tp)
                        found.append(MatchResult(w, sc_, "automation_id",
                                                  is_unique=True, elem_hash=_elem_hash(w)))
                        # M1: no return — continue to other windows
            except Exception:
                continue

        if not found:
            try:
                for hwnd in find_windows(title_re=".*")[:MAX_WINDOW_SCAN]:
                    try:
                        app = (self._app_cache.get(hwnd) or
                               Application(backend="uia").connect(handle=hwnd))
                        self._app_cache[hwnd] = app
                        win = app.window(handle=hwnd)
                        if t.process_name:
                            try:
                                import psutil
                                pname = psutil.Process(win.process_id()).name().lower()
                                if t.process_name.lower() not in pname: continue
                            except Exception: pass
                        descs = win.descendants(auto_id=t.automation_id)
                        for d in descs[:5]:
                            w = d.wrapper_object() if hasattr(d,"wrapper_object") else d
                            if not _validate_wrapper(w) or not _elem_visible(w): continue
                            if t.window_title:
                                try:
                                    wt = w.top_level_parent().window_text() or ""
                                    if not any(v.lower() in wt.lower()
                                               for v in _title_variants(t.window_title)):
                                        continue
                                except Exception: pass
                            sp  = _spatial_score(w, t.bbox)
                            rb  = _role_boost(t.control_type, getattr(t,"element_role",None))
                            tp  = _tp(t.control_type)
                            uniq= len(descs)==1
                            sc_ = _composite_score(1.0, sp,
                                  self._get_stats("automation_id_wide").success_rate,
                                  boost-10, sc, rb, tp)
                            found.append(MatchResult(w, sc_, "automation_id_wide",
                                                      is_unique=uniq, elem_hash=_elem_hash(w)))
                        # M1: no break — keep scanning
                    except Exception: continue
            except Exception: pass

        return found

    def _by_name_type(self, t: UITarget, exact: bool=True,
                       sc: float=0.5) -> list[MatchResult]:
        """M2: collects from child_window AND descendants — no early returns."""
        results  = []
        strategy = "semantic" if exact else "relaxed"
        boost    = self._get_stats(strategy).priority_boost
        name     = t.name or ""; conf_min = 0.85 if exact else 0.60
        name_re  = re.escape(name)

        for hwnd in _find_window_handles(t.window_title or "", t.process_name):
            try:
                app = (self._app_cache.get(hwnd) or
                       Application(backend="uia").connect(handle=hwnd))
                self._app_cache[hwnd] = app
                win = app.window(handle=hwnd)

                # child_window — M2: no early return
                try:
                    kw = {"control_type": t.control_type} if t.control_type else {}
                    elem = win.child_window(title_re=f".*{name_re}.*", **kw)
                    if elem.exists(timeout=0.4):
                        w = elem.wrapper_object()
                        if _validate_wrapper(w) and _elem_visible(w):
                            sim = _text_similarity(name, w.window_text() or "")
                            if sim >= conf_min:
                                sp  = _spatial_score(w, t.bbox)
                                rb  = _role_boost(t.control_type, getattr(t,"element_role",None))
                                tp  = _tp(t.control_type)
                                sc_ = _composite_score(sim, sp,
                                      self._get_stats(strategy).success_rate,
                                      boost, sc, rb, tp)
                                results.append(MatchResult(w, sc_, strategy,
                                    is_unique=True, elem_hash=_elem_hash(w)))
                                # M2: continue — still run descendants
                except Exception: pass

                # Descendants — M2: always, not just fallback
                try:
                    kw = {"control_type": t.control_type} if t.control_type else {}
                    descs = win.descendants(**kw) if kw else win.descendants()
                    for d in descs[:MAX_DESC_SEARCH]:
                        try:
                            w = d.wrapper_object() if hasattr(d,"wrapper_object") else d
                            if not _validate_wrapper(w) or not _elem_visible(w): continue
                            sim = _text_similarity(name, w.window_text() or "")
                            if sim < conf_min: continue
                            sp  = _spatial_score(w, t.bbox)
                            rb  = _role_boost(t.control_type, getattr(t,"element_role",None))
                            tp  = _tp(t.control_type)
                            sc_ = _composite_score(sim, sp,
                                  self._get_stats(strategy).success_rate,
                                  boost, sc, rb, tp)
                            results.append(MatchResult(w, sc_, f"{strategy}_desc",
                                is_unique=False, elem_hash=_elem_hash(w)))
                        except Exception: continue
                except Exception: pass
            except Exception: continue

        return results[:10]

    def _by_classname(self, t: UITarget, sc: float=0.5) -> list[MatchResult]:
        """M3: accumulates all classname matches — no early return."""
        results = []; boost = self._get_stats("classname").priority_boost
        for hwnd in _find_window_handles(t.window_title or ""):
            try:
                app = (self._app_cache.get(hwnd) or
                       Application(backend="uia").connect(handle=hwnd))
                self._app_cache[hwnd] = app
                win = app.window(handle=hwnd)
                try:
                    elem = win.child_window(class_name=t.class_name)
                    if elem.exists(timeout=0.4):
                        w = elem.wrapper_object()
                        if _validate_wrapper(w) and _elem_visible(w):
                            sp = _spatial_score(w, t.bbox)
                            rb = _role_boost(t.control_type, getattr(t,"element_role",None))
                            results.append(MatchResult(w, 65+boost+sp+rb, "classname",
                                                        elem_hash=_elem_hash(w)))
                            # M3: no return
                except Exception: pass
                try:
                    descs = win.descendants(class_name=t.class_name)
                    for d in descs[:10]:
                        w = d.wrapper_object() if hasattr(d,"wrapper_object") else d
                        if _validate_wrapper(w) and _elem_visible(w):
                            sp = _spatial_score(w, t.bbox)
                            rb = _role_boost(t.control_type, getattr(t,"element_role",None))
                            results.append(MatchResult(w, 60+boost+sp+rb, "classname_desc",
                                is_unique=len(descs)==1, elem_hash=_elem_hash(w)))
                except Exception: pass
            except Exception: continue
        return results

    def _by_ancestor(self, t: UITarget, sc: float=0.5) -> list[MatchResult]:
        """M4: scans all desc candidates — no early return."""
        if not t.ancestor_chain: return []
        parts    = t.ancestor_chain[0].split(":", 2)
        anc_text = parts[1].strip() if len(parts) > 1 else ""
        if not anc_text: return []

        results     = []; boost = self._get_stats("ancestor").priority_boost
        anch_names  = [a.name for a in (getattr(t,"anchor_elements",None) or [])
                       if getattr(a,"name",None)]
        try:
            for hwnd in find_windows(title_re=f".*{re.escape(anc_text[:20])}.*"):
                try:
                    app = (self._app_cache.get(hwnd) or
                           Application(backend="uia").connect(handle=hwnd))
                    self._app_cache[hwnd] = app
                    win = app.window(handle=hwnd)
                    kw  = {}
                    if t.control_type: kw["control_type"] = t.control_type
                    if t.class_name:   kw["class_name"]   = t.class_name
                    if not kw: continue
                    descs = win.descendants(**kw)
                    for d in descs[:20]:  # M4: all, no return on first
                        w = d.wrapper_object() if hasattr(d,"wrapper_object") else d
                        if not _validate_wrapper(w) or not _elem_visible(w): continue
                        sp   = _spatial_score(w, t.bbox)
                        ab   = self._anchor_confirmation(w, anch_names)
                        rb   = _role_boost(t.control_type, getattr(t,"element_role",None))
                        results.append(MatchResult(w, 55+boost+sp+rb+ab, "ancestor",
                                                    elem_hash=_elem_hash(w)))
                except Exception: continue
        except Exception: pass
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
        return min(bonus, 15.0)  # cap at 3 confirmed anchors

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
                monitors = sct.monitors[1:]
                capture_mon = monitors[0]
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
                sim        = _text_similarity(target.name, found_text)
                ok         = sim >= 0.65
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
        """M12: non-blocking; skips sleep when element already stable."""
        if not STABILITY_CHECK_ENABLED: return True
        if isinstance(element, tuple): return True
        raw = element.raw() if hasattr(element,"raw") else element
        try:
            r1 = raw.rectangle()
            # M12: only sleep if element might be animating (not enabled = loading)
            try:
                if raw.is_enabled():
                    time.sleep(STABILITY_WAIT_MS/1000)
            except Exception:
                time.sleep(STABILITY_WAIT_MS/1000)
            r2 = raw.rectangle()
            if abs(r1.left-r2.left)>2 or abs(r1.top-r2.top)>2: return False
            return _elem_visible(raw)
        except Exception:
            return False

    # Alias so _find_uia can call either name (historical and fast variant)
    _stability_check = _stability_check_fast
        
    # ── Context gate ──────────────────────────────────────────────────────────────────

    def _context_gate(self, results: list[MatchResult], target: UITarget) -> list[MatchResult]:
        """
        Score-penalise candidates that are not in the expected window/process.
        Candidates from the correct window get a +20 boost; wrong window get -25.
        """
        if not target.window_title and not target.process_name:
            return results
        out = []
        for r in results:
            if not r.is_wrapper:
                out.append(r)
                continue
            bonus = 0.0
            try:
                raw = r.element.raw() if hasattr(r.element, "raw") else r.element
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

    # ── Self-healing layer ─────────────────────────────────────────────────────────────

    def _self_heal(self, target: UITarget, event_id: int,
                   active_bonus: float = 0.0) -> Optional[object]:
        """
        Three-strategy self-healing layer, tried in order:
          A) Ctrl-type-free name match (removes ctrl_type constraint)
          B) Element-hash scan across all windows (identity-based match)
          C) Anchor-spatial search (finds element near a stable sibling)
        """
        logger.info("[MATCH] Event #{} SELF-HEAL starting", event_id)
        results: list[MatchResult] = []

        # ─ A: ctrl-type-free relaxed name match ──────────────────────────
        if target.name:
            try:
                t_loose = copy.copy(target)
                t_loose.control_type = None
                for r in self._by_name_type(t_loose, exact=False):
                    results.append(MatchResult(
                        r.element, r.score * 0.75 + active_bonus,
                        "heal_relaxed", r.is_unique, r.elem_hash))
            except Exception:
                pass

        # ─ B: element-hash scan ────────────────────────────────────────
        target_hash = getattr(target, "element_hash", None)
        if target_hash and UIA_OK:
            try:
                from .uia_enricher import _element_hash
                for hwnd in find_windows(title_re=".*")[:MAX_WINDOW_SCAN]:
                    try:
                        app = (self._app_cache.get(hwnd) or
                               Application(backend="uia").connect(handle=hwnd))
                        self._app_cache[hwnd] = app
                        win = app.window(handle=hwnd)
                        kw = {"control_type": target.control_type} if target.control_type else {}
                        descs = win.descendants(**kw) if kw else win.descendants()
                        for d in descs[:MAX_DESC_SEARCH]:
                            try:
                                w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                                if not _validate_wrapper(w) or not _elem_visible(w):
                                    continue
                                aid, nm, ct, cn = None, None, None, None
                                try: aid = w.automation_id()
                                except: pass
                                try: nm = w.window_text()
                                except: pass
                                try: ct = w.friendly_class_name()
                                except: pass
                                try: cn = w.class_name()
                                except: pass
                                if _element_hash(aid, nm, ct, cn) == target_hash:
                                    sp = _spatial_score(w, target.bbox)
                                    results.append(MatchResult(
                                        w, 65 + sp + active_bonus,
                                        "heal_hash", is_unique=True,
                                        elem_hash=target_hash))
                            except Exception:
                                continue
                    except Exception:
                        continue
            except Exception:
                pass

        # ─ C: anchor-spatial search ───────────────────────────────────
        anchors = getattr(target, "anchor_elements", None) or []
        anchor_names = [a.name for a in anchors if getattr(a, "name", None)]
        if anchor_names and target.window_title and UIA_OK:
            try:
                for hwnd in _find_window_handles(target.window_title):
                    try:
                        app = (self._app_cache.get(hwnd) or
                               Application(backend="uia").connect(handle=hwnd))
                        self._app_cache[hwnd] = app
                        win = app.window(handle=hwnd)
                        kw = {"control_type": target.control_type} if target.control_type else {}
                        descs = win.descendants(**kw) if kw else win.descendants()
                        for d in descs[:MAX_DESC_SEARCH]:
                            try:
                                w = d.wrapper_object() if hasattr(d, "wrapper_object") else d
                                if not _validate_wrapper(w) or not _elem_visible(w):
                                    continue
                                ab = self._anchor_confirmation(w, anchor_names)
                                if ab > 0:
                                    sp = _spatial_score(w, target.bbox)
                                    results.append(MatchResult(
                                        w, 45 + ab + sp + active_bonus,
                                        "heal_anchor", is_unique=False,
                                        elem_hash=_elem_hash(w)))
                            except Exception:
                                continue
                    except Exception:
                        continue
            except Exception:
                pass

        if not results:
            return None

        # Apply context gate and pick best
        results = self._context_gate(results, target)
        best = self._pick_best(results, 30.0)  # lower threshold for healed match
        if best:
            logger.info("[MATCH] Event #{} SELF-HEAL ✓ strategy={} score={:.0f}",
                        event_id, best.strategy, best.score)
            return self._wrap_safe(best.element, target)
        return None

    def _check_uniqueness(self, wrapper, target: UITarget, descs: list) -> bool:
      
        target_hash = getattr(target, "element_hash", None)
        if not target_hash:
            return len(descs) == 1
        try:
            from .uia_enricher import _element_hash
            auto_id    = None
            name       = None
            ctrl_type  = None
            class_name = None
            try: auto_id    = wrapper.automation_id()
            except Exception: pass
            try: name       = wrapper.window_text()
            except Exception: pass
            try: ctrl_type  = wrapper.friendly_class_name()
            except Exception: pass
            try: class_name = wrapper.class_name()
            except Exception: pass
            found_hash = _element_hash(auto_id, name, ctrl_type, class_name)
            if found_hash == target_hash:
                return True 
        except Exception:
            pass
        return len(descs) == 1

  
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

         
                if len(viable) >= 2:
                    elem0 = viable[0].element
                    elem1 = viable[1].element
                    try:
                        raw0 = elem0.raw() if hasattr(elem0, "raw") else elem0
                        raw1 = elem1.raw() if hasattr(elem1, "raw") else elem1
                        ct0  = (raw0.friendly_class_name() if callable(
                                getattr(raw0, "friendly_class_name", None)) else "") or ""
                        ct1  = (raw1.friendly_class_name() if callable(
                                getattr(raw1, "friendly_class_name", None)) else "") or ""
                        p0   = _ELEMENT_TYPE_PRIORITY.get(ct0, 50)
                        p1   = _ELEMENT_TYPE_PRIORITY.get(ct1, 50)
                        if abs(p0 - p1) >= 15:  # clear type priority difference
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

    def _coord_fallback(self, t: UITarget, event_id: int):
        rel = getattr(t, "relative_to_window", None)
        if rel and t.window_title:
            try:
                handles = _find_window_handles(t.window_title)
                if handles:
                    app = Application(backend="uia").connect(handle=handles[0])
                    win = app.window(handle=handles[0])
                    wr  = win.rectangle()
                    cx  = wr.left + rel["x"] + rel["w"] // 2
                    cy  = wr.top  + rel["y"] + rel["h"] // 2
                    logger.debug("[MATCH] Event #{} relative coord fallback ({},{})",
                                 event_id, cx, cy)
                    return (cx, cy)
            except Exception:
                pass

       
        raw = getattr(t, "raw_bbox", None)
        if raw:
            cx = (raw.left + raw.right)  // 2
            cy = (raw.top  + raw.bottom) // 2
            # raw_bbox is already in exact screen pixels at recording time.
            # Only scale if the primary DPI fundamentally changed between sessions.
            recorded_dpi = t.dpi_scale or 1.0
            current_dpi  = _primary_dpi()
            if abs(current_dpi - recorded_dpi) > 0.05:
                cx = int(cx * current_dpi / recorded_dpi)
                cy = int(cy * current_dpi / recorded_dpi)
            return (cx, cy)

        if t.bbox:
            # t.bbox is a normalized (unscaled) bbox in logical pixels.
            # Scale it strictly using the current monitor's local DPI.
            cx_logical = (t.bbox.left + t.bbox.right) / 2
            cy_logical = (t.bbox.top + t.bbox.bottom) / 2
            current_dpi = _dpi_for_point(int(cx_logical), int(cy_logical))
            cx = int(cx_logical * current_dpi)
            cy = int(cy_logical * current_dpi)
            return (cx, cy)

        if t.screen_x is not None:
            return (t.screen_x, t.screen_y)

        raise ElementNotFoundError(f"Event #{event_id}: no fallback coordinates", event_id)

