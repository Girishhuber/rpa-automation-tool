"""
BrowserBridge — CDP bridge with full reliability + structured debug logging.

New in this version:
  - find_candidates(): returns top-3 scored matches (not just first)
  - verify_element(): re-checks text/tag/role before returning coords
  - wait_for_dom_stable(): waits until no DOM mutations for ~300ms
  - Detailed [BROWSER] log prefix on every action for debugging
  - Viewport offset recomputed periodically (handles resize/move)
  - Shadow DOM + iframe traversal retained
  - Human-like typing retained
"""

from __future__ import annotations
import json
import queue
import random
import threading
import time
from dataclasses import dataclass
from typing import Any, Optional
from urllib.request import urlopen

from utils.logger import logger
from models.target import BrowserTarget

try:
    import websocket
    WS_AVAILABLE = True
except ImportError:
    WS_AVAILABLE = False
    logger.warning("[BROWSER] websocket-client not installed — pip install websocket-client")


# ─────────────────────────────────────────────────────────────────────────────
# Scored candidate
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class BrowserCandidate:
    cx:       int
    cy:       int
    score:    float        # 0–100
    strategy: str          # xpath_id / css_id / aria / text / tag
    tag:      str = ""
    text:     str = ""
    visible:  bool = True


# ─────────────────────────────────────────────────────────────────────────────
# JS function bodies (all use named params — no arguments[] scope issues)
# ─────────────────────────────────────────────────────────────────────────────

_JS_ELEMENT_AT = """\
(function(px, py) {
  var el = document.elementFromPoint(px, py);
  if (!el) return null;
  function xpath(e) {
    if (e.id) return '//*[@id="' + e.id + '"]';
    if (e === document.body) return '/html/body';
    var ix = 0, sibs = e.parentNode ? e.parentNode.childNodes : [];
    for (var i = 0; i < sibs.length; i++) {
      var s = sibs[i];
      if (s === e) return xpath(e.parentNode) + '/' + e.tagName.toLowerCase() + '[' + (ix+1) + ']';
      if (s.nodeType === 1 && s.tagName === e.tagName) ix++;
    }
  }
  function cssPath(e) {
    var parts = [];
    while (e && e.nodeType === 1) {
      var sel = e.tagName.toLowerCase();
      if (e.id) { parts.unshift('#' + e.id); break; }
      var cls = Array.from(e.classList||[]).filter(function(c){return !/[0-9]{3,}/.test(c);}).slice(0,2).join('.');
      if (cls) sel += '.' + cls;
      var n = 1, prev = e;
      while ((prev = prev.previousElementSibling)) if (prev.tagName === e.tagName) n++;
      if (n > 1) sel += ':nth-of-type(' + n + ')';
      parts.unshift(sel);
      e = e.parentElement;
      if (parts.length > 6) break;
    }
    return parts.join(' > ');
  }
  var rc = el.getBoundingClientRect();
  return {
    tag: el.tagName.toLowerCase(),
    id: el.id || null,
    name: el.getAttribute('name') || null,
    placeholder: el.getAttribute('placeholder') || null,
    aria_label: el.getAttribute('aria-label') || null,
    aria_role: el.getAttribute('role') || null,
    href: el.getAttribute('href') || null,
    inner_text: (el.innerText || el.textContent || '').trim().substring(0, 200),
    xpath: xpath(el),
    css: cssPath(el),
    rect: {left: rc.left, top: rc.top, right: rc.right, bottom: rc.bottom},
    is_visible: (rc.width > 0 && rc.height > 0)
  };
})"""

# Returns up to 3 candidates scored by match quality
_JS_FIND_CANDIDATES = """\
(function(xp, css, aria, txt, tagHint) {
  var candidates = [];

  function tryEl(el, strategy, baseScore) {
    if (!el) return;
    el.scrollIntoView({behavior: 'instant', block: 'center'});
    var rc = el.getBoundingClientRect();
    var visible = rc.width > 0 && rc.height > 0;
    candidates.push({
      cx: Math.round(rc.left + rc.width/2),
      cy: Math.round(rc.top + rc.height/2),
      score: baseScore - (visible ? 0 : 20),
      strategy: strategy,
      tag: el.tagName.toLowerCase(),
      text: (el.innerText || el.textContent || '').trim().substring(0, 60),
      visible: visible
    });
  }

  function searchIn(root, depth) {
    if (depth > 4) return;
    // XPath
    if (xp) { try { var r = root.evaluate(xp, root, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null); tryEl(r.singleNodeValue, 'xpath', 90); } catch(e){} }
    // CSS
    if (css) { try { tryEl(root.querySelector(css), 'css', 85); } catch(e){} }
    // ARIA
    if (aria) { try { tryEl(root.querySelector('[aria-label="'+aria+'"]'), 'aria', 80); } catch(e){} }
    // Text match
    if (txt) {
      var tag = tagHint || '*';
      var selector = (tag === '*') ? 'button,a,td,th,span,div,input,select,label,li' : tag;
      try {
        var all = root.querySelectorAll(selector);
        for (var i = 0; i < all.length && candidates.length < 5; i++) {
          var t = (all[i].innerText || all[i].textContent || '').trim();
          if (t === txt) { tryEl(all[i], 'exact_text', 75); }
          else if (t.indexOf(txt) !== -1) { tryEl(all[i], 'partial_text', 55); }
        }
      } catch(e) {}
    }
    // Shadow DOM
    try {
      var hosts = root.querySelectorAll ? root.querySelectorAll('*') : [];
      for (var j = 0; j < hosts.length; j++) {
        if (hosts[j].shadowRoot) searchIn(hosts[j].shadowRoot, depth + 1);
      }
    } catch(e) {}
  }

  searchIn(document, 0);

  // Search iframes
  try {
    var frames = document.querySelectorAll('iframe');
    for (var f = 0; f < frames.length; f++) {
      try { searchIn(frames[f].contentDocument, 0); } catch(e) {}
    }
  } catch(e) {}

  // Sort by score desc, deduplicate by cx+cy
  candidates.sort(function(a, b) { return b.score - a.score; });
  var seen = {};
  var unique = [];
  for (var k = 0; k < candidates.length; k++) {
    var key = candidates[k].cx + ',' + candidates[k].cy;
    if (!seen[key]) { seen[key] = true; unique.push(candidates[k]); }
    if (unique.length >= 3) break;
  }
  return unique;
})"""

# React/Vue-compatible value setter
_JS_SET_VALUE = """\
(function(xp, css, val) {
  var el = null;
  if (xp){try{var r=document.evaluate(xp,document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null);el=r.singleNodeValue;}catch(e){}}
  if (!el&&css){try{el=document.querySelector(css);}catch(e){}}
  if (!el) return false;
  el.focus();
  var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value');
  if (setter&&setter.set){
    setter.set.call(el, val);
    el.dispatchEvent(new Event('input',{bubbles:true}));
    el.dispatchEvent(new Event('change',{bubbles:true}));
  } else {
    el.value = val;
    el.dispatchEvent(new Event('change',{bubbles:true}));
  }
  return true;
})"""

# DOM mutation observer — resolves when no mutations for stable_ms
_JS_WAIT_DOM_STABLE = """\
(function(stableMs) {
  return new Promise(function(resolve) {
    var timer = null;
    var reset = function() {
      if (timer) clearTimeout(timer);
      timer = setTimeout(function() { observer.disconnect(); resolve(true); }, stableMs);
    };
    var observer = new MutationObserver(reset);
    observer.observe(document.body || document.documentElement, {
      childList: true, subtree: true, attributes: true
    });
    reset();
    setTimeout(function() { observer.disconnect(); resolve(true); }, stableMs * 10);
  });
})"""

_JS_VALIDATE_ELEMENT = """\
(function(xp, css) {
  var el = null;
  if (xp){try{var r=document.evaluate(xp,document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null);el=r.singleNodeValue;}catch(e){}}
  if (!el&&css){try{el=document.querySelector(css);}catch(e){}}
  if (!el) return {found:false};
  var rc=el.getBoundingClientRect();
  return {found:true, visible:(rc.width>0&&rc.height>0), enabled:!el.disabled,
          tag:el.tagName.toLowerCase(),
          text:(el.innerText||el.textContent||'').trim().substring(0,50)};
})"""

_JS_READY_STATE = "(function(){return document.readyState})"


class BrowserBridge:

    _SKIP_URLS = (
        "chrome-devtools://", "devtools://", "chrome-extension://",
        "about:", "data:", "chrome://",
    )

    def __init__(self, cdp_port: int = 9222):
        self._port     = cdp_port
        self._ws       = None
        self._ws_thread: Optional[threading.Thread] = None
        self._msg_id   = 0
        self._id_lock  = threading.Lock()
        self._pending: dict[int, tuple[queue.Queue, threading.Event]] = {}
        self._pending_lock = threading.Lock()
        self._connected    = False
        self._tab_url      = ""
        self._tab_id       = ""
        self._tab_title    = ""
        self._vp_offset:   Optional[dict] = None
        self._vp_ts:       float = 0.0
        self._VP_CACHE_S   = 2.0

    # ──────────────────────────────────────────────────────────────────
    # Connection
    # ──────────────────────────────────────────────────────────────────

    def connect(self) -> bool:
        if not WS_AVAILABLE:
            logger.error("[BROWSER] Cannot connect — websocket-client not installed")
            return False
        self.disconnect()
        try:
            tabs = self._list_tabs()
            tab  = self._pick_tab(tabs)
            if not tab:
                logger.warning("[BROWSER] No usable tab on port {}", self._port)
                return False
            ok = self._open_ws(tab)
            if ok:
                logger.info("[BROWSER] Connected → tab='{}' url='{}'",
                            self._tab_title[:40], self._tab_url[:60])
            return ok
        except Exception as exc:
            logger.warning("[BROWSER] Connect failed: {}", exc)
            return False

    def disconnect(self) -> None:
        self._connected  = False
        self._vp_offset  = None
        if self._ws:
            try:
                self._ws.close()
            except Exception:
                pass
        self._ws = None

    @property
    def is_connected(self) -> bool:
        return self._connected

    # ──────────────────────────────────────────────────────────────────
    # Recording
    # ──────────────────────────────────────────────────────────────────

    def get_element_at(self, vx: int, vy: int) -> Optional[BrowserTarget]:
        if not self._connected:
            return None
        try:
            result = self._call_js(_JS_ELEMENT_AT, [vx, vy])
            if not isinstance(result, dict):
                logger.debug("[BROWSER] get_element_at ({},{}) → no element", vx, vy)
                return None
            bt = BrowserTarget(
                xpath        = result.get("xpath"),
                css_selector = result.get("css"),
                element_id   = result.get("id"),
                name_attr    = result.get("name"),
                placeholder  = result.get("placeholder"),
                aria_label   = result.get("aria_label"),
                aria_role    = result.get("aria_role"),
                href         = result.get("href"),
                inner_text   = (result.get("inner_text") or "")[:120] or None,
                tag_name     = result.get("tag"),
                tab_url      = self._tab_url,
            )
            logger.info(
                "[BROWSER] Captured element at ({},{}) tag={} xpath={} css={} text='{}'",
                vx, vy,
                result.get("tag","?"),
                (result.get("xpath") or "")[:60],
                (result.get("css") or "")[:40],
                (result.get("inner_text") or "")[:30],
            )
            return bt
        except Exception as exc:
            logger.debug("[BROWSER] get_element_at failed: {}", exc)
            return None

    def get_page_url(self) -> str:
        if not self._connected:
            return ""
        try:
            r   = self._send("Page.getNavigationHistory", {})
            idx = r.get("currentIndex", 0)
            url = r.get("entries", [{}])[idx].get("url", "")
            logger.debug("[BROWSER] Current URL: {}", url)
            return url
        except Exception:
            return ""

    def screen_to_viewport(self, sx: int, sy: int, window_rect: dict) -> tuple[int, int]:
        off = self._get_viewport_offset(window_rect)
        vx, vy = sx - off["x"], sy - off["y"]
        logger.debug("[BROWSER] screen({},{}) → viewport({},{}) offset=({},{})",
                     sx, sy, vx, vy, off["x"], off["y"])
        return vx, vy

    def get_tab_list(self) -> list[dict]:
        try:
            return self._list_tabs()
        except Exception:
            return []

    # ──────────────────────────────────────────────────────────────────
    # Replay — multi-candidate finding
    # ──────────────────────────────────────────────────────────────────

    def bring_to_front(self) -> None:
        if not self._connected:
            return
        try:
            self._send("Target.activateTarget", {"targetId": self._tab_id})
            time.sleep(0.12)
            logger.debug("[BROWSER] Brought tab to front: {}", self._tab_title[:30])
        except Exception:
            pass

    def is_tab_focused(self) -> bool:
        try:
            return bool(self._call_js("(function(){return document.hasFocus()})", []))
        except Exception:
            return False

    def find_candidates(
        self,
        bt: BrowserTarget,
        timeout_ms: int = 8000,
    ) -> list[BrowserCandidate]:
        """
        Find up to 3 scored candidates for this browser target.
        Retries until timeout. Returns empty list if nothing found.
        """
        if not self._connected:
            logger.warning("[BROWSER] find_candidates: not connected")
            return []

        deadline = time.time() + timeout_ms / 1000
        attempt  = 0
        while time.time() < deadline:
            attempt += 1
            try:
                results = self._call_js(
                    _JS_FIND_CANDIDATES,
                    [bt.xpath, bt.css_selector, bt.aria_label, bt.inner_text, bt.tag_name],
                )
                if isinstance(results, list) and results:
                    candidates = [
                        BrowserCandidate(
                            cx=r.get("cx", 0), cy=r.get("cy", 0),
                            score=r.get("score", 0), strategy=r.get("strategy","?"),
                            tag=r.get("tag",""), text=r.get("text",""),
                            visible=r.get("visible", True),
                        )
                        for r in results
                    ]
                    logger.info(
                        "[BROWSER] find_candidates attempt={} → {} matches: {}",
                        attempt,
                        len(candidates),
                        "; ".join(
                            f"strategy={c.strategy} score={c.score:.0f} "
                            f"pos=({c.cx},{c.cy}) visible={c.visible} text='{c.text[:20]}'"
                            for c in candidates
                        ),
                    )
                    return candidates
            except Exception as exc:
                logger.debug("[BROWSER] find_candidates attempt={} error: {}", attempt, exc)

            time.sleep(0.4)

        logger.warning(
            "[BROWSER] find_candidates: timed out after {}ms. "
            "xpath={} css={} aria={} text='{}'",
            timeout_ms,
            bt.xpath, bt.css_selector, bt.aria_label, bt.inner_text,
        )
        return []

    def find_element(
        self,
        bt: BrowserTarget,
        retry: int = 2,
        retry_delay_ms: int = 400,
    ) -> Optional[tuple[int, int]]:
        """Find best element. Returns (cx, cy) or None."""
        candidates = self.find_candidates(bt, timeout_ms=retry_delay_ms * (retry + 1) + 2000)
        if not candidates:
            return None
        best = candidates[0]
        if not best.visible:
            logger.warning("[BROWSER] Best candidate not visible (score={:.0f})", best.score)
        return (best.cx, best.cy)

    def verify_element(self, bt: BrowserTarget) -> dict:
        """Re-check element existence, visibility, enabled state, tag, and text."""
        if not self._connected:
            return {"found": False, "visible": False, "enabled": False}
        try:
            r = self._call_js(_JS_VALIDATE_ELEMENT, [bt.xpath, bt.css_selector])
            if isinstance(r, dict):
                logger.debug(
                    "[BROWSER] verify_element: found={} visible={} enabled={} tag={} text='{}'",
                    r.get("found"), r.get("visible"), r.get("enabled"),
                    r.get("tag"), r.get("text","")[:20],
                )
                return r
        except Exception as exc:
            logger.debug("[BROWSER] verify_element error: {}", exc)
        return {"found": False, "visible": False, "enabled": False}

    def wait_for_element(
        self,
        bt: BrowserTarget,
        timeout_ms: int = 10000,
        poll_ms: int = 300,
    ) -> bool:
        """Poll until element appears in DOM."""
        deadline = time.time() + timeout_ms / 1000
        while time.time() < deadline:
            state = self.verify_element(bt)
            if state.get("found") and state.get("visible"):
                logger.debug("[BROWSER] wait_for_element: element ready")
                return True
            time.sleep(poll_ms / 1000)
        logger.warning("[BROWSER] wait_for_element: timed out {}ms", timeout_ms)
        return False

    def wait_for_dom_stable(self, stable_ms: int = 300, max_wait_ms: int = 3000) -> None:
        """Wait until DOM has no mutations for stable_ms. Prevents click-on-moving-element."""
        if not self._connected:
            return
        try:
            self._call_js(_JS_WAIT_DOM_STABLE, [stable_ms], timeout=max_wait_ms/1000 + 1)
            logger.debug("[BROWSER] DOM stable after wait ({}ms threshold)", stable_ms)
        except Exception as exc:
            logger.debug("[BROWSER] wait_for_dom_stable: {}", exc)

    def click_at_viewport(self, vx: int, vy: int) -> None:
        self.bring_to_front()
        if not self.is_tab_focused():
            try:
                self._call_js("(function(){document.body.click();})", [])
                time.sleep(0.05)
            except Exception:
                pass
        logger.info("[BROWSER] Clicking at viewport ({},{})", vx, vy)
        for etype in ("mousePressed", "mouseReleased"):
            self._send("Input.dispatchMouseEvent", {
                "type": etype, "x": vx, "y": vy,
                "button": "left", "clickCount": 1, "modifiers": 0,
            })
            time.sleep(0.025)

    def type_text_at(self, text: str, human_like: bool = True) -> None:
        self.bring_to_front()
        logger.info("[BROWSER] Typing {} chars (human_like={})", len(text), human_like)
        for ch in text:
            self._send("Input.dispatchKeyEvent", {"type": "keyDown", "text": ch, "unmodifiedText": ch})
            self._send("Input.dispatchKeyEvent", {"type": "keyUp",   "text": ch, "unmodifiedText": ch})
            time.sleep(random.uniform(0.02, 0.06) if human_like else 0.008)

    def set_value(self, bt: BrowserTarget, value: str) -> bool:
        if not self._connected:
            return False
        try:
            ok = bool(self._call_js(_JS_SET_VALUE, [bt.xpath, bt.css_selector, value]))
            logger.info("[BROWSER] set_value result={} xpath={} css={}",
                        ok, bt.xpath, bt.css_selector)
            return ok
        except Exception as exc:
            logger.error("[BROWSER] set_value failed: {}", exc)
            return False

    def navigate(self, url: str, wait: bool = True, timeout_ms: int = 15000) -> None:
        if not self._connected:
            return
        logger.info("[BROWSER] Navigating to {}", url)
        self._send("Page.navigate", {"url": url})
        if wait:
            self.wait_for_load(timeout_ms)

    def wait_for_load(self, timeout_ms: int = 15000) -> bool:
        deadline = time.time() + timeout_ms / 1000
        while time.time() < deadline:
            try:
                state = self._call_js(_JS_READY_STATE, [])
                if state == "complete":
                    time.sleep(0.15)
                    logger.debug("[BROWSER] Page load complete")
                    return True
            except Exception:
                pass
            time.sleep(0.25)
        logger.warning("[BROWSER] wait_for_load: page not ready after {}ms", timeout_ms)
        return False

    def get_selected_text(self) -> str:
        try:
            return self._call_js("(function(){return window.getSelection().toString()})", []) or ""
        except Exception:
            return ""

    def switch_to_tab(self, tab_id: str) -> bool:
        try:
            tab = next((t for t in self._list_tabs() if t.get("id") == tab_id), None)
            if tab:
                self._vp_offset = None
                return self._open_ws(tab)
        except Exception as exc:
            logger.error("[BROWSER] switch_to_tab failed: {}", exc)
        return False

    # ──────────────────────────────────────────────────────────────────
    # WebSocket management
    # ──────────────────────────────────────────────────────────────────

    def _open_ws(self, tab: dict) -> bool:
        ws_url = tab.get("webSocketDebuggerUrl")
        if not ws_url:
            return False
        self._tab_url   = tab.get("url", "")
        self._tab_id    = tab.get("id", "")
        self._tab_title = tab.get("title", "")
        self._vp_offset = None
        self._pending.clear()

        def on_message(ws, raw):
            try:
                msg    = json.loads(raw)
                msg_id = msg.get("id")
                if msg_id is not None:
                    with self._pending_lock:
                        entry = self._pending.get(msg_id)
                    if entry:
                        q, ev = entry
                        q.put(msg.get("result", {}))
                        ev.set()
            except Exception:
                pass

        self._ws = websocket.WebSocketApp(
            ws_url,
            on_message=on_message,
            on_open=lambda ws: setattr(self, "_connected", True),
            on_close=lambda ws, c, m: setattr(self, "_connected", False),
            on_error=lambda ws, e: logger.debug("[BROWSER] WS error: {}", e),
        )
        self._ws_thread = threading.Thread(
            target=lambda: self._ws.run_forever(), daemon=True
        )
        self._ws_thread.start()
        deadline = time.time() + 3.0
        while not self._connected and time.time() < deadline:
            time.sleep(0.05)
        return self._connected

    def _send(self, method: str, params: dict, timeout: float = 8.0) -> dict:
        if not self._connected or not self._ws:
            raise RuntimeError("CDP not connected")
        with self._id_lock:
            self._msg_id += 1
            mid = self._msg_id
        q:  queue.Queue    = queue.Queue()
        ev: threading.Event = threading.Event()
        with self._pending_lock:
            self._pending[mid] = (q, ev)
        try:
            self._ws.send(json.dumps({"id": mid, "method": method, "params": params}))
            if not ev.wait(timeout=timeout):
                raise TimeoutError(f"CDP {method} timed out ({timeout}s)")
            return q.get_nowait()
        finally:
            with self._pending_lock:
                self._pending.pop(mid, None)

    def _call_js(self, fn_body: str, args: list, timeout: float = 8.0) -> Any:
        args_json = json.dumps(args)[1:-1]
        expr      = f"({fn_body})({args_json})"
        result    = self._send("Runtime.evaluate", {
            "expression": expr, "returnByValue": True, "awaitPromise": True,
        }, timeout=timeout)
        if result.get("exceptionDetails"):
            logger.debug("[BROWSER] JS exception: {}", result["exceptionDetails"].get("text","?"))
            return None
        return result.get("result", {}).get("value")

    def _list_tabs(self) -> list[dict]:
        with urlopen(f"http://localhost:{self._port}/json", timeout=3) as r:
            return json.loads(r.read())

    def _pick_tab(self, tabs: list[dict]) -> Optional[dict]:
        for tab in tabs:
            if tab.get("type") != "page":
                continue
            url = tab.get("url", "")
            if any(url.startswith(p) for p in self._SKIP_URLS):
                continue
            if url in ("", "about:blank"):
                continue
            return tab
        return next((t for t in tabs if t.get("type") == "page"), None)

    def _get_viewport_offset(self, window_rect: dict) -> dict:
        now = time.time()
        if self._vp_offset is None or now - self._vp_ts > self._VP_CACHE_S:
            self._vp_offset = self._compute_offset(window_rect)
            self._vp_ts     = now
        return self._vp_offset

    def _compute_offset(self, window_rect: dict) -> dict:
        chrome_h = 90
        try:
            m    = self._send("Page.getLayoutMetrics", {}, timeout=3)
            vp   = m.get("visualViewport") or m.get("layoutViewport", {})
            wh   = window_rect.get("height", 0)
            cssH = vp.get("clientHeight", 0)
            if wh > 0 and cssH > 0:
                chrome_h = wh - cssH
        except Exception:
            pass
        offset = {
            "x": window_rect.get("left", 0),
            "y": window_rect.get("top",  0) + chrome_h,
        }
        logger.debug("[BROWSER] Viewport offset computed: x={} y={} (chrome_h={})",
                     offset["x"], offset["y"], chrome_h)
        return offset
