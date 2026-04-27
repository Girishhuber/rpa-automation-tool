
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


@dataclass
class BrowserCandidate:
    cx:       int
    cy:       int
    score:    float
    strategy: str
    tag:      str  = ""
    text:     str  = ""
    visible:  bool = True


BROWSER_MIN_ACCEPT_SCORE = 70.0
BROWSER_COORD_MIN_SCORE  = 70.0


# ─────────────────────────────────────────────────────────────────────────────
# JS snippets — ALL return plain values, no Promises
# (awaitPromise=False used everywhere except _JS_WAIT_DOM_STABLE)
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
    return '//' + e.tagName.toLowerCase();
  }
  function cssPath(e) {
    var parts = [];
    while (e && e.nodeType === 1) {
      var sel = e.tagName.toLowerCase();
      if (e.id) { parts.unshift('#' + e.id); break; }
      var cls = Array.from(e.classList||[])
        .filter(function(c){return !/[0-9]{3,}/.test(c) && c.length < 30;})
        .slice(0,2).join('.');
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


_JS_FIND_CANDIDATES = """\
(function(xp, css, aria, txt, tagHint) {
  var candidates = [];

  function tryEl(el, strategy, baseScore) {
    if (!el || !el.getBoundingClientRect) return;
    el.scrollIntoView({behavior: 'instant', block: 'center'});
    var rc = el.getBoundingClientRect();
    var visible = rc.width > 0 && rc.height > 0;
    candidates.push({
      cx: Math.round(rc.left + rc.width / 2),
      cy: Math.round(rc.top  + rc.height / 2),
      score: baseScore - (visible ? 0 : 20),
      strategy: strategy,
      tag: el.tagName.toLowerCase(),
      text: (el.innerText || el.textContent || '').trim().substring(0, 60),
      visible: visible
    });
  }

  // BB-2: always use document.evaluate for XPath — XPathResult only on document
  if (xp) {
    try {
      var r = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
      tryEl(r.singleNodeValue, 'xpath', 90);
    } catch(e) {}
  }

  function searchIn(root) {
    if (css) { try { tryEl(root.querySelector(css), 'css', 85); } catch(e) {} }
    if (aria) { try { tryEl(root.querySelector('[aria-label="' + aria + '"]'), 'aria', 80); } catch(e) {} }
    if (txt) {
      var selector = (tagHint && tagHint !== '*')
        ? tagHint
        : 'button,a,td,th,span,div,input,select,label,li,option';
      try {
        var all = root.querySelectorAll(selector);
        for (var i = 0; i < all.length && candidates.length < 8; i++) {
          var t = (all[i].innerText || all[i].textContent || '').trim();
          if (t === txt)            { tryEl(all[i], 'exact_text',   75); }
          else if (t.indexOf(txt) !== -1 && t.length < txt.length * 3) {
                                      tryEl(all[i], 'partial_text', 55); }
        }
      } catch(e) {}
    }
    // Shadow DOM
    try {
      var hosts = root.querySelectorAll('*');
      for (var j = 0; j < hosts.length; j++) {
        if (hosts[j].shadowRoot) searchIn(hosts[j].shadowRoot);
      }
    } catch(e) {}
  }

  searchIn(document);

  // Iframes
  try {
    var frames = document.querySelectorAll('iframe');
    for (var f = 0; f < frames.length; f++) {
      try { searchIn(frames[f].contentDocument || frames[f].contentWindow.document); } catch(e) {}
    }
  } catch(e) {}

  // Sort descending score, deduplicate by position
  candidates.sort(function(a, b) { return b.score - a.score; });
  var seen = {}, unique = [];
  for (var k = 0; k < candidates.length; k++) {
    var key = candidates[k].cx + ',' + candidates[k].cy;
    if (!seen[key]) { seen[key] = true; unique.push(candidates[k]); }
    if (unique.length >= 3) break;
  }
  return unique;
})"""

_JS_SET_VALUE = """\
(function(xp, css, val) {
  var el = null;
  if (xp) { try { var r = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null); el = r.singleNodeValue; } catch(e) {} }
  if (!el && css) { try { el = document.querySelector(css); } catch(e) {} }
  if (!el) return false;
  el.focus();
  var nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
  if (nativeSetter && nativeSetter.set) {
    nativeSetter.set.call(el, val);
  } else {
    el.value = val;
  }
  el.dispatchEvent(new Event('input',  {bubbles: true}));
  el.dispatchEvent(new Event('change', {bubbles: true}));
  return true;
})"""

_JS_VALIDATE_ELEMENT = """\
(function(xp, css) {
  var el = null;
  if (xp) { try { var r = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null); el = r.singleNodeValue; } catch(e) {} }
  if (!el && css) { try { el = document.querySelector(css); } catch(e) {} }
  if (!el) return {found: false, visible: false, enabled: false, tag: '', text: ''};
  var rc = el.getBoundingClientRect();
  return {
    found: true,
    visible: rc.width > 0 && rc.height > 0,
    enabled: !el.disabled,
    tag: el.tagName.toLowerCase(),
    text: (el.innerText || el.textContent || '').trim().substring(0, 50)
  };
})"""

# BB-5 FIX: DOM stable check returns plain boolean — no Promise, awaitPromise=False.
# Checks if page has been mutation-free for at least stable_ms by sampling.
_JS_DOM_MUTATION_COUNT = """\
(function() {
  if (!window.__rpaMutationCount) window.__rpaMutationCount = 0;
  if (!window.__rpaMutationObs) {
    window.__rpaMutationObs = new MutationObserver(function() {
      window.__rpaMutationCount++;
    });
    window.__rpaMutationObs.observe(document.body || document.documentElement,
      {childList: true, subtree: true, attributes: true, characterData: true});
  }
  return window.__rpaMutationCount;
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

        # BB-4 FIX: use Event for reliable open signal
        self._open_event = threading.Event()
        self._connected  = False

        self._tab_url   = ""
        self._tab_id    = ""
        self._tab_title = ""
        self._vp_offset: Optional[dict] = None
        self._vp_ts:     float = 0.0
        self._VP_CACHE_S = 2.0

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
        self._connected = False
        self._open_event.clear()
        self._vp_offset = None
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
        """Capture element identity at viewport coords. awaitPromise=False (BB-1)."""
        if not self._connected:
            return None
        try:
            result = self._call_js(_JS_ELEMENT_AT, [vx, vy], await_promise=False)
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
                "[BROWSER] Captured element @ ({},{}) tag={} id={} xpath='{}' css='{}' text='{}'",
                vx, vy,
                result.get("tag", "?"),
                result.get("id", ""),
                (result.get("xpath") or "")[:70],
                (result.get("css")   or "")[:50],
                (result.get("inner_text") or "")[:40],
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
        logger.debug("[BROWSER] screen({},{}) → viewport({},{}) chrome_offset=({},{})",
                     sx, sy, vx, vy, off["x"], off["y"])
        return vx, vy

    def get_tab_list(self) -> list[dict]:
        try:
            return self._list_tabs()
        except Exception:
            return []

    def get_selected_text(self) -> str:
        try:
            return self._call_js(
                "(function(){return window.getSelection().toString()})", [],
                await_promise=False
            ) or ""
        except Exception:
            return ""

    # ──────────────────────────────────────────────────────────────────
    # Replay
    # ──────────────────────────────────────────────────────────────────

    def bring_to_front(self) -> None:
        if not self._connected:
            return
        try:
            self._send("Target.activateTarget", {"targetId": self._tab_id})
            time.sleep(0.10)
            logger.debug("[BROWSER] Tab brought to front: {}", self._tab_title[:30])
        except Exception:
            pass

    def is_tab_focused(self) -> bool:
        try:
            return bool(self._call_js(
                "(function(){return document.hasFocus()})", [],
                await_promise=False
            ))
        except Exception:
            return False

    def find_candidates(
        self,
        bt: BrowserTarget,
        timeout_ms: int = 8000,
    ) -> list[BrowserCandidate]:
        """
        Find up to 3 scored candidates. Retries until timeout.
        BB-1: uses await_promise=False.
        BB-2: XPath always via document.evaluate in JS.
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
                    await_promise=False,  # BB-1: plain return value
                )
                if isinstance(results, list) and results:
                    candidates = [
                        BrowserCandidate(
                            cx       = r.get("cx", 0),
                            cy       = r.get("cy", 0),
                            score    = r.get("score", 0),
                            strategy = r.get("strategy", "?"),
                            tag      = r.get("tag", ""),
                            text     = r.get("text", ""),
                            visible  = r.get("visible", True),
                        )
                        for r in results
                    ]
                    logger.info(
                        "[BROWSER] find_candidates attempt={} → {} match(es): {}",
                        attempt,
                        len(candidates),
                        " | ".join(
                            f"strategy={c.strategy} score={c.score:.0f} "
                            f"pos=({c.cx},{c.cy}) vis={c.visible} text='{c.text[:25]}'"
                            for c in candidates
                        ),
                    )
                    return candidates
            except Exception as exc:
                logger.debug("[BROWSER] find_candidates attempt={} error: {}", attempt, exc)

            time.sleep(0.35)

        logger.warning(
            "[BROWSER] find_candidates TIMEOUT {}ms — xpath={} css={} aria={} text='{}'",
            timeout_ms,
            bt.xpath, bt.css_selector, bt.aria_label, (bt.inner_text or "")[:40],
        )
        return []

    def find_element(
        self,
        bt: BrowserTarget,
        retry: int = 2,
        retry_delay_ms: int = 400,
    ) -> Optional[tuple[int, int]]:
        """Find best element. Returns (cx, cy) in viewport coords or None."""
        candidates = self.find_candidates(
            bt, timeout_ms=retry_delay_ms * (retry + 1) + 2000
        )
        if not candidates:
            return None
        viable = [
            c for c in candidates
            if c.visible
            and c.score >= BROWSER_MIN_ACCEPT_SCORE
            and not (c.strategy == "coordinate" and c.score < BROWSER_COORD_MIN_SCORE)
        ]
        if not viable:
            best = candidates[0]
            logger.warning(
                "[BROWSER] Rejecting weak candidate strategy={} score={:.0f} visible={}",
                best.strategy, best.score, best.visible,
            )
            return None
        best = viable[0]
        return (best.cx, best.cy)


    def verify_element(self, bt: BrowserTarget) -> dict:
        """Re-check element existence + visibility before clicking."""
        if not self._connected:
            return {"found": False, "visible": False, "enabled": False}
        try:
            r = self._call_js(
                _JS_VALIDATE_ELEMENT, [bt.xpath, bt.css_selector],
                await_promise=False
            )
            if isinstance(r, dict):
                logger.debug(
                    "[BROWSER] verify: found={} visible={} enabled={} tag={} text='{}'",
                    r.get("found"), r.get("visible"), r.get("enabled"),
                    r.get("tag"), r.get("text", "")[:25],
                )
                return r
        except Exception as exc:
            logger.debug("[BROWSER] verify_element error: {}", exc)
        return {"found": False, "visible": False, "enabled": False}

    def wait_for_element(
        self,
        bt: BrowserTarget,
        timeout_ms: int = 10000,
        poll_ms: int = 250,
    ) -> bool:
        """Poll until element is found + visible."""
        deadline = time.time() + timeout_ms / 1000
        while time.time() < deadline:
            state = self.verify_element(bt)
            if state.get("found") and state.get("visible"):
                logger.debug("[BROWSER] wait_for_element: ready")
                return True
            time.sleep(poll_ms / 1000)
        logger.warning("[BROWSER] wait_for_element: timed out {}ms", timeout_ms)
        return False

    def wait_for_dom_stable(self, stable_ms: int = 300, max_wait_ms: int = 3000) -> None:
        """
        BB-5 FIX: Non-blocking DOM stability check using mutation count polling.
        Installs a MutationObserver that increments a counter on every mutation.
        We sample the counter twice with stable_ms gap — if unchanged, DOM is stable.
        Never uses awaitPromise=True so it doesn't block the CDP send thread.
        Fast path: skip wait entirely if readyState is already 'complete'.
        """
        if not self._connected:
            return
        try:
            # Fast path: page already loaded and settled
            state = self._call_js(_JS_READY_STATE, [], await_promise=False)
            if state != "complete":
                time.sleep(min(stable_ms / 1000, 0.5))
                logger.debug("[BROWSER] wait_for_dom_stable: page loading, waited {}ms", stable_ms)
                return

            # Install observer and sample mutation count
            count1 = self._call_js(_JS_DOM_MUTATION_COUNT, [], await_promise=False) or 0
            time.sleep(stable_ms / 1000)
            count2 = self._call_js(_JS_DOM_MUTATION_COUNT, [], await_promise=False) or 0

            if count1 == count2:
                logger.debug("[BROWSER] DOM stable (mutation_count={})", count1)
                return

            # Still mutating — wait again up to max_wait_ms total
            waited = stable_ms
            while waited < max_wait_ms:
                time.sleep(stable_ms / 1000)
                waited += stable_ms
                count3 = self._call_js(_JS_DOM_MUTATION_COUNT, [], await_promise=False) or 0
                if count3 == count2:
                    logger.debug("[BROWSER] DOM stable after {}ms (mutations={})", waited, count3)
                    return
                count2 = count3

            logger.debug("[BROWSER] DOM stable wait exceeded {}ms — proceeding anyway", max_wait_ms)
        except Exception as exc:
            logger.debug("[BROWSER] wait_for_dom_stable error: {}", exc)

    def click_at_viewport(self, vx: int, vy: int) -> None:
        """
        Click at viewport coordinates via CDP Input events.
        BB-3 FIX: removed document.body.click() focus fallback — it triggered
        unwanted click events. Use window.focus() instead.
        """
        self.bring_to_front()
        if not self.is_tab_focused():
            try:
                # BB-3: use window.focus() — no spurious click events
                self._call_js("(function(){window.focus();})", [], await_promise=False)
                time.sleep(0.04)
            except Exception:
                pass
        logger.info("[BROWSER] Click at viewport ({},{})", vx, vy)
        for etype in ("mousePressed", "mouseReleased"):
            self._send("Input.dispatchMouseEvent", {
                "type": etype, "x": vx, "y": vy,
                "button": "left", "clickCount": 1, "modifiers": 0,
            })
            time.sleep(0.025)

    def type_text_at(self, text: str, human_like: bool = True) -> None:
        self.bring_to_front()
        logger.info("[BROWSER] Typing {} chars human_like={}", len(text), human_like)
        for ch in text:
            self._send("Input.dispatchKeyEvent",
                       {"type": "keyDown", "text": ch, "unmodifiedText": ch})
            self._send("Input.dispatchKeyEvent",
                       {"type": "keyUp",   "text": ch, "unmodifiedText": ch})
            time.sleep(random.uniform(0.02, 0.06) if human_like else 0.008)

    def set_value(self, bt: BrowserTarget, value: str) -> bool:
        if not self._connected:
            return False
        try:
            ok = bool(self._call_js(_JS_SET_VALUE, [bt.xpath, bt.css_selector, value],
                                    await_promise=False))
            logger.info("[BROWSER] set_value result={} xpath='{}' css='{}'",
                        ok, (bt.xpath or "")[:60], (bt.css_selector or "")[:40])
            return ok
        except Exception as exc:
            logger.error("[BROWSER] set_value failed: {}", exc)
            return False

    def navigate(self, url: str, wait: bool = True, timeout_ms: int = 15000) -> None:
        if not self._connected:
            return
        logger.info("[BROWSER] Navigate → {}", url)
        self._send("Page.navigate", {"url": url})
        if wait:
            self.wait_for_load(timeout_ms)

    def wait_for_load(self, timeout_ms: int = 15000) -> bool:
        deadline = time.time() + timeout_ms / 1000
        while time.time() < deadline:
            try:
                state = self._call_js(_JS_READY_STATE, [], await_promise=False)
                if state == "complete":
                    time.sleep(0.15)
                    logger.debug("[BROWSER] Page load complete")
                    return True
            except Exception:
                pass
            time.sleep(0.25)
        logger.warning("[BROWSER] wait_for_load: not ready after {}ms", timeout_ms)
        return False

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
    # WebSocket (BB-4 fix: threading.Event for open signal)
    # ──────────────────────────────────────────────────────────────────

    def _open_ws(self, tab: dict) -> bool:
        ws_url = tab.get("webSocketDebuggerUrl")
        if not ws_url:
            logger.warning("[BROWSER] Tab has no webSocketDebuggerUrl")
            return False
        self._tab_url   = tab.get("url", "")
        self._tab_id    = tab.get("id", "")
        self._tab_title = tab.get("title", "")
        self._vp_offset = None
        self._pending.clear()
        self._open_event.clear()
        self._connected = False

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

        def on_open(ws):
            # BB-4 FIX: use Event, not lambda-setattr
            self._connected = True
            self._open_event.set()

        def on_close(ws, code, msg):
            self._connected = False

        def on_error(ws, err):
            logger.debug("[BROWSER] WS error: {}", err)

        self._ws = websocket.WebSocketApp(
            ws_url,
            on_message=on_message,
            on_open=on_open,
            on_close=on_close,
            on_error=on_error,
        )
        self._ws_thread = threading.Thread(
            target=lambda: self._ws.run_forever(), daemon=True
        )
        self._ws_thread.start()

        # BB-4: wait on Event, not busy-poll
        connected = self._open_event.wait(timeout=4.0)
        if not connected:
            logger.warning("[BROWSER] WS open timed out for {}", ws_url[:60])
        return connected

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

    def _call_js(
        self,
        fn_body: str,
        args: list,
        await_promise: bool = False,   # BB-1: caller decides, default False
        timeout: float = 8.0,
    ) -> Any:
        """
        Execute a JS function expression with positional args.
        BB-1: awaitPromise is now a parameter, defaulting to False.
              Only callers that use Promise-returning JS pass await_promise=True.
        """
        args_json = json.dumps(args)[1:-1]
        expr      = f"({fn_body})({args_json})"
        result    = self._send("Runtime.evaluate", {
            "expression":    expr,
            "returnByValue": True,
            "awaitPromise":  await_promise,
        }, timeout=timeout)
        if result.get("exceptionDetails"):
            detail = result["exceptionDetails"].get("text", "?")
            logger.debug("[BROWSER] JS exception: {}", detail)
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
        # fallback: any page tab
        return next((t for t in tabs if t.get("type") == "page"), None)

    def _get_viewport_offset(self, window_rect: dict) -> dict:
        now = time.time()
        if self._vp_offset is None or now - self._vp_ts > self._VP_CACHE_S:
            self._vp_offset = self._compute_offset(window_rect)
            self._vp_ts     = now
        return self._vp_offset

    def _compute_offset(self, window_rect: dict) -> dict:
        chrome_h = 90   # safe fallback
        try:
            m    = self._send("Page.getLayoutMetrics", {}, timeout=3)
            vp   = m.get("visualViewport") or m.get("layoutViewport", {})
            wh   = window_rect.get("height", 0)
            cssH = vp.get("clientHeight", 0)
            if wh > 0 and cssH > 0:
                chrome_h = max(0, wh - cssH)
        except Exception:
            pass
        offset = {
            "x": window_rect.get("left", 0),
            "y": window_rect.get("top",  0) + chrome_h,
        }
        logger.debug("[BROWSER] Viewport offset: x={} y={} chrome_h={}",
                     offset["x"], offset["y"], chrome_h)
        return offset