from __future__ import annotations

import queue
import threading
import time
from datetime import datetime, timezone
from typing import Callable, Optional

from models.event import (
    Event,
    MouseClickEvent, MouseDoubleClickEvent, MouseRightClickEvent,
    MouseScrollEvent, MouseDragEvent,
    KeyPressEvent, KeyComboEvent, TypeTextEvent, ClipboardPasteEvent,
    WindowFocusEvent, ScreenshotCheckpointEvent, ExplicitWaitEvent,
)
from models.target import UITarget
from utils.logger import logger

def classify_intent(payload) -> str:
    """Canonical intent classifier — import this everywhere instead of duplicating."""
    ptype = str(getattr(payload, "type", ""))
    if "navigate" in ptype or "browser_back" in ptype or "browser_forward" in ptype:
        return "navigation"
    if "process_launch" in ptype:
        return "navigation"
    if "clipboard" in ptype or "copy" in ptype or "paste" in ptype:
        return "clipboard"
    if "type" in ptype or "key" in ptype:
        return "input"
    if "click" in ptype or "drag" in ptype:
        return "selection"
    if "scroll" in ptype:
        return "scroll"        # was "navigation" — corrected
    if "window_focus" in ptype:
        return "system"
    if "excel" in ptype:
        return "input"
    if "screenshot" in ptype or "wait" in ptype:
        return "checkpoint"
    return "system"


_classify_intent = classify_intent


def _target_summary(target: Optional[UITarget]) -> str:
    """Short, human-readable description of a target for logging."""
    if not target:
        return "(no target)"
    parts = []
    backend = getattr(target, "backend", None)
    if backend:
        parts.append(f"backend={backend.value if hasattr(backend, 'value') else backend}")
    if target.process_name:
        parts.append(f"app={target.process_name}")
    if target.window_title:
        parts.append(f"window='{target.window_title[:30]}'")
    if target.control_type:
        parts.append(f"type={target.control_type}")
    if target.automation_id:
        parts.append(f"auto_id={target.automation_id}")
    elif target.name:
        parts.append(f"name='{target.name[:30]}'")
    if hasattr(target, "browser") and target.browser:
        bt = target.browser
        if bt.xpath:
            parts.append(f"xpath={bt.xpath[:60]}")
        elif bt.css_selector:
            parts.append(f"css={bt.css_selector[:40]}")
        elif bt.inner_text:
            parts.append(f"text='{bt.inner_text[:30]}'")
    return " | ".join(parts) if parts else "(unknown)"

_SENTINEL = object()   # signals worker to stop


class EventPipeline:
    
    QUEUE_MAX = 2_000
  
    _WORKER_POLL_S = 0.005   # 5 ms

    def __init__(
        self,
        consumer: Callable[[Event], None],
        debounce_scroll_ms: int = 80,
    ):
        self._consumer        = consumer
        self._debounce_scroll = debounce_scroll_ms
        self._event_id        = 0
        self._session_start_ms: Optional[int] = None

        # Bounded queue — drops oldest on overflow (back-pressure)
        self._q: queue.Queue = queue.Queue(maxsize=self.QUEUE_MAX)

        # Worker thread state
        self._worker_thread: Optional[threading.Thread] = None
        self._running = False

        # Overflow metrics
        self._overflow_drops: int = 0

        # Scroll accumulation (worker-thread only — no lock needed)
        self._pending_scroll: Optional[dict] = None
        self._last_scroll_ms: int = 0

    # ── Lifecycle ──────────────────────────────────────────────────────────

    def start(self) -> None:
        if self._running:
            return
        self._session_start_ms = self._now_ms()
        self._event_id = 0
        self._running = True
        self._worker_thread = threading.Thread(
            target=self._worker, daemon=True, name="EventPipelineWorker"
        )
        self._worker_thread.start()
        logger.info("[PIPELINE] Started — session clock reset")

    def stop(self) -> None:

        self._running = False
       
        try:
            self._q.put_nowait(_SENTINEL)
        except queue.Full:
            pass
        if self._worker_thread:
            self._worker_thread.join(timeout=5.0)
        if self._overflow_drops:
            logger.warning("[PIPELINE] {} events were dropped due to queue overflow during session.",
                           self._overflow_drops)
        logger.info("[PIPELINE] Stopped — {} events emitted", self._event_id)

    # ── Public emit API (called from hook threads) ─────────────────────────

    def emit_click(self, x: int, y: int, button: str = "left",
                   target: Optional[UITarget] = None) -> None:
        self._enqueue(("flush_scroll", None))
        self._enqueue(("emit", MouseClickEvent(x=x, y=y, button=button, target=target)))

    def emit_double_click(self, x: int, y: int,
                          target: Optional[UITarget] = None) -> None:
        self._enqueue(("flush_scroll", None))
        self._enqueue(("emit", MouseDoubleClickEvent(x=x, y=y, target=target)))

    def emit_right_click(self, x: int, y: int,
                         target: Optional[UITarget] = None) -> None:
        self._enqueue(("flush_scroll", None))
        self._enqueue(("emit", MouseRightClickEvent(x=x, y=y, target=target)))

    def emit_scroll(self, x: int, y: int, dx: int, dy: int,
                    target: Optional[UITarget] = None) -> None:
       
        self._enqueue(("scroll", {"x": x, "y": y, "dx": dx, "dy": dy, "target": target}))

    def emit_key_press(self, key: str, target: Optional[UITarget] = None) -> None:
        self._enqueue(("emit", KeyPressEvent(key=key, target=target)))

    def emit_key_combo(self, keys: list[str], target: Optional[UITarget] = None) -> None:
        self._enqueue(("emit", KeyComboEvent(keys=keys, target=target)))

    def emit_type_text(self, text: str, target: Optional[UITarget] = None,
                       clear_first: bool = False) -> None:
        self._enqueue(("emit", TypeTextEvent(text=text, target=target, clear_first=clear_first)))

    def emit_window_focus(self, title: str, process: str,
                          x: int, y: int, w: int, h: int) -> None:
        self._enqueue(("emit", WindowFocusEvent(
            window_title=title, process_name=process,
            x=x, y=y, width=w, height=h,
        )))

    def emit_screenshot(self, path: str, monitor: int = 0) -> None:
        self._enqueue(("emit", ScreenshotCheckpointEvent(path=path)))

    def emit_wait(self, duration_ms: int) -> None:
        self._enqueue(("emit", ExplicitWaitEvent(duration_ms=duration_ms)))

   
    def _enqueue(self, msg) -> None:
        try:
            self._q.put_nowait(msg)
        except queue.Full:
            # Back-pressure: drop the oldest message, log the overflow, and retry once.
            try:
                self._q.get_nowait()
                self._overflow_drops += 1
                logger.error(
                    "[PIPELINE] EVENT LOST - QUEUE OVERFLOW: {} events dropped so far "
                    "(QUEUE_MAX={}). Replay accuracy may be compromised.",
                    self._overflow_drops, self.QUEUE_MAX,
                )
            except queue.Empty:
                pass
            try:
                self._q.put_nowait(msg)
            except queue.Full:
                pass

  
    def _worker(self) -> None:
        """Single worker thread — processes messages in FIFO order."""
        while True:
            try:
                msg = self._q.get(timeout=self._WORKER_POLL_S)
            except queue.Empty:
                # Periodic auto-flush: emit accumulated scroll if idle
                self._maybe_autoflush_scroll()
                continue

            if msg is _SENTINEL:
                break

            kind, data = msg

            if kind == "emit":
                self._flush_pending_scroll()
                self._emit_direct(data)

            elif kind == "flush_scroll":
                self._flush_pending_scroll()

            elif kind == "scroll":
                self._accumulate_scroll(data)

        # Drain remaining items after sentinel
        while not self._q.empty():
            try:
                msg = self._q.get_nowait()
                if msg is _SENTINEL:
                    continue
                kind, data = msg
                if kind == "emit":
                    self._emit_direct(data)
                elif kind == "scroll":
                    self._accumulate_scroll(data)
            except queue.Empty:
                break

        # Final flush
        self._flush_pending_scroll()

    def _accumulate_scroll(self, s: dict) -> None:
        now = self._now_ms()
        dx, dy = s["dx"], s["dy"]

        # Determine dominant axis for this event
        new_axis = "y" if abs(dy) >= abs(dx) else "x"
        new_sign = (dy > 0) if new_axis == "y" else (dx > 0)

        if self._pending_scroll is None:
            # No pending — start fresh
            self._pending_scroll = dict(s)
            self._last_scroll_ms = now
            return

        p = self._pending_scroll
        elapsed = now - self._last_scroll_ms

        # Within debounce window → always merge (minor direction flips ignored)
        if elapsed < self._debounce_scroll:
            p["dx"] += dx
            p["dy"] += dy
            self._last_scroll_ms = now
            return

        # Outside debounce window — check axis and direction
        pending_axis = "y" if abs(p["dy"]) >= abs(p["dx"]) else "x"
        pending_sign = (p["dy"] > 0) if pending_axis == "y" else (p["dx"] > 0)

        if new_axis != pending_axis or new_sign != pending_sign:
            # Axis changed OR direction reversed → flush pending, start fresh
            self._flush_pending_scroll()
            self._pending_scroll = dict(s)
        else:
            # Same axis + same direction → accumulate
            p["dx"] += dx
            p["dy"] += dy

        self._last_scroll_ms = now

    def _maybe_autoflush_scroll(self) -> None:
        """Auto-flush pending scroll if it has been idle > 3× debounce window."""
        if self._pending_scroll is None:
            return
        elapsed = self._now_ms() - self._last_scroll_ms
        if elapsed > self._debounce_scroll * 3:
            self._flush_pending_scroll()

    def _flush_pending_scroll(self) -> None:
        if self._pending_scroll is None:
            return
        s = self._pending_scroll
        self._pending_scroll = None
        self._emit_direct(MouseScrollEvent(
            x=s["x"], y=s["y"], dx=s["dx"], dy=s["dy"], target=s["target"]
        ))

    def _emit_direct(self, payload) -> None:
        """Build Event, log it, call consumer. Runs only in worker thread."""
        self._event_id += 1
        now   = self._now_ms()
        ts_ms = now - (self._session_start_ms or now)
        intent = classify_intent(payload)

        event = Event(
            id=self._event_id,
            timestamp_ms=ts_ms,
            wall_time=datetime.now(timezone.utc).isoformat(),
            payload=payload,
        )

        # Structured log line
        target = getattr(payload, "target", None)
        tsum   = _target_summary(target) if target else ""
        ptype  = str(getattr(payload, "type", type(payload).__name__))
        extra  = ""
        if hasattr(payload, "text"):
            extra = f" text='{payload.text[:30]}'"
        elif hasattr(payload, "key"):
            extra = f" key={payload.key}"
        elif hasattr(payload, "keys"):
            extra = f" keys={'+'.join(payload.keys)}"
        elif hasattr(payload, "url"):
            extra = f" url={payload.url[:50]}"
        elif hasattr(payload, "x") and hasattr(payload, "y"):
            extra = f" pos=({payload.x},{payload.y})"

        logger.info(
            "[EVENT #{:04d}] type={:<30} intent={:<12}{} — {}",
            self._event_id, ptype, intent, extra, tsum,
        )

        try:
            self._consumer(event)
        except Exception as exc:
            logger.error("[PIPELINE] Consumer error on event #{}: {}", self._event_id, exc)

    @staticmethod
    def _now_ms() -> int:
        return int(time.perf_counter() * 1000)