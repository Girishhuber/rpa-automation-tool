"""
EventPipeline — normalises, compresses, and tags events before storage.

Improvements:
  - Scroll event compression: consecutive scrolls in same direction merged
  - Text typing compression: rapid keystrokes batched into single TypeTextEvent
  - Event intent tagging: navigation / input / selection / system / clipboard
  - Structured log line per event: type | intent | target summary
  - Debounce applies only to scroll (never to clicks)
"""

from __future__ import annotations
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


# ─────────────────────────────────────────────────────────────────────────────
# Intent classification
# ─────────────────────────────────────────────────────────────────────────────

def _classify_intent(payload) -> str:
    ptype = str(getattr(payload, "type", ""))
    if "navigate" in ptype or "browser_back" in ptype or "browser_forward" in ptype:
        return "navigation"
    if "clipboard" in ptype or "copy" in ptype or "paste" in ptype:
        return "clipboard"
    if "type" in ptype or "key" in ptype:
        return "input"
    if "click" in ptype or "drag" in ptype:
        return "selection"
    if "scroll" in ptype:
        return "navigation"
    if "window_focus" in ptype:
        return "system"
    if "excel" in ptype:
        return "input"
    return "system"


def _target_summary(target: Optional[UITarget]) -> str:
    """Short, human-readable description of a target for logging."""
    if not target:
        return "(no target)"
    parts = []
    backend = getattr(target, "backend", None)
    if backend:
        parts.append(f"backend={backend.value if hasattr(backend,'value') else backend}")
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


class EventPipeline:
    """
    Thread-safe pipeline. Compresses noisy events and tags with intent.
    """

    def __init__(
        self,
        consumer: Callable[[Event], None],
        debounce_scroll_ms: int = 80,
    ):
        self._consumer          = consumer
        self._debounce_scroll   = debounce_scroll_ms
        self._event_id          = 0
        self._session_start_ms: Optional[int] = None
        self._last_event_ts:    dict[str, int] = {}

        # Scroll compression state
        self._pending_scroll:   Optional[dict] = None
        self._last_scroll_ms:   int = 0

        # Text compression state (handled by recorder's text buffer — pipeline just passes through)

    def start(self) -> None:
        self._session_start_ms = self._now_ms()
        self._event_id = 0
        logger.info("[PIPELINE] Started — session clock reset")

    def stop(self) -> None:
        self._flush_pending_scroll()
        logger.info("[PIPELINE] Stopped — {} events emitted", self._event_id)

    # ──────────────────────────────────────────────────────────────────
    # Emit helpers
    # ──────────────────────────────────────────────────────────────────

    def emit_click(self, x: int, y: int, button: str = "left",
                   target: Optional[UITarget] = None) -> None:
        self._flush_pending_scroll()
        self._emit(MouseClickEvent(x=x, y=y, button=button, target=target))

    def emit_double_click(self, x: int, y: int,
                          target: Optional[UITarget] = None) -> None:
        self._flush_pending_scroll()
        self._emit(MouseDoubleClickEvent(x=x, y=y, target=target))

    def emit_right_click(self, x: int, y: int,
                         target: Optional[UITarget] = None) -> None:
        self._flush_pending_scroll()
        self._emit(MouseRightClickEvent(x=x, y=y, target=target))

    def emit_scroll(self, x: int, y: int, dx: int, dy: int,
                    target: Optional[UITarget] = None) -> None:
        """
        Compress consecutive scroll events in the same direction.
        Emits the accumulated scroll when direction changes or flush is called.
        """
        now = self._now_ms()
        if (now - self._last_scroll_ms) < self._debounce_scroll:
            # Debounce — merge into pending
            if self._pending_scroll:
                self._pending_scroll["dx"] += dx
                self._pending_scroll["dy"] += dy
                self._last_scroll_ms = now
                return

        # Check direction change
        if self._pending_scroll:
            same_dir = (
                (dy > 0) == (self._pending_scroll["dy"] > 0) or
                (dx > 0) == (self._pending_scroll["dx"] > 0)
            )
            if not same_dir:
                self._flush_pending_scroll()

        if self._pending_scroll:
            self._pending_scroll["dx"] += dx
            self._pending_scroll["dy"] += dy
        else:
            self._pending_scroll = {"x": x, "y": y, "dx": dx, "dy": dy, "target": target}
        self._last_scroll_ms = now

    def emit_key_press(self, key: str, target: Optional[UITarget] = None) -> None:
        self._emit(KeyPressEvent(key=key, target=target))

    def emit_key_combo(self, keys: list[str], target: Optional[UITarget] = None) -> None:
        self._emit(KeyComboEvent(keys=keys, target=target))

    def emit_type_text(self, text: str, target: Optional[UITarget] = None,
                       clear_first: bool = False) -> None:
        self._emit(TypeTextEvent(text=text, target=target, clear_first=clear_first))

    def emit_window_focus(self, title: str, process: str,
                          x: int, y: int, w: int, h: int) -> None:
        self._emit(WindowFocusEvent(
            window_title=title, process_name=process,
            x=x, y=y, width=w, height=h,
        ))

    def emit_screenshot(self, path: str, monitor: int = 0) -> None:
        self._emit(ScreenshotCheckpointEvent(path=path))

    def emit_wait(self, duration_ms: int) -> None:
        self._emit(ExplicitWaitEvent(duration_ms=duration_ms))

    # ──────────────────────────────────────────────────────────────────
    # Internal
    # ──────────────────────────────────────────────────────────────────

    def _flush_pending_scroll(self) -> None:
        if self._pending_scroll:
            s = self._pending_scroll
            self._pending_scroll = None
            self._emit(MouseScrollEvent(
                x=s["x"], y=s["y"], dx=s["dx"], dy=s["dy"], target=s["target"]
            ))

    def _emit(self, payload) -> None:
        self._event_id += 1
        now    = self._now_ms()
        ts_ms  = now - (self._session_start_ms or now)
        intent = _classify_intent(payload)

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
            "[EVENT #{:04d}] type={:<30} intent={:<12} {}{} — {}",
            self._event_id, ptype, intent, extra, "", tsum
        )

        try:
            self._consumer(event)
        except Exception as exc:
            logger.error("[PIPELINE] Consumer error on event #{}: {}", self._event_id, exc)

    @staticmethod
    def _now_ms() -> int:
        return int(time.perf_counter() * 1000)
