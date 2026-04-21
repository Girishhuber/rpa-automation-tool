"""
Recording overlay — transparent always-on-top window that shows:
  1. Animated click rings at every recorded click (red during record, green during replay)
  2. Floating event log — last N events in bottom-right corner
  3. Status bar — top strip showing recording state, event count, elapsed time

Uses tkinter with transparency. Runs in its own thread.
Never steals focus or blocks input.
"""

from __future__ import annotations
import math
import threading
import time
import tkinter as tk
from datetime import datetime
from typing import Optional

from utils.logger import logger


class ClickRing:
    """One animated ring that expands and fades at a click position."""
    def __init__(self, canvas: tk.Canvas, x: int, y: int, color: str, size: int):
        self.canvas = canvas
        self.x = x
        self.y = y
        self.color = color
        self.max_size = size
        self.current_size = 4
        self.alpha = 1.0
        self.done = False

        r = self.current_size // 2
        self.oval = canvas.create_oval(
            x - r, y - r, x + r, y + r,
            outline=color, width=3, fill="",
        )
        # Inner dot
        self.dot = canvas.create_oval(
            x - 4, y - 4, x + 4, y + 4,
            fill=color, outline="",
        )

    def step(self) -> bool:
        """Advance animation one frame. Returns False when done."""
        self.current_size += 4
        self.alpha = max(0.0, 1.0 - self.current_size / self.max_size)

        if self.current_size >= self.max_size or self.alpha <= 0:
            self.canvas.delete(self.oval)
            self.canvas.delete(self.dot)
            self.done = True
            return False

        r = self.current_size // 2
        self.canvas.coords(self.oval, self.x - r, self.y - r, self.x + r, self.y + r)

        # Simulate transparency by changing color opacity through shade
        shade = int(255 * self.alpha)
        hex_color = self.color
        try:
            # Parse hex color and fade it
            r_val = int(hex_color[1:3], 16)
            g_val = int(hex_color[3:5], 16)
            b_val = int(hex_color[5:7], 16)
            faded = f"#{int(r_val*self.alpha+255*(1-self.alpha)):02x}{int(g_val*self.alpha+255*(1-self.alpha)):02x}{int(b_val*self.alpha+255*(1-self.alpha)):02x}"
            self.canvas.itemconfig(self.oval, outline=faded, width=max(1, int(3 * self.alpha)))
        except Exception:
            pass
        return True


class RecordingOverlay:
    """
    Full-screen transparent overlay window. Always on top, click-through.
    Shows click rings, event log, and status bar.
    """

    def __init__(self, config):
        self._config = config
        self._root: Optional[tk.Tk] = None
        self._canvas: Optional[tk.Canvas] = None
        self._status_label: Optional[tk.Label] = None
        self._log_frame: Optional[tk.Frame] = None
        self._rings: list[ClickRing] = []
        self._log_lines: list[str] = []
        self._log_labels: list[tk.Label] = []
        self._thread: Optional[threading.Thread] = None
        self._running = False
        self._lock = threading.Lock()
        self._anim_job = None

        # State
        self._is_recording = False
        self._is_replaying = False
        self._event_count = 0
        self._start_time: Optional[float] = None
        self._status_text = "Ready"
        # Use a very distinctive chroma key for transparency.
        # "white" sometimes appears as a gray sheet if transparency fails.
        self._transparent_key = "#ff00ff"

    # ──────────────────────────────────────────────────────────────────
    # Lifecycle
    # ──────────────────────────────────────────────────────────────────

    def start(self) -> None:
        if not self._config.overlay.enabled:
            return
        self._running = True
        self._thread = threading.Thread(target=self._run_tk, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._running = False
        if self._root:
            try:
                self._root.after(0, self._root.destroy)
            except Exception:
                pass

    # ──────────────────────────────────────────────────────────────────
    # Public API — called from recorder/replayer threads
    # ──────────────────────────────────────────────────────────────────

    def set_recording(self, recording: bool) -> None:
        self._is_recording = recording
        self._is_replaying = False
        if recording:
            self._start_time = time.perf_counter()
            self._event_count = 0
        self._schedule(self._refresh_status)

    def set_replaying(self, replaying: bool) -> None:
        self._is_replaying = replaying
        self._is_recording = False
        if replaying:
            self._start_time = time.perf_counter()
        self._schedule(self._refresh_status)

    def flash_click(self, x: int, y: int, is_replay: bool = False) -> None:
        """Show an animated ring at screen position (x, y)."""
        color = (
            self._config.overlay.replay_color if is_replay
            else self._config.overlay.click_color
        )
        size = self._config.overlay.ring_size
        self._schedule(lambda: self._add_ring(x, y, color, size))

    def log_event(self, text: str) -> None:
        """Append a line to the floating event log."""
        self._event_count += 1
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"  {ts}  {text}"
        with self._lock:
            self._log_lines.append(line)
            if len(self._log_lines) > self._config.overlay.log_max_lines:
                self._log_lines.pop(0)
        self._schedule(self._refresh_log)
        self._schedule(self._refresh_status)

    def set_status(self, text: str) -> None:
        self._status_text = text
        self._schedule(self._refresh_status)

    # ──────────────────────────────────────────────────────────────────
    # Tkinter thread
    # ──────────────────────────────────────────────────────────────────

    def _run_tk(self) -> None:
        try:
            self._root = tk.Tk()
            self._root.title("RPA Overlay")
            self._root.withdraw()

            # Full-screen transparent window
            sw = self._root.winfo_screenwidth()
            sh = self._root.winfo_screenheight()
            self._root.geometry(f"{sw}x{sh}+0+0")
            self._root.overrideredirect(True)       # no title bar
            self._root.wm_attributes("-topmost", True)
            try:
                self._root.wm_attributes("-transparentcolor", self._transparent_key)
            except Exception:
                # If transparentcolor isn't supported, do not show the overlay
                # (better than blocking user input with a gray sheet).
                self._running = False
                return
            self._root.configure(bg=self._transparent_key)

            # Canvas covers full screen (transparent key color = transparent)
            self._canvas = tk.Canvas(
                self._root, width=sw, height=sh,
                bg=self._transparent_key, highlightthickness=0,
            )
            self._canvas.place(x=0, y=0)

            # Make window click-through on Windows
            try:
                import ctypes
                hwnd = self._root.winfo_id()

                GWL_EXSTYLE = -20
                WS_EX_LAYERED = 0x00080000
                WS_EX_TRANSPARENT = 0x00000020
                WS_EX_TOOLWINDOW = 0x00000080
                WS_EX_NOACTIVATE = 0x08000000

                user32 = ctypes.windll.user32

                # Prefer Ptr variants on 64-bit Python
                GetWindowLongPtrW = getattr(user32, "GetWindowLongPtrW", user32.GetWindowLongW)
                SetWindowLongPtrW = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)

                style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
                new_style = style | WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
                SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_style)

                # Ensure the window is shown without activation
                SW_SHOWNOACTIVATE = 4
                user32.ShowWindow(hwnd, SW_SHOWNOACTIVATE)
            except Exception:
                # If we can't guarantee click-through, disable overlay entirely.
                self._running = False
                return

            # Status bar (top-left)
            self._status_label = tk.Label(
                self._root,
                text="",
                bg="#1A1A2E", fg="#E0E0E0",
                font=("Segoe UI", 11, "bold"),
                padx=14, pady=6,
                anchor="w",
            )
            self._status_label.place(x=0, y=0)

            # Event log (bottom-right)
            self._log_frame = tk.Frame(self._root, bg="#1A1A2E", bd=0)
            self._log_frame.place(
                relx=1.0, rely=1.0,
                anchor="se",
                x=-12, y=-12,
            )

            self._root.deiconify()
            self._start_animation_loop()
            self._root.mainloop()
        except Exception as exc:
            logger.error("Overlay error: {}", exc)

    def _start_animation_loop(self) -> None:
        self._animate()

    def _animate(self) -> None:
        if not self._running or not self._root:
            return
        # Step all active rings
        dead = []
        for ring in self._rings:
            if not ring.step():
                dead.append(ring)
        for r in dead:
            self._rings.remove(r)

        # Update status elapsed
        if self._is_recording or self._is_replaying:
            self._refresh_status()

        self._anim_job = self._root.after(40, self._animate)   # ~25 fps

    def _add_ring(self, x: int, y: int, color: str, size: int) -> None:
        if self._canvas:
            ring = ClickRing(self._canvas, x, y, color, size)
            self._rings.append(ring)

    def _refresh_status(self) -> None:
        if not self._status_label:
            return
        elapsed = ""
        if self._start_time and (self._is_recording or self._is_replaying):
            secs = int(time.perf_counter() - self._start_time)
            elapsed = f"  {secs//60:02d}:{secs%60:02d}"

        if self._is_recording:
            icon = "⏺"
            color = self._config.overlay.recording_color
            state = f"RECORDING{elapsed}  |  {self._event_count} events"
            bg = "#8B0000"
        elif self._is_replaying:
            icon = "▶"
            color = self._config.overlay.replay_color
            state = f"REPLAYING{elapsed}  |  {self._event_count} events"
            bg = "#004400"
        else:
            icon = "●"
            state = self._status_text
            bg = "#1A1A2E"

        self._status_label.config(
            text=f"  {icon}  {state}  ",
            bg=bg,
        )

    def _refresh_log(self) -> None:
        if not self._log_frame:
            return
        # Destroy old labels
        for lbl in self._log_labels:
            lbl.destroy()
        self._log_labels.clear()

        if not self._config.overlay.show_event_log:
            return

        with self._lock:
            lines = list(self._log_lines)

        for line in lines:
            lbl = tk.Label(
                self._log_frame,
                text=line,
                bg="#1A1A2E",
                fg="#A8D8A8",
                font=("Consolas", 9),
                anchor="w",
                padx=4, pady=1,
            )
            lbl.pack(fill="x", side="bottom")
            self._log_labels.append(lbl)

    def _schedule(self, fn) -> None:
        """Schedule a function to run on the tkinter thread."""
        if self._root and self._running:
            try:
                self._root.after(0, fn)
            except Exception:
                pass