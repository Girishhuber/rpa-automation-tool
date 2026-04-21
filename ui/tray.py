from __future__ import annotations
import threading
import subprocess
import sys
from pathlib import Path
from typing import Optional

from utils.logger import logger
from utils.config import Config

try:
    import pystray
    from pystray import MenuItem as item, Menu
    from PIL import Image, ImageDraw
    PYSTRAY_AVAILABLE = True
except ImportError:
    PYSTRAY_AVAILABLE = False
    logger.warning("pystray not installed — tray UI unavailable")


def _make_icon(recording: bool = False) -> "Image.Image":
    """Generate a simple tray icon programmatically (no .ico file needed)."""
    from PIL import Image, ImageDraw
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    color = (220, 50, 50) if recording else (60, 130, 200)
    draw.ellipse([8, 8, size - 8, size - 8], fill=color)
    if recording:
        # Red dot with inner white square = recording indicator
        draw.rectangle([22, 22, size - 22, size - 22], fill="white")
    return img


class TrayApp:
    def __init__(
        self,
        config: Config,
        on_start_recording,
        on_stop_recording,
        on_replay_last,
        on_replay_picker,
        on_quit,
    ):
        self._config = config
        self._on_start_recording = on_start_recording
        self._on_stop_recording = on_stop_recording
        self._on_replay_last = on_replay_last
        self._on_replay_picker = on_replay_picker
        self._on_quit = on_quit

        self._is_recording = False
        self._status_text = "Ready"
        self._icon: Optional[pystray.Icon] = None

    def set_recording(self, recording: bool, status: str = "") -> None:
        self._is_recording = recording
        self._status_text = status or ("Recording..." if recording else "Ready")
        if self._icon:
            self._icon.icon = _make_icon(recording)
            self._icon.update_menu()

    def set_status(self, text: str) -> None:
        self._status_text = text
        if self._icon:
            self._icon.update_menu()

    def run(self) -> None:
        if not PYSTRAY_AVAILABLE:
            logger.error("pystray not available — cannot show tray icon")
            return

        self._icon = pystray.Icon(
            "rpa_tool",
            icon=_make_icon(False),
            title="RPA Recorder",
            menu=self._build_menu(),
        )
        logger.info("Tray icon started")
        self._icon.run()

    def stop(self) -> None:
        if self._icon:
            self._icon.stop()

    def _build_menu(self):
        return Menu(
            item(lambda _: self._status_text, lambda: None, enabled=False),
            Menu.SEPARATOR,
            item(
                "Start Recording",
                self._start_recording_action,
                enabled=lambda _: not self._is_recording,
            ),
            item(
                "Stop Recording",
                self._stop_recording_action,
                enabled=lambda _: self._is_recording,
            ),
            Menu.SEPARATOR,
            item("Replay Last Session", self._replay_last_action),
            item("Replay Session...", self._replay_picker_action),
            Menu.SEPARATOR,
            item("Open Log Folder", self._open_logs),
            Menu.SEPARATOR,
            item("Quit", self._quit_action),
        )


    def _start_recording_action(self, icon, query) -> None:
        threading.Thread(target=self._on_start_recording, daemon=True).start()

    def _stop_recording_action(self, icon, query) -> None:
        threading.Thread(target=self._on_stop_recording, daemon=True).start()

    def _replay_last_action(self, icon, query) -> None:
        threading.Thread(target=self._on_replay_last, daemon=True).start()

    def _replay_picker_action(self, icon, query) -> None:
        threading.Thread(target=self._on_replay_picker, daemon=True).start()

    def _open_logs(self, icon, query) -> None:
        log_dir = Path(self._config.storage.logs_dir)
        log_dir.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["explorer", str(log_dir)])

    def _quit_action(self, icon, query) -> None:
        self._on_quit()
        icon.stop()
