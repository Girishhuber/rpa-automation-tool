"""
Global hotkey manager using pynput.
Binds configurable hotkeys to recorder/replay actions.
Hotkeys work even when another window has focus.
"""

from __future__ import annotations
from typing import Callable, Optional

from utils.logger import logger

try:
    from pynput import keyboard
    PYNPUT_AVAILABLE = True
except ImportError:
    PYNPUT_AVAILABLE = False


class HotkeyManager:
    def __init__(self, config):
        self._config = config
        self._hotkeys: Optional[keyboard.GlobalHotKeys] = None
        self._bindings: dict[str, Callable] = {}

    def register(self, hotkey: str, callback: Callable) -> None:
        """Register a hotkey string (pynput format) to a callback."""
        self._bindings[hotkey] = callback
        logger.debug("Hotkey registered: {}", hotkey)

    def register_defaults(
        self,
        on_start: Callable,
        on_stop: Callable,
        on_replay: Callable,
        on_abort: Callable,
    ) -> None:
        hk = self._config.hotkeys
        self.register(hk.start_recording, on_start)
        self.register(hk.stop_recording, on_stop)
        self.register(hk.start_replay, on_replay)
        self.register(hk.abort_replay, on_abort)

    def start(self) -> None:
        if not PYNPUT_AVAILABLE:
            logger.warning("pynput not available — hotkeys disabled")
            return
        if not self._bindings:
            logger.warning("No hotkeys registered")
            return

        self._hotkeys = keyboard.GlobalHotKeys(self._bindings)
        self._hotkeys.start()
        logger.info("Hotkeys active: {}", list(self._bindings.keys()))

    def stop(self) -> None:
        if self._hotkeys:
            self._hotkeys.stop()
            logger.info("Hotkeys stopped")
