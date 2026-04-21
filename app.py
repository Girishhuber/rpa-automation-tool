from __future__ import annotations
import sys
import threading
from pathlib import Path
from typing import Optional

from utils.logger import logger, setup_logging
from utils.config import Config
from utils.errors import RPAError, RecorderError, ReplayError
from storage.session_store import SessionStore
from core.recorder import Recorder
from core.replayer import ReplayEngine
from models.session import Session
from ui.tray import TrayApp
from ui.hotkeys import HotkeyManager
from ui.dialogs import SessionPickerDialog, RecordingNameDialog


class App:
    def __init__(self, config: Config):
        self._config = config

        # Setup logging first
        setup_logging(
            log_dir=Path(config.storage.logs_dir),
            console=config.debug,
        )
        logger.info("RPA Tool starting up")

        # Storage
        self._store = SessionStore(
            sessions_dir=Path(config.storage.sessions_dir),
            db_path=Path(config.storage.db_path),
        )

        # State
        self._recorder: Optional[Recorder] = None
        self._active_session: Optional[Session] = None
        self._replay_engine: Optional[ReplayEngine] = None
        self._last_session_id: Optional[str] = None
        self._is_replaying = False
        # Overlay disabled (can block clicks on some systems).
        self._overlay = None

        # UI
        self._tray = TrayApp(
            config=config,
            on_start_recording=self.start_recording,
            on_stop_recording=self.stop_recording,
            on_replay_last=self.replay_last,
            on_replay_picker=self.show_replay_picker,
            on_quit=self.quit,
        )

        # Hotkeys
        self._hotkeys = HotkeyManager(config)
        self._hotkeys.register_defaults(
            on_start=self.start_recording,
            on_stop=self.stop_recording,
            on_replay=self.replay_last,
            on_abort=self.abort_replay,
        )

    def run(self) -> None:
        """Start hotkeys and the tray icon (blocking)."""
        self._hotkeys.start()
        logger.info(
            "Ready. Hotkeys — Record: {} | Stop: {} | Replay: {}",
            self._config.hotkeys.start_recording,
            self._config.hotkeys.stop_recording,
            self._config.hotkeys.start_replay,
        )
        self._tray.run()  # blocks until quit

    def quit(self) -> None:
        logger.info("Shutting down")
        self._hotkeys.stop()
        if self._recorder and self._recorder.is_recording:
            try:
                self.stop_recording()
            except Exception:
                pass

    def start_recording(self) -> None:
        if self._recorder and self._recorder.is_recording:
            logger.warning("Already recording — ignoring start request")
            return

        # Ask for session name
        dialog = RecordingNameDialog()
        name = dialog.show()
        if name is None:
            return  # user cancelled

        try:
            scr_dir = Path(self._config.storage.sessions_dir)
            self._recorder = Recorder(self._config, scr_dir, session_name=name)
            session = self._recorder.start()
            self._active_session = session
            self._tray.set_recording(True, f"Recording: {name}")
            logger.info("Recording started: {}", name)
        except RPAError as exc:
            logger.error("Failed to start recording: {}", exc)
            self._tray.set_status(f"Error: {exc}")

    def stop_recording(self) -> None:
        if not self._recorder or not self._recorder.is_recording:
            logger.warning("Not recording — ignoring stop request")
            return

        try:
            session = self._recorder.stop()
            self._active_session = session
            self._last_session_id = session.id
            self._store.save(session)
            self._tray.set_recording(False, f"Saved: {session.name} ({session.event_count()} events)")
            logger.info(
                "Recording saved: {} — {} events",
                session.name,
                session.event_count(),
            )
        except RPAError as exc:
            logger.error("Failed to stop recording: {}", exc)
            self._tray.set_status(f"Error saving: {exc}")

    def replay_last(self) -> None:
        if not self._last_session_id:
            logger.warning("No session recorded yet in this session")
            # Try loading the most recent from store
            sessions = self._store.list_sessions(limit=1)
            if not sessions:
                self._tray.set_status("No sessions found")
                return
            self._last_session_id = sessions[0]["id"]

        self._start_replay(self._last_session_id)

    def show_replay_picker(self) -> None:
        sessions = self._store.list_sessions(limit=100)
        if not sessions:
            self._tray.set_status("No sessions found")
            return

        dialog = SessionPickerDialog(
            sessions=sessions,
            on_select=self._start_replay,
            on_delete=self._store.delete,
        )
        dialog.show()

    def abort_replay(self) -> None:
        if self._replay_engine and self._is_replaying:
            self._replay_engine.abort()
            self._tray.set_status("Replay aborted")

    def _start_replay(self, session_id: str) -> None:
        if self._is_replaying:
            logger.warning("Already replaying")
            return

        threading.Thread(
            target=self._replay_worker,
            args=(session_id,),
            daemon=True,
        ).start()

    def _replay_worker(self, session_id: str) -> None:
        self._is_replaying = True
        try:
            session = self._store.load(session_id)
            self._tray.set_status(f"Replaying: {session.name}...")

            scr_dir = Path(self._config.storage.sessions_dir)
            self._replay_engine = ReplayEngine(
                config=self._config,
                screenshot_base_dir=scr_dir,
                on_progress=self._on_replay_progress,
                overlay=None,
            )

            result = self._replay_engine.replay(session)

            # Save result back to session
            session.replay_history.append(result)
            self._store.save(session)

            if result.success:
                self._tray.set_status(
                    f"Replay complete: {result.events_completed}/{result.events_total} events"
                )
                logger.info("Replay succeeded for session {}", session_id)
            else:
                self._tray.set_status(
                    f"Replay failed at event #{result.failed_event_id}"
                )
                logger.error(
                    "Replay failed for session {} at event #{}: {}",
                    session_id, result.failed_event_id, result.error_message,
                )

        except RPAError as exc:
            logger.error("Replay error: {}", exc)
            self._tray.set_status(f"Replay error: {exc}")
        finally:
            self._is_replaying = False
            self._replay_engine = None

    def _on_replay_progress(self, completed: int, total: int) -> None:
        pct = int(completed / total * 100) if total else 0
        self._tray.set_status(f"Replaying... {pct}% ({completed}/{total})")
