"""
RPA Tool — entry point.

Run with:
    python main.py
    python main.py --debug
    python main.py --record "My workflow"   # headless recording (no tray)
    python main.py --replay <session_id>    # headless replay
"""

from __future__ import annotations
import argparse
import ctypes
import sys
from pathlib import Path


def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def request_elevation() -> None:
    """Re-launch the current script with UAC elevation prompt."""
    params = " ".join(f'"{a}"' for a in sys.argv)
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, params, None, 1
    )
    sys.exit(0)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Windows activity recorder and replayer"
    )
    parser.add_argument("--debug", action="store_true", help="Enable debug console output")
    parser.add_argument("--config", default="config.toml", help="Path to config file")
    parser.add_argument("--record", metavar="NAME", help="Start recording immediately (headless)")
    parser.add_argument("--replay", metavar="SESSION_ID", help="Replay a session ID (headless)")
    parser.add_argument("--no-admin-check", action="store_true", help="Skip elevation check")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    # Check elevation — required to hook into elevated processes (ERP, UAC dialogs)
    if not args.no_admin_check and not is_admin():
        print(
            "RPA Tool works best with administrator privileges.\n"
            "Re-launching with elevation..."
        )
        request_elevation()
        return

    # Load config
    from utils.config import load_config
    config = load_config(Path(args.config))
    if args.debug:
        config.debug = True

    # --- Headless recording mode ---
    if args.record:
        _headless_record(config, args.record)
        return

    # --- Headless replay mode ---
    if args.replay:
        _headless_replay(config, args.replay)
        return

    # --- Normal tray app mode ---
    from app import App
    application = App(config)
    application.run()


def _headless_record(config, name: str) -> None:
    """Record until Ctrl+C is pressed, then save."""
    from utils.logger import setup_logging
    from pathlib import Path
    setup_logging(Path(config.storage.logs_dir), console=True)

    from core.recorder import Recorder
    from storage.session_store import SessionStore

    store = SessionStore(
        sessions_dir=Path(config.storage.sessions_dir),
        db_path=Path(config.storage.db_path),
    )

    recorder = Recorder(config, Path(config.storage.sessions_dir), session_name=name)
    recorder.start()

    print(f"\n  Recording: '{name}'")
    print("  Press Ctrl+C to stop...\n")

    try:
        import time
        while True:
            time.sleep(0.5)
            print(f"\r  Events captured: {recorder.event_count}", end="", flush=True)
    except KeyboardInterrupt:
        pass

    print("\n  Stopping...")
    session = recorder.stop()
    store.save(session)
    print(f"  Saved session: {session.id}  ({session.event_count()} events)")


def _headless_replay(config, session_id: str) -> None:
    """Replay a specific session by ID."""
    from utils.logger import setup_logging
    from pathlib import Path
    setup_logging(Path(config.storage.logs_dir), console=True)

    from core.replayer import ReplayEngine
    from storage.session_store import SessionStore

    store = SessionStore(
        sessions_dir=Path(config.storage.sessions_dir),
        db_path=Path(config.storage.db_path),
    )

    try:
        session = store.load(session_id)
    except Exception as exc:
        print(f"  Error loading session: {exc}")
        sys.exit(1)

    print(f"\n  Replaying: '{session.name}' ({session.event_count()} events)")

    def progress(done, total):
        pct = int(done / total * 100) if total else 0
        print(f"\r  Progress: {pct}% ({done}/{total})", end="", flush=True)

    engine = ReplayEngine(
        config=config,
        screenshot_base_dir=Path(config.storage.sessions_dir),
        on_progress=progress,
        overlay=None,
    )
    result = engine.replay(session)
    print()

    if result.success:
        print(f"  Replay complete — {result.events_completed} events in {result.duration_ms / 1000:.1f}s")
    else:
        print(f"  Replay FAILED at event #{result.failed_event_id}: {result.error_message}")
        sys.exit(1)


if __name__ == "__main__":
    main()
