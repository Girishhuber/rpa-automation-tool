from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import MagicMock

from models.session import Session, SessionStatus, SystemInfo, SCHEMA_VERSION
from models.event import Event, ExplicitWaitEvent
from core.replayer import ReplayEngine
from utils.config import Config


def test_replay_engine_calls_overlay_state(tmp_path: Path):
    config = Config()
    overlay = MagicMock()

    now = datetime.now(timezone.utc).isoformat()
    session = Session(
        name="Overlay wiring test",
        created_at=now,
        updated_at=now,
        status=SessionStatus.COMPLETE,
        system_info=SystemInfo(
            os_version="win",
            screen_width=1920,
            screen_height=1080,
            dpi_scale=1.0,
            monitor_count=1,
            python_version="3.11",
            schema_version=SCHEMA_VERSION,
        ),
    )
    session.events = [
        Event(
            id=1,
            timestamp_ms=0,
            wall_time=now,
            payload=ExplicitWaitEvent(duration_ms=1),
        ).model_dump(mode="json")
    ]

    engine = ReplayEngine(
        config=config,
        screenshot_base_dir=tmp_path,
        overlay=overlay,
    )
    result = engine.replay(session)

    assert result.success is True
    overlay.set_replaying.assert_any_call(True)
    overlay.set_replaying.assert_any_call(False)

