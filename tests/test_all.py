"""
Test suite for the RPA Tool.

Tests run without a live Windows session — they use fixture session files
and mock pywinauto/pynput calls.

Run: pytest tests/ -v
"""

import json
import pytest
from pathlib import Path
from datetime import datetime, timezone
from unittest.mock import MagicMock, patch

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.target import UITarget, BoundingBox
from models.event import (
    Event, EventType,
    MouseClickEvent, TypeTextEvent, KeyComboEvent,
    WindowFocusEvent, WaitEvent,
)
from models.session import Session, SessionStatus, SystemInfo, SCHEMA_VERSION
from storage.session_store import SessionStore
from core.event_pipeline import EventPipeline


# ===========================================================================
# Fixtures
# ===========================================================================

@pytest.fixture
def tmp_store(tmp_path):
    return SessionStore(
        sessions_dir=tmp_path / "sessions",
        db_path=tmp_path / "sessions" / "index.db",
    )


@pytest.fixture
def sample_system_info():
    return SystemInfo(
        os_version="10.0.22621",
        screen_width=1920,
        screen_height=1080,
        dpi_scale=1.0,
        monitor_count=1,
        python_version="3.11.0",
        schema_version=SCHEMA_VERSION,
    )


@pytest.fixture
def sample_session(sample_system_info):
    now = datetime.now(timezone.utc).isoformat()
    return Session(
        name="Test workflow",
        description="Automated test session",
        tags=["test", "excel"],
        created_at=now,
        updated_at=now,
        status=SessionStatus.COMPLETE,
        system_info=sample_system_info,
    )


@pytest.fixture
def sample_events():
    now = datetime.now(timezone.utc).isoformat()

    def make_event(id_, ts, payload):
        return Event(id=id_, timestamp_ms=ts, wall_time=now, payload=payload)

    target = UITarget(
        automation_id="btnSubmit",
        name="Submit",
        control_type="Button",
        window_title="My App",
        process_name="myapp.exe",
        bbox=BoundingBox(left=400, top=200, right=480, bottom=230),
    )

    return [
        make_event(1, 0,    WindowFocusEvent(window_title="My App", process_name="myapp.exe", x=0, y=0, width=800, height=600)),
        make_event(2, 500,  TypeTextEvent(text="hello world", target=UITarget(automation_id="txtInput", name="Input", control_type="Edit", window_title="My App"))),
        make_event(3, 1200, KeyComboEvent(keys=["ctrl", "a"])),
        make_event(4, 1500, MouseClickEvent(x=440, y=215, button="left", target=target)),
        make_event(5, 2000, WaitEvent(duration_ms=300)),
    ]


# ===========================================================================
# Model tests
# ===========================================================================

class TestUITarget:
    def test_match_confidence_automation_id(self):
        t = UITarget(automation_id="btn1", name="OK", control_type="Button")
        assert t.match_confidence() == "automation_id"

    def test_match_confidence_name_type(self):
        t = UITarget(name="OK", control_type="Button")
        assert t.match_confidence() == "name+type"

    def test_match_confidence_bbox_fallback(self):
        t = UITarget(bbox=BoundingBox(left=0, top=0, right=100, bottom=50))
        assert t.match_confidence() == "bbox"

    def test_match_confidence_none(self):
        t = UITarget()
        assert t.match_confidence() == "none"

    def test_bounding_box_properties(self):
        bbox = BoundingBox(left=100, top=200, right=300, bottom=400)
        assert bbox.width == 200
        assert bbox.height == 200
        assert bbox.center == (200, 300)


class TestEventModel:
    def test_click_event_roundtrip(self, sample_events):
        ev = sample_events[3]  # MouseClickEvent
        dumped = ev.model_dump(mode="json")
        restored = Event.model_validate(dumped)
        assert restored.id == ev.id
        assert restored.timestamp_ms == ev.timestamp_ms

    def test_type_text_event(self, sample_events):
        ev = sample_events[1]
        from models.event import TypeTextEvent
        assert isinstance(ev.payload, TypeTextEvent)
        assert ev.payload.text == "hello world"

    def test_event_discriminated_union(self):
        now = datetime.now(timezone.utc).isoformat()
        raw = {
            "id": 1,
            "timestamp_ms": 100,
            "wall_time": now,
            "payload": {
                "type": "mouse_click",
                "x": 100,
                "y": 200,
                "button": "left",
            }
        }
        ev = Event.model_validate(raw)
        from models.event import MouseClickEvent
        assert isinstance(ev.payload, MouseClickEvent)
        assert ev.payload.x == 100


class TestSessionModel:
    def test_session_defaults(self, sample_session):
        assert sample_session.status == SessionStatus.COMPLETE
        assert sample_session.replay_history == []
        assert sample_session.event_count() == 0

    def test_session_serialisation(self, sample_session, sample_events):
        sample_session.events = [e.model_dump(mode="json") for e in sample_events]
        dumped = json.loads(sample_session.model_dump_json())
        restored = Session.model_validate(dumped)
        assert restored.name == "Test workflow"
        assert len(restored.events) == 5

    def test_session_has_id(self, sample_session):
        assert len(sample_session.id) == 36  # UUID format


# ===========================================================================
# Storage tests
# ===========================================================================

class TestSessionStore:
    def test_save_and_load(self, tmp_store, sample_session, sample_events):
        sample_session.events = [e.model_dump(mode="json") for e in sample_events]
        tmp_store.save(sample_session)

        loaded = tmp_store.load(sample_session.id)
        assert loaded.id == sample_session.id
        assert loaded.name == "Test workflow"
        assert len(loaded.events) == 5

    def test_list_sessions(self, tmp_store, sample_session):
        tmp_store.save(sample_session)
        sessions = tmp_store.list_sessions()
        assert len(sessions) == 1
        assert sessions[0]["name"] == "Test workflow"

    def test_session_not_found(self, tmp_store):
        from utils.errors import SessionNotFoundError
        with pytest.raises(SessionNotFoundError):
            tmp_store.load("nonexistent-id")

    def test_delete_session(self, tmp_store, sample_session):
        tmp_store.save(sample_session)
        tmp_store.delete(sample_session.id)
        sessions = tmp_store.list_sessions()
        assert len(sessions) == 0

    def test_corrupt_session(self, tmp_store, sample_session):
        tmp_store.save(sample_session)
        path = tmp_store.session_file(sample_session.id)
        path.write_text("{ not valid json }", encoding="utf-8")

        from utils.errors import SessionCorruptError
        with pytest.raises(SessionCorruptError):
            tmp_store.load(sample_session.id)

    def test_schema_version_in_saved_file(self, tmp_store, sample_session):
        tmp_store.save(sample_session)
        path = tmp_store.session_file(sample_session.id)
        raw = json.loads(path.read_text(encoding="utf-8"))
        assert raw["system_info"]["schema_version"] == SCHEMA_VERSION

    def test_tags_filter(self, tmp_store, sample_session):
        tmp_store.save(sample_session)
        found = tmp_store.list_sessions(tag="excel")
        assert len(found) == 1
        not_found = tmp_store.list_sessions(tag="nonexistent")
        assert len(not_found) == 0


# ===========================================================================
# Event pipeline tests
# ===========================================================================

class TestEventPipeline:
    def test_emits_events(self):
        received = []
        pipeline = EventPipeline(consumer=received.append, debounce_ms=0)
        pipeline.start()

        pipeline.emit_click(100, 200, "left")
        pipeline.emit_type_text("hello")
        pipeline.emit_key_combo(["ctrl", "s"])

        pipeline.stop()
        assert len(received) == 3

    def test_event_ids_sequential(self):
        received = []
        pipeline = EventPipeline(consumer=received.append, debounce_ms=0)
        pipeline.start()

        for _ in range(5):
            pipeline.emit_click(0, 0)

        assert [e.id for e in received] == [1, 2, 3, 4, 5]

    def test_timestamps_are_non_negative(self):
        received = []
        pipeline = EventPipeline(consumer=received.append, debounce_ms=0)
        pipeline.start()
        pipeline.emit_click(0, 0)
        pipeline.emit_click(0, 0)
        pipeline.stop()
        assert all(e.timestamp_ms >= 0 for e in received)

    def test_scroll_debounce(self):
        import time
        received = []
        pipeline = EventPipeline(consumer=received.append, debounce_ms=200)
        pipeline.start()
        pipeline.emit_scroll(0, 0, 0, 1)
        time.sleep(0.001)
        pipeline.emit_scroll(0, 0, 0, 1)
        pipeline.stop()
        assert len(received) == 1

    def test_click_not_debounced(self):
        received = []
        pipeline = EventPipeline(consumer=received.append, debounce_ms=500)
        pipeline.start()
        pipeline.emit_click(0, 0)
        pipeline.emit_click(0, 0)  # clicks are never debounced
        pipeline.stop()
        assert len(received) == 2


# ===========================================================================
# Errors tests
# ===========================================================================

class TestErrors:
    def test_replay_error_carries_event_id(self):
        from utils.errors import ReplayError
        err = ReplayError("something failed", event_id=42)
        assert err.event_id == 42

    def test_element_not_found_is_replay_error(self):
        from utils.errors import ElementNotFoundError, ReplayError
        err = ElementNotFoundError("not found", event_id=7)
        assert isinstance(err, ReplayError)
        assert err.event_id == 7
