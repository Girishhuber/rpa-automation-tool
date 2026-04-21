from __future__ import annotations
import json, sqlite3, shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from models.session import Session, SCHEMA_VERSION
from utils.errors import SessionNotFoundError, SessionCorruptError, SchemaMismatchError
from utils.logger import logger

SUPPORTED = {"1.0", "2.0"}


class SessionStore:
    def __init__(self, sessions_dir: Path, db_path: Path):
        self.sessions_dir = sessions_dir
        self.db_path = db_path
        sessions_dir.mkdir(parents=True, exist_ok=True)
        db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _init_db(self):
        with self._db() as c:
            c.execute("""CREATE TABLE IF NOT EXISTS sessions (
                id TEXT PRIMARY KEY, name TEXT, created_at TEXT,
                updated_at TEXT, status TEXT, event_count INTEGER,
                duration_ms INTEGER, tags TEXT, path TEXT)""")
            c.execute("CREATE INDEX IF NOT EXISTS idx_created ON sessions(created_at DESC)")

    def _db(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def session_dir(self, sid: str) -> Path:
        return self.sessions_dir / sid

    def session_file(self, sid: str) -> Path:
        return self.session_dir(sid) / "session.json"

    def save(self, session: Session) -> None:
        path = self.session_file(session.id)
        path.parent.mkdir(parents=True, exist_ok=True)
        session.updated_at = datetime.now(timezone.utc).isoformat()
        path.write_text(json.dumps(session.model_dump(mode="json"), indent=2, ensure_ascii=False), encoding="utf-8")
        with self._db() as c:
            c.execute("""INSERT INTO sessions VALUES (?,?,?,?,?,?,?,?,?)
                ON CONFLICT(id) DO UPDATE SET name=excluded.name, updated_at=excluded.updated_at,
                status=excluded.status, event_count=excluded.event_count,
                duration_ms=excluded.duration_ms, tags=excluded.tags""",
                (session.id, session.name, session.created_at, session.updated_at,
                 str(session.status), len(session.events), session.duration_ms,
                 ",".join(session.tags), str(path)))

    def load(self, sid: str) -> Session:
        path = self.session_file(sid)
        if not path.exists():
            raise SessionNotFoundError(f"Session not found: {sid}")
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            raise SessionCorruptError(f"Cannot parse session {sid}: {e}") from e
        schema = raw.get("system_info", {}).get("schema_version", "unknown")
        if schema not in SUPPORTED:
            raise SchemaMismatchError(f"Schema '{schema}' not supported")
        try:
            return Session.model_validate(raw)
        except Exception as e:
            raise SessionCorruptError(f"Validation failed: {e}") from e

    def list_sessions(self, limit: int = 50, offset: int = 0, tag: Optional[str] = None) -> list[dict]:
        q = "SELECT * FROM sessions"
        params = []
        if tag:
            q += " WHERE tags LIKE ?"; params.append(f"%{tag}%")
        q += " ORDER BY created_at DESC LIMIT ? OFFSET ?"
        params += [limit, offset]
        with self._db() as c:
            return [dict(r) for r in c.execute(q, params).fetchall()]

    def delete(self, sid: str) -> None:
        d = self.session_dir(sid)
        if d.exists():
            shutil.rmtree(d)
        with self._db() as c:
            c.execute("DELETE FROM sessions WHERE id=?", (sid,))