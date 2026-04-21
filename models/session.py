from __future__ import annotations
from enum import Enum
from typing import Optional,List,Any
from pydantic import BaseModel, Field
import uuid

SCHEMA_VERSION = "2.0"

class SessionStatus(str, Enum):
    RECORDING = "recording"
    COMPLETE  = "complete"
    REPLAYING = "replaying"
    FAILED    = "failed"
    PAUSED = "Paused"
    
class ReplayMode(str, Enum):
    STRICT   = "strict"    
    ADAPTIVE = "adaptive"
    
class FailureStrategy(str, Enum):
    STOP   = "stop"
    SKIP   = "skip"
    RETRY  = "retry"

class SystemInfo(BaseModel):
    os_version: str
    screen_width: int
    screen_height: int
    dpi_scale: float
    monitor_count: int
    python_version: str
    machine_name: Optional[str] = None
    cpu_info: Optional[str] = None
    gpu_info: Optional[str] = None
    schema_version: str = SCHEMA_VERSION
    
class SessionSettings(BaseModel):
    screenshot_on_click: bool = True
    capture_scroll_events: bool = True

    
    replay_mode: ReplayMode = ReplayMode.ADAPTIVE
    failure_strategy: FailureStrategy = FailureStrategy.RETRY

    max_retries: int = 2
    retry_interval_ms: int = 500
    default_timeout_ms: int = 5000

    use_ocr_fallback: bool = True
    use_image_fallback: bool = True
    use_bbox_fallback: bool = True

  
    speed_factor: float = 1.0   

class ReplayResult(BaseModel):
    replayed_at: str
    success: bool
    events_total: int
    events_completed: int
    failed_event_id: Optional[int] = None
    error_message: Optional[str] = None
    duration_ms: int
    retries_performed: int = 0
    skipped_events: int = 0
    warnings: List[str] = Field(default_factory=list)
    
class SessionStats(BaseModel):
    total_events: int = 0
    click_count: int = 0
    typing_count: int = 0
    scroll_count: int = 0
    error_count: int = 0
    avg_event_interval_ms: Optional[float] = None
    
    
#environmental info
class EnvironmentSnapshot(BaseModel):
    active_apps: List[str] = Field(default_factory=list)
    screen_resolution: Optional[str] = None
    dpi_scale: Optional[float] = None
    

class Session(BaseModel):
    #identity
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    name: str
    description: str = ""
    tags: list[str] = []
    
    #Time
    created_at: str
    updated_at: str
    duration_ms: int = 0
    
    #Status
    status: SessionStatus = SessionStatus.RECORDING
    
    #System
    system_info: SystemInfo
    environment: Optional[EnvironmentSnapshot] = None
    
    #core data
    events: List[Any] = Field(default_factory=list)
    
    replay_history: List[ReplayResult] = Field(default_factory=list)
    
    #setting
    session_setings: SessionSettings= Field(default_factory=SessionSettings)
    
    stats:  SessionStats = Field(default_factory=SessionStats)  
    screenshot_on_click: bool = True
    capture_scroll_events: bool = True

    model_config = {"use_enum_values": True}

    def event_count(self) -> int:
        return len(self.events)
    def add_event(self, event: Any) -> None:
        self.events.append(event)
        self.stats.total_events += 1

    def add_replay_result(self, result: ReplayResult) -> None:
        self.replay_history.append(result)

    def last_replay(self) -> Optional[ReplayResult]:
        if not self.replay_history:
            return None
        return self.replay_history[-1]

    def success_rate(self) -> float:
        if not self.replay_history:
            return 0.0
        success = sum(1 for r in self.replay_history if r.success)
        return success / len(self.replay_history)