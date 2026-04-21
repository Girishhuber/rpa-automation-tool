from __future__ import annotations
import sys
from dataclasses import dataclass, field
from pathlib import Path

if sys.version_info >= (3, 11):
    import tomllib
else:
    try:
        import tomli as tomllib
    except ImportError:
        tomllib = None  # type: ignore

from .errors import ConfigError


@dataclass
class HotkeyConfig:
    start_recording: str = "<ctrl>+<shift>+r"
    stop_recording: str  = "<ctrl>+<shift>+s"
    start_replay: str    = "<ctrl>+<shift>+p"
    abort_replay: str    = "<ctrl>+<shift>+q"
    checkpoint: str      = "<ctrl>+<shift>+c"  


@dataclass
class RecorderConfig:
    capture_screenshots: bool = True
    screenshot_on_every_click: bool = True
    capture_scroll: bool = True
    capture_hover: bool = False         
    debounce_scroll_ms: int = 80
    debounce_hover_ms: int = 500
    max_session_minutes: int = 120
    text_flush_idle_ms: int = 800      
    capture_clipboard_content: bool = True  
    browser_cdp_port: int = 9222         
    detect_excel_cells: bool = True      

@dataclass
class ReplayConfig:
    speed: float = 1.0
    min_delay_ms: int = 80
    wait_timeout_ms: int = 15000
    retry_attempts: int = 3
    screenshot_on_failure: bool = True
    browser_action_delay_ms: int = 150 
    excel_action_delay_ms: int = 150  


@dataclass
class StorageConfig:
    sessions_dir: Path = field(default_factory=lambda: Path("sessions"))
    logs_dir: Path     = field(default_factory=lambda: Path("logs"))
    db_path: Path      = field(default_factory=lambda: Path("sessions/index.db"))
    max_sessions: int  = 500


@dataclass
class OverlayConfig:
    # Disabled by default because a full-screen tkinter overlay can block input
    # on some Windows setups if click-through/transparency fails.
    enabled: bool = False
    click_color: str = "#FF4444"        
    recording_color: str = "#FF0000"    
    replay_color: str = "#00AA44"       
    ring_size: int = 40                 
    ring_duration_ms: int = 600         
    show_event_log: bool = True       
    log_max_lines: int = 8


@dataclass
class Config:
    hotkeys: HotkeyConfig    = field(default_factory=HotkeyConfig)
    recorder: RecorderConfig = field(default_factory=RecorderConfig)
    replay: ReplayConfig     = field(default_factory=ReplayConfig)
    storage: StorageConfig   = field(default_factory=StorageConfig)
    overlay: OverlayConfig   = field(default_factory=OverlayConfig)
    debug: bool = False


def load_config(config_path: Path = Path("config.toml")) -> Config:
    cfg = Config()
    if not config_path.exists():
        return cfg
    if tomllib is None:
        raise ConfigError("tomllib not available. Install tomli: pip install tomli")
    try:
        with open(config_path, "rb") as f:
            data = tomllib.load(f)
    except Exception as exc:
        raise ConfigError(f"Failed to parse {config_path}: {exc}") from exc

    def _apply(section_data, cls, existing):
        if not section_data:
            return existing
        fields = {k: v for k, v in section_data.items() if hasattr(existing, k)}
        for k, v in fields.items():
            setattr(existing, k, v)
        return existing

    _apply(data.get("hotkeys"), HotkeyConfig, cfg.hotkeys)
    _apply(data.get("recorder"), RecorderConfig, cfg.recorder)
    _apply(data.get("replay"), ReplayConfig, cfg.replay)
    _apply(data.get("overlay"), OverlayConfig, cfg.overlay)
    if st := data.get("storage"):
        cfg.storage = StorageConfig(
            sessions_dir=Path(st.get("sessions_dir", "sessions")),
            logs_dir=Path(st.get("logs_dir", "logs")),
            db_path=Path(st.get("db_path", "sessions/index.db")),
            max_sessions=st.get("max_sessions", 500),
        )
    cfg.debug = data.get("debug", False)
    return cfg