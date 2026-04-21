# RPA Tool — Windows Activity Recorder & Replayer

A production-grade tool that records Windows user activity and replays it faithfully.  
Designed for **business applications** — Excel, Outlook, ERP systems, and standard Win32/UIA apps.

---

## Features

- Records mouse clicks, keyboard input, scroll, drag, and window focus changes  
- Captures **UI element identity** (AutomationId, Name, ControlType) — not just pixel coordinates  
- Replays using **5-strategy element matching** for robustness across window sizes and DPI changes  
- System tray app — always available, zero friction  
- Global hotkeys — start/stop without switching windows  
- Session history with SQLite index  
- Structured logging with automatic rotation  
- Single `.exe` deployment via PyInstaller — no Python needed on target machines  

---

## Requirements

- Windows 10 / 11  
- Python 3.11+ (for development)  
- Administrator privileges (required to hook into elevated processes)  

---

## Setup

```bash
# 1. Clone / download the project
cd rpa_tool

# 2. Create virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run (requests elevation automatically)
python main.py
```

---

## Usage

### Tray app (normal mode)

```bash
python main.py
```

Right-click the tray icon to access all actions.

### Global hotkeys (default)

| Action | Hotkey |
|--------|--------|
| Start recording | `Ctrl + Shift + R` |
| Stop recording  | `Ctrl + Shift + S` |
| Replay last     | `Ctrl + Shift + P` |
| Abort replay    | `Ctrl + Shift + Q` |

All hotkeys are configurable in `config.toml`.

### Headless / scripted mode

```bash
# Record until Ctrl+C
python main.py --record "Invoice entry workflow"

# Replay a specific session by ID
python main.py --replay <session-id>

# Debug mode (verbose console output)
python main.py --debug
```

---

## Project structure

```
rpa_tool/
├── core/
│   ├── recorder.py        # pynput hooks + UIA enrichment
│   ├── replayer.py        # event injection + sync waiting
│   ├── matcher.py         # 5-strategy element finder
│   ├── event_pipeline.py  # normalize, debounce, timestamp
│   └── screenshot.py      # mss capture + PIL comparison
├── models/
│   ├── target.py          # UITarget (element identity)
│   ├── event.py           # Event + all payload types
│   └── session.py         # Session (metadata + event list)
├── storage/
│   └── session_store.py   # JSON files + SQLite index
├── ui/
│   ├── tray.py            # pystray tray icon
│   ├── hotkeys.py         # global hotkey manager
│   └── dialogs.py         # session picker + name prompt
├── utils/
│   ├── logger.py          # loguru setup
│   ├── config.py          # TOML config loader
│   └── errors.py          # custom exception hierarchy
├── tests/
│   ├── conftest.py
│   ├── test_all.py
│   └── fixtures/
│       └── excel_workflow.json
├── sessions/              # recorded session files (auto-created)
├── logs/                  # rotating log files (auto-created)
├── app.py                 # application controller
├── main.py                # entry point
├── config.toml            # user configuration
├── requirements.txt
└── build.spec             # PyInstaller build config
```

---

## Session file format

Each session is stored as `sessions/<id>/session.json`:

```json
{
  "id": "uuid",
  "name": "Invoice entry",
  "schema_version": "1.0",
  "events": [
    {
      "id": 1,
      "timestamp_ms": 0,
      "payload": {
        "type": "mouse_click",
        "x": 412, "y": 208,
        "target": {
          "automation_id": "btnSubmit",
          "name": "Submit",
          "control_type": "Button",
          "window_title": "Invoice App"
        }
      }
    }
  ]
}
```

---

## Replay element matching — how it works

When replaying, the engine finds each element using this fallback chain:

| Priority | Strategy | Survives resize? | Survives theme change? |
|----------|----------|-----------------|----------------------|
| 1 | `AutomationId` | Yes | Yes |
| 2 | `Name` + `ControlType` | Yes | Yes |
| 3 | `ClassName` + window title | Yes | Mostly |
| 4 | Bounding box (DPI-adjusted) | No | Yes |
| 5 | Screenshot template match | No | No |

Always prefer apps where developers set `AutomationId` — most Office and line-of-business apps do.

---

## Build standalone EXE

```bash
pyinstaller build.spec
# Output: dist/rpa_tool.exe
```

The EXE includes all dependencies and requests UAC elevation automatically.  
Deploy by copying `rpa_tool.exe` to the target machine — no Python installation needed.

---

## Running tests

```bash
pytest tests/ -v
```

Tests run without a live Windows session — they use fixture JSON files and mock I/O.

---

## Configuration reference (`config.toml`)

| Key | Default | Description |
|-----|---------|-------------|
| `recorder.capture_screenshots` | `true` | Save screenshot on each click |
| `recorder.debounce_ms` | `50` | Ignore duplicate events within Nms |
| `replay.speed` | `1.0` | Playback speed (2.0 = 2x faster) |
| `replay.wait_timeout_ms` | `10000` | Max wait for element readiness |
| `replay.retry_attempts` | `3` | Retries per failed event |
| `storage.sessions_dir` | `sessions` | Where session files are stored |
| `storage.max_sessions` | `500` | Session cap before rotation |

---

## Known limitations

- **UAC dialogs** — Windows blocks automation of secure desktop prompts. These cannot be recorded or replayed.  
- **DirectX / game windows** — Not a target use case. UIA does not work with GPU-rendered surfaces.  
- **Browser content areas** — Web content inside Chrome/Edge uses its own accessibility tree. Some elements may not expose `AutomationId`. Use Name+Type matching.  
- **DPI changes between machines** — If you record on a 100% DPI machine and replay on 150%, bounding-box fallback coordinates are scaled automatically. AutomationId matching is DPI-independent.
