from __future__ import annotations
from enum import Enum
from typing import Optional,List,Any, Literal, Union, Annotated
from pydantic import BaseModel, Field
from .target import UITarget


class EventType(str, Enum):
   
    MOUSE_CLICK         = "mouse_click"
    MOUSE_DOUBLE_CLICK  = "mouse_double_click"
    MOUSE_RIGHT_CLICK   = "mouse_right_click"
    MOUSE_MIDDLE_CLICK  = "mouse_middle_click"
    MOUSE_SCROLL        = "mouse_scroll"
    MOUSE_DRAG          = "mouse_drag"
    MOUSE_HOVER         = "mouse_hover"

   
    KEY_PRESS           = "key_press"
    KEY_COMBO           = "key_combo"
    TYPE_TEXT           = "type_text"

    
    CLIPBOARD_COPY      = "clipboard_copy"
    CLIPBOARD_CUT       = "clipboard_cut"
    CLIPBOARD_PASTE     = "clipboard_paste"
    CLIPBOARD_PASTE_SPECIAL = "clipboard_paste_special"


    BROWSER_NAVIGATE    = "browser_navigate"
    BROWSER_TAB_SWITCH  = "browser_tab_switch"
    BROWSER_TAB_NEW     = "browser_tab_new"
    BROWSER_TAB_CLOSE   = "browser_tab_close"
    BROWSER_BACK        = "browser_back"
    BROWSER_FORWARD     = "browser_forward"
    BROWSER_REFRESH     = "browser_refresh"
    BROWSER_WAIT_LOAD   = "browser_wait_load"
    BROWSER_SELECT_TEXT = "browser_select_text"
    BROWSER_SCROLL_TO   = "browser_scroll_to"

    WINDOW_FOCUS        = "window_focus"
    WINDOW_OPEN         = "window_open"
    WINDOW_CLOSE        = "window_close"
    WINDOW_RESIZE       = "window_resize"
    WINDOW_MAXIMIZE     = "window_maximize"
    WINDOW_MINIMIZE     = "window_minimize"
    DIALOG_RESPONSE     = "dialog_response"  
    FILE_DIALOG         = "file_dialog"       
    DROPDOWN_SELECT     = "dropdown_select"   
    CHECKBOX_TOGGLE     = "checkbox_toggle"
    RADIO_SELECT        = "radio_select"

    EXCEL_CELL_SELECT   = "excel_cell_select"   
    EXCEL_RANGE_SELECT  = "excel_range_select"  
    EXCEL_SHEET_SWITCH  = "excel_sheet_switch" 


    SCREENSHOT_CHECKPOINT = "screenshot_checkpoint"
    EXPLICIT_WAIT       = "explicit_wait"
    PROCESS_LAUNCH      = "process_launch"
    HOTKEY_TRIGGER      = "hotkey_trigger"   
    
class ReplayMeta(BaseModel):
    wait_for: Optional[str] = None              
    timeout_ms: int = 5000
    retry_count: int = 2
    retry_interval_ms: int = 500
    continue_on_failure: bool = False
    
class Assertion(BaseModel):
    type: str          # "text_present", "element_visible"
    expected: Any
    
class ExecutionContext(BaseModel):
    active_app: Optional[str] = None
    active_window: Optional[str] = None
    process_id: Optional[int] = None
    
class BaseEvent(BaseModel):
    target: Optional[UITarget] = None

    
    replay: ReplayMeta = Field(default_factory=ReplayMeta)
    assertions: List[Assertion] = Field(default_factory=list)
    context: Optional[ExecutionContext] = None

    confidence_score: float = 0.0



class MouseClickEvent(BaseModel):
    type: Literal[EventType.MOUSE_CLICK] = EventType.MOUSE_CLICK
    x: int; y: int
    button: str = "left"
    target: Optional[UITarget] = None

class MouseDoubleClickEvent(BaseModel):
    type: Literal[EventType.MOUSE_DOUBLE_CLICK] = EventType.MOUSE_DOUBLE_CLICK
    x: int; y: int
    target: Optional[UITarget] = None

class MouseRightClickEvent(BaseModel):
    type: Literal[EventType.MOUSE_RIGHT_CLICK] = EventType.MOUSE_RIGHT_CLICK
    x: int; y: int
    target: Optional[UITarget] = None

class MouseMiddleClickEvent(BaseModel):
    type: Literal[EventType.MOUSE_MIDDLE_CLICK] = EventType.MOUSE_MIDDLE_CLICK
    x: int; y: int
    target: Optional[UITarget] = None

class MouseScrollEvent(BaseModel):
    type: Literal[EventType.MOUSE_SCROLL] = EventType.MOUSE_SCROLL
    x: int; y: int
    dx: int; dy: int   # dy<0 = scroll up, dy>0 = scroll down
    target: Optional[UITarget] = None

class MouseDragEvent(BaseModel):
    type: Literal[EventType.MOUSE_DRAG] = EventType.MOUSE_DRAG
    start_x: int; start_y: int
    end_x: int; end_y: int
    button: str = "left"
    duration_ms: int = 0
    start_target: Optional[UITarget] = None
    end_target: Optional[UITarget] = None

class MouseHoverEvent(BaseModel):
    type: Literal[EventType.MOUSE_HOVER] = EventType.MOUSE_HOVER
    x: int; y: int
    duration_ms: int = 500
    target: Optional[UITarget] = None

class KeyPressEvent(BaseModel):
    type: Literal[EventType.KEY_PRESS] = EventType.KEY_PRESS
    key: str   # "enter", "tab", "f2", "delete", "escape"
    target: Optional[UITarget] = None

class KeyComboEvent(BaseModel):
    type: Literal[EventType.KEY_COMBO] = EventType.KEY_COMBO
    keys: list[str]   # ["ctrl","c"], ["ctrl","shift","end"]
    target: Optional[UITarget] = None

class TypeTextEvent(BaseModel):
    type: Literal[EventType.TYPE_TEXT] = EventType.TYPE_TEXT
    text: str
    target: Optional[UITarget] = None
    clear_first: bool = False
    method: str = "keys"   # "keys" | "clipboard" | "set_value"
    cell_ref: Optional[str] = None
    sheet_name: Optional[str] = None
    force_plain_text: bool = False

class ClipboardCopyEvent(BaseModel):
    type: Literal[EventType.CLIPBOARD_COPY] = EventType.CLIPBOARD_COPY
    content: Optional[str] = None   # captured text content at record time
    target: Optional[UITarget] = None

class ClipboardCutEvent(BaseModel):
    type: Literal[EventType.CLIPBOARD_CUT] = EventType.CLIPBOARD_CUT
    content: Optional[str] = None
    target: Optional[UITarget] = None

class ClipboardPasteEvent(BaseModel):
    type: Literal[EventType.CLIPBOARD_PASTE] = EventType.CLIPBOARD_PASTE
    content: Optional[str] = None   # what was pasted (for verification)
    target: Optional[UITarget] = None

class ClipboardPasteSpecialEvent(BaseModel):
    type: Literal[EventType.CLIPBOARD_PASTE_SPECIAL] = EventType.CLIPBOARD_PASTE_SPECIAL
    paste_type: str = "values"   # "values" | "formats" | "all"
    target: Optional[UITarget] = None

class BrowserNavigateEvent(BaseModel):
    type: Literal[EventType.BROWSER_NAVIGATE] = EventType.BROWSER_NAVIGATE
    url: str
    wait_for_load: bool = True

class BrowserTabSwitchEvent(BaseModel):
    type: Literal[EventType.BROWSER_TAB_SWITCH] = EventType.BROWSER_TAB_SWITCH
    tab_index: int
    tab_url: Optional[str] = None
    tab_title: Optional[str] = None

class BrowserTabNewEvent(BaseModel):
    type: Literal[EventType.BROWSER_TAB_NEW] = EventType.BROWSER_TAB_NEW
    url: Optional[str] = None

class BrowserTabCloseEvent(BaseModel):
    type: Literal[EventType.BROWSER_TAB_CLOSE] = EventType.BROWSER_TAB_CLOSE
    tab_index: int

class BrowserBackEvent(BaseModel):
    type: Literal[EventType.BROWSER_BACK] = EventType.BROWSER_BACK

class BrowserForwardEvent(BaseModel):
    type: Literal[EventType.BROWSER_FORWARD] = EventType.BROWSER_FORWARD

class BrowserRefreshEvent(BaseModel):
    type: Literal[EventType.BROWSER_REFRESH] = EventType.BROWSER_REFRESH

class BrowserWaitLoadEvent(BaseModel):
    type: Literal[EventType.BROWSER_WAIT_LOAD] = EventType.BROWSER_WAIT_LOAD
    timeout_ms: int = 15000
    url_pattern: Optional[str] = None   # wait until URL matches

class BrowserSelectTextEvent(BaseModel):
    type: Literal[EventType.BROWSER_SELECT_TEXT] = EventType.BROWSER_SELECT_TEXT
    selected_text: str
    target: Optional[UITarget] = None

class BrowserScrollToEvent(BaseModel):
    type: Literal[EventType.BROWSER_SCROLL_TO] = EventType.BROWSER_SCROLL_TO
    target: Optional[UITarget] = None
    scroll_x: int = 0
    scroll_y: int = 0

class WindowFocusEvent(BaseModel):
    type: Literal[EventType.WINDOW_FOCUS] = EventType.WINDOW_FOCUS
    window_title: str
    process_name: str
    x: int; y: int; width: int; height: int

class WindowOpenEvent(BaseModel):
    type: Literal[EventType.WINDOW_OPEN] = EventType.WINDOW_OPEN
    window_title: str
    process_name: str
    launched_by: Optional[str] = None  

class WindowCloseEvent(BaseModel):
    type: Literal[EventType.WINDOW_CLOSE] = EventType.WINDOW_CLOSE
    window_title: str
    process_name: str

class WindowResizeEvent(BaseModel):
    type: Literal[EventType.WINDOW_RESIZE] = EventType.WINDOW_RESIZE
    window_title: str
    new_x: int; new_y: int; new_width: int; new_height: int

class WindowMaximizeEvent(BaseModel):
    type: Literal[EventType.WINDOW_MAXIMIZE] = EventType.WINDOW_MAXIMIZE
    window_title: str

class WindowMinimizeEvent(BaseModel):
    type: Literal[EventType.WINDOW_MINIMIZE] = EventType.WINDOW_MINIMIZE
    window_title: str

class DialogResponseEvent(BaseModel):
    type: Literal[EventType.DIALOG_RESPONSE] = EventType.DIALOG_RESPONSE
    dialog_title: str
    response: str   # "OK" | "Cancel" | "Yes" | "No" | "Retry"
    message_text: Optional[str] = None

class FileDialogEvent(BaseModel):
    type: Literal[EventType.FILE_DIALOG] = EventType.FILE_DIALOG
    dialog_type: str   # "open" | "save" | "folder"
    path: str
    filter_text: Optional[str] = None

class DropdownSelectEvent(BaseModel):
    type: Literal[EventType.DROPDOWN_SELECT] = EventType.DROPDOWN_SELECT
    selected_text: str
    selected_index: Optional[int] = None
    target: Optional[UITarget] = None

class CheckboxToggleEvent(BaseModel):
    type: Literal[EventType.CHECKBOX_TOGGLE] = EventType.CHECKBOX_TOGGLE
    checked: bool
    target: Optional[UITarget] = None

class RadioSelectEvent(BaseModel):
    type: Literal[EventType.RADIO_SELECT] = EventType.RADIO_SELECT
    option_text: str
    target: Optional[UITarget] = None

class ExcelCellSelectEvent(BaseModel):
    type: Literal[EventType.EXCEL_CELL_SELECT] = EventType.EXCEL_CELL_SELECT
    cell_ref: str          # "B4", "Sheet2!C12"
    sheet_name: Optional[str] = None
    target: Optional[UITarget] = None

class ExcelRangeSelectEvent(BaseModel):
    type: Literal[EventType.EXCEL_RANGE_SELECT] = EventType.EXCEL_RANGE_SELECT
    range_ref: str         # "B2:D10"
    sheet_name: Optional[str] = None

class ExcelSheetSwitchEvent(BaseModel):
    type: Literal[EventType.EXCEL_SHEET_SWITCH] = EventType.EXCEL_SHEET_SWITCH
    sheet_name: str
    sheet_index: int

# Screenshot screenshot/checkpoint event used by pipeline and storage.
# Single canonical model with both `monitor_index` and optional `label`.
class ScreenshotEvent(BaseModel):
    type: Literal[EventType.SCREENSHOT_CHECKPOINT] = EventType.SCREENSHOT_CHECKPOINT
    path: str
    monitor_index: int = 0
    label: Optional[str] = None

# Backwards-compatible name used in older code / storage
ScreenshotCheckpointEvent = ScreenshotEvent

class ExplicitWaitEvent(BaseModel):
    type: Literal[EventType.EXPLICIT_WAIT] = EventType.EXPLICIT_WAIT
    duration_ms: int

class ProcessLaunchEvent(BaseModel):
    type: Literal[EventType.PROCESS_LAUNCH] = EventType.PROCESS_LAUNCH
    executable: str
    arguments: list[str] = Field(default_factory=list)
    wait_for_window_title: Optional[str] = None

class HotkeyTriggerEvent(BaseModel):
    type: Literal[EventType.HOTKEY_TRIGGER] = EventType.HOTKEY_TRIGGER
    hotkey: str
    description: Optional[str] = None


EventPayload = Annotated[
    Union[
        MouseClickEvent, MouseDoubleClickEvent, MouseRightClickEvent,
        MouseMiddleClickEvent, MouseScrollEvent, MouseDragEvent, MouseHoverEvent,
        KeyPressEvent, KeyComboEvent, TypeTextEvent,
        ClipboardCopyEvent, ClipboardCutEvent, ClipboardPasteEvent, ClipboardPasteSpecialEvent,
        BrowserNavigateEvent, BrowserTabSwitchEvent, BrowserTabNewEvent,
        BrowserTabCloseEvent, BrowserBackEvent, BrowserForwardEvent,
        BrowserRefreshEvent, BrowserWaitLoadEvent, BrowserSelectTextEvent, BrowserScrollToEvent,
        WindowFocusEvent, WindowOpenEvent, WindowCloseEvent, WindowResizeEvent,
        WindowMaximizeEvent, WindowMinimizeEvent,
        DialogResponseEvent, FileDialogEvent, DropdownSelectEvent,
        CheckboxToggleEvent, RadioSelectEvent,
        ExcelCellSelectEvent, ExcelRangeSelectEvent, ExcelSheetSwitchEvent,
        ScreenshotEvent, ExplicitWaitEvent, ProcessLaunchEvent, HotkeyTriggerEvent,
    ],
    Field(discriminator="type"),
]


class Event(BaseModel):
    id: int
    timestamp_ms: int
    wall_time: str
    payload: EventPayload
    note: Optional[str] = None

    intent: Optional[str] = None
    action_group: Optional[str] = None

    model_config = {"use_enum_values": True}


WaitEvent = ExplicitWaitEvent