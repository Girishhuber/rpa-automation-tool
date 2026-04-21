from __future__ import annotations
from enum import Enum
from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field


class TargetBackend(str, Enum):
    UIA      = "uia"
    WIN32    = "win32"
    BROWSER  = "browser"
    OCR      = "ocr"
    IMAGE    = "image"
    FALLBACK = "fallback"



class BoundingBox(BaseModel):
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top

    @property
    def center(self) -> tuple[int, int]:
        return ((self.left + self.right) // 2, (self.top + self.bottom) // 2)

    def scale(self, factor: float) -> "BoundingBox":
        return BoundingBox(
            left=int(self.left * factor),
            top=int(self.top * factor),
            right=int(self.right * factor),
            bottom=int(self.bottom * factor),
        )


class BrowserTarget(BaseModel):
    xpath: Optional[str] = None
    css_selector: Optional[str] = None
    aria_label: Optional[str] = None
    aria_role: Optional[str] = None
    inner_text: Optional[str] = None
    tag_name: Optional[str] = None
    element_id: Optional[str] = None
    name_attr: Optional[str] = None
    placeholder: Optional[str] = None
    href: Optional[str] = None
    frame_xpath: Optional[str] = None
    tab_url: Optional[str] = None



class Selector(BaseModel):
    type: str        
    value: Any
    priority: int = 0
    confidence: float = 0.5


class UITarget(BaseModel):

    backend: TargetBackend = TargetBackend.UIA

    automation_id: Optional[str] = None
    name: Optional[str] = None
    control_type: Optional[str] = None
    class_name: Optional[str] = None
    window_title: Optional[str] = None
    process_name: Optional[str] = None
    framework_id: Optional[str] = None

    process_id: Optional[int] = None
    window_handle: Optional[int] = None

    browser: Optional[BrowserTarget] = None

    bbox: Optional[BoundingBox] = None
    screen_x: Optional[int] = None
    screen_y: Optional[int] = None

    monitor_index: int = 0
    dpi_scale: float = 1.0

 
    screenshot_ref: Optional[str] = None
    image_hash: Optional[str] = None
    template_path: Optional[str] = None

   
    ocr_text: Optional[str] = None


    ancestor_chain: List[str] = Field(default_factory=list)

   
    selectors: List[Selector] = Field(default_factory=list)

    confidence_score: float = 0.0
    confidence_reason: Optional[str] = None
    is_editable:Optional[bool] = None

    def build_selectors(self) -> None:
       
        selectors: List[Selector] = []

        if self.automation_id:
            selectors.append(Selector(
                type="automation_id",
                value=self.automation_id,
                priority=1,
                confidence=0.95
            ))

        if self.name and self.control_type:
            selectors.append(Selector(
                type="name_type",
                value={"name": self.name, "type": self.control_type},
                priority=2,
                confidence=0.85
            ))

        if self.class_name:
            selectors.append(Selector(
                type="class_name",
                value=self.class_name,
                priority=3,
                confidence=0.7
            ))

        # Browser
        if self.browser:
            if self.browser.xpath:
                selectors.append(Selector("xpath", self.browser.xpath, 1, 0.95))
            elif self.browser.css_selector:
                selectors.append(Selector("css", self.browser.css_selector, 2, 0.85))

        # Geometry fallback
        if self.bbox:
            selectors.append(Selector(
                type="bbox",
                value=self.bbox.model_dump(),
                priority=5,
                confidence=0.5
            ))

        # OCR fallback
        if self.ocr_text:
            selectors.append(Selector(
                type="ocr",
                value=self.ocr_text,
                priority=6,
                confidence=0.4
            ))

        # Image fallback
        if self.template_path:
            selectors.append(Selector(
                type="image",
                value=self.template_path,
                priority=7,
                confidence=0.3
            ))

        # Sort by priority
        self.selectors = sorted(selectors, key=lambda s: s.priority)

        if selectors:
            best = selectors[0]
            self.confidence_score = best.confidence
            self.confidence_reason = best.type



    def best_selector(self) -> Optional[Selector]:
        if not self.selectors:
            return None
        return sorted(self.selectors, key=lambda s: (s.priority, -s.confidence))[0]



    def is_browser_element(self) -> bool:
        return self.backend == TargetBackend.BROWSER

    

    def has_visual_fallback(self) -> bool:
        return bool(self.template_path or self.image_hash)



    def has_ocr_fallback(self) -> bool:
        return bool(self.ocr_text)


    def debug_summary(self) -> str:
        parts = []

        if self.backend:
            parts.append(f"backend={self.backend}")

        if self.process_name:
            parts.append(f"app={self.process_name}")

        if self.window_title:
            parts.append(f"window='{self.window_title[:30]}'")

        if self.control_type:
            parts.append(f"type={self.control_type}")

        if self.name:
            parts.append(f"name='{self.name[:30]}'")

        if self.automation_id:
            parts.append(f"id={self.automation_id}")

        if self.confidence_score:
            parts.append(f"conf={round(self.confidence_score, 2)}")

        return " | ".join(parts) if parts else "(unknown)"