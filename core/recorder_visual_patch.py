from __future__ import annotations

import time
from pathlib import Path
from typing import Optional, TYPE_CHECKING

from utils.logger import logger

if TYPE_CHECKING:
    from models.target import UITarget
    from .screenshot import ScreenCapture

ELEMENT_CROP_MARGIN = 8

ELEMENT_CROP_MIN_PX = 12

# FIX: raised from 600×400 → 800×500.  Compose windows, email bodies, and
# file-picker dialogs routinely exceed 600px wide; the old cap caused the crop
# to be silently skipped, leaving no screenshot_ref for those targets.
ELEMENT_CROP_MAX_W  = 800
ELEMENT_CROP_MAX_H  = 500

ELEMENT_CROP_PREFIX = "elem"


def _capture_element_crop(
    capture: "ScreenCapture",
    target: "UITarget",
    scr_dir: Path,
) -> Optional[str]:
    
    bbox = getattr(target, "bbox", None) or getattr(target, "raw_bbox", None)
    if not bbox:
        return None

    try:
        from .screenshot import ScreenCapture as SC
        dpi_scale = getattr(target, "dpi_scale", None) or 1.0

        # Convert logical to physical pixels
        left   = int(bbox.left   * dpi_scale)
        top    = int(bbox.top    * dpi_scale)
        right  = int(bbox.right  * dpi_scale)
        bottom = int(bbox.bottom * dpi_scale)

        width  = right  - left
        height = bottom - top

        if width  < ELEMENT_CROP_MIN_PX or height < ELEMENT_CROP_MIN_PX:
            return None
        if width  > ELEMENT_CROP_MAX_W  or height > ELEMENT_CROP_MAX_H:
            return None

        # Add margin
        left   = max(0, left   - ELEMENT_CROP_MARGIN)
        top    = max(0, top    - ELEMENT_CROP_MARGIN)
        width  = width  + ELEMENT_CROP_MARGIN * 2
        height = height + ELEMENT_CROP_MARGIN * 2

        ts    = int(time.time() * 1000)
        label = f"{ELEMENT_CROP_PREFIX}_{ts}"
        path  = capture.capture_element(left, top, width, height, label=label)
        if path and path.exists():
            return path.name
    except Exception as exc:
        logger.debug("[RECORDER] element crop failed: {}", exc)
    return None


def _enrich_target_with_visual(
    capture: Optional["ScreenCapture"],
    target: Optional["UITarget"],
    scr_dir: Optional[Path],
) -> None:
    
    if not capture or not target or not scr_dir:
        return
    if getattr(target, "screenshot_ref", None):
        return   # already captured

    bbox = getattr(target, "bbox", None)
    if not bbox:
        return

    # Skip containers and shell elements — not useful for visual matching
    ctrl = (target.control_type or "").lower()
    if ctrl in ("pane", "window", "group", "toolbar", "statusbar", "scrollbar"):
        return

    ref = _capture_element_crop(capture, target, scr_dir)
    if ref:
        try:
            target.screenshot_ref = ref
            logger.debug("[RECORDER] screenshot_ref={} attached to target name='{}'",
                         ref, (target.name or "")[:30])
        except Exception:
            pass   # target may be frozen/pydantic-strict


def apply_visual_patch(recorder_class) -> None:
    
    _original_maybe_screenshot = recorder_class._maybe_screenshot
    _original_on_click         = recorder_class._on_mouse_click

    def _patched_maybe_screenshot(self, target=None) -> None:
        
        _original_maybe_screenshot(self, target)

        if self._capture and target and self._scr_dir:
            _enrich_target_with_visual(self._capture, target, self._scr_dir)

    def _patched_on_mouse_click(self, x: int, y: int, button, pressed: bool) -> None:
        # FIX: call original first so _last_target is updated before we inspect it.
        _original_on_click(self, x, y, button, pressed)

        if pressed:
            return
        if not (self._capture and self._last_target and self._scr_dir):
            return
        if getattr(self._last_target, "screenshot_ref", None):
            # Already captured by _patched_maybe_screenshot — skip.
            return
        _enrich_target_with_visual(
            self._capture, self._last_target, self._scr_dir
        )

    recorder_class._maybe_screenshot = _patched_maybe_screenshot
    recorder_class._on_mouse_click   = _patched_on_mouse_click

    logger.info("[RECORDER] Visual patch applied — element crops will be saved for visual matching")