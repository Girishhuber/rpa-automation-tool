
from __future__ import annotations
from pathlib import Path
from typing import Optional
import hashlib
import time

try:
    import mss, mss.tools
    MSS_OK = True
except ImportError:
    MSS_OK = False

try:
    from PIL import Image
    PIL_OK = True
except ImportError:
    PIL_OK = False

from utils.logger import logger


class ScreenCapture:
    def __init__(self, output_dir: Path, fmt: str = "png"):
        self.output_dir = output_dir
        self.fmt = fmt
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def capture_full(self, monitor_index: int = 0) -> Optional[Path]:
        if not MSS_OK:
            return None
        try:
            with mss.mss() as sct:
                mon  = sct.monitors[min(monitor_index + 1, len(sct.monitors) - 1)]
                ts   = int(time.time() * 1000)
                path = self.output_dir / f"screen_{ts}.{self.fmt}"
                img  = sct.grab(mon)
                mss.tools.to_png(img.rgb, img.size, output=str(path))
                logger.debug("[SCREEN] Full capture → {}", path.name)
                return path
        except Exception as exc:
            logger.debug("[SCREEN] capture_full failed: {}", exc)
            return None

    def capture_element(self, left: int, top: int, width: int, height: int,
                        label: str = "element") -> Optional[Path]:
        """Capture just the element's bounding box — used as screenshot reference."""
        if not MSS_OK or width <= 0 or height <= 0:
            return None
        try:
            with mss.mss() as sct:
                ts   = int(time.time() * 1000)
                path = self.output_dir / f"{label}_{ts}.{self.fmt}"
                img  = sct.grab({"left": left, "top": top, "width": width, "height": height})
                mss.tools.to_png(img.rgb, img.size, output=str(path))
                logger.debug("[SCREEN] Element crop ({},{} {}x{}) → {}", left, top, width, height, path.name)
                return path
        except Exception as exc:
            logger.debug("[SCREEN] capture_element failed: {}", exc)
            return None

    def capture_region(self, left: int, top: int, width: int, height: int,
                       suffix: str = "region") -> Optional[Path]:
        return self.capture_element(left, top, width, height, label=suffix)

    def visual_hash(self, path: Path, size: int = 16) -> Optional[str]:
        """
        Perceptual hash of an image — 64-char hex string.
        Resize to size×size, convert to greyscale, hash pixel values.
        Fast for change detection without full OpenCV.
        """
        if not PIL_OK or not path.exists():
            return None
        try:
            img    = Image.open(path).convert("L").resize((size, size))
            pixels = list(img.getdata())
            avg    = sum(pixels) / len(pixels)
            bits   = "".join("1" if p > avg else "0" for p in pixels)
            return hashlib.md5(bits.encode()).hexdigest()
        except Exception:
            return None

    def compare_visual_hash(self, hash1: Optional[str], hash2: Optional[str]) -> bool:
        """Return True if hashes match (UI unchanged). None inputs → assume changed."""
        if hash1 is None or hash2 is None:
            return False
        return hash1 == hash2

    @staticmethod
    def monitor_info() -> list[dict]:
        if not MSS_OK:
            return [{"width": 1920, "height": 1080, "index": 0}]
        with mss.mss() as sct:
            return [
                {"index": i, "left": m["left"], "top": m["top"],
                 "width": m["width"], "height": m["height"]}
                for i, m in enumerate(sct.monitors[1:])
            ]
