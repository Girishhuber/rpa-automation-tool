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

        return self._capture_box(left, top, width, height, label)

    def capture_region(self, left: int, top: int, width: int, height: int,
                       suffix: str = "region") -> Optional[Path]:
        return self._capture_box(left, top, width, height, suffix)

    def _capture_box(self, left: int, top: int, width: int, height: int,
                     label: str) -> Optional[Path]:
        if not MSS_OK or width <= 0 or height <= 0:
            return None
        try:
            with mss.mss() as sct:
                ts   = int(time.time() * 1000)
                path = self.output_dir / f"{label}_{ts}.{self.fmt}"
                img  = sct.grab({"left": left, "top": top,
                                  "width": width, "height": height})
                mss.tools.to_png(img.rgb, img.size, output=str(path))
                logger.debug("[SCREEN] Crop ({},{} {}x{}) → {}",
                             left, top, width, height, path.name)
                return path
        except Exception as exc:
            logger.debug("[SCREEN] capture_box failed: {}", exc)
            return None

    # ── Perceptual hashing ────────────────────────────────────────────

    def visual_hash(self, path: Path, size: int = 16) -> Optional[str]:
     
        if not PIL_OK or not path or not path.exists():
            return None
        try:
            img    = Image.open(path).convert("L").resize((size, size), Image.LANCZOS)
            pixels = list(img.getdata())
            avg    = sum(pixels) / len(pixels)
            bits   = "".join("1" if p >= avg else "0" for p in pixels)
            value  = int(bits, 2)
            return f"{value:016x}"
        except Exception as exc:
            logger.debug("[SCREEN] visual_hash failed: {}", exc)
            return None

    def similarity(self, hash1: Optional[str], hash2: Optional[str]) -> float:
        if not hash1 or not hash2 or len(hash1) != 16 or len(hash2) != 16:
            return 0.0
        if hash1 == hash2:
            return 1.0
        try:
            v1 = int(hash1, 16)
            v2 = int(hash2, 16)
            diff = v1 ^ v2
            hamming = bin(diff).count("1")   # 0–64 differing bits
            return max(0.0, min(1.0, 1.0 - hamming / 64.0))
        except Exception:
            return 0.0

    def compare_visual_hash(self, hash1: Optional[str], hash2: Optional[str]) -> bool:
        # FIX: threshold lowered from 0.90 → 0.85.
        # At full-screen pHash (16×16 = 256 bits mapped to 64-bit hex), even a caret
        # blink or a scroll-bar thumb shift can consume 4-6 Hamming bits, dropping
        # similarity below 0.90 and triggering false "UI did not change" verdicts.
        # 0.85 tolerates minor chrome redraws while still catching real navigations.
        return self.similarity(hash1, hash2) >= 0.85

    @staticmethod
    def monitor_info() -> list[dict]:
        if not MSS_OK:
            return [{"width": 1920, "height": 1080, "index": 0,
                     "left": 0, "top": 0}]
        with mss.mss() as sct:
            return [
                {
                    "index":  i,
                    "left":   m["left"],
                    "top":    m["top"],
                    "width":  m["width"],
                    "height": m["height"],
                }
                for i, m in enumerate(sct.monitors[1:])
            ]