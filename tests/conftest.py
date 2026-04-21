"""
pytest configuration — adds project root to sys.path so imports resolve.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
