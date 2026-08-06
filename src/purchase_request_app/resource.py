from __future__ import annotations

import sys
from pathlib import Path


def project_root() -> Path:
    """Return the source project root or PyInstaller temporary bundle root."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parents[2]


def resource_path(relative_path: str | Path) -> Path:
    return project_root() / Path(relative_path)
