"""Tiny JSON settings store (remembers the output folder between runs)."""
from __future__ import annotations

import json
from pathlib import Path

PATH = Path.home() / ".checkrun" / "settings.json"


def load() -> dict:
    try:
        return json.loads(PATH.read_text())
    except (OSError, ValueError):
        return {}


def save(data: dict) -> None:
    try:
        PATH.parent.mkdir(parents=True, exist_ok=True)
        PATH.write_text(json.dumps(data, indent=2))
    except OSError:
        pass  # settings are a convenience; never block the app on them
