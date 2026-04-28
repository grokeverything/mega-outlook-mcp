"""Compatibility baseline for Outlook field detection."""

from __future__ import annotations

import json
from importlib import resources


def load_baseline() -> dict:
    with resources.files(__package__).joinpath("outlook_baseline.json").open("r", encoding="utf-8") as fh:
        return json.load(fh)


__all__ = ["load_baseline"]
