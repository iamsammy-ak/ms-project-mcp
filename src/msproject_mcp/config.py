"""Environment-driven settings. Read once at import time."""

from __future__ import annotations

import os


def _flag(name: str, default: bool = False) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _number(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


# Launching Project visibly means the user can watch what the agent is doing.
VISIBLE = _flag("MSPROJECT_MCP_VISIBLE", True)

# Arbitrary VBA execution is opt-in: it needs "Trust access to the VBA project
# object model" enabled in Project's Trust Center.
ALLOW_VBA = _flag("MSPROJECT_MCP_ALLOW_VBA", False)

# Default timeout for a single COM round-trip, in seconds. Leveling and
# recalculating very large projects can take a while.
CALL_TIMEOUT = _number("MSPROJECT_MCP_TIMEOUT", 180.0)

# Project to operate on when more than one file is open.
DEFAULT_PROJECT = os.environ.get("MSPROJECT_MCP_PROJECT", "").strip()

# Cap on rows returned by list tools, so a 40k-line schedule cannot flood a model.
MAX_ROWS = int(_number("MSPROJECT_MCP_MAX_ROWS", 1000))
