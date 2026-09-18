"""brain_store.py — excel_line v2: hierarchical tree store, max 10 rows per .xlsx file.

The implementation has moved to excel_line_core.brain_store. This module remains
as a backward-compatible import shim for Hermes and older tests that load
`excel_line.brain_store` or `brain_store` directly.
"""
from __future__ import annotations

import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)

from excel_line_core.brain_store import *  # noqa: F401,F403,E402
from excel_line_core.brain_store import (  # noqa: F401,E402
    BadBranch, BrainStore, FullError, MAX_ROWS, MASTER_V2, TITLE_CAP, CONTENT_CAP,
    V2_HDR, _flatten_rows, search_index,
)