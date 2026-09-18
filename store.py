"""store.py — persistence layer for the excel_line memory provider.

INTRODUCTION
    This module owns all Excel read/write for excel_line. It is the single
    source of truth for durable memory: it appends brief rows to a MASTER index
    and detailed records to per-zone workbooks, and answers search/read queries.
    It contains no agent-runtime imports, so it can run standalone (unit tests,
    the sub-agent worker) without loading the full provider.
"""
from __future__ import annotations

import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)

from excel_line_core.store import *  # noqa: F401,F403,E402
from excel_line_core.store import (  # noqa: F401,E402
    ExcelLineStore, _now, MASTER_NAME, INDEX_COLS, MEM_COLS, ZONE_DEFAULTS,
)