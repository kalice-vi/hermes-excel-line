"""worker.py — background indexer (sub-agent) for the excel_line memory provider.

The implementation has moved to excel_line_core.worker. This module remains
as a backward-compatible import shim for Hermes and older tests.
"""
from __future__ import annotations

import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)

from excel_line_core.worker import *  # noqa: F401,F403,E402
from excel_line_core.worker import (  # noqa: F401,E402
    _CLASSIFY_PROMPT, _apply, _classify, _classify_tristate, _read_records,
    index_while_logs_present, process_logs,
)