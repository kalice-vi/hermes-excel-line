"""Connector between excel_line flat-module layout and the excel_line_core package.

The repo is migrating from flat root modules (store.py, brain_store.py, worker.py)
to a pip-installable src layout where the same logic lives in
excel_line_core/{store,brain_store,worker}.py and the Hermes adapter lives in
excel_line_core/adapters/hermes.py (harness-agnostic core; adapter keeps the only
Hermes imports). This file is loaded by the repo's __init__.py (Hermes entry)
and by tests, and re-exports the new package so old import paths still work.
"""
from __future__ import annotations

import os
import sys

_ROOT = os.path.dirname(os.path.abspath(__file__))

def _ensure_pkg_importable() -> None:
    """Make `import excel_line_core` work even when pip install was not run."""
    pkg_dir = os.path.join(_ROOT, "excel_line_core")
    if os.path.isdir(pkg_dir) and _ROOT not in sys.path:
        sys.path.insert(0, _ROOT)

_ensure_pkg_importable()