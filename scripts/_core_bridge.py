"""Bridge so `python -m brain_server` and sibling scripts resolve the
excel_line_core package when running from a source checkout.

Importable from anywhere; carries no side effects. In an installed wheel the
package already lives on sys.path and this module is a no-op.
"""
from __future__ import annotations

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))            # .../scripts
_PLUGIN = os.path.dirname(_HERE)                              # excel_line plugin root

if os.path.isdir(os.path.join(_PLUGIN, "excel_line_core")):
    _CORE = os.path.join(_PLUGIN, "excel_line_core")
    if _CORE not in sys.path:
        sys.path.insert(0, _CORE)
    if _PLUGIN not in sys.path:
        sys.path.insert(0, _PLUGIN)
elif _PLUGIN not in sys.path:
    sys.path.insert(0, _PLUGIN)