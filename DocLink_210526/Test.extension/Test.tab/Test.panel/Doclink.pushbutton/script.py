# -*- coding: utf-8 -*-
"""
DocLink Import Manager + Excel Schedule Importer

Tab 1 – Import Excel ranges, Word pages, PDFs, or Images into the active Revit view
Tab 2 – Import Excel print areas into Revit Generic Model schedule headers

Both tabs track imports persistently per project file.
Re-open the tool to see previous imports and update from source.
"""

__title__ = "DocLink"
__doc__   = __doc__

# ── Add modules/ to sys.path so all module imports resolve ───────────────────
import os
import sys

_modules_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "modules")
if _modules_dir not in sys.path:
    sys.path.insert(0, _modules_dir)

# ── Entry point ───────────────────────────────────────────────────────────────
from main_window import main

if __name__ == "__main__":
    main()
