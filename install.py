#!/usr/bin/env python3
"""Shortcut for plugins/outlook/install.py (Zoo Code and other Agent Skills hosts). Run with --help for options."""
import runpy, sys
from pathlib import Path
sys.argv[0] = str(Path(__file__).resolve().parent / "plugins" / "outlook" / "install.py")
runpy.run_path(sys.argv[0], run_name="__main__")
