"""Pornire fără consolă: dublu-click (Windows asociază .pyw cu pythonw.exe)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pii_gui.app import main  # noqa: E402

main()
