# -*- coding: utf-8 -*-
"""PyInstaller launcher entry point (run from repo root)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx2pdf.cli import main

if __name__ == "__main__":
    raise SystemExit(main())