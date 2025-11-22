"""
config/config.py

Configuration module for report checking project.
Contains global constants and platform-specific environment setup.
"""

import sys
import os
from typing import Dict


# -----------------------------
# Target font for the project
# -----------------------------
TARGET_FONT = "Times New Roman"
ReportItem = Dict[str, object]   # {"run": Run, "paragraph_text": str, "reason": str}

# Font size constraints
MIN_FONT_SIZE = 12
MAX_FONT_SIZE = 14

# Pattern to detect title page (e.g., "Москва 2025 г.")
TITLE_PAGE_PATTERN = r"Москва\s+\d{4}\s+г\."

# Margin requirements in cm
TOP_MARGIN_CM = 2
BOTTOM_MARGIN_CM = 2
LEFT_MARGIN_CM = 2
RIGHT_MARGIN_CM = 1

# Indentation constants (in cm)
FIRST_LINE_INDENT_CM = 1.25
LEFT_INDENT_CM = 0
RIGHT_INDENT_CM = 0

# -----------------------------
# Configure Tcl/Tk environment on Windows
# -----------------------------
if sys.platform == "win32":
    # Set environment variables for Tcl/Tk
    os.environ["TCL_LIBRARY"] = r"C:\Users\user\AppData\Local\Programs\Python\Python313\tcl\tcl8.6"
    os.environ["TK_LIBRARY"] = r"C:\Users\user\AppData\Local\Programs\Python\Python313\tcl\tk8.6"

    # Enable High DPI awareness
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        # Fail silently if DPI setting fails
        pass
