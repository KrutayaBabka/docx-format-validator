"""
config.py

Configuration module for report checking project.
Contains global constants and platform-specific environment setup.
"""

import sys
import os


# -----------------------------
# Target font for the project
# -----------------------------
TARGET_FONT = "Times New Roman"


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
