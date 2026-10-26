"""Streamlit Cloud root entry point for Mega Pipes Quotation System."""

import sys
from pathlib import Path

# Ensure root directory is in sys.path
ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from streamlit_app.app import main

if __name__ == "__main__":
    main()
