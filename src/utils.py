import sys
import os
from pathlib import Path

def get_resource_path(relative_path: str) -> Path:
    """
    Get absolute path to resource, works for dev and for PyInstaller
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = Path(sys._MEIPASS)
    else:
        # In development, resources are in the project root
        # Assuming this file is in src/utils.py, root is parent of parent
        base_path = Path(__file__).parent.parent
        
    return base_path / relative_path
