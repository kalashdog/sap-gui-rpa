"""
Entry point for the GUI application.

The GUI has been reorganized into the gui/ package for better maintainability.
This file is kept as the PyInstaller entry point and for backwards compatibility.
"""
from gui.app import RpaGUI

if __name__ == "__main__":
    app = RpaGUI()
    app.mainloop()