# General utilities: resource_path() for PyInstaller, Unblock-File subprocess helper, and screen-dimension Tk init.

import os
import subprocess
import sys
import tkinter as tk


def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_screen_dimensions():
    """Return (screen_width, screen_height) using a temporary Tk root."""
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.destroy()
    return screen_width, screen_height


def unblock_file(filepath, warn_prefix="warn"):
    """Remove Windows 'Zone.Identifier' metadata (Protected View trigger)."""
    try:
        if os.name == "nt":  # only on Windows
            subprocess.run(
                [
                    "powershell",
                    "-Command",
                    f'Unblock-File -Path "{filepath}"'
                ],
                shell=True
            )
    except Exception as e:
        print(f"[{warn_prefix}] Could not unblock file: {e}")
