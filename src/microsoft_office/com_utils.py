"""Shared COM automation utilities for Microsoft Office control."""

import os

import pythoncom
import win32com.client


def get_or_create_app(prog_id: str, visible: bool = True):
    """Get a running Office app instance or create a new one.

    Args:
        prog_id: COM ProgID (e.g., "PowerPoint.Application")
        visible: Whether to make the app visible if creating new
    """
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject(prog_id)
    except Exception:
        app = win32com.client.Dispatch(prog_id)
        app.Visible = visible
    return app


def ensure_absolute_path(path: str) -> str:
    """Convert to absolute path (COM requires absolute paths)."""
    return os.path.abspath(path)


def rgb(r: int, g: int, b: int) -> int:
    """Convert RGB to COM color value (BGR format used by Office COM)."""
    return r + (g << 8) + (b << 16)
