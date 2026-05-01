"""Helper utilities for the UI."""

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from tkinter import messagebox

from app.config import DATE_FORMAT


def parse_positive_number(value: str, field_name: str) -> float | None:
    """Parse a strictly positive number (> 0)."""
    try:
        number = float(value)
    except ValueError:
        messagebox.showwarning("Validation", f"{field_name} must be a number.")
        return None

    if number <= 0:
        messagebox.showwarning("Validation", f"{field_name} must be greater than zero.")
        return None

    return number


def parse_non_negative_number(value: str, field_name: str) -> float | None:
    """Parse a non-negative number (>= 0)."""
    try:
        number = float(value)
    except ValueError:
        messagebox.showwarning("Validation", f"{field_name} must be a number.")
        return None

    if number < 0:
        messagebox.showwarning("Validation", f"{field_name} cannot be negative.")
        return None

    return number


def valid_date(value: str) -> bool:
    """Check if a date string matches the configured date format."""
    try:
        datetime.strptime(value, DATE_FORMAT)
        return True
    except ValueError:
        return False


def open_file(path: Path) -> None:
    """Open a file using the OS default application."""
    try:
        if sys.platform == "win32":
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=True)
        else:
            subprocess.run(["xdg-open", str(path)], check=True)
    except Exception as exc:
        messagebox.showerror("Open Failed", f"Could not open file:\n{exc}")
