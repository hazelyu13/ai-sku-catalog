import os
import base64
import platform
import subprocess
from io import BytesIO

from PIL import Image
import streamlit as st


def open_excel_file(excel_path: str) -> None:
    """Open the Excel file using the OS default program."""
    try:
        system = platform.system()
        if system == "Darwin":
            subprocess.call(["open", excel_path])
        elif system == "Windows":
            os.startfile(excel_path)  # type: ignore[attr-defined]
        else:
            subprocess.call(["xdg-open", excel_path])
        st.success("📂 Excel file opened!")
    except Exception as e:
        st.error(f"❌ Could not open Excel file.\n\n{e}")


def open_folder(path: str) -> None:
    """Open a folder in Finder / Explorer."""
    try:
        system = platform.system()
        if system == "Darwin":
            subprocess.call(["open", path])
        elif system == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.call(["xdg-open", path])
    except Exception as e:
        st.error(f"❌ Could not open folder.\n\n{e}")


def load_image_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def normalize_value(val):
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    # Collapse spaces and uppercase
    import re

    s = re.sub(r"\s+", " ", s)
    return s.upper()


def rotate_image_filelike(file_obj, degrees: int):
    """
    Rotate an uploaded image file-like object and return BytesIO of rotated image.
    This is for preview only (doesn't overwrite original).
    """
    file_obj.seek(0)
    img = Image.open(file_obj)
    rotated = img.rotate(degrees, expand=True)
    buf = BytesIO()
    rotated.save(buf, format="PNG")
    buf.seek(0)
    return buf
