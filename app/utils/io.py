"""Helpers de I/O. Lectura de archivos compartidos en Windows."""
from __future__ import annotations

from pathlib import Path


def read_file_shared(path: Path) -> bytes:
    """Lee un archivo aunque esté abierto por Excel/OneDrive (FILE_SHARE_*)."""
    import win32con
    import win32file

    handle = win32file.CreateFile(
        str(path),
        win32con.GENERIC_READ,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        0,
        None,
    )
    try:
        chunks: list[bytes] = []
        while True:
            err, data = win32file.ReadFile(handle, 1024 * 1024)
            if not data:
                break
            chunks.append(bytes(data))
        return b"".join(chunks)
    finally:
        handle.Close()
