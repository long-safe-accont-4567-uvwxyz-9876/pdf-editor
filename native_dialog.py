"""Windows 原生文件对话框（ctypes 直调 comdlg32，不依赖 tkinter）。

当前运行环境的 Python 未编译 tkinter，因此改用 Win32 API。
作为独立进程运行，避免与 Web 服务线程冲突。
"""

from __future__ import annotations

import ctypes
import os
import sys
from ctypes import wintypes

OFN_READONLY = 0x00000001
OFN_OVERWRITEPROMPT = 0x00000002
OFN_NOCHANGEDIR = 0x00000008
OFN_PATHMUSTEXIST = 0x00000800
OFN_FILEMUSTEXIST = 0x00001000
OFN_EXPLORER = 0x00080000
OFN_DONTADDTORECENT = 0x02000000


class OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hInstance", wintypes.HINSTANCE),
        ("lpstrFilter", wintypes.LPCWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter", wintypes.DWORD),
        ("nFilterIndex", wintypes.DWORD),
        ("lpstrFile", wintypes.LPWSTR),
        ("nMaxFile", wintypes.DWORD),
        ("lpstrFileTitle", wintypes.LPWSTR),
        ("nMaxFileTitle", wintypes.DWORD),
        ("lpstrInitialDir", wintypes.LPCWSTR),
        ("lpstrTitle", wintypes.LPCWSTR),
        ("Flags", wintypes.DWORD),
        ("nFileOffset", wintypes.WORD),
        ("nFileExtension", wintypes.WORD),
        ("lpstrDefExt", wintypes.LPCWSTR),
        ("lCustData", wintypes.LPARAM),
        ("lpfnHook", wintypes.LPVOID),
        ("lpTemplateName", wintypes.LPCWSTR),
        ("pvReserved", wintypes.LPVOID),
        ("dwReserved", wintypes.DWORD),
        ("FlagsEx", wintypes.DWORD),
    ]


def _run(mode: str, default_name: str, initial_dir: str,
         filter_str: str = "", title: str = "", default_ext: str = "pdf") -> str:
    if os.name != "nt":
        return ""

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    buf = ctypes.create_unicode_buffer(32768)
    if default_name:
        buf.value = default_name

    # 过滤器：双 \0 结尾
    flt = filter_str or "PDF 文件 (*.pdf)\0*.pdf\0所有文件 (*.*)\0*.*\0\0"

    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.hwndOwner = None
    ofn.lpstrFilter = flt
    ofn.nFilterIndex = 1
    ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
    ofn.nMaxFile = 32768
    ofn.lpstrInitialDir = initial_dir or None
    ofn.lpstrDefExt = default_ext or "pdf"

    if mode == "save":
        ofn.lpstrTitle = title or "导出 PDF"
        ofn.Flags = (OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR
                     | OFN_EXPLORER | OFN_DONTADDTORECENT)
        fn = ctypes.windll.comdlg32.GetSaveFileNameW
    else:
        ofn.lpstrTitle = title or "打开 PDF"
        ofn.Flags = (OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR
                     | OFN_EXPLORER | OFN_DONTADDTORECENT)
        fn = ctypes.windll.comdlg32.GetOpenFileNameW

    fn.restype = wintypes.BOOL
    fn.argtypes = [ctypes.POINTER(OPENFILENAMEW)]

    # 让对话框尽量出现在最前
    try:
        ctypes.windll.user32.SetForegroundWindow(ctypes.windll.kernel32.GetConsoleWindow())
    except Exception:
        pass

    if not fn(ctypes.byref(ofn)):
        return ""
    return buf.value


def main() -> None:
    mode = sys.argv[1] if len(sys.argv) > 1 else "open"
    default_name = sys.argv[2] if len(sys.argv) > 2 else ""
    initial_dir = sys.argv[3] if len(sys.argv) > 3 else ""
    filter_str = sys.argv[4] if len(sys.argv) > 4 else ""
    title = sys.argv[5] if len(sys.argv) > 5 else ""
    default_ext = sys.argv[6] if len(sys.argv) > 6 else "pdf"
    try:
        out = _run(mode, default_name, initial_dir, filter_str, title, default_ext)
    except Exception:
        out = ""
    sys.stdout.write(out or "")


if __name__ == "__main__":
    main()
