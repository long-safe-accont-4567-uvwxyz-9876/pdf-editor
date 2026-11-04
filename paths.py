"""统一路径解析：源码运行与 PyInstaller 打包后行为一致。"""

from __future__ import annotations

import os
import sys


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _data_dir() -> str:
    """只读资源目录（static / assets / 内置脚本）。"""
    if _is_frozen():
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _user_dir() -> str:
    """可写的用户数据目录。

    安装到 Program Files 后自身目录不可写，必须落到 LOCALAPPDATA。
    """
    if _is_frozen():
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(base, "PDFEditor")
    return os.path.dirname(os.path.abspath(__file__))


DATA_DIR = _data_dir()
USER_DIR = _user_dir()
WORK_DIR = os.path.join(USER_DIR, "work")
LOG_PATH = os.path.join(USER_DIR, "app.log")

os.makedirs(WORK_DIR, exist_ok=True)


def setup_logging() -> None:
    """窗口模式下没有控制台，把 stdout/stderr 重定向到日志文件。"""
    if not _is_frozen():
        return
    try:
        if sys.stdout is not None and sys.stderr is not None:
            return
    except Exception:
        pass
    try:
        f = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
        sys.stdout = f
        sys.stderr = f
    except Exception:
        pass
