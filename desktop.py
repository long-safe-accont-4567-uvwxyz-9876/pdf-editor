"""独立桌面窗口。

优先用 pywebview（系统 WebView2，独立任务栏图标）；
未安装时退回 Edge / Chrome 的 --app 模式，同样是独立窗口而非浏览器标签页。
"""

from __future__ import annotations

import os
import subprocess
import sys

TITLE = "PDF 编辑器"

_BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def _find_browser() -> str | None:
    for p in _BROWSERS:
        if os.path.exists(p):
            return p
    return None


_WINDOW = None


def get_window():
    """供后端调用原生文件对话框时使用。"""
    return _WINDOW


def open_window(url: str, width: int = 1440, height: int = 920) -> str:
    """打开独立窗口并阻塞到窗口关闭，返回使用的模式名。"""
    global _WINDOW
    try:
        import webview  # type: ignore

        _WINDOW = webview.create_window(
            TITLE, url,
            width=width, height=height,
            min_size=(960, 640),
            background_color="#EEF0F3",
            text_select=False,
        )
        webview.start()
        _WINDOW = None
        return "pywebview"
    except Exception:
        _WINDOW = None

    exe = _find_browser()
    if exe:
        args = [exe, f"--app={url}", f"--window-size={width},{height}",
                "--no-first-run", "--no-default-browser-check",
                f"--app-name={TITLE}"]
        try:
            proc = subprocess.Popen(args)
            proc.wait()
            return "browser-app"
        except Exception:
            pass

    import webbrowser
    webbrowser.open(url)
    return "browser-tab"


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8765/"
    print(open_window(target))
