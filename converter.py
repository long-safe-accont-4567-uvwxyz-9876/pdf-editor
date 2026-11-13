"""文档格式转换 —— PDF 与 Word 互转。

PDF → Word：使用 pdf2docx（保留段落、图片、表格与近似排版）。
Word → PDF：调用本机已安装的 Word / WPS 通过 COM 导出，保真度最高，
            不额外引入运行时依赖；两者都没装时给出明确提示。

两条路径都在本机完成，不上传任何内容。
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time

WORD_EXT = (".doc", ".docx", ".rtf", ".odt")


def _ensure_parent(path: str) -> None:
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)


def pdf_to_word(src: str, dst: str) -> dict:
    """PDF → docx。"""
    t0 = time.time()
    try:
        from pdf2docx import Converter
    except ImportError as e:
        raise RuntimeError("缺少 PDF 转 Word 组件（pdf2docx），请重新安装完整版本") from e

    _ensure_parent(dst)
    cv = Converter(src)
    try:
        cv.convert(dst)
    finally:
        try:
            cv.close()
        except Exception:
            pass

    if not os.path.exists(dst) or os.path.getsize(dst) == 0:
        raise RuntimeError("转换完成但未生成文件，可能是扫描版 PDF（纯图片，无文字层）")
    return {"ok": True, "path": dst, "seconds": round(time.time() - t0, 1)}


_PS_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$src = $env:XD_CONV_SRC
$dst = $env:XD_CONV_DST
$app = $null
foreach ($id in @('Word.Application', 'KWPS.Application', 'WPS.Application')) {
  try {
    $app = New-Object -ComObject $id -ErrorAction Stop
    break
  } catch { }
}
if ($null -eq $app) { exit 2 }
try {
  try { $app.Visible = $false } catch { }
  try { $app.DisplayAlerts = 0 } catch { }
  # 3 = msoAutomationSecurityForceDisable，禁止文档里的宏执行
  try { $app.AutomationSecurity = 3 } catch { }
  $doc = $app.Documents.Open($src, $false, $true)
  try {
    $doc.SaveAs2($dst, 17)
  } catch {
    $doc.SaveAs($dst, 17)
  }
  $doc.Close($false)
} finally {
  try { $app.Quit() } catch { }
}
"""


def _office_available() -> bool:
    if os.name != "nt":
        return False
    roots = [
        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
    ]
    if any(os.path.exists(p) for p in roots):
        return True
    try:
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe")
                return True
            except OSError:
                pass
    except Exception:
        pass
    # WPS 的注册表键
    try:
        import winreg
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for sub in (r"SOFTWARE\Kingsoft\Office", r"SOFTWARE\WOW6432Node\Kingsoft\Office"):
                try:
                    winreg.OpenKey(hive, sub)
                    return True
                except OSError:
                    pass
    except Exception:
        pass
    return False


def word_to_pdf(src: str, dst: str) -> dict:
    """Word → PDF，走本机 Office 的 COM 导出。"""
    if os.name != "nt":
        raise RuntimeError("仅支持 Windows")
    if not src.lower().endswith(WORD_EXT):
        raise RuntimeError("请选择 Word 文档（doc / docx / rtf / odt）")

    t0 = time.time()
    _ensure_parent(dst)

    script_path = os.path.join(tempfile.gettempdir(), "pdfeditor_conv.ps1")
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(_PS_SCRIPT)

    env = os.environ.copy()
    env["XD_CONV_SRC"] = os.path.abspath(src)
    env["XD_CONV_DST"] = os.path.abspath(dst)

    creationflags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    try:
        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-File", script_path],
            capture_output=True, text=True, timeout=240,
            creationflags=creationflags, env=env,
        )
    except subprocess.TimeoutExpired as e:
        raise RuntimeError("转换超时，文档可能过大或 Office 正在弹窗等待操作") from e
    finally:
        try:
            os.remove(script_path)
        except OSError:
            pass

    if r.returncode == 2:
        raise RuntimeError("未检测到 Microsoft Word 或 WPS，无法把 Word 转成 PDF。"
                           "请先安装其中之一，或改用「另存为 PDF」")
    if r.returncode != 0 or not os.path.exists(dst) or os.path.getsize(dst) == 0:
        tail = (r.stderr or r.stdout or "").strip().splitlines()
        detail = tail[-1][:160] if tail else "未知错误"
        raise RuntimeError(f"转换失败：{detail}")

    return {"ok": True, "path": dst, "seconds": round(time.time() - t0, 1)}


def can_convert() -> dict:
    try:
        import pdf2docx  # noqa: F401
        to_word = True
    except Exception:
        to_word = False
    return {"pdf2word": to_word, "word2pdf": _office_available()}
