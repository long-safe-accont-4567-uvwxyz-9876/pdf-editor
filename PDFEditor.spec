# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：产出 dist/PDFEditor/ 目录（含主程序）。"""

from PyInstaller.utils.hooks import collect_all

datas = [
    ("static", "static"),
    ("assets/icon.ico", "assets"),
]
binaries = []
hiddenimports = []

# pywebview / pythonnet 需要连同其原生组件一起收集，否则运行时报找不到 CLR
# pdf2docx / cv2 / docx 用于 PDF ↔ Word 转换，必须把数据文件与扩展模块一并带上
for pkg in ("webview", "pythonnet", "clr_loader", "pdf2docx", "docx", "cv2"):
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "tkinter", "matplotlib", "pandas", "scipy",
        "PyQt5", "PySide2", "PySide6", "IPython", "pytest",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PDFEditor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon="assets/icon.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="PDFEditor",
)
