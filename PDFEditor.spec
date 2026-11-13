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
# 注意：PDF→Word 走内置 pdf_to_word.py（仅依赖 pymupdf），
# 不再需要 pdf2docx / cv2 / docx，别把它们塞进包里白占 60MB
for pkg in ("webview", "pythonnet", "clr_loader"):
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

# PyMuPDF 2.1 自带一套 AI 版面分析模型（onnx，约 30MB），本编辑器只做文字/图片
# 的提取与写入，完全用不到它，排除掉能显著减小安装包
a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=[(p, dst) for (p, dst) in datas if "onnx" not in p.lower()],
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "tkinter", "matplotlib", "pandas", "scipy",
        "PyQt5", "PySide2", "PySide6", "IPython", "pytest",
        "pymupdf.pymupdf_layout", "pymupdf.layout",
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
