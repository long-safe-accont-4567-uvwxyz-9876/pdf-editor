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
# 转换全部内置：PDF→Word 用 pymupdf，Word→PDF 用 python-docx + reportlab，
# 都不需要 Word / WPS / LibreOffice，也不需要 pdf2docx / cv2
for pkg in ("webview", "pythonnet", "clr_loader", "reportlab", "docx"):
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

hiddenimports += [
    "word_to_pdf", "pdf_to_word", "converter",
    "reportlab.pdfbase._fontdata_enc_winansi",
    "reportlab.graphics.barcode",
]

# PyMuPDF 2.x 的 wheel 里带了一整套 AI 版面分析模型（layout/ 目录，约 50MB：
# onnx 模型 + 字体 + libtesseract 等）。本编辑器只做文字/图片的提取与写入，
# 这些模型一个都用不到，全部剔除，否则安装包会白胖一大圈。
_SKIP_PARTS = ("layout", "onnx", "mupdf-devel")


def _keep(entry) -> bool:
    """True = 保留。entry 形如 (source, dest_in_bundle)。"""
    try:
        src, dst = str(entry[0]).replace("\\", "/").lower(), str(entry[1]).replace("\\", "/").lower()
    except Exception:
        return True
    parts = (src + "/" + dst).split("/")
    return not any(p in parts for p in _SKIP_PARTS)


a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=[e for e in datas if _keep(e)],
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "tkinter", "matplotlib", "pandas", "scipy",
        "PyQt5", "PySide2", "PySide6", "IPython", "pytest",
        "pymupdf.pymupdf_layout", "pymupdf.layout",
        "pymupdf._mupdf_layout", "onnxruntime",
    ],
    noarchive=False,
)

# Analysis 之后再把 hook 自动塞进来的 layout 资源清一遍，
# PyMuPDF 的 hook 优先级高于 datas 过滤，只在 datas 里剔除不够
a.datas = [e for e in a.datas if _keep(e)]
a.binaries = [e for e in a.binaries if _keep(e)]

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
