"""文档格式转换 —— PDF 与 Word 互转，全部内置，不调用任何外部程序。

PDF → Word：pdf_to_word.py，只依赖 PyMuPDF，直接拼 OOXML
Word → PDF：word_to_pdf.py，python-docx 读结构 + ReportLab 排版

两个方向都不依赖 Microsoft Word / WPS / LibreOffice，也不联网。
"""

from __future__ import annotations

import os
import time

WORD_EXT = (".docx",)


def _ensure_parent(path: str) -> None:
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)


def pdf_to_word_fn(src: str, dst: str, progress=None) -> dict:
    """PDF → docx，走内置转换器。"""
    import pdf_to_word as engine  # 局部导入：避免与本函数同名遮蔽

    if not os.path.exists(src):
        raise RuntimeError("文件不存在：" + src)
    t0 = time.time()
    _ensure_parent(dst)
    r = engine.convert(src, dst, progress=progress)
    return {
        "ok": True,
        "path": r.get("path", dst),
        "seconds": round(time.time() - t0, 1),
        "pages": r.get("pages", 0),
        "chars": r.get("chars", 0),
        "images": r.get("images", 0),
        # 没有文字层 → 转换结果是整页图片，前端需要提示用户
        "scanned": bool(r.get("scanned")),
    }


def word_to_pdf_fn(src: str, dst: str, progress=None) -> dict:
    """Word → PDF，走内置渲染引擎，不需要装 Office。"""
    import word_to_pdf as engine

    if not os.path.exists(src):
        raise RuntimeError("文件不存在：" + src)
    # .doc / .rtf 是旧二进制或富文本封装，python-docx 解析不了
    if not src.lower().endswith(".docx"):
        raise RuntimeError("内置转换器只支持 .docx，请先把文档另存为 .docx 再转换")

    t0 = time.time()
    _ensure_parent(dst)
    try:
        r = engine.convert(src, dst, progress=progress)
    except Exception as e:
        # 把库的英文异常翻译成用户能看懂的话
        msg = str(e)
        if "Package not found" in msg or "not a zip" in msg.lower() or "BadZip" in msg:
            raise RuntimeError(
                "这个文件不是有效的 .docx（可能是改过扩展名的旧 .doc，或文件已损坏）。"
                "请用 Word 打开后另存为 .docx 再试。"
            ) from e
        raise
    return {
        "ok": True,
        "path": r.get("path", dst),
        "seconds": round(time.time() - t0, 1),
        "pages": r.get("pages", 0),
        "chars": r.get("chars", 0),
        "images": r.get("images", 0),
        "scanned": False,
    }


def can_convert() -> dict:
    """两个方向都是内置实现，始终可用。"""
    return {
        "pdf2word": True,
        "word2pdf": True,
        "word2pdf_exts": [".docx"],
        "engine": "builtin",
    }


# 兼容旧调用名
pdf_to_word = pdf_to_word_fn
word_to_pdf = word_to_pdf_fn
