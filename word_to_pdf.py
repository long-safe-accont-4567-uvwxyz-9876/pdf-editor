"""Word（.docx）→ PDF —— 完全内置渲染，不调用 Word / WPS / LibreOffice。

流程：
  1. 用 python-docx 读 docx 的结构（段落、表格、图片），保持原顺序
  2. 解析样式继承链，算出每段最终的字号 / 粗斜体 / 颜色 / 对齐 / 缩进 / 行距
  3. 用 ReportLab 的 Platypus 排版成 PDF（自动分页、表格跨页、图片等比缩放）

设计取舍：这是「忠实还原」而不是「像素复刻」。能做到段落样式、行内样式、
对齐、缩进、行距、表格、图片、列表、分页、页眉页脚。做不到 Word 的
复杂图文环绕、文本框、艺术字、宏表单——这些需要真正的排版引擎。
"""

from __future__ import annotations

import io
import os
import re
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, Image, PageBreak, PageTemplate,
    Paragraph as RLParagraph, Spacer, Table as RLTable, TableStyle,
)

FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")

# ---------------------------------------------------------------- 字体

# (字体族名, 常规文件, 粗体文件, 斜体文件, 粗斜体文件)
_FAMILIES = [
    ("雅黑", "msyh.ttc", "msyhbd.ttc", None, None),
    ("宋体", "simsun.ttc", "simsun.ttc", "simsun.ttc", "simsun.ttc"),
    ("黑体", "simhei.ttf", "simhei.ttf", "simhei.ttf", "simhei.ttf"),
    ("楷体", "simkai.ttf", "simkai.ttf", "simkai.ttf", "simkai.ttf"),
    ("仿宋", "simfang.ttf", "simfang.ttf", "simfang.ttf", "simfang.ttf"),
    ("等线", "Deng.ttf", "Dengb.ttf", "Deng.ttf", "Dengb.ttf"),
]

# 字体名（docx 里写的）→ 我们注册的族名
_NAME_MAP = [
    (("msyh", "yahei", "雅黑"), "雅黑"),
    (("simsun", "song", "宋体", "nsimsun"), "宋体"),
    (("simhei", "hei", "黑体"), "黑体"),
    (("simkai", "kai", "楷体"), "楷体"),
    (("simfang", "fang", "仿宋"), "仿宋"),
    (("deng", "等线"), "等线"),
]

_registered: dict[str, str] = {}
_ready = False


def _register_fonts() -> dict[str, str]:
    """注册中文字体，返回 {族名: 可用的常规字体名}。"""
    global _ready, _registered
    if _ready:
        return _registered

    for fam, regular, bold, italic, bolditalic in _FAMILIES:
        rp = os.path.join(FONT_DIR, regular)
        if not os.path.exists(rp):
            continue
        try:
            pdfmetrics.registerFont(TTFont(fam, rp))
        except Exception:
            continue
        # 变体：ReportLab 需要显式注册才能 <b>/<i> 生效
        for suffix, fn in (("-B", bold), ("-I", italic), ("-BI", bolditalic)):
            if not fn:
                continue
            fp = os.path.join(FONT_DIR, fn)
            if not os.path.exists(fp):
                continue
            try:
                pdfmetrics.registerFont(TTFont(fam + suffix, fp))
            except Exception:
                pass
        _registered[fam] = fam

    # 兜底：一个中文字体都没有，就用 reportlab 内置字体。
    # 注意不能把 "Helvetica" 当成族名塞进 _registered 的 key 位置，
    # 否则 <font name="雅黑"> 会让 reportlab 找不到字体而抛异常。
    if not _registered:
        _registered = {}

    _ready = True
    return _registered


def _fallback_font() -> str:
    """真正可用的兜底字体名。"""
    return "Helvetica"


def _pick_family(name: str | None) -> str:
    fams = _register_fonts()
    if not fams:
        return _fallback_font()
    n = (name or "").lower()
    for keys, fam in _NAME_MAP:
        if fam not in fams:
            continue
        for k in keys:
            if k in n:
                return fam
    # 默认：正文用宋体（贴近中文公文习惯），没有就用第一个可用的
    return "宋体" if "宋体" in fams else next(iter(fams))


def _resolve_font(name: str) -> str:
    """把任意字体名解析成 reportlab 里真实存在的字体，避免 KeyError。"""
    if not name:
        return _fallback_font()
    try:
        pdfmetrics.getFont(name)
        return name
    except Exception:
        return _fallback_font()


# ---------------------------------------------------------------- 单位换算

def _emu_to_pt(v) -> float | None:
    """python-docx 的 Length 是 EMU，1pt = 12700 EMU。"""
    if v is None:
        return None
    try:
        return float(v) / 12700.0
    except Exception:
        return None


def _twips_to_pt(v) -> float | None:
    if v is None:
        return None
    try:
        return float(v) / 20.0
    except Exception:
        return None


# ---------------------------------------------------------------- 样式解析

_ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: TA_LEFT,
    WD_ALIGN_PARAGRAPH.CENTER: TA_CENTER,
    WD_ALIGN_PARAGRAPH.RIGHT: TA_RIGHT,
    WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
    WD_ALIGN_PARAGRAPH.DISTRIBUTE: TA_JUSTIFY,
}


def _style_chain(style) -> list:
    """从当前样式一路沿 based_on 往上收集，用于继承查找。"""
    chain, seen = [], set()
    s = style
    while s is not None and id(s) not in seen:
        seen.add(id(s))
        chain.append(s)
        try:
            s = s.base_style
        except Exception:
            break
    return chain


def _para_effective(p: Paragraph) -> dict:
    """算出一个段落最终的格式（沿样式链继承）。"""
    pf = p.paragraph_format
    try:
        chain = _style_chain(p.style)
    except Exception:
        chain = []

    def first(attr, getter):
        # 先看段落自身，再沿样式链
        v = getter(pf)
        if v is not None:
            return v
        for st in chain:
            try:
                v = getter(st.paragraph_format)
            except Exception:
                v = None
            if v is not None:
                return v
        return None

    align = p.alignment
    if align is None:
        for st in chain:
            try:
                if st.paragraph_format.alignment is not None:
                    align = st.paragraph_format.alignment
                    break
            except Exception:
                pass

    size_pt = None
    for st in chain:
        try:
            if st.font.size:
                size_pt = _emu_to_pt(st.font.size)
                break
        except Exception:
            pass

    bold = None
    italic = None
    for st in chain:
        try:
            if bold is None and st.font.bold is not None:
                bold = bool(st.font.bold)
            if italic is None and st.font.italic is not None:
                italic = bool(st.font.italic)
        except Exception:
            pass

    color = None
    for st in chain:
        try:
            c = st.font.color
            if c and c.rgb is not None:
                color = c.rgb
                break
        except Exception:
            pass

    return {
        "align": _ALIGN_MAP.get(align, TA_LEFT),
        "size": size_pt or 10.5,                      # Word 中文默认五号=10.5pt
        "bold": bool(bold),
        "italic": bool(italic),
        "color": color,
        "first_line": _emu_to_pt(first("first_line_indent", lambda f: f.first_line_indent)),
        "left": _emu_to_pt(first("left_indent", lambda f: f.left_indent)) or 0.0,
        "right": _emu_to_pt(first("right_indent", lambda f: f.right_indent)) or 0.0,
        "space_before": _emu_to_pt(first("space_before", lambda f: f.space_before)),
        "space_after": _emu_to_pt(first("space_after", lambda f: f.space_after)),
        "line_spacing": first("line_spacing", lambda f: f.line_spacing),
        "style_name": _style_name(p),
    }


def _run_style(run, base: dict) -> dict:
    """单个 run 的最终样式：run 自身覆盖段落样式。"""
    size = _emu_to_pt(run.font.size) or base["size"]
    bold = run.bold if run.bold is not None else base["bold"]
    italic = run.italic if run.italic is not None else base["italic"]

    color = base["color"]
    try:
        c = run.font.color
        if c and c.rgb is not None:
            color = c.rgb
    except Exception:
        pass

    # 中文字体名常放在 rFonts 的 eastAsia 上，python-docx 的 font.name 读不到
    name = run.font.name
    if not name:
        try:
            rpr = run._element.find(qn("w:rPr"))
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None:
                    name = rf.get(qn("w:eastAsia")) or rf.get(qn("w:ascii"))
        except Exception:
            pass

    return {
        "size": size,
        "bold": bool(bold),
        "italic": bool(italic),
        "color": color,
        "family": _resolve_font(_pick_family(name)),
    }


# ---------------------------------------------------------------- 行内 HTML

def _esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


_FONT_RE = re.compile(r"(?<![-\w])(宋体|黑体|楷体|仿宋|等线|雅黑)(?![-\w])")


def _run_markup(text: str, st: dict) -> str:
    if not text:
        return ""
    body = _esc(text).replace("\n", "<br/>").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
    if st["color"] is not None:
        try:
            body = f'<font color="#{st["color"]}">{body}</font>'
        except Exception:
            pass
    fam = st["family"]
    if fam and fam != "Helvetica":
        # 字体族用 fontName 标签，避免与 <b>/<i> 冲突
        body = f'<font name="{fam}">{body}</font>'
    if st["bold"]:
        body = f"<b>{body}</b>"
    if st["italic"]:
        body = f"<i>{body}</i>"
    # 字号用 pt 绝对值，避免 reportlab 的 size 属性按相对档位缩放
    if abs(st["size"] - 10.5) > 0.01:
        body = f'<font size="{st["size"]:.1f}">{body}</font>'
    return body


def _para_markup(p: Paragraph, base: dict) -> str:
    parts = []
    for run in p.runs:
        if run.text:
            parts.append(_run_markup(run.text, _run_style(run, base)))
    # 链接、域控等非 run 文本兜底
    if not parts:
        return _esc(p.text or "")
    return "".join(parts)


# ---------------------------------------------------------------- 列表

_LIST_STYLES = ("list bullet", "list number", "list paragraph", "list")


def _style_name(p: Paragraph) -> str:
    """安全取段落样式名。我们自产的 docx 可能没有 styles.xml 里的对应定义，
    此时 p.style 会是 None，不能直接 .name。"""
    try:
        st = p.style
        if st is None:
            return ""
        return (st.name or "") or ""
    except Exception:
        return ""


def _is_list_item(p: Paragraph) -> tuple[bool, int, str]:
    """判断是否列表项，返回 (是否列表, 层级, 前缀)。"""
    name = _style_name(p).lower()
    if not any(k in name for k in _LIST_STYLES):
        return False, 0, ""

    # 从 numPr 读层级
    level = 0
    try:
        ppr = p._element.find(qn("w:pPr"))
        if ppr is not None:
            npr = ppr.find(qn("w:numPr"))
            if npr is not None:
                ilvl = npr.find(qn("w:ilvl"))
                if ilvl is not None:
                    level = int(ilvl.get(qn("w:val")) or 0)
    except Exception:
        pass

    bullet = "•" if "bullet" in name or "bullet" in name else "•"
    if "number" in name:
        bullet = "1."
    return True, level, bullet


# ---------------------------------------------------------------- 图片

def _iter_inline_shapes(paragraph: Paragraph, doc):
    """找出段落里的图片 run，返回 [(rId, width_pt, height_pt)]。"""
    out = []
    for run in paragraph.runs:
        try:
            for el in run._element.findall(qn("w:drawing")) + run._element.findall(qn("w:pict")):
                rid = None
                # 现代 drawing：a:blip r:embed
                for blip in el.iter(qn("a:blip")):
                    rid = blip.get(qn("r:embed"))
                    if rid:
                        break
                if not rid:
                    for imagedata in el.iter(qn("v:imagedata")):
                        rid = imagedata.get(qn("r:id"))
                        if rid:
                            break
                if not rid:
                    continue
                w = h = 0.0
                for ext in el.iter(qn("wp:extent")):
                    cx = ext.get("cx")
                    cy = ext.get("cy")
                    if cx and cy:
                        w = float(cx) / 12700.0
                        h = float(cy) / 12700.0
                        break
                out.append((rid, w, h))
        except Exception:
            continue
    return out


def _load_image(doc, rid: str) -> tuple[bytes, str] | None:
    try:
        part = doc.part.related_parts[rid]
        return part.blob, os.path.splitext(part.partname)[1].lower()
    except Exception:
        return None


# ---------------------------------------------------------------- 页面设置

def _page_setup(doc) -> dict:
    """读节属性：纸张、页边距、页眉页脚距离。"""
    setup = {
        "w": A4[0], "h": A4[1],
        "left": 72.0, "right": 72.0, "top": 72.0, "bottom": 72.0,
    }
    try:
        sec = doc.sections[0]
        def _pt(v, default):
            p = _emu_to_pt(v)
            return p if p and p > 0 else default
        setup["w"] = _pt(sec.page_width, A4[0])
        setup["h"] = _pt(sec.page_height, A4[1])
        setup["left"] = _pt(sec.left_margin, 72.0)
        setup["right"] = _pt(sec.right_margin, 72.0)
        setup["top"] = _pt(sec.top_margin, 72.0)
        setup["bottom"] = _pt(sec.bottom_margin, 72.0)
    except Exception:
        pass
    return setup


# ---------------------------------------------------------------- 表格

def _table_flowable(tbl: Table, base_size: float, avail_w: float):
    rows = []
    max_cols = 0
    for row in tbl.rows:
        cells = []
        for cell in row.cells:
            text = "\n".join(par.text for par in cell.paragraphs).strip()
            cells.append(_esc(text).replace("\n", "<br/>"))
        if cells:
            rows.append(cells)
            max_cols = max(max_cols, len(cells))
    if not rows or not max_cols:
        return None

    # 补齐列数，避免 ReportLab 报错
    for r in rows:
        while len(r) < max_cols:
            r.append("")

    style = ParagraphStyle(
        "tbl", fontName=_resolve_font(_pick_family(None)), fontSize=max(6, base_size - 1),
        leading=max(6, base_size - 1) * 1.45, wordWrap="CJK",
    )
    data = [[RLParagraph(c or "", style) for c in r] for r in rows]
    col_w = avail_w / max_cols
    t = RLTable(data, colWidths=[col_w] * max_cols, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#8a8f99")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f2f4f7")),
    ]))
    return t


# ---------------------------------------------------------------- 页眉页脚

def _draw_header_footer(canvas, doc, ctx: dict):
    canvas.saveState()
    canvas.setFont(_resolve_font(ctx.get("family")), 9)
    canvas.setFillColor(colors.HexColor("#666666"))

    setup = ctx["setup"]
    w = setup["w"]

    header = ctx.get("header") or ""
    if header:
        canvas.drawCentredString(w / 2.0, setup["h"] - setup["top"] * 0.55, header)

    footer = ctx.get("footer") or ""
    y = setup["bottom"] * 0.55
    if footer:
        canvas.drawCentredString(w / 2.0, y, footer)
    else:
        canvas.drawCentredString(w / 2.0, y, str(canvas.getPageNumber()))

    canvas.restoreState()


# ---------------------------------------------------------------- 主流程

def convert(src: str, dst: str, progress=None) -> dict:
    def _p(msg: str, frac: float) -> None:
        if progress:
            try:
                progress(msg, frac)
            except Exception:
                pass

    if not os.path.exists(src):
        raise RuntimeError("文件不存在：" + src)

    doc = Document(src)
    setup = _page_setup(doc)
    fams = _register_fonts()
    default_family = _pick_family(None)

    body_w = max(80.0, setup["w"] - setup["left"] - setup["right"])
    ctx = {
        "setup": setup,
        "family": default_family,
        "header": _header_text(doc),
        "footer": "",
    }

    path = dst
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)

    pdf = BaseDocTemplate(
        path, pagesize=(setup["w"], setup["h"]),
        leftMargin=setup["left"], rightMargin=setup["right"],
        topMargin=setup["top"], bottomMargin=setup["bottom"],
        title=os.path.splitext(os.path.basename(src))[0],
        author="",
    )
    frame = Frame(setup["left"], setup["bottom"], body_w,
                  max(40.0, setup["h"] - setup["top"] - setup["bottom"]),
                  id="body", showBoundary=0)
    pdf.addPageTemplates([PageTemplate(
        id="main", frames=[frame],
        onPage=lambda c, d: _draw_header_footer(c, d, ctx),
    )])

    story: list[Any] = []
    stats = {"paras": 0, "tables": 0, "images": 0, "chars": 0}

    _p("正在解析文档结构", 0.15)

    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            para = Paragraph(child, doc)
            _emit_paragraph(para, doc, story, setup, body_w, fams, default_family, stats)
        elif tag == "tbl":
            tbl = Table(child, doc)
            fl = _table_flowable(tbl, 10.5, body_w)
            if fl is not None:
                story.append(Spacer(1, 4))
                story.append(fl)
                story.append(Spacer(1, 6))
                stats["tables"] += 1

    if not story:
        story.append(Spacer(1, 1))

    _p("正在排版生成 PDF", 0.7)
    pdf.build(story)

    if not os.path.exists(path) or os.path.getsize(path) == 0:
        raise RuntimeError("转换失败：未能生成有效的 PDF")

    # 页数：BaseDocTemplate 构建完可以通过 page 属性拿到
    pages = 0
    try:
        pages = int(getattr(pdf, "page", 0) or 0)
    except Exception:
        pass
    if not pages:
        try:
            import pymupdf
            d = pymupdf.open(path)
            pages = d.page_count
            d.close()
        except Exception:
            pages = 0

    return {"ok": True, "path": path, "pages": pages, **stats}


def _header_text(doc) -> str:
    try:
        for sec in doc.sections:
            h = sec.header
            if h is not None and not h.is_linked_to_previous:
                t = " ".join(p.text for p in h.paragraphs).strip()
                if t:
                    return t
    except Exception:
        pass
    return ""


def _emit_paragraph(para, doc, story, setup, body_w, fams, default_family, stats):
    base = _para_effective(para)
    is_list, level, bullet = _is_list_item(para)

    text = para.text if para.text is not None else ""

    # 图片：按 run 在段落中的实际顺序处理。纯图片段落直接出图；
    # 图文混排时图片单独成行，文字保持原有的前后关系。
    shapes = _iter_inline_shapes(para, doc)
    has_text = bool((para.text or "").strip())

    if shapes and not has_text:
        for rid, w, h in shapes:
            _emit_image(rid, w, h, doc, story, setup, body_w, stats)
        return

    # 图文混排：先出文字段，再出图（Word 里这类图通常也独占一行）
    markup = _para_markup(para, base)

    if not (markup or "").strip():
        # 空段落保留为行高，维持段间距
        if text == "" and para.runs:
            story.append(Spacer(1, max(2, base["size"] * 0.6)))
        return

    stats["chars"] += len(text)

    # 字号护栏：异常值（如误把 EMU 当 pt）会把版面撑爆，直接夹紧到合理区间
    size = base["size"]
    if not size or size > 200 or size < 3:
        size = 10.5
    size = max(5.0, float(size))
    leading = size * 1.55
    ls = base["line_spacing"]
    # python-docx 的 line_spacing 有两种形态：
    #   - float/int：倍数（1.5、2）
    #   - Length（EMU）：绝对行距，必须先换算成 pt
    # 之前只判 float/int，Length 会落到「倍数」分支，leading 被算成几十万 pt 直接崩掉。
    if ls is not None:
        if isinstance(ls, (int, float)):
            mult = float(ls)
            if 0.5 <= mult <= 5:
                leading = size * 1.2 * mult
        else:
            pt_val = _emu_to_pt(ls)
            if pt_val and 1 <= pt_val <= 400:
                leading = pt_val

    indent = base["first_line"]
    if is_list:
        indent = 0
        markup = f'{"&nbsp;" * (level * 4)}{_esc(bullet)}&nbsp;{markup}'

    style = ParagraphStyle(
        f"p{stats['paras']}",
        fontName=_resolve_font(default_family),
        fontSize=size,
        leading=leading,
        alignment=base["align"],
        leftIndent=max(0.0, base["left"]) + (level * 12 if is_list else 0),
        rightIndent=max(0.0, base["right"]),
        firstLineIndent=indent or 0,
        spaceBefore=max(0.0, base["space_before"] or 0),
        spaceAfter=max(0.0, base["space_after"] if base["space_after"] is not None else 4),
        wordWrap="CJK",
        allowWidows=1,
        allowOrphans=1,
    )

    try:
        story.append(RLParagraph(markup, style))
        stats["paras"] += 1
    except Exception:
        # 极端情况：标记有问题时退回纯文本，保证内容不丢
        try:
            story.append(RLParagraph(_esc(text), style))
            stats["paras"] += 1
        except Exception:
            pass

    # 混排段落里的图片排在文字之后，避免图把文字顶到下方
    for rid, w, h in shapes:
        _emit_image(rid, w, h, doc, story, setup, body_w, stats)


def _emit_image(rid, w, h, doc, story, setup, body_w, stats) -> None:
    """把一张内嵌图片按版心宽度等比放进文档流。"""
    loaded = _load_image(doc, rid)
    if not loaded:
        return
    data, _ext = loaded

    disp_w = min(w if w > 0 else body_w, body_w)
    disp_h = (h * disp_w / w) if (w > 0 and h > 0) else disp_w * 0.75

    # 太高的图限制在版心内，避免一页塞不下
    max_h = max(40.0, setup["h"] - setup["top"] - setup["bottom"] - 24)
    if disp_h > max_h:
        disp_w = disp_w * (max_h / disp_h)
        disp_h = max_h

    try:
        story.append(Image(io.BytesIO(data), width=disp_w, height=disp_h))
        story.append(Spacer(1, 4))
        stats["images"] += 1
    except Exception:
        pass


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("用法: python word_to_pdf.py <input.docx> <output.pdf>")
        sys.exit(1)
    print(convert(sys.argv[1], sys.argv[2]))
