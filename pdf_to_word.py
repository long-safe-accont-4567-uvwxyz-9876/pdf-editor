"""PDF → Word（.docx）—— 不依赖任何第三方库。

思路：
  1. 逐页读取文字 span（文字、坐标、字号、字体、颜色、粗斜体）
  2. 按「同一水平基线 + 间距」把 span 归并成段落，同段内保留行内样式
  3. 收集页内图片，按原始分辨率裁出
  4. 直接拼 OOXML 写进 zip，生成真正的 .docx（Word/WPS 都能打开）

为什么不用 pdf2docx：它是重依赖（numpy / opencv / python-docx），
离线安装不上时整个功能就废了。这里只用标准库 + PyMuPDF，任何环境都能跑。
"""

from __future__ import annotations

import io
import os
import re
import zipfile

import pymupdf

# ---------------------------------------------------------------- 页面尺寸

# Word 用「缇」（twip）= 1/20 pt 作为长度单位
def _tw(v: float) -> int:
    return int(round(v * 20))


A4_W, A4_H = 595.28, 841.89          # PDF 点
MARGIN_PT = 72.0                      # 默认页边距 1 英寸

_W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
_WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
_A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
_PIC_NS = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'


def _xml_escape(s: str) -> str:
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    # 去掉 Word 不认的控制字符
    return "".join(ch for ch in s if ch >= " " or ch in "\t")


# ---------------------------------------------------------------- 字体映射

# PDF 内置/嵌入字体名 → (西文字体, 中文字体)
_FONT_MAP = [
    (("hei", "heiti", "黑体", "stheiti", "sourcehansans", "yahei", "msyh"), "SimHei", "黑体"),
    (("song", "simsun", "宋体", "nsimsun", "stsong", "serif-cjk"), "SimSun", "宋体"),
    (("kai", "kaiti", "楷体"), "KaiTi", "楷体"),
    (("fang", "fangsong", "仿宋"), "FangSong", "仿宋"),
    (("times", "timesnewroman"), "Times New Roman", "宋体"),
    (("arial", "helvetica", "helv"), "Arial", "宋体"),
    (("calibri",), "Calibri", "宋体"),
    (("cour", "courier", "mono"), "Courier New", "宋体"),
    (("georgia",), "Georgia", "宋体"),
    (("verdana",), "Verdana", "宋体"),
    (("cambria",), "Cambria", "宋体"),
]


def _map_font(pdf_font: str) -> tuple[str, str]:
    f = (pdf_font or "").lower()
    for keys, ascii_font, east_font in _FONT_MAP:
        for k in keys:
            if k in f:
                return ascii_font, east_font
    # 未知字体：含中文的文档默认宋体
    return "Times New Roman", "宋体"


def _color_hex(c: int) -> str:
    return f"{c & 0xFFFFFF:06X}"


# ---------------------------------------------------------------- 段落归并

def _spans_of_page(page: pymupdf.Page) -> list[dict]:
    """取页面所有可见文字 span，带上样式与几何信息。"""
    out: list[dict] = []
    try:
        dic = page.get_text("dict")
    except Exception:
        return out

    for blk in dic.get("blocks", []):
        if blk.get("type") != 0:
            continue
        for line in blk.get("lines", []):
            for sp in line.get("spans", []):
                text = sp.get("text") or ""
                if not text.strip():
                    continue
                bbox = sp.get("bbox") or (0, 0, 0, 0)
                flags = int(sp.get("flags") or 0)
                out.append({
                    "text": text,
                    "bbox": tuple(float(v) for v in bbox),
                    "size": float(sp.get("size") or 11),
                    "font": sp.get("font") or "",
                    "color": int(sp.get("color") or 0),
                    "bold": bool(flags & 16) or "bold" in (sp.get("font") or "").lower(),
                    "italic": bool(flags & 2),
                    "origin_y": float((sp.get("origin") or (0, bbox[3]))[1]),
                    "dir": tuple(line.get("dir") or (1, 0)),
                })
    out.sort(key=lambda s: (round(s["bbox"][1], 1), s["bbox"][0]))
    return out


def _page_flow(page: pymupdf.Page) -> list[dict]:
    """按自然阅读顺序返回页面的内容块（文字块 / 图片块）。

    PyMuPDF 的 blocks 已经是按阅读顺序排的，直接用它比拿坐标自己
    交叉排序可靠得多——后者容易把图和文字拆散、插错位置。
    """
    flow: list[dict] = []
    try:
        dic = page.get_text("dict")
    except Exception:
        dic = {"blocks": []}

    for blk in dic.get("blocks", []):
        bbox = tuple(float(v) for v in (blk.get("bbox") or (0, 0, 0, 0)))
        if blk.get("type") == 1:
            # 图片块
            if bbox[2] - bbox[0] >= 4 and bbox[3] - bbox[1] >= 4:
                flow.append({"kind": "image", "bbox": bbox})
            continue

        spans: list[dict] = []
        for line in blk.get("lines", []):
            for sp in line.get("spans", []):
                text = sp.get("text") or ""
                if not text.strip():
                    continue
                sb = tuple(float(v) for v in (sp.get("bbox") or (0, 0, 0, 0)))
                flags = int(sp.get("flags") or 0)
                spans.append({
                    "text": text,
                    "bbox": sb,
                    "size": float(sp.get("size") or 11),
                    "font": sp.get("font") or "",
                    "color": int(sp.get("color") or 0),
                    "bold": bool(flags & 16) or "bold" in (sp.get("font") or "").lower(),
                    "italic": bool(flags & 2),
                    "origin_y": float((sp.get("origin") or (0, sb[3]))[1]),
                    "dir": tuple(line.get("dir") or (1, 0)),
                })
        if spans:
            flow.append({"kind": "text", "spans": spans, "bbox": bbox})

    return flow


def _group_lines(spans: list[dict]) -> list[list[dict]]:
    """把 span 按基线聚成「行」。"""
    lines: list[list[dict]] = []
    for sp in spans:
        placed = False
        for ln in lines:
            ref = ln[0]
            if abs(sp["origin_y"] - ref["origin_y"]) <= max(2.0, ref["size"] * 0.35):
                ln.append(sp)
                placed = True
                break
        if not placed:
            lines.append([sp])
    for ln in lines:
        ln.sort(key=lambda s: s["bbox"][0])
    lines.sort(key=lambda ln: min(s["bbox"][1] for s in ln))
    return lines


def _group_paragraphs(lines: list[list[dict]]) -> list[list[list[dict]]]:
    """把行归并成段落。

    判定为同一段要求同时满足：
      - 行距在合理范围内（不与字号脱节）
      - 左右边界基本对齐
      - 字号不突变（标题→正文必须断开）
    """
    paras: list[list[list[dict]]] = []
    for ln in lines:
        if not paras:
            paras.append([ln])
            continue
        prev = paras[-1][-1]
        prev_x0 = min(s["bbox"][0] for s in prev)
        cur_x0 = min(s["bbox"][0] for s in ln)
        prev_y1 = max(s["bbox"][3] for s in prev)
        cur_y0 = min(s["bbox"][1] for s in ln)
        prev_size = max(s["size"] for s in prev) or 11
        cur_size = max(s["size"] for s in ln) or 11

        gap = cur_y0 - prev_y1
        # 行距按两行中较大的字号衡量，避免字号差异造成误判
        ref = max(prev_size, cur_size)

        same = (
            # 字号差超过 15% 视为不同层级（标题 / 正文）
            abs(prev_size - cur_size) / max(prev_size, cur_size) <= 0.15
            # 垂直间距：正常的行内换行（含 1.5 倍行距、少量额外行间距）
            and -ref * 0.6 <= gap <= ref * 1.2
            # 左边界：允许首行缩进，但不允许换段落式的错位
            and abs(cur_x0 - prev_x0) <= ref * 1.8
        )
        paras[-1].append(ln) if same else paras.append([ln])
    return paras


def _line_text(ln: list[dict]) -> str:
    """同一行内按位置拼接，跨度过大时补空格。"""
    parts: list[str] = []
    prev = None
    for sp in ln:
        if prev is not None:
            gap = sp["bbox"][0] - prev["bbox"][2]
            if gap > max(2.0, (prev["size"] or 11) * 0.35) and parts and not parts[-1].endswith(" "):
                parts.append(" ")
        parts.append(sp["text"])
        prev = sp
    return "".join(parts)


# ---------------------------------------------------------------- OOXML 片段

def _run_xml(text: str, sp: dict, first: bool) -> str:
    ascii_font, east_font = _map_font(sp["font"])
    sz = max(1, int(round(sp["size"] * 2)))          # 半磅
    rpr = [f'<w:rFonts w:ascii="{ascii_font}" w:hAnsi="{ascii_font}" w:eastAsia="{east_font}"/>']
    if sp["bold"]:
        rpr.append("<w:b/>")
    if sp["italic"]:
        rpr.append("<w:i/>")
    rpr.append(f'<w:color w:val="{_color_hex(sp["color"])}"/>')
    rpr.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    body = _xml_escape(text).replace("\n", "</w:t><w:br/><w:t>")
    return f'<w:r><w:rPr>{"".join(rpr)}</w:rPr><w:t xml:space="preserve">{body}</w:t></w:r>'


def _para_xml(lines: list[list[dict]], page_w: float, left: float, right: float) -> str:
    # 段落样式取首行首块
    first_sp = lines[0][0]
    size = first_sp["size"] or 11
    x0 = min(s["bbox"][0] for s in lines[0])

    ind_left = max(0, x0 - left)

    # 行距按原 PDF 的实测行间距算，而不是一律写死 1.5 倍。
    # 写死会让「PDF→Word→PDF」来回转时行距逐轮变大，最后多出空白页。
    # 用 exact 而不是 atLeast：atLeast 会在实测值之上再叠加字体行高，
    # 导致每转一轮行距就变一次，页数也跟着漂。
    leading = max(size, 1.0) * 1.2
    if len(lines) >= 2:
        gaps = []
        for i in range(1, len(lines)):
            y_prev = min(s["bbox"][1] for s in lines[i - 1])
            y_cur = min(s["bbox"][1] for s in lines[i])
            g = y_cur - y_prev
            if 0.5 * size <= g <= 4 * size:
                gaps.append(g)
        if gaps:
            gaps.sort()
            # 取偏小的那个（段落内的正常行距），避开段末/段首的大间距
            leading = gaps[0] if len(gaps) <= 2 else gaps[len(gaps) // 3]
    # 行距不能小于字号本身，否则文字会挤在一起
    leading = max(leading, size * 1.05)

    ppr = [
        f'<w:spacing w:line="{_tw(leading)}" w:lineRule="exact"'
        f' w:before="0" w:after="0"/>',
        f'<w:ind w:left="{_tw(ind_left)}"/>',
    ]

    # 段落内不写硬换行：让 Word 自己按页宽折行，用户改字/改字号时才不会错位。
    # 行与行直接拼接，仅在西文之间补空格（中文行末不需要）。
    runs: list[str] = []
    prev_last = ""
    for li, ln in enumerate(lines):
        for si, sp in enumerate(ln):
            text = sp["text"]
            if li > 0 and si == 0 and prev_last:
                # 上一行结尾是西文字符、本行开头也是西文 → 补空格
                if _needs_space(prev_last, text[:1]):
                    text = " " + text
            runs.append(_run_xml(text, sp, li == 0 and si == 0))
            prev_last = sp["text"][-1:] if sp["text"] else prev_last

    return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{"".join(runs)}</w:p>'


def _needs_space(prev_ch: str, next_ch: str) -> bool:
    """西文单词之间需要补空格，中日韩字符之间不需要。"""
    if not prev_ch or not next_ch:
        return False
    if prev_ch.isspace() or next_ch.isspace():
        return False

    def cjk(ch: str) -> bool:
        o = ord(ch)
        return 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF

    return not cjk(prev_ch) and not cjk(next_ch)


def _drawing_xml(rid: str, rid_px: int, cx_emu: int, cy_emu: int, doc_pr_id: int) -> str:
    return (
        '<w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx_emu}" cy="{cy_emu}"/>'
        f'<wp:docPr id="{doc_pr_id}" name="Picture {doc_pr_id}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic>'
        f'<pic:nvPicPr><pic:cNvPr id="{doc_pr_id}" name="img{rid_px}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm>'
        f'<a:off x="0" y="0"/><a:ext cx="{cx_emu}" cy="{cy_emu}"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    )


# ---------------------------------------------------------------- 主流程

def convert(src: str, dst: str, progress=None) -> dict:
    """把 PDF 转成 .docx。返回统计信息用于前端提示。"""
    def _p(msg: str, frac: float) -> None:
        if progress:
            try:
                progress(msg, frac)
            except Exception:
                pass

    doc = pymupdf.open(src)
    try:
        total = doc.page_count
        if total == 0:
            raise RuntimeError("PDF 没有页面")

        body_parts: list[str] = []
        media: list[tuple[str, bytes]] = []
        rels_parts: list[str] = []
        img_seq = [0]
        pic_ids = [1000]

        # 有文字页 / 纯图页 分开统计
        text_chars = 0
        image_pages = 0

        for pno in range(total):
            page = doc[pno]
            _p(f"正在解析第 {pno + 1}/{total} 页", 0.1 + 0.7 * pno / total)

            if pno > 0:
                body_parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

            pw, ph = page.rect.width, page.rect.height
            left, right = MARGIN_PT, max(MARGIN_PT, pw - MARGIN_PT)

            flow = _page_flow(page)
            text_spans = [s for blk in flow if blk["kind"] == "text" for s in blk["spans"]]
            text_chars += sum(len(s["text"]) for s in text_spans)

            has_img = any(blk["kind"] == "image" for blk in flow)

            # 既没有文字也没有图片：整页截图兜底，保证内容不丢
            if not text_spans and not has_img:
                image_pages += 1
                body_parts.append(_page_snapshot(
                    page, media, rels_parts, img_seq, pic_ids, pw,
                ))
                continue

            # 按内容块自然顺序输出：文字块成段、图片块单独成段
            for blk in flow:
                if blk["kind"] == "image":
                    body_parts.append(_image_paragraph(
                        page, blk["bbox"], media, rels_parts, img_seq, pic_ids, left, right,
                    ))
                    continue

                spans = blk["spans"]
                spans.sort(key=lambda s: (round(s["bbox"][1], 1), s["bbox"][0]))
                lines = _group_lines(spans)
                for para in _group_paragraphs(lines):
                    part = _para_xml(para, pw, left, right)
                    if part:
                        body_parts.append(part)

        _p("正在生成 Word 文档", 0.85)
        _write_docx(dst, body_parts, media, rels_parts)

        if not os.path.exists(dst) or os.path.getsize(dst) == 0:
            raise RuntimeError("转换失败：未能生成有效的 Word 文档")

        pages_with_text = sum(1 for part in body_parts if "<w:t" in part)
        return {
            "ok": True,
            "path": dst,
            "pages": total,
            "chars": text_chars,
            "images": img_seq[0],
            # 几乎没有文字、全靠截图兜底 → 提醒用户这是扫描版
            "scanned": text_chars < 20 * total,
        }
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _add_image_media(data: bytes, media, rels_parts, img_seq) -> tuple[str, int]:
    idx = img_seq[0]
    img_seq[0] += 1
    name = f"image{idx + 1}.png"
    media.append((name, data))
    rid = f"rIdImg{idx + 1}"
    rels_parts.append(
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/image" Target="media/{name}"/>'
    )
    return rid, idx


def _image_paragraph(page, bbox, media, rels_parts, img_seq, pic_ids, left, right) -> str:
    """把一块图片提取成 docx 图片段落（居中、等比缩放到页宽内）。"""
    x0, y0, x1, y1 = bbox
    w_pt, h_pt = x1 - x0, y1 - y0
    if w_pt < 2 or h_pt < 2:
        return ""
    try:
        clip = pymupdf.Rect(x0, y0, x1, y1)
        # 至少 2 倍渲染，避免小图放大后模糊
        zoom = max(2.0, min(6.0, (w_pt * 2.0) / max(1.0, w_pt)))
        pix = page.get_pixmap(clip=clip, matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        data = pix.tobytes("png")
    except Exception:
        return ""

    rid, idx = _add_image_media(data, media, rels_parts, img_seq)

    avail = max(40.0, right - left)
    disp_w = min(w_pt, avail)
    disp_h = h_pt * (disp_w / w_pt) if w_pt else h_pt
    pic_ids[0] += 1
    # EMU = pt * 12700
    return (
        '<w:p><w:pPr><w:jc w:val="center"/>'
        '<w:spacing w:before="120" w:after="120"/></w:pPr>'
        f'{_drawing_xml(rid, idx, int(disp_w * 12700), int(disp_h * 12700), pic_ids[0])}</w:p>'
    )


def _page_snapshot(page, media, rels_parts, img_seq, pic_ids, pw) -> str:
    """整页截图，用于纯图片/纯矢量页，保证内容不丢失。"""
    try:
        pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
        data = pix.tobytes("png")
    except Exception:
        return ""

    rid, idx = _add_image_media(data, media, rels_parts, img_seq)
    avail = max(40.0, pw - 2 * MARGIN_PT)
    disp_w = avail
    disp_h = avail * (pix.height / pix.width) if pix.width else avail
    pic_ids[0] += 1
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'{_drawing_xml(rid, idx, int(disp_w * 12700), int(disp_h * 12700), pic_ids[0])}</w:p>'
    )


def _write_docx(dst: str, body_parts: list[str], media, rels_parts) -> None:
    sect = (
        '<w:sectPr>'
        f'<w:pgSz w:w="{_tw(A4_W)}" w:h="{_tw(A4_H)}"/>'
        f'<w:pgMar w:top="{_tw(MARGIN_PT)}" w:right="{_tw(MARGIN_PT)}" '
        f'w:bottom="{_tw(MARGIN_PT)}" w:left="{_tw(MARGIN_PT)}" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr>'
    )

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {_W_NS} {_R_NS} {_WP_NS} {_A_NS} {_PIC_NS}><w:body>'
        + "".join(body_parts) + sect +
        '</w:body></w:document>'
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )

    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels_parts) +
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles {_W_NS}>'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/>'
        '</w:rPr></w:rPrDefault></w:docDefaults>'
        '</w:styles>'
    )

    d = os.path.dirname(os.path.abspath(dst))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles)
        for name, data in media:
            z.writestr(f"word/media/{name}", data)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("用法: python pdf_to_word.py <input.pdf> <output.docx>")
        sys.exit(1)
    print(convert(sys.argv[1], sys.argv[2]))
