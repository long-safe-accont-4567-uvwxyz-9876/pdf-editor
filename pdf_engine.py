"""PDF 编辑内核 —— 基于 PyMuPDF 的「擦除 + 原位重写」引擎。

原理与 WPS / Acrobat 的文字编辑一致：
  1. 读取页面文字 span（文本、坐标、字号、颜色、字体、粗斜体）
  2. 用 redaction 精确擦除目标区域，但保留矢量图形与图片（背景色块/底纹/照片不受伤）
  3. 在原基线位置用匹配的字体重新写入新文字

编辑以「操作集（ops）」记录，任何时刻都从原始文件重放，因此撤销/重做天然可靠。
"""

from __future__ import annotations

import os
import threading
import time

import pymupdf

# ---------------------------------------------------------------- 字体解析

FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")

from paths import WORK_DIR  # noqa: E402  （打包后必须落在可写目录）

# (关键词, 常规字体文件, 粗体字体文件, 展示名)
_FONT_RULES = [
    (("yahei", "msyh", "microsoftyahei", "微软雅黑"), "msyh.ttc", "msyhbd.ttc", "微软雅黑"),
    (("simhei", "heiti", "黑体", "stheiti", "sourcehansans"), "simhei.ttf", "simhei.ttf", "黑体"),
    (("simsun", "song", "宋体", "nsimsun", "stsong", "fzss", "fzshusong"), "simsun.ttc", "simsun.ttc", "宋体"),
    (("simkai", "kai", "楷体", "kaiti"), "simkai.ttf", "simkai.ttf", "楷体"),
    (("simfang", "fang", "仿宋", "fangsong"), "simfang.ttf", "simfang.ttf", "仿宋"),
    (("dengxian", "deng", "等线"), "Deng.ttf", "Dengb.ttf", "等线"),
    (("arial", "helvetica", "helv"), "arial.ttf", "arialbd.ttf", "Arial"),
    (("times", "timesnewroman", "timesroman"), "times.ttf", "timesbd.ttf", "Times New Roman"),
    (("calibri",), "calibri.ttf", "calibrib.ttf", "Calibri"),
    (("courier", "mono"), "cour.ttf", "courbd.ttf", "Courier New"),
    (("georgia",), "georgia.ttf", "georgiab.ttf", "Georgia"),
    (("verdana",), "verdana.ttf", "verdanab.ttf", "Verdana"),
    (("cambria",), "cambria.ttc", "cambriab.ttf", "Cambria"),
]


def _norm(name: str) -> str:
    return "".join(str(name or "").lower().split()).replace("-", "").replace("_", "")


def _has_cjk(text: str) -> bool:
    for ch in text or "":
        o = ord(ch)
        if 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF:
            return True
    return False


def resolve_font(font_name: str, text: str, bold: bool = False) -> tuple[str, str | None]:
    """PDF 字体名 + 文本内容 → (PyMuPDF 字体别名, 字体文件绝对路径或 None)。

    返回 None 表示使用 PyMuPDF 内置 Base14 字体（helv / tiro 等）。
    """
    n = _norm(font_name)

    for keys, regular, bolder, _label in _FONT_RULES:
        if any(k in n for k in keys):
            fname = bolder if bold else regular
            path = os.path.join(FONT_DIR, fname)
            if os.path.exists(path):
                return os.path.splitext(fname)[0], path
            break

    # 未命中关键词：按内容判定语言
    if _has_cjk(text):
        cand = "msyhbd.ttc" if bold else "msyh.ttc"
        path = os.path.join(FONT_DIR, cand)
        if os.path.exists(path):
            return "msyh", path
        path = os.path.join(FONT_DIR, "simsun.ttc")
        if os.path.exists(path):
            return "simsun", path
        return "china-s", None  # PyMuPDF 内置 CJK 兜底

    # 西文兜底：衬线特征走 times，其余 helv
    return ("tibo" if bold else "tiro", None) if "serif" in n else ("hebo" if bold else "helv", None)


def display_font(font_name: str, text: str = "") -> str:
    n = _norm(font_name)
    for keys, _, _, label in _FONT_RULES:
        if any(k in n for k in keys):
            return label
    return "微软雅黑" if _has_cjk(text) else "Arial"


# ---------------------------------------------------------------- 数据结构


def _color_int_to_rgb(c: int) -> list[int]:
    return [(c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF]


def _infer_align(span: dict, line: dict, block: dict, page_width: float = 595.0) -> str:
    """推断文字块的对齐方式。

    PDF 里没有"对齐"这个属性，只能靠位置反推。先用行/块内相对位置判断，
    单 span 撑满或压根无参照时，退到页面级留白来判断（标题、标语多居中）。
    """
    sx0, _, sx1, _ = span["bbox"]
    span_w = sx1 - sx0
    if span_w <= 0 or page_width <= 0:
        return "left"

    lx0, _, lx1, _ = line["bbox"]
    bx0, _, bx1, _ = block["bbox"]
    ref0, ref1 = (bx0, bx1) if (bx1 - bx0) > (lx1 - lx0) + 2 else (lx0, lx1)
    ref_w = ref1 - ref0
    tol = max(2.0, span["size"] * 0.15)

    # 一、行/块内有可比对象时，按相对留白判断
    if ref_w > 0 and abs(span_w - ref_w) > tol:
        lg, rg = sx0 - ref0, ref1 - sx1
        if lg > tol and rg > tol and abs(lg - rg) <= max(tol, ref_w * 0.06):
            return "center"
        if lg > tol * 2 and rg <= tol:
            return "right"
        if lg <= tol:
            return "left"

    # 二、退到页面级留白（覆盖标题、居中标语、右对齐页眉）
    pl, pr = sx0, page_width - sx1
    tol_p = max(3.0, page_width * 0.015)
    if pl > page_width * 0.06 and pr > page_width * 0.06 and abs(pl - pr) <= max(tol_p * 2, page_width * 0.035):
        return "center"
    if pr < pl * 0.45 and pl > page_width * 0.12:
        return "right"
    return "left"


_FONT_OBJ_CACHE: dict = {}


def get_font_obj(fontfile: str | None, alias: str | None):
    """缓存 pymupdf.Font 对象。

    字体文件（尤其 10MB 级的中文字体）解析一次约 30ms，逐个文字块重复解析
    会让"改几十处"的文档明显卡顿，因此全局复用。
    """
    key = fontfile or ("alias:" + str(alias))
    obj = _FONT_OBJ_CACHE.get(key)
    if obj is None:
        try:
            obj = pymupdf.Font(fontfile=fontfile) if fontfile else pymupdf.Font(fontname=alias)
        except Exception:
            obj = None
        _FONT_OBJ_CACHE[key] = obj
    return obj


def _wrap_text(font_obj, text: str, size: float, max_w: float) -> list[str]:
    """按实际字形宽度折行：中文逐字断，西文优先在空格处断。"""
    if not text:
        return [""]
    try:
        if font_obj.text_length(text, fontsize=size) <= max_w:
            return [text]
    except Exception:
        return [text]

    out: list[str] = []
    cur = ""
    for ch in text:
        if cur and font_obj.text_length(cur + ch, fontsize=size) > max_w:
            # 西文尽量不切断单词
            if ch.isascii() and ch.isalnum() and " " in cur:
                cut = cur.rfind(" ")
                if cut > len(cur) * 0.5:
                    out.append(cur[:cut])
                    cur = cur[cut + 1:] + ch
                    continue
            out.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        out.append(cur)
    return out


def _font_family(name: str) -> str:
    """归一化字体标识。

    Type3 文档常给每个字形分配独立字体对象（Type3 (8 0 R)、(9 0 R)…），
    直接比字符串会导致同一行文字永远无法合并，因此按家族归并。
    """
    n = str(name or "")
    if n.startswith("Type3"):
        return "Type3"
    return n


def _merge_line_spans(spans: list[dict]) -> list[dict]:
    """把同一行内样式一致、紧邻的 span 合并成一个可编辑块。

    某些 PDF（尤其是打印驱动、OCR、Type3 字体）会把每个字拆成独立 span，
    直接暴露给用户会变成"一次只能改一个字"。合并后才是正常的编辑粒度。
    """
    if len(spans) <= 1:
        return list(spans)

    out: list[dict] = []
    for sp in spans:
        if not out:
            out.append(dict(sp))
            continue

        prev = out[-1]
        psize = float(prev.get("size", 0)) or 1.0
        same_style = (
            _font_family(prev.get("font")) == _font_family(sp.get("font"))
            and prev.get("color") == sp.get("color")
            and prev.get("flags") == sp.get("flags")
            and abs(psize - float(sp.get("size", 0))) < 0.15
        )

        gap = sp["bbox"][0] - prev["bbox"][2]
        same_line = abs(sp["bbox"][1] - prev["bbox"][1]) < psize * 0.5
        # 同一列内几乎没有间隙；跨列/跨单元格会有明显留白
        close_enough = -psize * 0.8 < gap < psize * 0.9

        if same_style and same_line and close_enough:
            prev["text"] = prev.get("text", "") + sp.get("text", "")
            prev["bbox"] = [
                min(prev["bbox"][0], sp["bbox"][0]),
                min(prev["bbox"][1], sp["bbox"][1]),
                max(prev["bbox"][2], sp["bbox"][2]),
                max(prev["bbox"][3], sp["bbox"][3]),
            ]
        else:
            out.append(dict(sp))
    return out


class DocSession:
    """一份打开的文档及其全部编辑操作。"""

    def __init__(self, path: str, doc_id: str, orig_dir: str | None = None):
        self.id = doc_id
        self.src = path
        self.name = os.path.basename(path)
        # 用户是从哪个目录打开这份 PDF 的（拖拽/上传时为空）
        self.orig_dir = orig_dir or ""
        self.ops: list[dict] = []
        self.redo_stack: list[dict] = []
        self.version = 0
        self.lock = threading.RLock()
        self._spans: dict[int, list[dict]] = {}
        self._render_cache: dict[tuple, bytes] = {}
        self._built = None          # 重放结果缓存，供渲染复用
        self._built_ver = -1
        self._size_cache: tuple[int, float, float] | None = None
        self._build_baseline()

    # -- 基线 ---------------------------------------------------------
    def _open_src(self) -> pymupdf.Document:
        d = pymupdf.open(self.src)
        if d.needs_pass:
            d.authenticate("")
        return d

    def _build_baseline(self) -> None:
        doc = self._open_src()
        self._doc = doc
        self._size_cache = (doc.page_count, 0, 0)

    @property
    def page_count(self) -> int:
        return self._doc.page_count

    def page_size(self, index: int) -> list[float]:
        r = self._doc[index].rect
        w, h = r.width, r.height
        # 会话内已经旋转过的页，宽高需要对调
        for o in self.ops:
            if o.get("type") == "page" and o.get("action") == "rotate" and int(o.get("page", -1)) == index:
                if int(o.get("angle", 90)) % 180 == 90:
                    w, h = h, w
        return [round(w, 2), round(h, 2)]

    def all_page_sizes(self) -> list[list[float]]:
        return [self.page_size(i) for i in range(self.page_count)]

    # -- 文字提取 -----------------------------------------------------
    def spans(self, page_index: int) -> list[dict]:
        if page_index in self._spans:
            return self._spans[page_index]

        page = self._doc[page_index]
        out: list[dict] = []
        idx = 0
        raw = page.get_text("dict", sort=False)
        for block in raw.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in _merge_line_spans(line.get("spans", [])):
                    text = span.get("text", "")
                    if not text.strip():
                        continue
                    flags = span.get("flags", 0)
                    bold = bool(flags & 16) or "bold" in _norm(span.get("font", ""))
                    italic = bool(flags & 2)
                    size = round(float(span.get("size", 12)), 2)
                    origin = span.get("origin") or (span["bbox"][0], span["bbox"][3])
                    out.append({
                        "id": f"p{page_index}_s{idx}",
                        "page": page_index,
                        "text": text,
                        "bbox": [round(v, 2) for v in span["bbox"]],
                        "origin": [round(origin[0], 2), round(origin[1], 2)],
                        "size": size,
                        "color": _color_int_to_rgb(span.get("color", 0)),
                        "font": span.get("font", ""),
                        "fontLabel": display_font(span.get("font", ""), text),
                        "bold": bold,
                        "italic": italic,
                        "align": _infer_align(span, line, block, page.rect.width),
                        "lineHeight": round(size * 1.2, 2),
                    })
                    idx += 1
        self._spans[page_index] = out
        return out

    def find_span(self, page_index: int, span_id: str) -> dict | None:
        for s in self.spans(page_index):
            if s["id"] == span_id:
                return s
        return None

    # -- 操作集 -------------------------------------------------------
    def apply_op(self, op: dict) -> dict:
        typ = op.get("type")
        with self.lock:
            if typ == "text":
                result = self._op_text(op)
            elif typ == "delete":
                result = self._op_delete(op)
            elif typ == "redact":
                result = self._op_redact(op)
            elif typ == "freeText":
                result = self._op_free_text(op)
            elif typ == "image":
                result = self._op_image(op)
            elif typ == "imageUpdate":
                result = self._op_image_update(op)
            elif typ == "page":
                result = self._op_page(op)
            elif typ == "remove":
                result = self._op_remove(op)
            else:
                raise ValueError(f"未知操作类型: {typ}")
            self.version += 1
            self.redo_stack.clear()
            return result

    def _op_text(self, op: dict) -> dict:
        sid = op.get("spanId")
        page = int(op.get("page", 0))
        span = self.find_span(page, sid) if sid else None
        if span is None:
            raise ValueError(f"文字块不存在: {sid}")

        prev = next((o for o in self.ops if o.get("type") == "text" and o.get("spanId") == sid), None)
        self.ops = [o for o in self.ops if not (o.get("type") == "text" and o.get("spanId") == sid)]

        base = prev or {}
        text = str(op.get("text", base.get("text", span["text"])))
        size = float(op.get("size", base.get("size", span["size"])))
        color = op.get("color") or base.get("color") or span["color"]
        bold = bool(op.get("bold", base.get("bold", span["bold"])))
        italic = bool(op.get("italic", base.get("italic", span["italic"])))
        align = op.get("align") or base.get("align") or span["align"]
        font = op.get("font") or base.get("font") or span["font"]
        offset = op.get("offset") or base.get("offset") or [0.0, 0.0]
        line_height = float(op.get("lineHeight", base.get("lineHeight", 1.2)))

        unchanged = (
            text.strip() == span["text"].strip()
            and abs(float(offset[0])) < 0.01 and abs(float(offset[1])) < 0.01
            and abs(size - span["size"]) < 0.01
            and list(color) == list(span["color"])
            and bold == span["bold"] and italic == span["italic"]
            and align == span["align"] and font == span["font"]
            and abs(line_height - 1.2) < 0.01
        )
        if unchanged:
            return {"type": "text", "page": page, "spanId": sid, "text": span["text"]}

        merged = {
            "type": "text",
            "page": page,
            "spanId": sid,
            "bbox": span["bbox"],
            "origin": span["origin"],
            "size": size,
            "color": list(color),
            "bold": bold,
            "italic": italic,
            "align": align,
            "font": font,
            "lineHeight": line_height,
            "offset": [float(offset[0]), float(offset[1])],
            "text": text,
            "_spanText": span["text"],
        }
        self.ops.append(merged)
        return merged

    def _op_delete(self, op: dict) -> dict:
        page = int(op.get("page", 0))
        span = self.find_span(page, op.get("spanId"))
        if span is None:
            raise ValueError("文字块不存在")
        self.ops = [o for o in self.ops if o.get("spanId") != op.get("spanId")]
        rec = {
            "type": "delete", "page": page, "spanId": op["spanId"],
            "bbox": span["bbox"], "_spanText": span["text"],
        }
        self.ops.append(rec)
        return rec

    def _op_redact(self, op: dict) -> dict:
        """涂色遮盖：先真擦除原文字，再铺一块底色，视觉上等同于涂改带。"""
        page = int(op.get("page", 0))
        span = self.find_span(page, op.get("spanId"))
        if span is None:
            raise ValueError("文字块不存在")
        self.ops = [o for o in self.ops if o.get("spanId") != op.get("spanId")]
        fill = op.get("fill") or [1, 1, 1]
        fill = [max(0.0, min(1.0, float(c))) for c in fill][:3]
        rec = {
            "type": "redact", "page": page, "spanId": op["spanId"],
            "bbox": span["bbox"], "fill": fill, "_spanText": span["text"],
        }
        self.ops.append(rec)
        return rec

    def _op_free_text(self, op: dict) -> dict:
        """在空白处新增文字（不属于原文档任何文字块）。"""
        page = int(op.get("page", 0))
        size = float(op.get("size", 12))
        fid = op.get("id")

        # 已存在则原地更新，避免重复叠加
        if fid:
            exist = next((o for o in self.ops if o.get("type") == "freeText" and o.get("id") == fid), None)
            if exist is not None:
                for k in ("text", "size", "color", "bold", "italic", "align", "font", "lineHeight"):
                    if k in op:
                        exist[k] = op[k]
                if op.get("rect"):
                    exist["rect"] = [float(v) for v in op["rect"]]
                    exist["bbox"] = list(exist["rect"])
                    exist["origin"] = [exist["rect"][0], exist["rect"][1] + float(exist.get("size", 12))]
                return exist

        rect = [float(v) for v in op["rect"]]
        rec = {
            "type": "freeText",
            "id": op.get("id") or f"ft{int(time.time() * 1000) % 1000000}",
            "page": page,
            "rect": rect,
            "bbox": list(rect),
            "origin": [rect[0], rect[1] + size],
            "text": str(op.get("text", "新文字")),
            "size": size,
            "color": list(op.get("color") or [0, 0, 0]),
            "bold": bool(op.get("bold", False)),
            "italic": bool(op.get("italic", False)),
            "align": op.get("align") or "left",
            "font": op.get("font") or "微软雅黑",
            "lineHeight": float(op.get("lineHeight", 1.2)),
            "offset": [0.0, 0.0],
        }
        self.ops.append(rec)
        return rec

    def _op_image(self, op: dict) -> dict:
        rec = {
            "type": "image",
            "id": op.get("id") or f"im{int(time.time() * 1000) % 1000000}",
            "page": int(op["page"]),
            "rect": [float(v) for v in op["rect"]],
            "data": op["data"],
            "keepRatio": bool(op.get("keepRatio", True)),
        }
        self.ops.append(rec)
        return {k: v for k, v in rec.items() if k != "data"}

    def _op_image_update(self, op: dict) -> dict:
        target = next((o for o in self.ops if o.get("type") == "image" and o.get("id") == op.get("id")), None)
        if target is None:
            raise ValueError("图片不存在")
        if op.get("rect"):
            target["rect"] = [float(v) for v in op["rect"]]
        return {k: v for k, v in target.items() if k != "data"}

    def _op_remove(self, op: dict) -> dict:
        """删除一条由本编辑器新增的内容（新增文字 / 图片）。"""
        oid = op.get("id")
        before = len(self.ops)
        self.ops = [o for o in self.ops if o.get("id") != oid]
        return {"removed": before - len(self.ops), "id": oid}

    def _op_page(self, op: dict) -> dict:
        action = op.get("action")
        if action not in ("rotate", "delete", "insert"):
            raise ValueError(f"未知页面操作: {action}")
        rec = {
            "type": "page",
            "action": action,
            "page": int(op.get("page", 0)),
            "angle": int(op.get("angle", 0)),
            "width": float(op.get("width", 595)),
            "height": float(op.get("height", 842)),
        }
        # 同一页的旋转只保留最后一次；删除是幂等的
        if action == "rotate":
            self.ops = [o for o in self.ops
                        if not (o.get("type") == "page" and o.get("action") == "rotate" and o.get("page") == rec["page"])]
        if action == "delete":
            self.ops = [o for o in self.ops
                        if not (o.get("type") == "page" and o.get("action") == "delete" and o.get("page") == rec["page"])]
        self.ops.append(rec)
        return rec

    def undo(self) -> bool:
        with self.lock:
            if not self.ops:
                return False
            self.redo_stack.append(self.ops.pop())
            self.version += 1
            return True

    def redo(self) -> bool:
        with self.lock:
            if not self.redo_stack:
                return False
            self.ops.append(self.redo_stack.pop())
            self.version += 1
            return True

    def reset(self) -> None:
        with self.lock:
            self.ops.clear()
            self.redo_stack.clear()
            self.version += 1

    def ops_for(self, page: int) -> list[dict]:
        return [o for o in self.ops if o.get("page") == page]

    # -- 页面结构 ------------------------------------------------------
    def mutate_pages(self, action: str, page: int, width: float | None = None,
                     height: float | None = None) -> dict:
        """旋转 / 删除 / 新增页面。

        直接作用在基线文档上，而不是叠一层"显示期偏移"——否则文字层、
        缩略图、导出三者的页码会对不上。
        """
        with self.lock:
            doc = self._replay()

            if action == "rotate":
                if not (0 <= page < doc.page_count):
                    raise ValueError("页码超出范围")
                pg = doc[page]
                pg.set_rotation((pg.rotation + 90) % 360)
            elif action == "delete":
                if doc.page_count <= 1:
                    raise ValueError("至少需要保留一页")
                if not (0 <= page < doc.page_count):
                    raise ValueError("页码超出范围")
                doc.delete_page(page)
            elif action == "insert":
                ref = doc[0].rect
                doc.new_page(-1, width=width or ref.width, height=height or ref.height)
            else:
                raise ValueError(f"未知页面操作: {action}")

            tmp = os.path.join(WORK_DIR, f"base_{self.id}_{int(time.time() * 1000)}.pdf")
            doc.save(tmp, garbage=3, deflate=True)
            doc.close()

            self._drop_built()
            try:
                self._doc.close()
            except Exception:
                pass
            self.src = tmp
            self._doc = pymupdf.open(tmp)
            self.ops.clear()
            self.redo_stack.clear()
            self._spans.clear()
            self._render_cache.clear()
            self.version += 1

            return {
                "pageCount": self._doc.page_count,
                "pages": self.all_page_sizes(),
            }

    # -- 图片 / 检索 ---------------------------------------------------
    def images(self, page_index: int) -> list[dict]:
        return [{k: v for k, v in o.items() if k != "data"}
                for o in self.ops
                if o.get("type") == "image" and int(o.get("page", -1)) == page_index]

    def free_texts(self, page_index: int) -> list[dict]:
        return [o for o in self.ops
                if o.get("type") == "freeText" and int(o.get("page", -1)) == page_index]

    def find_text(self, query: str, limit: int = 500) -> list[dict]:
        """在整份文档里查找文字，返回命中的文字块。"""
        q = (query or "").strip()
        if not q:
            return []
        hits: list[dict] = []
        for p in range(self.page_count):
            for sp in self.spans(p):
                cur = self._current_text(sp)
                if q in cur:
                    hits.append({
                        "page": p,
                        "spanId": sp["id"],
                        "text": cur,
                        "bbox": sp["bbox"],
                        "count": cur.count(q),
                    })
                    if len(hits) >= limit:
                        return hits
        return hits

    def _current_text(self, span: dict) -> str:
        for o in self.ops:
            if o.get("type") == "text" and o.get("spanId") == span["id"]:
                return o.get("text", "")
            if o.get("type") in ("delete", "redact") and o.get("spanId") == span["id"]:
                return ""
        return span["text"]

    def replace_all(self, find: str, repl: str) -> int:
        q = (find or "")
        if not q:
            return 0
        total = 0
        for p in range(self.page_count):
            for sp in self.spans(p):
                cur = self._current_text(sp)
                if q in cur:
                    total += cur.count(q)
                    self.apply_op({
                        "type": "text", "page": p, "spanId": sp["id"],
                        "text": cur.replace(q, repl),
                    })
        return total

    # -- 重放 ---------------------------------------------------------
    def build(self) -> pymupdf.Document:
        """返回反映当前全部编辑的文档对象（带缓存，供渲染复用）。

        多页缩略图会连续命中这里，若每次都从头重放，页数一多就会明显卡顿。
        """
        with self.lock:
            if not self.ops:
                return self._doc
            if self._built is not None and self._built_ver == self.version:
                return self._built

            doc = self._replay()
            self._drop_built()
            self._built = doc
            self._built_ver = self.version
            return doc

    def _drop_built(self) -> None:
        if self._built is not None:
            try:
                self._built.close()
            except Exception:
                pass
            self._built = None
            self._built_ver = -1

    def _replay(self) -> pymupdf.Document:
        """从原始文件重放全部操作，生成一份全新的文档。"""
        doc = self._open_src()

        content_ops = [o for o in self.ops if o.get("type") != "page"]
        page_ops = [o for o in self.ops if o.get("type") == "page"]

        by_page: dict[int, list[dict]] = {}
        for op in content_ops:
            by_page.setdefault(int(op["page"]), []).append(op)

        for page_index, ops in by_page.items():
            if page_index >= doc.page_count:
                continue
            page = doc[page_index]

            # 1) 批量擦除（擦除会吞掉区域内图形，故关闭 graphics/images 处理）
            need_redact = False
            for op in ops:
                if op["type"] in ("text", "delete", "redact"):
                    b = op["bbox"]
                    rect = pymupdf.Rect(b[0] - 0.6, b[1] - 0.6, b[2] + 0.6, b[3] + 0.6)
                    page.add_redact_annot(rect, fill=None)
                    need_redact = True
            if need_redact:
                page.apply_redactions(
                    images=pymupdf.PDF_REDACT_IMAGE_NONE,
                    graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                    text=pymupdf.PDF_REDACT_TEXT_REMOVE,
                )

            # 2) 再写入新内容
            for op in ops:
                if op["type"] in ("text", "freeText"):
                    self._write_text(page, op)
                elif op["type"] == "image":
                    self._write_image(page, op)
                elif op["type"] == "redact":
                    b = op["bbox"]
                    r = pymupdf.Rect(b[0] - 0.6, b[1] - 0.6, b[2] + 0.6, b[3] + 0.6)
                    page.draw_rect(r, color=None, fill=tuple(op.get("fill") or (1, 1, 1)),
                                   overlay=True)

        # 3) 页面级操作：先旋转、再追加空白页、最后倒序删除（都基于原始页码）
        for o in page_ops:
            if o["action"] == "rotate" and 0 <= o["page"] < doc.page_count:
                page = doc[o["page"]]
                page.set_rotation((page.rotation + o["angle"]) % 360)

        for o in page_ops:
            if o["action"] == "insert":
                doc.new_page(-1, width=o["width"], height=o["height"])

        deletes = sorted({o["page"] for o in page_ops if o["action"] == "delete"}, reverse=True)
        for idx in deletes:
            if 0 <= idx < doc.page_count and doc.page_count > 1:
                doc.delete_page(idx)

        return doc

    def _write_text(self, page: pymupdf.Page, op: dict) -> None:
        text = op.get("text", "")
        if not text:
            return
        size = float(op["size"])
        color = tuple(c / 255.0 for c in op["color"])
        alias, fontfile = resolve_font(op.get("font", ""), text, op.get("bold", False))
        line_h = float(op.get("lineHeight", 1.2)) or 1.2
        lh = size * line_h
        off = op.get("offset") or [0.0, 0.0]
        dx0, dy0 = float(off[0]), float(off[1])
        x = float(op["origin"][0]) + dx0
        baseline = float(op["origin"][1]) + dy0
        b = op.get("bbox") or op.get("rect") or [x, baseline, x, baseline]
        bbox = [b[0] + dx0, b[1] + dy0, b[2] + dx0, b[3] + dy0]
        align = op.get("align", "left")

        # 用真实字形测宽，超出版心时自动折行，避免文字被页面裁掉
        fobj = get_font_obj(fontfile, alias)

        limit = max(page.rect.width - x - 16, size * 4)
        lines: list[str] = []
        for raw in text.split("\n"):
            if fobj and raw:
                lines.extend(_wrap_text(fobj, raw, size, limit))
            else:
                lines.append(raw)
        for i, line in enumerate(lines):
            if not line:
                continue
            y = baseline + i * lh
            dx = 0.0
            if align in ("center", "right"):
                try:
                    if fobj is not None:
                        w = fobj.text_length(line, fontsize=size)
                    else:
                        w = pymupdf.get_text_length(line, fontname=alias, fontsize=size)
                except Exception:
                    w = len(line) * size * 0.6
                if align == "center":
                    dx = (bbox[0] + bbox[2]) / 2.0 - w / 2.0 - x
                else:
                    dx = bbox[2] - w - x
            try:
                page.insert_text(
                    (x + dx, y), line,
                    fontsize=size, fontname=alias, fontfile=fontfile,
                    color=color, render_mode=0,
                )
            except Exception:
                # 字体异常时退回内置 CJK / Helvetica，保证不丢内容
                alias2, _ = ("china-s", None) if _has_cjk(line) else ("helv", None)
                page.insert_text((x + dx, y), line, fontsize=size, fontname=alias2, color=color)

    def _write_image(self, page: pymupdf.Page, op: dict) -> None:
        import base64
        data = op["data"]
        if "," in data and data.strip().startswith("data:"):
            data = data.split(",", 1)[1]
        raw = base64.b64decode(data)
        page.insert_image(pymupdf.Rect(op["rect"]), stream=raw, keep_proportion=op.get("keepRatio", True))

    # -- 渲染 ---------------------------------------------------------
    def render(self, page_index: int, dpi: int = 110) -> bytes:
        key = (page_index, dpi, self.version, len(self.ops))
        cached = self._render_cache.get(key)
        if cached is not None:
            return cached

        doc = self.build()
        page = doc[page_index]
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        data = pix.tobytes("png")

        self._render_cache[key] = data
        if len(self._render_cache) > 400:
            self._render_cache.clear()
            self._render_cache[key] = data
        return data

    def export(self, out_path: str) -> str:
        with self.lock:
            # 导出必须用独立副本：subset_fonts 会改写文档结构，
            # 复用渲染缓存会让后续预览出现缺字。
            doc = self._replay()
            # 关键：insert_text 会完整嵌入字体（中文字体可达 10~20MB），
            # 不做子集化会让一个 150KB 的文件膨胀到 10MB。
            try:
                doc.subset_fonts(fallback=False)
            except Exception:
                try:
                    doc.subset_fonts()
                except Exception:
                    pass
            doc.save(out_path, garbage=4, deflate=True, clean=True)
            doc.close()
        return out_path

    def close(self) -> None:
        self._drop_built()
        try:
            self._doc.close()
        except Exception:
            pass


# ---------------------------------------------------------------- 会话管理


class Engine:
    def __init__(self):
        self.sessions: dict[str, DocSession] = {}
        self._lock = threading.Lock()
        self._counter = 0

    def open(self, path: str, orig_dir: str | None = None) -> DocSession:
        path = os.path.abspath(path)
        if not os.path.exists(path):
            raise FileNotFoundError(path)
        with self._lock:
            for s in self.sessions.values():
                if s.src == path:
                    if orig_dir and not s.orig_dir:
                        s.orig_dir = orig_dir
                    return s
            self._counter += 1
            doc_id = f"d{self._counter}_{int(time.time())}"
            sess = DocSession(path, doc_id, orig_dir)
            self.sessions[doc_id] = sess
            return sess

    def get(self, doc_id: str) -> DocSession:
        sess = self.sessions.get(doc_id)
        if sess is None:
            raise KeyError("文档会话已失效，请重新打开")
        return sess

    def drop(self, doc_id: str) -> None:
        sess = self.sessions.pop(doc_id, None)
        if sess:
            sess.close()
