"""PDF 编辑器 —— 本地后端服务。

零第三方 Web 依赖，仅使用 Python 标准库 + PyMuPDF。
只监听 127.0.0.1，不对外暴露。
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
import traceback
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import desktop
import paths
import pdf_engine

paths.setup_logging()

APP_DIR = paths.DATA_DIR
STATIC_DIR = os.path.join(APP_DIR, "static")
WORK_DIR = paths.WORK_DIR
DIALOG_PY = os.path.join(APP_DIR, "native_dialog.py")

os.makedirs(WORK_DIR, exist_ok=True)

ENGINE = pdf_engine.Engine()
EXPORT_LOCK = threading.Lock()
LAST_SAVE_DIR = {"dir": ""}


def default_save_dir(sess) -> str:
    """导出时默认定位到用户自己的目录，而不是内部临时目录。"""
    home = os.path.expanduser("~")
    for cand in (LAST_SAVE_DIR["dir"], getattr(sess, "orig_dir", ""),
                 os.path.join(home, "Desktop"), os.path.join(home, "Documents"),
                 os.path.join(home, "Downloads")):
        if cand and os.path.isdir(cand):
            return cand
    return home

MIME = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
    ".json": "application/json; charset=utf-8",
}


def pick_dialog(mode: str, default_name: str = "", initial_dir: str = "") -> str:
    """弹出系统文件对话框。

    三种环境分别处理：
      1. 有 pywebview 窗口 → 用它的原生对话框（体验最好，且能挂到窗口上）
      2. PyInstaller 打包但走 Edge 应用模式 → 本进程直接调 Win32，不依赖外部解释器
      3. 源码运行 → 独立子进程，避免阻塞与焦点问题
    """
    # 1) pywebview
    try:
        win = desktop.get_window()
        if win is not None:
            import webview

            fd = getattr(webview, "FileDialog", None)
            if fd is not None:
                kind = fd.SAVE if mode == "save" else fd.OPEN
            else:
                kind = webview.SAVE_DIALOG if mode == "save" else webview.OPEN_DIALOG
            kwargs = {"file_types": ("PDF 文件 (*.pdf)", "所有文件 (*.*)")}
            if default_name:
                kwargs["save_filename"] = default_name
            if initial_dir and os.path.isdir(initial_dir):
                kwargs["directory"] = initial_dir
            res = win.create_file_dialog(kind, **kwargs)
            if not res:
                return ""
            return res if isinstance(res, str) else str(res[0])
    except Exception:
        pass

    # 2) 打包环境，直接调 Win32
    if getattr(sys, "frozen", False):
        try:
            import native_dialog
            return native_dialog._run(mode, default_name, initial_dir)
        except Exception:
            return ""

    # 3) 源码环境，子进程
    exe = sys.executable
    if os.name == "nt":
        cand = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if os.path.exists(cand):
            exe = cand
    args = [exe, DIALOG_PY, mode, default_name, initial_dir]
    creationflags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    try:
        r = subprocess.run(args, capture_output=True, text=True, timeout=300,
                           creationflags=creationflags, encoding="utf-8", errors="ignore")
        return (r.stdout or "").strip()
    except Exception:
        return ""


class Handler(BaseHTTPRequestHandler):
    server_version = "LocalPDF/1.0"
    protocol_version = "HTTP/1.1"

    # ------------------------------------------------------------ 基础
    def log_message(self, fmt, *args):  # 静默，避免控制台噪声
        pass

    def log_error(self, fmt, *args):
        pass

    def handle_one_request(self):
        # 浏览器刷新/切换页面会主动断开连接，属于正常现象，不必打印堆栈
        try:
            super().handle_one_request()
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
            self.close_connection = True

    def _send(self, code: int, body: bytes, ctype: str = "application/json; charset=utf-8",
              extra: dict | None = None) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if extra:
            for k, v in extra.items():
                self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionAbortedError):
            pass

    def _json(self, data, code: int = 200) -> None:
        self._send(code, json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def _err(self, msg: str, code: int = 400) -> None:
        self._json({"ok": False, "error": msg}, code)

    def _body_json(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            return {}
        raw = self.rfile.read(n)
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            return {}

    # ------------------------------------------------------------ GET
    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = urllib.parse.unquote(parsed.path)
        q = urllib.parse.parse_qs(parsed.query)

        try:
            if path.startswith("/api/"):
                return self._api_get(path, q)

            rel = "index.html" if path in ("/", "") else path.lstrip("/")
            full = os.path.normpath(os.path.join(STATIC_DIR, rel))
            if not full.startswith(os.path.normpath(STATIC_DIR)) or not os.path.isfile(full):
                return self._send(404, b"Not Found", "text/plain; charset=utf-8")

            ext = os.path.splitext(full)[1].lower()
            with open(full, "rb") as f:
                data = f.read()
            return self._send(200, data, MIME.get(ext, "application/octet-stream"))
        except KeyError as e:
            return self._err(str(e), 404)
        except Exception as e:
            traceback.print_exc()
            return self._err("服务器内部错误：" + str(e), 500)

    def _api_get(self, path: str, q: dict) -> None:
        def one(k, default=""):
            v = q.get(k)
            return v[0] if v else default

        if path == "/api/ping":
            return self._json({"app": "local-pdf-editor", "version": 1})

        if path == "/api/page":
            doc = ENGINE.get(one("doc"))
            n = int(one("n", "0"))
            dpi = max(48, min(400, int(one("dpi", "120"))))
            png = doc.render(n, dpi)
            return self._send(200, png, "image/png")

        if path == "/api/spans":
            doc = ENGINE.get(one("doc"))
            n = int(one("n", "0"))
            return self._json({
                "ok": True,
                "spans": doc.spans(n),
                "ops": [o for o in doc.ops if o.get("page") == n],
                "images": doc.images(n),
                "freeTexts": doc.free_texts(n),
                "width": doc.page_size(n)[0],
                "height": doc.page_size(n)[1],
            })

        if path == "/api/find":
            doc = ENGINE.get(one("doc"))
            return self._json({"ok": True, "hits": doc.find_text(one("q"))})

        if path == "/api/info":
            doc = ENGINE.get(one("doc"))
            return self._json({
                "ok": True,
                "docId": doc.id,
                "name": doc.name,
                "src": doc.src,
                "pageCount": doc.page_count,
                "pages": [doc.page_size(i) for i in range(doc.page_count)],
                "opCount": len(doc.ops),
                "canUndo": len(doc.ops) > 0,
                "canRedo": len(doc.redo_stack) > 0,
            })

        return self._err("未知接口", 404)

    # ------------------------------------------------------------ POST
    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = urllib.parse.unquote(parsed.path)
        q = urllib.parse.parse_qs(parsed.query)

        try:
            if path.startswith("/api/"):
                return self._api_post(path, q)
            return self._err("未知接口", 404)
        except FileNotFoundError as e:
            return self._err(f"文件不存在：{e}", 404)
        except KeyError as e:
            return self._err(str(e), 404)
        except Exception as e:
            traceback.print_exc()
            return self._err(str(e), 500)

    def _api_post(self, path: str, q: dict) -> None:
        if path == "/api/dialog/open":
            start = LAST_SAVE_DIR["dir"] or os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.isdir(start):
                start = os.path.expanduser("~")
            p = pick_dialog("open", initial_dir=start)
            if not p:
                return self._json({"ok": False, "cancelled": True})
            sess = ENGINE.open(p, orig_dir=os.path.dirname(os.path.abspath(p)))
            LAST_SAVE_DIR["dir"] = os.path.dirname(os.path.abspath(p))
            return self._json(self._doc_payload(sess))

        if path == "/api/openlocal":
            target = (data.get("path") or "").strip()
            if not target or not os.path.exists(target):
                return self._err("文件不存在：" + target, 404)
            sess = ENGINE.open(target, orig_dir=os.path.dirname(os.path.abspath(target)))
            return self._json(self._doc_payload(sess))

        if path == "/api/upload":
            name = (q.get("name") or ["upload.pdf"])[0]
            n = int(self.headers.get("Content-Length") or 0)
            safe = os.path.basename(name).replace("/", "_").replace("\\", "_")
            if not safe.lower().endswith(".pdf"):
                safe += ".pdf"
            sub = os.path.join(WORK_DIR, f"u{int(time.time()*1000)}")
            os.makedirs(sub, exist_ok=True)
            dest = os.path.join(sub, safe)
            remain = n
            with open(dest, "wb") as f:
                while remain > 0:
                    chunk = self.rfile.read(min(1 << 20, remain))
                    if not chunk:
                        break
                    f.write(chunk)
                    remain -= len(chunk)
            sess = ENGINE.open(dest)
            return self._json(self._doc_payload(sess))

        data = self._body_json()

        if path == "/api/edit":
            sess = ENGINE.get(data.get("doc", ""))
            result = sess.apply_op(data.get("op") or {})
            return self._json({"ok": True, "op": result, "opCount": len(sess.ops)})

        if path == "/api/undo":
            sess = ENGINE.get(data.get("doc", ""))
            return self._json({"ok": sess.undo(), "opCount": len(sess.ops)})

        if path == "/api/redo":
            sess = ENGINE.get(data.get("doc", ""))
            return self._json({"ok": sess.redo(), "opCount": len(sess.ops)})

        if path == "/api/replace":
            sess = ENGINE.get(data.get("doc", ""))
            n = sess.replace_all(data.get("find", ""), data.get("replace", ""))
            return self._json({"ok": True, "count": n, "opCount": len(sess.ops)})

        if path == "/api/pageop":
            sess = ENGINE.get(data.get("doc", ""))
            r = sess.mutate_pages(
                str(data.get("action", "")),
                int(data.get("page", 0)),
                data.get("width"),
                data.get("height"),
            )
            return self._json({"ok": True, **r, "opCount": len(sess.ops)})

        if path == "/api/reset":
            sess = ENGINE.get(data.get("doc", ""))
            sess.reset()
            return self._json({"ok": True, "opCount": 0})

        if path == "/api/export":
            sess = ENGINE.get(data.get("doc", ""))
            out = (data.get("path") or "").strip()
            if not out:
                stem = os.path.splitext(sess.name)[0]
                out = pick_dialog("save", default_name=f"{stem}_已编辑.pdf",
                                  initial_dir=default_save_dir(sess))
                if not out:
                    return self._json({"ok": False, "cancelled": True})
            with EXPORT_LOCK:
                sess.export(out)
            LAST_SAVE_DIR["dir"] = os.path.dirname(os.path.abspath(out))
            return self._json({"ok": True, "path": out, "dir": os.path.dirname(os.path.abspath(out))})

        if path == "/api/quickexport":
            sess = ENGINE.get(data.get("doc", ""))
            base_dir = getattr(sess, "orig_dir", "") or os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.isdir(base_dir):
                base_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            stem = os.path.splitext(sess.name)[0]
            out = os.path.join(base_dir, f"{stem}_已编辑.pdf")
            n = 2
            while os.path.exists(out):
                out = os.path.join(base_dir, f"{stem}_已编辑({n}).pdf")
                n += 1
            with EXPORT_LOCK:
                sess.export(out)
            LAST_SAVE_DIR["dir"] = base_dir
            return self._json({"ok": True, "path": out})

        if path == "/api/download":
            sess = ENGINE.get(data.get("doc", ""))
            tmp = os.path.join(WORK_DIR, f"export_{int(time.time()*1000)}.pdf")
            sess.export(tmp)
            with open(tmp, "rb") as f:
                blob = f.read()
            os.remove(tmp)
            stem = os.path.splitext(sess.name)[0]
            fname = urllib.parse.quote(f"{stem}_已编辑.pdf")
            return self._send(200, blob, "application/pdf",
                              {"Content-Disposition": f"attachment; filename*=UTF-8''{fname}"})

        if path == "/api/reveal":
            target = (data.get("path") or "").strip()
            if target and os.path.exists(target):
                try:
                    if os.name == "nt":
                        subprocess.Popen(["explorer", f"/select,{os.path.normpath(target)}"])
                    else:
                        subprocess.Popen(["open", "-R", target])
                except Exception as e:
                    return self._err(f"无法打开文件夹：{e}")
            return self._json({"ok": True})

        if path == "/api/close":
            ENGINE.drop(data.get("doc", ""))
            return self._json({"ok": True})

        return self._err("未知接口", 404)

    def _doc_payload(self, sess) -> dict:
        return {
            "ok": True,
            "docId": sess.id,
            "name": sess.name,
            "src": sess.src,
            "pageCount": sess.page_count,
            "pages": [sess.page_size(i) for i in range(sess.page_count)],
            "opCount": len(sess.ops),
        }


def find_port(start: int = 8765, tries: int = 40) -> int:
    """在 127.0.0.1 上找一个真正空闲的端口。

    Windows 下 SO_REUSEADDR 会允许两个进程绑定同一端口（请求被随机分发，
    表现为文档会话随机失效），因此这里必须用 SO_EXCLUSIVEADDRUSE 做探测。
    """
    import socket

    for p in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if os.name == "nt":
                s.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
            else:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind(("127.0.0.1", p))
                s.listen(1)
                return p
            except OSError:
                continue
    raise RuntimeError("找不到可用端口")


class LocalServer(ThreadingHTTPServer):
    daemon_threads = True
    # Windows 上绝不能开地址复用，否则会与已运行实例抢占同一端口
    allow_reuse_address = os.name != "nt"

    def server_bind(self):
        if os.name == "nt":
            import socket
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


APP_TAG = "local-pdf-editor"


def existing_instance() -> str | None:
    """若已有实例在 8765 运行，返回它的地址。"""
    import socket
    import urllib.request

    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.35)
            if s.connect_ex(("127.0.0.1", 8765)) != 0:
                return None
    except OSError:
        return None

    try:
        with urllib.request.urlopen("http://127.0.0.1:8765/api/ping", timeout=1.5) as r:
            if json.loads(r.read().decode()).get("app") == APP_TAG:
                return "http://127.0.0.1:8765/"
    except Exception:
        return None
    return None


def cleanup_work(max_age_hours: int = 24, max_items: int = 60) -> None:
    """清理过期的上传副本与页面基线文件，避免 work 目录无限膨胀。"""
    import shutil

    def drop(full, name):
        try:
            if os.path.isdir(full):
                shutil.rmtree(full, ignore_errors=True)
            elif name.endswith(".pdf"):
                os.remove(full)
        except Exception:
            pass

    try:
        items = [(os.path.join(WORK_DIR, n), n, os.path.getmtime(os.path.join(WORK_DIR, n)))
                 for n in os.listdir(WORK_DIR)]
    except Exception:
        return

    cutoff = time.time() - max_age_hours * 3600
    for full, name, mtime in items:
        if mtime < cutoff:
            drop(full, name)

    # 数量兜底：太旧的不论时间，保留最近的 max_items 个
    items.sort(key=lambda x: x[2], reverse=True)
    for full, name, _ in items[max_items:]:
        drop(full, name)


def main() -> None:
    cleanup_work()

    # 支持把 PDF 拖到 exe 图标上直接打开
    run_path = ""
    for a in sys.argv[1:]:
        if a.lower().endswith(".pdf") and os.path.exists(a):
            run_path = os.path.abspath(a)
            break

    running = existing_instance()
    if running and "--allow-multi" not in sys.argv:
        print("已有编辑器在运行，正在打开：", running)
        if "--no-window" not in sys.argv:
            url = running + (("?file=" + urllib.parse.quote(run_path)) if run_path else "")
            desktop.open_window(url)
        return

    port = find_port()
    httpd = LocalServer(("127.0.0.1", port), Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()

    url = f"http://127.0.0.1:{port}/"
    if run_path:
        url += "?file=" + urllib.parse.quote(run_path)

    print("=" * 52)
    print("  PDF 编辑器 已启动（本地运行）")
    print(f"  地址: {url}")
    print("  关闭窗口即退出程序")
    print("=" * 52)

    try:
        if "--no-window" in sys.argv or os.environ.get("PDF_EDITOR_NO_BROWSER"):
            # 开发/测试模式：不弹窗口，只跑服务
            while True:
                time.sleep(3600)
        else:
            desktop.open_window(url)
    except KeyboardInterrupt:
        pass
    finally:
        httpd.shutdown()
        httpd.server_close()
        print("已退出")


if __name__ == "__main__":
    main()
