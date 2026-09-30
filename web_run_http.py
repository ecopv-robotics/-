"""Same-origin, loopback-only transport for the local run console."""
import json
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from workbench_server import _Handler
from web_run_controller import UPLOAD_LIMIT

ASSETS = Path(__file__).resolve().parent / "web_console"


class RunHandler(_Handler):
    def _local_request(self):
        port = self.server.server_address[1]
        allowed = {f"127.0.0.1:{port}", f"localhost:{port}"}
        host = self.headers.get("Host", "")
        origin = self.headers.get("Origin")
        if host not in allowed or (origin is not None and origin != "http://" + host):
            self._send({"ok": False, "error": "只允许从本机工作台页面访问"}, 403)
            return False
        if self.headers.get("Sec-Fetch-Site") == "cross-site":
            self._send({"ok": False, "error": "不允许跨站请求"}, 403)
            return False
        return True

    def do_OPTIONS(self):
        self._send({"ok": False, "error": "运行控制台不开放跨域访问"}, 403)

    def do_GET(self):
        if not self._local_request():
            return
        parsed = urlparse(self.path)
        controller = self.server.run_controller
        try:
            assets = {"/run": ("index.html", "text/html"),
                      "/run/": ("index.html", "text/html"),
                      "/run/console.js": ("console.js", "application/javascript"),
                      "/run/console.css": ("console.css", "text/css")}
            if parsed.path in assets:
                name, mime = assets[parsed.path]
                self._send((ASSETS / name).read_bytes(), content_type=mime + "; charset=utf-8")
            elif parsed.path == "/api/run/state":
                after = int(parse_qs(parsed.query).get("after", ["0"])[0])
                self._send(controller.snapshot(after))
            elif parsed.path == "/api/run/download":
                path = controller.download(parse_qs(parsed.query).get("token", [""])[0])
                self._send_binary(path.read_bytes(), path.name,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", download=True)
            else:
                super().do_GET()
        except (ValueError, OSError) as exc:
            self._send({"ok": False, "error": controller._redact(exc)}, 400)

    def do_POST(self):
        if not self._local_request():
            return
        parsed = urlparse(self.path)
        controller = self.server.run_controller
        if getattr(controller, "closing", False):
            self._send({"ok": False, "error": "服务正在安全退出，不再接受新的修改"}, 409)
            return
        if not parsed.path.startswith("/api/run/"):
            if self.headers.get("Content-Type", "").split(";")[0] != "application/json":
                self._send({"ok": False, "error": "请使用 JSON 请求"}, 415)
                return
            # Do not change existing action/export semantics; they remain on the same origin.
            return super().do_POST()
        if self.headers.get("X-Run-Token") != controller.csrf:
            self._send({"ok": False, "error": "页面会话已更新，请刷新后重试"}, 403)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            uploading = parsed.path == "/api/run/upload"
            if length < 0 or length > (UPLOAD_LIMIT if uploading else 65536):
                raise ValueError("请求过大")
            self.connection.settimeout(30)
            raw = self.rfile.read(length)
            if len(raw) != length:
                raise ValueError("请求上传不完整，请重试")
            if uploading:
                query = parse_qs(parsed.query)
                result = controller.upload(query.get("role", [""])[0], query.get("name", [""])[0], raw)
            else:
                if self.headers.get("Content-Type", "").split(";")[0] != "application/json":
                    raise ValueError("请使用 JSON 请求")
                payload = json.loads(raw.decode("utf-8"))
                if not isinstance(payload, dict):
                    raise ValueError("请求格式不正确")
                if parsed.path == "/api/run/start":
                    result = controller.start(payload)
                elif parsed.path == "/api/run/stop":
                    result = controller.stop()
                elif parsed.path == "/api/run/settings":
                    result = controller.save_settings(payload)
                elif parsed.path == "/api/run/refresh":
                    result = controller.refresh_catalog()
                elif parsed.path == "/api/run/cache":
                    result = controller.use_cache(payload.get("token", ""))
                elif parsed.path == "/api/run/shutdown":
                    with controller.lock:
                        controller._idle()
                        controller.closing = True
                    result = {"ok": True, "message": "本机服务正在安全退出"}
                else:
                    self._send({"ok": False, "error": "接口不存在"}, 404)
                    return
            self._send(result)
        except (ValueError, OSError, TypeError) as exc:
            self._send({"ok": False, "error": controller._redact(exc)}, 400)
