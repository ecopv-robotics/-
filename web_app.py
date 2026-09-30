"""Browser entry point; desktop main.py remains available unchanged."""
import argparse
import json
import sys
import webbrowser
from urllib.request import urlopen

from PyQt5.QtCore import QCoreApplication, QLockFile, QTimer

from utils.runtime_paths import APP_ROOT
from web_run_controller import RunController, atomic_text, load_json
from workbench_server import WorkbenchServer


def main(argv=None):
    parser = argparse.ArgumentParser(description="本机邮件审核网页控制台")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args(argv)
    app = QCoreApplication(sys.argv[:1])
    storage = APP_ROOT / "storage"
    storage.mkdir(exist_ok=True)
    lock = QLockFile(str(storage / "gui-instance.lock"))
    lock.setStaleLockTime(0)
    info_path = storage / "web_runtime.json"
    if not lock.tryLock(100):
        # Never kill the desktop or an active query to obtain a web port.
        info = load_json(info_path)
        port = info.get("port")
        if isinstance(port, int) and 1 <= port <= 65535:
            try:
                url = f"http://127.0.0.1:{port}"
                with urlopen(url + "/api/run/state", timeout=2) as response:
                    if json.load(response).get("ok"):
                        if not args.no_browser:
                            webbrowser.open(url + "/run")
                        return 0
            except (OSError, ValueError):
                pass
        print("桌面程序或其他实例正在运行。请先正常关闭原桌面窗口，再启动网页版；不会强制停止任务。")
        return 2
    server = None
    try:
        controller = RunController(APP_ROOT)
        controller.closing = False
        server = WorkbenchServer(port=args.port, run_controller=controller)
        url = server.start(open_browser=False)
        atomic_text(info_path, json.dumps({"port": server.httpd.server_address[1], "build": controller.snapshot()["build"]}))
        print("网页运行控制台：" + url + "run", flush=True)
        if not args.no_browser:
            webbrowser.open(url + "run")
        timer = QTimer()
        def finish_shutdown():
            if not controller.closing or controller.busy:
                return
            try:
                server.store.wait_background_tasks(timeout=0.01)
            except RuntimeError:
                return  # Keep serving until the existing history write has really finished.
            app.quit()
        timer.timeout.connect(finish_shutdown)
        timer.start(250)
        return app.exec_()
    finally:
        if server:
            server.stop()
        lock.unlock()


if __name__ == "__main__":
    raise SystemExit(main())
