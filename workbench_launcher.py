"""独立启动可视化人工复核工作台。

用于技术人员不打开 PyQt GUI 时启动工作台。默认只绑定本机；
局域网部署时可传入 ``--host 0.0.0.0``，操作人员用 Edge 访问服务器地址。
"""

from __future__ import annotations

import argparse
import os
import socket
import time

from workbench_server import WorkbenchServer
from utils.runtime_paths import APP_ROOT


def main() -> None:
    os.chdir(APP_ROOT)
    parser = argparse.ArgumentParser(description="ECOPV 独立邮件复核工作台")
    parser.add_argument("--port", type=int, default=8765, help="本机服务端口，默认 8765")
    parser.add_argument(
        "--host",
        default=os.environ.get("ECOPV_WORKBENCH_HOST", "127.0.0.1"),
        help="监听地址；局域网部署传 0.0.0.0，默认只允许本机访问",
    )
    parser.add_argument(
        "--public-host",
        default=os.environ.get("ECOPV_WORKBENCH_PUBLIC_HOST", ""),
        help="启动提示/自动打开浏览器使用的地址；局域网部署可填服务器 IP",
    )
    parser.add_argument("--database", default="", help="可选：指定 workbench.db 路径")
    parser.add_argument("--no-browser", action="store_true", help="只启动服务，不自动打开浏览器")
    args = parser.parse_args()

    public_host = args.public_host
    if not public_host and args.host in {"0.0.0.0", "::"}:
        try:
            public_host = socket.gethostbyname(socket.gethostname())
        except OSError:
            public_host = "127.0.0.1"
    server = WorkbenchServer(
        port=args.port,
        database_path=args.database or None,
        host=args.host,
        public_host=public_host,
    )
    url = server.start(open_browser=not args.no_browser)
    print(f"ECOPV 人工复核工作台已启动：{url}")
    print(f"数据库：{server.store.database.path if server.store.database else '未启用（测试会话）'}")
    print("程序可通过 POST /api/ingest 增量写入 rows，工作台无需依赖 PyQt GUI。")
    print("关闭此窗口将停止本机工作台服务。")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        server.stop()


if __name__ == "__main__":
    main()
