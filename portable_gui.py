"""ECOPV 桌面 GUI 的独立启动入口。

便携发布包使用 PyInstaller 编译此入口，避免依赖操作人员电脑上的
Python、PyQt 或系统 PATH；运行时目录由 runtime_paths 统一解析。
"""

from __future__ import annotations

import os

from utils.runtime_paths import APP_ROOT


def main() -> None:
    os.chdir(APP_ROOT)
    from gui import main as gui_main

    gui_main()


if __name__ == "__main__":
    main()