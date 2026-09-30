"""ECOPV 独立工作台便携版入口。

这个入口只加载浏览器工作台及其数据层，不加载 PyQt、OCR 和邮件抓取模块，
从而可以打包成操作人员无需 Python 环境即可运行的轻量便携版。
"""

from __future__ import annotations

import os
import sys

from utils.runtime_paths import APP_ROOT


def main() -> None:
    os.chdir(APP_ROOT)
    from workbench_launcher import main as workbench_main

    workbench_main()


if __name__ == "__main__":
    main()
