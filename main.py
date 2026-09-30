"""主程序入口"""
import sys
import os

from utils.runtime_paths import APP_ROOT

# 确保项目根目录在 sys.path 中
ROOT_DIR = str(APP_ROOT)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)


def main():
    os.chdir(ROOT_DIR)
    # 便携版只打包一个运行时目录：把同一可执行文件复制为
    # ECOPV_Workbench.exe 时，按文件名直接进入独立工作台。
    if "--workbench" in sys.argv or os.path.splitext(os.path.basename(sys.argv[0]))[0].lower().endswith("workbench"):
        from workbench_launcher import main as workbench_main
        workbench_main()
        return
    from gui import main as gui_main
    gui_main()


if __name__ == "__main__":
    main()
