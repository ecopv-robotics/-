"""一次性导入开放平台凭据（clientKey / clientSecret）到 DPAPI 密文文件。

用法（推荐，值不落在命令行历史里）：
    py -3 utils/import_openapi_cred.py --env prod          # 交互式输入（getpass，屏幕不回显）
    py -3 utils/import_openapi_cred.py --env test

或从环境变量读（用于无人值守 / CI）：
    set CRM_CLIENT_KEY=...
    set CRM_CLIENT_SECRET=...
    py -3 utils/import_openapi_cred.py --env prod --from-env

脚本本身绝不打印明文，只回显脱敏结果。
"""

from __future__ import annotations

import argparse
import getpass
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from utils.openapi_credential import (  # noqa: E402
    configured_envs,
    mask,
    save_credential,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="导入开放平台凭据（DPAPI 加密落盘）")
    parser.add_argument("--env", default="prod", help="环境名：prod / test")
    parser.add_argument("--file", default="", help="密文文件路径，默认 storage/openapi.cred")
    parser.add_argument("--from-env", action="store_true", help="从 CRM_CLIENT_KEY / CRM_CLIENT_SECRET 读取")
    args = parser.parse_args()

    target = Path(args.file) if args.file else Path(__file__).resolve().parent.parent / "storage" / "openapi.cred"

    if args.from_env:
        key = os.environ.get("CRM_CLIENT_KEY", "").strip()
        secret = os.environ.get("CRM_CLIENT_SECRET", "").strip()
    else:
        print(f"导入环境: {args.env}")
        key = getpass.getpass("clientKey (输入不回显): ").strip()
        secret = getpass.getpass("clientSecret (输入不回显): ").strip()

    if not key or not secret:
        print("错误：clientKey / clientSecret 不能为空", file=sys.stderr)
        return 2

    try:
        save_credential(target, args.env, key, secret)
    except Exception as exc:
        print(f"错误：加密写入失败 {type(exc).__name__}: {exc}", file=sys.stderr)
        return 3

    # 明文用完即弃
    key = secret = ""
    print(f"已加密写入: {target}")
    print(f"  env        : {args.env}")
    print(f"  clientKey  : {mask('x' * 20)}")
    print(f"  当前已配置环境: {configured_envs(target)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
