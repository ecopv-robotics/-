"""开放平台连通性探针：验证签名方式 + token + 工单查询能否跑通。

用法：
    py -3 utils/probe_openapi.py                    # 用 config.yaml 的 env
    py -3 utils/probe_openapi.py --env test         # 指定环境
    py -3 utils/probe_openapi.py --company 关键词    # 带公司关键词查一次

输出全部脱敏：不打印 clientKey / clientSecret / sign / token 明文。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import yaml  # noqa: E402

from modules.openapi_client import OpenApiClient, OpenApiError  # noqa: E402
from utils.openapi_credential import mask, resolve_credential  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", default="", help="覆盖 config.yaml 的 env")
    ap.add_argument("--company", default="", help="公司名称/客户号关键词")
    ap.add_argument("--country", default="DE")
    ap.add_argument("--biz-type", default="1", help="1-WEEE / 2-电池法")
    ap.add_argument("--page-size", type=int, default=5)
    args = ap.parse_args()

    config = yaml.safe_load((ROOT / "config.yaml").read_text(encoding="utf-8")) or {}
    section = config.get("crm_api") or {}
    env_name = args.env or str(section.get("env") or "prod")
    profile = (section.get("profiles") or {}).get(env_name) or {}

    base_url = str(section.get("base_url") or profile.get("base_url") or "")
    client_id = str(section.get("client_id") or profile.get("client_id") or "")
    cred_file = ROOT / str(section.get("credential_file") or "storage/openapi.cred")

    cred = resolve_credential(cred_file, env_name)
    print(f"环境       : {env_name}")
    print(f"base_url   : {base_url}")
    print(f"clientId   : {client_id}")
    print(f"凭据来源   : {cred.source if cred else '未找到（先跑 utils/import_openapi_cred.py）'}")
    print(f"凭据脱敏   : {mask(cred.client_key) if cred else '-'} / {mask(cred.client_secret) if cred else '-'}")
    if not cred:
        return 2

    for mode in ("raw", "hex"):
        client = OpenApiClient(base_url, client_id, cred, logger=print, sign_mode=mode)
        print(f"\n--- 尝试签名模式: {mode} ---")
        try:
            token = client.token(force=True)
        except OpenApiError as exc:
            print(f"  失败: {str(exc)[:200]}")
            continue
        print(f"  取 token 成功: {mask(token)}")
        try:
            resp = client.query_register_list(
                customer_code=args.company,
                country_code=args.country,
                biz_type=args.biz_type,
                page_no=1,
                page_size=args.page_size,
            )
        except OpenApiError as exc:
            print(f"  查询失败: {str(exc)[:200]}")
            return 3
        result = resp.get("result") or {}
        records = result.get("records") if isinstance(result, dict) else None
        total = result.get("total") if isinstance(result, dict) else None
        print(f"  查询成功: success={resp.get('success')} total={total} 本页={len(records) if isinstance(records, list) else '?'}")
        if isinstance(records, list) and records:
            print("  首条记录的字段名:")
            print("   ", ", ".join(sorted(records[0].keys()))[:600])
        print(f"\n结论：签名模式应为 sign_mode: {mode}")
        return 0

    print("\n两种签名模式都失败；请核对 clientId / clientKey / clientSecret 与服务器时间。")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
