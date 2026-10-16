# -*- coding: utf-8 -*-
"""定向回归：整封业务都被过滤时，项目级过滤记录必须降级为整封过滤。

对应 2026-10-16 实测到的真实缺陷 —— IMAP 读到 66 封，界面「邮件总数」只显示 65。
原因是某封邮件的业务全部被过滤，记录却仍标 ``filter_scope="业务项目"``，而下游
（后端 workbench_server.py 的 total_mails、前端 workbench.html 的邮件总数公式）
都约定「业务项目」**不计入**邮件总数 —— 于是这封邮件两头都不算。

用例分工：
  - 部分过滤必须保持原语义（不能被改坏，这是既有多数场景）
  - 整封过滤必须降级（本次修复）
  - 同封多项目、混合场景、非 scope 字段、键类型、空输入等边界
"""
from modules.mail_filter import downgrade_orphan_project_filters


def _record(project: str = "德国一次性塑料税") -> dict:
    return {
        "filter_scope": "业务项目",
        "filter_object": project,
        "filter_company": "某某有限公司",
        "filter_agent": "",
        "filter_reason": f"明确非目标项目：{project}",
        "hard_filter": True,
    }


def test_partial_filter_keeps_project_scope():
    """同封邮件还有业务保留时，保持「业务项目」——原有语义，不能被改坏。"""
    record = _record()
    changed = downgrade_orphan_project_filters([("mail-A", record)], {"mail-A"})
    assert changed == 0
    assert record["filter_scope"] == "业务项目"


def test_whole_mail_filtered_is_downgraded():
    """整封邮件无剩余业务 → 降级为整封过滤，否则会被漏计。"""
    record = _record()
    changed = downgrade_orphan_project_filters([("mail-B", record)], {"mail-A"})
    assert changed == 1
    assert record["filter_scope"] == ""


def test_multiple_project_records_of_same_mail_all_downgraded():
    """同一封邮件有多个被过滤项目时，全部一起降级。"""
    records = [
        ("mail-C", _record("德国一次性塑料税")),
        ("mail-C", _record("德国一次性塑料税")),
    ]
    changed = downgrade_orphan_project_filters(records, set())
    assert changed == 2
    assert all(record["filter_scope"] == "" for _, record in records)


def test_mixed_only_orphans_downgraded():
    """混合场景：只降孤儿，不动还有其他业务的邮件。"""
    keeps = _record()
    orphan = _record()
    changed = downgrade_orphan_project_filters(
        [("mail-D", keeps), ("mail-E", orphan)], {"mail-D"}
    )
    assert changed == 1
    assert keeps["filter_scope"] == "业务项目"
    assert orphan["filter_scope"] == ""


def test_non_scope_fields_untouched():
    """只改 filter_scope，其它字段（客户/对象/原因/hard_filter）一律原样保留。"""
    record = _record()
    before = {k: v for k, v in record.items() if k != "filter_scope"}
    downgrade_orphan_project_filters([("mail-F", record)], set())
    after = {k: v for k, v in record.items() if k != "filter_scope"}
    assert after == before
    assert record["filter_object"] == "德国一次性塑料税"


def test_int_and_str_mail_keys_do_not_collide():
    """mail_key 可能是 int（gui.py 里用 enumerate 的 index 兜底）也可能是 str，不能互相误判。"""
    as_int = _record()
    as_str = _record()
    changed = downgrade_orphan_project_filters([(7, as_int), ("7", as_str)], {7})
    assert changed == 1
    assert as_int["filter_scope"] == "业务项目"
    assert as_str["filter_scope"] == ""


def test_empty_and_none_inputs_are_safe():
    """空输入 / 未提供存活集合都不应抛异常。"""
    assert downgrade_orphan_project_filters([], set()) == 0
    assert downgrade_orphan_project_filters([], None) == 0
    record = _record()
    assert downgrade_orphan_project_filters([("mail-G", record)], None) == 1
    assert record["filter_scope"] == ""


def test_downgraded_records_count_as_mail_in_downstream_formula():
    """端到端口径验证：降级后该记录会被下游计入邮件总数。

    复刻 workbench_server.py:2634-2639 与 workbench.html:2220 的公式：
        total_mails = len(mails) + (filtered_mails 中 scope != '业务项目' 的条数)
    修复前这封邮件算出 0 封（左），修复后应为 1 封（右）。
    """
    mails = []  # 该邮件已被整封过滤，不在活动邮件里
    record = _record()

    def total() -> int:
        return len(mails) + sum(1 for r in [record] if str(r.get("filter_scope") or "") != "业务项目")

    assert total() == 0  # 修复前：两头都不算
    downgrade_orphan_project_filters([("mail-H", record)], set())
    assert total() == 1  # 修复后：按整封过滤计一封
