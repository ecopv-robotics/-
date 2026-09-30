import json

import pytest

from modules.business_validator import inspect_row
from workbench_server import (
    WorkbenchStore, _detail_confirmation_ready,
    _is_non_company_customer_value, _prepare_workbench_row,
    _repair_attachment_fields,
)


@pytest.mark.parametrize("company", ["Fixture 327da5a9 Inc", "Fixture 0b15019e Ltd", "Fixture 72d0887b LLC"])
def test_english_companies_survive_business_validation_and_workbench(company):
    assert not _is_non_company_customer_value(company)
    assert not inspect_row({"客户": company, "项目": "波兰包装法", "需求": "注册"})
    row = _prepare_workbench_row({"客户公司名称": company, "标准化项目名称": "波兰包装法", "需求": "注册"})
    assert row["客户公司名称"] == company


@pytest.mark.parametrize("label", ["公司英文名称", "公司中文名称", "英文公司名称", "中文公司名称"])
def test_labels_cannot_be_ready_company_names(label):
    assert _is_non_company_customer_value(label)
    assert any(issue["field"] == "company" for issue in inspect_row(
        {"客户": label, "项目": "波兰包装法", "需求": "注册"}))
    assert _prepare_workbench_row({"客户公司名称": label})["客户公司名称"] == ""


def test_person_name_is_still_rejected():
    assert _is_non_company_customer_value("John Smith")
    assert any(issue["code"] == "COMPANY_PERSON_NAME" for issue in inspect_row(
        {"客户": "John Smith", "项目": "波兰包装法", "需求": "注册"}))


@pytest.mark.parametrize("source,body,expected", [
    ("", "", "Paris"),
    ("附件表格：application.xlsx / 申请表 第40行", "", "示例科技有限公司"),
    ("", "示例科技有限公司申请注册", "示例科技有限公司"),
])
def test_compact_preview_alone_cannot_prove_unique_applicant(source, body, expected):
    evidence = [{"filename": "application.xlsx", "preview_truncated": True,
                 "sheets": [{"sheet_name": "申请表", "rows": [
                     {"row_number": 40, "cells": ["公司中文名称", "示例科技有限公司"]}
                 ]}]}]
    row = {"客户公司名称": "Paris", "附件明细来源": source, "邮件正文": body,
           "附件证据": json.dumps(evidence, ensure_ascii=False)}
    assert _repair_attachment_fields(row)["客户公司名称"] == expected


def test_general_confirmation_keeps_unconfirmed_weee_in_pending_queue():
    # Build only the pure record projection; no live DB or history initialization.
    store = object.__new__(WorkbenchStore)
    row = {"_id": "sample", "_source": "工单待查", "代理": "示例代理", "置信度": "high",
           "客户公司名称": "示例科技有限公司", "标准化项目名称": "德国WEEE", "需求": "注册",
           "德国WEEE专项": "是", "德国WEEE品类明细": json.dumps([
               {"brand": "Example", "category": "小型设备", "category_class": "5"}
           ], ensure_ascii=False)}
    state = {"records": {"sample": {"status": "confirmed"}}}
    detail = store._record(row, state)
    assert detail["status"] == "ready"
    assert not detail["confirmation_ready"]
    assert not _detail_confirmation_ready(detail)
    assert state["records"]["sample"]["status"] == "confirmed"
    state["records"]["sample"]["weee_status"] = "confirmed"
    detail = store._record(row, state)
    assert detail["status"] == "confirmed"
    assert _detail_confirmation_ready(detail)
