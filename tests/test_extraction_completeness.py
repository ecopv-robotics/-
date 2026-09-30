"""申请人及公司项目关系必须完整保留；所有示例匿名且完全离线。"""
import socket

import pytest

from modules.field_extractor import FieldExtractor
from modules.business_validator import inspect_row


@pytest.fixture(autouse=True)
def deny_network(monkeypatch):
    def blocked(*args, **kwargs):
        raise AssertionError("Completeness tests must not access the network")
    monkeypatch.setattr(socket.socket, "connect", blocked)
    monkeypatch.setattr(socket.socket, "connect_ex", blocked)
    monkeypatch.setattr(socket, "create_connection", blocked)


def extractor(llm=None):
    return FieldExtractor({}, [], llm_client=llm, ocr_fallback=False)


def record(member, company, field="company_name_zh", sheet="申请信息"):
    return {"record_type": "epr_application", "attachment_name": member,
            "sheet_name": sheet, "customer": company, "company_field": field,
            "raw_text": f"申请公司字段 → {company}"}


def extract(subject, body="", attachments=None, llm=None):
    return extractor(llm).extract_fields({"subject": subject, "body_text": body,
                                         "attachments": attachments or []})


def pairs(rows):
    return [(row["客户"], row["项目"]) for row in rows]


def test_four_plain_forms_preserve_four_bilingual_applicants():
    class UnusedModel:
        enabled = True

        def extract_fields_llm(self, *args):
            raise AssertionError("Complete applicant groups must not request model extraction")

    records = []
    for index, company in enumerate(["甲甲科技有限公司", "乙乙贸易有限公司", "丙丙制造有限公司", "丁丁商贸有限公司"]):
        member = f"applicant-{index}/registration.xlsx"
        records.extend([record(member, company),
                        record(member, f"Applicant {index} Fixture 453f626e LLC", "company_name_en")])
    rows = extract("德国包装法注册4家", attachments=[
        {"filename": "applications.zip", "structured_records": records}], llm=UnusedModel())
    assert len(rows) == 4
    assert [row["客户"] for row in rows] == ["甲甲科技有限公司", "乙乙贸易有限公司", "丙丙制造有限公司", "丁丁商贸有限公司"]
    assert {row["项目"] for row in rows} == {"德国包装法"}
    assert all(row["字段LLM状态"] == "未调用" for row in rows)


def test_mixed_checked_and_plain_forms_keep_both_applicants():
    rows = extract("PL包装法2家新注册", attachments=[{
        "filename": "applications.zip",
        "structured_records": [record("first/泛欧EPR申请表.xlsx", "Fixture 327da5a9 Inc", "company_name_en"),
                               record("second/波兰EPR申请表.xlsx", "Fixture 72d0887b LLC", "company_name_en")],
        "epr_forms": [{"filename": "first/泛欧EPR申请表.xlsx", "projects": ["波兰包装法"]}],
    }])
    assert pairs(rows) == [("Fixture 327da5a9 Inc", "波兰包装法"), ("Fixture 72d0887b LLC", "波兰包装法")]


def test_plain_forms_with_multiple_global_projects_keep_unknown_mapping():
    rows = extract("德国WEEE和法国包装法注册", attachments=[{
        "filename": "applications.zip",
        "structured_records": [record("first.xlsx", "甲甲科技有限公司"), record("second.xlsx", "乙乙贸易有限公司")],
    }])
    assert pairs(rows) == [("甲甲科技有限公司", ""), ("乙乙贸易有限公司", "")]
    assert all(row["置信度"] == "需人工确认" for row in rows)


def test_plain_forms_use_member_project_without_global_cartesian_product():
    rows = extract("德国WEEE和法国包装法注册", attachments=[{
        "filename": "applications.zip",
        "structured_records": [record("first/德国WEEE申请表.xlsx", "甲甲科技有限公司"),
                               record("second/法国包装法申请表.xlsx", "乙乙贸易有限公司")],
    }])
    assert pairs(rows) == [("甲甲科技有限公司", "德国WEEE"), ("乙乙贸易有限公司", "法国包装法")]


def test_same_applicant_repeated_on_two_sheets_is_not_an_extra_company():
    rows = extract("德国包装法注册", attachments=[{
        "filename": "applications.zip",
        "structured_records": [record("first.xlsx", "甲甲科技有限公司", sheet="注册信息"),
                               record("first.xlsx", "甲甲科技有限公司", sheet="产品信息"),
                               record("second.xlsx", "乙乙贸易有限公司")],
    }])
    assert pairs(rows) == [("甲甲科技有限公司", "德国包装法"), ("乙乙贸易有限公司", "德国包装法")]


def test_distinct_applicants_on_distinct_sheets_remain_distinct():
    rows = extract("德国包装法注册", attachments=[{
        "filename": "applications.xlsx",
        "structured_records": [record("applications.xlsx", "甲甲科技有限公司", sheet="申请人甲"),
                               record("applications.xlsx", "乙乙贸易有限公司", sheet="申请人乙")],
    }])
    assert len(rows) == 2


def test_equal_member_names_in_different_archives_remain_distinct():
    rows = extract("德国包装法注册", attachments=[
        {"filename": "first.zip", "structured_records": [record("registration.xlsx", "甲甲科技有限公司")]},
        {"filename": "second.zip", "structured_records": [record("registration.xlsx", "乙乙贸易有限公司")]},
    ])
    assert pairs(rows) == [("甲甲科技有限公司", "德国包装法"), ("乙乙贸易有限公司", "德国包装法")]


@pytest.mark.parametrize("label", ["公司英文名称", "公司中文名称", "英文公司名称", "中文公司名称"])
def test_company_labels_are_not_applicants(label):
    assert extractor()._sanitize_customer_result({"customer": label, "source": "test"})["customer"] == ""


def test_unique_subject_project_is_inherited_by_plain_company_list():
    rows = extract("瑞典包装法5条新注册", "\n".join([
        "甲甲科技有限公司", "乙乙商贸行", "丙丙电子商行", "丁丁制造有限公司", "示例95bb7d04有限公司",
    ]))
    assert len(rows) == 5
    assert {row["项目"] for row in rows} == {"瑞典包装法"}


def test_html_space_numbered_list_keeps_business_name_suffix():
    companies = ["甲甲科技有限公司", "乙乙商贸行", "丙丙电子商行", "丁丁制造有限公司", "示例95bb7d04有限公司"]
    body = "\n".join(f"{index}.&nbsp;{name}\n瑞典包装法2026年新注册" for index, name in enumerate(companies, 1))
    rows = extract("瑞典包装法5条", body)
    assert pairs(rows) == [(name, "瑞典包装法") for name in companies]


def test_plain_company_list_does_not_include_signature_company():
    rows = extract("瑞典包装法新注册", "甲甲科技有限公司\n乙乙商贸行\nBest regards,\n代理示例咨询有限公司")
    assert pairs(rows) == [("甲甲科技有限公司", "瑞典包装法"), ("乙乙商贸行", "瑞典包装法")]


def test_partial_model_records_cannot_replace_five_applicants_with_three():
    companies = ["甲甲科技有限公司", "乙乙商贸有限公司", "丙丙制造有限公司", "丁丁贸易有限公司", "示例2fc42e17有限公司"]

    class OfflineResponse:
        enabled = True

        def extract_fields_llm(self, *args):
            return {"需求": "新注册", "记录": [
                {"客户": name, "项目": ["瑞典包装法"], "confidence": "high"}
                for name in companies[:3]
            ]}

    # 正文共用项目行使原有规则将关联标为待确认，从而触发记录级补充。
    rows = extract("瑞典包装法注册", "瑞典包装法\n" + "\n".join(companies), llm=OfflineResponse())
    assert pairs(rows) == [(name, "瑞典包装法") for name in companies]
    assert all(row["置信度"] == "需人工确认" for row in rows)


def test_mapping_guard_preserves_each_confirmed_company_project_pair():
    current = [{"customer": "甲甲科技有限公司", "projects": ["德国WEEE"]},
               {"customer": "乙乙贸易有限公司", "projects": ["法国包装法"]}]
    swapped = [{"customer": "甲甲科技有限公司", "projects": ["法国包装法"]},
               {"customer": "乙乙贸易有限公司", "projects": ["德国WEEE"]}]
    assert not extractor()._record_mapping_covers(current, swapped)


def test_ascii_country_codes_support_adjacent_chinese_without_matching_words():
    instance = extractor()
    assert [item["standard_name"] for item in instance._extract_projects_by_rules("PL包装法")] == ["波兰包装法"]
    assert instance._extract_projects_by_rules("SAMPLE包装法") == []


def test_two_cancellation_companies_with_inc_and_llc_are_kept():
    rows = extract("PL撤单2家", "代理+Fixture 327da5a9 Inc+波兰包装法\n代理+Fixture 72d0887b LLC+波兰包装法")
    assert pairs(rows) == [("Fixture 327da5a9 Inc", "波兰包装法"), ("Fixture 72d0887b LLC", "波兰包装法")]
    assert all(inspect_row(row) == [] for row in rows)


def test_chinese_company_service_prefix_is_removed_and_codes_stay_project_specific():
    rows = extract("德国WEEE EG3101 甲甲科技有限公司+德国电池 BG1901 甲甲科技有限公司注册")
    assert pairs(rows) == [("甲甲科技有限公司", "德国WEEE"), ("甲甲科技有限公司", "德国电池法")]
    assert [row["客户编号"] for row in rows] == ["EG3101", "BG1901"]


def test_attachment_applicants_do_not_remove_explicit_third_body_applicant():
    rows = extract("德国包装法注册", "\n".join([
        "甲甲科技有限公司+德国包装法", "乙乙贸易有限公司+德国包装法", "丙丙制造有限公司+德国包装法",
    ]), [{"filename": "batch.zip", "structured_records": [
        record("a.xlsx", "甲甲科技有限公司"), record("b.xlsx", "乙乙贸易有限公司"),
    ]}])
    assert pairs(rows) == [("甲甲科技有限公司", "德国包装法"), ("乙乙贸易有限公司", "德国包装法"), ("丙丙制造有限公司", "德国包装法")]


def test_single_explicit_body_applicant_is_not_lost_to_preferred_form_candidate():
    rows = extract("德国包装法注册", "丙丙制造有限公司+德国包装法", [{
        "filename": "batch.zip", "structured_records": [
            record("a.xlsx", "甲甲科技有限公司"), record("b.xlsx", "乙乙贸易有限公司"),
        ],
    }])
    assert {row["客户"] for row in rows} == {"甲甲科技有限公司", "乙乙贸易有限公司", "丙丙制造有限公司"}


@pytest.mark.parametrize("actual_sheet", [None, "申请人甲"])
def test_reordered_sheet_file_number_does_not_assign_project_to_wrong_applicant(actual_sheet):
    form = {"filename": "reordered.xlsx", "sheet": "xl/worksheets/sheet1.xml", "projects": ["德国包装法"]}
    if actual_sheet:
        form["sheet_name"] = actual_sheet
    groups = extractor()._application_record_groups([{
        "filename": "reordered.xlsx",
        "structured_records": [record("reordered.xlsx", "甲甲科技有限公司", sheet="申请人甲"),
                               record("reordered.xlsx", "乙乙贸易有限公司", sheet="申请人乙")],
        "sheets": [{"sheet_name": "申请人乙"}, {"sheet_name": "申请人甲"}],
        "epr_forms": [form],
    }], "注册资料", "", [])
    assert not any(g["customer"] == "乙乙贸易有限公司" and g["projects"] for g in groups)
    assert any(g["customer"] == "甲甲科技有限公司" and g["projects"] == ["德国包装法"] for g in groups) == bool(actual_sheet)
    if not actual_sheet:
        assert all(g["needs_review"] for g in groups)


def test_instruction_sheet_is_not_an_applicant_and_same_company_keeps_battery_order():
    rows = extract("德国WEEE EG3101\t甲甲科技有限公司+BG1901\t甲甲科技有限公司 小型非光伏设备+便携式电池+Brand", "资料查收", [{
        "filename": "EPR申请表.xlsx",
        "structured_records": [record("EPR申请表.xlsx", "甲甲科技有限公司"),
                               record("EPR申请表.xlsx", "Fixture e054fa73 LLC", "company_name_en"),
                               record("EPR申请表.xlsx", "针对西班牙包装法业务", "company_name", "资料清单"),
                               record("EPR申请表.xlsx", "针对德国WEEE、德国电池法、西班牙包装法业务", "company_name", "资料清单")],
    }])
    assert pairs(rows) == [("甲甲科技有限公司", "德国WEEE"), ("甲甲科技有限公司", "德国电池法")]
    assert [row["客户编号"] for row in rows] == ["EG3101", "BG1901"]
