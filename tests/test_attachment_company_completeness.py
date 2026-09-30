from __future__ import annotations

import io
import sys
import types
import zipfile

import pytest
from openpyxl import Workbook

from utils import attachment_parser as parser


def _xlsx_bytes(rows):
    workbook = Workbook()
    workbook.active.title = "申请信息"
    for row in rows:
        workbook.active.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


def _zip_bytes(members):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, payload in members:
            archive.writestr(name, payload)
    return buffer.getvalue()


def test_bilingual_labels_are_skipped_without_becoming_company_values():
    records = parser._xlsx_epr_company_records([
        ["公司中文名称", "Company name (in Chinese)", "示例甲贸易有限公司"],
        ["公司英文名称", "Company name (in English)", "Example Trading Ltd"],
    ], "申请信息")

    assert [(record["company_field"], record["customer"]) for record in records] == [
        ("company_name_zh", "示例甲贸易有限公司"),
        ("company_name_en", "Example Trading Ltd"),
    ]
    assert [(record["row_number"], record["column_number"]) for record in records] == [(1, 1), (2, 1)]


def test_inline_company_fields_read_the_value_in_the_same_cell():
    records = parser._xlsx_epr_company_records([
        ["公司中文名称：示例甲贸易有限公司"],
        ["Company name (in English): Example Trading Inc"],
        ["Applicant name: EXAMPLE APPLICANT LLC"],
    ], "申请信息")

    assert [(record["company_field"], record["customer"]) for record in records] == [
        ("company_name_zh", "示例甲贸易有限公司"),
        ("company_name_en", "Example Trading Inc"),
        ("applicant_name", "EXAMPLE APPLICANT LLC"),
    ]


def test_empty_chinese_field_does_not_capture_english_label_or_value():
    records = parser._xlsx_epr_company_records([
        ["公司中文名称", "公司英文名称", "Example Design Inc"],
    ], "申请信息")

    assert [(record["company_field"], record["customer"]) for record in records] == [
        ("company_name_en", "Example Design Inc"),
    ]


@pytest.mark.parametrize("label", [
    "公司中文名称", "公司英文名称", "中文公司名", "英文公司名",
    "Company name (in Chinese)", "Company name (in English)", "Applicant name",
])
def test_company_field_labels_are_never_candidate_values(label):
    assert parser._xlsx_customer_value(label) == ""


@pytest.mark.parametrize("company", [
    "Example Design Inc", "Example Trading Ltd", "Example Trading Limited", "Example Systems Corp",
])
def test_title_case_company_with_legal_suffix_is_not_rejected_as_a_person(company):
    assert parser._xlsx_customer_value(company) == company


def test_person_without_company_suffix_is_still_rejected():
    assert parser._xlsx_customer_value("Alice Example") == ""


@pytest.mark.parametrize("invalid_chinese", ["公司中文名称", "N/A", "法人姓名", "不适用"])
def test_invalid_chinese_column_falls_back_to_valid_english_company(invalid_chinese):
    records = parser._xlsx_structured_records([
        ["公司中文名称", "公司英文名称", "项目"],
        [invalid_chinese, "Example Trading Ltd", "德国包装法"],
    ], "主体清单")

    assert len(records) == 1
    assert records[0]["customer"] == "Example Trading Ltd"
    assert records[0]["row_number"] == 2


def test_duplicate_basenames_preserve_each_workbook_company_and_member_path(tmp_path):
    archive_path = tmp_path / "applications.zip"
    archive_path.write_bytes(_zip_bytes([
        ("first/application.xlsx", _xlsx_bytes([["公司中文名称", "示例甲贸易有限公司"]])),
        ("second/application.xlsx", _xlsx_bytes([["Company name (in English)", "Example Trading Inc"]])),
    ]))

    result = parser.parse_attachment(str(archive_path), archive_path.name)
    records = [record for record in result["structured_records"] if record["record_type"] == "epr_application"]

    assert {(record["attachment_name"], record["customer"]) for record in records} == {
        ("first/application.xlsx", "示例甲贸易有限公司"),
        ("second/application.xlsx", "Example Trading Inc"),
    }
    assert all(record["sheet_name"] == "申请信息" for record in records)
    assert len(result["sheets"]) == 2


@pytest.mark.parametrize("outer_format", ["zip", "rar"])
def test_nested_images_propagate_to_parent_without_running_ocr(tmp_path, monkeypatch, outer_format):
    monkeypatch.setattr(parser, "ATTACHMENT_CACHE_DIR", str(tmp_path / "isolated_attachments"))

    def unexpected_ocr(*args, **kwargs):
        pytest.fail("Attachment registration must not start OCR")

    monkeypatch.setattr(parser, "_get_ocr", unexpected_ocr)
    monkeypatch.setattr(parser, "_parse_image", unexpected_ocr)
    image_payload = b"anonymous-image-placeholder"
    inner_payload = _zip_bytes([("evidence/photo.png", image_payload)])
    outer_path = tmp_path / f"outer.{outer_format}"
    if outer_format == "zip":
        outer_path.write_bytes(_zip_bytes([("batch/inner.zip", inner_payload)]))
    else:
        # A fake RAR stream verifies the same aggregation path without requiring
        # an external RAR writer, a production archive, or an unpacking process.
        outer_path.write_bytes(b"anonymous-rar-placeholder")
        member = types.SimpleNamespace(
            filename="batch/inner.zip", file_size=len(inner_payload), isdir=lambda: False,
        )

        class FakeRarFile:
            def __init__(self, *args, **kwargs):
                pass

            def __enter__(self):
                return self

            def __exit__(self, *args):
                return False

            def infolist(self):
                return [member]

            def open(self, info):
                assert info is member
                return io.BytesIO(inner_payload)

        monkeypatch.setitem(sys.modules, "rarfile", types.SimpleNamespace(RarFile=FakeRarFile))
        monkeypatch.setattr(parser, "_configure_rar_tool", lambda module: "")

    result = parser.parse_attachment(str(outer_path), outer_path.name)

    assert len(result["pending_images"]) == 1
    image = result["pending_images"][0]
    assert image["member_path"] == "batch/inner.zip/evidence/photo.png"
    assert image["ocr_pending"] is True
    assert image["text_content"] == ""
    from pathlib import Path
    assert Path(image["filepath"]).read_bytes() == image_payload
    assert Path(image["filepath"]).parent == tmp_path / "isolated_attachments"
