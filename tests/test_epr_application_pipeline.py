from __future__ import annotations

import io
import json
import logging
import threading
import time
import zipfile

from openpyxl import Workbook

from gui import WorkerThread
from modules.field_extractor import FieldExtractor
from utils.attachment_parser import parse_attachment
from utils.workbook_preview import download_archive_member, preview_attachment
from workbench_server import WorkbenchStore


def _xlsx_bytes(rows, sheet_name="EPR申请表"):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for row in rows:
        sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _company_record(path, value, field="company_name_zh"):
    return {
        "record_type": "epr_application",
        "attachment_name": path,
        "customer": value,
        "company_field": field,
        "field_label": "公司中文名称",
        "raw_text": f"公司中文名称 → {value}",
    }


def test_epr_xlsx_label_value_extracts_company_without_legal_suffix(tmp_path):
    workbook_path = tmp_path / "EPR申请表.xlsx"
    workbook_path.write_bytes(_xlsx_bytes([
        ["公司中文名称（必填）", "示例0d3e2552商贸行"],
        ["Company name (in English)", "BLUE WHALE TRADING"],
    ]))

    parsed = parse_attachment(str(workbook_path), workbook_path.name)
    company_records = [
        item for item in parsed["structured_records"]
        if item.get("record_type") == "epr_application"
    ]

    assert [(item["company_field"], item["customer"]) for item in company_records] == [
        ("company_name_zh", "示例0d3e2552商贸行"),
        ("company_name_en", "BLUE WHALE TRADING"),
    ]


def test_nested_zip_parser_and_preview_keep_member_chain(tmp_path):
    workbook_bytes = _xlsx_bytes([["公司中文名称", "示例0d3e2552商贸行"]])
    inner_buffer = io.BytesIO()
    with zipfile.ZipFile(inner_buffer, "w") as inner:
        inner.writestr("forms/application.xlsx", workbook_bytes)

    outer_path = tmp_path / "outer.zip"
    with zipfile.ZipFile(outer_path, "w") as outer:
        outer.writestr("batch/forms.zip", inner_buffer.getvalue())

    parsed = parse_attachment(str(outer_path), outer_path.name)
    expected_path = "batch/forms.zip/forms/application.xlsx"
    assert [sheet["member_path"] for sheet in parsed["sheets"]] == [expected_path]
    company_record = next(
        item for item in parsed["structured_records"]
        if item.get("record_type") == "epr_application"
    )
    assert company_record["attachment_name"] == expected_path

    manifest = preview_attachment(outer_path, outer_path.name)
    workbook_member = next(item for item in manifest["members"] if item["previewable"])
    assert workbook_member["path"] == "batch/forms.zip → forms/application.xlsx"
    preview = preview_attachment(
        outer_path, outer_path.name, workbook_member["member_token"]
    )
    assert preview["sheets"][0]["rows"][0]["cells"][:2] == ["公司中文名称", "示例0d3e2552商贸行"]
    payload, filename = download_archive_member(
        outer_path, outer_path.name, workbook_member["member_token"]
    )
    assert filename == "application.xlsx"
    assert payload == workbook_bytes


def test_multiple_epr_forms_bind_each_company_and_project_independently():
    attachments = [{
        "filename": "mail.zip",
        "epr_forms": [
            {"filename": "batch/first.xlsx", "projects": ["德国WEEE"]},
            {"filename": "batch/second.xlsx", "projects": ["法国包装法"]},
        ],
        "structured_records": [
            _company_record("batch/first.xlsx", "示例ce093a2c商贸行"),
            _company_record("batch/second.xlsx", "第二家贸易行"),
        ],
    }]

    result = FieldExtractor({}, [])._extract_projects(
        "两个EPR申请", "", attachments, groups=[]
    )

    assert result["epr_forms_count"] == 2
    assert result["epr_form_groups"] is True
    assert [(group["customer"], group["projects"]) for group in result["groups"]] == [
        ("示例ce093a2c商贸行", ["德国WEEE"]),
        ("第二家贸易行", ["法国包装法"]),
    ]


def test_full_extraction_keeps_each_epr_application_candidate_scope():
    attachments = [{
        "filename": "mail.zip",
        "epr_forms": [
            {"filename": "first.xlsx", "projects": ["德国WEEE"]},
            {"filename": "second.xlsx", "projects": ["法国包装法"]},
        ],
        "structured_records": [
            _company_record("first.xlsx", "示例ce093a2c商贸行"),
            {**_company_record("first.xlsx", "Fixture b5cc172d LIMITED"), "company_field": "company_name_en"},
            _company_record("second.xlsx", "第二家贸易行"),
            {**_company_record("second.xlsx", "Fixture 1b515961 LIMITED"), "company_field": "company_name_en"},
        ],
    }]

    rows = FieldExtractor({}, [], ocr_fallback=False).extract_fields({
        "subject": "两家公司的德国WEEE和法国包装法注册",
        "body_text": "",
        "attachments": attachments,
    })

    assert [(row["客户"], row["项目"]) for row in rows] == [
        ("示例ce093a2c商贸行", "德国WEEE"),
        ("第二家贸易行", "法国包装法"),
    ]
    evidence = [json.loads(row["客户候选证据"]) for row in rows]
    assert [[item["source"] for item in group] for group in evidence] == [
        ["first.xlsx", "first.xlsx"],
        ["second.xlsx", "second.xlsx"],
    ]


def test_candidate_arbitration_runs_once_per_form_and_reuses_sibling_decision():
    candidates_a = [
        {"id": "epr-company-1", "value": "候选甲有限公司", "source": "batch/first.xlsx", "evidence": "A表字段"},
        {"id": "epr-company-2", "value": "候选乙有限公司", "source": "batch/first.xlsx", "evidence": "A表字段"},
    ]
    candidates_b = [
        {"id": "epr-company-1", "value": "候选甲有限公司", "source": "batch/second.xlsx", "evidence": "B表字段"},
        {"id": "epr-company-2", "value": "候选乙有限公司", "source": "batch/second.xlsx", "evidence": "B表字段"},
    ]
    rows = []
    for index, (source, candidates) in enumerate([
        ("batch/first.xlsx", candidates_a),
        ("batch/first.xlsx", candidates_a),
        ("batch/second.xlsx", candidates_b),
        ("batch/second.xlsx", candidates_b),
    ]):
        rows.append({
            "项目": f"项目{index + 1}",
            "客户": "示例c54a3ca7有限公司",
            "客户候选证据": json.dumps(candidates, ensure_ascii=False),
            "客户提取来源": "附件EPR申请表",
            "附件名称": source,
        })

    class FakeLLM:
        model = "offline-test"

        def __init__(self):
            self.calls = []

        def validate_extracted_fields_batch(self, items):
            self.calls.append(items)
            output = {}
            for item in items:
                options = item.get("company_candidates") or []
                selected = ""
                if options:
                    selected = "epr-company-1" if options[0]["source"].endswith("first.xlsx") else "epr-company-2"
                output[item["id"]] = {
                    "status": "valid",
                    "confidence": "high",
                    "issues": [],
                    "suggestions": {},
                    "selected_company_candidate_id": selected,
                    "reason": "离线测试",
                }
            return output

    fake_llm = FakeLLM()
    worker = WorkerThread({}, mode="stage1")
    result = worker._semantic_validate_rows(
        rows, fake_llm, logging.getLogger("epr-arbitration-test"),
        mail_keys=["MAIL-1"] * len(rows),
    )

    submitted_candidate_groups = [
        item["company_candidates"]
        for batch in fake_llm.calls for item in batch if item.get("company_candidates")
    ]
    assert len(submitted_candidate_groups) == 2
    assert [row["客户"] for row in result] == [
        "候选甲有限公司", "候选甲有限公司", "候选乙有限公司", "候选乙有限公司",
    ]
    assert [row["公司候选复核状态"] for row in result] == ["已仲裁"] * 4


def test_history_excel_rebuild_is_scheduled_outside_snapshot_request(tmp_path, monkeypatch):
    import workbench_server

    state_path = tmp_path / "review-state.json"
    monkeypatch.setattr(workbench_server, "REVIEW_STATE", state_path)
    monkeypatch.setattr(workbench_server, "HISTORY_STATE", tmp_path / "history-state.json")
    primary_path = tmp_path / "primary.xlsx"
    workbook = Workbook()
    workbook.active.title = "工单待查"
    workbook.active.append(["邮件编号", "客户公司名称", "标准化项目名称"])
    workbook.save(primary_path)
    store = WorkbenchStore(
        primary_path=str(primary_path),
        review_path=str(tmp_path / "review.xlsx"),
        filtered_path=str(tmp_path / "filtered.xlsx"),
        state_path=str(state_path),
        database_path=str(tmp_path / "workbench.sqlite"),
    )
    store._history_cache_summary = {
        "enabled": True,
        "message": "已缓存",
        "completed_count": 0,
        "unfinished_count": 0,
        "completed_path": "",
        "unfinished_path": "",
        "refreshing": False,
    }
    store._history_cache_fingerprint = (("old", 0, 0),)
    monkeypatch.setattr(store, "_history_input_fingerprint", lambda: (("new", 1, 1),))
    started = threading.Event()
    release = threading.Event()

    def slow_sync(_mails, _filtered):
        started.set()
        assert release.wait(2)
        return {
            "enabled": True,
            "message": "已完成",
            "completed_count": 1,
            "unfinished_count": 2,
            "completed_path": "completed.xlsx",
            "unfinished_path": "unfinished.xlsx",
        }

    monkeypatch.setattr(store, "_sync_persistent_history", slow_sync)
    try:
        snapshot = store.snapshot()
        assert snapshot["history"]["refreshing"] is True
        assert started.wait(1)
    finally:
        release.set()
        for _ in range(100):
            if not store._history_sync_pending:
                break
            time.sleep(0.01)
    assert store._history_sync_pending is False
