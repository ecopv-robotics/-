import json
import pytest

from modules.field_extractor import FieldExtractor
from utils.evidence_snapshot import encode_evidence_snapshot


@pytest.mark.parametrize("record_source", ["", "附件EPR申请表：1:company-0/application.xlsx / EPR申请表"])
def test_large_archive_snapshot_preserves_all_workbook_locators_and_valid_json(record_source):
    attachments = [{"filename": "batch.zip", "structured_records": [], "sheets": []}]
    for number in range(12):
        member = f"company-{number}/application.xlsx"
        attachments[0]["sheets"].append({
            "sheet_name": "EPR申请表", "member_path": member,
            "preview_rows": [{"row_number": row, "cells": ["表格内容" * 35] * 8}
                             for row in range(1, 81)],
        })
        attachments[0]["structured_records"].append({
            "record_type": "epr_application", "attachment_name": member,
            "sheet_name": "EPR申请表", "row_number": 40,
            "customer": f"Sample {number} Limited", "cells": ["公司英文名称", f"Sample {number} Limited"],
        })
    encoded = FieldExtractor._attachment_evidence_json(attachments, record_source)
    assert len(encoded) <= 30000
    restored = json.loads(encoded)
    assert len(restored[0]["sheets"]) == 12
    assert len(restored[0]["records"]) == 12
    assert len({sheet["member_path"] for sheet in restored[0]["sheets"]}) == 12
    assert all(any(row["row_number"] == 40 for row in sheet["rows"])
               for sheet in restored[0]["sheets"])
    assert restored[0]["preview_truncated"] is True
    assert len(attachments[0]["sheets"][0]["preview_rows"]) == 80


def test_small_snapshot_does_not_change_content():
    payload = [{"filename": "form.xlsx", "sheets": [], "records": [], "text": "正常"}]
    assert json.loads(encode_evidence_snapshot(payload)) == payload


def test_excessive_structured_evidence_is_explicitly_compacted_not_broken():
    payload = [{"filename": "large.zip", "sheets": [], "records": [
        {"customer": f"Example {index} Limited", "raw_text": "x" * 1000,
         "cells": ["y" * 1000]} for index in range(1500)]}]
    encoded = encode_evidence_snapshot(payload)
    assert len(encoded) <= 30000
    restored = json.loads(encoded)[0]
    assert restored["filename"] == "large.zip"
    assert restored["record_count"] == 1500
    assert restored["preview_truncated"] is True
