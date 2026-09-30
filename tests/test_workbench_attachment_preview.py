from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

import workbench_server
from workbench_server import WorkbenchStore


def test_archive_preview_modal_uses_valid_dialog_selector():
    html = (Path(__file__).resolve().parents[1] / "workbench.html").read_text(
        encoding="utf-8",
    )
    assert "modal.querySelector('[role=\"dialog\"]')" in html
    assert "modal.querySelector('[role=\"dialog]')" not in html


def _make_xlsx(path):
    workbook = Workbook()
    workbook.active["A1"] = "公司中文名称"
    workbook.active["B1"] = "测试公司有限公司"
    workbook.save(path)


def test_workbench_preview_and_member_download_require_mail_token(tmp_path, monkeypatch):
    monkeypatch.setattr(workbench_server, "APP_ROOT", tmp_path)
    cache = tmp_path / "cache" / "attachments"
    cache.mkdir(parents=True)
    workbook_path = tmp_path / "form.xlsx"
    _make_xlsx(workbook_path)
    token = "hash_form.zip"
    archive_path = cache / token
    with zipfile.ZipFile(archive_path, "w") as archive:
        archive.write(workbook_path, "EPR/application.xlsx")

    store = object.__new__(WorkbenchStore)
    store.snapshot = lambda: {"mails": [{
        "id": "mail-1",
        "details": [{"attachment_files": [{
            "filename": "forms.zip", "token": token, "attachment_index": 0,
        }]}],
    }]}

    manifest = store.attachment_preview("mail-1", token, "forms.zip")
    assert manifest["workbook_count"] == 1
    preview = store.attachment_preview(
        "mail-1", token, "forms.zip", "EPR/application.xlsx", 0, 40,
    )
    assert preview["sheets"][0]["rows"][0]["cells"] == ["公司中文名称", "测试公司有限公司"]
    payload, filename, _mime, _source = store.attachment_member_bytes(
        "mail-1", token, "forms.zip", "EPR/application.xlsx",
    )
    assert filename == "application.xlsx"
    assert payload

    with pytest.raises(FileNotFoundError, match="找不到对应邮件"):
        store.attachment_preview("other-mail", token, "forms.zip")
    with pytest.raises(FileNotFoundError, match="名称或索引"):
        store.attachment_preview("mail-1", "unrelated.zip", "forms.zip")
    with pytest.raises(FileNotFoundError, match="名称或索引"):
        store.attachment_preview("mail-1", token, "unrelated.zip")
