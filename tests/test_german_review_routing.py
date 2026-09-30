import pytest
from openpyxl import Workbook

from modules.mail_filter import MailFilter, is_internal_sender_address
from workbench_server import WorkbenchStore, _is_ecopv_internal_sender


@pytest.mark.parametrize("sender", [
    "fixture@intake.example.invalid", "Fixture <FIXTURE@intake.example.invalid>",
    "fixture@sub.intake.example.invalid", "fixture@ecopv-com.example.invalid",
    "fixture@ecopv-cn.example.invalid", "fixture@ecopv-net.example.invalid",
])
def test_internal_sender_shared_by_scan_and_historical_import(sender):
    assert is_internal_sender_address(sender)
    assert MailFilter._is_internal_sender({"sender_email": sender})
    assert _is_ecopv_internal_sender(sender)


@pytest.mark.parametrize("sender", [
    "fixture@example.test", "fixture-3142062a@example.test",
    "fixture@not-intake.example.invalid", "fixture@intake.example.invalid.example.test", "",
])
def test_external_sender_is_not_filtered_by_recipient_or_body(sender):
    assert not is_internal_sender_address(sender)
    assert not MailFilter._is_internal_sender({
        "sender_email": sender, "recipient": "fixture@intake.example.invalid",
        "body_text": "Quoted address fixture@intake.example.invalid",
    })


def test_internal_business_mail_is_hard_filtered_before_llm():
    # This branch runs before any configurable rule or network/LLM call.
    engine = MailFilter.__new__(MailFilter)
    engine._log = lambda message: None
    valid, filtered = engine.filter_mails([{
        "sender_email": "fixture@intake.example.invalid",
        "subject": "德国WEEE和电池法新增注册", "body_text": "注册资料",
    }])
    assert not valid
    assert len(filtered) == 1
    assert filtered[0]["hard_filter"] is True
    assert filtered[0]["llm_eligible"] is False


def test_old_internal_rows_remain_auditable_but_leave_active_queue(tmp_path):
    primary = tmp_path / "primary.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["邮件编号", "明细编号", "发件人邮箱", "发件日期", "邮件主题",
                  "代理", "客户公司名称", "标准化项目名称", "需求", "置信度"])
    for index, sender in enumerate(["fixture@intake.example.invalid", "fixture@example.test"]):
        sheet.append([f"MAIL-{index}", f"DETAIL-{index}", sender,
                      "2026-09-30 09:00:00", "Anonymous fixture", "Example Agent",
                      "Example LLC", "德国电池法", "注册", "high"])
    workbook.save(primary)
    workbook.close()
    original = primary.read_bytes()
    store = WorkbenchStore(primary_path=str(primary), review_path=str(tmp_path / "review.xlsx"),
                           filtered_path=str(tmp_path / "filtered.xlsx"),
                           state_path=str(tmp_path / "state.json"), test_mode=True)
    snapshot = store.snapshot()
    assert [mail["sender"] for mail in snapshot["mails"]] == ["fixture@example.test"]
    assert any(mail["sender"] == "fixture@intake.example.invalid" for mail in snapshot["filtered_mails"])
    assert primary.read_bytes() == original
