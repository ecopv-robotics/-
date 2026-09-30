from openpyxl import Workbook
from workbench_server import WorkbenchStore


def make_store(tmp_path):
    primary = tmp_path / "primary.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["邮件编号", "明细编号", "发件人邮箱", "发件日期", "邮件主题",
                  "代理", "客户公司名称", "标准化项目名称", "需求", "置信度"])
    for index in (1, 2):
        sheet.append(["MAIL-UI", f"DETAIL-{index}", "fixture@example.test",
                      "2026-09-30 09:00:00", "Anonymous UI fixture", "Example Agent",
                      f"Example {index} LLC", "法国包装法", "注册", "high"])
    workbook.save(primary)
    workbook.close()
    return WorkbenchStore(primary_path=str(primary), review_path=str(tmp_path / "review.xlsx"),
                          filtered_path=str(tmp_path / "filtered.xlsx"),
                          state_path=str(tmp_path / "state.json"), test_mode=True)


def test_unified_edit_saves_agent_and_company_only_to_target(tmp_path):
    store = make_store(tmp_path)
    mail = store.snapshot()["mails"][0]
    first, second = mail["details"]
    fields = {**second["fields"], "agent": "Replacement Agent", "company": "Fixture e5f74360 LLC"}
    result = store.action({"action": "edit", "mail_id": mail["id"],
                          "record_id": second["id"], "fields": fields, "reason": ""})
    assert result["ok"]
    updated = {detail["id"]: detail for detail in store.snapshot()["mails"][0]["details"]}
    assert updated[first["id"]]["fields"] == first["fields"]
    assert updated[second["id"]]["fields"]["agent"] == "Replacement Agent"
    assert updated[second["id"]]["fields"]["company"] == "Fixture e5f74360 LLC"
    assert updated[second["id"]]["status"] == "review"
    assert len(updated[second["id"]]["events"]) == 1


def test_whole_mail_completion_does_not_bypass_pending_review(tmp_path):
    store = make_store(tmp_path)
    mail = store.snapshot()["mails"][0]
    first, second = mail["details"]
    store.action({"action": "edit", "record_id": second["id"], "fields": second["fields"]})
    result = store.action({"action": "bulk_confirm", "mail_ids": [mail["id"]]})
    assert result["skipped_count"] == 1
    assert result["confirmed_details"] == 0
    store.action({"action": "confirm_detail", "record_id": second["id"]})
    result = store.action({"action": "bulk_confirm", "mail_ids": [mail["id"]]})
    assert result["skipped_count"] == 0
    assert all(detail["status"] == "confirmed" for detail in store.snapshot()["mails"][0]["details"])
